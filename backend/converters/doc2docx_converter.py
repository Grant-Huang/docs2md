"""
旧版文档格式升级转换器

将 .doc → .docx，.xls → .xlsx，使后续转换流程可以正常处理这些文件。
优先使用 LibreOffice headless（跨平台），Windows 上在 LibreOffice 不可用时
退化为 win32com（需要安装 Microsoft Office）。
"""
import asyncio
import platform
import shutil
import subprocess
from pathlib import Path
from typing import Callable, Awaitable, Optional

# 旧版扩展名 → 目标格式
_LEGACY_TO_FORMAT = {
    ".doc": "docx",
    ".xls": "xlsx",
}

# 旧版扩展名 → 新版扩展名
_LEGACY_TO_NEW_EXT = {
    ".doc": ".docx",
    ".xls": ".xlsx",
}


def _find_soffice() -> Optional[str]:
    """查找系统中的 LibreOffice / soffice 可执行文件"""
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return None


def _convert_with_libreoffice(input_path: Path, output_dir: Path) -> bool:
    """
    使用 LibreOffice headless 转换单个文件到 output_dir。
    返回是否成功。
    """
    soffice = _find_soffice()
    if not soffice:
        return False

    ext = input_path.suffix.lower()
    target_format = _LEGACY_TO_FORMAT.get(ext)
    if not target_format:
        return False

    try:
        result = subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to", target_format,
                "--outdir", str(output_dir),
                str(input_path.absolute()),
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )
        out_file = output_dir / (input_path.stem + "." + target_format)
        return result.returncode == 0 and out_file.exists()
    except (subprocess.TimeoutExpired, OSError):
        return False


def _convert_with_win32com(input_path: Path, output_dir: Path) -> bool:
    """
    使用 win32com（仅 Windows，需要安装 Microsoft Office）转换单个文件。
    返回是否成功。
    """
    try:
        import win32com.client as win32  # type: ignore[import]
    except ImportError:
        return False

    ext = input_path.suffix.lower()
    try:
        if ext == ".doc":
            app = win32.Dispatch("Word.Application")
            app.Visible = False
            try:
                out_path = str((output_dir / f"{input_path.stem}.docx").absolute())
                doc = app.Documents.Open(str(input_path.absolute()))
                doc.SaveAs(out_path, FileFormat=12)  # 12 = wdFormatXMLDocument
                doc.Close()
            finally:
                app.Quit()
            return (output_dir / f"{input_path.stem}.docx").exists()

        elif ext == ".xls":
            app = win32.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                out_path = str((output_dir / f"{input_path.stem}.xlsx").absolute())
                wb = app.Workbooks.Open(str(input_path.absolute()))
                wb.SaveAs(out_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook
                wb.Close()
            finally:
                app.Quit()
            return (output_dir / f"{input_path.stem}.xlsx").exists()

    except Exception:
        return False

    return False


def _convert_single(input_path: Path, output_dir: Path) -> dict:
    """
    转换单个旧版文档，优先用 LibreOffice，Windows 上退化为 win32com。
    返回 {"input", "output", "success"}。
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    success = _convert_with_libreoffice(input_path, output_dir)

    if not success and platform.system() == "Windows":
        success = _convert_with_win32com(input_path, output_dir)

    new_ext = _LEGACY_TO_NEW_EXT.get(input_path.suffix.lower(), "")
    out_file = output_dir / (input_path.stem + new_ext) if new_ext else None

    return {
        "input": str(input_path),
        "output": str(out_file) if (success and out_file) else None,
        "success": success,
    }


async def convert_legacy_dir(
    input_dir: Path,
    sse_callback: Optional[Callable[[dict], Awaitable[None]]] = None,
) -> list:
    """
    递归扫描 input_dir，将所有 .doc / .xls 文件就地转换为 .docx / .xlsx。
    转换结果输出到原文件所在目录（与原文件并排），不删除原文件。
    若同名新版文件已存在则跳过。

    返回每个文件的转换结果列表。
    """
    async def emit(data: dict):
        if sse_callback:
            await sse_callback(data)

    # 收集所有旧版文件，排除 Office 临时文件（~$ 前缀）
    legacy_files: list[Path] = []
    for ext in _LEGACY_TO_NEW_EXT:
        legacy_files.extend(
            f for f in input_dir.rglob(f"*{ext}")
            if not f.name.startswith("~$")
        )
    legacy_files.sort()

    if not legacy_files:
        return []

    await emit({
        "type": "debug",
        "content": f"发现 {len(legacy_files)} 个旧版文件（.doc/.xls），开始格式升级...",
    })

    results = []
    for i, fp in enumerate(legacy_files):
        new_ext = _LEGACY_TO_NEW_EXT[fp.suffix.lower()]
        out_file = fp.parent / (fp.stem + new_ext)

        if out_file.exists():
            await emit({"type": "debug", "content": f"  跳过（已存在）: {fp.name}"})
            results.append({
                "input": str(fp),
                "output": str(out_file),
                "success": True,
                "skipped": True,
            })
            continue

        await emit({
            "type": "debug",
            "content": f"  升级 ({i + 1}/{len(legacy_files)}): {fp.name} → {fp.stem}{new_ext}",
        })

        result = await asyncio.to_thread(_convert_single, fp, fp.parent)
        results.append(result)

        if result["success"]:
            await emit({"type": "debug", "content": f"  ✓ {fp.name} → {fp.stem}{new_ext}"})
            # 删除原始旧版文件，避免后续流程重复处理
            try:
                fp.unlink(missing_ok=True)
            except OSError:
                pass
        else:
            await emit({
                "type": "debug",
                "content": f"  ✗ {fp.name} 转换失败，将尝试直接转换（可能不支持）",
            })

    success_count = sum(1 for r in results if r["success"] and not r.get("skipped"))
    skip_count = sum(1 for r in results if r.get("skipped"))
    fail_count = sum(1 for r in results if not r["success"])

    summary = f"格式升级完成：{success_count} 个成功"
    if skip_count:
        summary += f"，{skip_count} 个跳过（已存在）"
    if fail_count:
        summary += f"，{fail_count} 个失败"
    await emit({"type": "debug", "content": summary})

    return results
