"""
旧版文档格式升级转换器

将 .doc → .docx，.xls → .xlsx，使后续转换流程可以正常处理这些文件。
优先使用 LibreOffice headless（跨平台），Windows 上在 LibreOffice 不可用时
退化为 win32com（需要安装 Microsoft Office）。
"""

from __future__ import annotations

import asyncio
import platform
import shutil
import subprocess
from pathlib import Path
from typing import Awaitable, Callable, Optional

_LEGACY_TO_FORMAT = {".doc": "docx", ".xls": "xlsx"}
_LEGACY_TO_NEW_EXT = {".doc": ".docx", ".xls": ".xlsx"}


def _find_soffice() -> Optional[str]:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path

    if platform.system() == "Darwin":
        import os

        mac_candidates = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        ]
        for p in mac_candidates:
            if Path(p).is_file() and os.access(p, os.X_OK):
                return p

    return None


def _convert_with_libreoffice(input_path: Path, output_dir: Path) -> bool:
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
                "--convert-to",
                target_format,
                "--outdir",
                str(output_dir),
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
                doc.SaveAs(out_path, FileFormat=12)
                doc.Close()
            finally:
                app.Quit()
            return (output_dir / f"{input_path.stem}.docx").exists()

        if ext == ".xls":
            app = win32.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                out_path = str((output_dir / f"{input_path.stem}.xlsx").absolute())
                wb = app.Workbooks.Open(str(input_path.absolute()))
                wb.SaveAs(out_path, FileFormat=51)
                wb.Close()
            finally:
                app.Quit()
            return (output_dir / f"{input_path.stem}.xlsx").exists()

    except Exception:
        return False

    return False


async def upgrade_legacy_files(
    files: list[Path],
    sse_callback: Optional[Callable[[dict], Awaitable[None]]] = None,
) -> list[Path]:
    async def emit(data: dict):
        if sse_callback:
            await sse_callback(data)

    out: list[Path] = []
    for p in files:
        ext = p.suffix.lower()
        if ext not in _LEGACY_TO_NEW_EXT:
            out.append(p)
            continue

        output_dir = p.parent
        new_path = output_dir / (p.stem + _LEGACY_TO_NEW_EXT[ext])

        if new_path.exists():
            await emit({"type": "debug", "content": f"跳过旧版（已存在新版）：{p.name}"})
            continue

        await emit({"type": "debug", "content": f"正在升级旧版格式：{p.name} -> {new_path.name}"})

        ok = await asyncio.to_thread(_convert_with_libreoffice, p, output_dir)
        if not ok and platform.system() == "Windows":
            ok = await asyncio.to_thread(_convert_with_win32com, p, output_dir)

        if ok and new_path.exists():
            out.append(new_path)
        else:
            out.append(p)
            await emit({"type": "debug", "content": f"旧版升级失败（继续按原文件处理）：{p.name}"})

    return out

