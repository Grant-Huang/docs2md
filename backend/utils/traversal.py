"""
目录遍历、输出路径映射、index 生成
"""
from pathlib import Path
from typing import List, Tuple

from backend.config import MAX_FILES_PER_BATCH

SUPPORTED_EXTENSIONS = {".docx", ".doc", ".xlsx", ".xls"}

# 旧版格式：有对应新版同名文件时跳过，避免重复处理
_LEGACY_EXTENSIONS = {".doc": ".docx", ".xls": ".xlsx"}


def collect_files(input_dir: Path) -> List[Path]:
    """
    递归收集支持的文档文件。
    若目录中同时存在 .doc 和同名 .docx（或 .xls 和同名 .xlsx），
    则跳过旧版文件，只保留新版（格式升级后的结果）。
    """
    files = []
    for p in input_dir.rglob("*"):
        if not p.is_file():
            continue
        ext = p.suffix.lower()
        if ext not in SUPPORTED_EXTENSIONS:
            continue
        # 旧版文件：若同名新版已存在则跳过
        if ext in _LEGACY_EXTENSIONS:
            new_ext = _LEGACY_EXTENSIONS[ext]
            if (p.parent / (p.stem + new_ext)).exists():
                continue
        files.append(p)
        if len(files) >= MAX_FILES_PER_BATCH:
            break
    return sorted(files)


def get_output_path(input_path: Path, input_dir: Path, output_dir: Path, format: str) -> Path:
    """根据输入路径计算输出路径，保持相对目录结构"""
    try:
        rel = input_path.relative_to(input_dir)
    except ValueError:
        rel = Path(input_path.name)
    ext = ".md" if format == "md" else ".txt"
    new_name = rel.stem + ext
    return output_dir / rel.parent / new_name


def generate_index_md(results: List[Tuple[Path, Path]], format: str) -> str:
    """生成 index.md 内容"""
    lines = ["# 转换结果索引\n", ""]
    for inp, out in results:
        name = out.stem
        rel = out.name
        lines.append(f"- [{name}]({rel})")
    return "\n".join(lines)


async def traverse_and_convert(
    input_dir: Path,
    output_dir: Path,
    format: str,
    sse_callback=None,
) -> List[dict]:
    """
    遍历输入目录，转换每个文件，生成 index。

    流程：
    1. 将目录中所有 .doc / .xls 升级为 .docx / .xlsx（LibreOffice 或 win32com）
    2. 收集所有可处理文件（优先使用升级后的新版文件）
    3. 依次调用 docx_converter / excel_converter 转为 Markdown/Text
    4. 生成 index.md

    返回每个文件的转换结果列表。
    """
    from backend.converters.doc2docx_converter import convert_legacy_dir
    from backend.converters.docx_converter import convert_docx
    from backend.converters.excel_converter import convert_excel

    # 阶段1：旧版格式升级
    await convert_legacy_dir(input_dir, sse_callback=sse_callback)

    # 阶段2：收集文件（旧版文件已被删除，直接收集新版）
    files = collect_files(input_dir)
    total = len(files)

    async def emit(data: dict):
        if sse_callback:
            await sse_callback(data)

    await emit({"type": "debug", "content": f"发现 {total} 个文件，开始转换..."})

    results = []
    converted = []  # (input_path, output_path)

    for i, fp in enumerate(files):
        await emit({"type": "debug", "content": f"解析第 {i + 1}/{total} 个文件：{fp.name}"})
        out_path = get_output_path(fp, input_dir, output_dir, format)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        ext = fp.suffix.lower()
        if ext in (".docx", ".doc"):
            r = await convert_docx(fp, out_path.parent, format, sse_callback=sse_callback)
        elif ext in (".xlsx", ".xls"):
            r = await convert_excel(fp, out_path.parent, format, sse_callback=sse_callback)
        else:
            continue

        if r.get("error"):
            results.append({"path": str(fp), "error": r["error"]})
        else:
            results.append({
                "path": str(fp),
                "output": r.get("path", ""),
                "content": r.get("content", "")[:500],
            })
            converted.append((fp, out_path))

    if converted:
        index_content = generate_index_md(converted, format)
        index_path = output_dir / "index.md"
        index_path.write_text(index_content, encoding="utf-8")
        results.append({"path": "index", "output": str(index_path), "content": index_content})

    return results
