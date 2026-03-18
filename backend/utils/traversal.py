"""
目录遍历、输出路径映射、index 生成
"""
from pathlib import Path
from typing import List, Tuple, AsyncGenerator

from backend.config import MAX_FILES_PER_BATCH

SUPPORTED_EXTENSIONS = {".docx", ".doc", ".xlsx", ".xls"}


def collect_files(input_dir: Path) -> List[Path]:
    """递归收集支持的文档文件"""
    files = []
    for p in input_dir.rglob("*"):
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS:
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
) -> List[dict]:
    """
    遍历输入目录，转换每个文件，生成 index
    返回每个文件的转换结果列表
    """
    from backend.converters.docx_converter import convert_docx
    from backend.converters.excel_converter import convert_excel

    files = collect_files(input_dir)
    results = []
    converted = []  # (input_path, output_path)

    for fp in files:
        out_path = get_output_path(fp, input_dir, output_dir, format)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        ext = fp.suffix.lower()
        if ext in (".docx", ".doc"):
            r = await convert_docx(fp, out_path.parent, format)
        elif ext in (".xlsx", ".xls"):
            r = await convert_excel(fp, out_path.parent, format)
        else:
            continue

        if r.get("error"):
            results.append({"path": str(fp), "error": r["error"]})
        else:
            results.append({"path": str(fp), "output": r.get("path", ""), "content": r.get("content", "")[:500]})
            converted.append((fp, out_path))

    if converted:
        index_content = generate_index_md(converted, format)
        index_path = output_dir / "index.md"
        index_path.write_text(index_content, encoding="utf-8")
        results.append({"path": "index", "output": str(index_path), "content": index_content})

    return results
