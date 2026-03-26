"""
目录遍历、输出路径映射、index 生成
（CLI 场景使用：不依赖服务端配置与仓库目录结构）
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import List, Tuple

from docs2md.converters.image_converter import IMAGE_EXTENSIONS

SUPPORTED_EXTENSIONS = {".docx", ".doc", ".xlsx", ".xls", ".pdf", ".txt"} | IMAGE_EXTENSIONS

# 旧版格式：有对应新版同名文件时跳过，避免重复处理
_LEGACY_EXTENSIONS = {".doc": ".docx", ".xls": ".xlsx"}


def _max_files_per_batch() -> int:
    # 兼容原服务端常量 MAX_FILES_PER_BATCH=200，同时支持环境变量覆盖
    try:
        return int(os.getenv("DOCS2MD_MAX_FILES_PER_BATCH", "200"))
    except ValueError:
        return 200


def collect_files(input_dir: Path) -> List[Path]:
    """
    递归收集支持的文档文件。
    若目录中同时存在 .doc 和同名 .docx（或 .xls 和同名 .xlsx），
    则跳过旧版文件，只保留新版（格式升级后的结果）。
    """
    files: List[Path] = []
    limit = _max_files_per_batch()
    for p in input_dir.rglob("*"):
        if not p.is_file():
            continue
        ext = p.suffix.lower()
        if ext not in SUPPORTED_EXTENSIONS:
            continue
        if ext in _LEGACY_EXTENSIONS:
            new_ext = _LEGACY_EXTENSIONS[ext]
            if (p.parent / (p.stem + new_ext)).exists():
                continue
        files.append(p)
        if len(files) >= limit:
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


def generate_index_md(results: List[Tuple[Path, Path]], format: str, output_root: Path) -> str:
    """生成 index.md，按原始目录结构分组列出所有转换文件。"""
    from collections import defaultdict

    groups: dict = defaultdict(list)

    for _inp, out in results:
        try:
            rel = out.relative_to(output_root)
        except ValueError:
            rel = Path(out.name)

        dir_key = rel.parent.as_posix()
        link = rel.as_posix()
        groups[dir_key].append((out.stem, link))

    lines = ["# 转换结果索引", ""]

    for dir_key in sorted(groups.keys()):
        title = dir_key if dir_key != "." else "根目录"
        lines.append(f"## {title}")
        lines.append("")
        for stem, link in sorted(groups[dir_key]):
            lines.append(f"- [{stem}]({link})")
        lines.append("")

    return "\n".join(lines).strip() + "\n"


async def traverse_and_convert(
    input_dir: Path,
    output_dir: Path,
    format: str,
    sse_callback=None,
) -> list[dict]:
    """
    遍历目录并批量转换，返回结果列表。
    返回元素格式与既有 CLI 兼容：成功包含 path/output，失败包含 error。
    """
    from docs2md.converters.doc2docx_converter import upgrade_legacy_files
    from docs2md.converters.docx_converter import convert_docx
    from docs2md.converters.excel_converter import convert_excel
    from docs2md.converters.pdf_converter import convert_pdf
    from docs2md.converters.txt_converter import convert_txt
    from docs2md.converters.image_converter import convert_image, IMAGE_EXTENSIONS

    async def emit(data: dict):
        if sse_callback:
            await sse_callback(data)

    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    await emit({"type": "debug", "content": "开始扫描目录..."})
    files = collect_files(input_dir)
    await emit({"type": "debug", "content": f"发现 {len(files)} 个待处理文件"})

    # 先做格式升级（.doc/.xls）
    await emit({"type": "debug", "content": "开始执行旧版格式升级（如需要）..."})
    upgraded = await upgrade_legacy_files(files, sse_callback=sse_callback)
    # 用升级后的文件列表继续处理
    files = upgraded

    results: list[dict] = []
    ok_pairs: List[Tuple[Path, Path]] = []

    for p in files:
        ext = p.suffix.lower()
        out_path = get_output_path(p, input_dir, output_dir, format)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        if ext in (".docx", ".doc"):
            r = await convert_docx(p, out_path.parent, format, sse_callback=sse_callback)
        elif ext in (".xlsx", ".xls"):
            r = await convert_excel(p, out_path.parent, format, sse_callback=sse_callback)
        elif ext == ".pdf":
            r = await convert_pdf(p, out_path.parent, format, sse_callback=sse_callback)
        elif ext == ".txt":
            r = await convert_txt(p, out_path.parent, format, sse_callback=sse_callback)
        elif ext in IMAGE_EXTENSIONS:
            r = await convert_image(p, out_path.parent, format, sse_callback=sse_callback)
        else:
            r = {"error": f"不支持的格式：{ext}", "path": str(p)}

        if r.get("error"):
            results.append({"path": str(p), "error": r["error"]})
            continue

        produced = Path(r.get("path", out_path))
        results.append({"path": str(p), "output": str(produced)})
        ok_pairs.append((p, produced))

    # 生成 index.md
    if format == "md":
        index_text = generate_index_md(ok_pairs, format, output_dir)
        index_path = output_dir / "index.md"
        index_path.write_text(index_text, encoding="utf-8")
        results.append({"path": "index", "output": str(index_path)})

    return results

