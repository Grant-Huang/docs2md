#!/usr/bin/env python3
"""
docs2md CLI — 将文档/图片转为 Markdown 或纯文本

用法：
  docs2md <输入文件或目录> [选项]

选项：
  -o, --output PATH    输出文件或目录（默认：<输入>.md 或 <输入>_out/）
  -f, --format         输出格式：md（默认）或 txt
  -q, --quiet          只输出错误，不打印进度

支持格式：
  文档  .docx .doc .xlsx .xls .pdf .txt
  图片  .png .jpg .jpeg .gif .webp .bmp .tiff .tif
"""

from __future__ import annotations

import argparse
import asyncio
import sys
from pathlib import Path


def make_sse_callback(quiet: bool):
    async def _cb(data: dict):
        if quiet:
            return
        t = data.get("type", "")
        msg = data.get("content", "")
        if t == "debug":
            print(f"  {msg}")
        elif t == "error":
            print(f"  [错误] {msg}", file=sys.stderr)

    return _cb


async def convert_one(input_path: Path, output_dir: Path, fmt: str, quiet: bool) -> bool:
    from docs2md.converters.docx_converter import convert_docx
    from docs2md.converters.excel_converter import convert_excel
    from docs2md.converters.pdf_converter import convert_pdf
    from docs2md.converters.txt_converter import convert_txt
    from docs2md.converters.image_converter import convert_image, IMAGE_EXTENSIONS

    cb = make_sse_callback(quiet)
    ext = input_path.suffix.lower()

    if ext in (".docx", ".doc"):
        r = await convert_docx(input_path, output_dir, fmt, sse_callback=cb)
    elif ext in (".xlsx", ".xls"):
        r = await convert_excel(input_path, output_dir, fmt, sse_callback=cb)
    elif ext == ".pdf":
        r = await convert_pdf(input_path, output_dir, fmt, sse_callback=cb)
    elif ext == ".txt":
        r = await convert_txt(input_path, output_dir, fmt, sse_callback=cb)
    elif ext in IMAGE_EXTENSIONS:
        r = await convert_image(input_path, output_dir, fmt, sse_callback=cb)
    else:
        print(f"[跳过] 不支持的格式：{input_path.name}", file=sys.stderr)
        return False

    if r.get("error"):
        print(f"[失败] {input_path.name}：{r['error']}", file=sys.stderr)
        return False

    out = r.get("path", "")
    if not quiet:
        print(f"  -> {out}")
    return True


async def convert_dir(input_dir: Path, output_dir: Path, fmt: str, quiet: bool) -> None:
    from docs2md.utils.traversal import traverse_and_convert

    cb = make_sse_callback(quiet)
    results = await traverse_and_convert(input_dir, output_dir, fmt, sse_callback=cb)

    ok = sum(1 for r in results if not r.get("error") and r.get("path") != "index")
    fail = sum(1 for r in results if r.get("error"))
    index_entry = next((r for r in results if r.get("path") == "index"), None)

    print(f"\n完成：{ok} 个成功，{fail} 个失败。")
    if index_entry:
        print(f"索引：{index_entry.get('output', '')}")

    if fail:
        print("\n失败列表：", file=sys.stderr)
        for r in results:
            if r.get("error"):
                print(f"  {r['path']}：{r['error']}", file=sys.stderr)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="docs2md",
        description="将文档/图片转为 Markdown 或纯文本",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例：
  docs2md report.pdf
  docs2md report.pdf -o output/report.md
  docs2md ./docs/ -o ./output/ -f txt
  docs2md photo.png --quiet""",
    )
    p.add_argument("input", metavar="INPUT", help="输入文件或目录")
    p.add_argument("-o", "--output", metavar="PATH", help="输出文件或目录")
    p.add_argument(
        "-f",
        "--format",
        choices=["md", "txt"],
        default="md",
        help="输出格式，默认 md",
    )
    p.add_argument("-q", "--quiet", action="store_true", help="只输出错误信息")
    return p


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    inp = Path(args.input).resolve()
    if not inp.exists():
        print(f"错误：路径不存在：{inp}", file=sys.stderr)
        sys.exit(1)

    fmt: str = args.format
    quiet: bool = args.quiet

    if inp.is_dir():
        if args.output:
            out = Path(args.output).resolve()
        else:
            out = inp.parent / (inp.name + "_out")
        out.mkdir(parents=True, exist_ok=True)

        if not quiet:
            print(f"输入目录：{inp}")
            print(f"输出目录：{out}")
            print(f"输出格式：{fmt}\n")

        asyncio.run(convert_dir(inp, out, fmt, quiet))
        return

    if args.output:
        out_path = Path(args.output).resolve()
        output_dir = out_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)
    else:
        output_dir = inp.parent

    if not quiet:
        print(f"输入文件：{inp}")
        print(f"输出目录：{output_dir}")
        print(f"输出格式：{fmt}\n")

    ok = asyncio.run(convert_one(inp, output_dir, fmt, quiet))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()

