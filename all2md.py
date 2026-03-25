#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
all2md — 文档资料知识化 CLI

用法：
  单文件转换
    python all2md.py input.docx
    python all2md.py input.xlsx -o output_dir/ --format txt

  目录批量转换（自动将 .doc/.xls 升级为 .docx/.xlsx 后再转换）
    python all2md.py input_dir/
    python all2md.py input_dir/ -o output_dir/ --format txt

支持格式：.docx .doc .xlsx .xls
"""
import argparse
import asyncio
import sys
from pathlib import Path


# ── 终端进度回调 ──────────────────────────────────────────────────────────────

def _make_cli_callback(verbose: bool):
    """返回一个将 SSE 消息打印到终端的异步回调"""
    async def callback(data: dict):
        kind = data.get("type", "")
        content = data.get("content", "")
        if kind == "debug":
            if verbose or not content.startswith("[正在解析图片"):
                print(f"  {content}", flush=True)
        elif kind == "error":
            print(f"  [错误] {content}", file=sys.stderr, flush=True)
        elif kind == "partial":
            pass  # CLI 不实时显示预览内容
    return callback


# ── 单文件转换 ────────────────────────────────────────────────────────────────

async def _convert_single(input_path: Path, output_dir: Path, fmt: str, verbose: bool) -> int:
    """转换单个文件，返回 0（成功）或 1（失败）"""
    from backend.converters.docx_converter import convert_docx
    from backend.converters.excel_converter import convert_excel

    ext = input_path.suffix.lower()
    callback = _make_cli_callback(verbose)

    print(f"转换: {input_path}  →  {output_dir}/")

    if ext in (".docx", ".doc"):
        result = await convert_docx(input_path, output_dir, fmt, sse_callback=callback)
    elif ext in (".xlsx", ".xls"):
        result = await convert_excel(input_path, output_dir, fmt, sse_callback=callback)
    else:
        print(f"不支持的格式: {ext}  （支持：.docx .doc .xlsx .xls）", file=sys.stderr)
        return 1

    if result.get("error"):
        print(f"[错误] {result['error']}", file=sys.stderr)
        return 1

    out_path = result.get("path", "")
    print(f"完成 → {out_path}")
    return 0


# ── 目录批量转换 ───────────────────────────────────────────────────────────────

async def _convert_dir(input_dir: Path, output_dir: Path, fmt: str, verbose: bool) -> int:
    """批量转换目录，返回失败文件数"""
    from backend.utils.traversal import traverse_and_convert

    output_dir.mkdir(parents=True, exist_ok=True)
    callback = _make_cli_callback(verbose)

    print(f"批量转换: {input_dir}  →  {output_dir}/")

    results = await traverse_and_convert(input_dir, output_dir, fmt, sse_callback=callback)

    ok = [r for r in results if not r.get("error") and r.get("path") != "index"]
    err = [r for r in results if r.get("error")]

    print(f"\n转换完成：{len(ok)} 个成功", end="")
    if err:
        print(f"，{len(err)} 个失败：")
        for r in err:
            print(f"  ✗ {r['path']} — {r['error']}", file=sys.stderr)
    else:
        print()

    idx = next((r for r in results if r.get("path") == "index"), None)
    if idx:
        print(f"索引文件 → {idx['output']}")

    return len(err)


# ── 主入口 ────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(
        prog="all2md",
        description="将 Word / Excel 文档转换为 Markdown 或纯文本",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例：
  python all2md.py report.docx
  python all2md.py data.xlsx --format txt -o results/
  python all2md.py docs/ -o output/ --format md --verbose
""",
    )
    parser.add_argument(
        "input",
        help="输入文件（.docx/.doc/.xlsx/.xls）或输入目录",
    )
    parser.add_argument(
        "-o", "--output",
        help="输出目录（默认：单文件同目录，目录模式为 <input>_output）",
    )
    parser.add_argument(
        "--format", "-f",
        choices=["md", "txt"],
        default="md",
        help="输出格式：md（Markdown，默认）或 txt（纯文本）",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="显示详细进度（含图片解析占位符信息）",
    )

    args = parser.parse_args()
    input_path = Path(args.input)

    if not input_path.exists():
        print(f"输入路径不存在: {input_path}", file=sys.stderr)
        return 1

    if input_path.is_dir():
        # 目录模式
        output_dir = Path(args.output) if args.output else input_path.parent / (input_path.name + "_output")
        return asyncio.run(_convert_dir(input_path, output_dir, args.format, args.verbose))
    else:
        # 单文件模式
        output_dir = Path(args.output) if args.output else input_path.parent
        return asyncio.run(_convert_single(input_path, output_dir, args.format, args.verbose))


if __name__ == "__main__":
    sys.exit(main())
