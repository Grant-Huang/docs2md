#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docs2md MCP Server

将 Word / Excel → Markdown 转换能力封装为 MCP 工具，
供 Claude 桌面版或任何 MCP 客户端直接调用。

启动方式：
  python mcp_server.py          # stdio 模式（标准 MCP 客户端对接）

Claude 桌面版配置（claude_desktop_config.json）：
  {
    "mcpServers": {
      "docs2md": {
        "command": "python",
        "args": ["/absolute/path/to/docs2md/mcp_server.py"]
      }
    }
  }
"""
import sys
from pathlib import Path

# 确保项目根目录在 sys.path 中，使 backend.* 包可以正常导入
_ROOT = Path(__file__).resolve().parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP(
    "docs2md",
    instructions="""
你可以使用以下工具将 Word / Excel 文档转换为 Markdown 或纯文本：
- convert_file：转换单个文件（.docx、.doc、.xlsx、.xls）
- convert_directory：批量转换整个目录（自动升级 .doc/.xls 旧版格式）

两个工具均支持 format="md"（默认，Markdown）或 format="txt"（纯文本）。
转换结果直接返回内容，并告知输出文件路径。
""",
)

_SUPPORTED = {".docx", ".doc", ".xlsx", ".xls"}


def _make_log_callback(logs: list) -> object:
    """返回一个收集 debug 消息的异步回调"""
    async def _cb(data: dict):
        if data.get("type") == "debug":
            logs.append(data["content"])
    return _cb


# ── 工具：单文件转换 ──────────────────────────────────────────────────────────

@mcp.tool()
async def convert_file(
    input_path: str,
    output_dir: str = "",
    format: str = "md",
) -> str:
    """将单个 Word 或 Excel 文件转换为 Markdown 或纯文本。

    Args:
        input_path: 输入文件的绝对路径（.docx / .doc / .xlsx / .xls）。
        output_dir: 输出目录的绝对路径。留空则输出到输入文件所在目录。
        format: 输出格式，"md"（Markdown，默认）或 "txt"（纯文本）。

    Returns:
        转换后的文本内容，以及输出文件路径。
    """
    from backend.converters.docx_converter import convert_docx
    from backend.converters.excel_converter import convert_excel

    inp = Path(input_path)
    if not inp.exists():
        return f"错误：文件不存在 — {input_path}"

    ext = inp.suffix.lower()
    if ext not in _SUPPORTED:
        return f"错误：不支持格式 {ext}，支持的格式：{', '.join(sorted(_SUPPORTED))}"

    if format not in ("md", "txt"):
        return "错误：format 必须为 'md' 或 'txt'"

    out_dir = Path(output_dir).resolve() if output_dir else inp.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    logs: list[str] = []
    cb = _make_log_callback(logs)

    if ext in (".docx", ".doc"):
        result = await convert_docx(inp, out_dir, format, sse_callback=cb)
    else:
        result = await convert_excel(inp, out_dir, format, sse_callback=cb)

    if result.get("error"):
        return f"转换失败：{result['error']}"

    out_path = result.get("path", "（路径未知）")
    content = result.get("content", "")

    header = f"已输出到：{out_path}\n\n"
    return header + content


# ── 工具：目录批量转换 ────────────────────────────────────────────────────────

@mcp.tool()
async def convert_directory(
    input_dir: str,
    output_dir: str = "",
    format: str = "md",
) -> str:
    """批量转换目录中所有 Word / Excel 文档为 Markdown 或纯文本。

    自动将 .doc / .xls 旧版格式升级为 .docx / .xlsx，完成后生成 index.md。

    Args:
        input_dir: 输入目录的绝对路径。
        output_dir: 输出目录的绝对路径。留空则输出到 <input_dir>_output/。
        format: 输出格式，"md"（Markdown，默认）或 "txt"（纯文本）。

    Returns:
        转换摘要：成功/失败文件列表及输出路径。
    """
    from backend.utils.traversal import traverse_and_convert

    inp = Path(input_dir)
    if not inp.exists() or not inp.is_dir():
        return f"错误：目录不存在 — {input_dir}"

    if format not in ("md", "txt"):
        return "错误：format 必须为 'md' 或 'txt'"

    out_dir = Path(output_dir).resolve() if output_dir else inp.parent / (inp.name + "_output")
    out_dir.mkdir(parents=True, exist_ok=True)

    logs: list[str] = []
    cb = _make_log_callback(logs)

    results = await traverse_and_convert(inp, out_dir, format, sse_callback=cb)

    ok = [r for r in results if not r.get("error") and r.get("path") != "index"]
    err = [r for r in results if r.get("error")]
    idx = next((r for r in results if r.get("path") == "index"), None)

    lines = [
        f"转换完成：{len(ok)} 个成功，{len(err)} 个失败",
        f"输出目录：{out_dir}",
    ]
    if idx:
        lines.append(f"索引文件：{idx['output']}")
    if ok:
        lines.append("\n已转换：")
        for r in ok:
            lines.append(f"  ✓ {Path(r['path']).name}  →  {r.get('output', '')}")
    if err:
        lines.append("\n失败：")
        for r in err:
            lines.append(f"  ✗ {Path(r['path']).name}  —  {r.get('error', '')}")

    return "\n".join(lines)


# ── 入口 ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    mcp.run()
