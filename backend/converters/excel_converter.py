"""
Excel 转 Markdown 转换器

- .xlsx / 新版格式：使用 MarkItDown（底层 openpyxl）
- .xls  / 旧版格式：使用 xlrd 原生读取（openpyxl 不支持 BIFF 格式）
多 sheet 合并为单文件，每个 sheet 用 ## sheet名 作为标题。
"""
import asyncio
from pathlib import Path
from typing import Optional, Callable, Awaitable


async def convert_excel(
    input_path: Path,
    output_dir: Path,
    format: str,
    sse_callback: Optional[Callable[[dict], Awaitable[None]]] = None,
) -> dict:
    """
    转换 Excel 为 Markdown 或纯文本。
    .xls 使用 xlrd，.xlsx 使用 MarkItDown。
    """
    async def emit(data: dict):
        if sse_callback:
            await sse_callback(data)

    try:
        await emit({"type": "debug", "content": f"正在转换: {input_path.name}"})

        ext = input_path.suffix.lower()
        if ext == ".xls":
            text = await asyncio.to_thread(_convert_xls_xlrd, input_path)
        else:
            text = await asyncio.to_thread(_convert_xlsx_markitdown, input_path)

        if not text or not text.strip():
            return {"error": "未能提取内容"}

        if format == "txt":
            text = _md_to_plain(text)

        output_dir.mkdir(parents=True, exist_ok=True)
        out_name = input_path.stem + (".md" if format == "md" else ".txt")
        out_path = output_dir / out_name
        out_path.write_text(text, encoding="utf-8")

        await emit({"type": "complete", "content": text[:5000] + ("..." if len(text) > 5000 else "")})
        return {"content": text, "path": str(out_path)}
    except Exception as e:
        await emit({"type": "error", "content": str(e)})
        return {"error": str(e)}


def _convert_xlsx_markitdown(input_path: Path) -> str:
    """使用 MarkItDown（openpyxl）转换 .xlsx 文件"""
    from markitdown import MarkItDown
    md = MarkItDown()
    result = md.convert(str(input_path))
    if not result or not getattr(result, "text_content", None):
        raise ValueError("MarkItDown 未能提取内容")
    return result.text_content


def _convert_xls_xlrd(input_path: Path) -> str:
    """使用 xlrd 读取 .xls 文件，输出 Markdown 表格"""
    try:
        import xlrd
    except ImportError:
        raise RuntimeError("缺少依赖 xlrd，请执行 pip install xlrd")

    wb = xlrd.open_workbook(str(input_path))
    parts: list[str] = []

    for sheet_name in wb.sheet_names():
        sheet = wb.sheet_by_name(sheet_name)
        if sheet.nrows == 0:
            continue

        parts.append(f"## {sheet_name}")
        parts.append("")

        rows: list[str] = []
        for row_idx in range(sheet.nrows):
            cells = []
            for col_idx in range(sheet.ncols):
                cell = sheet.cell(row_idx, col_idx)
                # 整数型数值去掉多余的 .0
                val = cell.value
                if isinstance(val, float) and val == int(val):
                    val = int(val)
                cells.append(str(val).strip().replace("|", "\\|"))
            rows.append("| " + " | ".join(cells) + " |")

        if rows:
            sep = "| " + " | ".join(["---"] * sheet.ncols) + " |"
            parts.append(rows[0])   # 表头
            parts.append(sep)
            parts.extend(rows[1:])  # 数据行

        parts.append("")

    return "\n".join(parts)


def _md_to_plain(md: str) -> str:
    """简单将 Markdown 转为纯文本"""
    import re
    t = re.sub(r"^#+\s*", "", md, flags=re.MULTILINE)
    lines = []
    for line in t.split("\n"):
        if line.strip().startswith("|") and line.strip().endswith("|"):
            cells = [c.strip() for c in line.split("|")[1:-1]]
            lines.append("\t".join(cells))
        else:
            lines.append(line)
    return "\n".join(lines)
