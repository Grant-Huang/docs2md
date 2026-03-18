"""
Excel 转 Markdown 转换器
使用 MarkItDown，多工作簿合并为单 md，表头为 sheet 名
"""
import asyncio
from pathlib import Path
from typing import Optional, Callable, Awaitable

from markitdown import MarkItDown


async def convert_excel(
    input_path: Path,
    output_dir: Path,
    format: str,
    sse_callback: Optional[Callable[[dict], Awaitable[None]]] = None,
) -> dict:
    """
    转换 Excel 为 Markdown 或纯文本
    多 sheet 合并为单文件，每个 sheet 用 ## sheet名 作为标题
    """
    async def emit(data: dict):
        if sse_callback:
            await sse_callback(data)

    try:
        await emit({"type": "debug", "content": f"正在转换: {input_path.name}"})

        md = MarkItDown()
        result = md.convert(str(input_path))

        if not result or not getattr(result, "text_content", None):
            return {"error": "未能提取内容"}

        text = result.text_content

        # MarkItDown 多 sheet 输出已包含表格，需确保每个 sheet 有标题
        # 若原始输出无 sheet 标题，可解析 result 结构补充
        # 默认 MarkItDown 会为每个 sheet 生成 ## SheetName 形式
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


def _md_to_plain(md: str) -> str:
    """简单将 Markdown 转为纯文本"""
    import re
    # 去掉 ## 标题
    t = re.sub(r"^#+\s*", "", md, flags=re.MULTILINE)
    # 表格行保留，去掉 | 分隔符的表格格式可简化
    lines = []
    for line in t.split("\n"):
        if line.strip().startswith("|") and line.strip().endswith("|"):
            cells = [c.strip() for c in line.split("|")[1:-1]]
            lines.append("\t".join(cells))
        else:
            lines.append(line)
    return "\n".join(lines)
