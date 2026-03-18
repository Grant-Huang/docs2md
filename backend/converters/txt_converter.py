"""
TXT 文件转换器
将纯文本文件直接写入输出目录，保留原始内容。
输出 md 格式时内容不变；输出 txt 格式时同样不变（已是纯文本）。
"""
from pathlib import Path
from typing import Optional, Callable, Awaitable


async def convert_txt(
    input_path: Path,
    output_dir: Path,
    format: str,
    sse_callback: Optional[Callable[[dict], Awaitable[None]]] = None,
) -> dict:
    """
    将 .txt 文件转为 Markdown 或纯文本输出。
    文件名保持原名，扩展名替换为 .md 或 .txt。
    """
    async def emit(data: dict):
        if sse_callback:
            await sse_callback(data)

    try:
        await emit({"type": "debug", "content": f"正在读取文本文件：{input_path.name}"})

        # 尝试 UTF-8，失败时退回 GBK（兼容 Windows 中文文本文件）
        try:
            content = input_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            content = input_path.read_text(encoding="gbk", errors="replace")

        output_dir.mkdir(parents=True, exist_ok=True)
        out_ext = ".md" if format == "md" else ".txt"
        out_path = output_dir / (input_path.stem + out_ext)
        out_path.write_text(content, encoding="utf-8")

        preview = content[:5000] + ("..." if len(content) > 5000 else "")
        await emit({"type": "complete", "content": preview})
        return {"content": content, "path": str(out_path)}

    except Exception as e:
        await emit({"type": "error", "content": str(e)})
        return {"error": str(e)}
