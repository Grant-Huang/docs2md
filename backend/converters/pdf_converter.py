"""
PDF 转 Markdown 转换器
支持文字型 PDF（直接提取文本）和图片型 PDF（逐页渲染后调用 Qwen3-VL 解析）
"""
import asyncio
import tempfile
from pathlib import Path
from typing import Optional, Callable, Awaitable

# 文字型判定：每页平均字符数低于此阈值则视为图片页
_TEXT_THRESHOLD = 50


async def convert_pdf(
    input_path: Path,
    output_dir: Path,
    format: str,
    sse_callback: Optional[Callable[[dict], Awaitable[None]]] = None,
) -> dict:
    """
    转换 PDF 为 Markdown 或纯文本。
    - 文字型页面：直接提取文本。
    - 图片型页面：渲染为 PNG 后调用 Qwen3-VL 解析。
    - 混合型 PDF：逐页判断，分别处理。
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return {"error": "缺少依赖 pymupdf，请执行 pip install pymupdf"}

    async def emit(data: dict):
        if sse_callback:
            await sse_callback(data)

    try:
        await emit({"type": "debug", "content": f"正在打开 PDF：{input_path.name}"})

        doc = fitz.open(str(input_path))
        total_pages = len(doc)
        await emit({"type": "debug", "content": f"共 {total_pages} 页，开始逐页处理..."})

        # 准备 assets 目录（用于存放图片型页面的截图）
        output_dir.mkdir(parents=True, exist_ok=True)
        assets_dir = output_dir / "assets"

        parts: list[str] = []
        has_image_pages = False

        for page_num in range(total_pages):
            page = doc[page_num]
            page_label = f"第 {page_num + 1}/{total_pages} 页"
            await emit({"type": "debug", "content": f"处理 {page_label}：{input_path.name}"})

            # 尝试提取文本
            text = page.get_text("text").strip()

            if len(text) >= _TEXT_THRESHOLD:
                # 文字型页面
                page_md = _text_to_md(text, page_num + 1)
                parts.append(page_md)
            else:
                # 图片型页面：渲染并交给 VL 解析
                has_image_pages = True
                assets_dir.mkdir(parents=True, exist_ok=True)
                img_name = f"{input_path.stem}_page_{page_num + 1:04d}.png"
                img_path = assets_dir / img_name

                # 以 2x 分辨率渲染页面
                mat = fitz.Matrix(2, 2)
                pix = page.get_pixmap(matrix=mat)
                pix.save(str(img_path))

                await emit({"type": "debug", "content": f"{page_label} 为图片型，调用 VL 解析..."})

                # 调用 Qwen3-VL（在线程中执行，避免阻塞事件循环）
                try:
                    from backend.services.qwen_vl import analyze_image
                    vl_result = await asyncio.to_thread(analyze_image, img_path)
                except Exception as e:
                    vl_result = f"[图片解析失败: {e}]"

                # 构造此页的 Markdown 内容
                rel_img = f"assets/{img_name}"
                page_md = (
                    f"## 第 {page_num + 1} 页\n\n"
                    f"![{img_name}]({rel_img})\n\n"
                    f"> {vl_result.replace(chr(10), chr(10) + '> ')}\n"
                )
                parts.append(page_md)

                await emit({"type": "partial", "content": "\n\n---\n\n".join(parts)})

        doc.close()

        content = "\n\n---\n\n".join(parts)

        if format == "txt":
            content = _md_to_plain(content)

        out_ext = ".md" if format == "md" else ".txt"
        out_path = output_dir / (input_path.stem + out_ext)
        out_path.write_text(content, encoding="utf-8")

        await emit({"type": "complete", "content": content[:5000] + ("..." if len(content) > 5000 else "")})
        return {"content": content, "path": str(out_path)}

    except Exception as e:
        await emit({"type": "error", "content": str(e)})
        return {"error": str(e)}


def _text_to_md(text: str, page_num: int) -> str:
    """将页面纯文本格式化为 Markdown（保留段落结构）"""
    lines = text.splitlines()
    # 过滤空行并合并段落
    cleaned = []
    for line in lines:
        stripped = line.strip()
        if stripped:
            cleaned.append(stripped)
        elif cleaned and cleaned[-1] != "":
            cleaned.append("")
    body = "\n".join(cleaned).strip()
    return f"## 第 {page_num} 页\n\n{body}"


def _md_to_plain(md: str) -> str:
    """简单去除 Markdown 标记，输出纯文本"""
    import re
    t = re.sub(r"^#+\s*", "", md, flags=re.MULTILINE)
    t = re.sub(r"!\[.*?\]\(.*?\)", "", t)
    t = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", t)
    t = re.sub(r"^>\s?", "", t, flags=re.MULTILINE)
    t = re.sub(r"---\n", "", t)
    return t.strip()
