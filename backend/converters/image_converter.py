"""
图片文件转 Markdown 转换器
支持 .png / .jpg / .jpeg / .gif / .webp / .bmp / .tiff 等常见格式。
生成的 Markdown 文件：
  - 文件名与原图片同名（扩展名改为 .md）
  - 将图片复制到 assets/ 子目录并在 md 中内嵌引用
  - 若配置了 Qwen3-VL，则追加 AI 解析描述
"""
import asyncio
import shutil
from pathlib import Path
from typing import Optional, Callable, Awaitable

# 支持的图片扩展名
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tiff", ".tif"}


async def convert_image(
    input_path: Path,
    output_dir: Path,
    format: str,
    sse_callback: Optional[Callable[[dict], Awaitable[None]]] = None,
) -> dict:
    """
    将图片文件转为 Markdown（或 txt）。
    - 图片复制到 output_dir/assets/<原文件名>
    - Markdown 中使用相对路径 ![name](assets/name) 引用
    - 若配置了 DASHSCOPE_API_KEY，调用 Qwen3-VL 追加图片描述
    """
    async def emit(data: dict):
        if sse_callback:
            await sse_callback(data)

    try:
        await emit({"type": "debug", "content": f"正在处理图片：{input_path.name}"})

        output_dir.mkdir(parents=True, exist_ok=True)
        assets_dir = output_dir / "assets"
        assets_dir.mkdir(parents=True, exist_ok=True)

        # 复制图片到 assets/
        dest_img = assets_dir / input_path.name
        shutil.copy2(str(input_path), str(dest_img))

        rel_img = f"assets/{input_path.name}"
        title = input_path.stem

        # 构造初步 Markdown（先不带解析结果）
        md_lines = [
            f"# {title}",
            "",
            f"![{input_path.name}]({rel_img})",
            "",
        ]

        # 调用 Qwen3-VL 解析图片
        await emit({"type": "debug", "content": f"调用 VL 解析图片：{input_path.name}"})
        try:
            from backend.services.qwen_vl import analyze_image
            vl_result = await asyncio.to_thread(analyze_image, dest_img)
            if vl_result and not vl_result.startswith("[未配置"):
                md_lines.append("## 图片内容描述")
                md_lines.append("")
                # 将多行结果转为 blockquote
                for line in vl_result.splitlines():
                    md_lines.append(f"> {line}" if line.strip() else ">")
                md_lines.append("")
        except Exception as e:
            md_lines.append(f"> [图片解析失败: {e}]")
            md_lines.append("")

        content = "\n".join(md_lines)

        if format == "txt":
            # txt 模式：去掉 Markdown 标记，保留图片路径和文字描述
            import re
            content = re.sub(r"^#+\s*", "", content, flags=re.MULTILINE)
            content = re.sub(r"^>\s?", "", content, flags=re.MULTILINE)
            content = content.strip()

        out_ext = ".md" if format == "md" else ".txt"
        out_path = output_dir / (input_path.stem + out_ext)
        out_path.write_text(content, encoding="utf-8")

        await emit({"type": "complete", "content": content[:5000] + ("..." if len(content) > 5000 else "")})
        return {"content": content, "path": str(out_path)}

    except Exception as e:
        await emit({"type": "error", "content": str(e)})
        return {"error": str(e)}
