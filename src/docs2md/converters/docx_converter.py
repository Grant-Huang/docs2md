"""
Word 文档转 Markdown 转换器
使用 python-docx，支持图片提取 + Qwen3-VL 解析后以引用形式插入
"""

from __future__ import annotations

import asyncio
import re
from pathlib import Path
from typing import Awaitable, Callable, List, Optional

from docx import Document
from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

from docs2md.services.qwen_vl import analyze_image, is_image_parse_enabled


def _dedupe_row(cells: List[str]) -> List[str]:
    """合并连续重复单元格"""
    if not cells:
        return cells
    result = [cells[0]]
    for c in cells[1:]:
        if c != result[-1]:
            result.append(c)
    return result


def _table_to_markdown(table: Table, output_format: str) -> str:
    """表格转 Markdown"""
    rows = []
    for row in table.rows:
        cells = [cell.text.strip().replace("\n", " ").replace("|", "\\|") for cell in row.cells]
        cells = _dedupe_row(cells)
        if cells:
            rows.append("| " + " | ".join(cells) + " |")
    if not rows:
        return ""
    col_count = max(len(r.split("|")) - 2 for r in rows) or 1
    sep = "| " + " | ".join(["---"] * col_count) + " |"
    return rows[0] + "\n" + sep + "\n" + "\n".join(rows[1:])


def _is_image_rid_attr(attr: str, val: str) -> bool:
    """判断是否为图片引用的属性：DrawingML 用 r:embed，VML 用 r:id"""
    if not val or not val.startswith("rId"):
        return False
    al = attr.lower()
    if al.endswith("}embed") or attr == "embed":
        return True
    if al.endswith("}id") or al == "id":
        return True
    return False


def _get_all_image_rids_from_document(doc: Document) -> List[str]:
    """从整个文档 body 中提取所有图片 rId（按出现顺序，去重）"""
    rids: List[str] = []
    seen = set()
    try:
        body = doc.element.body
        for elem in body.iter():
            for attr, val in elem.attrib.items():
                if _is_image_rid_attr(attr, val) and val not in seen:
                    seen.add(val)
                    rids.append(val)
                    break
    except Exception:
        pass
    return rids


def _build_rid_to_block_index(doc: Document, blocks: List) -> dict:
    """遍历 body XML，建立每个图片 rId 到块索引的映射。"""
    rid_to_block: dict = {}
    try:
        body = doc.element.body
        block_elements = [b._element for b in blocks]
        body_children = list(body.iterchildren())

        for elem in body.iter():
            for attr, val in elem.attrib.items():
                if not _is_image_rid_attr(attr, val) or val in rid_to_block:
                    continue
                block_idx = -1
                current = elem
                while current is not None:
                    for i, bel in enumerate(block_elements):
                        if bel is current:
                            block_idx = i
                            break
                    if block_idx >= 0:
                        break
                    parent = current.getparent()
                    if parent is body:
                        after_k = _find_insert_after_block(current, body_children, block_elements)
                        block_idx = -(after_k + 2) if after_k >= 0 else -1
                        break
                    current = parent
                rid_to_block[val] = block_idx
    except Exception:
        pass
    return rid_to_block


def _find_insert_after_block(elem, body_children: List, block_elements: List) -> int:
    """当图片在非块元素中时，找到其前一个块，插入到该块之后。"""
    try:
        for i, child in enumerate(body_children):
            if child is elem:
                for j in range(i - 1, -1, -1):
                    prev = body_children[j]
                    for k, bel in enumerate(block_elements):
                        if bel is prev:
                            return k
                return 0
        if elem.getparent() is not None:
            parent = elem.getparent()
            for i, child in enumerate(body_children):
                if child is parent:
                    for j in range(i - 1, -1, -1):
                        prev = body_children[j]
                        for k, bel in enumerate(block_elements):
                            if bel is prev:
                                return k
                    return 0
    except Exception:
        pass
    return 0


def _get_image_blob(part, rid: str) -> Optional[bytes]:
    """根据 rId 从指定 part 的 rels 获取图片二进制数据"""
    if rid not in part.rels:
        return None
    rel = part.rels[rid]
    try:
        target = rel.target_part
        if "ImagePart" not in type(target).__name__:
            return None
        return target.blob
    except Exception:
        return None


PLACEHOLDER_PREFIX = "[正在解析图片 "
PLACEHOLDER_SUFFIX = "...]"


def _format_image_block(rel_path: str, img_name: str, analysis: str, format: str) -> str:
    """生成图片占位符与可折叠解析结果块"""
    if format != "markdown":
        lines = [f"[图片: {rel_path}]"]
        if analysis and analysis != "pending":
            lines.append(analysis)
        else:
            lines.append(f"{PLACEHOLDER_PREFIX}{img_name}{PLACEHOLDER_SUFFIX}")
        return "\n".join(lines)

    lines = [f"![{img_name}]({rel_path})", "", "<details>", "<summary>点击展开内容</summary>", ""]
    if analysis and analysis != "pending":
        for line in analysis.split("\n"):
            lines.append("> " + line)
    else:
        lines.append(f"{PLACEHOLDER_PREFIX}{img_name}{PLACEHOLDER_SUFFIX}")
    lines.extend(["", "</details>"])
    return "\n".join(lines)


def _replace_placeholder(content: str, img_name: str, analysis: str) -> str:
    """将占位符替换为实际解析结果"""
    old = f"{PLACEHOLDER_PREFIX}{img_name}{PLACEHOLDER_SUFFIX}"
    new_lines = ["> " + line for line in analysis.split("\n") if line.strip()]
    new = "\n".join(new_lines) if new_lines else "[解析失败或为空]"
    return content.replace(old, new, 1)


def _suffix_from_part(part) -> str:
    """从 part 获取图片后缀"""
    try:
        content_type = getattr(part, "content_type", "") or ""
        if "png" in content_type:
            return ".png"
        if "jpeg" in content_type or "jpg" in content_type:
            return ".jpg"
        if "gif" in content_type:
            return ".gif"
    except Exception:
        pass
    return ".png"


async def convert_docx(
    input_path: Path,
    output_dir: Path,
    format: str,
    sse_callback: Optional[Callable[[dict], Awaitable[None]]] = None,
) -> dict:
    """转换 docx 为 Markdown 或纯文本。"""

    async def emit(data: dict):
        if sse_callback:
            await sse_callback(data)

    content_format = "markdown" if format in ("md", "markdown") else format

    try:
        await emit({"type": "debug", "content": f"正在转换: {input_path.name}"})

        if input_path.suffix.lower() == ".doc":
            return {"error": "不支持 .doc 格式，请使用 .docx"}

        await emit({"type": "debug", "content": "正在加载文档..."})
        doc = await asyncio.to_thread(Document, str(input_path))
        await emit({"type": "debug", "content": "文本解析中..."})

        output_dir.mkdir(parents=True, exist_ok=True)
        assets_dir = output_dir / "assets"
        assets_dir.mkdir(exist_ok=True)
        image_counter = 0

        all_rids = await asyncio.to_thread(_get_all_image_rids_from_document, doc)
        processed_rids = set()
        body_image_count = await asyncio.to_thread(
            lambda: sum(1 for rid in all_rids if _get_image_blob(doc.part, rid) is not None)
        )
        skip_count = len(all_rids) - body_image_count
        if not all_rids:
            await emit({"type": "debug", "content": "未发现图片，跳过图片解析"})
        else:
            msg = f"文档中共发现 {len(all_rids)} 个图片引用"
            if skip_count > 0:
                msg += f"，其中 {skip_count} 个在页眉页脚（不解析）"
            msg += f"，剩余 {body_image_count} 个将解析"
            await emit({"type": "debug", "content": msg})

        def iter_block_items(parent):
            if isinstance(parent, _Document):
                parent_elm = parent.element.body
            else:
                parent_elm = parent._element
            for child in parent_elm.iterchildren():
                if isinstance(child, CT_P):
                    yield Paragraph(child, parent)
                elif isinstance(child, CT_Tbl):
                    yield Table(child, parent)

        output_parts: List[str] = []

        header_footer_parts: List[str] = []
        if doc.sections:
            section = doc.sections[0]
            for attr in (
                "header",
                "first_page_header",
                "even_page_header",
                "footer",
                "first_page_footer",
                "even_page_footer",
            ):
                try:
                    hf = getattr(section, attr, None)
                    if hf is not None:
                        for para in hf.paragraphs:
                            t = para.text.strip()
                            if t:
                                header_footer_parts.append(t)
                except (AttributeError, KeyError):
                    pass

        images_to_parse: List[tuple] = []
        parse_enabled = is_image_parse_enabled()

        async def add_image_block(blob: bytes, rid_ref: str, source: str) -> str:
            nonlocal image_counter
            image_counter += 1
            ext = ".png"
            for r in doc.part.rels.values():
                if r.rId == rid_ref:
                    ext = _suffix_from_part(r.target_part)
                    break
            img_name = f"img_{image_counter}{ext}"
            img_path = assets_dir / img_name
            img_path.write_bytes(blob)
            if parse_enabled:
                images_to_parse.append((img_path, img_name))
            rel_path = f"assets/{img_name}"
            analysis = "pending" if parse_enabled else "[图片解析已禁用]"
            return _format_image_block(rel_path, img_name, analysis, content_format)

        blocks = list(iter_block_items(doc))
        rid_to_block_index = await asyncio.to_thread(_build_rid_to_block_index, doc, blocks)

        def rids_for_block(block_idx: int) -> List[str]:
            return [r for r in all_rids if rid_to_block_index.get(r, -1) == block_idx]

        def rids_after_block(block_idx: int) -> List[str]:
            return [r for r in all_rids if rid_to_block_index.get(r, -1) == -(block_idx + 2)]

        def rids_fallback() -> List[str]:
            return [r for r in all_rids if rid_to_block_index.get(r, -1) == -1]

        for block_idx, block in enumerate(blocks):
            for rid in rids_for_block(block_idx):
                if rid in processed_rids:
                    continue
                blob = _get_image_blob(doc.part, rid)
                if blob:
                    processed_rids.add(rid)
                    source = "段落" if hasattr(block, "text") else "表格"
                    output_parts.append(await add_image_block(blob, rid, source))
                    await emit({"type": "partial", "content": "\n\n".join(output_parts)})

            if hasattr(block, "text"):
                para = block
                text = para.text.strip()
                rids = rids_for_block(block_idx)

                if text:
                    if content_format == "markdown":
                        style_name = para.style.name if para.style else ""
                        if style_name.startswith("Heading"):
                            try:
                                match = re.search(r"\d+", style_name)
                                level = int(match.group()) if match else 1
                                level = min(max(level, 1), 6)
                                output_parts.append(f"{'#' * level} {text}")
                            except (ValueError, AttributeError):
                                output_parts.append(text)
                        else:
                            output_parts.append(text)
                    else:
                        output_parts.append(text)
                    await emit({"type": "partial", "content": "\n\n".join(output_parts)})
                elif content_format == "markdown" and not rids:
                    output_parts.append("")

                for rid in rids_after_block(block_idx):
                    if rid in processed_rids:
                        continue
                    blob = _get_image_blob(doc.part, rid)
                    if blob:
                        processed_rids.add(rid)
                        output_parts.append(await add_image_block(blob, rid, "段落后"))
                        await emit({"type": "partial", "content": "\n\n".join(output_parts)})

            else:
                table = block
                if content_format == "markdown":
                    md_table = _table_to_markdown(table, content_format)
                    if md_table:
                        output_parts.append(md_table)
                        output_parts.append("")
                else:
                    for row in table.rows:
                        cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                        cells = _dedupe_row(cells)
                        output_parts.append("\t".join(cells))

                await emit({"type": "partial", "content": "\n\n".join(output_parts)})

                for rid in rids_after_block(block_idx):
                    if rid in processed_rids:
                        continue
                    blob = _get_image_blob(doc.part, rid)
                    if blob:
                        processed_rids.add(rid)
                        output_parts.append(await add_image_block(blob, rid, "表格后"))
                        await emit({"type": "partial", "content": "\n\n".join(output_parts)})

        for rid in rids_fallback():
            if rid in processed_rids:
                continue
            blob = _get_image_blob(doc.part, rid)
            if blob:
                processed_rids.add(rid)
                output_parts.append(await add_image_block(blob, rid, "兜底"))
                await emit({"type": "partial", "content": "\n\n".join(output_parts)})

        if header_footer_parts:
            header_text = "\n".join(dict.fromkeys(header_footer_parts))
            if content_format == "markdown":
                output_parts.append("## 页眉页脚\n\n" + header_text)
            else:
                output_parts.append("=== 页眉页脚 ===\n" + header_text)
            await emit({"type": "partial", "content": "\n\n".join(output_parts)})

        parts_clean = [p for p in output_parts if p is not None]
        final_text = "\n\n".join(parts_clean).strip()
        final_text = re.sub(r"\n\s*\n\s*\n", "\n\n", final_text)
        if not final_text:
            return {"error": "未能提取到任何内容"}

        if images_to_parse:
            await emit({"type": "debug", "content": f"已检测到{len(images_to_parse)}个图片，已保存为图片资源"})
        elif not parse_enabled and image_counter > 0:
            await emit({"type": "debug", "content": "已禁用图片解析（通过环境变量），仅保留图片引用"})

        await emit({"type": "partial", "content": final_text})

        for i, (img_path, img_name) in enumerate(images_to_parse):
            await emit({"type": "debug", "content": f"解析图片（{i + 1}/{len(images_to_parse)}）：{img_name}..."})
            await asyncio.sleep(0)
            try:
                analysis = await asyncio.to_thread(analyze_image, img_path)
            except Exception as e:
                analysis = f"[解析失败: {e}]"
                await emit({"type": "debug", "content": f"  {img_name}: {e}"})
            final_text = _replace_placeholder(final_text, img_name, analysis or "[解析失败或为空]")
            await emit({"type": "partial", "content": final_text})

        out_name = input_path.stem + (".md" if format in ("md", "markdown") else ".txt")
        out_path = output_dir / out_name
        out_path.write_text(final_text, encoding="utf-8")

        await emit({"type": "debug", "content": f"提取完成，共 {len(final_text)} 字符"})
        if images_to_parse:
            await emit({"type": "debug", "content": f"所有图片解析完成 ({len(images_to_parse)} 个)"})

        return {"content": final_text, "path": str(out_path)}

    except Exception as e:
        await emit({"type": "error", "content": str(e)})
        return {"error": str(e)}

