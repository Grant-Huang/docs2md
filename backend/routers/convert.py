"""
转换 API 路由
"""
from fastapi import APIRouter, UploadFile, File, Form, HTTPException, Body
from fastapi.responses import StreamingResponse
import json
import asyncio
import time
import random
from pathlib import Path

from backend.config import UPLOADS_DIR, OUTPUT_DIR, MAX_FILE_SIZE
from backend.converters.docx_converter import convert_docx
from backend.converters.excel_converter import convert_excel

router = APIRouter()


def _validate_output_path(output_dir: str, base: Path) -> Path:
    """校验输出路径，禁止路径穿越"""
    if not output_dir or output_dir.strip() == "":
        return base
    try:
        p = Path(output_dir.strip())
        if ".." in p.parts or (p.is_absolute() and not str(p).startswith(str(base))):
            raise HTTPException(400, "非法输出路径")
        resolved = p if p.is_absolute() else (base / p)
        resolved = resolved.resolve()
        base_resolved = base.resolve()
        if not str(resolved).startswith(str(base_resolved)):
            raise HTTPException(400, "非法输出路径")
        return resolved
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(400, "非法输出路径")


@router.post("/convert")
async def convert_file(
    file: UploadFile = File(...),
    output_dir: str = Form(""),
    format: str = Form("md"),
):
    """单文件转换，返回 SSE 流"""
    if format not in ("md", "txt"):
        raise HTTPException(400, "format 必须为 md 或 txt")

    out_base = _validate_output_path(output_dir or "", OUTPUT_DIR)
    suffix = Path(file.filename or "file").suffix
    tmp_path = UPLOADS_DIR / f"{int(time.time()*1000)}_{random.randint(1000,9999)}{suffix}"

    queue: asyncio.Queue = asyncio.Queue()

    async def sse_callback(data: dict):
        await queue.put(data)

    async def run_convert():
        success = False
        try:
            if not tmp_path.exists():
                await queue.put({"type": "_done", "result": {"error": f"临时文件不存在: {tmp_path}"}})
                return

            await queue.put({"type": "debug", "content": "转换任务已启动"})

            ext = suffix.lower()
            if ext in (".docx", ".doc"):
                result = await convert_docx(tmp_path, out_base, format, sse_callback=sse_callback)
            elif ext in (".xlsx", ".xls"):
                result = await convert_excel(tmp_path, out_base, format, sse_callback=sse_callback)
            else:
                await queue.put({"type": "_done", "result": {"error": f"不支持格式: {ext}"}})
                return
            success = not result.get("error")
            await queue.put({"type": "_done", "result": result})
        except Exception as e:
            await queue.put({"type": "_done", "result": {"error": str(e)}})
        finally:
            # 仅转换成功时删除临时文件，失败时保留便于排查
            if success and tmp_path.exists():
                tmp_path.unlink(missing_ok=True)

    async def sse_stream():
        # 必须先完成 file.read()：multipart 解析会在此阻塞直到客户端上传完成。
        # 若先 yield 再 read，用户会看到「已连接」但实际卡在 read，造成误导。
        try:
            content = await asyncio.wait_for(file.read(), timeout=600)
        except asyncio.TimeoutError:
            yield f"data: {json.dumps({'type': 'error', 'content': '文件上传超时（10分钟），请检查网络或减小文件'}, ensure_ascii=False)}\n\n"
            return

        if len(content) > MAX_FILE_SIZE:
            yield f"data: {json.dumps({'type': 'error', 'content': f'文件超过 {MAX_FILE_SIZE // (1024*1024)}MB 限制'}, ensure_ascii=False)}\n\n"
            return

        if len(content) == 0:
            yield f"data: {json.dumps({'type': 'error', 'content': '文件为空，请检查上传'}, ensure_ascii=False)}\n\n"
            return

        try:
            UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
            tmp_path.write_bytes(content)
        except OSError as e:
            yield f"data: {json.dumps({'type': 'error', 'content': f'保存文件失败: {e}'}, ensure_ascii=False)}\n\n"
            return

        yield f"data: {json.dumps({'type': 'debug', 'content': f'文件已保存至 {tmp_path.name}，开始转换...'}, ensure_ascii=False)}\n\n"

        task = asyncio.create_task(run_convert())
        try:
            while True:
                try:
                    msg = await asyncio.wait_for(queue.get(), timeout=300)
                except asyncio.TimeoutError:
                    yield f"data: {json.dumps({'type': 'error', 'content': '转换超时'}, ensure_ascii=False)}\n\n"
                    break
                if msg.get("type") == "_done":
                    r = msg.get("result", {})
                    if r.get("error"):
                        yield f"data: {json.dumps({'type': 'error', 'content': r['error']}, ensure_ascii=False)}\n\n"
                    else:
                        yield f"data: {json.dumps({'type': 'complete', 'content': r.get('content', ''), 'path': r.get('path', '')}, ensure_ascii=False)}\n\n"
                    break
                yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
        finally:
            task.cancel()
            try:
                await task
            except asyncio.CancelledError:
                pass

    return StreamingResponse(
        sse_stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "Connection": "keep-alive"},
    )


@router.post("/convert-dir")
async def convert_directory(
    input_dir: str = Body(..., embed=True),
    output_dir: str = Body(..., embed=True),
    format: str = Body("md", embed=True),
):
    """目录批量转换（服务端路径）"""
    if format not in ("md", "txt"):
        raise HTTPException(400, "format 必须为 md 或 txt")

    inp = Path(input_dir)
    out = Path(output_dir)
    if not inp.exists() or not inp.is_dir():
        raise HTTPException(400, "输入目录不存在")
    out.mkdir(parents=True, exist_ok=True)

    from backend.utils.traversal import traverse_and_convert

    results = await traverse_and_convert(inp, out, format)
    return {"status": "ok", "results": results}


@router.post("/convert-dir-upload")
async def convert_directory_upload(
    files: list[UploadFile] = File(...),
    output_dir: str = Form(""),
    format: str = Form("md"),
):
    """目录批量转换（上传多个文件，保持目录结构）"""
    if format not in ("md", "txt"):
        raise HTTPException(400, "format 必须为 md 或 txt")

    out_base = _validate_output_path(output_dir or "", OUTPUT_DIR)
    tmp_root = UPLOADS_DIR / f"dir_{int(time.time()*1000)}_{random.randint(1000,9999)}"
    tmp_root.mkdir(parents=True, exist_ok=True)

    try:
        supported = {".docx", ".doc", ".xlsx", ".xls"}
        for f in files:
            name = f.filename or ""
            if not name or Path(name).suffix.lower() not in supported:
                continue
            rel = name
            target = tmp_root / rel
            target.parent.mkdir(parents=True, exist_ok=True)
            content = await f.read()
            if len(content) <= MAX_FILE_SIZE:
                target.write_bytes(content)

        from backend.utils.traversal import traverse_and_convert

        results = await traverse_and_convert(tmp_root, out_base, format)
        return {"status": "ok", "results": results}
    finally:
        import shutil
        if tmp_root.exists():
            shutil.rmtree(tmp_root, ignore_errors=True)
