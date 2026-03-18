"""
结果文件访问 API
"""
from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse, PlainTextResponse
from pathlib import Path

from backend.config import OUTPUT_DIR

router = APIRouter()


def _safe_path(relative: str) -> Path:
    """安全解析相对路径，禁止穿越"""
    p = Path(relative)
    if ".." in p.parts or p.is_absolute():
        raise HTTPException(400, "非法路径")
    full = (OUTPUT_DIR / p).resolve()
    if not str(full).startswith(str(OUTPUT_DIR.resolve())):
        raise HTTPException(400, "非法路径")
    return full


@router.get("/result/{path:path}")
async def get_result(path: str):
    """读取生成的文件内容"""
    fp = _safe_path(path)
    if not fp.exists():
        raise HTTPException(404, "文件不存在")
    if not fp.is_file():
        raise HTTPException(400, "不是文件")

    suffix = fp.suffix.lower()
    if suffix in (".md", ".txt"):
        return PlainTextResponse(fp.read_text(encoding="utf-8", errors="replace"))
    return FileResponse(fp)
