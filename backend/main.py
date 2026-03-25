"""
文档资料知识化APP FastAPI 入口
"""
from fastapi import FastAPI
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path

from backend.config import UPLOADS_DIR, OUTPUT_DIR, BASE_DIR
from backend.routers import convert, files

app = FastAPI(title="文档资料知识化APP", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(convert.router, prefix="/api", tags=["convert"])
app.include_router(files.router, prefix="/api", tags=["files"])


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.get("/api/skills-guide", response_class=Response)
def skills_guide():
    """返回 AI Skills Integration Guide（Markdown）。
    供 AI 系统、Agent 框架或技能平台读取，了解如何集成本服务的转换能力。
    """
    guide_path = BASE_DIR / "docs" / "ai_skills_guide.md"
    if not guide_path.exists():
        return Response("Guide not found", status_code=404, media_type="text/plain")
    return Response(
        guide_path.read_text(encoding="utf-8"),
        media_type="text/markdown; charset=utf-8",
    )


# 静态文件（mount 放最后，避免覆盖 /api）
if OUTPUT_DIR.exists():
    app.mount("/output", StaticFiles(directory=str(OUTPUT_DIR)), name="output")
if UPLOADS_DIR.exists():
    app.mount("/uploads", StaticFiles(directory=str(UPLOADS_DIR)), name="uploads")

frontend_dir = BASE_DIR / "frontend"
if frontend_dir.exists():
    app.mount("/", StaticFiles(directory=str(frontend_dir), html=True), name="frontend")
