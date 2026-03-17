"""
文档资料知识化APP FastAPI 入口
"""
from fastapi import FastAPI
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


# 静态文件（mount 放最后，避免覆盖 /api）
if OUTPUT_DIR.exists():
    app.mount("/output", StaticFiles(directory=str(OUTPUT_DIR)), name="output")
if UPLOADS_DIR.exists():
    app.mount("/uploads", StaticFiles(directory=str(UPLOADS_DIR)), name="uploads")

frontend_dir = BASE_DIR / "frontend"
if frontend_dir.exists():
    app.mount("/", StaticFiles(directory=str(frontend_dir), html=True), name="frontend")
