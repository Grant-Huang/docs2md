#!/usr/bin/env python3
"""启动 文档资料知识化APP 服务"""
import os
import uvicorn

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8001"))
    uvicorn.run(
        "backend.main:app",
        host="0.0.0.0",
        port=port,
        reload=True,
    )
