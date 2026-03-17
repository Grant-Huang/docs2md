# 文档资料知识化APP

将 Word (docx) 和 Excel 文档转换为 Markdown 或纯文本的工具。支持单文件、目录批量转换，docx 中的图片经 Qwen3-VL 解析后以引用形式插入。

## 功能

- **Docx**：python-docx 提取正文与表格，图片保存到 assets/，调用 Qwen3-VL 生成说明并以 blockquote 插入
- **Excel**：MarkItDown 多工作簿转 Markdown 表格，表头为 sheet 名
- **单文件 / 目录**：支持选择单个文件或整个目录
- **输出**：用户选择输出目录，支持 .md 或 .txt

## 安装

```bash
python -m venv .venv
.venv\Scripts\activate   # Windows
# source .venv/bin/activate  # Linux/Mac

pip install -r requirements.txt
```

## 配置

复制 `.env.example` 为 `.env`，填入 AI 配置：

```bash
cp .env.example .env
# 编辑 .env，设置 DASHSCOPE_API_KEY=sk-xxx
```

| 变量 | 说明 |
|------|------|
| `DASHSCOPE_API_KEY` | 通义千问 API Key，用于 docx 图片解析（可选，未配置时跳过图片解析） |
| `DASHSCOPE_BASE_URL` | DashScope 兼容端点（可选，默认已设置） |
| `QWEN_VL_MODEL` | 图片解析模型（可选，默认 qwen3-vl-plus） |

## 运行

```bash
python run.py
# 或
uvicorn backend.main:app --host 0.0.0.0 --port 8000
```

访问 http://localhost:8000

## 项目结构

```
all2markdown/
├── backend/           # FastAPI 后端
│   ├── main.py
│   ├── config.py
│   ├── routers/
│   ├── converters/
│   ├── services/     # Qwen VL
│   └── utils/
├── frontend/          # 静态前端
├── uploads/           # 临时上传
├── output/            # 默认输出
└── requirements.txt
```

## License

MIT
