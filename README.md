# docs2md — 文档知识化工具

将 Word（`.docx`）和 Excel（`.xlsx/.xls`）文档批量转换为 Markdown 或纯文本。
DOCX 中的图片由 Qwen3-VL 自动解析描述，以可折叠引用形式内嵌输出。

---

## 功能

| 格式 | 输出 | 说明 |
|------|------|------|
| `.docx` | `.md` / `.txt` | 正文、表格、标题样式；图片提取到 `assets/` 并由 Qwen3-VL 解析 |
| `.doc` | `.md` / `.txt` | 自动升级为 `.docx` 后转换（LibreOffice 或 win32com） |
| `.xlsx` / `.xls` | `.md` / `.txt` | 多工作表合并，每个 sheet 以 `## SheetName` 作标题 |

- **单文件转换**：上传单个文件，SSE 实时流式返回进度和结果
- **目录批量转换**：服务端路径 或 多文件上传（保持目录结构），完成后生成 `index.md`
- **图片智能解析**：三阶段渐进渲染——先出文字，再出图片链接占位，最后逐图替换 AI 解析结果

---

## 快速开始

```bash
# 1. 安装（CLI）
pip install docs2md

# 2. 系统依赖：LibreOffice（用于 .doc/.xls 旧版格式升级）
#    Ubuntu/Debian:
sudo apt install libreoffice
#    macOS:
brew install libreoffice
#    Windows: 安装 LibreOffice（https://www.libreoffice.org/）
#             或安装 Microsoft Office（退化方案，需 pywin32）

# 3. 配置（可选：启用图片解析）
# 需要在环境变量中提供 DASHSCOPE_API_KEY（以及可选的 DASHSCOPE_BASE_URL、QWEN_VL_MODEL）
# 例如（Linux/macOS）：
export DASHSCOPE_API_KEY="your_key"
```

> 说明：本仓库仍保留 FastAPI + 前端作为示例（见下文“服务端示例”），但 `pip install docs2md` 默认提供的是 CLI 能力。

---

## CLI 命令行

```bash
# 单文件转换（输出到同目录，默认 Markdown）
docs2md report.docx
docs2md data.xls -o results/ --format txt

# 目录批量转换（自动升级 .doc/.xls，生成 index.md）
docs2md docs/
docs2md docs/ -o output/ --format txt --quiet
```

| 参数 | 说明 |
|------|------|
| `input` | 输入文件（.docx/.doc/.xlsx/.xls）或目录 |
| `-o / --output` | 输出目录（单文件默认同目录；目录默认 `<input>_output/`） |
| `-f / --format` | `md`（Markdown，默认）或 `txt`（纯文本） |
| `-q / --quiet` | 只输出错误，不打印进度 |

---

## 服务端示例（FastAPI + 前端）

前端作为使用示例继续保留在仓库中（`frontend/`），可用来演示 SSE 进度与在线转换。

```bash
python -m venv .venv
source .venv/bin/activate      # Linux/Mac
# .venv\\Scripts\\activate      # Windows

pip install -e ".[server,legacy-win]"

cp .env.example .env
# 编辑 .env，填入 DASHSCOPE_API_KEY（可选，不填则跳过图片解析）

python run.py
```

访问 `http://localhost:8000`。

---

## MCP 集成（Claude 桌面版 / 任意 MCP 客户端）

通过 MCP 协议，将转换工具直接接入 Claude 桌面版，在对话中即可驱动文档转换。

---

## 发布到 PyPI

见 `docs/pypi_release.md`。

### 配置 Claude 桌面版

编辑 claude_desktop_config.json（重启 Claude 后生效）：

- **macOS**：`~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**：`%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "docs2md": {
      "command": "python",
      "args": ["/absolute/path/to/docs2md/mcp_server.py"]
    }
  }
}
```

如果使用虚拟环境，将 `"command"` 指向 venv 中的 python：

```json
{
  "mcpServers": {
    "docs2md": {
      "command": "/absolute/path/to/docs2md/.venv/bin/python",
      "args": ["/absolute/path/to/docs2md/mcp_server.py"]
    }
  }
}
```

### 可用工具

| 工具 | 参数 | 说明 |
|------|------|------|
| `convert_file` | `input_path`, `output_dir?`, `format?` | 转换单个文件 |
| `convert_directory` | `input_dir`, `output_dir?`, `format?` | 目录批量转换（含旧版格式升级） |

### 对话示例

> "把 /Users/me/documents/report.docx 转为 Markdown"
> "将 /Users/me/docs 整个文件夹批量转换，输出到 /Users/me/output"
> "转换 /data/table.xlsx，使用 txt 格式"

### AI Skills 集成文档

完整的 AI 集成指南（含 MCP 工具定义、OpenAI Function Calling JSON Schema、REST API 示例、LangChain 接入代码）可通过以下方式访问：

- **Web URL（服务运行时）：** `http://<host>:8000/api/skills-guide`
- **源文件：** `docs/ai_skills_guide.md`

将此 URL 发给任何 AI 平台或 Agent 框架，它们即可解析并创建对应的 skill / tool / action。

### 手动测试 MCP 服务器

```bash
# 或直接用 mcp dev 调试
mcp dev mcp_server.py
```

---

## 配置

| 变量 | 必填 | 默认值 | 说明 |
|------|------|--------|------|
| `DASHSCOPE_API_KEY` | 否 | — | 通义千问 API Key，缺失时跳过图片解析 |
| `DASHSCOPE_BASE_URL` | 否 | DashScope 兼容端点 | 可替换为其他 OpenAI 兼容接口 |
| `QWEN_VL_MODEL` | 否 | `qwen3-vl-plus` | VL 模型名称 |
| `UPLOADS_DIR` | 否 | `uploads/` | 临时文件目录 |
| `OUTPUT_DIR` | 否 | `output/` | 默认输出目录 |
| `MAX_FILE_SIZE` | 否 | `52428800`（50 MB） | 单文件上限（字节） |
| `MAX_FILES_PER_BATCH` | 否 | `200` | 目录批量转换文件数上限 |

### 如何禁用图片解析

支持以下环境变量开关，任意一个命中即禁用图片解析（VL）：

- `DOCS2MD_DISABLE_IMAGE_PARSE=1`
- `DOCS2MD_SKIP_IMAGE=1`
- `DOCS2MD_VL_ENABLED=0`

禁用后会保留图片文件和 Markdown 图片引用，但不再调用 VL 生成图片描述。

示例（Linux/macOS）：

```bash
export DOCS2MD_DISABLE_IMAGE_PARSE=1
docs2md report.docx
```

---

## 项目结构

```
docs2md/
├── backend/
│   ├── main.py                   # FastAPI 入口，挂载路由与静态资源
│   ├── config.py                 # 读取 .env 配置
│   ├── converters/
│   │   ├── doc2docx_converter.py # .doc/.xls → .docx/.xlsx（LibreOffice / win32com）
│   │   ├── docx_converter.py     # DOCX → Markdown/Text（异步，图片提取+VL解析）
│   │   └── excel_converter.py    # Excel → Markdown/Text（异步，MarkItDown）
│   ├── routers/
│   │   ├── convert.py            # /api/convert、/api/convert-dir、/api/convert-dir-upload
│   │   └── files.py              # /api/result/{path} 读取输出文件
│   ├── services/
│   │   └── qwen_vl.py            # 调用 Qwen3-VL 解析图片
│   └── utils/
│       └── traversal.py          # 目录遍历、输出路径映射、index.md 生成
├── frontend/                     # 静态前端（HTML/CSS/JS）
├── uploads/                      # 临时上传目录（运行时生成）
├── output/                       # 默认输出目录（运行时生成）
├── run.py                        # Web 服务启动脚本
├── all2md.py                     # CLI 入口
├── mcp_server.py                 # MCP 服务器（Claude 桌面版集成）
├── docs/
│   └── ai_skills_guide.md        # AI 集成指南（MCP / REST API / Function Calling）
├── requirements.txt
└── .env.example
```

---

## 架构设计

### 整体数据流

```
浏览器
  │  POST /api/convert（multipart/form-data）
  ▼
routers/convert.py
  │  写入 uploads/ 临时文件
  │  创建 asyncio.Queue 作为 SSE 消息总线
  ▼
converters/docx_converter.py  或  converters/excel_converter.py
  │  通过 sse_callback 向 Queue 推送进度消息
  │  调用 services/qwen_vl.py（同步，asyncio.to_thread 包装）
  ▼
StreamingResponse（text/event-stream）
  │  逐条 yield "data: {...}\n\n"
  ▼
浏览器实时渲染
```

**目录批量转换额外阶段（`/api/convert-dir` 和 `/api/convert-dir-upload`）：**

```
utils/traversal.py::traverse_and_convert()
  │
  ├─ 阶段1  converters/doc2docx_converter.py::convert_legacy_dir()
  │           递归扫描，将 .doc → .docx，.xls → .xlsx
  │           优先用 LibreOffice headless（subprocess）
  │           Windows 回退用 win32com（需安装 Microsoft Office）
  │           已存在同名新版文件则跳过
  │
  ├─ 阶段2  collect_files()
  │           收集 .docx/.xlsx（已升级的旧版文件不再重复处理）
  │
  └─ 阶段3  逐文件调用 docx_converter / excel_converter
```

### SSE 消息协议

每条消息均为 JSON 对象：

| `type` | 含义 |
|--------|------|
| `debug` | 进度日志（"正在加载文档..."、"解析图片 2/5..."） |
| `partial` | 当前已生成的完整内容快照（每次覆盖前一次） |
| `complete` | 转换完成，包含最终内容与输出文件路径 |
| `error` | 转换失败，包含错误信息 |

### DOCX 转换三阶段

```
阶段 1：文本提取
  ├─ 按文档顺序遍历段落 / 表格
  ├─ Heading 样式 → Markdown # 标题
  ├─ 表格 → Markdown 管道表（自动去重合并单元格）
  ├─ 图片 → 保存到 assets/imgN.png，生成「正在解析...」占位符
  └─ 页眉/页脚文字收集到文末独立小节

阶段 2：立即渲染
  └─ 将含占位符的完整文本通过 SSE 推送，用户即时可见内容与图片链接

阶段 3：逐图解析替换
  ├─ 逐个调用 Qwen3-VL 解析图片（asyncio.to_thread）
  ├─ 将占位符替换为实际解析结果（blockquote 格式）
  └─ 每替换一张即推送更新后的完整文本
```

### 图片定位算法

DOCX 图片可能出现在段落内（DrawingML `r:embed`）或 VML 内联（`r:id`），也可能位于 Content Control 等非块元素中。定位策略：

1. **精确定位**：遍历 XML body，将每个 `rId` 映射到最近的块级祖先索引
2. **邻近插入**：若图片在非块元素中，找到前一个块，插入到该块内容之后
3. **兜底追加**：无法定位的图片按 `all_rids` 顺序追加到正文末尾

### Qwen3-VL 图片解析

调用 DashScope OpenAI 兼容接口，图片以 Base64 data URL 传输，超时 120 秒。

内置四协议提示词，模型自动根据图片类型选择：

| 协议 | 图片类型 | 输出重点 |
|------|----------|----------|
| A | 软件界面 UI | 区域划分、交互元素、功能路径 |
| B | 系统架构图 | 组件清单、数据流 Markdown 表、拓扑结构 |
| C | 逻辑流程图 | 节点提取、分支路径、Mermaid 源码 |
| D | 数据/文字表格 | 逐行精确转录，处理合并单元格 |

### Excel 转换

使用 MarkItDown 库一次性完成转换，无需手动处理每个 sheet。
`txt` 格式时，通过 `_md_to_plain()` 将 Markdown 管道表转为 Tab 分隔纯文本。

### 目录批量转换

`utils/traversal.py` 负责：

1. `collect_files()`：`rglob` 递归收集 `.docx/.doc/.xlsx/.xls`，最多 `MAX_FILES_PER_BATCH` 个
2. `get_output_path()`：保持输入目录的相对层级结构映射到输出目录
3. 顺序调用各 converter（无并发，避免 VL API 限流）
4. `generate_index_md()`：生成 `index.md`，列出所有转换结果的相对链接

### 安全设计

- **路径穿越防护**：`_validate_output_path()` 检查 `..` 并验证 resolve 后的路径必须在 `OUTPUT_DIR` 内
- **文件大小限制**：上传时检查 `MAX_FILE_SIZE`（默认 50 MB）
- **临时文件清理**：转换成功后删除 `uploads/` 临时文件；失败时保留便于排查
- **目录批量上传**：处理完成后 `shutil.rmtree` 清理整个临时目录

---

## API 端点

### `POST /api/convert`
单文件转换，返回 SSE 流。

**表单参数：**
- `file`：上传文件（`.docx/.xlsx/.xls`）
- `output_dir`：输出目录（相对于 `OUTPUT_DIR`，可为空）
- `format`：`md`（默认）或 `txt`

**响应：** `text/event-stream`，每行 `data: {type, content[, path]}\n\n`

---

### `POST /api/convert-dir`
服务端目录批量转换（服务器本地路径）。

**JSON Body：**
```json
{ "input_dir": "/path/to/docs", "output_dir": "/path/to/out", "format": "md" }
```

**响应：** `{ "status": "ok", "results": [...] }`

---

### `POST /api/convert-dir-upload`
目录批量转换（上传多个文件，保持目录结构）。

**表单参数：**
- `files`：多文件上传，`filename` 含相对路径
- `output_dir`、`format`：同上

---

### `GET /api/result/{path}`
读取输出文件内容，支持路径穿越防护。

---

## License

MIT
