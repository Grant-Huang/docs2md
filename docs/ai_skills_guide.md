# docs2md — AI Skills Integration Guide

**Version:** 2.0
**Base URL:** `http://<host>:8000`
**OpenAPI Spec:** `GET /openapi.json`

This guide enables any AI system, agent, or skills platform to integrate docs2md's document conversion capabilities. Three integration paths are provided:

| Path | Best For |
|------|----------|
| [MCP Tools](#mcp-tools-claude--mcp-clients) | Claude Desktop, Claude Code, any MCP client |
| [REST API / Function Calling](#rest-api--function-calling) | OpenAI GPT Actions, LangChain, custom agents |
| [CLI](#cli-local-use) | Local scripts, shell agents, subprocess calls |

---

## Capabilities Overview

docs2md converts Word and Excel documents to Markdown or plain text.

| Input | Output | Notes |
|-------|--------|-------|
| `.docx` | `.md` / `.txt` | Preserves headings, tables, image descriptions |
| `.doc` | `.md` / `.txt` | Auto-upgraded to .docx via LibreOffice first |
| `.xlsx` / `.xls` | `.md` / `.txt` | Multi-sheet, each sheet as `## SheetName` |

**Key behaviors:**
- Images are extracted to `assets/`, analyzed by Qwen3-VL, inserted as collapsible Markdown
- Directory conversion upgrades all `.doc`/`.xls` before processing, then deletes the originals
- All conversion endpoints stream progress via Server-Sent Events (SSE)

---

## MCP Tools (Claude / MCP Clients)

### Server Startup

```bash
python mcp_server.py   # stdio mode — standard MCP transport
```

### Claude Desktop Configuration

Edit `claude_desktop_config.json` and restart Claude:

**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

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

With a virtual environment:

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

### MCP Tool Schemas

#### `convert_file`

```json
{
  "name": "convert_file",
  "description": "Convert a single Word or Excel document to Markdown or plain text. Supported formats: .docx, .doc, .xlsx, .xls. Returns the converted content and the output file path.",
  "inputSchema": {
    "type": "object",
    "properties": {
      "input_path": {
        "type": "string",
        "description": "Absolute path to the input file (.docx / .doc / .xlsx / .xls)."
      },
      "output_dir": {
        "type": "string",
        "description": "Absolute path to the output directory. Defaults to the same directory as the input file."
      },
      "format": {
        "type": "string",
        "enum": ["md", "txt"],
        "default": "md",
        "description": "Output format: 'md' for Markdown (default) or 'txt' for plain text."
      }
    },
    "required": ["input_path"]
  }
}
```

#### `convert_directory`

```json
{
  "name": "convert_directory",
  "description": "Batch-convert all Word and Excel files in a directory. Automatically upgrades .doc/.xls to .docx/.xlsx before converting and deletes the originals. Generates an index.md listing all converted files.",
  "inputSchema": {
    "type": "object",
    "properties": {
      "input_dir": {
        "type": "string",
        "description": "Absolute path to the input directory."
      },
      "output_dir": {
        "type": "string",
        "description": "Absolute path to the output directory. Defaults to <input_dir>_output/."
      },
      "format": {
        "type": "string",
        "enum": ["md", "txt"],
        "default": "md",
        "description": "Output format: 'md' for Markdown (default) or 'txt' for plain text."
      }
    },
    "required": ["input_dir"]
  }
}
```

### Example MCP Prompts

```
Convert /Users/alice/report.docx to Markdown.
Batch-convert the folder /Users/alice/docs and output to /Users/alice/output.
Convert /data/budget.xlsx using txt format.
```

---

## REST API / Function Calling

The service exposes three HTTP endpoints. Use these to build OpenAI Custom GPT Actions, LangChain tools, or any HTTP-based skill.

### Function Definitions (OpenAI Format)

```json
[
  {
    "type": "function",
    "function": {
      "name": "docs2md_convert_file",
      "description": "Upload a local document file to the docs2md service and convert it to Markdown or plain text. The file is sent as multipart/form-data. The response is a Server-Sent Event stream; the final 'complete' event contains the converted content.",
      "parameters": {
        "type": "object",
        "properties": {
          "file_path": {
            "type": "string",
            "description": "Local file path to read and upload (.docx / .doc / .xlsx / .xls)."
          },
          "format": {
            "type": "string",
            "enum": ["md", "txt"],
            "default": "md",
            "description": "Output format."
          },
          "output_dir": {
            "type": "string",
            "description": "Server-side output subdirectory (relative to OUTPUT_DIR). Optional."
          }
        },
        "required": ["file_path"]
      }
    }
  },
  {
    "type": "function",
    "function": {
      "name": "docs2md_convert_directory",
      "description": "Upload multiple files from a local directory to the docs2md service and batch-convert them. Returns an SSE stream with progress; the final 'complete' event contains a results array.",
      "parameters": {
        "type": "object",
        "properties": {
          "file_paths": {
            "type": "array",
            "items": { "type": "string" },
            "description": "List of local file paths to upload."
          },
          "format": {
            "type": "string",
            "enum": ["md", "txt"],
            "default": "md"
          }
        },
        "required": ["file_paths"]
      }
    }
  }
]
```

### Endpoint Reference

---

#### `POST /api/convert` — Single file conversion

**Request:** `multipart/form-data`

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `file` | binary | ✓ | File to convert (.docx / .doc / .xlsx / .xls, max 50 MB) |
| `format` | string | | `md` (default) or `txt` |
| `output_dir` | string | | Output subdirectory relative to server OUTPUT_DIR |

**Response:** `text/event-stream`

Each line: `data: <JSON>\n\n`

| `type` | Payload | Description |
|--------|---------|-------------|
| `debug` | `{ "content": "…" }` | Progress log message |
| `partial` | `{ "content": "…" }` | Current converted text snapshot (updates on each image parsed) |
| `complete` | `{ "content": "…", "path": "…" }` | Final content + server output file path |
| `error` | `{ "content": "…" }` | Conversion error |

**curl example:**

```bash
curl -N -X POST http://localhost:8000/api/convert \
  -F "file=@/path/to/report.docx" \
  -F "format=md"
```

**Python example (collect final result):**

```python
import httpx, json

def convert_file(file_path: str, format: str = "md") -> str:
    with open(file_path, "rb") as f:
        with httpx.stream(
            "POST", "http://localhost:8000/api/convert",
            data={"format": format},
            files={"file": (file_path, f)},
            timeout=600,
        ) as resp:
            for line in resp.iter_lines():
                if line.startswith("data: "):
                    msg = json.loads(line[6:])
                    if msg["type"] == "complete":
                        return msg["content"]
                    if msg["type"] == "error":
                        raise RuntimeError(msg["content"])
    return ""
```

---

#### `POST /api/convert-dir-upload` — Directory batch conversion

**Request:** `multipart/form-data`

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `files` | binary[] | ✓ | Files to convert; set `filename` to the relative path within the source folder to preserve structure |
| `format` | string | | `md` (default) or `txt` |
| `output_dir` | string | | Output subdirectory |

**Response:** `text/event-stream`

Same `debug` / `error` event format as single-file, plus a final:

```json
{ "type": "complete", "results": [
  { "path": "subdir/report.docx", "output": "/abs/path/out.md", "content": "…first 500 chars…" },
  { "path": "index", "output": "/abs/path/index.md", "content": "# 转换结果索引…" }
]}
```

**Python example:**

```python
import httpx, json
from pathlib import Path

def convert_directory(file_paths: list[str], format: str = "md") -> list[dict]:
    files = [
        ("files", (Path(p).name, open(p, "rb")))
        for p in file_paths
    ]
    with httpx.stream(
        "POST", "http://localhost:8000/api/convert-dir-upload",
        data={"format": format},
        files=files,
        timeout=600,
    ) as resp:
        for line in resp.iter_lines():
            if line.startswith("data: "):
                msg = json.loads(line[6:])
                if msg["type"] == "complete":
                    return msg["results"]
                if msg["type"] == "error":
                    raise RuntimeError(msg["content"])
    return []
```

---

#### `POST /api/convert-dir` — Server-side directory conversion

Converts a directory already present on the server (no upload needed).

**Request:** `application/json`

```json
{
  "input_dir": "/absolute/server/path/to/docs",
  "output_dir": "/absolute/server/path/to/output",
  "format": "md"
}
```

**Response:** `application/json`

```json
{
  "status": "ok",
  "results": [
    { "path": "report.docx", "output": "/path/output/report.md", "content": "…" }
  ]
}
```

---

### OpenAPI Specification

The full machine-readable OpenAPI 3.1 spec is available at runtime:

```
GET http://<host>:8000/openapi.json
```

Use this URL directly when creating an **OpenAI Custom GPT Action** — paste it into the "Import from URL" field in the GPT Action editor.

---

## CLI (Local Use)

For shell-based agents or scripts running on the same machine:

```bash
# Single file → Markdown
python /path/to/docs2md/all2md.py /path/to/report.docx

# Single file → plain text, custom output dir
python /path/to/docs2md/all2md.py /path/to/data.xls -o /path/to/out/ --format txt

# Directory batch conversion (auto-upgrades .doc/.xls, generates index.md)
python /path/to/docs2md/all2md.py /path/to/docs/

# With detailed progress
python /path/to/docs2md/all2md.py /path/to/docs/ -o /path/to/out/ --verbose
```

Exit code: `0` = success, `>0` = number of failed files.

| Argument | Description |
|----------|-------------|
| `input` | File (.docx/.doc/.xlsx/.xls) or directory |
| `-o / --output` | Output directory |
| `-f / --format` | `md` (default) or `txt` |
| `-v / --verbose` | Print detailed progress |

---

## Creating New Skills — Step-by-Step

### For Claude (MCP Skill)

1. Install: `pip install -r requirements.txt`
2. Copy the MCP config snippet above into `claude_desktop_config.json`
3. Restart Claude — the tools `convert_file` and `convert_directory` appear automatically

### For OpenAI Custom GPT

1. Start the docs2md web server: `python run.py`
2. In the GPT Action editor → **Import from URL** → enter `http://<host>:8000/openapi.json`
3. Adjust authentication and server URL as needed
4. The actions `/api/convert` and `/api/convert-dir-upload` will be available

### For LangChain / LlamaIndex

```python
from langchain.tools import StructuredTool
import httpx, json

def _convert(file_path: str, format: str = "md") -> str:
    with open(file_path, "rb") as f:
        with httpx.stream(
            "POST", "http://localhost:8000/api/convert",
            data={"format": format},
            files={"file": (file_path, f)},
            timeout=600,
        ) as resp:
            for line in resp.iter_lines():
                if line.startswith("data: "):
                    msg = json.loads(line[6:])
                    if msg["type"] == "complete":
                        return msg["content"]
    return ""

convert_tool = StructuredTool.from_function(
    func=_convert,
    name="convert_document",
    description="Convert a Word or Excel file to Markdown. Input: file_path (str), format ('md' or 'txt').",
)
```

### For Any OpenAI-Compatible Function Calling

Register the function definitions from the [Function Definitions](#function-definitions-openai-format) section above. Implement the function executor to call the REST endpoints and return the `content` field from the final `complete` SSE event.

---

## Notes

- **Image analysis** requires `DASHSCOPE_API_KEY` to be set in `.env`. Without it, images are extracted and linked but not analyzed.
- **Legacy format upgrade** (`.doc` / `.xls`) requires LibreOffice on Linux/macOS, or Microsoft Office + `pywin32` on Windows.
- **File size limit:** 50 MB per file.
- **Batch limit:** 200 files per directory conversion.
- **Output files** are written to the server's `output/` directory and accessible at `GET /output/<path>`.
