"""
Qwen3-VL 图片解析服务
通过 DashScope 兼容接口调用
配置从环境变量读取（不依赖包内路径或仓库结构）
"""

from __future__ import annotations

import base64
import os
import mimetypes
from pathlib import Path
from typing import Optional

from openai import OpenAI

# 图片解析提示词（用户需求）
IMAGE_PROMPT = """你是一位拥有 20 年经验的“全栈首席架构师”。
你的任务是针对用户上传的技术图片进行客观现状描述。
你必须保持中立，仅描述图中呈现的内容，严禁输出任何评价、优劣判断或改进建议。
处理流程：
分类识别： 首先识别图片属于以下哪一类：【软件界面 UI】、【系统架构图】、【逻辑流程图】或【数据表格】。
客观执行： 根据识别出的类型，自动启用下述对应的描述协议。

协议 A：[软件界面 UI]
布局拆解： 准确描述页面的区域划分（如：顶部导航、左侧菜单、中央内容区、右侧辅助栏）。
组件罗列： 识别并列出图中可见的所有交互元素（如：按钮文本、输入框标签、状态指示灯、数据图表等）及其当前呈现的状态。
流程还原： 仅根据界面按钮和箭头，描述其体现的功能路径，不做逻辑合理性分析。

协议 B：[系统架构图]
组件清单： 识别图中所有的技术组件（网关、数据库、中间件等）及其标注名称。
数据流描述： 提供一个 Markdown 表格，客观记录图中绘制的连接关系：【源组件】->【目标组件】|【图中标注的协议/接口】|【图中标注的传输内容】。
结构拓扑： 描述组件间的物理或逻辑层级关系（如：位于 DMZ 区、属于集群 A 等）。

协议 C：[逻辑流程图]
逻辑节点提取： 识别图中所有的【起始/终止符】、【处理动作】和【判断条件】。
路径追踪： 详细描述图中绘制的所有分支走向，包括所有的 Yes/No 路径。
Mermaid 源码： 必须提供完整的 Mermaid.js 源码，以 1:1 还原图中所示的逻辑结构。

协议 D：[数据/文字表格]
空间扫描： 请先扫描图片全貌，确认该表格共有 X 行 X 列（包含合并单元格）。
数据转录： 把每一行都视为一个独立的“精确还原”目标，遍历扫描获得的所有行，1:1 还原文字，包括编号（如 1、2、）、冒号、分号。忽略图片中的水印。
原样输出： 逐行完整输出原文描述中的客观事实，不做统计、分析和原因猜测，不要进行任何形式的总结、概括或解释。
分段输出： 由于表格内容极其丰富，请不要试图一次性总结，而是逐行进行精确转录。
处理合并单元格： 保持原图的行列结构。对于合并的行，请在对应的每一行 Markdown 中都重复填入该内容，确保数据结构完整。
特别注意： 严禁跳过任何一行。即使下方单元格内容重复或包含“现状/需求”等相似字样，也必须全部写出。
格式： 输出标准的 Mermaid.js 表格源码。

总体输出要求：
声明类型： 开篇请先声明识别到的图片类型：[软件界面 UI]/[系统架构图]/[逻辑流程图]/[数据表格/文档截图]/其他（简单说明）
禁止评论： 严禁使用“建议”、“推荐”、“风险”、“缺陷”、“优点”等主观词汇。
专业术语： 保持技术术语的准确性。
综合处理： 如果图片包含多种元素（如架构图中嵌套了流程），请同时执行相关协议的描述部分。"""


def _get_env(name: str, default: str = "") -> str:
    val = os.getenv(name)
    return val if val is not None and val != "" else default


def _get_mime(image_path: Path) -> str:
    mime, _ = mimetypes.guess_type(str(image_path))
    return mime or "application/octet-stream"


def is_image_parse_enabled() -> bool:
    """
    图片解析总开关。
    支持以下环境变量：
    - DOCS2MD_DISABLE_IMAGE_PARSE=1  -> 禁用
    - DOCS2MD_SKIP_IMAGE=1           -> 禁用
    - DOCS2MD_VL_ENABLED=0           -> 禁用
    """
    disable_image_parse = _get_env("DOCS2MD_DISABLE_IMAGE_PARSE", "").strip().lower()
    if disable_image_parse in {"1", "true", "yes", "on"}:
        return False

    skip_image = _get_env("DOCS2MD_SKIP_IMAGE", "").strip().lower()
    if skip_image in {"1", "true", "yes", "on"}:
        return False

    vl_enabled = _get_env("DOCS2MD_VL_ENABLED", "1").strip().lower()
    if vl_enabled in {"0", "false", "no", "off"}:
        return False

    return True


def analyze_image(image_path: Path, prompt: Optional[str] = None) -> str:
    """
    调用 Qwen3-VL 解析图片。
    必需环境变量：
    - DASHSCOPE_API_KEY
    可选环境变量：
    - DASHSCOPE_BASE_URL（默认 DashScope compatible-mode）
    - QWEN_VL_MODEL（默认 qwen3-vl-plus）
    """
    if not is_image_parse_enabled():
        return "[图片解析已禁用]"

    api_key = _get_env("DASHSCOPE_API_KEY", "")
    if not api_key:
        raise RuntimeError("缺少环境变量 DASHSCOPE_API_KEY，无法调用 VL 服务")

    base_url = _get_env(
        "DASHSCOPE_BASE_URL",
        "https://dashscope.aliyuncs.com/compatible-mode/v1",
    )
    model = _get_env("QWEN_VL_MODEL", "qwen3-vl-plus")

    client = OpenAI(api_key=api_key, base_url=base_url)
    data = base64.b64encode(image_path.read_bytes()).decode("utf-8")
    used_prompt = prompt or IMAGE_PROMPT
    mime = _get_mime(image_path)

    resp = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": used_prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{mime};base64,{data}"},
                    },
                ],
            }
        ],
    )
    try:
        return (resp.choices[0].message.content or "").strip()
    except Exception:
        return ""

