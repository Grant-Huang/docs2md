# docs2md PyPI 发布指南

本文档描述如何将 `docs2md` 构建并发布到 PyPI（或 TestPyPI）。

## 1. 前置条件

- 安装构建与上传工具：

```bash
python -m pip install --upgrade build twine
```

- 确认版本号：
  - `pyproject.toml` 中的 `[project].version`
  - `src/docs2md/__init__.py` 中的 `__version__`（建议与 `pyproject.toml` 同步）

## 2. 构建 sdist / wheel

在项目根目录执行：

```bash
python -m build
```

产物会生成在 `dist/` 目录下。

## 3. 校验包元数据

```bash
python -m twine check dist/*
```

确保输出中没有 error。

## 4. 发布到 TestPyPI（推荐先做）

```bash
python -m twine upload --repository testpypi dist/*
```

然后在一个干净环境验证安装与 CLI：

```bash
python -m venv .venv-test
source .venv-test/bin/activate
python -m pip install -i https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple docs2md
docs2md --help
```

## 5. 发布到 PyPI

```bash
python -m twine upload dist/*
```

## 6. 安装验证（用户视角）

```bash
pip install docs2md
docs2md --help
docs2md some.docx
```

## 7. 常见问题

- **包名已被占用**：如果 `docs2md` 在 PyPI 上已存在，需要修改 `pyproject.toml` 的 `name`，并同步 README。
- **依赖过大**：`markitdown` 的依赖链较长。如果后续希望更轻量，需要在 Excel 转换上做依赖拆分或改造。

