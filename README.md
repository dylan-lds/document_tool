# 制药文档内容提取服务

基于 Word（`.docx`）和 PDF（`.pdf`）的文档内容提取工具，适合大模型喂入场景。支持 MCP 和 OpenAPI 两种服务模式。

---

## 功能特性

- **Word 章节提取**：按章节名称模糊匹配提取内容，`*` 返回全文
- **Word 嵌套表格**：表格递归展开为 Markdown，嵌套子表格以 `[Nested Table L{深度}.{编号}]` 标注
- **PDF 全文提取**：按页组织正文和表格，表格结构化为 Markdown，适合大模型输入
- **双语标题合并**：中英文并行标题（如 `6.0 溶液配制 / 6.0 Preparation of Solutions`）自动合并为同一章节
- **本地 & 远程**：同时支持本地路径和 `http/https` 路径
- **内存缓存**：同一文件只解析一次，服务重启后清空
- **多平台接入**：支持 MCP（stdio / streamable-http）和 OpenAPI（兼容 Dify 等工具平台）

---

## 项目结构

```text
document_generate/
├── src/
│   ├── server.py                 # 服务入口，支持 MCP / OpenAPI 两种模式
│   └── utils/
│       ├── document_reader.py    # Word 读取、章节树解析、模糊匹配、缓存
│       └── pdf_reader.py         # PDF 读取、表格结构化、缓存
├── requirements.txt
├── pyproject.toml
└── README.md
```

---

## 安装依赖

```bash
# 推荐使用虚拟环境
python -m venv .venv
source .venv/bin/activate        # Linux / macOS
# .venv\Scripts\activate         # Windows

pip install -r requirements.txt
```

`requirements.txt` 已包含所有依赖（含 `pymupdf`），无需额外安装。

---

## 启动方式

### MCP 模式

适合通过管道调用或长连接接入的 MCP 客户端（如 Claude Desktop）。

**stdio 模式**（由 MCP 客户端管道启动）：

```bash
.venv/bin/python -m src.server --mode mcp --transport stdio
```

**streamable-http 模式**（本地常驻服务）：

```bash
.venv/bin/python -m src.server --mode mcp --transport streamable-http
```

默认监听 `http://127.0.0.1:8000`。

---

### OpenAPI 模式

适合 Dify 等通过 OpenAPI Schema 导入工具的平台。

```bash
.venv/bin/python -m src.server --mode openapi
```

默认地址：

| 端点 | 地址 |
|---|---|
| 服务 | `http://127.0.0.1:8001` |
| OpenAPI Schema | `http://127.0.0.1:8001/openapi.json` |

---

### 环境变量

| 变量 | 默认值 | 说明 |
|---|---|---|
| `MCP_HOST` | `127.0.0.1` | MCP 服务监听地址 |
| `MCP_PORT` | `8000` | MCP 服务端口 |
| `OPENAPI_HOST` | `127.0.0.1` | OpenAPI 服务监听地址 |
| `OPENAPI_PORT` | `8001` | OpenAPI 服务端口 |
| `HTTP_DOC_TIMEOUT` | `20` | 远程文档下载超时（秒） |
| `PDF_HTTP_TIMEOUT` | `30` | 远程 PDF 下载超时（秒） |
| `FUZZY_THRESHOLD` | `0.45` | 章节名模糊匹配阈值 |

可在项目根目录创建 `.env` 文件配置，格式示例：

```env
OPENAPI_HOST=0.0.0.0
OPENAPI_PORT=8001
FUZZY_THRESHOLD=0.5
```

---

## MCP 工具说明

### `get_section_content`

从 Word 文档中提取指定章节内容。

| 参数 | 类型 | 说明 |
|---|---|---|
| `file_path` | `string` | `.docx` 文件路径，支持本地路径或 HTTP(S) URL |
| `section_name` | `string` | 章节名称，支持模糊匹配；传 `*` 返回全文 |

示例：

```json
{
  "tool": "get_section_content",
  "arguments": {
    "file_path": "/path/to/分析方法.docx",
    "section_name": "溶液配制"
  }
}
```

---

### `get_pdf_content`

读取 PDF 全文内容，表格优先结构化为 Markdown，适合大模型输入。

| 参数 | 类型 | 说明 |
|---|---|---|
| `file_path` | `string` | `.pdf` 文件路径，支持本地路径或 HTTP(S) URL |

示例：

```json
{
  "tool": "get_pdf_content",
  "arguments": {
    "file_path": "/path/to/报告.pdf"
  }
}
```

---

## OpenAPI 接口说明

### 接口列表

| 方法 | 路径 | 说明 |
|---|---|---|
| `GET` | `/health` | 健康检查 |
| `GET/POST` | `/section-content` | 提取 Word 章节内容 |
| `GET/POST` | `/pdf-content` | 读取 PDF 全文及表格 |
| `GET` | `/openapi.json` | OpenAPI Schema（Dify 导入用） |

---

### GET /section-content

```bash
curl -sG 'http://127.0.0.1:8001/section-content' \
  --data-urlencode 'file_path=/path/to/分析方法.docx' \
  --data-urlencode 'section_name=溶液配制'
```

---

### POST /section-content

```bash
curl -s -X POST 'http://127.0.0.1:8001/section-content' \
  -H 'Content-Type: application/json' \
  -d '{"file_path":"/path/to/分析方法.docx","section_name":"溶液配制"}'
```

返回结构：

```json
{
  "status": "success",
  "file_path": "/path/to/分析方法.docx",
  "section_name": "溶液配制",
  "matched_sections": [
    {
      "title": "6.0 溶液配制",
      "level": 2,
      "score": 0.75,
      "content": "..."
    }
  ],
  "all_titles": ["1. 目的", "2. 范围", "..."],
  "total_chars": 6768,
  "message": null
}
```

---

### GET /pdf-content

```bash
curl -sG 'http://127.0.0.1:8001/pdf-content' \
  --data-urlencode 'file_path=/path/to/报告.pdf'
```

---

### POST /pdf-content

```bash
curl -s -X POST 'http://127.0.0.1:8001/pdf-content' \
  -H 'Content-Type: application/json' \
  -d '{"file_path":"/path/to/报告.pdf"}'
```

返回结构：

```json
{
  "status": "success",
  "file_path": "/path/to/报告.pdf",
  "content": "## 第 1 页\n\n[表格内容]\n### 表格 1\n| 项目 | 数值 |\n|---|---|\n| A | 123 |...",
  "tables": [
    "第 1 页 / 表格 1\n| 项目 | 数值 |\n|---|---|\n| A | 123 |"
  ],
  "message": null
}
```

---

## 表格输出说明

### Word 嵌套表格

单元格内的嵌套表格会递归展开，以 `[Nested Table L{深度}.{编号}]` 标注：

```
| 步骤 | 操作 | 备注 |
|---|---|---|
| 1 | 加水 | [Nested Table L2.1]<br>\| 规格 \| 数量 \|<br>\|---\|---\|<br>\| 500mL \| 1 \| |
| 2 | 溶解 | - |
```

### PDF 表格

每页的表格独立提取，输出为标准 Markdown：

```
## 第 1 页

[表格内容]
### 表格 1
| 项目 | 规格 | 备注 |
|---|---|---|
| 仪器A | XX型 | - |

[正文内容]
本品...
```

> 注意：部分扫描版 PDF 或结构复杂的表格可能解析不完整，建议人工核对。

---

## 本地测试（Word）

`document_reader.py` 提供命令行测试入口：

```bash
# 模糊匹配章节
.venv/bin/python src/utils/document_reader.py /path/to/doc.docx '溶液配制' --json

# 英文标题
.venv/bin/python src/utils/document_reader.py /path/to/doc.docx 'Preparation of Solutions' --json

# 编号标题
.venv/bin/python src/utils/document_reader.py /path/to/doc.docx '6.1' --json

# 多关键词（逗号或顿号分隔）
.venv/bin/python src/utils/document_reader.py /path/to/doc.docx '仪器设备、流动相' --json

# 返回全文
.venv/bin/python src/utils/document_reader.py /path/to/doc.docx '*' --json

# 远程文档
.venv/bin/python src/utils/document_reader.py 'https://example.com/demo.docx' '溶液配制' --json
```

---

## 停止服务

### Ctrl+C

大多数情况下直接 `Ctrl+C` 即可停止。

### 端口方式（如 Ctrl+C 无效）

```bash
# 查找端口占用（如 8001）
lsof -nP -iTCP:8001 -sTCP:LISTEN

# 按 PID 终止
kill -TERM <PID>

# 或直接按进程名终止
pkill -f "src.server"
```

---

## 注意事项

- 仅支持 `.docx` 和 `.pdf` 格式
- HTTP(S) 文档需可直接下载（不支持需登录认证的链接）
- 内存缓存不持久化，服务重启后清空
- Dify 工具导入地址：`http://<host>:8001/openapi.json`
