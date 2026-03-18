# 制药文档章节提取服务

基于 Word 文档结构解析的章节提取工具，支持：

- 按章节名称模糊匹配提取内容
- `*` 返回全文
- 本地 `.docx` 路径与 `http/https` `.docx` 地址
- 中英文并行标题文档
- 内存缓存，避免重复解析同一文档
- 两种接入方式：`MCP` 和 `OpenAPI`

---

## 进程管理与停止

### 停止 OpenAPI/MCP 服务

#### 方式一：端口/进程号
1. 查找端口占用（如8001）：
  ```bash
  lsof -nP -iTCP:8001 -sTCP:LISTEN
  ```
2. 杀死进程（假设PID为12345）：
  ```bash
  kill -TERM 12345
  ```
3. 或直接用 pkill：
  ```bash
  pkill -f src.server
  ```

#### 方式二：线程管理（Python 内部）
如果服务是以线程方式启动（如 FastAPI/Uvicorn 或 MCP），可在 Python 代码中优雅停止：

```python
import threading

# 获取所有线程
for t in threading.enumerate():
   print(t.name, t.is_alive())

# 停止主线程或目标线程
# 通常需设置线程停止标志（如 stop_event），并在主循环内检测
stop_event.set()  # 需在服务代码中实现
```

> 注意：直接杀死线程不安全，推荐通过事件/标志优雅退出。

#### 方式三：Ctrl+C/KeyboardInterrupt
部分情况下可直接 Ctrl+C 停止主线程，但有时无效。

### 常见问题

- 有时 Ctrl+C 无法停止服务，需用上述命令强制终止。
- 端口未释放时，重启前需确保进程已被杀死。
- 线程方式需在代码中实现停止标志，否则只能强制终止。
---

## 匹配规则

### 1. 章节名模糊匹配

系统优先按**模糊匹配**查找章节标题。

例如：

- `溶液配制` → 可命中 `6.0 溶液配制`
- `仪器设备` → 可命中 `5.1 仪器设备`
- `流动相` → 可命中最相关的 `6.1 流动相 A` 或对应最佳章节

### 2. 双语标题支持

如果文档是中英文并行标题，例如：

- `6.0 溶液配制`
- `6.0 Preparation of Solutions`

系统会自动将其识别为**同一个章节**，避免英文章节标题把中文标题正文“截断”。

### 3. 单关键词仅返回最佳匹配

当前逻辑已调整为：

- 对每个查询关键词，只返回**最佳匹配的一个章节**
- 不再返回大量相似章节

如果传入多个关键词，例如：

`仪器设备、流动相`

则会分别为每个关键词选取 1 个最佳章节，再合并去重返回。

---

## 文档缓存

为减少重复读取 Word 的开销，系统会把解析结果缓存在内存中。

- 同一本地文件：使用标准化绝对路径作为缓存 key
- 同一个 HTTP(S) URL：使用 URL 作为缓存 key

这意味着：

- 第一次读取会解析文档
- 后续再次提取章节时，会直接复用缓存结果

> 当前缓存仅保存在内存中，服务重启后会清空。

---

## 项目结构

```text
document_generate/
├── src/
│   ├── server.py              # 服务入口，支持 MCP / OpenAPI 两种模式
│   └── utils/
│       └── document_reader.py # Word 读取、章节树解析、匹配、缓存、本地测试入口
├── requirements.txt
├── pyproject.toml
└── README.md
```

---

## 安装依赖

```bash
cd document_generate
pip install -r requirements.txt
```

如果使用虚拟环境，推荐：

```bash
.venv/bin/pip install -r requirements.txt
```

---

## 启动方式

### 1. MCP 模式

#### stdio 模式

```bash
.venv/bin/python -m src.server --mode mcp --transport stdio
```

#### streamable-http 模式

```bash
.venv/bin/python -m src.server --mode mcp --transport streamable-http
```

### 2. OpenAPI 模式

用于 Dify 这类基于 OpenAPI Schema 导入工具的平台：

```bash
.venv/bin/python -m src.server --mode openapi
```

默认地址：

- 服务地址：`http://127.0.0.1:8001`
- OpenAPI Schema：`http://127.0.0.1:8001/openapi.json`

可通过环境变量修改：

```env
OPENAPI_HOST=127.0.0.1
OPENAPI_PORT=8001
MCP_HOST=127.0.0.1
MCP_PORT=8000
HTTP_DOC_TIMEOUT=20
FUZZY_THRESHOLD=0.45
```

---

## MCP 工具说明

### `get_section_content`

参数：

| 参数 | 类型 | 说明 |
|---|---|---|
| `file_path` | `string` | Word 文档路径，支持本地路径或 HTTP(S) URL |
| `section_name` | `string` | 章节名称，支持模糊匹配；传 `*` 返回全文 |

示例：

```json
{
  "tool": "get_section_content",
  "arguments": {
    "file_path": "/Users/liudashuai/workspace/aiAlign/普洛/分析方法.docx",
    "section_name": "溶液配制"
  }
}
```

---

## OpenAPI 接口说明

OpenAPI 模式提供以下接口：

| 方法 | 路径 | 说明 |
|---|---|---|
| `GET` | `/health` | 健康检查 |
| `GET` | `/section-content` | 查询章节内容 |
| `POST` | `/section-content` | 查询章节内容 |
| `GET` | `/openapi.json` | OpenAPI Schema |

### GET /section-content

参数：

- `file_path`
- `section_name`

示例：

```bash
curl -sG 'http://127.0.0.1:8001/section-content' \
  --data-urlencode 'file_path=/Users/liudashuai/workspace/aiAlign/普洛/分析方法.docx' \
  --data-urlencode 'section_name=溶液配制'
```

### POST /section-content

请求体：

```json
{
  "file_path": "/Users/liudashuai/workspace/aiAlign/普洛/分析方法.docx",
  "section_name": "溶液配制"
}
```

返回结构：

```json
{
  "status": "success",
  "file_path": "...",
  "section_name": "溶液配制",
  "matched_sections": [
    {
      "title": "6.0 溶液配制",
      "level": 2,
      "score": 0.75,
      "content": "..."
    }
  ],
  "all_titles": ["..."],
  "total_chars": 6768,
  "message": null
}
```

---

## Dify 接入方式

如果你使用的是旧版 Dify 自定义工具导入方式，可直接导入：

- `http://127.0.0.1:8001/openapi.json`

推荐使用：

- `GET /section-content` 或 `POST /section-content`

常用参数：

- `file_path`
- `section_name`

---

## 本地测试

`document_reader.py` 自带测试入口。

### 测试章节提取

```bash
.venv/bin/python src/utils/document_reader.py /Users/liudashuai/workspace/aiAlign/普洛/分析方法.docx '溶液配制' --json
```

### 测试英文标题

```bash
.venv/bin/python src/utils/document_reader.py /Users/liudashuai/workspace/aiAlign/普洛/分析方法.docx 'Preparation of Solutions' --json
```

### 测试编号标题

```bash
.venv/bin/python src/utils/document_reader.py /Users/liudashuai/workspace/aiAlign/普洛/分析方法.docx '6.1' --json
```

### 测试多关键词

```bash
.venv/bin/python src/utils/document_reader.py /Users/liudashuai/workspace/aiAlign/普洛/分析方法.docx '仪器设备、流动相' --json
```

### 测试全文返回

```bash
.venv/bin/python src/utils/document_reader.py /Users/liudashuai/workspace/aiAlign/普洛/分析方法.docx '*' --json
```

### 测试 HTTP 文档地址

```bash
.venv/bin/python src/utils/document_reader.py 'https://example.com/demo.docx' '溶液配制' --json
```

---

## 已验证的当前行为

- `溶液配制` 能命中 `6.0 溶液配制`
- 双语标题 `6.0 溶液配制 / 6.0 Preparation of Solutions` 会合并为同一章节
- 查询结果会带出该章节下的子章节内容，例如 `6.1 / 6.2 / 6.3 ...`
- OpenAPI Schema 可访问
- `GET /section-content` 可返回 JSON 结果

---

## 注意事项

- 当前仅支持 `.docx`
- HTTP 文档必须可直接下载
- 内存缓存不会持久化到磁盘
- 如果 OpenAPI 模式不可用，请确认已安装：

```bash
.venv/bin/pip install fastapi uvicorn
```
