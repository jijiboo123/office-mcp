# office-mcp

一个 MCP 服务器：让 AI agent 能读 Word、Excel、PPT、PDF。

模型能读代码、能读网页，但**读不了你的 `.docx`**。这个项目用 MCP 协议把办公文档接给 agent——接上之后，一句"看看这份报告讲了什么"就能直接出结论。

## 为什么是 MCP

| 方案 | 问题 |
|---|---|
| 写个脚本让 agent 调 | 每个 agent 都要单独适配 |
| 把文档转成 txt 再喂 | 丢结构（表格、标题层级、备注全没了） |
| **MCP 服务器** | 一次实现，任何支持 MCP 的客户端都能用 |

MCP（Model Context Protocol）是标准协议，`stdio` 上跑 JSON-RPC 2.0。接一次，Claude Desktop / Cline / VS Code Copilot / DeepSeek Harness 都能用。

## 架构

```
┌─────────────┐   JSON-RPC over stdio   ┌──────────────┐
│  AI agent   │ ◀─────────────────────▶ │  office-mcp  │
│ (MCP 客户端) │   initialize/tools/call │   server.py  │
└─────────────┘                         └──────┬───────┘
                                               │
                                    ┌──────────┴──────────┐
                                    ▼                     ▼
                              python-docx           openpyxl
                              python-pptx           pymupdf
```

## 四个工具

| 工具 | 输入 | 输出 |
|---|---|---|
| `doc_read` | `.docx` | Markdown：标题、段落、列表、表格（**按原文顺序**） |
| `sheet_read` | `.xlsx` | 列出所有工作表，指定表转 Markdown 表格（取公式**计算结果**） |
| `slide_read` | `.pptx` | 按页输出标题、正文要点、表格、演讲者备注 |
| `pdf_extract` | `.pdf` | 按页范围提文本；扫描件会明确说"没有文字层" |

所有输出都有截断上限，防止 300 页的文档把上下文撑爆。

## 快速开始

### 1. 装依赖

```bash
pip install -r requirements.txt
```

四个库各管一种格式，只读 Word 就只装 `python-docx`。

### 2. 自检

```bash
python server.py --selftest "D:\某处\报告.docx"
```

会列出工具清单，并真的读一次那份文档。

### 3. 接进 MCP 客户端

配置形如（各客户端字段名略有差异）：

```json
{
  "mcpServers": {
    "office": {
      "command": "python",
      "args": ["D:\\path\\to\\office-mcp\\server.py"]
    }
  }
}
```

> 注意：`command` 要用**装了上述依赖的那个 Python**。如果系统 Python 没装库，就填虚拟环境或其它解释器的绝对路径。

### 4. 手动验证协议

```bash
python examples/test_client.py
```

它会把完整对话跑一遍：`initialize` → `tools/list` → 三次 `tools/call`（真实文档）→ 两个错误场景。

实测输出：

```
[id=1] initialize -> {'name': 'office-mcp', 'version': '1.0.0'} 协议 2024-11-05
[id=2] tools/list -> 4 个工具: ['doc_read', 'sheet_read', 'slide_read', 'pdf_extract']

[id=3] ✅ 返回 631 字符:
     # 家庭用电分析报告.docx
     （家庭用电分析报告.docx：53 段、3 表）
     | 报告名称 | 家庭用电情况分析报告 |
```

## 协议实现说明

不用官方 SDK，手写不到 200 行，好处是**协议长什么样一目了然**：

```python
# 一条请求进来
{"jsonrpc": "2.0", "id": 3, "method": "tools/call",
 "params": {"name": "doc_read", "arguments": {"path": "报告.docx"}}}

# 一条响应出去
{"jsonrpc": "2.0", "id": 3,
 "result": {"content": [{"type": "text", "text": "# 报告.docx\n…"}], "isError": false}}
```

三个容易踩的点：

1. **stdout 只能出 JSON**。任何 `print` 调试都会污染协议流，日志必须走 `stderr`。
2. **通知没有 `id`，不能回复**。`notifications/initialized` 回了就会让客户端困惑。
3. **工具异常不能弄死服务器**。工具里任何异常都转成 `isError: true` 的响应，循环继续。

## 接入 DeepSeek Harness（实测记录）

我把这个服务接进 DSH 时踩了两个坑，都修好了，记录在这里省得别人再踩。

### 抓到的真实协议流

DSH 内置的 MCP 客户端比 2024-11-05 规范要新，握手时会先探测一个 `server/discover`：

```
IN  <<< {"method":"server/discover","params":{"_meta":{
             "io.modelcontextprotocol/protocolVersion":"2026-07-28", ...}}}
OUT >>> {"error":{"code":-32601,"message":"不支持的方法: server/discover"}}

IN  <<< {"method":"initialize","params":{"protocolVersion":"2025-11-25", ...}}
OUT >>> {"protocolVersion":"2024-11-05","capabilities":{"tools":{}}, ...}

IN  <<< {"jsonrpc":"2.0","method":"notifications/initialized"}
IN  <<< {"method":"tools/list","jsonrpc":"2.0","id":1}
OUT >>> {"id":1,"result":{"tools":[{"name":"doc_read", ...}]}}
```

**好消息**：回 2024-11-05 它仍然接受，向下兼容。拒掉 `server/discover` 也不影响握手。

### 坑一：Windows 上中文变乱码，还会崩

接上后工具描述显示成 `��ȡ Word`。原因是 Python 连管道时默认用 **GBK**，而客户端按 UTF-8 解码。

更麻烦的是会**直接崩**——文档里有 `✅` 这种 GBK 编不出来的字符：

```
'gbk' codec can't encode character '\u2705' in position 22
```

双保险修法（`server.py` 里的 `_force_utf8` + `_dumps`）：

```python
for stream in (sys.stdout, sys.stderr):
    stream.reconfigure(encoding="utf-8", errors="replace", newline="\n")

# 万一 reconfigure 没生效，就退化成纯 ASCII 输出（中文变 \uXXXX 转义）
ascii_only = (sys.stdout.encoding or "").lower().replace("-", "") != "utf8"
json.dumps(payload, ensure_ascii=ascii_only)
```

JSON 客户端解析转义后会还原成正常中文，所以**任何控制台编码下都不会崩**。

### 坑二：回复"通知"会被判违规

JSON-RPC 规定没有 `id` 的消息是通知，**不能回复**。我早期对未知通知也回了一条 `id: null` 的错误，客户端直接判定协议违规断开。

```python
is_notification = "id" not in request
...
if is_notification:
    return None
```

### 接入方式：配置型 bundle

DSH 里不需要写插件代码，一个纯配置文件就能挂上（样例见 [`examples/dsh-bundle/`](examples/dsh-bundle/)）：

```yaml
- insert:
    - id: office-mcp
      name: '@deepseek-ai/dsh-mcp-client'
      config:
        serverName: office
        transport: stdio
        command: '<装了依赖的那个 python 绝对路径>'
        args:
          - '<server.py 的绝对路径>'
        failOnStartupError: true
```

装好后工具名是 `mcp__office__doc_read` 这种格式。

### 一个必须知道的运维细节

**调试时如果手动杀掉 MCP 子进程，工具调用会一直报 `Connection closed`。** 这时客户端根本不会发 `tools/call`（抓包可见），靠开关插件重连也不够——得**完全退出并重启 DSH**。正常运行不会遇到，只有调试时才要注意。

## 实现细节

### Word 按原顺序输出

`document.paragraphs` 和 `document.tables` 是两个独立列表，分别遍历会**丢掉表格和段落的相对位置**。所以直接遍历 XML body：

```python
for child in document.element.body.iterchildren():
    tag = child.tag.split("}")[-1]
    if tag == "p":    ...Paragraph(child, document)
    elif tag == "tbl": ...Table(child, document)
```

### Excel 取公式结果而不是公式

```python
openpyxl.load_workbook(path, data_only=True)
```

代价：如果文件是程序生成、**从未被 Excel 打开过**，公式没有缓存值，读出来是 `None`。这是 Excel 的机制，不是 bug。

### 截断要说清楚

不是静默截断，而是明确告诉模型：

```
…（已截断，原文 24311 字符，只返回前 8000 字符）
```

这样模型知道"我没看到全部"，不会拿半份内容下结论。

## 已知限制

1. **只读**。写入工具还没做——写坏原文件是最糟的失败模式，要单独设计（先写副本、变更前备份）。
2. 不支持旧格式 `.doc` / `.xls`，请先另存为新格式。
3. **扫描件 PDF 提不出文字**（没有文字层），需要 OCR，本项目不做。
4. 复杂排版（艺术字、SmartArt、文本框嵌套）读不干净，只保证正文、表格、备注。
5. `slide_read` 不输出图表里的数据（图表是图片级对象）。

## Roadmap

- [ ] `doc_write`：Markdown → Word（先写副本，不覆盖原文件）
- [ ] `sheet_append`：往工作表追加行
- [ ] `office_convert`：docx ↔ md ↔ pdf
- [ ] 批量：一次读整个目录的文档并汇总

## License

MIT
