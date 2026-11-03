"""office-mcp 的工具实现：把各种办公文档读成模型能吃的 Markdown。

设计原则：
  - 只读优先（写坏了原文件是最糟的失败），写入单独实现且默认不覆盖
  - 所有输出都做过截断，防止一个 300 页的文档把上下文撑爆
  - 失败不抛裸异常，返回结构化错误让 MCP 层决定怎么报
"""

import os
import re

MAX_CHARS = 8000          # 单个工具默认返回上限
MAX_TABLE_ROWS = 100      # 表格最多返回多少行

HEADING = re.compile(r"^(?:Heading|标题)\s*(\d)$", re.IGNORECASE)


def clip(text, limit=MAX_CHARS):
    """截断并明确告知截断了多少"""
    if len(text) <= limit:
        return text
    return text[:limit] + f"\n\n…（已截断，原文 {len(text)} 字符，只返回前 {limit} 字符）"


def _error(message):
    """统一错误结构：isError 由 server 层转换成 MCP 的 isError"""
    return {"ok": False, "error": message}


# ---------------- Word ----------------

def _paragraph_markdown(paragraph, doc):
    """普通段落 → Markdown 行；标题按样式映射成 #，列表加 -"""
    text = paragraph.text.strip()
    if not text:
        return ""
    style_name = ""
    try:
        style_name = paragraph.style.name or ""
    except AttributeError:
        pass

    match = HEADING.match(style_name)
    if match:
        level = min(int(match.group(1)), 6)
        return "#" * level + " " + text
    if style_name.lower() == "title":
        return "# " + text
    if "List" in style_name or "列表" in style_name:
        try:
            if paragraph.paragraph_format.left_indent:
                return "  - " + text
        except AttributeError:
            pass
        return "- " + text
    return text


def _table_markdown(table, limit_rows=MAX_TABLE_ROWS):
    """word 表格 → Markdown 表格"""
    rows = []
    for index, row in enumerate(table.rows):
        if index >= limit_rows:
            rows.append(["…"] * len(table.columns))
            break
        cells = []
        for cell in row.cells:
            value = cell.text.replace("\n", " ").replace("|", "\\|").strip()
            cells.append(value)
        rows.append(cells)
    if not rows:
        return ""
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    lines = ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * width]
    for row in rows[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def doc_read(path, limit=MAX_CHARS):
    """读 Word 文档，输出 Markdown（段落、标题、列表、表格按原顺序）"""
    if not os.path.isfile(path):
        return _error(f"文件不存在: {path}")
    if not path.lower().endswith((".docx", ".docm")):
        return _error("只支持 .docx（旧的 .doc 请先另存为 .docx）")
    try:
        import docx
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError:
        return _error("缺少 python-docx，请先 pip install python-docx")

    try:
        document = docx.Document(path)
    except Exception as err:  # noqa: BLE001 - 文档损坏也常有，得给出人话
        return _error(f"打不开这个 docx: {err}")

    blocks = []
    for child in document.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            line = _paragraph_markdown(Paragraph(child, document), document)
            if line:
                blocks.append(line)
        elif tag == "tbl":
            table = _table_markdown(Table(child, document))
            if table:
                blocks.append("")
                blocks.append(table)
                blocks.append("")

    body = "\n".join(blocks).strip()
    stats = f"（{os.path.basename(path)}：{len(document.paragraphs)} 段、{len(document.tables)} 表）"
    return {"ok": True, "text": clip(f"# {os.path.basename(path)}\n{stats}\n\n{body}", limit)}

# ---------------- Excel ----------------

def sheet_read(path, sheet=None, limit=MAX_CHARS, max_rows=MAX_TABLE_ROWS):
    """读 Excel：列出所有工作表，把指定表转成 Markdown 表格"""
    if not os.path.isfile(path):
        return _error(f"文件不存在: {path}")
    if not path.lower().endswith((".xlsx", ".xlsm")):
        return _error("只支持 .xlsx（旧的 .xls 请先另存为 .xlsx）")
    try:
        import openpyxl
    except ImportError:
        return _error("缺少 openpyxl，请先 pip install openpyxl")

    try:
        # data_only=True 取公式的计算结果而不是公式本身；
        # 但如果是程序生成、从未被 Excel 打开过的文件，这里会是 None
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as err:  # noqa: BLE001
        return _error(f"打不开这个 xlsx: {err}")

    names = workbook.sheetnames
    target = sheet if sheet in names else names[0]
    if sheet and sheet not in names:
        note = f"（没有名为「{sheet}」的工作表，已改用第一张「{target}」）"
    else:
        note = ""

    worksheet = target in names and workbook[target]
    if not worksheet:
        workbook.close()
        return _error("工作簿里没有可读的工作表")

    rows, truncated_rows, empty_streak = [], False, 0
    for index, row in enumerate(worksheet.iter_rows(values_only=True)):
        if index >= max_rows:
            truncated_rows = True
            break
        values = ["" if v is None else str(v) for v in row]
        if not any(v.strip() for v in values):
            empty_streak += 1
            if empty_streak >= 5:      # 连续空行一般是格式残留，停
                break
            continue
        empty_streak = 0
        rows.append(values)
    workbook.close()

    if not rows:
        return {"ok": True, "text": f"# {os.path.basename(path)}\n工作表：{', '.join(names)}\n\n（「{target}」是空的）"}

    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    header, body = rows[0], rows[1:]
    lines = ["| " + " | ".join(header) + " |", "|" + "---|" * width]
    for row in body:
        lines.append("| " + " | ".join(c.replace("|", "\\|") for c in row) + " |")

    tail = ""
    if truncated_rows:
        tail = f"\n\n…（只返回前 {max_rows} 行）"
    text = (
        f"# {os.path.basename(path)}\n"
        f"工作表：{', '.join(names)}{note}\n"
        f"当前：「{target}」，{len(body)} 行数据\n\n"
        + "\n".join(lines)
        + tail
    )
    return {"ok": True, "text": clip(text, limit)}


# ---------------- PowerPoint ----------------

def slide_read(path, limit=MAX_CHARS):
    """读 PPT：按页输出标题 + 正文要点 + 备注"""
    if not os.path.isfile(path):
        return _error(f"文件不存在: {path}")
    if not path.lower().endswith((".pptx", ".pptm")):
        return _error("只支持 .pptx")
    try:
        from pptx import Presentation
    except ImportError:
        return _error("缺少 python-pptx，请先 pip install python-pptx")

    try:
        presentation = Presentation(path)
    except Exception as err:  # noqa: BLE001
        return _error(f"打不开这个 pptx: {err}")

    blocks = []
    for number, slide in enumerate(presentation.slides, 1):
        blocks.append(f"## 第 {number} 页")
        title = ""
        if slide.shapes.title is not None and slide.shapes.title.has_text_frame:
            title = slide.shapes.title.text_frame.text.strip()
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            text_frame = shape.text_frame
            shape_text = text_frame.text.strip()
            if not shape_text:
                continue
            if shape is slide.shapes.title or shape_text == title:
                continue
            for paragraph in text_frame.paragraphs:
                line = paragraph.text.strip()
                if line:
                    blocks.append(f"- {line}")
            if getattr(shape, "has_table", False) and shape.has_table:
                blocks.append(_table_markdown(shape.table))
        if not title and len(blocks) >= 1 and blocks[-1].startswith("## "):
            blocks.append("（这页没有标题占位符）")
        notes = ""
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes = slide.notes_slide.notes_text_frame.text.strip()
        except (AttributeError, KeyError):
            notes = ""
        if notes:
            blocks.append(f"> 备注：{notes}")
        blocks.append("")

    text = f"# {os.path.basename(path)}（{len(presentation.slides)} 页）\n\n" + "\n".join(blocks).strip()
    return {"ok": True, "text": clip(text, limit)}

# ---------------- PDF ----------------

def pdf_extract(path, pages=None, limit=MAX_CHARS):
    """提 PDF 文本。pages 形如 '1-5' 或 '3'，不填取前 10 页"""
    if not os.path.isfile(path):
        return _error(f"文件不存在: {path}")
    if not path.lower().endswith(".pdf"):
        return _error("只支持 .pdf")
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf   # 旧名字，兼容老版本
        except ImportError:
            return _error("缺少 pymupdf，请先 pip install pymupdf")

    try:
        document = pymupdf.open(path)
    except Exception as err:  # noqa: BLE001
        return _error(f"打不开这个 pdf: {err}")

    total = document.page_count
    start, end = 1, min(total, 10)
    if pages:
        match = re.match(r"^(\d+)\s*(?:-\s*(\d+))?$", str(pages).strip())
        if not match:
            document.close()
            return _error("pages 格式要写成 '3' 或 '1-5'")
        start = max(1, int(match.group(1)))
        end = min(total, int(match.group(2) or start))

    blocks = []
    for number in range(start - 1, end):
        try:
            text = document[number].get_text().strip()
        except Exception:  # noqa: BLE001 - 单页坏了不该整份失败
            text = "（这一页读取失败）"
        blocks.append(f"## 第 {number + 1} 页\n{text or '（这一页没有可提取的文字，可能是扫描件）'}")
    document.close()

    text = f"# {os.path.basename(path)}（共 {total} 页，显示第 {start}-{end} 页）\n\n" + "\n\n".join(blocks)
    return {"ok": True, "text": clip(text, limit)}


# ---------------- 工具注册表 ----------------

TOOLS = {
    "doc_read": {
        "func": doc_read,
        "description": "读取 Word 文档（.docx），输出 Markdown：标题、段落、列表、表格按原文顺序。适合让我看报告、总结、说明文档。",
        "schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "docx 文件的绝对路径"},
                "limit": {"type": "integer", "description": "最多返回多少字符，默认 8000"},
            },
            "required": ["path"],
        },
    },
    "sheet_read": {
        "func": sheet_read,
        "description": "读取 Excel 工作簿（.xlsx），输出指定工作表的 Markdown 表格。会先列出所有工作表名。",
        "schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "xlsx 文件的绝对路径"},
                "sheet": {"type": "string", "description": "工作表名，不填取第一张"},
                "max_rows": {"type": "integer", "description": "最多返回多少行数据，默认 100"},
            },
            "required": ["path"],
        },
    },
    "slide_read": {
        "func": slide_read,
        "description": "读取 PowerPoint（.pptx），按页输出标题、正文要点、表格和演讲者备注。",
        "schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "pptx 文件的绝对路径"},
                "limit": {"type": "integer", "description": "最多返回多少字符，默认 8000"},
            },
            "required": ["path"],
        },
    },
    "pdf_extract": {
        "func": pdf_extract,
        "description": "提取 PDF 文本，可按页范围（如 1-5）。扫描件没有文字层，会明确告诉你。",
        "schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "pdf 文件的绝对路径"},
                "pages": {"type": "string", "description": "页范围，如 '3' 或 '1-5'，不填取前 10 页"},
            },
            "required": ["path"],
        },
    },
}