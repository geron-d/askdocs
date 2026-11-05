"""Build the synthetic SupportDesk corpus.

Reads Markdown sources from corpus/src, renders manuals to PDF and admin
guides to DOCX, and writes a machine-readable FAQ plus a manifest.

Run with the repository root as the working directory:
    pip install -r corpus/requirements.txt
    python corpus/build.py

For Cyrillic PDFs a Unicode TTF font is required. The script looks in
/fonts first (Arial on Windows) and then in the DejaVu system directory.
"""

from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Preformatted,
    Spacer,
    Table,
    TableStyle,
)

BASE = Path(__file__).resolve().parent
SRC = BASE / "src"
MANUAL_DIR = BASE / "manual"
ADMIN_DIR = BASE / "admin"
VERSION = "1.0"
BUILD_DATE = date.today().isoformat()


def register_fonts() -> tuple[str, str]:
    candidates = [
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "DejaVuSans",
        ),
        ("/fonts/arial.ttf", "/fonts/arialbd.ttf", "Arial"),
    ]
    for regular, bold, name in candidates:
        if Path(regular).exists() and Path(bold).exists():
            pdfmetrics.registerFont(TTFont(name, regular))
            pdfmetrics.registerFont(TTFont(f"{name}-Bold", bold))
            return name, f"{name}-Bold"
    raise SystemExit(
        "No Unicode font found. Mount a font directory at /fonts with "
        "arial.ttf/arialbd.ttf, or install fonts-dejavu-core."
    )


FONT_REGULAR, FONT_BOLD = register_fonts()


def inline_pdf(text: str) -> str:
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"`(.+?)`", r'<font name="Courier">\1</font>', text)
    return text


def parse_markdown(path: Path) -> dict:
    lines = path.read_text(encoding="utf-8").splitlines()
    title = lines[0][2:].strip() if lines and lines[0].startswith("# ") else path.stem
    subtitle = ""
    blocks: list[dict] = []
    i = 0
    first_paragraph = True
    while i < len(lines):
        line = lines[i].rstrip()
        if not line:
            i += 1
            continue
        if line.startswith("# "):
            i += 1
            continue
        if line.startswith("## "):
            blocks.append({"type": "h2", "text": line[3:].strip()})
            first_paragraph = True
            i += 1
            continue
        if line.startswith("### "):
            blocks.append({"type": "h3", "text": line[4:].strip()})
            i += 1
            continue
        if line.startswith("```"):
            code: list[str] = []
            i += 1
            while i < len(lines) and not lines[i].startswith("```"):
                code.append(lines[i])
                i += 1
            i += 1
            blocks.append({"type": "code", "text": "\n".join(code)})
            continue
        if line.startswith("|"):
            rows: list[list[str]] = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                row = lines[i].strip().strip("|")
                cells = [cell.strip() for cell in row.split("|")]
                if not all(re.fullmatch(r":?-{3,}:?", cell) for cell in cells):
                    rows.append(cells)
                i += 1
            if rows:
                blocks.append({"type": "table", "header": rows[0], "rows": rows[1:]})
            continue
        if re.match(r"^- ", line):
            items: list[str] = []
            while i < len(lines) and re.match(r"^- ", lines[i].rstrip()):
                items.append(lines[i].rstrip()[2:].strip())
                i += 1
            blocks.append({"type": "ul", "items": items})
            continue
        if re.match(r"^\d+\. ", line):
            items = []
            while i < len(lines) and re.match(r"^\d+\. ", lines[i].rstrip()):
                items.append(re.sub(r"^\d+\.\s+", "", lines[i].rstrip()))
                i += 1
            blocks.append({"type": "ol", "items": items})
            continue
        if line.startswith("> "):
            blocks.append({"type": "quote", "text": line[2:].strip()})
            i += 1
            continue
        paragraph: list[str] = []
        while i < len(lines):
            current = lines[i].rstrip()
            if (
                not current
                or current.startswith("#")
                or current.startswith("```")
                or current.startswith("|")
                or re.match(r"^- ", current)
                or re.match(r"^\d+\. ", current)
                or current.startswith("> ")
            ):
                break
            paragraph.append(current)
            i += 1
        text = " ".join(paragraph).strip()
        if text:
            if first_paragraph and not subtitle:
                subtitle = text
                first_paragraph = False
            else:
                blocks.append({"type": "p", "text": text})
    return {"title": title, "subtitle": subtitle, "blocks": blocks, "source": path}


def pdf_styles() -> dict[str, ParagraphStyle]:
    styles = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "DocTitle",
            parent=styles["Title"],
            fontName=FONT_BOLD,
            fontSize=26,
            leading=32,
            textColor=colors.HexColor("#1F3A5F"),
            spaceAfter=14,
            alignment=0,
        ),
        "subtitle": ParagraphStyle(
            "DocSubtitle",
            parent=styles["Normal"],
            fontName=FONT_REGULAR,
            fontSize=12,
            leading=17,
            textColor=colors.HexColor("#4A5568"),
            spaceAfter=8,
        ),
        "meta": ParagraphStyle(
            "DocMeta",
            parent=styles["Normal"],
            fontName=FONT_REGULAR,
            fontSize=9,
            leading=13,
            textColor=colors.HexColor("#718096"),
        ),
        "h2": ParagraphStyle(
            "H2",
            parent=styles["Heading2"],
            fontName=FONT_BOLD,
            fontSize=15,
            leading=20,
            textColor=colors.HexColor("#1F3A5F"),
            spaceBefore=16,
            spaceAfter=7,
            keepWithNext=True,
        ),
        "h3": ParagraphStyle(
            "H3",
            parent=styles["Heading3"],
            fontName=FONT_BOLD,
            fontSize=11.5,
            leading=16,
            textColor=colors.HexColor("#2D3748"),
            spaceBefore=10,
            spaceAfter=5,
            keepWithNext=True,
        ),
        "body": ParagraphStyle(
            "Body",
            parent=styles["BodyText"],
            fontName=FONT_REGULAR,
            fontSize=10.5,
            leading=15.5,
            alignment=TA_LEFT,
            spaceAfter=7,
        ),
        "bullet": ParagraphStyle(
            "Bullet",
            parent=styles["BodyText"],
            fontName=FONT_REGULAR,
            fontSize=10.5,
            leading=15,
            leftIndent=14,
            bulletIndent=4,
            spaceAfter=3,
        ),
        "code": ParagraphStyle(
            "Code",
            parent=styles["Code"],
            fontName="Courier",
            fontSize=8,
            leading=11,
            textColor=colors.HexColor("#1A202C"),
            backColor=colors.HexColor("#F2F4F7"),
            borderPadding=6,
            spaceBefore=4,
            spaceAfter=8,
        ),
        "quote": ParagraphStyle(
            "Quote",
            parent=styles["BodyText"],
            fontName=FONT_REGULAR,
            fontSize=10,
            leading=15,
            leftIndent=12,
            textColor=colors.HexColor("#4A5568"),
            borderPadding=6,
            backColor=colors.HexColor("#F7FAFC"),
        ),
        "cell": ParagraphStyle(
            "Cell",
            parent=styles["BodyText"],
            fontName=FONT_REGULAR,
            fontSize=8,
            leading=11,
        ),
        "cell_bold": ParagraphStyle(
            "CellBold",
            parent=styles["BodyText"],
            fontName=FONT_BOLD,
            fontSize=8,
            leading=11,
        ),
    }


def make_pdf(doc: dict, out_path: Path) -> None:
    styles = pdf_styles()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc_title = doc["title"]
    story: list = []
    story.append(Spacer(1, 4 * cm))
    story.append(Paragraph(inline_pdf(doc_title), styles["title"]))
    if doc["subtitle"]:
        story.append(Paragraph(inline_pdf(doc["subtitle"]), styles["subtitle"]))
    story.append(Spacer(1, 1 * cm))
    story.append(
        Paragraph(
            f"Синтетический корпус демо · версия {VERSION} · {BUILD_DATE}",
            styles["meta"],
        )
    )
    story.append(PageBreak())

    headings = [b["text"] for b in doc["blocks"] if b["type"] == "h2"]
    if headings:
        story.append(Paragraph("Содержание", styles["h2"]))
        for number, heading in enumerate(headings, start=1):
            story.append(
                Paragraph(f"{number}. {inline_pdf(heading)}", styles["body"])
            )
        story.append(PageBreak())

    for block in doc["blocks"]:
        kind = block["type"]
        if kind == "h2":
            story.append(Paragraph(inline_pdf(block["text"]), styles["h2"]))
        elif kind == "h3":
            story.append(Paragraph(inline_pdf(block["text"]), styles["h3"]))
        elif kind == "p":
            story.append(Paragraph(inline_pdf(block["text"]), styles["body"]))
        elif kind == "ul":
            for item in block["items"]:
                story.append(
                    Paragraph(
                        inline_pdf(item),
                        styles["bullet"],
                        bulletText="•",
                    )
                )
        elif kind == "ol":
            for number, item in enumerate(block["items"], start=1):
                story.append(
                    Paragraph(
                        inline_pdf(item),
                        styles["bullet"],
                        bulletText=f"{number}.",
                    )
                )
        elif kind == "quote":
            story.append(Paragraph(inline_pdf(block["text"]), styles["quote"]))
        elif kind == "code":
            story.append(Preformatted(block["text"], styles["code"]))
        elif kind == "table":
            header = [
                Paragraph(inline_pdf(cell), styles["cell_bold"])
                for cell in block["header"]
            ]
            rows = [
                [Paragraph(inline_pdf(cell), styles["cell"]) for cell in row]
                for row in block["rows"]
            ]
            table = Table([header, *rows], repeatRows=1, hAlign="LEFT")
            table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF5")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
                        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E0")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 5),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ]
                )
            )
            story.append(KeepTogether([table, Spacer(1, 8)]))

    def footer(canvas, template):
        canvas.saveState()
        canvas.setFont(FONT_REGULAR, 8)
        canvas.setFillColor(colors.HexColor("#718096"))
        canvas.drawString(2 * cm, 1.2 * cm, "SupportDesk · синтетический корпус")
        canvas.drawRightString(
            A4[0] - 2 * cm, 1.2 * cm, f"{doc_title} · стр. {canvas.getPageNumber()}"
        )
        canvas.restoreState()

    pdf = BaseDocTemplate(
        str(out_path),
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm,
        title=doc_title,
        author="SupportDesk demo",
        subject="Synthetic helpdesk corpus",
    )
    frame = Frame(
        pdf.leftMargin,
        pdf.bottomMargin,
        pdf.width,
        pdf.height,
        id="normal",
    )
    pdf.addPageTemplates([PageTemplate(id="page", frames=[frame], onPage=footer)])
    pdf.build(story)


def add_page_number(paragraph) -> None:
    run = paragraph.add_run()
    field_begin = OxmlElement("w:fldChar")
    field_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    field_end = OxmlElement("w:fldChar")
    field_end.set(qn("w:fldCharType"), "end")
    run._r.append(field_begin)
    run._r.append(instr)
    run._r.append(field_end)


def add_docx_runs(paragraph, text: str) -> None:
    parts = re.split(r"(\*\*.+?\*\*|`.+?`)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith("`") and part.endswith("`"):
            run = paragraph.add_run(part[1:-1])
            run.font.name = "Courier New"
        else:
            paragraph.add_run(part)


def make_docx(doc: dict, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(6)

    for style_name in ("Heading 1", "Heading 2", "Title"):
        style = document.styles[style_name]
        style.font.name = "Arial"
        style.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = title.add_run(doc["title"])
    run.bold = True
    run.font.size = Pt(26)
    run.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)

    if doc["subtitle"]:
        paragraph = document.add_paragraph()
        run = paragraph.add_run(doc["subtitle"])
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    paragraph = document.add_paragraph()
    run = paragraph.add_run(
        f"Синтетический корпус демо · версия {VERSION} · {BUILD_DATE}"
    )
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x71, 0x80, 0x96)
    document.add_page_break()

    headings = [b["text"] for b in doc["blocks"] if b["type"] == "h2"]
    if headings:
        document.add_heading("Содержание", level=1)
        for number, heading in enumerate(headings, start=1):
            document.add_paragraph(f"{number}. {heading}")
        document.add_page_break()

    for block in doc["blocks"]:
        kind = block["type"]
        if kind == "h2":
            document.add_heading(block["text"], level=1)
        elif kind == "h3":
            document.add_heading(block["text"], level=2)
        elif kind == "p":
            paragraph = document.add_paragraph()
            add_docx_runs(paragraph, block["text"])
        elif kind == "ul":
            for item in block["items"]:
                paragraph = document.add_paragraph(style="List Bullet")
                add_docx_runs(paragraph, item)
        elif kind == "ol":
            for item in block["items"]:
                paragraph = document.add_paragraph(style="List Number")
                add_docx_runs(paragraph, item)
        elif kind == "quote":
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.left_indent = Cm(0.7)
            run = paragraph.add_run(block["text"])
            run.italic = True
        elif kind == "code":
            paragraph = document.add_paragraph()
            run = paragraph.add_run(block["text"])
            run.font.name = "Courier New"
            run.font.size = Pt(8)
        elif kind == "table":
            table = document.add_table(
                rows=1, cols=len(block["header"])
            )
            table.style = "Table Grid"
            for index, cell_text in enumerate(block["header"]):
                cell = table.rows[0].cells[index]
                cell.text = ""
                paragraph = cell.paragraphs[0]
                run = paragraph.add_run(cell_text)
                run.bold = True
                run.font.size = Pt(9)
            for row in block["rows"]:
                cells = table.add_row().cells
                for index, cell_text in enumerate(row):
                    cells[index].text = ""
                    paragraph = cells[index].paragraphs[0]
                    add_docx_runs(paragraph, cell_text)
                    for run in paragraph.runs:
                        run.font.size = Pt(9)
            document.add_paragraph()

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = footer.add_run(f"{doc['title']} · стр. ")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x71, 0x80, 0x96)
    add_page_number(footer)

    document.core_properties.title = doc["title"]
    document.core_properties.author = "SupportDesk demo"
    document.core_properties.subject = "Synthetic helpdesk corpus"
    document.save(str(out_path))


def write_faq() -> int:
    source = json.loads((SRC / "faq.json").read_text(encoding="utf-8"))
    ids = [item["id"] for item in source]
    if len(ids) != len(set(ids)):
        raise SystemExit("Duplicate FAQ ids")
    for item in source:
        for field in ("id", "category", "question", "answer"):
            if not item.get(field):
                raise SystemExit(f"FAQ item {item.get('id')} lacks {field}")
    payload = {
        "source": "faq.json",
        "doc_type": "faq",
        "version": VERSION,
        "generated": BUILD_DATE,
        "items": source,
    }
    (BASE / "faq.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return len(source)


def main() -> None:
    documents = []
    for path in sorted(SRC.glob("*.md")):
        doc = parse_markdown(path)
        if path.name.startswith("manual-"):
            out = MANUAL_DIR / f"{path.stem.removeprefix('manual-')}.pdf"
            make_pdf(doc, out)
            doc_type = "manual"
        elif path.name.startswith("admin-"):
            out = ADMIN_DIR / f"{path.stem.removeprefix('admin-')}.docx"
            make_docx(doc, out)
            doc_type = "admin"
        else:
            continue
        text = path.read_text(encoding="utf-8")
        documents.append(
            {
                "path": out.relative_to(BASE).as_posix(),
                "source": path.relative_to(BASE).as_posix(),
                "doc_type": doc_type,
                "title": doc["title"],
                "sections": sum(1 for b in doc["blocks"] if b["type"] == "h2"),
                "chars": len(text),
            }
        )
        print(f"wrote {out.relative_to(BASE)}")

    faq_count = write_faq()
    print(f"wrote faq.json ({faq_count} items)")
    manifest = {
        "version": VERSION,
        "generated": BUILD_DATE,
        "documents": documents,
        "faq_items": faq_count,
    }
    (BASE / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print("wrote manifest.json")


if __name__ == "__main__":
    main()
