#!/usr/bin/env python3
"""Convert the full RIIM markdown proposal into a polished Word document."""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

MD_PATH = Path("/workspace/proposal-riim/PROPOSAL_RIIM_Domba_Temanggung.md")
OUT_PATH = Path("/workspace/proposal-riim/PROPOSAL_RIIM_Domba_Temanggung.docx")
RAB_PATH = Path("/workspace/proposal-riim/RAB_Ringkas_RIIM_Domba_Temanggung.md")


def set_run_font(run, size=12, bold=False, italic=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def add_runs_with_markup(paragraph, text, size=12):
    """Support simple **bold** and *italic* markdown markers."""
    pattern = re.compile(r"(\*\*.+?\*\*|\*.+?\*)")
    parts = pattern.split(text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) >= 4:
            run = paragraph.add_run(part[2:-2])
            set_run_font(run, size=size, bold=True)
        elif part.startswith("*") and part.endswith("*") and len(part) >= 2:
            run = paragraph.add_run(part[1:-1])
            set_run_font(run, size=size, italic=True)
        else:
            run = paragraph.add_run(part)
            set_run_font(run, size=size)


def style_paragraph(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6, first_indent=None):
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if first_indent is not None:
        p.paragraph_format.first_line_indent = Cm(first_indent)


def add_heading(doc, text, level=1):
    # Strip markdown heading markers already handled by caller
    p = doc.add_paragraph()
    style_paragraph(p, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=8)
    sizes = {1: 14, 2: 12, 3: 12}
    run = p.add_run(text)
    set_run_font(run, size=sizes.get(level, 12), bold=True)
    return p


def add_center_line(doc, text, bold=False, size=12, space_after=4):
    p = doc.add_paragraph()
    style_paragraph(p, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=space_after)
    add_runs_with_markup(p, text, size=size)
    for run in p.runs:
        run.bold = bold or run.bold
    return p


def parse_table(lines, start_idx):
    """Parse a GitHub-style markdown table starting at start_idx. Returns (rows, next_idx)."""
    rows = []
    i = start_idx
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.startswith("|"):
            break
        # skip separator
        if re.match(r"^\|\s*:?-{3,}", line):
            i += 1
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        rows.append(cells)
        i += 1
    return rows, i


def add_table(doc, rows):
    if not rows:
        return
    cols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=cols)
    table.style = "Table Grid"
    for r_idx, row in enumerate(rows):
        for c_idx in range(cols):
            cell = table.rows[r_idx].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            style_paragraph(p, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2)
            text = row[c_idx] if c_idx < len(row) else ""
            add_runs_with_markup(p, text, size=10)
            if r_idx == 0:
                for run in p.runs:
                    run.bold = True
    doc.add_paragraph()


def add_code_block(doc, lines):
    p = doc.add_paragraph()
    style_paragraph(p, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=8)
    run = p.add_run("\n".join(lines))
    set_run_font(run, size=9)
    run.font.name = "Courier New"


def convert_markdown(doc, md_text, page_break_before_h1_after_cover=True):
    lines = md_text.replace("\r\n", "\n").split("\n")
    i = 0
    seen_first_h1 = False
    in_code = False
    code_buf = []

    while i < len(lines):
        line = lines[i]

        # code fences
        if line.strip().startswith("```"):
            if in_code:
                add_code_block(doc, code_buf)
                code_buf = []
                in_code = False
            else:
                in_code = True
            i += 1
            continue
        if in_code:
            code_buf.append(line)
            i += 1
            continue

        # horizontal rule / page-ish separator
        if line.strip() == "---":
            i += 1
            continue

        # tables
        if line.startswith("|"):
            rows, i = parse_table(lines, i)
            add_table(doc, rows)
            continue

        # headings
        m = re.match(r"^(#{1,3})\s+(.*)$", line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            # page break before main body after cover/pengesahan when hitting "1. Judul" or "SISTEMATIKA"
            if level == 1 and seen_first_h1 and page_break_before_h1_after_cover:
                # keep flowing; only break before "1. Judul Riset"
                if text.startswith("1. Judul") or text.startswith("SISTEMATIKA"):
                    doc.add_page_break()
            if level == 1:
                seen_first_h1 = True
                if text.startswith("HALAMAN PENGESAHAN") or text.startswith("PROPOSAL RISET"):
                    # cover/pengesahan titles centered
                    add_center_line(doc, text, bold=True, size=14, space_after=8)
                    i += 1
                    continue
            add_heading(doc, text, level=level)
            i += 1
            continue

        # blank
        if not line.strip():
            i += 1
            continue

        # bullet / numbered already as plain paragraphs with indent
        p = doc.add_paragraph()
        is_list = bool(re.match(r"^(\d+\.|-|\*)\s+", line.strip()))
        style_paragraph(
            p,
            align=WD_ALIGN_PARAGRAPH.JUSTIFY,
            space_after=6,
            first_indent=None if is_list else 1.0,
        )
        add_runs_with_markup(p, line.strip(), size=12)
        i += 1


def build():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.5)

    # Normal style
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

    md = MD_PATH.read_text(encoding="utf-8")
    # Drop the top-level duplicate "PROPOSAL..." if present; keep content as-is
    convert_markdown(doc, md)

    # Append RAB ringkas as appendix
    if RAB_PATH.exists():
        doc.add_page_break()
        add_heading(doc, "LAMPIRAN. RAB RINGKAS", level=1)
        convert_markdown(doc, RAB_PATH.read_text(encoding="utf-8"), page_break_before_h1_after_cover=False)

    doc.save(OUT_PATH)
    print(f"Saved Word proposal: {OUT_PATH}")


if __name__ == "__main__":
    build()
