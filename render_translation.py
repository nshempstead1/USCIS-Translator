#!/usr/bin/env python3
"""
USCIS Certified Translation Renderer
=====================================
Consumes structured JSON from the Gemini translation pipeline and renders
a finished .docx on the Corpus Localization LLC letterhead template.

Usage:
    python render_translation.py translation.json [--template template.docx] [--output output.docx]

Dependencies:
    pip install python-docx jsonschema
"""

import argparse
import json
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, Mm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

FONT_NAME = "Times New Roman"
FONT_SIZE_BODY = Pt(11)
FONT_SIZE_H1 = Pt(14)
FONT_SIZE_H2 = Pt(12)
FONT_SIZE_SMALL = Pt(10)
FONT_SIZE_VERY_SMALL = Pt(9)

# A4 usable width with 25mm left+right margins = 160mm
USABLE_WIDTH_MM = 160
USABLE_WIDTH = Mm(USABLE_WIDTH_MM)

# Table cell shading for separators / headers
SHADE_LIGHT = "F2F2F2"
SHADE_ILLEGIBLE = "FFFF00"

ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def set_cell_shading(cell, color_hex):
    """Apply background shading to a table cell."""
    shading = parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>'
    )
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    """Set individual borders on a cell. Each kwarg is a dict like
    {"sz": "4", "val": "single", "color": "000000"}."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
        tcPr.append(borders)
    for side, spec in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        if spec:
            el = parse_xml(
                f'<w:{side} {nsdecls("w")} w:val="{spec.get("val", "single")}" '
                f'w:sz="{spec.get("sz", "4")}" w:space="0" '
                f'w:color="{spec.get("color", "000000")}"/>'
            )
            existing = borders.find(qn(f"w:{side}"))
            if existing is not None:
                borders.remove(existing)
            borders.append(el)


def set_no_borders(cell):
    """Remove all borders from a cell."""
    spec = {"val": "none", "sz": "0", "color": "auto"}
    set_cell_border(cell, top=spec, bottom=spec, left=spec, right=spec)


def set_all_borders(cell):
    """Set solid black borders on all sides."""
    spec = {"val": "single", "sz": "4", "color": "000000"}
    set_cell_border(cell, top=spec, bottom=spec, left=spec, right=spec)


def apply_font(run, size=None, bold=False, italic=False, name=FONT_NAME, color=None):
    """Apply consistent font styling to a run."""
    run.font.name = name
    run.font.size = size or FONT_SIZE_BODY
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    # Force Times New Roman for East Asian fallback
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:cs"), name)


def add_styled_paragraph(doc_or_cell, text, size=None, bold=False, italic=False,
                         alignment=WD_ALIGN_PARAGRAPH.LEFT, space_after=Pt(6),
                         space_before=Pt(0)):
    """Add a paragraph with consistent styling. Works on Document or table Cell."""
    if hasattr(doc_or_cell, "add_paragraph"):
        para = doc_or_cell.add_paragraph()
    else:
        para = doc_or_cell.paragraphs[0] if doc_or_cell.paragraphs else doc_or_cell.add_paragraph()

    para.alignment = alignment
    para.paragraph_format.space_after = space_after
    para.paragraph_format.space_before = space_before
    para.paragraph_format.line_spacing = Pt(14)

    # Handle multi-line text (newlines in JSON strings)
    lines = text.split("\n")
    for i, line in enumerate(lines):
        run = para.add_run(line)
        apply_font(run, size=size, bold=bold, italic=italic)
        # Highlight [ILLEGIBLE] markers
        if "[ILLEGIBLE" in line:
            run.font.highlight_color = 7  # Yellow
        if i < len(lines) - 1:
            para.add_run().add_break()

    return para


def add_cell_text(cell, text, size=None, bold=False, italic=False,
                  alignment=WD_ALIGN_PARAGRAPH.LEFT):
    """Add text to a table cell with zero spacing (Word-safe)."""
    # Clear default empty paragraph
    for p in cell.paragraphs:
        p_element = p._element
        p_element.getparent().remove(p_element)

    para = cell.add_paragraph()
    para.alignment = alignment
    para.paragraph_format.space_after = Pt(0)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.line_spacing = Pt(13)

    lines = text.split("\n")
    for i, line in enumerate(lines):
        run = para.add_run(line)
        apply_font(run, size=size or FONT_SIZE_BODY, bold=bold, italic=italic)
        if "[ILLEGIBLE" in line:
            run.font.highlight_color = 7
        if i < len(lines) - 1:
            para.add_run().add_break()

    return para


def set_table_col_widths(table, widths_pct):
    """Set column widths as percentages of the usable page width."""
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if idx < len(widths_pct):
                width = Mm(USABLE_WIDTH_MM * widths_pct[idx] / 100)
                cell.width = width


def add_page_break(doc):
    """Insert a page break in the document."""
    from docx.enum.text import WD_BREAK
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(0)
    para.paragraph_format.space_before = Pt(0)
    run = para.add_run()
    run.add_break(WD_BREAK.PAGE)
    return para


# ---------------------------------------------------------------------------
# Element Renderers
# ---------------------------------------------------------------------------

def render_institution_header(doc, element):
    """Render a multi-column institution header as a borderless table."""
    columns = element["columns"]
    table = doc.add_table(rows=1, cols=len(columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    widths = [col["width_pct"] for col in columns]

    for idx, col in enumerate(columns):
        cell = table.cell(0, idx)
        set_no_borders(cell)
        align = ALIGN_MAP.get(col.get("align", "center"), WD_ALIGN_PARAGRAPH.CENTER)
        add_cell_text(cell, col["content"], bold=True, alignment=align)

    set_table_col_widths(table, widths)


def render_heading(doc, element):
    """Render a heading (level 1 or 2)."""
    level = element.get("level", 1)
    size = FONT_SIZE_H1 if level == 1 else FONT_SIZE_H2
    add_styled_paragraph(
        doc, element["text"],
        size=size, bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=Pt(8), space_before=Pt(8),
    )


def render_paragraph(doc, element):
    """Render a body paragraph with optional styling."""
    style = element.get("style", "")
    bold = "bold" in style
    italic = "italic" in style
    alignment = WD_ALIGN_PARAGRAPH.CENTER if "center" in style else WD_ALIGN_PARAGRAPH.LEFT
    size = FONT_SIZE_SMALL if "small" in style else FONT_SIZE_BODY

    add_styled_paragraph(
        doc, element["text"],
        size=size, bold=bold, italic=italic,
        alignment=alignment,
        space_after=Pt(6),
    )


def render_info_block(doc, element):
    """Render key-value pairs as a two-column or stacked layout."""
    pairs = element["pairs"]
    layout = element.get("layout", "two_column")

    if layout == "two_column":
        # Two pairs per row
        rows_needed = (len(pairs) + 1) // 2
        table = doc.add_table(rows=rows_needed, cols=4)
        table.autofit = False
        widths = [20, 30, 20, 30]

        pair_idx = 0
        for row_idx in range(rows_needed):
            row = table.rows[row_idx]
            for col_pair in range(2):
                if pair_idx < len(pairs):
                    label_cell = row.cells[col_pair * 2]
                    value_cell = row.cells[col_pair * 2 + 1]
                    set_no_borders(label_cell)
                    set_no_borders(value_cell)
                    add_cell_text(label_cell, pairs[pair_idx]["label"] + ":",
                                  bold=True, size=FONT_SIZE_BODY)
                    add_cell_text(value_cell, pairs[pair_idx]["value"],
                                  size=FONT_SIZE_BODY)
                    pair_idx += 1
                else:
                    for c in [row.cells[col_pair * 2], row.cells[col_pair * 2 + 1]]:
                        set_no_borders(c)

        set_table_col_widths(table, widths)
    else:
        # Stacked: one pair per row, label | value
        table = doc.add_table(rows=len(pairs), cols=2)
        table.autofit = False
        widths = [30, 70]

        for row_idx, pair in enumerate(pairs):
            label_cell = table.rows[row_idx].cells[0]
            value_cell = table.rows[row_idx].cells[1]
            set_no_borders(label_cell)
            set_no_borders(value_cell)
            add_cell_text(label_cell, pair["label"] + ":", bold=True)
            add_cell_text(value_cell, pair["value"])

        set_table_col_widths(table, widths)


def render_table(doc, element):
    """Render a bordered data table with headers, data rows, and separator rows."""
    headers = element["headers"]
    col_widths = element["column_widths_pct"]
    rows = element["rows"]
    num_cols = len(headers)

    # Count actual rows needed (header + data/separator rows)
    table = doc.add_table(rows=1 + len(rows), cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Header row
    header_row = table.rows[0]
    for idx, header_text in enumerate(headers):
        cell = header_row.cells[idx]
        set_all_borders(cell)
        set_cell_shading(cell, SHADE_LIGHT)
        add_cell_text(cell, header_text, bold=True, size=FONT_SIZE_BODY)

    # Data / separator rows
    for row_idx, row_data in enumerate(rows):
        doc_row = table.rows[row_idx + 1]

        if row_data["type"] == "separator":
            # Merge all cells for a separator row
            # We need to merge after accessing cells
            first_cell = doc_row.cells[0]
            last_cell = doc_row.cells[num_cols - 1]
            merged = first_cell.merge(last_cell)
            set_all_borders(merged)
            set_cell_shading(merged, SHADE_LIGHT)
            add_cell_text(merged, row_data["text"], bold=True, size=FONT_SIZE_BODY)
        elif row_data["type"] == "data":
            cells_data = row_data["cells"]
            for col_idx in range(num_cols):
                cell = doc_row.cells[col_idx]
                set_all_borders(cell)
                text = cells_data[col_idx] if col_idx < len(cells_data) else ""
                add_cell_text(cell, text, size=FONT_SIZE_BODY)

    set_table_col_widths(table, col_widths)


def render_summary_block(doc, element):
    """Render a summary block (totals, GPA) as a compact table."""
    pairs = element["pairs"]
    table = doc.add_table(rows=len(pairs), cols=2)
    table.autofit = False
    widths = [50, 50]

    for row_idx, pair in enumerate(pairs):
        label_cell = table.rows[row_idx].cells[0]
        value_cell = table.rows[row_idx].cells[1]
        set_all_borders(label_cell)
        set_all_borders(value_cell)
        set_cell_shading(label_cell, SHADE_LIGHT)
        add_cell_text(label_cell, pair["label"], bold=True)
        add_cell_text(value_cell, pair["value"], bold=True)

    set_table_col_widths(table, widths)


def render_grading_scale(doc, element):
    """Render a grading scale legend table."""
    title = element.get("title", "Grading Scale")
    rows_data = element["rows"]

    add_styled_paragraph(
        doc, title,
        size=FONT_SIZE_BODY, bold=True,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        space_after=Pt(4), space_before=Pt(8),
    )

    table = doc.add_table(rows=len(rows_data) + 1, cols=2)
    table.autofit = False
    widths = [30, 70]

    # Header
    set_all_borders(table.rows[0].cells[0])
    set_all_borders(table.rows[0].cells[1])
    set_cell_shading(table.rows[0].cells[0], SHADE_LIGHT)
    set_cell_shading(table.rows[0].cells[1], SHADE_LIGHT)
    add_cell_text(table.rows[0].cells[0], "Range", bold=True)
    add_cell_text(table.rows[0].cells[1], "Description", bold=True)

    for row_idx, row_data in enumerate(rows_data):
        r = table.rows[row_idx + 1]
        set_all_borders(r.cells[0])
        set_all_borders(r.cells[1])
        add_cell_text(r.cells[0], row_data["range"])
        add_cell_text(r.cells[1], row_data["description"])

    set_table_col_widths(table, widths)


def render_signature_block(doc, element):
    """Render a signature block as a borderless multi-column table."""
    signatures = element["signatures"]
    num_sigs = len(signatures)
    table = doc.add_table(rows=3, cols=num_sigs)
    table.autofit = False
    widths = [100 // num_sigs] * num_sigs

    for idx, sig in enumerate(signatures):
        align = ALIGN_MAP.get(sig.get("align", "center"), WD_ALIGN_PARAGRAPH.CENTER)

        # Row 0: signature placeholder
        cell_sig = table.rows[0].cells[idx]
        set_no_borders(cell_sig)
        add_cell_text(cell_sig, sig.get("signature", "[Signature]"),
                      italic=True, alignment=align)

        # Row 1: name with underline effect
        cell_name = table.rows[1].cells[idx]
        set_no_borders(cell_name)
        name = sig.get("name", "")
        if name:
            # Add a top border to simulate signature line
            set_cell_border(cell_name, top={"val": "single", "sz": "4", "color": "000000"})
            add_cell_text(cell_name, name, bold=True, alignment=align)

        # Row 2: title
        cell_title = table.rows[2].cells[idx]
        set_no_borders(cell_title)
        title = sig.get("title", "")
        if title:
            add_cell_text(cell_title, title, size=FONT_SIZE_SMALL, alignment=align)

    set_table_col_widths(table, widths)


def render_separator(doc, _element):
    """Render a horizontal rule."""
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(4)
    para.paragraph_format.space_before = Pt(4)
    # Add a bottom border to the paragraph to simulate <hr>
    pPr = para._element.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="4" w:space="1" w:color="000000"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def render_footer_text(doc, element):
    """Render footer/fine-print text."""
    style = element.get("style", "small")
    size = FONT_SIZE_VERY_SMALL if style == "very_small" else FONT_SIZE_SMALL
    italic = style == "italic"
    add_styled_paragraph(
        doc, element["text"],
        size=size, italic=italic,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        space_after=Pt(4), space_before=Pt(4),
    )


def render_placeholder(doc, element):
    """Render a placeholder (e.g., [QR Code], [Photo])."""
    align = ALIGN_MAP.get(element.get("align", "center"), WD_ALIGN_PARAGRAPH.CENTER)
    add_styled_paragraph(
        doc, element["content"],
        italic=True,
        alignment=align,
        space_after=Pt(6),
    )


def render_degree_title(doc, element):
    """Render a dual-language degree title."""
    add_styled_paragraph(
        doc, element["original"],
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        space_after=Pt(0),
    )
    add_styled_paragraph(
        doc, f'({element["english"]})',
        italic=True,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        space_after=Pt(6),
    )


# Dispatcher
ELEMENT_RENDERERS = {
    "institution_header": render_institution_header,
    "heading": render_heading,
    "paragraph": render_paragraph,
    "info_block": render_info_block,
    "table": render_table,
    "summary_block": render_summary_block,
    "grading_scale": render_grading_scale,
    "signature_block": render_signature_block,
    "separator": render_separator,
    "footer_text": render_footer_text,
    "placeholder": render_placeholder,
    "degree_title": render_degree_title,
}


# ---------------------------------------------------------------------------
# Certification Block
# ---------------------------------------------------------------------------

def render_certification_block(doc, metadata):
    """Append the USCIS certification statement at the end of the document."""
    add_styled_paragraph(doc, "", space_after=Pt(12))  # spacer

    render_separator(doc, {})

    add_styled_paragraph(
        doc, "CERTIFICATION OF TRANSLATION ACCURACY",
        size=FONT_SIZE_H2, bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=Pt(8), space_before=Pt(8),
    )

    source_lang = metadata.get("source_language_name", "the source language")
    subject = metadata.get("subject_name", "the subject")
    doc_title = metadata.get("document_title", "the document")

    cert_text = (
        f"I, the undersigned, hereby certify that I am competent to translate "
        f"from {source_lang} into English, and that the above translation of "
        f"the {doc_title} pertaining to {subject} is true and accurate to the "
        f"best of my abilities."
    )

    add_styled_paragraph(
        doc, cert_text,
        size=FONT_SIZE_BODY,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        space_after=Pt(16),
    )

    # Signature area
    sig_table = doc.add_table(rows=4, cols=2)
    sig_table.autofit = False

    labels = ["Signature:", "Printed Name:", "Date:", ""]
    for row_idx, label in enumerate(labels):
        label_cell = sig_table.rows[row_idx].cells[0]
        value_cell = sig_table.rows[row_idx].cells[1]
        set_no_borders(label_cell)
        set_no_borders(value_cell)
        if label:
            add_cell_text(label_cell, label, bold=True)
            add_cell_text(value_cell, "____________________________")

    set_table_col_widths(sig_table, [25, 75])


# ---------------------------------------------------------------------------
# Document Assembly
# ---------------------------------------------------------------------------

def validate_json(data):
    """Validate translation JSON against the schema. Returns errors list."""
    try:
        import jsonschema
        schema_path = Path(__file__).parent / "translation-output-schema.json"
        if schema_path.exists():
            with open(schema_path, "r", encoding="utf-8") as f:
                schema = json.load(f)
            jsonschema.validate(data, schema)
            return []
        else:
            print(f"Warning: Schema file not found at {schema_path}, skipping validation.",
                  file=sys.stderr)
            return []
    except ImportError:
        print("Warning: jsonschema not installed, skipping validation.", file=sys.stderr)
        return []
    except jsonschema.ValidationError as e:
        return [str(e.message)]


def render_document(data, template_path=None, output_path=None):
    """
    Render the translation JSON into a .docx file.

    Args:
        data: Parsed translation JSON (dict).
        template_path: Optional path to a .docx template with letterhead.
        output_path: Output .docx path.

    Returns:
        Path to the generated .docx file.
    """
    # Validate
    errors = validate_json(data)
    if errors:
        print("Schema validation errors:", file=sys.stderr)
        for err in errors:
            print(f"  - {err}", file=sys.stderr)
        print("Proceeding anyway — output may be malformed.", file=sys.stderr)

    # Load template or create blank document
    if template_path and Path(template_path).exists():
        doc = Document(template_path)
    else:
        doc = Document()
        if template_path:
            print(f"Warning: Template not found at {template_path}, using blank document.",
                  file=sys.stderr)

    # Set A4 page size on all sections
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.left_margin = Mm(25)
        section.right_margin = Mm(25)
        section.top_margin = Mm(25)
        section.bottom_margin = Mm(25)

    metadata = data.get("metadata", {})
    pages = data.get("pages", [])

    for page_idx, page in enumerate(pages):
        # Page break before every page except the first
        if page_idx > 0:
            para = doc.add_paragraph()
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.space_before = Pt(0)
            run = para.add_run()
            from docx.enum.text import WD_BREAK
            run.add_break(WD_BREAK.PAGE)

        # Continuation header (for split tables)
        continuation = page.get("continuation_header")
        if continuation:
            add_styled_paragraph(
                doc, continuation,
                size=FONT_SIZE_SMALL, italic=True,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                space_after=Pt(6),
            )

        # Render each element
        for element in page.get("elements", []):
            elem_type = element.get("type")
            renderer = ELEMENT_RENDERERS.get(elem_type)
            if renderer:
                renderer(doc, element)
            else:
                print(f"Warning: Unknown element type '{elem_type}', skipping.",
                      file=sys.stderr)

    # Append certification block
    render_certification_block(doc, metadata)

    # Determine output path
    if not output_path:
        subject = metadata.get("subject_name", "translation")
        safe_name = "".join(c if c.isalnum() or c in " -_" else "_" for c in subject)
        output_path = f"Certified_Translation_{safe_name}.docx"

    doc.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Render USCIS certified translation JSON to .docx"
    )
    parser.add_argument(
        "input",
        help="Path to translation JSON file",
    )
    parser.add_argument(
        "--template", "-t",
        help="Path to .docx template with letterhead (optional)",
        default=None,
    )
    parser.add_argument(
        "--output", "-o",
        help="Output .docx file path (optional, auto-generated if omitted)",
        default=None,
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Only validate the JSON without rendering",
    )

    args = parser.parse_args()

    # Load JSON
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    with open(input_path, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
        except json.JSONDecodeError as e:
            print(f"Error: Invalid JSON: {e}", file=sys.stderr)
            sys.exit(1)

    if args.validate_only:
        errors = validate_json(data)
        if errors:
            print("Validation FAILED:")
            for err in errors:
                print(f"  - {err}")
            sys.exit(1)
        else:
            print("Validation PASSED.")
            sys.exit(0)

    output = render_document(data, template_path=args.template, output_path=args.output)
    print(f"Generated: {output}")


if __name__ == "__main__":
    main()
