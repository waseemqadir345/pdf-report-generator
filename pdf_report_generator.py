"""
PDF Report Generator
====================
Upload any Excel/CSV file → Get a professional PDF report automatically.

Install required libraries first:
    pip install pandas reportlab openpyxl matplotlib

Usage:
    python pdf_report_generator.py
"""

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, HRFlowable
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from datetime import datetime
import sys
import os

# ─────────────────────────────────────────────
#  SETTINGS — client ke hisab se badlo
# ─────────────────────────────────────────────
COMPANY_NAME    = "Acme Corporation"
REPORT_TITLE    = "Sales Report"
PRIMARY_COLOR   = colors.HexColor("#1e3a6e")   # dark blue
ACCENT_COLOR    = colors.HexColor("#2d5aa0")   # medium blue
INPUT_FILE      = "data.xlsx"                  # ya "data.csv"
OUTPUT_PDF      = "report_output.pdf"
# ─────────────────────────────────────────────


def load_data(filepath):
    """Excel ya CSV dono load karta hai."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(filepath)
    elif ext in [".xlsx", ".xls"]:
        df = pd.read_excel(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}")
    return df


def make_header(company, title, styles):
    """Report ka top header banata hai."""
    date_str = datetime.now().strftime("%B %d, %Y")
    elems = []

    # Company name
    header_style = ParagraphStyle(
        "header",
        parent=styles["Normal"],
        fontSize=20,
        textColor=PRIMARY_COLOR,
        fontName="Helvetica-Bold",
        spaceAfter=4,
    )
    sub_style = ParagraphStyle(
        "sub",
        parent=styles["Normal"],
        fontSize=10,
        textColor=colors.HexColor("#8892a4"),
        spaceAfter=2,
    )
    date_style = ParagraphStyle(
        "date",
        parent=styles["Normal"],
        fontSize=9,
        textColor=colors.grey,
        alignment=TA_RIGHT,
    )

    elems.append(Paragraph(company, header_style))
    elems.append(Paragraph(title, sub_style))
    elems.append(Paragraph(f"Generated: {date_str}", date_style))
    elems.append(HRFlowable(width="100%", thickness=2, color=PRIMARY_COLOR, spaceAfter=12))
    return elems


def make_summary_table(df, styles):
    """Top summary stats — total rows, numeric columns ka sum/avg."""
    elems = []

    sec_style = ParagraphStyle(
        "section",
        parent=styles["Normal"],
        fontSize=11,
        textColor=PRIMARY_COLOR,
        fontName="Helvetica-Bold",
        spaceBefore=8,
        spaceAfter=6,
    )
    elems.append(Paragraph("Summary Overview", sec_style))

    # Sirf numeric columns
    num_cols = df.select_dtypes(include="number").columns.tolist()

    summary_data = [["Metric", "Value"]]
    summary_data.append(["Total Records", str(len(df))])
    summary_data.append(["Total Columns", str(len(df.columns))])

    for col in num_cols[:4]:          # max 4 numeric cols dikhao
        total = df[col].sum()
        avg   = df[col].mean()
        summary_data.append([f"{col} — Total", f"{total:,.2f}"])
        summary_data.append([f"{col} — Average", f"{avg:,.2f}"])

    t = Table(summary_data, colWidths=[9 * cm, 8 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",  (0, 0), (-1, 0), PRIMARY_COLOR),
        ("TEXTCOLOR",   (0, 0), (-1, 0), colors.white),
        ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, 0), 10),
        ("ALIGN",       (0, 0), (-1, -1), "LEFT"),
        ("FONTSIZE",    (0, 1), (-1, -1), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.HexColor("#f8f9fc"), colors.white]),
        ("GRID",        (0, 0), (-1, -1), 0.5, colors.HexColor("#e8ecf4")),
        ("TOPPADDING",  (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))
    elems.append(t)
    return elems


def make_data_table(df, styles, max_rows=20):
    """Actual data ka table — first max_rows rows."""
    elems = []

    sec_style = ParagraphStyle(
        "section2",
        parent=styles["Normal"],
        fontSize=11,
        textColor=PRIMARY_COLOR,
        fontName="Helvetica-Bold",
        spaceBefore=16,
        spaceAfter=6,
    )
    elems.append(Paragraph(f"Data Table (first {min(max_rows, len(df))} rows)", sec_style))

    # Headers + data
    headers = list(df.columns)
    rows    = df.head(max_rows).astype(str).values.tolist()
    table_data = [headers] + rows

    # Column widths — page width / number of columns
    page_w    = A4[0] - 4 * cm      # margins hata ke
    col_count = len(headers)
    col_w     = page_w / col_count

    t = Table(table_data, colWidths=[col_w] * col_count, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, 0), PRIMARY_COLOR),
        ("TEXTCOLOR",    (0, 0), (-1, 0), colors.white),
        ("FONTNAME",     (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, -1), 8),
        ("ALIGN",        (0, 0), (-1, -1), "LEFT"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.HexColor("#f8f9fc"), colors.white]),
        ("GRID",         (0, 0), (-1, -1), 0.4, colors.HexColor("#e0e4ee")),
        ("TOPPADDING",   (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",  (0, 0), (-1, -1), 6),
        ("OVERFLOW",     (0, 0), (-1, -1), "shrink"),
    ]))
    elems.append(t)
    return elems


def make_footer_note(styles):
    """Page ke neeche ek note."""
    note_style = ParagraphStyle(
        "note",
        parent=styles["Normal"],
        fontSize=8,
        textColor=colors.grey,
        spaceBefore=20,
    )
    return [
        HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#e0e4ee")),
        Paragraph(
            f"This report was automatically generated on "
            f"{datetime.now().strftime('%Y-%m-%d %H:%M')} | {COMPANY_NAME}",
            note_style,
        )
    ]


def generate_report(input_file, output_file):
    """Main function — sab kuch chalata hai."""
    print(f"📂 Loading data from: {input_file}")
    df = load_data(input_file)
    print(f"✅ Loaded {len(df)} rows and {len(df.columns)} columns")

    doc    = SimpleDocTemplate(
        output_file,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
    )
    styles = getSampleStyleSheet()

    # Sab elements ek list mein
    story = []
    story += make_header(COMPANY_NAME, REPORT_TITLE, styles)
    story.append(Spacer(1, 0.3 * cm))
    story += make_summary_table(df, styles)
    story.append(Spacer(1, 0.3 * cm))
    story += make_data_table(df, styles)
    story += make_footer_note(styles)

    doc.build(story)
    print(f"\n✅ PDF report saved: {output_file}")
    print(f"📄 Open '{output_file}' to see the result!")


# ─────────────────────────────────────────────
if __name__ == "__main__":
    # Command line se file name de sakte ho:
    # python pdf_report_generator.py mydata.xlsx myreport.pdf
    inp = sys.argv[1] if len(sys.argv) > 1 else INPUT_FILE
    out = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_PDF

    if not os.path.exists(inp):
        print(f"❌ File not found: {inp}")
        print("   Ek Excel ya CSV file banao aur uska naam upar INPUT_FILE mein likho.")
        sys.exit(1)

    generate_report(inp, out)
