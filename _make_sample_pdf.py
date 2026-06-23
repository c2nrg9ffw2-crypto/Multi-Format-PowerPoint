"""Run once to generate sample_report.pdf for testing."""
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)

ORANGE = colors.HexColor("#d97757")
DARK   = colors.HexColor("#141413")
LIGHT  = colors.HexColor("#faf9f5")
GRAY   = colors.HexColor("#6b6b6b")

styles = getSampleStyleSheet()

title_style = ParagraphStyle(
    "DocTitle",
    parent=styles["Title"],
    fontSize=18,
    textColor=DARK,
    spaceAfter=6,
)
section_style = ParagraphStyle(
    "Section",
    parent=styles["Heading2"],
    fontSize=13,
    textColor=ORANGE,
    spaceBefore=14,
    spaceAfter=4,
)
body_style = ParagraphStyle(
    "Body",
    parent=styles["Normal"],
    fontSize=10,
    textColor=DARK,
    leading=15,
    spaceAfter=6,
)
bullet_style = ParagraphStyle(
    "Bullet",
    parent=body_style,
    leftIndent=16,
    bulletIndent=4,
    spaceBefore=2,
    spaceAfter=2,
)


def section(title, body_text, bullets=None):
    items = [Paragraph(title, section_style), Paragraph(body_text, body_style)]
    if bullets:
        for b in bullets:
            items.append(Paragraph(b, bullet_style, bulletText="-"))
    items.append(Spacer(1, 6))
    return items


def make_table(headers, rows):
    data = [headers] + rows
    col_count = len(headers)
    col_w = [2.1 * inch] + [0.95 * inch] * (col_count - 1)

    t = Table(data, colWidths=col_w)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), ORANGE),
        ("TEXTCOLOR",  (0, 0), (-1, 0), LIGHT),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, -1), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f4f4")]),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#d4d4d0")),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


story = []

story.append(Paragraph("Annual Business Review 2024", title_style))
story.append(Spacer(1, 4))

story += section(
    "Executive Summary",
    "Fiscal year 2024 was a strong year. Total revenue reached 594,500 USD, "
    "a 21% year-over-year increase driven by product expansion and new enterprise "
    "contracts. Operating margins improved from 28% to 34% through disciplined "
    "cost management.",
    bullets=[
        "Revenue grew 21% YoY to USD 594,500",
        "Operating margin expanded from 28% to 34%",
        "Launched three new product lines in Q2 and Q3",
        "Expanded into two new geographic markets",
        "Headcount grew from 98 to 121",
    ],
)

story += section(
    "Market Overview",
    "The total addressable market expanded to 8.4 billion USD in 2024, up from "
    "6.9 billion in 2023. Two legacy incumbents lost combined market share of "
    "approximately 7 percentage points as customers migrated toward cloud-native "
    "solutions.",
    bullets=[
        "TAM grew 22% to USD 8.4 billion",
        "Cloud-native segment now 41% of total market",
        "Customer acquisition cost decreased 12%",
        "Net Promoter Score increased from 47 to 61",
    ],
)

story.append(PageBreak())

story += section(
    "Product Highlights",
    "Alpha Pro remains our flagship, accounting for 38% of unit sales. "
    "Gamma Basic exceeded launch targets by 40%, capturing the SMB segment. "
    "Delta Cloud entered general availability in Q4.",
    bullets=[
        "Alpha Pro: 4,320 units sold",
        "Gamma Basic: 6,150 units sold",
        "Delta Cloud: GA in Q4, 310 deals in pipeline",
        "Beta Suite: 2,870 units sold",
    ],
)

story += section(
    "Operational Metrics",
    "Infrastructure reliability reached 99.94% uptime, exceeding the 99.9% SLA. "
    "Support ticket volume grew 18% while median resolution time fell from 6.2h to 4.1h.",
)

story.append(make_table(
    ["Metric", "Q1", "Q2-Q3", "Q4"],
    [
        ["Uptime %",            "99.91", "99.95", "99.97"],
        ["Tickets Resolved",    "1,840", "4,210", "2,090"],
        ["Resolution Time (h)", "6.2",   "4.8",   "4.1"],
        ["CSAT Score",          "4.1",   "4.3",   "4.6"],
    ],
))

story.append(Spacer(1, 12))

story += section(
    "Outlook 2025",
    "Key priorities: Delta Cloud rollout, APAC expansion, and ISO 27001 certification. "
    "Revenue guidance for 2025 is USD 720k-760k, representing 21-28% growth.",
    bullets=[
        "Delta Cloud full rollout - Q1 2025",
        "APAC market entry - Q2 2025",
        "ISO 27001 certification - Q3 2025",
        "Revenue guidance: USD 720k-760k",
        "Headcount plan: +18 across Engineering and Sales",
    ],
)

doc = SimpleDocTemplate(
    "sample_report.pdf",
    pagesize=letter,
    leftMargin=0.8 * inch,
    rightMargin=0.8 * inch,
    topMargin=0.8 * inch,
    bottomMargin=0.8 * inch,
)
doc.build(story)
print("Created: sample_report.pdf")
