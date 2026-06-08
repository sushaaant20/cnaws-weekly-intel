from io import BytesIO
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.platypus import Paragraph, Table, TableStyle

PAGE_WIDTH, PAGE_HEIGHT = A4
PAGE_MARGIN = 24
HEADER_HEIGHT = 78
HEADER_INFO_HEIGHT = 36
FOOTER_HEIGHT = 24
COLUMN_GAP = 14

COLOR_HEADER = colors.HexColor("#0B1F2A")
COLOR_ACCENT = colors.HexColor("#A61E22")
COLOR_TEXT = colors.HexColor("#1A1A1A")
COLOR_TEXT_SECONDARY = colors.HexColor("#666666")
COLOR_BORDER = colors.HexColor("#D9E2EC")
COLOR_PANEL_BG = colors.HexColor("#FFFFFF")
COLOR_PANEL_TINT = colors.HexColor("#F8FAFC")
COLOR_GRID = colors.HexColor("#E5E7EB")

LOGO_DIR = Path(__file__).resolve().parent / "logo"
FONT_SEARCH_DIR = Path("C:/Windows/Fonts")


def _register_font(alias, filenames):
    for filename in filenames:
        font_path = FONT_SEARCH_DIR / filename
        if font_path.exists():
            pdfmetrics.registerFont(TTFont(alias, str(font_path)))
            return alias
    return None


MONTSERRAT_BOLD = _register_font("Montserrat-Bold", ["Montserrat-Bold.ttf", "montserrat-bold.ttf"])
MONTSERRAT_SEMIBOLD = _register_font(
    "Montserrat-SemiBold",
    ["Montserrat-SemiBold.ttf", "montserrat-semibold.ttf"],
)
INTER_REGULAR = _register_font(
    "Inter-Regular",
    ["Inter-Regular.ttf", "Inter_18pt-Regular.ttf", "Inter.ttf"],
)
INTER_SEMIBOLD = _register_font(
    "Inter-SemiBold",
    ["Inter-SemiBold.ttf", "Inter_18pt-SemiBold.ttf"],
)

FONT_TITLE = MONTSERRAT_BOLD or "Helvetica-Bold"
FONT_HEADING = MONTSERRAT_SEMIBOLD or "Helvetica-Bold"
FONT_BODY = INTER_REGULAR or "Helvetica"
FONT_BODY_BOLD = INTER_SEMIBOLD or "Helvetica-Bold"


def _report_title(report_context):
    return report_context.get("report_title", "CNAWS Intelligence Dashboard")


def _locate_logo():
    for extension in ("*.png", "*.jpg", "*.jpeg", "*.webp"):
        matches = sorted(LOGO_DIR.glob(extension))
        if matches:
            return matches[0]
    return None


def _format_number(value, decimals=0):
    if value is None:
        return "-"
    if decimals == 0:
        return f"{int(round(value)):,}"
    return f"{value:,.{decimals}f}"


def _format_pct(value):
    if value is None:
        return "New"
    return f"{value:+.1%}"


def _format_reporting_period(label):
    return str(label).replace(" to ", " - ").upper()


def _styles():
    base = getSampleStyleSheet()
    return {
        "body": ParagraphStyle(
            "Body",
            parent=base["BodyText"],
            fontName=FONT_BODY,
            fontSize=8.4,
            leading=11.2,
            textColor=COLOR_TEXT,
            spaceAfter=0,
        ),
        "small": ParagraphStyle(
            "Small",
            parent=base["BodyText"],
            fontName=FONT_BODY,
            fontSize=7.4,
            leading=9.4,
            textColor=COLOR_TEXT_SECONDARY,
            spaceAfter=0,
        ),
        "section": ParagraphStyle(
            "Section",
            parent=base["Heading2"],
            fontName=FONT_HEADING,
            fontSize=10.2,
            leading=12,
            textColor=COLOR_HEADER,
            spaceAfter=0,
            uppercase=True,
        ),
        "title": ParagraphStyle(
            "Title",
            parent=base["Heading1"],
            fontName=FONT_TITLE,
            fontSize=13.5,
            leading=15.5,
            textColor=colors.white,
            spaceAfter=0,
        ),
    }


def _draw_panel(canvas, x, y_top, width, height, fill=COLOR_PANEL_BG):
    canvas.saveState()
    canvas.setFillColor(fill)
    canvas.setStrokeColor(COLOR_BORDER)
    canvas.setLineWidth(0.8)
    canvas.roundRect(x, y_top - height, width, height, 12, stroke=1, fill=1)
    canvas.restoreState()


def _draw_section_heading(canvas, text, x, y_top, width, styles):
    heading = Paragraph(text, styles["section"])
    _, height = heading.wrap(width, 20)
    heading.drawOn(canvas, x, y_top - height)
    return y_top - height - 4


def _draw_paragraphs(canvas, lines, x, y_top, width, styles, bullet=False):
    y_cursor = y_top
    for line in lines:
        prefix = "&bull; " if bullet else ""
        paragraph = Paragraph(prefix + str(line), styles["body"])
        _, height = paragraph.wrap(width, y_cursor)
        paragraph.drawOn(canvas, x, y_cursor - height)
        y_cursor -= height + 3
    return y_cursor


def _draw_table(canvas, rows, x, y_top, width, height, column_widths, font_size=7.1):
    table = Table(rows, colWidths=[width * fraction for fraction in column_widths], repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), COLOR_HEADER),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), FONT_BODY_BOLD),
                ("FONTNAME", (0, 1), (-1, -1), FONT_BODY),
                ("FONTSIZE", (0, 0), (-1, -1), font_size),
                ("LEADING", (0, 0), (-1, -1), font_size + 1.8),
                ("GRID", (0, 0), (-1, -1), 0.45, COLOR_BORDER),
                ("BACKGROUND", (0, 1), (-1, -1), COLOR_PANEL_BG),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    _, wrapped_height = table.wrap(width, height)
    table.drawOn(canvas, x, y_top - wrapped_height)
    return y_top - wrapped_height


def _image_reader(image_bytes):
    if not image_bytes:
        return None
    return ImageReader(BytesIO(image_bytes))


def _draw_image(canvas, image_bytes, x, y_top, width, height, preserve=True):
    reader = _image_reader(image_bytes)
    if reader is None:
        return y_top

    image = PILImage.open(BytesIO(image_bytes))
    img_width, img_height = image.size
    draw_width = width
    draw_height = height
    if preserve and img_width and img_height:
        ratio = min(width / img_width, height / img_height)
        draw_width = img_width * ratio
        draw_height = img_height * ratio
    x_offset = x + (width - draw_width) / 2
    y_offset = y_top - draw_height - (height - draw_height) / 2
    canvas.drawImage(reader, x_offset, y_offset, draw_width, draw_height, preserveAspectRatio=True, mask="auto")
    return y_top - height


def _build_trend_chart_image(title, labels, previous_values, current_values, previous_color, current_color):
    figure, axis = plt.subplots(figsize=(4.6, 1.6), dpi=220)
    figure.patch.set_facecolor("white")
    axis.set_facecolor("white")
    axis.plot(
        labels,
        previous_values,
        color=previous_color,
        linewidth=3,
        linestyle="--",
        marker="o",
        markersize=4.2,
    )
    axis.plot(
        labels,
        current_values,
        color=current_color,
        linewidth=3,
        marker="o",
        markersize=4.2,
    )
    axis.set_title(title, loc="left", fontsize=9.5, color="#0B1F2A", fontweight="bold")
    axis.grid(axis="y", color="#E5E7EB", linewidth=0.9)
    axis.spines["top"].set_visible(False)
    axis.spines["right"].set_visible(False)
    axis.spines["left"].set_visible(False)
    axis.spines["bottom"].set_visible(False)
    axis.tick_params(axis="x", labelsize=7, colors="#666666", length=0)
    axis.tick_params(axis="y", labelsize=7, colors="#666666", length=0)
    axis.margins(x=0.03)
    plt.tight_layout(pad=0.7)
    buffer = BytesIO()
    figure.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(figure)
    return buffer.getvalue()


def _kpi_payload(report_context):
    summary = report_context["summary_metrics"]
    expansion = report_context["expansion_analysis"]
    return [
        ("Incidents", summary["incidents"]["current"], summary["incidents"]["delta"], f"Prev: {_format_number(summary['incidents']['previous'])}"),
        ("Casualties", summary["casualties"]["current"], summary["casualties"]["delta"], f"Prev: {_format_number(summary['casualties']['previous'])}"),
        ("Intensity", summary["intensity"]["current"], summary["intensity"]["delta"], f"Prev: {_format_number(summary['intensity']['previous'], 2)}"),
        ("Districts", summary["districts"]["current"], summary["districts"]["delta"], f"Prev: {_format_number(summary['districts']['previous'])}"),
        ("New Districts", expansion["new_districts"], expansion["new_districts"], f"Prev footprint: {expansion['previous_footprint']}"),
        ("Expansion Index", expansion["expansion_index"], expansion["expansion_index"], "Current spread rate"),
    ]


def _draw_kpi_cards(canvas, report_context, x, y_top, width, height):
    cards = _kpi_payload(report_context)
    card_gap = 8
    card_width = (width - (card_gap * 2)) / 3
    card_height = (height - card_gap) / 2
    for index, (label, value, delta, meta) in enumerate(cards):
        row = index // 3
        col = index % 3
        card_x = x + col * (card_width + card_gap)
        card_y = y_top - row * (card_height + card_gap)
        canvas.setFillColor(COLOR_PANEL_TINT)
        canvas.setStrokeColor(COLOR_BORDER)
        canvas.roundRect(card_x, card_y - card_height, card_width, card_height, 10, stroke=1, fill=1)
        canvas.setFont(FONT_HEADING, 7.2)
        canvas.setFillColor(COLOR_TEXT_SECONDARY)
        canvas.drawString(card_x + 8, card_y - 12, label.upper())
        canvas.setFont(FONT_TITLE, 15)
        canvas.setFillColor(COLOR_TEXT)
        decimals = 2 if isinstance(value, float) and not float(value).is_integer() else 0
        canvas.drawString(card_x + 8, card_y - 33, _format_number(value, decimals))
        canvas.setFont(FONT_BODY_BOLD, 7.2)
        delta_color = colors.HexColor("#0F766E") if delta > 0 else colors.HexColor("#A61E22") if delta < 0 else COLOR_TEXT_SECONDARY
        canvas.setFillColor(delta_color)
        delta_text = f"{delta:+.2f}" if isinstance(delta, float) and not float(delta).is_integer() else f"{delta:+,}"
        canvas.drawString(card_x + 8, card_y - 47, delta_text)
        canvas.setFont(FONT_BODY, 6.6)
        canvas.setFillColor(COLOR_TEXT_SECONDARY)
        canvas.drawString(card_x + 8, card_y - 58, meta)


def _draw_header_footer(canvas, report_context, page_number, total_pages, logo_path):
    canvas.saveState()
    canvas.setFillColor(COLOR_HEADER)
    canvas.rect(0, PAGE_HEIGHT - HEADER_HEIGHT, PAGE_WIDTH, HEADER_HEIGHT, stroke=0, fill=1)

    canvas.setFillColor(colors.white)
    canvas.setFont(FONT_HEADING, 8)
    canvas.drawString(PAGE_MARGIN, PAGE_HEIGHT - 18, "CENTRE FOR NEW AGE WARFARE STUDIES")
    canvas.setFont(FONT_TITLE, 15)
    brief_label = f"{report_context['period_label'].upper()} INTELLIGENCE BRIEF"
    canvas.drawString(PAGE_MARGIN, PAGE_HEIGHT - 38, brief_label)

    if logo_path and logo_path.exists():
        logo_reader = ImageReader(str(logo_path))
        canvas.drawImage(
            logo_reader,
            PAGE_WIDTH - PAGE_MARGIN - 64,
            PAGE_HEIGHT - 56,
            64,
            40,
            preserveAspectRatio=True,
            mask="auto",
        )

    info_y = PAGE_HEIGHT - HEADER_HEIGHT - 10
    canvas.setFillColor(COLOR_TEXT)
    canvas.setFont(FONT_HEADING, 8.2)
    canvas.drawString(PAGE_MARGIN, info_y, "PAKISTAN MILITANCY MONITORING")
    canvas.setFont(FONT_BODY_BOLD, 7.6)
    canvas.drawString(PAGE_MARGIN, info_y - 11, "CNAWS")
    canvas.setFont(FONT_BODY, 7.6)
    canvas.setFillColor(COLOR_TEXT_SECONDARY)
    canvas.drawRightString(
        PAGE_WIDTH - PAGE_MARGIN,
        info_y,
        _format_reporting_period(report_context["reporting_period_label"]),
    )
    canvas.drawRightString(PAGE_WIDTH - PAGE_MARGIN, info_y - 11, "REPORTING PERIOD")

    footer_y = FOOTER_HEIGHT + 6
    canvas.setStrokeColor(COLOR_BORDER)
    canvas.setLineWidth(0.7)
    canvas.line(PAGE_MARGIN, footer_y + 8, PAGE_WIDTH - PAGE_MARGIN, footer_y + 8)
    canvas.setFont(FONT_BODY, 7.2)
    canvas.setFillColor(COLOR_TEXT_SECONDARY)
    canvas.drawString(PAGE_MARGIN, footer_y - 1, "contact@cnaws.in")
    canvas.drawCentredString(PAGE_WIDTH / 2, footer_y - 1, f"Page {page_number} of {total_pages}")
    canvas.drawRightString(PAGE_WIDTH - PAGE_MARGIN, footer_y - 1, "cnaws.in")
    canvas.restoreState()


def _district_table_rows(report_context, limit=6):
    frame = report_context["district_breakdown"].head(limit).copy()
    rows = [["District", "Cur", "Prev", "Share", "Leth."]]
    for _, row in frame.iterrows():
        rows.append(
            [
                str(row["district"]),
                _format_number(row["current"]),
                _format_number(row["previous"]),
                f"{row['share'] * 100:.1f}%",
                f"{row['lethality']:.2f}",
            ]
        )
    return rows


def _event_table_rows(report_context, limit=7):
    frame = report_context["event_breakdown"].head(limit).copy()
    rows = [["Event", "Cur", "Prev", "Delta", "%"]]
    for _, row in frame.iterrows():
        rows.append(
            [
                str(row["event_type"]),
                _format_number(row["current"]),
                _format_number(row["previous"]),
                _format_number(row["delta"]),
                _format_pct(row["pct_change"]),
            ]
        )
    return rows


def _expansion_rows(report_context, limit=6):
    frame = report_context["expansion_analysis"]["table"].head(limit).copy()
    rows = [["District", "Incidents", "Casualties"]]
    if frame.empty:
        rows.append(["No new districts", "-", "-"])
        return rows
    for _, row in frame.iterrows():
        rows.append(
            [str(row["district"]), _format_number(row["incidents"]), _format_number(row["casualties"])]
        )
    return rows


def _incident_rows(report_context, limit=8):
    frame = report_context["incident_table"].head(limit).copy()
    rows = [["Date", "District", "Event", "Cas."]]
    if frame.empty:
        rows.append(["No mapped incidents", "-", "-", "-"])
        return rows
    for _, row in frame.iterrows():
        rows.append(
            [
                pd.to_datetime(row["date"]).strftime("%d %b"),
                str(row["district"]),
                str(row["event_type"]),
                _format_number(row["casualties"]),
            ]
        )
    return rows


def _high_impact_rows(report_context, limit=8):
    frame = report_context["high_impact_incidents"].head(limit).copy()
    rows = [["Date", "District", "Event", "Cas."]]
    if frame.empty:
        rows.append(["No high impact incidents", "-", "-", "-"])
        return rows
    for _, row in frame.iterrows():
        rows.append(
            [
                pd.to_datetime(row["date"]).strftime("%d %b"),
                str(row["district"]),
                str(row["event_type"]),
                _format_number(row["casualties_total"]),
            ]
        )
    return rows


def _draw_page_one(canvas, report_context, styles, logo_path):
    _draw_header_footer(canvas, report_context, 1, 2, logo_path)
    top = PAGE_HEIGHT - HEADER_HEIGHT - HEADER_INFO_HEIGHT - 8
    left_x = PAGE_MARGIN
    right_x = PAGE_WIDTH - PAGE_MARGIN - 190
    left_width = right_x - left_x - COLUMN_GAP
    right_width = PAGE_WIDTH - PAGE_MARGIN - right_x

    exec_height = 126
    _draw_panel(canvas, left_x, top, left_width, exec_height)
    y_cursor = _draw_section_heading(canvas, "Executive Assessment", left_x + 10, top - 8, left_width - 20, styles)
    exec_lines = [line for line in report_context["executive_summary"].splitlines() if line.strip()][1:]
    _draw_paragraphs(canvas, exec_lines[:5], left_x + 10, y_cursor - 2, left_width - 20, styles)
    canvas.setFont(FONT_BODY_BOLD, 8)
    canvas.setFillColor(COLOR_ACCENT)
    canvas.drawRightString(
        left_x + left_width - 10,
        top - 18,
        f"INTELLIGENCE SCORE {report_context['intelligence_score']['final_score']} / 100",
    )

    takeaways_top = top - exec_height - 10
    takeaways_height = 120
    _draw_panel(canvas, left_x, takeaways_top, left_width, takeaways_height)
    y_cursor = _draw_section_heading(canvas, "Key Takeaways", left_x + 10, takeaways_top - 8, left_width - 20, styles)
    _draw_paragraphs(canvas, report_context["key_takeaways"][:4], left_x + 10, y_cursor - 2, left_width - 20, styles, bullet=True)

    kpi_top = takeaways_top - takeaways_height - 10
    kpi_height = 138
    _draw_panel(canvas, left_x, kpi_top, left_width, kpi_height)
    _draw_section_heading(canvas, "Key Metrics", left_x + 10, kpi_top - 8, left_width - 20, styles)
    _draw_kpi_cards(canvas, report_context, left_x + 10, kpi_top - 22, left_width - 20, kpi_height - 34)

    trend_top = kpi_top - kpi_height - 10
    trend_height = 190
    _draw_panel(canvas, left_x, trend_top, left_width, trend_height)
    y_cursor = _draw_section_heading(canvas, "Operational Activity Trends", left_x + 10, trend_top - 8, left_width - 20, styles)
    trend = report_context["trend_payload"]
    militant_chart = _build_trend_chart_image(
        "Militant Activity",
        trend["labels"],
        trend["militant_previous"],
        trend["militant_current"],
        "#94A3B8",
        "#1F3B57",
    )
    security_chart = _build_trend_chart_image(
        "Security Operations",
        trend["labels"],
        trend["security_previous"],
        trend["security_current"],
        "#FCA5A5",
        "#A61E22",
    )
    chart_height = 66
    _draw_image(canvas, militant_chart, left_x + 8, y_cursor - 2, left_width - 16, chart_height)
    _draw_image(canvas, security_chart, left_x + 8, y_cursor - chart_height - 12, left_width - 16, chart_height)

    table_height = 170
    _draw_panel(canvas, right_x, top, right_width, table_height)
    y_cursor = _draw_section_heading(canvas, "Analytical Table", right_x + 10, top - 8, right_width - 20, styles)
    _draw_table(
        canvas,
        _district_table_rows(report_context, 6),
        right_x + 8,
        y_cursor - 2,
        right_width - 16,
        table_height - 24,
        [0.42, 0.12, 0.12, 0.16, 0.18],
    )

    tactical_top = top - table_height - 10
    tactical_height = 128
    _draw_panel(canvas, right_x, tactical_top, right_width, tactical_height)
    y_cursor = _draw_section_heading(canvas, "Tactical Shift Panel", right_x + 10, tactical_top - 8, right_width - 20, styles)
    tactical = report_context["tactical_shift"]
    tactical_lines = []
    if tactical["top_increase"] is not None:
        tactical_lines.append(
            f"Top Increase: {tactical['top_increase']['event_type']} ({_format_pct(tactical['top_increase']['pct_change'])})"
        )
    if tactical["top_decrease"] is not None:
        tactical_lines.append(
            f"Top Decrease: {tactical['top_decrease']['event_type']} ({_format_pct(tactical['top_decrease']['pct_change'])})"
        )
    tactical_lines.extend(tactical["observations"][:3])
    _draw_paragraphs(canvas, tactical_lines, right_x + 10, y_cursor - 2, right_width - 20, styles)

    map_top = tactical_top - tactical_height - 10
    map_height = 320
    _draw_panel(canvas, right_x, map_top, right_width, map_height)
    _draw_section_heading(canvas, "Geographic View", right_x + 10, map_top - 8, right_width - 20, styles)
    _draw_image(canvas, report_context["map_image"], right_x + 8, map_top - 24, right_width - 16, map_height - 32)


def _draw_page_two(canvas, report_context, styles, logo_path):
    _draw_header_footer(canvas, report_context, 2, 2, logo_path)
    top = PAGE_HEIGHT - HEADER_HEIGHT - HEADER_INFO_HEIGHT - 8
    left_x = PAGE_MARGIN
    right_x = PAGE_WIDTH - PAGE_MARGIN - 208
    left_width = right_x - left_x - COLUMN_GAP
    right_width = PAGE_WIDTH - PAGE_MARGIN - right_x

    expansion_top = top
    expansion_height = 144
    _draw_panel(canvas, left_x, expansion_top, left_width, expansion_height)
    y_cursor = _draw_section_heading(canvas, "Expansion Panel", left_x + 10, expansion_top - 8, left_width - 20, styles)
    expansion = report_context["expansion_analysis"]
    expansion_text = (
        f"New Districts: {expansion['new_districts']} | Expansion Index: {expansion['expansion_index']:.2f} | {expansion['tag']}"
    )
    paragraph = Paragraph(expansion_text, styles["body"])
    _, para_height = paragraph.wrap(left_width - 20, 30)
    paragraph.drawOn(canvas, left_x + 10, y_cursor - para_height)
    _draw_table(
        canvas,
        _expansion_rows(report_context, 5),
        left_x + 8,
        expansion_top - 48,
        left_width - 16,
        78,
        [0.52, 0.23, 0.25],
    )

    district_top = expansion_top - expansion_height - 10
    district_height = 188
    _draw_panel(canvas, left_x, district_top, left_width, district_height)
    y_cursor = _draw_section_heading(canvas, "District Intelligence", left_x + 10, district_top - 8, left_width - 20, styles)
    _draw_table(
        canvas,
        _district_table_rows(report_context, 8),
        left_x + 8,
        y_cursor - 2,
        left_width - 16,
        district_height - 24,
        [0.40, 0.12, 0.12, 0.18, 0.18],
    )

    geo_top = district_top - district_height - 10
    geo_height = 244
    _draw_panel(canvas, left_x, geo_top, left_width, geo_height)
    _draw_section_heading(canvas, "Geographic View", left_x + 10, geo_top - 8, left_width - 20, styles)
    _draw_image(canvas, report_context["map_image"], left_x + 10, geo_top - 24, left_width - 20, 92)
    y_cursor = geo_top - 124
    _draw_section_heading(canvas, "Incident Table", left_x + 10, y_cursor, left_width - 20, styles)
    _draw_table(
        canvas,
        _incident_rows(report_context, 7),
        left_x + 8,
        y_cursor - 16,
        left_width - 16,
        112,
        [0.18, 0.30, 0.34, 0.18],
    )

    event_top = top
    event_height = 214
    _draw_panel(canvas, right_x, event_top, right_width, event_height)
    y_cursor = _draw_section_heading(canvas, "Event And Tactical Analysis", right_x + 10, event_top - 8, right_width - 20, styles)
    _draw_table(
        canvas,
        _event_table_rows(report_context, 8),
        right_x + 8,
        y_cursor - 2,
        right_width - 16,
        event_height - 24,
        [0.42, 0.12, 0.12, 0.14, 0.20],
    )

    tactical_top = event_top - event_height - 10
    tactical_height = 140
    _draw_panel(canvas, right_x, tactical_top, right_width, tactical_height)
    y_cursor = _draw_section_heading(canvas, "Tactical Shift Interpretation", right_x + 10, tactical_top - 8, right_width - 20, styles)
    tactical_lines = report_context["tactical_shift"]["observations"][:3] + [
        report_context["tactical_shift"]["interpretation"]
    ]
    _draw_paragraphs(canvas, tactical_lines, right_x + 10, y_cursor - 2, right_width - 20, styles)

    high_top = tactical_top - tactical_height - 10
    high_height = 234
    _draw_panel(canvas, right_x, high_top, right_width, high_height)
    y_cursor = _draw_section_heading(canvas, "High Impact Incidents", right_x + 10, high_top - 8, right_width - 20, styles)
    _draw_table(
        canvas,
        _high_impact_rows(report_context, 8),
        right_x + 8,
        y_cursor - 2,
        right_width - 16,
        high_height - 24,
        [0.18, 0.28, 0.36, 0.18],
    )


def build_pdf_report(report_context):
    buffer = BytesIO()
    canvas = pdf_canvas.Canvas(buffer, pagesize=A4)
    styles = _styles()
    logo_path = _locate_logo()

    _draw_page_one(canvas, report_context, styles, logo_path)
    canvas.showPage()
    _draw_page_two(canvas, report_context, styles, logo_path)
    canvas.save()
    return buffer.getvalue()


def build_docx_report(report_context):
    raise NotImplementedError("DOCX export has been removed from the CNAWS intelligence brief workflow.")
