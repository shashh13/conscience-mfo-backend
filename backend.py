"""
Conscience MFO — PPTX Report Generator Backend
Deploy on Cloud Run, Railway, Render, or any Python host.

Install:  pip install flask pptxgenjs-python pptx
Run locally: python backend.py
POST /generate  →  { reportData: {...} }  →  { pptxBase64: "..." }

If you don't want to deploy a backend, the Apps Script falls back
to a Google Doc summary automatically.
"""

from flask import Flask, request, jsonify
import base64, io, math, json
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml import parse_xml
from lxml import etree
import copy

app = Flask(__name__)

# ── Colours ───────────────────────────────────────────────────────────────────
DARK_BLUE  = RGBColor(0x1A, 0x2A, 0x4A)
GOLD       = RGBColor(0xB8, 0x96, 0x2E)
LIGHT_GOLD = RGBColor(0xD4, 0xAF, 0x5A)
OFF_WHITE  = RGBColor(0xF7, 0xF4, 0xEF)
LIGHT_BG   = RGBColor(0xEE, 0xF3, 0xFA)
DARK_GREY  = RGBColor(0x37, 0x41, 0x51)
MID_GREY   = RGBColor(0x6B, 0x72, 0x80)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_ROW  = RGBColor(0xF3, 0xF4, 0xF6)
BORDER_CLR = RGBColor(0xE5, 0xE7, 0xEB)

SECTOR_PAL = {
    "Technology":  RGBColor(0x25, 0x63, 0xEB),
    "ETF":         RGBColor(0x7C, 0x3A, 0xED),
    "Industrials": RGBColor(0xD9, 0x77, 0x06),
    "Financials":  RGBColor(0x05, 0x96, 0x69),
    "Healthcare":  RGBColor(0xDC, 0x26, 0x26),
    "Energy":      RGBColor(0x08, 0x91, 0xB2),
    "Real Estate": RGBColor(0x93, 0x33, 0xEA),
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def rgb_hex(rgb):
    return "{:02X}{:02X}{:02X}".format(rgb[0], rgb[1], rgb[2])

def add_rect(slide, x, y, w, h, fill_rgb, border_rgb=None, border_pt=0):
    from pptx.util import Pt as Pt2
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if border_rgb and border_pt > 0:
        shape.line.color.rgb = border_rgb
        shape.line.width = Pt2(border_pt)
    else:
        shape.line.fill.background()
    return shape

def add_text_box(slide, x, y, w, h, text, font_size=10, bold=False, color=DARK_GREY,
                 align=PP_ALIGN.LEFT, font_name="Calibri", wrap=True, italic=False):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    txBox.word_wrap = wrap
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox

def add_bullet_text(slide, x, y, w, h, bullets, font_size=8, color=DARK_GREY):
    """Add a text box with bullet points."""
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    txBox.word_wrap = True
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for bullet_text in bullets:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_after = Pt(3)
        # Add bullet character manually
        run = p.add_run()
        run.text = "•  " + bullet_text
        run.font.name = "Calibri"
        run.font.size = Pt(font_size)
        run.font.color.rgb = color

# ── Slide builders ────────────────────────────────────────────────────────────
def build_slide1(prs, data):
    """Allocation table slide."""
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)

    W, H = 13.33, 7.5
    allocations = data["allocations"]
    grand_total = sum(a["alloc"] for a in allocations)
    sorted_allocs = sorted(allocations, key=lambda x: x["alloc"], reverse=True)

    # Background
    add_rect(slide, 0, 0, W, H, RGBColor(0xF8, 0xF9, 0xFB))

    # Header bar
    add_rect(slide, 0, 0, W, 0.72, DARK_BLUE)
    add_text_box(slide, 0.3, 0.08, 3, 0.56, "CONSCIENCE MFO",
                 font_size=13, bold=True, color=GOLD, align=PP_ALIGN.LEFT)
    add_text_box(slide, 3.1, 0.08, 6.5, 0.56,
                 "INVESTMENT REPORT  |  " + data["clientName"].upper(),
                 font_size=10, color=WHITE, align=PP_ALIGN.LEFT)
    add_text_box(slide, 9.8, 0.08, 3.2, 0.56,
                 data["reportDate"] + "  |  STRICTLY PRIVATE & CONFIDENTIAL",
                 font_size=8, color=RGBColor(0xAA, 0xBB, 0xCC), align=PP_ALIGN.RIGHT)

    # Snapshot bar
    add_rect(slide, 0, 0.72, W, 0.38, LIGHT_BG)
    snap = [
        "TOTAL INVESTMENT: USD {:,}".format(int(data["totalInvestment"])),
        "RISK APPETITE: " + data["riskAppetite"],
        "HORIZON: " + data["horizon"],
        "GOAL: " + data["goal"],
    ]
    for i, s in enumerate(snap):
        add_text_box(slide, 0.3 + i*3.2, 0.76, 3.1, 0.3, s,
                     font_size=8, bold=True, color=DARK_BLUE)

    # Section title
    add_rect(slide, 0.3, 1.28, 0.06, 0.26, GOLD)
    add_text_box(slide, 0.44, 1.22, 7, 0.36,
                 "RECOMMENDED PORTFOLIO ALLOCATION",
                 font_size=11, bold=True, color=DARK_BLUE)

    # Table layout
    table_top = 1.65
    row_h = 0.265
    col_w = [0.88, 3.28, 1.08, 0.82]
    left_x = 0.25
    right_x = 6.95
    headers = ["TICKER", "COMPANY", "ALLOC (USD)", "WEIGHT"]

    def draw_header(x):
        cx = x
        for i, h in enumerate(headers):
            add_rect(slide, cx, table_top, col_w[i], row_h, DARK_BLUE)
            add_text_box(slide, cx+0.02, table_top+0.04, col_w[i]-0.04, row_h-0.06,
                         h, font_size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            cx += col_w[i]

    def draw_rows(holdings, start_x):
        for idx, h in enumerate(holdings):
            ry = table_top + row_h + idx * row_h
            sec_color = SECTOR_PAL.get(h["sector"], RGBColor(0x88, 0x88, 0x88))
            row_bg = WHITE if idx % 2 == 0 else LIGHT_ROW
            cx = start_x

            # Sector stripe
            add_rect(slide, cx, ry, 0.05, row_h, sec_color)
            # Ticker cell
            add_rect(slide, cx, ry, col_w[0], row_h, row_bg, BORDER_CLR, 0.3)
            add_text_box(slide, cx+0.07, ry+0.04, col_w[0]-0.08, row_h-0.07,
                         h["ticker"], font_size=8, bold=True, color=DARK_BLUE)
            cx += col_w[0]
            # Name cell
            add_rect(slide, cx, ry, col_w[1], row_h, row_bg, BORDER_CLR, 0.3)
            add_text_box(slide, cx+0.06, ry+0.04, col_w[1]-0.08, row_h-0.07,
                         h["name"], font_size=8, color=DARK_GREY)
            cx += col_w[1]
            # Alloc cell
            add_rect(slide, cx, ry, col_w[2], row_h, row_bg, BORDER_CLR, 0.3)
            add_text_box(slide, cx, ry+0.04, col_w[2], row_h-0.07,
                         "${:,}".format(int(h["alloc"])), font_size=8, bold=True,
                         color=DARK_BLUE, align=PP_ALIGN.CENTER)
            cx += col_w[2]
            # Weight cell
            wt = "{:.1f}%".format(h["alloc"] / grand_total * 100)
            add_rect(slide, cx, ry, col_w[3], row_h, row_bg, BORDER_CLR, 0.3)
            add_text_box(slide, cx, ry+0.04, col_w[3], row_h-0.07,
                         wt, font_size=8, color=MID_GREY, align=PP_ALIGN.CENTER)

    draw_header(left_x)
    draw_header(right_x)
    left_col  = sorted_allocs[:14]
    right_col = sorted_allocs[14:]
    draw_rows(left_col, left_x)
    draw_rows(right_col, right_x)

    # Sector legend
    legend_y = table_top + row_h * 15 + 0.12
    add_text_box(slide, left_x, legend_y, 1.1, 0.22, "SECTOR KEY:",
                 font_size=7, bold=True, color=DARK_GREY)
    for i, (sec, col) in enumerate(SECTOR_PAL.items()):
        lx = left_x + 1.15 + i * 1.7
        add_rect(slide, lx, legend_y+0.04, 0.13, 0.13, col)
        add_text_box(slide, lx+0.17, legend_y, 1.5, 0.22, sec,
                     font_size=7, color=DARK_GREY)

    # Footer
    add_rect(slide, 0, H-0.32, W, 0.32, DARK_BLUE)
    add_text_box(slide, 0, H-0.28, W, 0.24,
                 "Conscience MFO  |  Confidential — For Addressee Only  |  Not Investment Advice",
                 font_size=7, color=RGBColor(0x88, 0xAA, 0xCC), align=PP_ALIGN.CENTER)
    add_text_box(slide, W-0.55, H-0.28, 0.45, 0.24, "1 / 2",
                 font_size=7, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)

    return slide


def build_slide2(prs, data):
    """Charts + analysis slide."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    W, H = 13.33, 7.5
    allocations = data["allocations"]
    grand_total = sum(a["alloc"] for a in allocations)

    # Compute sector totals
    sector_totals = {}
    for a in allocations:
        sector_totals[a["sector"]] = sector_totals.get(a["sector"], 0) + a["alloc"]
    sorted_sectors = sorted(sector_totals.items(), key=lambda x: x[1], reverse=True)

    # Background
    add_rect(slide, 0, 0, W, H, RGBColor(0xF8, 0xF9, 0xFB))

    # Slim header
    add_rect(slide, 0, 0, W, 0.5, DARK_BLUE)
    add_text_box(slide, 0.3, 0.06, 2.5, 0.38, "CONSCIENCE MFO",
                 font_size=11, bold=True, color=GOLD)
    add_text_box(slide, 2.9, 0.06, 6, 0.38,
                 data["clientName"].upper(),
                 font_size=9, color=WHITE)
    add_text_box(slide, 10, 0.06, 3.1, 0.38,
                 "Total AUM: ${:,}  |  {}".format(int(data["totalInvestment"]), data["reportDate"]),
                 font_size=8, color=RGBColor(0xAA, 0xBB, 0xCC), align=PP_ALIGN.RIGHT)

    # Market Context banner
    ctx_y = 0.56
    add_rect(slide, 0, ctx_y, W, 0.82, LIGHT_BG)
    add_rect(slide, 0.22, ctx_y+0.10, 0.05, 0.62, GOLD)
    add_text_box(slide, 0.35, ctx_y+0.05, 3.5, 0.25,
                 "MARKET CONTEXT  |  " + data["reportDate"].upper(),
                 font_size=9.5, bold=True, color=DARK_BLUE)
    add_text_box(slide, 0.35, ctx_y+0.30, 12.7, 0.46,
                 data["marketContext"], font_size=8.5, color=DARK_GREY, wrap=True)

    # Zone boundaries
    content_top = 1.50
    chart_zone_bottom = 5.08
    rat_hdr_y  = 5.08
    rat_body_y = 5.40
    rat_body_h = H - 0.32 - rat_body_y - 0.02

    chart_col_w = 5.7
    chart_col_x = 0.25
    chart_zone_h = chart_zone_bottom - content_top

    # Sector Allocation title
    add_rect(slide, chart_col_x, content_top, 0.05, 0.22, GOLD)
    add_text_box(slide, chart_col_x+0.1, content_top, 3, 0.22,
                 "SECTOR ALLOCATION", font_size=9, bold=True, color=DARK_BLUE)

    # Single large donut chart — full chart zone height
    donut_h = chart_zone_h - 0.28  # full zone minus title
    donut_top = content_top + 0.28
    chart_data = ChartData()
    chart_data.categories = [s for s, v in sorted_sectors]
    chart_data.add_series("Allocation", [v for s, v in sorted_sectors])
    donut = slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT,
        Inches(chart_col_x), Inches(donut_top),
        Inches(chart_col_w), Inches(donut_h),
        chart_data
    ).chart

    donut.has_legend = True
    donut.legend.position = 2  # right
    donut.legend.include_in_layout = False
    donut.series[0].data_labels.show_percentage = True
    donut.series[0].data_labels.number_format = "0%"

    # Color the donut segments
    for i, (sec, _) in enumerate(sorted_sectors):
        pt = donut.series[0].points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = SECTOR_PAL.get(sec, RGBColor(0x88,0x88,0x88))

    # Right column — Diversification blocks
    div_x = chart_col_x + chart_col_w + 0.28
    div_w = W - div_x - 0.22
    div_top = content_top

    add_rect(slide, div_x, div_top, 0.05, 0.22, GOLD)
    add_text_box(slide, div_x+0.1, div_top, div_w, 0.22,
                 "PORTFOLIO DIVERSIFICATION", font_size=9, bold=True, color=DARK_BLUE)

    div_blocks = [
        {
            "title": "CURRENCY DIVERSIFICATION",
            "bullets": data.get("currencyDiversification", [
                "Portfolio denominated primarily in USD with strategic EUR and CAD exposure.",
                "For INR-based investors, USD/EUR assets provide a structural hedge against rupee depreciation.",
                "Multi-currency construction reduces dependence on any single currency pair.",
            ])
        },
        {
            "title": "GEOGRAPHIC DIVERSIFICATION",
            "bullets": data.get("geographicDiversification", [
                "Holdings span six distinct markets across US, Europe, Canada, and Emerging Markets.",
                "EM ETFs provide cross-border exposure without single-country regulatory risk.",
                "Construction ensures returns are not correlated to any single economic cycle.",
            ])
        },
        {
            "title": "SECTOR DIVERSIFICATION",
            "bullets": data.get("sectorDiversification", [
                "Seven distinct sectors: Technology, ETFs, Industrials, Financials, Healthcare, Energy, Real Estate.",
                "Multi-sector allocation mitigates industry-specific drawdowns.",
                "Multiple return drivers remain active across different market regimes.",
            ])
        },
    ]

    div_zone_h = chart_zone_bottom - div_top - 0.26
    title_bar_h = 0.28
    block_gap = 0.06
    block_h = (div_zone_h - 2 * block_gap) / 3

    for i, blk in enumerate(div_blocks):
        by = div_top + 0.26 + i * (block_h + block_gap)

        # Title bar
        add_rect(slide, div_x, by, div_w, title_bar_h, DARK_BLUE)
        add_text_box(slide, div_x+0.12, by+0.03, div_w-0.15, title_bar_h-0.05,
                     blk["title"], font_size=8.5, bold=True, color=WHITE)

        # Body
        body_h = block_h - title_bar_h
        add_rect(slide, div_x, by+title_bar_h, div_w, body_h, WHITE, BORDER_CLR, 0.3)
        add_bullet_text(slide, div_x+0.1, by+title_bar_h+0.05,
                        div_w-0.18, body_h-0.08, blk["bullets"], font_size=8)

    # Investment Rationale header strip
    add_rect(slide, 0, rat_hdr_y, W, 0.3, LIGHT_BG)
    add_rect(slide, 0.25, rat_hdr_y+0.04, 0.05, 0.22, GOLD)
    add_text_box(slide, 0.38, rat_hdr_y, 5, 0.3,
                 "INVESTMENT RATIONALE", font_size=9, bold=True, color=DARK_BLUE)

    # Three rationale columns
    rat_col_w = (W - 0.5) / 3
    rat_cols = [
        {"title": data.get("rationaleCol1Title", ""), "body": data.get("rationaleCol1Body", "")},
        {"title": data.get("rationaleCol2Title", ""), "body": data.get("rationaleCol2Body", "")},
        {"title": data.get("rationaleCol3Title", ""), "body": data.get("rationaleCol3Body", "")},
    ]
    for i, col in enumerate(rat_cols):
        rx = 0.25 + i * rat_col_w
        add_text_box(slide, rx, rat_body_y, rat_col_w-0.12, 0.28,
                     col["title"], font_size=8.5, bold=True, color=GOLD, wrap=True)
        add_text_box(slide, rx, rat_body_y+0.3, rat_col_w-0.12, rat_body_h-0.3,
                     col["body"], font_size=8, color=DARK_GREY, wrap=True)
        if i < 2:
            # Vertical divider
            add_rect(slide, rx+rat_col_w-0.06, rat_body_y,
                     0.005, rat_body_h, BORDER_CLR)

    # Footer
    add_rect(slide, 0, H-0.32, W, 0.32, DARK_BLUE)
    add_text_box(slide, 0, H-0.28, W, 0.24,
                 "Conscience MFO  |  Confidential — For Addressee Only  |  Not Investment Advice",
                 font_size=7, color=RGBColor(0x88,0xAA,0xCC), align=PP_ALIGN.CENTER)
    add_text_box(slide, W-0.55, H-0.28, 0.45, 0.24, "2 / 2",
                 font_size=7, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)

    return slide


# ── Flask endpoint ─────────────────────────────────────────────────────────────
@app.route("/generate", methods=["POST"])
def generate():
    body = request.get_json()
    data = body.get("reportData", {})

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    build_slide1(prs, data)
    build_slide2(prs, data)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    encoded = base64.b64encode(buf.read()).decode("utf-8")

    return jsonify({"pptxBase64": encoded})


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
