#!/usr/bin/env python3
"""Build the Insider Threat Campaign executive scorecard PPTX (slide 1)."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---- Palette (Exabeam brand) ----
INK        = RGBColor(0x0A, 0x0A, 0x0A)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
PAPER      = RGBColor(0xF5, 0xF5, 0xF5)
LINE       = RGBColor(0xE0, 0xE0, 0xE0)
MUTED      = RGBColor(0x70, 0x6E, 0x6B)
BLUE       = RGBColor(0x00, 0x6B, 0xFF)
BLUE_DARK  = RGBColor(0x00, 0x3F, 0xCC)
GREEN      = RGBColor(0x00, 0x9D, 0x00)
GREEN_DK   = RGBColor(0x10, 0x6D, 0x00)
GREEN_BG   = RGBColor(0xEA, 0xF6, 0xEE)
AMBER      = RGBColor(0xB5, 0x7A, 0x00)
AMBER_BG   = RGBColor(0xFD, 0xF2, 0xDD)
BLUE_BG    = RGBColor(0xEA, 0xF2, 0xFC)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank = prs.slide_layouts[6]

def set_fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()

def add_rect(slide, x, y, w, h, color, radius=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    set_fill(shp, color)
    shp.shadow.inherit = False
    return shp

def add_text(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT,
             font='Arial', anchor=MSO_ANCHOR.TOP, spacing=None, italic=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = font
    r.font.color.rgb = color
    if spacing:
        rPr = r._r.get_or_add_rPr()
        rPr.set('spc', str(spacing))
    return box

def add_pill(slide, x, y, w, h, text, bg, fg):
    shp = add_rect(slide, x, y, w, h, bg, radius=0.5)
    tf = shp.text_frame
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size = Pt(10.5)
    r.font.bold = True
    r.font.name = 'Arial'
    r.font.color.rgb = fg
    return shp

slide = prs.slides.add_slide(blank)
add_rect(slide, 0, 0, 13.333, 7.5, WHITE)

# ---- Header band ----
add_rect(slide, 0, 0, 13.333, 1.15, INK)
add_text(slide, 0.55, 0.18, 9.5, 0.5, "INSIDER THREAT CAMPAIGN", 22, WHITE, bold=True, font='Arial', spacing=40)
add_text(slide, 0.55, 0.66, 9.5, 0.35, "Executive Scorecard  ·  Jul 1 – Sep 30, 2026", 12.5, RGBColor(0xC9,0xC8,0xDE))
add_text(slide, 9.9, 0.30, 2.9, 0.3, "Prepared: Oct 1, 2026", 10, RGBColor(0x9C,0x9B,0xB4), align=PP_ALIGN.RIGHT)
add_text(slide, 9.9, 0.58, 2.9, 0.3, "Marketing Analytics", 10, RGBColor(0x9C,0x9B,0xB4), align=PP_ALIGN.RIGHT)

# ---- Overall status strip ----
add_rect(slide, 0.55, 1.35, 12.23, 0.62, GREEN_BG, radius=0.25)
add_text(slide, 0.85, 1.35, 4.2, 0.62, "●  ON TRACK", 14, GREEN_DK, bold=True, anchor=MSO_ANCHOR.MIDDLE)
add_text(slide, 5.0, 1.35, 7.5, 0.62,
         "4 of 6 KPIs met or exceeded goal  ·  Pipeline +37% vs. target  ·  Opportunity creation trailing goal by 14%",
         12, INK, anchor=MSO_ANCHOR.MIDDLE)

# ---- KPI cards ----
kpis = [
    ("Spend vs. Budget",  "$184.3K", "of $250K budget", "73.7%", "ON TRACK", BLUE_BG, BLUE_DARK),
    ("Leads Generated",   "3,842",   "Goal: 3,500",      "110%",  "EXCEEDED", GREEN_BG, GREEN_DK),
    ("MQLs",              "1,120",   "Goal: 1,000",      "112%",  "EXCEEDED", GREEN_BG, GREEN_DK),
    ("Opportunities",     "214",     "Goal: 250",        "86%",   "AT RISK",  AMBER_BG, AMBER),
    ("Pipeline Generated","$4.8M",   "Goal: $3.5M",      "137%",  "EXCEEDED", GREEN_BG, GREEN_DK),
    ("Campaign ROI",      "551%",    "Goal: 400%",       "138%",  "EXCEEDED", GREEN_BG, GREEN_DK),
]
card_w = 1.93
gap = 0.12
start_x = 0.55
card_y = 2.15
card_h = 1.95
for i, (label, val, sub, pct, status, sbg, sfg) in enumerate(kpis):
    x = start_x + i*(card_w+gap)
    add_rect(slide, x, card_y, card_w, card_h, PAPER, radius=0.06)
    add_text(slide, x+0.14, card_y+0.14, card_w-0.28, 0.4, label.upper(), 9, MUTED, bold=True, spacing=10)
    add_text(slide, x+0.14, card_y+0.46, card_w-0.28, 0.55, val, 24, INK, bold=True)
    add_text(slide, x+0.14, card_y+0.98, card_w-0.28, 0.3, sub, 9.5, MUTED)
    add_pill(slide, x+0.14, card_y+1.38, card_w-0.28, 0.32, status + "  ·  " + pct, sbg, sfg)

# ---- Key Takeaways (left) ----
tak_y = 4.20
add_text(slide, 0.55, tak_y, 6.9, 0.35, "KEY TAKEAWAYS", 12, INK, bold=True, spacing=15)
add_rect(slide, 0.55, tak_y+0.38, 0.5, 0.03, BLUE)
takeaways = [
    "Lead and pipeline generation significantly exceeded targets (pipeline +37% vs. goal), driven by Paid Search, Email Nurture, and the new Physical Events channel.",
    "Opportunity creation is trailing goal (214 vs. 250) despite strong top-of-funnel volume — signals an MQL-to-SQL handoff gap worth reviewing with Sales.",
    "PathFactory content scoring directly influenced 148 opportunities (~$1.1M pipeline), validating the TOFU → MOFU → BOFU nurture model.",
    "Campaign is under budget at 73.7% spend utilization with one month remaining — room to reinvest in top-performing tactics (Whitepaper, Webinar Series).",
]
ty = tak_y + 0.55
for t in takeaways:
    add_rect(slide, 0.55, ty+0.09, 0.09, 0.09, BLUE)
    add_text(slide, 0.80, ty, 6.6, 0.5, t, 10.5, RGBColor(0x3E,0x3E,0x3C))
    ty += 0.54

# ---- Right column: pipeline by tactic mini ranking ----
rx = 7.85
add_text(slide, rx, tak_y, 4.9, 0.35, "TOP TACTICS BY PIPELINE", 12, INK, bold=True, spacing=15)
add_rect(slide, rx, tak_y+0.38, 0.5, 0.03, BLUE)
tactics = [
    ("Gated Whitepaper", 1240000, "#0d366b"),
    ("Webinar Series", 980000, "#104281"),
    ("ABM Direct Mail", 760000, "#184f95"),
    ("LinkedIn InMail", 610000, "#1c5cab"),
    ("Exec Roundtables", 520000, "#2a78d6"),
]
max_v = tactics[0][1]
by = tak_y + 0.62
bar_area_w = 3.55
label_w = 1.65
for name, val, hexcolor in tactics:
    rgb = RGBColor(int(hexcolor[1:3],16), int(hexcolor[3:5],16), int(hexcolor[5:7],16))
    add_text(slide, rx, by, label_w, 0.28, name, 9.5, RGBColor(0x3E,0x3E,0x3C))
    bw = max(0.15, (val/max_v) * bar_area_w)
    add_rect(slide, rx+label_w+0.1, by+0.02, bw, 0.24, rgb, radius=0.3)
    add_text(slide, rx+label_w+0.1+bw+0.08, by, 1.3, 0.28, "${:.2f}M".format(val/1_000_000), 9.5, INK, bold=True)
    by += 0.42

# ---- Footer ----
add_rect(slide, 0, 7.18, 13.333, 0.32, PAPER)
add_text(slide, 0.55, 7.20, 9, 0.28,
         "Prototype scorecard — sample data for illustration only. Not a real Exabeam campaign report.",
         8.5, MUTED, italic=True)
add_text(slide, 10.5, 7.20, 2.3, 0.28, "Page 1 of 2", 8.5, MUTED, align=PP_ALIGN.RIGHT)

prs.save('/private/tmp/claude-501/-Users-scottclinton/b975b1c5-0616-4716-ad2d-ab6bcf229c90/scratchpad/campaign-dashboard/Insider_Threat_Campaign_Scorecard.pptx')
print("Saved.")
