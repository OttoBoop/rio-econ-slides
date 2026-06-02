"""
build_pptx.py — Rio de Janeiro Economic Development Slides
Generates output/rio_econ_slides.pptx from slide_content.py

Run: python scripts/build_pptx.py
"""

import sys
import os
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Cm

# ── path setup ──────────────────────────────────────────────────────────────
ROOT = Path(__file__).parent.parent
ASSETS = ROOT / "assets"
OUTPUT = ROOT / "output"
OUTPUT.mkdir(exist_ok=True)

sys.path.insert(0, str(ROOT / "scripts"))
from slide_content import SLIDES

# ── design tokens ────────────────────────────────────────────────────────────
NAVY     = RGBColor(0x00, 0x2B, 0x5E)   # dark navy background
BLUE     = RGBColor(0x00, 0x7A, 0xC2)   # accent blue
AMBER    = RGBColor(0xE8, 0x7D, 0x00)   # break slide accent
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BG = RGBColor(0xF4, 0xF7, 0xFB)   # near-white for data slides
DARK_TXT = RGBColor(0x0D, 0x1B, 0x3E)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# widescreen 16:9
prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

BLANK = prs.slide_layouts[6]  # truly blank layout


# ── helpers ──────────────────────────────────────────────────────────────────

def add_rect(slide, l, t, w, h, color):
    shape = slide.shapes.add_shape(1, l, t, w, h)  # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_textbox(slide, l, t, w, h, text, size=18, bold=False, color=WHITE,
                align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txb = slide.shapes.add_textbox(l, t, w, h)
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return txb


def add_multiline_textbox(slide, l, t, w, h, lines, size=16, color=WHITE,
                           bold=False, align=PP_ALIGN.LEFT, line_spacing_pt=None):
    """lines: list of (text, bold_override, size_override) or plain strings."""
    txb = slide.shapes.add_textbox(l, t, w, h)
    tf = txb.text_frame
    tf.word_wrap = True
    first = True
    for item in lines:
        if isinstance(item, str):
            text, b, s = item, bold, size
        else:
            text, b, s = item
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        if line_spacing_pt:
            p.line_spacing = Pt(line_spacing_pt)
        run = p.add_run()
        run.text = text
        run.font.size = Pt(s)
        run.font.bold = b
        run.font.color.rgb = color
        run.font.name = "Calibri"
    return txb


def logo_bar(slide, logo_path=None):
    """Thin navy footer bar with logo text."""
    bar_h = Inches(0.45)
    bar_t = SLIDE_H - bar_h
    add_rect(slide, 0, bar_t, SLIDE_W, bar_h, NAVY)
    add_textbox(slide, Inches(0.2), bar_t + Pt(4), Inches(8), bar_h,
                "PREFEITURA RIO  ·  Secretaria de Desenvolvimento Econômico, Inovação e Serviços",
                size=9, color=WHITE, align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(9), bar_t + Pt(4), Inches(4), bar_h,
                "Coleção Estudos Cariocas · IPP · 2026",
                size=9, color=RGBColor(0xAA, 0xCC, 0xFF), align=PP_ALIGN.RIGHT)


def accent_line(slide, t_offset=Inches(1.55), color=BLUE):
    add_rect(slide, Inches(0.5), t_offset, Inches(12.33), Inches(0.04), color)


# ── slide builders ────────────────────────────────────────────────────────────

def build_title(slide_data):
    sl = prs.slides.add_slide(BLANK)

    # full navy background
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, NAVY)

    # left blue accent bar
    add_rect(sl, 0, 0, Inches(0.18), SLIDE_H, BLUE)

    # SMDEIS logo image (top-right)
    logo_path = ASSETS / "logos" / "smdeis_header.jpeg"
    if logo_path.exists():
        sl.shapes.add_picture(str(logo_path), Inches(9.5), Inches(0.25), Inches(3.6), Inches(0.9))

    # main title
    add_textbox(sl, Inches(0.5), Inches(1.4), Inches(9), Inches(2.2),
                slide_data["title"], size=38, bold=True, color=WHITE)

    # blue underline
    add_rect(sl, Inches(0.5), Inches(3.7), Inches(5), Inches(0.05), BLUE)

    # subtitle
    add_textbox(sl, Inches(0.5), Inches(3.85), Inches(9), Inches(0.9),
                slide_data["subtitle"], size=18, color=RGBColor(0xAA, 0xCC, 0xFF))

    # presenter
    add_textbox(sl, Inches(0.5), Inches(4.9), Inches(9), Inches(0.5),
                slide_data["presenter"], size=20, bold=True, color=WHITE)
    add_textbox(sl, Inches(0.5), Inches(5.45), Inches(10), Inches(0.7),
                slide_data["role"], size=13, color=RGBColor(0xAA, 0xCC, 0xFF))

    # event line at bottom
    add_textbox(sl, Inches(0.5), Inches(6.65), Inches(9), Inches(0.4),
                slide_data["event"], size=11, color=RGBColor(0x77, 0xAA, 0xDD), italic=True)


def build_content(slide_data):
    sl = prs.slides.add_slide(BLANK)
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, LIGHT_BG)
    add_rect(sl, 0, 0, SLIDE_W, Inches(1.5), NAVY)
    add_rect(sl, 0, 0, Inches(0.10), SLIDE_H, BLUE)

    add_textbox(sl, Inches(0.3), Inches(0.18), Inches(12.5), Inches(0.8),
                slide_data["title"], size=30, bold=True, color=WHITE)

    if slide_data.get("subtitle"):
        add_textbox(sl, Inches(0.3), Inches(0.95), Inches(12), Inches(0.45),
                    slide_data["subtitle"], size=14, color=RGBColor(0xAA, 0xCC, 0xFF))

    bullet_t = Inches(1.65)
    for bullet in slide_data.get("bullets", []):
        txb = slide.shapes.add_textbox if False else sl.shapes.add_textbox(
            Inches(0.5), bullet_t, Inches(12.3), Inches(0.55))
        tf = txb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = "▸  " + bullet
        run.font.size = Pt(16)
        run.font.color.rgb = DARK_TXT
        run.font.name = "Calibri"
        bullet_t += Inches(0.72)

    if slide_data.get("footer_note"):
        add_rect(sl, Inches(0.3), Inches(6.6), Inches(12.5), Inches(0.04), BLUE)
        add_textbox(sl, Inches(0.3), Inches(6.68), Inches(12.5), Inches(0.45),
                    slide_data["footer_note"], size=12, italic=True,
                    color=RGBColor(0x00, 0x4C, 0x97))

    logo_bar(sl)


def build_stats(slide_data):
    sl = prs.slides.add_slide(BLANK)
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(sl, 0, 0, Inches(0.10), SLIDE_H, BLUE)

    add_textbox(sl, Inches(0.3), Inches(0.18), Inches(12.5), Inches(0.8),
                slide_data["title"], size=30, bold=True, color=WHITE)
    accent_line(sl, t_offset=Inches(1.05), color=BLUE)

    stats = slide_data["stats"]
    col_w = SLIDE_W / len(stats)
    for i, s in enumerate(stats):
        cx = Inches(0.3) + i * col_w
        # value
        add_textbox(sl, cx, Inches(1.4), col_w - Inches(0.2), Inches(1.4),
                    s["value"], size=34, bold=True, color=BLUE,
                    align=PP_ALIGN.CENTER)
        # label
        add_textbox(sl, cx, Inches(2.85), col_w - Inches(0.2), Inches(0.6),
                    s["label"], size=16, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER)
        # detail
        add_textbox(sl, cx + Inches(0.15), Inches(3.55), col_w - Inches(0.5), Inches(0.8),
                    s["detail"], size=12, color=RGBColor(0xAA, 0xCC, 0xFF),
                    align=PP_ALIGN.CENTER)
        # vertical divider (except last)
        if i < len(stats) - 1:
            add_rect(sl, cx + col_w - Inches(0.02), Inches(1.4),
                     Inches(0.02), Inches(3.0), RGBColor(0x33, 0x55, 0x88))

    add_textbox(sl, Inches(0.3), Inches(6.6), Inches(12.5), Inches(0.4),
                slide_data.get("source", ""), size=10,
                color=RGBColor(0x77, 0xAA, 0xDD), italic=True)
    logo_bar(sl)


def build_chart_slide(slide_data):
    sl = prs.slides.add_slide(BLANK)
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, LIGHT_BG)
    add_rect(sl, 0, 0, SLIDE_W, Inches(1.4), NAVY)
    add_rect(sl, 0, 0, Inches(0.10), SLIDE_H, BLUE)

    add_textbox(sl, Inches(0.3), Inches(0.12), Inches(12.5), Inches(0.75),
                slide_data["title"], size=26, bold=True, color=WHITE)
    add_textbox(sl, Inches(0.3), Inches(0.88), Inches(12.5), Inches(0.45),
                slide_data["subtitle"], size=12, color=RGBColor(0xAA, 0xCC, 0xFF))

    graph_path = ASSETS / "graphs" / slide_data["graph"]
    if graph_path.exists():
        sl.shapes.add_picture(str(graph_path), Inches(0.4), Inches(1.45), Inches(9.5), Inches(4.8))
    else:
        add_textbox(sl, Inches(0.4), Inches(2.5), Inches(9.5), Inches(1),
                    f"[Graph not found: {slide_data['graph']}]", size=14,
                    color=RGBColor(0xCC, 0x00, 0x00))

    # callout box on right
    add_rect(sl, Inches(10.1), Inches(1.45), Inches(3.0), Inches(3.0), NAVY)
    add_textbox(sl, Inches(10.2), Inches(1.6), Inches(2.8), Inches(2.7),
                slide_data.get("callout", ""), size=13, color=WHITE,
                wrap=True)

    add_textbox(sl, Inches(0.3), Inches(6.6), Inches(12.5), Inches(0.4),
                slide_data.get("source", ""), size=10,
                color=RGBColor(0x55, 0x77, 0xAA), italic=True)
    logo_bar(sl)


def build_axes_grid(slide_data):
    sl = prs.slides.add_slide(BLANK)
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(sl, 0, 0, Inches(0.10), SLIDE_H, BLUE)

    add_textbox(sl, Inches(0.3), Inches(0.15), Inches(12.5), Inches(0.75),
                slide_data["title"], size=28, bold=True, color=WHITE)
    add_textbox(sl, Inches(0.3), Inches(0.9), Inches(12), Inches(0.45),
                slide_data["intro"], size=13, color=RGBColor(0xAA, 0xCC, 0xFF))
    accent_line(sl, Inches(1.38), BLUE)

    axes = slide_data["axes"]
    cols, rows = 6, 2
    cell_w = (SLIDE_W - Inches(0.6)) / cols
    cell_h = Inches(2.0)
    start_t = Inches(1.5)

    for idx, ax in enumerate(axes):
        col = idx % cols
        row = idx // cols
        cx = Inches(0.3) + col * cell_w
        cy = start_t + row * (cell_h + Inches(0.12))

        # cell background
        bg = RGBColor(0x00, 0x3B, 0x7A) if idx != 11 else RGBColor(0x5C, 0x2E, 0x00)
        add_rect(sl, cx, cy, cell_w - Inches(0.08), cell_h, bg)

        # number
        add_textbox(sl, cx + Inches(0.08), cy + Inches(0.1), cell_w - Inches(0.16), Inches(0.45),
                    ax["num"], size=11, bold=True, color=BLUE, align=PP_ALIGN.LEFT)
        # label
        add_textbox(sl, cx + Inches(0.08), cy + Inches(0.45), cell_w - Inches(0.16), Inches(1.3),
                    ax["label"], size=13, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # note about axis 12
    add_textbox(sl, Inches(0.3), Inches(6.55), Inches(12.5), Inches(0.4),
                slide_data.get("note", ""), size=10,
                color=RGBColor(0xFF, 0xAA, 0x55), italic=True)
    logo_bar(sl)


def build_axis_detail(slide_data):
    sl = prs.slides.add_slide(BLANK)
    is_security = slide_data.get("axis_num") == "12"
    header_color = AMBER if is_security else NAVY
    accent_color = AMBER if is_security else BLUE

    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, LIGHT_BG)
    add_rect(sl, 0, 0, SLIDE_W, Inches(1.55), header_color)
    add_rect(sl, 0, 0, Inches(0.10), SLIDE_H, accent_color)

    num_bg = accent_color
    add_rect(sl, Inches(0.3), Inches(0.15), Inches(0.65), Inches(0.65), num_bg)
    add_textbox(sl, Inches(0.3), Inches(0.15), Inches(0.65), Inches(0.65),
                slide_data["axis_num"], size=20, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)

    add_textbox(sl, Inches(1.1), Inches(0.15), Inches(11), Inches(0.5),
                slide_data["axis_label"].upper(), size=11, bold=True,
                color=accent_color if not is_security else WHITE)
    add_textbox(sl, Inches(1.1), Inches(0.62), Inches(11), Inches(0.75),
                slide_data["title"], size=24, bold=True, color=WHITE)

    if slide_data.get("disclaimer"):
        add_rect(sl, Inches(0.3), Inches(1.55), Inches(12.5), Inches(0.4),
                 RGBColor(0x5C, 0x2E, 0x00))
        add_textbox(sl, Inches(0.4), Inches(1.6), Inches(12.3), Inches(0.35),
                    slide_data["disclaimer"], size=10, color=RGBColor(0xFF, 0xCC, 0x66))
        item_start_t = Inches(2.05)
    else:
        item_start_t = Inches(1.7)

    items = slide_data.get("items", [])
    item_h = (Inches(4.8) - (item_start_t - Inches(1.7))) / max(len(items), 1)
    for item in items:
        add_rect(sl, Inches(0.3), item_start_t, Inches(3.8), item_h - Inches(0.08),
                 NAVY)
        add_textbox(sl, Inches(0.38), item_start_t + Inches(0.05),
                    Inches(3.65), item_h - Inches(0.15),
                    item["program"], size=13, bold=True, color=WHITE)

        add_textbox(sl, Inches(4.3), item_start_t + Inches(0.05),
                    Inches(8.7), item_h - Inches(0.15),
                    item["detail"], size=13, color=DARK_TXT)

        add_rect(sl, Inches(0.3), item_start_t + item_h - Inches(0.06),
                 Inches(12.5), Inches(0.015), RGBColor(0xCC, 0xDD, 0xEE))
        item_start_t += item_h

    if slide_data.get("callout"):
        add_rect(sl, Inches(0.3), Inches(6.3), Inches(12.5), Inches(0.05), accent_color)
        add_textbox(sl, Inches(0.3), Inches(6.38), Inches(12.5), Inches(0.5),
                    "➤  " + slide_data["callout"], size=13, bold=True,
                    color=accent_color if not is_security else AMBER)

    logo_bar(sl)


def build_remaining_axes(slide_data):
    sl = prs.slides.add_slide(BLANK)
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, LIGHT_BG)
    add_rect(sl, 0, 0, SLIDE_W, Inches(1.1), NAVY)
    add_rect(sl, 0, 0, Inches(0.10), SLIDE_H, BLUE)

    add_textbox(sl, Inches(0.3), Inches(0.18), Inches(12.5), Inches(0.75),
                slide_data["title"], size=26, bold=True, color=WHITE)

    remaining = slide_data["remaining"]
    cols, rows = 3, 2
    cell_w = (SLIDE_W - Inches(0.6)) / cols
    cell_h = Inches(2.5)
    start_t = Inches(1.2)

    for idx, ax in enumerate(remaining):
        col = idx % cols
        row = idx // cols
        cx = Inches(0.3) + col * cell_w
        cy = start_t + row * (cell_h + Inches(0.1))

        add_rect(sl, cx, cy, cell_w - Inches(0.1), cell_h, NAVY)
        add_textbox(sl, cx + Inches(0.1), cy + Inches(0.08),
                    cell_w - Inches(0.2), Inches(0.35),
                    ax["num"], size=11, bold=True, color=BLUE)
        add_textbox(sl, cx + Inches(0.1), cy + Inches(0.38),
                    cell_w - Inches(0.2), Inches(0.55),
                    ax["label"], size=14, bold=True, color=WHITE)
        add_rect(sl, cx + Inches(0.1), cy + Inches(0.97),
                 cell_w - Inches(0.3), Inches(0.025), BLUE)
        add_textbox(sl, cx + Inches(0.1), cy + Inches(1.05),
                    cell_w - Inches(0.2), Inches(1.35),
                    ax["detail"], size=11, color=RGBColor(0xAA, 0xCC, 0xFF))

    logo_bar(sl)


def build_break_slide(slide_data):
    sl = prs.slides.add_slide(BLANK)
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, DARK_TXT)
    # amber horizontal band
    add_rect(sl, 0, Inches(2.8), SLIDE_W, Inches(1.9), AMBER)

    add_textbox(sl, Inches(1.0), Inches(1.3), Inches(11), Inches(1.4),
                slide_data["title"], size=52, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)
    add_textbox(sl, Inches(1.0), Inches(2.9), Inches(11), Inches(1.7),
                slide_data["subtitle"], size=36, bold=True, color=DARK_TXT,
                align=PP_ALIGN.CENTER)
    add_textbox(sl, Inches(1.5), Inches(5.0), Inches(10), Inches(1.5),
                slide_data["note"], size=14, color=RGBColor(0xAA, 0xCC, 0xFF),
                align=PP_ALIGN.CENTER, italic=True)


def build_forward(slide_data):
    # reuse content builder logic
    sl = prs.slides.add_slide(BLANK)
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(sl, 0, 0, Inches(0.10), SLIDE_H, BLUE)

    add_textbox(sl, Inches(0.3), Inches(0.15), Inches(12.5), Inches(0.75),
                slide_data["title"], size=28, bold=True, color=WHITE)
    add_textbox(sl, Inches(0.3), Inches(0.9), Inches(12), Inches(0.45),
                slide_data.get("subtitle", ""), size=14,
                color=RGBColor(0xAA, 0xCC, 0xFF))
    accent_line(sl, Inches(1.38), BLUE)

    bullet_t = Inches(1.52)
    for bullet in slide_data.get("bullets", []):
        txb = sl.shapes.add_textbox(Inches(0.5), bullet_t, Inches(12.3), Inches(0.65))
        tf = txb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "▸  " + bullet
        run.font.size = Pt(15)
        run.font.color.rgb = WHITE
        run.font.name = "Calibri"
        bullet_t += Inches(0.75)

    if slide_data.get("footer_note"):
        add_textbox(sl, Inches(0.3), Inches(6.6), Inches(12.5), Inches(0.45),
                    slide_data["footer_note"], size=11,
                    color=RGBColor(0x77, 0xAA, 0xDD), italic=True)
    logo_bar(sl)


def build_why_slide(slide_data):
    sl = prs.slides.add_slide(BLANK)
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, LIGHT_BG)
    add_rect(sl, 0, 0, SLIDE_W, Inches(1.1), NAVY)
    add_rect(sl, 0, 0, Inches(0.10), SLIDE_H, BLUE)

    add_textbox(sl, Inches(0.3), Inches(0.18), Inches(12.5), Inches(0.75),
                slide_data["title"], size=26, bold=True, color=WHITE)

    points = slide_data["points"]
    cols, rows = 3, 2
    cell_w = (SLIDE_W - Inches(0.6)) / cols
    cell_h = Inches(2.6)
    start_t = Inches(1.15)

    for idx, pt in enumerate(points):
        col = idx % cols
        row = idx // cols
        cx = Inches(0.3) + col * cell_w
        cy = start_t + row * (cell_h + Inches(0.06))

        add_rect(sl, cx, cy, cell_w - Inches(0.1), cell_h, NAVY)
        # icon + label header
        add_textbox(sl, cx + Inches(0.12), cy + Inches(0.12),
                    cell_w - Inches(0.22), Inches(0.5),
                    pt["label"], size=15, bold=True, color=BLUE)
        add_rect(sl, cx + Inches(0.12), cy + Inches(0.65),
                 cell_w - Inches(0.3), Inches(0.02), BLUE)
        add_textbox(sl, cx + Inches(0.12), cy + Inches(0.75),
                    cell_w - Inches(0.22), Inches(1.65),
                    pt["detail"], size=12, color=WHITE)

    logo_bar(sl)


def build_bio_slide(slide_data):
    sl = prs.slides.add_slide(BLANK)
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(sl, 0, 0, Inches(0.10), SLIDE_H, BLUE)

    # photo on left
    photo_path = ASSETS / "photos" / slide_data["photo"]
    if photo_path.exists():
        sl.shapes.add_picture(str(photo_path), Inches(0.5), Inches(0.9), Inches(3.8), Inches(5.0))

    # name + title
    add_textbox(sl, Inches(4.8), Inches(1.0), Inches(8.2), Inches(0.9),
                slide_data["name"], size=36, bold=True, color=WHITE)
    add_textbox(sl, Inches(4.8), Inches(1.95), Inches(8.2), Inches(0.5),
                slide_data["title"], size=18, color=BLUE)

    accent_line(sl, Inches(2.6), BLUE)

    add_textbox(sl, Inches(4.8), Inches(2.75), Inches(8.2), Inches(3.5),
                slide_data["bio"], size=15, color=RGBColor(0xCC, 0xDD, 0xFF))

    # SMDEIS logo bottom-left area
    logo_path = ASSETS / "logos" / "smdeis_header.jpeg"
    if logo_path.exists():
        sl.shapes.add_picture(str(logo_path), Inches(0.5), Inches(6.2), Inches(3.5), Inches(0.85))

    logo_bar(sl)


def build_multi_axis(slide_data):
    """2 or 3 axes per slide, each with a bulleted program list below."""
    sl = prs.slides.add_slide(BLANK)
    add_rect(sl, 0, 0, SLIDE_W, SLIDE_H, LIGHT_BG)
    add_rect(sl, 0, 0, SLIDE_W, Inches(1.1), NAVY)
    add_rect(sl, 0, 0, Inches(0.10), SLIDE_H, BLUE)

    add_textbox(sl, Inches(0.3), Inches(0.15), Inches(12.5), Inches(0.75),
                slide_data["title"], size=24, bold=True, color=WHITE)

    axes = slide_data["axes"]
    n = len(axes)
    col_w = (SLIDE_W - Inches(0.6)) / n
    start_t = Inches(1.15)

    for i, ax in enumerate(axes):
        cx = Inches(0.3) + i * col_w
        col_inner_w = col_w - Inches(0.15)

        # axis number badge
        add_rect(sl, cx, start_t, Inches(0.5), Inches(0.45), BLUE)
        add_textbox(sl, cx, start_t, Inches(0.5), Inches(0.45),
                    ax["num"], size=14, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER)

        # axis label
        add_textbox(sl, cx + Inches(0.58), start_t + Inches(0.02),
                    col_inner_w - Inches(0.58), Inches(0.55),
                    ax["label"], size=14, bold=True, color=DARK_TXT)

        # divider
        add_rect(sl, cx, start_t + Inches(0.52), col_inner_w,
                 Inches(0.025), BLUE)

        # programs list
        prog_t = start_t + Inches(0.62)
        for prog in ax.get("programs", []):
            txb = sl.shapes.add_textbox(cx + Inches(0.1), prog_t,
                                        col_inner_w - Inches(0.1), Inches(0.8))
            tf = txb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = "▸ " + prog
            run.font.size = Pt(11)
            run.font.color.rgb = DARK_TXT
            run.font.name = "Calibri"
            prog_t += Inches(0.9)

        # vertical divider between columns
        if i < n - 1:
            add_rect(sl, cx + col_w - Inches(0.075), start_t,
                     Inches(0.015), Inches(5.9), RGBColor(0xCC, 0xDD, 0xEE))

    logo_bar(sl)


# ── dispatch ─────────────────────────────────────────────────────────────────

BUILDERS = {
    "title":         build_title,
    "content":       build_content,
    "stats":         build_stats,
    "chart_slide":   build_chart_slide,
    "axes_grid":     build_axes_grid,
    "axis_detail":   build_axis_detail,
    "multi_axis":    build_multi_axis,
    "remaining_axes": build_remaining_axes,
    "break_slide":   build_break_slide,
    "forward":       build_forward,
    "why_slide":     build_why_slide,
    "bio_slide":     build_bio_slide,
}


def main():
    for slide_data in SLIDES:
        stype = slide_data["type"]
        builder = BUILDERS.get(stype)
        if builder is None:
            print(f"  WARNING: no builder for type '{stype}' (slide id: {slide_data.get('id')})")
            continue
        print(f"  Building [{slide_data['id']}] ({stype}) ...")
        builder(slide_data)

    out = OUTPUT / "rio_econ_slides.pptx"
    prs.save(str(out))
    print(f"\n✓ Saved: {out}")
    print(f"  Slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
