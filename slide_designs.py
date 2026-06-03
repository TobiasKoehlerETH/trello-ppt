#!/usr/bin/env python3
"""Render three improved design options for the To do / In progress / Completed
status slide into one .pptx (one slide per option) for visual comparison.

Importance is encoded with Red / Yellow / Green (High / Medium / Low); column
headers are a neutral slate so colour means *only* priority. Uses live board
data when available, otherwise demo data. Output stays in the git-ignored
output/ folder.
"""

from __future__ import annotations

import datetime as _dt
import math
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from trello_to_ppt import HIGH, MEDIUM, LOW, NONE, LEVELS

HERE = Path(__file__).parent
OUT_PPTX = HERE / "output" / "designs.pptx"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
M = Inches(0.55)
BRAND_RED = RGBColor(0xE2, 0x00, 0x12)
SLATE = RGBColor(0x33, 0x3F, 0x50)
INK = RGBColor(0x1F, 0x2A, 0x37)
MUTED = RGBColor(0x77, 0x80, 0x8A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
HAIRLINE = RGBColor(0xE3, 0xE6, 0xEA)
FONT = "Segoe UI"
FONT_SB = "Segoe UI Semibold"

COLS_TOP = Inches(1.62)
COLS_BOTTOM = Inches(7.02)
GAP = Inches(0.30)
HEADER_H = Inches(0.52)


# --------------------------------------------------------------------------- #
# small helpers
# --------------------------------------------------------------------------- #
def _tint(color: RGBColor, ratio: float) -> RGBColor:
    return RGBColor(*(int(c + (255 - c) * ratio) for c in color))


def _no_shadow(shape):
    shape.shadow.inherit = False


def _soft_shadow(shape):
    spPr = shape._element.spPr
    for tag in ("a:effectLst",):
        existing = spPr.find(qn(tag))
        if existing is not None:
            spPr.remove(existing)
    xml = (
        '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:outerShdw blurRad="50000" dist="22000" dir="5400000" rotWithShape="0">'
        '<a:srgbClr val="111827"><a:alpha val="22000"/></a:srgbClr>'
        '</a:outerShdw></a:effectLst>'
    )
    spPr.append(etree.fromstring(xml))


def _rect(slide, x, y, w, h, fill, *, rounded=False, line=None, line_w=0.75, shadow=False):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE, x, y, w, h
    )
    if rounded:
        try:
            shape.adjustments[0] = 0.10
        except Exception:
            pass
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(line_w)
    if shadow:
        _soft_shadow(shape)
    else:
        _no_shadow(shape)
    return shape


def _text(slide, x, y, w, h, runs, *, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
          wrap=True, pad=0.0):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(pad)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.alignment = align
    for text, size, color, bold, font in runs:
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.bold = bold
        r.font.name = font
    return box


def _dot(slide, x, y, d, color):
    o = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    o.fill.solid()
    o.fill.fore_color.rgb = color
    o.line.fill.background()
    _no_shadow(o)
    return o


def _est_h(name, cpl=30, base=0.46, per=0.27, maxlines=3):
    lines = min(maxlines, max(1, math.ceil(len(name) / cpl)))
    return Inches(base + per * (lines - 1))


# --------------------------------------------------------------------------- #
# shared slide chrome (title, date, legend, column headers)
# --------------------------------------------------------------------------- #
def _chrome(slide, columns, option_label):
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, WHITE)  # white background

    _rect(slide, M, Inches(0.42), Inches(1.0), Inches(0.08), BRAND_RED)  # accent
    _text(slide, M, Inches(0.52), Inches(7), Inches(0.7),
          [("Overview", 30, INK, True, FONT_SB)], anchor=MSO_ANCHOR.TOP)

    today = _dt.date.today().strftime("%d.%m.%Y")
    _text(slide, SLIDE_W - Inches(3.3), Inches(0.55), Inches(2.75), Inches(0.4),
          [(today, 13, MUTED, False, FONT)], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP)

    # legend
    lx = SLIDE_W - Inches(5.6)
    ly = Inches(1.04)
    for level, label in ((HIGH, "High"), (MEDIUM, "Medium"), (LOW, "Low")):
        _dot(slide, lx, ly + Inches(0.04), Inches(0.16), LEVELS[level]["color"])
        _text(slide, lx + Inches(0.22), ly - Inches(0.03), Inches(1.3), Inches(0.3),
              [(label, 11, MUTED, False, FONT)], anchor=MSO_ANCHOR.TOP)
        lx += Inches(1.55)

    _text(slide, M, SLIDE_H - Inches(0.42), Inches(8), Inches(0.3),
          [(option_label, 11, MUTED, False, FONT)], anchor=MSO_ANCHOR.TOP)

    n = len(columns)
    col_w = int((SLIDE_W - 2 * M - GAP * (n - 1)) / n)
    xs = [M + i * (col_w + GAP) for i in range(n)]
    for i, (title, cards) in enumerate(columns):
        hdr = _rect(slide, xs[i], COLS_TOP, col_w, HEADER_H, SLATE, rounded=True)
        _text(slide, xs[i] + Inches(0.22), COLS_TOP, col_w - Inches(1.0), HEADER_H,
              [(title, 14, WHITE, True, FONT_SB)])
        _text(slide, xs[i] + col_w - Inches(0.95), COLS_TOP, Inches(0.72), HEADER_H,
              [(str(len(cards)), 14, _tint(SLATE, 0.55), True, FONT_SB)], align=PP_ALIGN.RIGHT)
    return xs, col_w


def _more(slide, x, y, w, remaining):
    _text(slide, x + Inches(0.1), y, w, Inches(0.3),
          [(f"+{remaining} more", 11, MUTED, False, FONT)], anchor=MSO_ANCHOR.TOP)


def _lay_cards(slide, xs, col_w, columns, render, *, gap=Inches(0.16), reserve=Inches(0.34)):
    top = COLS_TOP + HEADER_H + Inches(0.18)
    for i, (_title, cards) in enumerate(columns):
        def fit(limit):
            y, k = top, 0
            for name, _lvl in cards:
                h = _est_h(name)
                if y + h > limit:
                    break
                y, k = y + h + gap, k + 1
            return k
        shown = fit(COLS_BOTTOM)
        if shown < len(cards):
            shown = fit(COLS_BOTTOM - reserve)
        y = top
        for name, level in cards[:shown]:
            y += render(slide, xs[i], y, col_w, name, level) + gap
        if shown < len(cards):
            _more(slide, xs[i], y, col_w, len(cards) - shown)


# --------------------------------------------------------------------------- #
# Option A — tinted cards with a flush left accent bar (flat / modern)
# --------------------------------------------------------------------------- #
def _card_A(slide, x, y, w, name, level):
    h = _est_h(name)
    base = LEVELS[level]["color"]
    _rect(slide, x, y, w, h, _tint(base, 0.86), line=_tint(base, 0.45), line_w=0.75)
    _rect(slide, x, y, Inches(0.12), h, base)
    _text(slide, x + Inches(0.30), y, w - Inches(0.42), h,
          [(name, 13, INK, True, FONT)])
    return h


# --------------------------------------------------------------------------- #
# Option B — white rounded cards, soft shadow, colour tab on the left
# --------------------------------------------------------------------------- #
def _card_B(slide, x, y, w, name, level):
    h = _est_h(name)
    base = LEVELS[level]["color"]
    _rect(slide, x, y, w, h, WHITE, rounded=True, line=HAIRLINE, line_w=0.75, shadow=True)
    _rect(slide, x + Inches(0.1), y + Inches(0.09), Inches(0.1), h - Inches(0.18),
          base, rounded=True)
    _text(slide, x + Inches(0.34), y, w - Inches(0.46), h,
          [(name, 13, INK, True, FONT)])
    return h


# --------------------------------------------------------------------------- #
# Option C — minimal rows: priority dot + title on alternating tint
# --------------------------------------------------------------------------- #
def _card_C(slide, x, y, w, name, level):
    h = _est_h(name, base=0.5, per=0.26)
    base = LEVELS[level]["color"]
    _rect(slide, x, y, w, h, _tint(base, 0.93))
    _dot(slide, x + Inches(0.16), y + (h - Inches(0.18)) / 2, Inches(0.18), base)
    _text(slide, x + Inches(0.46), y, w - Inches(0.58), h,
          [(name, 13, INK, True, FONT)])
    return h


# --------------------------------------------------------------------------- #
def _blank(prs):
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return min(prs.slide_layouts, key=lambda l: len(l.placeholders))


def build(columns) -> Path:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    options = [
        ("Option A — Tinted cards with priority accent bar", _card_A),
        ("Option B — White cards, soft shadow & colour tab", _card_B),
        ("Option C — Minimal rows with priority dots", _card_C),
    ]
    for label, render in options:
        slide = prs.slides.add_slide(_blank(prs))
        xs, col_w = _chrome(slide, columns, label)
        _lay_cards(slide, xs, col_w, columns, render)
    OUT_PPTX.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT_PPTX))
    return OUT_PPTX


def _demo_columns():
    return [
        ("To do", [("Define Q3 roadmap", HIGH), ("Draft onboarding guide", MEDIUM),
                   ("Order spare sensors", HIGH), ("Clean up backlog labels", LOW),
                   ("Book venue for workshop", NONE)]),
        ("In progress", [("Payment service refactor", HIGH),
                         ("Search indexing improvements", MEDIUM),
                         ("Migrate CI to new runners", MEDIUM),
                         ("Update dependency versions", LOW)]),
        ("Completed", [("Ship auth token refresh", HIGH),
                       ("Fix board sync race condition", MEDIUM),
                       ("Add PPTX export pipeline", LOW),
                       ("Write API documentation", NONE)]),
    ]


def main() -> int:
    columns = None
    try:
        from trello_integration import load_dotenv
        from overview_deck import fetch_overview_columns
        load_dotenv()
        columns = fetch_overview_columns("Tasks", days=3)
        if not any(c for _, c in columns):
            columns = None
    except Exception as exc:  # noqa: BLE001
        print(f"  (live data unavailable: {exc}; using demo data)")
    if columns is None:
        columns = _demo_columns()
    out = build(columns)
    print(f"Wrote {out} ({sum(len(c) for _, c in columns)} cards)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
