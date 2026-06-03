#!/usr/bin/env python3
"""Build a 2-slide Trello status deck on top of a branded example template.

* **Slide 1** is taken verbatim from the template (the title slide); only the
  date is refreshed to today (``DD.MM.YYYY``).
* **Slide 2 ("Overview")** keeps the template's 3-column table header
  (To do / In progress / Completed) and is filled with one textbox per card.
  Each card's **border colour encodes importance** — Red (High) / Yellow
  (Medium) / Green (Low), grey when unrated.

Importance comes from each card's Trello labels (see
:func:`trello_to_ppt.classify_importance`). With no board / invalid credentials,
``--demo`` renders sample data so you can preview the layout.

Examples
--------
    python overview_deck.py --demo
    python overview_deck.py "My Board" -o output/status_deck.pptx
"""

from __future__ import annotations

import argparse
import datetime as _dt
import math
import re
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

from trello_integration import load_dotenv, trello_request
from trello_to_ppt import (
    AMBER, GREEN, GREY, HIGH, INK, LOW, MEDIUM, MUTED, NONE, RED, WHITE,
    LEVELS, classify_importance, resolve_board,
)

HERE = Path(__file__).parent
DEFAULT_TEMPLATE = HERE / "example.pptx"
DEFAULT_OUTPUT = HERE / "output" / "status_deck.pptx"

CARD_FONT = "Futura Md BT"          # matches the template
CARD_FONT_PT = 14
CARD_BORDER_PT = 2.25
DATE_RE = re.compile(r"^\d{1,2}\.\d{1,2}\.\d{4}$")

# Which Trello lists (lower-cased names) feed each template column.
# "Completed" is special-cased: it shows cards MOVED into those lists within the
# last N days, not everything currently sitting there.
COLUMN_SOURCES: dict[str, list[str]] = {
    "To do": ["today"],
    "In progress": ["top priority"],
    "Completed": ["done"],
}
COLUMN_ORDER = tuple(COLUMN_SOURCES)
COMPLETED_COLUMN = "Completed"


# --------------------------------------------------------------------------- #
# Slide 1: refresh the date
# --------------------------------------------------------------------------- #
def refresh_date(slide, date_str: str) -> bool:
    """Replace the first ``DD.MM.YYYY`` run on the slide with ``date_str``."""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if DATE_RE.match(run.text.strip()):
                    run.text = date_str
                    return True
    return False


# --------------------------------------------------------------------------- #
# Slide 2: fill the table with importance-bordered card textboxes
# --------------------------------------------------------------------------- #
def _remove_textboxes(slide) -> None:
    """Drop the template's placeholder card textboxes (keep title + table)."""
    for shape in list(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
            shape._element.getparent().remove(shape._element)


def _find_table(slide):
    for shape in slide.shapes:
        if shape.has_table:
            return shape
    return None


def _clear_header_fill(table) -> None:
    """Remove background fill from the column header row."""
    for cell in table.rows[0].cells:
        cell.fill.background()


def _estimate_height(text: str, chars_per_line: int = 26, max_lines: int = 3) -> int:
    lines = min(max_lines, max(1, math.ceil(len(text) / chars_per_line)))
    return Inches(0.40 + 0.30 * (lines - 1))


def _tint(color, ratio: float):
    """Blend ``color`` toward white (ratio 0 = original, 1 = white)."""
    return RGBColor(*(int(c + (255 - c) * ratio) for c in color))


def _add_card(slide, left, top, width, text: str, level: int):
    """Option A card: an importance-tinted card with a solid left accent bar + bold title."""
    height = _estimate_height(text)
    base = LEVELS[level]["color"]

    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = _tint(base, 0.86)
    card.line.color.rgb = _tint(base, 0.45)
    card.line.width = Pt(0.75)
    card.shadow.inherit = False

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.12), height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = base
    bar.line.fill.background()
    bar.shadow.inherit = False

    box = slide.shapes.add_textbox(left + Inches(0.30), top, width - Inches(0.42), height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.04)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    run = tf.paragraphs[0].add_run()
    run.text = text
    run.font.name = CARD_FONT
    run.font.size = Pt(CARD_FONT_PT)
    run.font.bold = True
    run.font.color.rgb = INK
    return height


def _legend_slide2(slide) -> None:
    """Small Red/Yellow/Green priority legend at the top-right of the overview slide."""
    dia = Inches(0.15)
    y = Inches(0.74)
    x = Inches(8.95)
    for level, lbl in ((HIGH, "High"), (MEDIUM, "Medium"), (LOW, "Low")):
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y + Inches(0.04), dia, dia)
        dot.fill.solid()
        dot.fill.fore_color.rgb = LEVELS[level]["color"]
        dot.line.fill.background()
        dot.shadow.inherit = False
        box = slide.shapes.add_textbox(x + Inches(0.22), y - Inches(0.02), Inches(1.05), Inches(0.32))
        box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        run = box.text_frame.paragraphs[0].add_run()
        run.text = lbl
        run.font.name = CARD_FONT
        run.font.size = Pt(12)
        run.font.color.rgb = MUTED
        x += Inches(1.35)


def _add_more(slide, left, top, width, remaining: int) -> None:
    box = slide.shapes.add_textbox(left, top, width, Inches(0.3))
    run = box.text_frame.paragraphs[0].add_run()
    run.text = f"+{remaining} more"
    run.font.name = CARD_FONT
    run.font.size = Pt(11)
    run.font.italic = True
    run.font.color.rgb = MUTED


def fill_overview(slide, columns: list[tuple[str, list[tuple[str, int]]]]) -> None:
    """Lay importance-bordered card textboxes under each table column."""
    table_shape = _find_table(slide)
    if table_shape is None:
        raise SystemExit("Template slide 2 has no table to fill.")

    _clear_header_fill(table_shape.table)

    t_left, t_top = table_shape.left, table_shape.top
    t_width, t_height = table_shape.width, table_shape.height
    header_h = table_shape.table.rows[0].height
    n = len(columns)
    col_w = int(t_width / n)
    card_w = col_w - Inches(0.26)
    inset = int((col_w - card_w) / 2)
    body_top = t_top + header_h + Inches(0.12)
    body_bottom = t_top + t_height - Inches(0.12)
    gap = Inches(0.14)

    def count_fit(cards, limit):
        y, k = body_top, 0
        for name, _level in cards:
            h = _estimate_height(name)
            if y + h > limit:
                break
            y, k = y + h + gap, k + 1
        return k

    more_reserve = Inches(0.34)
    for i, (_title, cards) in enumerate(columns):
        x = t_left + i * col_w + inset
        shown = count_fit(cards, body_bottom)
        if shown < len(cards):  # leave room so "+N more" never overlaps a card
            shown = count_fit(cards, body_bottom - more_reserve)
        y = body_top
        for name, level in cards[:shown]:
            y += _add_card(slide, x, y, card_w, name, level) + gap
        if shown < len(cards):
            _add_more(slide, x + Inches(0.05), y, card_w, len(cards) - shown)


# --------------------------------------------------------------------------- #
# Data
# --------------------------------------------------------------------------- #
def fetch_overview_columns(board: str, days: int = 7) -> list[tuple[str, list[tuple[str, int]]]]:
    """Build the three template columns from the live board.

    To-do / In-progress columns list the cards currently in their source lists
    (highest importance first). The Completed column lists cards moved into the
    Done list(s) within the last ``days``.
    """
    board_id, _name, _url = resolve_board(board)
    lists = trello_request(
        "GET", f"/boards/{board_id}/lists", params={"fields": "name", "filter": "open"}
    )
    name_to_ids: dict[str, set[str]] = {}
    for lst in lists:
        name_to_ids.setdefault(lst.get("name", "").strip().lower(), set()).add(lst["id"])

    cards = trello_request(
        "GET", f"/boards/{board_id}/cards",
        params={"fields": "name,idList,labels,dateLastActivity", "filter": "open"},
    )
    by_id = {c["id"]: c for c in cards}

    def label(card: dict[str, Any]) -> tuple[str, int]:
        return (card.get("name") or "(untitled)").strip(), classify_importance(card.get("labels"))

    columns: list[tuple[str, list[tuple[str, int]]]] = []
    for column in COLUMN_ORDER:
        source_ids = set().union(*(name_to_ids.get(n, set()) for n in COLUMN_SOURCES[column])) or set()

        if column == COMPLETED_COLUMN:
            items = _completed_in_window(board_id, source_ids, by_id, days)
        else:
            picked = [label(c) for c in cards if c["idList"] in source_ids]
            picked.sort(key=lambda item: item[1], reverse=True)  # High importance first
            items = picked
        columns.append((column, items))
    return columns


def _completed_in_window(board_id, done_ids, by_id, days: int) -> list[tuple[str, int]]:
    if not done_ids:
        return []
    since = _dt.datetime.now(_dt.timezone.utc) - _dt.timedelta(days=days)
    actions = trello_request(
        "GET", f"/boards/{board_id}/actions",
        params={"filter": "updateCard", "since": since.strftime("%Y-%m-%dT%H:%M:%SZ"), "limit": "1000"},
    )
    items, seen = [], set()
    for action in actions or []:  # newest first → first hit per card is its latest move
        data = action.get("data", {})
        after = data.get("listAfter") or {}
        card = data.get("card") or {}
        cid = card.get("id")
        if after.get("id") in done_ids and cid and cid not in seen:
            seen.add(cid)
            source = by_id.get(cid, {})
            name = (source.get("name") or card.get("name") or "(untitled)").strip()
            items.append((name, classify_importance(source.get("labels"))))
    return items


def _demo_columns() -> list[tuple[str, list[tuple[str, int]]]]:
    return [
        ("To do", [
            ("Define Q3 roadmap", HIGH),
            ("Draft onboarding guide", MEDIUM),
            ("Clean up backlog labels", LOW),
        ]),
        ("In progress", [
            ("Payment service refactor", HIGH),
            ("Search indexing improvements", MEDIUM),
            ("Migrate CI to new runners", MEDIUM),
            ("Update dependency versions", LOW),
        ]),
        ("Completed", [
            ("Ship auth token refresh", HIGH),
            ("Fix board sync race condition", MEDIUM),
            ("Add PPTX export pipeline", LOW),
            ("Write API documentation", LOW),
        ]),
    ]


# --------------------------------------------------------------------------- #
# Build
# --------------------------------------------------------------------------- #
def build_deck(columns, template_path, output_path, date_str: str) -> Path:
    # Cards with no priority label (NONE) are ignored on the slide.
    columns = [(title, [(n, lvl) for n, lvl in cards if lvl != NONE]) for title, cards in columns]

    prs = Presentation(str(template_path))
    if len(prs.slides) < 2:
        raise SystemExit("Template must have at least 2 slides (title + overview).")

    if not refresh_date(prs.slides[0], date_str):
        print("  note: no DD.MM.YYYY date found on slide 1 to update")

    overview = prs.slides[1]
    _remove_textboxes(overview)
    fill_overview(overview, columns)
    _legend_slide2(overview)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("board", nargs="?", help="Trello board id or exact name")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Base .pptx (default example.pptx)")
    parser.add_argument("--output", "-o", default=str(DEFAULT_OUTPUT), help="Output .pptx path")
    parser.add_argument("--date", default=_dt.date.today().strftime("%d.%m.%Y"),
                        help="Title-slide date (default today, DD.MM.YYYY)")
    parser.add_argument("--days", type=int, default=7,
                        help="Look-back window for the Completed column (default 7)")
    parser.add_argument("--demo", action="store_true", help="Use sample data, no Trello calls")
    args = parser.parse_args(argv)

    if args.demo or not args.board:
        if not args.demo:
            print("No board given — rendering demo data. Pass a board name for live data.")
        columns = _demo_columns()
    else:
        load_dotenv()
        columns = fetch_overview_columns(args.board, days=args.days)

    out = build_deck(columns, args.template, args.output, args.date)
    total = sum(1 for _, cards in columns for _n, lvl in cards if lvl != NONE)
    print(f"Wrote {out}  (date {args.date}, {total} prioritised cards across {len(columns)} columns)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
