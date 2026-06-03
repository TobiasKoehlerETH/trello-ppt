#!/usr/bin/env python3
"""Generate a PowerPoint weekly-status slide for a Trello kanban board.

The slide answers three questions at a glance:

* **What got done?**  Cards moved into the "Done" list in the last 7 days.
* **What matters most now?**  The single highest-priority open card ("Top
  priority").
* **How important is each item?**  A Red / Yellow / Green dot, derived from each
  card's Trello labels (label *name* such as High/Medium/Low, or label *colour*
  red / orange-yellow / green).

Credentials live in ``.env`` — see :mod:`trello_integration`. If a local
``template.pptx`` exists, the slide is drawn onto it; otherwise the script falls
back to a blank widescreen deck.

Examples
--------
    # By board id or exact name
    python trello_to_ppt.py 5f2a9c1b3e4d5a6b7c8d9e0f
    python trello_to_ppt.py "Product Roadmap"

    # Tweak the window, the Done list name, template, and output
    python trello_to_ppt.py "Product Roadmap" --days 7 --done-list "Done" \
        --template brand.pptx -o output/status.pptx

    # Render built-in sample data, no Trello calls (handy for previewing)
    python trello_to_ppt.py --demo
"""

from __future__ import annotations

import argparse
import datetime as _dt
import re
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from trello_integration import load_dotenv, trello_request


HERE = Path(__file__).parent
DEFAULT_TEMPLATE = HERE / "template.pptx"
DEFAULT_OUTPUT = HERE / "output" / "board_summary.pptx"

# Slide geometry (widescreen 16:9).
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.5)

# Base palette.
INK = RGBColor(0x17, 0x24, 0x3E)        # near-black navy, body text
MUTED = RGBColor(0x6B, 0x72, 0x80)      # grey, subtitle / hints
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PANEL_BG = RGBColor(0xF4, 0xF5, 0xF7)   # light grey panel

# Importance levels and their Red / Yellow / Green styling.
HIGH, MEDIUM, LOW, NONE = 3, 2, 1, 0
RED = RGBColor(0xE0, 0x3E, 0x36)
AMBER = RGBColor(0xF2, 0xB1, 0x14)
GREEN = RGBColor(0x3F, 0xA9, 0x4A)
GREY = RGBColor(0x9A, 0xA0, 0xA6)
LEVELS: dict[int, dict[str, Any]] = {
    HIGH: {"name": "High", "color": RED, "text": WHITE},
    MEDIUM: {"name": "Medium", "color": AMBER, "text": INK},
    LOW: {"name": "Low", "color": GREEN, "text": WHITE},
    NONE: {"name": "Unrated", "color": GREY, "text": WHITE},
}

# Importance inferred from label names first, then label colours.
_NAME_HIGH = ("high", "urgent", "critical", "blocker", "p0", "p1")
_NAME_MEDIUM = ("medium", "med", "normal", "p2")
_NAME_LOW = ("low", "minor", "trivial", "nice", "p3", "p4")
_COLOR_LEVEL = {"red": HIGH, "orange": MEDIUM, "yellow": MEDIUM, "green": LOW, "lime": LOW}

# A Trello object id is a 24-char hex string; anything else is treated as a name.
_ID_RE = re.compile(r"^[0-9a-fA-F]{24}$")


# --------------------------------------------------------------------------- #
# Importance helpers
# --------------------------------------------------------------------------- #
def classify_importance(labels: list[dict[str, Any]] | None) -> int:
    """Map a card's Trello labels to an importance level (HIGH/MEDIUM/LOW/NONE)."""
    level = NONE
    for label in labels or []:
        name = (label.get("name") or "").strip().lower()
        color = (label.get("color") or "").strip().lower()
        if any(token in name for token in _NAME_HIGH):
            current = HIGH
        elif any(token in name for token in _NAME_MEDIUM):
            current = MEDIUM
        elif any(token in name for token in _NAME_LOW):
            current = LOW
        else:
            current = _COLOR_LEVEL.get(color.split("_")[0], NONE)
        level = max(level, current)
    return level


# --------------------------------------------------------------------------- #
# Data: pull the board state from Trello
# --------------------------------------------------------------------------- #
def resolve_board(board: str) -> tuple[str, str, str]:
    """Return ``(board_id, name, url)`` for a board id or case-insensitive name."""
    if _ID_RE.match(board):
        data = trello_request("GET", f"/boards/{board}", params={"fields": "name,url"})
        return board, data.get("name", board), data.get("url", "")

    boards = trello_request(
        "GET",
        "/members/me/boards",
        params={"fields": "name,url", "filter": "open"},
    )
    for entry in boards:
        if entry.get("name", "").lower() == board.lower():
            return entry["id"], entry["name"], entry.get("url", "")

    available = ", ".join(repr(b.get("name", "")) for b in boards) or "none found"
    raise SystemExit(f"No open board named {board!r}. Your open boards: {available}")


def _parse_trello_date(value: str | None) -> _dt.datetime | None:
    if not value:
        return None
    try:
        return _dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def fetch_board_summary(board: str, *, days: int = 7, done_list: str = "Done") -> dict[str, Any]:
    """Build the weekly-status summary for ``board``.

    * ``completed`` — cards moved into the Done list within the last ``days``.
    * ``top_priority`` — the highest-importance open card not already in Done.
    * ``list_counts`` — open card count per list (the ongoing board state).
    """
    board_id, name, url = resolve_board(board)
    now = _dt.datetime.now(_dt.timezone.utc)
    since = now - _dt.timedelta(days=days)

    lists = trello_request(
        "GET",
        f"/boards/{board_id}/lists",
        params={"fields": "name", "filter": "open"},
    )
    list_names = {lst["id"]: lst.get("name", "") for lst in lists}
    done_ids = {lid for lid, lname in list_names.items() if lname.strip().lower() == done_list.lower()}

    cards = trello_request(
        "GET",
        f"/boards/{board_id}/cards",
        params={"fields": "name,due,idList,dateLastActivity,labels", "filter": "open"},
    )
    by_id = {card["id"]: card for card in cards}

    # Per-list open counts in board order (ongoing state).
    counts = {lst["id"]: 0 for lst in lists}
    for card in cards:
        if card["idList"] in counts:
            counts[card["idList"]] += 1
    list_counts = [(list_names[lst["id"]], counts[lst["id"]]) for lst in lists]

    # Completed in the last `days`: card-move actions whose destination is Done.
    completed: list[dict[str, Any]] = []
    seen: set[str] = set()
    if done_ids:
        actions = trello_request(
            "GET",
            f"/boards/{board_id}/actions",
            params={
                "filter": "updateCard",
                "since": since.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "limit": "1000",
            },
        )
        # Actions arrive newest-first, so the first hit per card is its latest move.
        for action in actions or []:
            data = action.get("data", {})
            after = data.get("listAfter") or {}
            card = data.get("card") or {}
            card_id = card.get("id")
            if after.get("id") not in done_ids or not card_id or card_id in seen:
                continue
            seen.add(card_id)
            source = by_id.get(card_id, {})
            completed.append(
                {
                    "name": source.get("name") or card.get("name") or "(untitled)",
                    "date": _parse_trello_date(action.get("date")),
                    "level": classify_importance(source.get("labels")),
                }
            )
        completed.sort(key=lambda item: item["date"] or now, reverse=True)

    # Top priority: highest-importance open card that is NOT already in Done.
    def priority_key(card: dict[str, Any]) -> tuple:
        level = classify_importance(card.get("labels"))
        due = _parse_trello_date(card.get("due"))
        activity = _parse_trello_date(card.get("dateLastActivity")) or now
        # Higher level first; then soonest due (None last); then most recent activity.
        return (level, due is not None, -(due or now.replace(year=9999)).timestamp(), activity.timestamp())

    candidates = [c for c in cards if c["idList"] not in done_ids]
    top_priority = None
    if candidates:
        best = max(candidates, key=priority_key)
        top_priority = {
            "name": best.get("name") or "(untitled)",
            "list_name": list_names.get(best["idList"], ""),
            "due": _parse_trello_date(best.get("due")),
            "level": classify_importance(best.get("labels")),
        }

    return {
        "board_name": name,
        "board_url": url,
        "generated_at": _dt.datetime.now(),
        "window_days": days,
        "done_list": done_list,
        "completed": completed,
        "top_priority": top_priority,
        "list_counts": list_counts,
        "open_total": len(cards),
    }


# --------------------------------------------------------------------------- #
# Rendering
# --------------------------------------------------------------------------- #
def _blank_layout(prs: Presentation):
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() == "blank":
            return layout
    return min(prs.slide_layouts, key=lambda layout: len(layout.placeholders))


def _style_box(shape, fill: RGBColor, *, line: RGBColor | None = None) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    shape.shadow.inherit = False


def _text(box_or_frame, text: str, *, size: int, color: RGBColor, bold: bool = False,
          italic: bool = False, align=None, new: bool = False):
    tf = box_or_frame.text_frame if hasattr(box_or_frame, "text_frame") else box_or_frame
    para = tf.add_paragraph() if new else tf.paragraphs[0]
    if align is not None:
        para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return para, run


def _fmt_date(value: _dt.datetime | None) -> str:
    return value.strftime("%b %d") if value else ""


def _dot(slide, left, top, size, color: RGBColor):
    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
    _style_box(dot, color)
    return dot


def _legend(slide) -> None:
    box = slide.shapes.add_textbox(SLIDE_W - MARGIN - Inches(5.2), Inches(0.5), Inches(5.2), Inches(0.4))
    para = box.text_frame.paragraphs[0]
    para.alignment = PP_ALIGN.RIGHT
    for level, gap in ((HIGH, "   "), (MEDIUM, "   "), (LOW, "")):
        meta = LEVELS[level]
        dot = para.add_run()
        dot.text = "● "
        dot.font.size = Pt(12)
        dot.font.color.rgb = meta["color"]
        label = para.add_run()
        label.text = f"{meta['name']}{gap}"
        label.font.size = Pt(12)
        label.font.color.rgb = MUTED


def _chip(slide, right_edge, top, level: int):
    """A small importance chip aligned to ``right_edge``; returns its left x."""
    meta = LEVELS[level]
    width = Inches(1.15)
    left = right_edge - width
    chip = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, Inches(0.34))
    _style_box(chip, meta["color"])
    tf = chip.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _text(tf, meta["name"].upper(), size=11, color=meta["text"], bold=True, align=PP_ALIGN.CENTER)
    return left


def build_status_slide(
    summary: dict[str, Any],
    template_path: str | Path | None = DEFAULT_TEMPLATE,
    output_path: str | Path = DEFAULT_OUTPUT,
    max_completed: int = 6,
) -> Path:
    if template_path and Path(template_path).exists():
        prs = Presentation(str(template_path))
    else:
        prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    slide = prs.slides.add_slide(_blank_layout(prs))

    content_w = SLIDE_W - 2 * MARGIN

    # Title + subtitle + legend.
    title = slide.shapes.add_textbox(MARGIN, Inches(0.3), content_w - Inches(5.2), Inches(0.7))
    _text(title, f"{summary['board_name']} — Weekly Status", size=28, color=INK, bold=True)

    when = summary["generated_at"].strftime("%Y-%m-%d %H:%M")
    sub = slide.shapes.add_textbox(MARGIN, Inches(0.98), content_w, Inches(0.4))
    _text(
        sub,
        f"{len(summary['completed'])} cards completed in the last {summary['window_days']} days"
        f"  ·  {summary['open_total']} open  ·  generated {when}",
        size=13, color=MUTED,
    )
    _legend(slide)

    # ---- Top priority banner ------------------------------------------------
    top = summary.get("top_priority")
    by = Inches(1.5)
    bh = Inches(1.3)
    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, MARGIN, by, content_w, bh)
    _style_box(panel, PANEL_BG)
    level = top["level"] if top else NONE
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN, by, Inches(0.16), bh)
    _style_box(accent, LEVELS[level]["color"])

    label_box = slide.shapes.add_textbox(MARGIN + Inches(0.4), by + Inches(0.13), Inches(4.0), Inches(0.35))
    _text(label_box, "TOP PRIORITY", size=12, color=MUTED, bold=True)

    if top:
        _chip(slide, SLIDE_W - MARGIN - Inches(0.25), by + Inches(0.17), level)
        name_box = slide.shapes.add_textbox(MARGIN + Inches(0.4), by + Inches(0.46), content_w - Inches(2.0), Inches(0.6))
        name_box.text_frame.word_wrap = True
        _text(name_box, top["name"], size=22, color=INK, bold=True)
        meta_bits = [f"In {top['list_name']}"] if top["list_name"] else []
        if top["due"]:
            meta_bits.append(f"due {_fmt_date(top['due'])}")
        meta_bits.append(f"{LEVELS[level]['name']} importance")
        meta_box = slide.shapes.add_textbox(MARGIN + Inches(0.4), by + Inches(0.95), content_w - Inches(1.0), Inches(0.3))
        _text(meta_box, "   ·   ".join(meta_bits), size=12, color=MUTED)
    else:
        name_box = slide.shapes.add_textbox(MARGIN + Inches(0.4), by + Inches(0.5), content_w - Inches(1.0), Inches(0.5))
        _text(name_box, "No open cards to prioritise.", size=18, color=MUTED)

    # ---- Completed in the last N days --------------------------------------
    heading = slide.shapes.add_textbox(MARGIN, Inches(3.1), content_w, Inches(0.4))
    _text(
        heading,
        f"Completed in the last {summary['window_days']} days  ·  {len(summary['completed'])}",
        size=16, color=INK, bold=True,
    )

    completed = summary["completed"]
    row_top = Inches(3.66)
    row_h = Inches(0.46)
    if not completed:
        empty = slide.shapes.add_textbox(MARGIN + Inches(0.05), row_top, content_w, Inches(0.4))
        _text(empty, "No cards completed in this window.", size=13, color=MUTED, italic=True)
    else:
        for index, item in enumerate(completed[:max_completed]):
            y = row_top + row_h * index
            meta = LEVELS[item["level"]]
            _dot(slide, MARGIN + Inches(0.05), y + Inches(0.07), Inches(0.2), meta["color"])
            name_box = slide.shapes.add_textbox(MARGIN + Inches(0.4), y, content_w - Inches(2.3), Inches(0.4))
            name_box.text_frame.word_wrap = False
            _text(name_box, item["name"], size=13, color=INK)
            date_box = slide.shapes.add_textbox(SLIDE_W - MARGIN - Inches(1.7), y, Inches(1.7), Inches(0.4))
            _text(date_box, _fmt_date(item["date"]), size=12, color=MUTED, align=PP_ALIGN.RIGHT)
        if len(completed) > max_completed:
            y = row_top + row_h * max_completed
            more = slide.shapes.add_textbox(MARGIN + Inches(0.4), y, content_w, Inches(0.4))
            _text(more, f"+{len(completed) - max_completed} more", size=12, color=MUTED, italic=True)

    # ---- Ongoing board state strip -----------------------------------------
    bits = [f"{lname} {count}" for lname, count in summary["list_counts"][:7] if lname]
    if bits:
        strip = slide.shapes.add_textbox(MARGIN, Inches(6.78), content_w, Inches(0.4))
        _text(strip, "Board now:   " + "    ·    ".join(bits), size=11, color=MUTED)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# --------------------------------------------------------------------------- #
# Demo data (no API calls)
# --------------------------------------------------------------------------- #
def _demo_summary() -> dict[str, Any]:
    today = _dt.datetime.now()
    days_ago = lambda n: today - _dt.timedelta(days=n)  # noqa: E731
    completed = [
        {"name": "Ship auth token refresh", "date": days_ago(1), "level": HIGH},
        {"name": "Fix board sync race condition", "date": days_ago(2), "level": HIGH},
        {"name": "Add PPTX export pipeline", "date": days_ago(2), "level": MEDIUM},
        {"name": "Write API documentation", "date": days_ago(4), "level": LOW},
        {"name": "Tidy project scaffolding", "date": days_ago(5), "level": LOW},
        {"name": "Set up CI cache", "date": days_ago(6), "level": MEDIUM},
        {"name": "Archive stale spike cards", "date": days_ago(6), "level": NONE},
    ]
    return {
        "board_name": "Product Roadmap (Demo)",
        "board_url": "https://trello.com/b/demo",
        "generated_at": today,
        "window_days": 7,
        "done_list": "Done",
        "completed": completed,
        "top_priority": {
            "name": "Resolve payment webhook outage",
            "list_name": "In Progress",
            "due": today + _dt.timedelta(days=1),
            "level": HIGH,
        },
        "list_counts": [("Backlog", 9), ("To Do", 5), ("In Progress", 3), ("Review", 2), ("Done", 7)],
        "open_total": 19,
    }


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("board", nargs="?", help="Trello board id or exact board name")
    parser.add_argument("--days", type=int, default=7, help="Look-back window for completed cards (default 7)")
    parser.add_argument("--done-list", default="Done", help='Name of the completed list (default "Done")')
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Optional local base .pptx template")
    parser.add_argument("--output", "-o", default=str(DEFAULT_OUTPUT), help="Output .pptx path")
    parser.add_argument("--max-completed", type=int, default=6, help="Max completed cards listed (default 6)")
    parser.add_argument("--demo", action="store_true", help="Render built-in sample data without calling Trello")
    args = parser.parse_args(argv)

    if args.demo:
        summary = _demo_summary()
    else:
        if not args.board:
            parser.error("a board id/name is required (or pass --demo)")
        load_dotenv()
        summary = fetch_board_summary(args.board, days=args.days, done_list=args.done_list)

    out = build_status_slide(summary, args.template, args.output, args.max_completed)
    top = summary.get("top_priority")
    print(
        f"Wrote {out}\n"
        f"  completed (last {summary['window_days']}d): {len(summary['completed'])}\n"
        f"  top priority: {top['name'] if top else 'none'}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
