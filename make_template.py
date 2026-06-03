#!/usr/bin/env python3
"""Create a local ``template.pptx`` for trello_to_ppt.py.

The template carries the look and feel (slide size, theme, fonts, master
background). ``trello_to_ppt.py`` opens it and *adds* a status slide onto it, so
to rebrand the output you only need to replace ``template.pptx`` with your own
PowerPoint file — no code changes required. Generated ``.pptx`` templates are
ignored by git.

Run this once to regenerate the default template:

    python make_template.py
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches


def main() -> None:
    # Start from python-pptx's default Office theme (Calibri, standard palette)
    # and resize it to widescreen 16:9 so generated slides fill modern displays.
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    out = Path(__file__).with_name("template.pptx")
    prs.save(str(out))
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
