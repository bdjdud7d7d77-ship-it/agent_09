---
name: mk-ppt
description: Build or edit PowerPoint (.pptx) decks in this project with python-pptx — slides, text, tables, charts, images, shapes, notes. Use this whenever the user asks for a PPT, 발표자료, 슬라이드, 프레젠테이션, deck, or .pptx file, wants research notes or a report turned into slides, or wants an existing .pptx changed or inspected, even if they don't say "python-pptx".
---

# mk-ppt

Create `.pptx` decks with python-pptx through a reproducible generator script, following this project's layout rules.

## Project conventions (why they matter)

The project keeps work grouped by one kebab-case `<topic>` and expects every deliverable to be regenerable:

| What | Where |
|---|---|
| Source notes (if any) | `research/<topic>/*.md` |
| Generator script | `scripts/make_<topic>_ppt.py` (snake_case topic in the file name) |
| Output deck | `reports/<topic>/<topic_snake>_slides.pptx` |

- Never overwrite an existing file. The helper `unique_path()` picks `_v2`, `_v3`, ... automatically. If the generator script name is taken, save the new script as `make_<topic>_ppt_v2.py`.
- Scripts must run from any directory, so resolve paths from `__file__`, not the cwd.
- Use ASCII file and folder names; Korean names break in the Windows shell and git.
- Slide text follows the user's language (usually Korean). Use a Korean-capable font (default `맑은 고딕`). The helpers set the East Asian font slot too, because setting only `font.name` leaves Hangul in the theme font.

## Workflow

1. **Check the library.** Run `python -c "import pptx; print(pptx.__version__)"`. If it fails, tell the user and install with `python -m pip install python-pptx`.
2. **Plan the deck.** Decide the topic slug, the number of slides, and each slide's title, content type (bullets / table / chart / image / two-column) and speaker notes. If source notes exist in `research/<topic>/`, build from them. Keep about 3–6 bullets per slide; dense slides are unreadable when projected.
3. **Write the generator** at `scripts/make_<topic>_ppt.py`. Import the bundled helpers instead of re-writing boilerplate:

   ```python
   import os, sys
   ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
   sys.path.insert(0, os.path.join(ROOT, ".claude", "skills", "mk-ppt", "scripts"))
   from pptx_helpers import *  # new_deck, add_title_slide, add_bullet_slide, ...

   prs = new_deck()  # 16:9
   add_title_slide(prs, "제목", "부제")
   add_bullet_slide(prs, "개요", ["항목 1", ("세부 항목", 1), "항목 2"], notes="발표자 노트")
   add_table_slide(prs, "비교", [["구분", "A", "B"], ["값", "1", "2"]])
   add_chart_slide(prs, "추이", ["2023", "2024"], {"매출": [10, 14]}, kind="column")
   out = save(prs, os.path.join(ROOT, "reports", "<topic>"), "<topic_snake>_slides")
   print(out)
   ```

   For anything the helpers don't cover (freeforms, connectors, gradients, hyperlinks, click actions, editing existing decks), use python-pptx directly. See `references/python-pptx-api.md`.
4. **Run it.** `python scripts/make_<topic>_ppt.py`
5. **Verify.** `python .claude/skills/mk-ppt/scripts/inspect_pptx.py <output.pptx>` prints each slide's shapes and text. Check the slide count, that no placeholder is left empty ("Click to add..." boxes show up in the projector), and that text isn't obviously overflowing (long bullets → split the slide).
6. **Report.** Tell the user the output path, the slide count, and a one-line outline. Add new files to `CLAUDE.md` → "Project structure" if a new folder kind appeared.

## Editing an existing deck

Open with `Presentation(path)`, change it, then save under a new name with `save(...)` (never in place). Use `inspect_pptx.py` first to learn the shape names and indices.

## Limits

python-pptx can't do animations, transitions, SmartArt, PDF/image export, or legacy `.ppt` files. There's no public API to delete, reorder, or duplicate slides; `references/python-pptx-api.md` has the XML workarounds. Say so plainly instead of faking it.
