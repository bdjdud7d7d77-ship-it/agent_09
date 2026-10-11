---
name: mk-ppt
description: Create PowerPoint (.pptx) presentations in this project with python-pptx — title, section, bullet, two-column, table, chart, KPI-stat, image and closing slides with speaker notes, Korean fonts and a consistent theme. Use this skill whenever the user asks for a PPT, PowerPoint, 발표 자료, 슬라이드, 프레젠테이션, deck, or wants research notes, a report or any content turned into slides — even if they don't say ".pptx" or "python-pptx". Also use it when the user asks what python-pptx can or cannot do for a deck.
---

# mk-ppt — build .pptx decks with python-pptx

This skill turns content (a topic, the user's notes, or files in `research/` / `report/`) into a clean 16:9 deck. The heavy lifting is done by a bundled script so every deck has the same look and the same Korean-font handling; your job is to plan the story, write a JSON spec, run the script, and check the result.

## Step 0 — read the PPT rules first

Before anything else in this skill, read and follow `.claude/rules/ppt-rule-v1/ppt-rule-v1.md`. It overrides this skill wherever they differ. Order of priority for decks:

1. `.claude/rules/ppt-rule-v1/ppt-rule-v1.md` (and `CLAUDE.md`)
2. The design `.md` the user provides (e.g. `references/ppt/slide-design-system.md`)
3. This skill's defaults (themes, font, layouts)

In short: 8 slides by default, follow the user's design file (ask for one if none was given), and never let text overflow or overlap — render and fix before reporting.

## Project rules that apply here

This project has its own rules in `CLAUDE.md`. They take priority, and the important ones for decks are:

- **Todo list first, then wait for approval.** Before building, report the plan: slide outline (title of each slide + its type), source files, output path, and open decisions (theme, slide count, audience). Only build after the user approves.
- **Save decks in `slides/`.** Never leave a deck in the project root.
- **Never overwrite.** The script automatically appends `_v2`, `_v3`, … if the file exists, so you can always pass the plain name.
- **Respond in Korean**, and write slide text in Korean unless the user asks otherwise.
- The JSON spec is a temporary file — write it to the scratchpad directory, not the project.

## Workflow

1. **Gather the content and the design file.** Read the source material the user points to. If they only give a topic, ask whether to research it first or use only what they provided — slides with invented numbers are worse than no slides. Read the user's design `.md`; if none was given, ask for one and use this skill's default theme only if the user says so (ppt-rule-v1 §2).
2. **Plan the story.** One message per slide. **8 slides by default** unless the user asks for another count (ppt-rule-v1 §1): title → (agenda or section) → content slides → summary/next steps → closing. Pick the slide type that fits each message (see table below): numbers over time → `chart`; 2–4 headline numbers → `stats`; comparisons → `two_column` or `table`; everything else → `bullets`.
3. **Report the outline as the todo list** and wait for approval (project rule).
4. **Write the spec** (JSON, UTF-8) to the scratchpad.
5. **Run the script** from the project root:
   ```bash
   python .claude/skills/mk-ppt/scripts/build_pptx.py <scratchpad>/spec.json
   ```
   It saves to `slides/<filename>` (or `--output <path>`), re-opens the file to verify it, and prints one line per slide.
6. **Check visually — required (ppt-rule-v1 §3).** PowerPoint is installed on this machine, so export every slide to PNG and view it with the Read tool:
   ```powershell
   New-Item -ItemType Directory -Force "<scratchpad>\png" | Out-Null
   $app = New-Object -ComObject PowerPoint.Application
   $pres = $app.Presentations.Open("<full path>.pptx", $true, $false, $false)
   foreach ($sl in $pres.Slides) { $sl.Export("<scratchpad>\png\s$($sl.SlideIndex).png", "PNG", 1920, 1080) }
   $pres.Close(); $app.Quit()
   ```
   Look for text overflowing its box, text/shape overlap, words broken across lines, crowded slides and unreadable labels; fix the spec (shorten text, split the slide) and rebuild until none remain. Ignore very faint thin streaks running from text down through light boxes — that is a PowerPoint PNG-export artifact, not part of the file.
7. **Report** the saved path, the slide list, and anything the user asked for that python-pptx could not do.

## Spec format

```json
{
  "title": "Deck title (document property)",
  "author": "춘식",
  "filename": "AI_발전_발표.pptx",
  "theme": "navy",
  "font": "Noto Sans KR",
  "slides": [ { "type": "...", "notes": "optional speaker notes", ... } ]
}
```

`theme`: `navy` (default), `charcoal`, `green`, `orange`, `pink`. `font` defaults to Noto Sans KR and is applied to Latin, East Asian and complex-script text (python-pptx alone only sets the Latin font, which breaks Korean rendering).

| type | fields | use for |
|---|---|---|
| `title` | `title`, `subtitle`, `author`, `date` | first slide |
| `section` | `title`, `subtitle`, `number` (e.g. "01") | chapter dividers |
| `bullets` | `title`, `bullets`: list of strings or `{"text", "level": 0/1, "bold"}` | general content |
| `two_column` | `title`, `left`/`right`: `{"heading", "bullets"}` | pros/cons, before/after, A vs B |
| `table` | `title`, `columns`, `rows` (list of lists), `col_widths` (relative, optional), `caption` | structured data, timelines |
| `chart` | `title`, `chart_type`, `categories`, `series`: `[{"name", "values"}]`, `number_format`, `legend`, `data_labels`, `chart_title`, `takeaways` (bullets shown beside the chart), `source` | numeric trends and shares |
| `stats` | `title`, `items`: up to 4 `{"value", "label"}`, `caption` | headline KPIs |
| `image` | `title`, `path` (relative to project root or absolute), `caption` | screenshots, diagrams |
| `closing` | `title` (default 감사합니다), `subtitle` | last slide |

`chart_type`: `column`, `column_stacked`, `bar`, `bar_stacked`, `line`, `line_plain`, `pie`, `doughnut`, `area`, `radar`. For `pie`/`doughnut` use one series; labels show percentages by default (`"show_percentage": false` shows values). Use `number_format` like `"0%"`, `"#,##0"`, `"0.0"`.

## Writing good slides

- **Keep text short.** Up to ~6 bullets per slide and ~40 Korean characters per bullet (~22 in a `two_column` column, ~14 in chart `takeaways`). Font size shrinks automatically as bullets increase, but beyond 8 it gets hard to read — split the slide instead.
- **Korean wraps mid-word.** PowerPoint breaks Hangul at any character, so a line that is slightly too long ends as "난/제". In narrow boxes, size text to fit on one line rather than relying on wrapping.
- **`stats` tiles are small.** Keep `value` to ~7 characters with 4 tiles (e.g. "1.7억", "$5,817억", "73:23") and `label` to ~12 characters; move the rest to `caption` or `notes`. Long values are shrunk automatically, but a short value reads better.
- **Titles state the point**, not the topic: "기업 AI 도입률 88% 돌파" beats "AI 도입 현황".
- **Put detail in `notes`**, so the slide stays clean and the presenter still has the full story.
- **Cite sources** for numbers via `source` (charts) or `caption` (tables, stats).
- Tables over ~12 rows or 6 columns become unreadable; split or summarise.

## What python-pptx cannot do

Tell the user up front if they ask for any of these, and offer an alternative:

- Animations and slide transitions (only via raw XML — fragile; suggest adding them in PowerPoint)
- Deleting, reordering or duplicating slides through the official API (possible with XML manipulation; rebuilding from the spec is simpler here)
- Rendering slides to images/PDF by itself (on this machine use PowerPoint COM as in step 6)
- SmartArt, 3D/stock/surface charts as new charts, opening legacy `.ppt`, running or editing VBA macros
- Fine-grained shadow/glow/3D effects (only "inherit or not")

For anything the script does not cover (e.g. editing an existing deck, placeholders of a corporate template, hyperlinks, connectors, freeforms, video, OLE objects), write python-pptx code directly. `references/python_pptx_api.md` lists the API by area with short snippets — read it when you need something outside the spec format.
