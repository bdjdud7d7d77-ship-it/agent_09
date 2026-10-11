# Slide Design System — "Orange Can-do"

Design reference for building `.pptx` decks (used with the `mk-ppt` skill). Based on the orange business template (`슬라이드 템플릿.jpg`).

## 1. Concept

- Clean, energetic business tone: **white canvas + one strong orange accent**.
- Plenty of white space and thin lines; color is used only to highlight key info.
- Slide ratio **16:9** (13.333 × 7.5 in).

## 2. Color

| Token | HEX | Use |
|---|---|---|
| `primary` | `#F26B1D` | Accent bars, icons, chart fills, highlighted row |
| `primary-dark` | `#D9480F` | Section labels ("CONTENTS A"), hover/emphasis text |
| `primary-light` | `#FDE3D3` | Secondary chart segments, soft backgrounds |
| `red-accent` | `#E8344E` | Second bar in comparisons (use sparingly) |
| `text-dark` | `#222222` | Titles, headings |
| `text-body` | `#666666` | Body text |
| `text-muted` | `#999999` | Captions, English sub-copy |
| `line` | `#E5E5E5` | Dividers, table rules, chart tracks |
| `bg` | `#FFFFFF` | Slide background |
| `chip-dark` | `#333333` | Dark pill labels (top-right tag, "Check" circle) |

Rule: orange ≤ ~20% of the slide area; never put two saturated colors side by side except in charts.

## 3. Typography

| Role | Font | Size | Style |
|---|---|---|---|
| Cover title | Pretendard ExtraBold | 44–54 pt | Bold **italic**, `text-dark` |
| Slide title (header) | Pretendard ExtraBold | 18–20 pt | Bold italic, uppercase English OK |
| Section label | Pretendard Bold | 13–14 pt | `primary-dark` or `text-dark` |
| Body (Korean) | Pretendard / 맑은 고딕 | 11–12 pt | Regular, `text-body` |
| Caption / English sub-copy | Pretendard | 9–10 pt | Regular, `text-muted` |
| Big numbers (%) | Pretendard ExtraBold | 24–28 pt | `text-dark` |

- Fallback font: **맑은 고딕 (Malgun Gothic)**.
- Body ≥ 11 pt; never shrink below this — cut text instead.
- Line spacing 1.2–1.3; left-align body, center-align text under charts/icons.

## 4. Layout

**Cover slide**
- Left half: large rounded "pill" (right side fully rounded) filled with an orange-tinted photo + white headline.
- Right: big italic title + one-line muted subtitle.
- Decorations: thin diagonal orange lines (background), concentric thin circles at the right edge, small orange ring mark.
- Bottom-left: two small pill chips (orange / dark) for presenter or date.

**Content slide header (every body slide)**
- Top-left: small orange tag (rounded rect) with "CONTENTS", then italic slide title, then a thin `line` divider to the right.
- Top-right: small dark pill chip (section or page label).
- Thin orange line along the bottom edge.

**Grid & spacing**
- Margins: 0.5 in left/right, 0.4 in top/bottom.
- Content area starts ~1.1 in from top (below header).
- Use 3-column (cards / charts) or 2-column (text + visual) grids; equal gutters ≥ 0.3 in.

## 5. Components

| Component | Description |
|---|---|
| Icon donut card | Orange partial ring around white circle with icon, downward orange triangle pointer, label + 2-line text below. Use ×3. |
| Pie/donut chart | Orange + `primary-light` segments, white center with "CONTENTS" label, small orange icon circle on top. |
| Percent ring | Thin orange arc showing a value, big % number inside, thin leader line to a label below. |
| Image circle | Large circular photo crop with small orange/white icon circles orbiting it; text items on the left. |
| Table | Header row bold, thin `line` rules; one **highlighted row** filled `primary` with white text; dot-rating column. |
| Bar comparison | Short horizontal bars (`primary`, `red-accent`) with % value at right; bracket line grouping items. |
| Check box | Orange "Check" tab + light gray box for notes/summary, placed at slide bottom. |
| Pill chip | Fully rounded small label, orange or `chip-dark`, white 9 pt text. |

## 6. Do / Don't

- **Do** keep one message per slide, max 3–4 items.
- **Do** align icons, charts and texts on a shared center line.
- **Do** use the same header on every body slide.
- **Don't** use gradients, shadows heavier than a soft blur, or more than 2 fonts.
- **Don't** let text overlap shapes or exceed its box (see `ppt-rule-v1`).
