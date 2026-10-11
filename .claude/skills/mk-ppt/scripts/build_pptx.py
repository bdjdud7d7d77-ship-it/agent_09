"""Build a .pptx deck from a JSON spec with python-pptx.

Usage (run from anywhere; relative paths resolve against the project root):
    python .claude/skills/mk-ppt/scripts/build_pptx.py spec.json
    python .claude/skills/mk-ppt/scripts/build_pptx.py spec.json --output slides/my_deck.pptx

The output is never overwritten: if the target exists, _v2, _v3, ... is appended.
See SKILL.md for the spec format.
"""

import argparse
import json
import sys
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

PROJECT_ROOT = Path(__file__).resolve().parents[4]
DEFAULT_DIR = PROJECT_ROOT / "slides"

SLIDE_W, SLIDE_H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.6)
CONTENT_TOP = Inches(1.55)
CONTENT_H = SLIDE_H - CONTENT_TOP - Inches(0.7)
CONTENT_W = SLIDE_W - 2 * MARGIN

THEMES = {
    "navy": {
        "primary": "1F3864", "accent": "2E75B6", "light": "EEF3FA",
        "text": "222222", "muted": "7F7F7F", "on_primary": "FFFFFF",
        "series": ["1F3864", "2E75B6", "7FA7D6", "F2A541", "5B9279", "C0504D"],
    },
    "charcoal": {
        "primary": "2D2D2D", "accent": "E07A1F", "light": "F4F4F4",
        "text": "222222", "muted": "808080", "on_primary": "FFFFFF",
        "series": ["2D2D2D", "E07A1F", "8C8C8C", "F2B880", "4F81BD", "9BBB59"],
    },
    "green": {
        "primary": "1E5631", "accent": "4C9A2A", "light": "EEF5EC",
        "text": "222222", "muted": "7F7F7F", "on_primary": "FFFFFF",
        "series": ["1E5631", "4C9A2A", "A4DE02", "F2A541", "2E75B6", "7F7F7F"],
    },
    "orange": {
        "primary": "C55A11", "accent": "ED7D31", "light": "FBE5D6",
        "text": "222222", "muted": "7F7F7F", "on_primary": "FFFFFF", "subtitle": "FDE9D9", "meta": "F8CBAD",
        "series": ["C55A11", "ED7D31", "F4B183", "7F7F7F", "2E75B6", "5B9279"],
    },
    "pink": {
        "primary": "C2185B", "accent": "EC407A", "light": "FCE4EC",
        "text": "222222", "muted": "7F7F7F", "on_primary": "FFFFFF", "subtitle": "F8BBD0", "meta": "F48FB1",
        "series": ["C2185B", "EC407A", "F48FB1", "7F7F7F", "8E24AA", "5B9279"],
    },
}

CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column_stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "bar_stacked": XL_CHART_TYPE.BAR_STACKED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "line_plain": XL_CHART_TYPE.LINE,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "area": XL_CHART_TYPE.AREA,
    "radar": XL_CHART_TYPE.RADAR_MARKERS,
}


def rgb(hex_str):
    return RGBColor.from_string(hex_str)


class Builder:
    def __init__(self, spec):
        self.spec = spec
        self.theme = THEMES.get(spec.get("theme", "navy"), THEMES["navy"])
        self.font = spec.get("font", "Noto Sans KR")
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = SLIDE_W, SLIDE_H
        self.blank = self.prs.slide_layouts[6]
        self.page = 0

    # ---------- low-level helpers ----------
    def set_font(self, font, size=None, bold=None, color=None):
        """Apply the deck font to Latin, East Asian and complex scripts.

        python-pptx's font.name only sets <a:latin>, so Korean text would fall
        back to the theme font; <a:ea>/<a:cs> are added here explicitly.
        """
        font.name = self.font
        rpr = font._rPr
        latin = rpr.find(qn("a:latin"))
        prev = latin
        for tag in ("a:ea", "a:cs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = rpr.makeelement(qn(tag), {})
                prev.addnext(el)
            el.set("typeface", self.font)
            prev = el
        if size is not None:
            font.size = Pt(size)
        if bold is not None:
            font.bold = bold
        if color is not None:
            font.color.rgb = rgb(color)

    def textbox(self, slide, left, top, width, height, text, size=18, bold=False,
                color=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, fill=None):
        box = slide.shapes.add_textbox(left, top, width, height)
        if fill:
            box.fill.solid()
            box.fill.fore_color.rgb = rgb(fill)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        lines = text if isinstance(text, list) else [text]
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            run = p.add_run()
            run.text = str(line)
            self.set_font(run.font, size, bold, color or self.theme["text"])
        return box

    def rect(self, slide, left, top, width, height, color, shape=MSO_SHAPE.RECTANGLE):
        shp = slide.shapes.add_shape(shape, left, top, width, height)
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(color)
        shp.line.fill.background()
        # Drop the theme style reference: it adds effects that PowerPoint can
        # render as faint artifacts on flat background boxes.
        style = shp._element.find(qn("p:style"))
        if style is not None:
            shp._element.remove(style)
        return shp

    def new_slide(self, data):
        slide = self.prs.slides.add_slide(self.blank)
        self.page += 1
        if data.get("notes"):
            slide.notes_slide.notes_text_frame.text = data["notes"]
        return slide

    def header(self, slide, title):
        self.textbox(slide, MARGIN, Inches(0.45), CONTENT_W, Inches(0.8), title,
                     size=30, bold=True, color=self.theme["primary"],
                     anchor=MSO_ANCHOR.MIDDLE)
        self.rect(slide, MARGIN, Inches(1.25), Inches(1.2), Inches(0.06), self.theme["accent"])
        self.textbox(slide, SLIDE_W - MARGIN - Inches(1), SLIDE_H - Inches(0.55),
                     Inches(1), Inches(0.35), str(self.page), size=11,
                     color=self.theme["muted"], align=PP_ALIGN.RIGHT)

    def bullets(self, slide, left, top, width, height, items, base_size=None, fill=None):
        items = [i if isinstance(i, dict) else {"text": i, "level": 0} for i in items]
        n = len(items)
        size = base_size or (28 if n <= 3 else 24 if n <= 5 else 20 if n <= 7 else 18 if n <= 9 else 16)
        box = slide.shapes.add_textbox(left, top, width, height)
        if fill:
            # A transparent text box over a filled box renders with faint
            # streaks in PowerPoint; matching the fill avoids them.
            box.fill.solid()
            box.fill.fore_color.rgb = rgb(fill)
        tf = box.text_frame
        tf.word_wrap = True
        for i, item in enumerate(items):
            level = int(item.get("level", 0))
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(10 if level == 0 else 4)
            marker = "■ " if level == 0 else "– "
            r1 = p.add_run()
            r1.text = ("    " * level) + marker
            self.set_font(r1.font, size - 6 if level == 0 else size - 2, False,
                          self.theme["accent"] if level == 0 else self.theme["muted"])
            r2 = p.add_run()
            r2.text = item["text"]
            self.set_font(r2.font, size if level == 0 else size - 3, item.get("bold", False),
                          self.theme["text"])
        return box

    # ---------- slide types ----------
    def s_title(self, d):
        s = self.new_slide(d)
        self.rect(s, 0, 0, SLIDE_W, SLIDE_H, self.theme["primary"])
        self.rect(s, MARGIN, Inches(3.55), Inches(1.6), Inches(0.08), self.theme["accent"])
        self.textbox(s, MARGIN, Inches(1.6), CONTENT_W, Inches(1.9), d["title"], size=44,
                     bold=True, color=self.theme["on_primary"], anchor=MSO_ANCHOR.BOTTOM)
        if d.get("subtitle"):
            self.textbox(s, MARGIN, Inches(3.85), CONTENT_W, Inches(1), d["subtitle"],
                         size=22, color=self.theme.get("subtitle", "D9E2F3"))
        meta = " · ".join(x for x in (d.get("author"), d.get("date")) if x)
        if meta:
            self.textbox(s, MARGIN, SLIDE_H - Inches(1.2), CONTENT_W, Inches(0.5), meta,
                         size=14, color=self.theme.get("meta", "BFCBE0"))

    def s_section(self, d):
        s = self.new_slide(d)
        self.rect(s, 0, 0, Inches(0.35), SLIDE_H, self.theme["accent"])
        self.rect(s, Inches(0.35), 0, SLIDE_W - Inches(0.35), SLIDE_H, self.theme["light"])
        if d.get("number"):
            self.textbox(s, MARGIN + Inches(0.3), Inches(2.0), Inches(3), Inches(1),
                         str(d["number"]), size=48, bold=True, color=self.theme["accent"])
        self.textbox(s, MARGIN + Inches(0.3), Inches(3.0), CONTENT_W, Inches(1.2), d["title"],
                     size=38, bold=True, color=self.theme["primary"])
        if d.get("subtitle"):
            self.textbox(s, MARGIN + Inches(0.3), Inches(4.2), CONTENT_W, Inches(0.8),
                         d["subtitle"], size=20, color=self.theme["muted"])

    def s_bullets(self, d):
        s = self.new_slide(d)
        self.header(s, d["title"])
        self.bullets(s, MARGIN, CONTENT_TOP, CONTENT_W, CONTENT_H, d["bullets"])

    def s_two_column(self, d):
        s = self.new_slide(d)
        self.header(s, d["title"])
        gap = Inches(0.4)
        col_w = (CONTENT_W - gap) // 2
        for i, col in enumerate((d["left"], d["right"])):
            left = MARGIN + i * (col_w + gap)
            self.rect(s, left, CONTENT_TOP, col_w, Inches(0.6), self.theme["primary"])
            self.textbox(s, left + Inches(0.2), CONTENT_TOP, col_w - Inches(0.4), Inches(0.6),
                         col.get("heading", ""), size=20, bold=True,
                         color=self.theme["on_primary"], anchor=MSO_ANCHOR.MIDDLE)
            self.rect(s, left, CONTENT_TOP + Inches(0.6), col_w, CONTENT_H - Inches(0.6),
                      self.theme["light"])
            self.bullets(s, left + Inches(0.2), CONTENT_TOP + Inches(0.8), col_w - Inches(0.4),
                         CONTENT_H - Inches(1), col.get("bullets", []),
                         base_size=20 if len(col.get("bullets", [])) <= 4 else 18 if len(col.get("bullets", [])) <= 6 else 16,
                         fill=self.theme["light"])

    def s_table(self, d):
        s = self.new_slide(d)
        self.header(s, d["title"])
        cols, rows = d["columns"], d["rows"]
        n_rows = len(rows) + 1
        row_h = Inches(0.5) if n_rows <= 8 else Inches(0.4)
        height = Emu(int(row_h) * n_rows)
        tbl = s.shapes.add_table(n_rows, len(cols), MARGIN, CONTENT_TOP, CONTENT_W, height).table
        widths = d.get("col_widths")
        if widths:
            total = sum(widths)
            for i, w in enumerate(widths):
                tbl.columns[i].width = Emu(int(CONTENT_W * w / total))
        size = 16 if n_rows <= 8 else 13 if n_rows <= 12 else 11
        for r in range(n_rows):
            tbl.rows[r].height = row_h
            values = cols if r == 0 else rows[r - 1]
            for c in range(len(cols)):
                cell = tbl.cell(r, c)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.fill.solid()
                if r == 0:
                    cell.fill.fore_color.rgb = rgb(self.theme["primary"])
                else:
                    cell.fill.fore_color.rgb = rgb(self.theme["light"] if r % 2 else "FFFFFF")
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = str(values[c]) if c < len(values) else ""
                self.set_font(run.font, size, r == 0,
                              self.theme["on_primary"] if r == 0 else self.theme["text"])
        if d.get("caption"):
            self.textbox(s, MARGIN, SLIDE_H - Inches(1.05), CONTENT_W, Inches(0.4),
                         d["caption"], size=12, color=self.theme["muted"])

    def s_chart(self, d):
        s = self.new_slide(d)
        self.header(s, d["title"])
        kind = d.get("chart_type", "column")
        if kind not in CHART_TYPES:
            raise ValueError(f"Unsupported chart_type '{kind}'. Use one of {sorted(CHART_TYPES)}")
        data = CategoryChartData(number_format=d.get("number_format", "General"))
        data.categories = d["categories"]
        for ser in d["series"]:
            data.add_series(ser["name"], ser["values"])
        has_side = bool(d.get("takeaways"))
        width = Inches(8.2) if has_side else CONTENT_W
        gf = s.shapes.add_chart(CHART_TYPES[kind], MARGIN, CONTENT_TOP, width, CONTENT_H, data)
        chart = gf.chart
        self.set_font(chart.font, 14, color=self.theme["text"])
        chart.has_title = bool(d.get("chart_title"))
        if chart.has_title:
            chart.chart_title.text_frame.text = d["chart_title"]
            self.set_font(chart.chart_title.text_frame.paragraphs[0].runs[0].font, 16, True,
                          self.theme["text"])
        single_pie = kind in ("pie", "doughnut")
        multi = len(d["series"]) > 1
        chart.has_legend = d.get("legend", single_pie or multi)
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
        if d.get("data_labels", True):
            plot = chart.plots[0]
            plot.has_data_labels = True
            dl = plot.data_labels
            self.set_font(dl.font, 12, color=self.theme["text"])
            if single_pie:
                self.set_font(dl.font, 14, True, "FFFFFF")
                dl.show_percentage = d.get("show_percentage", True)
                dl.show_value = not dl.show_percentage
                dl.number_format = "0%" if dl.show_percentage else d.get("number_format", "General")
                dl.number_format_is_linked = False
        palette = self.theme["series"]
        if single_pie:
            for i, point in enumerate(chart.plots[0].series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = rgb(palette[i % len(palette)])
        else:
            for i, ser in enumerate(chart.plots[0].series):
                fmt = ser.format
                if kind.startswith("line") or kind == "radar":
                    fmt.line.color.rgb = rgb(palette[i % len(palette)])
                    fmt.line.width = Pt(2.5)
                    if kind != "line_plain":
                        ser.marker.format.fill.solid()
                        ser.marker.format.fill.fore_color.rgb = rgb(palette[i % len(palette)])
                        ser.marker.format.line.color.rgb = rgb(palette[i % len(palette)])
                else:
                    fmt.fill.solid()
                    fmt.fill.fore_color.rgb = rgb(palette[i % len(palette)])
            if kind in ("column", "column_stacked", "bar", "bar_stacked"):
                chart.plots[0].gap_width = 80
            if kind != "radar":
                va = chart.value_axis
                va.has_major_gridlines = True
                va.major_gridlines.format.line.color.rgb = rgb("D9D9D9")
                va.format.line.fill.background()
        if has_side:
            left = MARGIN + width + Inches(0.3)
            side_w = CONTENT_W - width - Inches(0.3)
            self.rect(s, left, CONTENT_TOP, side_w, CONTENT_H, self.theme["light"])
            self.bullets(s, left + Inches(0.2), CONTENT_TOP + Inches(0.25), side_w - Inches(0.4),
                         CONTENT_H - Inches(0.5), d["takeaways"], base_size=16,
                         fill=self.theme["light"])
        if d.get("source"):
            self.textbox(s, MARGIN, SLIDE_H - Inches(0.6), Inches(10), Inches(0.35),
                         "출처: " + d["source"] if not d["source"].lower().startswith(("source", "출처"))
                         else d["source"], size=11, color=self.theme["muted"])

    def s_stats(self, d):
        s = self.new_slide(d)
        self.header(s, d["title"])
        items = d["items"][:4]
        gap = Inches(0.35)
        w = (CONTENT_W - gap * (len(items) - 1)) // len(items)
        h = Inches(2.6)
        top = CONTENT_TOP + Inches(0.4)
        for i, it in enumerate(items):
            left = MARGIN + i * (w + gap)
            self.rect(s, left, top, w, h, self.theme["light"])
            self.rect(s, left, top, w, Inches(0.08), self.theme["accent"])
            # Shrink long values so they stay on one line inside the tile.
            n_chars = len(str(it["value"])) * (4 / len(items))
            value_size = 44 if n_chars <= 6 else 36 if n_chars <= 8 else 30 if n_chars <= 10 else 24
            self.textbox(s, left, top + Inches(0.35), w, Inches(1.2), it["value"], size=value_size,
                         bold=True, color=self.theme["primary"], align=PP_ALIGN.CENTER,
                         anchor=MSO_ANCHOR.MIDDLE, fill=self.theme["light"])
            self.textbox(s, left + Inches(0.15), top + Inches(1.6), w - Inches(0.3), h - Inches(1.6),
                         it["label"], size=16 if len(it["label"]) <= 12 else 14,
                         color=self.theme["text"], align=PP_ALIGN.CENTER, fill=self.theme["light"])
        if d.get("caption"):
            self.textbox(s, MARGIN, top + h + Inches(0.5), CONTENT_W, Inches(1), d["caption"],
                         size=16, color=self.theme["muted"], align=PP_ALIGN.CENTER)

    def s_image(self, d):
        s = self.new_slide(d)
        self.header(s, d["title"])
        path = Path(d["path"])
        if not path.is_absolute():
            path = PROJECT_ROOT / path
        pic = s.shapes.add_picture(str(path), MARGIN, CONTENT_TOP)
        max_w, max_h = CONTENT_W, CONTENT_H - (Inches(0.5) if d.get("caption") else 0)
        scale = min(max_w / pic.width, max_h / pic.height, 1.0)
        pic.width, pic.height = int(pic.width * scale), int(pic.height * scale)
        pic.left = int((SLIDE_W - pic.width) / 2)
        if d.get("caption"):
            self.textbox(s, MARGIN, pic.top + pic.height + Inches(0.1), CONTENT_W, Inches(0.4),
                         d["caption"], size=12, color=self.theme["muted"], align=PP_ALIGN.CENTER)

    def s_closing(self, d):
        s = self.new_slide(d)
        self.rect(s, 0, 0, SLIDE_W, SLIDE_H, self.theme["primary"])
        self.textbox(s, MARGIN, Inches(2.6), CONTENT_W, Inches(1.3), d.get("title", "감사합니다"),
                     size=44, bold=True, color=self.theme["on_primary"], align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE)
        if d.get("subtitle"):
            self.textbox(s, MARGIN, Inches(4.0), CONTENT_W, Inches(0.8), d["subtitle"], size=20,
                         color=self.theme.get("subtitle", "D9E2F3"), align=PP_ALIGN.CENTER)

    def build(self):
        core = self.prs.core_properties
        core.title = self.spec.get("title", "")
        core.author = self.spec.get("author", "")
        for i, d in enumerate(self.spec["slides"], 1):
            handler = getattr(self, "s_" + d.get("type", "bullets"), None)
            if handler is None:
                raise ValueError(f"Slide {i}: unknown type '{d.get('type')}'")
            handler(d)
        return self.prs


def unique_path(path):
    """Return path, or path with _v2, _v3, ... appended if it already exists."""
    if not path.exists():
        return path
    n = 2
    while True:
        cand = path.with_name(f"{path.stem}_v{n}{path.suffix}")
        if not cand.exists():
            return cand
        n += 1


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", help="JSON spec file")
    ap.add_argument("--output", help="output .pptx (default: slides/<spec 'filename'>)")
    args = ap.parse_args()

    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    out = Path(args.output or spec.get("filename") or "presentation.pptx")
    if not out.is_absolute():
        out = (PROJECT_ROOT / out) if out.parent != Path(".") else DEFAULT_DIR / out
    if out.suffix.lower() != ".pptx":
        out = out.with_suffix(".pptx")
    out.parent.mkdir(parents=True, exist_ok=True)
    out = unique_path(out)

    Builder(spec).build().save(out)

    # Re-open to verify the file is valid and report what was written.
    check = Presentation(out)
    print(f"Saved: {out}")
    print(f"Slides: {len(check.slides)}")
    for i, slide in enumerate(check.slides, 1):
        texts = [sh.text_frame.text for sh in slide.shapes
                 if sh.has_text_frame and sh.text_frame.text.strip()]
        label = texts[0].replace("\n", " ")[:50] if texts else ""
        extras = [name for name, hit in (
            ("chart", any(sh.has_chart for sh in slide.shapes)),
            ("table", any(sh.has_table for sh in slide.shapes)),
            ("notes", slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip()),
        ) if hit]
        kind = spec["slides"][i - 1].get("type", "bullets")
        print(f"  {i:>2}. [{kind}] {label}" + (f"  ({', '.join(extras)})" if extras else ""))

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
