# python-pptx API cheat sheet (v1.0.x)

Read the section you need. Units are EMU; use `Inches`, `Cm`, `Pt`, `Mm`, `Emu` from `pptx.util`.

## Contents
1. Presentation file
2. Slides
3. Shapes
4. Text
5. Tables
6. Charts
7. Fill, line, color
8. Click actions and hyperlinks
9. Reading an existing deck
10. Unsupported features and workarounds

## 1. Presentation file

```python
from pptx import Presentation
prs = Presentation()                 # new (default template, 4:3 → set 16:9 below)
prs = Presentation("in.pptx")        # open existing (also accepts a file-like object)
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
cp = prs.core_properties             # title, author, subject, keywords, created, modified, revision
cp.title, cp.author = "Title", "Name"
prs.slide_masters, prs.slide_layouts # masters and the first master's layouts
prs.save("out.pptx")                 # or a BytesIO stream
```

Default layouts: 0 Title, 1 Title+Content, 2 Section Header, 3 Two Content, 4 Comparison, 5 Title Only, 6 Blank, 7 Content w/ Caption, 8 Picture w/ Caption.

## 2. Slides

```python
slide = prs.slides.add_slide(prs.slide_layouts[1])
for s in prs.slides: s.slide_id, s.name, s.slide_layout.name
slide.follow_master_background = False
fill = slide.background.fill; fill.solid(); fill.fore_color.rgb = RGBColor(0xF5, 0xF5, 0xF5)
slide.notes_slide.notes_text_frame.text = "speaker notes"
slide.placeholders            # by idx: slide.placeholders[1]
prs.slide_master.shapes       # shapes on the master
```

## 3. Shapes

```python
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR, MSO_SHAPE_TYPE
sh = slide.shapes
sh.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)   # 180+ autoshapes
sh.add_textbox(l, t, w, h)
pic = sh.add_picture("img.png", l, t, width=None, height=None)  # keeps aspect if one side given
pic.crop_left = 0.1                                        # crop fractions
c = sh.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
c.begin_connect(shape_a, 3); c.end_connect(shape_b, 1)      # glue to connection points
ff = sh.build_freeform(x, y); ff.add_line_segments([(x2, y2), (x3, y3)], close=True); ff.convert_to_shape()
grp = sh.add_group_shape(); grp.shapes.add_shape(...)
sh.add_movie("clip.mp4", l, t, w, h, poster_frame_image="poster.png", mime_type="video/mp4")
sh.add_ole_object("data.xlsx", "Excel.Sheet.12", l, t)      # embedded OLE object
shape.left, shape.top, shape.width, shape.height, shape.rotation, shape.name
shape.adjustments[0] = 0.2                                 # e.g. rounded-corner radius
shape.shape_type == MSO_SHAPE_TYPE.PICTURE                  # also TABLE, CHART, GROUP, ...
placeholder.insert_picture("img.png")                      # picture placeholder
```

Delete a shape: `shape._element.getparent().remove(shape._element)`.

## 4. Text

```python
tf = shape.text_frame
tf.text = "first paragraph"
p = tf.add_paragraph(); p.text = "second"; p.level = 1
p.alignment = PP_ALIGN.CENTER; p.line_spacing = 1.2; p.space_before = Pt(6); p.space_after = Pt(6)
r = p.add_run(); r.text = "run"
f = r.font; f.name, f.size, f.bold, f.italic, f.underline = "맑은 고딕", Pt(18), True, False, True
f.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
r.hyperlink.address = "https://example.com"
tf.word_wrap = True
tf.margin_left = Inches(0.1)                     # also right/top/bottom
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE   # or SHAPE_TO_FIT_TEXT / NONE
tf.fit_text(font_family="Malgun Gothic", max_size=24)  # computes size; needs the font file
"line\vbreak"                                    # \v = line break inside a paragraph
```

Korean text: `font.name` sets only the Latin slot. Also set `<a:ea typeface=...>` (see `set_font` in `scripts/pptx_helpers.py`).

## 5. Tables

```python
gf = slide.shapes.add_table(rows, cols, l, t, w, h); tbl = gf.table
tbl.columns[0].width = Inches(2); tbl.rows[0].height = Inches(0.5)
cell = tbl.cell(0, 0); cell.text = "x"
cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(...)
cell.margin_left = Inches(0.05); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
cell.merge(tbl.cell(1, 1)); cell.is_merge_origin; cell.is_spanned; cell.split()
tbl.first_row = True; tbl.horz_banding = True; tbl.first_col = False
```

## 6. Charts

```python
from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_MARKER_STYLE
cd = CategoryChartData(); cd.categories = ["A", "B"]; cd.add_series("S1", (1.2, 3.4))
chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, l, t, w, h, cd).chart
xy = XyChartData(); s = xy.add_series("pts"); s.add_data_point(1.0, 2.0)      # XY_SCATTER
bb = BubbleChartData(); s = bb.add_series("b"); s.add_data_point(1, 2, 10)     # BUBBLE
chart.replace_data(new_cd)                         # swap data in existing chart
chart.has_title = True; chart.chart_title.text_frame.text = "Title"
chart.has_legend = True; chart.legend.position = XL_LEGEND_POSITION.BOTTOM; chart.legend.include_in_layout = False
va = chart.value_axis; va.minimum_scale, va.maximum_scale, va.major_unit = 0, 100, 20
va.has_major_gridlines = True; va.tick_labels.number_format = '0"%"'; va.tick_labels.number_format_is_linked = False
va.has_title = True; va.axis_title.text_frame.text = "Unit"; chart.category_axis.visible = True
plot = chart.plots[0]; plot.gap_width = 80; plot.overlap = 0; plot.vary_by_categories = False
plot.has_data_labels = True; dl = plot.data_labels; dl.show_value = True; dl.show_percentage = True; dl.position = XL_LABEL_POSITION.OUTSIDE_END
ser = plot.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = RGBColor(...)
ser.format.line.width = Pt(2.5); ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = 8; ser.smooth = False
ser.points[1].format.fill.solid()                  # per-point formatting
chart.chart_style = 10; chart.font.size = Pt(12)
```

Types include COLUMN/BAR (clustered, stacked, 100%), LINE (+markers), PIE, DOUGHNUT, AREA, XY_SCATTER, BUBBLE, RADAR, and 3-D variants.

## 7. Fill, line, color

```python
shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)
shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1; shape.fill.fore_color.brightness = 0.4
shape.fill.gradient(); shape.fill.gradient_angle = 90; stops = shape.fill.gradient_stops  # stops[0].color.rgb, .position
shape.fill.patterned(); shape.fill.pattern = MSO_PATTERN.CROSS; shape.fill.background()  # no fill
shape.line.color.rgb = RGBColor(...); shape.line.width = Pt(1.5); shape.line.dash_style = MSO_LINE.DASH
shape.shadow.inherit = False                        # only shadow control available
```

## 8. Click actions and hyperlinks

```python
from pptx.enum.action import PP_ACTION
shape.click_action.hyperlink.address = "https://example.com"
shape.click_action.target_slide = prs.slides[3]     # jump to slide
shape.click_action.action  # PP_ACTION.NEXT_SLIDE / PREVIOUS_SLIDE / FIRST_SLIDE / LAST_SLIDE ... (read)
shape.click_over_action    # mouse-over action, same API
```

## 9. Reading an existing deck

Run `python .claude/skills/mk-ppt/scripts/inspect_pptx.py file.pptx`, or iterate `prs.slides` → `slide.shapes` → `has_text_frame` / `has_table` / `has_chart` / `shape_type == PICTURE` (`shape.image.blob`, `shape.image.ext`).

## 10. Unsupported features and workarounds

| Feature | Status |
|---|---|
| Animations, transitions | Not supported. Would need raw XML (`p:timing`, `p:transition`); avoid. |
| SmartArt | Not supported; build from shapes + connectors instead. |
| Shadow / reflection / 3-D detail | Only `shadow.inherit`. Raw XML otherwise. |
| New masters or layouts | Not supported; start from a template `.pptx` that has them. |
| PDF / image export, thumbnails | Not supported. LibreOffice: `soffice --headless --convert-to pdf file.pptx`. |
| Comments | Not supported. |
| Legacy `.ppt` | Not supported; convert to `.pptx` first. |
| Delete slide | Workaround below. |
| Reorder slide | Workaround below. |
| Duplicate slide | No API; add a slide with the same layout and copy shape XML (`copy.deepcopy(sh._element)` into `new.shapes._spTree`). Charts and pictures need their relationships copied too, so prefer regenerating. |

```python
# delete slide i
sldIdLst = prs.slides._sldIdLst
sld = sldIdLst[i]; prs.part.drop_rel(sld.rId); sldIdLst.remove(sld)
# move slide from i to j
sld = sldIdLst[i]; sldIdLst.remove(sld); sldIdLst.insert(j, sld)
```
