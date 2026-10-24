# python-pptx API 요약표 (v1.0.x)

필요한 절만 읽는다. 단위는 EMU이며, `pptx.util`의 `Inches`, `Cm`, `Pt`, `Mm`, `Emu`를 쓴다.

## 목차
1. 프레젠테이션 파일
2. 슬라이드
3. 도형
4. 텍스트
5. 표
6. 차트
7. 채우기, 선, 색상
8. 클릭 동작과 하이퍼링크
9. 기존 발표자료 읽기
10. 지원하지 않는 기능과 우회 방법

## 1. 프레젠테이션 파일

```python
from pptx import Presentation
prs = Presentation()                 # 새 파일 (기본 템플릿은 4:3 → 아래처럼 16:9로 설정)
prs = Presentation("in.pptx")        # 기존 파일 열기 (파일 형식 객체도 가능)
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
cp = prs.core_properties             # 제목, 작성자, 주제, 키워드, 만든 날짜, 수정 날짜, 리비전
cp.title, cp.author = "Title", "Name"
prs.slide_masters, prs.slide_layouts # 마스터, 첫 번째 마스터의 레이아웃
prs.save("out.pptx")                 # BytesIO 스트림도 가능
```

기본 레이아웃: 0 제목, 1 제목+내용, 2 구역 머리글, 3 콘텐츠 2개, 4 비교, 5 제목만, 6 빈 화면, 7 캡션 있는 콘텐츠, 8 캡션 있는 그림.

## 2. 슬라이드

```python
slide = prs.slides.add_slide(prs.slide_layouts[1])
for s in prs.slides: s.slide_id, s.name, s.slide_layout.name
slide.follow_master_background = False
fill = slide.background.fill; fill.solid(); fill.fore_color.rgb = RGBColor(0xF5, 0xF5, 0xF5)
slide.notes_slide.notes_text_frame.text = "speaker notes"   # 발표자 노트
slide.placeholders            # idx로 접근: slide.placeholders[1]
prs.slide_master.shapes       # 마스터의 도형
```

## 3. 도형

```python
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR, MSO_SHAPE_TYPE
sh = slide.shapes
sh.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)   # 자동 도형 180여 종
sh.add_textbox(l, t, w, h)
pic = sh.add_picture("img.png", l, t, width=None, height=None)  # 한쪽만 주면 비율 유지
pic.crop_left = 0.1                                        # 자르기 비율
c = sh.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
c.begin_connect(shape_a, 3); c.end_connect(shape_b, 1)      # 연결점에 붙이기
ff = sh.build_freeform(x, y); ff.add_line_segments([(x2, y2), (x3, y3)], close=True); ff.convert_to_shape()
grp = sh.add_group_shape(); grp.shapes.add_shape(...)
sh.add_movie("clip.mp4", l, t, w, h, poster_frame_image="poster.png", mime_type="video/mp4")
sh.add_ole_object("data.xlsx", "Excel.Sheet.12", l, t)      # OLE 개체 삽입
shape.left, shape.top, shape.width, shape.height, shape.rotation, shape.name
shape.adjustments[0] = 0.2                                 # 예: 모서리 둥글기
shape.shape_type == MSO_SHAPE_TYPE.PICTURE                  # TABLE, CHART, GROUP 등도 있음
placeholder.insert_picture("img.png")                      # 그림 자리표시자
```

도형 삭제: `shape._element.getparent().remove(shape._element)`.

## 4. 텍스트

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
tf.margin_left = Inches(0.1)                     # right/top/bottom도 있음
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE   # 또는 SHAPE_TO_FIT_TEXT / NONE
tf.fit_text(font_family="Malgun Gothic", max_size=24)  # 글꼴 크기 계산, 글꼴 파일 필요
"line\vbreak"                                    # \v = 문단 안 줄 바꿈
```

한국어 텍스트: `font.name`은 라틴 글꼴 슬롯만 설정한다. `<a:ea typeface=...>`도 함께 설정한다(`scripts/pptx_helpers.py`의 `set_font` 참고).

## 5. 표

```python
gf = slide.shapes.add_table(rows, cols, l, t, w, h); tbl = gf.table
tbl.columns[0].width = Inches(2); tbl.rows[0].height = Inches(0.5)
cell = tbl.cell(0, 0); cell.text = "x"
cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(...)
cell.margin_left = Inches(0.05); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
cell.merge(tbl.cell(1, 1)); cell.is_merge_origin; cell.is_spanned; cell.split()   # 병합, 분할
tbl.first_row = True; tbl.horz_banding = True; tbl.first_col = False
```

## 6. 차트

```python
from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_MARKER_STYLE
cd = CategoryChartData(); cd.categories = ["A", "B"]; cd.add_series("S1", (1.2, 3.4))
chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, l, t, w, h, cd).chart
xy = XyChartData(); s = xy.add_series("pts"); s.add_data_point(1.0, 2.0)      # XY_SCATTER
bb = BubbleChartData(); s = bb.add_series("b"); s.add_data_point(1, 2, 10)     # BUBBLE
chart.replace_data(new_cd)                         # 기존 차트 데이터 교체
chart.has_title = True; chart.chart_title.text_frame.text = "Title"
chart.has_legend = True; chart.legend.position = XL_LEGEND_POSITION.BOTTOM; chart.legend.include_in_layout = False
va = chart.value_axis; va.minimum_scale, va.maximum_scale, va.major_unit = 0, 100, 20
va.has_major_gridlines = True; va.tick_labels.number_format = '0"%"'; va.tick_labels.number_format_is_linked = False
va.has_title = True; va.axis_title.text_frame.text = "Unit"; chart.category_axis.visible = True
plot = chart.plots[0]; plot.gap_width = 80; plot.overlap = 0; plot.vary_by_categories = False
plot.has_data_labels = True; dl = plot.data_labels; dl.show_value = True; dl.show_percentage = True; dl.position = XL_LABEL_POSITION.OUTSIDE_END
ser = plot.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = RGBColor(...)
ser.format.line.width = Pt(2.5); ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = 8; ser.smooth = False
ser.points[1].format.fill.solid()                  # 데이터 포인트별 서식
chart.chart_style = 10; chart.font.size = Pt(12)
```

차트 종류: 세로/가로 막대(묶은형, 누적, 100% 누적), 꺾은선(표식 포함), 원형, 도넛형, 영역형, 분산형(XY_SCATTER), 거품형, 방사형과 각각의 3차원 형식.

## 7. 채우기, 선, 색상

```python
shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(0x12, 0x34, 0x56)
shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1; shape.fill.fore_color.brightness = 0.4
shape.fill.gradient(); shape.fill.gradient_angle = 90; stops = shape.fill.gradient_stops  # stops[0].color.rgb, .position
shape.fill.patterned(); shape.fill.pattern = MSO_PATTERN.CROSS; shape.fill.background()  # 채우기 없음
shape.line.color.rgb = RGBColor(...); shape.line.width = Pt(1.5); shape.line.dash_style = MSO_LINE.DASH
shape.shadow.inherit = False                        # 그림자는 이것만 제어 가능
```

## 8. 클릭 동작과 하이퍼링크

```python
from pptx.enum.action import PP_ACTION
shape.click_action.hyperlink.address = "https://example.com"
shape.click_action.target_slide = prs.slides[3]     # 해당 슬라이드로 이동
shape.click_action.action  # PP_ACTION.NEXT_SLIDE / PREVIOUS_SLIDE / FIRST_SLIDE / LAST_SLIDE ... (읽기)
shape.click_over_action    # 마우스를 올렸을 때 동작, 사용법 동일
```

## 9. 기존 발표자료 읽기

`python .claude/skills/mk-ppt/scripts/inspect_pptx.py file.pptx`를 실행한다. 또는 `prs.slides` → `slide.shapes` → `has_text_frame` / `has_table` / `has_chart` / `shape_type == PICTURE`(`shape.image.blob`, `shape.image.ext`) 순서로 돌며 읽는다.

## 10. 지원하지 않는 기능과 우회 방법

| 기능 | 상태 |
|---|---|
| 애니메이션, 화면 전환 | 지원하지 않음. XML(`p:timing`, `p:transition`)을 직접 다뤄야 하므로 피한다. |
| SmartArt | 지원하지 않음. 도형과 연결선으로 대신 만든다. |
| 그림자·반사·3D 세부 효과 | `shadow.inherit`만 가능. 그 외에는 XML 직접 수정. |
| 새 마스터·레이아웃 | 지원하지 않음. 해당 레이아웃이 있는 템플릿 `.pptx`에서 시작한다. |
| PDF·이미지 내보내기, 미리보기 | 지원하지 않음. LibreOffice: `soffice --headless --convert-to pdf file.pptx`. |
| 메모(댓글) | 지원하지 않음. |
| 구형 `.ppt` | 지원하지 않음. 먼저 `.pptx`로 변환한다. |
| 슬라이드 삭제 | 아래 우회 방법 참고. |
| 슬라이드 순서 변경 | 아래 우회 방법 참고. |
| 슬라이드 복제 | API 없음. 같은 레이아웃으로 슬라이드를 추가하고 도형 XML을 복사한다(`copy.deepcopy(sh._element)`를 `new.shapes._spTree`에 넣기). 차트와 그림은 관계(relationship)도 복사해야 하므로 다시 생성하는 편이 낫다. |

```python
# i번째 슬라이드 삭제
sldIdLst = prs.slides._sldIdLst
sld = sldIdLst[i]; prs.part.drop_rel(sld.rId); sldIdLst.remove(sld)
# i번째 슬라이드를 j 위치로 이동
sld = sldIdLst[i]; sldIdLst.remove(sld); sldIdLst.insert(j, sld)
```
