---
name: mk-ppt
description: 이 프로젝트에서 python-pptx로 PowerPoint(.pptx) 발표자료를 만들거나 수정한다. 슬라이드, 텍스트, 표, 차트, 이미지, 도형, 발표자 노트를 다룬다. 사용자가 PPT, 발표자료, 슬라이드, 프레젠테이션, 덱, .pptx 파일을 요청하거나, 자료조사 노트나 보고서를 슬라이드로 바꾸고 싶어 하거나, 기존 .pptx를 수정하거나 살펴보고 싶어 할 때는 "python-pptx"라고 말하지 않더라도 항상 이 스킬을 사용한다.
---

# mk-ppt

이 프로젝트의 폴더 규칙에 맞춰, 다시 실행할 수 있는 생성 스크립트를 통해 python-pptx로 `.pptx` 발표자료를 만든다.

## 프로젝트 규칙 (중요한 이유)

이 프로젝트는 작업을 kebab-case `<topic>` 하나로 묶고, 모든 결과물을 다시 생성할 수 있어야 한다.

| 무엇 | 위치 |
|---|---|
| 원본 노트 (있을 경우) | `research/<topic>/*.md` |
| 생성 스크립트 | `scripts/make_<topic>_ppt.py` (파일 이름의 topic은 snake_case) |
| 결과 발표자료 | `reports/<topic>/<topic_snake>_slides.pptx` |

- 기존 파일을 덮어쓰지 않는다. 보조 함수 `unique_path()`가 `_v2`, `_v3`, ... 이름을 자동으로 고른다. 생성 스크립트 이름이 이미 있으면 새 스크립트를 `make_<topic>_ppt_v2.py`로 저장한다.
- 스크립트는 어느 폴더에서 실행해도 동작해야 하므로, 경로는 현재 폴더가 아니라 `__file__`을 기준으로 찾는다.
- 파일과 폴더 이름은 ASCII로 짓는다. 한글 이름은 Windows 셸과 git에서 깨진다.
- 슬라이드 본문은 사용자의 언어(보통 한국어)를 따른다. 한글을 지원하는 글꼴(기본 `맑은 고딕`)을 쓴다. `font.name`만 설정하면 한글이 테마 글꼴로 남기 때문에, 보조 함수는 동아시아 글꼴 슬롯까지 함께 설정한다.

## 작업 순서

1. **라이브러리 확인.** `python -c "import pptx; print(pptx.__version__)"`를 실행한다. 실패하면 사용자에게 알리고 `python -m pip install python-pptx`로 설치한다.
2. **발표자료 구성.** 주제 slug, 슬라이드 수, 슬라이드별 제목, 내용 유형(글머리표 / 표 / 차트 / 이미지 / 2단), 발표자 노트를 정한다. `research/<topic>/`에 원본 노트가 있으면 그것을 바탕으로 만든다. 슬라이드당 글머리표는 3~6개 정도로 한다. 빽빽한 슬라이드는 화면에 띄웠을 때 읽기 어렵다.
3. **생성 스크립트 작성.** `scripts/make_<topic>_ppt.py`에 작성한다. 반복 코드를 새로 짜지 말고 함께 제공되는 보조 함수를 불러온다.

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

   보조 함수가 다루지 않는 기능(자유형 도형, 연결선, 그라데이션, 하이퍼링크, 클릭 동작, 기존 발표자료 수정)은 python-pptx를 직접 사용한다. `references/python-pptx-api.md`를 참고한다.
4. **실행.** `python scripts/make_<topic>_ppt.py`
5. **확인.** `python .claude/skills/mk-ppt/scripts/inspect_pptx.py <output.pptx>`가 슬라이드마다 도형과 텍스트를 출력한다. 슬라이드 수를 확인하고, 비어 있는 자리표시자가 없는지 확인한다("클릭하여 텍스트 추가" 상자가 화면에 그대로 보인다). 텍스트가 눈에 띄게 넘치지 않는지도 확인한다(글머리표가 길면 슬라이드를 나눈다).
6. **보고.** 사용자에게 결과 파일 경로, 슬라이드 수, 한 줄 개요를 알린다. 새로운 종류의 폴더가 생겼다면 `CLAUDE.md`의 "Project structure"에 추가한다.

## 기존 발표자료 수정

`Presentation(path)`로 열어 수정한 뒤, `save(...)`로 새 이름으로 저장한다(원본 위에 저장하지 않는다). 먼저 `inspect_pptx.py`로 도형 이름과 번호를 파악한다.

## 한계

python-pptx는 애니메이션, 화면 전환, SmartArt, PDF·이미지 내보내기, 구형 `.ppt` 파일을 지원하지 않는다. 슬라이드 삭제, 순서 변경, 복제를 위한 공개 API도 없다. 이에 대한 XML 우회 방법은 `references/python-pptx-api.md`에 있다. 흉내 내지 말고 안 된다고 분명히 말한다.
