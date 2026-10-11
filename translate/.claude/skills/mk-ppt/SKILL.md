---
name: mk-ppt-translation
description: Korean translation of .claude/skills/mk-ppt/SKILL.md for reading only. Not a skill; never invoke.
disable-model-invocation: true
user-invocable: false
---

> 번역본 안내: 위 frontmatter는 이 파일이 스킬로 잘못 인식되지 않도록 자동 호출을 끈 설정입니다. 원본의 `name`은 `mk-ppt`이고, `description`의 번역은 다음과 같습니다.
>
> 이 프로젝트에서 python-pptx로 PowerPoint(.pptx) 발표 자료를 만듭니다. 제목, 섹션, 글머리, 2단, 표, 차트, 핵심 수치(KPI), 이미지, 마무리 슬라이드와 발표자 노트, 한글 글꼴, 일관된 테마를 지원합니다. 사용자가 PPT, PowerPoint, 발표 자료, 슬라이드, 프레젠테이션, deck을 요청하거나 자료조사 노트·보고서 등 어떤 내용이든 슬라이드로 바꾸고 싶어 할 때 ".pptx"나 "python-pptx"를 말하지 않더라도 이 스킬을 사용합니다. python-pptx로 발표 자료에 무엇을 할 수 있고 없는지 물을 때도 사용합니다.

# mk-ppt — python-pptx로 .pptx 발표 자료 만들기

이 스킬은 내용(주제, 사용자의 메모, `research/`·`report/`의 파일)을 깔끔한 16:9 발표 자료로 바꿉니다. 실제 생성은 번들 스크립트가 맡기 때문에 모든 자료가 같은 디자인과 같은 한글 글꼴 처리를 갖습니다. 여러분의 역할은 흐름을 기획하고, JSON 명세를 쓰고, 스크립트를 실행하고, 결과를 확인하는 것입니다.

## 0단계 — PPT 규칙부터 읽기

이 스킬의 다른 어떤 단계보다 먼저 `.claude/rules/ppt-rule-v1/ppt-rule-v1.md`를 읽고 따릅니다. 이 스킬과 내용이 다르면 규칙이 우선합니다. 발표 자료의 우선순위는 다음과 같습니다.

1. `.claude/rules/ppt-rule-v1/ppt-rule-v1.md` (및 `CLAUDE.md`)
2. 사용자가 제공한 디자인 `.md` (예: `references/ppt/slide-design-system.md`)
3. 이 스킬의 기본값 (테마, 글꼴, 레이아웃)

요약: 기본 8장, 사용자의 디자인 파일을 따르고(없으면 요청), 글이 넘치거나 겹치지 않게 합니다. 보고 전에 렌더링해서 고칩니다.

## 여기에 적용되는 프로젝트 규칙

이 프로젝트는 `CLAUDE.md`에 자체 규칙이 있으며 그것이 우선합니다. 발표 자료와 관련해 중요한 규칙은 다음과 같습니다.

- **Todo list를 먼저 보고하고 승인을 기다립니다.** 만들기 전에 계획을 보고합니다: 슬라이드 개요(각 슬라이드 제목과 종류), 원본 파일, 저장 경로, 미정 사항(테마, 슬라이드 수, 청중). 사용자가 승인한 뒤에만 만듭니다.
- **발표 자료는 `slides/`에 저장합니다.** 프로젝트 루트에 두지 않습니다.
- **덮어쓰지 않습니다.** 파일이 이미 있으면 스크립트가 자동으로 `_v2`, `_v3`, …를 붙이므로 항상 기본 이름을 넘기면 됩니다.
- **한국어로 답하고**, 사용자가 달리 요청하지 않으면 슬라이드 문구도 한국어로 씁니다.
- JSON 명세는 임시 파일이므로 프로젝트가 아니라 scratchpad 디렉터리에 씁니다.

## 작업 순서

1. **내용과 디자인 파일 수집.** 사용자가 가리킨 원본 자료를 읽습니다. 주제만 주어졌다면 먼저 조사할지, 주어진 내용만 쓸지 묻습니다. 지어낸 숫자가 들어간 슬라이드는 슬라이드가 없는 것보다 나쁩니다. 사용자의 디자인 `.md`를 읽고, 없으면 요청합니다. 이 스킬의 기본 테마는 사용자가 원할 때만 씁니다(ppt-rule-v1 §2).
2. **흐름 기획.** 슬라이드 하나에 메시지 하나. 사용자가 다른 장수를 요청하지 않으면 **기본 8장**(ppt-rule-v1 §1): 제목 → (목차 또는 섹션) → 내용 슬라이드 → 요약/다음 단계 → 마무리. 메시지에 맞는 슬라이드 종류를 고릅니다(아래 표 참고): 시간에 따른 수치 → `chart`, 핵심 수치 2–4개 → `stats`, 비교 → `two_column` 또는 `table`, 나머지 → `bullets`.
3. **개요를 todo list로 보고**하고 승인을 기다립니다(프로젝트 규칙).
4. **명세 작성** (JSON, UTF-8)을 scratchpad에 저장합니다.
5. **프로젝트 루트에서 스크립트 실행**:
   ```bash
   python .claude/skills/mk-ppt/scripts/build_pptx.py <scratchpad>/spec.json
   ```
   `slides/<filename>`(또는 `--output <경로>`)에 저장하고, 파일을 다시 열어 검증한 뒤 슬라이드마다 한 줄씩 출력합니다.
6. **눈으로 확인 — 필수(ppt-rule-v1 §3).** 이 컴퓨터에는 PowerPoint가 설치되어 있으므로 모든 슬라이드를 PNG로 내보내 Read 도구로 확인합니다:
   ```powershell
   New-Item -ItemType Directory -Force "<scratchpad>\png" | Out-Null
   $app = New-Object -ComObject PowerPoint.Application
   $pres = $app.Presentations.Open("<전체 경로>.pptx", $true, $false, $false)
   foreach ($sl in $pres.Slides) { $sl.Export("<scratchpad>\png\s$($sl.SlideIndex).png", "PNG", 1920, 1080) }
   $pres.Close(); $app.Quit()
   ```
   상자를 넘치는 텍스트, 글과 도형의 겹침, 줄바꿈으로 잘린 단어, 빽빽한 슬라이드, 읽기 어려운 레이블을 찾아 명세를 고치고(문구 줄이기, 슬라이드 나누기) 문제가 없어질 때까지 다시 만듭니다. 텍스트에서 연한 상자 아래로 이어지는 아주 옅고 가는 줄무늬는 PowerPoint PNG 내보내기 잡티이며 파일 자체의 문제가 아니므로 무시합니다.
7. **보고:** 저장 경로, 슬라이드 목록, 그리고 요청 중 python-pptx로 할 수 없었던 것.

## 명세 형식

```json
{
  "title": "발표 제목 (문서 속성)",
  "author": "춘식",
  "filename": "AI_발전_발표.pptx",
  "theme": "navy",
  "font": "Noto Sans KR",
  "slides": [ { "type": "...", "notes": "선택: 발표자 노트", ... } ]
}
```

`theme`: `navy`(기본), `charcoal`, `green`, `orange`, `pink`. `font`의 기본값은 Noto Sans KR이며 라틴, 동아시아, 복합 문자 모두에 적용됩니다(python-pptx만으로는 라틴 글꼴만 설정되어 한글 표시가 깨집니다).

| type | 필드 | 용도 |
|---|---|---|
| `title` | `title`, `subtitle`, `author`, `date` | 첫 슬라이드 |
| `section` | `title`, `subtitle`, `number`(예: "01") | 장 구분 |
| `bullets` | `title`, `bullets`: 문자열 또는 `{"text", "level": 0/1, "bold"}` 목록 | 일반 내용 |
| `two_column` | `title`, `left`/`right`: `{"heading", "bullets"}` | 장단점, 전후 비교, A vs B |
| `table` | `title`, `columns`, `rows`(리스트의 리스트), `col_widths`(상대값, 선택), `caption` | 구조화된 데이터, 연표 |
| `chart` | `title`, `chart_type`, `categories`, `series`: `[{"name", "values"}]`, `number_format`, `legend`, `data_labels`, `chart_title`, `takeaways`(차트 옆에 표시할 글머리), `source` | 수치 추이와 비중 |
| `stats` | `title`, `items`: 최대 4개의 `{"value", "label"}`, `caption` | 핵심 KPI |
| `image` | `title`, `path`(프로젝트 루트 기준 상대 경로 또는 절대 경로), `caption` | 스크린샷, 도식 |
| `closing` | `title`(기본값 감사합니다), `subtitle` | 마지막 슬라이드 |

`chart_type`: `column`, `column_stacked`, `bar`, `bar_stacked`, `line`, `line_plain`, `pie`, `doughnut`, `area`, `radar`. `pie`/`doughnut`은 계열을 하나만 씁니다. 레이블은 기본적으로 백분율을 보여 줍니다(`"show_percentage": false`이면 값 표시). `number_format`은 `"0%"`, `"#,##0"`, `"0.0"`처럼 씁니다.

## 좋은 슬라이드 쓰기

- **문구는 짧게.** 슬라이드당 글머리 약 6개, 글머리당 한글 약 40자까지(`two_column`의 한 열에서는 약 22자, 차트 `takeaways`에서는 약 14자). 글머리가 늘면 글꼴 크기가 자동으로 줄지만 8개를 넘으면 읽기 어려우니 슬라이드를 나눕니다.
- **한글은 단어 중간에서 줄이 바뀝니다.** PowerPoint는 한글을 글자 단위로 줄바꿈하므로 조금만 길어도 "난/제"처럼 끊깁니다. 좁은 상자에서는 줄바꿈에 기대지 말고 한 줄에 들어가도록 문구 길이를 맞춥니다.
- **`stats` 타일은 작습니다.** 타일 4개일 때 `value`는 약 7자("1.7억", "$5,817억", "73:23" 등), `label`은 약 12자로 줄이고 나머지는 `caption`이나 `notes`로 옮깁니다. 긴 값은 자동으로 작아지지만 짧은 값이 더 잘 읽힙니다.
- **제목은 주제가 아니라 요점을 말합니다**: "AI 도입 현황"보다 "기업 AI 도입률 88% 돌파".
- **자세한 내용은 `notes`에** 넣어 슬라이드는 깔끔하게 두고 발표자는 전체 내용을 갖도록 합니다.
- 수치에는 `source`(차트)나 `caption`(표, 수치)으로 **출처를 밝힙니다**.
- 12행 또는 6열을 넘는 표는 읽기 어렵습니다. 나누거나 요약합니다.

## python-pptx로 할 수 없는 것

사용자가 다음을 요청하면 먼저 알리고 대안을 제시합니다.

- 애니메이션과 슬라이드 전환(원시 XML로만 가능하고 불안정함. PowerPoint에서 추가하도록 안내)
- 공식 API로 슬라이드 삭제, 순서 변경, 복제(XML 조작으로 가능하지만 여기서는 명세로 다시 만드는 편이 간단함)
- 자체적으로 슬라이드를 이미지/PDF로 렌더링(이 컴퓨터에서는 6단계처럼 PowerPoint COM 사용)
- SmartArt, 3D/주식형/표면형 차트 새로 만들기, 구버전 `.ppt` 열기, VBA 매크로 실행·편집
- 세밀한 그림자/광선/3D 효과(“상속 여부”만 가능)

스크립트가 다루지 않는 작업(예: 기존 자료 편집, 회사 템플릿의 플레이스홀더, 하이퍼링크, 커넥터, 자유형, 동영상, OLE 개체)은 python-pptx 코드를 직접 작성합니다. `references/python_pptx_api.md`에 분야별 API와 짧은 예제가 있으니, 명세 형식 밖의 기능이 필요할 때 읽으세요.
