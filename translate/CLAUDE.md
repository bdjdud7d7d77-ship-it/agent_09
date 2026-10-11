# CLAUDE.md

이 파일은 이 저장소에서 작업할 때 Claude Code(claude.ai/code)에게 제공하는 가이드입니다.

## 정체성

- 이름: 춘식
- 사용자 입력 언어와 관계없이 항상 한글로 답변합니다.

## 프로젝트 구조

이 폴더는 코드 프로젝트가 아니라 자료조사와 보고서 작성을 위한 작업 공간입니다.

```
agent_09/
├── CLAUDE.md      # 프로젝트 규칙 (이 파일)
├── .gitignore     # Office 잠금 파일(~$*)과 OS 임시 파일 제외
├── .claude/
│   ├── rules/
│   │   └── ppt-rule-v1/   # .pptx 제작 시 항상 로드되는 규칙
│   │       └── ppt-rule-v1.md
│   └── skills/
│       └── mk-ppt/    # 이 프로젝트 전용 스킬: python-pptx로 .pptx 발표 자료 생성
│           ├── SKILL.md
│           ├── references/python_pptx_api.md
│           └── scripts/build_pptx.py
├── references/    # 산출물 제작용 디자인 참고 자료 .md (영어)
│   └── ppt/
│       └── slide-design-system.md   # 오렌지 슬라이드 디자인 시스템 (색상, 폰트, 레이아웃)
├── research/      # 자료조사 결과 .md (영어)
├── report/        # 최종 보고서 .docx (한글)
├── slides/        # 발표 자료 .pptx (한글), mk-ppt 스킬로 생성
└── translate/     # 모든 .md의 한글 번역본 (같은 상대 경로)
```

- 자료조사 결과(`.md`)는 `research/`에 저장합니다.
- 보고서(`.docx`)는 `report/`에 저장합니다. 산출물을 프로젝트 루트에 두지 않습니다.
- 발표 자료(`.pptx`)는 `mk-ppt` 스킬(`.claude/skills/mk-ppt/`)을 사용해 `slides/`에 저장합니다.
- `.pptx` 작업에는 `.claude/rules/ppt-rule-v1/ppt-rule-v1.md`를 가장 먼저 적용합니다. 우선순위: ppt-rule-v1 → 사용자의 디자인 `.md` → `mk-ppt` 기본값.
- 사용자가 요청하지 않는 한 기존 산출물을 덮어쓰지 않습니다. 대신 `_v2`, `_v3`, … 처럼 접미사를 붙인 새 버전을 만듭니다(예: `ai_development_research_v2.md`, `AI_발전_보고서_v2.docx`).

## 작업 방식

1. **Todo list 먼저 보고**: 사용자가 프롬프트로 작업을 요청하면, 작업을 시작하기 전에 항상 관련 todo list를 만들어 사용자에게 먼저 보고합니다. 작업 단계, 산출물, 저장 위치, 사용자가 결정해야 할 사항을 포함합니다.
2. **승인 후 실행**: 사용자가 todo list를 읽고 승인한 뒤에만 작업을 시작합니다. 수정을 요청받으면 목록을 고쳐 다시 보고하고, 다시 승인을 기다립니다.
3. **마무리 구조 정리**: 작업 결과로 폴더나 파일이 추가되었다면, 마무리하기 전에 항상 Claude Code가 가장 잘 인식할 수 있는 구조로 정리합니다.
   - 파일을 알맞은 용도별 폴더로 옮기고, 프로젝트 루트에는 산출물을 남기지 않습니다.
   - 새 폴더가 추가되면 이 파일(과 번역본)의 **프로젝트 구조** 트리와 설명을 갱신합니다.
   - 모든 `.md` 파일이 `translate/`에 대응하는 번역본을 갖고 있는지 확인합니다.
   - 완료 보고에 최종 폴더 트리를 포함합니다.

## 마크다운 파일 정책

- 지금부터 작성되는 모든 `.md` 파일은 영어로 작성합니다.
- 모든 `.md` 파일에 대해 `c:\Users\SBS\agent1004\agent_09\translate`에 동일한 상대 경로/파일명으로 한글 번역본을 저장합니다(예: `research/x.md` → `translate/research/x.md`). 번역본은 원본과 동시에 만듭니다.
- `.md` 파일이 수정되거나 삭제될 때마다 변경 사항을 확인하여 `translate`에 있는 대응 번역 파일에도 동일하게 수정(또는 삭제)을 반영하여 두 파일을 동기화합니다.
