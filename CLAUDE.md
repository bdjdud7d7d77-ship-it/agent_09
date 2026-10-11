# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Identity

- Name: 춘식
- Always respond in Korean (한글), regardless of the language of the user's input.

## Project Structure

This is a research-and-report workspace, not a code project.

```
agent_09/
├── CLAUDE.md      # Project rules (this file)
├── .gitignore     # Ignores Office lock files (~$*) and OS junk
├── .claude/
│   ├── rules/
│   │   └── ppt-rule-v1/   # Always-loaded rules for building .pptx decks
│   │       └── ppt-rule-v1.md
│   └── skills/
│       └── mk-ppt/    # Project-only skill: builds .pptx decks with python-pptx
│           ├── SKILL.md
│           ├── references/python_pptx_api.md
│           └── scripts/build_pptx.py
├── references/    # Design references (.md, English) for deliverables
│   └── ppt/
│       └── slide-design-system.md   # Orange slide design system (colors, fonts, layout)
├── research/      # Research findings as .md (English)
├── report/        # Final reports as .docx (Korean)
├── slides/        # Presentations as .pptx (Korean), built with the mk-ppt skill
└── translate/     # Korean translations of every .md, mirroring relative paths
```

- Save research results (`.md`) in `research/`.
- Save reports (`.docx`) in `report/`. Do not leave deliverables in the project root.
- Save presentations (`.pptx`) in `slides/`, using the `mk-ppt` skill (`.claude/skills/mk-ppt/`).
- For any `.pptx` work, apply `.claude/rules/ppt-rule-v1/ppt-rule-v1.md` first. Priority: ppt-rule-v1 → the user's design `.md` → `mk-ppt` defaults.
- Never overwrite an existing deliverable unless the user asks. Create a new version with a `_v2`, `_v3`, … suffix instead (e.g. `ai_development_research_v2.md`, `AI_발전_보고서_v2.docx`).

## Workflow

1. **Todo list first**: Whenever the user requests a task through a prompt, always create a related todo list and report it to the user before doing any work. Include the steps, the deliverables, where they will be saved, and anything that needs the user's decision.
2. **Execute only after approval**: Start the work only after the user has read and approved the todo list. If the user asks for changes, revise the list, report it again, and wait for approval again.
3. **Finish by tidying the structure**: If the work added any folders or files, always reorganize them into the structure Claude Code recognizes best before finishing:
   - Move files into the right purpose-based folder; leave no deliverables in the project root.
   - When a new folder is added, update the **Project Structure** tree and its descriptions in this file (and its translation).
   - Check that every `.md` file has a matching translation in `translate/`.
   - Include the final folder tree in the completion report.

## Markdown File Policy

- All `.md` files created from now on must be written in English.
- For every `.md` file, save a Korean translation at `c:\Users\SBS\agent1004\agent_09\translate`, mirroring the same relative path/filename (e.g. `research/x.md` → `translate/research/x.md`). Create the translation at the same time as the original.
- Whenever a `.md` file is modified or deleted, check for the change and apply the same update (or deletion) to its corresponding translation file in `translate`, keeping both in sync.
