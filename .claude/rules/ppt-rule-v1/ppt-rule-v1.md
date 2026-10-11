# PPT Rules (v1)

These rules apply whenever a PowerPoint (.pptx) presentation is created or edited in this project, including when the `mk-ppt` skill is used.

## 1. Default slide count

- Build **8 slides** by default.
- Use a different number only when the user explicitly asks for one.
- Plan the outline (title of each slide + its type) to fit the slide count before building.

## 2. Design reference

- Before building, read the design `.md` file(s) the user provides (colors, fonts, layout, tone, etc.) and follow them.
- When the design file conflicts with the default theme of the `mk-ppt` skill, the user's design file wins.
- If no design file has been provided, ask the user for one before building; use the `mk-ppt` default theme only if the user says to.

## 3. Readability: no overflow, no overlap

- Never let text run outside the slide or its text box.
- Never let text and shapes, or shapes and other shapes, overlap in a way that hurts readability.
- Keep text short: cut wording, reduce bullet count, or split content across slides instead of shrinking the font below a readable size.
- After building, render the slides to images and check every slide for overflow, overlap, broken line wraps and crowding; fix the spec and rebuild until none remain.
