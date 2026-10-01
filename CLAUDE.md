# SSB Interview Master Notes

Aditya's SSB interview prep notes. The user pastes long notes on a topic, and we turn them into **concise notes**. Those notes go into one Word file, `SSB_Master_Notes.docx`, that grows to hundreds of topics.

## Workflow (every time a topic is shared)
1. Make the notes concise (rules below). Save them as `topics/<short-slug>.md`. If the topic already exists, update that file and its `updated:` date instead of making a second file.
2. Run `node build/build.js` (if `require('docx')` fails, run `npm install` first). It regenerates `SSB_Master_Notes.docx` with the index, categories and every topic.
3. Commit the `.md` and the rebuilt `.docx` together and push.

## Topic file format
```
---
title: G20 (Group of Twenty)
category: International Organisations & Groupings
updated: YYYY-MM-DD
---

## Section heading
- **Key term**: compressed fact · fact · fact
  - optional sub-bullet (2-space indent)
| Col | Col |
|---|---|
| cell | cell |
```
Only these elements are supported: `##` headings, `-` bullets (2 levels), pipe tables, plain lines, and `**bold**`.
Use the category names in `CATEGORY_ORDER` in `build/build.js`. A new category is allowed, and it is added after the listed ones.

## Concise-notes style (the user wants these SHORT: about 1–2 pages per topic)
- Keep **every fact**: names, numbers, dates, places, themes, acronyms. Cut the filler, repetition, "don't say X" padding, long model-answer prose, rapid-fire tables that repeat facts, and ASCII mind maps.
- Write in telegraphic fragments, not sentences. Use `→` for cause or sequence, `·` to separate items, `=` for definitions, `vs` for contrasts.
- **Bold** the must-know facts (numbers, dates, names). Put ★ on high-priority sections.
- Usual section order. Skip sections that don't apply:
  1. Core facts
  2. How it works / background
  3. Timeline or table, only if it's denser than bullets
  4. India link ★
  5. Key issues, one line each
  6. Comparisons
  7. India angle / defence-security link
  8. Strengths vs limitations
  9. IO questions → answer lines, each answer in about one line
- Answers stay balanced and SSB-safe: India's position, international law, no extreme opinions.
- Keep the source's facts. Don't add new claims unless you are sure of them, and mark time-sensitive facts with the "as of" date in `updated`.
