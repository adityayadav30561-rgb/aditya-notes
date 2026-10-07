# SSB Interview Question Sheets

Aditya's SSB interview question lists. The user pastes a ChatGPT response, and we filter it down to **questions only**. Each question also carries a short model answer (see below); where only the user knows the facts, the answer says "(Fill.)".

There are two sheets. Each one has its own source file and output file:

| Sheet | Source | Output | What goes in it |
|---|---|---|---|
| PIQ questions | `questions.md` | `SSB_Question_Sheet.docx` | Questions about the user's PIQ entries: family, education, residence, work, sports and so on |
| Service knowledge | `service_knowledge.md` | `SSB_Service_Knowledge.docx` | Armed Forces knowledge: organisation, ranks, commands, roles and similar |

## Workflow (every time a response is shared)
1. Pick the right sheet. PIQ entries go in the PIQ sheet; service knowledge (Army/Navy/Air Force organisation etc.) goes in the service knowledge sheet.
2. Pull out only the questions. Drop answers, tips, explanations, intros, notes, tables and emojis.
3. Filter them:
   - Remove duplicates and near-duplicates. This includes questions already in that sheet's source file.
   - Turn implied questions into clear, plain questions.
   - Keep the wording short and natural, the way an Interviewing Officer (IO) would ask.
4. Add them to the source file under a fitting `## Section`. Use an existing section if one fits; otherwise add a new section.
5. Add a short answer to each new question after ` :: ` (written for a compact printout). Use the user's known PIQ facts, a general model answer for opinion questions, "(Fill.)" or "(Fill: hint)" where only the user knows, and "(verify)" for facts that change often.
6. Run `node build/build.js` (if `require('docx')` fails, run `npm install` first). It rebuilds both sheets. Each one is a single compact table (# · Question · Answer), with a shaded row for each section. It is laid out for printing at 7 pt.
7. Commit the source file and the rebuilt `.docx` together, push, and send the user the updated `.docx`.

## Source file format
```
## Section name
- Question one? :: Short answer.
- Question two? :: (Fill.)
```
Plain text only, with no bold or sub-bullets.

## Bulk answering helpers
- `python3 build/number.py <source.md> [from] [to]` prints questions with their sheet numbers.
- Answers can be written as `N|answer` lines in `answers/*.txt` and merged with `python3 build/merge_answers.py <source.md> <answer files...>`. The source `.md` stays the source of truth.
