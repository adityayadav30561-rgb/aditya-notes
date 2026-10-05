# SSB Interview Question Sheets

Aditya's SSB interview question lists. The user pastes a ChatGPT response, and we filter it down to **questions only**. The user fills in the answers later.

There are two sheets. Each one has its own source file and output file:

| Sheet | Source | Output | What goes in it |
|---|---|---|---|
| PIQ questions | `questions.md` | `SSB_Question_Sheet.docx` | Questions about the user's PIQ entries: family, education, residence, work, sports and so on |
| Service knowledge | `service_knowledge.md` | `SSB_Service_Knowledge.docx` | Armed Forces knowledge: organisation, ranks, commands, roles and similar |

## Lecturette preparation
There is also a separate lecturette document: `lecturettes.md` builds `SSB_Lecturettes.docx`. This one holds lecturette topic flows, not questions. When the user shares a lecturette topic, add it as a new `## Topic name` section. Write one line per stage in the form `- Stage: point → point → point`, keeping the user's points and order. Keep points short and drop tips, sources and commentary. `node build/build.js` puts each topic on its own page as a table (# · Stage · Flow), with each point on its own line. Commit `lecturettes.md` and `SSB_Lecturettes.docx` together.

## Workflow (every time a response is shared)
1. Pick the right sheet. PIQ entries go in the PIQ sheet; service knowledge (Army/Navy/Air Force organisation etc.) goes in the service knowledge sheet.
2. Pull out only the questions. Drop answers, tips, explanations, intros, notes, tables and emojis.
3. Filter them:
   - Remove duplicates and near-duplicates. This includes questions already in that sheet's source file.
   - Turn implied questions into clear, plain questions.
   - Keep the wording short and natural, the way an Interviewing Officer (IO) would ask.
4. Add them to the source file under a fitting `## Section`. Use an existing section if one fits; otherwise add a new section.
5. Run `node build/build.js` (if `require('docx')` fails, run `npm install` first). It rebuilds both sheets. Each one is a single compact table (# · Question · Answer), with a shaded row for each section and an empty Answer column to fill in.
6. Commit the source file and the rebuilt `.docx` together, push, and send the user the updated `.docx`.

## Source file format
```
## Section name
- Question one?
- Question two?
```
Plain text only, with no bold, answers or sub-bullets.
