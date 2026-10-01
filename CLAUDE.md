# SSB Interview Question Sheet

Aditya's SSB interview question list. The user pastes a ChatGPT response, and we filter it down to **questions only**. The questions go into `SSB_Question_Sheet.docx`, and the user fills in the answers later.

## Workflow (every time a response is shared)
1. Pull out only the questions. Drop answers, tips, explanations, intros, notes, tables and emojis.
2. Filter them:
   - Remove duplicates and near-duplicates. This includes questions already in `questions.md`.
   - Turn implied questions into clear, plain questions.
   - Keep the wording short and natural, the way an Interviewing Officer (IO) would ask.
3. Add them to `questions.md` under a fitting `## Section`. Use an existing section if one fits; otherwise add a new section.
4. Run `node build/build.js` (if `require('docx')` fails, run `npm install` first). It regenerates `SSB_Question_Sheet.docx`, with questions numbered and blank space after each one.
5. Commit `questions.md` and the rebuilt `.docx` together, push, and send the user the `.docx`.

## questions.md format
```
## Section name
- Question one?
- Question two?
```
Plain text only, with no bold, answers or sub-bullets.
