(c) 2026 James Hooke. Confidential. Internal use only. Not for redistribution.

# Auslan Teacher Layer Prompt v1.0
## Step 3 of the Auslan pipeline | Turns the unit document into the documents the teacher plans, teaches and records from | Runs inside the PPTX lesson generator repo

# 0. Why this step exists

The unit document from step 1 is the deep record: about a hundred pages, correct, and too much to teach from. The principal's instruction is to keep it and to add a short pick-up layer generated from it. The teacher's instruction is what that layer looks like: tables, landscape, left to right, the same labelled parts in the same place every time, so he finds things by shape. He has ADHD and he said so; a document he has to read top to bottom is a document he reads once.

This step produces five things from one unit document. None of them adds content. Every word in them is lifted from the unit document; if something is missing there, the fix is in step 1, not here.

1. The unit-at-a-glance: one landscape page, one row per week.
2. The lesson pages: one landscape page per lesson, stages left to right.
3. The weekly planner: one landscape page per term, all cohorts side by side, linked to the lesson pages.
4. The evidence tracker: one spreadsheet per unit, one sheet per class, designed for the evidence and imported into iDoceo as is.
5. The CRT review pack: the unit's review lesson as a standalone document a non-signing teacher can run.

Authority order is the same as step 2: the pasted unit document, then this prompt, then CLAUDE.md.

# 1. Inputs

Paste with this prompt:

- The unit document, whole. Sections 7, 8, 9, 10, 13 and 14 are the ones this step reads.
- Which cohorts to produce for (Challenge and Enrichment get separate lesson pages from a shared unit document).
- The term folder: the folder name where every Term 4 file will live, so hyperlinks resolve. If not supplied, use `output/<UnitFolder>/Teacher Layer/` and say so.
- For the weekly planner only: every unit document running that term, or their section 8 tables, one per cohort.
- Class lists, if the teacher wants names pre-filled in the tracker. Otherwise the tracker ships with numbered blank rows.

# 2. Shared rules for every document here

- Landscape A4, about 1.5 cm margins, fixed-layout tables with set column widths. Tables are the format; prose is the exception.
- One row per lesson or week. The same columns in the same order on every page of the same document, so the eye lands in the same place each time.
- Bold label, then content. Nothing over two lines in a cell without becoming a list.
- The unit document's icons come across. They are the only non-ASCII characters allowed, and each one marks the same thing it marks in the unit document.
- Plain ASCII otherwise: no em dashes, en dashes, curly quotes or ellipsis characters. Say: lines keep the Say: label with no quotation marks.
- Every file name is human-readable, unit first, spaces not underscores: `Deaf Sports Enrichment At A Glance.docx`, `Deaf Sports Enrichment Lesson 3.docx`, `Term 4 Weekly Planner.docx`, `Deaf Sports Enrichment Evidence Tracker.xlsx`, `Deaf Sports CRT Review Pack.docx`.
- Word documents are built with `python scripts/md_to_docx.py in.md out.docx --landscape --plain --fontsize=9`. That is the tested command for every document in this section: landscape A4, no title page or contents field, 9 point body, 1cm top and bottom margins. Spreadsheets are built with openpyxl. There is no pandoc here and no pdftoppm; render with LibreOffice at `/c/Program Files/LibreOffice/program/soffice.exe --headless --convert-to pdf` then PyMuPDF, and look at every page before handing it over. A collapsed column is invisible in the file and obvious in the render.
- Nothing here is student-facing. Nothing here is a weekly summary sheet in the sense the school banned for classroom teachers; these are the specialist's own planning views, requested by the principal and the teacher by name.

# 3. The unit-at-a-glance

One landscape page per unit per cohort. Source: the unit document's section 8 table, which was written to be this page.

Columns, exactly, left to right: Week | Topic | Vocab | Learning intention and success criteria | Activities | Assessment and materials.

- Vocab is the five production signs for that week, gloss and English meaning, one per line.
- Learning intention and success criteria cell: the bold labels on their own lines, one criterion per line.
- Activities: the game by appendix number and title, and the You Do task in one line.
- Assessment and materials: the evidence piece landing that week if any (Piece 1, Piece 2, exit rotation), then the materials line from the run sheet.
- Below the table, on a second page, the differentiation table: one row per week, Support and Extend lifted from the weekly plans, then the band calibration, the lost-lesson rule and the full-hour additions as labelled lines. Printed double sided that is one sheet: the term on the front, how to adapt it on the back.
- The review weeks are one labelled paragraph under the table, not table rows. They are the same lesson run twice and two more rows push the term off page one.
- If the term opens with a completion block or closes with review weeks, they are rows in this table too, so the page shows the whole term.
- Header: unit title, cohort, term, the unit anchor in its exact words, the evidence-complete week and the report week.

The teaching weeks must fit on one page at 9 point. If they do not, the unit document's section 8 has too much in its cells: shorten the learning intention to one rendered line and trim the criteria, never the font below 9. Two pages total is correct, teaching weeks on the first and differentiation on the second; three is a failure.

# 4. The lesson pages

One landscape page per lesson per cohort. Source: the weekly plan in section 9, and only that plan.

Layout: a single table running left to right, one column per stage in teaching order: Do Now | LI and SC | I Do | We Do | You Do | Exit ticket. Each column head carries the stage icon, the stage name and its minutes.

Inside each column, top to bottom, in this fixed order and with these bold labels wherever the stage has them:

- What happens (one line from the run sheet)
- Say: lines the stage scripts. A page may drop a whole Say: line, or its opening clause, to fit. It may never reword one: every word it keeps is identical to the unit document, and if a line is shortened on the page it is shortened in the unit document too, so the two never disagree.
- The game, by number and title, or the task
- Check: the decision point's cue, the 80 percent move and the pivot, each on its own line (primary decision point only; the light check is one line)
- Protocol (the week's Deaf-friendly protocol, in the stage where it is practised)

**Nothing goes under the table.** A strip of labelled lines below it costs five or six rendered lines and pushes every lesson onto a second page. Each adaptation goes inside the stage column where it is used, appended after a blank line at the bottom of that cell:

- LI and SC column, the shortest one, takes **Band calibration**, **If you are running late**, **If you have the full hour** and **No-prep fallback**.
- We Do column takes the week's **Deaf-friendly protocol**.
- You Do column takes **Support** and **Extend**.
- Exit ticket column takes the **Care note**, when the plan has one.

Materials goes in the header line above the table. The journal reflection, the cultural question and the exit rotation are already the Exit ticket column's own content.

Header: unit title, cohort, lesson number and title, week, the anchor in its exact words, the learning intention.

The page is complete when a teacher who has read the unit document once could run the lesson from this page alone. It is wrong when it says less than the plan in a place the plan was specific: a pivot reduced to "re-teach" has lost the plan.

**One lesson is one page.** If a lesson runs over, the I Do column is always the cause. Cut in this order: the light check to one clause, then any Say: line that is not the anchor restatement, the deliberate error or a cue script. Never drop a whole stage and never go below 9 point.

Deliver the seven pages as one file with a page break between lessons. Split them into one file per lesson only when the weekly planner's links need a target.

# 5. The weekly planner

One landscape page per term. Source: the section 8 tables of every unit running that term, one per cohort, plus the completion block and review week rows.

Columns, exactly, left to right, in the school's cohort order: Term and week | Foundation | Discovery | Challenge | Enrichment.

Each cohort cell for a week carries, in this order, bold labels on their own lines:

- Topic
- Learning intention
- Success criteria (three lines)
- Main activity (the game by number and title, or the task)
- Link: a hyperlink to that cohort's lesson page for that week, relative to the term folder

Discovery is one column; where Year 1 and Year 2 differ, the calibration note is one line under the success criteria, not a second column.

Timetable cells are not generated. The teacher's timetable changes; the planner shows what each cohort is doing in a given week, and he reads which class is in front of him.

Rows for weeks with no lesson for a cohort (assessment finished, report writing) carry the review lesson's game set for that week or the word "Review", never a blank.

If the planner runs past one page, reduce the success criteria to the second criterion only and say so in the header; never drop the links.

# 6. The evidence tracker

One workbook per unit per cohort, `<Unit> <Cohort> Evidence Tracker.xlsx`, one sheet per class, plus a sheet named Key. Source: section 13 of the unit document.

Each class sheet:

- Column A: student name (pre-filled if a class list was supplied, otherwise numbered 1 to the class size).
- Column B: Rotation week, the week that student is due for the exit ticket, assigned so that each week has one quarter of the class and every student is due at least twice before the evidence-complete week. Students are not told this; the sheet is teacher-only.
- Then one column per evidence piece: Piece 1 filmed task (rubric level), Piece 2 receptive test (score out of the item count), Cultural reflections (count of weeks a reflection was recorded, Challenge and Enrichment only), Exit rotation (dates seen).
- Then a Notes column.
- Then one column per ACARA checklist row for the cohort's year levels (five per year, headed Y4.1 to Y4.5 with the full statement in a comment), rated for the term. These are the report descriptors: the checklist is what he reports on every semester, so this block is the report, as a class grid, imported into iDoceo unchanged.
- Data validation on the Piece 1 and checklist columns: a drop-down of E / C / P (Emerging, Consolidating, Proficient), the checklist's own scale. No four-level scale anywhere.
- Frozen header row and frozen name column. Column widths set so every header reads without wrapping past two lines.

The Key sheet: the 13.6 evidence map in full (checklist row wording, lessons, evidence piece, what E, C and P look like in this unit, ACARA code, Victorian mapping); and one line saying the sheet is designed to import into iDoceo unchanged.

Do not build formulas that compute a report level. The teacher decides the level; the sheet holds the evidence.

# 7. The CRT review pack

One document per unit, `<Unit> CRT Review Pack.docx`, portrait is acceptable here because it is read by somebody else. Source: the CRT review lesson in section 9 and the games it names in section 10.

It is written to a teacher with no Auslan, who has the class for one lesson and needs to run it without modelling a sign. It carries:

1. One page: what this lesson is (game-based review of signs the class has already been taught, nothing new), the three rules of the room (voices off during games, eyes on the signer, wait for eyes before you start), and what to do if a student asks how to sign something (open Sign It! to the page on the list; do not guess).
2. The term vocabulary list: English word, gloss, Sign It! page, in a table. No sign images.
3. Each game in the review set as a full-page rules card: the game bank entry rewritten as numbered steps for a stand-in, the team and points structure, the timing, the safety lines, and one line saying which signs it uses from the list.
4. A points sheet for the board.
5. One line at the end: "Leave a note for the Auslan teacher: which games ran, which signs the class struggled with."

Kahoot and similar online quizzes are not built by this pipeline. If the unit document names one the teacher already runs, the pack lists it by name with its link as an optional round; it never creates one.

# 8. Request format

```
Produce: at-a-glance, lesson pages, tracker, CRT pack   (or: weekly planner)
Unit prefix: deafsport
Cohorts: Challenge, Enrichment
Term folder: <path or name>
Class lists: <none / pasted>

<paste the unit document>
```

For the weekly planner, paste every cohort's section 8 table for the term and name the term folder so the links resolve.

# 9. QA

- Render every Word document and the first sheet of the workbook to images. Look at one lesson page, the at-a-glance and the planner in full. A column that has collapsed, a cell that has wrapped past its row, or a link that points at a file that does not exist is a failed build.
- Open the workbook and test one drop-down.
- Diff every Say: line on a lesson page against the unit document. A dropped line or a dropped opening clause is fine; a reworded one is a defect, and so is a shortened line that was not shortened in the unit document as well.
- Check that every lesson page's stage minutes total 50 and match the run sheet.
- List in the final summary every hyperlink target that does not yet exist in the term folder, so James knows which files still have to land there.
- Run the review panel (`IMPORTANT/REVIEW_PANEL.md`) on the finished documents, judged as the pages a teacher plans and teaches from, and rebuild its agreed changes. The Say: lines stay word for word with the unit document; a panel finding about one goes back to the unit document, not into these pages.
