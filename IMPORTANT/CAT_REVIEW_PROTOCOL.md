# CAT Review Protocol: Keep, Revise, Remove

Source: *DCEPS CAT Review Protocol.docx*, Diamond Creek East Primary School, 2026 AIP, 27 October 2026.
This file is the school's protocol, transcribed. When James asks for a CAT (common assessment task) to be
reviewed or improved, follow it. The school's wording wins over anything here; the last section is how
Claude applies it.

## Purpose

Every CAT and its rubric gets one decision: **keep, revise or remove**. A CAT stays only if it tells us
something we act on, and its rubric lets two teachers land on the same judgement.

Run the protocol once per CAT. Decisions feed the assessment schedule review (RAP 5), so a removed CAT
comes off the schedule.

## How the school runs it

Team leaders run it in a SIT meeting, each checking their own team's CATs. About 20 minutes per CAT.

1. Before SIT, gather the task, the rubric and three work samples (one below, one at, one above).
2. Also before SIT, the team leader and one teacher place each sample on the rubric separately, row by
   row, without talking.
3. Compare. Same column on most rows means the rubric works. Note any row where you split.
4. In SIT, answer the five checks: Yes, Partly or No.
5. Make the call and fill one row of the record sheet. Take any revision back to the team.

## The five checks

| # | Check | Yes looks like |
|---|---|---|
| 1 | **Purpose.** Does it inform a decision we make? | Names the decision: reporting judgement, intervention, extension or next teaching. Gives us something no other assessment on the schedule already does. |
| 2 | **Curriculum.** Does it match VC 2.0 for the level? | Every row traces to the VC 2.0 achievement standard. The rubric states the point in the year it is judged against (start or end of year). |
| 3 | **Task fit.** Can the task show every column? | Students have a chance to show Below through to Well Above on every row. Any gap is flagged on the rubric. |
| 4 | **Clear At standard.** Do two teachers agree? | The shaded At column says what students do, not how well. Each column adds one visible step. |
| 5 | **Family ready.** Can a parent read it? | Plain language, a "How to read" section, task examples in italics, and the note that it is one snapshot, not the semester result. |

## Making the call

Check 1 decides keep or remove. Checks 2 to 5 decide keep or revise.

| Decision | When | Then |
|---|---|---|
| **Keep** | Yes on all five | Save it to the CAT bank. Small wording fixes are still a keep. |
| **Revise** | Yes on check 1; Partly or No on any of checks 2 to 5 | Fix only what failed. Name who and by when. Re-run the sample test once fixed. |
| **Remove** | No on check 1, or fixing it would take longer than writing a new one | Take it off the assessment schedule and tell the RAP 5 review. |

## Record sheet (one row per CAT)

| CAT | Level and term | Checks not met | Decision | What changes | Who | By |
|---|---|---|---|---|---|---|
| The Strange Discovery (writing) | Grade 4, Term 1 | 4 | Revise | Rewrite the word choice row | | |
| Decimals, Factors and Negative Numbers | Grade 5, Term 4 | 3 | Revise | Add two task questions | | |

## Worked example (the house standard)

Both exemplars are one fix away from a keep. Both already show the house standard: **shaded At column,
point in the year stated, task examples in italics, gaps flagged honestly, and the snapshot note.**

| Check | Grade 4 Writing: The Strange Discovery | Grade 5 Maths: Decimals, Factors and Negative Numbers |
|---|---|---|
| 1 Purpose | Yes. Start-of-year writing baseline. Speaks to the Years 3-6 writing gender gap (15% of males above expected vs 33% of females, 2025). | Yes. End-of-year reporting evidence. Speaks to the maths gender gap (19% of females above expected vs 36% of males, 2025). |
| 2 Curriculum | Yes. At is set at end of Grade 3 (progression point 3.0) because it runs in Term 1. | Yes. At is set at end of Grade 5 (progression point 5.0); Above and Well Above reach into Grade 6. Team confirms rows against VC 2.0. |
| 3 Task fit | Yes. An open story lets students show every column. | Partly. No questions for Towards on factors or Well Above on estimating. A student just below standard on factors cannot be placed. |
| 4 Clear At standard | Partly. Word choice barely steps: "several precise words", "precise words", "precise vocabulary". Two teachers will split on it. | Yes. Worked examples (0.601 vs 0.70, factor pairs of 24) make At easy to agree on. |
| 5 Family ready | Yes. | Yes. |
| Decision | Revise the rubric. Rewrite the word choice row so each column adds one visible step. | Revise the task. Add one Towards question (multiples, does it divide evenly) and one Well Above question (estimate, then check a decimal calculation). |

Sources named in the protocol: Grade 4 Writing Rubric for Families; Grade 5 Maths Rubric for Families;
Diamond Creek East PS 2026 School Review Report (teacher judgement gender data).

## How Claude applies it

- **Default output is two short sentences per CAT**: the decision, the one or two macro reasons (wrong kind
  of task, too long, too broad, visually weak next to what the pipeline makes now), and what to replace or
  fix it with. No nitpicks. Give the full five-check breakdown and record-sheet row only when James asks.
- **Shared continuum across grades is fine**: a Grade 6 rubric that moves the Grade 5 columns down one
  level (Grade 5 Above = Grade 6 At) is normal practice. Do not flag it.
- **Elastik first for other writing types**: any writing CAT that isn't narrative or persuasive (information
  report, procedure, explanation and so on) waits until Elastik supports marking that text type. Say so in
  the verdict (James, 27 Oct 2026).
- **Band at the point in the year the task runs**, not the end of year (school decision, per James 27 Oct 2026;
  a Term 1 Grade 5 At is 4.0, Term 2 is 4.5). Flag any CAT banded at end of year.
- **Check 1 is the gate.** If the purpose is No, recommend Remove and stop; do not polish a CAT the school
  should drop. If it duplicates another assessment on the schedule, say which one, or say the schedule was
  not supplied and the team must confirm.
- **Fix only what failed.** A Revise changes the failing check and nothing else. Do not rewrite passing rows
  for style; "small wording fixes are still a keep".
- **Check 2**: trace every rubric row to the VC 2.0 achievement standard for the level and name the
  standard it maps to. Confirm the rubric states start or end of year, and that At matches that point
  (a Term 1 CAT judges At against the end of the previous year).
- **Check 3**: walk every rubric row and name the task item a student would use to show each column,
  Below to Well Above. An empty cell is a gap: either add a task item or flag the gap on the rubric.
- **Check 4 (the sample test)**: Claude cannot replace the two-teacher blind placement. If work samples are
  supplied, place each one row by row and show the reasoning, and name rows where a reasonable second
  teacher could land a column away. If none are supplied, judge the wording: the At column says what
  students *do* (observable), not how well (no "good", "effective", "precise" carrying the whole
  distinction), and each column adds one visible step. Always tell the team to re-run the real sample
  test after a revision.
- **Check 5**: plain language, a "How to read" section, task examples in italics, and the snapshot note
  ("one snapshot, not the semester result"). Missing any of the four is Partly.
- **Rewrites follow the worked example**: when rewriting a rubric row, each column adds one visible step a
  parent could see in the work, with an italic task example.
- Revised rubrics or tasks as files: Word via `scripts/md_to_docx.py` (see the docx toolchain memory), unless
  James asks for another format.
