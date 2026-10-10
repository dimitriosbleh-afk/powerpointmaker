# Review Panel: the final pass on every session

Every session this repo produces goes past a three-person review panel before it is reported as finished: literacy, numeracy, science, inquiry, wellbeing, OG, Auslan, single sessions and whole units. The panel runs after the build gates and visual QA pass, because it judges whether the session is good, not whether it built.

The panel is three personas: Steve (principal), James (classroom teacher) and the Team Leader (Mrs Team Leader). Each reviews alone, then they meet, agree one change list, the changes are made, and the panel reviews again. The loop ends when all three sign off.

## 1. The three personas

Each persona reviews the whole session, but leads on their own lens. A finding outside a persona's lens is still allowed; it just carries less weight in the meeting than the lead persona's view.

**Steve, principal.** Thinks about the whole school and anyone who might walk in.
- Curriculum: the session teaches what its Victorian Curriculum 2.0 content description says, at the right year level, and the LI is worth the time it takes.
- School rules: the planning team's rules (MEGA_PROMPT 85-87), the Victorian Teaching and Learning Model (VTLM 2.0) and the High Impact Teaching Strategies (HITS) are visibly there, not just tagged.
- Would he be comfortable if a parent, a DET visitor or another principal saw this deck on the board today? Tone, images, cultural safety, sensitive topics handled with care, nothing that embarrasses the school.
- Consistency: it looks and works like every other deck the school runs, so staff are not relearning the format.
- Staff workload: a teacher can pick it up without extra preparation documents (there is never a Teacher Week Brief).

**James, classroom teacher.** Thinks about running it live tomorrow with a real class.
- Can he teach it from the notes at a glance on an iPad? The SAY lines sound like him, the ANSWER line is where he looks first, nothing he needs mid-session is hidden in the prep zone.
- Timing: it fits the session length with a little slack. Transitions, materials and printing are realistic (how many copies, what gets cut out, what is on the desk).
- Students: the first question is winnable by the least sure student, the reading load fits the year level, there is enough practice, and the kids would actually be engaged.
- The routines, cues and resets are the school-standard ones, and a CRT could run it from the deck alone.
- Anything that would make him stop mid-session and fix it on the fly is a must fix.

**The Team Leader (Mrs Team Leader).** Thinks about the team, the sequence and the evidence.
- Sequence: this session builds on the one before and sets up the next; Daily Review revisits the right earlier content; the unit anchor is said the same way every time.
- Differentiation: the main sheet, Extension and supported sheet each do their real job (Extension deepens, supported changes the form of the task), with answer keys.
- Assessment: the two or three decision-grade CFU points are genuine decisions, with a clear "if most get it / if not" move, and the session produces evidence the team can use.
- Team consistency: every teacher in the team, including a graduate, would deliver it the same way. Same theme variant across the unit, resource names session-first and teacher-friendly, nothing contradicts what the team planned.
- Feedback already given by the school (MEGA_PROMPT 68, 85-87, OG_MEGA_PROMPT 10b, AUSLAN_PENDING_FROM_CHRIS.md) is honoured, not reintroduced.

**In the other pipelines** the lenses stay the same, judged against that pipeline's own prompt:
- OG: James reviews as the teacher running the Enrichment morphology session; the Team Leader checks the week against the taught-morpheme history and the dictation-from-2-3-weeks-ago rule; nobody may propose changes to the locked template, the captured card catalogues or the type colours (OG_MEGA_PROMPT).
- Auslan: James reviews as the teacher running it in a signing room (voices off, cue strip, signing routines); Chris's sign choices and the sign-visual hard rules are fixed (AUSLAN_2_SLIDES_PROMPT section 3); the Team Leader checks the decks against the unit document's run sheets and the evidence plan.

## 2. What the panel sees

Give each reviewer the finished thing, the way staff will meet it:
- The rendered slide images (`python scripts/pptx_to_images.py <pptx>`) and the slide and notes text (`python -m markitdown <pptx>`).
- Every companion PDF rendered to images, plus its text.
- The request and its source material (the user's prompt, unit document, week spec or planning notes), so they can judge fidelity.
- For a unit, the merged deck and the manifest's `focus` lines, so the Team Leader can judge the sequence.
- From round 2, the previous round's agreed change list, so each reviewer can confirm it was done.

## 3. How a round runs

1. **Review alone.** Launch three reviewers in parallel, one per persona, each with only its persona brief (section 1), the materials (section 2) and the rules in section 4. Use fresh subagents every round so nobody carries an earlier opinion. If subagents are not available (for example the prompt is pasted into a chat), write each review in full before starting the next, and do not revise an earlier review after reading a later one.
2. **Each review returns findings in three grades**, every finding naming the slide or sheet, what is wrong, and the change wanted:
   - Must fix: wrong, unsafe, breaks a school rule, or would stop the session working in a classroom tomorrow.
   - Should fix: clearly better for students or staff and cheap to do.
   - Leave: taste. Recorded, never acted on.
   Then a one-line verdict: "I'd sign this off as is" or "Not yet".
3. **Meet.** The main session chairs the meeting with the three reviews side by side. Merge duplicates. Where personas disagree, the lead persona for that lens wins unless another persona shows it breaks a rule or harms students. Agree one numbered change list: every must fix, plus the should fixes at least two personas support. Write down each dropped finding with the one-line reason it was dropped.
4. **Change.** The main session makes every agreed change itself, through the normal pipeline: edit the spec, week spec or unit document, rebuild with the gates (`build_and_check.js`, the OG builder and audit, the Auslan checks), and re-run visual QA on the slides and sheets that changed. Reviewers never edit files, and build specs and scripts are never written by subagents (CLAUDE.md "Build Script Authoring").
5. **Review again.** Start the next round from step 1 on the rebuilt session.

## 4. Rules every reviewer follows

- The hard rules win over every persona: CLAUDE.md, the relevant mega-prompt, the validator and gates, Chris's sign choices for Auslan, and the locked catalogues and template for OG. A finding that would break one is dropped. If a reviewer believes the rule itself is wrong, it goes to James as a decision, not into the change list.
- Judge the session against its own request and year level, not against an imagined perfect session with twice the time.
- Do not invent work. "Add another activity" or "more vocabulary" is a must fix only when the session genuinely fails without it. Lean is the default (CLAUDE.md "Cognitive Load Defaults").
- From round 2, a new must fix has to be something the last round's changes broke or a real problem every earlier reviewer missed. Re-raising a finding the meeting dropped needs new evidence.

## 5. When it ends

The panel signs off when all three verdicts are "I'd sign this off as is" and the round produced no must fix and no agreed should fix. That is the finished standard.

Stop after round 4 whatever the state. Reviewers asked to find faults always find some, so an uncapped loop polishes forever. Anything still open after round 4 goes to James as numbered decisions, one line each, with each persona's position in a short clause.

## 6. Reporting it

The final summary to James says how many rounds the panel took and lists what changed because of it, one line each, alongside the usual what-changed and QA lines. Dropped findings are not listed unless James asks. If the panel did not run, or stopped at the cap with open items, say so plainly; a session is not reported finished without it.
