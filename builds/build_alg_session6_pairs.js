"use strict";

// Algebra Unit — Session 6: Pairs of Unknowns in Complex Equations
// Session 6 of 6, Grade 5/6 Numeracy, Variant 0
// DR: Mixed Review (factor pairs, LCM, estimation)
// Fluency: Order of Operations Countdown
// VC2M6A02: recognising that [] + [] = 12 has multiple solutions

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 0);
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 6;
const UNIT_TITLE = "Algebra: Finding Unknown Values";
const FOOTER = "Session 6 of 6 | Algebra: Unknown Values | Year 5/6 Maths";
const OUT_DIR = "output/ALG6_Session6_Pairs_of_Unknowns";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "6 problems: additive pairs + multi-operation unknowns.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Tables pre-structured for systematic listing.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Unknown pairs and patterns — discovering rules and coordinate connections.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- This is it -- our final session. You have learned number properties, order of operations, brackets, and how to find unknowns. Today we combine EVERYTHING.
- Today's challenge: equations with TWO unknowns. And sometimes, there is more than one right answer.

DO:
- Display the title slide. Whiteboards ready.

TEACHER NOTES:
Lesson 6 is the culminating lesson. It addresses VC2M6A02: "finding unknown values in numerical equations involving one or more operations" and specifically the elaboration about multiple solutions: "recognising that [] + [] = 12 has multiple solutions and could be 1 + 11, 2 + 10, 3 + 9, ..." This lesson also introduces equations like 6 + 4 x 8 = 6 x T + [] where students apply all prior learning.

WATCH FOR:
- Students who are excited by the open-ended nature of multiple solutions.
- Students who are unsettled by "more than one right answer" -- reassure that this is normal in algebra.

[Maths: Planning | VTLM 2.0: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review: Mixed review -- we are cycling back across all our Daily Review topics.
- Problem 1: Find the unknown pair: [] x T = 24 where [] > T. Give TWO possible answers. [e.g., 8x3, 6x4, 12x2, 24x1]
- Problem 2: What is the LCM of 6 and 8? [24]
- Problem 3: Estimate 489 + 312 by rounding to hundreds. [500 + 300 = 800]

DO:
- Display the slide. Students work on whiteboards, 20 seconds per problem. "Show me!"
- Problem 1 previews today's lesson -- multiple solutions for a pair of unknowns.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Show me your TWO solutions for Problem 1!"
- Scan for: two valid factor pairs on 80%+ of boards.
PROCEED: If 80%+ find two valid pairs, move to Fluency.
PIVOT: If students can only find one pair (or give the same pair twice like 4x6 and 6x4), reteach: "[] x T = 24 with [] > T means the first number is bigger. Let us list: 24x1, 12x2, 8x3, 6x4. Four pairs!" Re-check: "[] x T = 36, [] > T. Find two pairs."

TEACHER NOTES:
Daily Review cycles back across the unit's topics: factor pairs (L1-2), LCM (NP unit), estimation (L4-5). Problem 1 is deliberately chosen to preview the lesson -- multiple valid solutions for a pair of unknowns.

WATCH FOR:
- Students who list 6x4 and 4x6 as different -- clarify the [] > T constraint.
- Students who struggle with LCM -- brief reteach if needed, but do not derail.

[Maths: Daily Review | VTLM 2.0: Retention & Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers.
- 1) Possible pairs ([] > T): 24x1, 12x2, 8x3, 6x4.
- 2) LCM(6,8) = 24.
- 3) About 800 (500+300).

DO:
- Click to reveal answers. Students self-check.

TEACHER NOTES:
Problem 1 previews the multiple-solutions concept central to today's lesson.

WATCH FOR:
- Students who found all 4 factor pairs -- strong preparation for today.

[Maths: Daily Review Answers | VTLM 2.0: Retention & Recall]`;

const NOTES_FLUENCY = `SAY:
- Countdown challenge! I will put an expression on the board. Evaluate it -- race against the timer. 20 seconds each.
- Expression 1: 3 + 2 x 5. [13]
- Expression 2: (4 + 1) x 6. [30]
- Expression 3: 12 / 3 + 5 x 2. [14]
- Expression 4: (8 - 3) x (2 + 4). [30]
- Speed round -- show answer on board when timer hits zero.

DO:
- Display expressions one at a time with a 20-second countdown.
- "Show me!" after each. Quick pace -- this is fluency, not teaching.

TEACHER NOTES:
This fluency round consolidates Lesson 5's order of operations content. Expressions 3 and 4 are more complex -- E3 has two operations at the same level (/ and x) interleaved with +, and E4 has two sets of brackets.

WATCH FOR:
- Expression 3 is the trickiest -- students must evaluate 12 / 3 AND 5 x 2 before adding.
- Students who ace all four -- they are ready for today's challenges.

[Maths: Fluency | VTLM 2.0: Retention & Recall]`;

const NOTES_LI_SC = `SAY:
- Read from slide: "We are learning to find pairs of unknown values in equations with multiple operations."
- SC1: I can find multiple solutions when an equation has two unknowns.
- SC2: I can use number properties and order of operations to find unknowns in complex equations.
- SC3: I can systematically list and check all possible solutions.
- SC3 is the big one -- being SYSTEMATIC means you do not miss any solutions.

DO:
- Display the slide. Point to each SC.

TEACHER NOTES:
SC1 introduces the concept of multiple solutions -- a significant shift from earlier lessons where each equation had exactly one answer. SC2 connects to all prior content. SC3 emphasises systematic thinking -- listing all possibilities rather than guessing.

WATCH FOR:
- Students who seem worried about "multiple answers" -- normalise: "In real algebra, many equations have more than one answer. Your job is to find ALL of them."

[Maths: Planning -- Curriculum Alignment | VTLM 2.0: Planning]`;

const NOTES_IDO1 = `SAY:
- Let us start with a puzzle: [] + T = 12. What could [] and T be?
- I will be SYSTEMATIC. Start with [] = 1. Then T = 11. Next, [] = 2, T = 10. Then 3 + 9, 4 + 8, 5 + 7, 6 + 6.
- That is 6 pairs if we use whole numbers from 1 to 11 (assuming both are positive).
- If [] = T, we get 6 + 6 = 12. If [] is not equal to T, we get 5 pairs.
- Now: What about [] + [] = 12? This is different -- both unknowns are the SAME number.
- [] + [] = 12 means 2 x [] = 12, so [] = 6. Only ONE solution.
- See the difference? [] + T (different letters) means different values are OK. [] + [] (same letter) means SAME value.

DO:
- Display the pairs systematically in a table on the slide.
- Highlight the distinction between [] + T and [] + [].

TEACHER NOTES:
This is the direct curriculum example: "recognising that [] + [] = 12 has multiple solutions and could be 1 + 11, 2 + 10, 3 + 9, ..." The systematic listing demonstrates the organised approach students need.

MISCONCEPTIONS:
- Misconception: "[] + T = 12 has only one answer."
  Why: All previous lessons had exactly one answer. Students expect one right answer.
  Impact: Students will stop after finding one pair, missing the richness of the problem.
  Quick correction: "Finding ONE pair is great. But can you find ANOTHER? And another? How many can you find? Be systematic!"

WATCH FOR:
- Students who give one pair and stop -- push: "Can you find them ALL?"
- Students who raise 0 + 12 -- validate if you accept zero in your class context.

[Maths: Launch -- Explicit Instruction (I Do) | VTLM 2.0: Explicit Explanation & Modelling]`;

const NOTES_IDO2 = `SAY:
- Now the big one. Ready? 6 + 4 x 8 = 6 x T + []. Find T and [].
- First, I evaluate the left side. Remember order of operations: 4 x 8 = 32, then 6 + 32 = 38.
- So the equation becomes: 38 = 6 x T + [].
- Strategy: Try values for T and see what [] must be.
- If T = 1: 6 x 1 + [] = 38. 6 + [] = 38. [] = 32.
- If T = 2: 6 x 2 + [] = 38. 12 + [] = 38. [] = 26.
- If T = 3: 6 x 3 + [] = 38. 18 + [] = 38. [] = 20.
- If T = 6: 6 x 6 + [] = 38. 36 + [] = 38. [] = 2.
- Verify T = 6, [] = 2: 6 x 6 + 2 = 36 + 2 = 38. Yes.
- This equation has MANY solutions. But some are more interesting. T = 6, [] = 2 is nice because the numbers are small.

DO:
- Display the equation. Show the systematic substitution.
- Build a table on the slide: T, 6 x T, [] = 38 - 6T.

TEACHER NOTES:
This equation combines everything: order of operations (evaluating the left side), substitution (trying values for T), and inverse operations (finding [] given T). Multiple valid solutions exist -- restrict to positive whole numbers at this level.

MISCONCEPTIONS:
- Misconception: "There should be one answer."
  Why: Prior experience. Every equation before this had one solution.
  Impact: Students stop searching after finding any valid pair.
  Quick correction: "With TWO unknowns, you often get MANY solutions. Your job is to find them systematically."

WATCH FOR:
- Students overwhelmed by the complexity -- break it down: "Left side first. Then try numbers for T."
- Students who try T = 7: 6 x 7 = 42 > 38, so [] would be negative -- discuss the boundary.

[Maths: Launch -- Explicit Instruction (I Do) | VTLM 2.0: Explicit Explanation & Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. [] x T = 20, both whole numbers greater than 1. On your whiteboard, list ALL possible pairs.

DO:
- Display question. 30 seconds. "Show me!"
- Scan for systematic listing on boards.

CFU CHECKPOINT:
Technique: Show Me Boards + Cold Call
Script:
- "List ALL pairs on your board. Show me!"
- Scan for: at least 2 valid pairs on 80%+.
PROCEED: If 80%+ find at least 2 pairs with systematic listing, move to We Do.
PIVOT: If students find only one pair or are not systematic, reteach: "Start from the smallest factor greater than 1. Does 2 x ? = 20? Yes, 2 x 10. Does 3 x ? = 20? No. Does 4 x ? = 20? Yes, 4 x 5. Keep going until the factors start repeating." Re-check: "[] x T = 18, both > 1."

TEACHER NOTES:
This CFU checks whether students can systematically find multiple solutions. The constraint "greater than 1" limits the search space and prevents students from defaulting to 1 x 20.

WATCH FOR:
- Students who list 2x10 and 10x2 as the same -- discuss whether order matters.
- Students who try every number from 2 to 19 -- efficient but time-consuming. Guide them to stop when factors repeat.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_CFU_A = `SAY:
- The pairs are: 2x10, 4x5, 5x4, 10x2. If order matters, that is 4 pairs. If [] > T, just 2: 10x2 and 5x4.
- Follow-up: Why is 1x20 not included? [Because both must be greater than 1]

DO:
- Click to reveal. Discuss.

TEACHER NOTES:
Constraints limit the solution set. Understanding constraints is an algebraic skill.

WATCH FOR:
- Students who found all pairs systematically -- excellent.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_WEDO1_Q = `SAY:
- Together: [] + T = 15, where [] > T and both are whole numbers >= 1.
- Ask: What is the smallest T can be? [1] And then [] = ? [14]
- Let us list systematically. T = 1, [] = 14. T = 2, [] = 13. Keep going...
- Ask: When do we stop? [When [] = T would happen -- at 7 + 8, since 8 > 7. If [] must be GREATER, 8 + 7 is the last valid pair]
- So we have: 14+1, 13+2, 12+3, 11+4, 10+5, 9+6, 8+7. Seven pairs.

DO:
- Display equation. Build the list together through Cold Call. Click to reveal full table.

CFU CHECKPOINT:
Technique: Cold Call
Script:
- "[Name], give me one pair." "[Name], give me another."
PROCEED: If students contribute correct pairs fluently, move to PP2.
PIVOT: If students repeat pairs or cannot find new ones, reteach using the systematic approach from smallest T upward.

TEACHER NOTES:
This problem parallels I Do 1 ([] + T = 12) but with a different target. The key learning is the systematic listing strategy and knowing when to stop.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Partially filled table: "T = 1, [] = ___. T = 2, [] = ___. ..." Students fill in.
EXTENDING PROMPT:
- Task: "What if we allow [] = T? How many MORE solutions do we get? What if we allow 0?"

WATCH FOR:
- Students who forget the constraint [] > T and list duplicates.
- Students who see the pattern: as T increases by 1, [] decreases by 1.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO1_A = `SAY:
- 14+1, 13+2, 12+3, 11+4, 10+5, 9+6, 8+7 = 7 pairs. Stop at T=7 because [] must be > T.

DO:
- Click to reveal full list.

TEACHER NOTES:
Seven pairs when [] > T. The stopping point is key.

WATCH FOR:
- Students who see the pattern -- ready for complex equations.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO2_Q = `SAY:
- Now the complex one. Boards: 5 + 3 x T = []. Find THREE pairs where T is 1, 2, or 3.
- Remember: order of operations! 3 x T happens FIRST, then + 5.
- T = 1: 5 + 3 x 1 = 5 + 3 = 8. So [] = 8.
- T = 2: 5 + 3 x 2 = 5 + 6 = 11. So [] = 11.
- T = 3: 5 + 3 x 3 = 5 + 9 = 14. So [] = 14.
- Notice the pattern: [] goes up by 3 each time. Why? [Because we are adding another group of 3]

DO:
- Display equation. 60 seconds on boards. Reveal all three pairs.
- Highlight the pattern: [] increases by 3 as T increases by 1.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Write all three pairs. Show me!"
- Scan for: all three correct on 80%+.
PROCEED: If 80%+ find all three correctly (applying order of operations), move to hinge.
PIVOT: If students get wrong answers (e.g., adding before multiplying), reteach: "3 x T happens FIRST. THEN add 5. So for T = 2: 3 x 2 = 6, then 6 + 5 = 11."

TEACHER NOTES:
This We Do combines order of operations (from L5) with paired unknowns (today). The pattern ([] increases by 3 as T increases by 1) connects to the concept of a linear relationship.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Table with structure: "T = 1: 3 x 1 = ___. ___ + 5 = ___. [] = ___." For each value.
EXTENDING PROMPT:
- Task: "Predict: What is [] when T = 10? When T = 100? Can you write a RULE connecting [] and T?"

WATCH FOR:
- Students who see the +3 pattern and predict without calculating -- excellent mathematical reasoning.
- Students who apply incorrect order of operations -- redirect to L5 learning.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO2_A = `SAY:
- T=1 -> []=8. T=2 -> []=11. T=3 -> []=14. Pattern: [] goes up by 3!

DO:
- Click to reveal all three pairs with pattern highlighted.

TEACHER NOTES:
The +3 pattern foreshadows linear relationships in later algebra.

WATCH FOR:
- Students who can predict T=4 -> []=17 without calculating -- they see the pattern.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Gate check. 4 x T + [] = 19. If T = 3, what is []?
- Show me fingers: A) 7, B) 19, C) 31, D) 12.

DO:
- Display options. 15 seconds. "Show fingers." Scan. Reveal.

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- "T = 3. What is []? A=7, B=19, C=31, D=12. Fingers NOW."
- Scan for: A on 80%+.
PROCEED: If 80%+ choose A, ready for You Do.
PIVOT: If D (students just calculated 4 x 3 and stopped), reteach: "4 x 3 = 12. But the equation says 12 + [] = 19. You still need to find []: 19 - 12 = 7." Re-check: "2 x T + [] = 15. If T = 4, what is []?"

TEACHER NOTES:
Distractors: A (correct), B (wrote the target 19), C (added 12+19), D (just calculated 4x3 without finishing). The most diagnostic wrong answer is D.

WATCH FOR:
- Students choosing D -- they are halfway there but forgot the second step.
- Quick correct answers -- ready for independent work.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_HINGE_A = `SAY:
- A -- [] = 7. Because 4 x 3 = 12. 12 + [] = 19. [] = 7.

DO:
- Click to reveal.

TEACHER NOTES:
Two-step solution: multiply first (order of operations), then subtract to find [].

WATCH FOR:
- Students who got it -- ready for You Do.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_YOUDO = `SAY:
- Read from slide: "First: Identify all unknowns and constraints. Next: Try values systematically (start from the smallest). Then: Record all solutions and verify each one."
- Complete all problems on your worksheet.

DO:
- Distribute Session 6 Worksheet. Circulate. Focus on systematic listing and correct order of operations.

TEACHER NOTES:
The worksheet has 6 graduated problems: 1-2 are additive pairs ([] + T = target, with constraints), 3-4 combine multiplication and addition (substitution type), 5-6 are the complex multi-operation type.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 6 Enabling Scaffold with tables pre-structured for systematic substitution.
EXTENDING PROMPT:
- Task: Session 6 Extension investigation connecting pairs of unknowns to coordinate plots and rules.

WATCH FOR:
- Students who find some solutions but not all -- push for systematic approach.
- Students who can solve but cannot verify -- emphasise substitution check.

[Maths: Summarise -- Independent Practice (You Do) | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Final exit ticket of the unit. Two questions. 3 minutes.

DO:
- Display exit ticket. Students work silently. Collect after 3 minutes.

TEACHER NOTES:
Q1 assesses SC1+SC3 (find all solutions systematically). Q2 assesses SC2 (combined operations with paired unknowns). This exit ticket data gives a holistic picture of the unit -- compare with L1 exit ticket data to measure growth.

WATCH FOR:
- Students who could not find single unknowns in L1 but can now find paired unknowns -- significant progress.

[Maths: Summarise -- Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Success criteria check -- one last time.
- SC1: I can find multiple solutions when an equation has two unknowns. Thumbs.
- SC2: I can use number properties and order of operations to find unknowns. Thumbs.
- SC3: I can systematically list and check all possible solutions. Thumbs.
- Turn and talk: What was the BIGGEST thing you learned across all 6 sessions?
- You started this unit not knowing what an "unknown" was. Now you can solve equations with TWO unknowns, use brackets, apply properties, and work systematically. That is real algebra.
- Well done. Give yourselves a round of applause.

DO:
- Display closing slide. Run thumbs for each SC. Celebrate.

TEACHER NOTES:
This closing should feel celebratory -- the unit is complete. The summary statement connects L1 (simple unknowns) to L6 (paired unknowns in complex equations) to show the learning journey.

WATCH FOR:
- Students who are proud of their progress -- celebrate.

[Maths: Monitor Progress & Feedback | VTLM 2.0: Monitor Progress]`;

const NOTES_RESOURCES = `SAY:
- Printable resources for today.

DO:
- Print Session 6 Worksheet (class set), Session 6 Enabling Scaffold (enabling), Session 6 Extension (extending).

TEACHER NOTES:
Session 6 Worksheet has 6 problems. Session 6 Enabling Scaffold provides table scaffolds. Session 6 Extension extends to coordinate plotting and rules.

WATCH FOR:
- Ensure Session 6 Extension is available for early finishers.

[Maths: Planning -- Preparation | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // ── Slide 1: Title ──
  titleSlide(pres, UNIT_TITLE, "Session 6: Pairs of Unknowns in Complex Equations", "Multiple Solutions and Systematic Thinking | Session 6 of 6 | Year 5/6", NOTES_TITLE);

  // ── Slides 2-3: Daily Review (withReveal) — Mixed Review ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Mixed Review", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText([
        { text: "1.  [] x T = 24 where [] > T. Find TWO pairs.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 4 } },
        { text: "     Think: factor pairs of 24", options: { fontSize: 12, color: C.MUTED, breakLine: true, paraSpaceAfter: 14 } },
        { text: "2.  What is the LCM of 6 and 8?", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 4 } },
        { text: "     Lowest Common Multiple", options: { fontSize: 12, color: C.MUTED, breakLine: true, paraSpaceAfter: 14 } },
        { text: "3.  Estimate: 489 + 312 (round to hundreds)", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 4 } },
        { text: "     489 approx ?   312 approx ?", options: { fontSize: 12, color: C.MUTED } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 8.5, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) e.g. 8x3, 6x4, 12x2    2) LCM=24    3) ~800 (500+300)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // ── Slide 4: Fluency — Order of Operations Countdown ──
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Order of Operations Countdown", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

    const expressions = [
      { label: "E1", eq: "3 + 2 x 5" },
      { label: "E2", eq: "(4 + 1) x 6" },
      { label: "E3", eq: "12 / 3 + 5 x 2" },
      { label: "E4", eq: "(8 - 3) x (2 + 4)" },
    ];

    expressions.forEach((ex, i) => {
      const exY = CONTENT_TOP + i * 0.72;
      addCard(s, 0.5, exY, 9, 0.58, { strip: C.PRIMARY });
      s.addText(ex.label, {
        x: 0.75, y: exY + 0.05, w: 0.8, h: 0.48,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, valign: "middle",
      });
      s.addText(ex.eq, {
        x: 1.6, y: exY + 0.05, w: 7.6, h: 0.48,
        fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0, valign: "middle",
      });
    });

    s.addText("20 seconds per expression.  Show your answer on your board!", {
      x: 0.5, y: CONTENT_TOP + 3.1, w: 9, h: 0.4,
      fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FLUENCY);
  })();

  // ── Slide 5: LI/SC ──
  liSlide(pres,
    ["We are learning to find pairs of unknown values in equations with multiple operations"],
    [
      "I can find multiple solutions when an equation has two unknowns",
      "I can use number properties and order of operations to find unknowns in complex equations",
      "I can systematically list and check all possible solutions",
    ],
    NOTES_LI_SC, FOOTER);

  // ── Slide 6: I Do 1 — Multiple Solutions: [] + T = 12 ──
  workedExSlide(pres, 2, "I Do", "Multiple Solutions: [] + T = 12",
    [
      "How many ways can two whole",
      "numbers add to 12?",
      "",
      "Systematic approach ([] > T, both >= 1):",
      "",
      "  11+1, 10+2, 9+3, 8+4, 7+5",
      "",
      "That is 5 pairs where [] > T.",
      "(Plus 6+6 if [] = T is allowed.)",
      "",
      "KEY: [] + [] = 12 -> only ONE",
      "solution: [] = 6",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("[] + T = 12", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      // Table of pairs
      const pairs = ["11 + 1", "10 + 2", "9 + 3", "8 + 4", "7 + 5", "6 + 6"];
      pairs.forEach((p, i) => {
        const pY = lg.panelTopPadded + 0.45 + i * 0.35;
        slide.addText(p, {
          x: lg.rightX + 0.5, y: pY, w: lg.rightW - 1.0, h: 0.3,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, align: "center",
        });
      });
      // Label
      addTextOnShape(slide, "Many solutions!", {
        x: lg.rightX + 0.5, y: lg.panelTopPadded + 2.6, w: lg.rightW - 1.0, h: 0.35, rectRadius: 0.06,
        fill: { color: C.ACCENT },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // ── Slide 7: I Do 2 — Complex: 6 + 4 x 8 = 6 x T + [] ──
  workedExSlide(pres, 2, "I Do", "Complex: 6 + 4 x 8 = 6 x T + []",
    [
      "Left side first (order of operations):",
      "  4 x 8 = 32, then 6 + 32 = 38",
      "",
      "So: 6 x T + [] = 38",
      "",
      "Try T = 1: 6 + [] = 38 -> [] = 32",
      "Try T = 2: 12 + [] = 38 -> [] = 26",
      "Try T = 3: 18 + [] = 38 -> [] = 20",
      "Try T = 6: 36 + [] = 38 -> [] = 2",
      "",
      "Multiple valid pairs!",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
      slide.addText("6 + 4 x 8 = 6 x T + []", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      // Table header
      const cols = ["T", "6 x T", "[]"];
      const colW = (lg.rightW - 0.4) / 3;
      cols.forEach((col, i) => {
        slide.addShape("rect", {
          x: lg.rightX + 0.2 + i * colW, y: lg.panelTopPadded + 0.45, w: colW, h: 0.3,
          fill: { color: C.PRIMARY }, line: { color: C.WHITE, width: 1 },
        });
        slide.addText(col, {
          x: lg.rightX + 0.2 + i * colW, y: lg.panelTopPadded + 0.45, w: colW, h: 0.3,
          fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true, align: "center", valign: "middle", margin: 0,
        });
      });
      // Table rows
      const rows = [["1", "6", "32"], ["2", "12", "26"], ["3", "18", "20"], ["6", "36", "2"]];
      rows.forEach((row, ri) => {
        row.forEach((val, ci) => {
          const cellY = lg.panelTopPadded + 0.75 + ri * 0.38;
          slide.addShape("rect", {
            x: lg.rightX + 0.2 + ci * colW, y: cellY, w: colW, h: 0.35,
            fill: { color: C.WHITE }, line: { color: C.MUTED, width: 0.5 },
          });
          slide.addText(val, {
            x: lg.rightX + 0.2 + ci * colW, y: cellY, w: colW, h: 0.35,
            fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
          });
        });
      });
      // Label
      addTextOnShape(slide, "Multiple valid pairs!", {
        x: lg.rightX + 0.4, y: lg.panelTopPadded + 2.55, w: lg.rightW - 0.8, h: 0.35, rectRadius: 0.06,
        fill: { color: C.ACCENT },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // ── Slides 8-9: CFU (withReveal) ──
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check: List All Pairs", "Show Me Boards",
      "[] x T = 20\n(both whole numbers > 1)\n\nList ALL possible pairs on your board.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Pairs: 2x10, 4x5, 5x4, 10x2 (if order matters) or 2x10, 4x5 (if [] > T)", {
        x: 0.5, y: 4.0, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // ── Slides 10-11: We Do PP1 (withReveal) — [] + T = 15 ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "[] + T = 15 ([] > T, both >= 1)", { y: 0.65, fontSize: 20, color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9, 3.2, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "[] + T = 15  ([] > T)", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 10 } },
        { text: "Start from the smallest T and work up:", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 10 } },
        { text: "T = 1, [] = ___    T = 2, [] = ___    T = 3, [] = ___", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 6 } },
        { text: "T = 4, [] = ___    T = 5, [] = ___    T = 6, [] = ___", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 6 } },
        { text: "T = 7, [] = ___    STOP! Why do we stop here?", options: { fontSize: 14, color: C.ALERT, bold: true } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.2, w: 8.5, h: 2.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "14+1, 13+2, 12+3, 11+4, 10+5, 9+6, 8+7 = 7 pairs. Stop: [] must be > T.", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // ── Slides 12-13: We Do PP2 (withReveal) — 5 + 3 x T = [] ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "5 + 3 x T = []", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9, 3.2, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "5 + 3 x T = []", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 8 } },
        { text: "Find [] when T = 1, 2, and 3:", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 4 } },
        { text: "Remember: multiplication BEFORE addition!", options: { fontSize: 13, color: C.ALERT, breakLine: true, bold: true, paraSpaceAfter: 14 } },
        { text: "T = 1: 5 + 3 x 1 = 5 + ___ = ___", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 8 } },
        { text: "T = 2: 5 + 3 x 2 = 5 + ___ = ___", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 8 } },
        { text: "T = 3: 5 + 3 x 3 = 5 + ___ = ___", options: { fontSize: 15, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.2, w: 8.5, h: 2.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "T=1 -> []=8    T=2 -> []=11    T=3 -> []=14    (Pattern: [] goes up by 3!)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // ── Slides 14-15: Hinge CFU (withReveal) ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "4 x T + [] = 19.  If T = 3, what is []?", { color: C.ALERT, fontSize: 18 });

      s.addShape("roundRect", {
        x: 0.5, y: CONTENT_TOP, w: 2.4, h: 0.40, rectRadius: 0.08,
        fill: { color: C.ALERT },
      });
      s.addText("Finger Voting", {
        x: 0.5, y: CONTENT_TOP, w: 2.4, h: 0.40,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      const qY = CONTENT_TOP + 0.56;
      addCard(s, 0.5, qY, 9, 2.8, { strip: C.ALERT, fill: C.WHITE });
      const options = ["A)  7", "B)  19", "C)  31", "D)  12"];
      s.addText(options.map((opt, i) => ({
        text: opt,
        options: { fontSize: 18, color: C.CHARCOAL, breakLine: i < options.length - 1, paraSpaceAfter: 14 },
      })), {
        x: 0.9, y: qY + 0.2, w: 8.2, h: 1.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      s.addText("Hold up 1, 2, 3, or 4 fingers.", {
        x: 0.9, y: qY + 2.1, w: 8.2, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A)  [] = 7     (4 x 3 = 12, 12 + [] = 19, [] = 7)", {
        x: 0.5, y: 4.0, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // ── Slide 16: You Do ──
  workedExSlide(pres, 4, "You Do", "Pairs of Unknowns",
    [
      "First: Identify unknowns and constraints.",
      "Next: Try values systematically",
      "      (start from the smallest).",
      "Then: Record ALL solutions and verify.",
      "",
      "1.  [] + T = 18 ([] > T >= 1).",
      "    List all pairs.",
      "2.  [] x T = 36 (both > 1).",
      "    List all pairs.",
      "3.  2 x T + [] = 15.",
      "    Find [] when T = 1, 2, 3, 4.",
      "",
      "Complete all 6 problems on your",
      "worksheet.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.6, { strip: C.ALERT });
      slide.addText("Be Systematic!", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Start from smallest value", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Check every possibility", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Stop when values repeat", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 0.9,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // ── Slide 17: Exit Ticket ──
  exitTicketSlide(pres,
    [
      "[] + T = 20 where [] > T >= 1. List ALL possible pairs. How many are there?",
      "3 x T + [] = 22. If T = 4, what is []? Show your working.",
    ],
    NOTES_EXIT, FOOTER);

  // ── Slide 18: Closing ──
  closingSlide(pres,
    "What was the BIGGEST thing you learned across all 6 sessions?",
    [
      "I can find multiple solutions for paired unknowns",
      "I can use properties and order of operations to solve complex equations",
      "I can systematically list and verify all solutions",
    ],
    NOTES_CLOSING);

  // ── Slide 19: Resources ──
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // ── Write PPTX ──
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG6_Session6_Pairs_of_Unknowns.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Session 6 Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Pairs of Unknowns",
      color: C.NAVY,
      lessonInfo: "Session 6 of 6 | Algebra: Unknown Values | Year 5/6",
    });
    y = addTipBox(doc, "Be systematic! Start from the smallest value and work up. Check every possibility. Verify each solution by substituting back into the equation.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Additive Pairs", y, { color: C.NAVY });
    y = addProblem(doc, 1, "[] + T = 18 ([] > T, both >= 1). List ALL pairs.", y, { color: C.NAVY, writeLines: [{ label: "Pairs:" }, { label: "" }, { label: "How many pairs?" }] });
    y = addProblem(doc, 2, "[] x T = 36 (both > 1). List ALL pairs.", y, { color: C.NAVY, writeLines: [{ label: "Pairs:" }, { label: "" }, { label: "How many pairs?" }] });

    y = addSectionHeading(doc, "Section B: Substitution", y, { color: C.NAVY });
    y = addProblem(doc, 3, "2 x T + [] = 15. Find [] when T = 1, 2, 3, and 4.", y, { color: C.NAVY, writeLines: [{ label: "T=1: [] =" }, { label: "T=2: [] =" }, { label: "T=3: [] =" }, { label: "T=4: [] =" }] });
    y = addProblem(doc, 4, "3 x T + [] = 22. Find [] when T = 1, 3, 5, and 7.", y, { color: C.NAVY, writeLines: [{ label: "T=1: [] =" }, { label: "T=3: [] =" }, { label: "T=5: [] =" }, { label: "T=7: [] =" }] });

    y = addSectionHeading(doc, "Section C: Complex Equations", y, { color: C.NAVY });
    y = addProblem(doc, 5, "4 x T + [] = 30. Find all whole number pairs where [] >= 0.", y, { color: C.NAVY, writeLines: [{ label: "Systematic list:" }, { label: "" }, { label: "" }] });
    y = addProblem(doc, 6, "(T + 2) x [] = 24. Find at least 3 pairs where both > 0.", y, { color: C.NAVY, writeLines: [{ label: "Pairs:" }, { label: "" }, { label: "Verify one:" }] });

    addPdfFooter(doc, "Session 6 | Algebra: Finding Unknown Values | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Session 6 Enabling Scaffold
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Tables for Systematic Listing",
      color: C.TEAL,
      lessonInfo: "Session 6 of 6 | Algebra: Unknown Values | Year 5/6",
    });
    y = addTipBox(doc, "Use these tables to work through each problem systematically. Fill in one row at a time. The first row is done for you in each table.", y, { color: C.TEAL });

    // Problem 1 scaffold
    y = addSectionHeading(doc, "Problem 1: [] + T = 18 ([] > T, both >= 1)", y, { color: C.NAVY });

    const headers1 = ["T", "[] = 18 - T", "[] > T?"];
    const colW1 = 130;
    const tableX = 80;
    // Header row
    headers1.forEach((h, i) => {
      doc.rect(tableX + i * colW1, y, colW1, 22).fill("#" + C.PRIMARY).stroke("#" + C.PRIMARY);
      doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF").text(h, tableX + i * colW1 + 5, y + 5, { width: colW1 - 10, align: "center" });
    });
    y += 22;
    // First row filled
    const firstRow = ["1", "17", "Yes"];
    firstRow.forEach((val, i) => {
      doc.rect(tableX + i * colW1, y, colW1, 22).lineWidth(0.5).stroke("#999999");
      doc.fontSize(10).font("Sans").fillColor("#333333").text(val, tableX + i * colW1 + 5, y + 5, { width: colW1 - 10, align: "center" });
    });
    y += 22;
    // Empty rows
    for (let r = 0; r < 8; r++) {
      headers1.forEach((_, i) => {
        doc.rect(tableX + i * colW1, y, colW1, 22).lineWidth(0.5).stroke("#999999");
      });
      y += 22;
    }
    y += 10;

    // Problem 3 scaffold
    y = addSectionHeading(doc, "Problem 3: 2 x T + [] = 15", y, { color: C.NAVY });

    const headers3 = ["T", "2 x T", "[] = 15 - (2xT)"];
    headers3.forEach((h, i) => {
      doc.rect(tableX + i * colW1, y, colW1, 22).fill("#" + C.PRIMARY).stroke("#" + C.PRIMARY);
      doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF").text(h, tableX + i * colW1 + 5, y + 5, { width: colW1 - 10, align: "center" });
    });
    y += 22;
    // First row filled
    const firstRow3 = ["1", "2", "13"];
    firstRow3.forEach((val, i) => {
      doc.rect(tableX + i * colW1, y, colW1, 22).lineWidth(0.5).stroke("#999999");
      doc.fontSize(10).font("Sans").fillColor("#333333").text(val, tableX + i * colW1 + 5, y + 5, { width: colW1 - 10, align: "center" });
    });
    y += 22;
    for (let r = 0; r < 4; r++) {
      headers3.forEach((_, i) => {
        doc.rect(tableX + i * colW1, y, colW1, 22).lineWidth(0.5).stroke("#999999");
      });
      y += 22;
    }
    y += 10;

    y = addTipBox(doc, "Remember: 2 x T happens FIRST (order of operations), THEN add to get 15. So [] = 15 - (2 x T).", y, { color: C.TEAL });

    addPdfFooter(doc, "Session 6 | Enabling Scaffold | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Session 6 Extension — Unknown Pairs & Patterns
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Discovering Rules and Coordinate Connections",
      color: C.NAVY,
      lessonInfo: "Session 6 of 6 | Algebra: Unknown Values | Year 5/6",
    });

    y = addSectionHeading(doc, "From Pairs to Patterns", y, { color: C.NAVY });
    y = addBodyText(doc, "When you found pairs for 5 + 3 x T = [], you discovered that as T goes up by 1, [] goes up by 3. This is a PATTERN -- and patterns like this can be drawn as pictures called coordinate graphs.", y);
    y += 5;

    y = addSectionHeading(doc, "Worked Example: 5 + 3 x T = []", y, { color: C.NAVY });
    y = addBodyText(doc, "T = 1, [] = 8.   T = 2, [] = 11.   T = 3, [] = 14.   T = 4, [] = 17.", y);
    y = addBodyText(doc, "These pairs can be written as coordinates: (1, 8), (2, 11), (3, 14), (4, 17).", y);
    y = addBodyText(doc, "The RULE connecting T and [] is: [] = 5 + 3 x T. This is called a linear rule because the pattern makes a straight line when plotted.", y);
    y += 5;

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addTipBox(doc, "For each equation, find pairs, write them as coordinates, and try to write the rule.", y, { color: C.TEAL });

    y = addProblem(doc, 1, "2 + 4 x T = []. Find [] for T = 1, 2, 3, 4, 5.", y, { color: C.NAVY, writeLines: [{ label: "Pairs:" }, { label: "Coordinates:" }, { label: "Rule: [] = " }] });
    y = addProblem(doc, 2, "10 - 2 x T = []. Find [] for T = 1, 2, 3, 4.", y, { color: C.NAVY, writeLines: [{ label: "Pairs:" }, { label: "Coordinates:" }, { label: "Rule: [] = " }, { label: "What happens as T gets bigger?" }] });
    y = addProblem(doc, 3, "Make your own equation: ___ + ___ x T = []. Choose your numbers, find 5 pairs, and write the rule.", y, { color: C.NAVY, writeLines: [{ label: "Equation:" }, { label: "Pairs:" }, { label: "Rule: [] = " }, { label: "What is the pattern?" }] });

    y = addSectionHeading(doc, "The Big Connection", y, { color: C.NAVY });
    y = addBodyText(doc, "In high school, you will study these rules using letters instead of symbols:", y);
    y = addBodyText(doc, "[] = 5 + 3 x T becomes y = 3x + 5", y, { italic: true });
    y = addBodyText(doc, "You are already doing algebra -- just with different notation!", y);

    addPdfFooter(doc, "Session 6 | Extension Investigation | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("\nSession 6 build complete!");
}

build().catch(err => { console.error(err); process.exit(1); });
