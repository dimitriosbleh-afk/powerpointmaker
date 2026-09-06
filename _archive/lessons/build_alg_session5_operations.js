"use strict";

// Algebra Unit — Session 5: Order of Operations & Brackets
// Session 5 of 6, Grade 5/6 Numeracy, Variant 0
// DR: Estimation Strategies
// Fluency: True or False Sprint
// VC2M6A02: recognising that 6 + 4 x 8 is not the same as (6 + 4) x 8

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

const SESSION = 5;
const UNIT_TITLE = "Algebra: Finding Unknown Values";
const FOOTER = "Session 5 of 6 | Algebra: Unknown Values | Year 5/6 Maths";
const OUT_DIR = "output/ALG6_Session5_Order_of_Operations";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "8 problems: evaluation + unknowns with order of operations.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Step-by-step undo framework for each problem.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Exploring how brackets change answers.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we level up. Until now, every equation had ONE operation -- multiplication or division. Today we tackle equations with TWO operations: addition AND multiplication, mixed together.
- When you see 6 + 4 x 8, your brain might want to go left to right. But maths has rules about which operation comes first. Today we learn those rules and use brackets to change them.

DO:
- Display the title slide. Whiteboards ready.

TEACHER NOTES:
Lesson 5 marks the shift from VC2M5A02 (single-operation unknowns) to VC2M6A02 (multi-operation equations). The curriculum explicitly states students should recognise "that 6 + 4 x 8 is not the same as (6 + 4) x 8." This lesson establishes the order of operations convention and introduces brackets as an override tool.

WATCH FOR:
- Students who have encountered BODMAS/BIMDAS before -- some may have partial understanding. Build on it.

[Maths: Planning | VTLM 2.0: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review: Estimation strategies.
- Problem 1: A school orders 48 boxes of pencils with 12 pencils each. Estimate the total. [50 x 12 = 600, or 48 x 10 = 480]
- Problem 2: 397 x 6. Estimate by rounding one factor. [400 x 6 = 2,400]
- Problem 3: 2,847 / 7. Estimate. [2,800 / 7 = 400]

DO:
- Display the slide. Students work on whiteboards, 15 seconds per problem. "Show me!"
- Accept multiple valid estimates -- emphasise the strategy, not one "correct" estimate.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Estimate on your board. Show me!"
- Scan for: reasonable estimates using rounding on 80%+ of boards.
PROCEED: If 80%+ produce reasonable estimates, move to Fluency.
PIVOT: If students calculate exactly instead of estimating, reteach: "Estimating means rounding to make the calculation easy. 397 is close to 400. 400 x 6 is easy mental maths -- 2,400." Re-check: "Estimate 612 x 4."

TEACHER NOTES:
Daily Review targets: "Four Processes -- I can use appropriate estimation strategies." This connects to algebra -- students will need estimation to verify whether solutions to multi-operation equations are reasonable.

WATCH FOR:
- Students who round both factors (48 x 12 -> 50 x 10) -- valid but may lose accuracy. Discuss trade-offs.
- Students who cannot decide what to round to -- guide: "Round to the nearest easy number."

[Maths: Daily Review | VTLM 2.0: Retention & Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers.
- 1) About 600 (50x12) or 480 (48x10). 2) About 2,400 (400x6). 3) About 400 (2800/7).

DO:
- Click to reveal estimates. Students self-check.

TEACHER NOTES:
Multiple valid estimates are acceptable -- the strategy matters more than one exact answer.

WATCH FOR:
- Students who estimated well -- estimation skills are solid.

[Maths: Daily Review Answers | VTLM 2.0: Retention & Recall]`;

const NOTES_FLUENCY = `SAY:
- True or False sprint. I will show a statement. Decide if it is true or false. Show thumbs up for true, thumbs down for false.
- Statement 1: 3 x 4 = 4 x 3. [True -- commutative]
- Statement 2: 5 + 3 x 2 = 16. [False -- it is 11, not 16. But students may say true if they go left to right!]
- Hold that thought on Statement 2. We will come back to it.
- Statement 3: (2 + 3) x 4 = 2 x 4 + 3 x 4. [True -- distributive property]
- Statement 4: 10 - 3 + 2 = 10 - 5 = 5. [False -- it is 9. Left to right for + and -]

DO:
- Display statements one at a time. Quick pace -- 5 seconds per statement.
- Statement 2 is the hook -- do not resolve it yet. Say "Interesting -- we will explore this."

TEACHER NOTES:
This True/False sprint is designed to surface the order-of-operations misconception organically. Statement 2 (5 + 3 x 2) will likely split the class -- students who go left to right will get 16, while those who multiply first get 11. This creates productive cognitive conflict that motivates the new learning.

WATCH FOR:
- How the class splits on Statement 2 -- the ratio tells you how much reteaching you will need.
- Students who already know "multiplication first" -- they may have prior exposure.

[Maths: Fluency | VTLM 2.0: Retention & Recall]`;

const NOTES_LI_SC = `SAY:
- Read from slide: "We are learning to use the order of operations and brackets to evaluate and solve equations with more than one operation."
- SC1: I know which operations to do first when there are no brackets.
- SC2: I can use brackets to change which operation happens first.
- SC3: I can find unknown values in equations with mixed operations.

DO:
- Display the slide. Point to each SC.
- "SC3 is what we are building towards -- but we need SC1 and SC2 first."

TEACHER NOTES:
SC1 establishes the convention (multiplication/division before addition/subtraction). SC2 introduces brackets as an override. SC3 applies both to find unknowns -- linking back to the unit's central skill.

WATCH FOR:
- Students who think "order of operations" is a fixed algorithm (BODMAS) to memorise -- it is a convention, not a law of nature.

[Maths: Planning -- Curriculum Alignment | VTLM 2.0: Planning]`;

const NOTES_IDO1 = `SAY:
- Back to Statement 2 from our sprint: 5 + 3 x 2. Some of you said 16, some said 11. Who is right?
- Mathematics has a convention -- a rule everyone agrees on -- about the ORDER of operations.
- The rule: Multiplication and division come BEFORE addition and subtraction.
- 5 + 3 x 2. I see addition AND multiplication. Multiplication first: 3 x 2 = 6. THEN addition: 5 + 6 = 11.
- So 5 + 3 x 2 = 11, NOT 16.
- Let me try another: 6 + 4 x 8.
- Multiplication first: 4 x 8 = 32. Then addition: 6 + 32 = 38.
- NOT 10 x 8 = 80. The 6 + 4 does NOT happen first.
- Does 38 make sense? 4 x 8 is about 32, plus a bit more. Yes.

DO:
- Display the slide showing both interpretations visually.
- Circle the multiplication to emphasise "this happens first."
- Show the WRONG way crossed out.

TEACHER NOTES:
This is the curriculum example from VC2M6A02. The convention is that multiplication/division have higher precedence than addition/subtraction. The wrong interpretation (left to right) is shown and explicitly crossed out.

MISCONCEPTIONS:
- Misconception: "You always work left to right."
  Why: Reading left to right is natural and reinforced in literacy. Students transfer this to maths.
  Impact: Students will get wrong answers on every multi-operation expression.
  Quick correction: "In reading, left to right is correct. In maths, the OPERATION tells you what to do first, not the position. x and / jump the queue."

WATCH FOR:
- Students who got 16 during the sprint -- they now see why it is wrong.
- Students who ask "Why does multiplication go first?" -- "It is a convention -- an agreement. Like driving on the left."

[Maths: Launch -- Explicit Instruction (I Do) | VTLM 2.0: Explicit Explanation & Modelling]`;

const NOTES_IDO2 = `SAY:
- But what if I WANT to add first? What if the problem NEEDS me to add before I multiply?
- That is what BRACKETS are for. Brackets say: DO THIS FIRST.
- Watch: (6 + 4) x 8. The brackets around 6 + 4 say: add first.
- Inside brackets first: 6 + 4 = 10. Then multiply: 10 x 8 = 80.
- Compare: 6 + 4 x 8 = 38. But (6 + 4) x 8 = 80. Same numbers, different answer.
- The brackets CHANGED the order of operations.
- Let me try: 3 x (5 + 2). Brackets first: 5 + 2 = 7. Then: 3 x 7 = 21.
- Without brackets: 3 x 5 + 2 = 15 + 2 = 17. With brackets: 21. Different answers. Brackets matter.

DO:
- Display the slide with both expressions side by side.
- Highlight brackets in a contrasting colour.
- Show the two different answers prominently.

TEACHER NOTES:
This directly addresses the curriculum example: "6 + 4 x 8 is not the same as (6 + 4) x 8." The side-by-side comparison makes the impact of brackets concrete.

MISCONCEPTIONS:
- Misconception: "Brackets do not change anything -- they are just decoration."
  Why: Students may have seen brackets used merely for grouping without changing meaning.
  Impact: Students will ignore brackets and get wrong answers.
  Quick correction: "In maths, brackets are an INSTRUCTION. They say: do this part first. They are not optional."

WATCH FOR:
- The "aha" moment when students see that the same numbers give different answers.
- Students who connect this to the distributive property from Lesson 3 -- validate but keep focus on order of operations.

[Maths: Launch -- Explicit Instruction (I Do) | VTLM 2.0: Explicit Explanation & Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. What is 2 + 5 x 4? On your whiteboard. 15 seconds.
- Now what is (2 + 5) x 4? Same numbers, but brackets.

DO:
- Display first expression. 15 seconds. "Show me!" Scan. Reveal.
- Display second expression. 15 seconds. "Show me!" Scan. Reveal.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Solve 2 + 5 x 4 on your board. Show me!"
- Scan for: 22 on 80%+.
- Then: "(2 + 5) x 4. Show me!"
- Scan for: 28 on 80%+.
PROCEED: If 80%+ correct on both, students understand the convention. Move to We Do.
PIVOT: If students get 28 for both (ignoring order of operations on the first), reteach: "Without brackets: x first, then +. WITH brackets: inside first, then outside." Use the mnemonic: "x and / are VIPs -- they go first unless brackets override them." Re-check: "3 + 6 x 2. What is the answer? Now (3 + 6) x 2?"

TEACHER NOTES:
The paired questions test both conventions: without brackets (multiplication first) and with brackets (brackets first). Students who get both right have grasped the core concept.

WATCH FOR:
- Students who get the first right but hesitate on the second -- they may be second-guessing themselves.
- Students who get both right confidently -- they are ready for unknowns.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_CFU_A = `SAY:
- 1. 2 + 5 x 4 = 2 + 20 = 22. Multiplication first: 5 x 4 = 20. Then add.
- 2. (2 + 5) x 4 = 7 x 4 = 28. Brackets first: 2 + 5 = 7. Then multiply.

DO:
- Click to reveal both answers.

TEACHER NOTES:
Same numbers, different answers. This is the key lesson: the order you compute in changes the result.

WATCH FOR:
- Students who now understand the difference -- ready for unknowns.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_WEDO1_Q = `SAY:
- Together: Find the value of [] in 3 + [] x 2 = 13.
- Ask: What is the order of operations here? [Multiplication first, then addition]
- So [] x 2 happens first, then we add 3. That means [] x 2 = 13 - 3 = 10.
- Ask: If [] x 2 = 10, what is []? [[] = 5]
- Verify: 3 + 5 x 2 = 3 + 10 = 13. Yes.
- Key insight: I worked BACKWARDS from the answer. The last operation to happen was + 3, so I UNDO it first by subtracting 3.

DO:
- Display the equation. Cold Call through each step. Click to reveal.

CFU CHECKPOINT:
Technique: Cold Call
Script:
- "[Name], which operation happens first?" "[Name], what do we undo first?"
PROCEED: If students identify the operations correctly and can undo them.
PIVOT: If students try 3 + [] = something first (wrong order), reteach: "Multiplication happens FIRST. So I need to undo the LAST thing that happened -- the addition -- before I can find []."

TEACHER NOTES:
This introduces the "working backwards" (inverse operations) strategy for multi-operation equations. The key insight: to find an unknown, undo operations in REVERSE order.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Scaffold: "13 - 3 = ___. So [] x 2 = ___. [] = ___ / 2 = ___."
EXTENDING PROMPT:
- Task: "Find []: 7 + [] x 3 = 25. Then make up your own equation with the same structure."

WATCH FOR:
- Students who try to subtract 13 - 2 first -- they are undoing in the wrong order.
- Students who see the "undo" logic -- they are thinking algebraically.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO1_A = `SAY:
- Last op: + 3. Undo: 13 - 3 = 10. [] x 2 = 10. [] = 5. Verify: 3 + 5 x 2 = 13. Yes.

DO:
- Click to reveal solution.

TEACHER NOTES:
Confirm the "working backwards" strategy.

WATCH FOR:
- Students who see the undo pattern -- ready for brackets problems.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO2_Q = `SAY:
- Now with brackets: Find [] in ([] + 4) x 3 = 21.
- Boards: What happens first here? [Inside brackets: [] + 4]
- So ([] + 4) is calculated first, then multiplied by 3.
- Working backwards: the LAST operation was x 3. Undo it: 21 / 3 = 7.
- So [] + 4 = 7.
- Ask: What is []? [[] = 3]
- Verify: (3 + 4) x 3 = 7 x 3 = 21. Yes.

DO:
- Display equation. 30 seconds think time. Step through with Cold Call. Click to reveal.

CFU CHECKPOINT:
Technique: Show Me Boards + Cold Call
Script:
- "What do we undo first -- the + 4 or the x 3?" [x 3]
- "Write [] on your board. Show me!"
- Scan for: [] = 3 on 80%+.
PROCEED: If 80%+ correct, move to hinge.
PIVOT: If students get confused about which to undo first, reteach: "The LAST operation to happen in ([] + 4) x 3 is the x 3. We always undo the LAST operation first. 21 / 3 = 7. Now [] + 4 = 7. Undo + 4: 7 - 4 = 3."

TEACHER NOTES:
With brackets, the undoing order changes. The key heuristic: "Last operation done = first operation undone."

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: "21 / 3 = ___. So [] + 4 = ___. [] = ___ - 4 = ___."
EXTENDING PROMPT:
- Task: "Find []: ([] + 5) x 4 = 36. Then: [] x (3 + 2) = 45. How is the second one different?"

WATCH FOR:
- Students who can follow the undo steps but do not understand WHY -- push for explanation.
- Students confusing bracket equations with non-bracket equations.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO2_A = `SAY:
- Last op: x 3. Undo: 21 / 3 = 7. [] + 4 = 7. [] = 3. Verify: (3 + 4) x 3 = 21. Yes.

DO:
- Click to reveal solution.

TEACHER NOTES:
Brackets change which operation is "last" and therefore which to undo first.

WATCH FOR:
- Students who got it -- ready for the hinge.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Gate check. What is [] in: 4 + [] x 5 = 24?
- Show me fingers: A) 4, B) 28, C) 100, D) 1.

DO:
- Display options. 15 seconds. "Show fingers." Scan. Reveal.

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- "What is []? A=4, B=28, C=100, D=1. Fingers NOW."
- Scan for: A on 80%+.
PROCEED: If 80%+ choose A, ready for You Do.
PIVOT: If B (students went left to right), reteach: "Remember: multiplication happens BEFORE addition. So x 5 happens first, then + 4. Undo: 24 - 4 = 20. 20 / 5 = 4." Re-check: "2 + [] x 3 = 14. What is []?"

TEACHER NOTES:
Distractors: A (correct), B (24+4=28), C (if they multiply), D (guess). The most revealing wrong answer is B, indicating left-to-right thinking.

WATCH FOR:
- Students choosing B -- they are still thinking left to right. Need targeted reteaching.
- Quick correct answers -- ready for independent practice.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_HINGE_A = `SAY:
- A -- [] = 4. Multiplication first: [] x 5. Then + 4. Working backwards: 24 - 4 = 20. 20 / 5 = 4.

DO:
- Click to reveal.

TEACHER NOTES:
The working-backwards strategy with correct order of operations.

WATCH FOR:
- Students who now understand.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_YOUDO = `SAY:
- Read from slide: "First: Identify the operations and their order. Next: Work backwards -- undo the last operation first. Then: Verify by substituting back."
- Complete all problems on your worksheet.

DO:
- Distribute Session 5 Worksheet. Circulate. Focus on whether students correctly identify the order of operations.

TEACHER NOTES:
The worksheet has 8 graduated problems: 1-2 are evaluation (no unknowns -- just order of operations practice), 3-5 are unknowns without brackets, 6-8 are unknowns with brackets.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 5 Enabling Scaffold with "undo steps" pre-structured.
EXTENDING PROMPT:
- Task: Session 5 Extension investigation on creating equations that have different answers with and without brackets.

WATCH FOR:
- Students who can evaluate expressions (Q1-2) but struggle to find unknowns (Q3+) -- the backwards-working step is new.
- Students who confuse bracket and non-bracket rules.

[Maths: Summarise -- Independent Practice (You Do) | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Two questions. 3 minutes.

DO:
- Display exit ticket. Students work silently. Collect after 3 minutes.

TEACHER NOTES:
Q1 assesses SC1+SC3 (evaluate using correct order, find an unknown). Q2 assesses SC2 (impact of brackets). Data informs Lesson 6.

WATCH FOR:
- Students who confuse the two conventions.

[Maths: Summarise -- Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Success criteria check.
- SC1: I know which operations to do first when there are no brackets. Thumbs.
- SC2: I can use brackets to change which operation happens first. Thumbs.
- SC3: I can find unknown values in equations with mixed operations. Thumbs.
- Turn and talk: In your own words, what do brackets DO in a maths expression?
- Tomorrow -- our final lesson. Two unknowns in one equation. The ultimate challenge.

DO:
- Display closing slide. Run thumbs for each SC.

TEACHER NOTES:
The forward look to Lesson 6 signals the culmination of the unit. Students who are solid on order of operations and brackets will find the transition to paired unknowns manageable.

WATCH FOR:
- Thumbs-down on SC1 -- these students need extra consolidation.

[Maths: Monitor Progress & Feedback | VTLM 2.0: Monitor Progress]`;

const NOTES_RESOURCES = `SAY:
- Printable resources for today.

DO:
- Print Session 5 Worksheet (class set), Session 5 Enabling Scaffold (enabling), Session 5 Extension (extending).

TEACHER NOTES:
Session 5 Worksheet has 8 problems. Session 5 Enabling Scaffold provides undo-step scaffolding. Session 5 Extension explores bracket effects.

WATCH FOR:
- Ensure Session 5 Extension is available for early finishers.

[Maths: Planning -- Preparation | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // ── Slide 1: Title ──
  titleSlide(pres, UNIT_TITLE, "Session 5: Order of Operations & Brackets", "Why 6 + 4 x 8 is not (6 + 4) x 8 | Session 5 of 6 | Year 5/6", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // ── Slides 2-3: Daily Review (withReveal) — Estimation ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Estimation Strategies", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText([
        { text: "1.  48 boxes x 12 pencils each. Estimate the total.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 4 } },
        { text: "     Round one factor to make it easy", options: { fontSize: 12, color: C.MUTED, breakLine: true, paraSpaceAfter: 14 } },
        { text: "2.  397 x 6. Estimate by rounding.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 4 } },
        { text: "     397 approx ?", options: { fontSize: 12, color: C.MUTED, breakLine: true, paraSpaceAfter: 14 } },
        { text: "3.  2,847 / 7. Estimate.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 4 } },
        { text: "     What nearby number divides easily by 7?", options: { fontSize: 12, color: C.MUTED } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 8.5, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) ~600 (50x12)    2) ~2,400 (400x6)    3) ~400 (2800/7)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // ── Slide 4: Fluency — True or False Sprint ──
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "True or False Sprint", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

    const statements = [
      { label: "Statement 1", eq: "3 x 4 = 4 x 3" },
      { label: "Statement 2", eq: "5 + 3 x 2 = 16" },
      { label: "Statement 3", eq: "(2 + 3) x 4 = 2 x 4 + 3 x 4" },
      { label: "Statement 4", eq: "10 - 3 + 2 = 5" },
    ];

    statements.forEach((st, i) => {
      const stY = CONTENT_TOP + i * 0.75;
      addCard(s, 0.5, stY, 9, 0.62, { strip: i === 1 ? C.ALERT : C.PRIMARY });
      s.addText(st.label, {
        x: 0.75, y: stY + 0.05, w: 1.8, h: 0.5,
        fontSize: 12, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });
      s.addText(st.eq, {
        x: 2.5, y: stY + 0.05, w: 6.8, h: 0.5,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0, valign: "middle",
      });
    });

    s.addText("Thumbs up = TRUE.   Thumbs down = FALSE.", {
      x: 0.5, y: CONTENT_TOP + 3.2, w: 9, h: 0.4,
      fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FLUENCY);
  })();

  // ── Slide 5: LI/SC ──
  liSlide(pres,
    ["We are learning to use the order of operations and brackets to evaluate and solve equations with more than one operation"],
    [
      "I know which operations to do first when there are no brackets",
      "I can use brackets to change which operation happens first",
      "I can find unknown values in equations with mixed operations",
    ],
    NOTES_LI_SC, FOOTER);

  // ── Slide 6: I Do 1 — Order of Operations ──
  workedExSlide(pres, 2, "I Do", "Order of Operations: 6 + 4 x 8",
    [
      "Expression: 6 + 4 x 8",
      "",
      "Step 1 -- Identify operations:",
      "  Addition (+) and multiplication (x)",
      "",
      "Step 2 -- Multiplication first:",
      "  4 x 8 = 32",
      "",
      "Step 3 -- Then addition:",
      "  6 + 32 = 38",
      "",
      "Answer: 6 + 4 x 8 = 38",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.PRIMARY });
      slide.addText("Compare:", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      // Correct
      addTextOnShape(slide, "6 + 4 x 8 = 38", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.5, w: lg.rightW - 0.4, h: 0.55, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("x before +", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.1, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
      });
      // Wrong
      addTextOnShape(slide, "10 x 8 = 80  WRONG", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.5, w: lg.rightW - 0.4, h: 0.55, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("Left to right = WRONG", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.1, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
      });
    }
  );

  // ── Slide 7: I Do 2 — Brackets Change the Order ──
  workedExSlide(pres, 2, "I Do", "Brackets Change the Order: (6 + 4) x 8",
    [
      "Expression: (6 + 4) x 8",
      "",
      "Step 1 -- Brackets first:",
      "  6 + 4 = 10",
      "",
      "Step 2 -- Then multiply:",
      "  10 x 8 = 80",
      "",
      "Compare:",
      "  6 + 4 x 8 = 38",
      "  (6 + 4) x 8 = 80",
      "",
      "Same numbers, different answer!",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.SECONDARY });
      slide.addText("Side by Side", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      // Without brackets
      slide.addText("Without brackets", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.3, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, margin: 0,
      });
      addTextOnShape(slide, "6 + 4 x 8 = 38", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.75, w: lg.rightW - 0.4, h: 0.45, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      // With brackets
      slide.addText("With brackets", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.35, w: lg.rightW - 0.3, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, margin: 0,
      });
      addTextOnShape(slide, "(6 + 4) x 8 = 80", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.65, w: lg.rightW - 0.4, h: 0.45, rectRadius: 0.06,
        fill: { color: C.ACCENT },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      // Key insight
      addTextOnShape(slide, "Brackets = DO THIS FIRST", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.25, w: lg.rightW - 0.6, h: 0.35, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // ── Slides 8-9: CFU (withReveal) ──
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check: Order of Operations", "Show Me Boards",
      "Solve both on your board:\n\n1.  2 + 5 x 4 = ?\n\n2.  (2 + 5) x 4 = ?",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1.  2 + 5 x 4 = 2 + 20 = 22         2.  (2 + 5) x 4 = 7 x 4 = 28", {
        x: 0.5, y: 4.0, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // ── Slides 10-11: We Do PP1 (withReveal) — 3 + [] x 2 = 13 ──
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "3 + [] x 2 = 13",
      [
        "3 + [] x 2 = 13",
        "",
        "Step 1: Which operation happens LAST?",
        "",
        "Step 2: Undo the last operation:",
        "  13 - ___ = ___",
        "",
        "Step 3: So [] x 2 = ___.",
        "  Find []: ___ / 2 = ___",
        "",
        "Verify: 3 + ___ x 2 = ?",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.SECONDARY });
        slide.addText("Working Backwards:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "Last operation done = first to undo", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "x before + means + is LAST", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Undo + first, then undo x", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.0,
          fontFace: FONT_B, margin: 0, valign: "top",
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "Last op: + 3. Undo: 13-3=10. [] x 2=10. []=5. Verify: 3+5x2=13", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // ── Slides 12-13: We Do PP2 (withReveal) — ([] + 4) x 3 = 21 ──
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "([] + 4) x 3 = 21",
      [
        "([] + 4) x 3 = 21",
        "",
        "Step 1: Which operation happens LAST?",
        "  (Hint: brackets change the order!)",
        "",
        "Step 2: Undo the last operation:",
        "  21 / ___ = ___",
        "",
        "Step 3: So [] + 4 = ___.",
        "  Find []: ___ - 4 = ___",
        "",
        "Verify: (___ + 4) x 3 = ?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.SECONDARY });
        slide.addText("Brackets Change Undo Order:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
          fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "Brackets are done FIRST", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "So x 3 is done LAST", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Undo x 3 first, then undo + 4", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.0,
          fontFace: FONT_B, margin: 0, valign: "top",
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "Last op: x3. Undo: 21/3=7. []+4=7. []=3. Verify: (3+4)x3=21", {
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
      addTitle(s, "Find []:  4 + [] x 5 = 24", { color: C.ALERT });

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
      const options = ["A)  4", "B)  28", "C)  100", "D)  1"];
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
      addTextOnShape(slide, "A)  [] = 4     (24 - 4 = 20, 20 / 5 = 4)", {
        x: 0.5, y: 4.0, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // ── Slide 16: You Do ──
  workedExSlide(pres, 4, "You Do", "Order of Operations & Unknowns",
    [
      "First: Identify operations and their order.",
      "Next: Work backwards -- undo the last",
      "      operation first.",
      "Then: Verify by substituting back.",
      "",
      "1.  Evaluate: 8 + 3 x 6 = ?",
      "2.  Evaluate: (8 + 3) x 6 = ?",
      "3.  Find []: 5 + [] x 3 = 20",
      "4.  Find []: ([] + 6) x 2 = 18",
      "",
      "Complete all 8 problems on your worksheet.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.ALERT });
      slide.addText("Key Rule:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "No brackets: x and / FIRST", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Brackets: inside FIRST", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "To find []: undo LAST op first", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // ── Slide 17: Exit Ticket ──
  exitTicketSlide(pres,
    [
      "Find []: 7 + [] x 4 = 27. Show your working.",
      "Explain: Why does 6 + 4 x 8 give a different answer to (6 + 4) x 8?",
    ],
    NOTES_EXIT, FOOTER);

  // ── Slide 18: Closing ──
  closingSlide(pres,
    "In your own words, what do brackets DO in a maths expression?",
    [
      "I know which operations to do first (x and / before + and -)",
      "I can use brackets to change which operation happens first",
      "I can find unknown values in mixed-operation equations",
    ],
    NOTES_CLOSING);

  // ── Slide 19: Resources ──


  // ── Write PPTX ──
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG6_Session5_Order_of_Operations.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Session 5 Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Order of Operations & Brackets",
      color: C.NAVY,
      lessonInfo: "Session 5 of 6 | Algebra: Unknown Values | Year 5/6",
    });
    y = addTipBox(doc, "Remember: x and / happen BEFORE + and -. Brackets override the order -- do brackets FIRST. To find unknowns, work BACKWARDS: undo the last operation first.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Evaluate (no unknowns)", y, { color: C.NAVY });
    y = addProblem(doc, 1, "8 + 3 x 6 = ___", y, { color: C.NAVY, writeLines: [{ label: "Working:" }] });
    y = addProblem(doc, 2, "(8 + 3) x 6 = ___", y, { color: C.NAVY, writeLines: [{ label: "Working:" }] });

    y = addSectionHeading(doc, "Section B: Find [] (no brackets)", y, { color: C.NAVY });
    y = addProblem(doc, 3, "5 + [] x 3 = 20", y, { color: C.NAVY, writeLines: [{ label: "Undo steps:" }, { label: "[] =" }] });
    y = addProblem(doc, 4, "2 + [] x 4 = 22", y, { color: C.NAVY, writeLines: [{ label: "Undo steps:" }, { label: "[] =" }] });
    y = addProblem(doc, 5, "[] x 6 + 7 = 43", y, { color: C.NAVY, writeLines: [{ label: "Undo steps:" }, { label: "[] =" }] });

    y = addSectionHeading(doc, "Section C: Find [] (with brackets)", y, { color: C.NAVY });
    y = addProblem(doc, 6, "([] + 6) x 2 = 18", y, { color: C.NAVY, writeLines: [{ label: "Undo steps:" }, { label: "[] =" }] });
    y = addProblem(doc, 7, "([] + 3) x 5 = 40", y, { color: C.NAVY, writeLines: [{ label: "Undo steps:" }, { label: "[] =" }] });
    y = addProblem(doc, 8, "4 x ([] - 2) = 28", y, { color: C.NAVY, writeLines: [{ label: "Undo steps:" }, { label: "[] =" }] });

    addPdfFooter(doc, "Session 5 | Algebra: Finding Unknown Values | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Session 5 Enabling Scaffold
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Step-by-Step Undo Framework",
      color: C.TEAL,
      lessonInfo: "Session 5 of 6 | Algebra: Unknown Values | Year 5/6",
    });
    y = addTipBox(doc, "For each problem: identify which operation happens LAST, undo it first, then undo the next. The first one is done for you.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Example (done for you)", y, { color: C.NAVY });
    y = addBodyText(doc, "3 + [] x 2 = 13", y, { fontSize: 12 });
    y = addBodyText(doc, "Operations: + and x. Order: x first, then +. So + is LAST.", y);
    y = addBodyText(doc, "Undo + first: 13 - 3 = 10", y);
    y = addBodyText(doc, "So [] x 2 = 10", y);
    y = addBodyText(doc, "Undo x: 10 / 2 = 5", y);
    y = addBodyText(doc, "[] = 5. Check: 3 + 5 x 2 = 3 + 10 = 13. Yes!", y);
    y += 8;

    y = addSectionHeading(doc, "Now you try:", y, { color: C.NAVY });
    const problems = [
      { eq: "5 + [] x 3 = 20", hint: "Operations: + and x. Last operation: +" },
      { eq: "2 + [] x 4 = 22", hint: "Operations: + and x. Last operation: +" },
      { eq: "([] + 6) x 2 = 18", hint: "Brackets change it! Last operation: x 2" },
      { eq: "([] + 3) x 5 = 40", hint: "Brackets! Last operation: x 5" },
    ];
    problems.forEach((p, i) => {
      y = addBodyText(doc, (i + 1) + ". " + p.eq, y, { fontSize: 12 });
      y = addBodyText(doc, "   " + p.hint, y, { fontSize: 10, italic: true, color: "6B7280" });
      y = addWriteLine(doc, "Undo last operation: ", y);
      y = addWriteLine(doc, "So: ", y);
      y = addWriteLine(doc, "Undo next: ", y);
      y = addWriteLine(doc, "[] = ___. Check: ", y);
      y += 8;
      if (y > 700) { doc.addPage(); y = 50; }
    });

    addPdfFooter(doc, "Session 5 | Enabling Scaffold | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Session 5 Extension — Brackets Investigation
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "How Do Brackets Change Answers?",
      color: C.NAVY,
      lessonInfo: "Session 5 of 6 | Algebra: Unknown Values | Year 5/6",
    });

    y = addSectionHeading(doc, "The Big Idea", y, { color: C.NAVY });
    y = addBodyText(doc, "Adding brackets to an expression can change the answer completely. Your job: investigate WHEN brackets change the answer and WHEN they do not.", y);
    y += 5;

    y = addSectionHeading(doc, "Worked Example", y, { color: C.NAVY });
    y = addBodyText(doc, "Expression: 3 + 4 x 2", y);
    y = addBodyText(doc, "Without brackets: 3 + 4 x 2 = 3 + 8 = 11", y);
    y = addBodyText(doc, "With brackets around 3+4: (3 + 4) x 2 = 7 x 2 = 14", y);
    y = addBodyText(doc, "With brackets around 4x2: 3 + (4 x 2) = 3 + 8 = 11", y);
    y = addBodyText(doc, "Notice: brackets around 4 x 2 did NOT change the answer! Why? Because multiplication already happens first.", y);
    y += 5;

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addTipBox(doc, "For each expression: evaluate without brackets, then add brackets in DIFFERENT positions. Record when the answer changes and when it stays the same.", y, { color: C.TEAL });

    y = addProblem(doc, 1, "5 + 2 x 6. Try brackets in two positions.", y, { color: C.NAVY, writeLines: [{ label: "No brackets:" }, { label: "Brackets position 1:" }, { label: "Brackets position 2:" }, { label: "Did the answer change?" }] });
    y = addProblem(doc, 2, "8 - 3 + 2 x 4. Try at least three bracket positions.", y, { color: C.NAVY, writeLines: [{ label: "No brackets:" }, { label: "Position 1:" }, { label: "Position 2:" }, { label: "Position 3:" }] });
    y = addProblem(doc, 3, "Create your own expression with 3 operations. Show that different bracket positions give different answers.", y, { color: C.NAVY, writeLines: [{ label: "Expression:" }, { label: "Version 1:" }, { label: "Version 2:" }, { label: "Version 3:" }] });

    y = addSectionHeading(doc, "The Pattern", y, { color: C.NAVY });
    y = addBodyText(doc, "When do brackets change the answer? When do they NOT change it? Write your rule:", y);
    y = addWriteLine(doc, "My rule: ", y);
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 5 | Extension Investigation | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("\nSession 5 build complete!");
}

build().catch(err => { console.error(err); process.exit(1); });
