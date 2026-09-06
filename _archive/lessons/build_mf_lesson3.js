"use strict";

// Multiplication Facts Unit — Lesson 3: x6, x8 (Double Known Facts)
// Session 3, Grade 3/4 Numeracy, Variant 0
// x6 = double x3, x8 = double x4
// DR: Addition Algorithm (Vertical, 4- & 5-digits)
// Fluency: Subtraction Algorithm (Vertical, 3- & 4-digits)

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade34", 0);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar,
  withReveal, STAGE_COLORS,
} = T;

const SESSION = 3;
const FOOTER = "Multiplication Facts | Session 3 of 4 | Grade 3/4 Numeracy";
const OUT_DIR = "output/MF_Session3_Doubling_x6_x8";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - x6 and x8 multiplication facts.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-filled doubling bridge for x6 and x8 from x3 and x4.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into factor pairs.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Session 1 we nailed x2, x5, x10. Session 2 we built x3 and x4 from x2.
- Today we go further. We are building x6 and x8 from the facts we already know
- The pattern: x6 = double your x3 fact. x8 = double your x4 fact.
- You are building a chain of knowledge. Each new table connects to the last one.

DO:
- Display title slide
- Have whiteboards ready

TEACHER NOTES:
Session 3 of 4. x6 = 2 x (x3) and x8 = 2 x (x4). These derived strategies chain from Session 2. Students who secured x3 and x4 will find this accessible. Students who struggled with Session 2 may need the enabling scaffold.

WATCH FOR:
- Students who did not consolidate x3 or x4 -- they will struggle with the doubling bridge. Check exit ticket data from Session 2.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily review -- addition algorithm
- Four problems on your whiteboards. Vertical algorithm, rename where needed.

DO:
- Display problems. 90 seconds each.
- Circulate and check

TEACHER NOTES:
DR continues addition algorithm with 4- and 5-digit numbers. Fresh problems for spaced retrieval.

WATCH FOR:
- By Session 3, most students should be accurate with 4-digit addition. 5-digit with multiple renamings is the stretch.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers. Tick correct, fix errors.

DO:
- Reveal answers. Students self-check.

TEACHER NOTES:
Tick-and-fix cycle.

WATCH FOR:
- Improvement trend across Sessions 1-3

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency -- subtraction algorithm, 2 minutes, whiteboards

DO:
- Display problems. Time 2 minutes.

TEACHER NOTES:
Subtraction fluency continues. Problems include renaming across zeros and multi-step renaming.

WATCH FOR:
- Students who have improved since Session 1 fluency -- note progress

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention together
- Read from slide: "We are learning to use doubling to work out x6 and x8 multiplication facts from x3 and x4"
- Success criteria -- which one connects to what we learned yesterday? [SC1 -- using x3 and x4 as building blocks]

DO:
- Choral read LI and SCs

TEACHER NOTES:
This lesson explicitly chains from Session 2. SC1 is the floor (recalling x3 and x4), SC2 is the core (doubling to get x6 and x8), SC3 extends to choosing strategies flexibly.

WATCH FOR:
- Students who look unsure when x3 and x4 are mentioned -- they may need the enabling scaffold from Session 2 as a refresher

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_X6 = `SAY:
- First: x6. Six groups of something.
- Here is the big idea: 6 = 2 x 3. So x6 = double your x3 fact.
- Watch: 6 x 7
- Step 1: What is 3 x 7? I know this from yesterday: 21.
- Step 2: Double 21. That is 42.
- So 6 x 7 = 42.
- Let me check: 7, 14, 21, 28, 35, 42. The 6th number. Yes!

DO:
- Write each step clearly
- Point to the visual showing the connection: x3 -> double -> x6
- Emphasise: "You already know x3. Just double it."

TEACHER NOTES:
x6 = 2 x (x3) is the cleanest derived strategy. The array visual shows 6 rows as two groups of 3 rows. Students who automated x3 in Session 2 will find this fast.

MISCONCEPTIONS:
- Misconception: Students double the wrong number (doubling 7 instead of 21 for 6 x 7)
  Why: Confusion about what to double -- the group size vs the x3 product
  Impact: Produces systematically wrong answers for all x6 facts
  Quick correction: "You are doubling the ANSWER to x3, not the number being multiplied. 3 x 7 = 21. Double 21, not double 7."

WATCH FOR:
- Students who cannot recall x3 facts quickly enough to use them as a bridge -- they may need a quick x3 review before continuing

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_X8 = `SAY:
- Now x8. Same idea.
- 8 = 2 x 4. So x8 = double your x4 fact.
- Watch: 8 x 6
- Step 1: What is 4 x 6? Double-double: double 6 = 12, double 12 = 24.
- Step 2: Double 24. That is 48.
- So 8 x 6 = 48.
- Another way to think about it: x8 is triple-double. Double, double, double.

DO:
- Write the steps
- Show the chain visually: 6 -> 12 -> 24 -> 48 (three doublings)
- This is the most complex derived strategy so far -- go slowly

TEACHER NOTES:
x8 = double x4 = double (double-double) = triple-double. Some students will prefer to think of x8 as three doublings rather than finding x4 first then doubling. Both are valid. The triple-double is actually fewer mental steps for students who are fast doublers.

WATCH FOR:
- Students who lose track of the chain (doubling three times) -- suggest they write each step down rather than doing it all mentally
- Readiness signal: students following along and predicting the next step

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Cold call time.
- I am going to ask a question. When I say your name, give me the answer AND the strategy.
- What is 6 x 5?

DO:
- Cold call a student
- Expected: "30. I know 3 x 5 = 15, then double 15 = 30."
- If the student says 30 but uses skip counting, acknowledge and ask: "Can you also show me the doubling-from-x3 way?"

CFU CHECKPOINT:
Technique: Cold Call
Script:
- Say: "6 x 5. [Name], what is 6 x 5 and how did you work it out?"
- Scan for: 30, with strategy "3 x 5 = 15, double 15 = 30."
PROCEED: If the cold-called student (and 2-3 follow-ups) demonstrate the doubling strategy, proceed.
PIVOT: Most likely misconception -- student gives 30 via skip counting but cannot explain the doubling link. Reteach: "6 is double 3. So if I know 3 groups of 5 = 15, then 6 groups of 5 = double 15 = 30. The x3 fact is the stepping stone."

TEACHER NOTES:
Cold call varies the CFU technique from Sessions 1-2. Using 6 x 5 is deliberate -- students already know x5 from Session 1, so the anchor fact (3 x 5 = 15) should be accessible.

WATCH FOR:
- Students who freeze when cold called -- reassure: "Take your time. Start with: what is 3 x 5?"

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Partner work. Two problems on your whiteboards.
- Problem 1: 6 x 9. Use the double-x3 strategy. Show both steps.
- Problem 2: 8 x 5. Use the double-x4 strategy. Show both steps.

DO:
- Partners work together. 90 seconds.
- Circulate: check for correct intermediate steps, not just final answers

TEACHER NOTES:
We Do pairs one x6 and one x8. Different numbers from the I Do. Requiring written steps makes the doubling chain visible.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the enabling scaffold which pre-fills the x3 or x4 intermediate step. Students only need to double the result. For 6 x 9: "3 x 9 = 27" is given; students write "Double 27 = ___".
- Extra Notes: Distribute Session 3 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Find a x6 fact and a x8 fact that give the same answer. Are there more than one pair? Show your working.
- Extra Notes: Students discover common products (e.g., 6 x 8 = 48 = 8 x 6) and explore commutativity in a new way.

WATCH FOR:
- Students who double the wrong intermediate value -- the most common error in this strategy chain
- Partners where one does all the work -- redirect: "Both of you write the steps"

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Check your working.
- 6 x 9: Step 1: 3 x 9 = 27. Step 2: Double 27 = 54. Answer: 54.
- 8 x 5: Step 1: 4 x 5 = 20. Step 2: Double 20 = 40. Answer: 40.
- Quick tip for doubling 27: double 20 = 40, double 7 = 14, total = 54.

DO:
- Reveal answers. Cold call for strategy steps.
- Model the partitioning strategy for doubling larger numbers (like 27)

TEACHER NOTES:
Doubling two-digit numbers is a sub-skill that may need explicit attention. Partitioning (double the tens, double the ones, add) is the most reliable strategy.

WATCH FOR:
- Students who got the right answer but struggled with the doubling step -- model partitioning

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Four corners.
- I am going to show you 8 x 7.
- Corner 1: answer is 54. Corner 2: answer is 56. Corner 3: answer is 58. Corner 4: answer is 64.
- Think first, then move to your corner.

DO:
- Display the four options. Students move to corners.
- Correct answer: 56 (Corner 2). 4 x 7 = 28, double 28 = 56.
- Distractor analysis: 54 = common error from 6 x 9 carryover. 58 = adding 2 instead of doubling. 64 = 8 x 8.

CFU CHECKPOINT:
Technique: Four Corners
Script:
- Say: "8 x 7. Corner 1: 54. Corner 2: 56. Corner 3: 58. Corner 4: 64. Think, then move."
- Scan for: most students in Corner 2 (56).
PROCEED: If 80%+ go to Corner 2, move to You Do.
PIVOT: Most likely misconception -- students choose 54 (confusing 8 x 7 with 6 x 9) or 64 (confusing 8 x 7 with 8 x 8). Reteach: "8 x 7: start with 4 x 7. Double-double: double 7 = 14, double 14 = 28. That is 4 x 7. Now double 28 to get 8 x 7. Double 20 = 40, double 8 = 16, total = 56."

TEACHER NOTES:
Four Corners gets students moving and makes the response visible. The distractors are designed to catch specific errors: adjacent-fact confusion and operation errors.

WATCH FOR:
- Students who hesitate between corners -- they may be guessing. Ask: "What is 4 x 7 first?"

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice. Session 3 Worksheet.
- First: identify whether the problem is x6 or x8.
- Next: find the x3 or x4 fact first.
- Then: double it to get your answer.
- 8-10 minutes. Go.

DO:
- Distribute worksheets
- Enabling students get the Session 3 Enabling Scaffold
- Circulate, prioritise students who struggled at the hinge check

TEACHER NOTES:
You Do mixes x6 and x8 problems with different numbers from the We Do. The mixed section also includes x2, x3, x4, x5, x10 review to consolidate all facts learned so far.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete the enabling scaffold with pre-filled x3/x4 intermediate steps. Students only need to double.
- Extra Notes: Distribute Session 3 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate factor pairs -- all the multiplication facts that give a particular product. Find all factor pairs of 24, 36, and 48. Record them systematically.
- Extra Notes: Distribute Session 3 Extension PDF. Self-contained investigation.

WATCH FOR:
- Students who can do x6 but struggle with x8 (because x8 requires recalling x4 which itself is derived) -- suggest they write the full chain
- Students completing all problems correctly and quickly -- direct to extension

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Four questions. On your own, 3 minutes.

DO:
- Display exit ticket. Students work in workbooks.
- Collect/sort for Session 4 planning.

TEACHER NOTES:
Q1-Q2 assess x6 (double x3). Q3-Q4 assess x8 (double x4). Q4 asks students to explain the strategy chain, assessing SC2.

WATCH FOR:
- Students who skip the intermediate step and guess -- they need the strategy reinforced
- Students who get all 4 correct with clear working -- they are ready for Session 4 consolidation

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Success criteria check.
- SC1: I can recall my x3 and x4 facts to use as stepping stones. Thumbs?
- SC2: I can double x3 to get x6 and double x4 to get x8. Thumbs?
- SC3: I can explain the doubling chain for any x6 or x8 fact. Thumbs?
- Turn and talk: We have now covered x2, x3, x4, x5, x6, x8, x10. Which table do you feel most confident with? Which needs more practice?

DO:
- Display SCs. Thumbs for each.
- Turn and Talk 30 seconds, cold call 2-3 students
- Note: tomorrow is x9 plus consolidation of everything

TEACHER NOTES:
By Session 3, students should see the interconnected web of multiplication facts. The Turn and Talk question prompts metacognitive self-assessment that informs Session 4 differentiation.

WATCH FOR:
- Students who are still thumbs-down on SC1 -- they have a gap in x3/x4 that needs targeted intervention before Session 4

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Resources for today are linked here

DO:
- Point out each resource

TEACHER NOTES:
Resource slide for Session 3 companion PDFs.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, "Multiplication Facts", "Session 3: Doubling Known Facts — x6, x8", "Grade 3/4 Numeracy | Session 3 of 4", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      T.addTitle(s, "Addition Algorithm Review", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const problems = [
        "1.  5 128 + 3 469 =",
        "2.  7 694 + 4 758 =",
        "3.  31 246 + 9 875 =",
        "4.  56 789 + 34 567 =",
      ];
      s.addText(problems.map((p, i) => ({
        text: p,
        options: { fontSize: 15, color: C.CHARCOAL, breakLine: i < problems.length - 1, paraSpaceAfter: 12 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Vertical algorithm. Rename where needed.", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: 1.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Answers:  1) 8 597    2) 12 452    3) 41 121    4) 91 356", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  T.addTitle(sFluency, "Subtraction Algorithm Practice", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  const subProblems = [
    "1.  695 - 348 =",
    "2.  903 - 467 =",
    "3.  3 142 - 1 578 =",
    "4.  5 000 - 2 389 =",
    "5.  7 206 - 3 859 =",
    "6.  9 010 - 4 567 =",
  ];
  sFluency.addText(subProblems.map((p, i) => ({
    text: p,
    options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < subProblems.length - 1, paraSpaceAfter: 8 },
  })), {
    x: 0.75, y: CONTENT_TOP + 0.12, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, 1.6, { strip: C.SECONDARY });
  sFluency.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Set out vertically. Watch for zeros!", options: { fontSize: 13, color: C.CHARCOAL } },
  ], {
    x: 5.45, y: CONTENT_TOP + 0.12, w: 3.8, h: 1.36,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to use doubling to work out x6 and x8 multiplication facts from x3 and x4"],
    [
      "I can recall my x3 and x4 facts to use as stepping stones",
      "I can double x3 to get x6 and double x4 to get x8",
      "I can explain the doubling chain for any x6 or x8 fact",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — x6
  workedExSlide(pres, 2, "I Do", "x6 Facts: Double Your x3",
    [
      "6 = 2 x 3, so x6 = double x3",
      "",
      "Example: 6 x 7",
      "  Step 1: 3 x 7 = 21  (x3 fact from Session 2)",
      "  Step 2: Double 21 = 42",
      "  Answer: 6 x 7 = 42",
      "",
      "Check: skip count 7, 14, 21, 28, 35, 42. Yes!",
    ],
    NOTES_IDO_X6, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.PRIMARY });
      slide.addText("The Strategy Chain", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      // Chain: x3 fact -> double -> x6 answer
      const chainY = lg.panelTopPadded + 0.5;
      addTextOnShape(slide, "3 x 7 = 21", {
        x: lg.rightX + 0.2, y: chainY, w: 1.7, h: 0.45, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
      slide.addText("double", {
        x: lg.rightX + 2.0, y: chainY + 0.1, w: 0.6, h: 0.25,
        fontSize: 10, fontFace: FONT_B, color: C.ALERT, margin: 0, align: "center", bold: true,
      });
      addTextOnShape(slide, "6 x 7 = 42", {
        x: lg.rightX + 2.7, y: chainY, w: 1.4, h: 0.45, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });

      // Array: 6 rows of 7, top 3 in one colour, bottom 3 in another
      const dotR = 0.12;
      const dotGap = 0.22;
      const arrX = lg.rightX + 0.3;
      const arrY = lg.panelTopPadded + 1.15;
      for (let row = 0; row < 6; row++) {
        for (let col = 0; col < 7; col++) {
          slide.addShape("roundRect", {
            x: arrX + col * dotGap, y: arrY + row * dotGap,
            w: dotR, h: dotR, rectRadius: dotR / 2,
            fill: { color: row < 3 ? C.SECONDARY : C.PRIMARY },
          });
        }
      }
      slide.addText("3 x 7 = 21", {
        x: lg.rightX + 0.15, y: arrY + 1.4, w: 1.6, h: 0.2,
        fontSize: 9, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText("+ 3 x 7 = 21", {
        x: lg.rightX + 1.8, y: arrY + 1.4, w: 1.6, h: 0.2,
        fontSize: 9, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
    }
  );

  // Slide 7: I Do — x8
  workedExSlide(pres, 2, "I Do", "x8 Facts: Double Your x4",
    [
      "8 = 2 x 4, so x8 = double x4",
      "",
      "Example: 8 x 6",
      "  Step 1: 4 x 6 = 24  (double-double from Session 2)",
      "  Step 2: Double 24 = 48",
      "  Answer: 8 x 6 = 48",
      "",
      "Or think: triple-double (double, double, double)",
      "  6 -> 12 -> 24 -> 48",
    ],
    NOTES_IDO_X8, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.ACCENT });
      slide.addText("Triple-Double Chain", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });
      // Chain: 6 -> 12 -> 24 -> 48
      const chainY = lg.panelTopPadded + 0.5;
      const nums = ["6", "12", "24", "48"];
      const colors = [C.MUTED, C.SECONDARY, C.PRIMARY, C.ALERT];
      const bw = 0.7;
      const gap = 0.28;
      nums.forEach((n, i) => {
        const bx = lg.rightX + 0.15 + i * (bw + gap);
        addTextOnShape(slide, n, {
          x: bx, y: chainY, w: bw, h: 0.42, rectRadius: 0.06,
          fill: { color: colors[i] },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        if (i < 3) {
          slide.addText("x2", {
            x: bx + bw + 0.02, y: chainY + 0.08, w: gap - 0.04, h: 0.2,
            fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, align: "center", bold: true,
          });
        }
      });
      // Labels
      slide.addText("x2", {
        x: lg.rightX + 0.3, y: chainY + 0.5, w: 1.0, h: 0.22,
        fontSize: 10, fontFace: FONT_B, color: C.SECONDARY, margin: 0, align: "center",
      });
      slide.addText("x4", {
        x: lg.rightX + 1.3, y: chainY + 0.5, w: 1.0, h: 0.22,
        fontSize: 10, fontFace: FONT_B, color: C.PRIMARY, margin: 0, align: "center",
      });
      slide.addText("x8", {
        x: lg.rightX + 2.3, y: chainY + 0.5, w: 1.2, h: 0.22,
        fontSize: 10, fontFace: FONT_B, color: C.ALERT, margin: 0, align: "center",
      });
    }
  );

  // Slide 8-9: CFU 1 (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Cold Call",
      "6 x 5 = ?\n\nBe ready to explain your strategy.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "6 x 5 = 30   (3 x 5 = 15, double 15 = 30)", {
        x: 1.5, y: 4.0, w: 7, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do (withReveal)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your Turn — With a Partner",
      [
        "Show your strategy steps on your whiteboard.",
        "",
        "Problem 1 (double x3):",
        "  6 x 9 = ___",
        "  Step 1: 3 x 9 = ___",
        "  Step 2: Double ___ = ___",
        "",
        "Problem 2 (double x4):",
        "  8 x 5 = ___",
        "  Step 1: 4 x 5 = ___",
        "  Step 2: Double ___ = ___",
      ],
      NOTES_WEDO_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1) 6 x 9 = 54 (3x9=27, dbl 27=54)   2) 8 x 5 = 40 (4x5=20, dbl 20=40)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check — Four Corners", "Four Corners",
      "8 x 7 = ?\n\nCorner 1: 54\nCorner 2: 56\nCorner 3: 58\nCorner 4: 64",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Corner 2: 56   (4 x 7 = 28, double 28 = 56)", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "You Do — x6 and x8 Facts",
    [
      "First: Is it x6 or x8?",
      "Next: Find the x3 or x4 fact first.",
      "Then: Double it to get your answer.",
      "",
      "Show your working steps.",
      "8-10 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.ALERT });
      slide.addText("Strategy Reminder", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      const reminders = [
        { label: "x6", desc: "Double your x3 fact", color: C.PRIMARY },
        { label: "x8", desc: "Double your x4 fact", color: C.ACCENT },
      ];
      reminders.forEach((r, i) => {
        const ry = lg.panelTopPadded + 0.5 + i * 0.7;
        addTextOnShape(slide, r.label, {
          x: lg.rightX + 0.2, y: ry, w: 0.8, h: 0.4, rectRadius: 0.06,
          fill: { color: r.color },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(r.desc, {
          x: lg.rightX + 1.15, y: ry, w: lg.rightW - 1.35, h: 0.4,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
        });
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "6 x 8 = ___  (show your steps)",
      "8 x 4 = ___  (show your steps)",
      "6 x 12 = ___",
      "Explain: How does knowing x3 help you with x6?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "We now know x2, x3, x4, x5, x6, x8, and x10. Which table do you feel most confident with? Which needs more practice? Tell your partner.",
    [
      "I can recall x3 and x4 facts as stepping stones",
      "I can double x3 to get x6 and double x4 to get x8",
      "I can explain the doubling chain for x6 or x8 facts",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MF_Session3_Doubling_x6_x8.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Multiplication Facts — x6 and x8",
      color: C.NAVY, lessonInfo: "Session 3 of 4 | Grade 3/4 Numeracy",
    });
    y = addTipBox(doc, "x6 = double your x3 fact. x8 = double your x4 fact. Show your steps!", y, { color: C.TEAL });
    y = addSectionHeading(doc, "x6 Facts (Double x3)", y, { color: C.NAVY });
    y = addProblem(doc, 1, "6 x 4 = ___    Step 1: 3 x 4 = ___    Step 2: Double ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "6 x 6 = ___    Step 1: 3 x 6 = ___    Step 2: Double ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "6 x 8 = ___    Steps: ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "6 x 11 = ___   Steps: ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "x8 Facts (Double x4)", y, { color: C.NAVY });
    y = addProblem(doc, 5, "8 x 3 = ___    Step 1: 4 x 3 = ___    Step 2: Double ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "8 x 6 = ___    Step 1: 4 x 6 = ___    Step 2: Double ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 7, "8 x 9 = ___    Steps: ___", y, { color: C.NAVY });
    y = addProblem(doc, 8, "8 x 12 = ___   Steps: ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Mixed Review (All Tables So Far)", y, { color: C.NAVY });
    y = addProblem(doc, 9, "3 x 8 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 10, "5 x 7 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 11, "4 x 9 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 12, "10 x 6 = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 13, "Is 6 x 8 the same as 8 x 6? Prove it using your strategies.", y, { color: C.NAVY, writeLines: [{ label: "" }] });
    addPdfFooter(doc, "Session 3 | Multiplication Facts | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Multiplication Facts — x6 and x8", color: C.NAVY,
      lessonInfo: "Session 3 of 4 | Grade 3/4 Numeracy",
    });
    y = addSectionHeading(doc, "x6", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 6 x 4 = 24 (3x4=12, dbl 12=24)", y);
    y = addBodyText(doc, "2. 6 x 6 = 36 (3x6=18, dbl 18=36)", y);
    y = addBodyText(doc, "3. 6 x 8 = 48 (3x8=24, dbl 24=48)", y);
    y = addBodyText(doc, "4. 6 x 11 = 66 (3x11=33, dbl 33=66)", y);
    y = addSectionHeading(doc, "x8", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 8 x 3 = 24 (4x3=12, dbl 12=24)", y);
    y = addBodyText(doc, "6. 8 x 6 = 48 (4x6=24, dbl 24=48)", y);
    y = addBodyText(doc, "7. 8 x 9 = 72 (4x9=36, dbl 36=72)", y);
    y = addBodyText(doc, "8. 8 x 12 = 96 (4x12=48, dbl 48=96)", y);
    y = addSectionHeading(doc, "Mixed", y, { color: C.NAVY });
    y = addBodyText(doc, "9. 3 x 8 = 24   10. 5 x 7 = 35   11. 4 x 9 = 36   12. 10 x 6 = 60", y);
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "13. Yes. 6 x 8: 3x8=24, dbl=48. 8 x 6: 4x6=24, dbl=48. Both = 48. Commutative property.", y);
    addPdfFooter(doc, "Session 3 | Multiplication Facts | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Doubling Bridge for x6 and x8", color: C.TEAL,
      lessonInfo: "Session 3 of 4 | Grade 3/4 Numeracy",
    });
    y = addTipBox(doc, "Step 1 is done for you. Just double the answer to find the x6 or x8 fact.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "x6 = Double x3", y, { color: C.NAVY });
    y = addBodyText(doc, "6 x 4:  3 x 4 = 12    Double 12 = ___", y);
    y = addBodyText(doc, "6 x 6:  3 x 6 = 18    Double 18 = ___", y);
    y = addBodyText(doc, "6 x 7:  3 x 7 = 21    Double 21 = ___", y);
    y = addBodyText(doc, "6 x 8:  3 x 8 = 24    Double 24 = ___", y);
    y = addBodyText(doc, "6 x 9:  3 x 9 = 27    Double 27 = ___", y);
    y = addSectionHeading(doc, "x8 = Double x4", y, { color: C.NAVY });
    y = addBodyText(doc, "8 x 3:  4 x 3 = 12    Double 12 = ___", y);
    y = addBodyText(doc, "8 x 5:  4 x 5 = 20    Double 20 = ___", y);
    y = addBodyText(doc, "8 x 6:  4 x 6 = 24    Double 24 = ___", y);
    y = addBodyText(doc, "8 x 7:  4 x 7 = 28    Double 28 = ___", y);
    y = addBodyText(doc, "8 x 9:  4 x 9 = 36    Double 36 = ___", y);
    y = addTipBox(doc, "Tip for doubling big numbers: double the tens, double the ones, add them. Example: double 27 = double 20 (40) + double 7 (14) = 54.", y, { color: C.TEAL });
    addPdfFooter(doc, "Session 3 | Enabling Scaffold | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Factor Pairs
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Factor Pairs Investigation", color: C.NAVY,
      lessonInfo: "Session 3 of 4 | Grade 3/4 Numeracy",
    });
    y = addSectionHeading(doc, "What Are Factor Pairs?", y, { color: C.NAVY });
    y = addBodyText(doc, "A factor pair is two numbers that multiply together to make a product.", y);
    y = addBodyText(doc, "Example: The factor pairs of 12 are:", y);
    y = addBodyText(doc, "  1 x 12 = 12", y);
    y = addBodyText(doc, "  2 x 6 = 12", y);
    y = addBodyText(doc, "  3 x 4 = 12", y);
    y = addBodyText(doc, "So 12 has three factor pairs: (1,12), (2,6), (3,4).", y);
    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "Find ALL the factor pairs of these numbers:", y);
    y = addWriteLine(doc, "Factor pairs of 24:", y);
    y = addWriteLine(doc, "", y);
    y = addWriteLine(doc, "Factor pairs of 36:", y);
    y = addWriteLine(doc, "", y);
    y = addWriteLine(doc, "Factor pairs of 48:", y);
    y = addWriteLine(doc, "", y);
    y = addSectionHeading(doc, "Think About It", y, { color: C.NAVY });
    y = addBodyText(doc, "Which number has the most factor pairs? Why do you think that is?", y);
    y = addWriteLine(doc, "I think:", y);
    addPdfFooter(doc, "Session 3 | Extension | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 3 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
