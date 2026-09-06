"use strict";

// Algebra Unit — Session 6: Order of Operations with Brackets
// Week 2 Session 1, Grade 5/6 Numeracy, Variant 0
// DR: Solving Equations with Multiplication, Division, and Operations
// Fluency: Order of operations warm-up
// VC2M6A02 — brackets and order of operations

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
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, STAGE_COLORS,
} = T;

const SESSION = 6;
const FOOTER = "Algebra | Session 6 of 10 | Grade 5/6 Numeracy";
const OUT_DIR = "output/ALG_Session6_Order_of_Operations";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - order of operations with brackets.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Step-by-step BODMAS scaffold with brackets highlighted.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into how brackets change answers.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const NOTES_TITLE = `SAY:
- Welcome to Week 2 of our algebra unit
- Last week we learned properties and finding unknowns with single operations
- This week we level up -- brackets, multiple operations, and order of operations
- By Friday you will be solving equations that look really complex but follow clear rules

DO:
- Display title slide

TEACHER NOTES:
Session 6 of 10. Week 2 shifts to Level 6 content (VC2M6A02). This lesson introduces brackets and the agreed order of operations. Students need to understand why an agreed order matters before they can find unknowns in complex equations.

WATCH FOR:
- Students who have heard of BODMAS/PEMDAS but may have incomplete understanding

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily review. Last week's skills -- find the unknowns
- Use your evaluate-target-solve strategy

DO:
- Display equation-solving problems from Week 1
- Allow 90 seconds

TEACHER NOTES:
DR retrieves Week 1 unknown-finding skills. These are the building blocks for Week 2 -- students need fast, confident equation solving before adding brackets.

WATCH FOR:
- Students who have retained the strategy from last week
- Students who need a reminder of evaluate-target-solve

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers
- Ask: What was your first step each time? [Evaluate the known side]

DO:
- Reveal answers, students self-check

TEACHER NOTES:
Reinforcing the evaluate-first strategy from Week 1.

WATCH FOR:
- Students who solved all four quickly -- Week 1 skills are solid

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency -- let's see if you can spot the right order
- For each expression, calculate the answer. Be careful about what to do first!
- Some have brackets, some do not

DO:
- Display 8 expressions
- Time 2 minutes

TEACHER NOTES:
This fluency pre-tests order of operations awareness. Students who calculate left-to-right without considering operation hierarchy will get some wrong. This sets up the I Do perfectly.

WATCH FOR:
- Students who get different answers for expressions without brackets -- the need for agreed rules becomes obvious

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention
- Read from slide: "We are learning to use brackets and the order of operations to solve number sentences correctly"
- Brackets tell us what to do first. Without an agreed order, we get different answers for the same expression.
- Read the success criteria

DO:
- Choral read LI and SC

TEACHER NOTES:
The motivating question: "Without rules, 3 + 4 x 2 could be 14 or 11 -- which is right?" This creates the need for an agreed order of operations.

WATCH FOR:
- Students who confidently say "multiplication first" -- they may already know the convention

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. What is 3 + 4 x 2?
- If I go left to right: 3 + 4 = 7, then 7 x 2 = 14
- But if I do multiplication first: 4 x 2 = 8, then 3 + 8 = 11
- Two different answers for the same expression! That is a problem.
- We need an agreed order. The rule is: multiplication and division BEFORE addition and subtraction
- So the correct answer is 11. Multiplication first, then addition.
- Now what if I WANT to add first? That is where brackets come in.
- (3 + 4) x 2. Brackets first: 3 + 4 = 7. Then 7 x 2 = 14.
- Brackets override the normal order.

DO:
- Write both interpretations clearly
- Circle the multiplication to show it goes first
- Show brackets changing the answer
- Write the order: Brackets -> Multiplication/Division -> Addition/Subtraction

TEACHER NOTES:
The I Do demonstrates WHY we need order of operations -- without it, expressions are ambiguous. Brackets provide explicit control over evaluation order.

MISCONCEPTIONS:
- Misconception: Always work left to right
  Why: Students apply reading order to mathematics
  Impact: Produces incorrect results whenever multiplication/division appears after addition/subtraction
  Quick correction: "In maths, we do not always go left to right. Multiplication and division are more powerful -- they go first. Brackets are the boss -- they go before everything."

WATCH FOR:
- Students who are surprised that the answer changes with brackets
- Students who already knew multiplication goes first -- affirm and extend

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- A bigger example: 40 / 2 x (4 + 6) = ?
- Step 1: Brackets first. (4 + 6) = 10.
- Step 2: Now I have 40 / 2 x 10.
- Multiplication and division are equal rank -- go left to right when they are the same level.
- 40 / 2 = 20. Then 20 x 10 = 200.
- Answer: 200.
- The order: brackets first, then x and / left to right, then + and - left to right.

DO:
- Write each step, circling what gets done at each stage
- Emphasise: x and / are equal rank, + and - are equal rank
- When same rank, go left to right

TEACHER NOTES:
This example directly matches the curriculum elaboration. The key insight is that multiplication and division have equal precedence and are resolved left to right, not multiplication before division.

WATCH FOR:
- Students who try to do multiplication before division (thinking x always comes first) -- clarify equal rank
- Readiness signal: students calling out each step before you write it

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboard: 5 + 3 x 4 = ?
- Remember the order! 15 seconds.

DO:
- Students write on whiteboards
- Scan for 17 (not 32)

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "5 + 3 x 4. Show me your answer."
- Scan for: 17 (3 x 4 = 12, then 5 + 12 = 17).
PROCEED: If 80%+ show 17, continue.
PIVOT: Most likely misconception -- students write 32 (adding 5+3=8 first, then 8x4=32). Reteach: "Multiplication is more powerful than addition. Do 3 x 4 FIRST: that is 12. THEN add 5: 5 + 12 = 17. Not left to right -- multiplication first."

TEACHER NOTES:
Classic order of operations check. The 32 vs 17 split reveals who understands operation hierarchy.

WATCH FOR:
- Students who write 32 -- they are going left to right
- Students who write 17 confidently -- order of operations is understood

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- With your partner, solve each expression
- Show your working -- which operation did you do first?

DO:
- Display four expressions
- Partners work on whiteboards
- Circulate, check the order of steps
- 2 minutes

TEACHER NOTES:
We Do mixes expressions with and without brackets. Students must apply the order of operations consistently.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 6 Enabling Scaffold which highlights the order step by step. Each expression has arrows showing which operation to do first, second, third. Students fill in the calculation at each step.
- Extra Notes: Distribute Session 6 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate how brackets change answers. Given the expression 2 + 3 x 4 - 1, place brackets in different positions and calculate. How many different answers can you create?
- Extra Notes: Distribute Session 6 Extension PDF.

WATCH FOR:
- Students who do brackets first correctly but then go left-to-right for the rest
- Readiness signal: correct answers with clear step-by-step working

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check.
- 1: (8 + 2) x 5. Brackets: 10. Then 10 x 5 = 50.
- 2: 20 - 4 x 3. Multiply first: 4 x 3 = 12. Then 20 - 12 = 8.
- 3: 6 x (10 - 3). Brackets: 7. Then 6 x 7 = 42.
- 4: 15 + 36 / 6 - 2. Divide first: 36/6 = 6. Then 15 + 6 - 2 = 19.

DO:
- Reveal one at a time, cold call
- For each, ask: "What did you do first? Why?"

TEACHER NOTES:
Problem 4 is the most complex -- it requires dividing before adding/subtracting, then working left to right. Students who get this have a solid grasp.

WATCH FOR:
- Problem 2 errors (students adding 20-4 first) -- operation hierarchy not yet automatic

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Multiple choice.
- What is 12 / (2 + 4)?
- A: 10    B: 2    C: 8    D: 6
- Show me which letter on your whiteboard.

DO:
- Students write letter
- Scan for B (brackets first: 2+4=6, then 12/6=2)

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "12 / (2 + 4). Write the letter. A=10, B=2, C=8, D=6."
- Scan for: B (brackets: 2+4=6, then 12/6=2).
PROCEED: If 80%+ show B, move to You Do.
PIVOT: Most likely misconception -- students choose D (they calculated 12/2=6, ignoring brackets). Reteach: "Brackets are the boss. (2+4) = 6. Now the expression is 12/6. What is 12 divided by 6? [2]. The brackets change what we divide by."

TEACHER NOTES:
Distractor analysis: A=forgot brackets, subtracted; C=forgot brackets, added wrong; D=ignored brackets, divided by 2 only.

WATCH FOR:
- Students who choose D -- brackets are not being processed first

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice. Your worksheet has expressions to solve.
- First: Find any brackets and solve those. Next: Do multiplication and division (left to right). Then: Do addition and subtraction (left to right).
- Show all steps. 8 minutes.

DO:
- Distribute Session 6 Worksheet
- Circulate, distribute scaffolds/extensions as needed

TEACHER NOTES:
You Do uses different expressions from We Do, with increasing complexity. Section C introduces writing brackets to make a target answer.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Step-by-step scaffold with operation order arrows.
- Extra Notes: Distribute Session 6 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Brackets investigation -- how many different answers from one expression?
- Extra Notes: Distribute Session 6 Extension PDF.

WATCH FOR:
- Students who show all steps clearly vs those who skip steps and make errors
- Readiness signal: completing Section A correctly without prompting

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Three questions. 3 minutes.

DO:
- Display questions, students work alone, collect

TEACHER NOTES:
Q1 tests basic brackets (SC1). Q2 tests operation hierarchy without brackets (SC2). Q3 tests explaining why brackets change the answer (SC3).

WATCH FOR:
- Students who get Q1 and Q2 but struggle with Q3 -- they can calculate but not yet explain

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Success criteria check.
- SC1: I can solve expressions with brackets by doing brackets first. Thumbs?
- SC2: I can apply the correct order of operations without brackets. Thumbs?
- SC3: I can explain why we need an agreed order of operations. Thumbs?
- Turn and talk: Why do mathematicians agree on an order of operations?

DO:
- Thumbs for each SC
- Turn and Talk
- Cold call 2 students

TEACHER NOTES:
The Turn and Talk targets the WHY -- without agreed rules, expressions are ambiguous. This is a foundational understanding for all future algebra.

WATCH FOR:
- Students who articulate "so everyone gets the same answer" -- that is the key insight

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Resources linked here

DO:
- Point out resources

TEACHER NOTES:
Resource slide.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "Algebra: Finding Unknown Values", "Session 6: Order of Operations & Brackets", "Grade 5/6 Numeracy | Session 6 of 10 | Week 2", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // DR
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]); addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Solving Equations Review", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const probs = ["1.  4 x 7 = ? x 2", "2.  6 x 9 = 108 / ?", "3.  ? x 5 = 45", "4.  5 x 8 = ? x 4"];
      s.addText(probs.map((p, i) => ({
        text: p, options: { fontSize: 15, color: C.CHARCOAL, breakLine: i < probs.length - 1, paraSpaceAfter: 14 },
      })), { x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "Evaluate - Target - Solve", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Show your working.", options: { fontSize: 14, color: C.CHARCOAL } },
      ], { x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addFooter(s, FOOTER); s.addNotes(NOTES_DR_Q); return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) ?=14  2) ?=2  3) ?=9  4) ?=10", {
        x: 1.5, y: 4.55, w: 7, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency
  const sFl = pres.addSlide();
  addTopBar(sFl, STAGE_COLORS["1"]); addStageBadge(sFl, 1, "Fluency");
  addTitle(sFl, "Operation Order Warm-Up", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
  addCard(sFl, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  const fl = ["1.  2 + 3 x 4 =", "2.  (2 + 3) x 4 =", "3.  10 - 6 / 2 =", "4.  (10 - 6) / 2 =", "5.  3 x 5 + 2 =", "6.  3 x (5 + 2) =", "7.  8 / 4 + 6 =", "8.  8 / (4 + 6) ="];
  sFl.addText(fl.map((p, i) => ({
    text: p, options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < fl.length - 1, paraSpaceAfter: 4 },
  })), { x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2, fontFace: FONT_B, margin: 0, valign: "top" });
  addCard(sFl, 5.2, CONTENT_TOP, 4.3, 1.8, { strip: C.SECONDARY });
  sFl.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Be careful about what to do FIRST!", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "14, 20, 7, 2, 17, 21, 8, 0.8", options: { fontSize: 11, color: C.MUTED } },
  ], { x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.6, fontFace: FONT_B, margin: 0, valign: "top" });
  addFooter(sFl, FOOTER); sFl.addNotes(NOTES_FLUENCY);

  // LI/SC
  liSlide(pres,
    ["We are learning to use brackets and the order of operations to solve number sentences correctly"],
    [
      "I can solve expressions with brackets by doing brackets first",
      "I can apply the correct order: brackets, then x and /, then + and -",
      "I can explain why we need an agreed order of operations",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1
  workedExSlide(pres, 2, "I Do", "Why Order Matters",
    [
      "What is 3 + 4 x 2?",
      "",
      "Left to right: 3 + 4 = 7, then 7 x 2 = 14",
      "Multiply first: 4 x 2 = 8, then 3 + 8 = 11",
      "",
      "Two different answers! We need a rule.",
      "",
      "The agreed order:",
      "  1. Brackets first",
      "  2. x and / (left to right)",
      "  3. + and - (left to right)",
      "",
      "Correct answer: 11 (multiply before adding)",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Order of Operations", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      const order = [
        { num: "1st", label: "BRACKETS ( )", col: C.ALERT },
        { num: "2nd", label: "x  and  /", col: C.PRIMARY },
        { num: "3rd", label: "+  and  -", col: C.SECONDARY },
      ];
      order.forEach((o, i) => {
        const oy = lg.panelTopPadded + 0.45 + i * 0.55;
        addTextOnShape(slide, o.num, {
          x: lg.rightX + 0.2, y: oy, w: 0.6, h: 0.38, rectRadius: 0.06,
          fill: { color: o.col },
        }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(o.label, {
          x: lg.rightX + 0.9, y: oy, w: lg.rightW - 1.1, h: 0.38,
          fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, margin: 0, valign: "middle", bold: true,
        });
      });

      // Brackets override
      addTextOnShape(slide, "Brackets are the BOSS!", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.2, w: lg.rightW - 0.6, h: 0.4, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // I Do 2
  workedExSlide(pres, 2, "I Do", "A Bigger Example",
    [
      "40 / 2 x (4 + 6) = ?",
      "",
      "Step 1: Brackets first",
      "  (4 + 6) = 10",
      "",
      "Step 2: x and / left to right",
      "  40 / 2 = 20",
      "  20 x 10 = 200",
      "",
      "Answer: 200",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.SECONDARY });
      slide.addText("Step-by-Step", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      const steps = [
        "40 / 2 x (4 + 6)",
        "     Brackets: (4 + 6) = 10",
        "= 40 / 2 x 10",
        "     Division: 40 / 2 = 20",
        "= 20 x 10",
        "     Multiply: 20 x 10 = 200",
        "= 200",
      ];
      slide.addText(steps.map((st, i) => ({
        text: st, options: { fontSize: 12, color: st.includes("=") && !st.includes(":") ? C.PRIMARY : C.CHARCOAL, bold: st.startsWith("="), breakLine: i < steps.length - 1, paraSpaceAfter: 2 },
      })), { x: lg.rightX + 0.25, y: lg.panelTopPadded + 0.4, w: lg.rightW - 0.5, h: 2.2, fontFace: FONT_B, margin: 0, valign: "top" });
    }
  );

  // CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Show Me Boards",
      "5 + 3 x 4 = ?\n\nRemember the order!",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3 x 4 = 12 first, then 5 + 12 = 17   (not 32!)", {
        x: 0.8, y: 4.0, w: 8.4, h: 0.55, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your Turn — Order of Operations",
      [
        "Solve each expression. Show your steps.",
        "",
        "1.   (8 + 2) x 5 =",
        "2.   20 - 4 x 3 =",
        "3.   6 x (10 - 3) =",
        "4.   15 + 36 / 6 - 2 =",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.SECONDARY });
        slide.addText("Order Reminder:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "1st: Brackets ( )", options: { bullet: true, fontSize: 13, color: C.ALERT, bold: true, breakLine: true } },
          { text: "2nd: x and / (left to right)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "3rd: + and - (left to right)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.5, h: 0.9, fontFace: FONT_B, margin: 0, valign: "top" });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1) 50   2) 8   3) 42   4) 19", {
        x: 1.5, y: 4.55, w: 7, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // CFU 2 Hinge
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check", "Show Me Boards",
      "12 / (2 + 4) = ?\n\nA: 10   B: 2   C: 8   D: 6",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "B: Brackets first (2+4=6), then 12 / 6 = 2", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "First: Find brackets and solve those.",
      "Next: Do x and / (left to right).",
      "Then: Do + and - (left to right).",
      "",
      "Show all steps.",
      "8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.6, { strip: C.ALERT });
      slide.addText("The Order:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText("1. Brackets ( )\n2. x  /  (left to right)\n3. +  -  (left to right)", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.6, h: 0.9,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
    }
  );

  exitTicketSlide(pres,
    [
      "(7 + 3) x 6 = ___  (show steps)",
      "18 - 2 x 5 + 1 = ___  (show steps)",
      "Explain: Why does 3 + 4 x 2 = 11, not 14?",
    ],
    NOTES_EXIT, FOOTER);

  closingSlide(pres,
    "Why do mathematicians agree on an order of operations? Tell your partner.",
    [
      "I can solve expressions with brackets first",
      "I can apply the correct order: brackets, x/, +-",
      "I can explain why we need agreed rules",
    ],
    NOTES_CLOSING);



  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG_Session6_Order_of_Operations.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Order of Operations with Brackets", color: C.NAVY, lessonInfo: "Session 6 of 10 | Grade 5/6" });
    y = addTipBox(doc, "Order: 1. Brackets  2. Multiplication and Division (left to right)  3. Addition and Subtraction (left to right). Show all steps!", y, { color: C.TEAL });
    y = addSectionHeading(doc, "Section A: Brackets First", y, { color: C.NAVY });
    y = addProblem(doc, 1, "(5 + 3) x 7 =", y, { color: C.NAVY });
    y = addProblem(doc, 2, "4 x (9 - 2) =", y, { color: C.NAVY });
    y = addProblem(doc, 3, "(12 + 8) / 5 =", y, { color: C.NAVY });
    y = addProblem(doc, 4, "3 x (6 + 4) - 5 =", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Section B: No Brackets", y, { color: C.NAVY });
    y = addProblem(doc, 5, "7 + 6 x 3 =", y, { color: C.NAVY });
    y = addProblem(doc, 6, "24 / 4 + 8 =", y, { color: C.NAVY });
    y = addProblem(doc, 7, "15 - 3 x 2 + 4 =", y, { color: C.NAVY });
    y = addProblem(doc, 8, "8 x 5 - 12 / 3 =", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Section C: Place the Brackets", y, { color: C.NAVY });
    y = addProblem(doc, 9, "Add brackets to make it true: 2 + 3 x 4 = 20", y, { color: C.NAVY });
    y = addProblem(doc, 10, "Add brackets to make it true: 8 - 2 x 3 = 18", y, { color: C.NAVY });
    y = addProblem(doc, 11, "Add brackets to make it true: 24 / 4 + 2 = 4", y, { color: C.NAVY });
    y = addProblem(doc, 12, "Add brackets to make it true: 5 + 10 / 5 = 3", y, { color: C.NAVY });
    addPdfFooter(doc, "Session 6 | Algebra | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Order of Operations", color: C.NAVY, lessonInfo: "Session 6 of 10" });
    y = addSectionHeading(doc, "A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. (5+3)x7 = 8x7 = 56", y);
    y = addBodyText(doc, "2. 4x(9-2) = 4x7 = 28", y);
    y = addBodyText(doc, "3. (12+8)/5 = 20/5 = 4", y);
    y = addBodyText(doc, "4. 3x(6+4)-5 = 3x10-5 = 30-5 = 25", y);
    y = addSectionHeading(doc, "B", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 7+6x3 = 7+18 = 25", y);
    y = addBodyText(doc, "6. 24/4+8 = 6+8 = 14", y);
    y = addBodyText(doc, "7. 15-3x2+4 = 15-6+4 = 13", y);
    y = addBodyText(doc, "8. 8x5-12/3 = 40-4 = 36", y);
    y = addSectionHeading(doc, "C", y, { color: C.NAVY });
    y = addBodyText(doc, "9. (2+3) x 4 = 5x4 = 20", y);
    y = addBodyText(doc, "10. (8-2) x 3 = 6x3 = 18", y);
    y = addBodyText(doc, "11. 24 / (4+2) = 24/6 = 4", y);
    y = addBodyText(doc, "12. (5+10) / 5 = 15/5 = 3", y);
    addPdfFooter(doc, "Session 6 | Answer Key | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Step-by-Step Order Guide", color: C.TEAL, lessonInfo: "Session 6 of 10" });
    y = addTipBox(doc, "Follow the arrows. At each step, circle the operation you do FIRST, then calculate.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "The Order", y, { color: C.NAVY });
    y = addBodyText(doc, "Step 1: Do BRACKETS first -> ( )", y);
    y = addBodyText(doc, "Step 2: Do MULTIPLY and DIVIDE next -> x /  (left to right)", y);
    y = addBodyText(doc, "Step 3: Do ADD and SUBTRACT last -> + -  (left to right)", y);
    y += 10;
    y = addSectionHeading(doc, "Example: (8 + 2) x 5", y, { color: C.NAVY });
    y = addBodyText(doc, "Circle brackets first: (8 + 2) = 10", y);
    y = addBodyText(doc, "Now: 10 x 5 = 50", y);
    y += 10;
    y = addSectionHeading(doc, "Practice", y, { color: C.NAVY });
    const probs = ["(5 + 3) x 7", "4 x (9 - 2)", "7 + 6 x 3", "24 / 4 + 8"];
    probs.forEach((p, i) => {
      y = addBodyText(doc, (i + 1) + ". " + p, y);
      y = addBodyText(doc, "   Circle first operation: ___", y);
      y = addBodyText(doc, "   After step 1: ___", y);
      y = addBodyText(doc, "   After step 2: ___", y);
      y = addBodyText(doc, "   Answer: ___", y);
      y += 8;
    });
    addPdfFooter(doc, "Session 6 | Enabling Scaffold | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF: " + ENABLING_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "Brackets Change Everything!", color: C.NAVY, lessonInfo: "Session 6 of 10" });
    y = addSectionHeading(doc, "The Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "Take this expression: 2 + 3 x 4 - 1", y);
    y = addBodyText(doc, "Without brackets, the answer is: 2 + 12 - 1 = 13", y);
    y = addBodyText(doc, "But by placing brackets in different positions, you can create different answers!", y);
    y = addSectionHeading(doc, "Try These Bracket Placements", y, { color: C.NAVY });
    y = addProblem(doc, 1, "(2 + 3) x 4 - 1 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "2 + 3 x (4 - 1) = ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "(2 + 3) x (4 - 1) = ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "2 + (3 x 4 - 1) = ___", y, { color: C.NAVY });
    y = addProblem(doc, 5, "(2 + 3 x 4) - 1 = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Your Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "Take the expression: 1 + 2 x 3 + 4 x 5", y);
    y = addBodyText(doc, "How many different answers can you create by placing brackets?", y);
    y = addBodyText(doc, "List each bracket placement and its answer:", y);
    y = addWriteLine(doc, "", y); y = addWriteLine(doc, "", y); y = addWriteLine(doc, "", y);
    y = addWriteLine(doc, "", y); y = addWriteLine(doc, "", y);
    y = addTipBox(doc, "You can use one set of brackets, two sets, or even nested brackets like ((2+3) x 4). Each arrangement can give a different answer!", y, { color: C.TEAL });
    addPdfFooter(doc, "Session 6 | Extension | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 6 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
