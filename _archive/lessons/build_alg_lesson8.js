"use strict";

// Algebra Unit — Session 8: Finding Unknowns with Multiple Operations
// Week 2 Session 3, Grade 5/6 Numeracy, Variant 0
// DR: Exploring Mathematical Patterns and Algorithms
// VC2M6A02 — find unknown values in equations with brackets and combinations of operations

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
const { C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP, titleSlide, liSlide, cfuSlide, closingSlide, workedExSlide, exitTicketSlide, addStageBadge, addTextOnShape, addCard, addFooter, addTopBar, addTitle, withReveal, STAGE_COLORS } = T;

const SESSION = 8;
const FOOTER = "Algebra | Session 8 of 10 | Grade 5/6 Numeracy";
const OUT_DIR = "output/ALG_Session8_Unknowns_Multiple_Operations";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Finding unknowns in multi-operation equations with brackets.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Step-by-step scaffold: simplify brackets, then solve.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Multi-step equations with two unknowns.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

const NOTES_TITLE = `SAY:
- Session 8. We have order of operations and equivalent sentences with brackets
- Today we combine everything: finding unknowns in equations with brackets AND multiple operations
- This is the full algebra challenge

DO:
- Display title slide

TEACHER NOTES:
Session 8 of 10. Students find unknown values in equations that include brackets and combinations of all four operations. This is the core Level 6 outcome.

WATCH FOR:
- Students who are confident from Sessions 6-7

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily review. Patterns and algorithms
- Find the pattern rule and the missing value

DO:
- Display problems, 90 seconds

TEACHER NOTES:
DR retrieves pattern work from a different strand.

WATCH FOR:
- Students applying multiplicative thinking to patterns

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers

DO:
- Reveal, self-check

TEACHER NOTES:
Quick tick-and-fix.

WATCH FOR:
- Pattern recognition fluency

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency -- mixed operations with brackets. Simplify each.

DO:
- Display 8 expressions, 2 minutes

TEACHER NOTES:
Fluency drills order of operations with brackets -- prerequisite for today's unknown-finding.

WATCH FOR:
- Students processing brackets first consistently

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention
- Read from slide: "We are learning to find unknown values in equations that use brackets and multiple operations"
- Today is detective work with complex equations

DO:
- Choral read

TEACHER NOTES:
This lesson combines order of operations (Session 6) with unknown-finding (Session 4). The bracket adds a layer of complexity.

WATCH FOR:
- Students who connect this to the evaluate-target-solve strategy

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. Find ?: 5 x (? + 3) = 40
- I know the whole expression equals 40. And 5 is multiplied by the bracket.
- So the bracket must equal 40 / 5 = 8.
- ? + 3 = 8. So ? = 5.
- Check: 5 x (5 + 3) = 5 x 8 = 40. Correct!
- The strategy: work backwards from the outside in.
- Start with what you know (the total), undo the outer operation, then solve the bracket.

DO:
- Write each step, showing the "working backwards" process
- Emphasise: undo the outer operation first

TEACHER NOTES:
"Working backwards" or "undoing" is the inverse-operation strategy. When the unknown is inside brackets, students must peel off the outer operation first, then solve what remains.

MISCONCEPTIONS:
- Misconception: Try to solve the bracket first (but you cannot because ? is unknown)
  Why: Students are trained to do brackets first (from Session 6) and try to evaluate (? + 3) directly
  Impact: Students get stuck because they cannot simplify an expression with an unknown
  Quick correction: "You are right that brackets usually go first. But here we CANNOT simplify the bracket because ? is unknown. Instead, work from the OUTSIDE in -- figure out what the bracket must equal by undoing the outer operation."

WATCH FOR:
- Students who try to do brackets first and get stuck -- redirect to "outside in"
- Students who see the inverse operation immediately

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Harder one. Find ?: 3 x ? + 7 = 25
- No brackets this time, but order of operations still matters
- I need to undo the operations in REVERSE order
- Last operation (in order of ops): + 7. Undo: 25 - 7 = 18.
- Previous operation: 3 x ?. So 3 x ? = 18. ? = 6.
- Check: 3 x 6 + 7 = 18 + 7 = 25. Correct!

DO:
- Write each step
- Show the "reverse order of operations" strategy
- Emphasise: undo in REVERSE order

TEACHER NOTES:
When there are no brackets and the unknown is in a multiplication, students must identify the order of operations for the expression, then reverse it. Addition/subtraction (lowest priority) gets undone first, then multiplication/division.

WATCH FOR:
- Students who try to subtract 7 from both sides instinctively -- affirm this as correct inverse thinking
- Students who try to divide 25 by 3 first -- they are not reversing the order correctly

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Find ?: 4 x (? + 2) = 28
- Work backwards: what must the bracket equal? Then find ?.
- 15 seconds.

DO:
- Students write on whiteboards
- Scan for ? = 5

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "4 x (? + 2) = 28. Work backwards. Show me."
- Scan for: 5 (bracket = 28/4 = 7, then ? + 2 = 7, ? = 5).
PROCEED: If 80%+ show 5, continue.
PIVOT: Most likely error -- students write 7 (they found the bracket value but forgot to subtract 2). Reteach: "Good -- you found the bracket equals 7. But ? + 2 = 7. ? is not 7 -- ? is 7 minus 2. ? = 5."

TEACHER NOTES:
Tests the two-step backwards solving: undo outer operation, then solve the bracket.

WATCH FOR:
- Students who stop at 7 (bracket value) without finishing
- Students who get 5 with full working

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- With your partner, find the unknown in each equation
- Work backwards from what you know
- Show all steps

DO:
- Display four problems, 3 minutes

TEACHER NOTES:
We Do mixes bracket unknowns and non-bracket unknowns. Students must choose the right approach.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use Session 8 Enabling Scaffold which guides each step: (1) identify the outer operation, (2) undo it, (3) solve what remains. First two problems are modelled.
- Extra Notes: Distribute Session 8 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Multi-step equations with two unknowns. Given 2 x ? + 3 x triangle = 26, and knowing triangle = 4, find ?. Then create your own two-unknown equations.
- Extra Notes: Distribute Session 8 Extension PDF.

WATCH FOR:
- Students who solve bracket equations but struggle without brackets
- Readiness signal: completing all four with clear working

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check.
- 1: 6 x (? + 1) = 42. Bracket = 42/6 = 7. ?+1 = 7. ? = 6.
- 2: 2 x ? - 5 = 11. Undo -5: 2x? = 16. ? = 8.
- 3: (? - 3) x 5 = 35. Bracket = 35/5 = 7. ?-3 = 7. ? = 10.
- 4: 48 / (? + 2) = 8. Bracket = 48/8 = 6. ?+2 = 6. ? = 4.

DO:
- Reveal one at a time, cold call

TEACHER NOTES:
Problem 4 uses division with the bracket as divisor -- this requires recognising 48/bracket = 8, so bracket = 48/8 = 6.

WATCH FOR:
- Problem 4 errors (students dividing 48 by 2 instead of finding what bracket equals)

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge. Find ?: 5 x ? + 10 = 45
- Finger vote for the answer.
- 1 finger: ? = 7    2 fingers: ? = 9    3 fingers: ? = 5    4 fingers: ? = 8

DO:
- Students vote

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "5 x ? + 10 = 45. Vote: 1=7, 2=9, 3=5, 4=8."
- Scan for: 1 finger (? = 7). Undo +10: 5x? = 35. ? = 7.
PROCEED: If 80%+ show 1 finger, move to You Do.
PIVOT: Students choose 2 (? = 9, from 45/5 = 9 -- they divided without undoing the +10 first). Reteach: "Order matters. In 5x?+10, the +10 happens LAST in order of operations. So undo it FIRST: 45-10=35. Now 5x?=35. ?=7."

TEACHER NOTES:
Distractor 2 (9) is from dividing 45/5 directly without undoing +10. This tests whether students understand reverse order.

WATCH FOR:
- Students who chose 2 -- they need the reverse-order concept reinforced

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice. Mixed equations with unknowns.
- Show all working using the backwards strategy.
- 8 minutes.

DO:
- Distribute worksheet, circulate

TEACHER NOTES:
You Do has different numbers from We Do.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Guided step-by-step scaffold.
- Extra Notes: Distribute Session 8 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Two-unknown equations investigation.
- Extra Notes: Distribute Session 8 Extension PDF.

WATCH FOR:
- Students who show clear working vs those who guess
- Readiness signal: completing 8+ problems correctly

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. 3 questions. 3 minutes.

DO:
- Display, collect

TEACHER NOTES:
Q1: bracket unknown. Q2: no-bracket unknown with order of ops. Q3: explain strategy.

WATCH FOR:
- Q2 errors (reverse-order mistakes)

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- SC check. Thumbs.
- SC1: I can find unknowns inside brackets. Thumbs?
- SC2: I can find unknowns in multi-operation equations. Thumbs?
- SC3: I can explain the backwards strategy. Thumbs?
- Turn and talk: When you see an equation with an unknown, what is your first step?

DO:
- Thumbs, Turn and Talk, cold call

TEACHER NOTES:
The Turn and Talk targets strategy articulation. "Work backwards from what you know" is the key insight.

WATCH FOR:
- Students who articulate "undo the last operation first"

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Resources linked

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

  titleSlide(pres, "Algebra: Finding Unknown Values", "Session 8: Unknowns with Multiple Operations", "Grade 5/6 Numeracy | Session 8 of 10 | Week 2", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]); addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Patterns and Algorithms", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText(["1.  3, 9, 27, 81, ___  (rule: ?)", "2.  256, 64, 16, 4, ___  (rule: ?)", "3.  2, 6, 18, 54, ___  (rule: ?)", "4.  1000, 100, 10, ___  (rule: ?)"].map((p, i) => ({
        text: p, options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < 3, paraSpaceAfter: 14 },
      })), { x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "Find the rule and next number.", options: { fontSize: 16, bold: true, color: C.SECONDARY } },
      ], { x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addFooter(s, FOOTER); s.addNotes(NOTES_DR_Q); return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 243 (x3)  2) 1 (/4)  3) 162 (x3)  4) 1 (/10)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  const sFl = pres.addSlide();
  addTopBar(sFl, STAGE_COLORS["1"]); addStageBadge(sFl, 1, "Fluency");
  addTitle(sFl, "Mixed Operations Practice", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
  addCard(sFl, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFl.addText(["1. 3 x (7 + 2) =", "2. 4 x 6 - 10 =", "3. (15 - 5) x 3 =", "4. 24 / (8 - 2) =", "5. 7 + 5 x 8 =", "6. (3 + 9) / 4 =", "7. 6 x 5 + 3 x 2 =", "8. 50 - (4 + 6) x 3 ="].map((p, i) => ({
    text: p, options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < 7, paraSpaceAfter: 4 },
  })), { x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2, fontFace: FONT_B, margin: 0, valign: "top" });
  addCard(sFl, 5.2, CONTENT_TOP, 4.3, 1.6, { strip: C.SECONDARY });
  sFl.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "27, 14, 30, 4, 47, 3, 36, 20", options: { fontSize: 11, color: C.MUTED } },
  ], { x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.4, fontFace: FONT_B, margin: 0, valign: "top" });
  addFooter(sFl, FOOTER); sFl.addNotes(NOTES_FLUENCY);

  liSlide(pres,
    ["We are learning to find unknown values in equations that use brackets and multiple operations"],
    [
      "I can find an unknown value inside brackets in an equation",
      "I can find unknowns in equations with multiple operations by working backwards",
      "I can check my answer by substituting it back into the original equation",
    ],
    NOTES_LI_SC, FOOTER);

  workedExSlide(pres, 2, "I Do", "Working Backwards: Unknown in Brackets",
    [
      "Find ?: 5 x (? + 3) = 40",
      "",
      "Step 1: What does the bracket times 5 equal? 40.",
      "  So the bracket = 40 / 5 = 8",
      "",
      "Step 2: ? + 3 = 8",
      "  So ? = 8 - 3 = 5",
      "",
      "Check: 5 x (5 + 3) = 5 x 8 = 40. Correct!",
      "",
      "Strategy: work from OUTSIDE IN.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.PRIMARY });
      slide.addText("Backwards Strategy", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      const steps = [
        { n: "1", l: "Find the TOTAL (given)" },
        { n: "2", l: "UNDO the outer operation" },
        { n: "3", l: "SOLVE what remains" },
        { n: "4", l: "CHECK by substituting back" },
      ];
      steps.forEach((st, i) => {
        const sy = lg.panelTopPadded + 0.4 + i * 0.48;
        addTextOnShape(slide, st.n, {
          x: lg.rightX + 0.2, y: sy, w: 0.35, h: 0.32, rectRadius: 0.06, fill: { color: C.PRIMARY },
        }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(st.l, {
          x: lg.rightX + 0.65, y: sy, w: lg.rightW - 0.85, h: 0.32,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
        });
      });
    }
  );

  workedExSlide(pres, 2, "I Do", "No Brackets: Reverse Order of Operations",
    [
      "Find ?: 3 x ? + 7 = 25",
      "",
      "Order of operations says: multiply first, then add.",
      "To undo, reverse the order:",
      "",
      "Step 1: Undo the LAST operation (+7)",
      "  25 - 7 = 18. So 3 x ? = 18.",
      "",
      "Step 2: Undo the multiplication",
      "  ? = 18 / 3 = 6",
      "",
      "Check: 3 x 6 + 7 = 18 + 7 = 25. Correct!",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.SECONDARY });
      slide.addText("Reverse Order", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      slide.addText("Forward order:\n  1. Multiply   2. Add\n\nReverse order (to solve):\n  1. Undo Add   2. Undo Multiply", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.6, h: 1.5,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
    }
  );

  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Show Me Boards",
      "Find ?:\n\n4 x (? + 2) = 28",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Bracket = 28/4 = 7.  ? + 2 = 7.  ? = 5.", {
        x: 0.8, y: 4.0, w: 8.4, h: 0.55, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Find the Unknown",
      [
        "Work backwards to find ? in each.",
        "Show all steps.",
        "",
        "1.   6 x (? + 1) = 42",
        "2.   2 x ? - 5 = 11",
        "3.   (? - 3) x 5 = 35",
        "4.   48 / (? + 2) = 8",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.SECONDARY });
        slide.addText("Steps:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "1. Find the total", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "2. Undo the outer operation", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "3. Solve what remains", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "4. Check!", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.5, h: 1.0, fontFace: FONT_B, margin: 0, valign: "top" });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1) ?=6  2) ?=8  3) ?=10  4) ?=4", {
        x: 1.5, y: 4.55, w: 7, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check", "Finger Voting",
      "5 x ? + 10 = 45\n\n1: ?=7    2: ?=9    3: ?=5    4: ?=8",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1: ? = 7.  Undo +10: 5x?=35. ?=35/5=7.", {
        x: 0.8, y: 4.0, w: 8.4, h: 0.55, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "First: Identify if ? is in brackets or not.",
      "Next: Work backwards -- undo operations in reverse.",
      "Then: Solve and CHECK by substituting back.",
      "",
      "Show all steps. 8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.4, { strip: C.ALERT });
      slide.addText("Backwards Strategy:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText("Total -> Undo outer -> Solve inner -> Check", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.6, h: 0.5,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
    }
  );

  exitTicketSlide(pres,
    [
      "Find ?: 3 x (? + 4) = 27  (show working)",
      "Find ?: 7 x ? - 3 = 39  (show working)",
      "Explain your strategy: how do you work backwards to find an unknown?",
    ],
    NOTES_EXIT, FOOTER);

  closingSlide(pres,
    "When you see an equation with an unknown, what is your first step? Tell your partner.",
    [
      "I can find unknowns inside brackets",
      "I can work backwards with multiple operations",
      "I can check by substituting back",
    ],
    NOTES_CLOSING);



  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG_Session8_Unknowns_Multiple_Operations.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Unknowns with Multiple Operations", color: C.NAVY, lessonInfo: "Session 8 of 10 | Grade 5/6" });
    y = addTipBox(doc, "Strategy: Work BACKWARDS. Undo the last operation first, then solve what remains. Always CHECK!", y, { color: C.TEAL });
    y = addSectionHeading(doc, "A: Unknown in Brackets", y, { color: C.NAVY });
    y = addProblem(doc, 1, "7 x (? + 2) = 63        ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "5 x (? - 4) = 30        ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "(? + 5) x 6 = 54        ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "36 / (? + 1) = 4        ? = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "B: No Brackets", y, { color: C.NAVY });
    y = addProblem(doc, 5, "4 x ? + 8 = 36          ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "6 x ? - 12 = 30         ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 7, "? x 5 + 15 = 50         ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 8, "8 x ? - 20 = 44         ? = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "C: Mixed Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 9, "3 x (? + 7) - 6 = 24    ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 10, "(? x 4) + 10 = 34      ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 11, "50 / (? + 5) = 5       ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 12, "Write your own equation where ? = 9. Include brackets.", y, { color: C.NAVY });
    y = addWriteLine(doc, "", y);
    addPdfFooter(doc, "Session 8 | Algebra | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName)); console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Unknowns with Multiple Operations", color: C.NAVY, lessonInfo: "Session 8" });
    y = addSectionHeading(doc, "A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. Bracket=63/7=9, ?+2=9, ?=7", y); y = addBodyText(doc, "2. Bracket=30/5=6, ?-4=6, ?=10", y);
    y = addBodyText(doc, "3. Bracket=54/6=9, ?+5=9, ?=4", y); y = addBodyText(doc, "4. Bracket=36/4=9, ?+1=9, ?=8", y);
    y = addSectionHeading(doc, "B", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 4x?=36-8=28, ?=7", y); y = addBodyText(doc, "6. 6x?=30+12=42, ?=7", y);
    y = addBodyText(doc, "7. ?x5=50-15=35, ?=7", y); y = addBodyText(doc, "8. 8x?=44+20=64, ?=8", y);
    y = addSectionHeading(doc, "C", y, { color: C.NAVY });
    y = addBodyText(doc, "9. 3x(?+7)=24+6=30, bracket=10, ?+7=10, ?=3", y);
    y = addBodyText(doc, "10. ?x4=34-10=24, ?=6", y); y = addBodyText(doc, "11. ?+5=50/5=10, ?=5", y);
    addPdfFooter(doc, "Session 8 | Answer Key | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName)); console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Backwards Strategy Guide", color: C.TEAL, lessonInfo: "Session 8" });
    y = addTipBox(doc, "Work BACKWARDS: undo the last operation first.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "Example: 5 x (? + 3) = 40", y, { color: C.NAVY });
    y = addBodyText(doc, "The outer operation is x5. Undo: 40 / 5 = 8. Bracket = 8.", y);
    y = addBodyText(doc, "Inside bracket: ? + 3 = 8. Undo +3: ? = 8 - 3 = 5.", y);
    y = addBodyText(doc, "Check: 5 x (5+3) = 5 x 8 = 40. Correct!", y); y += 10;
    y = addSectionHeading(doc, "Example: 3 x ? + 7 = 25", y, { color: C.NAVY });
    y = addBodyText(doc, "Last operation: +7. Undo: 25 - 7 = 18. So 3 x ? = 18.", y);
    y = addBodyText(doc, "Undo x3: ? = 18 / 3 = 6.", y);
    y = addBodyText(doc, "Check: 3 x 6 + 7 = 18 + 7 = 25. Correct!", y); y += 10;
    y = addSectionHeading(doc, "Now you try", y, { color: C.NAVY });
    const probs = ["7 x (? + 2) = 63", "4 x ? + 8 = 36", "(? + 5) x 6 = 54", "6 x ? - 12 = 30"];
    probs.forEach((p, i) => {
      y = addBodyText(doc, (i + 1) + ". " + p, y);
      y = addBodyText(doc, "   Outer operation: ___  Undo it: ___", y);
      y = addBodyText(doc, "   Remaining: ___  Solve: ? = ___", y);
      y = addBodyText(doc, "   Check: ___", y); y += 8;
    });
    addPdfFooter(doc, "Session 8 | Enabling Scaffold | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName)); console.log("PDF: " + ENABLING_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "Two-Unknown Equations", color: C.NAVY, lessonInfo: "Session 8" });
    y = addSectionHeading(doc, "Equations with Two Unknowns", y, { color: C.NAVY });
    y = addBodyText(doc, "Sometimes an equation has TWO unknowns. If you know one, you can find the other.", y);
    y = addSectionHeading(doc, "Worked Example", y, { color: C.NAVY });
    y = addBodyText(doc, "2 x ? + 3 x triangle = 26, and triangle = 4.", y);
    y = addBodyText(doc, "Substitute: 2 x ? + 3 x 4 = 26", y);
    y = addBodyText(doc, "Simplify: 2 x ? + 12 = 26", y);
    y = addBodyText(doc, "Undo +12: 2 x ? = 14. ? = 7.", y);
    y = addSectionHeading(doc, "Your Turn", y, { color: C.NAVY });
    y = addProblem(doc, 1, "5 x ? + 2 x triangle = 36, triangle = 3. Find ?.", y, { color: C.NAVY });
    y = addWriteLine(doc, "", y);
    y = addProblem(doc, 2, "? x 4 + triangle x 6 = 50, triangle = 5. Find ?.", y, { color: C.NAVY });
    y = addWriteLine(doc, "", y);
    y = addProblem(doc, 3, "3 x (? + triangle) = 30, triangle = 6. Find ?.", y, { color: C.NAVY });
    y = addWriteLine(doc, "", y);
    y = addSectionHeading(doc, "Create Your Own", y, { color: C.NAVY });
    y = addBodyText(doc, "Write a two-unknown equation. Give the value of one unknown. Challenge a partner to find the other.", y);
    y = addWriteLine(doc, "", y);
    addPdfFooter(doc, "Session 8 | Extension | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName)); console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 8 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
