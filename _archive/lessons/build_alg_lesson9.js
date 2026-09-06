"use strict";

// Algebra Unit — Session 9: Finding Pairs of Unknown Values
// Week 2 Session 4, Grade 5/6 Numeracy, Variant 0
// DR: Exploring Mathematical Patterns and Algorithms
// VC2M6A02 — finding pairs of unknown values that make equations true

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

const SESSION = 9;
const FOOTER = "Algebra | Session 9 of 10 | Grade 5/6 Numeracy";
const OUT_DIR = "output/ALG_Session9_Pairs_of_Unknowns";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Finding pairs of unknowns that make equations true.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Systematic table scaffold for finding pairs.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Investigation into equations with infinite solutions.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

const NOTES_TITLE = `SAY:
- Session 9. Nearly there!
- Today is a different challenge. Instead of finding ONE unknown, you find PAIRS of numbers that work
- Some equations have many solutions, not just one

DO:
- Display title slide

TEACHER NOTES:
Session 9 of 10. This lesson directly addresses the curriculum elaboration: "finding pairs of unknown values in numerical equations that make the equation hold true; for example, listing possible combinations of natural numbers that make this statement true: 6 + 4 x 8 = 6 x triangle + square."

WATCH FOR:
- Students who are surprised that equations can have multiple solutions

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily review -- patterns
- Find the rule and missing number

DO:
- Display problems, 90 seconds

TEACHER NOTES:
DR retrieves pattern work.

WATCH FOR:
- Pattern fluency

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers

DO:
- Reveal, self-check

TEACHER NOTES:
Tick-and-fix.

WATCH FOR:
- Confidence level

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency -- solve these equations quickly

DO:
- Display 8 single-unknown equations, 2 minutes

TEACHER NOTES:
Fluency with single unknowns prepares for the step up to pairs.

WATCH FOR:
- Students who solve automatically

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention
- Read from slide: "We are learning to find pairs of unknown values that make equations true"
- Sometimes an equation has TWO unknowns. There is not just one answer -- there are MANY pairs that work.
- Read our success criteria

DO:
- Choral read

TEACHER NOTES:
The shift from single to paired unknowns is significant. Students must understand that with two unknowns and one equation, there are multiple valid solutions (infinitely many in general, but we restrict to natural numbers).

WATCH FOR:
- Students who expect exactly one answer -- the mindset shift is important

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. Here is the equation: 6 + 4 x 8 = 6 x triangle + square
- First, work out the left side: 6 + 4 x 8 = 6 + 32 = 38
- So we need: 6 x triangle + square = 38
- Now, what pairs of natural numbers (triangle, square) work?
- If triangle = 1: 6 x 1 + square = 38, square = 32. Pair: (1, 32)
- If triangle = 2: 6 x 2 + square = 38, square = 26. Pair: (2, 26)
- If triangle = 3: 6 x 3 + square = 38, square = 20. Pair: (3, 20)
- If triangle = 5: 6 x 5 + square = 38, square = 8. Pair: (5, 8)
- If triangle = 6: 6 x 6 + square = 38, square = 2. Pair: (6, 2)
- Multiple solutions! I find them by trying different values of triangle systematically.

DO:
- Write the systematic table approach
- Show how fixing one unknown lets you calculate the other
- Emphasise "systematic" -- try triangle = 1, 2, 3, etc.

TEACHER NOTES:
The systematic table approach (fix one variable, calculate the other) is the core strategy. This directly matches the curriculum elaboration.

MISCONCEPTIONS:
- Misconception: There is only one correct answer
  Why: All prior equation work had unique solutions
  Impact: Students stop after finding one pair and think they are done
  Quick correction: "This equation has MANY solutions. Your job is to find as many as you can. Keep going with the next value of triangle."

WATCH FOR:
- Students who find one pair and stop
- Students who see the systematic pattern (triangle goes up by 1, square goes down by 6)

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Let me try a simpler one first. triangle + square = 10 (natural numbers only)
- triangle = 1, square = 9. triangle = 2, square = 8. ... triangle = 9, square = 1.
- That is 9 pairs!
- Now with multiplication: triangle x square = 24
- Factor pairs of 24: 1x24, 2x12, 3x8, 4x6, 6x4, 8x3, 12x2, 24x1
- So there are 8 pairs.
- Notice: for addition equations, there are MANY pairs. For multiplication, the pairs are the factor pairs of the product.

DO:
- Show both examples
- Draw the table for each
- Connect factor pairs to multiplication equations

TEACHER NOTES:
Two types of paired-unknown equations: additive (many solutions) and multiplicative (solutions are factor pairs). Students should recognise which type they are working with.

WATCH FOR:
- Students who connect factor pairs from Session 4 to this work
- Students who see the pattern in additive equations

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Find THREE pairs of natural numbers for: triangle + square = 12
- Write them on your whiteboard. 20 seconds.

DO:
- Students write
- Scan for three valid pairs (e.g., 1+11, 2+10, 3+9...)

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "triangle + square = 12. Write THREE pairs. Show me."
- Scan for: any three valid pairs from (1,11), (2,10), (3,9), ..., (11,1).
PROCEED: If 80%+ show three valid pairs, continue.
PIVOT: Students write only one pair or invalid pairs. Reteach: "Be systematic. Start with triangle = 1: what does square have to be? [11]. Now triangle = 2: square = ? [10]. Keep going!"

TEACHER NOTES:
Tests the systematic approach and the understanding that multiple solutions exist.

WATCH FOR:
- Students who find exactly three and stop vs those who want to find all
- Students who give duplicate pairs (1+11 and 11+1 -- discuss: are these different?)

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- With your partner, find as many pairs as you can for each equation
- Be systematic! Use a table.

DO:
- Display two equations
- Partners work together, 3 minutes

TEACHER NOTES:
We Do has one additive and one multiplicative equation. Students should recognise the different structure.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use Session 9 Enabling Scaffold with pre-drawn tables. Column 1 has triangle values pre-filled (1, 2, 3, ...). Students calculate the corresponding square value for each row.
- Extra Notes: Distribute Session 9 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate equations with constraints. In triangle x 3 + square = 20, find all natural number pairs. Then add the constraint "triangle must be even" -- how does that change the solution set?
- Extra Notes: Distribute Session 9 Extension PDF.

WATCH FOR:
- Students who use systematic tables vs random guessing
- Readiness signal: finding all pairs for both equations

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's share. 2 x triangle + square = 20.
- triangle=1: square=18. triangle=2: square=16. ...triangle=9: square=2.
- How many pairs? [9 pairs where both are natural numbers]
- triangle x square = 36.
- Factor pairs: 1x36, 2x18, 3x12, 4x9, 6x6, 9x4, 12x3, 18x2, 36x1.

DO:
- Reveal, cold call
- Discuss: how did you know when to stop?

TEACHER NOTES:
For the additive equation, stop when square would be 0 or negative. For multiplicative, stop when you have found all factor pairs.

WATCH FOR:
- Students who found all pairs for both

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge. triangle x square = 20. How many pairs of natural numbers are there?
- 1 finger: 4 pairs    2 fingers: 6 pairs    3 fingers: 8 pairs

DO:
- Students vote

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "triangle x square = 20. How many natural number pairs? 1=4, 2=6, 3=8."
- Scan for: 2 fingers (6 pairs: 1x20, 2x10, 4x5, 5x4, 10x2, 20x1).
PROCEED: If 80%+ show 2 fingers, move to You Do.
PIVOT: Students show 1 finger (4 pairs -- they only counted 1x20, 2x10, 4x5 and missed the reverses). Reteach: "Remember, 4x5 and 5x4 are different pairs because triangle and square are different positions. triangle=4, square=5 is different from triangle=5, square=4."

TEACHER NOTES:
Whether order matters depends on interpretation. In this unit, (4,5) and (5,4) are distinct because triangle and square are different unknowns.

WATCH FOR:
- Students who debate whether order matters -- great mathematical discussion

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice. Find pairs for each equation.
- Be systematic -- use tables!
- 8 minutes.

DO:
- Distribute worksheet, circulate

TEACHER NOTES:
You Do uses different equations from We Do.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pre-drawn tables with first column filled in.
- Extra Notes: Distribute Session 9 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Constrained solutions investigation.
- Extra Notes: Distribute Session 9 Extension PDF.

WATCH FOR:
- Students who use systematic tables consistently

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. 3 questions. 3 minutes.

DO:
- Display, collect

TEACHER NOTES:
Q1: find 3 pairs for additive. Q2: find all pairs for multiplicative. Q3: explain strategy.

WATCH FOR:
- Q2 -- factor pair fluency

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- SC check. Thumbs.
- SC1: I can find multiple pairs of values that make an equation true. Thumbs?
- SC2: I can use a systematic table to find all pairs. Thumbs?
- SC3: I can explain why some equations have many solutions. Thumbs?
- Turn and talk: What is the difference between finding ONE unknown and finding PAIRS?

DO:
- Thumbs, Turn and Talk, cold call

TEACHER NOTES:
The key insight: one equation with two unknowns has multiple solutions. This is foundational for secondary algebra (systems of equations).

WATCH FOR:
- Students who articulate "two unknowns means many answers"

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

  titleSlide(pres, "Algebra: Finding Unknown Values", "Session 9: Pairs of Unknown Values", "Grade 5/6 Numeracy | Session 9 of 10 | Week 2", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]); addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Patterns and Algorithms", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText(["1.  5, 15, 45, 135, ___  (rule: ?)", "2.  1024, 256, 64, 16, ___  (rule: ?)", "3.  7, 21, 63, 189, ___  (rule: ?)", "4.  10000, 1000, 100, ___  (rule: ?)"].map((p, i) => ({
        text: p, options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < 3, paraSpaceAfter: 14 },
      })), { x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([{ text: "Find the rule and next number.", options: { fontSize: 16, bold: true, color: C.SECONDARY } }], { x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addFooter(s, FOOTER); s.addNotes(NOTES_DR_Q); return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 405 (x3)  2) 4 (/4)  3) 567 (x3)  4) 10 (/10)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  const sFl = pres.addSlide();
  addTopBar(sFl, STAGE_COLORS["1"]); addStageBadge(sFl, 1, "Fluency");
  addTitle(sFl, "Quick Equations", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
  addCard(sFl, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFl.addText(["1.  ? x 6 = 42", "2.  35 / ? = 7", "3.  (? + 3) x 4 = 32", "4.  5 x ? - 8 = 27", "5.  ? x 8 = 72", "6.  60 / ? = 12", "7.  3 x (? + 1) = 18", "8.  4 x ? + 6 = 30"].map((p, i) => ({
    text: p, options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < 7, paraSpaceAfter: 4 },
  })), { x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2, fontFace: FONT_B, margin: 0, valign: "top" });
  addCard(sFl, 5.2, CONTENT_TOP, 4.3, 1.6, { strip: C.SECONDARY });
  sFl.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "7, 5, 5, 7, 9, 5, 5, 6", options: { fontSize: 11, color: C.MUTED } },
  ], { x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.4, fontFace: FONT_B, margin: 0, valign: "top" });
  addFooter(sFl, FOOTER); sFl.addNotes(NOTES_FLUENCY);

  liSlide(pres,
    ["We are learning to find pairs of unknown values that make equations true"],
    [
      "I can find multiple pairs of values that make an equation true",
      "I can use a systematic table to find all natural number solutions",
      "I can explain why some equations have many solutions",
    ],
    NOTES_LI_SC, FOOTER);

  workedExSlide(pres, 2, "I Do", "Pairs of Unknowns: The Curriculum Example",
    [
      "6 + 4 x 8 = 6 x triangle + square",
      "Left side: 6 + 32 = 38",
      "So: 6 x triangle + square = 38",
      "",
      "Try triangle = 1: 6 + square = 38, square = 32",
      "Try triangle = 2: 12 + square = 38, square = 26",
      "Try triangle = 3: 18 + square = 38, square = 20",
      "Try triangle = 5: 30 + square = 38, square = 8",
      "Try triangle = 6: 36 + square = 38, square = 2",
      "",
      "Multiple solutions! Be systematic.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Solution Table", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      // Simple table
      const colW = (lg.rightW - 0.4) / 2;
      const tblX = lg.rightX + 0.2;
      const tblY = lg.panelTopPadded + 0.4;
      const rowH = 0.28;
      // Header
      slide.addShape("rect", { x: tblX, y: tblY, w: colW, h: rowH, fill: { color: C.PRIMARY }, line: { color: C.WHITE, width: 1 } });
      slide.addText("Triangle", { x: tblX, y: tblY, w: colW, h: rowH, fontSize: 11, fontFace: FONT_B, color: C.WHITE, align: "center", valign: "middle", bold: true, margin: 0 });
      slide.addShape("rect", { x: tblX + colW, y: tblY, w: colW, h: rowH, fill: { color: C.PRIMARY }, line: { color: C.WHITE, width: 1 } });
      slide.addText("Square", { x: tblX + colW, y: tblY, w: colW, h: rowH, fontSize: 11, fontFace: FONT_B, color: C.WHITE, align: "center", valign: "middle", bold: true, margin: 0 });
      // Rows
      const pairs = [[1,32],[2,26],[3,20],[4,14],[5,8],[6,2]];
      pairs.forEach((pair, i) => {
        const ry = tblY + rowH * (i + 1);
        slide.addShape("rect", { x: tblX, y: ry, w: colW, h: rowH, fill: { color: C.WHITE }, line: { color: C.MUTED, width: 0.5 } });
        slide.addText(String(pair[0]), { x: tblX, y: ry, w: colW, h: rowH, fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0 });
        slide.addShape("rect", { x: tblX + colW, y: ry, w: colW, h: rowH, fill: { color: C.WHITE }, line: { color: C.MUTED, width: 0.5 } });
        slide.addText(String(pair[1]), { x: tblX + colW, y: ry, w: colW, h: rowH, fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0 });
      });
    }
  );

  workedExSlide(pres, 2, "I Do", "Additive vs Multiplicative Pairs",
    [
      "Additive: triangle + square = 10",
      "  (1,9) (2,8) (3,7) ... (9,1) -- 9 pairs!",
      "",
      "Multiplicative: triangle x square = 24",
      "  Factor pairs: (1,24) (2,12) (3,8) (4,6)",
      "  (6,4) (8,3) (12,2) (24,1) -- 8 pairs",
      "",
      "Additive equations: many pairs",
      "Multiplicative: pairs are factor pairs",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.3, { strip: C.SECONDARY });
      slide.addText("Addition: lots of pairs!", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText("triangle + square = 10\n9 natural number pairs", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.35, w: lg.rightW - 0.6, h: 0.6,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });

      addCard(slide, lg.rightX, lg.panelTopPadded + 1.5, lg.rightW, 1.3, { strip: C.ACCENT });
      slide.addText("Multiplication: factor pairs!", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.56, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
      });
      slide.addText("triangle x square = 24\n8 natural number pairs", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.85, w: lg.rightW - 0.6, h: 0.6,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
    }
  );

  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Show Me Boards",
      "Find THREE pairs of natural numbers:\n\ntriangle + square = 12",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Any 3 from: (1,11) (2,10) (3,9) (4,8) (5,7) (6,6) (7,5) ...", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Find All Pairs",
      [
        "Find ALL natural number pairs for each:",
        "",
        "1.   2 x triangle + square = 20",
        "     (use a table: triangle = 1, 2, 3, ...)",
        "",
        "2.   triangle x square = 36",
        "     (list factor pairs of 36)",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.SECONDARY });
        slide.addText("Strategy:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "Fix one unknown (e.g., triangle = 1)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Calculate the other", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Try the next value", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Stop when square <= 0", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.5, h: 1.0, fontFace: FONT_B, margin: 0, valign: "top" });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1) 9 pairs: (1,18)(2,16)...(9,2)  |  2) Factor pairs of 36: (1,36)(2,18)(3,12)(4,9)(6,6)...", {
        x: 0.3, y: 4.55, w: 9.4, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge — How Many Pairs?", "Finger Voting",
      "triangle x square = 20\n\nHow many natural number pairs?\n1: 4 pairs    2: 6 pairs    3: 8 pairs",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "2: 6 pairs — (1,20)(2,10)(4,5)(5,4)(10,2)(20,1)", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "First: Identify the equation type (add or multiply).",
      "Next: Set up a systematic table.",
      "Then: Find ALL natural number pairs.",
      "",
      "Show your table. 8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.4, { strip: C.ALERT });
      slide.addText("Be Systematic!", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Additive: try 1, 2, 3, ...", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Multiplicative: list factor pairs", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.5, h: 0.7, fontFace: FONT_B, margin: 0, valign: "top" });
    }
  );

  exitTicketSlide(pres,
    [
      "Find 3 natural number pairs for: triangle + square = 15",
      "Find ALL natural number pairs for: triangle x square = 18",
      "Explain why triangle + square = 10 has more solutions than triangle x square = 10.",
    ],
    NOTES_EXIT, FOOTER);

  closingSlide(pres,
    "What is the difference between finding one unknown and finding pairs of unknowns? Tell your partner.",
    [
      "I can find multiple pairs that make an equation true",
      "I can use systematic tables to find all solutions",
      "I can explain why some equations have many solutions",
    ],
    NOTES_CLOSING);



  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG_Session9_Pairs_of_Unknowns.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs — kept focused
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Pairs of Unknown Values", color: C.NAVY, lessonInfo: "Session 9 of 10 | Grade 5/6" });
    y = addTipBox(doc, "For each equation: set up a table with one unknown in column 1. Calculate the other unknown for each row.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "A: Additive Pairs", y, { color: C.NAVY });
    y = addProblem(doc, 1, "triangle + square = 18. Find ALL natural number pairs.", y, { color: C.NAVY });
    y = addWriteLine(doc, "Pairs:", y);
    y = addProblem(doc, 2, "3 x triangle + square = 30. Find ALL natural number pairs.", y, { color: C.NAVY });
    y = addWriteLine(doc, "Pairs:", y);
    y = addSectionHeading(doc, "B: Multiplicative Pairs", y, { color: C.NAVY });
    y = addProblem(doc, 3, "triangle x square = 30. List all factor pairs.", y, { color: C.NAVY });
    y = addWriteLine(doc, "Pairs:", y);
    y = addProblem(doc, 4, "triangle x square = 48. List all factor pairs.", y, { color: C.NAVY });
    y = addWriteLine(doc, "Pairs:", y);
    y = addSectionHeading(doc, "C: Mixed", y, { color: C.NAVY });
    y = addProblem(doc, 5, "5 + 3 x triangle = 2 x square. If triangle = 4, find square.", y, { color: C.NAVY });
    y = addProblem(doc, 6, "2 x triangle + 3 x square = 24. Find 3 natural number pairs.", y, { color: C.NAVY });
    y = addWriteLine(doc, "Pairs:", y);
    addPdfFooter(doc, "Session 9 | Algebra | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName)); console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Pairs of Unknown Values", color: C.NAVY, lessonInfo: "Session 9" });
    y = addSectionHeading(doc, "A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. (1,17)(2,16)(3,15)...(17,1) = 17 pairs", y);
    y = addBodyText(doc, "2. (1,27)(2,24)(3,21)(4,18)(5,15)(6,12)(7,9)(8,6)(9,3) = 9 pairs", y);
    y = addSectionHeading(doc, "B", y, { color: C.NAVY });
    y = addBodyText(doc, "3. (1,30)(2,15)(3,10)(5,6)(6,5)(10,3)(15,2)(30,1) = 8 pairs", y);
    y = addBodyText(doc, "4. (1,48)(2,24)(3,16)(4,12)(6,8)(8,6)(12,4)(16,3)(24,2)(48,1) = 10 pairs", y);
    y = addSectionHeading(doc, "C", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 5+3x4=17, 2xsquare=17, square=8.5 (not a natural number). Try triangle=5: 5+15=20, square=10.", y);
    y = addBodyText(doc, "6. (3,6)(6,4)(9,2) are three valid pairs.", y);
    addPdfFooter(doc, "Session 9 | Answer Key | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName)); console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Systematic Table for Pairs", color: C.TEAL, lessonInfo: "Session 9" });
    y = addTipBox(doc, "Fill in the table. Column 1 is given. Calculate column 2.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "Example: triangle + square = 10", y, { color: C.NAVY });
    y = addBodyText(doc, "triangle=1, square=9. triangle=2, square=8. ... Continue until square=1.", y);
    y += 10;
    y = addSectionHeading(doc, "1. triangle + square = 18", y, { color: C.NAVY });
    for (let i = 1; i <= 10; i++) {
      y = addBodyText(doc, "   triangle = " + i + ",  square = 18 - " + i + " = ___", y);
    }
    y += 5;
    y = addSectionHeading(doc, "2. triangle x square = 30", y, { color: C.NAVY });
    y = addBodyText(doc, "List factor pairs: 1 x ___, 2 x ___, 3 x ___, 5 x ___, ...", y);
    y = addWriteLine(doc, "Factor pairs:", y);
    addPdfFooter(doc, "Session 9 | Enabling Scaffold | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName)); console.log("PDF: " + ENABLING_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "Constrained Solutions", color: C.NAVY, lessonInfo: "Session 9" });
    y = addSectionHeading(doc, "Adding Constraints", y, { color: C.NAVY });
    y = addBodyText(doc, "Equations can have constraints that limit the solutions.", y);
    y = addSectionHeading(doc, "Investigation 1", y, { color: C.NAVY });
    y = addBodyText(doc, "triangle x 3 + square = 20. Find all natural number pairs.", y);
    y = addWriteLine(doc, "All pairs:", y);
    y = addBodyText(doc, "Now add constraint: triangle must be EVEN. Which pairs remain?", y);
    y = addWriteLine(doc, "Constrained pairs:", y);
    y = addSectionHeading(doc, "Investigation 2", y, { color: C.NAVY });
    y = addBodyText(doc, "triangle + square = 20, AND triangle > square. Find all pairs.", y);
    y = addWriteLine(doc, "Pairs:", y);
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "Can you find an equation with exactly 1 natural number pair? Exactly 2?", y);
    y = addWriteLine(doc, "1 pair:", y);
    y = addWriteLine(doc, "2 pairs:", y);
    y = addTipBox(doc, "Hint: Multiplication equations with prime numbers have very few factor pairs!", y, { color: C.TEAL });
    addPdfFooter(doc, "Session 9 | Extension | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName)); console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 9 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
