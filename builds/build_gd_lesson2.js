"use strict";

// Geometry and Data Unit (Year 5/6 Numeracy) - Lesson 2: Symmetry and nets.
// Year 5 VC2M5SP03 (symmetry) and VC2M5SP01 (nets) + Year 6 VC2M6SP01 (cross-sections).
// Uses kinder square paper for hands-on folding and cutting.
// Daily Review: Solving equations with multiplication, division and operations.
// Fluency: Multiplication facts (slightly less scaffold than yesterday - gradual release).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 2);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 8;
const UNIT_TITLE = "Geometry and Data: Shapes, Tessellations, Data";
const FOOTER = `Geometry and Data | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/GD_Lesson2_Symmetry_and_Nets";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 2 Symmetry and Nets",
  "Draw lines of symmetry. Match nets to 3D solids.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 2 Answer Key",
  "Worked answers for the Symmetry and Nets sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// =====================================================================
// Teacher notes
// =====================================================================

const NOTES_TITLE = `SAY:
- Yesterday we moved shapes. Today we look INSIDE shapes for lines of symmetry, and we open up 3D shapes into flat nets.
- Some of you may remember symmetry. If this feels new, that is okay.

DO:
- Hand out kinder square paper to each student.
- Whiteboards ready.

TEACHER NOTES:
Lesson 2 of 8. Today combines symmetry and nets in one session because nets give us a natural place to talk about symmetry of faces.

WATCH FOR:
- Students who already know "line of symmetry" - secure for extension.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- One kinder square paper per student (for folding).
- One Symmetry and Nets sheet per student.
- Scissors per pair (for the optional cube net cut-out).

DO:
- Print one student sheet per student.
- Distribute kinder squares.

TEACHER NOTES:
The kinder square is the manipulative for the folding investigation. It can be cut into rectangles, triangles or kept square. Keep a couple of pre-cut shapes (equilateral triangle, hexagon) for the extension.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review.
- Solve each equation on your whiteboard.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Same skill as yesterday - solve for n. Slightly larger numbers and a two-step equation.

WATCH FOR:
- Students who write the inverse - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 5 x n = 45, so n = 9.
- n divided by 6 = 7, so n = 42.
- 3 x n - 4 = 20, so 3n = 24, n = 8.

DO:
- Click to reveal.

TEACHER NOTES:
Quick reveal. Note any student who can explain why we add 4 before dividing by 3 in the last one.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Multiplication facts.
- Slightly harder than yesterday. Whisper-answer then write.

DO:
- Display the three prompts.
- 25 seconds.

TEACHER NOTES:
Day two of multiplication fluency. Less scaffold than Day 1 - smaller time window. Includes a 2-digit by 1-digit.

WATCH FOR:
- Students who recall doubles to help (12 x 4 as double of 12 x 2) - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 8 x 7 = 56.
- 9 x 9 = 81.
- 12 x 4 = 48.

DO:
- Click to reveal.

TEACHER NOTES:
Brisk reveal.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Hold up your kinder square.
- Fold it in half so the corners match exactly.
- That fold line is a LINE OF SYMMETRY.
- Now open it and fold a different way. Did you find another line?
- A square has 4 lines of symmetry. Today we find lines of symmetry for other shapes, and we open up 3D solids into nets.

DO:
- Mirror the folds with a kinder square at the front.
- Have students compare with a partner.

TEACHER NOTES:
This launch makes the line of symmetry concrete. Students who feel the corners match know the fold is a line of symmetry. Half-half-rotation folds (corner to corner) often surprise them.

WATCH FOR:
- Students who find all 4 folds on their square - secure.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to identify lines of symmetry and match 3D solids to their nets.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 - lines of symmetry on a simple shape. SC2 - the core target - match a net to its solid. SC3 - parallel cross-sections of a right prism (Year 6 stretch).

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Three key words today.
- Line of symmetry - a line where the shape folds onto itself exactly.
- Net - a flat shape that folds up into a 3D solid.
- Cross-section - the flat shape you see when you slice a solid.

DO:
- Point at each visual.
- Have students whisper each word back.

TEACHER NOTES:
"Cross-section" is the Year 6 stretch term. Anchor it concretely: cutting a loaf of bread; cutting a carrot; cutting a Toblerone box across.

WATCH FOR:
- Students who can name each term with the image - tracking.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_IDO_SYMMETRY = `SAY:
- Watch me. I draw a rectangle. How many lines of symmetry?
- I can fold it left-right. That works. 1 line.
- I can fold it top-bottom. That works too. 2 lines.
- Can I fold corner to corner? Let me check. No - the corners do not match.
- A rectangle has exactly 2 lines of symmetry.

DO:
- Demonstrate with a kinder square folded into a rectangle, or use a separate piece of paper.
- Show each fold one at a time.

TEACHER NOTES:
The rectangle catches lots of students - they try the diagonals because the diagonals work for a square. Use the actual paper to prove it.

MISCONCEPTIONS:
- Misconception: A rectangle has 4 lines of symmetry (like a square).
  Why: Students remember the square has 4 lines and assume any rectangle does too.
  Impact: Wrong on the worksheet.
  Quick correction: Fold the paper diagonally - corners do not match.

WATCH FOR:
- Students who fold to check - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_NET = `SAY:
- Watch me. Here is a cube.
- If I cut along some edges and unfold it flat, I get a NET.
- A cube net has 6 squares. The most common shape is a plus or cross.
- Watch how I can fold these 6 squares back into the cube.

DO:
- Use the pre-cut cross-shaped net and fold it up.
- Point at each face: top, bottom, front, back, left, right.

TEACHER NOTES:
There are 11 different cube nets. Today's I Do uses the cross-shaped one. Other arrangements work too; the test is whether they fold up without overlap.

WATCH FOR:
- Students who can name the bottom face before you fold - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- Three different arrangements of six squares are on the slide.
- Which one is a valid net of a cube?
- Write A, B or C.

DO:
- Display the prompt.
- 45 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me on three, two, one, show."
- Scan for: B.
PROCEED: If 80% have B, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count 6 squares and assume any 6 will do.
- Reteach: "Mentally fold each one. The L shape leaves a face hanging open at the back."
- Re-check: Show a new pair (T-shape vs straight line of 6). Which folds?

TEACHER NOTES:
B is the cross-shape (the classic cube net). A has a corner overlap. C has a face hanging open.

WATCH FOR:
- Students who write B - secure.
- Students who write A or C - need the fold demonstration.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO = `SAY:
- With me. We will fold our kinder squares and find every line of symmetry.
- Step 1: fold left to right. Did the edges meet? Yes - that is one line.
- Step 2: open it. Fold top to bottom. Same? Yes - line two.
- Step 3: open it. Fold corner to corner (a diagonal). Same? Yes - line three.
- Step 4: open it. Fold the other diagonal. Yes - line four.
- A square has 4 lines of symmetry.

DO:
- Mirror with your own kinder square.
- Pause after each fold for the class to check.

TEACHER NOTES:
The square's 4 lines are: vertical, horizontal, two diagonals. Use the actual paper - this is a tactile lesson.

WATCH FOR:
- Students who feel for matching corners - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- A square has 4 lines of symmetry: two through the sides, two through the corners.

DO:
- Click to reveal.

TEACHER NOTES:
This is the most lines of symmetry of any of the common 4-sided shapes. Compare to rectangle (2) and parallelogram (0).

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CROSS_SECTION = `SAY:
- Year 6 idea.
- A right prism is a solid where the two ends are the same shape and the sides are rectangles.
- If you slice it parallel to the ends, every slice gives the SAME shape.
- That shape is called the cross-section.
- A triangular prism sliced anywhere along its length gives a triangle. A cube sliced parallel gives a square.

DO:
- Use a hand action - slicing.
- Show on a triangular prism (a Toblerone box works well).

TEACHER NOTES:
This is the Year 6 SP01 idea. Tie it to the net work: each "side" of the prism in its net unfolds to the rectangle of the side face; the end faces are the cross-section.

WATCH FOR:
- Students who can name the cross-section of a cube and a triangular prism - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_YOUDO = `SAY:
- Your turn. Take the Symmetry and Nets sheet.
- Part A - draw the lines of symmetry on each shape.
- Part B - match the net to the 3D solid.
- Extension - sketch the cross-section of a triangular prism.

DO:
- Distribute the sheet.
- Circulate. Have a kinder square ready to lend to students who want to check by folding.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Do Part A only. Use the kinder square to check each shape.
- Extra Notes: Encourage students to physically fold matching shapes.
EXTENDING PROMPT:
- Task: Find as many different cube nets as you can. Sketch them on grid paper. There are 11.
- Extra Notes: Use the language "valid" and "not valid".

TEACHER NOTES:
Different content from We Do. Same three core ideas.

WATCH FOR:
- Students who fold to check - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Q1: How many lines of symmetry does a rectangle have?
- Q2: Name the 3D solid that has six square faces.

DO:
- Display the prompt.
- 90 seconds.

TEACHER NOTES:
Exit ticket assesses SC2. Look for "2" and "cube".

WATCH FOR:
- Students who get both - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria for today.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: what is the difference between a square's and a rectangle's lines of symmetry?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea: a line of symmetry is a fold that makes the shape match itself exactly. A net is the open version of a 3D solid. Tomorrow we use transformations on tessellations.

[General: Closing | VTLM 2.0: Reflection]`;

// =====================================================================
// Helpers
// =====================================================================

// Draw a basic 2D shape (square / rectangle / equilateral triangle / regular hexagon).
function drawShapeAt(slide, kind, x, y, w, h, opts) {
  const o = opts || {};
  const color = o.color || C.PRIMARY;
  if (kind === "square") {
    slide.addShape("rect", {
      x, y, w, h, fill: { color: o.fill || C.BG_LIGHT },
      line: { color, width: 2.5 },
    });
  } else if (kind === "rectangle") {
    slide.addShape("rect", {
      x, y, w, h, fill: { color: o.fill || C.BG_LIGHT },
      line: { color, width: 2.5 },
    });
  } else if (kind === "triangle") {
    // Approximate equilateral triangle: use three lines.
    const cx = x + w / 2;
    const top = { x: cx, y };
    const bl  = { x, y: y + h };
    const br  = { x: x + w, y: y + h };
    [[top, bl], [bl, br], [br, top]].forEach(([a, b]) => {
      slide.addShape("line", {
        x: Math.min(a.x, b.x), y: Math.min(a.y, b.y),
        w: Math.abs(b.x - a.x), h: Math.abs(b.y - a.y),
        line: { color, width: 2.5 },
        flipV: (a.x < b.x) === (a.y < b.y) ? false : true,
      });
    });
  } else if (kind === "hexagon") {
    // Approximate regular hexagon with six points.
    const cx = x + w / 2;
    const cy = y + h / 2;
    const rx = w / 2;
    const ry = h / 2;
    const pts = [];
    for (let i = 0; i < 6; i++) {
      const ang = (Math.PI / 3) * i + Math.PI / 6;
      pts.push({ x: cx + rx * Math.cos(ang), y: cy + ry * Math.sin(ang) });
    }
    for (let i = 0; i < 6; i++) {
      const a = pts[i];
      const b = pts[(i + 1) % 6];
      slide.addShape("line", {
        x: Math.min(a.x, b.x), y: Math.min(a.y, b.y),
        w: Math.abs(b.x - a.x), h: Math.abs(b.y - a.y),
        line: { color, width: 2.5 },
        flipV: (a.x < b.x) === (a.y < b.y) ? false : true,
      });
    }
  }
}

// Draw lines of symmetry as dashed overlays on a centred shape.
function drawSymmetryLines(slide, x, y, w, h, lines, color) {
  // lines: array of "v","h","d1","d2"
  const cx = x + w / 2;
  const cy = y + h / 2;
  lines.forEach((kind) => {
    if (kind === "v") {
      slide.addShape("line", {
        x: cx, y: y - 0.05, w: 0, h: h + 0.10,
        line: { color, width: 1.5, dashType: "dash" },
      });
    } else if (kind === "h") {
      slide.addShape("line", {
        x: x - 0.05, y: cy, w: w + 0.10, h: 0,
        line: { color, width: 1.5, dashType: "dash" },
      });
    } else if (kind === "d1") {
      slide.addShape("line", {
        x: x - 0.05, y: y - 0.05, w: w + 0.10, h: h + 0.10,
        line: { color, width: 1.5, dashType: "dash" },
      });
    } else if (kind === "d2") {
      slide.addShape("line", {
        x: x - 0.05, y: y - 0.05, w: w + 0.10, h: h + 0.10,
        line: { color, width: 1.5, dashType: "dash" },
        flipV: true,
      });
    }
  });
}

// Draw a small cube-net pattern at (x,y) using cell size s.
// pattern is an array of [col,row] offsets for filled squares.
function drawNetPattern(slide, x, y, s, pattern, color) {
  pattern.forEach(([c, r]) => {
    slide.addShape("rect", {
      x: x + c * s, y: y + r * s, w: s, h: s,
      fill: { color: C.BG_LIGHT },
      line: { color, width: 2 },
    });
  });
}

// =====================================================================
// Build
// =====================================================================

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Symmetry and nets",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Solving equations",
      ["5 x n = 45", "n / 6 = 7", "3 x n - 4 = 20"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "n = 9         n = 42         n = 8", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal
  withReveal(
    () => fluencySlide(pres, "Fluency: Multiplication facts",
      ["8 x 7", "9 x 9", "12 x 4"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "56       81       48", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Launch", { color: C.PRIMARY });
    addTitle(s, "Fold your kinder square");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    // Left: Big visual - 3 squares with fold lines
    const startX = 1.2;
    const cardY = CONTENT_TOP + 0.30;
    const size = 1.5;
    const gap = 0.50;
    [
      { lines: ["v"], label: "Fold left to right" },
      { lines: ["v","h"], label: "Add top to bottom" },
      { lines: ["v","h","d1","d2"], label: "Add the diagonals" },
    ].forEach((item, i) => {
      const sx = startX + i * (size + gap + 0.60);
      const sy = cardY + 0.30;
      drawShapeAt(s, "square", sx, sy, size, size, { color: C.PRIMARY });
      drawSymmetryLines(s, sx, sy, size, size, item.lines, C.ALERT);
      s.addText(item.label, {
        x: sx - 0.30, y: sy + size + 0.15, w: size + 0.60, h: 0.45,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0, italic: true,
      });
    });

    s.addText("A square has 4 lines of symmetry.", {
      x: 0.7, y: SAFE_BOTTOM - 0.50, w: 8.6, h: 0.35,
      fontSize: 16, fontFace: FONT_H, color: C.PRIMARY,
      bold: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to identify lines of symmetry and match 3D solids to their nets.",
    [
      "I can find at least one line of symmetry on a simple 2D shape.",
      "I can match a net to its 3D solid for cubes and rectangular prisms.",
      "I can describe the cross-section you get when you slice a right prism parallel to its end.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key vocabulary (with simple visuals)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Symmetry, net, cross-section");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const items = [
      { word: "Line of symmetry", simple: "fold line", detail: "Where the shape folds onto itself.", color: C.PRIMARY },
      { word: "Net",              simple: "flat version", detail: "Folds up into a 3D solid.",       color: C.ACCENT },
      { word: "Cross-section",    simple: "slice", detail: "The shape you see when you cut a solid.", color: C.ALERT },
    ];
    const cellW = 2.7, gap = 0.20;
    const totalW = cellW * 3 + gap * 2;
    const startX = (10 - totalW) / 2;
    const cardY = CONTENT_TOP + 0.30;
    const cardH = 2.50;

    items.forEach((it, i) => {
      const cx = startX + i * (cellW + gap);
      addCard(s, cx, cardY, cellW, cardH, { strip: it.color, fill: C.WHITE });
      addTextOnShape(s, it.word, {
        x: cx + 0.15, y: cardY + 0.20, w: cellW - 0.30, h: 0.60, rectRadius: 0.10,
        fill: { color: it.color },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(it.simple, {
        x: cx, y: cardY + 0.95, w: cellW, h: 0.40,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", margin: 0,
      });
      s.addText(it.detail, {
        x: cx + 0.15, y: cardY + 1.50, w: cellW - 0.30, h: 0.85,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 10: I Do - Rectangle has 2 lines of symmetry
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "How many lines of symmetry?", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Three example shapes with their lines drawn
    const cardY = CONTENT_TOP + 0.25;
    const examples = [
      { kind: "square",    w: 1.5, h: 1.5, lines: ["v","h","d1","d2"], label: "Square: 4 lines", color: C.PRIMARY },
      { kind: "rectangle", w: 2.0, h: 1.2, lines: ["v","h"],           label: "Rectangle: 2 lines", color: C.ACCENT },
      { kind: "triangle",  w: 1.6, h: 1.5, lines: ["v"],               label: "Equilateral: 3 (we show 1)", color: C.ALERT },
    ];
    const startX = 1.2;
    const gap = 0.65;
    let cx = startX;
    examples.forEach((ex) => {
      const cy = cardY + 0.45;
      drawShapeAt(s, ex.kind, cx, cy, ex.w, ex.h, { color: ex.color });
      drawSymmetryLines(s, cx, cy, ex.w, ex.h, ex.lines, C.ALERT);
      s.addText(ex.label, {
        x: cx - 0.20, y: cy + ex.h + 0.25, w: ex.w + 0.40, h: 0.50,
        fontSize: 12, fontFace: FONT_H, color: ex.color,
        bold: true, align: "center", margin: 0,
      });
      cx += ex.w + gap;
    });

    s.addText("Test by folding: if the edges meet exactly, the fold is a line of symmetry.", {
      x: 0.7, y: SAFE_BOTTOM - 0.45, w: 8.6, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_SYMMETRY);
  })();

  // Slide 11: I Do - Net of a cube
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "A net of a cube has 6 squares", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Left: cube (use a simple isometric drawing with 3 rhombi-as-squares)
    const cubeX = 1.5;
    const cubeY = CONTENT_TOP + 0.60;
    const cubeSize = 1.2;
    // front face
    slide_face(s, cubeX, cubeY + cubeSize * 0.30, cubeSize, cubeSize, C.PRIMARY, C.WHITE);
    // top face (skewed approximation)
    slide_face(s, cubeX + cubeSize * 0.30, cubeY, cubeSize, cubeSize * 0.30, C.PRIMARY, C.BG_LIGHT);
    // right face
    slide_face(s, cubeX + cubeSize, cubeY + cubeSize * 0.30, cubeSize * 0.30, cubeSize, C.PRIMARY, C.BG_LIGHT);
    s.addText("Cube", {
      x: cubeX - 0.30, y: cubeY + cubeSize + 0.45, w: cubeSize + 0.60, h: 0.40,
      fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });

    // Right: cross-shaped net
    const netX = 5.5;
    const netY = CONTENT_TOP + 0.40;
    const netS = 0.55;
    // Pattern: cross
    const cross = [[1,0],[0,1],[1,1],[2,1],[3,1],[1,2]];
    drawNetPattern(s, netX, netY, netS, cross, C.SECONDARY);
    s.addText("Net of a cube", {
      x: netX - 0.30, y: netY + 3 * netS + 0.20, w: 4 * netS + 0.60, h: 0.40,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    // Arrow between
    s.addShape("line", {
      x: cubeX + cubeSize + 0.7, y: cubeY + cubeSize * 0.65,
      w: netX - (cubeX + cubeSize + 0.7) - 0.10, h: 0,
      line: { color: C.ACCENT, width: 2.5, endArrowType: "triangle" },
    });
    s.addText("unfold", {
      x: cubeX + cubeSize + 0.70, y: cubeY + cubeSize * 0.65 - 0.40,
      w: netX - (cubeX + cubeSize + 0.7), h: 0.30,
      fontSize: 11, fontFace: FONT_B, color: C.ACCENT, italic: true,
      align: "center", margin: 0,
    });

    s.addText("6 square faces. Fold along the edges to make the cube.", {
      x: 0.7, y: SAFE_BOTTOM - 0.45, w: 8.6, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_NET);
  })();

  // Slides 12-13: CFU + reveal — Which is a valid cube net?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Which one is a valid cube net?", { color: C.ALERT });
      addTextOnShape(s, "CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      const netY = CONTENT_TOP + 0.40;
      const netS = 0.40;
      // Option A: L shape (invalid - overlap)
      const xa = 1.3;
      [[0,0],[0,1],[1,1],[2,1],[0,2],[0,3]].forEach(([c,r]) => {
        s.addShape("rect", {
          x: xa + c * netS, y: netY + r * netS, w: netS, h: netS,
          fill: { color: C.BG_LIGHT }, line: { color: C.MUTED, width: 1.5 },
        });
      });
      s.addText("A", {
        x: xa, y: netY + 4 * netS + 0.10, w: 3 * netS, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });

      // Option B: cross (valid)
      const xb = 4.2;
      [[1,0],[0,1],[1,1],[2,1],[3,1],[1,2]].forEach(([c,r]) => {
        s.addShape("rect", {
          x: xb + c * netS, y: netY + r * netS, w: netS, h: netS,
          fill: { color: C.BG_LIGHT }, line: { color: C.MUTED, width: 1.5 },
        });
      });
      s.addText("B", {
        x: xb, y: netY + 3 * netS + 0.10, w: 4 * netS, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });

      // Option C: straight line of 6 (invalid)
      const xc = 7.0;
      [[0,0],[1,0],[2,0],[3,0],[4,0],[5,0]].forEach(([c,r]) => {
        s.addShape("rect", {
          x: xc + c * netS * 0.55, y: netY + r * netS, w: netS * 0.55, h: netS,
          fill: { color: C.BG_LIGHT }, line: { color: C.MUTED, width: 1.5 },
        });
      });
      s.addText("C", {
        x: xc, y: netY + netS + 0.10, w: 6 * netS * 0.55, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });

      s.addText("On your boards: write A, B or C.", {
        x: 0.7, y: SAFE_BOTTOM - 0.45, w: 8.6, h: 0.30,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "B is correct. The cross shape folds into a cube.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do — Find all lines of symmetry on a square (using paper folding language)
  withReveal(
    () => {
      const s = pres.addSlide();
      const stageColor = STAGE_COLORS["3"];
      addTopBar(s, stageColor);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "How many lines of symmetry does a square have?", { color: stageColor });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

      // Big blank square in the middle
      const sq = 2.4;
      const sx = 5 - sq / 2;
      const sy = CONTENT_TOP + 0.40;
      drawShapeAt(s, "square", sx, sy, sq, sq, { color: C.SECONDARY });

      s.addText("Fold your kinder square. Find every fold where the edges match.", {
        x: 0.7, y: SAFE_BOTTOM - 0.45, w: 8.6, h: 0.30,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      const sq = 2.4;
      const sx = 5 - sq / 2;
      const sy = CONTENT_TOP + 0.40;
      drawSymmetryLines(slide, sx, sy, sq, sq, ["v","h","d1","d2"], C.ALERT);
      addTextOnShape(slide, "A square has 4 lines of symmetry.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // Slide 16: Year 6 stretch - Cross sections of a right prism
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Cross-sections of right prisms (Year 6)", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    const cardY = CONTENT_TOP + 0.30;
    // Triangular prism
    const px1 = 1.4;
    slide_face(s, px1, cardY + 0.5, 1.7, 1.0, C.PRIMARY, C.BG_LIGHT);
    s.addShape("line", {
      x: px1, y: cardY + 0.5, w: 0.5, h: -0.4,
      line: { color: C.PRIMARY, width: 2 },
    });
    s.addShape("line", {
      x: px1 + 1.7, y: cardY + 0.5, w: 0.5, h: -0.4,
      line: { color: C.PRIMARY, width: 2 },
    });
    s.addShape("line", {
      x: px1 + 0.85, y: cardY + 1.5, w: 0.5, h: -0.4,
      line: { color: C.PRIMARY, width: 2 },
    });
    s.addShape("line", {
      x: px1 + 0.5, y: cardY + 0.1, w: 1.7, h: 0,
      line: { color: C.PRIMARY, width: 2 },
    });
    s.addText("Triangular prism\nslice = triangle", {
      x: px1 - 0.20, y: cardY + 1.85, w: 2.10, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "top", margin: 0,
    });

    // Cube
    const px2 = 4.2;
    slide_face(s, px2, cardY + 0.5, 1.4, 1.0, C.ACCENT, C.BG_LIGHT);
    s.addShape("line", {
      x: px2, y: cardY + 0.5, w: 0.5, h: -0.4,
      line: { color: C.ACCENT, width: 2 },
    });
    s.addShape("line", {
      x: px2 + 1.4, y: cardY + 0.5, w: 0.5, h: -0.4,
      line: { color: C.ACCENT, width: 2 },
    });
    s.addShape("line", {
      x: px2 + 1.4, y: cardY + 1.5, w: 0.5, h: -0.4,
      line: { color: C.ACCENT, width: 2 },
    });
    s.addShape("line", {
      x: px2 + 0.5, y: cardY + 0.1, w: 1.4, h: 0,
      line: { color: C.ACCENT, width: 2 },
    });
    s.addText("Cube\nslice = square", {
      x: px2 - 0.20, y: cardY + 1.85, w: 1.80, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "top", margin: 0,
    });

    // Hexagonal prism
    const px3 = 6.8;
    drawShapeAt(s, "hexagon", px3, cardY + 0.5, 1.5, 1.0, { color: C.ALERT });
    s.addText("Hexagonal prism\nslice = hexagon", {
      x: px3 - 0.20, y: cardY + 1.85, w: 1.90, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "top", margin: 0,
    });

    s.addText("Slice parallel to the end. The cross-section matches the end face.", {
      x: 0.7, y: SAFE_BOTTOM - 0.45, w: 8.6, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CROSS_SECTION);
  })();

  // Slide 17: You Do
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["4"];
    addTopBar(s, stageColor);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: lines of symmetry and nets", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    s.addText([
      { text: "First: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "draw the lines of symmetry on each shape (Part A).", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Next: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "match the net to its 3D solid (Part B).", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Then: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "sketch the cross-section of a triangular prism.", options: { fontSize: 15, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    // Right side preview
    const px = 5.30;
    const py = CONTENT_TOP + 0.20;
    const sw = 4.10;
    addCard(s, px, py, sw, 3.10, { strip: stageColor, fill: C.WHITE });
    // 4 quick shapes
    drawShapeAt(s, "square",    px + 0.30, py + 0.50, 1.0, 1.0, { color: C.PRIMARY });
    drawShapeAt(s, "rectangle", px + 1.70, py + 0.50, 1.2, 0.8, { color: C.PRIMARY });
    drawShapeAt(s, "triangle",  px + 0.40, py + 1.80, 1.0, 0.9, { color: C.ACCENT });
    drawShapeAt(s, "hexagon",   px + 1.80, py + 1.80, 1.2, 1.0, { color: C.ALERT });

    s.addText("Use the Symmetry and Nets sheet.", {
      x: 5.30, y: SAFE_BOTTOM - 0.40, w: 4.10, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.MUTED,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 18: Exit Ticket
  exitTicketSlide(pres,
    [
      "How many lines of symmetry does a rectangle have?",
      "Name the 3D solid that has six square faces.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 19: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the difference between a square's and a rectangle's lines of symmetry?",
      scItems: [
        "I can find at least one line of symmetry on a simple 2D shape.",
        "I can match a net to its 3D solid for cubes and rectangular prisms.",
        "I can describe the cross-section you get when you slice a right prism parallel to its end.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "GD_Lesson2_Symmetry_and_Nets.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Draw lines of symmetry. Match each net to a 3D solid.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "A line of symmetry is a fold where the shape matches itself exactly. " +
      "A net is the flat version of a 3D solid. Use your kinder square to check by folding.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Part A — Lines of symmetry", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Draw every line of symmetry on each shape. Write how many you found.", y);
    y = addWriteLine(doc, "a) Square:           number of lines = ___", y);
    y = addWriteLine(doc, "b) Rectangle:        number of lines = ___", y);
    y = addWriteLine(doc, "c) Equilateral tri:  number of lines = ___", y);
    y = addWriteLine(doc, "d) Regular hexagon:  number of lines = ___", y);
    y = addWriteLine(doc, "e) Parallelogram:    number of lines = ___", y);

    y = addSectionHeading(doc, "Part B — Match the net to the solid", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Draw a line from each net to its 3D solid.", y);
    y = addBodyText(doc, "Net 1: a cross made of 6 squares.", y);
    y = addBodyText(doc, "Net 2: 2 triangles + 3 rectangles in a row.", y);
    y = addBodyText(doc, "Net 3: 2 circles and a long rectangle.", y);
    y = addBodyText(doc, "Solids: cube, triangular prism, cylinder.", y);
    y = addWriteLine(doc, "Net 1 -> ____________", y);
    y = addWriteLine(doc, "Net 2 -> ____________", y);
    y = addWriteLine(doc, "Net 3 -> ____________", y);

    y = addSectionHeading(doc, "Part C — Cross-section (extension)", y, { color: C.ACCENT });
    y = addBodyText(doc, "A triangular prism is sliced parallel to its triangular end. Sketch the cross-section.", y);
    y = addWriteLine(doc, "Cross-section shape: ____________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Symmetry and Nets | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Symmetry and Nets sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Part A — Lines of symmetry", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a) Square = 4 lines (two through sides, two through corners).", y);
    y = addBodyText(doc, "b) Rectangle = 2 lines (horizontal and vertical only). NOT 4 - the diagonals do not fold cleanly.", y);
    y = addBodyText(doc, "c) Equilateral triangle = 3 lines (one from each corner to the opposite midpoint).", y);
    y = addBodyText(doc, "d) Regular hexagon = 6 lines (three through opposite corners, three through opposite midpoints).", y);
    y = addBodyText(doc, "e) Parallelogram (non-rectangle) = 0 lines.", y);

    y = addSectionHeading(doc, "Part B — Net to solid", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Net 1 -> cube.", y);
    y = addBodyText(doc, "Net 2 -> triangular prism.", y);
    y = addBodyText(doc, "Net 3 -> cylinder.", y);

    y = addSectionHeading(doc, "Part C — Cross-section", y, { color: C.ACCENT });
    y = addBodyText(doc, "The cross-section of a triangular prism (sliced parallel to the triangular end) is a triangle.", y);
    y = addBodyText(doc, "The cross-section matches the end face of any right prism when sliced parallel.", y);

    y = addTipBox(doc,
      "Watch for: students who say a rectangle has 4 lines of symmetry (it has 2); students who confuse the net of a triangular prism with a pyramid; students who slice a prism perpendicular to its length and get a rectangle, not the end shape.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

// Tiny helper: a single filled rectangle face (kept lower-case so the cube call sites read naturally).
function slide_face(slide, x, y, w, h, lineColor, fillColor) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: fillColor },
    line: { color: lineColor, width: 2 },
  });
}

build().catch((err) => { console.error(err); process.exit(1); });
