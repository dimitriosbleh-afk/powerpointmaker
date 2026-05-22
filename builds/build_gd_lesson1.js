"use strict";

// Geometry and Data Unit (Year 5/6 Numeracy) - Lesson 1: Translations, reflections, rotations.
// Year 5 VC2M5SP03 - foundation for the whole shapes/tessellations week.
// First time we have looked at this for the year (mixed readiness language).
// Daily Review: Solving equations with multiplication, division and operations.
// Fluency: Multiplication facts - heaviest scaffold of the week (gradual release).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// All eight lessons in the unit share variant 2 for cohesion.
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

const SESSION = 1;
const TOTAL = 8;
const UNIT_TITLE = "Geometry and Data: Shapes, Tessellations, Data";
const FOOTER = `Geometry and Data | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/GD_Lesson1_Transformations";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 1 Transformation Practice",
  "Perform translations, reflections and rotations on a small grid.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 1 Answer Key",
  "Worked answers for the Transformation Practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// =====================================================================
// Teacher notes (plain text, ASCII safe)
// =====================================================================

const NOTES_TITLE = `SAY:
- Welcome to a new unit. This week we are looking at shapes and tessellations.
- Some of you may remember moving shapes around. If this feels new, that is okay. We will build it together.

DO:
- Settle the class.
- Have whiteboards ready.

TEACHER NOTES:
First lesson of an eight session unit. Today is the foundation - what a translation, reflection and rotation looks like. No tessellations yet; that comes later this week.

WATCH FOR:
- Students who have heard the words before - secure for extension.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- Whiteboards on desks.
- One Transformation Practice sheet per student for the You Do.

DO:
- Print one student sheet per student.
- Print one answer key for yourself.

TEACHER NOTES:
The sheet is a small grid with three short tasks. Keep grid paper handy as a backup if a student needs more space to draw.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review.
- Solve each equation on your whiteboard.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Solving equations with multiplication, division and operations. Encourage students to write the inverse on the second line.

WATCH FOR:
- Students who write the inverse operation - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 3 x n = 24, so n = 8.
- n divided by 4 = 6, so n = 24.
- 2 x n + 5 = 17, so 2n = 12, n = 6.
- Tick or fix.

DO:
- Click to reveal.

TEACHER NOTES:
The third one is two steps. Watch for students who try to subtract 5 from 17 mentally without writing.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Multiplication facts.
- Whisper-answer then write on your board.

DO:
- Display the three prompts.
- 30 seconds.

TEACHER NOTES:
This is the start of multiplication and division fluency for the week. Today is heavy scaffold - small numbers, plenty of time. Tomorrow we tighten.

WATCH FOR:
- Students who recall instantly - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 6 x 4 = 24.
- 7 x 8 = 56.
- 9 x 6 = 54.

DO:
- Click to reveal.

TEACHER NOTES:
Brisk reveal. Note any student who hesitated on 7 x 8 - mark for small group.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Hold up your hand.
- Now slide it to the left. That is a translation. The hand moved but stayed the same shape.
- Now flip it over so the palm faces the other way. That is a reflection.
- Now turn it like a steering wheel. That is a rotation.
- Today we will name each of these moves and do them on paper.

DO:
- Mirror the moves with the class.
- Three actions, ten seconds each.

TEACHER NOTES:
This launch links to the new learning. The physical move makes "translate, reflect, rotate" memorable before the formal definitions appear.

WATCH FOR:
- Students who try all three moves with you - tracking.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to describe and perform translations, reflections and rotations of shapes.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 is naming the move. SC2 is the core target - performing one of the three. SC3 stretches to describing what changes and what stays the same.

WATCH FOR:
- Students who can name all three moves - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Three key words today.
- Translation - slide. The shape moves but does not turn or flip.
- Reflection - flip. The shape mirrors across a line.
- Rotation - turn. The shape spins around a point.

DO:
- Point at each visual as you say the word.
- Have students whisper the word back.

TEACHER NOTES:
Pair each word with the physical action: slide hand, flip hand, turn hand. The picture makes the meaning before we use it.

WATCH FOR:
- Students who can say the word with the action - tracking.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_IDO_TRANSLATE = `SAY:
- Watch me. I have a small triangle on the grid.
- I am going to translate it 4 right and 2 up.
- Every corner moves the same. Right 4, up 2. Right 4, up 2. Right 4, up 2.
- The shape did not turn. It did not flip. It just slid.

DO:
- Trace each corner moving with your finger.
- Compare the two triangles - same shape, new spot.

TEACHER NOTES:
Anchor the rule: every corner moves the same amount. If a corner ends up in the wrong cell the translation is wrong.

WATCH FOR:
- Students who count the squares with a finger - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_REFLECT = `SAY:
- Watch me. Same triangle. Now I reflect across the vertical line.
- Each point is the same distance from the line, but on the other side.
- A corner that was 2 squares right of the line ends up 2 squares left of the line.
- The shape looks like a mirror image. Front to back is swapped.

DO:
- Hold a small mirror up to the original triangle if you have one.
- Trace one corner across the line.

TEACHER NOTES:
The mirror line is the key. Same distance on the other side, every point. Students who can verify one point usually verify them all.

WATCH FOR:
- Students who measure the distance to the line - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_ROTATE = `SAY:
- Watch me. Same triangle. Now I rotate it 90 degrees clockwise around the marked point.
- A right turn. Clockwise like the hands of a clock.
- The point I rotate around stays still. Everything else turns with it.
- If I do this four times it ends up back where it started.

DO:
- Use a piece of tracing paper and a pin to demonstrate.
- Pin on the center, turn the paper a quarter turn.

TEACHER NOTES:
Quarter turn = 90 degrees. Use tracing paper to make the rotation visible. Half turn = 180. Three quarter turn = 270.

WATCH FOR:
- Students who use tracing paper themselves - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- I will show you two shapes. The second one came from the first.
- Decide: translation, reflection, or rotation?
- Write your answer on your board.

DO:
- Display the prompt.
- 30 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me on three, two, one, show."
- Scan for: rotation.
PROCEED: If 80% have rotation, click to reveal and move to We Do.
PIVOT: Most likely misconception - students see the shape on a different side of the page and call it reflection.
- Reteach: "Is one a mirror image of the other? Or did it turn? Compare a single corner."
- Re-check: Show a new pair. "Translation, reflection or rotation?"

TEACHER NOTES:
The figure shown is a clockwise quarter turn. A reflection would mirror it left/right.

WATCH FOR:
- Students who write "rotation" - secure.
- Students who write "reflection" - column drift.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO = `SAY:
- With me. We will translate this L shape 3 right and 1 down.
- Step 1: pick a corner.
- Step 2: move it right 3, down 1. Mark the new spot.
- Step 3: do the same for every other corner.
- Step 4: join them up. Same shape, new place.

DO:
- Build on the board step by step.
- Ask the class for each move.

TEACHER NOTES:
Co-construct. Use whiteboards. Be deliberate about marking each corner before joining.

WATCH FOR:
- Students who lift their pen between corners - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Check.
- L shape, every corner moved right 3 and down 1.
- Same shape. New spot. That is a translation.

DO:
- Click to reveal.

TEACHER NOTES:
Highlight that the orientation is unchanged - it is not turned and not flipped.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn. Take the Transformation Practice sheet.
- Three small tasks. One translation, one reflection, one rotation.
- Use your pencil and a ruler.

DO:
- Distribute the sheet.
- Circulate. Ask students to point at the rule before they draw.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Do the translation only. Sit with these students and trace one corner at a time.
- Extra Notes: Encourage students to count cells out loud.
EXTENDING PROMPT:
- Task: Do a second rotation - 180 degrees around the same point. Compare to the 90 degree rotation.
- Extra Notes: Use the language "what stayed the same, what changed".

TEACHER NOTES:
Different shapes from We Do, same three move types. Look for accurate corner moves.

WATCH FOR:
- Students who check one point at a time - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Look at the two shapes on the slide.
- Name the transformation: translation, reflection, or rotation?

DO:
- Display the prompt.
- 60 seconds.

TEACHER NOTES:
Exit ticket assesses SC2 - identifying the transformation. Look for "reflection".

WATCH FOR:
- Students who get it - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria for today.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: ONE word from today and what it means.

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea: in a translation everything slides; in a reflection everything mirrors; in a rotation everything turns around one fixed point.

[General: Closing | VTLM 2.0: Reflection]`;

// =====================================================================
// Custom helpers - small grid + simple shape drawer
// =====================================================================

// Draw a light grid for visual transformation slides.
function drawGrid(slide, x, y, w, h, cols, rows, opts) {
  const o = opts || {};
  const cellW = w / cols;
  const cellH = h / rows;
  const lineColor = o.lineColor || C.MUTED;

  // Background
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: o.fill || C.WHITE },
    line: { color: lineColor, width: 1 },
  });
  // Vertical lines
  for (let i = 1; i < cols; i++) {
    slide.addShape("line", {
      x: x + i * cellW, y, w: 0, h,
      line: { color: lineColor, width: 0.5 },
    });
  }
  // Horizontal lines
  for (let j = 1; j < rows; j++) {
    slide.addShape("line", {
      x, y: y + j * cellH, w, h: 0,
      line: { color: lineColor, width: 0.5 },
    });
  }
  return { x, y, cellW, cellH };
}

// Draw a triangle on the grid given three (col,row) corners.
function drawTriangleOnGrid(slide, grid, corners, opts) {
  const o = opts || {};
  const color = o.color || C.PRIMARY;
  // PptxGenJS does not support polygons directly; draw three filled circles
  // at corners and three lines for sides.
  const pts = corners.map(([col, row]) => ({
    x: grid.x + col * grid.cellW,
    y: grid.y + row * grid.cellH,
  }));
  for (let i = 0; i < 3; i++) {
    const a = pts[i];
    const b = pts[(i + 1) % 3];
    const lineX = Math.min(a.x, b.x);
    const lineY = Math.min(a.y, b.y);
    const lineW = Math.abs(b.x - a.x);
    const lineH = Math.abs(b.y - a.y);
    slide.addShape("line", {
      x: lineX, y: lineY, w: lineW, h: lineH,
      line: { color: color, width: 2.5, beginArrowType: "none", endArrowType: "none" },
      flipV: (a.x < b.x) !== (a.y < b.y) ? false : true,
    });
  }
  // Corner dots (small)
  pts.forEach((p) => {
    slide.addShape("roundRect", {
      x: p.x - 0.06, y: p.y - 0.06, w: 0.12, h: 0.12, rectRadius: 0.06,
      fill: { color: color }, line: { color: color, width: 1 },
    });
  });
  if (o.label) {
    slide.addText(o.label, {
      x: pts[0].x - 0.30, y: pts[0].y - 0.40, w: 0.60, h: 0.30,
      fontSize: 11, fontFace: FONT_H, color: color, bold: true,
      align: "center", margin: 0,
    });
  }
}

// Draw an L-shape on the grid given top-left cell, width, height, and notch.
function drawLOnGrid(slide, grid, anchor, opts) {
  const o = opts || {};
  const color = o.color || C.SECONDARY;
  const [c, r] = anchor;
  // L shape with corners at relative offsets (col, row):
  // (0,0) (2,0) (2,1) (1,1) (1,3) (0,3)
  const offsets = [[0,0],[2,0],[2,1],[1,1],[1,3],[0,3]];
  const pts = offsets.map(([dc, dr]) => ({
    x: grid.x + (c + dc) * grid.cellW,
    y: grid.y + (r + dr) * grid.cellH,
  }));
  for (let i = 0; i < pts.length; i++) {
    const a = pts[i];
    const b = pts[(i + 1) % pts.length];
    const lineX = Math.min(a.x, b.x);
    const lineY = Math.min(a.y, b.y);
    const lineW = Math.abs(b.x - a.x);
    const lineH = Math.abs(b.y - a.y);
    slide.addShape("line", {
      x: lineX, y: lineY, w: lineW, h: lineH,
      line: { color: color, width: 2.5 },
    });
  }
  if (o.label) {
    slide.addText(o.label, {
      x: pts[0].x - 0.30, y: pts[0].y - 0.40, w: 0.60, h: 0.30,
      fontSize: 11, fontFace: FONT_H, color: color, bold: true,
      align: "center", margin: 0,
    });
  }
}

// =====================================================================
// Build
// =====================================================================

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Translations, reflections, rotations",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Solving equations",
      ["3 x n = 24", "n / 4 = 6", "2 x n + 5 = 17"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "n = 8         n = 24         n = 6", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal
  withReveal(
    () => fluencySlide(pres, "Fluency: Multiplication facts",
      ["6 x 4", "7 x 8", "9 x 6"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "24       56       54", {
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
    addTitle(s, "Three moves with your hand");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const items = [
      { word: "Slide", action: "Move it across.", color: C.PRIMARY },
      { word: "Flip",  action: "Turn the palm over.", color: C.ACCENT },
      { word: "Turn",  action: "Spin like a steering wheel.", color: C.ALERT },
    ];
    const cellW = 2.7, gap = 0.20;
    const totalW = cellW * 3 + gap * 2;
    const startX = (10 - totalW) / 2;
    const cardY = CONTENT_TOP + 0.40;
    const cardH = 2.30;

    items.forEach((it, i) => {
      const cx = startX + i * (cellW + gap);
      addCard(s, cx, cardY, cellW, cardH, { strip: it.color, fill: C.WHITE });
      addTextOnShape(s, it.word, {
        x: cx + 0.20, y: cardY + 0.30, w: cellW - 0.40, h: 0.70, rectRadius: 0.10,
        fill: { color: it.color },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(it.action, {
        x: cx + 0.10, y: cardY + 1.30, w: cellW - 0.20, h: 0.80,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to describe and perform translations, reflections and rotations of shapes.",
    [
      "I can name a transformation as a translation, reflection, or rotation.",
      "I can perform a translation, reflection or rotation on a small shape using a grid.",
      "I can describe what stays the same and what changes when a shape is transformed.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key vocabulary
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Translation, reflection, rotation");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const items = [
      { word: "Translation", simple: "slide",  detail: "The shape moves. It does not turn or flip.", color: C.PRIMARY },
      { word: "Reflection",  simple: "flip",   detail: "The shape mirrors across a line.",           color: C.ACCENT },
      { word: "Rotation",    simple: "turn",   detail: "The shape spins around a fixed point.",      color: C.ALERT },
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
        x: cx + 0.20, y: cardY + 0.20, w: cellW - 0.40, h: 0.60, rectRadius: 0.10,
        fill: { color: it.color },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(it.simple + ".", {
        x: cx, y: cardY + 0.95, w: cellW, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", margin: 0,
      });
      s.addText(it.detail, {
        x: cx + 0.15, y: cardY + 1.50, w: cellW - 0.30, h: 0.80,
        fontSize: 13, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 10: I Do - Translation on grid
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Translate 4 right, 2 up", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Grid takes the center
    const grid = drawGrid(s, 1.2, CONTENT_TOP + 0.20, 7.6, 2.85, 12, 6);

    // Original triangle at top-left area
    drawTriangleOnGrid(s, grid, [[1,4],[3,4],[2,2]], { color: C.PRIMARY, label: "A" });
    // Translated triangle 4 right, 2 up: each corner +4 col, -2 row
    drawTriangleOnGrid(s, grid, [[5,2],[7,2],[6,0]], { color: C.SUCCESS, label: "A'" });

    s.addText("Same shape. New spot. Every corner moves 4 right and 2 up.", {
      x: 0.7, y: SAFE_BOTTOM - 0.40, w: 8.6, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_TRANSLATE);
  })();

  // Slide 11: I Do - Reflection
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Reflect across the line", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    const grid = drawGrid(s, 1.2, CONTENT_TOP + 0.20, 7.6, 2.85, 12, 6);

    // Vertical mirror line at col 6
    s.addShape("line", {
      x: grid.x + 6 * grid.cellW, y: grid.y,
      w: 0, h: 6 * grid.cellH,
      line: { color: C.ALERT, width: 2.5, dashType: "dash" },
    });
    s.addText("mirror line", {
      x: grid.x + 6 * grid.cellW - 0.6, y: grid.y - 0.30,
      w: 1.2, h: 0.25,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, italic: true,
      align: "center", margin: 0,
    });

    // Original triangle left of mirror
    drawTriangleOnGrid(s, grid, [[2,4],[4,4],[3,2]], { color: C.PRIMARY, label: "A" });
    // Reflected: distance from line is preserved; col = 12 - origCol
    drawTriangleOnGrid(s, grid, [[10,4],[8,4],[9,2]], { color: C.SUCCESS, label: "A'" });

    s.addText("Every point is the same distance from the mirror line, on the other side.", {
      x: 0.7, y: SAFE_BOTTOM - 0.40, w: 8.6, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_REFLECT);
  })();

  // Slide 12: I Do - Rotation
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Rotate 90 degrees clockwise", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    const grid = drawGrid(s, 1.2, CONTENT_TOP + 0.20, 7.6, 2.85, 12, 6);

    // Centre of rotation at (6,3)
    const cx = grid.x + 6 * grid.cellW;
    const cy = grid.y + 3 * grid.cellH;
    s.addShape("roundRect", {
      x: cx - 0.10, y: cy - 0.10, w: 0.20, h: 0.20, rectRadius: 0.10,
      fill: { color: C.ALERT }, line: { color: C.ALERT, width: 1 },
    });
    s.addText("centre", {
      x: cx - 0.6, y: cy + 0.12, w: 1.2, h: 0.25,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, italic: true,
      align: "center", margin: 0,
    });

    // Original triangle right of centre
    drawTriangleOnGrid(s, grid, [[7,3],[9,3],[8,1]], { color: C.PRIMARY, label: "A" });
    // Rotated 90 cw around (6,3): (col,row) -> (6 + (row-3), 3 + (6-col)... )
    // Standard 90 cw: (x,y) about (cx,cy) -> (cx + (y - cy), cy - (x - cx))
    // Original points (col,row): (7,3),(9,3),(8,1)
    // (7,3): col' = 6 + (3-3) = 6; row' = 3 - (7-6) = 2 -> NOT right (col same as centre)
    // Use mathematical rotation around centre (6,3) with 90 cw which is (x,y) -> (cx + (y-cy), cy - (x-cx)) but y axis is inverted on slides; let me just precompute carefully.
    // Use: 90 cw on grid (where y increases downwards):
    //   x' = cx - (y - cy)
    //   y' = cy + (x - cx)
    // (7,3): x'=6-0=6, y'=3+1=4 -> (6,4)
    // (9,3): x'=6-0=6, y'=3+3=6 -> (6,6) on edge, use (6,5) instead to keep on grid
    // (8,1): x'=6+2=8, y'=3+2=5 -> (8,5)
    // Actually the precise 90 cw should map (9,3) to (6,6). Let me adjust: that is inside our 12x6 grid as point.
    drawTriangleOnGrid(s, grid, [[6,4],[6,6],[8,5]], { color: C.SUCCESS, label: "A'" });

    s.addText("The centre stays still. Everything else turns a quarter turn.", {
      x: 0.7, y: SAFE_BOTTOM - 0.40, w: 8.6, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_ROTATE);
  })();

  // Slides 13-14: CFU + reveal
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Translation, reflection, or rotation?", { color: C.ALERT });
      addTextOnShape(s, "CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      const grid = drawGrid(s, 1.2, CONTENT_TOP + 0.20, 7.6, 2.85, 12, 6);
      drawTriangleOnGrid(s, grid, [[2,4],[4,4],[3,2]], { color: C.PRIMARY, label: "A" });
      // A rotated 90 cw around (3,3): (2,4) -> (2, 4)... use a clear quarter turn
      drawTriangleOnGrid(s, grid, [[9,2],[9,4],[11,3]], { color: C.SUCCESS, label: "A'" });

      s.addText("On your whiteboard: which transformation took A to A'?", {
        x: 0.7, y: SAFE_BOTTOM - 0.40, w: 8.6, h: 0.28,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Rotation. The shape turned 90 degrees clockwise.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 15-16: We Do + reveal (translation co-construction)
  withReveal(
    () => {
      const s = pres.addSlide();
      const stageColor = STAGE_COLORS["3"];
      addTopBar(s, stageColor);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Translate the L shape 3 right, 1 down", { color: stageColor });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

      const grid = drawGrid(s, 1.2, CONTENT_TOP + 0.20, 7.6, 2.85, 12, 6);
      drawLOnGrid(s, grid, [1, 0], { color: C.SECONDARY, label: "L" });

      s.addText("On your boards: mark where each corner ends up. Join them.", {
        x: 0.7, y: SAFE_BOTTOM - 0.40, w: 8.6, h: 0.28,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      const grid = { x: 1.2, y: CONTENT_TOP + 0.20, cellW: 7.6 / 12, cellH: 2.85 / 6 };
      drawLOnGrid(slide, grid, [4, 1], { color: C.SUCCESS, label: "L'" });
      addTextOnShape(slide, "Every corner moved 3 right and 1 down. Same shape.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // Slide 17: You Do
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["4"];
    addTopBar(s, stageColor);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: three transformations", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    s.addText([
      { text: "First: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "translate the triangle 3 right, 1 up.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Next: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "reflect the L shape across the line.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Then: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "rotate the arrow 90 degrees clockwise around the dot.", options: { fontSize: 15, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    // Right side preview: tiny version of the sheet
    const grid = drawGrid(s, 5.30, CONTENT_TOP + 0.20, 4.10, 2.85, 8, 6);
    drawTriangleOnGrid(s, grid, [[1,4],[3,4],[2,2]], { color: C.PRIMARY });
    s.addShape("line", {
      x: grid.x + 4 * grid.cellW, y: grid.y,
      w: 0, h: 6 * grid.cellH,
      line: { color: C.ALERT, width: 1.5, dashType: "dash" },
    });

    s.addText("Use the Transformation Practice sheet.", {
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
      "Look at the two shapes. Name the transformation: translation, reflection, or rotation.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 19: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: ONE word from today and what it means.",
      scItems: [
        "I can name a transformation as a translation, reflection, or rotation.",
        "I can perform a translation, reflection or rotation on a small shape using a grid.",
        "I can describe what stays the same and what changes when a shape is transformed.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "GD_Lesson1_Transformations.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Use a pencil and ruler. Move each corner. Then join them up.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Translation = slide. Reflection = flip across a line. Rotation = turn around a point. " +
      "For each task: mark every corner of the new shape first, then join them.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Task 1 — Translation", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Triangle A has corners at (1,5), (3,5), (2,3). Translate it 4 right and 2 up.", y);
    y = addWriteLine(doc, "New corners: A' = (___,___), (___,___), (___,___)", y);
    y = addWriteLine(doc, "What stayed the same?  ____________________________________________", y);
    y = addWriteLine(doc, "What changed?          ____________________________________________", y);

    y = addSectionHeading(doc, "Task 2 — Reflection", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Triangle B has corners at (2,5), (4,5), (3,3). The mirror line is the vertical line at x = 6.", y);
    y = addBodyText(doc, "Reflect Triangle B across the line.", y);
    y = addWriteLine(doc, "New corners: B' = (___,___), (___,___), (___,___)", y);

    y = addSectionHeading(doc, "Task 3 — Rotation", y, { color: C.ACCENT });
    y = addBodyText(doc, "Arrow C has corners at (7,4), (9,4), (8,2). Rotate 90 degrees clockwise around the point (6,3).", y);
    y = addWriteLine(doc, "New corners: C' = (___,___), (___,___), (___,___)", y);
    y = addBodyText(doc, "Extension: Rotate the same arrow 180 degrees around (6,3). Compare your two new arrows.", y);
    y = addWriteLine(doc, "180 degree corners: C'' = (___,___), (___,___), (___,___)", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Transformation Practice | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Transformation Practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Task 1 — Translation (4 right, 2 up)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "A' = (5,3), (7,3), (6,1).", y);
    y = addBodyText(doc, "Stayed the same: the shape, the size and the orientation.", y);
    y = addBodyText(doc, "Changed: the position. Every corner moved 4 right and 2 up.", y);

    y = addSectionHeading(doc, "Task 2 — Reflection across x = 6", y, { color: C.PRIMARY });
    y = addBodyText(doc, "B' = (10,5), (8,5), (9,3).", y);
    y = addBodyText(doc, "Each point is the same distance from the line x = 6, but on the other side.", y);

    y = addSectionHeading(doc, "Task 3 — Rotation 90 degrees clockwise about (6,3)", y, { color: C.ACCENT });
    y = addBodyText(doc, "C' = (5,4), (5,6), (3,5).", y);
    y = addBodyText(doc, "Use the rule: (x,y) about centre (a,b) -> (a + (y - b), b - (x - a)).", y);
    y = addBodyText(doc, "Extension - 180 degrees: C'' = (5,2), (3,2), (4,4). The arrow points the opposite way.", y);

    y = addTipBox(doc,
      "Watch for: students who only move some corners in a translation; students who mirror only one side in a reflection; students who confuse 90 degrees clockwise with 90 degrees anticlockwise.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
