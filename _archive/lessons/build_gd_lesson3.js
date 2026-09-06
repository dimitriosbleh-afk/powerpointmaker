"use strict";

// Geometry and Data Unit (Year 5/6 Numeracy) - Lesson 3: Tessellation exploration.
// Year 6 VC2M6SP03 - recognise and use combinations of transformations to create tessellations.
// Today: WHICH regular polygons tessellate and WHY.
// Daily Review: Adding and subtracting fractions and decimals.
// Fluency: Division facts from multiplication (gradual release - less scaffold than mult).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 2);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 8;
const UNIT_TITLE = "Geometry and Data: Shapes, Tessellations, Data";
const FOOTER = `Geometry and Data | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/GD_Lesson3_Tessellation_Exploration";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 3 Tessellation Investigation",
  "Investigate which regular polygons tessellate and why.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 3 Answer Key",
  "Worked answers for the Tessellation Investigation.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// =====================================================================
// Teacher notes
// =====================================================================

const NOTES_TITLE = `SAY:
- Today we look at tessellations - shapes that fit together with no gaps and no overlaps.
- Some of you may have seen tessellations in bathroom tiles or honeycomb. If this feels new, that is okay.

DO:
- Settle.
- Whiteboards ready.

TEACHER NOTES:
Lesson 3 of 8. Today is the WHY of tessellations - which regular polygons tessellate and what the angles tell us. Tomorrow students CREATE their own tessellation for the classroom display.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- Whiteboards on desks.
- One Tessellation Investigation sheet per student.
- Optional: paper cut-outs of an equilateral triangle, square, regular pentagon and regular hexagon if you have them.

DO:
- Print one student sheet per student.
- Print one answer key.

TEACHER NOTES:
Kinder square paper can stand in for the square cut-out. Cut a few equilateral triangles in advance if you can.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review.
- Add or subtract on your whiteboard.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Adding and subtracting fractions and decimals. Same denominator on the fractions; tenths on the decimals.

WATCH FOR:
- Students who line up the decimal points - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 3/8 + 2/8 = 5/8.
- 0.7 + 0.4 = 1.1.
- 0.95 - 0.6 = 0.35.

DO:
- Click to reveal.

TEACHER NOTES:
Quick reveal. The third one catches students who think 0.6 only takes from the tenths.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Division facts.
- Use the matching multiplication fact.
- Whisper-answer then write.

DO:
- Display the three prompts.
- 25 seconds.

TEACHER NOTES:
Day three. Today's release: connect division facts to multiplication facts. If you know 7 x 8 = 56, you know 56 / 7 = 8.

WATCH FOR:
- Students who say the matching multiplication fact under their breath - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 56 divided by 7 = 8.
- 72 divided by 9 = 8.
- 48 divided by 6 = 8.

DO:
- Click to reveal.

TEACHER NOTES:
Note the same answer on all three. That is on purpose - the same answer means students focused on the strategy not the digit.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Look at the picture - bathroom tiles, honeycomb.
- Notice: no gaps, no overlaps. Each shape meets its neighbours edge to edge.
- That is called a tessellation.
- Today's question - which shapes will do this?

DO:
- Display the visual.

TEACHER NOTES:
Two real-world tessellations on the slide. Use them to build the meaning before we explore which shapes work.

WATCH FOR:
- Students who name another tessellation they have seen - tracking.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to investigate which regular polygons tessellate and why.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 - name a shape that tessellates. SC2 - use the angle around a point to test a regular polygon. SC3 - explain in words why some shapes do not tessellate.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Three key words today.
- Tessellation - shapes that fit together with no gaps and no overlaps.
- Regular polygon - all sides equal, all angles equal.
- Interior angle - the angle inside the shape at each corner.

DO:
- Point at each visual.

TEACHER NOTES:
Anchor "regular" - it means equal-everything. A regular hexagon has six equal sides and six equal angles. A non-regular hexagon might tessellate or might not.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_IDO_360 = `SAY:
- Watch me. For shapes to tessellate, the angles meeting at any point must add to exactly 360 degrees.
- 360 is a full turn. No gap, no overlap.
- If they add to less than 360, there is a gap. If more, they overlap.

DO:
- Use the visual on the slide - the angles around the centre point.
- Draw the full turn with your finger.

TEACHER NOTES:
This is the core idea. Every other test in the lesson depends on this. Make sure students see "360 degrees" written down before moving on.

WATCH FOR:
- Students who can say "360 degrees around a point" - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_TRIANGLE = `SAY:
- Equilateral triangle. Each angle is 60 degrees.
- 6 triangles meeting at a point - 6 times 60 = 360. Tessellates.

DO:
- Show the diagram - 6 triangles around the centre.
- Highlight 6 x 60.

TEACHER NOTES:
The equilateral triangle is the first test case. The angle 60 divides 360 cleanly six times.

WATCH FOR:
- Students who can do 60 x 6 in their head - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_SQUARE = `SAY:
- Square. Each angle is 90 degrees.
- 4 squares meeting at a point - 4 times 90 = 360. Tessellates.

DO:
- Show the diagram - 4 squares around the centre.

TEACHER NOTES:
The square is the easiest example. Most students will visualise floor tiles.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_HEXAGON = `SAY:
- Regular hexagon. Each angle is 120 degrees.
- 3 hexagons meeting at a point - 3 times 120 = 360. Tessellates.
- This is the honeycomb shape.

DO:
- Show the diagram - 3 hexagons around the centre.

TEACHER NOTES:
Bees use this for a reason - the hexagon packs tightly with no wasted material.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- A regular pentagon has interior angles of 108 degrees.
- Do regular pentagons tessellate? Yes or No on your board.

DO:
- Display the prompt.
- 60 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me on three, two, one, show."
- Scan for: NO.
PROCEED: If 80% have NO, click to reveal and move to We Do.
PIVOT: Most likely misconception - students guess YES because all five-sided shapes look as regular as triangles or squares.
- Reteach: Try 108 x 3 = 324 (gap) and 108 x 4 = 432 (overlap). Neither = 360.
- Re-check: Show a regular octagon (135 degrees). Does it tessellate?

TEACHER NOTES:
108 x 3 = 324, leaves a gap. 108 x 4 = 432, overlaps. Pentagons do not tessellate alone.

WATCH FOR:
- Students who multiply 108 x 3 = 324 and explain the gap - secure.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO = `SAY:
- With me. We test a regular octagon.
- Each interior angle is 135 degrees.
- Test: how many around a point?
- 135 x 1 = 135. Not 360.
- 135 x 2 = 270. Not 360.
- 135 x 3 = 405. Too much. Overlap.
- Conclusion - regular octagons alone do not tessellate.

DO:
- Build the test on the board.
- Use the structure: multiply, compare, conclude.

TEACHER NOTES:
The octagon is the test that closes the loop. It only works in combination with a square (octagon + square is a classic semi-regular tessellation - extension if time).

WATCH FOR:
- Students who set out their working in three lines - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Result - regular octagons do not tessellate alone.
- 135 does not divide 360 cleanly.

DO:
- Click to reveal.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn. Take the Tessellation Investigation sheet.
- For each shape - work out the angle around a point.
- Decide - tessellates or not.

DO:
- Distribute the sheet.
- Circulate. Ask "what is your test?" before drawing conclusions.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Stick with Parts A and B (triangle, square, hexagon). Use the formula 360 / angle.
- Extra Notes: Sit with these students. Confirm the multiplication is correct.
EXTENDING PROMPT:
- Task: Investigate a SEMI-REGULAR tessellation - use a square + a regular octagon. Sketch the arrangement and check the angles add to 360.
- Extra Notes: 90 + 135 + 135 = 360.

TEACHER NOTES:
Different content from We Do. Same structured test.

WATCH FOR:
- Students who write the test on each line - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Q1: Name TWO regular polygons that tessellate by themselves.
- Q2: Why does a regular pentagon NOT tessellate? Use the angle.

DO:
- Display the prompt.
- 2 minutes.

TEACHER NOTES:
Exit ticket assesses SC2 and SC3. Q1 - any two of triangle, square, hexagon. Q2 - 108 degrees does not divide 360 cleanly.

WATCH FOR:
- Students who give a number reason in Q2 - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria for today.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: WHY do shapes need to add to 360 around a point?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea - angles around a point must add to 360 degrees for a tessellation to work. Tomorrow students design their own tessellation art using combinations of transformations.

[General: Closing | VTLM 2.0: Reflection]`;

// =====================================================================
// Helpers - draw regular polygons and tessellation arrangements
// =====================================================================

// Draw a regular polygon with n sides centred at (cx, cy) with radius r.
function regularPolygonPoints(cx, cy, r, n, rotation) {
  const pts = [];
  const rot = rotation || -Math.PI / 2;
  for (let i = 0; i < n; i++) {
    const ang = rot + (2 * Math.PI / n) * i;
    pts.push({ x: cx + r * Math.cos(ang), y: cy + r * Math.sin(ang) });
  }
  return pts;
}

function drawPolygon(slide, pts, opts) {
  const o = opts || {};
  const color = o.color || C.PRIMARY;
  const width = o.width || 2;
  for (let i = 0; i < pts.length; i++) {
    const a = pts[i];
    const b = pts[(i + 1) % pts.length];
    const minX = Math.min(a.x, b.x);
    const minY = Math.min(a.y, b.y);
    slide.addShape("line", {
      x: minX, y: minY,
      w: Math.abs(b.x - a.x), h: Math.abs(b.y - a.y),
      line: { color, width },
      flipV: (a.x < b.x) === (a.y < b.y) ? false : true,
    });
  }
}

// Draw n copies of a regular polygon meeting at a centre - used for the angle-around-a-point visuals.
function drawTessellationCluster(slide, cx, cy, r, n, copies, color) {
  // Each polygon shares the centre as one vertex. Place polygons rotated around centre.
  // Strategy: draw 'copies' polygons whose first vertex is at the centre, rotated by k*(360/copies) degrees.
  for (let k = 0; k < copies; k++) {
    const offset = (2 * Math.PI / copies) * k;
    // Place polygon centre so that one of its vertices lies at (cx, cy).
    // The polygon centre is at distance r from this vertex, in the direction "outward".
    const dirX = Math.cos(offset);
    const dirY = Math.sin(offset);
    // Polygon centre is at (cx + r*dirX, cy + r*dirY) but oriented so its vertex at angle (offset + 180) is at the centre.
    const polyCx = cx + r * dirX;
    const polyCy = cy + r * dirY;
    // Rotation: the vertex pointing inward should be at angle (offset + pi) from poly centre.
    // For a regular polygon with vertices at -pi/2 + 2*pi*i/n, we want -pi/2 = (offset + pi) effectively rotate by ...
    // Simplest: rotate so first vertex points back to centre.
    const polyRotation = offset + Math.PI - Math.PI / 2 * (1 + 0);
    // Actually: each polygon should be rotated so one of its vertices is at the centre. Use rotation = offset + pi (the direction from polyCentre back to (cx,cy)).
    // standard rotation in regularPolygonPoints is -pi/2 + 2pi*i/n. Setting rotation = offset + pi - pi/2 ... easier: just use rotation = offset + pi.
    // For an n-gon to have its first vertex at angle offset+pi from its own centre, rotation must equal that minus 0.
    const rot = offset + Math.PI;
    const pts = regularPolygonPoints(polyCx, polyCy, r, n, rot);
    drawPolygon(slide, pts, { color, width: 1.8 });
  }
}

// =====================================================================
// Build
// =====================================================================

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 3: Which shapes tessellate?",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions and decimals",
      ["3/8 + 2/8", "0.7 + 0.4", "0.95 - 0.6"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "5/8         1.1         0.35", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal
  withReveal(
    () => fluencySlide(pres, "Fluency: Division (use the matching mult fact)",
      ["56 div 7", "72 div 9", "48 div 6"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "8        8        8", {
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
    addTitle(s, "No gaps, no overlaps");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    // Left: a square tile pattern
    const tileX = 1.2;
    const tileY = CONTENT_TOP + 0.45;
    const tileSize = 0.40;
    for (let r = 0; r < 5; r++) {
      for (let c = 0; c < 6; c++) {
        s.addShape("rect", {
          x: tileX + c * tileSize, y: tileY + r * tileSize, w: tileSize, h: tileSize,
          fill: { color: (r + c) % 2 ? C.PRIMARY : C.ACCENT },
          line: { color: C.WHITE, width: 1 },
        });
      }
    }
    s.addText("Floor tiles", {
      x: tileX, y: tileY + 5 * tileSize + 0.15, w: 6 * tileSize, h: 0.35,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", margin: 0,
    });

    // Right: a hexagon pattern
    const hexX = 5.8;
    const hexY = CONTENT_TOP + 0.45;
    const hexR = 0.28;
    const hexW = hexR * Math.sqrt(3);
    for (let r = 0; r < 4; r++) {
      for (let c = 0; c < 4; c++) {
        const cx = hexX + c * hexW + (r % 2 ? hexW / 2 : 0);
        const cy = hexY + r * hexR * 1.5;
        const pts = regularPolygonPoints(cx, cy, hexR, 6, 0);
        drawPolygon(s, pts, { color: C.ALERT, width: 1.5 });
      }
    }
    s.addText("Honeycomb", {
      x: hexX, y: hexY + 5 * hexR * 1.5 + 0.05, w: 4 * hexW, h: 0.35,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", margin: 0,
    });

    s.addText("Every shape meets its neighbours edge to edge. Today: which shapes can do this?", {
      x: 0.7, y: SAFE_BOTTOM - 0.50, w: 8.6, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to investigate which regular polygons tessellate and why.",
    [
      "I can name at least one shape that tessellates.",
      "I can use the angle around a point to test whether a regular polygon tessellates.",
      "I can explain in words why some regular polygons do not tessellate.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key vocabulary
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Tessellation, regular polygon, interior angle");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const items = [
      { word: "Tessellation",   simple: "no gaps, no overlaps", detail: "Shapes fit together like floor tiles.",  color: C.PRIMARY },
      { word: "Regular polygon", simple: "all sides equal",     detail: "All angles equal, all sides equal.",     color: C.ACCENT },
      { word: "Interior angle", simple: "angle inside",         detail: "The angle inside the shape at a corner.", color: C.ALERT },
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

  // Slide 10: I Do - the 360 rule
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "The 360 rule", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Centre point with a full turn arrow
    const cx = 3.0;
    const cy = CONTENT_TOP + 1.60;
    const r = 1.10;
    // draw a full circle (ROUNDED_RECT trick)
    s.addShape("ellipse", {
      x: cx - r, y: cy - r, w: 2 * r, h: 2 * r,
      fill: { color: C.BG_LIGHT, transparency: 30 },
      line: { color: stageColor, width: 2 },
    });
    s.addText("360°", {
      x: cx - 0.8, y: cy - 0.3, w: 1.6, h: 0.6,
      fontSize: 28, fontFace: FONT_H, color: stageColor, bold: true,
      align: "center", margin: 0,
    });

    // Right: explanation
    addCard(s, 5.20, CONTENT_TOP + 0.25, 4.30, 2.90, { strip: C.ACCENT, fill: C.WHITE });
    s.addText([
      { text: "Around any point: 360 degrees.", options: { fontSize: 16, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "= 360", options: { fontSize: 14, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "  the shapes fit (tessellate).", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "< 360", options: { fontSize: 14, color: C.ALERT, bold: true, breakLine: true } },
      { text: "  there is a gap.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "> 360", options: { fontSize: 14, color: C.ALERT, bold: true, breakLine: true } },
      { text: "  the shapes overlap.", options: { fontSize: 13, color: C.CHARCOAL } },
    ], {
      x: 5.40, y: CONTENT_TOP + 0.45, w: 4.00, h: 2.60,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_360);
  })();

  // Slide 11: I Do - Equilateral triangle (6 x 60 = 360)
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Equilateral triangle: 6 x 60 = 360", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Left: 6 triangles meeting at a centre
    const cx = 2.8;
    const cy = CONTENT_TOP + 1.60;
    drawTessellationCluster(s, cx, cy, 0.95, 3, 6, C.PRIMARY);
    // Centre dot
    s.addShape("roundRect", {
      x: cx - 0.08, y: cy - 0.08, w: 0.16, h: 0.16, rectRadius: 0.08,
      fill: { color: C.ALERT }, line: { color: C.ALERT, width: 1 },
    });

    // Right: working
    addCard(s, 5.20, CONTENT_TOP + 0.30, 4.30, 2.80, { strip: C.SUCCESS, fill: C.WHITE });
    s.addText([
      { text: "Triangle interior angle = 60°", options: { fontSize: 16, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "60 x 6 = 360", options: { fontSize: 22, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "Six triangles fit around a point.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
      { text: "Tessellates.", options: { fontSize: 16, color: C.SUCCESS, bold: true } },
    ], {
      x: 5.40, y: CONTENT_TOP + 0.50, w: 4.00, h: 2.40,
      fontFace: FONT_B, valign: "top", margin: 0, align: "center",
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_TRIANGLE);
  })();

  // Slide 12: I Do - Square (4 x 90 = 360)
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Square: 4 x 90 = 360", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Left: 4 squares meeting at a centre
    const cx = 2.8;
    const cy = CONTENT_TOP + 1.60;
    const sz = 0.85;
    [
      { x: cx - sz, y: cy - sz }, { x: cx, y: cy - sz },
      { x: cx - sz, y: cy }, { x: cx, y: cy },
    ].forEach((p) => {
      s.addShape("rect", {
        x: p.x, y: p.y, w: sz, h: sz,
        fill: { color: C.BG_LIGHT }, line: { color: C.ACCENT, width: 2 },
      });
    });
    // Centre dot
    s.addShape("roundRect", {
      x: cx - 0.08, y: cy - 0.08, w: 0.16, h: 0.16, rectRadius: 0.08,
      fill: { color: C.ALERT }, line: { color: C.ALERT, width: 1 },
    });

    addCard(s, 5.20, CONTENT_TOP + 0.30, 4.30, 2.80, { strip: C.SUCCESS, fill: C.WHITE });
    s.addText([
      { text: "Square interior angle = 90°", options: { fontSize: 16, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "90 x 4 = 360", options: { fontSize: 22, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "Four squares fit around a point.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
      { text: "Tessellates.", options: { fontSize: 16, color: C.SUCCESS, bold: true } },
    ], {
      x: 5.40, y: CONTENT_TOP + 0.50, w: 4.00, h: 2.40,
      fontFace: FONT_B, valign: "top", margin: 0, align: "center",
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_SQUARE);
  })();

  // Slide 13: I Do - Hexagon (3 x 120 = 360)
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Regular hexagon: 3 x 120 = 360", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Left: 3 hexagons meeting at a centre
    const cx = 2.8;
    const cy = CONTENT_TOP + 1.60;
    drawTessellationCluster(s, cx, cy, 0.85, 6, 3, C.ALERT);
    s.addShape("roundRect", {
      x: cx - 0.08, y: cy - 0.08, w: 0.16, h: 0.16, rectRadius: 0.08,
      fill: { color: C.PRIMARY }, line: { color: C.PRIMARY, width: 1 },
    });

    addCard(s, 5.20, CONTENT_TOP + 0.30, 4.30, 2.80, { strip: C.SUCCESS, fill: C.WHITE });
    s.addText([
      { text: "Hexagon interior angle = 120°", options: { fontSize: 16, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "120 x 3 = 360", options: { fontSize: 22, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "Three hexagons fit around a point.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
      { text: "Tessellates - this is honeycomb.", options: { fontSize: 16, color: C.SUCCESS, bold: true } },
    ], {
      x: 5.40, y: CONTENT_TOP + 0.50, w: 4.00, h: 2.40,
      fontFace: FONT_B, valign: "top", margin: 0, align: "center",
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_HEXAGON);
  })();

  // Slides 14-15: CFU + reveal — Pentagon
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Regular pentagon: yes or no?", { color: C.ALERT });
      addTextOnShape(s, "CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Pentagon visual
      const cx = 2.8;
      const cy = CONTENT_TOP + 1.60;
      const pts = regularPolygonPoints(cx, cy, 1.00, 5);
      drawPolygon(s, pts, { color: C.PRIMARY, width: 2.5 });
      s.addText("Each angle = 108°", {
        x: cx - 1.2, y: cy + 1.20, w: 2.4, h: 0.40,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });

      // Right: the question prompt
      addCard(s, 5.20, CONTENT_TOP + 0.30, 4.30, 2.80, { strip: C.ALERT, fill: C.WHITE });
      s.addText([
        { text: "Each interior angle of a regular pentagon = 108°", options: { fontSize: 14, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Does 108 fit cleanly around 360?", options: { fontSize: 15, color: C.CHARCOAL, italic: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "On your board: YES or NO.", options: { fontSize: 16, color: C.ALERT, bold: true } },
      ], {
        x: 5.40, y: CONTENT_TOP + 0.55, w: 4.00, h: 2.30,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "NO. 108 x 3 = 324 (gap). 108 x 4 = 432 (overlap). Does not tessellate.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 16-17: We Do + reveal — Octagon
  withReveal(
    () => {
      const s = pres.addSlide();
      const stageColor = STAGE_COLORS["3"];
      addTopBar(s, stageColor);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Test: regular octagon (135°)", { color: stageColor });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

      // Octagon
      const cx = 2.8;
      const cy = CONTENT_TOP + 1.60;
      const pts = regularPolygonPoints(cx, cy, 1.00, 8);
      drawPolygon(s, pts, { color: C.SECONDARY, width: 2.5 });
      s.addText("Each angle = 135°", {
        x: cx - 1.2, y: cy + 1.20, w: 2.4, h: 0.40,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Right: structured working space
      addCard(s, 5.20, CONTENT_TOP + 0.30, 4.30, 2.80, { strip: stageColor, fill: C.WHITE });
      s.addText([
        { text: "Test it together:", options: { fontSize: 15, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "135 x 1 = ___", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "135 x 2 = ___", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "135 x 3 = ___", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Does any equal 360?", options: { fontSize: 14, color: C.CHARCOAL, italic: true } },
      ], {
        x: 5.40, y: CONTENT_TOP + 0.50, w: 4.00, h: 2.40,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "135, 270, 405. None equal 360. Regular octagons alone do NOT tessellate.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // Slide 18: You Do
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["4"];
    addTopBar(s, stageColor);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: investigate four shapes", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    s.addText([
      { text: "First: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "for each shape, find its interior angle.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Next: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "test if it divides cleanly into 360.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Then: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "write tessellates / does not tessellate.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Extension: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "design a tile pattern using 2 shapes.", options: { fontSize: 15, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    // Right preview: a small angle table summary
    addCard(s, 5.30, CONTENT_TOP + 0.20, 4.10, 3.10, { strip: stageColor, fill: C.WHITE });
    const tableHeaders = ["Shape", "Angle", "Tessellates?"];
    const tableRows = [
      ["Equilateral triangle", "60°", "yes"],
      ["Square", "90°", "yes"],
      ["Regular pentagon", "108°", "no"],
      ["Regular hexagon", "120°", "yes"],
    ];
    const tableX = 5.45;
    const tableY = CONTENT_TOP + 0.40;
    const colW = [1.95, 0.80, 1.05];
    const rowH = 0.42;
    // Header
    let tx = tableX;
    tableHeaders.forEach((h, i) => {
      s.addShape("rect", {
        x: tx, y: tableY, w: colW[i], h: rowH,
        fill: { color: C.PRIMARY }, line: { color: C.WHITE, width: 1 },
      });
      s.addText(h, {
        x: tx, y: tableY, w: colW[i], h: rowH,
        fontSize: 11, fontFace: FONT_H, color: C.WHITE,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
      tx += colW[i];
    });
    // Rows
    tableRows.forEach((row, ri) => {
      tx = tableX;
      row.forEach((cell, ci) => {
        s.addShape("rect", {
          x: tx, y: tableY + (ri + 1) * rowH, w: colW[ci], h: rowH,
          fill: { color: ri % 2 ? C.BG_LIGHT : C.WHITE }, line: { color: C.MUTED, width: 0.5 },
        });
        s.addText(cell, {
          x: tx, y: tableY + (ri + 1) * rowH, w: colW[ci], h: rowH,
          fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
          align: "center", valign: "middle", margin: 0,
        });
        tx += colW[ci];
      });
    });

    s.addText("Use the Tessellation Investigation sheet.", {
      x: 5.30, y: SAFE_BOTTOM - 0.40, w: 4.10, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.MUTED,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Name TWO regular polygons that tessellate by themselves.",
      "Why does a regular pentagon NOT tessellate? Use the angle in your answer.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: WHY do shapes need to add to 360 around a point?",
      scItems: [
        "I can name at least one shape that tessellates.",
        "I can use the angle around a point to test whether a regular polygon tessellates.",
        "I can explain in words why some regular polygons do not tessellate.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "GD_Lesson3_Tessellation_Exploration.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Test each regular polygon. Decide if it tessellates by itself.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "The angles meeting at any point in a tessellation must add to 360 degrees. " +
      "Test = multiply the interior angle. If a whole number of copies makes exactly 360, it tessellates.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Part A — Test each regular polygon", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Interior angles given. Work out which fit cleanly into 360 degrees.", y);
    y = addBodyText(doc, "Equilateral triangle (60°):", y);
    y = addWriteLine(doc, "60 x ___ = 360 ?     Tessellates (yes / no): ___", y);
    y = addBodyText(doc, "Square (90°):", y);
    y = addWriteLine(doc, "90 x ___ = 360 ?     Tessellates (yes / no): ___", y);
    y = addBodyText(doc, "Regular pentagon (108°):", y);
    y = addWriteLine(doc, "108 x ___ closest to 360?     Tessellates (yes / no): ___", y);
    y = addBodyText(doc, "Regular hexagon (120°):", y);
    y = addWriteLine(doc, "120 x ___ = 360 ?    Tessellates (yes / no): ___", y);

    y = addSectionHeading(doc, "Part B — Explain", y, { color: C.PRIMARY });
    y = addBodyText(doc, "In your own words, why does a regular pentagon not tessellate?", y);
    y = addWriteLine(doc, "_____________________________________________________________", y);
    y = addWriteLine(doc, "_____________________________________________________________", y);

    y = addSectionHeading(doc, "Part C — Combination tile (extension)", y, { color: C.ACCENT });
    y = addBodyText(doc, "A semi-regular tessellation uses 2 different shapes. A square (90°) and a regular octagon (135°) tile together.", y);
    y = addBodyText(doc, "Check: how many of each meet at a point?", y);
    y = addWriteLine(doc, "___ squares + ___ octagons    Total angle = ___ degrees.", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Tessellation Investigation | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Tessellation Investigation.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Part A — Tests", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Equilateral triangle: 60 x 6 = 360. Tessellates (yes).", y);
    y = addBodyText(doc, "Square: 90 x 4 = 360. Tessellates (yes).", y);
    y = addBodyText(doc, "Regular pentagon: 108 x 3 = 324 (gap). 108 x 4 = 432 (overlap). Does not tessellate (no).", y);
    y = addBodyText(doc, "Regular hexagon: 120 x 3 = 360. Tessellates (yes).", y);

    y = addSectionHeading(doc, "Part B — Explanation", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Sample answer: A regular pentagon does not tessellate because its interior angle (108 degrees) does not divide 360 evenly. Three pentagons leave a gap; four pentagons overlap.", y);

    y = addSectionHeading(doc, "Part C — Combination tile", y, { color: C.ACCENT });
    y = addBodyText(doc, "1 square (90°) + 2 octagons (135° each) meet at a point.", y);
    y = addBodyText(doc, "90 + 135 + 135 = 360 degrees.", y);
    y = addBodyText(doc, "This is one of the classic semi-regular tessellations - you see it in some tile floors.", y);

    y = addTipBox(doc,
      "Watch for: students who say a pentagon tessellates because three almost reach 360; students who multiply angles incorrectly; students who confuse a regular hexagon (tessellates) with a regular heptagon (does not).",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
