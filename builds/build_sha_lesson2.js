"use strict";

// 2D Shapes Unit — Lesson 2: Straight and Curved Sides
// Year 2 Numeracy | Variant 2 (Slate & Copper)
// VC2M2SP01 - sort shapes by straight vs curved sides
// Daily Review: Length - order pencils
// Fluency: Doubles to 10

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade2", weekToVariant(3));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  workedExSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 5;
const UNIT_TITLE = "2D Shapes";
const FOOTER = `2D Shapes | Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`;
const OUT_DIR = "output/SHA_Lesson2_Straight_and_Curved";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 2 Sort Straight or Curved",
  "Look at each shape. Tick straight only or has a curve.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 2 Answer Key",
  "Which shapes are straight only and which have a curve.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

function drawShape(slide, shapeKey, x, y, w, h, fill, line) {
  const lineColor = line || C.CHARCOAL;
  const fillColor = fill || C.SECONDARY;
  let type = "rect";
  if (shapeKey === "triangle")  type = "triangle";
  else if (shapeKey === "square" || shapeKey === "rectangle") type = "rect";
  else if (shapeKey === "pentagon") type = "pentagon";
  else if (shapeKey === "hexagon")  type = "hexagon";
  else if (shapeKey === "circle" || shapeKey === "oval") type = "ellipse";
  else if (shapeKey === "trapezoid") type = "trapezoid";
  else if (shapeKey === "diamond") type = "diamond";
  slide.addShape(type, {
    x, y, w, h,
    fill: { color: fillColor },
    line: { color: lineColor, width: 2.5 },
  });
}

// Draw a labelled shape card — small visual box with the shape and a name underneath.
function drawShapeCard(slide, x, y, w, h, shapeKey, label, color) {
  // Card background
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.10,
    fill: { color: C.WHITE },
    line: { color: C.MUTED, width: 1 },
  });
  // Shape area
  const shapeY = y + 0.10;
  const labelH = 0.40;
  const shapeH = h - labelH - 0.20;
  const shapeW = Math.min(w - 0.30, shapeH);
  const shapeX = x + (w - shapeW) / 2;
  drawShape(slide, shapeKey, shapeX, shapeY, shapeW, shapeH, color || C.SECONDARY);
  // Label
  slide.addText(label, {
    x, y: y + h - labelH - 0.04, w, h: labelH,
    fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

function drawPencil(slide, x, y, w, h, color) {
  slide.addShape("rect", {
    x: x + h * 0.6, y, w: w - h * 1.0, h,
    fill: { color: color || C.ACCENT },
    line: { color: C.CHARCOAL, width: 1.5 },
  });
  slide.addShape("triangle", {
    x: x, y, w: h * 0.6, h,
    fill: { color: "F2D08C" },
    line: { color: C.CHARCOAL, width: 1.5 },
    rotate: 270,
  });
  slide.addShape("rect", {
    x: x + w - h * 0.4, y, w: h * 0.4, h,
    fill: { color: "D64545" },
    line: { color: C.CHARCOAL, width: 1.5 },
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Yesterday we counted sides and corners.
- Today we look at sides more closely.
- Some sides are straight. Some sides are curved.

DO:
- Have a square card and a circle card ready to hold up.
- Settle students on the floor before clicking past the title.

TEACHER NOTES:
Lesson 2 of 5. Builds on yesterday's "side" language. Today we sort shapes by the kind of side - straight or curved.

WATCH FOR:
- Students who already use "round" - link it to "curved".

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the three pencils.
- They all start at the same line on the left.
- Whisper to your partner: which pencil is the longest?

DO:
- Allow 15 seconds for whisper.
- Cold call 1-2 students before clicking the reveal.

TEACHER NOTES:
Daily Review focus: length. Same start line so students compare end positions. Continues yesterday's compare-by-end strategy.

WATCH FOR:
- Students who say "the green one" because it is bigger looking - prompt: "Look at where each one ends."

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check. The blue pencil is the longest.
- The yellow is in the middle. The red is the shortest.

DO:
- Click to reveal the answer bar.
- Run your finger along each pencil.

TEACHER NOTES:
Tick and fix. Note any students who confused length with thickness or colour.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. We are practising doubles.
- I will say a number. You whisper its double, then say it loud.
- Whisper, then say. Ready?
- "Double 4." (whisper) "8."
- "Double 5." (whisper) "10."
- "Double 6." (whisper) "12."

DO:
- Lead choral response on each prompt.
- Repeat each card 2-3 times faster.

TEACHER NOTES:
Brisk doubles to 10. Not new teaching. Whisper-then-say slows guessing and gives all students a moment to recall.

WATCH FOR:
- Students who hesitate on double 6 - common stumbling spot.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to sort shapes by their sides.
- Some sides are straight. Some sides are curved.
- Let us read the success criteria.

DO:
- Choral read the LI and each SC.
- Hold up a triangle (all straight) and a circle (all curved) as you say "straight" and "curved".

TEACHER NOTES:
SC1 is reachable for all - just point. SC2 is the core target. SC3 needs students to sort with reasoning.

WATCH FOR:
- Students who repeat "straight" and "curved" - language is forming.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. Two shapes here. A triangle and a circle.
- The triangle has straight sides. Look - 1, 2, 3 straight sides.
- The circle has one curved side. It goes all the way around.
- Triangle: all straight. Circle: all curved.

DO:
- Trace each side of the triangle with your finger, saying "straight".
- Trace around the circle with your finger, saying "curved".
- Repeat: "straight" or "curved" while pointing.

TEACHER NOTES:
Establish the language with the two clearest examples - triangle for all-straight, circle for all-curved. The "all the way around" tracing motion helps students see a circle has only one side.

MISCONCEPTIONS:
- Misconception: Students think a circle has many sides because their finger keeps moving.
  Why: They are counting movement, not sides.
  Impact: They report the circle has lots of sides.
  Quick correction: "One side. Your finger never lifts and never turns a corner."

WATCH FOR:
- Students who say "round" for circle - accept and link to "curved".

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Look at this shape.
- Does it have a curved side? Thumbs up for yes. Thumbs down for no.

DO:
- Display the oval.
- Wait 5 seconds.
- Say: "Thumbs up or thumbs down... show."

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Curved side? Thumbs up if yes."
- Scan for: thumbs UP. The oval has one curved side.
PROCEED: If 80% show thumbs up, click to reveal and move to We Do.
PIVOT: Most likely misconception - students confuse oval with rectangle.
- Reteach: "Trace the edge with your finger. Did your finger turn a corner?"
- Re-check: Show again, ask same question.

TEACHER NOTES:
The oval has one curved side - same as a circle. The reveal confirms the language.

WATCH FOR:
- Confident thumbs up - secure.
- Mixed thumbs - reteach with finger trace.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Three shapes here. A square. A circle. A triangle.
- We are going to sort them into two groups.
- Straight only on this side. Has a curve on that side.
- With your partner, whisper which shape goes in which group.

DO:
- Allow 30 seconds for partner whisper.
- Cold call 1-2 students to suggest a sort.
- Click to reveal the sort.

TEACHER NOTES:
We Do: classroom sort. The shapes use the same names from yesterday so the focus is on the kind of side, not on the name.

WATCH FOR:
- Pairs who put the circle on the curve side and the others on straight - secure.
- Pairs who put the triangle on the curve side - prompt: "Trace it. Did your finger turn?"

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check. Square and triangle: straight only.
- Circle: has a curve.

DO:
- Click to reveal the sort.
- Run your finger along each shape as you name the group.

TEACHER NOTES:
The reveal shows the correct sort. Students who got it can lead a partner who did not.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Look at the four shapes.
- On your whiteboard, write the names or numbers of the shapes that have a curve.

DO:
- Allow 90 seconds.
- Walk and scan whiteboards.
- Cold call 2 students to share.

TEACHER NOTES:
You Do mixes 4 shapes. Two have curves (circle, oval). Two are straight only (hexagon, pentagon). The independent task is "find the curves".

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Just point to one shape with a curve. Say "curve" out loud.
EXTENDING PROMPT:
- Task: After your sort, draw a new shape that has both a straight side AND a curved side.

WATCH FOR:
- Students who pick circle and oval - on track.
- Students who pick the hexagon - they may be confused by the slanted sides; prompt: "Slanted is still straight. No corner means a curve."

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. Look at the shape.
- On your whiteboard, write: straight only OR has a curve.

DO:
- Allow 60 seconds.
- Collect whiteboards or photo for records.

TEACHER NOTES:
Exit ticket assesses SC2 (core). The shape is a half-circle - it has a straight side AND a curved side, so the answer is "has a curve". Accept "both" as correct - the student is noticing both features.

WATCH FOR:
- Students who write "has a curve" or "both" - on track.
- Students who write "straight only" - they missed the curve; reteach with finger trace.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today we sorted shapes by their sides.
- Straight or curved.
- Show me thumbs up if you can find a curved side.
- Turn and tell your partner: name a shape with all straight sides.

DO:
- Read the success criteria.
- Use thumbs up / sideways / down.
- Cold call 1-2 students for the partner share.

TEACHER NOTES:
Self-assessment data informs Lesson 3 grouping.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- The sort sheet is for after our session.

DO:
- Print the Lesson 2 Sort Straight or Curved sheet, one per student.
- Have shape cards ready for partner work.

TEACHER NOTES:
One printed student resource. Mini-whiteboards do most of the active work.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Straight and Curved Sides",
    `Year 2 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slides 2-3: Daily Review with reveal — Order three pencils
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Which Pencil Is the Longest?", { color: C.ACCENT });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

      const startX = 1.0;
      const lineY = CONTENT_TOP + 0.30;
      const lineH = 3.20;
      s.addShape("line", {
        x: startX - 0.05, y: lineY, w: 0, h: lineH,
        line: { color: C.CHARCOAL, width: 2 },
      });
      s.addText("Start", {
        x: startX - 0.7, y: lineY + lineH / 2 - 0.15, w: 0.6, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, align: "right", margin: 0,
      });

      // Three pencils — different lengths
      drawPencil(s, startX, lineY + 0.35, 6.5, 0.50, "1976D2"); // longest blue
      s.addText("Blue", {
        x: 7.7, y: lineY + 0.40, w: 1.2, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: "1976D2", bold: true, valign: "middle", margin: 0,
      });

      drawPencil(s, startX, lineY + 1.30, 4.6, 0.50, "F2A52C"); // middle yellow
      s.addText("Yellow", {
        x: 5.8, y: lineY + 1.35, w: 1.4, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: "B86E10", bold: true, valign: "middle", margin: 0,
      });

      drawPencil(s, startX, lineY + 2.25, 3.0, 0.50, "D64545"); // shortest red
      s.addText("Red", {
        x: 4.2, y: lineY + 2.30, w: 1.2, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: "D64545", bold: true, valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Blue is longest.   Yellow is middle.   Red is shortest.", {
        x: 0.7, y: 4.45, w: 8.6, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Doubles to 10
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Doubles to 10", { color: C.ACCENT });

  // Three big prompt cards: double 4, double 5, double 6
  const prompts = ["double 4", "double 5", "double 6"];
  const cardY = CONTENT_TOP + 0.30;
  const cardH = 2.80;
  const cardGap = 0.20;
  const totalW = 9.0;
  const cardW = (totalW - cardGap * 2) / 3;
  prompts.forEach((p, i) => {
    const x = 0.5 + i * (cardW + cardGap);
    addCard(sFluency, x, cardY, cardW, cardH, { strip: C.ACCENT });
    sFluency.addText(p, {
      x, y: cardY + 0.20, w: cardW, h: cardH - 0.40,
      fontSize: 44, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  // Below: whisper-then-say cue
  addCard(sFluency, 1.5, cardY + cardH + 0.20, 7.0, 0.55, { strip: C.PRIMARY });
  sFluency.addText("Whisper the answer. Then say it loud.", {
    x: 1.7, y: cardY + cardH + 0.22, w: 6.6, h: 0.50,
    fontSize: 20, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to sort shapes by their sides.",
    [
      "I can find a straight side on a shape.",
      "I can find a curved side on a shape.",
      "I can sort shapes into straight only and has a curve.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Triangle (all straight) + Circle (all curved)
  const sIDo = pres.addSlide();
  addTopBar(sIDo, STAGE_COLORS["2"]);
  addStageBadge(sIDo, 2, "I Do");
  addTitle(sIDo, "Straight Sides and Curved Sides", { color: STAGE_COLORS["2"] });

  // Two big shape cards side by side. Each card has a coloured header banner
  // (STRAIGHT / CURVED), a shape in the middle, and a small caption below.
  const idoCardW = 4.30;
  const idoCardH = SAFE_BOTTOM - CONTENT_TOP - 0.10;
  const idoCardY = CONTENT_TOP + 0.05;
  const bannerH = 0.55;

  function drawIDoCard(slide, x, y, w, h, header, headerColor, shapeKey, shapeColor, caption) {
    // Card background
    slide.addShape("roundRect", {
      x, y, w, h, rectRadius: 0.10,
      fill: { color: C.WHITE },
      line: { color: C.MUTED, width: 1 },
    });
    // Header banner
    addTextOnShape(slide, header, {
      x, y, w, h: bannerH, rectRadius: 0.10,
      fill: { color: headerColor },
    }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    // Shape
    const shapeY = y + bannerH + 0.10;
    const captionH = 0.40;
    const shapeBoxH = h - bannerH - captionH - 0.20;
    const shapeBoxW = Math.min(w - 0.40, shapeBoxH);
    const shapeX = x + (w - shapeBoxW) / 2;
    drawShape(slide, shapeKey, shapeX, shapeY, shapeBoxW, shapeBoxH, shapeColor);
    // Caption below shape
    slide.addText(caption, {
      x, y: y + h - captionH - 0.04, w, h: captionH,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }

  drawIDoCard(sIDo, 0.6, idoCardY, idoCardW, idoCardH,
    "STRAIGHT", C.PRIMARY, "triangle", C.SECONDARY, "all straight sides");
  drawIDoCard(sIDo, 5.10, idoCardY, idoCardW, idoCardH,
    "CURVED", C.ACCENT, "circle", C.ACCENT, "one curved side");

  addFooter(sIDo, FOOTER);
  sIDo.addNotes(NOTES_IDO);
  runSlideDiagnostics(sIDo, pres, { respectSafeBottom: false });

  // Slides 7-8: CFU with reveal — Has this oval got a curved side?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Has It Got a Curved Side?", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Big oval centred
      drawShape(s, "oval", 3.50, CONTENT_TOP + 0.25, 3.0, 2.0, C.SECONDARY);

      // Prompt below
      s.addText("Thumbs up for yes. Thumbs down for no.", {
        x: 0.7, y: 4.45, w: 8.6, h: 0.55,
        fontSize: 24, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Thumbs UP. The oval has one curved side.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 9-10: We Do with reveal — Sort 3 shapes
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Sort the Three Shapes", { color: STAGE_COLORS["3"] });

      // Top strip: three shapes to sort
      const shapeStripY = CONTENT_TOP + 0.10;
      const stripH = 1.40;
      addCard(s, 0.5, shapeStripY, 9.0, stripH, { strip: STAGE_COLORS["3"] });

      const shapeW = 1.20;
      const shapeY = shapeStripY + 0.10;
      drawShape(s, "square",   1.30, shapeY, shapeW, shapeW, C.PRIMARY);
      s.addText("square",   { x: 1.0, y: shapeY + shapeW + 0.02, w: shapeW + 0.60, h: 0.25,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

      drawShape(s, "circle",   4.40, shapeY, shapeW, shapeW, C.ACCENT);
      s.addText("circle",   { x: 4.10, y: shapeY + shapeW + 0.02, w: shapeW + 0.60, h: 0.25,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

      drawShape(s, "triangle", 7.50, shapeY, shapeW, shapeW, C.SECONDARY);
      s.addText("triangle", { x: 7.20, y: shapeY + shapeW + 0.02, w: shapeW + 0.60, h: 0.25,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

      // Two empty sort columns
      const sortY = shapeStripY + stripH + 0.20;
      const sortH = SAFE_BOTTOM - sortY - 0.10;
      const sortW = 4.30;
      // Left column: STRAIGHT
      addCard(s, 0.5, sortY, sortW, sortH, { strip: C.PRIMARY });
      addTextOnShape(s, "STRAIGHT only", {
        x: 0.5, y: sortY, w: sortW, h: 0.40,
        fill: { color: C.PRIMARY },
      }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true });
      // Right column: CURVED
      addCard(s, 5.20, sortY, sortW, sortH, { strip: C.ACCENT });
      addTextOnShape(s, "HAS a curve", {
        x: 5.20, y: sortY, w: sortW, h: 0.40,
        fill: { color: C.ACCENT },
      }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      // Reveal: place shapes in their columns
      const sortY = CONTENT_TOP + 0.10 + 1.40 + 0.20;
      const sortH = SAFE_BOTTOM - sortY - 0.10;
      // Inside left column: square + triangle
      const leftShapeY = sortY + 0.55;
      drawShape(slide, "square",   1.20, leftShapeY, 1.20, 1.20, C.PRIMARY);
      drawShape(slide, "triangle", 3.00, leftShapeY, 1.20, 1.20, C.SECONDARY);
      // Inside right column: circle
      drawShape(slide, "circle",   6.85, leftShapeY, 1.20, 1.20, C.ACCENT);
      // Tag at bottom of slide
      addTextOnShape(slide, "Square and triangle = straight only.   Circle = has a curve.", {
        x: 0.5, y: SAFE_BOTTOM - 0.05, w: 9.0, h: 0.0,
        rectRadius: 0,
        fill: { color: C.SUCCESS },
      }, { fontSize: 1, color: C.WHITE });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 11: You Do — Find the curves (4 shapes)
  const sYouDo = pres.addSlide();
  addTopBar(sYouDo, STAGE_COLORS["4"]);
  addStageBadge(sYouDo, 4, "You Do");
  addTitle(sYouDo, "Your Turn: Find the Curves", { color: STAGE_COLORS["4"] });

  // Left card: instructions
  addCard(sYouDo, 0.5, CONTENT_TOP, 4.30, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["4"] });
  sYouDo.addText([
    { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Write the names of the shapes with a curved side.", options: { fontSize: 18, color: C.CHARCOAL } },
  ], {
    x: 0.70, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  // Right card: 4 shapes labelled
  addCard(sYouDo, 5.0, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
  // 2x2 grid of shapes
  const ydW = 1.55;
  const ydX1 = 5.30;
  const ydX2 = 7.45;
  const ydY1 = CONTENT_TOP + 0.20;
  const ydY2 = CONTENT_TOP + 1.95;
  drawShape(sYouDo, "circle",   ydX1, ydY1, ydW, ydW, C.ACCENT);
  sYouDo.addText("circle", { x: ydX1 - 0.20, y: ydY1 + ydW + 0.02, w: ydW + 0.40, h: 0.25,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  drawShape(sYouDo, "hexagon",  ydX2, ydY1, ydW, ydW, C.SECONDARY);
  sYouDo.addText("hexagon", { x: ydX2 - 0.20, y: ydY1 + ydW + 0.02, w: ydW + 0.40, h: 0.25,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  drawShape(sYouDo, "oval",     ydX1, ydY2, ydW + 0.20, ydW, C.ACCENT);
  sYouDo.addText("oval", { x: ydX1 - 0.20, y: ydY2 + ydW + 0.02, w: ydW + 0.40, h: 0.25,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  drawShape(sYouDo, "pentagon", ydX2, ydY2, ydW, ydW, C.PRIMARY);
  sYouDo.addText("pentagon", { x: ydX2 - 0.20, y: ydY2 + ydW + 0.02, w: ydW + 0.40, h: 0.25,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  addFooter(sYouDo, FOOTER);
  sYouDo.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(sYouDo, pres, { respectSafeBottom: false });

  // Slide 12: Exit Ticket — half circle (D-shape) — has both a straight side and a curved side.
  // Use a custom slide with the shape as the visual.
  const sExit = pres.addSlide();
  sExit.background = { color: C.BG_CARD };
  sExit.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.ASSESS || C.ALERT } });
  addBadge(sExit, "Exit Ticket", { color: C.ASSESS || C.ALERT });
  addTitle(sExit, "Show What You Know", { color: C.ASSESS || C.ALERT });

  // Left: the shape — a D shape (half-circle) drawn manually as ellipse + rectangle mask is too complex;
  // we use a chord shape: ellipse with the bottom half hidden by a white rectangle on top of it.
  addCard(sExit, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ASSESS || C.ALERT });
  // Draw a "D" by using a chord cut from an ellipse:
  // ellipse w=2.6 h=2.6, then a white rectangle covers the top half so only the bottom half (D rotated) shows.
  // Simpler: rotate -90 so D points right.
  // Cleanest approach: half-ellipse using "chord" shape if available, otherwise compose with two shapes.
  const dCx = 2.75;
  const dCy = CONTENT_TOP + 1.80;
  const dR = 1.20;
  // Bottom half-circle effect: draw full ellipse, then mask top half with a card-coloured rect.
  drawShape(sExit, "circle", dCx - dR, dCy - dR, dR * 2, dR * 2, C.SECONDARY);
  // Mask top half with a BG_CARD rectangle (so the visible part is a D opening downward → looks like a half-moon)
  sExit.addShape("rect", {
    x: dCx - dR - 0.05, y: dCy - dR - 0.05, w: dR * 2 + 0.10, h: dR + 0.05,
    fill: { color: C.WHITE },
    line: { color: C.WHITE, width: 0 },
  });
  // Now draw the straight (top) edge as a thick line
  sExit.addShape("line", {
    x: dCx - dR, y: dCy, w: dR * 2, h: 0,
    line: { color: C.CHARCOAL, width: 2.5 },
  });
  // And re-stroke the curved arc by placing a transparent ellipse outline (re-stroke trick: replace
  // by a coloured ellipse with no fill — pptxgenjs has no "noFill", so re-add the ellipse with white
  // strokes... simpler approach: just leave the ellipse as-is and accept that the inside is one colour).
  // The visual reads as "half circle" because the masked half is white.

  // Right: write prompt
  addCard(sExit, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
  sExit.addText([
    { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Does this shape have a curve?", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Write yes or no.", options: { fontSize: 26, bold: true, color: C.CHARCOAL } },
  ], {
    x: 5.4, y: CONTENT_TOP + 0.20, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sExit, FOOTER);
  sExit.addNotes(NOTES_EXIT);
  runSlideDiagnostics(sExit, pres, { respectSafeBottom: false });

  // Slide 13: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: name a shape with all straight sides.",
      scItems: [
        "I can find a straight side on a shape.",
        "I can find a curved side on a shape.",
        "I can sort shapes into straight only and has a curve.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 14: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "SHA_Lesson2_Straight_and_Curved.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Worksheet — sort shapes
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Look at each shape. Tick the box that matches.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addTipBox(doc,
      "Trace each shape with your finger. If your finger turns at a corner, the side is straight. If your finger does not lift or turn, the side is curved.",
      y, { color: C.ACCENT });

    // Page width 612, margin 50, content width 512
    const colShape = 90;
    const colName  = 130;
    const colStr   = 140;
    const colCur   = 140;
    const tableX = 50;
    const tableY = y + 8;
    const rowH = 75;

    // Header row
    doc.fillColor("#" + C.PRIMARY).rect(tableX, tableY, colShape + colName + colStr + colCur, 30).fill();
    doc.fillColor("#FFFFFF").fontSize(13).font("Sans-Bold");
    doc.text("Shape",         tableX + 8, tableY + 9, { width: colShape - 16, align: "center" });
    doc.text("Name",          tableX + colShape + 8, tableY + 9, { width: colName - 16, align: "center" });
    doc.text("Straight only", tableX + colShape + colName + 8, tableY + 9, { width: colStr - 16, align: "center" });
    doc.text("Has a curve",   tableX + colShape + colName + colStr + 8, tableY + 9, { width: colCur - 16, align: "center" });

    const shapes = [
      { name: "Triangle", drawer: (cx, cy) => {
          doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.SECONDARY);
          doc.moveTo(cx, cy - 25).lineTo(cx + 25, cy + 22).lineTo(cx - 25, cy + 22).closePath().fillAndStroke();
      } },
      { name: "Circle", drawer: (cx, cy) => {
          doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.ACCENT);
          doc.circle(cx, cy, 26).fillAndStroke();
      } },
      { name: "Square", drawer: (cx, cy) => {
          doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.PRIMARY);
          doc.rect(cx - 24, cy - 24, 48, 48).fillAndStroke();
      } },
      { name: "Oval", drawer: (cx, cy) => {
          doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.ACCENT);
          doc.ellipse(cx, cy, 32, 22).fillAndStroke();
      } },
      { name: "Pentagon", drawer: (cx, cy) => {
          doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.PRIMARY);
          const r = 26; const start = -Math.PI / 2;
          doc.moveTo(cx + r * Math.cos(start), cy + r * Math.sin(start));
          for (let i = 1; i < 5; i += 1) {
            const a = start + i * 2 * Math.PI / 5;
            doc.lineTo(cx + r * Math.cos(a), cy + r * Math.sin(a));
          }
          doc.closePath().fillAndStroke();
      } },
    ];

    shapes.forEach((shape, i) => {
      const rowY = tableY + 30 + i * rowH;
      doc.lineWidth(0.8).strokeColor("#" + C.MUTED)
        .rect(tableX, rowY, colShape + colName + colStr + colCur, rowH).stroke();
      // Vertical separators
      doc.moveTo(tableX + colShape, rowY).lineTo(tableX + colShape, rowY + rowH).stroke();
      doc.moveTo(tableX + colShape + colName, rowY).lineTo(tableX + colShape + colName, rowY + rowH).stroke();
      doc.moveTo(tableX + colShape + colName + colStr, rowY).lineTo(tableX + colShape + colName + colStr, rowY + rowH).stroke();
      // Shape
      shape.drawer(tableX + colShape / 2, rowY + rowH / 2);
      // Name
      doc.fillColor("#333333").fontSize(13).font("Sans-Bold")
        .text(shape.name, tableX + colShape + 8, rowY + rowH / 2 - 8, { width: colName - 16, align: "center" });
      // Tick boxes
      doc.lineWidth(1.4).strokeColor("#333333").fillColor("#FFFFFF");
      doc.rect(tableX + colShape + colName + colStr / 2 - 14, rowY + rowH / 2 - 14, 28, 28).fillAndStroke();
      doc.rect(tableX + colShape + colName + colStr + colCur / 2 - 14, rowY + rowH / 2 - 14, 28, 28).fillAndStroke();
    });

    addPdfFooter(doc, `Lesson ${SESSION} | Sort Straight or Curved | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Which shapes are straight only and which have a curve.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addSectionHeading(doc, "Straight only", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Triangle, square, pentagon.", y);
    y = addSectionHeading(doc, "Has a curve", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Circle, oval.", y);
    y = addTipBox(doc,
      "If a shape has any curved edge, even one, it goes in the has-a-curve group.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
