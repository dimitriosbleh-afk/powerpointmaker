"use strict";

// 2D Shapes Unit — Lesson 1: Sides and Corners
// Year 2 Numeracy | Variant 2 (Slate & Copper)
// VC2M2SP01 - recognise, compare and classify shapes by features
// Daily Review: Length - which is longer?
// Fluency: Skip counting by 2s

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
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, dailyReviewSlide, fluencySlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics,
  addNumberTrack,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 5;
const UNIT_TITLE = "2D Shapes";
const FOOTER = `2D Shapes | Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`;
const OUT_DIR = "output/SHA_Lesson1_Sides_and_Corners";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 1 Sides and Corners",
  "Count the sides and corners of 4 shapes. Write the numbers in the boxes.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 1 Answer Key",
  "Sides and corners for each shape on the worksheet.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a coloured 2D shape on the slide. shapeKey is one of:
//   "triangle" | "square" | "rectangle" | "pentagon" | "hexagon" | "circle"
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
  slide.addShape(type, {
    x, y, w, h,
    fill: { color: fillColor },
    line: { color: lineColor, width: 2.5 },
  });
}

// Draw a shape AND number-label every corner with a small chip.
// cornerPositions are RELATIVE positions (0..1) inside the shape's bounding box.
// PptxGenJS hexagon = horizontal regular hexagon (flat top/bottom).
// Corners sit at the 6 vertex positions of that shape, not at edge midpoints.
const CORNER_POSITIONS = {
  triangle: [[0.50, 0.06], [0.06, 0.94], [0.94, 0.94]],
  square:   [[0.06, 0.06], [0.94, 0.06], [0.94, 0.94], [0.06, 0.94]],
  rectangle:[[0.06, 0.10], [0.94, 0.10], [0.94, 0.90], [0.06, 0.90]],
  pentagon: [[0.50, 0.04], [0.95, 0.42], [0.79, 0.96], [0.21, 0.96], [0.05, 0.42]],
  hexagon:  [[0.25, 0.05], [0.75, 0.05], [0.97, 0.50], [0.75, 0.95], [0.25, 0.95], [0.03, 0.50]],
};

function drawShapeWithCorners(slide, shapeKey, x, y, w, h, opts) {
  const o = opts || {};
  const fill = o.fill || C.SECONDARY;
  drawShape(slide, shapeKey, x, y, w, h, fill);
  const positions = CORNER_POSITIONS[shapeKey] || [];
  const dotSize = o.dotSize || 0.30;
  const showLabels = o.showLabels !== false;
  positions.forEach((p, i) => {
    const cx = x + w * p[0];
    const cy = y + h * p[1];
    slide.addShape("roundRect", {
      x: cx - dotSize / 2, y: cy - dotSize / 2, w: dotSize, h: dotSize,
      rectRadius: dotSize / 2,
      fill: { color: C.ALERT },
      line: { color: C.WHITE, width: 1.2 },
    });
    if (showLabels) {
      slide.addText(String(i + 1), {
        x: cx - dotSize / 2, y: cy - dotSize / 2, w: dotSize, h: dotSize,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
    }
  });
}

// Draw a horizontal pencil shape (for Daily Review length comparison).
function drawPencil(slide, x, y, w, h, color) {
  // Pencil body
  slide.addShape("rect", {
    x: x + h * 0.6, y, w: w - h * 1.0, h,
    fill: { color: color || C.ACCENT },
    line: { color: C.CHARCOAL, width: 1.5 },
  });
  // Pencil tip (triangle on left)
  slide.addShape("triangle", {
    x: x, y, w: h * 0.6, h,
    fill: { color: "F2D08C" },
    line: { color: C.CHARCOAL, width: 1.5 },
    rotate: 270,
  });
  // Eraser (square on right)
  slide.addShape("rect", {
    x: x + w - h * 0.4, y, w: h * 0.4, h,
    fill: { color: "D64545" },
    line: { color: C.CHARCOAL, width: 1.5 },
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Good morning everyone. We are starting a new maths topic today.
- Our topic is 2D shapes - flat shapes you can draw on paper.
- Today we are looking at sides and corners.

DO:
- Have a triangle, square, pentagon and hexagon shape card ready to hold up.
- Settle students on the floor before clicking past the title.

TEACHER NOTES:
Lesson 1 of 5 in the 2D Shapes unit. Some students may already know the words "sides" and "corners". Build the language with everyone using the labelled visuals.

WATCH FOR:
- Students who already name shapes - use them in turn-and-tells later.
- Students who look unsure - the I Do labels every side and corner.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the two pencils.
- One is longer. One is shorter.
- Whisper to your partner: which pencil is longer? Blue or green?

DO:
- Let students whisper for 10 seconds.
- Ask 1-2 students to share before clicking the reveal.

TEACHER NOTES:
Daily Review focus: length comparison. Both pencils start at the same line on the left so students compare by the right end. This is a quick warm up, not new teaching.

WATCH FOR:
- Students who say "blue" - they are comparing the right ends correctly.
- Students who say "green because it is bigger" - prompt: "Look at where they end."

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check. The blue pencil is longer.
- Both pencils start at the same line. The blue one ends further along.

DO:
- Click to reveal the answer bar.
- Run your finger along both pencils to show the comparison.

TEACHER NOTES:
Tick and fix. Take note of any students who looked at thickness or colour instead of length.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. We are going to skip count by 2s.
- Stand up. Count with me. Start at 0. Stop at 20.
- 0... 2... 4... 6...
- One more time, a little faster.

DO:
- Lead choral count 0, 2, 4 ... 20 with claps on each number.
- Repeat 2-3 times, increasing pace each round.
- Sit students down before the next slide.

TEACHER NOTES:
Brisk automaticity routine. The number track highlights the count-by-2s pattern. Not new teaching - keep it fast.

WATCH FOR:
- Students who lose track at 12-14 - common stumbling spot.
- Students who say "1, 2, 3" instead - prompt: "Skip the odd numbers, jump 2 each time."

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Let us read the learning intention together.
- We are learning to count the sides and corners of 2D shapes.
- A side is one of the straight or curved lines around the shape.
- A corner is where two sides meet.

DO:
- Choral read the LI.
- Choral read each "I can" statement.
- Hold up one shape card and run your finger along a side, then point to a corner.

TEACHER NOTES:
SC1 is reachable for all - just pointing. SC2 is the core target. SC3 stretches students to count both features on the same shape.

WATCH FOR:
- Students who repeat "side" and "corner" while pointing - language is forming.
- Students who look unsure - the I Do will model both.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. This is a triangle.
- I am going to count the sides. A side is one of the straight lines around the shape.
- 1 side. 2 sides. 3 sides. A triangle has 3 sides.
- Now I count the corners. A corner is where two sides meet.
- 1 corner. 2 corners. 3 corners. A triangle has 3 corners.
- Say with me: 3 sides, 3 corners.

DO:
- Trace each side with your finger, counting aloud.
- Touch each numbered corner dot, counting aloud.
- Repeat the language pattern: "3 sides, 3 corners" three times.

TEACHER NOTES:
First exposure to "side" and "corner" with formal counting. Move slowly. Students often miss a side because they lose track of the start - that is why the corners are numbered.

MISCONCEPTIONS:
- Misconception: Students count one corner twice when they go around the shape.
  Why: They start counting and forget where they began.
  Impact: They report 4 corners on a triangle.
  Quick correction: "Look at the numbers on each corner. Stop when you see number 1 again."

WATCH FOR:
- Students who count "1, 2, 3" with you - language is forming.
- Students who count edges of the corner dots - prompt: "One side is one straight line."

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Look at this square.
- How many sides does the square have?
- Show me on your fingers. Ready... show.

DO:
- Display the square with numbered corners.
- Wait 5 seconds.
- Say: "Show me on your fingers... ready... show."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "How many sides? Show me on your fingers... show."
- Scan for: 4 fingers held up.
PROCEED: If 80% show 4 fingers, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count corners instead of sides.
- Reteach: "A side is one straight line. Watch me trace 1 side... 2 sides... 3 sides... 4 sides."
- Re-check: "Show me on fingers again."

TEACHER NOTES:
Fingers are the fastest every-student response. Watch the room for the 4-finger sweep.

WATCH FOR:
- Students who hold up 4 quickly - secure.
- Students who hold up 8 - they counted both sides and corners; reteach the difference.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. This shape is called a pentagon.
- We are going to count the sides together.
- I will trace. You count out loud.
- Then we will count the corners.

DO:
- Trace each side slowly, students count aloud.
- Pause at the last side - "How many sides?"
- Touch each numbered corner, students count.
- On their whiteboards: "Write: ___ sides, ___ corners."

TEACHER NOTES:
We Do uses a pentagon - a shape they may not have named before. The language pattern from I Do (sides, then corners) carries directly across.

WATCH FOR:
- Students who say "5 sides, 5 corners" - secure with the language.
- Students who write only one number - prompt: "Both - sides AND corners."

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check. A pentagon has 5 sides.
- A pentagon has 5 corners.
- Same number of sides as corners. Always true for shapes with straight sides.

DO:
- Click to reveal.
- Repeat the language pattern with the class.

TEACHER NOTES:
The "same number" pattern is a useful early generalisation but do not labour it. We will revisit it in later lessons.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Look at the hexagon.
- On your whiteboard, write: ___ sides, ___ corners.
- When you finish, hold your whiteboard up.

DO:
- Allow 60 seconds.
- Walk and scan whiteboards.
- Cold call 1-2 students to share their count.

TEACHER NOTES:
Hexagon is a new shape name for many. The numbered corners are a built-in scaffold. The goal today is the count, not memorising the name.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the triangle from the I Do. Count sides and corners again.
EXTENDING PROMPT:
- Task: Draw your own shape with 4 sides on the back of your whiteboard. Count the corners.

WATCH FOR:
- Students who write 6 sides, 6 corners - on track.
- Students who write 12 - they counted dots and sides as the same thing; reteach the difference.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. Look at the shape.
- On your whiteboard, write: ___ sides, ___ corners.

DO:
- Allow 60 seconds.
- Collect whiteboards or take a quick photo.

TEACHER NOTES:
Exit ticket assesses SC2 (core target). The shape is a pentagon (5 sides, 5 corners). The corners are numbered to support students who lose track.

WATCH FOR:
- Students who write 5 sides, 5 corners - on track.
- Students who write only one number - they attended to one feature only.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- We learned new words today. Sides. Corners.
- A side is a line around the shape.
- A corner is where two sides meet.
- Show me thumbs up if you can count sides on a shape.
- Turn and tell your partner: how many corners on a square?

DO:
- Read the success criteria with the class.
- Use thumbs up / sideways / down.
- Cold call 1-2 for the partner share.

TEACHER NOTES:
Self-assessment data informs Lesson 2 grouping. Most students should reach SC1 today; SC2 and SC3 build over the unit.

WATCH FOR:
- Strong thumbs up - they have the language.
- Sideways or down - small group focus tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Today we used shape cards and the worksheet.

DO:
- Print the Lesson 1 Sides and Corners sheet, one per student.
- Have shape cards (triangle, square, pentagon, hexagon) ready for partner work.

TEACHER NOTES:
One printed student resource for this lesson. Shape cards and mini-whiteboards do most of the active work.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Sides and Corners",
    `Year 2 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 2-3: Daily Review with reveal — Length: which pencil is longer?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Which Pencil Is Longer?", { color: C.ACCENT });

      // Hero card spanning the slide
      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

      // Start line for both pencils (left edge alignment)
      const startX = 1.0;
      const lineY = CONTENT_TOP + 0.45;
      const lineH = 3.0;
      s.addShape("line", {
        x: startX - 0.05, y: lineY, w: 0, h: lineH,
        line: { color: C.CHARCOAL, width: 2 },
      });
      s.addText("Start", {
        x: startX - 0.7, y: lineY + lineH / 2 - 0.15, w: 0.6, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, align: "right", margin: 0,
      });

      // Blue pencil (longer) — top
      drawPencil(s, startX, lineY + 0.55, 6.5, 0.55, "1976D2");
      s.addText("Blue", {
        x: 7.7, y: lineY + 0.60, w: 1.2, h: 0.45,
        fontSize: 22, fontFace: FONT_H, color: "1976D2", bold: true, valign: "middle", margin: 0,
      });

      // Green pencil (shorter) — bottom
      drawPencil(s, startX, lineY + 1.85, 4.6, 0.55, "2E7D32");
      s.addText("Green", {
        x: 5.8, y: lineY + 1.90, w: 1.4, h: 0.45,
        fontSize: 22, fontFace: FONT_H, color: "2E7D32", bold: true, valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "The blue pencil is longer.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Skip count by 2s with number track
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Skip Count by 2s to 20", { color: C.ACCENT });

  // Number track 0-10 row 1
  const trackY = CONTENT_TOP + 0.40;
  addNumberTrack(sFluency, 0.5, trackY, 9.0, 0, 10, [0, 2, 4, 6, 8, 10], { cellH: 0.85, fontSize: 28 });
  // Row 2: 11-20 (11-20 highlight only the multiples of 2)
  addNumberTrack(sFluency, 0.5, trackY + 1.05, 9.0, 11, 20, [12, 14, 16, 18, 20], { cellH: 0.85, fontSize: 28 });

  // Stand up and clap card at bottom
  addCard(sFluency, 1.5, trackY + 2.20, 7.0, 0.85, { strip: C.PRIMARY });
  sFluency.addText("Stand up. Count by 2s. Clap each number.", {
    x: 1.7, y: trackY + 2.22, w: 6.6, h: 0.80,
    fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to count the sides and corners of 2D shapes.",
    [
      "I can point to a side of a shape.",
      "I can point to a corner of a shape.",
      "I can count the sides and corners on a shape.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Triangle: 3 sides, 3 corners
  workedExSlide(pres, 2, "I Do", "A Triangle Has 3 Sides and 3 Corners",
    [
      "This is a triangle.",
      "",
      "Count the sides:",
      "1, 2, 3 sides.",
      "",
      "Count the corners:",
      "1, 2, 3 corners.",
      "",
      "3 sides. 3 corners.",
    ],
    NOTES_IDO, FOOTER,
    (slide, lg) => {
      // Triangle with numbered corners on the right
      drawShapeWithCorners(slide, "triangle",
        lg.rightX + 0.30, lg.panelTopPadded + 0.10,
        lg.rightW - 0.60, 3.55,
        { fill: C.SECONDARY }
      );
    }
  );

  // Slides 7-8: CFU with reveal — How many sides on a square?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How Many Sides?", { color: C.ALERT });

      // CHECK stamp top right
      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      // Hero card
      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Square in the middle, sized to leave room for the prompt below.
      drawShapeWithCorners(s, "square", 3.75, CONTENT_TOP + 0.20, 2.5, 2.5,
        { fill: C.PRIMARY });

      // Prompt below — sits inside the card, clear of the square.
      s.addText("Show me on your fingers.", {
        x: 0.7, y: 4.45, w: 8.6, h: 0.55,
        fontSize: 26, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "4 sides.   And 4 corners.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 9-10: We Do with reveal — Pentagon
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Count the Pentagon Together", { color: STAGE_COLORS["3"] });

      // Left card: pentagon visual
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      drawShapeWithCorners(s, "pentagon", 1.30, CONTENT_TOP + 0.20, 2.9, 2.9,
        { fill: C.SECONDARY });
      s.addText("pentagon", {
        x: 0.6, y: 4.50, w: 4.3, h: 0.40,
        fontSize: 20, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Right card: prompt
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "____ sides", options: { fontSize: 30, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "____ corners", options: { fontSize: 30, bold: true, color: C.CHARCOAL } },
      ], {
        x: 5.4, y: CONTENT_TOP + 0.30, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.55,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "5 sides.   5 corners.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 11: You Do — Hexagon
  workedExSlide(pres, 4, "You Do", "Your Turn: Count the Hexagon",
    [
      "First: Look at the shape.",
      "",
      "Next: Count the sides.",
      "",
      "Then: Count the corners.",
      "",
      "Write on your whiteboard:",
      "____ sides",
      "____ corners",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      drawShapeWithCorners(slide, "hexagon",
        lg.rightX + 0.30, lg.panelTopPadded + 0.10,
        lg.rightW - 0.60, 3.55,
        { fill: C.PRIMARY });
    }
  );

  // Slide 12: Exit Ticket — Pentagon (single shape, write count)
  // Use a custom slide rather than exitTicketSlide because we want a visual.
  const sExit = pres.addSlide();
  sExit.background = { color: C.BG_CARD };
  sExit.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.ASSESS || C.ALERT } });
  addBadge(sExit, "Exit Ticket", { color: C.ASSESS || C.ALERT });
  addTitle(sExit, "Show What You Know", { color: C.ASSESS || C.ALERT });

  // Left: shape
  addCard(sExit, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ASSESS || C.ALERT });
  drawShapeWithCorners(sExit, "pentagon", 1.30, CONTENT_TOP + 0.30, 2.9, 2.9,
    { fill: C.SECONDARY });

  // Right: write prompt
  addCard(sExit, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
  sExit.addText([
    { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "____ sides", options: { fontSize: 30, bold: true, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "____ corners", options: { fontSize: 30, bold: true, color: C.CHARCOAL } },
  ], {
    x: 5.4, y: CONTENT_TOP + 0.30, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.55,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sExit, FOOTER);
  sExit.addNotes(NOTES_EXIT);
  runSlideDiagnostics(sExit, pres, { respectSafeBottom: false });

  // Slide 13: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: how many corners on a square?",
      scItems: [
        "I can point to a side of a shape.",
        "I can point to a corner of a shape.",
        "I can count the sides and corners on a shape.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 14: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "SHA_Lesson1_Sides_and_Corners.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Worksheet — count sides and corners of 4 shapes.
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Look at each shape. Count the sides. Count the corners. Write the numbers.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addTipBox(doc,
      "A side is one line around the shape. A corner is where two sides meet. Touch each one as you count.",
      y, { color: C.ACCENT });

    // Draw a 2x2 grid of shapes with answer lines.
    // PDFKit drawing primitives. Page width is 612 pts (Letter), margin 50.
    const pageX = 50;
    const pageW = 512;
    const cellW = pageW / 2;
    const cellH = 215;
    const startY = y + 8;

    function drawShapeCellPdf(cx, cy, label, drawFn) {
      // Cell border
      doc.lineWidth(1).strokeColor("#" + C.MUTED)
        .rect(cx, cy, cellW - 10, cellH).stroke();
      // Shape name label
      doc.fillColor("#" + C.PRIMARY).fontSize(13).font("Sans-Bold")
        .text(label, cx + 12, cy + 10);
      // Shape drawing area (centred)
      const shapeBoxX = cx + 16;
      const shapeBoxY = cy + 36;
      const shapeBoxW = (cellW - 10) - 32;
      const shapeBoxH = 110;
      drawFn(shapeBoxX, shapeBoxY, shapeBoxW, shapeBoxH);
      // Answer lines
      const lineY = cy + 158;
      doc.fillColor("#333333").fontSize(11).font("Sans-Bold");
      doc.text("Sides: ____", cx + 18, lineY);
      doc.text("Corners: ____", cx + 18, lineY + 22);
    }

    function drawTrianglePdf(x, y, w, h) {
      const cx = x + w / 2;
      doc.lineWidth(2).strokeColor("#333333").fillColor("#" + C.SECONDARY);
      doc.moveTo(cx, y).lineTo(x + w, y + h).lineTo(x, y + h).closePath().fillAndStroke();
    }
    function drawSquarePdf(x, y, w, h) {
      const side = Math.min(w, h);
      const sx = x + (w - side) / 2;
      const sy = y;
      doc.lineWidth(2).strokeColor("#333333").fillColor("#" + C.PRIMARY)
        .rect(sx, sy, side, side).fillAndStroke();
    }
    function drawPolygonPdf(x, y, w, h, sides, color) {
      const cx = x + w / 2;
      const cy = y + h / 2;
      const r = Math.min(w, h) / 2 - 2;
      doc.lineWidth(2).strokeColor("#333333").fillColor("#" + color);
      const startAngle = -Math.PI / 2;
      doc.moveTo(cx + r * Math.cos(startAngle), cy + r * Math.sin(startAngle));
      for (let i = 1; i < sides; i += 1) {
        const a = startAngle + (i * 2 * Math.PI) / sides;
        doc.lineTo(cx + r * Math.cos(a), cy + r * Math.sin(a));
      }
      doc.closePath().fillAndStroke();
    }

    // Row 1
    drawShapeCellPdf(pageX,            startY, "Triangle",
      (sx, sy, sw, sh) => drawTrianglePdf(sx, sy, sw, sh));
    drawShapeCellPdf(pageX + cellW,    startY, "Square",
      (sx, sy, sw, sh) => drawSquarePdf(sx, sy, sw, sh));
    // Row 2
    drawShapeCellPdf(pageX,            startY + cellH + 8, "Pentagon",
      (sx, sy, sw, sh) => drawPolygonPdf(sx, sy, sw, sh, 5, C.ACCENT));
    drawShapeCellPdf(pageX + cellW,    startY + cellH + 8, "Hexagon",
      (sx, sy, sw, sh) => drawPolygonPdf(sx, sy, sw, sh, 6, C.SECONDARY));

    addPdfFooter(doc, `Lesson ${SESSION} | Sides and Corners | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Sides and corners for each shape.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addSectionHeading(doc, "Triangle", y, { color: C.PRIMARY });
    y = addBodyText(doc, "3 sides. 3 corners.", y);
    y = addSectionHeading(doc, "Square", y, { color: C.PRIMARY });
    y = addBodyText(doc, "4 sides. 4 corners.", y);
    y = addSectionHeading(doc, "Pentagon", y, { color: C.PRIMARY });
    y = addBodyText(doc, "5 sides. 5 corners.", y);
    y = addSectionHeading(doc, "Hexagon", y, { color: C.PRIMARY });
    y = addBodyText(doc, "6 sides. 6 corners.", y);
    y = addTipBox(doc,
      "For shapes with straight sides, the number of sides always equals the number of corners.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
