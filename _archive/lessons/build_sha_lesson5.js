"use strict";

// 2D Shapes Unit — Lesson 5: Same Shape, Different Position
// Year 2 Numeracy | Variant 2 (Slate & Copper)
// VC2M2SP01 - rotation does not change a shape; apply to a sports field
// Daily Review: Vertical addition (24 + 13)
// Fluency: 10 more / 10 less

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

const SESSION = 5;
const TOTAL = 5;
const UNIT_TITLE = "2D Shapes";
const FOOTER = `2D Shapes | Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`;
const OUT_DIR = "output/SHA_Lesson5_Same_Shape_Different_Position";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 5 Find the Shapes",
  "Look at the basketball court. Colour the rectangles. Circle the circle.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 5 Answer Key",
  "Where the rectangles and circle are on the court.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

function drawShape(slide, shapeKey, x, y, w, h, fill, line, opts) {
  const o = opts || {};
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
  const shape = {
    x, y, w, h,
    fill: { color: fillColor },
    line: { color: lineColor, width: 2.5 },
  };
  if (o.rotate != null) shape.rotate = o.rotate;
  slide.addShape(type, shape);
}

// Draw a basketball-court style rectangle with key rectangles, centre circle,
// and the half-court line. Suitable for Year 2 visual analysis.
function drawBasketballCourt(slide, x, y, w, h) {
  // Court (outer rectangle)
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: "F2D08C" },          // basketball-court tan
    line: { color: C.CHARCOAL, width: 3 },
  });
  // Half-court line (vertical, middle)
  slide.addShape("line", {
    x: x + w / 2, y, w: 0, h,
    line: { color: C.CHARCOAL, width: 2 },
  });
  // Centre circle
  const centreR = Math.min(w, h) * 0.10;
  slide.addShape("ellipse", {
    x: x + w / 2 - centreR, y: y + h / 2 - centreR, w: centreR * 2, h: centreR * 2,
    fill: { color: "F2D08C" },
    line: { color: C.CHARCOAL, width: 2 },
  });
  // Left key rectangle
  const keyW = w * 0.18;
  const keyH = h * 0.45;
  slide.addShape("rect", {
    x, y: y + (h - keyH) / 2, w: keyW, h: keyH,
    fill: { color: "F2D08C" },
    line: { color: C.CHARCOAL, width: 2 },
  });
  // Right key rectangle
  slide.addShape("rect", {
    x: x + w - keyW, y: y + (h - keyH) / 2, w: keyW, h: keyH,
    fill: { color: "F2D08C" },
    line: { color: C.CHARCOAL, width: 2 },
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- All week we have looked at 2D shapes.
- Today is the last lesson in our unit.
- We learn one more big idea: a shape stays the same shape no matter how you turn it.
- We will use everything we know to find shapes on a basketball court.

DO:
- Have a square card and a triangle card ready to rotate as you talk.
- Settle students on the floor.

TEACHER NOTES:
Lesson 5 of 5. Two ideas join up here: orientation does not change the shape, and shapes are everywhere - including sports fields. The court visual links lesson concepts to a real space.

WATCH FOR:
- Students who say a turned square is a "diamond" - that is the misconception we work on today.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Vertical addition.
- Look at the sum: 24 plus 13.
- Add the ones first. Then add the tens.
- Whisper to your partner. Then write your answer.

DO:
- Display 24 + 13 stacked.
- Allow 30 seconds for whisper and write.
- Walk and check.

TEACHER NOTES:
Daily Review focus: vertical addition. Ones first, then tens. No regrouping needed - 4+3=7 ones, 2+1=3 tens, total 37.

WATCH FOR:
- Students who add tens first - prompt: "Ones column first."
- Students who write 73 - they reversed tens and ones.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check. 4 plus 3 is 7. Write 7 in the ones.
- 2 tens plus 1 ten is 3 tens. Write 3 in the tens.
- The answer is 37.

DO:
- Click to reveal each column step.
- Trace each column with your finger.

TEACHER NOTES:
Tick and fix. Note any students who used the wrong column order or reversed digits.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. 10 more or 10 less.
- I will show you a number. Whisper the answer. Then write it.
- "10 more than 47."
- "10 less than 80."
- "10 more than 26."

DO:
- Display each prompt one at a time.
- Allow 10 seconds per prompt.
- Walk and check.

TEACHER NOTES:
Brisk practice. 10 more / less changes the tens digit only - this links to the place value work from earlier in the unit.

WATCH FOR:
- Students who add or subtract 1 instead of 10 - reteach with the tens column.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning that turning a shape does not change what it is.
- A square is still a square if you turn it. A triangle is still a triangle if you flip it.
- Let us read the success criteria.

DO:
- Choral read the LI and each SC.
- Pick up a square card. Slowly rotate it 45 degrees. "Still a square."
- Pick up a triangle card. Flip it upside down. "Still a triangle."

TEACHER NOTES:
SC1 is reachable for all - just naming a shape in different positions. SC2 is the core target. SC3 stretches students to use a feature word to describe shapes on the court.

WATCH FOR:
- Students who say "diamond" for the rotated square - we will fix that in I Do.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. Four triangles here. They all look different.
- One points up. One points down. One points sideways. One is on a slant.
- But they all have 3 sides and 3 corners.
- They are all triangles.
- Turning a shape does not change what it is.

DO:
- Point to each triangle and trace the sides.
- Say "3 sides, 3 corners - still a triangle" for each one.

TEACHER NOTES:
First idea: orientation does not change the shape. The 4 triangles are all the same kind of triangle, just turned. The features (3 sides, 3 corners) stay the same.

MISCONCEPTIONS:
- Misconception: Students think a turned shape becomes a different shape.
  Why: They learned each shape facing one way (a square always sits flat).
  Impact: They call a rotated square a "diamond" and treat it as a different shape.
  Quick correction: "Count the sides. Count the corners. Are they the same? Then it is the same shape."

WATCH FOR:
- Students who name "triangle" for each one - secure.
- Students who hesitate on the upside-down triangle - reteach with finger trace.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Look at this shape.
- It is sitting on a corner. It looks like a diamond.
- But is it still a square? Thumbs up for yes. Thumbs down for no.

DO:
- Display a square rotated 45 degrees.
- Wait 5 seconds.
- Say: "Square or not? Show me with thumbs."

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Still a square? Thumbs up if yes."
- Scan for: thumbs UP. The shape has 4 equal sides and 4 corners - still a square.
PROCEED: If 80% show thumbs up, click to reveal.
PIVOT: Most likely misconception - students think rotation makes it a "diamond".
- Reteach: "Count the sides. 1, 2, 3, 4. All the same length. 4 corners. It is a square."
- Re-check: Show again, ask "Still a square?"

TEACHER NOTES:
This is the hinge. The 'diamond' misconception is very common. The reveal confirms the answer and shows the same square in both positions.

WATCH FOR:
- Confident thumbs up - they are using features.
- Mixed - reteach with side count.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Look at this basketball court.
- I can see lots of shapes I know.
- Whisper to your partner: name one shape you can see on the court.
- Then we will name them all together.

DO:
- Display the basketball court mockup.
- Allow 30 seconds for partner whisper.
- Cold call 2-3 students.

TEACHER NOTES:
The court has: outer rectangle (the court), 2 small rectangles (the keys), 1 circle (the centre circle). Students notice the rectangles and the circle. The half-court line is just a line.

WATCH FOR:
- Students who say "rectangle" or "circle" - secure.
- Students who say "square" for the keys - prompt: "Are the sides all the same length? Or are some longer?"

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check. The whole court is a rectangle.
- Two rectangles inside - one at each end.
- One circle in the middle.

DO:
- Click to reveal the labels.
- Trace each shape on the court as you say its name.

TEACHER NOTES:
Reveal labels each shape. Some students may also notice the half-court line - accept "line" but note it is not a 2D shape on its own.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Look at the soccer field on your screen.
- On your whiteboard, write the names of the shapes you can see.
- Try to find at least 2 shapes.

DO:
- Allow 90 seconds.
- Walk and scan whiteboards.
- Cold call 2 students to share.

TEACHER NOTES:
The soccer field has: outer rectangle (the field), centre circle, two smaller rectangles (penalty areas). Students should find at least 2 shape names.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Just point to one rectangle and one circle on the field.
EXTENDING PROMPT:
- Task: Use a feature word to describe one shape. For example, "the field is a rectangle with 2 pairs of parallel sides".

WATCH FOR:
- Students who write "rectangle" and "circle" - on track.
- Students who write "square" for the field - prompt: "Trace the long side and the short side. Same length?"

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task in our unit.
- Look at the two shapes on the screen.
- One is sitting flat. One is turned.
- On your whiteboard, write: same OR different.

DO:
- Allow 60 seconds.
- Collect whiteboards or photo for records.

TEACHER NOTES:
Exit ticket assesses SC2 (core). Two squares - one upright, one rotated 45 degrees. Both are squares. The answer is "same".

WATCH FOR:
- Students who write "same" - on track.
- Students who write "different" - they used appearance, not features. Reteach with side count.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- That is the end of our 2D shapes unit.
- We learned about sides, corners, straight, curved, opposite, parallel, equal, and same shape when turned.
- Show me thumbs up if you can name a shape no matter how it is turned.
- Turn and tell your partner: name one shape you saw on the basketball court.

DO:
- Read the success criteria.
- Use thumbs up / sideways / down.
- Cold call 1-2 students for the share.

TEACHER NOTES:
End of unit. Self-assessment data informs the next maths topic. Most students should be secure with naming shapes by features.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- The find-the-shapes sheet is for after our session.

DO:
- Print the Lesson 5 Find the Shapes sheet, one per student.
- Have coloured pencils available.

TEACHER NOTES:
One printed student resource. Mini-whiteboards do most of the live work.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 5: Same Shape, Different Position",
    `Year 2 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 2-3: Daily Review with reveal — Vertical addition 24 + 13
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Vertical Addition", { color: C.ACCENT });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

      // Big stacked sum centred
      const sumX = 4.0;
      const sumY = CONTENT_TOP + 0.40;
      const fontSize = 70;
      // Tens / Ones header
      s.addText("tens", {
        x: sumX, y: sumY, w: 0.85, h: 0.30,
        fontSize: 14, fontFace: FONT_B, color: C.MUTED,
        align: "center", margin: 0,
      });
      s.addText("ones", {
        x: sumX + 0.85, y: sumY, w: 0.85, h: 0.30,
        fontSize: 14, fontFace: FONT_B, color: C.MUTED,
        align: "center", margin: 0,
      });
      // Top number 24
      s.addText("2", { x: sumX, y: sumY + 0.30, w: 0.85, h: 0.85,
        fontSize, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0 });
      s.addText("4", { x: sumX + 0.85, y: sumY + 0.30, w: 0.85, h: 0.85,
        fontSize, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0 });
      // Plus sign + bottom number 13
      s.addText("+", { x: sumX - 0.55, y: sumY + 1.20, w: 0.45, h: 0.85,
        fontSize, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0 });
      s.addText("1", { x: sumX, y: sumY + 1.20, w: 0.85, h: 0.85,
        fontSize, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0 });
      s.addText("3", { x: sumX + 0.85, y: sumY + 1.20, w: 0.85, h: 0.85,
        fontSize, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0 });
      // Underline
      s.addShape("line", {
        x: sumX - 0.10, y: sumY + 2.10, w: 1.85, h: 0,
        line: { color: C.CHARCOAL, width: 3 },
      });

      // Prompt
      s.addText("On your whiteboard: write the answer.", {
        x: 0.7, y: 4.45, w: 8.6, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      // Reveal the answer underneath the line.
      const sumX = 4.0;
      const sumY = CONTENT_TOP + 0.40;
      slide.addText("3", { x: sumX, y: sumY + 2.20, w: 0.85, h: 0.85,
        fontSize: 70, fontFace: FONT_H, color: C.SUCCESS, bold: true, align: "center", margin: 0 });
      slide.addText("7", { x: sumX + 0.85, y: sumY + 2.20, w: 0.85, h: 0.85,
        fontSize: 70, fontFace: FONT_H, color: C.SUCCESS, bold: true, align: "center", margin: 0 });
      addTextOnShape(slide, "Ones first: 4 + 3 = 7. Then tens: 2 + 1 = 3. Answer: 37.", {
        x: 0.5, y: 4.45, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — 10 more / 10 less (3 prompts)
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "10 More or 10 Less", { color: C.ACCENT });

  const prompts = ["10 more than 47", "10 less than 80", "10 more than 26"];
  const cardY = CONTENT_TOP + 0.30;
  const cardH = 2.80;
  const cardGap = 0.20;
  const totalW = 9.0;
  const cardW = (totalW - cardGap * 2) / 3;
  prompts.forEach((p, i) => {
    const x = 0.5 + i * (cardW + cardGap);
    addCard(sFluency, x, cardY, cardW, cardH, { strip: C.ACCENT });
    sFluency.addText(p, {
      x: x + 0.10, y: cardY + 0.20, w: cardW - 0.20, h: cardH - 0.40,
      fontSize: 28, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  addCard(sFluency, 1.5, cardY + cardH + 0.15, 7.0, 0.55, { strip: C.PRIMARY });
  sFluency.addText("Whisper the answer. Then write it.", {
    x: 1.7, y: cardY + cardH + 0.22, w: 6.6, h: 0.50,
    fontSize: 20, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning that turning a shape does not change what it is.",
    [
      "I can name a shape no matter how it is turned.",
      "I can find shapes on a sports field.",
      "I can use a feature word to describe a shape.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Triangle in 4 different orientations
  const sIDo = pres.addSlide();
  addTopBar(sIDo, STAGE_COLORS["2"]);
  addStageBadge(sIDo, 2, "I Do");
  addTitle(sIDo, "Still a Triangle", { color: STAGE_COLORS["2"] });

  // 4 triangles in a row, different orientations.
  const triCardX = 0.5;
  const triCardY = CONTENT_TOP + 0.10;
  const triCardW = 9.0;
  const triCardH = SAFE_BOTTOM - triCardY - 0.10;
  addCard(sIDo, triCardX, triCardY, triCardW, triCardH, { strip: STAGE_COLORS["2"] });

  const triSize = 1.50;
  const triSlots = [0.95, 3.00, 5.05, 7.10];  // x-positions for 4 triangles
  const triRotations = [0, 90, 180, 270];
  const triLabels = ["points up", "points right", "points down", "points left"];
  const triTopY = triCardY + 0.45;
  triSlots.forEach((tx, i) => {
    drawShape(sIDo, "triangle", tx, triTopY, triSize, triSize, C.SECONDARY, null, { rotate: triRotations[i] });
    sIDo.addText(triLabels[i], {
      x: tx - 0.15, y: triTopY + triSize + 0.15, w: triSize + 0.30, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", margin: 0,
    });
  });

  // Caption strip across the bottom of the card
  addTextOnShape(sIDo, "All 4 are triangles. 3 sides, 3 corners — still a triangle.", {
    x: triCardX + 0.30, y: triCardY + triCardH - 0.65, w: triCardW - 0.60, h: 0.50, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

  addFooter(sIDo, FOOTER);
  sIDo.addNotes(NOTES_IDO);
  runSlideDiagnostics(sIDo, pres, { respectSafeBottom: false });

  // Slides 7-8: CFU with reveal — Is this still a square? (square rotated 45°)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Is This Still a Square?", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Diamond shape — same as a square rotated 45 degrees
      drawShape(s, "diamond", 4.00, CONTENT_TOP + 0.20, 2.0, 2.30, C.PRIMARY);

      // Prompt
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
      addTextOnShape(slide, "Thumbs UP. 4 equal sides + 4 corners. Still a square.", {
        x: 0.5, y: 4.45, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 9-10: We Do with reveal — Basketball court
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Find the Shapes on the Basketball Court", { color: STAGE_COLORS["3"] });

      // Big court visual
      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      // Court drawn in centre
      const courtX = 1.30;
      const courtY = CONTENT_TOP + 0.40;
      const courtW = 7.40;
      const courtH = 2.60;
      drawBasketballCourt(s, courtX, courtY, courtW, courtH);

      // Prompt below
      s.addText("Whisper to your partner: name one shape you can see.", {
        x: 0.7, y: 4.45, w: 8.6, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      // Reveal labels — small chips placed inside the court so they don't
      // overlap the prompt at the bottom of the slide.
      const courtX = 1.30;
      const courtY = CONTENT_TOP + 0.40;
      const courtW = 7.40;
      const courtH = 2.60;
      // Whole court label (top-right corner, just inside the court)
      addTextOnShape(slide, "rectangle = court", {
        x: courtX + courtW - 1.95, y: courtY + 0.10, w: 1.80, h: 0.30, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true });
      // Centre circle label (inside the court, above the centre circle)
      addTextOnShape(slide, "circle", {
        x: courtX + courtW / 2 - 0.40, y: courtY + 0.10, w: 0.80, h: 0.30, rectRadius: 0.06,
        fill: { color: C.ACCENT },
      }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true });
      // Left key label (inside the left key)
      addTextOnShape(slide, "rectangle", {
        x: courtX + 0.05, y: courtY + 0.10, w: 1.10, h: 0.30, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true });
      // Right key label (inside the right key)
      addTextOnShape(slide, "rectangle", {
        x: courtX + courtW * 0.18 + 0.10, y: courtY + 0.10, w: 1.10, h: 0.30, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 11: You Do — Soccer field, find the shapes.
  const sYouDo = pres.addSlide();
  addTopBar(sYouDo, STAGE_COLORS["4"]);
  addStageBadge(sYouDo, 4, "You Do");
  addTitle(sYouDo, "Your Turn: Find the Shapes", { color: STAGE_COLORS["4"] });

  // Left card: soccer field mockup
  addCard(sYouDo, 0.5, CONTENT_TOP, 5.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["4"] });

  // Soccer field — outer rectangle + centre circle + 2 small rectangles
  const fieldX = 0.80;
  const fieldY = CONTENT_TOP + 0.30;
  const fieldW = 4.90;
  const fieldH = 3.10;
  // Field
  sYouDo.addShape("rect", {
    x: fieldX, y: fieldY, w: fieldW, h: fieldH,
    fill: { color: "5CAA5C" },
    line: { color: C.WHITE, width: 3 },
  });
  // Half line
  sYouDo.addShape("line", {
    x: fieldX + fieldW / 2, y: fieldY, w: 0, h: fieldH,
    line: { color: C.WHITE, width: 2 },
  });
  // Centre circle
  const cR = Math.min(fieldW, fieldH) * 0.14;
  sYouDo.addShape("ellipse", {
    x: fieldX + fieldW / 2 - cR, y: fieldY + fieldH / 2 - cR, w: cR * 2, h: cR * 2,
    fill: { color: "5CAA5C" },
    line: { color: C.WHITE, width: 2 },
  });
  // Penalty rectangles (one each end)
  const penW = fieldW * 0.16;
  const penH = fieldH * 0.55;
  sYouDo.addShape("rect", {
    x: fieldX, y: fieldY + (fieldH - penH) / 2, w: penW, h: penH,
    fill: { color: "5CAA5C" },
    line: { color: C.WHITE, width: 2 },
  });
  sYouDo.addShape("rect", {
    x: fieldX + fieldW - penW, y: fieldY + (fieldH - penH) / 2, w: penW, h: penH,
    fill: { color: "5CAA5C" },
    line: { color: C.WHITE, width: 2 },
  });

  // Right card: prompt
  addCard(sYouDo, 6.2, CONTENT_TOP, 3.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
  sYouDo.addText([
    { text: "On your whiteboard:", options: { fontSize: 20, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Write the names of the shapes you can see.", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Find at least 2 shapes.", options: { fontSize: 16, color: C.MUTED, italic: true } },
  ], {
    x: 6.40, y: CONTENT_TOP + 0.20, w: 2.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sYouDo, FOOTER);
  sYouDo.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(sYouDo, pres, { respectSafeBottom: false });

  // Slide 12: Exit Ticket — Same or different shape?
  const sExit = pres.addSlide();
  sExit.background = { color: C.BG_CARD };
  sExit.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.ASSESS || C.ALERT } });
  addBadge(sExit, "Exit Ticket", { color: C.ASSESS || C.ALERT });
  addTitle(sExit, "Same Shape or Different?", { color: C.ASSESS || C.ALERT });

  // Left card: two shapes — square upright, square rotated (diamond)
  addCard(sExit, 0.5, CONTENT_TOP, 5.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ASSESS || C.ALERT });

  drawShape(sExit, "square", 1.10, CONTENT_TOP + 0.65, 1.70, 1.70, C.PRIMARY);
  sExit.addText("Shape A", {
    x: 1.10, y: CONTENT_TOP + 2.50, w: 1.70, h: 0.30,
    fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
  });

  drawShape(sExit, "diamond", 3.50, CONTENT_TOP + 0.50, 1.95, 2.10, C.PRIMARY);
  sExit.addText("Shape B", {
    x: 3.45, y: CONTENT_TOP + 2.80, w: 2.05, h: 0.30,
    fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
  });

  // Right card: prompt
  addCard(sExit, 6.2, CONTENT_TOP, 3.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
  sExit.addText([
    { text: "On your whiteboard:", options: { fontSize: 20, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Are A and B the same shape?", options: { fontSize: 20, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Write same OR different.", options: { fontSize: 22, color: C.CHARCOAL, bold: true } },
  ], {
    x: 6.40, y: CONTENT_TOP + 0.20, w: 2.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sExit, FOOTER);
  sExit.addNotes(NOTES_EXIT);
  runSlideDiagnostics(sExit, pres, { respectSafeBottom: false });

  // Slide 13: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: name one shape you saw on the basketball court.",
      scItems: [
        "I can name a shape no matter how it is turned.",
        "I can find shapes on a sports field.",
        "I can use a feature word to describe a shape.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 14: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "SHA_Lesson5_Same_Shape_Different_Position.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Worksheet — colour the shapes on a basketball court.
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Look at the basketball court. Colour the rectangles. Circle the circle.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addTipBox(doc,
      "Trace each shape with your finger first. Use one colour for the rectangles. Draw a ring around the circle.",
      y, { color: C.ACCENT });

    // Court drawing
    const cx = 60;
    const cy = y + 20;
    const cw = 480;
    const ch = 240;
    // Court outer
    doc.lineWidth(2.5).strokeColor("#333333").fillColor("#FFFFFF")
      .rect(cx, cy, cw, ch).fillAndStroke();
    // Half line
    doc.moveTo(cx + cw / 2, cy).lineTo(cx + cw / 2, cy + ch).stroke();
    // Centre circle
    doc.circle(cx + cw / 2, cy + ch / 2, 32).stroke();
    // Left key
    const kw = cw * 0.18;
    const kh = ch * 0.45;
    doc.rect(cx, cy + (ch - kh) / 2, kw, kh).stroke();
    // Right key
    doc.rect(cx + cw - kw, cy + (ch - kh) / 2, kw, kh).stroke();

    // Below: writing prompt
    const promptY = cy + ch + 30;
    doc.fillColor("#" + C.PRIMARY).fontSize(13).font("Sans-Bold")
      .text("How many rectangles did you find?", 60, promptY);
    doc.lineWidth(1).strokeColor("#333333")
      .moveTo(60, promptY + 50).lineTo(540, promptY + 50).stroke();

    addPdfFooter(doc, `Lesson ${SESSION} | Find the Shapes | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Where the rectangles and circle are on the court.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addSectionHeading(doc, "Rectangles", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "Three rectangles. The whole court is one rectangle. Each end of the court has a small rectangle (the key).",
      y);
    y = addSectionHeading(doc, "Circle", y, { color: C.PRIMARY });
    y = addBodyText(doc, "One circle. In the middle of the court (the centre circle).", y);
    y = addTipBox(doc,
      "The half-court line is a line, not a 2D shape on its own. It splits the court rectangle into two smaller rectangles.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 5 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
