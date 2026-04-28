"use strict";

// 2D Shapes Unit — Lesson 3: Opposite and Parallel Sides
// Year 2 Numeracy | Variant 2 (Slate & Copper)
// VC2M2SP01 - opposite and parallel sides
// Daily Review: Place value - importance of zero (30 vs 3)
// Fluency: Number formation / place value (write 47, 74)

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

const SESSION = 3;
const TOTAL = 5;
const UNIT_TITLE = "2D Shapes";
const FOOTER = `2D Shapes | Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`;
const OUT_DIR = "output/SHA_Lesson3_Opposite_and_Parallel";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 3 Find Parallel Sides",
  "Look at each shape. Colour one pair of parallel sides.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 3 Answer Key",
  "Where the parallel sides are on each shape.");

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
  else if (shapeKey === "parallelogram") type = "parallelogram";
  slide.addShape(type, {
    x, y, w, h,
    fill: { color: fillColor },
    line: { color: lineColor, width: 2.5 },
  });
}

// Highlight a single edge between two points with a thick coloured line.
function highlightEdge(slide, x1, y1, x2, y2, color) {
  slide.addShape("line", {
    x: x1, y: y1, w: x2 - x1, h: y2 - y1,
    line: { color, width: 8 },
  });
}

// Highlight the rectangle's four sides as two parallel pairs.
// Top + bottom = colour A. Left + right = colour B.
function highlightRectSides(slide, x, y, w, h, opts) {
  const o = opts || {};
  const colorA = o.colorA || "D64545";  // top + bottom
  const colorB = o.colorB || "1976D2";  // left + right
  if (o.showTopBottom !== false) {
    highlightEdge(slide, x, y, x + w, y, colorA);              // top
    highlightEdge(slide, x, y + h, x + w, y + h, colorA);      // bottom
  }
  if (o.showLeftRight !== false) {
    highlightEdge(slide, x, y, x, y + h, colorB);              // left
    highlightEdge(slide, x + w, y, x + w, y + h, colorB);      // right
  }
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Yesterday we sorted shapes by straight or curved sides.
- Today we use two new words: opposite and parallel.
- These help us describe sides on a shape.

DO:
- Have a rectangle card and a hexagon card ready.
- Settle students on the floor before clicking past the title.

TEACHER NOTES:
Lesson 3 of 5. New vocabulary: opposite, parallel. Concrete first - opposite means "on the other side". Parallel means "go the same way and never meet". Use the long-thin pencil image as a parallel cue.

WATCH FOR:
- Students who already use "opposite" - link to "across from".

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Two numbers. 30 and 3.
- The 0 in 30 is doing a job. It tells us there are no ones.
- Without the 0, the 3 is just three.
- Whisper to your partner: which number is bigger, 30 or 3?

DO:
- Display 30 and 3 in big numerals.
- Allow 10 seconds for whisper.
- Cold call before clicking the reveal.

TEACHER NOTES:
Daily Review focus: place value - importance of zero. The zero is a place holder. Without it, the digit moves.

WATCH FOR:
- Students who say 3 because "3 is more digits" - prompt: "Look at the place. The 3 in 30 is in the tens place."

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check. 30 is bigger.
- The zero holds the ones place. It says: no ones, but 3 tens.
- 30 is three tens. 3 is just three ones.

DO:
- Click to reveal.
- Point to the 3 in 30 and the empty ones place.

TEACHER NOTES:
Tick and fix. Note any students who chose 3 - they need place value support.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Place value writing.
- I will say a number. You write it on your whiteboard. Quick and neat.
- "Forty-seven."
- "Seventy-four."
- "Ninety."

DO:
- Say each number clearly.
- Allow 10 seconds per number.
- Walk and check digit order and formation.

TEACHER NOTES:
Brisk place value writing. Focus on tens digit FIRST, then ones. The "ninety" prompt checks the zero from the daily review - students should write 90, not 9.

WATCH FOR:
- Students who write 74 instead of 47 - tens-ones reversal is the most common error.
- Students who write 9 instead of 90 - reteach with the daily review zero example.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning about opposite and parallel sides.
- Opposite means on the other side.
- Parallel means the sides go the same way and never meet.
- Let us read the success criteria.

DO:
- Choral read the LI and each SC.
- Hold a rectangle card and trace top and bottom: "opposite sides".
- Trace top and bottom again: "parallel - they go the same way".

TEACHER NOTES:
Two new words today. Build them concretely. SC1 is reachable for all. SC2 is the core target. SC3 needs students to spot parallel sides on a new shape.

WATCH FOR:
- Students who repeat "opposite" and "parallel" - language is forming.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. This is a rectangle.
- The top and the bottom are opposite sides. They are across from each other.
- Look - they go the same way. They will never meet. They are parallel.
- The left and the right are also opposite. Also parallel. They go the same way too.
- A rectangle has two pairs of parallel sides.

DO:
- Trace the top side, then the bottom side, saying "opposite, parallel".
- Trace the left side, then the right side, saying "opposite, parallel".
- Repeat: "two pairs of parallel sides".

TEACHER NOTES:
First exposure to "parallel". The colour-coded edges show the two pairs - red across from red (top/bottom), blue across from blue (left/right). Students do not need to count yet - the goal is the language and the picture.

MISCONCEPTIONS:
- Misconception: Students think parallel means "the same length".
  Why: They notice the matching colours and assume same colour = same length.
  Impact: They reject parallel when sides are different lengths (eg. trapezium).
  Quick correction: "Parallel is about direction. The sides go the same way. They can be different lengths."

WATCH FOR:
- Students who repeat "parallel" while pointing - language is forming.
- Students who say "they are equal" - accept and reframe to "they go the same way".

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Look at the two lines.
- Are they parallel? Thumbs up for yes. Thumbs down for no.

DO:
- Display two parallel horizontal lines.
- Wait 5 seconds.
- Say: "Parallel? Show me with thumbs."

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Are these lines parallel? Thumbs up if yes."
- Scan for: thumbs UP. The two horizontal lines never meet.
PROCEED: If 80% show thumbs up, click to reveal and move to We Do.
PIVOT: Most likely misconception - students think parallel needs same length.
- Reteach: "Parallel is about direction. Both lines point the same way. They never meet, even if you make them longer."
- Re-check: Show same prompt again.

TEACHER NOTES:
The two lines have different lengths but go the same way - tests for the misconception that parallel means equal.

WATCH FOR:
- Confident thumbs up - secure with direction over length.
- Mixed - reteach with finger trace and the "extend the line" prompt.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. This shape is a trapezium.
- A trapezium has only ONE pair of parallel sides.
- Find them with me. Trace the top, then the bottom.
- Top and bottom go the same way - they are parallel.
- The two slanted sides are NOT parallel. They would meet if we made them longer.

DO:
- Trace the top, then the bottom, saying "parallel".
- Trace the left slant, then the right slant, saying "not parallel - they would meet".
- On their whiteboards: "Write: which sides are parallel?"

TEACHER NOTES:
Trapezium is chosen because it has only ONE parallel pair. This contrasts with the rectangle's two pairs and prevents the rule "all opposite sides are parallel" from forming.

WATCH FOR:
- Students who say "top and bottom" - secure.
- Students who say "all sides" - prompt: "Trace the slanted sides. Would they meet?"

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check. The top and the bottom are parallel.
- The two slanted sides are not parallel.
- A trapezium has just one pair of parallel sides.

DO:
- Click to reveal. The top and bottom are highlighted red.
- Repeat: "one pair of parallel sides".

TEACHER NOTES:
Reveals the parallel pair. Students who got it can lead a partner who is unsure.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Look at the hexagon.
- A hexagon has parallel sides too.
- On your whiteboard, draw the hexagon and colour ONE pair of parallel sides.

DO:
- Allow 90 seconds.
- Walk and scan whiteboards.
- Cold call 1-2 students to share which pair they coloured.

TEACHER NOTES:
A regular hexagon has 3 pairs of parallel sides. Any one is correct. Students need to spot a pair that "goes the same way" - opposite sides on a regular hexagon match.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Look at the rectangle from the I Do. Point to one pair of parallel sides.
EXTENDING PROMPT:
- Task: How many pairs of parallel sides can you find on the hexagon? Try to find more than one.

WATCH FOR:
- Students who colour the top and bottom flat sides - on track.
- Students who colour two sides that meet at a corner - prompt: "Trace them. Would they meet?"

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. Look at the rectangle.
- On your whiteboard, draw two arrows on a pair of parallel sides.

DO:
- Allow 60 seconds.
- Collect whiteboards or photo for records.

TEACHER NOTES:
Exit ticket assesses SC2 (core). Students should mark a pair of opposite parallel sides - either top + bottom OR left + right is correct. Both pairs is a stretch answer.

WATCH FOR:
- Students who mark either pair - on track.
- Students who mark adjacent sides - they confused opposite with next-to.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today we used two new words. Opposite. Parallel.
- Opposite means across from.
- Parallel means going the same way and never meeting.
- Show me thumbs up if you can find a pair of parallel sides.
- Turn and tell your partner: how many pairs of parallel sides on a rectangle?

DO:
- Read the success criteria.
- Use thumbs up / sideways / down.
- Cold call 1-2 students for the share.

TEACHER NOTES:
Self-assessment data informs Lesson 4 grouping.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- The find-parallel sheet is for after our session.

DO:
- Print the Lesson 3 Find Parallel Sides sheet, one per student.
- Have shape cards (rectangle, trapezium, hexagon) ready.

TEACHER NOTES:
One printed student resource. Mini-whiteboards do most of the active work.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 3: Opposite and Parallel Sides",
    `Year 2 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slides 2-3: Daily Review with reveal — 30 vs 3 (zero is a place holder)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Which Is Bigger?", { color: C.ACCENT });

      // Two big numeral cards
      const cardY = CONTENT_TOP + 0.30;
      const cardH = 3.10;
      const cardW = 4.0;
      addCard(s, 0.7, cardY, cardW, cardH, { strip: C.ACCENT });
      s.addText("30", {
        x: 0.7, y: cardY + 0.20, w: cardW, h: cardH - 0.40,
        fontSize: 200, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      addCard(s, 5.3, cardY, cardW, cardH, { strip: C.ACCENT });
      s.addText("3", {
        x: 5.3, y: cardY + 0.20, w: cardW, h: cardH - 0.40,
        fontSize: 200, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      // Prompt below
      s.addText("Show me on your whiteboard: 30 or 3?", {
        x: 0.7, y: cardY + cardH + 0.10, w: 8.6, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "30 is bigger. The 0 means no ones, but 3 tens.", {
        x: 0.5, y: 4.45, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Place value writing (3 prompts: 47, 74, 90)
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Write the Number", { color: C.ACCENT });

  const prompts = ["forty-seven", "seventy-four", "ninety"];
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
      fontSize: 32, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });

  addCard(sFluency, 1.5, cardY + cardH + 0.20, 7.0, 0.55, { strip: C.PRIMARY });
  sFluency.addText("Write the number on your whiteboard. Tens first.", {
    x: 1.7, y: cardY + cardH + 0.22, w: 6.6, h: 0.50,
    fontSize: 20, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning about opposite and parallel sides on a shape.",
    [
      "I can point to opposite sides on a shape.",
      "I can point to parallel sides on a shape.",
      "I can find a pair of parallel sides on a new shape.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Rectangle with two pairs of parallel sides highlighted
  const sIDo = pres.addSlide();
  addTopBar(sIDo, STAGE_COLORS["2"]);
  addStageBadge(sIDo, 2, "I Do");
  addTitle(sIDo, "A Rectangle Has 2 Pairs of Parallel Sides", { color: STAGE_COLORS["2"] });

  // Left card: explanation
  addCard(sIDo, 0.5, CONTENT_TOP, 4.30, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["2"] });
  sIDo.addText([
    { text: "Opposite", options: { fontSize: 24, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "= across from each other.", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Parallel", options: { fontSize: 24, bold: true, color: C.ACCENT, breakLine: true } },
    { text: "= go the same way.", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
    { text: "Never meet.", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Top and bottom: parallel.", options: { fontSize: 16, color: "B02525", bold: true, breakLine: true } },
    { text: "Left and right: parallel.", options: { fontSize: 16, color: "1976D2", bold: true } },
  ], {
    x: 0.70, y: CONTENT_TOP + 0.15, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.30,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  // Right: Rectangle with edges highlighted in two colours
  const rectX = 5.50;
  const rectY = CONTENT_TOP + 0.50;
  const rectW = 3.80;
  const rectH = 2.70;
  drawShape(sIDo, "rectangle", rectX, rectY, rectW, rectH, C.SECONDARY);
  highlightRectSides(sIDo, rectX, rectY, rectW, rectH, {
    colorA: "D64545",  // top + bottom red
    colorB: "1976D2",  // left + right blue
  });

  addFooter(sIDo, FOOTER);
  sIDo.addNotes(NOTES_IDO);
  runSlideDiagnostics(sIDo, pres, { respectSafeBottom: false });

  // Slides 7-8: CFU with reveal — Are these lines parallel?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Are These Lines Parallel?", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Two horizontal lines, different lengths, same direction (parallel)
      s.addShape("line", {
        x: 1.5, y: CONTENT_TOP + 0.90, w: 6.5, h: 0,
        line: { color: "1976D2", width: 10 },
      });
      s.addShape("line", {
        x: 1.5, y: CONTENT_TOP + 2.20, w: 4.5, h: 0,
        line: { color: "1976D2", width: 10 },
      });

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
      addTextOnShape(slide, "Thumbs UP. Both lines go the same way. They never meet.", {
        x: 0.5, y: 4.45, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 9-10: We Do with reveal — Trapezium with one pair of parallel sides
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Find the Parallel Sides", { color: STAGE_COLORS["3"] });

      // Left card: trapezium visual
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      drawShape(s, "trapezoid", 1.20, CONTENT_TOP + 0.30, 3.10, 2.70, C.SECONDARY);
      s.addText("trapezium", {
        x: 0.6, y: CONTENT_TOP + 3.20, w: 4.3, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Right card: prompt
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Which 2 sides are parallel?", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Trace each side first.", options: { fontSize: 16, color: C.MUTED, italic: true } },
      ], {
        x: 5.4, y: CONTENT_TOP + 0.20, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      // PptxGenJS trapezoid: bottom is the long side, top is short, slants on left/right.
      // Highlight the top and bottom edges — those are the parallel pair.
      const tx = 1.20, ty = CONTENT_TOP + 0.30, tw = 3.10, th = 2.70;
      // top of trapezoid — by default the top is at ~25% in from each side at y=ty.
      // Use the bounding box top edge (drawing the top side as a horizontal line at the top of the bbox).
      // The PptxGenJS trapezoid's top side runs from approximately (tx + tw*0.25, ty) to (tx + tw*0.75, ty).
      slide.addShape("line", {
        x: tx + tw * 0.25, y: ty, w: tw * 0.5, h: 0,
        line: { color: "D64545", width: 8 },
      });
      // bottom side: full bottom edge
      slide.addShape("line", {
        x: tx, y: ty + th, w: tw, h: 0,
        line: { color: "D64545", width: 8 },
      });
      // Reveal banner
      addTextOnShape(slide, "Top and bottom are parallel. One pair only.", {
        x: 0.5, y: 4.45, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 11: You Do — Hexagon, find one pair of parallel sides
  workedExSlide(pres, 4, "You Do", "Your Turn: Find a Parallel Pair",
    [
      "First: Look at the hexagon.",
      "",
      "Next: Find two sides that go the same way.",
      "",
      "Then: On your whiteboard, draw the hexagon.",
      "Colour one pair of parallel sides.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      // Hexagon on the right
      drawShape(slide, "hexagon",
        lg.rightX + 0.30, lg.panelTopPadded + 0.10,
        lg.rightW - 0.60, 3.55,
        C.SECONDARY);
    }
  );

  // Slide 12: Exit Ticket — Rectangle, mark a parallel pair
  const sExit = pres.addSlide();
  sExit.background = { color: C.BG_CARD };
  sExit.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.ASSESS || C.ALERT } });
  addBadge(sExit, "Exit Ticket", { color: C.ASSESS || C.ALERT });
  addTitle(sExit, "Show What You Know", { color: C.ASSESS || C.ALERT });

  // Left: rectangle
  addCard(sExit, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ASSESS || C.ALERT });
  drawShape(sExit, "rectangle", 0.90, CONTENT_TOP + 0.70, 3.70, 2.10, C.SECONDARY);

  // Right: prompt
  addCard(sExit, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
  sExit.addText([
    { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Draw the rectangle.", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Colour 2 sides", options: { fontSize: 24, color: C.CHARCOAL, bold: true, breakLine: true } },
    { text: "that are parallel.", options: { fontSize: 24, color: C.CHARCOAL, bold: true } },
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
      reflectionPrompt: "Turn and tell: how many pairs of parallel sides on a rectangle?",
      scItems: [
        "I can point to opposite sides on a shape.",
        "I can point to parallel sides on a shape.",
        "I can find a pair of parallel sides on a new shape.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 14: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "SHA_Lesson3_Opposite_and_Parallel.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Worksheet — colour one pair of parallel sides on each shape
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "For each shape, colour one pair of parallel sides.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addTipBox(doc,
      "Parallel sides go the same way. They never meet, even if you make them longer. Trace each side with your finger first.",
      y, { color: C.ACCENT });

    const pageX = 50;
    const pageW = 512;
    const cellW = pageW / 2;
    const cellH = 215;
    const startY = y + 8;

    function drawCell(cx, cy, label, drawShapeFn) {
      doc.lineWidth(1).strokeColor("#" + C.MUTED)
        .rect(cx, cy, cellW - 10, cellH).stroke();
      doc.fillColor("#" + C.PRIMARY).fontSize(13).font("Sans-Bold")
        .text(label, cx + 12, cy + 10);
      drawShapeFn(cx + 16, cy + 36, cellW - 42, cellH - 60);
    }

    // Each shape is unfilled (just outline) so children can colour the parallel pair.
    function drawRectanglePdf(x, y, w, h) {
      doc.lineWidth(2.5).strokeColor("#333333").fillColor("#FFFFFF");
      doc.rect(x + 20, y + 20, w - 40, h - 40).fillAndStroke();
    }
    function drawTrapeziumPdf(x, y, w, h) {
      doc.lineWidth(2.5).strokeColor("#333333").fillColor("#FFFFFF");
      const padX = 10;
      const slant = 35;
      doc.moveTo(x + padX + slant, y + 20)
        .lineTo(x + w - padX - slant, y + 20)
        .lineTo(x + w - padX, y + h - 20)
        .lineTo(x + padX, y + h - 20).closePath().fillAndStroke();
    }
    function drawHexagonPdf(x, y, w, h) {
      doc.lineWidth(2.5).strokeColor("#333333").fillColor("#FFFFFF");
      const cx = x + w / 2; const cy = y + h / 2;
      const rx = (w - 30) / 2; const ry = (h - 30) / 2;
      // Horizontal hexagon (flat top/bottom).
      doc.moveTo(cx - rx * 0.5, cy - ry)
        .lineTo(cx + rx * 0.5, cy - ry)
        .lineTo(cx + rx, cy)
        .lineTo(cx + rx * 0.5, cy + ry)
        .lineTo(cx - rx * 0.5, cy + ry)
        .lineTo(cx - rx, cy).closePath().fillAndStroke();
    }
    function drawSquarePdf(x, y, w, h) {
      doc.lineWidth(2.5).strokeColor("#333333").fillColor("#FFFFFF");
      const side = Math.min(w, h) - 40;
      const sx = x + (w - side) / 2;
      const sy = y + (h - side) / 2;
      doc.rect(sx, sy, side, side).fillAndStroke();
    }

    drawCell(pageX,         startY,             "Rectangle", drawRectanglePdf);
    drawCell(pageX + cellW, startY,             "Trapezium", drawTrapeziumPdf);
    drawCell(pageX,         startY + cellH + 8, "Hexagon",   drawHexagonPdf);
    drawCell(pageX + cellW, startY + cellH + 8, "Square",    drawSquarePdf);

    addPdfFooter(doc, `Lesson ${SESSION} | Find Parallel Sides | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Where the parallel sides are on each shape.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addSectionHeading(doc, "Rectangle", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Two pairs of parallel sides. Top and bottom is one pair. Left and right is the other pair.", y);
    y = addSectionHeading(doc, "Trapezium", y, { color: C.PRIMARY });
    y = addBodyText(doc, "One pair of parallel sides. The top and the bottom go the same way. The two slanted sides are not parallel.", y);
    y = addSectionHeading(doc, "Hexagon", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Three pairs of parallel sides. Each side has an opposite side that goes the same way.", y);
    y = addSectionHeading(doc, "Square", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Two pairs of parallel sides. The same pattern as a rectangle.", y);
    y = addTipBox(doc,
      "Parallel sides do not need to be the same length. They just need to go the same way and never meet.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
