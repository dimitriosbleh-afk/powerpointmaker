"use strict";

// Geometry and Data Unit (Year 5/6 Numeracy) - Lesson 4: Tessellation art.
// Year 6 VC2M6SP03 - recognise and use COMBINATIONS of transformations to create tessellations.
// Students build a final piece for the classroom display.
// Daily Review: Adding and subtracting fractions and decimals.
// Fluency: Division (missing factor / harder).

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

const SESSION = 4;
const TOTAL = 8;
const UNIT_TITLE = "Geometry and Data: Shapes, Tessellations, Data";
const FOOTER = `Geometry and Data | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/GD_Lesson4_Tessellation_Art";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 4 Tessellation Art Template",
  "Cut, slide, paste, repeat. Build a tessellation for the classroom display.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 4 Teacher Notes",
  "Step-by-step guide for the classroom display and assessment look-fors.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// =====================================================================
// Teacher notes
// =====================================================================

const NOTES_TITLE = `SAY:
- Last shape lesson of the week.
- Today you design your own tessellation. It goes up on the wall.
- We will use transformations - translations, reflections, rotations - to make a single shape that tiles forever.

DO:
- Settle.
- Have kinder square paper, scissors and tape ready per student.

TEACHER NOTES:
Lesson 4 of 8. This is the creative payoff lesson. Students learn a simple Escher-style modify-a-square technique, then build a tile and repeat it across paper for a classroom display.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- One kinder square paper per student.
- Scissors per student.
- Sticky tape per pair.
- One Tessellation Art Template per student.
- A4 paper to glue/draw the final pattern (one per student).

DO:
- Set up tape and scissors at desk.
- Stack of A4 paper at front for the final display work.

TEACHER NOTES:
The final display can be glued onto coloured backing paper. Students often want to make their tessellation more elaborate than the time allows - keep them focused on getting a clean repeating tile first.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review.
- Add or subtract on your whiteboard.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Adding and subtracting fractions and decimals - second day. Mixed denominators on one prompt today.

WATCH FOR:
- Students who find a common denominator on the fraction question - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 1/2 + 1/4 = 2/4 + 1/4 = 3/4.
- 1.25 + 0.6 = 1.85.
- 2.4 - 0.75 = 1.65.

DO:
- Click to reveal.

TEACHER NOTES:
Note any student who tried to add unlike denominators directly.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Division. Some are missing-factor style.
- Whisper-answer then write.

DO:
- Display the three prompts.
- 25 seconds.

TEACHER NOTES:
Day four of mult/div fluency. Less scaffold today - missing factor stretches students who are secure.

WATCH FOR:
- Students who solve 6 x ___ = 54 by thinking in 6s - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 81 divided by 9 = 9.
- 6 x ___ = 54  ->  9.
- 144 divided by 12 = 12.

DO:
- Click to reveal.

TEACHER NOTES:
Brisk reveal.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Yesterday we learned WHICH shapes tessellate.
- Today we use transformations to make our OWN tessellating shape.
- Watch - I start with a square (which tessellates), cut a bit off, then TRANSLATE that bit to the opposite side.
- The new shape still tessellates. And it can be anything - a fish, a bird, a leaf.

DO:
- Demonstrate with a kinder square cut and shifted.
- Show the resulting shape can repeat by translation.

TEACHER NOTES:
This launch is the WOW moment. Once students see the cut/shift trick they are hooked. Make sure the cut piece is taped back exactly on the opposite side, not flipped.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to use combinations of transformations to design a tessellation pattern.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 - identify the transformation used. SC2 - create one tile and one repeat. SC3 - design a full pattern for display.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Quick vocabulary check from yesterday.
- Translation - slide.
- Reflection - flip.
- Rotation - turn.
- And tessellation - shapes that fit with no gaps and no overlaps.

DO:
- Have students whisper each word back.

TEACHER NOTES:
This is retrieval rather than new vocabulary. Anchor before the I Do.

[General: Vocabulary Recall | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_IDO_TECHNIQUE = `SAY:
- Watch closely.
- Step 1 - start with a square.
- Step 2 - draw and cut a shape from the LEFT side.
- Step 3 - slide that exact piece to the RIGHT side and tape it on.
- Step 4 - now my new tile is no longer a square. But it still tessellates - because what I took off one side, I put on the other.
- Translation. Slide. Same direction every time.

DO:
- Cut a kinder square and demonstrate.
- Show the tile fits next to a copy of itself.

TEACHER NOTES:
This is the core technique. The cut must come from the LEFT, move RIGHT without flipping. Translation only - no rotation, no reflection. If students rotate, the shape will not tile.

WATCH FOR:
- Students who tape the cut piece without flipping - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_TOPBOTTOM = `SAY:
- One more move.
- I can do the same with the top and bottom.
- Cut a shape from the TOP, slide it to the BOTTOM, tape on.
- Now my tile has wiggly edges all around. But it still tessellates - top meets bottom, left meets right.

DO:
- Demonstrate the top/bottom modification on top of the left/right one.

TEACHER NOTES:
Combining both is what makes it look more art-like. The tile is no longer a square but the transformation logic guarantees it tiles.

WATCH FOR:
- Students who can name the two translations - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- I cut a piece from the LEFT side of a square. To make sure the new tile tessellates, where do I put the piece?
- A: tape it to the top.
- B: tape it to the right side.
- C: tape it to the bottom.

DO:
- Display the prompt.
- 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me on three, two, one, show."
- Scan for: B.
PROCEED: If 80% have B, click to reveal and move to We Do.
PIVOT: Most likely misconception - students think any side works.
- Reteach: "Translation means slide in one direction. Left and right are opposite sides - the cut and the paste must be on opposite sides."
- Re-check: Show a piece cut from the top. Where does it go? Bottom.

TEACHER NOTES:
B - opposite side. The cut and the tape go on opposite sides for translation.

WATCH FOR:
- Students who write B - secure.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO = `SAY:
- With me, with your own kinder square. Step by step.
- Step 1: hold your square. Decide where to cut on the LEFT side.
- Step 2: cut a curve or zigzag from top to bottom of the left side. Do NOT cut off a corner.
- Step 3: slide the piece to the RIGHT side. Tape it on. Same height.
- Step 4: look at your tile. Hold it next to your partner's and see how the tiles could meet.

DO:
- Mirror with your own square.
- Pause between each step.

TEACHER NOTES:
Walk between desks to check the cut piece is taped at the same height. If a student tapes higher or lower, the tile will not fit cleanly.

WATCH FOR:
- Students whose tape line matches the cut height exactly - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Hold your tile up. Notice your wiggly left and right edges match exactly.
- That is the rule. The shape you take from one side is the shape you give to the other.

DO:
- Click to reveal.

TEACHER NOTES:
This is also the moment to remind students that this is the easiest tessellation method - translation only. Reflection and rotation methods are extension.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn. Take the Tessellation Art Template.
- Use your tile from the kinder square.
- Trace it onto the A4 paper.
- Slide it one tile-width across. Trace again.
- Keep going. Fill the page.
- Then decorate. Make each tile look like a bird, a fish, a leaf or something you can imagine.

DO:
- Distribute A4 paper and the template guide.
- Set the timer for the main work.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Trace and slide only. Skip the decoration. Or stick with a 2x2 grid instead of filling the page.
- Extra Notes: Sit with these students. Help them get the slide direction consistent.
EXTENDING PROMPT:
- Task: Use ROTATION as well as translation. Cut from the left and rotate the piece 90 degrees before taping it to the top. Show how the new tile tessellates around a centre.
- Extra Notes: This produces an Escher-style pattern.

TEACHER NOTES:
Different content from We Do - students apply the technique to make their own art. The classroom display is the goal.

WATCH FOR:
- Students who line up their tiles edge to edge with no gaps - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Q1: Which transformation does the basic cut-and-slide method use?
- Q2: If I cut from the TOP of my square, where must I tape the piece for it to tessellate?

DO:
- Display the prompt.
- 60 seconds.

TEACHER NOTES:
Exit ticket assesses SC2. Look for "translation" and "the bottom".

WATCH FOR:
- Students who answer both - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria for today.
- Show me thumbs up, sideways, or down for each one.
- Hold up your finished tile or partial display so the class can see.

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
Celebrate the work. Tomorrow we shift to data - collecting real data at the school boundary.

[General: Closing | VTLM 2.0: Reflection]`;

// =====================================================================
// Helpers
// =====================================================================

// Draw a tile with a wave/curve on one edge to suggest the modified-square technique.
function drawWavyTile(slide, x, y, w, h, opts) {
  const o = opts || {};
  const color = o.color || C.PRIMARY;
  // Outline rough rectangle with wavy left/right edges (approximation using small ticks).
  // Top edge
  slide.addShape("line", { x, y, w, h: 0, line: { color, width: 2.5 } });
  // Bottom edge
  slide.addShape("line", { x, y: y + h, w, h: 0, line: { color, width: 2.5 } });
  // Left edge: indent (3 small lines)
  const segs = 4;
  for (let i = 0; i < segs; i++) {
    const y1 = y + (h * i) / segs;
    const y2 = y + (h * (i + 1)) / segs;
    const indent = (i % 2 === 0) ? 0 : 0.20;
    slide.addShape("line", {
      x: x + indent, y: y1, w: 0, h: y2 - y1,
      line: { color, width: 2.5 },
    });
    if (i < segs - 1) {
      const nextIndent = ((i + 1) % 2 === 0) ? 0 : 0.20;
      slide.addShape("line", {
        x: x + Math.min(indent, nextIndent), y: y2, w: Math.abs(indent - nextIndent), h: 0,
        line: { color, width: 2.5 },
      });
    }
  }
  // Right edge: outdent matching indent (mirror)
  for (let i = 0; i < segs; i++) {
    const y1 = y + (h * i) / segs;
    const y2 = y + (h * (i + 1)) / segs;
    const outdent = (i % 2 === 0) ? 0 : 0.20;
    slide.addShape("line", {
      x: x + w + outdent, y: y1, w: 0, h: y2 - y1,
      line: { color, width: 2.5 },
    });
    if (i < segs - 1) {
      const nextOutdent = ((i + 1) % 2 === 0) ? 0 : 0.20;
      slide.addShape("line", {
        x: x + w + Math.min(outdent, nextOutdent), y: y2, w: Math.abs(outdent - nextOutdent), h: 0,
        line: { color, width: 2.5 },
      });
    }
  }
}

// =====================================================================
// Build
// =====================================================================

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 4: Tessellation art for the classroom display",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions and decimals",
      ["1/2 + 1/4", "1.25 + 0.6", "2.4 - 0.75"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3/4         1.85         1.65", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal
  withReveal(
    () => fluencySlide(pres, "Fluency: Division and missing factor",
      ["81 div 9", "6 x ? = 54", "144 div 12"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "9         9         12", {
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
    addTitle(s, "Make your own tessellating tile");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    // Four-step storyboard
    const stepY = CONTENT_TOP + 0.50;
    const stepW = 1.85;
    const gap = 0.20;
    const startX = (10 - (4 * stepW + 3 * gap)) / 2;

    const steps = [
      { label: "1. Start with a square", action: "square" },
      { label: "2. Cut a piece off LEFT", action: "cut" },
      { label: "3. Slide piece to RIGHT", action: "shifted" },
      { label: "4. Repeat across page", action: "repeat" },
    ];

    steps.forEach((step, i) => {
      const sx = startX + i * (stepW + gap);
      const sy = stepY;
      // Card
      addCard(s, sx, sy, stepW, 2.50, { strip: C.PRIMARY, fill: C.WHITE });
      // Mini illustration
      const illX = sx + 0.30;
      const illY = sy + 0.30;
      const illW = stepW - 0.60;
      if (step.action === "square") {
        s.addShape("rect", {
          x: illX, y: illY, w: illW, h: illW,
          fill: { color: C.BG_LIGHT }, line: { color: C.PRIMARY, width: 2 },
        });
      } else if (step.action === "cut") {
        s.addShape("rect", {
          x: illX, y: illY, w: illW, h: illW,
          fill: { color: C.BG_LIGHT }, line: { color: C.PRIMARY, width: 2 },
        });
        // Dashed cut line
        s.addShape("line", {
          x: illX + 0.15, y: illY + illW * 0.25, w: 0, h: illW * 0.50,
          line: { color: C.ALERT, width: 2, dashType: "dash" },
        });
        s.addText("cut", {
          x: illX - 0.05, y: illY + illW * 0.35, w: 0.50, h: 0.30,
          fontSize: 10, fontFace: FONT_B, color: C.ALERT, italic: true,
          align: "left", margin: 0,
        });
      } else if (step.action === "shifted") {
        drawWavyTile(s, illX, illY, illW, illW, { color: C.PRIMARY });
        s.addText("slide", {
          x: illX + 0.30, y: illY + illW + 0.05, w: 0.80, h: 0.25,
          fontSize: 10, fontFace: FONT_B, color: C.ACCENT, italic: true,
          align: "left", margin: 0,
        });
      } else if (step.action === "repeat") {
        // Three small wavy tiles in a row
        const ts = (illW - 0.20) / 3;
        for (let k = 0; k < 3; k++) {
          drawWavyTile(s, illX + k * ts, illY + 0.10, ts - 0.05, ts, { color: C.SUCCESS });
        }
      }
      // Label
      s.addText(step.label, {
        x: sx + 0.05, y: sy + 2.05, w: stepW - 0.10, h: 0.40,
        fontSize: 11, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to use combinations of transformations to design a tessellation pattern.",
    [
      "I can identify the transformation used when a tile is repeated.",
      "I can create one tile and one repeat using the cut and slide technique.",
      "I can design a full tessellation pattern for display using my own tile.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Vocabulary recall
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Vocabulary Recall", { color: C.PRIMARY });
    addTitle(s, "Translation, reflection, rotation, tessellation");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const items = [
      { word: "Translation", simple: "slide", color: C.PRIMARY },
      { word: "Reflection",  simple: "flip",  color: C.ACCENT },
      { word: "Rotation",    simple: "turn",  color: C.ALERT },
      { word: "Tessellation", simple: "no gaps, no overlaps", color: C.SUCCESS },
    ];
    const cellW = 2.10, gap = 0.15;
    const totalW = cellW * 4 + gap * 3;
    const startX = (10 - totalW) / 2;
    const cardY = CONTENT_TOP + 0.50;
    const cardH = 2.20;

    items.forEach((it, i) => {
      const cx = startX + i * (cellW + gap);
      addCard(s, cx, cardY, cellW, cardH, { strip: it.color, fill: C.WHITE });
      addTextOnShape(s, it.word, {
        x: cx + 0.10, y: cardY + 0.25, w: cellW - 0.20, h: 0.60, rectRadius: 0.10,
        fill: { color: it.color },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(it.simple, {
        x: cx, y: cardY + 1.05, w: cellW, h: 0.90,
        fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 10: I Do - cut and slide (left to right)
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Cut from the LEFT, slide to the RIGHT", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Three-frame storyboard
    const fY = CONTENT_TOP + 0.40;
    const fSize = 1.50;
    const fGap = 0.50;
    const frames = 3;
    const totalW = frames * fSize + (frames - 1) * (fGap + 0.5);
    const startX = (10 - totalW) / 2;

    // Frame 1: square with dashed cut
    const x1 = startX;
    s.addShape("rect", {
      x: x1, y: fY, w: fSize, h: fSize,
      fill: { color: C.BG_LIGHT }, line: { color: C.PRIMARY, width: 2 },
    });
    s.addShape("line", {
      x: x1 + 0.20, y: fY + 0.25, w: 0, h: fSize - 0.50,
      line: { color: C.ALERT, width: 2.5, dashType: "dash" },
    });
    s.addText("Step 1: cut the left edge", {
      x: x1 - 0.20, y: fY + fSize + 0.10, w: fSize + 0.40, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", margin: 0,
    });

    // Arrow
    s.addShape("line", {
      x: x1 + fSize + 0.10, y: fY + fSize / 2,
      w: 0.70, h: 0,
      line: { color: C.ALERT, width: 2.5, endArrowType: "triangle" },
    });
    s.addText("slide", {
      x: x1 + fSize + 0.05, y: fY + fSize / 2 - 0.35, w: 0.80, h: 0.30,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, italic: true,
      align: "center", margin: 0,
    });

    // Frame 2: piece on the right + remaining square missing piece
    const x2 = x1 + fSize + 0.80;
    drawWavyTile(s, x2, fY, fSize - 0.20, fSize, { color: C.PRIMARY });
    s.addText("Step 2: tape on the right", {
      x: x2 - 0.20, y: fY + fSize + 0.10, w: fSize + 0.40, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", margin: 0,
    });

    // Arrow
    s.addShape("line", {
      x: x2 + fSize + 0.10, y: fY + fSize / 2,
      w: 0.70, h: 0,
      line: { color: C.SUCCESS, width: 2.5, endArrowType: "triangle" },
    });
    s.addText("repeat", {
      x: x2 + fSize + 0.05, y: fY + fSize / 2 - 0.35, w: 0.80, h: 0.30,
      fontSize: 11, fontFace: FONT_B, color: C.SUCCESS, italic: true,
      align: "center", margin: 0,
    });

    // Frame 3: two wavy tiles side by side
    const x3 = x2 + fSize + 0.80;
    drawWavyTile(s, x3, fY, (fSize - 0.20) / 2, fSize, { color: C.SUCCESS });
    drawWavyTile(s, x3 + (fSize - 0.20) / 2, fY, (fSize - 0.20) / 2, fSize, { color: C.SUCCESS });
    s.addText("Step 3: tile across", {
      x: x3 - 0.20, y: fY + fSize + 0.10, w: fSize + 0.40, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", margin: 0,
    });

    s.addText("Transformation used: translation (slide).", {
      x: 0.7, y: SAFE_BOTTOM - 0.50, w: 8.6, h: 0.35,
      fontSize: 14, fontFace: FONT_H, color: stageColor, bold: true,
      align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_TECHNIQUE);
  })();

  // Slide 11: I Do - also top and bottom
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Cut top, slide to bottom (the second move)", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Left: square with two cut marks
    const x0 = 1.4;
    const y0 = CONTENT_TOP + 0.45;
    const sz = 2.20;
    s.addShape("rect", {
      x: x0, y: y0, w: sz, h: sz,
      fill: { color: C.BG_LIGHT }, line: { color: C.PRIMARY, width: 2 },
    });
    // Vertical dashed cut on left
    s.addShape("line", {
      x: x0 + 0.20, y: y0 + 0.30, w: 0, h: sz - 0.60,
      line: { color: C.ALERT, width: 2, dashType: "dash" },
    });
    // Horizontal dashed cut on top
    s.addShape("line", {
      x: x0 + 0.30, y: y0 + 0.20, w: sz - 0.60, h: 0,
      line: { color: C.ACCENT, width: 2, dashType: "dash" },
    });
    s.addText("Cut from left -> slide right", {
      x: x0 - 0.30, y: y0 + sz + 0.15, w: sz + 0.60, h: 0.30,
      fontSize: 12, fontFace: FONT_B, color: C.ALERT, italic: true,
      align: "center", margin: 0,
    });
    s.addText("Cut from top -> slide bottom", {
      x: x0 - 0.30, y: y0 + sz + 0.50, w: sz + 0.60, h: 0.30,
      fontSize: 12, fontFace: FONT_B, color: C.ACCENT, italic: true,
      align: "center", margin: 0,
    });

    // Right: explanation
    addCard(s, 5.20, CONTENT_TOP + 0.30, 4.30, 2.80, { strip: stageColor, fill: C.WHITE });
    s.addText([
      { text: "Combine two translations:", options: { fontSize: 16, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "1. Cut shape from LEFT, tape on RIGHT.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "2. Cut shape from TOP, tape on BOTTOM.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Result: a tile with wiggly edges all around. Still tessellates.", options: { fontSize: 13, color: C.SUCCESS, bold: true, italic: true } },
    ], {
      x: 5.40, y: CONTENT_TOP + 0.50, w: 4.00, h: 2.40,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_TOPBOTTOM);
  })();

  // Slides 12-13: CFU + reveal
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Where does the cut piece go?", { color: C.ALERT });
      addTextOnShape(s, "CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      s.addText([
        { text: "You cut a piece from the LEFT side of a square.", options: { fontSize: 18, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Where do you tape it so the new tile tessellates?", options: { fontSize: 16, color: C.CHARCOAL, italic: true, breakLine: true } },
        { text: "", options: { fontSize: 12, breakLine: true } },
        { text: "A.   Top", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "B.   Right side", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "C.   Bottom", options: { fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 1.0, y: CONTENT_TOP + 0.40, w: 8.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.80,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "B. Right side. Translation moves opposite to opposite.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do (with reveal showing the finished tile pattern)
  withReveal(
    () => {
      const s = pres.addSlide();
      const stageColor = STAGE_COLORS["3"];
      addTopBar(s, stageColor);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Together: make your tile", { color: stageColor });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

      s.addText([
        { text: "Step 1: ", options: { fontSize: 16, color: C.ALERT, bold: true, breakLine: false } },
        { text: "hold your kinder square.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Step 2: ", options: { fontSize: 16, color: C.ALERT, bold: true, breakLine: false } },
        { text: "cut a wave or zigzag from the LEFT side.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Step 3: ", options: { fontSize: 16, color: C.ALERT, bold: true, breakLine: false } },
        { text: "SLIDE the piece to the RIGHT.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Step 4: ", options: { fontSize: 16, color: C.ALERT, bold: true, breakLine: false } },
        { text: "tape it on at the SAME HEIGHT.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Step 5: ", options: { fontSize: 16, color: C.ALERT, bold: true, breakLine: false } },
        { text: "hold up your tile.", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      // Right: blank tile preview
      const tx = 5.40;
      const ty = CONTENT_TOP + 0.30;
      const tw = 4.0;
      addCard(s, tx, ty, tw, 2.80, { strip: stageColor, fill: C.WHITE });
      const tileSize = 2.0;
      drawWavyTile(s, tx + (tw - tileSize) / 2, ty + (2.80 - tileSize) / 2, tileSize - 0.20, tileSize, { color: stageColor });
      s.addText("Your tile will look like this", {
        x: tx, y: ty + 2.80 - 0.30, w: tw, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      // Reveal: show the tile repeated three times
      const ry = 4.55;
      const tileSize = 0.50;
      const startX = 5 - (3 * tileSize) / 2;
      // Actually use the success bar approach
      addTextOnShape(slide, "Hold tiles edge to edge with a partner. They fit.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // Slide 16: You Do - create the display piece
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["4"];
    addTopBar(s, stageColor);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: build the display", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    s.addText([
      { text: "First: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "trace your tile onto A4 paper.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Next: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "slide it across. Trace again. No gaps.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Then: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "fill the page in rows.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Finally: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "decorate each tile. Bird, fish, leaf, lizard.", options: { fontSize: 15, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    // Right: a 2x3 tessellation preview
    const px = 5.30, py = CONTENT_TOP + 0.20;
    const pw = 4.10, ph = 3.10;
    addCard(s, px, py, pw, ph, { strip: stageColor, fill: C.WHITE });
    const tileW = (pw - 0.40) / 3;
    const tileH = (ph - 0.40) / 2;
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < 3; c++) {
        drawWavyTile(s,
          px + 0.20 + c * tileW,
          py + 0.20 + r * tileH,
          tileW - 0.05, tileH,
          { color: c % 2 ? C.PRIMARY : C.ACCENT });
      }
    }
    s.addText("Use the Tessellation Art Template for tracing.", {
      x: px, y: SAFE_BOTTOM - 0.40, w: pw, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Which transformation does the basic cut-and-slide method use?",
      "If I cut from the TOP of my square, where must I tape the piece for the tile to tessellate?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Hold up your tile or your partial display. Who used a second cut as well?",
      scItems: [
        "I can identify the transformation used when a tile is repeated.",
        "I can create one tile and one repeat using the cut and slide technique.",
        "I can design a full tessellation pattern for display using my own tile.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "GD_Lesson4_Tessellation_Art.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Make a tessellating tile. Trace it across the page. Decorate.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Translation = slide opposite to opposite. " +
      "If you cut from the LEFT side of your square, tape the piece on the RIGHT at the SAME height. " +
      "If you cut from the TOP, tape the piece on the BOTTOM at the SAME position across.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Step 1 — Make your tile", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Use your kinder square. Cut a wave or zigzag from the LEFT side, top to bottom. Slide it to the RIGHT side and tape it on.", y);
    y = addBodyText(doc, "Optional second cut: cut from the TOP and tape on the BOTTOM.", y);
    y = addWriteLine(doc, "Transformation used:  _______________________", y);

    y = addSectionHeading(doc, "Step 2 — Trace your tile across the page", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Place your tile on a piece of A4 paper. Trace it. Then SLIDE the tile one width to the right and trace again. Keep going.", y);
    y = addBodyText(doc, "Check: no gaps, no overlaps. The right edge of one tile matches the left edge of the next.", y);

    y = addSectionHeading(doc, "Step 3 — Fill the page", y, { color: C.PRIMARY });
    y = addBodyText(doc, "After the first row, move down one tile height. Trace another row. Keep going until the page is full.", y);

    y = addSectionHeading(doc, "Step 4 — Decorate (for the classroom display)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Each tile should look the same. Add eyes, scales, lines, colour - turn each tile into a bird, fish, leaf, lizard, or anything you can imagine.", y);
    y = addWriteLine(doc, "What did your tile become?  _______________________", y);
    y = addBodyText(doc, "Title your piece. Sign your name. Hand in for the classroom display.", y);

    y = addSectionHeading(doc, "Stretch — Use rotation as well", y, { color: C.ACCENT });
    y = addBodyText(doc, "Instead of sliding the cut piece to the OPPOSITE side, ROTATE it 90 degrees first and tape it to the NEXT side. The tile now repeats around a centre point. Sketch your design here.", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Tessellation Art Template | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Step-by-step guide and assessment look-fors for the display.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Materials per student", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Kinder square paper, scissors, sticky tape, A4 paper for the final tessellation, coloured pencils or markers.", y);

    y = addSectionHeading(doc, "Common mistakes to spot", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1. Cut piece taped on the WRONG side (flipped or rotated). The tile will not tessellate.", y);
    y = addBodyText(doc, "2. Cut piece taped at a DIFFERENT height. The tile leaves gaps when repeated.", y);
    y = addBodyText(doc, "3. Corner of the square cut off. The shape is no longer a closed figure.", y);

    y = addSectionHeading(doc, "Assessment look-fors", y, { color: C.PRIMARY });
    y = addBodyText(doc, "SC1 - student can name the transformation as translation (slide).", y);
    y = addBodyText(doc, "SC2 - student's tile has been traced at least twice, edge to edge, with no gap or overlap.", y);
    y = addBodyText(doc, "SC3 - student has filled an A4 sheet with their tessellation and decorated it for the display.", y);

    y = addSectionHeading(doc, "Display ideas", y, { color: C.ACCENT });
    y = addBodyText(doc, "Glue each finished tessellation onto coloured backing paper. Arrange as a large patchwork on the wall. Add a title strip: \"Year 5/6 Tessellation Display.\"", y);
    y = addBodyText(doc, "Optional - label each piece with the student's name and the animal/object their tile became.", y);

    y = addTipBox(doc,
      "Watch the time. Some students will want to perfect their tile and never tile the page. Set a soft deadline: 'Your tile must be on the A4 paper at least twice by [time].' Decoration can finish next session.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Teacher Notes | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
