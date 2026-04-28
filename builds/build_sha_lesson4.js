"use strict";

// 2D Shapes Unit — Lesson 4: Sorting Shapes by Features
// Year 2 Numeracy | Variant 2 (Slate & Copper)
// VC2M2SP01 - sort shapes by number of sides, equal sides
// Daily Review: Place value - non-standard renaming (3 tens 12 ones)
// Fluency: Patterns

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
  addBaseTenBlocks,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const TOTAL = 5;
const UNIT_TITLE = "2D Shapes";
const FOOTER = `2D Shapes | Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`;
const OUT_DIR = "output/SHA_Lesson4_Sorting_Shapes";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 4 Sort by Sides",
  "Cut out the shapes. Sort them into two groups: 4 sides or not 4 sides.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 4 Answer Key",
  "Which shapes have 4 sides and which do not.");

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

function drawShapeLabelled(slide, shapeKey, x, y, w, h, fill, label) {
  const labelH = 0.30;
  drawShape(slide, shapeKey, x, y, w, h - labelH - 0.05, fill);
  if (label) {
    slide.addText(label, {
      x, y: y + h - labelH, w, h: labelH,
      fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "middle", margin: 0,
    });
  }
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- We have learned about sides, corners, straight, curved, and parallel.
- Today we use those words to sort shapes into groups.
- We are sorting by features.

DO:
- Have shape cards ready: triangle, square, rectangle, hexagon, pentagon, circle.
- Settle students on the floor before clicking past the title.

TEACHER NOTES:
Lesson 4 of 5. Pulls together the language from lessons 1-3 into a sorting task. Students choose ONE feature and sort by it.

WATCH FOR:
- Students who can name several features - cold call them in CFU.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the blocks.
- I see 3 tens. I see 12 ones.
- Whisper to your partner: what number is this?

DO:
- Display 3 tens rods and 12 ones cubes.
- Allow 20 seconds for whisper.
- Cold call before clicking the reveal.

TEACHER NOTES:
Daily Review focus: place value - non-standard renaming. 12 ones is 1 ten and 2 ones, so 3 tens + 12 ones = 4 tens + 2 ones = 42.

WATCH FOR:
- Students who say 312 - they are reading the digits side by side; reteach with the regrouping.
- Students who say 42 - secure with renaming.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check. 3 tens and 12 ones.
- 12 ones is the same as 1 ten and 2 ones.
- So we have 4 tens and 2 ones.
- The number is 42.

DO:
- Click to reveal.
- Group 10 of the ones cubes into a circle to show "this group is one ten".

TEACHER NOTES:
Tick and fix. Note any students who said 312 - they need place value support before Year 3.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Number patterns.
- Look at the pattern. What comes next?
- Whisper to your partner. Then write it on your whiteboard.

DO:
- Display 3 patterns.
- Walk and check each one.

TEACHER NOTES:
Brisk patterns. Skip counts by 2, 5, 10. Students should recognise the jump in each pattern.

WATCH FOR:
- Students who write 7 for the 2s pattern - they added 1 instead of 2.
- Students who write 25 for the 10s pattern - they added 5 instead of 10.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to sort shapes using their features.
- A feature is something we can see on the shape - like number of sides.
- Let us read the success criteria.

DO:
- Choral read the LI and each SC.
- Hold up a square (4 sides, all equal) and a triangle (3 sides, equal).

TEACHER NOTES:
SC1 is reachable for all - just count and group. SC2 is the core target. SC3 needs students to choose their own sorting rule.

WATCH FOR:
- Students who repeat "feature" - new word; build it concretely.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. Six shapes here.
- I am going to sort them by number of sides.
- 4 sides on this side. NOT 4 sides on that side.
- Square has 4 sides - 4 sides group.
- Rectangle has 4 sides - 4 sides group.
- Triangle has 3 sides - not 4 sides.
- Pentagon has 5 sides - not 4 sides.
- Hexagon has 6 sides - not 4 sides.
- Circle has 1 curved side - not 4 sides.

DO:
- Move each shape into its column as you say its sides count.
- Repeat the rule: "If it has 4 sides, it goes left. If it does not, it goes right."

TEACHER NOTES:
I Do models the sorting routine: name the rule, check each shape, move it. The rule is the feature being used.

MISCONCEPTIONS:
- Misconception: Students sort by appearance, not feature.
  Why: They group "all the round things" without thinking about sides.
  Impact: They put oval and circle together but miss that the rule was sides.
  Quick correction: "What is our rule? Count the sides first."

WATCH FOR:
- Students who copy the rule before checking - secure with the routine.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Look at the trapezium.
- Our rule is: 4 sides or not 4 sides.
- Where does the trapezium go? Point to the side of the room.
- Left for 4 sides. Right for not 4 sides.

DO:
- Display the trapezium.
- Wait 5 seconds.
- Say: "Left or right? Point now."

CFU CHECKPOINT:
Technique: Point to the Group
Script:
- Say: "Left side for 4 sides. Right side for not 4 sides. Point now."
- Scan for: students pointing LEFT (trapezium has 4 sides).
PROCEED: If 80% point left, click to reveal.
PIVOT: Most likely misconception - students think trapezium has 3 sides because of the slanted look.
- Reteach: "Trace each side. 1, 2, 3, 4. It has 4 sides."
- Re-check: Show again, ask "How many sides?"

TEACHER NOTES:
The trapezium is included to challenge the rule - it does not look like a square or a rectangle, but it still has 4 sides.

WATCH FOR:
- Confident left points - secure with the count rule.
- Right points - prompt: "Trace and count again."

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Four shapes here.
- New rule: all sides EQUAL or NOT all equal.
- A square has 4 sides, all the same length. All equal.
- A rectangle has 4 sides, but the long sides are longer. NOT all equal.
- With your partner, whisper which group each shape belongs to.

DO:
- Allow 30 seconds for partner whisper.
- Cold call before clicking the reveal.

TEACHER NOTES:
A new rule - equal sides. The shapes shown: square (equal), triangle (equilateral - equal), rectangle (not equal), trapezium (not equal). The reveal places them in two columns.

WATCH FOR:
- Pairs who put the rectangle in "equal" - prompt: "Trace the long side. Trace the short side. Same length?"

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check. Square: all equal. Triangle: all equal.
- Rectangle: not equal. Trapezium: not equal.

DO:
- Click to reveal.
- Trace one side of the rectangle and the long side - "see the difference?"

TEACHER NOTES:
Reveal places the shapes in their groups. The triangle here is shown as equilateral. Note that not all triangles have equal sides.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Four shapes here. Pick a rule of your own.
- It could be: 4 sides or not, has a curve or not, has parallel sides or not.
- On your whiteboard, write your rule. Then list which shapes go in each group.

DO:
- Allow 2 minutes.
- Walk and scan whiteboards.
- Cold call 2 students to share their rule.

TEACHER NOTES:
Students choose their own sorting rule. This stretches them to use the feature words from lessons 1-3.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the same rule from the I Do: 4 sides or not 4 sides.
EXTENDING PROMPT:
- Task: Sort the shapes using TWO rules. For example: 4 sides AND all equal.

WATCH FOR:
- Students who choose "has a curve" or "4 sides" - secure rules.
- Students who pick a rule that does not work for the shapes - prompt: "Does your rule split the shapes into two groups?"

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. Look at the four shapes.
- On your whiteboard, write the name of the shape with all sides EQUAL.

DO:
- Allow 60 seconds.
- Collect whiteboards or photo for records.

TEACHER NOTES:
Exit ticket assesses SC2 (core). The shape with all sides equal is the square. The triangle here is NOT equilateral - it has unequal sides.

WATCH FOR:
- Students who write "square" - on track.
- Students who write "triangle" - they assumed all triangles have equal sides; reteach with side tracing.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today we sorted shapes by their features.
- Number of sides is one feature. Equal sides is another.
- Show me thumbs up if you can sort shapes using a rule.
- Turn and tell your partner: name two features of a square.

DO:
- Read the success criteria.
- Use thumbs up / sideways / down.
- Cold call 1-2 students for the share.

TEACHER NOTES:
Self-assessment data informs Lesson 5 grouping.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- The cut-and-sort sheet is for after our session.

DO:
- Print the Lesson 4 Sort by Sides sheet, one per student.
- Have scissors and the printed sheet ready.

TEACHER NOTES:
One printed student resource - cut and sort. Mini-whiteboards are used during the You Do.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 4: Sorting Shapes by Features",
    `Year 2 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slides 2-3: Daily Review with reveal — Non-standard renaming (3 tens 12 ones)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "What Number Is This?", { color: C.ACCENT });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

      // Base-ten blocks: 3 tens and 12 ones
      // unit ~0.20 to fit in card. tens rods = 0.20 wide x 2.0 tall
      const blocksY = CONTENT_TOP + 0.50;
      addBaseTenBlocks(s, 1.0, blocksY, 0, 3, 12, { unit: 0.20, color: C.SECONDARY });

      // Labels
      s.addText("3 tens", {
        x: 1.0, y: blocksY + 2.20, w: 1.4, h: 0.30,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      s.addText("12 ones", {
        x: 2.6, y: blocksY + 2.20, w: 6.5, h: 0.30,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "left", margin: 0,
      });

      // Prompt
      s.addText("On your whiteboard: write the number.", {
        x: 0.7, y: 4.45, w: 8.6, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "12 ones = 1 ten + 2 ones.   So 3 tens + 12 ones = 42.", {
        x: 0.5, y: 4.45, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — 3 number patterns
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "What Comes Next?", { color: C.ACCENT });

  const patterns = ["2, 4, 6, ___", "10, 20, 30, ___", "5, 10, 15, ___"];
  const cardY = CONTENT_TOP + 0.30;
  const cardH = 2.80;
  const cardGap = 0.20;
  const totalW = 9.0;
  const cardW = (totalW - cardGap * 2) / 3;
  patterns.forEach((p, i) => {
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
  sFluency.addText("Whisper the next number. Then write it.", {
    x: 1.7, y: cardY + cardH + 0.22, w: 6.6, h: 0.50,
    fontSize: 20, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to sort shapes using their features.",
    [
      "I can find shapes that have the same number of sides.",
      "I can find shapes with all sides the same length.",
      "I can sort shapes into two groups using one feature.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Sort 6 shapes by "4 sides or not"
  const sIDo = pres.addSlide();
  addTopBar(sIDo, STAGE_COLORS["2"]);
  addStageBadge(sIDo, 2, "I Do");
  addTitle(sIDo, "Sort by Number of Sides", { color: STAGE_COLORS["2"] });

  // Two columns: 4 SIDES (left) and NOT 4 SIDES (right)
  const colY = CONTENT_TOP + 0.10;
  const colH = SAFE_BOTTOM - colY - 0.10;
  const colW = 4.30;
  // Left column
  addCard(sIDo, 0.5, colY, colW, colH, { strip: C.PRIMARY });
  addTextOnShape(sIDo, "4 sides", {
    x: 0.5, y: colY, w: colW, h: 0.45,
    fill: { color: C.PRIMARY },
  }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
  // Right column
  addCard(sIDo, 5.20, colY, colW, colH, { strip: C.ACCENT });
  addTextOnShape(sIDo, "not 4 sides", {
    x: 5.20, y: colY, w: colW, h: 0.45,
    fill: { color: C.ACCENT },
  }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

  // Place shapes inside the columns. Two rows fit between header (y=colY+0.45)
  // and column bottom (y=colY+colH=5.00). Row pitch ~1.45, shape h=1.20, label h=0.25.
  const shapeRowY1 = colY + 0.55;
  const shapeRowY2 = colY + 2.00;
  const sw = 1.10;
  const shapeH = 1.20;  // includes the 0.25 label band drawn by drawShapeLabelled
  // Left column: square, rectangle (both have 4 sides)
  drawShapeLabelled(sIDo, "square",    0.95, shapeRowY1, sw, shapeH, C.SECONDARY, "square");
  drawShapeLabelled(sIDo, "rectangle", 2.55, shapeRowY1, sw + 0.30, shapeH, C.SECONDARY, "rectangle");
  drawShapeLabelled(sIDo, "trapezoid", 1.75, shapeRowY2, sw + 0.30, shapeH, C.SECONDARY, "trapezium");

  // Right column: triangle, pentagon, hexagon, circle
  drawShapeLabelled(sIDo, "triangle", 5.55, shapeRowY1, sw, shapeH, C.SECONDARY, "triangle");
  drawShapeLabelled(sIDo, "pentagon", 7.25, shapeRowY1, sw, shapeH, C.SECONDARY, "pentagon");
  drawShapeLabelled(sIDo, "hexagon",  5.55, shapeRowY2, sw, shapeH, C.SECONDARY, "hexagon");
  drawShapeLabelled(sIDo, "circle",   7.25, shapeRowY2, sw, shapeH, C.SECONDARY, "circle");

  addFooter(sIDo, FOOTER);
  sIDo.addNotes(NOTES_IDO);
  runSlideDiagnostics(sIDo, pres, { respectSafeBottom: false });

  // Slides 7-8: CFU with reveal — Where does the trapezium go?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Where Does the Trapezium Go?", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Big trapezium centred
      drawShape(s, "trapezoid", 3.50, CONTENT_TOP + 0.30, 3.0, 2.20, C.SECONDARY);

      // Prompt
      s.addText("Point left for 4 sides. Point right for not 4 sides.", {
        x: 0.7, y: 4.45, w: 8.6, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "LEFT — 4 sides. Trace and count: 1, 2, 3, 4.", {
        x: 0.5, y: 4.45, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 9-10: We Do with reveal — Sort by all-equal sides
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "All Sides Equal — Or Not?", { color: STAGE_COLORS["3"] });

      // Top strip: 4 shapes
      const stripY = CONTENT_TOP + 0.10;
      const stripH = 1.40;
      addCard(s, 0.5, stripY, 9.0, stripH, { strip: STAGE_COLORS["3"] });

      const sw = 1.10;
      const shapeY = stripY + 0.10;
      drawShape(s, "square",    1.00, shapeY, sw, sw, C.PRIMARY);
      s.addText("square", { x: 0.7, y: shapeY + sw + 0.02, w: sw + 0.60, h: 0.25,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

      drawShape(s, "triangle",  3.20, shapeY, sw, sw, C.SECONDARY);
      s.addText("triangle", { x: 2.90, y: shapeY + sw + 0.02, w: sw + 0.60, h: 0.25,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

      drawShape(s, "rectangle", 5.30, shapeY, 1.50, sw, C.PRIMARY);
      s.addText("rectangle", { x: 5.10, y: shapeY + sw + 0.02, w: 1.90, h: 0.25,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

      drawShape(s, "trapezoid", 7.50, shapeY, sw + 0.20, sw, C.SECONDARY);
      s.addText("trapezium", { x: 7.20, y: shapeY + sw + 0.02, w: sw + 0.80, h: 0.25,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

      // Two empty sort columns
      const sortY = stripY + stripH + 0.20;
      const sortH = SAFE_BOTTOM - sortY - 0.10;
      const colW = 4.30;
      addCard(s, 0.5, sortY, colW, sortH, { strip: C.PRIMARY });
      addTextOnShape(s, "ALL sides equal", {
        x: 0.5, y: sortY, w: colW, h: 0.40,
        fill: { color: C.PRIMARY },
      }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true });
      addCard(s, 5.20, sortY, colW, sortH, { strip: C.ACCENT });
      addTextOnShape(s, "NOT all equal", {
        x: 5.20, y: sortY, w: colW, h: 0.40,
        fill: { color: C.ACCENT },
      }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      const stripY = CONTENT_TOP + 0.10;
      const stripH = 1.40;
      const sortY = stripY + stripH + 0.20;
      const inColY = sortY + 0.55;
      const sw = 1.0;
      // Left column: square + triangle (both equal)
      drawShape(slide, "square",    1.10, inColY, sw, sw, C.PRIMARY);
      drawShape(slide, "triangle",  2.80, inColY, sw, sw, C.SECONDARY);
      // Right column: rectangle + trapezium (not equal)
      drawShape(slide, "rectangle", 5.80, inColY, 1.40, sw, C.PRIMARY);
      drawShape(slide, "trapezoid", 7.65, inColY, sw + 0.20, sw, C.SECONDARY);
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 11: You Do — Sort 4 shapes with own rule
  const sYouDo = pres.addSlide();
  addTopBar(sYouDo, STAGE_COLORS["4"]);
  addStageBadge(sYouDo, 4, "You Do");
  addTitle(sYouDo, "Your Turn: Sort with Your Own Rule", { color: STAGE_COLORS["4"] });

  // Left card: instructions
  addCard(sYouDo, 0.5, CONTENT_TOP, 4.30, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["4"] });
  sYouDo.addText([
    { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Pick a rule.", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Could be:", options: { fontSize: 14, color: C.MUTED, italic: true, breakLine: true } },
    { text: "4 sides or not", options: { fontSize: 14, color: C.MUTED, breakLine: true } },
    { text: "has a curve or not", options: { fontSize: 14, color: C.MUTED, breakLine: true } },
    { text: "all equal or not", options: { fontSize: 14, color: C.MUTED, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Write your rule.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
    { text: "List the shapes in each group.", options: { fontSize: 16, color: C.CHARCOAL } },
  ], {
    x: 0.70, y: CONTENT_TOP + 0.15, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.30,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  // Right card: 4 labelled shapes
  addCard(sYouDo, 5.0, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
  const ydW = 1.55;
  const ydX1 = 5.30;
  const ydX2 = 7.45;
  const ydY1 = CONTENT_TOP + 0.20;
  const ydY2 = CONTENT_TOP + 1.95;
  drawShape(sYouDo, "pentagon", ydX1, ydY1, ydW, ydW, C.SECONDARY);
  sYouDo.addText("pentagon", { x: ydX1 - 0.20, y: ydY1 + ydW + 0.02, w: ydW + 0.40, h: 0.25,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  drawShape(sYouDo, "circle",   ydX2, ydY1, ydW, ydW, C.ACCENT);
  sYouDo.addText("circle", { x: ydX2 - 0.20, y: ydY1 + ydW + 0.02, w: ydW + 0.40, h: 0.25,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  drawShape(sYouDo, "rectangle", ydX1, ydY2, ydW + 0.25, ydW * 0.75, C.PRIMARY);
  sYouDo.addText("rectangle", { x: ydX1 - 0.20, y: ydY2 + ydW + 0.02, w: ydW + 0.40, h: 0.25,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  drawShape(sYouDo, "hexagon",  ydX2, ydY2, ydW, ydW, C.SECONDARY);
  sYouDo.addText("hexagon", { x: ydX2 - 0.20, y: ydY2 + ydW + 0.02, w: ydW + 0.40, h: 0.25,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  addFooter(sYouDo, FOOTER);
  sYouDo.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(sYouDo, pres, { respectSafeBottom: false });

  // Slide 12: Exit Ticket — Which shape has all sides equal?
  const sExit = pres.addSlide();
  sExit.background = { color: C.BG_CARD };
  sExit.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.ASSESS || C.ALERT } });
  addBadge(sExit, "Exit Ticket", { color: C.ASSESS || C.ALERT });
  addTitle(sExit, "Which Shape Has All Sides Equal?", { color: C.ASSESS || C.ALERT });

  // Left: 4 labelled shapes — square (equal), rectangle (not), scalene-style triangle (not), trapezium (not)
  addCard(sExit, 0.5, CONTENT_TOP, 5.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ASSESS || C.ALERT });
  const exW = 1.40;
  const exX1 = 0.85;
  const exX2 = 3.45;
  const exY1 = CONTENT_TOP + 0.20;
  const exY2 = CONTENT_TOP + 1.95;
  drawShape(sExit, "square",    exX1, exY1, exW, exW, C.PRIMARY);
  sExit.addText("square", { x: exX1 - 0.20, y: exY1 + exW + 0.02, w: exW + 0.40, h: 0.25,
    fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  drawShape(sExit, "rectangle", exX2, exY1, exW + 0.30, exW * 0.7, C.SECONDARY);
  sExit.addText("rectangle", { x: exX2 - 0.20, y: exY1 + exW + 0.02, w: exW + 0.70, h: 0.25,
    fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  drawShape(sExit, "trapezoid", exX1, exY2, exW + 0.20, exW * 0.8, C.SECONDARY);
  sExit.addText("trapezium", { x: exX1 - 0.20, y: exY2 + exW + 0.02, w: exW + 0.60, h: 0.25,
    fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  // Use a non-isosceles triangle effect — a wide right-angled triangle (rtTriangle) — to differentiate from equilateral.
  sExit.addShape("rtTriangle", {
    x: exX2, y: exY2, w: exW + 0.30, h: exW * 0.85,
    fill: { color: C.PRIMARY },
    line: { color: C.CHARCOAL, width: 2.5 },
  });
  sExit.addText("triangle", { x: exX2 - 0.20, y: exY2 + exW + 0.02, w: exW + 0.70, h: 0.25,
    fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

  // Right: prompt
  addCard(sExit, 6.2, CONTENT_TOP, 3.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
  sExit.addText([
    { text: "On your whiteboard:", options: { fontSize: 20, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Write the name of the shape with all sides equal.", options: { fontSize: 18, color: C.CHARCOAL } },
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
      reflectionPrompt: "Turn and tell: name two features of a square.",
      scItems: [
        "I can find shapes that have the same number of sides.",
        "I can find shapes with all sides the same length.",
        "I can sort shapes into two groups using one feature.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 14: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "SHA_Lesson4_Sorting_Shapes.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Worksheet — cut and sort by 4 sides or not
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Cut out the shapes. Sort them: 4 sides or not 4 sides.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addTipBox(doc,
      "Trace and count the sides of each shape before you cut. Then glue or place each shape in its group.",
      y, { color: C.ACCENT });

    // Two sort columns at the top
    const colY = y + 8;
    const colW = 240;
    const colH = 240;
    // Left column box
    doc.lineWidth(2).strokeColor("#" + C.PRIMARY).fillColor("#" + C.PRIMARY)
      .rect(50, colY, colW, 26).fillAndStroke();
    doc.fillColor("#FFFFFF").fontSize(14).font("Sans-Bold")
      .text("4 sides", 50, colY + 7, { width: colW, align: "center" });
    doc.fillColor("#FFFFFF");
    doc.lineWidth(1).strokeColor("#" + C.PRIMARY).fillColor("#FFFFFF")
      .rect(50, colY + 26, colW, colH - 26).fillAndStroke();

    // Right column box
    doc.lineWidth(2).strokeColor("#" + C.ACCENT).fillColor("#" + C.ACCENT)
      .rect(50 + colW + 32, colY, colW, 26).fillAndStroke();
    doc.fillColor("#FFFFFF").fontSize(14).font("Sans-Bold")
      .text("not 4 sides", 50 + colW + 32, colY + 7, { width: colW, align: "center" });
    doc.fillColor("#FFFFFF");
    doc.lineWidth(1).strokeColor("#" + C.ACCENT).fillColor("#FFFFFF")
      .rect(50 + colW + 32, colY + 26, colW, colH - 26).fillAndStroke();

    // Cut-out shapes section
    const cutY = colY + colH + 20;
    doc.fillColor("#" + C.PRIMARY).fontSize(13).font("Sans-Bold")
      .text("Cut these shapes:", 50, cutY);
    doc.fillColor("#333333");
    const cutShapeY = cutY + 22;
    const cutSize = 60;
    const cutGap = 18;
    // Drawing helpers (filled outlined)
    function drawTrianglePdf(cx, cy) {
      doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.SECONDARY);
      doc.moveTo(cx, cy - 30).lineTo(cx + 30, cy + 25).lineTo(cx - 30, cy + 25).closePath().fillAndStroke();
    }
    function drawSquarePdf(cx, cy) {
      doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.PRIMARY);
      doc.rect(cx - 28, cy - 28, 56, 56).fillAndStroke();
    }
    function drawCirclePdf(cx, cy) {
      doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.ACCENT);
      doc.circle(cx, cy, 30).fillAndStroke();
    }
    function drawHexagonPdf(cx, cy) {
      doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.SECONDARY);
      const rx = 30; const ry = 26;
      doc.moveTo(cx - rx * 0.5, cy - ry)
        .lineTo(cx + rx * 0.5, cy - ry)
        .lineTo(cx + rx, cy)
        .lineTo(cx + rx * 0.5, cy + ry)
        .lineTo(cx - rx * 0.5, cy + ry)
        .lineTo(cx - rx, cy).closePath().fillAndStroke();
    }
    function drawRectanglePdf(cx, cy) {
      doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.PRIMARY);
      doc.rect(cx - 36, cy - 22, 72, 44).fillAndStroke();
    }
    function drawTrapeziumPdf(cx, cy) {
      doc.lineWidth(1.6).strokeColor("#333333").fillColor("#" + C.SECONDARY);
      doc.moveTo(cx - 20, cy - 25).lineTo(cx + 20, cy - 25)
        .lineTo(cx + 35, cy + 25).lineTo(cx - 35, cy + 25).closePath().fillAndStroke();
    }
    const cutShapes = [
      drawTrianglePdf, drawSquarePdf, drawCirclePdf, drawHexagonPdf, drawRectanglePdf, drawTrapeziumPdf,
    ];
    cutShapes.forEach((fn, i) => {
      const cx = 80 + i * (cutSize + cutGap);
      // Dotted cut box
      doc.dash(3, { space: 3 }).lineWidth(0.8).strokeColor("#777777")
        .rect(cx - cutSize / 2, cutShapeY, cutSize, cutSize).stroke();
      doc.undash();
      fn(cx, cutShapeY + cutSize / 2);
    });

    addPdfFooter(doc, `Lesson ${SESSION} | Sort by Sides | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Which shapes have 4 sides and which do not.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addSectionHeading(doc, "4 sides", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Square, rectangle, trapezium.", y);
    y = addSectionHeading(doc, "Not 4 sides", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Triangle (3 sides), hexagon (6 sides), circle (1 curved side).", y);
    y = addTipBox(doc,
      "A trapezium does not look like a square or a rectangle, but it still has 4 sides. Always count the sides.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
