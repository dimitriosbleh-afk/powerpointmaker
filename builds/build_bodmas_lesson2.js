"use strict";

// BODMAS Unit (Year 5/6 Numeracy) - Lesson 2: Multiplication and Division come first.
// Builds directly on Lesson 1's "x before +". Uses arrays as the visual model.
// Daily Review: 2D shapes (sides and angles).
// Number Fluency: Adding 2-digit numbers.
// Cohort note: students are weaker on additive thinking. After x and ÷ steps
// we keep the addition piece friendly so the BODMAS reasoning is the load.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(7));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addArray,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 5;
const UNIT_TITLE = "BODMAS: Order of Operations";
const FOOTER = `BODMAS | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/BODMAS_Lesson2_Multiply_And_Divide_First";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PRACTICE_RES = makeSessionResource(SESSION,
  "Lesson 2 Multiply and Divide First Practice",
  "Six expressions with x or ÷ followed by + or -. Students underline the first step.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 2 Answer Key",
  "Teacher reference with worked answers and watch-fors.");
const RESOURCE_ITEMS = [PRACTICE_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes ----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to BODMAS.
- Yesterday we learned WHY we need a shared order, and we met the BODMAS rule.
- Today we are zooming in on the M and the D - multiplication and division - and how they always come before plus and minus.

DO:
- Have counters and the BODMAS reference card on each desk.
- Settle students before clicking on.

TEACHER NOTES:
Lesson 2 of 5. The threshold idea today is that x and ÷ are at the same level, and BOTH come before + and -. Use arrays as the visual model.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- You will use counters, your whiteboard, and a practice sheet later in the lesson.

DO:
- Have a pot of 24 counters per pair on each desk.
- Print one practice sheet per student.
- Print one answer key for yourself.

TEACHER NOTES:
One practice sheet and one answer key. The bulk of the lesson runs on counters, arrays and whiteboards.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are revising 2D shapes.
- Look at the shape on the right.
- On your whiteboard, name the shape and write how many sides it has.

DO:
- Display the prompt.
- 60 seconds.
- Walk and scan.

TEACHER NOTES:
2D shapes prior learning. The shape shown is a pentagon (5 sides). Watch for students confusing pentagon with hexagon.

WATCH FOR:
- Students who write hexagon - quick correction: "Count the sides one by one."

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check yours.
- This shape is a pentagon. It has 5 sides and 5 angles.
- Tick or fix.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Pent = five. Reinforce the prefix link.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Adding two-digit numbers again.
- Whisper-answer each one, then write it on your board.

DO:
- Display the three prompts.
- 45 seconds.

TEACHER NOTES:
Adding 2-digit numbers stays as the fluency focus all week. Students who finger-counted yesterday should start to chunk by tens today.

WATCH FOR:
- Students who add the tens first then the ones - secure strategy.
- Students still counting on - small group focus.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 27 + 31 = 58.
- 44 + 25 = 69.
- 38 + 17 = 55.
- Tick and fix.

DO:
- Click to reveal.

TEACHER NOTES:
38 + 17 needs regrouping. Note students who got 45 - they forgot to carry the ten.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Quick recap from yesterday. Look at 4 + 3 x 2.
- Show me on your fingers - 1 finger for plus first, 2 fingers for times first.
- Then write the answer on your board.

DO:
- Display the prompt.
- Show me on fingers.
- 30 seconds.

TEACHER NOTES:
The launch is a 30-second retrieval from yesterday. Looking for 2 fingers and answer 10. If most students show 1 finger or answer 14, slow down and rebuild the rule with counters before continuing.

WATCH FOR:
- 2 fingers up across the room - secure, move on.
- 1 finger up - small group reteach during You Do.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning that multiplication and division come before addition and subtraction.
- Read each I can statement with me.

DO:
- Choral read.
- Point to the relevant letters on the BODMAS card on each desk.

TEACHER NOTES:
SC1 is recall - every student can name M and D as coming before A and S. SC2 is core - apply the rule with both x and ÷. SC3 stretches to subtraction in the same expression, which the cohort is strong on.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Two key words for today.
- An ARRAY is a tidy grid of rows and columns. We use arrays to show multiplication.
- LEFT TO RIGHT means we read from this side to this side. We will use it when two operations are at the same level.

DO:
- Hold up an array of counters at the front (3 rows of 4).
- Point left to right as you say "left to right".

TEACHER NOTES:
Arrays unlock the visual scaffold. Left-to-right will land properly in Lesson 4 - just plant it today.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_IDO1 = `SAY:
- Watch me. 5 + 3 x 4.
- BODMAS says I do the times first. So I build 3 groups of 4 in an array.
- Count with me: 4, 8, 12.
- 3 x 4 is 12.
- Now I add the 5. 12 + 5 is 17.
- Done. The answer is 17.

DO:
- Build the array on the board: 3 rows of 4 counters.
- Tap each row as you count.
- Add 5 counters off to the side.
- Write 17 in green.

TEACHER NOTES:
This is the I Do for x first then +. Slow down on the array - the counters carry the meaning of 3 x 4 = 12. After the array, the addition is friendly.

WATCH FOR:
- Students who track each row with you - secure.
- Students still adding first - flag for small group support.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Same idea works with divide.
- 20 ÷ 4 + 3.
- BODMAS says do the ÷ first. 20 shared between 4 groups gives 5 in each group.
- Now I add the 3. 5 + 3 is 8.
- Done. The answer is 8.

DO:
- Share 20 counters into 4 equal groups.
- Hold up one group: "5 in each".
- Write the answer.

TEACHER NOTES:
Use sharing language for division. Many Year 5/6 students still rely on this concrete meaning. After the ÷ step we have 5 + 3 - a friendly addition.

MISCONCEPTIONS:
- Misconception: Students do the + first because it comes first when you read left to right.
  Why: Reading habit.
  Impact: Wrong answer of 20 ÷ 7 which is messy.
  Quick correction: "BODMAS is the order, not reading order. D comes before A."

WATCH FOR:
- Students who explain "÷ first because D is before A" - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. 12 ÷ 3 + 2.
- On your whiteboard:
- 1. Write which operation comes first - divide or plus.
- 2. Write the answer.

DO:
- Display the prompt.
- 45 seconds.
- Walk and scan boards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your boards on three, two, one, show."
- Scan for: "divide" and answer 6.
PROCEED: If 80% write divide first and 6, click to reveal and move to We Do.
PIVOT: Most likely misconception - students add first.
- Reteach: Share 12 counters into 3 equal groups at the front. "How many in each group? 4. Now add 2."
- Re-check: "Try 15 ÷ 5 + 4 on your board."

TEACHER NOTES:
The single CFU decides movement. Share counters live if you need to pivot.

WATCH FOR:
- Boards showing 6 with "divide first" - secure.
- Boards showing 4.2 or similar messy numbers - they added first.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- Check yours.
- Divide first. 12 ÷ 3 = 4.
- Then 4 + 2 = 6.
- Tick and fix.

DO:
- Click to reveal.

TEACHER NOTES:
Look for students who self-correct from messy answer to 6.

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. 10 - 2 x 3.
- This one has minus. Be careful.
- Step 1 - which one first?
- Step 2 - do the times.
- Step 3 - do the minus.

DO:
- Display the expression.
- 30 seconds partner talk.
- Build the array at the front: 2 rows of 3 counters.
- Cold call a pair.

TEACHER NOTES:
The cohort is strong on subtraction, so this We Do plays to that strength. The harder bit is resisting the urge to subtract first. Watch for "10 - 2 = 8, then 8 x 3 = 24" - the classic BODMAS slip.

WATCH FOR:
- Pairs who say times first - secure.
- Pairs who say minus first - prompt back to the BODMAS card.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 2 x 3 = 6.
- 10 - 6 = 4.
- Done.
- Tick and fix.

DO:
- Click to reveal.
- Build the array, then take 6 away from 10.

TEACHER NOTES:
The subtraction step is friendly for this cohort. The BODMAS step is what we are practising.

[Stage 3: We Do Answer | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the practice sheet.
- For each expression, first UNDERLINE the operation you do first, then work it out step by step.
- Use counters or your whiteboard. Show your working.

DO:
- Distribute the practice sheet.
- Circulate. Listen for "M or D first" reasoning.
- Sit with the small group who pivoted earlier.

TEACHER NOTES:
Six expressions, all with either x or ÷ followed by + or -. The additive piece is kept friendly so BODMAS is the load.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use counters for each expression. Build the array first, then add or subtract.
- Extra Notes: Sit with these students. Do the first two together.
EXTENDING PROMPT:
- Task: For the bonus task at the bottom of the sheet, write a number sentence with BOTH x and ÷ in it. Work out the answer using BODMAS.
- Extra Notes: Encourage left-to-right between x and ÷ since they sit at the same level - we will teach this formally in Lesson 4.

WATCH FOR:
- Students who underline x or ÷ first - secure.
- Students who underline + or - first - bring back to the BODMAS card.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task on your whiteboard.
- Work out 16 ÷ 4 + 5. Underline the step you did first.

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards or photograph.

TEACHER NOTES:
Exit ticket assesses SC2 - applying the rule with division. Expected answer: 9 (since 16 ÷ 4 = 4, then 4 + 5 = 9).

WATCH FOR:
- Students who write 9 with the divide underlined - secure.
- Students who write 16 ÷ 9 = some messy number - bring back to the rule tomorrow.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Success criteria check.
- Thumbs up, sideways, or down for each one.
- Turn and tell your partner: which letters always come before A and S in BODMAS?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call 1-2 students.

TEACHER NOTES:
Threshold idea is "M and D before A and S". Students who can say this are ready for Lesson 3 (brackets).

[General: Closing | VTLM 2.0: Reflection]`;

// --- Helpers ---------------------------------------------------------------

// Draw a pentagon
function drawPentagon(slide, cx, cy, size) {
  const points = [];
  for (let i = 0; i < 5; i += 1) {
    const ang = -Math.PI / 2 + i * (2 * Math.PI / 5);
    const px = cx + Math.cos(ang) * size;
    const py = cy + Math.sin(ang) * size;
    points.push({ x: px, y: py });
  }
  // Draw as 5 lines
  for (let i = 0; i < 5; i += 1) {
    const p1 = points[i];
    const p2 = points[(i + 1) % 5];
    slide.addShape("line", {
      x: p1.x, y: p1.y, w: p2.x - p1.x, h: p2.y - p1.y,
      line: { color: C.PRIMARY, width: 3 },
    });
  }
}

// Draw an expression in a large styled card.
function drawExpression(slide, x, y, w, h, exprText) {
  addTextOnShape(slide, exprText, {
    x, y, w, h, rectRadius: 0.12,
    fill: { color: C.BG_LIGHT || "F6F6F6" },
    line: { color: C.CHARCOAL, width: 1.4 },
  }, {
    fontSize: 50, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
  });
}

// Draw a row of counters for additive scaffolding.
function drawCounters(slide, x, y, count, opts) {
  const o = opts || {};
  const dotSize = o.dotSize || 0.24;
  const gap = o.gap || 0.06;
  const color = o.color || C.PRIMARY;
  for (let i = 0; i < count; i += 1) {
    const cx = x + i * (dotSize + gap);
    slide.addShape("roundRect", {
      x: cx, y, w: dotSize, h: dotSize, rectRadius: dotSize / 2,
      fill: { color }, line: { color, width: 0.2 },
    });
  }
}

// --- Build ------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Multiplication and division come first",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - 2D shape (pentagon)
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Name this 2D shape", [
      "Name this shape.",
      "How many sides does it have?",
    ], NOTES_DR_Q, FOOTER, (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.ACCENT });
      slide.addText("This shape", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.34,
        fontSize: 16, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      drawPentagon(slide, lg.rightX + lg.rightW / 2, lg.panelTopPadded + 1.85, 0.95);
    }),
    (slide) => {
      addTextOnShape(slide, "Pentagon  |  5 sides, 5 angles", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    },
  );

  // Slides 5-6: Fluency + reveal - 2-digit addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Adding 2-digit numbers",
      ["27 + 31", "44 + 25", "38 + 17"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "58      69      55", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    },
  );

  // Slide 7: Launch - retrieval from yesterday
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Launch", { color: C.PRIMARY });
    addTitle(s, "Remember from yesterday?");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    drawExpression(s, 2.0, CONTENT_TOP + 0.35, 6.0, 0.90, "4  +  3  x  2");

    s.addText([
      { text: "Show me on your fingers.", options: { fontSize: 18, color: C.CHARCOAL, bold: true } },
    ], {
      x: 0.85, y: CONTENT_TOP + 1.50, w: 8.30, h: 0.35,
      fontFace: FONT_B, align: "center", margin: 0,
    });

    // Two side-by-side options
    const optY = CONTENT_TOP + 1.95;
    const colW = 4.10;
    addTextOnShape(s, "1 finger\n+ first", {
      x: 0.85, y: optY, w: colW, h: 0.85, rectRadius: 0.10,
      fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.4 },
    }, { fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, bold: true });
    addTextOnShape(s, "2 fingers\nx first", {
      x: 5.05, y: optY, w: colW, h: 0.85, rectRadius: 0.10,
      fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.4 },
    }, { fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, bold: true });

    s.addText("Then write the answer on your board.", {
      x: 0.7, y: SAFE_BOTTOM - 0.45, w: 8.6, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning that multiplication and division come before addition and subtraction.",
    [
      "I can name M and D as the operations that come before A and S in BODMAS.",
      "I can work out an expression with x or ÷ followed by + or -.",
      "I can explain WHY we do the M or D step first using an array or sharing.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key vocabulary
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Two words for today");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const wordY = CONTENT_TOP + 0.30;
    const cellW = 4.20;
    const gap = 0.30;
    const startX = (10 - (cellW * 2 + gap)) / 2;

    addTextOnShape(s, "array", {
      x: startX, y: wordY, w: cellW, h: 0.55, rectRadius: 0.10,
      fill: { color: C.PRIMARY },
    }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Mini array for "array"
    addArray(s, startX + 1.00, wordY + 0.85, 3, 4, { cellSize: 0.32, color: C.PRIMARY });
    s.addText("rows and columns of dots\nused for multiplication", {
      x: startX, y: wordY + 2.20, w: cellW, h: 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "top", margin: 0,
    });

    // Left-to-right card
    const lrX = startX + cellW + gap;
    addTextOnShape(s, "left to right", {
      x: lrX, y: wordY, w: cellW, h: 0.55, rectRadius: 0.10,
      fill: { color: C.PRIMARY },
    }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Arrow visual
    slide_arrow(s, lrX + 0.50, wordY + 1.30, cellW - 1.00, 0.50);

    s.addText("read this way when two operations\nare at the same level (eg. x and ÷)", {
      x: lrX, y: wordY + 2.20, w: cellW, h: 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 10: I Do (1) - 5 + 3 x 4 with array
  workedExSlide(pres, 2, "I Do", "5 + 3 x 4  - times first",
    [
      "Step 1.  Which one first?",
      "x is M.  M comes before A.",
      "So do the times first.",
      "",
      "Step 2.  3 x 4 = 12.",
      "(Look at the array.)",
      "",
      "Step 3.  Add the 5.",
      "12 + 5 = 17.",
      "",
      "Answer: 17.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("3 x 4 array", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.32,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Array
      addArray(slide, lg.rightX + 0.85, lg.panelTopPadded + 0.55, 3, 4,
        { cellSize: 0.48, color: C.ACCENT });

      addTextOnShape(slide, "= 12", {
        x: lg.rightX + 0.30, y: lg.panelTopPadded + 2.45, w: lg.rightW - 0.60, h: 0.50, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    });

  // Slide 11: I Do (2) - 20 ÷ 4 + 3
  workedExSlide(pres, 2, "I Do", "20 ÷ 4 + 3  - divide first",
    [
      "Step 1.  Which one first?",
      "÷ is D.  D comes before A.",
      "So do the divide first.",
      "",
      "Step 2.  20 ÷ 4 = 5.",
      "(20 shared into 4 equal groups.)",
      "",
      "Step 3.  Add the 3.",
      "5 + 3 = 8.",
      "",
      "Answer: 8.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("Share 20 into 4 groups", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.2, h: 0.32,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Four groups of 5 counters
      const grpY = lg.panelTopPadded + 0.55;
      for (let g = 0; g < 4; g += 1) {
        const gy = grpY + g * 0.46;
        slide.addShape("roundRect", {
          x: lg.rightX + 0.20, y: gy, w: lg.rightW - 0.40, h: 0.38, rectRadius: 0.08,
          fill: { color: C.BG_LIGHT || "F6F6F6" },
          line: { color: C.ACCENT, width: 1 },
        });
        drawCounters(slide, lg.rightX + 0.35, gy + 0.08, 5,
          { dotSize: 0.20, gap: 0.10, color: C.ACCENT });
      }

      addTextOnShape(slide, "5 in each", {
        x: lg.rightX + 0.30, y: lg.panelTopPadded + 2.50, w: lg.rightW - 0.60, h: 0.42, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    });

  // Slides 12-13: CFU + reveal - 12 ÷ 3 + 2
  withReveal(
    () => cfuSlide(pres, "CFU", "Which operation first?", "Show Me Boards",
      "12 ÷ 3 + 2\n\nOn your whiteboard:\n\n1.  Write which operation comes FIRST.\n2.  Write the answer.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "÷ first.  12 ÷ 3 = 4.  Then 4 + 2 = 6.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    },
  );

  // Slides 14-15: We Do + reveal - 10 - 2 x 3 (subtraction strength)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Together: 10 - 2 x 3", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      drawExpression(s, 2.0, CONTENT_TOP + 0.30, 6.0, 0.85, "10  -  2  x  3");

      // Two-column step layout
      const stepY = CONTENT_TOP + 1.40;
      const stepH = 1.65;
      const colW = 4.10;

      // Left column: Step 1 + Step 2 with array
      addTextOnShape(s, "Step 1.  x first.", {
        x: 0.85, y: stepY, w: colW, h: 0.45, rectRadius: 0.08,
        fill: { color: C.ACCENT },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

      s.addText("Step 2.  2 x 3 = 6.", {
        x: 0.85, y: stepY + 0.55, w: colW, h: 0.34,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
        align: "left", margin: 0,
      });
      addArray(s, 1.85, stepY + 0.95, 2, 3, { cellSize: 0.32, color: C.ACCENT });

      // Right column: Step 3
      addTextOnShape(s, "Step 3.  Now minus.", {
        x: 5.05, y: stepY, w: colW, h: 0.45, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

      s.addText("10 - 6 = ?", {
        x: 5.05, y: stepY + 0.65, w: colW, h: 0.55,
        fontSize: 24, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", margin: 0,
      });

      s.addText("Take 6 away from 10.", {
        x: 5.05, y: stepY + 1.25, w: colW, h: 0.32,
        fontSize: 13, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "2 x 3 = 6.  10 - 6 = 4.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    },
  );

  // Slide 16: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.20, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "underline the operation you do first.   ", options: { fontSize: 16, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "use counters or your board to do that step.   ", options: { fontSize: 16, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "finish the sentence.", options: { fontSize: 16, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.15, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.40;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Remember", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "M and D before A and S.", {
      x: 1.5, y: panelY + 0.55, w: 7.0, h: 0.55, rectRadius: 0.10,
      fill: { color: C.PRIMARY },
    }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

    s.addText("Counters and arrays are on your desk - use them.", {
      x: 0.85, y: panelY + 1.30, w: 8.30, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Work out  16 ÷ 4 + 5.  Underline the step you did first.",
      "Tell your partner why M and D come before A and S.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: which letters always come BEFORE A and S in BODMAS?",
      scItems: [
        "I can name M and D as the operations that come before A and S in BODMAS.",
        "I can work out an expression with x or ÷ followed by + or -.",
        "I can explain WHY we do the M or D step first using an array or sharing.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BODMAS_Lesson2_Multiply_And_Divide_First.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDF: Practice sheet --------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: PRACTICE_RES.name });
    let y = addPdfHeader(doc, PRACTICE_RES.name, {
      subtitle: "Underline the step you do FIRST. Then work it out.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addTipBox(doc,
      "BODMAS reminder: M and D before A and S. Counters and arrays will help you see the multiply or divide step.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Part A — Underline and solve", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  6 + 4 x 3   =   _____", y);
    y = addWriteLine(doc, "b)  8 + 2 x 5   =   _____", y);
    y = addWriteLine(doc, "c)  15 ÷ 3 + 4  =   _____", y);
    y = addWriteLine(doc, "d)  18 ÷ 2 + 6  =   _____", y);
    y = addWriteLine(doc, "e)  10 - 3 x 2  =   _____", y);
    y = addWriteLine(doc, "f)  20 - 2 x 5  =   _____", y);

    y = addSectionHeading(doc, "Part B — Write and explain", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "Look at:  9 + 6 ÷ 3.  A student wrote 9 + 6 = 15, then 15 ÷ 3 = 5.",
      y);
    y = addWriteLine(doc, "What did they do wrong? __________________________________________", y);
    y = addWriteLine(doc, "What is the correct answer? ______________________________________", y);

    y = addSectionHeading(doc, "Part C — Bonus (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc,
      "Write a number sentence with BOTH x and ÷ in it. Work out the answer using BODMAS.",
      y);
    y = addWriteLine(doc, "Sentence: _____________________________________________________", y);
    y = addWriteLine(doc, "Answer: _______________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Multiply and Divide First | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, PRACTICE_RES.fileName));
    console.log("PDF written: " + PRACTICE_RES.fileName);
  })();

  // --- PDF: Answer Key ------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 2 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Part A answers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  6 + 4 x 3   =   6 + 12  =  18.", y);
    y = addBodyText(doc, "b)  8 + 2 x 5   =   8 + 10  =  18.", y);
    y = addBodyText(doc, "c)  15 ÷ 3 + 4  =   5 + 4   =  9.", y);
    y = addBodyText(doc, "d)  18 ÷ 2 + 6  =   9 + 6   =  15.", y);
    y = addBodyText(doc, "e)  10 - 3 x 2  =   10 - 6  =  4.", y);
    y = addBodyText(doc, "f)  20 - 2 x 5  =   20 - 10 =  10.", y);

    y = addSectionHeading(doc, "Part B answer", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "The student added before dividing. BODMAS says D before A. The correct working is:",
      y);
    y = addBodyText(doc, "9 + 6 ÷ 3  =  9 + 2  =  11.", y);

    y = addSectionHeading(doc, "Part C", y, { color: C.ACCENT });
    y = addBodyText(doc,
      "Look for sentences with both x and ÷ where the student does them left to right at the same level.",
      y);

    y = addTipBox(doc,
      "Watch for: students who add or subtract first - reteach with arrays and the BODMAS card. Students who get part B but with a wrong final number - check the multiplication or division step itself.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

// Helper: arrow shape (rendered as a line plus a triangle head)
function slide_arrow(slide, x, y, w, h) {
  slide.addShape("line", {
    x, y: y + h / 2, w: w - 0.25, h: 0,
    line: { color: C.PRIMARY, width: 4 },
  });
  slide.addShape("triangle", {
    x: x + w - 0.35, y: y + h / 2 - 0.18, w: 0.35, h: 0.36,
    rotate: 90,
    fill: { color: C.PRIMARY },
  });
}

build().catch((err) => { console.error(err); process.exit(1); });
