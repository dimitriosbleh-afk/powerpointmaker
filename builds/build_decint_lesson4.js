"use strict";

// Decimals & Integers Unit (Year 6 Numeracy) - Lesson 4: Integers in context
// VC2M6N01. Integers in real contexts: temperature, sea level/elevation, profit and loss.
// Daily Review: Shapes + Data. Fluency: vertical addition algorithm with money.
// Carries the unit's Year 8 Extension Challenge resource.
// Unit variant fixed across all 4 lessons for cohesion.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide,
  addNumberLine,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const TOTAL = 4;
const UNIT_TITLE = "Decimals and Integers";
const FOOTER = `Decimals & Integers | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/DecInt_Lesson4_Integers_In_Context";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Integers in Context",
  "Use integers for sea level, temperature and profit/loss, then order results around zero.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the integers in context practice sheet.");
const Y8_RES = makeSessionResource(SESSION,
  "Year 8 Extension Challenge",
  "Stretch tasks: decimals to thousandths, integer operations, profit running totals, and the coordinate plane.");
const Y8_KEY_RES = makeSessionResource(SESSION,
  "Year 8 Extension Answer Key",
  "Worked answers for the Year 8 extension challenge.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, Y8_RES, Y8_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 4, the last in our decimals and integers unit.
- Today we put integers to work in real life: sea level, temperature, and money.
- We will use positive and negative numbers to describe above and below, and profit and loss.

DO:
- Have whiteboards, markers and a printed number line ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 4 of 4. This applies the integer number line from Lesson 3 to real contexts. A student who only missed Lesson 3 can still join - the launch and I Do rebuild the number line in context.

WATCH FOR:
- Students who connect negatives to a freezer, a basement or a bank balance - draw that out.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- There is the usual practice sheet, and a separate Year 8 challenge for anyone ready to go further.

DO:
- Print one practice sheet and answer key per student.
- Print a few copies of the Year 8 extension challenge for early finishers.
- Have whiteboards, markers and a printed number line ready.

TEACHER NOTES:
Two student resources today: the lesson practice sheet (with built-in enabling and extension) and a separate Year 8 Extension Challenge for students working well beyond level. Both have answer keys.

CATCH-UP NOTE:
This lesson applies Lesson 3. A returning student can use the printed number line and the sea level diagram to access the contexts. The enabling section keeps the numbers small.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review, two quick ones.
- Shapes: how many vertices does the shape on the left have?
- Data: on the column graph, which bar is the tallest?
- Write both answers on your whiteboard.

DO:
- Display the shape and the graph.
- 60 seconds.
- Walk and check both answers.

TEACHER NOTES:
Prior geometry and data, not today's integers. A vertex is a corner where sides meet. Reading a column graph means comparing bar heights.

WATCH FOR:
- Students who count vertices correctly - secure.
- Students who read the graph by height, not by guess - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- The hexagon has 6 vertices - 6 corners.
- On the graph, bar C is the tallest, with 6 goals.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick-and-fix.

TEACHER NOTES:
Vertices equals corners equals number of sides for a polygon. For the graph, the tallest bar shows the most. Note any student who reads the graph wrong for small group data work.

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Vertical addition with money.
- Line up the decimal points so dollars sit under dollars and cents under cents.
- Add and write the total.

DO:
- Display the three prompts.
- 2 minutes.
- Scan for lined-up decimal points.

TEACHER NOTES:
Same vertical algorithm, now with money to two decimal places. Lining up the decimal point is the key habit. This ties our decimal work to a real context.

WATCH FOR:
- Students who line up the points - secure.
- Students who forget the cents column makes a dollar when it reaches 100 - watch the carry.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours, in dollars.
- 12.50 plus 8.75 is 21.25.
- 45.60 plus 9.40 is 55.00.
- 7.05 plus 23.95 is 31.00.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The second and third make exact dollar amounts - the cents carry into a new dollar. That is the point to coach.

WATCH FOR:
- Students who self-correct - secure.
- Students who drop the carry into the dollars - reteach with the placeholder zero in cents.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember our number line going below zero.
- Where do we see below zero in real life?
- A lift goes up to level 3 and down to basement 2. Sea level is zero - a cliff is above, a diver is below.
- A shop can make a profit or a loss. Today we use integers for all of these.

DO:
- Point up and down as you give each example.
- Take two or three real examples from students.
- Land on the idea: positive is one direction, negative is the other.

TEACHER NOTES:
This launch connects the abstract number line from Lesson 3 to concrete contexts before the learning intention. Above and below, profit and loss, are the everyday meanings of positive and negative.

WATCH FOR:
- Students who offer a real below-zero example - strong transfer.
- Students unsure - the sea level diagram in the I Do makes it concrete.

[Stage: Launch | VTLM 2.0: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to use integers to describe and compare real situations above and below zero.
- Now the three I can statements. Read them with me.

DO:
- Choral read the LI and SC.
- Hold up the printed number line.

TEACHER NOTES:
SC1 is achievable for everyone - choosing a positive or negative sign for an above or below situation. SC2 is the core target the exit ticket checks. SC3 stretches to ordering several real results around zero.

WATCH FOR:
- Students who can match a sign to a situation - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Words for today.
- Above zero is positive: a cliff is +3 metres, a profit is positive.
- Below zero is negative: a diver is -2 metres, a loss is negative.
- Sea level is zero - the boundary between above and below.
- The sign shows direction: which side of zero we are on.

DO:
- Point to above and below on the diagram.
- Have students show thumbs up for positive, thumbs down for negative.

TEACHER NOTES:
The key idea is that the sign carries meaning - direction from a zero that we choose (sea level, ground floor, break even). Keep it to these words.

WATCH FOR:
- Students who match sign to situation - secure.
- Students who say a loss is positive because the number looks normal - stress loss is below zero.

[General: Key Vocabulary | VTLM 2.0: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Watch me use sea level as our zero.
- Above sea level is positive. The gull is at +3 metres.
- Below sea level is negative. The diver is at -2 metres. The wreck is at -5 metres.
- The diver is below sea level, so its height is negative.
- The wreck at -5 is lower than the diver at -2, because -5 is further down.

DO:
- Point to sea level zero first.
- Move up to +3, then down to -2 and -5.
- Compare -5 and -2 by how far below zero they are.

TEACHER NOTES:
Sea level is the chosen zero. This makes negative numbers concrete - below the line. Lower equals more negative equals smaller, exactly as on the vertical number line.

MISCONCEPTIONS:
- Misconception: students think -5 metres is higher than -2 metres because 5 is bigger.
  Why: they read the digit, not the direction.
  Impact: they order depths backwards.
  Quick correction: -5 is further BELOW sea level, so it is lower and smaller.

WATCH FOR:
- Students who say the wreck is lowest - secure.
- Students who say -5 is highest - re-point down the scale.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Watch me use money. A market stall.
- On a good day it makes a profit. We write profit as a positive: +5 dollars.
- On a bad day it makes a loss. We write loss as a negative: -3 dollars.
- Break even, no profit and no loss, is zero.
- A bigger loss is a smaller number: -8 is less than -3.

DO:
- Place +5 to the right of zero and -3 to the left on the line.
- Stress profit is positive, loss is negative, break even is zero.
- Compare -8 and -3 as sizes of loss.

TEACHER NOTES:
Money makes the sign meaningful. Profit is positive, loss is negative, break even is zero. A larger loss sits further left, so it is the smaller number.

MISCONCEPTIONS:
- Misconception: students think a 8 dollar loss is better than a 3 dollar loss because 8 is bigger.
  Why: they ignore that the number is negative.
  Impact: they rank losses backwards.
  Quick correction: -8 is further left than -3, so -8 is the smaller number - a worse result.

WATCH FOR:
- Students who say -8 is the worse result - secure.
- Students who say -8 is better - re-point left of zero.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- A diver is at -8 metres. A fish is at -3 metres.
- Which one is deeper, that is, lower in the water?
- Write the depth and a reason.

DO:
- Display the prompt and the scale.
- 45 seconds.
- Walk and scan for lower-is-more-negative reasoning.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: write who is deeper on three, two, one, show.
- Scan for: the diver at -8, because -8 is further below zero.
PROCEED: If about 80 percent write the diver, click to reveal and move to We Do.
PIVOT: Most likely misconception - students pick the fish because -3 looks closer to the surface.
- Reteach: on the vertical scale, -8 is further DOWN than -3, so the diver is deeper.
- Re-check: which is further below zero, and which is deeper?

TEACHER NOTES:
Depth is the context for ordering negatives. Deeper equals lower equals more negative equals smaller.

WATCH FOR:
- Students who pick the diver with reasoning - secure.
- Students who pick the fish - depth misconception, reteach on the scale.

[Stage 2: CFU | VTLM 2.0: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Sea level is zero.
- Order these heights from LOWEST to HIGHEST: +3, -5, 0, -2 metres.
- Lowest is furthest below sea level. Highest is furthest above.

DO:
- Display the four heights and the line.
- 90 seconds.
- Listen for furthest-below reasoning.

TEACHER NOTES:
Ordering elevations around sea level. -5 is lowest (deepest), +3 is highest. Reading the line left to right gives the order.

WATCH FOR:
- Pairs who put -5 lowest - secure.
- Pairs who put -2 below -5 - reteach: -5 is further below sea level.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Lowest to highest: -5, then -2, then 0, then +3.
- -5 metres is deepest, +3 metres is highest above sea level.

DO:
- Click to reveal.
- Trace left to right along the line.

TEACHER NOTES:
The reveal shows ordering elevations is reading the number line left to right. Lowest sits furthest left (furthest below zero).

WATCH FOR:
- Students who self-correct - secure.
- Students who reversed the negatives - small group before You Do.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, with money.
- A stall's results over four days: +20, -15, 0, -5 dollars.
- Order them from WORST to BEST.
- Worst is the biggest loss. Best is the biggest profit.

DO:
- Display the four results and the money line (counting in fives).
- 75 seconds.
- Listen for biggest-loss-is-smallest reasoning.

TEACHER NOTES:
Profit and loss ordering. Worst equals biggest loss equals most negative. -15 is worst, +20 is best.

WATCH FOR:
- Pairs who put -15 worst - secure.
- Pairs who put -5 below -15 - reteach: -15 is further left, a bigger loss.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Worst to best: -15, then -5, then 0, then +20.
- -15 dollars is the biggest loss, +20 dollars is the biggest profit.

DO:
- Click to reveal the order on the line.
- Point to how far each sits from zero.

TEACHER NOTES:
If pairs struggled, walk the money line from the left before You Do.

WATCH FOR:
- Students who justify worst-is-most-negative - ready for independent work.
- Students who reversed it - enabling group for You Do.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1: write the integer for each real situation (above or below zero).
- Section 2: compare two situations with greater than or less than.
- Section 3: order each set, from lowest or worst to highest or best.
- If you finish, try the extension, or ask for the Year 8 challenge.

DO:
- Distribute the practice sheet.
- Circulate and listen for sign-and-direction reasoning.
- Offer the Year 8 challenge to students who finish early.

TEACHER NOTES:
Different contexts from the We Do, same idea: choose the sign for direction, then order using the number line.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: The number line is drawn for you. Do Sections 1 and 2 only, pointing to above or below zero each time.
- Extra Notes: Sit with these students and write the first integer together.
EXTENDING PROMPT:
- Task: Extension - order a mixed set of temperatures, depths and money results, then answer a short profit-running-total question.
- Extra Notes: Students who finish the extension take the separate Year 8 Extension Challenge.

WATCH FOR:
- Students who match the sign to the situation - secure.
- Students who rank losses or depths backwards - prompt them to use the line.

[Stage 4: You Do | VTLM 2.0: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task of the unit, on your whiteboard.
- A morning temperature is -4 degrees. The afternoon is 6 degrees. Which is colder, and how do you know?
- Then order these results from worst to best: -10, 5, -2, 0 dollars.

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses SC2 - comparing and ordering integers in context. Look for -4 named as colder (further below zero) and -10 as the worst result.

WATCH FOR:
- Students who order -10, -2, 0, 5 - secure.
- Students who put -2 worse than -10 - direction misconception, note for next unit.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria one last time for this unit.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: give one real example of a negative number and say what it means.

DO:
- Read each I can statement.
- Use thumbs.
- Cold call two or three students for the partner share.

TEACHER NOTES:
The threshold idea is that the sign shows direction from a chosen zero - above or below, profit or loss - and ordering still reads along the number line. This closes the unit; note any students for spaced review next unit.

WATCH FOR:
- Strong thumbs up across all three - the unit has landed.
- Sideways or down on SC2 - flag for spaced review in the next unit.

[General: Closing | VTLM 2.0: Reflection]`;

// --- Helpers -----------------------------------------------------------------

// Vertical number line (sea level / thermometer). labels[0] is the TOP (largest).
function addVerticalNumberLine(slide, cx, yTop, h, labels, markedIdx, opts) {
  const o = opts || {};
  const n = labels.length - 1;
  if (n <= 0) return;
  const intervalH = h / n;
  const lineColor = o.lineColor || C.CHARCOAL;

  slide.addShape("line", { x: cx, y: yTop, w: 0, h: h, line: { color: lineColor, width: 2.5 } });

  labels.forEach((lbl, i) => {
    const ty = yTop + i * intervalH;
    slide.addShape("line", { x: cx - 0.09, y: ty, w: 0.18, h: 0, line: { color: lineColor, width: 2 } });
    if (lbl !== "") {
      const isZero = o.zeroIdx === i;
      slide.addText(String(lbl), {
        x: cx - 1.0, y: ty - 0.12, w: 0.82, h: 0.24,
        fontSize: o.fontSize || 12, fontFace: FONT_B,
        color: isZero ? C.PRIMARY : C.CHARCOAL, bold: isZero,
        align: "right", valign: "middle", margin: 0,
      });
    }
  });

  (markedIdx || []).forEach((m) => {
    const idx = typeof m === "object" ? m.idx : m;
    const my = yTop + idx * intervalH;
    slide.addShape("roundRect", {
      x: cx - 0.09, y: my - 0.09, w: 0.18, h: 0.18, rectRadius: 0.09,
      fill: { color: (typeof m === "object" && m.color) || o.markColor || C.ALERT },
    });
    if (typeof m === "object" && m.label) {
      slide.addText(String(m.label), {
        x: cx + 0.18, y: my - 0.13, w: 1.5, h: 0.26,
        fontSize: 12, fontFace: FONT_B, color: m.color || C.ALERT, bold: true,
        align: "left", valign: "middle", margin: 0,
      });
    }
  });
}

// Small column graph for the Daily Review data prompt.
function addMiniColumnGraph(slide, x0, baselineY, values, labels, opts) {
  const o = opts || {};
  const barW = o.barW || 0.55;
  const gap = o.gap || 0.30;
  const unit = o.unit || 0.22; // inches per value unit
  const fill = o.fill || C.PRIMARY;
  // baseline
  slide.addShape("line", {
    x: x0 - 0.1, y: baselineY, w: values.length * (barW + gap) + 0.1, h: 0,
    line: { color: C.CHARCOAL, width: 1.5 },
  });
  values.forEach((v, i) => {
    const bx = x0 + i * (barW + gap);
    const bh = v * unit;
    slide.addShape("rect", {
      x: bx, y: baselineY - bh, w: barW, h: bh,
      fill: { color: fill }, line: { color: C.CHARCOAL, width: 0.75 },
    });
    slide.addText(String(labels[i] || ""), {
      x: bx - 0.1, y: baselineY + 0.04, w: barW + 0.2, h: 0.26,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", margin: 0,
    });
  });
}

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 4: Integers in context",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Shapes + Data
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Daily Review: Shapes and data", { color: STAGE_COLORS["1"] });

      // Left card - shape (vertices)
      addCard(s, 0.5, CONTENT_TOP, 4.3, 2.95, { strip: STAGE_COLORS["1"] });
      s.addText("Shapes", {
        x: 0.7, y: CONTENT_TOP + 0.1, w: 3.9, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: STAGE_COLORS["1"], bold: true, margin: 0,
      });
      s.addShape("hexagon", {
        x: 1.7, y: CONTENT_TOP + 0.55, w: 1.9, h: 1.55,
        fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("How many vertices (corners)?", {
        x: 0.7, y: CONTENT_TOP + 2.25, w: 3.9, h: 0.5,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      // Right card - data (column graph)
      addCard(s, 5.0, CONTENT_TOP, 4.5, 2.95, { strip: STAGE_COLORS["1"] });
      s.addText("Data: goals scored", {
        x: 5.2, y: CONTENT_TOP + 0.1, w: 4.1, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: STAGE_COLORS["1"], bold: true, margin: 0,
      });
      addMiniColumnGraph(s, 5.55, CONTENT_TOP + 2.0, [4, 3, 6, 2], ["A", "B", "C", "D"],
        { fill: C.SECONDARY, unit: 0.20 });
      s.addText("Which bar is the tallest?", {
        x: 5.2, y: CONTENT_TOP + 2.4, w: 4.1, h: 0.45,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Hexagon = 6 vertices       Tallest bar = C (6 goals)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - vertical addition with money
  withReveal(
    () => fluencySlide(pres, "Fluency: Vertical addition (money)",
      ["12.50 + 8.75", "45.60 + 9.40", "7.05 + 23.95"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "21.25        55.00        31.00", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - below zero in real life
  contentSlide(pres, "Launch", C.ACCENT, "Where do we see below zero?",
    [
      "A lift: level 3 up, basement 2 down.",
      "Sea level: a cliff above, a diver below.",
      "Money: a profit or a loss.",
      "",
      "Positive one way, negative the other.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.ACCENT });
      slide.addText("Sea level = zero", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.26,
        fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
      });
      addVerticalNumberLine(slide, lg.rightX + 1.7, lg.panelTopPadded + 0.5, 2.2,
        ["3", "2", "1", "0", "-1", "-2", "-3"],
        [{ idx: 0, label: "cliff +3", color: C.SECONDARY }, { idx: 5, label: "diver -2", color: C.ALERT }],
        { zeroIdx: 3, fontSize: 12 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to use integers to describe and compare real situations above and below zero.",
    [
      "I can choose a positive or negative sign for an above or below situation.",
      "I can compare two real situations using integers and a number line.",
      "I can order several results around zero, such as depths or profit and loss.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Above and below zero",
    [
      "Above zero = positive: a cliff +3 m, a profit.",
      "Below zero = negative: a diver -2 m, a loss.",
      "Sea level (or break even) = zero, the boundary.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
      slide.addText("The sign shows direction", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
      });
      const ry0 = lg.panelTopPadded + 0.6;
      addTextOnShape(slide, "+  above / profit / warmer", {
        x: lg.rightX + 0.3, y: ry0, w: lg.rightW - 0.6, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "0  sea level / break even", {
        x: lg.rightX + 0.3, y: ry0 + 0.66, w: lg.rightW - 0.6, h: 0.45, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "-  below / loss / colder", {
        x: lg.rightX + 0.3, y: ry0 + 1.22, w: lg.rightW - 0.6, h: 0.55, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10: I Do #1 - sea level / elevation (vertical scale)
  workedExSlide(pres, 2, "I Do", "Sea level is our zero",
    [
      "Above sea level: positive.",
      "Below sea level: negative.",
      "",
      "Gull at +3 m (above).",
      "Diver at -2 m (below).",
      "Wreck at -5 m (further below).",
      "",
      "-5 is lower than -2.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.PRIMARY });
      slide.addText("Above + / Below -", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.26,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addVerticalNumberLine(slide, lg.rightX + 1.55, lg.panelTopPadded + 0.5, 2.3,
        ["3", "2", "1", "0", "-1", "-2", "-3", "-4", "-5"],
        [
          { idx: 0, label: "gull +3", color: C.SECONDARY },
          { idx: 5, label: "diver -2", color: C.ALERT },
          { idx: 8, label: "wreck -5", color: C.ALERT },
        ],
        { zeroIdx: 3, fontSize: 12 });
    }
  );

  // Slide 11: I Do #2 - money / profit and loss (wide number line)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Profit and loss", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.9, { strip: C.PRIMARY });
    s.addText([
      { text: "Profit is positive (+).   Loss is negative (-).   Break even is 0.", options: { fontSize: 18, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "A good day: +5 dollars profit.   A bad day: -3 dollars loss.", options: { fontSize: 17, color: C.CHARCOAL, breakLine: true } },
      { text: "A bigger loss is a smaller number:  -8 < -3.", options: { fontSize: 17, color: C.ALERT, bold: true } },
    ], {
      x: 0.8, y: CONTENT_TOP + 0.22, w: 8.4, h: 1.5, fontFace: FONT_B, valign: "middle", margin: 0,
    });

    const nlY = 3.5;
    addNumberLine(s, 0.9, nlY, 8.4,
      ["-5", "-4", "-3", "-2", "-1", "0", "1", "2", "3", "4", "5"], [2, 10]);
    s.addText("The orange dots show a -3 dollar loss and a +5 dollar profit.", {
      x: 0.8, y: nlY + 0.5, w: 8.4, h: 0.3,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO2);
  })();

  // Slides 12-13: CFU + reveal - which diver is deeper (vertical scale)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Who is deeper?", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 8.0, y: 0.20, w: 1.4, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 19, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 9, breakLine: true } },
        { text: "Diver at -8 m.", options: { fontSize: 20, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "Fish at -3 m.", options: { fontSize: 20, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 9, breakLine: true } },
        { text: "Who is deeper (lower)?", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: true } },
        { text: "Write a reason.", options: { fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 0.72, y: CONTENT_TOP + 0.20, w: 3.85, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.1, CONTENT_TOP, 4.4, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("Depth scale (m)", {
        x: 5.25, y: CONTENT_TOP + 0.12, w: 4.1, h: 0.26,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addVerticalNumberLine(s, 6.6, CONTENT_TOP + 0.55, 2.7,
        ["0", "-1", "-2", "-3", "-4", "-5", "-6", "-7", "-8"],
        [{ idx: 3, label: "fish", color: C.SECONDARY }, { idx: 8, label: "diver", color: C.ALERT }],
        { zeroIdx: 0, fontSize: 11 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "The diver is deeper. -8 is further below zero than -3, so -8 < -3.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - order elevations (wide line)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Order heights: lowest to highest", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 1.9, { strip: C.SECONDARY });
      s.addText([
        { text: "Sea level = 0.  Order these heights:   +3,   -5,   0,   -2 m", options: { fontSize: 18, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Lowest = furthest BELOW sea level.   Highest = furthest ABOVE.", options: { fontSize: 16, color: C.SECONDARY } },
      ], {
        x: 0.8, y: CONTENT_TOP + 0.25, w: 8.4, h: 1.4, fontFace: FONT_B, valign: "middle", margin: 0,
      });

      const nlY = 3.45;
      addNumberLine(s, 0.9, nlY, 8.4,
        ["-5", "-4", "-3", "-2", "-1", "0", "1", "2", "3", "4", "5"], [0, 3, 5, 8]);
      s.addText("Each orange dot is one of the four heights.", {
        x: 0.8, y: nlY + 0.5, w: 8.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Lowest to highest:   -5  <  -2  <  0  <  +3   (metres)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - order profit/loss (money line in fives)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Order results: worst to best", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 1.9, { strip: C.SECONDARY });
      s.addText([
        { text: "A stall's four days (dollars):   +20,   -15,   0,   -5", options: { fontSize: 18, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Worst = biggest loss (most negative).   Best = biggest profit.", options: { fontSize: 16, color: C.SECONDARY } },
      ], {
        x: 0.8, y: CONTENT_TOP + 0.25, w: 8.4, h: 1.4, fontFace: FONT_B, valign: "middle", margin: 0,
      });

      const nlY = 3.45;
      addNumberLine(s, 0.9, nlY, 8.4,
        ["-20", "-15", "-10", "-5", "0", "5", "10", "15", "20"], [1, 3, 4, 8]);
      s.addText("The money line counts in fives. Each orange dot is one day's result.", {
        x: 0.8, y: nlY + 0.5, w: 8.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Worst to best:   -15  <  -5  <  0  <  +20   (dollars)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // Slide 18: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "write the integer for each real situation.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "compare two situations with > or <.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "order each set, lowest/worst to highest/best.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.4;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember", {
      x: 0.7, y: panelY + 0.13, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
    });
    addTextOnShape(s, "The sign shows direction from zero. Greater = further RIGHT / higher / more profit.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

    addNumberLine(s, 0.9, panelY + 1.55, 8.2,
      ["-10", "-8", "-6", "-4", "-2", "0", "2", "4", "6", "8", "10"], []);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Morning is -4 degrees, afternoon is 6 degrees. Which is colder, and how do you know?",
      "Order from worst to best (dollars): -10, 5, -2, 0.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: give one real example of a negative number and say what it means.",
      scItems: [
        "I can choose a positive or negative sign for an above or below situation.",
        "I can compare two real situations using integers and a number line.",
        "I can order several results around zero, such as depths or profit and loss.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecInt_Lesson4_Integers_In_Context.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Use integers for sea level, temperature and profit/loss, then order results.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "The sign shows direction from zero. Above / profit / warmer = positive. Below / loss / colder = negative. A bigger loss or a deeper point is a SMALLER number: -8 < -3.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "A diver 6 m below sea level is at -6. A gull 2 m above is at +2. The diver's number is negative because it is below zero.",
      y);

    y = addSectionHeading(doc, "Section 1 - Write the integer for each situation", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  4 metres above sea level         = ______", y);
    y = addWriteLine(doc, "b)  a 7 dollar loss                  = ______", y);
    y = addWriteLine(doc, "c)  3 degrees below zero             = ______", y);
    y = addWriteLine(doc, "d)  break even (no profit, no loss)  = ______", y);

    y = addSectionHeading(doc, "Section 2 - Compare with > or <", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  a diver at -5 m ____ a diver at -9 m   (which is higher?)", y);
    y = addWriteLine(doc, "b)  -2 degrees ____ 3 degrees", y);
    y = addWriteLine(doc, "c)  a 12 dollar loss ____ a 4 dollar loss   (-12 ____ -4)", y);

    y = addSectionHeading(doc, "Section 3 - Order each set", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  Heights (lowest to highest): 2, -6, 0, -1   ->  ____ , ____ , ____ , ____", y);
    y = addWriteLine(doc, "b)  Results worst to best ($): 8, -3, -10, 0  ->  ____ , ____ , ____ , ____", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order this mixed set, smallest to largest:  -7, 4, 0, -2, 9, -11", y);
    y = addWriteLine(doc, "Order: ____________________________________________________________", y);
    y = addBodyText(doc, "A stall makes a 6 dollar profit, then a 10 dollar loss, then a 7 dollar profit. What is the running total?", y);
    y = addWriteLine(doc, "Answer: ___________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Integers in Context | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the integers in context practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Write the integer", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  +4      b)  -7      c)  -3      d)  0", y);

    y = addSectionHeading(doc, "Section 2 - Compare", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  -5 > -9 (a diver at -5 m is higher)      b)  -2 < 3      c)  -12 < -4", y);

    y = addSectionHeading(doc, "Section 3 - Order", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Lowest to highest: -6, -1, 0, 2", y);
    y = addBodyText(doc, "b)  Worst to best: -10, -3, 0, 8", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order: -11, -7, -2, 0, 4, 9.", y);
    y = addBodyText(doc, "Running total: +6, then 6 - 10 = -4 (a 4 dollar loss), then -4 + 7 = +3. Final: 3 dollar profit.", y);

    y = addTipBox(doc,
      "Watch for: students who rank a 9 dollar loss as better than a 3 dollar loss; students who think -5 m is deeper than -9 m; students who forget the sign when ordering.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension Challenge ------------------------------------------------
  await (async () => {
    const doc = createPdf({ title: Y8_RES.name });
    let y = addPdfHeader(doc, Y8_RES.name, {
      subtitle: "For students ready to go beyond Year 6 - decimals, integer operations, and the coordinate plane.",
      color: C.SECONDARY,
      lessonInfo: `Decimals & Integers Unit | Year 8 stretch`,
    });
    y = addTipBox(doc,
      "These tasks reach into Year 7 and Year 8. Use everything you know about place value, the number line and the sign as direction. Show your reasoning.",
      y, { color: C.SECONDARY });

    y = addSectionHeading(doc, "1 - Decimals to thousandths and between", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  Order smallest to largest: 0.305, 0.35, 0.035, 0.3, 0.3005", y);
    y = addWriteLine(doc, "    Order: __________________________________________________", y);
    y = addWriteLine(doc, "b)  Write a decimal that is between 3.14 and 3.15: __________", y);
    y = addWriteLine(doc, "c)  0.7 divided by 10 = ______      and      0.7 multiplied by 10 = ______", y);

    y = addSectionHeading(doc, "2 - Integer operations (temperature change)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Use a number line if it helps. Moving up adds, moving down subtracts.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  -3 + 5 = ______       b)  4 - 7 = ______       c)  -2 - 6 = ______", y);
    y = addWriteLine(doc, "d)  The temperature is -4 degrees and rises 9 degrees. New temperature: ______", y);

    y = addSectionHeading(doc, "3 - Profit and loss running total", y, { color: C.PRIMARY });
    y = addBodyText(doc, "A stall's week: Mon +12, Tue -8, Wed -5, Thu +3, Fri +10 (dollars).", y);
    y = addWriteLine(doc, "a)  Running total after each day: ___ , ___ , ___ , ___ , ___", y);
    y = addWriteLine(doc, "b)  Did the week end in profit or loss, and by how much? __________", y);

    doc.addPage();
    y = 50;
    y = addSectionHeading(doc, "4 - The coordinate plane (four quadrants)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Plot and label these points on the grid below: P(3, 2), Q(-4, 1), R(-2, -3), S(1, -4).", y);
    y = drawCoordinateGrid(doc, y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Which point is in the third quadrant (bottom-left)? ______", y);

    addPdfFooter(doc, `Year 8 Extension Challenge | Decimals & Integers Unit`);
    await writePdf(doc, path.join(OUT_DIR, Y8_RES.fileName));
    console.log("PDF written: " + Y8_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: Y8_KEY_RES.name });
    let y = addPdfHeader(doc, Y8_KEY_RES.name, {
      subtitle: "Worked answers for the Year 8 extension challenge.",
      color: C.SECONDARY,
      lessonInfo: `Decimals & Integers Unit | Year 8 stretch`,
    });

    y = addSectionHeading(doc, "1 - Decimals", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  0.035, 0.3, 0.3005, 0.305, 0.35   (lined up: 0.0350, 0.3000, 0.3005, 0.3050, 0.3500).", y);
    y = addBodyText(doc, "b)  Any value between 3.14 and 3.15, for example 3.145.", y);
    y = addBodyText(doc, "c)  0.7 divided by 10 = 0.07.   0.7 multiplied by 10 = 7.", y);

    y = addSectionHeading(doc, "2 - Integer operations", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  -3 + 5 = 2      b)  4 - 7 = -3      c)  -2 - 6 = -8      d)  -4 + 9 = 5 degrees.", y);

    y = addSectionHeading(doc, "3 - Running total", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  +12, +4, -1, +2, +12.", y);
    y = addBodyText(doc, "b)  The week ended in profit, +12 dollars.", y);

    y = addSectionHeading(doc, "4 - Coordinate plane", y, { color: C.PRIMARY });
    y = addBodyText(doc, "P(3, 2) is top-right (quadrant 1). Q(-4, 1) is top-left (quadrant 2). R(-2, -3) is bottom-left (quadrant 3). S(1, -4) is bottom-right (quadrant 4).", y);
    y = addBodyText(doc, "The point in the third quadrant (bottom-left) is R(-2, -3).", y);

    addPdfFooter(doc, `Year 8 Extension Answer Key | Decimals & Integers Unit`);
    await writePdf(doc, path.join(OUT_DIR, Y8_KEY_RES.fileName));
    console.log("PDF written: " + Y8_KEY_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

// Draw a simple 4-quadrant coordinate grid on the PDF (returns y after the grid).
function drawCoordinateGrid(doc, y, opts) {
  const o = opts || {};
  const { hex } = require("../themes/pdf_helpers");
  const margin = 50;
  const size = 230;            // grid square size in points
  const cells = 10;            // -5..5
  const step = size / cells;   // 23 pt per unit
  const x0 = margin + (495.28 - size) / 2;
  // Page-fit guard: keep the whole grid on one page.
  if (y + size + 30 > 772) { doc.addPage(); y = margin; }
  const yTop = y + 6;
  const color = hex(o.color || "1B3A6B");
  const lineCol = "#9CA3AF";

  // light gridlines
  doc.save();
  doc.lineWidth(0.4).strokeColor(lineCol);
  for (let i = 0; i <= cells; i++) {
    doc.moveTo(x0 + i * step, yTop).lineTo(x0 + i * step, yTop + size).stroke();
    doc.moveTo(x0, yTop + i * step).lineTo(x0 + size, yTop + size - size + i * step).stroke();
  }
  doc.restore();

  // axes (bold) through the centre
  const cx = x0 + size / 2;
  const cy = yTop + size / 2;
  doc.save();
  doc.lineWidth(1.4).strokeColor(color);
  doc.moveTo(x0, cy).lineTo(x0 + size, cy).stroke();   // x-axis
  doc.moveTo(cx, yTop).lineTo(cx, yTop + size).stroke(); // y-axis
  doc.restore();

  // axis number labels (-5, 5 on each axis)
  doc.fontSize(8).font("Sans").fillColor("#4B5563");
  doc.text("-5", x0 - 2, cy + 3, { width: step, align: "left" });
  doc.text("5", x0 + size - step, cy + 3, { width: step, align: "right" });
  doc.text("5", cx + 4, yTop - 1, { width: 14, align: "left" });
  doc.text("-5", cx + 4, yTop + size - 9, { width: 14, align: "left" });
  doc.fontSize(9).font("Sans-Bold").fillColor(color);
  doc.text("x", x0 + size + 4, cy - 4);
  doc.text("y", cx - 3, yTop - 12);

  return yTop + size + 12;
}

build().catch((err) => { console.error(err); process.exit(1); });
