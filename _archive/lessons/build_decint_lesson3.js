"use strict";

// Decimals & Integers Unit (Year 6 Numeracy) - Lesson 3: Integers on number lines
// VC2M6N01. Extend the number line below zero; locate and order integers; relationship to zero.
// Daily Review: Shapes (count the sides). Fluency: vertical addition algorithm with decimals.
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

const SESSION = 3;
const TOTAL = 4;
const UNIT_TITLE = "Decimals and Integers";
const FOOTER = `Decimals & Integers | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/DecInt_Lesson3_Integers_On_Number_Lines";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Integers on Number Lines",
  "Place integers on horizontal and vertical number lines, compare them, and order sets around zero.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the integers on number lines practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 3 of our unit. Today we meet a brand new kind of number.
- We are going below zero, to negative numbers, called integers.
- We will place them on a number line that goes left and right, and one that goes up and down.

DO:
- Have whiteboards, markers and a printed number line ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 3 of 4. This is a fresh strand - integers. It does not depend on the decimal lessons, so a student who missed Lessons 1 or 2 can fully access today.

WATCH FOR:
- Students who have seen negative numbers on a thermometer - draw that out in the launch.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do section.

DO:
- Print one practice sheet and answer key per student.
- Have whiteboards, markers and a printed number line that includes negatives.
- Optional: a large thermometer picture for the launch.

TEACHER NOTES:
One student practice sheet plus answer key. The sheet has a built-in enabling start (a number line already drawn) and an extension that orders larger sets and reasons about distance from zero.

CATCH-UP NOTE:
Integers are a new topic, so a returning student starts fresh here with everyone. No decimal knowledge is needed today. Hand a returner the printed number line and they are ready.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are naming shapes by their sides.
- Count the sides of each shape and write the number on your whiteboard.

DO:
- Display the three shapes.
- 45 seconds.
- Walk and check counting.

TEACHER NOTES:
Prior geometry, not today's integers. Counting sides keeps shape names fresh. A is an octagon, B is a trapezoid, C is a hexagon.

WATCH FOR:
- Students who count carefully - secure.
- Students who miscount the octagon - prompt them to touch each side once.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Shape A, the octagon - 8 sides.
- Shape B, the trapezoid - 4 sides.
- Shape C, the hexagon - 6 sides.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick-and-fix.

TEACHER NOTES:
Note any student who confuses hexagon and octagon - quick small group with shape cards.

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Vertical addition, now with decimals.
- The big rule: line up the decimal points, so tenths sit under tenths.
- Set each one out and add.

DO:
- Display the three prompts.
- 2 minutes.
- Scan for lined-up decimal points.

TEACHER NOTES:
Same vertical algorithm as the whole unit, applied to decimals. Lining up the decimal point is the same as lining up place value columns. This bridges our decimal work into the algorithm.

WATCH FOR:
- Students who line up the points - secure.
- Students who right-align the digits instead of the points - reteach: points under points.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 3.4 plus 2.75 is 6.15.
- 12.6 plus 4.38 is 16.98.
- 0.85 plus 9.5 is 10.35.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Adding a placeholder zero helps: 3.4 becomes 3.40, 9.5 becomes 9.50. That keeps the columns lined up.

WATCH FOR:
- Students who self-correct - secure.
- Students who add 3.4 and 2.75 as if both had two places without lining up - reteach with the placeholder zero.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Picture a cold morning. The thermometer reads 5 degrees. Then it drops.
- It passes 4, 3, 2, 1, and zero. What happens if it keeps dropping?
- It goes below zero, to negative numbers. Negative 1, negative 2.
- Today we put numbers below zero onto a number line.

DO:
- Point up and down the thermometer scale as you speak.
- Take a couple of real examples - fridge, freezer, frosty morning.
- Reveal the idea of below zero.

TEACHER NOTES:
This launch uses temperature to activate the everyday idea of below zero before the learning intention. Negative numbers feel natural on a thermometer.

WATCH FOR:
- Students who name a freezer temperature like negative 18 - strong prior knowledge.
- Students unsure - reassure, the number line will make it clear.

[Stage: Launch | VTLM 2.0: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to place and compare integers on a number line, including below zero.
- Now the three I can statements. Read them with me.

DO:
- Choral read the LI and SC.
- Hold up the printed number line that includes negatives.

TEACHER NOTES:
SC1 is achievable for everyone - placing a positive and a negative either side of zero. SC2 is the core target the exit ticket checks. SC3 stretches to ordering a set and reasoning about distance from zero.

WATCH FOR:
- Students who can repeat integer and negative - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Words for today.
- Integers are whole numbers and their negatives: -3, -2, -1, 0, 1, 2, 3. No fractions or decimals.
- A negative number is below zero. We write a minus sign in front: negative three is -3.
- The sign shows direction from zero. Plus is to the right or up. Minus is to the left or down.
- Zero is the boundary - not positive, not negative.

DO:
- Point to each part of a number line as you say it.
- Have students show left with one hand for negative, right for positive.

TEACHER NOTES:
The key idea is that the sign shows direction from zero. Keep it to these words. We use this directly when we place numbers either side of zero.

WATCH FOR:
- Students who can point the correct direction for a sign - secure.
- Students who read -3 as just three - stress the minus means below zero.

[General: Key Vocabulary | VTLM 2.0: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Watch me build a number line that goes below zero.
- Zero sits in the middle. Positives go to the right, negatives to the left.
- Watch how I place positive 3. I count three steps right of zero.
- Now negative 3. I count three steps left of zero.
- Positive 3 and negative 3 are the same distance from zero, just opposite directions.
- And negative 3 is less than zero is less than positive 3.

DO:
- Point to zero first, then step right for +3, then left for -3.
- Stress same distance, opposite direction.
- Say negative 3 is less than 0 is less than positive 3.

TEACHER NOTES:
This is the core model. Numbers get bigger to the right and smaller to the left, straight through zero. The mirror image of +3 and -3 around zero is the key picture.

MISCONCEPTIONS:
- Misconception: students think -3 is bigger than -1 because 3 is bigger than 1.
  Why: they read the digit and ignore the sign and direction.
  Impact: they order negatives backwards.
  Quick correction: on the line, -3 is further LEFT than -1, so -3 is smaller.

WATCH FOR:
- Students who place -3 left of zero - secure.
- Students who place -3 to the right - re-point: minus means go left.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- The number line can also stand up. This is how a thermometer works.
- Zero is here. Above zero is warmer and positive. Below zero is colder and negative.
- Watch me find negative 2. I start at zero and step down two.
- Negative 2 is below zero. Negative 4 is even further down, so it is even colder.
- Up means bigger, down means smaller - the same idea, just vertical.

DO:
- Trace up for positive, down for negative.
- Place negative 2 below zero.
- Compare negative 2 and negative 4 by how far down they are.

TEACHER NOTES:
The vertical number line is the thermometer and the lift. Same rule as horizontal: further from zero in the minus direction means smaller. Lower equals colder equals smaller.

MISCONCEPTIONS:
- Misconception: students think negative 4 is warmer than negative 2.
  Why: 4 is bigger than 2 as a digit.
  Impact: they compare temperatures backwards.
  Quick correction: negative 4 is further down the thermometer, so it is colder and smaller.

WATCH FOR:
- Students who say negative 4 is colder - secure.
- Students who say negative 4 is warmer - re-point down the scale.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Which is greater: negative 2 or negative 5?
- Mark both on the number line, then circle the greater one.

DO:
- Display the prompt and the line.
- 45 seconds.
- Walk and scan for left-right placement.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: circle the greater number on three, two, one, show.
- Scan for: -2 circled, placed to the right of -5.
PROCEED: If about 80 percent circle -2, click to reveal and move to We Do.
PIVOT: Most likely misconception - students circle -5 because 5 is bigger than 2.
- Reteach: mark both on the line. -2 is closer to zero, so it is greater. -5 is further left.
- Re-check: which is further left, and which is greater?

TEACHER NOTES:
This is the central integer misconception. The number line resolves it: greater means further right, even among negatives.

WATCH FOR:
- Students who circle -2 with reasoning - secure.
- Students who circle -5 - digit-not-sign misconception, reteach on the line.

[Stage 2: CFU | VTLM 2.0: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Order these from smallest to largest: -4, 2, -1, 0, 3.
- Picture each one on the number line.
- Smallest is furthest left, largest is furthest right.

DO:
- Display the five numbers and a blank line.
- 90 seconds.
- Listen for furthest-left reasoning.

TEACHER NOTES:
Ordering mixes negatives, zero and positives. Furthest left is smallest. -4 is furthest left; 3 is furthest right.

WATCH FOR:
- Pairs who start from the left of the line - secure.
- Pairs who put -1 before -4 - reteach: -4 is further left, so smaller.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Reading left to right on the line: -4, then -1, then 0, then 2, then 3.
- So smallest to largest is -4, -1, 0, 2, 3.

DO:
- Click to reveal.
- Trace left to right along the line.

TEACHER NOTES:
The reveal shows the order is just reading the number line left to right. Smallest sits furthest left.

WATCH FOR:
- Students who self-correct - secure.
- Students who reversed the negatives - small group before You Do.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, on the vertical line - temperatures.
- Order these from coldest to warmest: 3 degrees, -4 degrees, -1 degrees, 0 degrees.
- Coldest is furthest down, warmest is furthest up.

DO:
- Display the four temperatures and the thermometer scale.
- 75 seconds.
- Listen for furthest-down reasoning.

TEACHER NOTES:
Same ordering on a vertical scale. Coldest equals lowest equals smallest. -4 is coldest, 3 is warmest.

WATCH FOR:
- Pairs who put -4 at the bottom - secure.
- Pairs who put -1 below -4 - reteach: -4 is further down the thermometer.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- From coldest to warmest: -4, then -1, then 0, then 3.
- -4 is furthest down, so coldest. 3 is furthest up, so warmest.

DO:
- Click to reveal the marked thermometer.
- Point to how far each sits from zero.

TEACHER NOTES:
If pairs struggled, walk the scale from the bottom up before You Do.

WATCH FOR:
- Students who justify coldest-is-lowest - ready for independent work.
- Students who reversed it - enabling group for You Do.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1: mark each integer on the number line.
- Section 2: compare two integers with greater than or less than.
- Section 3: order each set from smallest to largest.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for further-left, further-down reasoning.
- Cold call one or two students to compare two negatives.

TEACHER NOTES:
Different numbers from the We Do, same idea: place on the line, then greater means further right or further up.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: The number line is already drawn with marks. Do Sections 1 and 2 only, pointing to the line each time.
- Extra Notes: Sit with these students and place the first integer together.
EXTENDING PROMPT:
- Task: Extension - order a set of six integers, and answer which of two integers is further from zero (distance), with a reason.
- Extra Notes: Distance from zero is a new idea - separate it from greater or smaller. Strong students can preview the Year 8 challenge sheet in the unit.

WATCH FOR:
- Students who order using the line - secure.
- Students who order negatives backwards - prompt them to read left to right.

[Stage 4: You Do | VTLM 2.0: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Mark -3 and 4 on a number line, with zero in the middle.
- Then order these from smallest to largest: -2, 1, -5, 0.

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses SC2 - placing integers either side of zero and ordering a mixed set. Look for -3 to the left of zero and -5 placed as the smallest.

WATCH FOR:
- Students who order -5, -2, 0, 1 - secure.
- Students who put -2 before -5 - digit-not-sign misconception, revisit next lesson.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: on a number line, how do you know which of two numbers is greater?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that greater means further right (or further up), even for negatives, because the sign shows direction from zero. Students who can say this are ready for integers in real contexts next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on SC2 - small group revision at the start of Lesson 4.

[General: Closing | VTLM 2.0: Reflection]`;

// --- Helpers -----------------------------------------------------------------

// Vertical number line (thermometer / lift). labels[0] is the TOP (largest value).
function addVerticalNumberLine(slide, cx, yTop, h, labels, markedIdx, opts) {
  const o = opts || {};
  const n = labels.length - 1;
  if (n <= 0) return;
  const intervalH = h / n;
  const lineColor = o.lineColor || C.CHARCOAL;

  slide.addShape("line", { x: cx, y: yTop, w: 0, h: h, line: { color: lineColor, width: 2.5 } });
  slide.addShape("line", { x: cx - 0.1, y: yTop + 0.1, w: 0.1, h: -0.1, line: { color: lineColor, width: 2 } });
  slide.addShape("line", { x: cx, y: yTop + 0.1, w: 0.1, h: -0.1, line: { color: lineColor, width: 2 } });

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
        x: cx + 0.18, y: my - 0.13, w: 1.3, h: 0.26,
        fontSize: 12, fontFace: FONT_B, color: m.color || C.ALERT, bold: true,
        align: "left", valign: "middle", margin: 0,
      });
    }
  });
}

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 3: Integers on number lines",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Shapes (count the sides)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Daily Review: How many sides?", { color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      const shapeY = CONTENT_TOP + 0.45;
      s.addShape("octagon", {
        x: 1.35, y: shapeY, w: 1.8, h: 1.6,
        fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("A", { x: 1.35, y: shapeY + 1.7, w: 1.8, h: 0.4, fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });

      s.addShape("trapezoid", {
        x: 4.05, y: shapeY + 0.25, w: 1.9, h: 1.2,
        fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("B", { x: 4.05, y: shapeY + 1.7, w: 1.9, h: 0.4, fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });

      s.addShape("hexagon", {
        x: 6.95, y: shapeY, w: 1.8, h: 1.6,
        fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("C", { x: 6.95, y: shapeY + 1.7, w: 1.8, h: 0.4, fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });

      s.addText("Count the sides of each shape. Write the number on your whiteboard.", {
        x: 0.7, y: SAFE_BOTTOM - 0.5, w: 8.6, h: 0.35,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A octagon = 8     B trapezoid = 4     C hexagon = 6", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - vertical addition with decimals
  withReveal(
    () => fluencySlide(pres, "Fluency: Vertical addition (decimals)",
      ["3.4 + 2.75", "12.6 + 4.38", "0.85 + 9.5"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "6.15        16.98        10.35", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - below zero (thermometer)
  contentSlide(pres, "Launch", C.ACCENT, "What is colder than zero?",
    [
      "A thermometer reads 5 degrees.",
      "It drops past 4, 3, 2, 1, 0...",
      "",
      "What happens if it keeps dropping?",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.ACCENT });
      slide.addText("Below zero", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
      });
      addVerticalNumberLine(slide, lg.rightX + 2.1, lg.panelTopPadded + 0.55, 2.0,
        ["3", "2", "1", "0", "-1", "-2", "-3"], [{ idx: 4, label: "below zero", color: C.ALERT }, { idx: 5, color: C.ALERT }, { idx: 6, color: C.ALERT }],
        { zeroIdx: 3, fontSize: 13 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to place and compare integers on a number line, including below zero.",
    [
      "I can place a positive and a negative integer either side of zero.",
      "I can compare two integers using a number line and the words greater and smaller.",
      "I can order a set of integers and explain it using distance and direction from zero.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Integers and the sign",
    [
      "Integers: ... -2, -1, 0, 1, 2 ... (whole numbers and their negatives).",
      "Negative = below zero, written with a minus: -3.",
      "The sign shows direction from zero: + right/up, - left/down.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
      slide.addText("The sign points the way", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
      });
      const nlY = lg.panelTopPadded + 1.3;
      addNumberLine(slide, lg.rightX + 0.55, nlY, lg.rightW - 1.1,
        ["-2", "-1", "0", "1", "2"], [2]);
      addTextOnShape(slide, "<-- negative", {
        x: lg.rightX + 0.35, y: nlY + 0.45, w: 1.7, h: 0.34, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
      addTextOnShape(slide, "positive -->", {
        x: lg.rightX + lg.rightW - 2.05, y: nlY + 0.45, w: 1.7, h: 0.34, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 10: I Do #1 - wide horizontal number line with +3 and -3
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Either side of zero", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.9, { strip: C.PRIMARY });
    s.addText([
      { text: "Zero sits in the middle of the line.", options: { fontSize: 18, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "+3 is 3 steps RIGHT of zero.   -3 is 3 steps LEFT of zero.", options: { fontSize: 17, color: C.CHARCOAL, breakLine: true } },
      { text: "Same distance, opposite direction.   ", options: { fontSize: 17, color: C.PRIMARY, bold: true } },
      { text: "So  -3 < 0 < +3.", options: { fontSize: 17, color: C.ALERT, bold: true } },
    ], {
      x: 0.8, y: CONTENT_TOP + 0.22, w: 8.4, h: 1.5, fontFace: FONT_B, valign: "middle", margin: 0,
    });

    const nlY = 3.5;
    addNumberLine(s, 0.9, nlY, 8.4,
      ["-5", "-4", "-3", "-2", "-1", "0", "1", "2", "3", "4", "5"], [2, 8]);
    s.addText("The orange dots show -3 and +3: mirror images around zero.", {
      x: 0.8, y: nlY + 0.5, w: 8.4, h: 0.3,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO1);
  })();

  // Slide 11: I Do #2 - vertical number line (thermometer)
  workedExSlide(pres, 2, "I Do", "The line can stand up",
    [
      "A thermometer is a vertical line.",
      "",
      "Above zero: warmer, positive.",
      "Below zero: colder, negative.",
      "",
      "-2 is two steps below zero.",
      "-4 is even further down -",
      "even colder, even smaller.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.PRIMARY });
      slide.addText("Up = warmer   Down = colder", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.26,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addVerticalNumberLine(slide, lg.rightX + 1.85, lg.panelTopPadded + 0.50, 2.3,
        ["4", "3", "2", "1", "0", "-1", "-2", "-3", "-4"], [{ idx: 6, label: "-2 here", color: C.ALERT }],
        { zeroIdx: 4, fontSize: 12 });
    }
  );

  // Slides 12-13: CFU + reveal - which is greater, -2 or -5 (wide line)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Which is greater?", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 8.0, y: 0.20, w: 1.4, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 1.6, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:  which is greater,  -2  or  -5 ?", options: { fontSize: 20, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Mark both on the line. Circle the greater one.", options: { fontSize: 16, color: C.ALERT, bold: true } },
      ], {
        x: 0.8, y: CONTENT_TOP + 0.2, w: 8.4, h: 1.15, fontFace: FONT_B, valign: "middle", margin: 0,
      });

      const nlY = 3.45;
      addNumberLine(s, 0.9, nlY, 8.4,
        ["-6", "-5", "-4", "-3", "-2", "-1", "0", "1", "2"], []);
      s.addText("Mark -2 and -5. Which one sits further right?", {
        x: 0.8, y: nlY + 0.5, w: 8.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "-2 is greater. It is closer to zero (further right). -5 is further left.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - order integers on a wide horizontal line
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Order from smallest to largest", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 1.9, { strip: C.SECONDARY });
      s.addText([
        { text: "With your partner.  Order:   -4,   2,   -1,   0,   3", options: { fontSize: 19, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Smallest = furthest LEFT.    Largest = furthest RIGHT.", options: { fontSize: 16, color: C.SECONDARY } },
      ], {
        x: 0.8, y: CONTENT_TOP + 0.25, w: 8.4, h: 1.4, fontFace: FONT_B, valign: "middle", margin: 0,
      });

      const nlY = 3.45;
      addNumberLine(s, 0.9, nlY, 8.4,
        ["-5", "-4", "-3", "-2", "-1", "0", "1", "2", "3", "4", "5"], [1, 4, 5, 7, 8]);
      s.addText("Each orange dot is one of the five numbers.", {
        x: 0.8, y: nlY + 0.5, w: 8.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Smallest to largest:   -4  <  -1  <  0  <  2  <  3", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - order temperatures on a vertical line
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Coldest to warmest",
      [
        "With your partner.",
        "",
        "Order these temperatures:",
        "3, -4, -1, 0 degrees.",
        "",
        "Coldest = furthest DOWN.",
        "Warmest = furthest UP.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.SECONDARY });
        slide.addText("Thermometer scale", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.26,
          fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        addVerticalNumberLine(slide, lg.rightX + 1.85, lg.panelTopPadded + 0.50, 2.3,
          ["4", "3", "2", "1", "0", "-1", "-2", "-3", "-4"], [],
          { zeroIdx: 4, fontSize: 12 });
      }),
    (slide) => {
      addTextOnShape(slide, "Coldest to warmest:   -4  <  -1  <  0  <  3", {
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
      { text: "mark each integer on the number line.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "compare two integers with > or <.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "order each set, smallest to largest.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Greater = further RIGHT (or UP). The sign shows direction from zero.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    addNumberLine(s, 0.9, panelY + 1.55, 8.2,
      ["-5", "-4", "-3", "-2", "-1", "0", "1", "2", "3", "4", "5"], []);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Mark -3 and 4 on a number line with zero in the middle.",
      "Order from smallest to largest: -2, 1, -5, 0.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: on a number line, how do you know which of two numbers is greater?",
      scItems: [
        "I can place a positive and a negative integer either side of zero.",
        "I can compare two integers using a number line and the words greater and smaller.",
        "I can order a set of integers and explain it using distance and direction from zero.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecInt_Lesson3_Integers_On_Number_Lines.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Place, compare, and order integers using number lines.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Greater means further RIGHT on a horizontal line, or further UP on a vertical line. The sign shows direction from zero: + is right/up, - is left/down. -5 is smaller than -2.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Compare -2 and -5. On the line, -2 is further RIGHT (closer to zero) than -5. So -2 > -5.",
      y);

    y = addSectionHeading(doc, "Section 1 - Mark each integer on the number line", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Use the line below, which runs from -6 to 6. Mark and label each point.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  -4       b)  3       c)  -1       d)  5", y);

    y = addSectionHeading(doc, "Section 2 - Compare with > or <", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  -3 ____ 2", y);
    y = addWriteLine(doc, "b)  -6 ____ -1", y);
    y = addWriteLine(doc, "c)  0 ____ -4", y);

    y = addSectionHeading(doc, "Section 3 - Order each set from smallest to largest", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3, -2, 0, -5   ->  ____ , ____ , ____ , ____", y);
    y = addWriteLine(doc, "b)  -1, -7, 4, -3  ->  ____ , ____ , ____ , ____", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order from smallest to largest:  2, -8, -3, 6, 0, -1", y);
    y = addWriteLine(doc, "Order: ____________________________________________________________", y);
    y = addBodyText(doc, "Which is FURTHER FROM ZERO: -6 or 4? (This asks about distance, not greater or smaller.) Explain.", y);
    y = addWriteLine(doc, "Answer: ___________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Integers on Number Lines | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the integers on number lines practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Mark each integer", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  -4 is 4 steps left of zero.   b)  3 is 3 steps right.   c)  -1 is 1 step left.   d)  5 is 5 steps right.", y);

    y = addSectionHeading(doc, "Section 2 - Compare", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  -3 < 2      b)  -6 < -1      c)  0 > -4", y);

    y = addSectionHeading(doc, "Section 3 - Order (smallest to largest)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  -5, -2, 0, 3", y);
    y = addBodyText(doc, "b)  -7, -3, -1, 4", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order: -8, -3, -1, 0, 2, 6.", y);
    y = addBodyText(doc, "Further from zero: -6 is 6 away from zero; 4 is only 4 away. So -6 is further from zero, even though -6 is the smaller number.", y);

    y = addTipBox(doc,
      "Watch for: students who think -5 is greater than -2 because 5 is bigger; students who place negatives to the right of zero; students who confuse 'smaller' with 'further from zero'.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
