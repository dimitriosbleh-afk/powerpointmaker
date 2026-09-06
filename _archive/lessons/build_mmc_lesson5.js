"use strict";

// Metric Measurement & Unit Conversion (Year 6 Numeracy)
// Session 5 of 5: Choose, convert & compare — consolidation across length,
// mass and capacity.
// VC2M5M01 + VC2M6M01 — choosing units; converting with the correct operation;
//            comparing measurements by converting to a common unit first.
// Daily Review: Tessellations (prior learning).
// Fluency: Division bracket (long-division layout). NOT the bus stop method.
//
// Unit routine (consistent across all 5 sessions):
//   SET THE STAIRS  ->  COUNT THE STEPS  ->  MULTIPLY OR DIVIDE.
// This consolidation session re-anchors the routine (catch-up friendly) and
// adds the new move: convert to a common unit, THEN compare or solve.
// Ships the Year 8 Extension Challenge for students ready to push further.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Whole-unit palette: every session of this unit uses the SAME variant (0).
const UNIT_VARIANT = 0;
const T = createTheme("numeracy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, addRevealAnswerBar,
  STAGE_COLORS,
} = T;

const SESSION = 5;
const TOTAL = 5;
const UNIT_TITLE = "Metric Measurement & Conversion";
const FOOTER = `Metric Measurement | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/MMC_Session5_Choose_Convert_Compare";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Choose Convert Compare Practice",
  "Choose a unit, convert, and compare measurements in real problems.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the choose-convert-compare practice sheet.");
const EXTENSION_RES = makeSessionResource(SESSION,
  "Year 8 Extension Challenge",
  "Stretch task: area, volume and rate conversions for students ready to push on.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXTENSION_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back, mathematicians. This is our final session, where it all comes together.
- We will choose units, convert between them, and use converting to compare and solve real problems.
- The big new move today: when two measurements use different units, convert them to the same unit first.

DO:
- Have whiteboards and markers ready.
- Keep the printed practice sheet face down until the You Do section.

TEACHER NOTES:
Session 5 of 5. A consolidation lesson. It re-anchors the staircase routine so students who missed sessions can still join, and adds the compare-by-converting move. The Year 8 Extension Challenge is ready for early finishers.

WATCH FOR:
- Students who compare numbers without checking the units first. That is the habit we break today.

[General: Title | VTLM 2.0: Attention, focus and regulation]`;

const NOTES_RESOURCES = `SAY:
- Here is what we are using today.
- Whiteboards for most of the lesson, and the practice sheet near the end.
- There is also a Year 8 Extension Challenge for anyone ready to push further.

DO:
- Print one practice sheet and one answer key per student who needs paper.
- Print a few Year 8 Extension Challenge sheets for early finishers.

TEACHER NOTES:
Three resources today: a consolidation practice sheet, its answer key, and the Year 8 Extension Challenge (area, volume and rate conversions). The extension is a stretch task and can be offered any week, not only today.

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_DR_Q = `SAY:
- Quick warm up on tessellations from our geometry work.
- Read each one, then write your answer on your whiteboard.
- These are from earlier learning, not today's measurement work.

DO:
- Display the three prompts and the tessellation pattern.
- Allow about 90 seconds.

TEACHER NOTES:
Daily Review revisits tessellations. The pattern shown is squares tiling with no gaps or overlaps. This is prior learning and retrieval, separate from today's conversion focus.

WATCH FOR:
- Students who name squares, triangles or hexagons - secure.
- Students who think every shape tessellates - revisit gaps and overlaps.

[Stage 1: Daily Review | VTLM 2.0: Retention and recall]`;

const NOTES_DR_A = `SAY:
- Let us check together.
- Squares, equilateral triangles, rectangles and regular hexagons all tessellate on their own.
- A regular pentagon does not tessellate on its own - false.
- Shapes tessellate when they fit together with no gaps and no overlaps, and the angles meet to make 360 degrees at a point.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The key idea is no gaps, no overlaps, angles meeting to 360 at a vertex.

WATCH FOR:
- Confident self-correction - secure.
- Students who said a pentagon tessellates - show the gap it leaves.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency now. Division using the division bracket.
- Set each one out with the division bracket and work from the left.
- Whisper your answer, then write it.

DO:
- Display the three divisions.
- Allow about 80 seconds.

TEACHER NOTES:
Fluency builds division using the division bracket, the same focus across all five sessions. Please call it the division bracket. Brisk retrieval, not new teaching.

WATCH FOR:
- Clean carrying of remainders - secure.
- Dropped remainders - reteach the carry step with the division bracket.

[Stage 1: Fluency | VTLM 2.0: Retention and recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check your division brackets.
- 176 divided by 8 is 22.
- 261 divided by 9 is 29.
- 210 divided by 6 is 35.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
261 divided by 9: 26 divided by 9 is 2 remainder 8, carry the 8 to make 81, 81 divided by 9 is 9.

WATCH FOR:
- Clean carrying - secure.
- Answers near 30 for 261 divided by 9 - re-check the carry.

[Stage 1: Fluency Answers | VTLM 2.0: Retention and recall]`;

const NOTES_LAUNCH = `SAY:
- Here is a puzzle. Which is longer: 1500 metres or 2 kilometres?
- You cannot compare them straight away, because they use different units.
- Turn and tell your partner: how could you make them easy to compare?

DO:
- Display the two measurements.
- Give 40 seconds of partner talk.
- Take responses. Steer towards converting one of them.

TEACHER NOTES:
The launch sets up today's big move: convert to a common unit, then compare. 2 km is 2000 m, which is longer than 1500 m. Do not rush to the answer; let students realise they must convert first.

WATCH FOR:
- Students who suggest converting km to m - exactly the move.
- Students who say 1500 is bigger than 2 so 1500 m wins - the unit trap we fix today.

[Stage 1: Launch | VTLM 2.0: Knowledge and memory]`;

const NOTES_LI_SC = `SAY:
- Here is what we are learning today.
- We are learning to convert measurements so we can compare them and solve problems.
- Read the success criteria with me.

DO:
- Choral read the learning intention.
- Read each I can statement.

TEACHER NOTES:
SC1 is reachable for everyone - choosing a unit and converting. SC2 is the core target the exit ticket assesses - converting to a common unit to compare. SC3 stretches into multi-step problems. Tier labels stay off the slide.

WATCH FOR:
- Students repeating the language back - tracking with us.

[General: LI/SC | VTLM 2.0: Planning]`;

const NOTES_RECAP = `SAY:
- Quick recap of our routine, in case you missed a session or want a reminder.
- Set the stairs: largest unit at the top, smallest at the bottom.
- Count the steps between the units. For mass and capacity every step is times one thousand. For length the steps are times one thousand, times one hundred, then times ten.
- Down to a smaller unit, multiply. Up to a bigger unit, divide.

DO:
- Point to the staircase as you say each move.
- Have students whisper the three moves to a partner.

TEACHER NOTES:
This slide re-anchors the whole unit in one place. It is the catch-up safety net: a student who missed Sessions 2 to 4 can convert from this slide alone. Length is shown; remind them mass and capacity use the same routine with x1000 steps.

WATCH FOR:
- Students who recite the three moves - secure.
- Students still unsure of direction - anchor: smaller unit means more of them.

[General: Recap | VTLM 2.0: Knowledge and memory]`;

const NOTES_IDO1 = `SAY:
- Let us solve our launch puzzle properly. Which is longer: 1500 metres or 2 kilometres?
- I cannot compare different units, so I convert. I will change kilometres into metres.
- Kilometres down to metres is times one thousand. 2 times 1000 is 2000.
- Now both are in metres. 2000 metres against 1500 metres.
- 2000 is more than 1500, so 2 kilometres is longer. Converting first made it easy.

DO:
- Show 2 km becoming 2000 m on the panel.
- Underline both measurements in metres before comparing.

TEACHER NOTES:
This is the core skill of the lesson: convert to a common unit, then compare. Always convert to the smaller unit if it keeps whole numbers, but either works.

MISCONCEPTIONS:
- Misconception: 1500 m is longer because 1500 is bigger than 2.
  Why: students compare the numbers, ignoring the units.
  Impact: they choose the wrong measurement.
  Quick correction: you cannot compare until the units match. Convert first, then compare.

WATCH FOR:
- Students who convert before comparing - secure.
- Students who compare 1500 and 2 directly - stop and convert first.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a multi-step problem. A bottle holds 2 litres. You drink 600 millilitres. How much is left, in millilitres?
- The units are different, so I convert first. 2 litres down to millilitres is times one thousand, so 2000 millilitres.
- Now both are in millilitres. 2000 take away 600 is 1400.
- So 1400 millilitres are left. Convert first, then calculate.

DO:
- Show 2 L becoming 2000 mL on the panel.
- Write 2000 minus 600 equals 1400.

TEACHER NOTES:
A two-step problem: convert, then operate. Students often subtract before converting and get a nonsense answer. The order matters.

MISCONCEPTIONS:
- Misconception: 2 minus 600 or 2 take 600 in mixed units.
  Why: students subtract before converting.
  Impact: impossible or negative answers.
  Quick correction: make the units match first, then subtract.

WATCH FOR:
- Students who convert then subtract - secure.
- Students who subtract across units - reorder the steps.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO3 = `SAY:
- One more comparison, this time with mass. Which is heavier: half a kilogram or 800 grams?
- Different units, so I convert. Half a kilogram, that is 0.5 kilograms, down to grams is times one thousand. 0.5 times 1000 is 500.
- Now both are in grams. 500 grams against 800 grams.
- 800 is more than 500, so 800 grams is heavier.

DO:
- Show 0.5 kg becoming 500 g on the panel.
- Compare 500 g and 800 g side by side.

TEACHER NOTES:
Reinforces convert-then-compare with a decimal and a different attribute. Shows the routine transfers across length, mass and capacity.

MISCONCEPTIONS:
- Misconception: half a kilogram must be heavier because a kilogram is the big unit.
  Why: students judge by the unit name, not the amount.
  Impact: they pick the wrong one.
  Quick correction: convert to the same unit. 0.5 kg is only 500 g, which is less than 800 g.

WATCH FOR:
- Students who convert the decimal correctly - secure and ready for depth.
- Students unsure about 0.5 times 1000 - re-show the three-place move.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU1_Q = `SAY:
- Quick check on your whiteboards.
- Which is longer: 3 metres or 250 centimetres? Convert to the same unit first, then tell me which.
- Show me on three, two, one.

DO:
- Display the prompt.
- Allow about 40 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Convert, then show me which is longer, on three, two, one, show."
- Scan for: 3 m (since 3 m is 300 cm, more than 250 cm).
PROCEED: If 80 percent show 3 m with 300 cm shown, click to reveal and move on.
PIVOT: Most likely misconception - students compare 3 and 250 directly and pick 250 cm.
- Reteach: "Convert first. 3 metres is 300 centimetres. 300 is more than 250, so 3 metres is longer."
- Re-check: "Which is longer: 4 m or 450 cm?"

TEACHER NOTES:
Checks the convert-then-compare habit before guided practice.

WATCH FOR:
- 3 m with 300 cm shown - secure.
- 250 cm chosen - they compared without converting.

[Stage 2: CFU | VTLM 2.0: Supported Application]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. A jug holds 1.5 litres. You pour 400 millilitres into a cup. How many millilitres are left in the jug?
- Step one, convert the litres to millilitres. Step two, subtract. Show your working with your partner.

DO:
- Keep the steps visible.
- Walk and scan. Prompt with "convert first" where needed.
- Take 60 seconds before revealing.

TEACHER NOTES:
Guided multi-step practice. Convert 1.5 L to 1500 mL, then subtract 400. Watch for students subtracting before converting.

WATCH FOR:
- Pairs who convert then subtract - secure.
- Pairs who try 1.5 minus 400 - reorder to convert first.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 1.5 litres is 1500 millilitres. Times one thousand, down to a smaller unit.
- 1500 take away 400 is 1100.
- So 1100 millilitres are left.

DO:
- Click to reveal.
- Ask one pair to explain why they converted first.

TEACHER NOTES:
The reveal restates convert, then subtract. 1100 mL is the answer.

WATCH FOR:
- Self-correction to 1100 mL - secure.
- Answers like 1.1 or 1460 - they mixed units or subtracted first.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Anya says 1 kilogram is lighter than 900 grams, because 1 is smaller than 900.
- Do not just fix it. Decide: true or false?
- Thumbs up for true, thumbs down for false.

DO:
- Display the claim.
- Give 15 seconds of thinking time.

CFU CHECKPOINT:
Technique: Thumbs Up or Down
Script:
- Say: "Thumbs up if true, thumbs down if false."
- Scan for: thumbs down. 1 kg is 1000 g, which is more than 900 g.
PROCEED: If 80 percent thumbs down with a reason, click to reveal.
PIVOT: Most likely misconception - comparing the numbers instead of converting to the same unit.
- Reteach: "Convert first. 1 kilogram is 1000 grams. 1000 is more than 900, so 1 kilogram is heavier."
- Re-check: "Which is heavier: 2 kg or 1500 g?"

TEACHER NOTES:
This is the threshold idea for the whole unit: you must convert to the same unit before comparing. The number alone does not tell you which is bigger.

WATCH FOR:
- Confident thumbs down naming 1000 g - secure.
- Thumbs up - revisit convert before comparing.

[Stage 2: CFU Hinge | VTLM 2.0: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Time to work on your own. Take the practice sheet.
- Section one, choose and convert. Section two, compare by converting to the same unit. Section three, multi-step problems.
- Always convert to the same unit before you compare or calculate.

DO:
- Distribute the practice sheet.
- Circulate. Sit with the enabling group first.
- Offer the Year 8 Extension Challenge to anyone who finishes confidently.

TEACHER NOTES:
The sheet builds from converting to comparing to multi-step problems. Section two is the threshold - comparing by converting first.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the staircase box and the worked first item. Do the comparing items with both measurements already part-converted.
- Extra Notes: Sit with this group for the first compare item.
EXTENDING PROMPT:
- Task: The Year 8 Extension Challenge - area, volume and rate conversions. A real stretch beyond Year 6.
- Extra Notes: Encourage them to explain why a squared unit converts by 100 times 100, not just 100.

WATCH FOR:
- Students who convert before comparing - secure.
- Students who compare raw numbers - return to convert first.

[Stage 4: You Do | VTLM 2.0: Mastery and application]`;

const NOTES_EXIT = `SAY:
- Last task, on your own, on your whiteboard.
- One, which is longer: 2 kilometres or 1800 metres? Convert to the same unit and explain how you know.
- Two, a bottle holds 1.5 litres. You drink 750 millilitres. How many millilitres are left?

DO:
- Display the prompt.
- Allow about 4 minutes.
- Collect whiteboards or take a photo for your records.

TEACHER NOTES:
Exit ticket assesses SC2 (converting to a common unit to compare) and reaches SC3 (a multi-step problem). Internal target is SC2. Do not display any SC label to students.

WATCH FOR:
- 2 km is longer (2000 m more than 1800 m) and 750 mL left - secure.
- A comparison made without converting - revisit convert first next unit.

[Stage 5: Exit Ticket | VTLM 2.0: Mastery and application]`;

const NOTES_CLOSING = `SAY:
- Look back at our success criteria for the whole unit.
- Show me thumbs for each one - got it, getting there, or need more practice.
- Turn and tell your partner: what is the one rule you will never forget about converting units?

DO:
- Read each I can statement.
- Use thumbs to self-assess.
- Note who is still unsure so you can plan revisits.

TEACHER NOTES:
The unit's big idea: a smaller unit needs a bigger number, and you must match the units before you compare. Celebrate that they can now choose units, convert, and compare across length, mass and capacity.

WATCH FOR:
- Students who state the direction rule clearly - secure.
- Students unsure - note for a short revisit and offer the staircase reminder.

[General: Closing | VTLM 2.0: Mastery and application]`;

// ─── Visual helpers (build-script local) ─────────────────────────────────────

// The metric staircase: units largest (top) to smallest (bottom), with the
// multiply/divide factor between each step. Unit hero visual; recurs every
// session as the catch-up anchor.
function metricStaircase(slide, lg, opts) {
  const o = opts || {};
  const x = lg.rightX;
  const w = lg.rightW;
  const y0 = lg.panelTopPadded;
  const bottom = o.bottom != null ? o.bottom : lg.safeBottom;
  const color = o.color || C.PRIMARY;
  const steps = o.steps;
  const cardH = bottom - y0;
  addCard(slide, x, y0, w, cardH, { strip: color });

  slide.addText(o.title || "Metric staircase", {
    x: x + 0.18, y: y0 + 0.10, w: w - 0.36, h: 0.28,
    fontSize: 14.5, fontFace: FONT_H, color, bold: true, align: "center", margin: 0,
  });

  const innerTop = y0 + 0.44;
  const keyH = 0.66;
  const region = bottom - innerTop - keyH - 0.06;
  const n = steps.length;
  const pillH = 0.38;
  const gap = (region - n * pillH) / (n - 1);
  const pillW = w - 1.1;
  const pillX = x + (w - pillW) / 2;

  let cy = innerTop;
  steps.forEach((st, i) => {
    if (i > 0) {
      slide.addText([
        { text: "× " + st.f, options: { color: C.SUCCESS, bold: true } },
        { text: "      ÷ " + st.f, options: { color: C.ALERT, bold: true } },
      ], {
        x: pillX, y: cy - gap, w: pillW, h: gap,
        fontSize: 13, fontFace: FONT_B, align: "center", valign: "middle", margin: 0,
      });
    }
    addTextOnShape(slide, st.u, {
      x: pillX, y: cy, w: pillW, h: pillH, rectRadius: 0.08,
      fill: { color: i === 0 ? color : C.SECONDARY },
    }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    cy += pillH + gap;
  });

  const keyY = bottom - keyH + 0.02;
  addCard(slide, x + 0.14, keyY, w - 0.28, keyH - 0.08, { fill: C.BG_LIGHT });
  slide.addText([
    { text: "Down to a smaller unit:  × multiply", options: { color: C.SUCCESS, bold: true, breakLine: true } },
    { text: "Up to a bigger unit:  ÷ divide", options: { color: C.ALERT, bold: true } },
  ], {
    x: x + 0.24, y: keyY + 0.04, w: w - 0.48, h: keyH - 0.16,
    fontSize: 12, fontFace: FONT_B, valign: "middle", margin: 0,
  });
}

// A worked conversion box: start value, an operation pill (x or div), and the
// result (or "?" on a question slide). Adapts to a constrained bottom so it
// can clear a reveal answer bar.
function convertPanel(slide, lg, opts) {
  const o = opts || {};
  const x = lg.rightX;
  const w = lg.rightW;
  const y0 = lg.panelTopPadded;
  const bottom = o.bottom != null ? o.bottom : lg.safeBottom;
  const color = o.color || C.PRIMARY;
  const cardH = bottom - y0;
  addCard(slide, x, y0, w, cardH, { strip: color });

  slide.addText(o.title || "Convert", {
    x: x + 0.18, y: y0 + 0.10, w: w - 0.36, h: 0.28,
    fontSize: 14.5, fontFace: FONT_H, color, bold: true, align: "center", margin: 0,
  });

  const innerX = x + 0.45;
  const innerW = w - 0.9;
  const startY = y0 + 0.50;
  const captionReserve = o.caption ? 0.56 : 0.12;
  const region = bottom - startY - captionReserve;
  const boxH = region * 0.30;
  const opH = region * 0.22;
  const gap = (region - 2 * boxH - opH) / 2;

  addTextOnShape(slide, o.start, {
    x: innerX, y: startY, w: innerW, h: boxH, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 23, fontFace: FONT_H, color: C.WHITE, bold: true });

  const opColor = o.divide ? C.ALERT : C.SUCCESS;
  const opText = (o.divide ? "÷ " : "× ") + o.factor;
  addTextOnShape(slide, opText, {
    x: innerX + innerW * 0.18, y: startY + boxH + gap, w: innerW * 0.64, h: opH, rectRadius: 0.08,
    fill: { color: opColor },
  }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });

  addTextOnShape(slide, o.end, {
    x: innerX, y: startY + boxH + gap + opH + gap, w: innerW, h: boxH, rectRadius: 0.08,
    fill: { color: o.endFill || C.SECONDARY },
  }, { fontSize: 23, fontFace: FONT_H, color: C.WHITE, bold: true });

  if (o.caption) {
    const capY = startY + region + 0.10;
    slide.addText(o.caption, {
      x: x + 0.22, y: capY, w: w - 0.44, h: bottom - capY - 0.08,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "top", margin: 0,
    });
  }
}

// A neat checkerboard of squares — a clear example of a tessellation (repeating
// shape, no gaps or overlaps) for the Daily Review visual anchor.
function tessGrid(slide, lg) {
  const x = lg.rightX;
  const w = lg.rightW;
  const y0 = lg.panelTopPadded;
  const bottom = lg.safeBottom;
  const cardH = bottom - y0;
  addCard(slide, x, y0, w, cardH, { strip: C.ACCENT });
  slide.addText("Shapes that tessellate", {
    x: x + 0.18, y: y0 + 0.10, w: w - 0.36, h: 0.30,
    fontSize: 14.5, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });
  const cols = 6;
  const rows = 4;
  const gridTop = y0 + 0.54;
  const avail = bottom - gridTop - 0.18;
  const cell = Math.min((w - 0.5) / cols, avail / rows);
  const gridW = cell * cols;
  const gridH = cell * rows;
  const gx = x + (w - gridW) / 2;
  const gy = gridTop + (avail - gridH) / 2;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const fill = ((r + c) % 2 === 0) ? C.PRIMARY : C.SECONDARY;
      slide.addShape("rect", {
        x: gx + c * cell, y: gy + r * cell, w: cell, h: cell,
        fill: { color: fill }, line: { color: C.WHITE, width: 1 },
      });
    }
  }
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, UNIT_TITLE, "Session 5: Choose, convert & compare",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3-4. Daily Review + reveal (Tessellations)
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Tessellations",
      [
        "Name one shape that tessellates on its own.",
        "True or false: a regular pentagon tessellates on its own.",
        "What makes shapes tessellate? (one idea)",
      ],
      NOTES_DR_Q, FOOTER,
      (slide, lg) => tessGrid(slide, lg)),
    (slide) => {
      addRevealAnswerBar(slide,
        ["Square / triangle / hexagon", "False", "No gaps or overlaps"],
        { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // 5-6. Fluency + reveal (division bracket)
  withReveal(
    () => fluencySlide(pres, "Fluency: Division bracket",
      ["176 ÷ 8", "261 ÷ 9", "210 ÷ 6"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["22", "29", "35"], { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // 7. Launch
  contentSlide(pres, "Launch", C.SECONDARY,
    "Which is longer?",
    [
      "1500 m   or   2 km?",
      "You can't compare different units yet.",
      "How could you make them easy to compare?",
      "Tell your partner.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Make the units match",
        start: "2 km", factor: "1000", divide: false, end: "2000 m",
        caption: "Now compare 2000 m with 1500 m — which is longer?",
      });
    });

  // 8. LI / SC
  liSlide(pres,
    "We are learning to convert measurements so we can compare them and solve problems.",
    [
      "I can choose a sensible unit and convert between units.",
      "I can convert two measurements to the same unit to compare them.",
      "I can solve a multi-step measurement problem and explain my steps.",
    ],
    NOTES_LI_SC, FOOTER);

  // 9. Recap — the routine (catch-up anchor)
  contentSlide(pres, "Recap", C.PRIMARY,
    "Our conversion routine",
    [
      "1. Set the stairs (largest to smallest).",
      "2. Count the steps between the units.",
      "3. Smaller unit → multiply.  Bigger unit → divide.",
      "",
      "Mass & capacity: every step × 1000.",
    ],
    NOTES_RECAP, FOOTER,
    (slide, lg) => {
      metricStaircase(slide, lg, {
        title: "Length staircase",
        color: C.PRIMARY,
        steps: [{ u: "km" }, { u: "m", f: 1000 }, { u: "cm", f: 100 }, { u: "mm", f: 10 }],
      });
    });

  // 10. I Do 1 — compare by converting
  workedExSlide(pres, 2, "I Do", "Which is longer: 1500 m or 2 km?",
    [
      "Different units — convert first.",
      "Change 2 km into metres: 2 × 1000 = 2000.",
      "",
      "Now compare in the same unit:",
      "2000 m  vs  1500 m.",
      "So 2 km is longer.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Make units match",
        start: "2 km", factor: "1000", divide: false, end: "2000 m",
        caption: "2000 m > 1500 m, so 2 km is longer.",
      });
    });

  // 11. I Do 2 — multi-step problem
  workedExSlide(pres, 2, "I Do", "2 L bottle, drink 600 mL — how much left?",
    [
      "Convert first: 2 L = 2 × 1000 = 2000 mL.",
      "",
      "Now subtract in the same unit:",
      "2000 − 600 = 1400.",
      "So 1400 mL are left.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Convert, then subtract",
        start: "2 L", factor: "1000", divide: false, end: "2000 mL",
        caption: "2000 − 600 = 1400 mL left.",
      });
    });

  // 12. I Do 3 — compare mass with a decimal
  workedExSlide(pres, 2, "I Do", "Which is heavier: 0.5 kg or 800 g?",
    [
      "Different units — convert first.",
      "Change 0.5 kg into grams: 0.5 × 1000 = 500.",
      "",
      "Now compare in the same unit:",
      "500 g  vs  800 g.",
      "So 800 g is heavier.",
    ],
    NOTES_IDO3, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Make units match",
        start: "0.5 kg", factor: "1000", divide: false, end: "500 g",
        caption: "500 g < 800 g, so 800 g is heavier.",
      });
    });

  // 13-14. CFU 1 + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Compare them", "Show Me Boards",
      "Which is longer: 3 m or 250 cm?\n\nConvert to the same unit first.",
      NOTES_CFU1_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "3 m is longer.  3 m = 300 cm, and 300 > 250.",
        { color: C.SUCCESS, label: "Answer", fontSize: 20 });
    }
  );

  // 15-16. We Do + reveal
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "1.5 L jug, pour 400 mL — how much left?",
      [
        "With your partner.",
        "",
        "Step 1: convert 1.5 L into mL.",
        "Step 2: subtract 400.",
        "Step 3: write the amount left in mL.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        convertPanel(slide, lg, {
          title: "Convert, then subtract",
          start: "1.5 L", factor: "1000", divide: false, end: "1500 mL",
          bottom: 4.15,
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "1100 mL left.   (1.5 L = 1500 mL, then 1500 − 400 = 1100)",
        { color: C.SUCCESS, label: "Answer", fontSize: 18 });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // 17-18. CFU hinge + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "True or false?", "Thumbs Up or Thumbs Down",
      "Anya says 1 kg is lighter than 900 g, because 1 is smaller than 900.\n\nTrue or false?",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        "FALSE. 1 kg = 1000 g, and 1000 > 900. Convert to the same unit before comparing.",
        { color: C.ALERT, label: "Check", fontSize: 17 });
    }
  );

  // 19. You Do — practice sheet
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 1 — choose & convert.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 2 — compare by converting.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 3 — multi-step problems.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.40;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("The golden rule", {
      x: 0.7, y: panelY + 0.14, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "1.  Make the units MATCH before you compare or calculate.", {
      x: 1.0, y: panelY + 0.54, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "2.  Smaller unit → multiply.   Bigger unit → divide.", {
      x: 1.0, y: panelY + 1.06, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "3.  Finished? Try the Year 8 Extension Challenge.", {
      x: 1.0, y: panelY + 1.58, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // 20. Exit Ticket
  exitTicketSlide(pres,
    [
      "Which is longer: 2 km or 1800 m? Convert to the same unit and explain how you know.",
      "A bottle holds 1.5 L. You drink 750 mL. How many millilitres are left?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 21. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the one rule you will never forget about converting units?",
      scItems: [
        "I can choose a sensible unit and convert between units.",
        "I can convert two measurements to the same unit to compare them.",
        "I can solve a multi-step measurement problem and explain my steps.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MMC_Session5_Choose_Convert_Compare.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  await generatePdfs();
  console.log("Session 5 build complete.");
}

// ─── PDFs ─────────────────────────────────────────────────────────────────

async function generatePdfs() {
  // ── Worksheet ──────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Choose a unit, convert, and compare. Make the units match before you compare or calculate.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addTipBox(doc,
      "Staircases:  LENGTH km, m, cm, mm (x1000, x100, x10).  MASS tonne, kg, g, mg (all x1000).  CAPACITY ML, kL, L, mL (all x1000).  DOWN to a smaller unit multiply, UP to a bigger unit divide.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Choose and convert", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   Best unit for the length of a pool:  _______    Then 25 m = _______ cm", y);
    y = addWriteLine(doc, "b)   3 kg = _______ g", y);
    y = addWriteLine(doc, "c)   4000 mL = _______ L", y);

    y = addSectionHeading(doc, "Section 2 — Compare (convert to the same unit first)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   Which is longer:  3 m  or  250 cm?     Longer: _______   (worked: 3 m = 300 cm)", y);
    y = addWriteLine(doc, "b)   Which is heavier:  0.5 kg  or  800 g?  Heavier: _______", y);
    y = addWriteLine(doc, "c)   Which is more:  1.2 L  or  1150 mL?    More: _______", y);

    y = addSectionHeading(doc, "Section 3 — Multi-step problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  A ribbon is 3 m long. You cut off 80 cm. How many centimetres are left?", y);
    y = addWriteLine(doc, "      Working: ______________________   Answer: _______ cm", y);
    y = addBodyText(doc, "b)  A jug holds 2 L. You pour out 750 mL. How many millilitres are left?", y);
    y = addWriteLine(doc, "      Working: ______________________   Answer: _______ mL", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Order from smallest to largest, converting to one unit first:  1.4 m,  120 cm,  1350 mm.", y);
    y = addWriteLine(doc, "      ______________ , ______________ , ______________", y);

    addPdfFooter(doc, `Session ${SESSION} | Choose, Convert, Compare | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // ── Answer Key ─────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the choose-convert-compare practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Choose and convert", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  metres for the pool; 25 m = 2500 cm.   b)  3000 g.   c)  4 L.", y);

    y = addSectionHeading(doc, "Section 2 — Compare", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3 m (3 m = 300 cm, more than 250 cm).   b)  800 g (0.5 kg = 500 g).   c)  1.2 L (1.2 L = 1200 mL, more than 1150 mL).", y);

    y = addSectionHeading(doc, "Section 3 — Multi-step", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3 m = 300 cm, 300 - 80 = 220 cm left.", y);
    y = addBodyText(doc, "b)  2 L = 2000 mL, 2000 - 750 = 1250 mL left.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Convert to cm: 1.4 m = 140 cm, 120 cm, 1350 mm = 135 cm. Smallest to largest: 120 cm, 1350 mm, 1.4 m.", y);

    y = addTipBox(doc,
      "Watch for: students who compare the raw numbers without converting. Always make the units match first. Check decimal moves on conversions.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // ── Year 8 Extension Challenge ─────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: EXTENSION_RES.name });
    let y = addPdfHeader(doc, EXTENSION_RES.name, {
      subtitle: "A stretch task beyond Year 6: area, volume and rate conversions.",
      color: C.SECONDARY,
      lessonInfo: `Session ${SESSION} | Aimed at Year 8 | Metric Measurement`,
    });

    y = addTipBox(doc,
      "Squared and cubed units scale differently. Because 1 m = 100 cm, an AREA of 1 m2 = 100 x 100 = 10 000 cm2. For VOLUME, 1 cm3 = 1 mL, and 1 L = 1000 cm3, so 1 m3 = 1000 L.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Area conversions", y, { color: C.SECONDARY });
    y = addWriteLine(doc, "a)   1 m2     =   _______ cm2      (worked: 100 x 100 = 10 000)", y);
    y = addWriteLine(doc, "b)   3 m2     =   _______ cm2", y);
    y = addWriteLine(doc, "c)   50 000 cm2 =   _______ m2", y);

    y = addSectionHeading(doc, "Section 2 — Volume and capacity links", y, { color: C.SECONDARY });
    y = addWriteLine(doc, "a)   1 cm3    =   _______ mL", y);
    y = addWriteLine(doc, "b)   2 L      =   _______ cm3", y);
    y = addWriteLine(doc, "c)   1 m3     =   _______ L", y);

    y = addSectionHeading(doc, "Section 3 — Rates and multi-step", y, { color: C.SECONDARY });
    y = addBodyText(doc, "a)  A car travels 90 km in 1 hour. How many metres does it travel each minute?", y);
    y = addWriteLine(doc, "      Working: _________________________   Answer: _______ m/min", y);
    y = addBodyText(doc, "b)  A tap fills 250 mL every 10 seconds. How many litres does it fill in 2 minutes?", y);
    y = addWriteLine(doc, "      Working: _________________________   Answer: _______ L", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.PRIMARY });
    y = addBodyText(doc, "A fish tank is 60 cm long, 30 cm wide and 40 cm deep. Find its volume in cm3, then its capacity in litres.", y);
    y = addWriteLine(doc, "      Volume = _______ cm3        Capacity = _______ L", y);

    y = addTipBox(doc,
      "Teacher answers. S1: a) 10 000 cm2  b) 30 000 cm2  c) 5 m2.  S2: a) 1 mL  b) 2000 cm3  c) 1000 L.  S3: a) 90 km = 90 000 m, div 60 = 1500 m/min  b) 120 s div 10 = 12 lots of 250 mL = 3000 mL = 3 L.  Challenge: 60 x 30 x 40 = 72 000 cm3 = 72 000 mL = 72 L.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Session ${SESSION} | Year 8 Extension Challenge | Metric Measurement`);
    await writePdf(doc, path.join(OUT_DIR, EXTENSION_RES.fileName));
    console.log("PDF written: " + EXTENSION_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
