"use strict";

// Metric Measurement & Unit Conversion (Year 6 Numeracy)
// Session 2 of 5: Converting length (mm, cm, m, km).
// VC2M6M01 — using the correct operation when converting between units;
//            recognising equivalence such as 1.25 m = 125 cm.
// Daily Review: Decimal Operations and Place Value Mastery (prior learning).
// Fluency: Division bracket (long-division layout). NOT the bus stop method.
//
// Unit routine (consistent across all 5 sessions):
//   SET THE STAIRS  ->  COUNT THE STEPS  ->  MULTIPLY OR DIVIDE.
// Self-contained for catch-up: this session re-teaches the staircase method
// from scratch with fresh worked examples. A student who missed Session 1 can
// still access today's lesson.

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

const SESSION = 2;
const TOTAL = 5;
const UNIT_TITLE = "Metric Measurement & Conversion";
const FOOTER = `Metric Measurement | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/MMC_Session2_Converting_Length";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Converting Length Practice",
  "Convert between mm, cm, m and km using multiply or divide.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the converting-length practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back, mathematicians. Last session we chose the right unit. Today we change between units.
- Our job is to convert lengths - to swap metres for centimetres, or centimetres for metres - without changing the actual length.
- By the end you will know exactly when to multiply and when to divide.

DO:
- Have whiteboards and markers ready.
- Keep the printed practice sheet face down until the You Do section.

TEACHER NOTES:
Session 2 of 5. Today introduces the conversion routine: Set the stairs, Count the steps, Multiply or divide. This is the core method for the whole unit and is re-taught from scratch, so students who missed Session 1 can still access it.

WATCH FOR:
- Students who think converting changes the length. It does not - the same length is just written in a different unit.

[General: Title | VTLM 2.0: Attention, focus and regulation]`;

const NOTES_RESOURCES = `SAY:
- Here is what we are using today.
- Whiteboards for most of the lesson, and the practice sheet near the end.

DO:
- Print one practice sheet and one answer key per student who needs paper.
- Have a metre ruler handy to show 100 centimetres in a metre.

TEACHER NOTES:
One student resource today plus an answer key. The metre ruler is a strong anchor for 1 m = 100 cm.

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_DR_Q = `SAY:
- Quick warm up on decimals and place value.
- Read each one, then write your answer on your whiteboard.
- These are from earlier learning.

DO:
- Display the three prompts.
- Allow about 90 seconds.
- Watch the multiply-by-ten prompt closely.

TEACHER NOTES:
Daily Review revisits Decimal Operations and Place Value Mastery. Multiplying by 10 and rounding decimals are exactly the moves we use when converting today.

WATCH FOR:
- Students who move every digit one place when multiplying by 10 - secure.
- Students who just add a zero to a decimal - reteach place value movement.

[Stage 1: Daily Review | VTLM 2.0: Retention and recall]`;

const NOTES_DR_A = `SAY:
- Let us check together.
- 9.2 minus 4.65 is 4.55. Line up the points.
- 3.4 times 10 is 34. Every digit moves one place to the left.
- 7.38 rounded to one decimal place is 7.4.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Multiplying by 10 by moving digits, not adding a zero, is the bridge to converting today.

WATCH FOR:
- Confident self-correction - secure.
- 3.4 times 10 written as 3.40 - revisit place value.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency now. Division using the division bracket.
- Set each one out with the division bracket and work from the left.
- Whisper your answer, then write it.

DO:
- Display the three divisions.
- Allow about 80 seconds.

TEACHER NOTES:
Fluency builds division using the division bracket, the same focus across all five sessions. Please call it the division bracket, not anything else. Brisk retrieval, not new teaching.

WATCH FOR:
- Clean carrying of remainders to the next digit - secure.
- Dropped remainders - reteach the carry step with the division bracket.

[Stage 1: Fluency | VTLM 2.0: Retention and recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check your division brackets.
- 156 divided by 4 is 39.
- 203 divided by 7 is 29.
- 288 divided by 8 is 36.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
203 divided by 7 has a zero in the tens of the working: 20 divided by 7 is 2 remainder 6, carry the 6 to make 63, 63 divided by 7 is 9.

WATCH FOR:
- Clean handling of the carry - secure.
- Answers near 30 for 203 divided by 7 - the carry was mishandled.

[Stage 1: Fluency Answers | VTLM 2.0: Retention and recall]`;

const NOTES_LAUNCH = `SAY:
- Picture a metre ruler in your hands. How many centimetres are along it? One hundred.
- So one metre is the same length as one hundred centimetres. We did not change the length, only the unit.
- Now think: how many centimetres in two metres? In half a metre? Tell your partner how you worked it out.

DO:
- Hold up a metre ruler if you have one.
- Give 40 seconds of partner talk.
- Take responses. Listen for times one hundred.

TEACHER NOTES:
The launch connects what students know - 100 cm in a metre - to today's rule. The big idea: converting keeps the length the same and just changes the unit. Two metres is 200 cm; half a metre is 50 cm.

WATCH FOR:
- Students who multiply by 100 to get centimetres - exactly today's move.
- Students who think 2 m is 102 cm - they added instead of scaling.

[Stage 1: Launch | VTLM 2.0: Knowledge and memory]`;

const NOTES_LI_SC = `SAY:
- Here is what we are learning today.
- We are learning to convert between metric units of length.
- Read the success criteria with me.

DO:
- Choral read the learning intention.
- Read each I can statement.

TEACHER NOTES:
SC1 is reachable for everyone - knowing which units are bigger and smaller. SC2 is the core target the exit ticket assesses - converting with multiply or divide. SC3 stretches into decimal conversions. Tier labels stay off the slide.

WATCH FOR:
- Students repeating the language back - tracking with us.

[General: LI/SC | VTLM 2.0: Planning]`;

const NOTES_VOCAB = `SAY:
- Two key words today.
- Convert means to change a measurement from one unit to another without changing the actual amount.
- Equivalent means equal in value. 1.25 metres and 125 centimetres are equivalent - the same length, written two ways.

DO:
- Point to the staircase as you say convert.
- Have students say convert and equivalent back to you.

TEACHER NOTES:
These two words carry the whole unit. Keep returning to the idea that the length does not change, only the unit and the number.

WATCH FOR:
- Students who explain that the length stays the same - secure.
- Students who think a bigger number means a bigger length - 125 cm is not longer than 1.25 m; revisit equivalence.

[General: Key Vocabulary | VTLM 2.0: Knowledge and memory]`;

const NOTES_IDO1 = `SAY:
- Let us convert 3 metres into centimetres together. Watch my routine.
- First I set the stairs - largest at the top: km, m, cm, mm.
- Next I count the steps from metres to centimetres. It is one step, and that step is times one hundred.
- I am going down to a smaller unit, so I multiply. 3 times 100 is 300.
- So 3 metres is 300 centimetres. Same length, new unit.

DO:
- Point to m and cm on the staircase as you count the step.
- Write 3 times 100 equals 300 on the board.

TEACHER NOTES:
This is the I Do for the multiply direction. Anchor the rule: down to a smaller unit means more of them, so multiply. The staircase is the map we set up first every time.

MISCONCEPTIONS:
- Misconception: to get centimetres you divide because centimetres are small.
  Why: students link small unit with the smaller operation.
  Impact: 3 m becomes 0.03 cm, which is clearly wrong.
  Quick correction: a smaller unit means you need MORE of them, so the number grows - multiply.

WATCH FOR:
- Students who multiply going down - secure.
- Students who divide - point to the staircase and ask: more or fewer centimetres than metres?

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now the other direction. Convert 450 centimetres into metres.
- Set the stairs. Count the steps from cm up to m - one step, times one hundred.
- This time I am going up to a bigger unit, so I divide. 450 divided by 100.
- Dividing by 100 moves every digit two places. 450 divided by 100 is 4.5.
- So 450 centimetres is 4.5 metres. Notice the decimal - that is why our place value warm up mattered.

DO:
- Point to cm and m on the panel.
- Show 450 divided by 100 equals 4.5 and say each digit moves two places.

TEACHER NOTES:
This is the divide direction. Up to a bigger unit means fewer of them, so divide. Connect dividing by 100 to moving digits two places, linking back to Daily Review.

MISCONCEPTIONS:
- Misconception: you just delete the zeros, so 450 becomes 45.
  Why: it works for some whole-number cases by accident.
  Impact: students get 45 m instead of 4.5 m.
  Quick correction: dividing by 100 moves the digits two places - 450 becomes 4.5, not 45.

WATCH FOR:
- Students who get 4.5 m - secure.
- Students who get 45 m or 0.45 m - re-count the place value moves.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO3 = `SAY:
- One more, and this is an important one. Convert 1.25 metres into centimetres.
- Set the stairs. Metres down to centimetres is one step, times one hundred.
- Down to a smaller unit, so I multiply. 1.25 times 100 moves every digit two places to the left.
- 1.25 times 100 is 125. So 1.25 metres equals 125 centimetres.
- These two measurements are equivalent - exactly the same length, written two ways.

DO:
- Write 1.25 times 100 equals 125.
- Say the word equivalent and point to both forms.

TEACHER NOTES:
This is the curriculum's equivalence example: 1.25 m equals 125 cm. It pulls together multiplying by 100 and the idea of equivalence. A strong model to return to in the You Do.

MISCONCEPTIONS:
- Misconception: 1.25 m must be more than 125 cm because 1.25 looks bigger... or smaller.
  Why: students compare the numbers, not the lengths.
  Impact: they doubt a correct conversion.
  Quick correction: they are the same length. The number changed because the unit changed.

WATCH FOR:
- Students who say equivalent and explain why - secure and ready for depth.
- Students unsure about the decimal - re-show the two-place move.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU1_Q = `SAY:
- Quick check on your whiteboards.
- Convert 6 metres into centimetres.
- Show me on three, two, one.

DO:
- Display the prompt.
- Allow about 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me 6 metres in centimetres on three, two, one, show."
- Scan for: 600 cm.
PROCEED: If 80 percent show 600, click to reveal and move on.
PIVOT: Most likely misconception - students divide because centimetres feel small.
- Reteach: "Down to a smaller unit means more of them. 6 metres is 600 centimetres. Multiply by 100."
- Re-check: "Now convert 9 metres into centimetres."

TEACHER NOTES:
Checks the multiply direction before guided practice.

WATCH FOR:
- 600 cm - secure.
- 0.06 cm or 60 cm - re-count the step and the direction.

[Stage 2: CFU | VTLM 2.0: Supported Application]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Convert 320 centimetres into metres.
- Step one, set the stairs. Step two, count the steps from cm to m. Step three, decide multiply or divide, then calculate.
- Work it with your partner on your whiteboard.

DO:
- Keep the steps visible.
- Walk and scan. Prompt with "bigger or smaller unit" where needed.
- Take 60 seconds before revealing.

TEACHER NOTES:
Guided practice in the divide direction. Up to a bigger unit means divide by 100. Watch for the decimal answer.

WATCH FOR:
- Pairs who divide by 100 to get 3.2 m - secure.
- Pairs who multiply - point them up the staircase: fewer metres than centimetres.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- Centimetres up to metres is one step, times one hundred, and we are going up, so divide.
- 320 divided by 100 is 3.2.
- So 320 centimetres is 3.2 metres.

DO:
- Click to reveal.
- Ask one pair to explain the direction.

TEACHER NOTES:
The reveal restates the direction and the calculation. 3.2 m shows the decimal place value at work.

WATCH FOR:
- Self-correction to 3.2 m - secure.
- Answers of 32 m or 0.32 m - re-count the place value move.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. A student wrote: 200 centimetres equals 20000 metres.
- Do not just fix it. Decide: true or false?
- Thumbs up for true, thumbs down for false.

DO:
- Display the claim.
- Give 15 seconds of thinking time.

CFU CHECKPOINT:
Technique: Thumbs Up or Down
Script:
- Say: "Thumbs up if true, thumbs down if false."
- Scan for: thumbs down. 200 cm is only 2 m.
PROCEED: If 80 percent thumbs down with a reason, click to reveal.
PIVOT: Most likely misconception - students multiply when going up to a bigger unit.
- Reteach: "Centimetres up to metres means FEWER of them, so divide. 200 divided by 100 is 2 metres."
- Re-check: "What is 200 centimetres in metres? Show me."

TEACHER NOTES:
This checks the threshold idea: direction decides the operation. Going to a bigger unit divides. 20000 m would be a 20 kilometre length - clearly not 2 metres of something.

WATCH FOR:
- Confident thumbs down naming divide - secure.
- Thumbs up - revisit bigger unit means fewer, so divide.

[Stage 2: CFU Hinge | VTLM 2.0: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Time to work on your own. Take the practice sheet.
- Section one, metres to centimetres. Section two, centimetres to metres. Section three, a mix with kilometres and millimetres.
- Every time: set the stairs, count the steps, then multiply or divide.

DO:
- Distribute the practice sheet.
- Circulate. Sit with the enabling group first.
- Watch for the direction decision: bigger or smaller unit.

TEACHER NOTES:
The sheet sequences the multiply direction, then the divide direction, then a mix. Section two is the threshold - dividing when going to a bigger unit.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the staircase box and the worked first item at the top of the sheet. Stay with metres and centimetres only.
- Extra Notes: Do the first item in each section with this group before they continue.
EXTENDING PROMPT:
- Task: The challenge box - a two-step word problem and an ordering task that needs converting to a common unit first.
- Extra Notes: Push for converting everything to the same unit before comparing.

WATCH FOR:
- Students who decide the direction before calculating - secure.
- Students who always multiply - return to the bigger or smaller unit question.

[Stage 4: You Do | VTLM 2.0: Mastery and application]`;

const NOTES_EXIT = `SAY:
- Last task, on your own, on your whiteboard.
- One, convert 7 metres into centimetres.
- Two, convert 250 centimetres into metres.
- Set the stairs and decide multiply or divide before you calculate.

DO:
- Display the prompt.
- Allow about 3 minutes.
- Collect whiteboards or take a photo for your records.

TEACHER NOTES:
Exit ticket assesses SC2 (converting with multiply or divide). Internal target is SC2. Do not display any SC label to students.

WATCH FOR:
- 700 cm and 2.5 m - secure.
- A correct first answer but wrong direction on the second - revisit the bigger-unit-divide rule in Session 3.

[Stage 5: Exit Ticket | VTLM 2.0: Mastery and application]`;

const NOTES_CLOSING = `SAY:
- Look back at our success criteria.
- Show me thumbs for each one - got it, getting there, or need more practice.
- Turn and tell your partner: how do you decide whether to multiply or divide?

DO:
- Read each I can statement.
- Use thumbs to self-assess.
- Note who is still unsure for tomorrow.

TEACHER NOTES:
The big idea is direction decides the operation: down to a smaller unit multiply, up to a bigger unit divide. Students who can say that are ready for mass next session, which uses the same routine.

WATCH FOR:
- Students who explain the direction rule - secure.
- Students who only remember one direction - revisit in Session 3's launch.

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

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, UNIT_TITLE, "Session 2: Converting length",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3-4. Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Decimals",
      [
        "Calculate 9.2 − 4.65",
        "What is 3.4 × 10?",
        "Round 7.38 to one decimal place",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["4.55", "34", "7.4"], { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // 5-6. Fluency + reveal (division bracket)
  withReveal(
    () => fluencySlide(pres, "Fluency: Division bracket",
      ["156 ÷ 4", "203 ÷ 7", "288 ÷ 8"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["39", "29", "36"], { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // 7. Launch
  contentSlide(pres, "Launch", C.SECONDARY,
    "One metre — how many cm?",
    [
      "Picture a metre ruler.",
      "How many centimetres along it?",
      "What about 2 metres? Half a metre?",
      "Tell your partner how you worked it out.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Same length, new unit",
        start: "1 m", factor: "100", divide: false, end: "100 cm",
        caption: "Converting keeps the length the same. Only the unit changes.",
      });
    });

  // 8. LI / SC
  liSlide(pres,
    "We are learning to convert between metric units of length.",
    [
      "I can say which units of length are bigger and smaller.",
      "I can convert between units of length using multiply or divide.",
      "I can convert lengths that include decimals and explain my steps.",
    ],
    NOTES_LI_SC, FOOTER);

  // 9. Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.PRIMARY,
    "Convert and equivalent",
    [
      "Convert = change to another unit, same length.",
      "Equivalent = equal in value.",
      "",
      "1.25 m and 125 cm are equivalent — the same length, two ways.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      metricStaircase(slide, lg, {
        title: "Length staircase",
        color: C.PRIMARY,
        steps: [{ u: "km" }, { u: "m", f: 1000 }, { u: "cm", f: 100 }, { u: "mm", f: 10 }],
      });
    });

  // 10. I Do 1 — m -> cm (multiply)
  workedExSlide(pres, 2, "I Do", "Convert 3 m into cm",
    [
      "Set the stairs: km, m, cm, mm.",
      "Count the steps: m to cm is × 100.",
      "Going DOWN to a smaller unit → multiply.",
      "",
      "3 × 100 = 300.",
      "So 3 m = 300 cm.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      metricStaircase(slide, lg, {
        title: "Length staircase",
        color: C.PRIMARY,
        steps: [{ u: "km" }, { u: "m", f: 1000 }, { u: "cm", f: 100 }, { u: "mm", f: 10 }],
      });
    });

  // 11. I Do 2 — cm -> m (divide)
  workedExSlide(pres, 2, "I Do", "Convert 450 cm into m",
    [
      "Count the steps: cm to m is × 100.",
      "Going UP to a bigger unit → divide.",
      "",
      "450 ÷ 100 = 4.5.",
      "Every digit moves two places.",
      "So 450 cm = 4.5 m.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Convert",
        start: "450 cm", factor: "100", divide: true, end: "4.5 m",
        caption: "Up to a bigger unit means fewer of them — so the number gets smaller.",
      });
    });

  // 12. I Do 3 — 1.25 m -> cm (equivalence)
  workedExSlide(pres, 2, "I Do", "Convert 1.25 m into cm",
    [
      "Count the steps: m to cm is × 100.",
      "Going DOWN to a smaller unit → multiply.",
      "",
      "1.25 × 100 = 125.",
      "So 1.25 m = 125 cm.",
      "These are equivalent — the same length.",
    ],
    NOTES_IDO3, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Equivalent lengths",
        start: "1.25 m", factor: "100", divide: false, end: "125 cm",
        caption: "1.25 m and 125 cm are the same length, written two ways.",
      });
    });

  // 13-14. CFU 1 + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Convert it", "Show Me Boards",
      "Convert 6 m into centimetres.",
      NOTES_CFU1_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "600 cm   (6 × 100 — down to a smaller unit, multiply)",
        { color: C.SUCCESS, label: "Answer", fontSize: 20 });
    }
  );

  // 15-16. We Do + reveal
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Convert 320 cm into m",
      [
        "With your partner.",
        "",
        "Step 1: set the stairs.",
        "Step 2: count the steps (cm to m is × 100).",
        "Step 3: bigger unit → divide. Calculate.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        convertPanel(slide, lg, {
          title: "Convert",
          start: "320 cm", factor: "100", divide: true, end: "?",
          bottom: 4.15,
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "3.2 m   (320 ÷ 100 — up to a bigger unit, divide)",
        { color: C.SUCCESS, label: "Answer", fontSize: 20 });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // 17-18. CFU hinge + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "True or false?", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n200 cm = 20 000 m\n\nIs this true or false?",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        "FALSE. 200 cm = 2 m (200 ÷ 100). Up to a bigger unit, you divide.",
        { color: C.ALERT, label: "Check", fontSize: 18 });
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
      { text: "Section 1 — metres to centimetres.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 2 — centimetres to metres.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 3 — a mix with km and mm.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.40;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Every time, the same routine", {
      x: 0.7, y: panelY + 0.14, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "1.  Set the stairs: km, m, cm, mm.", {
      x: 1.0, y: panelY + 0.54, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "2.  Count the steps (× 1000, × 100 or × 10).", {
      x: 1.0, y: panelY + 1.06, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "3.  Smaller unit → multiply.   Bigger unit → divide.", {
      x: 1.0, y: panelY + 1.58, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // 20. Exit Ticket
  exitTicketSlide(pres,
    [
      "Convert 7 m into centimetres.",
      "Convert 250 cm into metres.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 21. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how do you decide whether to multiply or divide?",
      scItems: [
        "I can say which units of length are bigger and smaller.",
        "I can convert between units of length using multiply or divide.",
        "I can convert lengths that include decimals and explain my steps.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MMC_Session2_Converting_Length.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  await generatePdfs();
  console.log("Session 2 build complete.");
}

// ─── PDFs ─────────────────────────────────────────────────────────────────

async function generatePdfs() {
  // ── Worksheet ──────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Convert between mm, cm, m and km. Set the stairs, count the steps, multiply or divide.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addTipBox(doc,
      "Staircase (largest to smallest): km, m, cm, mm.  Steps: km-m is x1000, m-cm is x100, cm-mm is x10.  DOWN to a smaller unit: multiply.  UP to a bigger unit: divide.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Metres to centimetres (x 100)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   4 m   =   _______ cm      (worked: 4 x 100 = 400)", y);
    y = addWriteLine(doc, "b)   7 m   =   _______ cm", y);
    y = addWriteLine(doc, "c)   2.5 m =   _______ cm", y);

    y = addSectionHeading(doc, "Section 2 — Centimetres to metres (div 100)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   600 cm =   _______ m", y);
    y = addWriteLine(doc, "b)   350 cm =   _______ m", y);
    y = addWriteLine(doc, "c)   80 cm   =   _______ m", y);

    y = addSectionHeading(doc, "Section 3 — A mix (watch the unit!)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   3 km   =   _______ m", y);
    y = addWriteLine(doc, "b)   5 cm   =   _______ mm", y);
    y = addWriteLine(doc, "c)   1.25 m =   _______ cm", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc, "1.  A ribbon is 2.4 m long. Jess cuts off 80 cm. How many centimetres of ribbon are left? (Convert to the same unit first.)", y);
    y = addWriteLine(doc, "      Working: ______________________   Answer: _______ cm", y);
    y = addBodyText(doc, "2.  Order from shortest to longest:  1.5 m,  120 cm,  1450 mm.", y);
    y = addWriteLine(doc, "      ______________ , ______________ , ______________", y);

    addPdfFooter(doc, `Session ${SESSION} | Converting Length | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // ── Answer Key ─────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the converting-length practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Metres to centimetres", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  400 cm (worked).   b)  700 cm.   c)  250 cm.", y);

    y = addSectionHeading(doc, "Section 2 — Centimetres to metres", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  6 m.   b)  3.5 m.   c)  0.8 m.", y);

    y = addSectionHeading(doc, "Section 3 — A mix", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3000 m (3 x 1000).   b)  50 mm (5 x 10).   c)  125 cm (1.25 x 100).", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc, "1.  2.4 m = 240 cm. 240 - 80 = 160 cm left.", y);
    y = addBodyText(doc, "2.  Convert to cm: 1.5 m = 150 cm, 120 cm, 1450 mm = 145 cm. Shortest to longest: 120 cm, 1450 mm, 1.5 m.", y);

    y = addTipBox(doc,
      "Watch for: students who always multiply. The direction decides the operation - down to a smaller unit multiply, up to a bigger unit divide. Check the decimal moves: 250 cm is 2.5 m, not 25 m.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
