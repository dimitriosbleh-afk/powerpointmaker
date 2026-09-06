"use strict";

// Metric Measurement & Unit Conversion (Year 6 Numeracy)
// Session 3 of 5: Converting mass (mg, g, kg, tonne).
// VC2M6M01 — using the correct operation when converting between units of mass;
//            recognising equivalence such as 1.25 kg = 1250 g.
// Daily Review: Practical Problem Solving and Financial Reasoning (prior learning).
// Fluency: Division bracket (long-division layout). NOT the bus stop method.
//
// Unit routine (consistent across all 5 sessions):
//   SET THE STAIRS  ->  COUNT THE STEPS  ->  MULTIPLY OR DIVIDE.
// Self-contained for catch-up: this session re-teaches the staircase method
// from scratch. A student who missed Session 1 or 2 can still access today.
// Mass is friendlier than length: every step is x1000.

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

const SESSION = 3;
const TOTAL = 5;
const UNIT_TITLE = "Metric Measurement & Conversion";
const FOOTER = `Metric Measurement | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/MMC_Session3_Converting_Mass";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Converting Mass Practice",
  "Convert between mg, g, kg and tonne using multiply or divide.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the converting-mass practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back, mathematicians. We can convert length. Today we convert mass - how heavy things are.
- The good news is mass uses the exact same routine, and every step is the same size.
- By the end you will convert between milligrams, grams, kilograms and tonnes with confidence.

DO:
- Have whiteboards and markers ready.
- Keep the printed practice sheet face down until the You Do section.

TEACHER NOTES:
Session 3 of 5. The conversion routine is re-taught from scratch, so students who missed earlier sessions can still access today. The one new idea: for mass, every step is x1000, which is simpler than length.

WATCH FOR:
- Students who think a heavier object needs a bigger number. The number depends on the unit, not just the object.

[General: Title | VTLM 2.0: Attention, focus and regulation]`;

const NOTES_RESOURCES = `SAY:
- Here is what we are using today.
- Whiteboards for most of the lesson, and the practice sheet near the end.

DO:
- Print one practice sheet and one answer key per student who needs paper.
- Have a 1 kg object handy, such as a bag of sugar or rice.

TEACHER NOTES:
One student resource today plus an answer key. A real 1 kg object anchors 1 kg = 1000 g.

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_DR_Q = `SAY:
- Quick warm up on money problems.
- Read each one, then write your answer on your whiteboard.
- These are from earlier learning.

DO:
- Display the three prompts.
- Allow about 90 seconds.
- Scan for the change-from-ten prompt.

TEACHER NOTES:
Daily Review revisits Practical Problem Solving and Financial Reasoning. Multi-step money thinking and estimation keep these skills warm.

WATCH FOR:
- Students who find change by counting up from the price - secure.
- Students who add instead of subtract for change - reteach the take-away.

[Stage 1: Daily Review | VTLM 2.0: Retention and recall]`;

const NOTES_DR_A = `SAY:
- Let us check together.
- Four drinks at 2 dollars 75 each is 11 dollars.
- Change from 10 dollars for a 6 dollar 40 item is 3 dollars 60.
- Estimating, 4 dollars 95 plus 3 dollars 10 plus 6 dollars 85 rounds to 5 plus 3 plus 7, about 15 dollars.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Estimation by rounding to the nearest dollar mirrors today's idea of sensible-sized numbers.

WATCH FOR:
- Confident self-correction - secure.
- Change found by adding up - fine if it lands on 3 dollars 60.

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
- 175 divided by 5 is 35.
- 216 divided by 6 is 36.
- 294 divided by 7 is 42.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
294 divided by 7: 29 divided by 7 is 4 remainder 1, carry the 1 to make 14, 14 divided by 7 is 2.

WATCH FOR:
- Clean carrying - secure.
- Answers near 40 for 294 divided by 7 - re-check the carry.

[Stage 1: Fluency Answers | VTLM 2.0: Retention and recall]`;

const NOTES_LAUNCH = `SAY:
- Picture a bag of sugar. It is 1 kilogram. How many grams is that? One thousand.
- So 1 kilogram is the same mass as 1000 grams. We did not change the mass, only the unit.
- Now think: how many grams in 2 kilograms? In half a kilogram? Tell your partner.

DO:
- Hold up a 1 kg object if you have one.
- Give 40 seconds of partner talk.
- Take responses. Listen for times one thousand.

TEACHER NOTES:
The launch connects what students know - 1000 g in a kilogram - to today's rule. 2 kg is 2000 g; half a kilogram is 500 g. Same routine as length, bigger step.

WATCH FOR:
- Students who multiply by 1000 to get grams - exactly today's move.
- Students who think 2 kg is 1002 g - they added instead of scaling.

[Stage 1: Launch | VTLM 2.0: Knowledge and memory]`;

const NOTES_LI_SC = `SAY:
- Here is what we are learning today.
- We are learning to convert between metric units of mass.
- Read the success criteria with me.

DO:
- Choral read the learning intention.
- Read each I can statement.

TEACHER NOTES:
SC1 is reachable for everyone - knowing which mass units are bigger and smaller. SC2 is the core target the exit ticket assesses - converting with multiply or divide. SC3 stretches into decimal conversions. Tier labels stay off the slide.

WATCH FOR:
- Students repeating the language back - tracking with us.

[General: LI/SC | VTLM 2.0: Planning]`;

const NOTES_VOCAB = `SAY:
- Mass means how heavy something is. We measure it in milligrams, grams, kilograms and tonnes.
- Notice the staircase. For mass, every single step is times one thousand - kilograms to grams, grams to milligrams, and tonnes to kilograms.
- That makes mass a friendly set to convert: the step is always the same size.

DO:
- Point to the mass staircase as you name each step.
- Have students say the four units in order, largest to smallest.

TEACHER NOTES:
The key idea today: every mass step is x1000. Contrast gently with length, where the steps were x1000, x100 and x10. Same routine, simpler steps.

WATCH FOR:
- Students who notice the steps are all x1000 - secure.
- Students who carry over x100 from length - point to the mass staircase factors.

[General: Key Vocabulary | VTLM 2.0: Knowledge and memory]`;

const NOTES_IDO1 = `SAY:
- Let us convert 3 kilograms into grams together. Watch my routine.
- First I set the stairs: tonne, kg, g, mg.
- Next I count the steps from kilograms to grams. One step, and that step is times one thousand.
- I am going down to a smaller unit, so I multiply. 3 times 1000 is 3000.
- So 3 kilograms is 3000 grams. Same mass, new unit.

DO:
- Point to kg and g on the staircase.
- Write 3 times 1000 equals 3000 on the board.

TEACHER NOTES:
This is the I Do for the multiply direction. Down to a smaller unit means more of them, so multiply. The staircase is the map we set up first.

MISCONCEPTIONS:
- Misconception: to get grams you divide because grams are small.
  Why: students link a small unit with the smaller operation.
  Impact: 3 kg becomes 0.003 g, which is clearly wrong.
  Quick correction: a smaller unit means you need MORE of them, so the number grows - multiply.

WATCH FOR:
- Students who multiply going down - secure.
- Students who divide - ask: more or fewer grams than kilograms?

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now the other direction. Convert 4500 grams into kilograms.
- Set the stairs. Count the steps from grams up to kilograms - one step, times one thousand.
- I am going up to a bigger unit, so I divide. 4500 divided by 1000.
- Dividing by 1000 moves every digit three places. 4500 divided by 1000 is 4.5.
- So 4500 grams is 4.5 kilograms.

DO:
- Point to g and kg on the panel.
- Show 4500 divided by 1000 equals 4.5.

TEACHER NOTES:
This is the divide direction. Up to a bigger unit means fewer of them, so divide. Connect dividing by 1000 to moving digits three places.

MISCONCEPTIONS:
- Misconception: just delete three zeros, so 4500 becomes 4.
  Why: it works for round thousands by accident.
  Impact: students lose the 5, getting 4 kg instead of 4.5 kg.
  Quick correction: dividing by 1000 moves the digits three places - 4500 becomes 4.5.

WATCH FOR:
- Students who get 4.5 kg - secure.
- Students who get 4 kg or 45 kg - re-count the place value moves.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO3 = `SAY:
- One more, with a decimal. Convert 1.25 kilograms into grams.
- Set the stairs. Kilograms down to grams is one step, times one thousand.
- Down to a smaller unit, so I multiply. 1.25 times 1000 moves every digit three places.
- 1.25 times 1000 is 1250. So 1.25 kilograms equals 1250 grams.
- These are equivalent - exactly the same mass, written two ways.

DO:
- Write 1.25 times 1000 equals 1250.
- Say the word equivalent and point to both forms.

TEACHER NOTES:
This is the mass equivalence example: 1.25 kg equals 1250 g. It pulls together multiplying by 1000 and equivalence. A strong model to return to in the You Do.

MISCONCEPTIONS:
- Misconception: 1.25 kg must weigh less than 1250 g because 1.25 is a small number.
  Why: students compare the numbers, not the masses.
  Impact: they doubt a correct conversion.
  Quick correction: they are the same mass. The number changed because the unit changed.

WATCH FOR:
- Students who say equivalent and explain why - secure and ready for depth.
- Students unsure about the decimal - re-show the three-place move.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU1_Q = `SAY:
- Quick check on your whiteboards.
- Convert 5 kilograms into grams.
- Show me on three, two, one.

DO:
- Display the prompt.
- Allow about 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me 5 kilograms in grams on three, two, one, show."
- Scan for: 5000 g.
PROCEED: If 80 percent show 5000, click to reveal and move on.
PIVOT: Most likely misconception - students divide because grams feel small.
- Reteach: "Down to a smaller unit means more of them. 5 kilograms is 5000 grams. Multiply by 1000."
- Re-check: "Now convert 8 kilograms into grams."

TEACHER NOTES:
Checks the multiply direction before guided practice.

WATCH FOR:
- 5000 g - secure.
- 0.005 g or 500 g - re-count the step and the direction.

[Stage 2: CFU | VTLM 2.0: Supported Application]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Convert 2500 grams into kilograms.
- Step one, set the stairs. Step two, count the steps from g to kg. Step three, decide multiply or divide, then calculate.
- Work it with your partner on your whiteboard.

DO:
- Keep the steps visible.
- Walk and scan. Prompt with "bigger or smaller unit" where needed.
- Take 60 seconds before revealing.

TEACHER NOTES:
Guided practice in the divide direction. Up to a bigger unit means divide by 1000. Watch for the decimal answer.

WATCH FOR:
- Pairs who divide by 1000 to get 2.5 kg - secure.
- Pairs who multiply - point them up the staircase: fewer kilograms than grams.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- Grams up to kilograms is one step, times one thousand, and we are going up, so divide.
- 2500 divided by 1000 is 2.5.
- So 2500 grams is 2.5 kilograms.

DO:
- Click to reveal.
- Ask one pair to explain the direction.

TEACHER NOTES:
The reveal restates the direction and the calculation. 2.5 kg shows the decimal place value at work.

WATCH FOR:
- Self-correction to 2.5 kg - secure.
- Answers of 25 kg or 0.25 kg - re-count the place value move.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. A student wrote: 3 kilograms equals 0.003 grams.
- Do not just fix it. Decide: true or false?
- Thumbs up for true, thumbs down for false.

DO:
- Display the claim.
- Give 15 seconds of thinking time.

CFU CHECKPOINT:
Technique: Thumbs Up or Down
Script:
- Say: "Thumbs up if true, thumbs down if false."
- Scan for: thumbs down. 3 kg is 3000 g, a big number, not a tiny one.
PROCEED: If 80 percent thumbs down with a reason, click to reveal.
PIVOT: Most likely misconception - students divide when going down to a smaller unit.
- Reteach: "Kilograms down to grams means MORE of them, so multiply. 3 times 1000 is 3000 grams."
- Re-check: "What is 3 kilograms in grams? Show me."

TEACHER NOTES:
This checks the threshold idea: direction decides the operation. Going to a smaller unit multiplies, giving a bigger number. 0.003 g is a speck of dust, not 3 kg.

WATCH FOR:
- Confident thumbs down naming multiply - secure.
- Thumbs up - revisit smaller unit means more, so multiply.

[Stage 2: CFU Hinge | VTLM 2.0: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Time to work on your own. Take the practice sheet.
- Section one, kilograms to grams. Section two, grams to kilograms. Section three, a mix with tonnes and milligrams.
- Every time: set the stairs, count the steps, then multiply or divide.

DO:
- Distribute the practice sheet.
- Circulate. Sit with the enabling group first.
- Watch for the direction decision: bigger or smaller unit.

TEACHER NOTES:
The sheet sequences the multiply direction, then the divide direction, then a mix. Remember every mass step is x1000.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the staircase box and the worked first item at the top of the sheet. Stay with kilograms and grams only.
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
- One, convert 4 kilograms into grams.
- Two, convert 6000 grams into kilograms.
- Set the stairs and decide multiply or divide before you calculate.

DO:
- Display the prompt.
- Allow about 3 minutes.
- Collect whiteboards or take a photo for your records.

TEACHER NOTES:
Exit ticket assesses SC2 (converting with multiply or divide). Internal target is SC2. Do not display any SC label to students.

WATCH FOR:
- 4000 g and 6 kg - secure.
- A correct first answer but wrong direction on the second - revisit the bigger-unit-divide rule in Session 4.

[Stage 5: Exit Ticket | VTLM 2.0: Mastery and application]`;

const NOTES_CLOSING = `SAY:
- Look back at our success criteria.
- Show me thumbs for each one - got it, getting there, or need more practice.
- Turn and tell your partner: how is converting mass the same as converting length?

DO:
- Read each I can statement.
- Use thumbs to self-assess.
- Note who is still unsure for tomorrow.

TEACHER NOTES:
The big idea is the routine transfers: same staircase, same direction rule, just a different attribute. Students who see that are ready for capacity next session.

WATCH FOR:
- Students who explain the routine transfers - secure.
- Students who treat mass as brand new - revisit the shared routine in Session 4's launch.

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
  titleSlide(pres, UNIT_TITLE, "Session 3: Converting mass",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3-4. Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Money problems",
      [
        "A drink costs $2.75. How much for 4?",
        "Pay for a $6.40 item with $10. What change?",
        "Estimate $4.95 + $3.10 + $6.85 (nearest $)",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["$11.00", "$3.60", "about $15"], { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // 5-6. Fluency + reveal (division bracket)
  withReveal(
    () => fluencySlide(pres, "Fluency: Division bracket",
      ["175 ÷ 5", "216 ÷ 6", "294 ÷ 7"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["35", "36", "42"], { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // 7. Launch
  contentSlide(pres, "Launch", C.SECONDARY,
    "One kilogram — how many grams?",
    [
      "Picture a 1 kg bag of sugar.",
      "How many grams is that?",
      "What about 2 kg? Half a kilogram?",
      "Tell your partner how you worked it out.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Same mass, new unit",
        start: "1 kg", factor: "1000", divide: false, end: "1000 g",
        caption: "Converting keeps the mass the same. Only the unit changes.",
      });
    });

  // 8. LI / SC
  liSlide(pres,
    "We are learning to convert between metric units of mass.",
    [
      "I can say which units of mass are bigger and smaller.",
      "I can convert between units of mass using multiply or divide.",
      "I can convert masses that include decimals and explain my steps.",
    ],
    NOTES_LI_SC, FOOTER);

  // 9. Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.PRIMARY,
    "Mass — and every step is × 1000",
    [
      "Mass = how heavy something is.",
      "Units: mg, g, kg, tonne.",
      "",
      "Every mass step is × 1000 — simpler than length!",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      metricStaircase(slide, lg, {
        title: "Mass staircase",
        color: C.PRIMARY,
        steps: [{ u: "tonne" }, { u: "kg", f: 1000 }, { u: "g", f: 1000 }, { u: "mg", f: 1000 }],
      });
    });

  // 10. I Do 1 — kg -> g (multiply)
  workedExSlide(pres, 2, "I Do", "Convert 3 kg into g",
    [
      "Set the stairs: tonne, kg, g, mg.",
      "Count the steps: kg to g is × 1000.",
      "Going DOWN to a smaller unit → multiply.",
      "",
      "3 × 1000 = 3000.",
      "So 3 kg = 3000 g.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      metricStaircase(slide, lg, {
        title: "Mass staircase",
        color: C.PRIMARY,
        steps: [{ u: "tonne" }, { u: "kg", f: 1000 }, { u: "g", f: 1000 }, { u: "mg", f: 1000 }],
      });
    });

  // 11. I Do 2 — g -> kg (divide)
  workedExSlide(pres, 2, "I Do", "Convert 4500 g into kg",
    [
      "Count the steps: g to kg is × 1000.",
      "Going UP to a bigger unit → divide.",
      "",
      "4500 ÷ 1000 = 4.5.",
      "Every digit moves three places.",
      "So 4500 g = 4.5 kg.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Convert",
        start: "4500 g", factor: "1000", divide: true, end: "4.5 kg",
        caption: "Up to a bigger unit means fewer of them — so the number gets smaller.",
      });
    });

  // 12. I Do 3 — 1.25 kg -> g (equivalence)
  workedExSlide(pres, 2, "I Do", "Convert 1.25 kg into g",
    [
      "Count the steps: kg to g is × 1000.",
      "Going DOWN to a smaller unit → multiply.",
      "",
      "1.25 × 1000 = 1250.",
      "So 1.25 kg = 1250 g.",
      "These are equivalent — the same mass.",
    ],
    NOTES_IDO3, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Equivalent masses",
        start: "1.25 kg", factor: "1000", divide: false, end: "1250 g",
        caption: "1.25 kg and 1250 g are the same mass, written two ways.",
      });
    });

  // 13-14. CFU 1 + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Convert it", "Show Me Boards",
      "Convert 5 kg into grams.",
      NOTES_CFU1_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "5000 g   (5 × 1000 — down to a smaller unit, multiply)",
        { color: C.SUCCESS, label: "Answer", fontSize: 20 });
    }
  );

  // 15-16. We Do + reveal
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Convert 2500 g into kg",
      [
        "With your partner.",
        "",
        "Step 1: set the stairs.",
        "Step 2: count the steps (g to kg is × 1000).",
        "Step 3: bigger unit → divide. Calculate.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        convertPanel(slide, lg, {
          title: "Convert",
          start: "2500 g", factor: "1000", divide: true, end: "?",
          bottom: 4.15,
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "2.5 kg   (2500 ÷ 1000 — up to a bigger unit, divide)",
        { color: C.SUCCESS, label: "Answer", fontSize: 20 });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // 17-18. CFU hinge + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "True or false?", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n3 kg = 0.003 g\n\nIs this true or false?",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        "FALSE. 3 kg = 3000 g (3 × 1000). Down to a smaller unit, you multiply.",
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
      { text: "Section 1 — kilograms to grams.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 2 — grams to kilograms.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 3 — a mix with tonnes and mg.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "1.  Set the stairs: tonne, kg, g, mg.", {
      x: 1.0, y: panelY + 0.54, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "2.  Count the steps — for mass, always × 1000.", {
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
      "Convert 4 kg into grams.",
      "Convert 6000 g into kilograms.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 21. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how is converting mass the same as converting length?",
      scItems: [
        "I can say which units of mass are bigger and smaller.",
        "I can convert between units of mass using multiply or divide.",
        "I can convert masses that include decimals and explain my steps.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MMC_Session3_Converting_Mass.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  await generatePdfs();
  console.log("Session 3 build complete.");
}

// ─── PDFs ─────────────────────────────────────────────────────────────────

async function generatePdfs() {
  // ── Worksheet ──────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Convert between mg, g, kg and tonne. Set the stairs, count the steps, multiply or divide.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addTipBox(doc,
      "Staircase (largest to smallest): tonne, kg, g, mg.  For mass, EVERY step is x1000.  DOWN to a smaller unit: multiply.  UP to a bigger unit: divide.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Kilograms to grams (x 1000)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   2 kg   =   _______ g      (worked: 2 x 1000 = 2000)", y);
    y = addWriteLine(doc, "b)   5 kg   =   _______ g", y);
    y = addWriteLine(doc, "c)   1.5 kg =   _______ g", y);

    y = addSectionHeading(doc, "Section 2 — Grams to kilograms (div 1000)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   3000 g =   _______ kg", y);
    y = addWriteLine(doc, "b)   4500 g =   _______ kg", y);
    y = addWriteLine(doc, "c)   750 g   =   _______ kg", y);

    y = addSectionHeading(doc, "Section 3 — A mix (watch the unit!)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   2 tonne =   _______ kg", y);
    y = addWriteLine(doc, "b)   4 g     =   _______ mg", y);
    y = addWriteLine(doc, "c)   1.25 kg =   _______ g", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc, "1.  A recipe needs 250 g of flour. You make 4 batches. How many kilograms of flour is that? (Convert at the end.)", y);
    y = addWriteLine(doc, "      Working: ______________________   Answer: _______ kg", y);
    y = addBodyText(doc, "2.  Order from lightest to heaviest:  0.5 kg,  750 g,  1200 mg.", y);
    y = addWriteLine(doc, "      ______________ , ______________ , ______________", y);

    addPdfFooter(doc, `Session ${SESSION} | Converting Mass | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // ── Answer Key ─────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the converting-mass practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Kilograms to grams", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2000 g (worked).   b)  5000 g.   c)  1500 g.", y);

    y = addSectionHeading(doc, "Section 2 — Grams to kilograms", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3 kg.   b)  4.5 kg.   c)  0.75 kg.", y);

    y = addSectionHeading(doc, "Section 3 — A mix", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2000 kg (2 x 1000).   b)  4000 mg (4 x 1000).   c)  1250 g (1.25 x 1000).", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc, "1.  250 x 4 = 1000 g = 1 kg.", y);
    y = addBodyText(doc, "2.  Convert to grams: 0.5 kg = 500 g, 750 g, 1200 mg = 1.2 g. Lightest to heaviest: 1200 mg, 0.5 kg, 750 g.", y);

    y = addTipBox(doc,
      "Watch for: students who always multiply. The direction decides the operation - down to a smaller unit multiply, up to a bigger unit divide. Check the decimal moves: 4500 g is 4.5 kg, not 45 kg.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
