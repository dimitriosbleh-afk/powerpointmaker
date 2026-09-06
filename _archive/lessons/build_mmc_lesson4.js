"use strict";

// Metric Measurement & Unit Conversion (Year 6 Numeracy)
// Session 4 of 5: Converting capacity (mL, L, kL, ML) and equivalence.
// VC2M6M01 — using the correct operation when converting between units of
//            capacity; recognising equivalence such as 1.25 L = 1250 mL.
// Daily Review: Practical Problem Solving and Financial Reasoning (prior learning).
// Fluency: Division bracket (long-division layout). NOT the bus stop method.
//
// Unit routine (consistent across all 5 sessions):
//   SET THE STAIRS  ->  COUNT THE STEPS  ->  MULTIPLY OR DIVIDE.
// Self-contained for catch-up: this session re-teaches the staircase method
// from scratch. A student who missed earlier sessions can still access today.

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

const SESSION = 4;
const TOTAL = 5;
const UNIT_TITLE = "Metric Measurement & Conversion";
const FOOTER = `Metric Measurement | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/MMC_Session4_Converting_Capacity";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Converting Capacity Practice",
  "Convert between mL, L, kL and ML, and show equivalent capacities.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the converting-capacity practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back, mathematicians. We can convert length and mass. Today we convert capacity - how much a container holds.
- Same routine again, and like mass, every step is times one thousand.
- We will also prove when two capacities are equivalent.

DO:
- Have whiteboards and markers ready.
- Keep the printed practice sheet face down until the You Do section.

TEACHER NOTES:
Session 4 of 5. The conversion routine is re-taught from scratch, so students who missed earlier sessions can still access today. Today we lean into equivalence, such as 1.25 L equals 1250 mL.

WATCH FOR:
- Students who confuse capacity units with mass units. Litres measure how much it holds, not how heavy it is.

[General: Title | VTLM 2.0: Attention, focus and regulation]`;

const NOTES_RESOURCES = `SAY:
- Here is what we are using today.
- Whiteboards for most of the lesson, and the practice sheet near the end.

DO:
- Print one practice sheet and one answer key per student who needs paper.
- Have a 1 litre bottle and a measuring jug handy if you can.

TEACHER NOTES:
One student resource today plus an answer key. A 1 litre bottle anchors 1 L = 1000 mL.

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_DR_Q = `SAY:
- Quick warm up on money problems.
- Read each one, then write your answer on your whiteboard.
- These are from earlier learning.

DO:
- Display the three prompts.
- Allow about 90 seconds.

TEACHER NOTES:
Daily Review revisits Practical Problem Solving and Financial Reasoning. Sharing, subtracting and estimating keep the skills warm.

WATCH FOR:
- Students who divide to share the ticket cost - secure.
- Students who guess - prompt them to write the operation first.

[Stage 1: Daily Review | VTLM 2.0: Retention and recall]`;

const NOTES_DR_A = `SAY:
- Let us check together.
- Three tickets for 27 dollars means one ticket is 9 dollars. 27 divided by 3.
- A 40 dollar jacket reduced by 15 dollars is 25 dollars.
- Estimating, 19 dollars 80 plus 9 dollars 95 rounds to 20 plus 10, about 30 dollars.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Dividing 27 by 3 is a nice link to today, where dividing converts to a bigger unit.

WATCH FOR:
- Confident self-correction - secure.
- A subtraction slip on the jacket - re-check 40 take 15.

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
- 168 divided by 8 is 21.
- 234 divided by 6 is 39.
- 315 divided by 5 is 63.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
315 divided by 5: 31 divided by 5 is 6 remainder 1, carry the 1 to make 15, 15 divided by 5 is 3.

WATCH FOR:
- Clean carrying - secure.
- Answers near 60 for 315 divided by 5 - re-check the carry.

[Stage 1: Fluency Answers | VTLM 2.0: Retention and recall]`;

const NOTES_LAUNCH = `SAY:
- Picture a large water bottle. It holds 1 litre. How many millilitres is that? One thousand.
- So 1 litre is the same capacity as 1000 millilitres. We did not change how much it holds, only the unit.
- Now think: how many millilitres in 2 litres? In half a litre? Tell your partner.

DO:
- Hold up a 1 litre bottle if you have one.
- Give 40 seconds of partner talk.
- Take responses. Listen for times one thousand.

TEACHER NOTES:
The launch connects what students know - 1000 mL in a litre - to today's rule. 2 L is 2000 mL; half a litre is 500 mL. Same routine as mass.

WATCH FOR:
- Students who multiply by 1000 to get millilitres - exactly today's move.
- Students who think 2 L is 1002 mL - they added instead of scaling.

[Stage 1: Launch | VTLM 2.0: Knowledge and memory]`;

const NOTES_LI_SC = `SAY:
- Here is what we are learning today.
- We are learning to convert between metric units of capacity.
- Read the success criteria with me.

DO:
- Choral read the learning intention.
- Read each I can statement.

TEACHER NOTES:
SC1 is reachable for everyone - knowing which capacity units are bigger and smaller. SC2 is the core target the exit ticket assesses - converting with multiply or divide. SC3 stretches into proving equivalence. Tier labels stay off the slide.

WATCH FOR:
- Students repeating the language back - tracking with us.

[General: LI/SC | VTLM 2.0: Planning]`;

const NOTES_VOCAB = `SAY:
- Capacity means how much a container can hold. We measure it in millilitres, litres, kilolitres and megalitres.
- Notice the staircase. Like mass, every capacity step is times one thousand.
- Remember equivalent means equal in value. 1.25 litres and 1250 millilitres are equivalent.

DO:
- Point to the capacity staircase as you name each step.
- Have students say the four units in order, largest to smallest.

TEACHER NOTES:
The key ideas today: every capacity step is x1000, and equivalence. Keep returning to the idea that the amount held does not change, only the unit.

WATCH FOR:
- Students who notice the steps are all x1000 - secure.
- Students who think a bigger number means more liquid - 1250 mL is not more than 1.25 L; revisit equivalence.

[General: Key Vocabulary | VTLM 2.0: Knowledge and memory]`;

const NOTES_IDO1 = `SAY:
- Let us convert 2 litres into millilitres together. Watch my routine.
- First I set the stairs: ML, kL, L, mL.
- Next I count the steps from litres to millilitres. One step, times one thousand.
- I am going down to a smaller unit, so I multiply. 2 times 1000 is 2000.
- So 2 litres is 2000 millilitres. Same capacity, new unit.

DO:
- Point to L and mL on the staircase.
- Write 2 times 1000 equals 2000 on the board.

TEACHER NOTES:
This is the I Do for the multiply direction. Down to a smaller unit means more of them, so multiply.

MISCONCEPTIONS:
- Misconception: to get millilitres you divide because millilitres are small.
  Why: students link a small unit with the smaller operation.
  Impact: 2 L becomes 0.002 mL, which is clearly wrong.
  Quick correction: a smaller unit means you need MORE of them, so the number grows - multiply.

WATCH FOR:
- Students who multiply going down - secure.
- Students who divide - ask: more or fewer millilitres than litres?

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now the other direction. Convert 3500 millilitres into litres.
- Set the stairs. Count the steps from millilitres up to litres - one step, times one thousand.
- I am going up to a bigger unit, so I divide. 3500 divided by 1000.
- Dividing by 1000 moves every digit three places. 3500 divided by 1000 is 3.5.
- So 3500 millilitres is 3.5 litres.

DO:
- Point to mL and L on the panel.
- Show 3500 divided by 1000 equals 3.5.

TEACHER NOTES:
This is the divide direction. Up to a bigger unit means fewer of them, so divide. Connect dividing by 1000 to moving digits three places.

MISCONCEPTIONS:
- Misconception: just delete three zeros, so 3500 becomes 3.
  Why: it works for round thousands by accident.
  Impact: students lose the 5, getting 3 L instead of 3.5 L.
  Quick correction: dividing by 1000 moves the digits three places - 3500 becomes 3.5.

WATCH FOR:
- Students who get 3.5 L - secure.
- Students who get 3 L or 35 L - re-count the place value moves.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO3 = `SAY:
- One more, the equivalence example. Convert 1.25 litres into millilitres.
- Set the stairs. Litres down to millilitres is one step, times one thousand.
- Down to a smaller unit, so I multiply. 1.25 times 1000 moves every digit three places.
- 1.25 times 1000 is 1250. So 1.25 litres equals 1250 millilitres.
- These are equivalent - exactly the same amount, written two ways.

DO:
- Write 1.25 times 1000 equals 1250.
- Say the word equivalent and point to both forms.

TEACHER NOTES:
This is the capacity equivalence example: 1.25 L equals 1250 mL. It pulls together multiplying by 1000 and equivalence. A strong model to return to in the You Do.

MISCONCEPTIONS:
- Misconception: 1250 mL must be more liquid than 1.25 L because 1250 is a bigger number.
  Why: students compare the numbers, not the amounts.
  Impact: they doubt a correct conversion.
  Quick correction: they are the same amount. The number changed because the unit changed.

WATCH FOR:
- Students who say equivalent and explain why - secure and ready for depth.
- Students unsure about the decimal - re-show the three-place move.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU1_Q = `SAY:
- Quick check on your whiteboards.
- Convert 4 litres into millilitres.
- Show me on three, two, one.

DO:
- Display the prompt.
- Allow about 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me 4 litres in millilitres on three, two, one, show."
- Scan for: 4000 mL.
PROCEED: If 80 percent show 4000, click to reveal and move on.
PIVOT: Most likely misconception - students divide because millilitres feel small.
- Reteach: "Down to a smaller unit means more of them. 4 litres is 4000 millilitres. Multiply by 1000."
- Re-check: "Now convert 7 litres into millilitres."

TEACHER NOTES:
Checks the multiply direction before guided practice.

WATCH FOR:
- 4000 mL - secure.
- 0.004 mL or 400 mL - re-count the step and the direction.

[Stage 2: CFU | VTLM 2.0: Supported Application]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Convert 2500 millilitres into litres.
- Step one, set the stairs. Step two, count the steps from mL to L. Step three, decide multiply or divide, then calculate.
- Work it with your partner on your whiteboard.

DO:
- Keep the steps visible.
- Walk and scan. Prompt with "bigger or smaller unit" where needed.
- Take 60 seconds before revealing.

TEACHER NOTES:
Guided practice in the divide direction. Up to a bigger unit means divide by 1000. Watch for the decimal answer.

WATCH FOR:
- Pairs who divide by 1000 to get 2.5 L - secure.
- Pairs who multiply - point them up the staircase: fewer litres than millilitres.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- Millilitres up to litres is one step, times one thousand, and we are going up, so divide.
- 2500 divided by 1000 is 2.5.
- So 2500 millilitres is 2.5 litres.

DO:
- Click to reveal.
- Ask one pair to explain the direction.

TEACHER NOTES:
The reveal restates the direction and the calculation. 2.5 L shows the decimal place value at work.

WATCH FOR:
- Self-correction to 2.5 L - secure.
- Answers of 25 L or 0.25 L - re-count the place value move.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. A student wrote: 250 millilitres equals 250000 litres.
- Do not just fix it. Decide: true or false?
- Thumbs up for true, thumbs down for false.

DO:
- Display the claim.
- Give 15 seconds of thinking time.

CFU CHECKPOINT:
Technique: Thumbs Up or Down
Script:
- Say: "Thumbs up if true, thumbs down if false."
- Scan for: thumbs down. 250 mL is less than a litre, only 0.25 L.
PROCEED: If 80 percent thumbs down with a reason, click to reveal.
PIVOT: Most likely misconception - students multiply when going up to a bigger unit.
- Reteach: "Millilitres up to litres means FEWER of them, so divide. 250 divided by 1000 is 0.25 litres."
- Re-check: "What is 250 millilitres in litres? Show me."

TEACHER NOTES:
This checks the threshold idea: direction decides the operation. Going to a bigger unit divides. 250000 L is an enormous tank, not a small cup.

WATCH FOR:
- Confident thumbs down naming divide - secure.
- Thumbs up - revisit bigger unit means fewer, so divide.

[Stage 2: CFU Hinge | VTLM 2.0: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Time to work on your own. Take the practice sheet.
- Section one, litres to millilitres. Section two, millilitres to litres. Section three, a mix with kilolitres and decimals.
- Every time: set the stairs, count the steps, then multiply or divide.

DO:
- Distribute the practice sheet.
- Circulate. Sit with the enabling group first.
- Watch for the direction decision: bigger or smaller unit.

TEACHER NOTES:
The sheet sequences the multiply direction, then the divide direction, then a mix. Every capacity step is x1000.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the staircase box and the worked first item at the top of the sheet. Stay with litres and millilitres only.
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
- One, convert 6 litres into millilitres.
- Two, convert 1500 millilitres into litres.
- Set the stairs and decide multiply or divide before you calculate.

DO:
- Display the prompt.
- Allow about 3 minutes.
- Collect whiteboards or take a photo for your records.

TEACHER NOTES:
Exit ticket assesses SC2 (converting with multiply or divide). Internal target is SC2. Do not display any SC label to students.

WATCH FOR:
- 6000 mL and 1.5 L - secure.
- A correct first answer but wrong direction on the second - revisit the bigger-unit-divide rule in Session 5.

[Stage 5: Exit Ticket | VTLM 2.0: Mastery and application]`;

const NOTES_CLOSING = `SAY:
- Look back at our success criteria.
- Show me thumbs for each one - got it, getting there, or need more practice.
- Turn and tell your partner: what does it mean to say 1.25 litres and 1250 millilitres are equivalent?

DO:
- Read each I can statement.
- Use thumbs to self-assess.
- Note who is still unsure for tomorrow.

TEACHER NOTES:
The big idea is equivalence: the same amount can be written in different units. Next session we put length, mass and capacity together in real problems.

WATCH FOR:
- Students who explain equivalence in their own words - secure.
- Students who think a bigger number means more - revisit with the 1.25 L example.

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
  titleSlide(pres, UNIT_TITLE, "Session 4: Converting capacity",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3-4. Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Money problems",
      [
        "3 tickets cost $27 in total. Price of one?",
        "A $40 jacket is reduced by $15. New price?",
        "Estimate $19.80 + $9.95 (nearest $)",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["$9", "$25", "about $30"], { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // 5-6. Fluency + reveal (division bracket)
  withReveal(
    () => fluencySlide(pres, "Fluency: Division bracket",
      ["168 ÷ 8", "234 ÷ 6", "315 ÷ 5"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["21", "39", "63"], { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // 7. Launch
  contentSlide(pres, "Launch", C.SECONDARY,
    "One litre — how many mL?",
    [
      "Picture a 1 litre water bottle.",
      "How many millilitres is that?",
      "What about 2 litres? Half a litre?",
      "Tell your partner how you worked it out.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Same amount, new unit",
        start: "1 L", factor: "1000", divide: false, end: "1000 mL",
        caption: "Converting keeps the amount the same. Only the unit changes.",
      });
    });

  // 8. LI / SC
  liSlide(pres,
    "We are learning to convert between metric units of capacity.",
    [
      "I can say which units of capacity are bigger and smaller.",
      "I can convert between units of capacity using multiply or divide.",
      "I can show that two capacities are equivalent, like 1.25 L and 1250 mL.",
    ],
    NOTES_LI_SC, FOOTER);

  // 9. Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.PRIMARY,
    "Capacity — every step is × 1000",
    [
      "Capacity = how much a container holds.",
      "Units: mL, L, kL, ML.",
      "",
      "Equivalent = equal in value (1.25 L = 1250 mL).",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      metricStaircase(slide, lg, {
        title: "Capacity staircase",
        color: C.PRIMARY,
        steps: [{ u: "ML" }, { u: "kL", f: 1000 }, { u: "L", f: 1000 }, { u: "mL", f: 1000 }],
      });
    });

  // 10. I Do 1 — L -> mL (multiply)
  workedExSlide(pres, 2, "I Do", "Convert 2 L into mL",
    [
      "Set the stairs: ML, kL, L, mL.",
      "Count the steps: L to mL is × 1000.",
      "Going DOWN to a smaller unit → multiply.",
      "",
      "2 × 1000 = 2000.",
      "So 2 L = 2000 mL.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      metricStaircase(slide, lg, {
        title: "Capacity staircase",
        color: C.PRIMARY,
        steps: [{ u: "ML" }, { u: "kL", f: 1000 }, { u: "L", f: 1000 }, { u: "mL", f: 1000 }],
      });
    });

  // 11. I Do 2 — mL -> L (divide)
  workedExSlide(pres, 2, "I Do", "Convert 3500 mL into L",
    [
      "Count the steps: mL to L is × 1000.",
      "Going UP to a bigger unit → divide.",
      "",
      "3500 ÷ 1000 = 3.5.",
      "Every digit moves three places.",
      "So 3500 mL = 3.5 L.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Convert",
        start: "3500 mL", factor: "1000", divide: true, end: "3.5 L",
        caption: "Up to a bigger unit means fewer of them — so the number gets smaller.",
      });
    });

  // 12. I Do 3 — 1.25 L -> mL (equivalence)
  workedExSlide(pres, 2, "I Do", "Convert 1.25 L into mL",
    [
      "Count the steps: L to mL is × 1000.",
      "Going DOWN to a smaller unit → multiply.",
      "",
      "1.25 × 1000 = 1250.",
      "So 1.25 L = 1250 mL.",
      "These are equivalent — the same amount.",
    ],
    NOTES_IDO3, FOOTER,
    (slide, lg) => {
      convertPanel(slide, lg, {
        title: "Equivalent capacities",
        start: "1.25 L", factor: "1000", divide: false, end: "1250 mL",
        caption: "1.25 L and 1250 mL hold the same amount, written two ways.",
      });
    });

  // 13-14. CFU 1 + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Convert it", "Show Me Boards",
      "Convert 4 L into millilitres.",
      NOTES_CFU1_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "4000 mL   (4 × 1000 — down to a smaller unit, multiply)",
        { color: C.SUCCESS, label: "Answer", fontSize: 20 });
    }
  );

  // 15-16. We Do + reveal
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Convert 2500 mL into L",
      [
        "With your partner.",
        "",
        "Step 1: set the stairs.",
        "Step 2: count the steps (mL to L is × 1000).",
        "Step 3: bigger unit → divide. Calculate.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        convertPanel(slide, lg, {
          title: "Convert",
          start: "2500 mL", factor: "1000", divide: true, end: "?",
          bottom: 4.15,
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "2.5 L   (2500 ÷ 1000 — up to a bigger unit, divide)",
        { color: C.SUCCESS, label: "Answer", fontSize: 20 });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // 17-18. CFU hinge + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "True or false?", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n250 mL = 250 000 L\n\nIs this true or false?",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        "FALSE. 250 mL = 0.25 L (250 ÷ 1000). Up to a bigger unit, you divide.",
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
      { text: "Section 1 — litres to millilitres.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 2 — millilitres to litres.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 3 — a mix with kL and decimals.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "1.  Set the stairs: ML, kL, L, mL.", {
      x: 1.0, y: panelY + 0.54, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "2.  Count the steps — for capacity, always × 1000.", {
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
      "Convert 6 L into millilitres.",
      "Convert 1500 mL into litres.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 21. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what does it mean to say 1.25 L and 1250 mL are equivalent?",
      scItems: [
        "I can say which units of capacity are bigger and smaller.",
        "I can convert between units of capacity using multiply or divide.",
        "I can show that two capacities are equivalent, like 1.25 L and 1250 mL.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MMC_Session4_Converting_Capacity.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  await generatePdfs();
  console.log("Session 4 build complete.");
}

// ─── PDFs ─────────────────────────────────────────────────────────────────

async function generatePdfs() {
  // ── Worksheet ──────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Convert between mL, L, kL and ML. Set the stairs, count the steps, multiply or divide.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addTipBox(doc,
      "Staircase (largest to smallest): ML, kL, L, mL.  For capacity, EVERY step is x1000.  DOWN to a smaller unit: multiply.  UP to a bigger unit: divide.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Litres to millilitres (x 1000)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   2 L   =   _______ mL      (worked: 2 x 1000 = 2000)", y);
    y = addWriteLine(doc, "b)   5 L   =   _______ mL", y);
    y = addWriteLine(doc, "c)   1.5 L =   _______ mL", y);

    y = addSectionHeading(doc, "Section 2 — Millilitres to litres (div 1000)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   3000 mL =   _______ L", y);
    y = addWriteLine(doc, "b)   4500 mL =   _______ L", y);
    y = addWriteLine(doc, "c)   750 mL   =   _______ L", y);

    y = addSectionHeading(doc, "Section 3 — A mix (watch the unit!)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   2 kL   =   _______ L", y);
    y = addWriteLine(doc, "b)   1.25 L =   _______ mL", y);
    y = addWriteLine(doc, "c)   6000 L =   _______ kL", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc, "1.  A jug holds 1.5 L. You pour out 600 mL. How many millilitres are left? (Convert first.)", y);
    y = addWriteLine(doc, "      Working: ______________________   Answer: _______ mL", y);
    y = addBodyText(doc, "2.  Order from smallest to largest:  0.4 L,  450 mL,  1.2 L.", y);
    y = addWriteLine(doc, "      ______________ , ______________ , ______________", y);

    addPdfFooter(doc, `Session ${SESSION} | Converting Capacity | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // ── Answer Key ─────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the converting-capacity practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Litres to millilitres", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2000 mL (worked).   b)  5000 mL.   c)  1500 mL.", y);

    y = addSectionHeading(doc, "Section 2 — Millilitres to litres", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3 L.   b)  4.5 L.   c)  0.75 L.", y);

    y = addSectionHeading(doc, "Section 3 — A mix", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2000 L (2 x 1000).   b)  1250 mL (1.25 x 1000).   c)  6 kL (6000 div 1000).", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc, "1.  1.5 L = 1500 mL. 1500 - 600 = 900 mL left.", y);
    y = addBodyText(doc, "2.  Convert to mL: 0.4 L = 400 mL, 450 mL, 1.2 L = 1200 mL. Smallest to largest: 0.4 L, 450 mL, 1.2 L.", y);

    y = addTipBox(doc,
      "Watch for: students who always multiply. The direction decides the operation - down to a smaller unit multiply, up to a bigger unit divide. Check the decimal moves: 4500 mL is 4.5 L, not 45 L.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
