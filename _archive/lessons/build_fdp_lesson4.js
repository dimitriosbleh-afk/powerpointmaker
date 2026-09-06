"use strict";

// Mastering Fractions, Decimals and Percentages (Year 6 Numeracy) - Lesson 4 of 4.
// Estimate with fractions, decimals and percentages, and use estimates to check
// whether an answer is reasonable. Round to friendly forms (0.3 is about 1/3;
// 52% is about 1/2); common percentages 10%, 25%, 30%, 50%, 1%. Verify discounts
// by estimating. VC2M6N08 / VC2M5N04.
// Daily Review: mixed FDP from the unit (prior). Fluency: decimal column
// addition with ragged places.
// Unit variant fixed (variant 1) for cohesion across the four lessons.
// CATCH-UP: the launch rebuilds "round to a friendly number, then estimate"
// from a shopping example; worksheet Section 1 re-grounds it. Ships the unit's
// single Year 8 extension worksheet as a stretch resource.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(2)); // variant 1, fixed for the unit
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide, dailyReviewSlide,
  addTenthsStrip,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics, getContrastColor,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const TOTAL = 4;
const UNIT_TITLE = "Mastering Fractions, Decimals & Percentages";
const FOOTER = `Fractions, Decimals & Percentages | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FDP_Lesson4_Estimating_With_FDP";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Estimate with fractions, decimals and percentages, and check if answers are reasonable.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 4 practice sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Percentage increase, decrease and working backwards - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Shared bar-model helper (unit visual anchor) ---------------------------
function partBar(slide, x, y, w, h, parts, shaded, opts) {
  const o = opts || {};
  const segW = w / parts;
  const fill = o.fill || C.SUCCESS;
  for (let i = 0; i < parts; i++) {
    const isShaded = i < shaded;
    slide.addShape("rect", {
      x: x + i * segW, y, w: segW, h,
      fill: { color: isShaded ? fill : C.WHITE },
      line: { color: C.PRIMARY, width: 1.25 },
    });
    if (o.segLabels && o.segLabels[i] != null) {
      slide.addText(String(o.segLabels[i]), {
        x: x + i * segW, y, w: segW, h,
        fontSize: o.segFontSize || 14, fontFace: FONT_B,
        color: isShaded ? getContrastColor(fill) : C.CHARCOAL,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
    }
  }
}

function stepPills(slide, x, y, w, rows, opts) {
  const o = opts || {};
  const h = o.h || 0.52;
  const gap = o.gap || 0.12;
  const fs = o.fontSize || 16;
  rows.forEach((r, i) => {
    addTextOnShape(slide, r.text, {
      x, y: y + i * (h + gap), w, h, rectRadius: 0.08,
      fill: { color: r.color || C.PRIMARY },
    }, { fontSize: fs, fontFace: FONT_H, color: C.WHITE, bold: true });
  });
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to the last lesson of our fractions, decimals and percentages unit.
- We can already find fractions and percentages of amounts, and work out discounts.
- Today we learn to estimate sensibly and, just as importantly, to use an estimate to check whether an answer is reasonable.

DO:
- Have whiteboards and markers ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 4 of 4. The anchor is "round to a friendly number, estimate, then check". Estimation is the safety net for all the exact work in Lessons 1 to 3.

WATCH FOR:
- Students who think estimating is just guessing - reassure them it is a smart, quick, close answer.

[General: Title | Element: Attention, focus and regulation]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- There is a practice sheet, and a Year 8 extension for anyone who is flying.

DO:
- Print one practice sheet and one answer key per student.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student practice sheet, an answer key, and the unit's Year 8 extension. The extension covers percentage increase, decrease and working backwards.

CATCH-UP NOTE:
A student who missed earlier lessons can still access today. The launch rebuilds "round to a friendly number, then estimate" from a shopping example that stands alone. Section 1 of the worksheet re-grounds it. A returner needs only the worksheet and one minute with you to start.

[General: Resources | Element: Enabling Learning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are warming up everything from this unit.
- Read each one and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for the find-one-part-then-scale move.

TEACHER NOTES:
Daily Review pulls together a fraction of an amount, a fraction-to-percent conversion, and a 10% calculation - all prior to today's estimation.

WATCH FOR:
- Students who use the unit methods fluently - secure.
- Students who stall - point them to a quick bar sketch.

[Stage 1: Daily Review | Element: Retention and recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- One third of 60 is 20.
- Nought point two five is 25 percent.
- 10 percent of 80 dollars is 8 dollars.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
These three retrieve the core of the unit. Note any student who multiplies instead of dividing for 10%.

WATCH FOR:
- Students who self-correct quickly - secure.

[Stage 1: Daily Review Answers | Element: Retention and recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Adding decimals with the vertical algorithm.
- These have different numbers of decimal places. Line up the points, and fill empty places with a zero if it helps.
- Tenths under tenths, hundredths under hundredths.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up decimal points and zero place-fillers.

TEACHER NOTES:
Ragged decimals are exactly where alignment matters most. Writing 0.4 as 0.40 to match places is a fair strategy and connects to estimating with decimals today.

WATCH FOR:
- Students who line up the points - secure.
- Students who right-align the digits - reteach: points under points, fill with a zero.

[Stage 1: Fluency | Element: Retention and recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 1.4 plus 0.35 is 1.75.
- 2.5 plus 0.75 is 3.25.
- 3.6 plus 0.4 is 4.0, which we can write as 4.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The third lands on a whole number, which surprises some students. 4.0 is the same as 4.

WATCH FOR:
- Students who self-correct - secure.
- Students who misalign 0.35 - small group focus.

[Stage 1: Fluency Answers | Element: Retention and recall]`;

const NOTES_LAUNCH = `SAY:
- Imagine a shirt priced at 19 dollars 95, with 50 percent off. Do we really need the exact answer to know roughly what we will pay?
- 19 dollars 95 is about 20 dollars. Half off 20 dollars is about 10 dollars. So we will pay roughly 10 dollars.
- That is estimating: round to a friendly number, then use an easy fraction or percentage.
- An estimate is not the exact answer, but it tells us straight away if an answer makes sense.

DO:
- Write $19.95, round it to $20, then halve to about $10.
- Have students chorus "round to a friendly number, then estimate".
- Stress: estimate first, then we can check the exact answer is sensible.

TEACHER NOTES:
This launch builds estimation from a shopping example everyone can picture, so it does not depend on earlier lessons. It is the catch-up bridge for today.

WATCH FOR:
- Students who round and halve quickly - strong number sense.
- Students unsure - model rounding $19.95 to $20 explicitly.

[Stage: Launch | Element: Knowledge and memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to estimate with fractions, decimals and percentages, and to use estimates to check if an answer is reasonable.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Point to the words "friendly", "estimate" and "reasonable".

TEACHER NOTES:
The first criterion is reachable for everyone - round to a friendly number. The second is the core target the exit ticket checks. The third stretches to judging whether a given answer is reasonable.

WATCH FOR:
- Students who can say "round, then estimate, then check" - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- An estimate is a sensible, close answer that is quick to work out. It is not exact.
- About, or approximately, means near to. The wavy equals sign means approximately equal.
- Reasonable means the answer makes sense - it is close to our estimate.

DO:
- Point to each word as you say it.
- Have students say "estimate first, then check it is reasonable" once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. The key shift is that an estimate is a tool for checking, not a lazy answer.

WATCH FOR:
- Students who can explain estimate versus exact - secure.
- Students who think estimate means guess - reframe as a smart, close answer.

[General: Key Vocabulary | Element: Knowledge and memory]`;

const NOTES_IDO1 = `SAY:
- Let us work through one together. I want to estimate nought point three of 180.
- Nought point three is very close to one third. So nought point three of 180 is about one third of 180.
- One third of 180 is 180 divided by 3, which is 60. So my estimate is about 60.
- The exact answer is 54, which is close to 60. The friendly fraction gave me a fast, sensible estimate.

DO:
- Show 0.3 as three tenths next to 1/3 to highlight how close they are.
- Write "about 1/3 of 180 = 60".
- Note the exact 54 underneath to show the estimate is close.

TEACHER NOTES:
This is the elaboration example exactly. Swapping a decimal for a near friendly fraction is a powerful estimation move. Close, not exact, is the point.

MISCONCEPTIONS:
- Misconception: students think an estimate must equal the exact answer.
  Why: school maths usually rewards the exact answer.
  Impact: they distrust a good estimate of 60 because the exact is 54.
  Quick correction: an estimate is meant to be close and fast, not exact - 60 is a great estimate for 54.

WATCH FOR:
- Students who swap 0.3 for 1/3 - secure.
- Students who try to be exact - remind them this is an estimate.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now we use estimation to check. A toy costs 39 dollars 95 with 25 percent off, and the website says you save 9 dollars 99. Is that reasonable?
- Round 39 dollars 95 to 40 dollars. 25 percent is one quarter, so about 40 divided by 4, which is 10 dollars off.
- So a saving of about 10 dollars is what we expect. The website says 9 dollars 99, which is about 10 dollars. That is reasonable.
- Estimating first gives us a number to check against, so we catch silly mistakes.

DO:
- Round $39.95 to $40, then quarter it to about $10.
- Compare to the stated $9.99 saving.
- Mark the answer as reasonable.

TEACHER NOTES:
This is the verify-by-estimating idea. The estimate is the yardstick. If the stated answer is wildly different from the estimate, something is wrong.

MISCONCEPTIONS:
- Misconception: students think they must redo the exact calculation to check.
  Why: they trust only exact arithmetic.
  Impact: checking becomes slow, so they skip it.
  Quick correction: a quick estimate is enough to spot a wrong answer; you do not need the exact value to check it is sensible.

WATCH FOR:
- Students who compare estimate to stated answer - secure.
- Students who ignore the estimate - prompt them to use it as the check.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Estimate 48 percent of 200.
- Round the percentage to a friendly one first.

DO:
- Display the prompt.
- Give 45 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 48% is about 50%, half of 200 is 100. Estimate about 100.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students try to work out 48% exactly.
- Reteach: 48% is very close to 50%, which is one half. Half of 200 is 100. About 100.
- Re-check: what friendly percentage is 48% closest to, and what is half of 200?

TEACHER NOTES:
The move is rounding 48% up to 50%. A student doing long multiplication has missed the point of estimation.

WATCH FOR:
- Students who write about 100 - secure.
- Students who grind out 96 exactly - praise the accuracy, but coach the faster estimate.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Estimate 9 percent of 80 dollars.
- 9 percent is very close to a friendly percentage. Which one?
- Whisper to your partner: what is 10 percent of 80 dollars?

DO:
- Display "Estimate 9% of $80".
- Give 75 seconds.
- Listen for "9% is about 10%".

TEACHER NOTES:
9% rounds to 10%, which is dividing by 10. 10% of 80 is 8 dollars, so about 8 dollars. The exact is 7.20, close to the estimate.

WATCH FOR:
- Pairs who round 9% to 10% - secure.
- Pairs who try exact - coach the friendly round.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- 9 percent is about 10 percent. 10 percent of 80 dollars is 8 dollars.
- So 9 percent of 80 dollars is about 8 dollars. The exact answer is 7 dollars 20, which is close.
- Round to a friendly percentage, then it is easy.

DO:
- Click to reveal.
- Note the estimate 8 dollars is close to the exact 7 dollars 20.

TEACHER NOTES:
A great moment to value a close estimate over a slow exact answer.

WATCH FOR:
- Students who self-correct - secure.
- Students who distrust the estimate - remind them it is meant to be close.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, and this time we are detectives.
- A calculator says 30 percent of 60 dollars is 1 dollar 80. Estimate first, then decide: is that reasonable?
- Find a quick estimate for 30 percent of 60 dollars.

DO:
- Display "Calculator: 30% of $60 = $1.80. Reasonable?".
- Give 90 seconds.
- Listen for "10% of 60 is 6 dollars, so 30% is about 18 dollars".

TEACHER NOTES:
10% of 60 is 6 dollars, 30% is 3 times 6, which is 18 dollars. The calculator answer of 1 dollar 80 is ten times too small - probably a misplaced decimal point. Estimation catches it.

WATCH FOR:
- Students who say "not reasonable" with a reason - secure.
- Students who trust the calculator - prompt: what did your estimate say?

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- 10 percent of 60 dollars is 6 dollars, so 30 percent is about 18 dollars.
- The calculator says 1 dollar 80, which is ten times too small. That is not reasonable - the decimal point is in the wrong place.
- Our estimate caught the mistake. That is exactly why we estimate.

DO:
- Click to reveal.
- Celebrate the estimate catching the error.

TEACHER NOTES:
This is the headline of the lesson: an estimate is a powerful check. The real answer should be about 18 dollars.

WATCH FOR:
- Students who explain the error - ready for independent work.
- Students who accepted 1 dollar 80 - revisit estimating as a check.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 warms up by rounding to friendly numbers.
- Section 2 estimates fractions, decimals and percentages of amounts.
- Section 3 is detective work: is the answer reasonable? If you finish, try the Year 8 extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for "round to friendly, then estimate".
- Cold call one or two students to justify a reasonable-or-not decision.

TEACHER NOTES:
Different numbers from the We Do, same moves: round to a friendly fraction or percentage, estimate, then use the estimate to check.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 gives the friendly number to round to. Students just round, then estimate. This is also the rebuild for any returning student.
- Extra Notes: Sit with these students and round the first number together.
EXTENDING PROMPT:
- Task: The detective questions in Section 3, plus the Year 8 extension on percentage increase, decrease and working backwards.
- Extra Notes: Push students to explain how far off an unreasonable answer is and why.

WATCH FOR:
- Students who estimate and check fluently - secure.
- Students who try to be exact - re-anchor on "close and fast".

[Stage 4: You Do | Element: Mastery and application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- A 60 dollar jacket has 25 percent off.
- Estimate the discount. Then decide: is a 15 dollar discount reasonable?

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - estimate a percentage of an amount and judge reasonableness. 25% is one quarter, 60 divided by 4 is 15 dollars, so a 15 dollar discount is reasonable. The SC target is SC2.

WATCH FOR:
- Students who estimate $15 and say reasonable - secure.
- Students who cannot judge reasonableness - revisit estimate-as-check next unit.

[Stage 5: Exit Ticket | Element: Mastery and application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria one last time for the unit.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: how can an estimate help you check an answer?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is "estimate first, then use it to check the answer is reasonable". This completes the unit: students can now find, connect, apply and check fractions, decimals and percentages.

WATCH FOR:
- Strong thumbs up across all three - the unit has landed.
- Sideways or down - note who needs a small group next unit.

[General: Closing | Element: Retention and recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 4: Estimating to check answers",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - mixed FDP from the unit
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Unit recap",
      [
        "Find 1/3 of 60.",
        "Write 0.25 as a percentage.",
        "Find 10% of $80.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "20          25%          $8", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal column addition (ragged places)
  withReveal(
    () => fluencySlide(pres, "Fluency: Adding decimals",
      ["1.4 + 0.35", "2.5 + 0.75", "3.6 + 0.4"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1.75        3.25        4.0", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - round to a friendly number (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "About how much?",
    [
      "A shirt is $19.95, 50% off.",
      "We don't need the exact answer yet.",
      "",
      "$19.95 is about $20.  Half off = about $10.",
      "Round to a friendly number, then estimate.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.85, { strip: C.ACCENT });
      slide.addText("Round, then estimate", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      stepPills(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.55, lg.rightW - 0.60, [
        { text: "$19.95  ~  $20", color: C.PRIMARY },
        { text: "50% off = half = ~ $10", color: C.SECONDARY },
        { text: "Pay about $10", color: C.SUCCESS },
      ], { h: 0.56, gap: 0.14, fontSize: 16 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to estimate with fractions, decimals and percentages, and to use estimates to check if an answer is reasonable.",
    [
      "I can round a number or percentage to a friendly one to estimate.",
      "I can estimate a fraction, decimal or percentage of an amount.",
      "I can use an estimate to decide whether a calculated answer is reasonable.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Estimate = a sensible, close answer (not exact)",
      "About / approximately (~) = near to, roughly",
      "Reasonable = it makes sense, close to your estimate",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.SECONDARY });
      slide.addText("Friendly numbers", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      stepPills(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.55, lg.rightW - 0.60, [
        { text: "0.3 is about 1/3", color: C.PRIMARY },
        { text: "48% is about 50%", color: C.SECONDARY },
      ], { h: 0.56, gap: 0.14, fontSize: 16 });
    }
  );

  // Slide 10: I Do #1 - estimate 0.3 of 180
  workedExSlide(pres, 2, "I Do", "Estimate 0.3 of 180",
    [
      "0.3 is very close to 1/3.",
      "So 0.3 of 180 is about 1/3 of 180.",
      "1/3 of 180 = 180 ÷ 3 = 60.",
      "",
      "Estimate: about 60.",
      "(Exact 0.3 x 180 = 54 - close.)",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.05, { strip: C.PRIMARY });
      slide.addText("0.3 is close to 1/3", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      // 0.3 as a tenths strip
      slide.addText("0.3 = three tenths", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.46, w: lg.rightW - 0.4, h: 0.24,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });
      addTenthsStrip(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.74, 3.0, 3,
        { fillColor: C.SECONDARY });
      // 1/3 as a 3-part bar
      slide.addText("1/3 = one of three parts", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.30, w: lg.rightW - 0.4, h: 0.24,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });
      partBar(slide, lg.rightX + 0.25, lg.panelTopPadded + 1.58, 3.0, 0.55, 3, 1,
        { fill: C.SECONDARY });
      addTextOnShape(slide, "About 1/3 of 180 = 60", {
        x: lg.rightX + 0.35, y: lg.panelTopPadded + 2.32, w: lg.rightW - 0.70, h: 0.54, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 11: I Do #2 - estimate to check a discount
  workedExSlide(pres, 2, "I Do", "Estimate to check a saving",
    [
      "Toy $39.95, 25% off. Saving says $9.99.",
      "Round $39.95 to $40.",
      "25% = 1/4, so about $40 ÷ 4 = $10 off.",
      "",
      "$9.99 is about $10 - that is reasonable.",
      "The estimate is our check.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Estimate, then check", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      stepPills(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.50, lg.rightW - 0.60, [
        { text: "$39.95 ~ $40", color: C.PRIMARY },
        { text: "25% = 1/4 ~ $10 off", color: C.SECONDARY },
      ], { h: 0.52, gap: 0.12, fontSize: 15.5 });
      addTextOnShape(slide, "Saving $9.99 ~ $10  REASONABLE", {
        x: lg.rightX + 0.30, y: lg.panelTopPadded + 1.90, w: lg.rightW - 0.60, h: 0.80, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: CFU + reveal - estimate 48% of 200
  withReveal(
    () => cfuSlide(pres, "CFU", "Estimate 48% of 200",
      { technique: "Show Me Boards",
        question: "On your whiteboard:\n\nRound 48% to a friendly percentage, then estimate 48% of 200." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "48% is about 50%, and half of 200 is about 100", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - estimate 9% of $80
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Estimate 9% of $80",
      [
        "With your partner.",
        "",
        "1.  9% is close to which friendly %?",
        "2.  Find that friendly % of $80.",
        "3.  Read your estimate.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.SECONDARY });
        slide.addText("9% is about 10%", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.40,
          fontSize: 21, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        slide.addText("10% means divide by 10.", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.58, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
          align: "center", margin: 0,
        });
        addTenthsStrip(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.98, 3.0, 1,
          { fillColor: C.SECONDARY });
        slide.addText("Round the percentage to friendly first.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.70, w: lg.rightW - 0.30, h: 0.40,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "9% of $80 ~ 10% of $80 = $8   (exact $7.20)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - is $1.80 reasonable?
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Detective: is it reasonable?",
      [
        "With your partner.",
        "",
        "A calculator says:",
        "30% of $60 = $1.80.",
        "Estimate first. Is that reasonable?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.SECONDARY });
        slide.addText("Estimate to check", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        stepPills(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.55, lg.rightW - 0.60, [
          { text: "10% of $60 = $6", color: C.PRIMARY },
          { text: "30% is about $18", color: C.SECONDARY },
        ], { h: 0.56, gap: 0.14, fontSize: 16 });
        slide.addText("Does $1.80 sit near $18?", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.90, w: lg.rightW - 0.30, h: 0.40,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "Not reasonable! 30% is about $18, not $1.80 (decimal slip)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 16, fontFace: FONT_H, color: getContrastColor(C.ALERT), bold: true });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // Slide 18: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: estimate & check", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "round to a friendly number.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "estimate the answer.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "is the given answer reasonable?", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.4;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember", {
      x: 0.7, y: panelY + 0.13, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "Round to friendly, estimate, then check it is reasonable.", {
      x: 1.2, y: panelY + 0.52, w: 7.6, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Friendly-rounding reference pills
    const refs = [["0.3 ~ 1/3", C.SUCCESS], ["48% ~ 50%", C.PRIMARY], ["9% ~ 10%", C.SECONDARY]];
    const rw = 2.5, rgap = 0.2, rx0 = (10 - (rw * 3 + rgap * 2)) / 2;
    refs.forEach((r, i) => {
      addTextOnShape(s, r[0], {
        x: rx0 + i * (rw + rgap), y: panelY + 1.20, w: rw, h: 0.6, rectRadius: 0.08,
        fill: { color: r[1] },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "A $60 jacket has 25% off. Estimate the discount.",
      "Is a $15 discount reasonable? Explain.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how can an estimate help you check an answer?",
      scItems: [
        "I can round a number or percentage to a friendly one to estimate.",
        "I can estimate a fraction, decimal or percentage of an amount.",
        "I can use an estimate to decide whether a calculated answer is reasonable.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FDP_Lesson4_Estimating_With_FDP.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Estimate with fractions, decimals and percentages, and check if answers are reasonable.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Round to a friendly number or percentage, then estimate. Useful swaps: 0.3 is about 1/3, 0.5 is exactly 1/2, 48% is about 50%, 9% is about 10%. Use your estimate to check an answer is reasonable.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Estimate 0.49 of 60. 0.49 is about 1/2. Half of 60 is 30. So 0.49 of 60 is about 30.",
      y);

    y = addSectionHeading(doc, "Section 1 - Round to the friendly number (started for you)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  0.3 is about  1/3        b)  0.51 is about  ______ (fraction)", y);
    y = addWriteLine(doc, "c)  52% is about  ______%        d)  9% is about  ______%", y);
    y = addWriteLine(doc, "e)  $29.95 is about  $______        f)  24% is about  ______%", y);

    y = addSectionHeading(doc, "Section 2 - Estimate the amount", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  0.3 of 90 is about  ______", y);
    y = addWriteLine(doc, "b)  52% of 80 is about  ______", y);
    y = addWriteLine(doc, "c)  9% of $50 is about  ______", y);
    y = addWriteLine(doc, "d)  48% of 200 is about  ______", y);

    y = addSectionHeading(doc, "Section 3 - Detective: is it reasonable?", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  A calculator says 50% of $48 = $2.40. Reasonable?  ______  Why?", y);
    y = addWriteLine(doc, "b)  A sign says 25% off $40 saves $10. Reasonable?  ______  Why?", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addSectionHeading(doc, "Challenge (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "A jacket is $79.95 with 30% off. Estimate the sale price. Then explain how you would know if $55 was a reasonable sale price.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Estimating with FDP | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 4 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Round to the friendly number", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/3        b)  1/2        c)  50%        d)  10%        e)  $30        f)  25%", y);

    y = addSectionHeading(doc, "Section 2 - Estimate the amount", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  about 30 (1/3 of 90)        b)  about 40 (1/2 of 80)        c)  about $5 (10% of $50)        d)  about 100 (1/2 of 200)", y);

    y = addSectionHeading(doc, "Section 3 - Detective", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Not reasonable. 50% of $48 is half = $24, not $2.40. The decimal point is ten times out.", y);
    y = addBodyText(doc, "b)  Reasonable. 25% of $40 is one quarter = $10. The saving matches.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "$79.95 is about $80. 30% off: 10% = $8, 30% = $24 off, so sale price is about $56. $55 is close to $56, so it is a reasonable sale price.", y);

    y = addTipBox(doc,
      "Watch for: students who try to be exact instead of estimating; students who trust a calculator answer without checking; students who round the wrong way (e.g. 9% to 1% instead of 10%).",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension (unit-wide stretch resource)
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Percentage increase, decrease and working backwards.",
      color: C.SECONDARY,
      lessonInfo: `Year 8 challenge | extends Year 6 VC2M6N07 / VC2M6N08`,
    });
    y = addTipBox(doc,
      "You already find a percentage of an amount. Year 8 goes further: increasing and decreasing by a percentage, finding the original amount after a change, and expressing one number as a percentage of another.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example - working backwards", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "After a 20% discount, a game costs $40. What was the original price? A 20% discount means you pay 80% of the original. So 80% = $40, which means 10% = $5, and 100% = $50. The original price was $50.",
      y);

    y = addSectionHeading(doc, "Section 1 - Increase and decrease", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  Increase $40 by 15%. New amount = _______", y);
    y = addWriteLine(doc, "b)  Decrease 120 kg by 30%. New amount = _______", y);
    y = addWriteLine(doc, "c)  A $250 phone rises 8%. New price = _______", y);

    y = addSectionHeading(doc, "Section 2 - Working backwards", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  After 25% off, a coat costs $60. Original price = _______", y);
    y = addWriteLine(doc, "b)  After a 10% rise, a fare is $22. Original fare = _______", y);
    y = addWriteLine(doc, "c)  60 is 75% of a number. The number = _______", y);

    y = addSectionHeading(doc, "Section 3 - One amount as a percentage of another", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  18 out of 24 as a percentage = _______", y);
    y = addWriteLine(doc, "b)  A team won 13 of 20 games. Win percentage = _______", y);

    y = addSectionHeading(doc, "Section 4 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "A price goes up 10%, then down 10%. Is it back to the start? Explain.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers:  S1  a) $46  b) 84 kg  c) $270.   S2  a) $80 (75% = 60, 100% = 80)  b) $20 (110% = 22, 100% = 20)  c) 80.   S3  a) 75%  b) 65%.   S4  No. Down 10% is taken from the higher amount, so it ends slightly below the start (e.g. $100 -> $110 -> $99).",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Year 8 Extension | Percentages | extends Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
