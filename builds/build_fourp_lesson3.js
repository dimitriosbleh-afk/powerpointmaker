"use strict";

// Four Processes (Year 6 Numeracy) - Lesson 3 of 5.
// The four processes with decimals. Multiply/divide a decimal by powers of 10
// (digits move, point stays). Estimate FIRST by rounding to friendly numbers,
// then multiply a decimal by a two-digit number and check the answer is close
// to the estimate. VC2M6N06.
// Daily Review: Fractions & Decimals (prior). Fluency: decimal x 1-digit algorithm.
// Unit variant fixed (variant 3, Ocean Logic) across all 5 lessons for cohesion.
// Catch-up: the launch rebuilds estimation from whole numbers everyone can do,
// and worksheet Section 1 rebuilds the powers-of-10 move. No session assumes the
// one before; the decimal place value anchor is re-taught here for returners.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 3); // variant 3 (Ocean Logic), fixed for the unit
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide, dailyReviewSlide,
  addPlaceValueChart, addDecimalDot,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 5;
const UNIT_TITLE = "Four Processes";
const FOOTER = `Four Processes | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FourP_Lesson3_Decimal_Operations";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Multiply and divide decimals by powers of 10, then estimate and calculate.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 3 practice sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Multiplying decimals by decimals and estimating with significant figures - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Helpers -----------------------------------------------------------------
const PV_HTOTH = ["Hundreds", "Tens", "Ones", "Tenths"];

function pvChart(slide, x, y, headers, values, opts) {
  const o = opts || {};
  const chart = addPlaceValueChart(slide, x, y, headers, values, {
    totalW: o.totalW || 3.6,
    valH: o.valH || 0.50,
    hdrH: o.hdrH || 0.34,
    headerColor: o.headerColor || C.PRIMARY,
  });
  if (o.dotAfter != null) {
    addDecimalDot(slide, chart, o.dotAfter, { position: "baseline" });
  }
  return chart;
}

// Estimate -> Calculate -> Check, three labelled rows. The shared visual for
// every estimation slide in this lesson so students read it the same way.
function addEstimateRows(slide, lg, top, rows) {
  rows.forEach((r, i) => {
    const ry = top + i * 0.82;
    slide.addText(r.label, {
      x: lg.rightX + 0.2, y: ry, w: lg.rightW - 0.4, h: 0.24,
      fontSize: 12, fontFace: FONT_H, color: C.MUTED, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(slide, r.expr, {
      x: lg.rightX + 0.30, y: ry + 0.26, w: lg.rightW - 0.60, h: 0.48, rectRadius: 0.07,
      fill: { color: r.color },
    }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });
  });
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to Four Processes. We have linked multiplication and division, and used the properties to calculate smartly.
- Today we put the four processes to work with decimals.
- The big habit today is estimate first, then calculate, then check the answer is close. A quick estimate catches silly mistakes.

DO:
- Have whiteboards, markers and the printed place value chart ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 3 of 5. The hero habit is estimate, calculate, check. The powers-of-10 move is re-taught today so returning students can rejoin.

WATCH FOR:
- Students who look unsure - that is expected. Reassure them: if this feels new, that is okay, we build it together.

[General: Title | Element: Planning and Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end. The Year 8 extension is there for anyone who is flying.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards, markers and the printed place value chart ready.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student practice sheet, an answer key, and a Year 8 extension. Most teaching happens on whiteboards with the place value chart.

CATCH-UP NOTE:
A student who missed earlier sessions can still access today. The launch rebuilds estimation from whole numbers, and Section 1 of the worksheet rebuilds the powers-of-10 move. A returner only needs the printed chart and one minute with you to start. No session in this unit assumes the student saw the one before.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are warming up our fractions and decimals from earlier work.
- Read each one carefully and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for how students convert and order decimals.

TEACHER NOTES:
Daily Review is prior learning, not today's new content. Fractions and decimals keep ticking over.

WATCH FOR:
- Students who order decimals by place value - secure.
- Students who think 0.85 is smaller than 0.8 - reteach lining up tenths in the reveal.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 7/10 is 0.7.
- Smallest to largest: 0.08, then 0.8, then 0.85.
- 0.25 is the same as 1/4.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The ordering item rewards lining up place value. 0.08 is eight hundredths, smaller than eight tenths.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students who order by digit count - small group focus.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Multiplying a decimal by a single digit using the algorithm.
- Set each one out vertically on your whiteboard and multiply.
- Keep your columns lined up and remember where the decimal point sits.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for clean column alignment and the point kept in line.

TEACHER NOTES:
Fluency this unit is the multiplication algorithm with decimals. Strong single-digit fluency makes today's two-digit multiplication much easier.

WATCH FOR:
- Students who keep the decimal point lined up - secure.
- Students who lose the point - prompt: one decimal place in, one decimal place out.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 6.4 times 5 is 32.0, which we write as 32.
- 3.7 times 8 is 29.6.
- 5.6 times 4 is 22.4.
- Tick and fix.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
6.4 times 5 lands on 32.0; accept 32. Keep it brisk.

WATCH FOR:
- Students who self-correct - secure.
- Students whose point drifts - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember rounding to make numbers friendly. Watch this.
- I do not know 4.2 times 38 yet, but I can estimate. Round 4.2 to 4 and 38 to 40.
- 4 times 40 is 160. So the answer should be ABOUT 160.
- That estimate is like a target. When I do the real calculation, my answer should land near 160. If it lands at 16 or 1600, I know I have slipped a decimal place.

DO:
- Show the rounding of each number.
- Have students chorus 'round, then multiply the friendly numbers'.
- Bridge: 'estimate first, so we can check our real answer'.

TEACHER NOTES:
This launch starts from rounding and whole-number multiplication everyone can access, then sets up the estimate-calculate-check habit. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who round and estimate confidently - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to multiply and divide decimals, and to estimate first so we can check our answer is reasonable.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Hold up the printed place value chart.

TEACHER NOTES:
The first criterion is reachable for everyone - move the digits for powers of 10. The second is the core target the exit ticket checks - estimating. The third stretches to a full two-digit multiplication checked against the estimate.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- Estimate means a close, quick answer using friendly numbers.
- Round means change a number to a nearby friendly one, like a multiple of 10 or 100.
- Powers of 10 are 10, 100 and 1000. Multiplying or dividing by them moves the digits.

DO:
- Point to each word as you say it.
- Have students say 'estimate first, then check' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. The phrase 'estimate first, then check' is the anchor for today.

WATCH FOR:
- Students who can round to a friendly number - secure.
- Students who think estimate means guess - clarify: a smart estimate uses rounding.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at this one together. Multiplying or dividing by a power of 10 just moves the digits.
- 3.6 times 100. Every digit moves two places to the LEFT. The 3 goes to hundreds, the 6 goes to tens, and a zero holds the ones place. So 3.6 times 100 is 360.
- Now divide. 36 divided by 10 moves every digit one place to the RIGHT. The 3 goes to ones, the 6 goes to tenths. So 36 divided by 10 is 3.6.
- The important thing: the digits move, the point stays still. We do not move the point.

DO:
- Fill the after chart cell by cell, tracing the move with your finger.
- Say 'times moves left, divide moves right, the point stays'.
- Circle the placeholder zero in 360.

TEACHER NOTES:
This re-teaches the powers-of-10 move so any returning student can rejoin, and it is the tool we use to scale numbers today. Resist 'move the decimal point' - describe the digits moving.

MISCONCEPTIONS:
- Misconception: students think multiplying by 100 means 'add two zeros'.
  Why: it works for whole numbers.
  Impact: they write 3.6 times 100 as 3.600, which is still 3.6.
  Quick correction: show each digit moving two places left on the chart.

WATCH FOR:
- Students who move the digits correctly - secure.
- Students who add zeros to the end - re-point to the chart move.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now the main event: multiplying a decimal by a two-digit number, and checking it.
- First I ESTIMATE. 4.2 times 38: round to 4 times 40, which is 160. My answer should be near 160.
- Then I CALCULATE. I multiply as if there is no point: 42 times 38 is 1596. There is one decimal place in 4.2, so there is one decimal place in my answer: 159.6.
- Then I CHECK. Is 159.6 close to my estimate of 160? Yes. So I am confident.
- Estimate, calculate, check. The estimate is what tells me the point is in the right spot.

DO:
- Reveal the three rows one at a time: estimate, calculate, check.
- Point to the single decimal place in 4.2 and then in 159.6.
- Have students chorus 'estimate, calculate, check'.

TEACHER NOTES:
The estimate does real work here: it tells students whether the decimal point is in a sensible place. 159.6 near 160 confirms it; 15.96 or 1596 would not.

MISCONCEPTIONS:
- Misconception: students misplace the decimal point in the answer.
  Why: they multiply 42 x 38 then guess where the point goes.
  Impact: answers like 15.96 or 1596 instead of 159.6.
  Quick correction: count decimal places in the question (one), so one in the answer; the estimate of 160 confirms it.

WATCH FOR:
- Students whose answer lands near the estimate - secure.
- Students whose answer is ten times out - use the estimate to relocate the point.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards. Estimate 6.1 times 49.
- Round each number to something friendly, then multiply.
- You do not need the exact answer yet - just the estimate.

DO:
- Display the prompt.
- Give 45 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 6 times 50 = 300 (rounding 6.1 to 6 and 49 to 50).
PROCEED: If about 80 percent reach about 300, click to reveal and move to We Do.
PIVOT: Most likely misconception - students round 49 to 40 or keep 6.1 as it is.
- Reteach: round to the NEAREST friendly number. 49 is nearly 50, 6.1 is nearly 6. 6 times 50 is 300.
- Re-check: what is 49 rounded to the nearest ten?

TEACHER NOTES:
The trap is rounding the wrong way or not rounding at all. A good estimate uses the nearest friendly numbers.

WATCH FOR:
- Students who reach about 300 - secure.
- Students who round 49 down to 40 - prompt for the nearest ten.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. 5.2 times 39.
- First estimate together: round to friendly numbers.
- Then we will calculate and check it lands near the estimate.

DO:
- Display 5.2 x 39.
- Give 90 seconds for the estimate and the calculation.
- Listen for 'about 200' as the estimate.

TEACHER NOTES:
Same estimate-calculate-check move as the I Do with new numbers. Get the estimate first, then the calculation.

WATCH FOR:
- Pairs who estimate 5 times 40 = 200 - secure.
- Pairs who jump straight to calculating - prompt them to estimate first.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Estimate: 5 times 40 is 200.
- Calculate: 52 times 39 is 2028. One decimal place in 5.2, so one in the answer: 202.8.
- Check: 202.8 is close to 200. The point is in the right place.

DO:
- Click to reveal.
- Point to the estimate and the matching answer.

TEACHER NOTES:
Reveal restates the habit. 202.8 sits just above 200 because we rounded both numbers up a little.

WATCH FOR:
- Students whose answer lands near 200 - secure.
- Students who land on 20.28 or 2028 - use the estimate to relocate the point.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, this time a division in a real situation.
- A 2.5 metre ribbon is shared into 10 equal pieces. How long is each piece?
- Think: dividing by 10 moves the digits which way? And does your answer make sense?

DO:
- Display the ribbon problem.
- Give 75 seconds.
- Watch for students who multiply instead of divide.

TEACHER NOTES:
This brings division by a power of 10 into a context and asks for a sense check. 2.5 divided by 10 is 0.25 metres. Each piece is much smaller than the whole, which makes sense.

WATCH FOR:
- Students who get 0.25 m and say it makes sense - secure.
- Students who get 25 m - prompt: can a piece be longer than the whole ribbon?

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Dividing by 10 moves the digits one place to the right. 2.5 divided by 10 is 0.25.
- Each piece is 0.25 metres, which is 25 centimetres.
- Sense check: each piece is much shorter than the 2.5 metre ribbon. That is exactly what we expect.

DO:
- Click to reveal.
- Connect 0.25 m to 25 cm so the size feels real.

TEACHER NOTES:
The reasonableness check is the point: dividing makes each share smaller, so a piece must be less than the whole.

WATCH FOR:
- Students who check the size makes sense - ready for independent work.
- Students who give an impossible answer - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 warms up with multiplying and dividing by powers of 10.
- Section 2 is estimate, then calculate, then check.
- Section 3 asks you to spot an answer with the point in the wrong place. If you finish, try the extension box.

DO:
- Distribute the practice sheet.
- Circulate and listen for 'estimate, calculate, check'.
- Cold call one or two students to explain how an estimate caught a misplaced point.

TEACHER NOTES:
Different numbers from the We Do, same habit: estimate first, calculate, then check against the estimate.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 has the place value chart printed and the first move started. Move the digits, the point stays.
- Extra Notes: Sit with these students and move the first digit together. This is also the rebuild for any returning student.
EXTENDING PROMPT:
- Task: The extension box asks students to multiply a decimal by a three-digit number and to write a sentence explaining how their estimate proved the point was in the right place.
- Extra Notes: Students who are ready can move on to the Year 8 extension sheet on multiplying decimals by decimals.

WATCH FOR:
- Students who estimate before calculating - secure.
- Students who calculate without checking - prompt: does it land near your estimate?

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- Estimate 4.8 times 21 by rounding to friendly numbers.
- Then say whether 100.8 is a reasonable answer, and explain.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - estimate a decimal multiplication and judge reasonableness. Look for 5 times 20 = 100, and a yes because 100.8 is close to 100.

WATCH FOR:
- Students who estimate 100 and accept 100.8 - secure.
- Students who cannot judge reasonableness - revisit the estimate-check habit at the start of Lesson 4.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: how does estimating first help you check a decimal answer?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is estimate first, then check the real answer lands near the estimate. Students who do this are ready to find fractions and percentages of quantities next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on the core criterion - small group revision at the start of Lesson 4.

[General: Closing | Element: Retention and Recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 3: The four processes with decimals",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - fractions & decimals
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions & decimals",
      [
        "Write 7/10 as a decimal.",
        "Order from smallest: 0.8, 0.08, 0.85",
        "Write 0.25 as a fraction.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "0.7          0.08, 0.8, 0.85          1/4", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal x 1-digit
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal x single digit",
      ["6.4 x 5", "3.7 x 8", "5.6 x 4"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "32.0        29.6        22.4", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - estimate from whole numbers (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "Estimate first",
    [
      "I do not know 4.2 x 38 yet.",
      "Round: 4.2 -> 4 and 38 -> 40.",
      "4 x 40 = 160.",
      "",
      "So the answer is about 160.",
      "That is my target to check against.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.ACCENT });
      slide.addText("Round, then multiply", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      const rows = [["4.2  ->  4", C.PRIMARY], ["38  ->  40", C.PRIMARY]];
      const ry0 = lg.panelTopPadded + 0.55;
      rows.forEach((r, i) => {
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.7, y: ry0 + i * 0.56, w: lg.rightW - 1.4, h: 0.46, rectRadius: 0.07,
          fill: { color: r[1] },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
      addTextOnShape(slide, "4 x 40 = 160", {
        x: lg.rightX + 0.5, y: ry0 + 1.30, w: lg.rightW - 1.0, h: 0.52, rectRadius: 0.07,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to multiply and divide decimals, and to estimate first so we can check our answer is reasonable.",
    [
      "I can multiply or divide a decimal by 10, 100 or 1000.",
      "I can estimate a decimal multiplication by rounding to friendly numbers.",
      "I can multiply a decimal by a two-digit number and check it against my estimate.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Estimate = a close, quick answer",
      "Round = change to a nearby friendly number",
      "Powers of 10 = 10, 100, 1000 (digits move)",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
      slide.addText("Estimate first, then check", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.32,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [["Estimate", C.PRIMARY], ["Calculate", C.SECONDARY], ["Check", C.ACCENT]];
      const ry0 = lg.panelTopPadded + 0.58;
      rows.forEach((r, i) => {
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.8, y: ry0 + i * 0.56, w: lg.rightW - 1.6, h: 0.46, rectRadius: 0.07,
          fill: { color: r[1] },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slide 10: I Do #1 - powers of 10 (catch-up anchor)
  workedExSlide(pres, 2, "I Do", "Powers of 10: the digits move",
    [
      "3.6 x 100: digits move 2 places LEFT.",
      "3 -> hundreds, 6 -> tens, 0 holds ones.",
      "3.6 x 100 = 360.",
      "",
      "36 ÷ 10: digits move 1 place RIGHT.",
      "36 ÷ 10 = 3.6.",
      "",
      "The digits move. The point stays.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.PRIMARY });
      slide.addText("3.6 x 100", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      pvChart(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.42, PV_HTOTH,
        ["", "", "3", "6"], { totalW: lg.rightW - 0.60, valH: 0.44, hdrH: 0.30, dotAfter: 2 });
      slide.addText("becomes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.20, w: lg.rightW - 0.4, h: 0.24,
        fontSize: 11.5, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      pvChart(slide, lg.rightX + 0.30, lg.panelTopPadded + 1.46, PV_HTOTH,
        ["3", "6", "0", ""], { totalW: lg.rightW - 0.60, valH: 0.44, hdrH: 0.30, dotAfter: 2 });
      addTextOnShape(slide, "3.6 x 100 = 360       36 ÷ 10 = 3.6", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.58, w: lg.rightW - 0.40, h: 0.52, rectRadius: 0.07,
        fill: { color: C.SECONDARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 11: I Do #2 - estimate, calculate, check (4.2 x 38)
  workedExSlide(pres, 2, "I Do", "Estimate, calculate, check: 4.2 x 38",
    [
      "Estimate: round to 4 x 40 = 160.",
      "Calculate: 42 x 38 = 1596.",
      "One decimal place in, one out: 159.6.",
      "Check: 159.6 is close to 160.",
      "",
      "The estimate shows the point is right.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.05, { strip: C.PRIMARY });
      addEstimateRows(slide, lg, lg.panelTopPadded + 0.18, [
        { label: "ESTIMATE", expr: "4 x 40 = 160", color: C.PRIMARY },
        { label: "CALCULATE", expr: "42 x 38 = 1596 -> 159.6", color: C.SECONDARY },
        { label: "CHECK", expr: "159.6 is close to 160", color: C.SUCCESS },
      ]);
    }
  );

  // Slides 12-13: CFU + reveal - estimate 6.1 x 49
  withReveal(
    () => cfuSlide(pres, "CFU", "Estimate 6.1 x 49",
      { technique: "Show Me Boards",
        question: "On your whiteboard: round each number to a friendly one and estimate 6.1 x 49.\n\nYou do not need the exact answer." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "6 x 50 = 300   (round 6.1 to 6 and 49 to 50)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - estimate then calculate 5.2 x 39
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Estimate then calculate: 5.2 x 39",
      [
        "With your partner.",
        "",
        "1.  Estimate: round to friendly numbers.",
        "2.  Calculate: multiply, then place the point.",
        "3.  Check: does it land near your estimate?",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.SECONDARY });
        slide.addText("5.2 x 39", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.42,
          fontSize: 28, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        slide.addText("Estimate box:", {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.70, w: lg.rightW - 0.6, h: 0.26,
          fontSize: 12.5, fontFace: FONT_H, color: C.MUTED, bold: true, margin: 0,
        });
        slide.addShape("roundRect", {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.98, w: lg.rightW - 0.6, h: 0.50,
          fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
        });
        slide.addText("Calculate box:", {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.58, w: lg.rightW - 0.6, h: 0.26,
          fontSize: 12.5, fontFace: FONT_H, color: C.MUTED, bold: true, margin: 0,
        });
        slide.addShape("roundRect", {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.86, w: lg.rightW - 0.6, h: 0.50,
          fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
        });
      }),
    (slide) => {
      addTextOnShape(slide, "Estimate 5 x 40 = 200.  52 x 39 = 2028 -> 202.8.  Close to 200.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - ribbon division (2.5 ÷ 10)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "A real situation: sharing a ribbon",
      [
        "With your partner.",
        "",
        "A 2.5 m ribbon is cut into 10 equal pieces.",
        "How long is each piece?",
        "Does your answer make sense?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.SECONDARY });
        slide.addText("2.5 m shared into 10", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        // A simple ribbon split into 10 segments.
        const rx = lg.rightX + 0.30, rw = lg.rightW - 0.60, ry = lg.panelTopPadded + 0.62, seg = rw / 10;
        for (let i = 0; i < 10; i++) {
          slide.addShape("rect", {
            x: rx + i * seg, y: ry, w: seg, h: 0.5,
            fill: { color: i % 2 === 0 ? C.SECONDARY : C.ACCENT }, line: { color: C.WHITE, width: 1 },
          });
        }
        slide.addText("10 equal pieces", {
          x: lg.rightX + 0.2, y: ry + 0.58, w: lg.rightW - 0.4, h: 0.26,
          fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
          align: "center", margin: 0,
        });
        addTextOnShape(slide, "2.5 ÷ 10 = ?", {
          x: lg.rightX + 0.85, y: ry + 0.94, w: lg.rightW - 1.7, h: 0.50, rectRadius: 0.07,
          fill: { color: C.PRIMARY },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      }),
    (slide) => {
      addTextOnShape(slide, "2.5 ÷ 10 = 0.25 m (25 cm) each. Smaller than the whole - it makes sense.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "multiply and divide by powers of 10.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "estimate, calculate, check.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "spot the misplaced decimal point.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Estimate first. Then calculate. Then check it lands near your estimate.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    pvChart(s, 2.5, panelY + 1.18, PV_HTOTH, ["", "", "", ""],
      { totalW: 5.0, valH: 0.55, hdrH: 0.38, dotAfter: 2 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Estimate 4.8 x 21 by rounding to friendly numbers.",
      "Is 100.8 a reasonable answer? Explain.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how does estimating first help you check a decimal answer?",
      scItems: [
        "I can multiply or divide a decimal by 10, 100 or 1000.",
        "I can estimate a decimal multiplication by rounding to friendly numbers.",
        "I can multiply a decimal by a two-digit number and check it against my estimate.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FourP_Lesson3_Decimal_Operations.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Multiply and divide decimals by powers of 10, then estimate and calculate.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Powers of 10 move the digits: x10 one place left, x100 two places left; div 10 one place right. The point stays still. For two-digit multiplication: estimate by rounding, calculate, then check the answer lands near your estimate.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "3.4 x 29: estimate 3 x 30 = 90. Calculate 34 x 29 = 986, one decimal place: 98.6. Check: 98.6 is close to 90. The point is in the right place.",
      y);

    y = addSectionHeading(doc, "Section 1 - Powers of 10 (started for you)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  4.7 x 100 = 470        (each digit moves 2 places left)", y);
    y = addWriteLine(doc, "b)  2.5 x 10 = _______", y);
    y = addWriteLine(doc, "c)  63 ÷ 10 = _______", y);
    y = addWriteLine(doc, "d)  8.2 x 1000 = _______", y);

    y = addSectionHeading(doc, "Section 2 - Estimate, calculate, check", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3.6 x 42.   Estimate: ______   Calculate: ______   Check: ______", y);
    y = addWriteLine(doc, "b)  5.1 x 19.   Estimate: ______   Calculate: ______   Check: ______", y);
    y = addWriteLine(doc, "c)  2.4 x 48.   Estimate: ______   Calculate: ______   Check: ______", y);

    y = addSectionHeading(doc, "Section 3 - Spot the misplaced point", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "A student says 6.3 x 41 = 25.83. Estimate to show this is wrong, then give the correct answer.", y);
    y = addWriteLine(doc, "Estimate: ____________   Correct answer: ____________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "Multiply by a three-digit number: 3.2 x 125 = ______   (estimate first)", y);
    y = addWriteLine(doc, "Explain how your estimate proved the decimal point was in the right place.", y);

    addPdfFooter(doc, `Lesson ${SESSION} | The four processes with decimals | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 3 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Powers of 10", y, { color: C.PRIMARY });
    y = addBodyText(doc, "b)  25        c)  6.3        d)  8200", y);

    y = addSectionHeading(doc, "Section 2 - Estimate, calculate, check", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Estimate 4 x 40 = 160; 36 x 42 = 1512 -> 151.2; close to 160.", y);
    y = addBodyText(doc, "b)  Estimate 5 x 20 = 100; 51 x 19 = 969 -> 96.9; close to 100.", y);
    y = addBodyText(doc, "c)  Estimate 2 x 50 = 100; 24 x 48 = 1152 -> 115.2; close to 100.", y);

    y = addSectionHeading(doc, "Section 3 - Spot the misplaced point", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Estimate 6 x 40 = 240, so 25.83 is far too small. 63 x 41 = 2583 -> 258.3. The correct answer is 258.3 (the point was two places out).", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Estimate 3 x 125 = 375; 32 x 125 = 4000 -> 400.0 = 400. Close to 375. The estimate confirms the point is after the hundreds.", y);

    y = addTipBox(doc,
      "Watch for: students who 'add zeros' for powers of 10; students who skip the estimate; students whose answer is ten times too big or small because the point is misplaced - the estimate is the fix.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Multiplying decimals by decimals.",
      color: C.SECONDARY,
      lessonInfo: `Lesson ${SESSION} | Year 8 challenge | extends Year 6 VC2M6N06`,
    });
    y = addTipBox(doc,
      "When you multiply two decimals, count the decimal places in BOTH numbers and add them. Estimate first to check the point ends up in a sensible place.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "3.2 x 1.5: estimate 3 x 2 = 6. Multiply 32 x 15 = 480. One decimal place plus one decimal place is two: 4.80. So 3.2 x 1.5 = 4.8, close to 6.",
      y);

    y = addSectionHeading(doc, "Section 1 - Decimal times decimal", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2.4 x 1.5 = ______        b)  6.3 x 0.4 = ______        c)  0.7 x 0.8 = ______", y);

    y = addSectionHeading(doc, "Section 2 - Estimate to place the point", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  4.6 x 2.5.   Estimate: ______   Answer: ______", y);
    y = addWriteLine(doc, "b)  12.5 x 0.8.  Estimate: ______   Answer: ______", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Why is 0.7 x 0.8 SMALLER than both 0.7 and 0.8? Explain using the meaning of multiplying by a number less than 1.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) 3.6  b) 2.52  c) 0.56.   S2  a) estimate 5 x 2.5 ~ 12, answer 11.5  b) estimate 12 x 1 = 12, answer 10.   S3  Multiplying by a number less than 1 means taking a part of it, so the result is smaller than what you started with.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Year 8 Extension | Decimal x decimal`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
