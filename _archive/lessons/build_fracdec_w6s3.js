"use strict";

// Fractions & Decimals - Week 6 (Year 6 Numeracy) - Session 3: Adding & subtracting decimals with estimation
// VC2M5N04. Estimate by rounding, then add/subtract decimals using place value (line up the points).
// Daily Review: Metric measurement & unit conversion. Fluency: subtraction algorithm.
// Anchor models: vertical place-value algorithm + number line. Variant weekToVariant(6).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(6));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, dailyReviewSlide, fluencySlide,
  addStageBadge, addRevealAnswerBar,
  addNumberLine,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 4;
const WEEK = 6;
const UNIT_TITLE = "Fractions & Decimals";
const FOOTER = `Fractions & Decimals | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracDec_W6S3_Add_Subtract_Decimals";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Adding and Subtracting Decimals Practice",
  "Estimate, then add and subtract decimals by lining up the points. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Adding and Subtracting Decimals Answer Key",
  "Worked answers for the adding and subtracting decimals practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back. We have been working with fractions and percentages.
- Today we switch to decimals - estimating first, then adding and subtracting carefully.

DO:
- Have whiteboards and markers ready.
- Settle the class before you click on.

TEACHER NOTES:
Week 6 Session 3 of 4. The vertical place-value algorithm is the anchor. A returning student needs the Week 6 Catch-Up Card and a quick reminder to line up the points.

WATCH FOR:
- Students who line up the last digit instead of the decimal point - the main error today.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- Today's materials are the practice sheet only.
- The practice sheet has an easier start and a challenge built in.

DO:
- Print one practice sheet per student and keep the answer key.
- Have whiteboards ready.

TEACHER NOTES:
One student practice sheet plus its answer key. Most teaching is on whiteboards.

CATCH-UP NOTE:
A returner can join here. The launch re-shows estimating by rounding, and the Week 6 Catch-Up Card lists the line-up-the-points method with a worked example.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Measurement and unit conversion, not today's decimal sums.
- Answer each on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and check the conversion direction.

TEACHER NOTES:
Measurement retrieval, which is also a decimal context. Good warm-up for today.

WATCH FOR:
- Students who convert correctly - secure.
- Students who go the wrong way - prompt: bigger or smaller number?

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 1.25 metres is 125 centimetres.
- 450 grams is 0.45 kilograms.
- 2 litres take away 750 millilitres is 1250 millilitres.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The last item needs 2 L thought of as 2000 mL before subtracting. A nice bridge to lining up units today.

WATCH FOR:
- Students who self-correct - secure.
- Students who subtract without converting - reteach common units first.

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Decimal subtraction, vertical.
- Line up the decimal points, subtract, regroup if needed.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up points.

TEACHER NOTES:
Subtraction fluency continues this week, and it directly supports today's lesson.

WATCH FOR:
- Students who regroup correctly - secure.
- Students who flip a column - coach borrowing.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 9.4 minus 2.7 is 6.7.
- 15.2 minus 8.6 is 6.6.
- 7.05 minus 3.4 is 3.65.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
7.05 minus 3.4 needs 3.4 read as 3.40 to line up. Coach the placeholder zero.

WATCH FOR:
- Students who self-correct - secure.
- Students who misalign 3.4 - small group focus.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Before we calculate, let us estimate. About how much is 3.7 plus 2.2?
- I round each number to the nearest whole. 3.7 rounds to 4, and 2.2 rounds to 2.
- So my estimate is about 4 plus 2, which is 6. The real answer should be near 6.
- Estimating first tells us if our exact answer is sensible. That is today's first habit.

DO:
- Round each number out loud.
- Take the estimate from the class before any exact work.

TEACHER NOTES:
This launch builds the estimate-first habit. The estimate is a safety check for the exact answer.

WATCH FOR:
- Students who round then add - strong start.
- Students who try to calculate exactly straight away - redirect to rounding first.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to estimate, then add and subtract decimals using place value.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is estimating by rounding. Second is adding by lining up the points, the core target. Third is subtracting with regrouping.

WATCH FOR:
- Students who can say line up the points - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- Two words for today.
- To estimate is to find an answer that is close enough, quickly, usually by rounding.
- To round a decimal to the nearest whole, look at the tenths: 5 or more rounds up, less than 5 rounds down.

DO:
- Round one or two decimals to the nearest whole together.
- Say estimate means close enough.

TEACHER NOTES:
Vocabulary after the learning intention. Estimating supports the exact work; it does not replace it.

WATCH FOR:
- Students who round using the tenths - secure.
- Students who round 3.7 down to 3 - reteach the rounding digit.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us add 4.6 and 3.75. First, estimate: 4.6 rounds to 5, 3.75 rounds to 4, so about 9.
- Now the exact answer. I line up the decimal points, one under the other.
- I write 4.6 as 4.60 so both have two decimal places. Then I add from the right: 0 and 5 is 5, 6 and 7 is 13 carry 1, 4 and 3 and the carried 1 is 8.
- The answer is 8.35. That is close to my estimate of 9, so it is sensible.

DO:
- Write the estimate first.
- Line up the points, add the placeholder zero, add from the right.
- Check the answer against the estimate.

TEACHER NOTES:
The core move: estimate, line up the points, add a placeholder zero so places match, then add. Compare to the estimate.

MISCONCEPTIONS:
- Misconception: students line up the last digits, so 4.6 + 3.75 becomes 4.6 lined under 3.75 wrongly.
  Why: they treat decimals like whole numbers.
  Impact: tenths add to hundredths and the answer is wrong.
  Quick correction: line up the decimal points; fill the empty place with a zero.

WATCH FOR:
- Students who line up the points and check the estimate - secure.
- Students who right-align the digits - reteach lining up the point.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a subtraction: 5.3 take away 2.85. Estimate first: 5 take away 3 is about 2.
- I line up the points and write 5.3 as 5.30 so both have two decimal places.
- Now I subtract from the right. 0 take away 5 - I need to regroup. I borrow to make it work, then continue.
- The answer is 2.45, which is close to my estimate of 2. Sensible.

DO:
- Write the estimate.
- Line up the points and add the placeholder zero.
- Subtract from the right, regrouping where needed.

TEACHER NOTES:
Subtraction with a placeholder zero and regrouping. The estimate confirms the answer is in the right range.

MISCONCEPTIONS:
- Misconception: students do 0 take away 5 as 5, avoiding the regroup.
  Why: they subtract the smaller from the larger regardless of order.
  Impact: a wrong answer, often too big.
  Quick correction: you cannot take 5 from 0; regroup from the next column first.

WATCH FOR:
- Students who regroup correctly - secure.
- Students who flip the column - reteach borrowing.

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- First estimate 6.9 plus 2.2, then work out the exact answer.
- Write both your estimate and your exact answer.

DO:
- Display the prompt.
- Give about 60 seconds.
- Walk and scan both numbers.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me your estimate and exact answer on three, two, one, show.
- Scan for: estimate about 9, exact 9.1.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students get 8.11 by mis-lining the points.
- Reteach: line up the points; 6.9 plus 2.2, tenths 9 and 2 make 11, carry 1; answer 9.1.
- Re-check: what is 9 tenths plus 2 tenths?
- Use the vertical layout to confirm.

TEACHER NOTES:
Estimate then exact. The estimate of 9 flags any wildly wrong exact answer.

WATCH FOR:
- Students who get 9.1 near their estimate - secure.
- Students far from their estimate - reteach lining up.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Estimate, then add 7.4 and 2.85.
- Round each first, then line up the points and add.

DO:
- Display 7.4 + 2.85.
- Give about 90 seconds for partners.
- Listen for an estimate near 10.

TEACHER NOTES:
Estimate about 7 plus 3 is 10. Exact 10.25. The placeholder zero on 7.40 matters.

WATCH FOR:
- Pairs who line up and add a zero - secure.
- Pairs who right-align digits - point to the decimal points.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Estimate: 7 plus 3 is about 10.
- Exact: 7.40 plus 2.85 is 10.25. Close to 10, so sensible.

DO:
- Click to reveal.
- Compare the exact answer to the estimate.

TEACHER NOTES:
The estimate and the exact answer agree, which builds confidence in the method.

WATCH FOR:
- Students who self-correct - secure.
- Students who misalign - small group before You Do.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, a subtraction. Estimate, then work out 8.2 take away 3.65.
- Round each first, then line up and subtract, regrouping where needed.

DO:
- Display 8.2 - 3.65.
- Give about 90 seconds.
- Listen for an estimate near 4.

TEACHER NOTES:
Estimate 8 take away 4 is about 4. Exact 4.55. Needs 8.20 and regrouping.

WATCH FOR:
- Pairs who write 8.20 and regroup - secure.
- Pairs who avoid the regroup - reteach borrowing.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Estimate: 8 take away 4 is about 4.
- Exact: 8.20 take away 3.65 is 4.55. Close to 4, so sensible.

DO:
- Click to reveal.
- Compare exact to estimate.

TEACHER NOTES:
Subtraction with regrouping confirmed by the estimate. Next session brings fractions, decimals and percentages together.

WATCH FOR:
- Students who get 4.55 - ready for independent work.
- Students unsure - one more subtraction before You Do.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- The first section is estimate only - round and add or subtract the wholes.
- The next section is exact addition. The last is exact subtraction.
- Always write your estimate first, then check your exact answer against it.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for estimate first, line up the points.
- Cold call one or two students to compare an exact answer to its estimate.

TEACHER NOTES:
Different numbers from the We Do, same habit: estimate, line up the points, add a placeholder zero, then calculate and check.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the grid-lined section where the decimal points are already lined up, and do Section 1 and the first row of Section 2.
- Extra Notes: Sit with these students and line up the first one together.
EXTENDING PROMPT:
- Task: Extension - add and subtract decimals to thousandths, and solve a two-step money or measurement problem. Early finishers may start the Year 8 Extension Challenge later this week.
- Extra Notes: Push estimating first every time.

WATCH FOR:
- Students who line up and check the estimate - secure.
- Students who right-align digits - prompt back to the decimal point.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- First estimate 9.6 take away 4.75, then work out the exact answer.
- Write both your estimate and your exact answer.

DO:
- Display the prompt.
- Give about 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - estimate, then subtract decimals accurately. Estimate about 5; exact 4.85.

WATCH FOR:
- Students who get about 5 then 4.85 - secure.
- Students whose exact answer is far from their estimate - revisit lining up at the start of Session 4.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: why do we line up the decimal points before adding or subtracting?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is estimate first, line up the decimal points so places match, then calculate and check against the estimate.

WATCH FOR:
- Strong thumbs across all three - move at pace next session.
- Sideways or down on the core criterion - quick revision at the start of Session 4.

[Retention and Recall]`;

// --- Helpers -----------------------------------------------------------------

// Right-aligned vertical algorithm (decimal points align when decimals match).
function verticalSum(slide, x, y, w, lines, opts) {
  const o = opts || {};
  const lineH = o.lineH || 0.6;
  const fontSize = o.fontSize || 32;
  lines.forEach((ln, i) => {
    const ly = y + i * lineH;
    if (ln.rule) {
      slide.addShape("line", {
        x: x, y: ly - 0.03, w: w, h: 0,
        line: { color: C.CHARCOAL, width: 2 },
      });
    }
    slide.addText(String(ln.text), {
      x: x, y: ly, w: w, h: lineH,
      fontSize, fontFace: FONT_H, color: ln.color || C.CHARCOAL, bold: true,
      align: "right", valign: "middle", margin: 0,
    });
  });
}

// Right panel showing estimate chip + vertical algorithm.
function sumPanel(slide, lg, headerText, estimateText, lines, opts) {
  const o = opts || {};
  const cardH = o.cardH || 2.9;
  addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: o.strip || C.PRIMARY });
  slide.addText(headerText, {
    x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
    fontSize: 15, fontFace: FONT_H, color: o.strip || C.PRIMARY, bold: true,
    align: "center", margin: 0,
  });
  addTextOnShape(slide, estimateText, {
    x: lg.rightX + 0.35, y: lg.panelTopPadded + 0.5, w: lg.rightW - 0.7, h: 0.45, rectRadius: 0.06,
    fill: { color: C.SECONDARY },
  }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
  const sumW = 2.2;
  const sumX = lg.rightX + (lg.rightW - sumW) / 2;
  verticalSum(slide, sumX, lg.panelTopPadded + 1.12, sumW, lines,
    { fontSize: o.sumFontSize || 30, lineH: 0.55, resultColor: o.resultColor });
}

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Week 6 Session 3: Adding and subtracting decimals",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Metric measurement
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Metric measurement",
      [
        "Convert: 1.25 m = ____ cm",
        "Convert: 450 g = ____ kg",
        "2 L - 750 mL = ____ mL",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "125 cm     0.45 kg     1250 mL", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal subtraction
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal subtraction",
      ["9.4 - 2.7", "15.2 - 8.6", "7.05 - 3.4"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "6.7        6.6        3.65", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - estimate first
  contentSlide(pres, "Launch", C.ACCENT, "Estimate first: about how much?",
    [
      "3.7 + 2.2 = about how much?",
      "",
      "Round: 4 + 2 = 6.",
      "The real answer should be near 6.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.3, { strip: C.ACCENT });
      slide.addText("3.7  ->  4", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.3, w: lg.rightW - 0.4, h: 0.5,
        fontSize: 24, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      slide.addText("2.2  ->  2", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.85, w: lg.rightW - 0.4, h: 0.5,
        fontSize: 24, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      addTextOnShape(slide, "Estimate: about 6", {
        x: lg.rightX + 0.7, y: lg.panelTopPadded + 1.5, w: lg.rightW - 1.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to estimate, then add and subtract decimals using place value.",
    [
      "I can estimate a decimal sum or difference by rounding.",
      "I can add decimals by lining up the decimal points.",
      "I can subtract decimals, regrouping when needed.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Estimate & round",
    [
      "Estimate: an answer that is close enough, found quickly.",
      "Round to the nearest whole: look at the tenths (5+ rounds up).",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.SECONDARY });
      const rows = [["3.7", "4"], ["2.2", "2"], ["6.5", "7"]];
      const ry0 = lg.panelTopPadded + 0.25;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.55;
        slide.addText(r[0] + "  rounds to  " + r[1], {
          x: lg.rightX + 0.2, y: ry, w: lg.rightW - 0.4, h: 0.45,
          fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10: I Do #1 - estimate then add 4.6 + 3.75
  workedExSlide(pres, 2, "I Do", "Estimate, then add: 4.6 + 3.75",
    [
      "Estimate: 5 + 4 = about 9.",
      "",
      "Line up the decimal points.",
      "Write 4.6 as 4.60.",
      "Add from the right.",
      "",
      "Answer 8.35 - close to 9, sensible.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      sumPanel(slide, lg, "Line up the points", "Estimate: 5 + 4 = 9", [
        { text: "4.60" },
        { text: "+ 3.75" },
        { text: "8.35", rule: true, color: C.SECONDARY },
      ], { strip: C.PRIMARY, cardH: 2.9 });
    }
  );

  // Slide 11: I Do #2 - estimate then subtract 5.3 - 2.85
  workedExSlide(pres, 2, "I Do", "Estimate, then subtract: 5.3 - 2.85",
    [
      "Estimate: 5 - 3 = about 2.",
      "",
      "Line up the decimal points.",
      "Write 5.3 as 5.30.",
      "Subtract from the right, regroup.",
      "",
      "Answer 2.45 - close to 2, sensible.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      sumPanel(slide, lg, "Line up the points", "Estimate: 5 - 3 = 2", [
        { text: "5.30" },
        { text: "- 2.85" },
        { text: "2.45", rule: true, color: C.SECONDARY },
      ], { strip: C.PRIMARY, cardH: 2.9 });
    }
  );

  // Slides 12-13: CFU + reveal - estimate then 6.9 + 2.2
  withReveal(
    () => cfuSlide(pres, "CFU", "Estimate, then add 6.9 + 2.2", "Show Me Boards",
      "Estimate 6.9 + 2.2, then work out the exact answer.\n\nWrite BOTH your estimate and your exact answer.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "Estimate: 7 + 2 = about 9.    Exact: 9.1",
        { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 7.4 + 2.85
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Estimate, then add: 7.4 + 2.85",
      [
        "With your partner.",
        "",
        "1.  Estimate by rounding.",
        "2.  Line up the points (7.40).",
        "3.  Add, then check the estimate.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        sumPanel(slide, lg, "Line up the points", "Estimate: 7 + 3 = 10", [
          { text: "7.40" },
          { text: "+ 2.85" },
          { text: "?", rule: true, color: C.ALERT },
        ], { strip: C.SECONDARY, cardH: 2.9 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "7.40 + 2.85 = 10.25   (estimate was 10)", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - 8.2 - 3.65
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Estimate, then subtract: 8.2 - 3.65",
      [
        "With your partner.",
        "",
        "1.  Estimate by rounding.",
        "2.  Line up the points (8.20).",
        "3.  Subtract, regroup, check.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        sumPanel(slide, lg, "Line up the points", "Estimate: 8 - 4 = 4", [
          { text: "8.20" },
          { text: "- 3.65" },
          { text: "?", rule: true, color: C.ALERT },
        ], { strip: C.SECONDARY, cardH: 2.9 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "8.20 - 3.65 = 4.55   (estimate was 4)", { color: C.SUCCESS });
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
      { text: "estimate by rounding.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "add decimals, lining up the points.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "subtract, regrouping when needed.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Estimate first. Line up the decimal points. Fill empty places with a zero. Check against your estimate.", {
      x: 0.85, y: panelY + 0.52, w: 8.3, h: 0.55, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

    verticalSum(s, 4.0, panelY + 1.25, 2.0, [
      { text: "4.60" },
      { text: "+ 3.75" },
      { text: "8.35", rule: true, color: C.SECONDARY },
    ], { fontSize: 26, lineH: 0.48 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Estimate 9.6 - 4.75, then work out the exact answer.",
      "Write both your estimate and your exact answer.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why do we line up the decimal points before adding or subtracting?",
      scItems: [
        "I can estimate a decimal sum or difference by rounding.",
        "I can add decimals by lining up the decimal points.",
        "I can subtract decimals, regrouping when needed.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDec_W6S3_Add_Subtract_Decimals.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Estimate, then add and subtract decimals by lining up the decimal points.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Estimate first by rounding each decimal to the nearest whole. Then line up the decimal points, fill empty places with a zero, and add or subtract from the right. Check your exact answer is close to your estimate.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "4.6 + 3.75: estimate 5 + 4 = 9. Line up: 4.60 + 3.75 = 8.35. Close to 9, so it is sensible.",
      y);

    y = addSectionHeading(doc, "Section 1 - Estimate only (round to the nearest whole)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3.7 + 2.2 is about ______        b)  8.1 - 2.9 is about ______", y);
    y = addWriteLine(doc, "c)  6.8 + 4.3 is about ______        d)  9.7 - 5.2 is about ______", y);

    y = addSectionHeading(doc, "Section 2 - Add (estimate first, then line up the points)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  4.6 + 3.75 = ______      b)  5.4 + 2.8 = ______      c)  7.25 + 1.9 = ______", y);
    y = addWriteLine(doc, "d)  6.08 + 3.4 = ______      e)  12.5 + 8.65 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Subtract (estimate first, regroup when needed)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  5.3 - 2.85 = ______      b)  8.2 - 3.65 = ______      c)  9 - 4.7 = ______", y);
    y = addWriteLine(doc, "d)  7.05 - 2.4 = ______      e)  15.2 - 8.75 = ______", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "A long-jumper jumps 4.85 m, then 5.12 m. How much further was the second jump? Then add both jumps.", y);
    y = addWriteLine(doc, "Difference: ______        Total: ______", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Adding and Subtracting Decimals | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the adding and subtracting decimals practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Estimate only", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  about 6 (4 + 2)      b)  about 5 (8 - 3)      c)  about 11 (7 + 4)      d)  about 5 (10 - 5)", y);

    y = addSectionHeading(doc, "Section 2 - Add", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  8.35      b)  8.2      c)  9.15      d)  9.48      e)  21.15", y);

    y = addSectionHeading(doc, "Section 3 - Subtract", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2.45      b)  4.55      c)  4.3      d)  4.65      e)  6.45", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Difference: 5.12 - 4.85 = 0.27 m. Total: 4.85 + 5.12 = 9.97 m.", y);

    y = addTipBox(doc,
      "Watch for: students who right-align digits instead of the decimal point; students who skip the placeholder zero; students who avoid regrouping by flipping a column.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Week 6 Session 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
