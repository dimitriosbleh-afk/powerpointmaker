"use strict";

// Fractions & Decimals - Week 6 (Year 6 Numeracy) - Session 1: Percentages - 100% is the whole
// VC2M6N03 / percentages. 100% = one whole; connect familiar percentages to fractions and decimals.
// Daily Review: Fractions & Decimals. Fluency: subtraction algorithm.
// Anchor models: 100-grid (hundredths) + fraction wall + number line. Variant weekToVariant(6).

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
  addFractionStripSet, addNumberLine, addAreaModel,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 4;
const WEEK = 6;
const UNIT_TITLE = "Fractions & Decimals";
const FOOTER = `Fractions & Decimals | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracDec_W6S1_Percentages_Whole";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Percentages Practice",
  "Connect percentages to fractions and decimals using the 100-grid. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Percentages Answer Key",
  "Worked answers for the percentages practice sheet.");
const CATCHUP_RES = makeSessionResource(SESSION,
  "Week 6 Catch-Up Card",
  "One-page recap of the four Week 6 ideas with a worked example each. For a student returning after an absence.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, CATCHUP_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to Week 6. Last week we built and combined fractions.
- This week we connect fractions to decimals and percentages. Today we start with percentages and what 100% really means.

DO:
- Have whiteboards, markers and the printed 100-grid ready.
- Settle the class before you click on.

TEACHER NOTES:
Week 6 Session 1 of 4. The 100-grid is the new anchor this week, alongside the fraction wall and number line. A returning student needs the printed 100-grid and the Week 6 Catch-Up Card.

WATCH FOR:
- Students who think a percentage is a totally new kind of number - reassure, it is just hundredths.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- Today's materials are the practice sheet and a printed 100-grid.
- The practice sheet has an easier start and a challenge built in.

DO:
- Print one practice sheet per student and keep the answer key.
- Have the printed 100-grid and whiteboards ready.

TEACHER NOTES:
One student practice sheet plus its answer key, and the Week 6 Catch-Up Card for absences.

CATCH-UP NOTE:
A student returning this week can pick up here. Every Week 6 lesson re-shows the 100-grid or number line in the launch, and the Week 6 Catch-Up Card summarises percentages, comparing fractions, decimals, and the FDP links with a worked example each.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Fractions and decimals from last week and earlier.
- Answer each on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and check reasoning.

TEACHER NOTES:
Spaced retrieval of fraction and decimal facts. This sets up today's links to percentages.

WATCH FOR:
- Students who know 3/4 is 0.75 - secure.
- Students who think 0.07 is bigger than 0.7 - revisit place value.

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Three-quarters as a decimal is 0.75.
- 0.7 is bigger than 0.07; 0.7 is seven tenths, 0.07 is seven hundredths.
- Four-eighths simplifies to one half.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The comparison item checks decimal place value. The simplify item links to equivalent fractions.

WATCH FOR:
- Students who self-correct - secure.
- Students who pick 0.07 - reteach tenths versus hundredths.

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. This week we practise the subtraction algorithm with decimals.
- Line up the decimal points, subtract, regroup if you need to.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up points and regrouping.

TEACHER NOTES:
Fluency this whole week is decimal subtraction. Lining up the point is the key habit, same as adding.

WATCH FOR:
- Students who line up the points - secure.
- Students who subtract the smaller digit from the larger regardless of order - coach regrouping.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 8.5 minus 3.2 is 5.3.
- 12.4 minus 5.7 is 6.7.
- 9.06 minus 2.4 is 6.66.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
9.06 minus 2.4 needs 2.4 read as 2.40 to line up. Coach the placeholder zero.

WATCH FOR:
- Students who self-correct - secure.
- Students who misalign 2.4 - small group focus.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Look at this 100-grid. The whole grid is one hundred small squares - that is 100 percent, the whole thing.
- Per cent just means out of 100. So 100 percent is all 100 squares.
- Now I shade 50 squares. That is 50 out of 100, which we call 50 percent.
- And 50 out of 100 is the same as one half. Today we connect percentages to fractions and decimals.

DO:
- Point around the whole grid as you say 100 percent.
- Shade or point to 50 squares and name it 50 percent.

TEACHER NOTES:
This launch grounds percent as hundredths of one whole. 100% is the whole; 50% is half of it.

WATCH FOR:
- Students who see 50 shaded as one half - strong start.
- Students unsure what percent means - repeat: out of 100.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning that 100 percent is the whole, and to connect percentages to fractions and decimals.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is saying what 100% means. Second is matching common percentages to fractions and decimals, the core target. Third is writing any percentage as a fraction and a decimal.

WATCH FOR:
- Students who can say out of 100 - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- One key word today.
- Per cent means out of 100. The percent sign stands for out of 100.
- So 25 percent means 25 out of 100, which we can write as the fraction 25 over 100, or the decimal 0.25.

DO:
- Write the percent sign and say out of 100.
- Show 25% as 25/100 and 0.25.

TEACHER NOTES:
Vocabulary after the learning intention. Percent is simply hundredths, so every percentage already has a fraction and decimal form.

WATCH FOR:
- Students who read the percent sign as out of 100 - secure.
- Students who think percent is separate from fractions - link them on the grid.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us connect 50 percent to a fraction and a decimal.
- 50 percent means 50 out of 100. On the grid that is half the squares shaded.
- As a fraction that is 50 over 100, which simplifies to one half.
- As a decimal, 50 hundredths is 0.5.
- So 50 percent equals one half equals 0.5. Three names, one amount.

DO:
- Shade 50 on the grid.
- Write 50% = 50/100 = 1/2 = 0.5 as you say it.

TEACHER NOTES:
The core triple: percent to fraction to decimal, all read off the 100-grid. Keep the three forms side by side.

MISCONCEPTIONS:
- Misconception: students write 50% as 0.50 meaning fifty, or as 5.0.
  Why: they are unsure where the decimal point goes.
  Impact: wrong size by a factor of ten or a hundred.
  Quick correction: percent is hundredths, so 50 percent is 50 hundredths, 0.5.

WATCH FOR:
- Students who give all three forms - secure.
- Students who write 5.0 - back to hundredths on the grid.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now two more we should just know. 25 percent and 10 percent.
- 25 out of 100 shaded is one quarter of the grid. So 25 percent is 25 over 100, which is one quarter, and 0.25 as a decimal.
- 10 out of 100 shaded is one column. So 10 percent is 10 over 100, which is one tenth, and 0.1 as a decimal.
- These are worth memorising because they come up everywhere.

DO:
- Shade 25 squares, then 10 squares, on grids.
- Write each triple: 25% = 1/4 = 0.25 and 10% = 1/10 = 0.1.

TEACHER NOTES:
These two benchmarks plus 50% cover most everyday percentage work. Anchor them on the grid.

MISCONCEPTIONS:
- Misconception: students think 25% is 0.25 but cannot see it is one quarter.
  Why: they treat the forms separately.
  Impact: they cannot reason flexibly later.
  Quick correction: 25 of the 100 squares is one of four equal quarters of the grid.

WATCH FOR:
- Students who know all three benchmarks - secure.
- Students unsure of 10% - show the single column.

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Write 75 percent as a fraction and as a decimal.
- Use the grid in your head: 75 out of 100 shaded.

DO:
- Display the prompt.
- Give about 45 seconds.
- Walk and scan the boards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me the fraction and decimal on three, two, one, show.
- Scan for: 75/100 = 3/4, and 0.75.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students write 0.075, or cannot simplify 75/100.
- Reteach: 75 percent is 75 hundredths, which is 0.75. On the grid, 75 of 100 is three of the four quarters, so 3/4.
- Re-check: how many hundredths is 75 percent?
- Use the grid to confirm.

TEACHER NOTES:
75% to 3/4 is the discriminator. Students who only memorised 50%, 25%, 10% must reason from the grid.

WATCH FOR:
- Students who write 3/4 and 0.75 - secure.
- Students who write 0.075 - reteach hundredths.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Write 20 percent as a fraction and a decimal.
- Picture 20 out of 100 shaded. What fraction is that, and what decimal?

DO:
- Display 20% with a grid.
- Give about 90 seconds for partners.
- Listen for one-fifth and 0.2.

TEACHER NOTES:
20% = 20/100 = 1/5 = 0.2. The simplifying to fifths is the new step.

WATCH FOR:
- Pairs who simplify 20/100 to 1/5 - secure.
- Pairs who stop at 20/100 - prompt them to simplify.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- 20 percent is 20 out of 100, which is 20 over 100.
- 20 over 100 simplifies to one fifth. As a decimal that is 0.2.
- So 20 percent equals one fifth equals 0.2.

DO:
- Click to reveal.
- Show 20/100 simplifying to 1/5.

TEACHER NOTES:
Reinforce simplifying. Five twenties make 100, so 20/100 is one of five equal parts.

WATCH FOR:
- Students who self-correct - secure.
- Students who cannot simplify - small group before You Do.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together. This grid has 40 squares shaded.
- What percentage is that? And what fraction and decimal?

DO:
- Display the grid with 40 shaded.
- Give about 75 seconds.
- Listen for 40 percent, two-fifths, 0.4.

TEACHER NOTES:
Reading from the grid to all three forms. 40/100 simplifies to 2/5.

WATCH FOR:
- Pairs who give 40 percent, 2/5, 0.4 - secure.
- Pairs who miscount the shaded squares - count columns of ten.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- 40 out of 100 shaded is 40 percent.
- As a fraction that is 40 over 100, which simplifies to two-fifths. As a decimal it is 0.4.

DO:
- Click to reveal.
- Show four full columns shaded as 40.

TEACHER NOTES:
Counting columns of ten makes the 40 quick to see. Next session we compare and order fractions on a line.

WATCH FOR:
- Students who read all three forms - ready for independent work.
- Students unsure - one more grid before You Do.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- The first section shades the grid to match a percentage.
- The next section writes percentages as fractions and decimals.
- The last section matches the three forms.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for out of 100, and the simplifying step.
- Cold call one or two students to explain a triple.

TEACHER NOTES:
Different percentages from the We Do, same method: read the grid, write hundredths, simplify the fraction.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed 100-grid and do Section 1 and the 50%, 25%, 10% rows of Section 2.
- Extra Notes: Sit with these students and shade the first grid together.
EXTENDING PROMPT:
- Task: Extension - write percentages like 5%, 35% and 100% as fractions and decimals, and a percentage over 100 such as 150%. Early finishers may start the Year 8 Extension Challenge later this week.
- Extra Notes: Push the idea that 100% is the whole, so 150% is more than one whole.

WATCH FOR:
- Students who give all three forms and simplify - secure.
- Students who forget the decimal point place - prompt hundredths.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Write 40 percent as a fraction and as a decimal.
- Then explain what 100 percent means.

DO:
- Display the prompt.
- Give about 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - matching a percentage to a fraction and decimal. Look for 40/100 = 2/5 = 0.4, and 100% described as the whole.

WATCH FOR:
- Students who write 2/5 and 0.4 - secure.
- Students who cannot say 100% is the whole - revisit at the start of Session 2.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: what does 100 percent mean, and what is 50 percent as a fraction?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that percent means out of 100, 100% is the whole, and every percentage has a fraction and decimal form read off the grid.

WATCH FOR:
- Strong thumbs across all three - move at pace next session.
- Sideways or down on the core criterion - quick revision at the start of Session 2.

[Retention and Recall]`;

// --- Helpers -----------------------------------------------------------------

// 100-grid (10x10 area model) in a right-column panel with a header and label.
function gridPanel(slide, lg, headerText, filledTenths, extraHundredths, labelText, opts) {
  const o = opts || {};
  const cardH = o.cardH || 2.9;
  addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: o.strip || C.PRIMARY });
  slide.addText(headerText, {
    x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
    fontSize: 15, fontFace: FONT_H, color: o.strip || C.PRIMARY, bold: true,
    align: "center", margin: 0,
  });
  const gridSize = Math.min(cardH - 1.0, 2.0);
  const gx = lg.rightX + (lg.rightW - gridSize) / 2;
  const gy = lg.panelTopPadded + 0.5;
  addAreaModel(slide, gx, gy, gridSize, filledTenths, extraHundredths || 0);
  if (labelText) {
    slide.addText(labelText, {
      x: lg.rightX + 0.2, y: gy + gridSize + 0.08, w: lg.rightW - 0.4, h: 0.4,
      fontSize: 15, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", margin: 0, fit: "shrink", shrinkText: true,
    });
  }
}

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Week 6 Session 1: Percentages - 100% is the whole",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Fractions & decimals
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions & decimals",
      [
        "Write 3/4 as a decimal",
        "Which is bigger: 0.7 or 0.07?",
        "Simplify 4/8",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "0.75     0.7     1/2", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal subtraction
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal subtraction",
      ["8.5 - 3.2", "12.4 - 5.7", "9.06 - 2.4"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "5.3        6.7        6.66", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - 100-grid, 50 shaded
  contentSlide(pres, "Launch", C.ACCENT, "What does 100% mean?",
    [
      "The whole grid = 100 squares = 100%.",
      "",
      "Per cent means 'out of 100'.",
      "50 shaded = 50% = one half.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      gridPanel(slide, lg, "100-grid: 50 shaded", 5, 0, "50 out of 100 = 50%",
        { strip: C.ACCENT, cardH: 2.9 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning that 100% is the whole and to connect percentages to fractions and decimals.",
    [
      "I can say what 100% means.",
      "I can match 50%, 25% and 10% to a fraction and a decimal.",
      "I can write a percentage as a fraction and a decimal.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Per cent = out of 100",
    [
      "Per cent (%) means 'out of 100'.",
      "25% = 25 out of 100.",
      "= 25/100 (fraction) = 0.25 (decimal).",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      gridPanel(slide, lg, "25 out of 100", 2, 5, "25% = 1/4 = 0.25",
        { strip: C.SECONDARY, cardH: 2.9 });
    }
  );

  // Slide 10: I Do #1 - 50% triple
  workedExSlide(pres, 2, "I Do", "50% = one half = 0.5",
    [
      "50% means 50 out of 100.",
      "On the grid: half the squares.",
      "",
      "Fraction: 50/100 = 1/2",
      "Decimal: 50 hundredths = 0.5",
      "",
      "So 50% = 1/2 = 0.5.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      gridPanel(slide, lg, "50 shaded", 5, 0, "50% = 1/2 = 0.5",
        { strip: C.PRIMARY, cardH: 2.9 });
    }
  );

  // Slide 11: I Do #2 - 25% and 10%
  workedExSlide(pres, 2, "I Do", "Two more to know: 25% and 10%",
    [
      "25% = 25 out of 100 = 1/4 = 0.25",
      "(one quarter of the grid)",
      "",
      "10% = 10 out of 100 = 1/10 = 0.1",
      "(one column of the grid)",
      "",
      "Worth remembering - they appear everywhere.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.PRIMARY });
      const g = 1.55;
      addAreaModel(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.45, g, 2, 5);
      slide.addText("25% = 1/4", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.45 + g + 0.02, w: g + 0.2, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      addAreaModel(slide, lg.rightX + lg.rightW - g - 0.3, lg.panelTopPadded + 0.45, g, 1, 0);
      slide.addText("10% = 1/10", {
        x: lg.rightX + lg.rightW - g - 0.4, y: lg.panelTopPadded + 0.45 + g + 0.02, w: g + 0.2, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
    }
  );

  // Slides 12-13: CFU + reveal - 75%
  withReveal(
    () => cfuSlide(pres, "CFU", "Write 75% as a fraction and a decimal", "Show Me Boards",
      "Write 75% as a fraction and as a decimal.\n\nPicture 75 out of 100 shaded.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "75% = 75/100 = 3/4 = 0.75", { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 20%
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Write 20% as a fraction and decimal",
      [
        "With your partner.",
        "",
        "1.  20 out of 100 shaded.",
        "2.  Write the fraction, then simplify.",
        "3.  Write the decimal.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        gridPanel(slide, lg, "20 shaded", 2, 0, "20% = ?",
          { strip: C.SECONDARY, cardH: 2.9 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "20% = 20/100 = 1/5 = 0.2", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - read 40 from the grid
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Read the grid: what percentage?",
      [
        "With your partner.",
        "",
        "1.  Count the shaded squares.",
        "2.  Write the percentage.",
        "3.  Write the fraction and decimal.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        gridPanel(slide, lg, "How many shaded?", 4, 0, "____ %",
          { strip: C.SECONDARY, cardH: 2.9 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "40 shaded = 40% = 40/100 = 2/5 = 0.4", { color: C.SUCCESS });
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
      { text: "shade the grid to match the percentage.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "write percentages as fractions and decimals.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "match the three forms.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Per cent means out of 100. Write it as hundredths, then simplify the fraction.", {
      x: 0.95, y: panelY + 0.52, w: 8.1, h: 0.5, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    const benchY = panelY + 1.2;
    const b = ["50% = 1/2 = 0.5", "25% = 1/4 = 0.25", "10% = 1/10 = 0.1"];
    const bw = 2.7;
    b.forEach((t, i) => {
      addTextOnShape(s, t, {
        x: 0.85 + i * (bw + 0.2), y: benchY, w: bw, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Write 40% as a fraction and as a decimal.",
      "Explain what 100% means.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what does 100% mean, and what is 50% as a fraction?",
      scItems: [
        "I can say what 100% means.",
        "I can match 50%, 25% and 10% to a fraction and a decimal.",
        "I can write a percentage as a fraction and a decimal.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDec_W6S1_Percentages_Whole.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Connect percentages to fractions and decimals using the 100-grid.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Per cent means out of 100. A percentage is just hundredths, so it has a fraction form (over 100, then simplify) and a decimal form. 100% is the whole. Benchmarks: 50% = 1/2 = 0.5, 25% = 1/4 = 0.25, 10% = 1/10 = 0.1.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "75%: 75 out of 100 = 75/100. Simplify by dividing top and bottom by 25: 3/4. Decimal: 0.75. So 75% = 3/4 = 0.75.",
      y);

    y = addSectionHeading(doc, "Section 1 - Shade and name (use your 100-grid)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  Shade 50%. That is ______ squares.    b)  Shade 30%. That is ______ squares.", y);
    y = addWriteLine(doc, "c)  Shade 10%. That is ______ squares.    d)  Shade 100%. That is ______ squares.", y);

    y = addSectionHeading(doc, "Section 2 - Write each percentage as a fraction and a decimal", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  50% = ______ = ______        b)  25% = ______ = ______", y);
    y = addWriteLine(doc, "c)  10% = ______ = ______        d)  20% = ______ = ______", y);
    y = addWriteLine(doc, "e)  75% = ______ = ______        f)  60% = ______ = ______", y);

    y = addSectionHeading(doc, "Section 3 - Match the three forms", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Join each percentage to its fraction and its decimal.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "25%   1%   50%   10%        1/2   1/4   1/100   1/10        0.5   0.25   0.01   0.1", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Write these as fractions and decimals: 5%, 35%, 150%. (Hint: 100% is one whole, so 150% is more than one whole.)", y);
    y = addWriteLine(doc, "5% = ______ = ______   35% = ______ = ______   150% = ______ = ______", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Percentages | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the percentages practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Shade and name", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  50 squares      b)  30 squares      c)  10 squares      d)  100 squares", y);

    y = addSectionHeading(doc, "Section 2 - Percentage to fraction and decimal", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/2 = 0.5      b)  1/4 = 0.25      c)  1/10 = 0.1      d)  1/5 = 0.2", y);
    y = addBodyText(doc, "e)  3/4 = 0.75      f)  3/5 = 0.6", y);

    y = addSectionHeading(doc, "Section 3 - Match the three forms", y, { color: C.PRIMARY });
    y = addBodyText(doc, "25% = 1/4 = 0.25.   1% = 1/100 = 0.01.   50% = 1/2 = 0.5.   10% = 1/10 = 0.1.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "5% = 5/100 = 1/20 = 0.05.   35% = 35/100 = 7/20 = 0.35.   150% = 150/100 = 1 1/2 = 1.5.", y);

    y = addTipBox(doc,
      "Watch for: students who write 75% as 0.075; students who do not simplify the fraction; students who think a percentage cannot be more than 100.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: CATCHUP_RES.name });
    let y = addPdfHeader(doc, CATCHUP_RES.name, {
      subtitle: "Missed a lesson? Here are the four big ideas of Week 6 with one worked example each.",
      color: C.SECONDARY,
      lessonInfo: `Week ${WEEK} | Year 6 Numeracy | Catch-up`,
      showNameDate: false,
    });
    y = addTipBox(doc,
      "Our tools this week are the 100-GRID (100 squares = the whole = 100%), the NUMBER LINE (compare and order), and the FDP triple (fraction = decimal = percentage). Keep this card in your book.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Session 1 - Percentages", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Per cent means out of 100. 100% is the whole. 50% = 1/2 = 0.5, 25% = 1/4 = 0.25, 10% = 1/10 = 0.1.", y);

    y = addSectionHeading(doc, "Session 2 - Compare and order fractions", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Put fractions on one 0-1 number line, or rename to the same bottom. 1/4 = 3/12, 1/3 = 4/12, 1/2 = 6/12, so 1/4 < 1/3 < 1/2.", y);

    y = addSectionHeading(doc, "Session 3 - Add and subtract decimals", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Estimate first by rounding. Then line up the decimal points and add or subtract. 4.6 + 3.75: estimate 5 + 4 = 9; exact 8.35.", y);

    y = addSectionHeading(doc, "Session 4 - Fractions, decimals and percentages together", y, { color: C.PRIMARY });
    y = addBodyText(doc, "The same amount has three forms. 3/4 = 0.75 = 75%. To compare, change to the same form. 25% off means take 1/4 away.", y);

    addPdfFooter(doc, `Week ${WEEK} | Catch-Up Card | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, CATCHUP_RES.fileName));
    console.log("PDF written: " + CATCHUP_RES.fileName);
  })();

  console.log("Week 6 Session 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
