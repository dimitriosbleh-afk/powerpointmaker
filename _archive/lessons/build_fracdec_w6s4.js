"use strict";

// Fractions & Decimals - Week 6 (Year 6 Numeracy) - Session 4: Fractions, decimals & percentages together
// VC2M6N03 / percentages. Move between F, D and P to compare and to solve real-world problems.
// Daily Review: Practical problem solving & financial reasoning. Fluency: subtraction algorithm.
// Anchor models: 100-grid + FDP triple. Variant weekToVariant(6). Ships the Year 8 extension for Week 6.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(6));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, dailyReviewSlide, fluencySlide,
  addStageBadge, addRevealAnswerBar,
  addAreaModel,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, STAGE_COLORS,
} = T;

const SESSION = 4;
const TOTAL = 4;
const WEEK = 6;
const UNIT_TITLE = "Fractions & Decimals";
const FOOTER = `Fractions & Decimals | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracDec_W6S4_FDP_Together";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Fractions Decimals Percentages Practice",
  "Complete FDP triples, compare across forms, and solve discount problems. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Fractions Decimals Percentages Answer Key",
  "Worked answers for the fractions, decimals and percentages practice sheet.");
const Y8_RES = makeSessionResource(SESSION,
  "Year 8 Extension Challenge",
  "Beyond Year 6 - percentage of amounts, increase and decrease, finding the whole, recurring decimals, and reasoning.");
const Y8_KEY_RES = makeSessionResource(SESSION,
  "Year 8 Extension Answer Key",
  "Worked answers and look-fors for the Year 8 extension challenge.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, Y8_RES, Y8_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back, and the last session of the unit.
- Today we bring fractions, decimals and percentages together and use them to solve everyday problems like discounts.

DO:
- Have whiteboards, markers and the printed 100-grid ready.
- Settle the class before you click on.

TEACHER NOTES:
Week 6 Session 4 of 4. Pulls the two weeks together. A returning student needs the Week 6 Catch-Up Card and the printed 100-grid.

WATCH FOR:
- Students who can do one form but not switch between them - today is all about switching.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- Today there is a practice sheet, plus a Year 8 Extension Challenge for anyone ready to push further.
- The practice sheet has an easier start and a challenge built in.

DO:
- Print one practice sheet per student and keep the answer key.
- Print a few Year 8 Extension Challenge sheets for early finishers.

TEACHER NOTES:
Two student sheets with answer keys. The Year 8 sheet is optional and aimed beyond Year 6 - percentage of amounts, increase and decrease, and finding the whole.

CATCH-UP NOTE:
A returner can still join. The launch re-shows the fraction-decimal-percentage links, and the Week 6 Catch-Up Card lists the FDP triple with a worked example.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Money and everyday problem solving, which fits today's lesson well.
- Answer each on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and listen for sensible money reasoning.

TEACHER NOTES:
Financial reasoning retrieval, a natural bridge to today's discounts.

WATCH FOR:
- Students who find 10 percent as one tenth - secure.
- Students who guess the change - prompt counting up from the price.

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 10 percent off 50 dollars is 5 dollars off.
- Change from 20 dollars for an item costing 13 dollars 50 is 6 dollars 50.
- 4 dollars 85 rounds to 5 dollars.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The discount item links straight to today. Ten percent is one tenth, so divide by ten.

WATCH FOR:
- Students who self-correct - secure.
- Students who struggle with change - model counting up.

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Decimal subtraction, vertical, last time this unit.
- Line up the decimal points, subtract, regroup if needed.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up points.

TEACHER NOTES:
Decimal subtraction one more time. It supports working out change and discounts.

WATCH FOR:
- Students who regroup correctly - secure.
- Students who flip a column - coach borrowing.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 16.4 minus 8.7 is 7.7.
- 20 minus 6.25 is 13.75.
- 11.3 minus 4.85 is 6.45.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
20 minus 6.25 needs 20 read as 20.00 to line up. Coach the placeholders.

WATCH FOR:
- Students who self-correct - secure.
- Students who misalign whole numbers - small group focus.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- A jumper is 25 percent off. Let us think about what that means in different ways.
- 25 percent is the same as one quarter, and the same as 0.25.
- So 25 percent off means a quarter of the price comes off.
- Today we move smoothly between fractions, decimals and percentages, and use them in real problems.

DO:
- Show 25% next to 1/4 and 0.25.
- Ask: how much of the price comes off if it is a quarter off?

TEACHER NOTES:
This launch reactivates the FDP links from Session 1 and points them at a real discount.

WATCH FOR:
- Students who say 25 percent is a quarter - strong start.
- Students unsure - show the 100-grid with 25 shaded.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to move between fractions, decimals and percentages to solve everyday problems.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is completing an FDP triple. Second is choosing the easier form to compare, the core target. Third is solving a discount or comparison problem.

WATCH FOR:
- Students who can name the three forms - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- One key idea today: equivalent forms.
- The same amount can be written as a fraction, a decimal and a percentage.
- One half is 0.5 is 50 percent. They are three names for the same amount.

DO:
- Show the triple 1/2 = 0.5 = 50%.
- Have students read all three forms.

TEACHER NOTES:
Vocabulary after the learning intention. The big idea is that the three forms are interchangeable.

WATCH FOR:
- Students who read all three forms fluently - secure.
- Students who treat them as separate - link them on the grid.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us complete a full triple for three-quarters.
- As a fraction it is three-quarters. On the 100-grid that is 75 of the 100 squares.
- 75 out of 100 is 0.75 as a decimal, and 75 percent.
- So three-quarters equals 0.75 equals 75 percent. One amount, three forms.

DO:
- Shade 75 on the grid.
- Write the triple: 3/4 = 0.75 = 75%.

TEACHER NOTES:
The core triple, read off the 100-grid. Keep all three forms visible together.

MISCONCEPTIONS:
- Misconception: students write 3/4 as 0.34 or 3.4.
  Why: they read the digits, not the value.
  Impact: wrong decimal and percentage.
  Quick correction: three-quarters is 75 hundredths, which is 0.75.

WATCH FOR:
- Students who give all three forms - secure.
- Students who write 0.34 - back to the grid and hundredths.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a real discount. A jacket costs 40 dollars and is 25 percent off.
- 25 percent is one quarter. A quarter of 40 is 10, so 10 dollars comes off.
- 40 take away 10 is 30. The jacket now costs 30 dollars.
- Choosing the fraction form, one quarter, made the discount easy to work out.

DO:
- Write 25% = 1/4.
- Find one quarter of 40, then subtract from the price.

TEACHER NOTES:
This shows why switching forms is useful: a quarter of 40 is friendlier than 0.25 times 40 for mental work.

MISCONCEPTIONS:
- Misconception: students subtract 25 from 40, treating the percent as dollars.
  Why: they ignore that percent is a fraction of the price.
  Impact: a wrong discount.
  Quick correction: 25 percent OF 40 means one quarter of 40, which is 10.

WATCH FOR:
- Students who find a quarter of 40 - secure.
- Students who take 25 dollars off - reteach percent of an amount.

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Complete the triple for one-fifth. Write it as a decimal and a percentage.

DO:
- Display the prompt.
- Give about 45 seconds.
- Walk and scan the boards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me the decimal and the percentage on three, two, one, show.
- Scan for: 1/5 = 0.2 = 20%.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students write 1/5 as 0.15 or 15%.
- Reteach: one-fifth is the same as two-tenths, which is 0.2, and 20 out of 100, which is 20 percent.
- Re-check: how many tenths is one-fifth?
- Use the 100-grid to confirm.

TEACHER NOTES:
1/5 to 0.2 to 20% is the discriminator. Renaming to tenths makes it clear.

WATCH FOR:
- Students who write 0.2 and 20% - secure.
- Students who write 0.15 - reteach via tenths.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Which is bigger, 0.7 or three-quarters?
- Change them into the same form so you can compare.

DO:
- Display 0.7 and 3/4.
- Give about 90 seconds for partners.
- Listen for 3/4 = 0.75.

TEACHER NOTES:
The strategy is to choose one form and compare. 3/4 = 0.75, which is more than 0.7.

WATCH FOR:
- Pairs who convert 3/4 to 0.75 - secure.
- Pairs who think 0.7 is bigger because seven - reteach via 0.70 vs 0.75.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Three-quarters as a decimal is 0.75.
- 0.75 is bigger than 0.7, because 75 hundredths is more than 70 hundredths.
- So three-quarters is bigger than 0.7.

DO:
- Click to reveal.
- Compare 0.70 and 0.75 on the board.

TEACHER NOTES:
Writing 0.7 as 0.70 makes the comparison obvious. Same form, then compare.

WATCH FOR:
- Students who self-correct - secure.
- Students unsure - line up the hundredths.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, a discount problem.
- A 60 dollar jacket is 25 percent off. What is the new price?
- Use the fraction form to make it easy.

DO:
- Display the problem.
- Give about 90 seconds.
- Listen for one quarter of 60 is 15.

TEACHER NOTES:
25 percent is one quarter. A quarter of 60 is 15, so 15 dollars off. New price 45 dollars.

WATCH FOR:
- Pairs who find a quarter of 60 then subtract - secure.
- Pairs who take 25 dollars off - reteach percent of an amount.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- 25 percent is one quarter. A quarter of 60 is 15.
- So 15 dollars comes off. 60 take away 15 is 45. The new price is 45 dollars.

DO:
- Click to reveal.
- Restate the new price.

TEACHER NOTES:
Switching 25 percent to one quarter made this a friendly division. That is the power of moving between forms.

WATCH FOR:
- Students who get 45 dollars - ready for independent work.
- Students unsure - one more discount before You Do.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- The first section completes fraction-decimal-percentage triples.
- The next section compares two numbers by changing them to the same form.
- The last section solves discount problems.
- If you finish, try the extension, or the Year 8 Extension Challenge.

DO:
- Distribute the practice sheet.
- Circulate and listen for change to the same form, and percent of an amount.
- Cold call one or two students to explain a discount.

TEACHER NOTES:
Different values from the We Do, same skills: complete the triple, compare in one form, find a percent of an amount.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the benchmark card (1/2, 1/4, 1/10, 1/5) and the 100-grid, and do Section 1 and the first comparison in Section 2.
- Extra Notes: Sit with these students and complete the first triple together.
EXTENDING PROMPT:
- Task: Extension - find 10 percent and 5 percent of amounts, and try the Year 8 Extension Challenge with percentage increase and finding the whole.
- Extra Notes: Push the idea that 10 percent is one tenth, so divide by ten.

WATCH FOR:
- Students who switch forms confidently - secure.
- Students who take the percent off as dollars - reteach percent of an amount.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Complete the triple for two-fifths: write it as a decimal and a percentage.
- Then say which is bigger, 0.3 or one-quarter, and how you know.

DO:
- Display the prompt.
- Give about 4 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - moving between forms and comparing. 2/5 = 0.4 = 40%; 0.3 is bigger than 1/4 = 0.25.

WATCH FOR:
- Students who give 0.4, 40% and pick 0.3 with reasoning - secure.
- Students who cannot switch forms - revisit with the Week 6 Catch-Up Card.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again, and think back across the whole unit.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: why is it useful to change a percentage into a fraction or decimal?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that fractions, decimals and percentages are three forms of the same amount, and choosing the easiest form makes comparing and discount problems simple.

WATCH FOR:
- Strong thumbs across all three - the unit has landed.
- Sideways or down on the core criterion - small-group revision with the Catch-Up Card.

[Retention and Recall | Mastery and Application]`;

// --- Helpers -----------------------------------------------------------------

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
      fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", margin: 0, fit: "shrink", shrinkText: true,
    });
  }
}

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Week 6 Session 4: Fractions, decimals and percentages together",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Financial reasoning
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Money problems",
      [
        "$50 item, 10% off. Discount = ?",
        "Change from $20 for a $13.50 item?",
        "Round $4.85 to the nearest dollar",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "$5     $6.50     $5", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal subtraction
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal subtraction",
      ["16.4 - 8.7", "20 - 6.25", "11.3 - 4.85"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "7.7        13.75        6.45", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - 25% off
  contentSlide(pres, "Launch", C.ACCENT, "25% off - what does that mean?",
    [
      "25% = 1/4 = 0.25.",
      "",
      "25% off means a quarter comes off.",
      "Same amount, three forms.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      gridPanel(slide, lg, "25 out of 100", 2, 5, "25% = 1/4 = 0.25",
        { strip: C.ACCENT, cardH: 2.9 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to move between fractions, decimals and percentages to solve everyday problems.",
    [
      "I can complete a fraction-decimal-percentage triple.",
      "I can choose the easier form to compare two numbers.",
      "I can solve a discount or comparison problem.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Equivalent forms (F = D = P)",
    [
      "The same amount has three forms.",
      "Fraction = decimal = percentage.",
      "1/2 = 0.5 = 50%.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
      const rows = [
        ["1/2", "0.5", "50%"],
        ["1/4", "0.25", "25%"],
        ["1/10", "0.1", "10%"],
      ];
      const ry0 = lg.panelTopPadded + 0.3;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.6;
        const cw = (lg.rightW - 0.5) / 3;
        r.forEach((cell, j) => {
          slide.addText(cell, {
            x: lg.rightX + 0.25 + j * cw, y: ry, w: cw, h: 0.5,
            fontSize: 18, fontFace: FONT_H, color: j === 0 ? C.PRIMARY : C.CHARCOAL, bold: true,
            align: "center", valign: "middle", margin: 0,
          });
        });
      });
    }
  );

  // Slide 10: I Do #1 - FDP triple for 3/4
  workedExSlide(pres, 2, "I Do", "Complete the triple: 3/4",
    [
      "Fraction: 3/4.",
      "On the grid: 75 of 100 squares.",
      "",
      "Decimal: 75 hundredths = 0.75.",
      "Percentage: 75%.",
      "",
      "So 3/4 = 0.75 = 75%.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      gridPanel(slide, lg, "75 shaded", 7, 5, "3/4 = 0.75 = 75%",
        { strip: C.PRIMARY, cardH: 2.9 });
    }
  );

  // Slide 11: I Do #2 - discount problem
  workedExSlide(pres, 2, "I Do", "Discount: 25% off $40",
    [
      "25% = 1/4 (the easy form).",
      "",
      "1/4 of $40 = $10 off.",
      "$40 - $10 = $30.",
      "",
      "The jacket now costs $30.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.PRIMARY });
      const cx = lg.rightX + lg.rightW / 2;
      slide.addText("$40", {
        x: cx - 1.0, y: lg.panelTopPadded + 0.2, w: 2.0, h: 0.55,
        fontSize: 28, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      addTextOnShape(slide, "25% off = 1/4 of 40 = $10 off", {
        x: lg.rightX + 0.35, y: lg.panelTopPadded + 0.85, w: lg.rightW - 0.7, h: 0.5, rectRadius: 0.06,
        fill: { color: C.SECONDARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
      slide.addText("$40 - $10", {
        x: cx - 1.2, y: lg.panelTopPadded + 1.5, w: 2.4, h: 0.45,
        fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      addTextOnShape(slide, "New price = $30", {
        x: lg.rightX + 0.6, y: lg.panelTopPadded + 2.0, w: lg.rightW - 1.2, h: 0.5, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: CFU + reveal - complete triple for 1/5
  withReveal(
    () => cfuSlide(pres, "CFU", "Complete the triple for 1/5", "Show Me Boards",
      "Complete the triple for 1/5.\n\n1/5  =  ______ (decimal)  =  ______ (percentage)",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "1/5 = 0.2 = 20%   (1/5 = 2/10)", { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - compare 0.7 and 3/4
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Which is bigger: 0.7 or 3/4?",
      [
        "With your partner.",
        "",
        "1.  Change to the same form.",
        "2.  3/4 = ? as a decimal.",
        "3.  Compare the hundredths.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
        slide.addText("0.7", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.35, w: lg.rightW - 0.4, h: 0.6,
          fontSize: 34, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
        });
        slide.addText("vs", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.95, w: lg.rightW - 0.4, h: 0.4,
          fontSize: 18, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", margin: 0,
        });
        slide.addText("3/4", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.35, w: lg.rightW - 0.4, h: 0.6,
          fontSize: 34, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "3/4 = 0.75, and 0.75 > 0.70, so 3/4 is bigger", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - discount $60 at 25% off
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Discount: 25% off a $60 jacket",
      [
        "With your partner.",
        "",
        "1.  25% = which fraction?",
        "2.  Find that fraction of $60.",
        "3.  Subtract to find the new price.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
        slide.addText("$60 jacket", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.3, w: lg.rightW - 0.4, h: 0.55,
          fontSize: 26, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
        });
        addTextOnShape(slide, "25% off", {
          x: lg.rightX + 1.0, y: lg.panelTopPadded + 1.0, w: lg.rightW - 2.0, h: 0.6, rectRadius: 0.08,
          fill: { color: C.ALERT },
        }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText("New price = ?", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.75, w: lg.rightW - 0.4, h: 0.45,
          fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "1/4 of $60 = $15 off,  so new price = $45", { color: C.SUCCESS });
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
      { text: "complete the F-D-P triples.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "compare by changing to one form.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "solve the discount problems.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Fraction = decimal = percentage. Change to the same form to compare. Percent OF an amount = a fraction of it.", {
      x: 0.8, y: panelY + 0.52, w: 8.4, h: 0.55, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

    const benchY = panelY + 1.25;
    const b = ["1/2 = 0.5 = 50%", "1/4 = 0.25 = 25%", "1/5 = 0.2 = 20%"];
    const bw = 2.7;
    b.forEach((t, i) => {
      addTextOnShape(s, t, {
        x: 0.85 + i * (bw + 0.2), y: benchY, w: bw, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 13.5, fontFace: FONT_H, color: C.WHITE, bold: true });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Complete the triple: 2/5 = ____ (decimal) = ____ (percentage).",
      "Which is bigger, 0.3 or 1/4? How do you know?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why is it useful to change a percentage into a fraction or decimal?",
      scItems: [
        "I can complete a fraction-decimal-percentage triple.",
        "I can choose the easier form to compare two numbers.",
        "I can solve a discount or comparison problem.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDec_W6S4_FDP_Together.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Complete FDP triples, compare across forms, and solve discount problems.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "A fraction, a decimal and a percentage can all name the same amount. To compare two numbers, change them to the same form. To find a percentage OF an amount, use the fraction form (25% off = a quarter off). Benchmarks: 1/2 = 0.5 = 50%, 1/4 = 0.25 = 25%, 1/5 = 0.2 = 20%, 1/10 = 0.1 = 10%.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "3/4: on the 100-grid that is 75 of 100, so 0.75 and 75%. So 3/4 = 0.75 = 75%.",
      y);

    y = addSectionHeading(doc, "Section 1 - Complete the triple", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Fill in the two missing forms. The first one is started for you.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  1/2 = 0.5 = 50%        b)  1/4 = ______ = ______        c)  3/4 = ______ = ______", y);
    y = addWriteLine(doc, "d)  ______ = 0.2 = ______   e)  ______ = ______ = 10%        f)  2/5 = ______ = ______", y);

    y = addSectionHeading(doc, "Section 2 - Compare (circle the bigger one)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  0.7  or  3/4        b)  1/2  or  40%        c)  0.3  or  1/4", y);
    y = addWriteLine(doc, "d)  60%  or  0.65       e)  1/5  or  0.25", y);

    y = addSectionHeading(doc, "Section 3 - Discount problems (show your working)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  A $40 jacket is 25% off. New price? ______", y);
    y = addWriteLine(doc, "b)  An $80 scooter is 10% off. New price? ______", y);
    y = addWriteLine(doc, "c)  A $30 game is 50% off. New price? ______", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "A $120 bike is 25% off, then another 10% off the sale price. What is the final price? (Hint: do one step at a time.)", y);
    y = addWriteLine(doc, "Final price: ______", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Fractions, Decimals and Percentages | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the fractions, decimals and percentages practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Complete the triple", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/2 = 0.5 = 50%   b)  1/4 = 0.25 = 25%   c)  3/4 = 0.75 = 75%", y);
    y = addBodyText(doc, "d)  1/5 = 0.2 = 20%   e)  1/10 = 0.1 = 10%   f)  2/5 = 0.4 = 40%", y);

    y = addSectionHeading(doc, "Section 2 - Compare", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3/4 (0.75 > 0.7)   b)  1/2 (50% > 40%)   c)  0.3 (0.3 > 0.25)   d)  0.65 (65% > 60%)   e)  0.25 (0.25 > 0.2)", y);

    y = addSectionHeading(doc, "Section 3 - Discounts", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/4 of 40 = 10 off, so $30.   b)  1/10 of 80 = 8 off, so $72.   c)  1/2 of 30 = 15 off, so $15.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "25% off $120 = $90. Then 10% off $90 = $9 off, so $81. (Note: two discounts are not the same as 35% off.)", y);

    y = addTipBox(doc,
      "Watch for: students who write 3/4 as 0.34; students who take the percent off as dollars; students who add the two discounts instead of applying them in turn.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension Challenge (Week 6)
  await (async () => {
    const doc = createPdf({ title: Y8_RES.name });
    let y = addPdfHeader(doc, Y8_RES.name, {
      subtitle: "Week 6 challenge - aimed beyond Year 6, for students ready to push further.",
      color: C.ASSESS,
      lessonInfo: "Week 6 | Fractions, decimals, percentages | Year 8 reach task",
    });
    y = addTipBox(doc,
      "These problems go past Year 6: percentage of amounts, percentage increase and decrease, finding the whole from a percentage, and recurring decimals. Show all working.",
      y, { color: C.ASSESS });

    y = addSectionHeading(doc, "Part A - Percentage of an amount", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  15% of $80 = ______        b)  30% of 250 = ______        c)  12.5% of 64 = ______", y);

    y = addSectionHeading(doc, "Part B - Increase and decrease", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  Increase $200 by 15% = ______        b)  Decrease 60 kg by 20% = ______", y);

    y = addSectionHeading(doc, "Part C - Find the whole", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  20% of a number is 14. The number is ______", y);
    y = addWriteLine(doc, "b)  A sale price of $45 is 75% of the original. The original price is ______", y);

    y = addSectionHeading(doc, "Part D - Recurring decimals and ordering", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  Write 1/3 and 2/3 as decimals.        b)  Order: 0.65, 2/3, 70%, 0.6", y);

    y = addSectionHeading(doc, "Part E - Reasoning", y, { color: C.ACCENT });
    y = addBodyText(doc, "A shop adds 10% GST to a $60 item, then offers 10% off the new total. Is the final price the same as the original $60? Explain why or why not.", y);
    y = addLinedArea(doc, y + 4, 3, {});

    addPdfFooter(doc, "Week 6 | Year 8 Extension Challenge | Fractions, decimals, percentages");
    await writePdf(doc, path.join(OUT_DIR, Y8_RES.fileName));
    console.log("PDF written: " + Y8_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: Y8_KEY_RES.name });
    let y = addPdfHeader(doc, Y8_KEY_RES.name, {
      subtitle: "Worked answers and look-fors for the Week 6 Year 8 extension challenge.",
      color: C.ASSESS,
      lessonInfo: "Week 6 | Fractions, decimals, percentages | Year 8 reach task",
    });

    y = addSectionHeading(doc, "Part A - Percentage of an amount", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  15% of 80 = 12 (10% = 8, 5% = 4).   b)  30% of 250 = 75.   c)  12.5% of 64 = 8 (1/8 of 64).", y);

    y = addSectionHeading(doc, "Part B - Increase and decrease", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  15% of 200 = 30, so 200 + 30 = $230.   b)  20% of 60 = 12, so 60 - 12 = 48 kg.", y);

    y = addSectionHeading(doc, "Part C - Find the whole", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  20% = 14, so 10% = 7, so 100% = 70.   b)  75% = $45, so 25% = $15, so 100% = $60.", y);

    y = addSectionHeading(doc, "Part D - Recurring decimals and ordering", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/3 = 0.333... and 2/3 = 0.666...   b)  Order: 0.6, 0.65, 2/3 (0.666...), 70% (0.7).", y);

    y = addSectionHeading(doc, "Part E - Reasoning", y, { color: C.ACCENT });
    y = addBodyText(doc, "No. 60 + 10% = 66; then 66 - 10% = 66 - 6.60 = 59.40. The second 10% is taken from the larger 66, so the final price is $59.40, not $60. Look for: correct two-step working, and the insight that the percentages act on different amounts.", y);

    addPdfFooter(doc, "Week 6 | Year 8 Extension Answer Key | Fractions, decimals, percentages");
    await writePdf(doc, path.join(OUT_DIR, Y8_KEY_RES.fileName));
    console.log("PDF written: " + Y8_KEY_RES.fileName);
  })();

  console.log("Week 6 Session 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
