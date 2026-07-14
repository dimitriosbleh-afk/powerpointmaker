"use strict";

// Mastering Fractions, Decimals and Percentages (Year 6 Numeracy) - Lesson 3 of 4.
// Find a percentage of a quantity; percentage discounts and sale prices.
// Friendly percentages: 10% = divide by 10, 25% = divide by 4, 50% = divide by
// 2. Non-friendly via 10%: 30% = find 10%, then x3 (the VC2M6N07 elaboration,
// equivalence to 0.3). Online toy-sale context. VC2M6N07.
// Daily Review: fractions & decimals (prior). Fluency: decimal column addition.
// Unit variant fixed (variant 1) for cohesion across the four lessons.
// CATCH-UP: the launch rebuilds "10% = divide by 10" from a price tag, and
// worksheet Section 1 re-grounds it, so a returner can rejoin at this lesson.

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

const SESSION = 3;
const TOTAL = 4;
const UNIT_TITLE = "Mastering Fractions, Decimals & Percentages";
const FOOTER = `Fractions, Decimals & Percentages | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FDP_Lesson3_Percentage_Discounts";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Find percentages of amounts and work out toy-sale discounts and sale prices.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 3 practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

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

// Stacked step pills used on the discount worked examples.
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
- Welcome back to our fractions, decimals and percentages unit.
- Last lesson we learned that percentages, decimals and fractions can name the same amount.
- Today we put that to work in the shops: finding a percentage of a price, working out a discount, and finding the new sale price.

DO:
- Have whiteboards and markers ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 3 of 4. The anchor is "find 10% first, then scale" - the same find-one-part-then-scale move from Lesson 1, now with percentages.

WATCH FOR:
- Students who find this new - reassure them, finding 10% is just dividing by 10.

[General: Title | Element: Attention, focus and regulation]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet has an online toy sale on it for the You Do. Everything else is whiteboards.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards and markers ready.

TEACHER NOTES:
One student practice sheet with a toy-sale table, and an answer key. Most teaching happens on whiteboards.

CATCH-UP NOTE:
A student who missed earlier lessons can still access today. The launch rebuilds "10 percent means divide by 10" from a price tag, which stands on its own. Section 1 of the worksheet re-grounds it. A returner needs only the worksheet and one minute with you to start.

[General: Resources | Element: Enabling Learning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are warming up our fractions and decimals.
- Read each one and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for how students convert between forms.

TEACHER NOTES:
Daily Review is prior learning. Converting between fractions and decimals, plus a fraction of an amount, keep earlier skills warm and lead naturally into today's percentages.

WATCH FOR:
- Students who convert smoothly - secure.
- Students who stall on 0.6 - prompt them: six tenths.

[Stage 1: Daily Review | Element: Retention and recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Nought point six as a fraction is six tenths, which simplifies to three fifths.
- One quarter of 40 is 10.
- Nought point five plus nought point two five is nought point seven five.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Accept 6/10 or 3/5. The decimal addition keeps the fluency habit alive in a money-friendly form.

WATCH FOR:
- Students who self-correct quickly - secure.

[Stage 1: Daily Review Answers | Element: Retention and recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Adding decimals with the vertical algorithm, money style.
- Set each one out vertically and line up the decimal points.
- Keep dollars under dollars and cents under cents.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up decimal points.

TEACHER NOTES:
Money is decimals lined up at the point. This is exactly the place value habit we use when we add a discount back to check a sale price.

WATCH FOR:
- Students who line up the points - secure.
- Students who lose the cents place - remind them two places after the point for money.

[Stage 1: Fluency | Element: Retention and recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 12.50 plus 7.25 is 19.75.
- 4.99 plus 3.00 is 7.99.
- 8.75 plus 6.50 is 15.25.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
These all keep two decimal places, like money. Watch the carry from cents into dollars in the third one.

WATCH FOR:
- Students who self-correct - secure.
- Students who drop a cents digit - small group focus.

[Stage 1: Fluency Answers | Element: Retention and recall]`;

const NOTES_LAUNCH = `SAY:
- Here is a toy with a price tag of 40 dollars.
- 10 percent means 10 out of every 100, which is one tenth. To find one tenth of an amount, we divide by 10.
- So 10 percent of 40 dollars is 40 divided by 10, which is 4 dollars.
- Finding 10 percent is our friendly first step. Once we have 10 percent, we can build almost any percentage.

DO:
- Point to the price tag and say "10 percent is one tenth, so divide by 10".
- Write 40 divided by 10 equals 4 dollars.
- Have students chorus "10 percent means divide by 10".

TEACHER NOTES:
This launch rebuilds the key move from a single price tag, so it does not depend on earlier lessons. It is the catch-up bridge for the discount work.

WATCH FOR:
- Students who divide by 10 confidently - strong prior knowledge.
- Students unsure - connect back to 10% = 1/10 from last lesson, then divide.

[Stage: Launch | Element: Knowledge and memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to find a percentage of an amount and to work out discounts and sale prices.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Point to the words "10 percent" and "sale price".

TEACHER NOTES:
The first criterion is reachable for everyone - find 10% by dividing by 10. The second is the core target the exit ticket checks. The third stretches to choosing the most efficient method.

WATCH FOR:
- Students who can say "find 10%, then scale" - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- A discount is money taken off the price - the amount you save.
- The sale price is what you actually pay: the original price minus the discount.
- The word of in maths means times. 10 percent of 40 dollars means nought point one times 40.

DO:
- Point to each word as you say it.
- Have students say "sale price equals price minus discount" once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. The most common slip is stopping at the discount and forgetting to subtract it to get the sale price.

WATCH FOR:
- Students who can define discount and sale price - secure.
- Students who confuse the two - anchor with "discount is what you save, sale price is what you pay".

[General: Key Vocabulary | Element: Knowledge and memory]`;

const NOTES_IDO1 = `SAY:
- Let us work through the three friendly percentages on a 40 dollar toy.
- 10 percent is one tenth, so divide by 10: 40 divided by 10 is 4 dollars.
- 25 percent is one quarter, so divide by 4: 40 divided by 4 is 10 dollars.
- 50 percent is one half, so divide by 2: 40 divided by 2 is 20 dollars.
- These three are quick because each one is just a friendly division.

DO:
- Show the tenths bar with one part shaded for 10%.
- Reveal the three pills: divide by 10, divide by 4, divide by 2.
- Have students chorus the three divisions.

TEACHER NOTES:
These three are the toolkit. 10% by dividing by 10 is the most important because it builds the others, like 30% next.

MISCONCEPTIONS:
- Misconception: students multiply by 10 to find 10%.
  Why: percentage feels like it should make the number bigger.
  Impact: 10% of 40 becomes 400.
  Quick correction: 10% is one tenth, a small part, so we divide. Show one part of the tenths bar.

WATCH FOR:
- Students who divide correctly - secure.
- Students who multiply - re-point to the single shaded tenth.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a real discount. A 40 dollar toy has 30 percent off. How much do we save, and what do we pay?
- 30 percent is the same as nought point three. The neat way is to find 10 percent first, then scale.
- 10 percent of 40 is 4 dollars. 30 percent is three lots of that, so 3 times 4 is 12 dollars off.
- The sale price is the price minus the discount: 40 dollars minus 12 dollars is 28 dollars.

DO:
- Shade 3 parts of the tenths bar for 30%.
- Reveal the steps: 10% = $4, then x3 = $12 off, then sale price = $28.
- Circle 28 dollars as the answer to "what do we pay".

TEACHER NOTES:
This is the elaboration method exactly: 30% via its equivalence to 0.3, found by dividing by 10 then multiplying by 3. Always finish by subtracting to get the sale price.

MISCONCEPTIONS:
- Misconception: students give 12 dollars as the final answer and forget to subtract.
  Why: they found the discount and stopped.
  Impact: they report the saving as the price.
  Quick correction: ask "is the toy 12 dollars now, or 28 dollars?" Price minus discount.

WATCH FOR:
- Students who find $12 then subtract to $28 - secure.
- Students who stop at $12 - prompt for the sale price.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Find 10 percent of 60 dollars. Then find 30 percent.

DO:
- Display the prompt.
- Give 45 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 10% = 6 dollars, then 30% = 3 x 6 = 18 dollars.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students multiply 60 by 10, or stop at 10%.
- Reteach: 10% is divide by 10, so 6 dollars. 30% is three lots of 10%, so 3 times 6 is 18 dollars.
- Re-check: what is 10% of 60, and how many lots of it is 30%?

TEACHER NOTES:
The two-step "find 10%, then scale" is the move to watch. A student who writes 600 has multiplied instead of divided.

WATCH FOR:
- Students who write 6 then 18 - secure.
- Students who stop at 6 - they found 10% but did not scale, prompt them.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. A 24 dollar game has 25 percent off.
- 25 percent is one quarter. What is the friendly division?
- Whisper to your partner: what is 24 divided by 4?

DO:
- Display the 24 dollar game, 25% off, above the quarters bar.
- Give 75 seconds.
- Listen for "divide by 4".

TEACHER NOTES:
25% is a friendly one - divide by 4. 24 divided by 4 is 6 dollars off. Then sale price is 24 minus 6.

WATCH FOR:
- Pairs who divide by 4 - secure.
- Pairs who try 10% first - that works too (10% = 2.40, x2.5), but the quarter is quicker; let both stand.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- 25 percent of 24 dollars is 24 divided by 4, which is 6 dollars off.
- The sale price is 24 minus 6, which is 18 dollars.
- Discount 6 dollars, you pay 18 dollars.

DO:
- Click to reveal.
- Restate: discount is what you save, sale price is what you pay.

TEACHER NOTES:
If a pair used 10% then scaled, celebrate it - choosing a method is the depth criterion.

WATCH FOR:
- Students who self-correct - secure.
- Students who report 6 dollars as the price - remind them to subtract.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together. A 50 dollar scooter has 30 percent off.
- Use our neat method: find 10 percent first, then scale to 30 percent.
- Then do not forget the last step - the sale price.

DO:
- Display the 50 dollar scooter, 30% off, above the tenths bar.
- Give 90 seconds.
- Watch for students who find the discount but forget the sale price.

TEACHER NOTES:
10% of 50 is 5 dollars. 30% is 3 times 5, which is 15 dollars off. Sale price is 50 minus 15, which is 35 dollars.

WATCH FOR:
- Students who write 35 dollars - secure.
- Students who stop at 15 dollars - prompt for the sale price.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- 10 percent of 50 dollars is 5 dollars. 30 percent is 3 times 5, which is 15 dollars off.
- The sale price is 50 minus 15, which is 35 dollars.
- Find 10 percent, scale, then subtract. Same recipe every time.

DO:
- Click to reveal.
- Run the three steps once more together.

TEACHER NOTES:
If many stopped at the discount, do one more quick "and what do you pay" before the You Do.

WATCH FOR:
- Students who complete all three steps - ready for independent work.
- Students who stop at the discount - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- There is an online toy sale on it. Work out the discount and the sale price for each toy.
- Section 1 starts with finding 10 percent. Section 2 is the sale. If you finish, try the challenge.

DO:
- Distribute the practice sheet.
- Circulate and listen for "find 10%, then scale, then subtract".
- Cold call one or two students to explain a sale price.

TEACHER NOTES:
The toy-sale table uses 10%, 25%, 50% and 30% discounts - the friendly ones plus the build-from-10% one. Different prices from the We Do, same method.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 is started - find 10% of each price by dividing by 10. This is also the rebuild for any returning student.
- Extra Notes: Sit with these students and divide the first price by 10 together.
EXTENDING PROMPT:
- Task: The challenge box asks for a 15% discount (find 10%, then half of it again) and to compare two sale offers to decide the better buy.
- Extra Notes: Push them to justify which toy is the better deal and why.

WATCH FOR:
- Students who complete all three steps - secure.
- Students who multiply to find 10% - re-anchor with one tenth.

[Stage 4: You Do | Element: Mastery and application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- A 20 dollar toy has 30 percent off.
- Find the discount and the sale price.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - find a percentage discount and the sale price. Look for 10% = 2 dollars, 30% = 6 dollars off, sale price 14 dollars. The SC target is SC2.

WATCH FOR:
- Students who write $6 off and $14 to pay - secure.
- Students who write $6 only - they forgot the sale price, revisit at the start of Lesson 4.

[Stage 5: Exit Ticket | Element: Mastery and application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: what are the steps to find a 30 percent discount and the sale price?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is "find 10%, scale, then subtract for the sale price". Students who can say this are ready to estimate and check answers next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on the core criterion - small group revision at the start of Lesson 4.

[General: Closing | Element: Retention and recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 3: Percentage of a quantity & discounts",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - fractions & decimals
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions & decimals",
      [
        "Write 0.6 as a fraction.",
        "Find 1/4 of 40.",
        "0.5 + 0.25 = ?",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "6/10 = 3/5          10          0.75", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal column addition (money)
  withReveal(
    () => fluencySlide(pres, "Fluency: Adding decimals",
      ["12.50 + 7.25", "4.99 + 3.00", "8.75 + 6.50"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "19.75        7.99        15.25", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - 10% = divide by 10 (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "10% means divide by 10",
    [
      "A toy is priced at $40.",
      "10% = 10 out of 100 = one tenth.",
      "One tenth means divide by 10.",
      "10% of $40 = $40 ÷ 10 = $4.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.ACCENT });
      slide.addText("Find 10% first", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      // Price tag
      addTextOnShape(slide, "$40", {
        x: lg.rightX + 1.30, y: lg.panelTopPadded + 0.52, w: 1.6, h: 0.70, rectRadius: 0.10,
        fill: { color: C.PRIMARY },
      }, { fontSize: 30, fontFace: FONT_H, color: C.WHITE, bold: true });
      // Tenths bar, one tenth shaded
      addTenthsStrip(slide, lg.rightX + 0.25, lg.panelTopPadded + 1.45, 3.0, 1,
        { fillColor: C.SECONDARY });
      addTextOnShape(slide, "$40 ÷ 10 = $4", {
        x: lg.rightX + 0.45, y: lg.panelTopPadded + 2.10, w: lg.rightW - 0.90, h: 0.52, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to find a percentage of an amount and to work out discounts and sale prices.",
    [
      "I can find 10% of an amount by dividing by 10.",
      "I can find a percentage discount (like 25% or 30%) and work out the sale price.",
      "I can choose an efficient method and explain my discount thinking.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Discount = money OFF the price (what you save)",
      "Sale price = price - discount (what you pay)",
      "'of' = times (10% of $40 = 0.1 x 40)",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.1, { strip: C.SECONDARY });
      slide.addText("The recipe", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      stepPills(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.55, lg.rightW - 0.60, [
        { text: "1. Find the percentage", color: C.PRIMARY },
        { text: "2. Price - discount = sale price", color: C.SUCCESS },
      ], { h: 0.56, gap: 0.14, fontSize: 14.5 });
    }
  );

  // Slide 10: I Do #1 - friendly percentages of $40
  workedExSlide(pres, 2, "I Do", "Friendly percentages of $40",
    [
      "10% = ÷ 10  ->  $4",
      "25% = ÷ 4  ->  $10",
      "50% = ÷ 2  ->  $20",
      "",
      "Each one is a friendly division.",
      "10% is the most useful - it builds the rest.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.15, { strip: C.PRIMARY });
      slide.addText("10% = one tenth of $40", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addTenthsStrip(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.50, 3.0, 1,
        { fillColor: C.SECONDARY });
      stepPills(slide, lg.rightX + 0.30, lg.panelTopPadded + 1.15, lg.rightW - 0.60, [
        { text: "10% = $40 ÷ 10 = $4", color: C.SUCCESS },
        { text: "25% = $40 ÷ 4 = $10", color: C.PRIMARY },
        { text: "50% = $40 ÷ 2 = $20", color: C.SECONDARY },
      ], { h: 0.52, gap: 0.10, fontSize: 15.5 });
    }
  );

  // Slide 11: I Do #2 - 30% off $40
  workedExSlide(pres, 2, "I Do", "30% off a $40 toy",
    [
      "30% = 0.3.  Find 10% first.",
      "10% of $40 = $4.",
      "30% = 3 x $4 = $12 OFF.",
      "",
      "Sale price = $40 - $12 = $28.",
      "Always finish with the sale price.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("30% = 3 lots of 10%", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addTenthsStrip(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.50, 3.0, 3,
        { fillColor: C.SECONDARY });
      stepPills(slide, lg.rightX + 0.30, lg.panelTopPadded + 1.15, lg.rightW - 0.60, [
        { text: "10% = $4", color: C.PRIMARY },
        { text: "30% = 3 x $4 = $12 off", color: C.SECONDARY },
        { text: "Sale price = $40 - $12 = $28", color: C.SUCCESS },
      ], { h: 0.52, gap: 0.10, fontSize: 15 });
    }
  );

  // Slides 12-13: CFU + reveal - 10% then 30% of $60
  withReveal(
    () => cfuSlide(pres, "CFU", "Find 10% of $60, then 30%",
      { technique: "Show Me Boards",
        question: "On your whiteboard:\n\n10% of $60 = ?     then     30% of $60 = ?" },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "10% = $6,  so 30% = 3 x $6 = $18", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 25% off $24
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "25% off a $24 game",
      [
        "With your partner.",
        "",
        "1.  25% = one quarter = ÷ 4.",
        "2.  Find the discount.",
        "3.  Work out the sale price.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.SECONDARY });
        slide.addText("$24 game, 25% off", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.40,
          fontSize: 20, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        slide.addText("25% = one quarter. Share into 4 parts.", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.56, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
          align: "center", margin: 0,
        });
        partBar(slide, lg.rightX + 0.20, lg.panelTopPadded + 0.94, lg.rightW - 0.40, 0.72,
          4, 1, { fill: C.SECONDARY });
        slide.addText("One quarter is the discount.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.80, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "25% of $24 = $6 off.  Sale price = $24 - $6 = $18", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - 30% off $50
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "30% off a $50 scooter",
      [
        "With your partner.",
        "",
        "1.  Find 10% of $50.",
        "2.  Scale to 30% (x 3).",
        "3.  Sale price = price - discount.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.SECONDARY });
        slide.addText("$50 scooter, 30% off", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.40,
          fontSize: 19, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        slide.addText("30% = 3 lots of 10%.", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.56, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
          align: "center", margin: 0,
        });
        addTenthsStrip(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.98, 3.0, 3,
          { fillColor: C.SECONDARY });
        slide.addText("Three tenths shaded = 30%.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.70, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "10% = $5, 30% = $15 off.  Sale price = $50 - $15 = $35", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // Slide 18: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: the online toy sale", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "find 10% of each price (÷10).   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "work out the discount.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "price - discount = sale price.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Find 10%, scale, then subtract.  Sale price = price - discount.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Friendly-percentage reference pills
    const refs = [["10% = ÷10", C.SUCCESS], ["25% = ÷4", C.PRIMARY], ["50% = ÷2", C.SECONDARY]];
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
      "A $20 toy has 30% off.",
      "Find the discount and the sale price.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what are the steps to find a 30% discount and the sale price?",
      scItems: [
        "I can find 10% of an amount by dividing by 10.",
        "I can find a percentage discount (like 25% or 30%) and work out the sale price.",
        "I can choose an efficient method and explain my discount thinking.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FDP_Lesson3_Percentage_Discounts.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find percentages of amounts and work out toy-sale discounts and sale prices.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Find 10% by dividing by 10. Build other percentages from it: 30% = 3 lots of 10%. Friendly ones: 25% = divide by 4, 50% = divide by 2. Always finish: sale price = price - discount.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "30% off a $60 toy. 10% of 60 = $6. 30% = 3 x $6 = $18 off. Sale price = $60 - $18 = $42.",
      y);

    y = addSectionHeading(doc, "Section 1 - Find 10% first (started for you)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  10% of $30 = $30 ÷ 10 = $3", y);
    y = addWriteLine(doc, "b)  10% of $50 = $50 ÷ 10 = _______", y);
    y = addWriteLine(doc, "c)  10% of $80 = _______", y);
    y = addWriteLine(doc, "d)  10% of $120 = _______", y);

    y = addSectionHeading(doc, "Section 2 - The online toy sale (discount and sale price)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Robot $40, 10% off   |   Puzzle $24, 25% off   |   Bike $80, 50% off   |   Drone $50, 30% off", y);
    y = addWriteLine(doc, "a)  Robot:  discount = _______   sale price = _______", y);
    y = addWriteLine(doc, "b)  Puzzle: discount = _______   sale price = _______", y);
    y = addWriteLine(doc, "c)  Bike:   discount = _______   sale price = _______", y);
    y = addWriteLine(doc, "d)  Drone:  discount = _______   sale price = _______", y);

    y = addSectionHeading(doc, "Challenge (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "A $60 toy has 15% off. (Hint: find 10%, then half of it again for the extra 5%.) Discount = _______  Sale price = _______", y);
    y = addWriteLine(doc, "Better buy: Toy A $50 with 30% off, or Toy B $45 with 20% off? Show why.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Percentage Discounts | Year 6 Numeracy`);
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

    y = addSectionHeading(doc, "Section 1 - Find 10%", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  $3        b)  $5        c)  $8        d)  $12", y);

    y = addSectionHeading(doc, "Section 2 - The online toy sale", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Robot: 10% of $40 = $4 off; sale price $36.", y);
    y = addBodyText(doc, "b)  Puzzle: 25% of $24 = $24 ÷ 4 = $6 off; sale price $18.", y);
    y = addBodyText(doc, "c)  Bike: 50% of $80 = $80 ÷ 2 = $40 off; sale price $40.", y);
    y = addBodyText(doc, "d)  Drone: 10% of $50 = $5; 30% = 3 x $5 = $15 off; sale price $35.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "15% of $60: 10% = $6, 5% = $3, so 15% = $9 off; sale price $51.", y);
    y = addBodyText(doc, "Better buy: Toy A 30% off $50 = $15 off, pay $35. Toy B 20% off $45 = $9 off, pay $36. Toy A is cheaper by $1.", y);

    y = addTipBox(doc,
      "Watch for: students who multiply to find 10% (10% of 40 = 400); students who give the discount as the final answer and forget the sale price; students who subtract the percentage from the price (40 - 30) instead of the dollar discount.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
