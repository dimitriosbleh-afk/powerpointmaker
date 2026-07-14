"use strict";

// Mastering Fractions, Decimals and Percentages (Year 6 Numeracy) - Lesson 2 of 4.
// Connect familiar fractions, decimals and percentages. Big ideas: percent
// means "out of 100"; 100% is the whole; 1/2 = 0.5 = 50% (same amount, three
// names). Use percentages to compare relative size. VC2M6N07 / VC2M5N04.
// Daily Review: operating with fractions (prior). Fluency: decimal column
// addition (unit-wide fluency focus).
// Unit variant fixed (variant 1) for cohesion with Lessons 1, 3 and 4.
// CATCH-UP: the launch rebuilds "percent = out of 100" from a 100-grid, and
// worksheet Section 1 re-grounds the same idea, so a returner can rejoin here.

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
  addAreaModel, addTenthsStrip,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics, getContrastColor,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 4;
const UNIT_TITLE = "Mastering Fractions, Decimals & Percentages";
const FOOTER = `Fractions, Decimals & Percentages | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FDP_Lesson2_Connecting_FDP";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Match familiar fractions, decimals and percentages, and compare sizes.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 2 practice sheet.");
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

// Three-name equivalence row: fraction | decimal | percent pills.
function equivRow(slide, x, y, frac, dec, pct, opts) {
  const o = opts || {};
  const colW = o.colW || 1.2;
  const gap = o.gap || 0.12;
  const h = o.h || 0.46;
  const fs = o.fontSize || 15;
  const cells = [[frac, C.PRIMARY], [dec, C.SECONDARY], [pct, C.SUCCESS]];
  cells.forEach((c, i) => {
    addTextOnShape(slide, c[0], {
      x: x + i * (colW + gap), y, w: colW, h, rectRadius: 0.07,
      fill: { color: c[1] },
    }, { fontSize: fs, fontFace: FONT_H, color: C.WHITE, bold: true });
  });
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to our fractions, decimals and percentages unit.
- Last lesson we found a fraction of an amount by finding one part then scaling.
- Today we learn that a fraction, a decimal and a percentage can be three names for the very same amount, and we use percentages to compare sizes.

DO:
- Have whiteboards and markers ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 2 of 4. The anchor today is "same amount, three names" and "percent means out of 100". This unlocks the discount work in Lesson 3.

WATCH FOR:
- Students who find this new - reassure them, we build it from a 100 grid everyone can see.

[General: Title | Element: Attention, focus and regulation]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end. Everything else is whiteboards and partner talk.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards and markers ready.

TEACHER NOTES:
One student practice sheet and an answer key. Most teaching happens on whiteboards with a drawn 100 grid.

CATCH-UP NOTE:
A student who missed Lesson 1 can still access today. The launch rebuilds "percent means out of 100" from a 100 grid, which does not depend on Lesson 1. Section 1 of the worksheet re-grounds the same idea. A returner needs only the worksheet and one minute with you to start.

[General: Resources | Element: Enabling Learning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are warming up our work operating with fractions.
- Read each one and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for how students handle the same denominators.

TEACHER NOTES:
Daily Review is prior learning, not today's connections. Adding and subtracting fractions with the same denominator, plus a fraction of an amount from Lesson 1, keep fraction sense ticking over.

WATCH FOR:
- Students who keep the denominator the same when adding - secure.
- Students who add denominators too - reteach: the parts stay the same size.

[Stage 1: Daily Review | Element: Retention and recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- One quarter plus two quarters is three quarters.
- Five sixths minus one sixth is four sixths, which simplifies to two thirds.
- One half of 12 is 6.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
For the subtraction, accept 4/6 or 2/3. Note any student who changed the denominator when adding.

WATCH FOR:
- Students who self-correct quickly - secure.

[Stage 1: Daily Review Answers | Element: Retention and recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Adding decimals with the vertical algorithm.
- Set each one out vertically and line up the decimal points.
- Tenths under tenths, hundredths under hundredths.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up decimal points.

TEACHER NOTES:
Same fluency every lesson this unit. Lining up the point is the place value habit we reuse with money in Lesson 3.

WATCH FOR:
- Students who line up the points - secure.
- Students who line up the last digit instead - reteach: points under points.

[Stage 1: Fluency | Element: Retention and recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 3.4 plus 5.8 is 9.2.
- 7.25 plus 2.5 is 9.75.
- 6.7 plus 0.45 is 7.15.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
In the second and third, a missing place can trip students up. Remind them they can write a zero to fill an empty place, for example 2.5 as 2.50.

WATCH FOR:
- Students who self-correct - secure.
- Students who misalign 0.45 - small group focus.

[Stage 1: Fluency Answers | Element: Retention and recall]`;

const NOTES_LAUNCH = `SAY:
- Here is a grid of 100 small squares. The whole grid is one whole.
- If every square is shaded, that is 100 out of 100. We call that 100 percent. Percent means out of 100.
- So 100 percent is just another way of saying the whole, all of it.
- Now I shade half the grid - 50 squares. That is 50 out of 100, which is 50 percent. And it is also one half. Same amount, two names.

DO:
- Point to the full grid and say "100 out of 100 is 100 percent, the whole".
- Shade or reveal 50 squares and say "50 out of 100 is 50 percent, which is one half".
- Have students chorus "percent means out of 100".

TEACHER NOTES:
This launch builds the core idea from a 100 grid that every student can see, regardless of whether they were here for Lesson 1. It is the catch-up bridge for the unit's second half.

WATCH FOR:
- Students who say "50 percent is a half" - strong connection.
- Students unsure - keep pointing at the count out of 100.

[Stage: Launch | Element: Knowledge and memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to connect familiar fractions, decimals and percentages, and to use percentages to compare size.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Point to the words "out of 100" and "compare".

TEACHER NOTES:
The first criterion is reachable for everyone - percent means out of 100. The second is the core target the exit ticket checks. The third stretches to comparing and ordering with percentages.

WATCH FOR:
- Students who can say "same amount, three names" - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- Percent means per hundred, or out of 100. The percent sign is a reminder of the two zeros in 100.
- 100 percent means the whole, everything.
- Equivalent means the same value, even when it is written as a fraction, a decimal or a percentage.

DO:
- Point to each word as you say it.
- Have students say "percent means out of 100" once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. Keep it to these three. The percent sign with its two little zeros is a useful memory hook for "out of 100".

WATCH FOR:
- Students who can explain 100 percent as the whole - secure.
- Students who think a percentage is always small - show 100 percent and beyond is possible.

[General: Key Vocabulary | Element: Knowledge and memory]`;

const NOTES_IDO1 = `SAY:
- Let us work through one together. I want to show that one half, nought point five, and 50 percent are the same amount.
- Here is one whole split into two equal parts. One part shaded is one half.
- That one shaded part is also five tenths, and five tenths written as a decimal is nought point five.
- And out of 100, half is 50, so it is 50 percent. Three names, one amount.

DO:
- Draw the two-part bar and shade one part.
- Write the three names underneath: 1/2, then 0.5, then 50%.
- Say "same amount, three names" as you finish.

TEACHER NOTES:
Use the bar from Lesson 1 so the representation is familiar. The point is equivalence: we are not changing the amount, only how we name it.

MISCONCEPTIONS:
- Misconception: students read 0.5 as "nought point five" and think it is smaller than 50 percent.
  Why: 0.5 looks like a small number and 50 looks big.
  Impact: they rank 0.5 below 50 percent.
  Quick correction: both are half of the same whole - show them on the one bar.

WATCH FOR:
- Students who match all three names - secure.
- Students who treat them as different sizes - return to the single shaded bar.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Here are the familiar ones worth knowing by heart.
- One tenth is nought point one is 10 percent. One fifth is nought point two is 20 percent.
- One quarter is nought point two five is 25 percent. One half is nought point five is 50 percent. Three quarters is nought point seven five is 75 percent.
- One that is not exact: one third is about 33 and a third percent. We keep that one as "about a third".

DO:
- Reveal the table row by row.
- Have students read each trio across: fraction, decimal, percent.
- Flag the thirds row as the "not exact" one.

TEACHER NOTES:
These familiar equivalents are the toolkit for discounts next lesson. The thirds row shows that not every fraction gives a tidy percentage, which is honest and prepares students for estimation in Lesson 4.

MISCONCEPTIONS:
- Misconception: students convert 1/4 to 14 percent by reading the digits.
  Why: they glue the numerator and denominator together.
  Impact: wrong percentage every time.
  Quick correction: one quarter of 100 is 25, so 25 percent. Use the 100 grid.

WATCH FOR:
- Students who recall the trios quickly - ready to compare.
- Students who guess - send them back to "what is this fraction of 100".

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Write one quarter as a decimal and as a percentage.

DO:
- Display the prompt.
- Give 45 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 0.25 and 25 percent.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students write 0.4 or 14 percent.
- Reteach: one quarter of 100 squares is 25 squares, so 25 percent. As a decimal that is 25 hundredths, nought point two five.
- Re-check: how many squares out of 100 is one quarter?

TEACHER NOTES:
The trap is gluing digits or guessing the decimal. Anchor everything to "how many out of 100".

WATCH FOR:
- Students who write 0.25 and 25 percent - secure.
- Students who write 14 percent - the glue-the-digits error, reteach with the grid.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. 30 squares of the 100 grid are shaded.
- What is that as a percentage, as a decimal, and as a fraction?
- Whisper to your partner: out of 100, how many are shaded?

DO:
- Display the shaded 30/100 grid.
- Give 75 seconds.
- Listen for "30 out of 100".

TEACHER NOTES:
Reading straight off the grid keeps the meaning visible. 30 out of 100 is 30 percent, which is 0.30 and 30/100 (or 3/10).

WATCH FOR:
- Pairs who say 30 percent - secure.
- Pairs who miscount the columns - prompt them to count full columns of ten.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- 30 out of 100 is 30 percent.
- As a decimal that is nought point three. As a fraction it is 30 hundredths, which simplifies to three tenths.
- Same amount, three names again.

DO:
- Click to reveal.
- Point to the three full columns as "three tenths".

TEACHER NOTES:
Accept 30/100 or 3/10. The grid makes the simplification visible - three full columns of ten.

WATCH FOR:
- Students who self-correct - secure.
- Students who write 0.30 as "thirty" - remind them it is three tenths.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, and this time we compare.
- Which is larger: two fifths, or 35 percent?
- The trick is to put them in the same language. Turn two fifths into a percentage first.

DO:
- Display the two amounts: 2/5 and 35%.
- Give 90 seconds.
- Listen for "two fifths is 40 percent".

TEACHER NOTES:
This is the "use percentages to compare" criterion. Two fifths is 40 out of 100, which is 40 percent, and 40 percent beats 35 percent. Converting to a common form is the key move.

WATCH FOR:
- Students who convert 2/5 to 40 percent - secure.
- Students who pick 35 percent because 35 looks bigger than the small numbers in 2/5 - reteach: change to the same language first.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Two fifths is the same as 40 hundredths, so 40 percent.
- 40 percent is larger than 35 percent.
- So two fifths is larger. Comparing is easy once both are percentages.

DO:
- Click to reveal.
- Restate the move: change to the same language, then compare.

TEACHER NOTES:
If students struggled, do one more quick comparison, for example 1/4 versus 30 percent, before the You Do.

WATCH FOR:
- Students who compare confidently - ready for independent work.
- Students who compare without converting - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 warms up by reading amounts off a 100 grid.
- Section 2 matches familiar fractions, decimals and percentages.
- Section 3 compares sizes. If you finish, try the challenge with the thirds.

DO:
- Distribute the practice sheet.
- Circulate and listen for "out of 100" and "same amount, three names".
- Cold call one or two students to justify a comparison.

TEACHER NOTES:
Different amounts from the We Do, same ideas: percent is out of 100, equivalents are three names, compare by converting to the same form.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 gives a shaded 100 grid for each amount. Read straight off the grid. This is also the rebuild for any returning student.
- Extra Notes: Sit with these students and count full columns of ten together.
EXTENDING PROMPT:
- Task: The challenge box uses one third and two thirds, which are not exact. Students estimate the percentage and explain why it is "about" not "exactly".
- Extra Notes: Connect this to Lesson 4 estimation.

WATCH FOR:
- Students who convert fluently - secure.
- Students who glue digits - re-anchor with the 100 grid.

[Stage 4: You Do | Element: Mastery and application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- Write three quarters as a decimal and as a percentage.

DO:
- Display the prompt.
- Give 2 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - match a familiar fraction to its decimal and percentage. Look for 0.75 and 75 percent. The SC target is SC2.

WATCH FOR:
- Students who write 0.75 and 75 percent - secure.
- Students who write 0.34 or 34 percent - the glue-the-digits error, revisit at the start of Lesson 3.

[Stage 5: Exit Ticket | Element: Mastery and application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: what does percent mean, and what is 100 percent?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is "percent means out of 100, and familiar fractions, decimals and percentages are three names for the same amount". This sets up discounts next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on the core criterion - small group revision at the start of Lesson 3.

[General: Closing | Element: Retention and recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Connecting fractions, decimals & percentages",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - operating with fractions
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions",
      [
        "1/4 + 2/4 = ?",
        "5/6 - 1/6 = ?",
        "Find 1/2 of 12.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3/4          4/6 = 2/3          6", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal column addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Adding decimals",
      ["3.4 + 5.8", "7.25 + 2.5", "6.7 + 0.45"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "9.2        9.75        7.15", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - percent means out of 100 (100-grid, catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "Percent means out of 100",
    [
      "100 small squares = one whole.",
      "All shaded = 100% = the whole.",
      "50 shaded = 50% = one half.",
      "Today: same amount, three names.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.05, { strip: C.ACCENT });
      slide.addText("50 out of 100 = 50%", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      // 100-grid, 50 shaded (5 full columns)
      const gSize = 2.0;
      const gX = lg.rightX + (lg.rightW - gSize) / 2;
      addAreaModel(slide, gX, lg.panelTopPadded + 0.52, gSize, 5, 0, { fillColor: C.SECONDARY });
      slide.addText("Half the whole = 50% = 0.5", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.62, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to connect familiar fractions, decimals and percentages, and to use percentages to compare size.",
    [
      "I can explain that a percentage means out of 100 and that 100% is the whole.",
      "I can match a familiar fraction to its decimal and percentage (like 1/2 = 0.5 = 50%).",
      "I can use percentages to compare and order amounts, and explain which is larger.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Percent = per hundred = out of 100",
      "100% = the whole (everything)",
      "Equivalent = same value, shown a different way",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.SECONDARY });
      slide.addText("Same amount, three names", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      equivRow(slide, lg.rightX + 0.35, lg.panelTopPadded + 0.62, "1/2", "0.5", "50%",
        { colW: 1.0, gap: 0.14, h: 0.55, fontSize: 16 });
      slide.addText("A fraction, a decimal, a percentage.", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.32, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 10: I Do #1 - 1/2 = 0.5 = 50%
  workedExSlide(pres, 2, "I Do", "1/2 = 0.5 = 50%",
    [
      "One whole split into 2 equal parts.",
      "One part shaded = one half.",
      "",
      "That part is also 5 tenths = 0.5.",
      "Out of 100, half is 50 = 50%.",
      "Same amount, three names.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("One half, three ways", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      partBar(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.58, lg.rightW - 0.50, 0.80,
        2, 1, { fill: C.SUCCESS });
      equivRow(slide, lg.rightX + 0.30, lg.panelTopPadded + 1.70, "1/2", "0.5", "50%",
        { colW: 1.05, gap: 0.12, h: 0.60, fontSize: 17 });
      slide.addText("Not three sizes - one amount.", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.45, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 11: I Do #2 - the familiar equivalence table
  workedExSlide(pres, 2, "I Do", "Familiar ones to know",
    [
      "Learn these by heart.",
      "Read each row across.",
      "",
      "1/3 is about 33 1/3% - not exact.",
      "We keep that one as 'about a third'.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.35, { strip: C.PRIMARY });
      // Header
      const cols = ["Fraction", "Decimal", "Percent"];
      const colW = 1.05, gap = 0.12;
      const tableX = lg.rightX + 0.30;
      cols.forEach((c, i) => {
        slide.addText(c, {
          x: tableX + i * (colW + gap), y: lg.panelTopPadded + 0.10, w: colW, h: 0.26,
          fontSize: 11.5, fontFace: FONT_B, color: C.PRIMARY, bold: true,
          align: "center", margin: 0,
        });
      });
      const rows = [
        ["1/10", "0.1", "10%"],
        ["1/5", "0.2", "20%"],
        ["1/4", "0.25", "25%"],
        ["1/2", "0.5", "50%"],
        ["3/4", "0.75", "75%"],
      ];
      const ry0 = lg.panelTopPadded + 0.44;
      rows.forEach((r, i) => {
        equivRow(slide, tableX, ry0 + i * 0.54, r[0], r[1], r[2],
          { colW, gap, h: 0.46, fontSize: 14 });
      });
    }
  );

  // Slides 12-13: CFU + reveal - 1/4 as decimal and percent
  withReveal(
    () => cfuSlide(pres, "CFU", "Write 1/4 as a decimal and a percentage",
      { technique: "Show Me Boards",
        question: "On your whiteboard:\n\n1/4 = ____ (decimal) = ____% " },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1/4 = 0.25 = 25%   (one quarter of 100 squares is 25)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 30/100
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Read 30 out of 100",
      [
        "With your partner.",
        "",
        "1.  How many out of 100 are shaded?",
        "2.  Write it as a percentage.",
        "3.  Then as a decimal and a fraction.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
        slide.addText("How much is shaded?", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        const gSize = 2.05;
        const gX = lg.rightX + (lg.rightW - gSize) / 2;
        addAreaModel(slide, gX, lg.panelTopPadded + 0.55, gSize, 3, 0, { fillColor: C.SECONDARY });
        slide.addText("Count the full columns of ten.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.68, w: lg.rightW - 0.30, h: 0.28,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "30 out of 100 = 30% = 0.3 = 3/10", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - compare 2/5 vs 35%
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Which is larger?",
      [
        "With your partner.",
        "",
        "Compare:  2/5  or  35% ?",
        "",
        "Tip: change 2/5 into a percentage first.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.SECONDARY });
        slide.addText("Same language, then compare", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addTextOnShape(slide, "2/5", {
          x: lg.rightX + 0.45, y: lg.panelTopPadded + 0.60, w: 1.35, h: 0.70, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText("vs", {
          x: lg.rightX + 1.85, y: lg.panelTopPadded + 0.60, w: 0.55, h: 0.70,
          fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        addTextOnShape(slide, "35%", {
          x: lg.rightX + 2.40, y: lg.panelTopPadded + 0.60, w: 1.35, h: 0.70, rectRadius: 0.08,
          fill: { color: C.ACCENT },
        }, { fontSize: 24, fontFace: FONT_H, color: getContrastColor(C.ACCENT), bold: true });
        slide.addText("Turn 2/5 into hundredths to compare.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.55, w: lg.rightW - 0.30, h: 0.50,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "2/5 = 40%, so 2/5 is larger than 35%", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "read amounts off the 100 grid.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "match fraction, decimal and percent.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "compare two amounts and explain.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Percent means out of 100.  Same amount, three names.  Compare in the same language.", {
      x: 0.9, y: panelY + 0.52, w: 8.2, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Reference equivalence row
    equivRow(s, 3.0, panelY + 1.20, "1/4 = 0.25 = 25%", "1/2 = 0.5 = 50%", "3/4 = 0.75 = 75%",
      { colW: 1.3, gap: 0.10, h: 0.5, fontSize: 11 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Write 3/4 as a decimal and as a percentage.",
      "Use the 100 grid idea if it helps.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what does percent mean, and what is 100%?",
      scItems: [
        "I can explain that a percentage means out of 100 and that 100% is the whole.",
        "I can match a familiar fraction to its decimal and percentage (like 1/2 = 0.5 = 50%).",
        "I can use percentages to compare and order amounts, and explain which is larger.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FDP_Lesson2_Connecting_FDP.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Match familiar fractions, decimals and percentages, and compare sizes.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Percent means out of 100. A fraction, a decimal and a percentage can be three names for the same amount. To compare, change both to the same form (usually a percentage) first.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Familiar equivalents to remember", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "1/10 = 0.1 = 10%     1/5 = 0.2 = 20%     1/4 = 0.25 = 25%     1/2 = 0.5 = 50%     3/4 = 0.75 = 75%",
      y);

    y = addSectionHeading(doc, "Section 1 - Read off the 100 grid (started for you)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  60 squares shaded =  60%  =  0.6  =  6/10", y);
    y = addWriteLine(doc, "b)  25 squares shaded =  ____%  =  ____  =  ____", y);
    y = addWriteLine(doc, "c)  70 squares shaded =  ____%  =  ____  =  ____", y);

    y = addSectionHeading(doc, "Section 2 - Match the three names", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2 = ____ (decimal) = ____%", y);
    y = addWriteLine(doc, "b)  1/4 = ____ (decimal) = ____%", y);
    y = addWriteLine(doc, "c)  1/10 = ____ (decimal) = ____%", y);
    y = addWriteLine(doc, "d)  3/4 = ____ (decimal) = ____%", y);
    y = addWriteLine(doc, "e)  1/5 = ____ (decimal) = ____%", y);

    y = addSectionHeading(doc, "Section 3 - Compare (circle the larger)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2   or   40%        b)  3/4   or   70%        c)  1/5   or   25%", y);

    y = addSectionHeading(doc, "Challenge (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "One third is about 33 1/3% and two thirds is about 66 2/3%. Why are these 'about' and not exact? Use the idea of out of 100.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Connecting FDP | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 2 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Read off the 100 grid", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  60%, 0.6, 6/10        b)  25%, 0.25, 25/100 (= 1/4)        c)  70%, 0.7, 7/10", y);

    y = addSectionHeading(doc, "Section 2 - Match the three names", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  0.5, 50%        b)  0.25, 25%        c)  0.1, 10%        d)  0.75, 75%        e)  0.2, 20%", y);

    y = addSectionHeading(doc, "Section 3 - Compare", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/2 = 50%, larger than 40%.        b)  3/4 = 75%, larger than 70%.        c)  25% is larger (1/5 = 20%).", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "One third of 100 is 33.33..., which does not stop, so we write 'about 33 1/3%'. Two thirds is about 66 2/3%. They are not exact because 100 does not divide evenly by 3.", y);

    y = addTipBox(doc,
      "Watch for: students who glue digits (1/4 read as 14%); students who think 0.5 is smaller than 50%; students who compare without converting to the same form.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
