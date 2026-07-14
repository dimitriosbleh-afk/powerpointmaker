"use strict";

// Term 3 Maths Review (Year 6 Numeracy) - Session 3 of 4.
// FRACTIONS, DECIMALS & PERCENTAGES. Reviews Week 6 (VC2M5N04 + VC2M6N03) and
// Week 9 (VC2M6N07 + VC2M5N04 + VC2M6N08).
//   Block A - 100% is the whole; connect F, D and P (1/2 = 0.5 = 50%).
//   Block B - find a fraction or percentage of a quantity (find one part, scale).
//   Block C - percentage discounts, then estimate to check.
// Daily Review: Understanding & Operating with Fractions (prior, Week 8).
// Fluency: subtraction vertical algorithm with decimals (unit-wide focus).
// Unit variant fixed (variant 2) across all 4 sessions for cohesion.
// Catch-up: the launch starts from a 100-grid everyone can read, and worksheet
// Section 1 re-grounds FDP equivalence. No session assumes the one before.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 2); // variant 2, fixed for the unit
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide, dailyReviewSlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  addAreaModel, addTenthsStrip,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 4;
const UNIT_TITLE = "Term 3 Maths Review";
const FOOTER = `Term 3 Maths Review | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/TermRev_Lesson3_Fractions_Decimals_Percentages";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Connect fractions, decimals and percentages, find a part of a quantity, and work out discounts.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Session 3 review sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Percentage increase and decrease, and finding the original amount - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Small reusable visual ---------------------------------------------------

// FDP equivalence table: a header row (Fraction / Decimal / Percent) and one
// row per equivalent set, so students see the same amount written three ways.
function drawFdpTable(slide, x, y, w, rows) {
  const headers = ["Fraction", "Decimal", "Percent"];
  const cols = 3;
  const cellW = w / cols;
  const hdrH = 0.36, cellH = 0.5, gap = 0.08;
  headers.forEach((h, c) => {
    slide.addShape("rect", {
      x: x + c * cellW, y, w: cellW, h: hdrH,
      fill: { color: C.PRIMARY }, line: { color: C.WHITE, width: 1 },
    });
    slide.addText(h, {
      x: x + c * cellW, y, w: cellW, h: hdrH,
      fontSize: 11, fontFace: FONT_B, color: C.WHITE,
      bold: true, align: "center", valign: "middle", margin: 0, fit: "shrink",
    });
  });
  rows.forEach((r, ri) => {
    const ry = y + hdrH + 0.08 + ri * (cellH + gap);
    r.forEach((d, c) => {
      slide.addShape("rect", {
        x: x + c * cellW, y: ry, w: cellW, h: cellH,
        fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1 },
      });
      slide.addText(String(d), {
        x: x + c * cellW, y: ry, w: cellW, h: cellH,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
    });
  });
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to our Term 3 review. Today we connect fractions, decimals and percentages.
- These are three ways of writing the same amount. One half, nought point five, fifty per cent - all the same.
- We will swap between them, find a part of an amount, and work out a discount in a sale.

DO:
- Have whiteboards, markers and the review sheet ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Session 3 of 4. A review, so retrieve and firm up. Keep the hundred grid and the find-one-part-then-scale idea front and centre.

WATCH FOR:
- Students who think percentages are brand new - reassure them. It is just out of a hundred.

[General: Title | Element: Planning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The review sheet is for the You Do near the end. The Year 8 extension is there for anyone who is flying.

DO:
- Print one review sheet and one answer key per student.
- Have whiteboards and markers ready for every check.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student review sheet, an answer key, and a Year 8 extension on percentage increase and decrease. Most of the lesson runs on whiteboards.

CATCH-UP NOTE:
A student who missed earlier work can still access today. The launch starts from a hundred grid anyone can read, and Section 1 of the review sheet rebuilds the fraction, decimal and percentage links from scratch. A returner needs only a whiteboard and one minute with you to start. No session in this unit assumes the student saw the one before.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Warming up our fraction work from last session - comparing, adding and subtracting.
- Read each one and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for renaming to a common denominator.

TEACHER NOTES:
Daily Review is prior learning. Yesterday's fraction skills keep ticking over and feed straight into today's fraction-decimal-percentage links.

WATCH FOR:
- Students who rename 3/5 to tenths to compare - secure.
- Students who add the bottoms - address it in the reveal.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 3/5 is the same as 6/10, so 7/10 is larger.
- One quarter plus two quarters is three quarters.
- Five eighths take away one eighth is four eighths, which is one half.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The compare item rewards renaming to tenths. The add and subtract items reward keeping the denominator and working only with the tops.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students who add or subtract denominators - quick small group reminder.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Subtraction set out vertically, with decimals.
- Set each one out in columns on your whiteboard and subtract.
- Line up the decimal points so the places sit under each other.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up columns and the point kept in line.

TEACHER NOTES:
Fluency this whole unit is the subtraction algorithm with decimals. Decimals are one of the three forms we connect today, so this warms that up too.

WATCH FOR:
- Students who line up the points and fill gaps with zeros - secure.
- Students who right-align the digits instead of the points - fix it in the reveal.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 7.3 take away 4.8 is 2.5.
- 10.00 take away 3.65 is 6.35.
- 8.4 take away 2.85 is 5.55. Writing 8.4 as 8.40 helps line it up.
- Tick and fix.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The 8.4 minus 2.85 item is the key one. Fill the empty place with a zero to keep the columns honest.

WATCH FOR:
- Students who self-correct - secure.
- Students whose columns drift - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember the hundred grid. The whole grid is one hundred small squares - that is 100 per cent, the whole thing.
- I have shaded fifty squares. Look at how much is shaded: it is exactly half the grid.
- So this one shaded part has three names. As a fraction it is one half. As a decimal it is nought point five. As a percentage it is fifty per cent.
- Same amount, three ways to write it. That is the big idea today.

DO:
- Point to the whole grid (100 per cent), then to the fifty shaded squares.
- Ask: what fraction is shaded? What decimal? What percentage?
- Bridge: 'fractions, decimals and percentages are three names for the same amount'.

TEACHER NOTES:
This launch starts from a picture anyone can read, then connects it to the FDP links. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who name all three forms - strong prior knowledge.
- Students unsure of the decimal - 50 out of 100 is 0.50, which is 0.5; we build this today.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are reviewing how fractions, decimals and percentages connect, and using them to find a part of an amount.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Point out that the first one is something everyone can do today.

TEACHER NOTES:
The first criterion is reachable for everyone - name a shaded grid three ways. The second is the core target the exit ticket checks. The third stretches to working out a discount and checking it.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- Per cent means out of a hundred. So fifty per cent is fifty out of a hundred.
- Equivalent means equal in value - one half, nought point five and fifty per cent are equivalent.
- A discount is an amount taken off a price in a sale.

DO:
- Point to each word as you say it.
- Have students say 'per cent means out of a hundred' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. Per cent is the anchor word - keep linking it to out of a hundred.

WATCH FOR:
- Students who connect 25% to a quarter - secure.
- Students who read 50% as the number fifty - remind them it is out of a hundred.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO_A = `SAY:
- Let us line up the families that connect fractions, decimals and percentages.
- One half is nought point five, which is fifty per cent. Picture half the hundred grid shaded.
- One quarter is nought point two five, which is twenty five per cent. A quarter of the grid.
- One tenth is nought point one, which is ten per cent. One column of the grid.
- The percentage is just the decimal out of a hundred. Nought point five is fifty hundredths, so fifty per cent.

DO:
- Point to each row of the table as you say it.
- Link each one back to the hundred grid from the launch.
- Have students chorus 'the percentage is the decimal out of a hundred'.

TEACHER NOTES:
This is the core of Block A. These three families - halves, quarters, tenths - cover most of what students need. Keep tying back to the grid.

MISCONCEPTIONS:
- Misconception: 0.5 is the same as 5%.
  Why: students read the 5 and attach a per cent sign.
  Impact: a percentage ten times too small.
  Quick correction: 0.5 is 5 tenths, which is 50 hundredths, so 50%.

WATCH FOR:
- Students who convert decimal to percent by thinking hundredths - secure.
- Students who write 0.5 as 5% - back to the grid, half is 50 squares.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_A_Q = `SAY:
- Quick check on your whiteboards.
- Write three quarters as a decimal and as a percentage.
- Picture three quarters of the hundred grid if it helps.

DO:
- Display the prompt.
- Give 60 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 0.75 and 75%.
PROCEED: If about 80 percent have it right, click to reveal and move to Block B.
PIVOT: Most likely misconception - students write 0.75 but then 7.5% or 0.75%.
- Reteach: 0.75 is 75 hundredths; hundredths ARE per cent, so 75%.
- Re-check: how many hundredths is 0.75?

TEACHER NOTES:
The trap is the decimal to percentage step. Hundredths and per cent are the same thing.

WATCH FOR:
- Students with 0.75 and 75% - secure.
- Students with 7.5% - reteach hundredths equal per cent.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_IDO_B = `SAY:
- Now we find a percentage of an amount. Find thirty per cent of forty.
- The smart move is to find a friendly part first. Ten per cent is easy - it is one tenth. One tenth of forty is four.
- Thirty per cent is three lots of ten per cent. So I take my four and multiply by three: three times four is twelve.
- So thirty per cent of forty is twelve. Find one part, then scale up.

DO:
- Point to the bar split into tenths; each tenth is worth four.
- Shade three tenths and count: four, eight, twelve.
- Have students whisper 'find ten per cent, then scale' to a partner.

TEACHER NOTES:
This is VC2M6N07. Ten per cent is the friendly part - dividing by ten - then multiply for the percentage you need.

MISCONCEPTIONS:
- Misconception: 30% of 40 is 30, or is found by subtracting.
  Why: students grab the numbers without finding a part.
  Impact: an answer that ignores the whole.
  Quick correction: find 10% first (4), then take 3 lots of it.

WATCH FOR:
- Students who find 10% then scale - secure.
- Students who guess - back to the tenths bar.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_WEDO_B_Q = `SAY:
- Your turn with a partner on your whiteboards.
- One: find twenty five per cent of sixty.
- Two: find one third of eighteen.

DO:
- Display the two prompts.
- Give 75 seconds.
- Listen for find one part, then scale.

TEACHER NOTES:
25% is a quarter, so divide by four. One third means divide by three. Both are find-one-part moves.

WATCH FOR:
- Pairs who link 25% to a quarter - secure.
- Pairs stuck - prompt: what fraction is 25%? What do you divide by?

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO_B_A = `SAY:
- Let us check.
- Twenty five per cent of sixty: a quarter of sixty is sixty divided by four, which is fifteen.
- One third of eighteen: eighteen divided by three is six.
- Both used find one part of the amount.

DO:
- Click to reveal.
- Re-say the find-one-part move for each.

TEACHER NOTES:
If pairs struggled, model 25% of 60 as 10% is 6, plus another 10% is 12, plus 5% is 3, total 15 - or just a quarter of 60.

WATCH FOR:
- Students who self-correct - secure.
- Students who multiply instead of divide - small group focus.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_IDO_C = `SAY:
- Now a real sale. A toy costs forty dollars and it is thirty per cent off. How much do you pay?
- First I find the discount. Ten per cent of forty dollars is four dollars. Thirty per cent is three lots, so twelve dollars off.
- Then I take it off the price: forty dollars take away twelve dollars is twenty eight dollars. You pay twenty eight dollars.
- Let me check it makes sense. Thirty per cent is close to a third, and a third of forty is about thirteen. Twelve is close to that, so my discount looks right.

DO:
- Point to the 10% chip, then the 30% chip, then the price you pay.
- Say the estimate out loud to model checking.
- Have students whisper the two steps: find the discount, take it off.

TEACHER NOTES:
This is VC2M6N07 plus VC2M6N08 - find the discount with the find-one-part move, then estimate to check the answer is sensible.

MISCONCEPTIONS:
- Misconception: 30% off means you pay 30% of the price.
  Why: students find the discount and stop, or confuse off with of.
  Impact: they report the discount as the price.
  Quick correction: the discount comes OFF; pay the original minus the discount.

WATCH FOR:
- Students who subtract the discount from the price - secure.
- Students who stop at the discount - prompt: that is how much you save, not what you pay.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_WEDO_C_Q = `SAY:
- Your turn with a partner. A game costs sixty dollars and it is twenty five per cent off.
- Work out the discount, then what you pay.
- Then check: is your answer sensible?

DO:
- Display the prompt.
- Give about 90 seconds.
- Listen for the two steps and a quick estimate.

TEACHER NOTES:
25% is a quarter. A quarter of 60 dollars is 15 dollars off, so you pay 45 dollars.

WATCH FOR:
- Pairs who find the discount then subtract - secure.
- Pairs who report 15 dollars as the price - prompt: that is the saving.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO_C_A = `SAY:
- Let us check.
- Twenty five per cent of sixty dollars is a quarter of sixty, which is fifteen dollars off.
- Sixty dollars take away fifteen dollars is forty five dollars. You pay forty five dollars.
- Does it make sense? A quarter off leaves three quarters, and three quarters of sixty is forty five. It checks out.

DO:
- Click to reveal.
- Point out the three-quarters check.

TEACHER NOTES:
If many reported 15 dollars as the price, do one more like 10% off 50 dollars before the You Do.

WATCH FOR:
- Students who self-correct - ready for independent work.
- Students who still report the discount - enabling group for the You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the review sheet.
- Section 1 warms up by connecting fractions, decimals and percentages - that is for everyone.
- Section 2 finds a part of an amount. Section 3 works out discounts and checks them.
- If you finish, try the challenge box, then the Year 8 extension.

DO:
- Distribute the review sheet.
- Circulate and listen for find-one-part-then-scale language.
- Cold call one or two students to explain an answer.

TEACHER NOTES:
Different numbers from the We Do, same moves. Section 1 is the rebuild for any returning student.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 gives a hundred grid and one worked family. Students shade and name the amount three ways.
- Extra Notes: Sit with these students and link each form to the grid. This is also the rebuild for a returner.
EXTENDING PROMPT:
- Task: The challenge box asks students to find which is the better buy: 20% off, or a third off.
- Extra Notes: Students who are ready move on to the Year 8 extension on percentage increase and decrease.

WATCH FOR:
- Students who find one part then scale - secure.
- Students who confuse off with of - prompt the take-it-off step.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- One: write one quarter as a percentage.
- Two: find twenty per cent of fifty.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - connect a fraction to a percentage, and find a percentage of an amount. Look for 25% and 10.

WATCH FOR:
- Students who write 25% and find 10 - secure.
- Students who write 4% or find 20 - revisit at the start of Session 4.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: how would you find ten per cent of an amount, and how does that help you find thirty per cent?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that fractions, decimals and percentages are three names for the same amount, and finding one friendly part lets you scale to any percentage. Students who can say this are ready for problem solving and measurement next session.

WATCH FOR:
- Strong thumbs up across all three - move at pace next session.
- Sideways or down on the core criterion - small group revision at the start of Session 4.

[General: Closing | Element: Retention and Recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 3: Fractions, decimals and percentages",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - operating with fractions
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Operating with fractions",
      [
        "Which is larger: 3/5 or 7/10?",
        "Work out 1/4 + 2/4.",
        "Work out 5/8 - 1/8.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "7/10 is larger       1/4 + 2/4 = 3/4       5/8 - 1/8 = 4/8 = 1/2", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - subtraction with decimals
  withReveal(
    () => fluencySlide(pres, "Fluency: Subtraction with decimals",
      ["7.3 - 4.8", "10.00 - 3.65", "8.4 - 2.85"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "2.5        6.35        5.55", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - 50 of 100 shaded = 0.5 = 1/2 = 50% (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "One shaded part, three names",
    [
      "The whole grid = 100 squares = 100%.",
      "50 squares are shaded.",
      "",
      "As a fraction: 1/2.",
      "As a decimal: 0.5.",
      "As a percentage: 50%.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.ACCENT });
      const gridSize = 2.0;
      const gx = lg.rightX + (lg.rightW - gridSize) / 2;
      const gy = lg.panelTopPadded + 0.4;
      addAreaModel(slide, gx, gy, gridSize, 5, 0);
      slide.addText("50 of 100 shaded = 0.5 = 1/2 = 50%", {
        x: lg.rightX + 0.1, y: gy + gridSize + 0.18, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are reviewing how fractions, decimals and percentages connect, and using them to find a part of an amount.",
    [
      "I can name a shaded grid as a fraction, decimal and percentage.",
      "I can find a fraction or percentage of a quantity.",
      "I can work out a discount and check that my answer is sensible.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Per cent = out of 100 (50% = 50 out of 100)",
      "Equivalent = same value (1/2 = 0.5 = 50%)",
      "Discount = an amount taken off the price",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.1, { strip: C.SECONDARY });
      slide.addText("5 tenths shaded", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
      });
      addTenthsStrip(slide, lg.rightX + 0.35, lg.panelTopPadded + 0.62, lg.rightW - 1.2, 5);
      slide.addText("= 0.5 = 50%", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.25, w: lg.rightW - 0.4, h: 0.34,
        fontSize: 17, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
    }
  );

  // Slide 10: I Do A - FDP equivalence families
  workedExSlide(pres, 2, "I Do", "Three names for the same amount",
    [
      "100% is the whole.",
      "Per cent means out of 100.",
      "1/2 = 0.5 = 50%.",
      "1/4 = 0.25 = 25%.",
      "1/10 = 0.1 = 10%.",
      "",
      "The percentage is the decimal out of 100.",
    ],
    NOTES_IDO_A, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.PRIMARY });
      slide.addText("Same amount, three forms", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      drawFdpTable(slide, lg.rightX + 0.2, lg.panelTopPadded + 0.5, lg.rightW - 0.4,
        [["1/2", "0.5", "50%"], ["1/4", "0.25", "25%"], ["1/10", "0.1", "10%"]]);
    }
  );

  // Slides 11-12: CFU A + reveal - convert a fraction to decimal and percent
  withReveal(
    () => cfuSlide(pres, "CFU", "Three names for 3/4",
      { technique: "Show Me Boards",
        question: "Write 3/4 as a decimal and as a percentage.\n\nPicture three quarters of the hundred grid." },
      NOTES_CFU_A_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3/4 = 0.75 = 75%", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: I Do B - find a percentage of a quantity (find one part, scale)
  workedExSlide(pres, 2, "I Do", "Find 30% of 40",
    [
      "Find a friendly part first.",
      "10% of 40 = 4 (one tenth).",
      "30% = 3 lots of 10%.",
      "3 x 4 = 12.",
      "",
      "So 30% of 40 = 12.",
    ],
    NOTES_IDO_B, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.PRIMARY });
      slide.addText("40 split into tenths (each = 4)", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addTenthsStrip(slide, lg.rightX + 0.35, lg.panelTopPadded + 0.7, lg.rightW - 1.2, 3);
      addTextOnShape(slide, "3 tenths = 3 x 4 = 12", {
        x: lg.rightX + 0.6, y: lg.panelTopPadded + 1.5, w: lg.rightW - 1.2, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do B + reveal - find a part of a quantity
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your turn: find a part",
      [
        "With your partner.",
        "",
        "1.  Find 25% of 60.",
        "2.  Find 1/3 of 18.",
        "",
        "Find one part, then scale.",
      ],
      NOTES_WEDO_B_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
        slide.addText("25% = a quarter\n1/3 = divide by 3", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.25, w: lg.rightW - 0.4, h: 0.9,
          fontSize: 17, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        addTextOnShape(slide, "What do you\ndivide by?", {
          x: lg.rightX + 0.6, y: lg.panelTopPadded + 1.3, w: lg.rightW - 1.2, h: 0.8, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      }),
    (slide) => {
      addTextOnShape(slide, "25% of 60 = 15        1/3 of 18 = 6", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_B_A);
    }
  );

  // Slide 16: I Do C - percentage discount + estimate to check
  workedExSlide(pres, 2, "I Do", "A $40 toy is 30% off",
    [
      "Step 1 - find the discount.",
      "10% of $40 = $4, so 30% = $12 off.",
      "Step 2 - take it off the price.",
      "$40 - $12 = $28. You pay $28.",
      "",
      "Check: 30% is about 1/3; 1/3 of $40 is about $13. Close.",
    ],
    NOTES_IDO_C, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.PRIMARY });
      slide.addText("Find the discount, then take it off", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 12.5, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      const chips = [
        ["10% of $40 = $4", C.SECONDARY],
        ["30% = $12 off", C.SECONDARY],
        ["You pay $28", C.ALERT],
      ];
      const cy0 = lg.panelTopPadded + 0.62;
      chips.forEach((ch, i) => {
        addTextOnShape(slide, ch[0], {
          x: lg.rightX + 0.45, y: cy0 + i * 0.6, w: lg.rightW - 0.9, h: 0.46, rectRadius: 0.07,
          fill: { color: ch[1] },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
      slide.addText("Estimate check: about 1/3 off = about $13", {
        x: lg.rightX + 0.2, y: cy0 + 1.92, w: lg.rightW - 0.4, h: 0.34,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slides 17-18: We Do C + reveal - discount problem
  withReveal(
    () => cfuSlide(pres, "We Do", "A $60 game is 25% off",
      { technique: "Partner whiteboards",
        question: "Work out the discount, then what you pay.\n\nThen check: is your answer sensible?" },
      NOTES_WEDO_C_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "25% of $60 = $15 off, so you pay $45", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_C_A);
    }
  );

  // Slide 19: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: review sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "connect the fractions, decimals and percentages.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "find a part of each amount.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "work out the discounts and check them.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.4;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember", {
      x: 0.7, y: panelY + 0.13, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
    });
    addTextOnShape(s, "Find 10% first (divide by 10), then scale. A discount comes OFF the price.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    s.addText("Reference families:  1/2 = 0.5 = 50%    1/4 = 0.25 = 25%    1/10 = 0.1 = 10%", {
      x: 0.7, y: panelY + 1.18, w: 8.6, h: 0.34,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 20: Exit Ticket
  exitTicketSlide(pres,
    [
      "Write 1/4 as a percentage.",
      "Find 20% of 50.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 21: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how would you find 10% of an amount, and how does that help you find 30%?",
      scItems: [
        "I can name a shaded grid as a fraction, decimal and percentage.",
        "I can find a fraction or percentage of a quantity.",
        "I can work out a discount and check that my answer is sensible.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "TermRev_Lesson3_Fractions_Decimals_Percentages.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Connect fractions, decimals and percentages, find a part, and work out discounts.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Fractions, decimals and percentages are three names for the same amount. Per cent means out of 100. To find a percentage of an amount, find a friendly part first (10% is one tenth), then scale. A discount comes OFF the price.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Find 30% of 40. 10% of 40 = 4. 30% is 3 lots of 10%, so 3 x 4 = 12. So 30% of 40 = 12.",
      y);

    y = addSectionHeading(doc, "Section 1 - Connect F, D and P (for everyone)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2 = ______ (decimal) = ______ (percent)", y);
    y = addWriteLine(doc, "b)  0.25 = ______ (fraction) = ______ (percent)", y);
    y = addWriteLine(doc, "c)  10% = ______ (fraction) = ______ (decimal)", y);

    y = addSectionHeading(doc, "Section 2 - Find a part of an amount", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  10% of 70 = ______", y);
    y = addWriteLine(doc, "b)  25% of 40 = ______", y);
    y = addWriteLine(doc, "c)  1/3 of 24 = ______", y);
    y = addWriteLine(doc, "d)  30% of 50 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Discounts (work it out, then check)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  A $20 book is 50% off. You pay ______", y);
    y = addWriteLine(doc, "b)  A $80 scooter is 25% off. You pay ______", y);
    y = addWriteLine(doc, "c)  A $50 jacket is 10% off. You pay ______", y);

    y = addSectionHeading(doc, "Challenge (optional)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "A $30 item: which saves more, 20% off or a third off? Show your check.  ______", y);

    addPdfFooter(doc, `Session ${SESSION} | Fractions, decimals and percentages | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Session 3 review sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Connect F, D and P", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/2 = 0.5 = 50%.   b)  0.25 = 1/4 = 25%.   c)  10% = 1/10 = 0.1.", y);

    y = addSectionHeading(doc, "Section 2 - Find a part", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  7        b)  10        c)  8        d)  15  (10% of 50 = 5, x 3 = 15)", y);

    y = addSectionHeading(doc, "Section 3 - Discounts", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  50% off $20 = $10 off, pay $10.   b)  25% off $80 = $20 off, pay $60.   c)  10% off $50 = $5 off, pay $45.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "A third off $30 = $10 off; 20% off $30 = $6 off. A third off saves more. Check: a third is bigger than 20%, so it should save more.", y);

    y = addTipBox(doc,
      "Watch for: students who read 0.5 as 5%; students who report the discount as the price (off vs of); students who forget to scale after finding 10%.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Percentage increase and decrease, and finding the original amount.",
      color: C.SECONDARY,
      lessonInfo: `Session ${SESSION} | Year 8 challenge | extends Year 6 VC2M6N07`,
    });
    y = addTipBox(doc,
      "A percentage increase adds on; a decrease takes off. After a change you can think in multipliers: 15% off means you pay 85%, or x 0.85. To find the original from the new price, divide by that multiplier.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "A jacket is 20% off and now costs $48. You paid 80% of the original, so original = 48 divided by 0.8 = $60.",
      y);

    y = addSectionHeading(doc, "Section 1 - Increase and decrease", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  Increase $40 by 15% = ______", y);
    y = addWriteLine(doc, "b)  Decrease $90 by 30% = ______", y);
    y = addWriteLine(doc, "c)  Increase $250 by 8% = ______", y);

    y = addSectionHeading(doc, "Section 2 - Find the original", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  After 25% off, a price is $30. Original = ______", y);
    y = addWriteLine(doc, "b)  After 10% off, a price is $63. Original = ______", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "A price goes up 10%, then down 10%. Is it back to the start? Explain.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) $46  b) $63  c) $270.   S2  a) $40 (30 / 0.75)  b) $70 (63 / 0.9).   S3  No - up 10% then down 10% lands at 99% of the start, because the 10% down is taken off a bigger amount.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Session ${SESSION} | Year 8 Extension | Percentage change`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Session 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
