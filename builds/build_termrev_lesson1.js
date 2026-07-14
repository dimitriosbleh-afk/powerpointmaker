"use strict";

// Term 3 Maths Review (Year 6 Numeracy) - Session 1 of 4.
// NUMBER & PLACE VALUE: decimals and integers. Reviews Week 1 (Advanced Number
// Concepts, VC2M5N01 + VC2M6N01) and Week 3 (Decimal Operations, VC2M6N06).
//   Block A - compare and order decimals to thousandths (line up place value).
//   Block B - multiply and divide decimals by powers of 10 (digits move, point stays).
//   Block C - locate decimals and integers on a number line (negatives included).
// Daily Review: Mastering Fractions, Decimals & Percentages (prior, Week 9).
// Fluency: subtraction vertical algorithm with decimals (unit-wide focus).
// Unit variant fixed (variant 2) across all 4 sessions for cohesion.
// Catch-up: the launch starts from two area models everyone can read, and
// worksheet Section 1 re-grounds place value. No session assumes the one before.

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
  addPlaceValueChart, addNumberLine, addAreaModel,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 4;
const UNIT_TITLE = "Term 3 Maths Review";
const FOOTER = `Term 3 Maths Review | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/TermRev_Lesson1_Decimals_And_Integers";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Compare and order decimals, multiply and divide by powers of 10, and place integers on a number line.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Session 1 review sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Powers of 10 and standard form (scientific notation) - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Small reusable visuals --------------------------------------------------

// Aligned place-value comparison: each row is one decimal written digit by
// digit in Ones / Tenths / Hundredths / Thousandths cells so students compare
// column by column. The tenths column is highlighted because it usually
// decides the order. A small dot marks the decimal point on each row.
function drawDecimalCompare(slide, x, y, w, rows) {
  const cols = 4;
  const headers = ["Ones", "Tenths", "Hundredths", "Thousandths"];
  const cellW = w / cols;
  const hdrH = 0.36, cellH = 0.5, gap = 0.10;
  headers.forEach((h, c) => {
    slide.addShape("rect", {
      x: x + c * cellW, y, w: cellW, h: hdrH,
      fill: { color: C.PRIMARY }, line: { color: C.WHITE, width: 1 },
    });
    slide.addText(h, {
      x: x + c * cellW, y, w: cellW, h: hdrH,
      fontSize: c === 0 ? 9 : 8, fontFace: FONT_B, color: C.WHITE,
      bold: true, align: "center", valign: "middle", margin: 0, fit: "shrink",
    });
  });
  rows.forEach((r, ri) => {
    const ry = y + hdrH + 0.08 + ri * (cellH + gap);
    r.forEach((d, c) => {
      const hi = c === 1;
      slide.addShape("rect", {
        x: x + c * cellW, y: ry, w: cellW, h: cellH,
        fill: { color: hi ? C.SECONDARY : C.WHITE }, line: { color: C.PRIMARY, width: 1 },
      });
      slide.addText(String(d), {
        x: x + c * cellW, y: ry, w: cellW, h: cellH,
        fontSize: 22, fontFace: FONT_H, color: hi ? C.WHITE : C.CHARCOAL,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
    });
    // Decimal point on the ones/tenths boundary.
    slide.addShape("roundRect", {
      x: x + cellW - 0.05, y: ry + cellH - 0.17, w: 0.10, h: 0.10, rectRadius: 0.05,
      fill: { color: C.ALERT },
    });
  });
}

// Mark an integer value on a horizontal -5..5 number line drawn at (lineX, lineY)
// with width lineW. Adds a coloured dot and a label below.
function markInteger(slide, lineX, lineY, lineW, value, color) {
  const frac = (value + 5) / 10;
  const mx = lineX + frac * lineW;
  slide.addShape("roundRect", {
    x: mx - 0.08, y: lineY - 0.08, w: 0.16, h: 0.16, rectRadius: 0.08,
    fill: { color: color || C.ALERT },
  });
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to our Term 3 review. Over the next four sessions we bring back everything we learned this term and make it stick.
- Today is all about numbers themselves - decimals and integers. Where they sit, how big they are, and what happens when we multiply or divide by ten, a hundred, a thousand.
- If some of this feels rusty, that is exactly what review is for. We build it back together.

DO:
- Have whiteboards, markers and the review sheet ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Session 1 of 4. This is a review unit, so the aim is to retrieve and firm up, not to teach brand new content. Keep the pace warm and confident.

WATCH FOR:
- Students who say they have forgotten - reassure them. Review brings it back fast.

[General: Title | Element: Planning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The review sheet is for the You Do near the end. The Year 8 extension is there for anyone who is flying.

DO:
- Print one review sheet and one answer key per student.
- Have whiteboards and markers ready for Daily Review, Fluency and every check.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student review sheet, an answer key, and a Year 8 extension on powers of 10 and standard form. Most of the lesson runs on whiteboards.

CATCH-UP NOTE:
A student who missed earlier work can still access today. The launch starts from two pictures of a decimal that anyone can read, and Section 1 of the review sheet rebuilds place value from scratch. A returner needs only a whiteboard and one minute with you to start. No session in this unit assumes the student saw the one before.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are warming up our most recent work - finding a fraction or a percentage of an amount.
- Read each one and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for how students find one part first, then scale.

TEACHER NOTES:
Daily Review is prior learning, not today's content. Finding a part of a quantity keeps ticking over from Week 9 because we lean on it again next session.

WATCH FOR:
- Students who find a quarter by dividing by 4 - secure.
- Students who confuse 10% and 1/10 - that is fine, they are the same thing; name it in the reveal.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- A quarter of 20 means share 20 into 4 equal groups, so 5.
- 10% of 80 is one tenth of 80, which is 8.
- 25% is the same as a quarter, so 25% of 40 is 10.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The link to make explicit: 10% is one tenth, 25% is one quarter. Finding one part then scaling is the move we reuse all unit.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students who multiply instead of divide - revisit find one part first.

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
Fluency this whole unit is the subtraction algorithm with decimals. Lining up the points is the same place value habit we review on the slides today.

WATCH FOR:
- Students who line up the points and fill gaps with zeros - secure.
- Students who right-align the digits instead of the points - common error, fix it in the reveal.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 6.4 take away 2.7 is 3.7.
- 8.50 take away 3.65 is 4.85.
- 7.0 take away 1.8 is 5.2. Writing 7.0 as 7.00 helps line it up.
- Tick and fix.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The 7.0 minus 1.8 item is the key one. Filling the empty place with a zero keeps the columns honest. Keep it brisk.

WATCH FOR:
- Students who self-correct - secure.
- Students whose columns drift - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember work like this. Two long-jump marks: one of 0.5 metres and one of 0.45 metres. Which jump is longer?
- A trap voice in our head says 0.45 because 45 is a big number. Let us look instead of guess.
- On the left, 0.5 fills five whole columns - fifty little squares. On the right, 0.45 fills four columns and five more - forty-five squares.
- Fifty squares beats forty-five, so 0.5 is longer. More digits after the point does not mean a bigger number. That idea drives today.

DO:
- Point to each shaded grid and count the columns together.
- Ask for a thumbs up or down: is 0.5 bigger than 0.45?
- Bridge: 'to compare decimals we look at place value, not how many digits there are'.

TEACHER NOTES:
This launch starts from a picture anyone can read, then connects it to today's compare-and-order work. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who count the squares and decide confidently - strong prior knowledge.
- Students still tempted by digit count - reassure, the grids settle it; we build the place value habit this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are reviewing how to read, compare and move decimals and integers using place value.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Point out that the first one is something everyone can do today.

TEACHER NOTES:
The first criterion is reachable for everyone - compare two decimals. The second is the core target the exit ticket checks. The third stretches to explaining the powers-of-ten effect in place value words.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Four words for today.
- Tenths, hundredths and thousandths are the places after the decimal point. Each step right is ten times smaller.
- An integer is a whole number. It can be positive, negative, or zero - no fraction part.
- So 3, 0 and negative 4 are all integers; 3.5 is not.

DO:
- Point to each word as you say it.
- Have students say 'each place to the right is ten times smaller' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. Thousandths and integer are the two anchor words for today.

WATCH FOR:
- Students who can name the place after hundredths - secure.
- Students who think a negative cannot be an integer - clarify with the number line later.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO_A = `SAY:
- Let us compare three decimals together: 0.6, 0.45 and 0.275.
- Watch how I line them up by place value. I write each one to thousandths: 0.600, 0.450, 0.275.
- Now I compare from the left, the biggest place first. Tenths: 6, then 4, then 2. That order already decides it.
- 0.6 has the most tenths, so it is largest. 0.275 has the fewest tenths, so it is smallest. Smallest to largest: 0.275, 0.45, 0.6.
- Notice 0.275 looks the longest but is the smallest. Place value wins, not digit count.

DO:
- Point to the tenths column on the chart as you compare.
- Write the three numbers padded to thousandths so the columns line up.
- Have students chorus 'compare the tenths first'.

TEACHER NOTES:
This is the core move of Block A. Always padding to the same number of places makes the comparison fair and visible.

MISCONCEPTIONS:
- Misconception: a decimal with more digits is bigger, so 0.275 beats 0.6.
  Why: students read 275 as larger than 6.
  Impact: they order decimals back to front.
  Quick correction: line up place value and compare the tenths first.

WATCH FOR:
- Students who compare place by place - secure.
- Students who count digits - send them back to the tenths column.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_A_Q = `SAY:
- Quick check on your whiteboards.
- Order these from smallest to largest: 0.3, 0.275, 0.31.
- Line them up by place value first.

DO:
- Display the three decimals.
- Give 60 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 0.275, 0.3, 0.31 in that order.
PROCEED: If about 80 percent have it right, click to reveal and move to Block B.
PIVOT: Most likely misconception - students put 0.275 last because it has the most digits.
- Reteach: write all three to thousandths (0.300, 0.275, 0.310) and compare the tenths, then hundredths.
- Re-check: which has the fewest tenths?

TEACHER NOTES:
The trap is 0.275 looking biggest. A student who lines up place value will see it has only 2 tenths and is smallest.

WATCH FOR:
- Students with 0.275 first - secure.
- Students with 0.275 last - reteach with padded zeros.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_IDO_B = `SAY:
- Now we review multiplying and dividing decimals by ten, a hundred, a thousand.
- Here is 6.3. Watch what happens, and watch the decimal point - it stays put. The digits move.
- 6.3 times 10: every digit moves one place to a bigger column, so we get 63.
- 6.3 times 100: the digits move two places, so 630.
- 6.3 divided by 10: the digits move one place to a smaller column, so 0.63.
- We are not sliding the point along. The digits change place value, and that is the proper way to say it.

DO:
- Point to the chart, then to each result chip as you say it.
- Trace one digit, the 6, moving columns.
- Have students whisper 'digits move, the point stays' to a partner.

TEACHER NOTES:
This is VC2M6N06. Say the effect in place value language - digits become ten times bigger or smaller - not 'move the point'. The chart makes the movement visible.

MISCONCEPTIONS:
- Misconception: to multiply by 10 you 'add a zero'.
  Why: it works for whole numbers and gets over-generalised.
  Impact: 6.3 times 10 becomes 6.30, which is wrong.
  Quick correction: move the digits one place; 6.3 becomes 63.

WATCH FOR:
- Students who move digits by place - secure.
- Students who 'add a zero' to a decimal - reteach with the chart.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_WEDO_B_Q = `SAY:
- Your turn with me, on your whiteboards with a partner.
- Work out 5.2 times 100, and 0.6 divided by 10.
- Say what happens to the digits each time.

DO:
- Display the two prompts.
- Give 75 seconds.
- Listen for place value language, not 'add a zero' or 'move the point'.

TEACHER NOTES:
Same move as the I Do with new numbers. 0.6 divided by 10 is the Week 1 example - six tenths becomes six hundredths, 0.06.

WATCH FOR:
- Pairs who explain the digit movement - secure.
- Pairs who guess - prompt them back to the place value chart.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO_B_A = `SAY:
- Let us check.
- 5.2 times 100: the digits move two places, so 520.
- 0.6 divided by 10: six tenths become six hundredths, so 0.06.
- The point never moved. The digits changed place value.

DO:
- Click to reveal.
- Re-say the place value movement for each.

TEACHER NOTES:
If pairs 'added zeros' or 'moved the point', model the digit movement once more before Block C.

WATCH FOR:
- Students who self-correct - secure.
- Students who still add zeros - small group focus.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_IDO_C = `SAY:
- A number line orders any numbers, including decimals between decimals and negative numbers.
- Look at the top line. It zooms between 2.33 and 2.34. The number 2.335 sits exactly halfway between them, so 2.33 is less than 2.335 is less than 2.34.
- Now the bottom line. It runs through zero into negatives. Negative 3 sits three steps to the left of zero. Positive 2 sits two steps to the right.
- The rule on any line: further left is smaller, further right is larger. So negative 3 is less than zero is less than 2.

DO:
- Point to 2.335 halfway between the two marks.
- Point to negative 3 on the left and 2 on the right of zero.
- Have students say 'further left is smaller'.

TEACHER NOTES:
This reviews two Week 1 ideas at once - locating a three-place decimal between two marks, and placing integers around zero. Keep tying position to size.

MISCONCEPTIONS:
- Misconception: negative 5 is bigger than negative 3 because 5 is bigger.
  Why: students compare the digits and ignore the sign.
  Impact: they order negatives back to front.
  Quick correction: point to the line - negative 5 is further left, so it is smaller.

WATCH FOR:
- Students who read position as size - secure.
- Students who say negative 5 is biggest - back to the line, further left is smaller.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_WEDO_C_Q = `SAY:
- Your turn with a partner. Use the number line.
- Place these five numbers, then write them in order from smallest to largest: negative 4, negative 1, 3, negative 3, 2.
- Remember, further left is smaller.

DO:
- Display the prompt and the line.
- Give about 90 seconds.
- Listen for students using the line, not just the digits.

TEACHER NOTES:
This pulls integers together. The two negatives, negative 4 and negative 3, are the discriminator - negative 4 is further left, so smaller.

WATCH FOR:
- Pairs who order the negatives correctly - secure.
- Pairs who put negative 4 after negative 3 - prompt them to the line.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO_C_A = `SAY:
- Let us check on the line.
- From the far left: negative 4, then negative 3, then negative 1, then 2, then 3.
- The two negatives are the tricky pair. Negative 4 is further left than negative 3, so it is smaller.

DO:
- Click to reveal the marked line and the order.
- Run a finger left to right along the order.

TEACHER NOTES:
If many ordered the negatives back to front, do one more set before the You Do.

WATCH FOR:
- Students who self-correct - ready for independent work.
- Students still flipping the negatives - enabling group for the You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the review sheet.
- Section 1 warms up by comparing and ordering decimals - that is for everyone.
- Section 2 multiplies and divides by powers of ten. Section 3 places integers on a line.
- If you finish, try the challenge box, then the Year 8 extension.

DO:
- Distribute the review sheet.
- Circulate and listen for place value language.
- Cold call one or two students to explain an answer.

TEACHER NOTES:
Different numbers from the We Do, same moves. Section 1 is the rebuild for any returning student.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 gives the place value headers and one worked comparison. Students pad to thousandths and compare the tenths.
- Extra Notes: Sit with these students and point to the tenths column. This is also the rebuild for a returner.
EXTENDING PROMPT:
- Task: The challenge box asks students to explain, in place value words, why 4.2 times 100 is not 4.200.
- Extra Notes: Students who are ready move on to the Year 8 extension on standard form.

WATCH FOR:
- Students who compare by place value - secure.
- Students who count digits or add zeros - prompt back to the chart.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- One: order these from smallest to largest - 0.4, 0.39, 0.405.
- Two: work out 3.7 times 100.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - compare and order decimals, and multiply by a power of ten. Look for 0.39, 0.4, 0.405 and 370.

WATCH FOR:
- Students who order by place value and get 370 - secure.
- Students who put 0.405 first or write 3.700 - revisit at the start of Session 2.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: when we multiply 5.6 by 10, what happens to the digits and what happens to the point?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that place value decides everything - how big a decimal is, and what happens when we multiply or divide by powers of ten. Students who can say 'the digits move, the point stays' are ready for fractions next session.

WATCH FOR:
- Strong thumbs up across all three - move at pace next session.
- Sideways or down on the core criterion - small group revision at the start of Session 2.

[General: Closing | Element: Retention and Recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 1: Decimals, place value and integers",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - finding a part of a quantity
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Find part of an amount",
      [
        "Find 1/4 of 20.",
        "Find 10% of 80.",
        "Find 25% of 40.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1/4 of 20 = 5       10% of 80 = 8       25% of 40 = 10", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - subtraction with decimals
  withReveal(
    () => fluencySlide(pres, "Fluency: Subtraction with decimals",
      ["6.4 - 2.7", "8.50 - 3.65", "7.0 - 1.8"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3.7        4.85        5.2", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - which is longer, 0.5 or 0.45 (area models, catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "Which is longer: 0.5 m or 0.45 m?",
    [
      "Two long-jump marks.",
      "0.5 fills five whole columns.",
      "0.45 fills four columns and five more.",
      "",
      "50 squares beats 45.",
      "More digits does not mean bigger.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.ACCENT });
      const gridSize = 1.55;
      const g1x = lg.rightX + 0.35;
      const g2x = lg.rightX + lg.rightW - gridSize - 0.35;
      const gy = lg.panelTopPadded + 0.55;
      slide.addText("0.5", {
        x: g1x, y: lg.panelTopPadded + 0.14, w: gridSize, h: 0.32,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      slide.addText("0.45", {
        x: g2x, y: lg.panelTopPadded + 0.14, w: gridSize, h: 0.32,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addAreaModel(slide, g1x, gy, gridSize, 5, 0);
      addAreaModel(slide, g2x, gy, gridSize, 4, 5);
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are reviewing how to read, compare and move decimals and integers using place value.",
    [
      "I can compare two decimals using place value.",
      "I can order decimals and multiply or divide them by powers of 10.",
      "I can place integers on a number line and explain the powers-of-10 effect.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Tenths, hundredths, thousandths = the places after the point",
      "Each step right = 10 times smaller",
      "Integer = a whole number (can be negative or zero)",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
      slide.addText("0.001 is 10x smaller than 0.01", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.32,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
      });
      const rows = [["3", C.PRIMARY], ["0", C.PRIMARY], ["-4", C.SECONDARY]];
      const ry0 = lg.panelTopPadded + 0.62;
      rows.forEach((r, i) => {
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.55, y: ry0 + i * 0.52, w: 0.9, h: 0.42, rectRadius: 0.07,
          fill: { color: r[1] },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
      slide.addText("...are all integers.\n3.5 is not.", {
        x: lg.rightX + 1.65, y: ry0 + 0.08, w: lg.rightW - 1.85, h: 1.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      });
    }
  );

  // Slide 10: I Do A - compare and order decimals
  workedExSlide(pres, 2, "I Do", "Compare and order: 0.6, 0.45, 0.275",
    [
      "Line them up by place value.",
      "Write each to thousandths: 0.600, 0.450, 0.275.",
      "Compare the biggest place first.",
      "Tenths: 6, then 4, then 2.",
      "",
      "Smallest to largest: 0.275, 0.45, 0.6.",
      "0.275 looks longest but is smallest.",
    ],
    NOTES_IDO_A, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.1, { strip: C.PRIMARY });
      slide.addText("Compare the tenths first", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      drawDecimalCompare(slide, lg.rightX + 0.2, lg.panelTopPadded + 0.5, lg.rightW - 0.4,
        [["0", "6", "0", "0"], ["0", "4", "5", "0"], ["0", "2", "7", "5"]]);
    }
  );

  // Slides 11-12: CFU A + reveal - order three decimals
  withReveal(
    () => cfuSlide(pres, "CFU", "Order from smallest to largest",
      { technique: "Show Me Boards",
        question: "Order these from smallest to largest:\n\n0.3        0.275        0.31\n\nLine them up by place value first." },
      NOTES_CFU_A_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "0.275,  0.3,  0.31", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: I Do B - multiply and divide by powers of 10
  workedExSlide(pres, 2, "I Do", "Powers of 10: digits move, point stays",
    [
      "Start with 6.3.",
      "x 10: digits move one place -> 63.",
      "x 100: digits move two places -> 630.",
      "÷ 10: digits move one place smaller -> 0.63.",
      "",
      "The point stays. The digits change place value.",
    ],
    NOTES_IDO_B, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("6.3", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.34,
        fontSize: 20, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addPlaceValueChart(slide, lg.rightX + 1.05, lg.panelTopPadded + 0.55,
        ["Ones", "Tenths"], ["6", "3"], { totalW: 2.0 });
      const chips = [
        ["6.3 x 10 = 63", C.SECONDARY],
        ["6.3 x 100 = 630", C.SECONDARY],
        ["6.3 ÷ 10 = 0.63", C.PRIMARY],
      ];
      const cy0 = lg.panelTopPadded + 1.85;
      chips.forEach((ch, i) => {
        addTextOnShape(slide, ch[0], {
          x: lg.rightX + 0.45, y: cy0 + i * 0.36, w: lg.rightW - 0.9, h: 0.30, rectRadius: 0.06,
          fill: { color: ch[1] },
        }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slides 14-15: We Do B + reveal - powers of 10 pair
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your turn: powers of 10",
      [
        "With your partner.",
        "",
        "1.  5.2 x 100 = ?",
        "2.  0.6 ÷ 10 = ?",
        "",
        "Say what happens to the digits.",
      ],
      NOTES_WEDO_B_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
        slide.addText("Move the digits,\nnot the point", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.2, w: lg.rightW - 0.4, h: 0.8,
          fontSize: 17, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        addTextOnShape(slide, "0.6 ÷ 10\n= 6 hundredths", {
          x: lg.rightX + 0.6, y: lg.panelTopPadded + 1.2, w: lg.rightW - 1.2, h: 0.9, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      }),
    (slide) => {
      addTextOnShape(slide, "5.2 x 100 = 520        0.6 ÷ 10 = 0.06", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_B_A);
    }
  );

  // Slide 16: I Do C - decimals and integers on number lines (full width)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Number lines: decimals and integers", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.15, { strip: STAGE_COLORS["2"] });
    s.addText([
      { text: "A number line orders any numbers. ", options: { fontSize: 15, color: C.CHARCOAL } },
      { text: "2.335 sits halfway between 2.33 and 2.34.   ", options: { fontSize: 15, color: C.PRIMARY, bold: true } },
      { text: "Negatives sit left of zero: -3 < 0 < 2.   ", options: { fontSize: 15, color: C.SECONDARY, bold: true } },
      { text: "Further left is smaller.", options: { fontSize: 15, color: C.CHARCOAL, italic: true } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.5, h: 0.9, fontFace: FONT_B, valign: "middle", margin: 0,
    });

    s.addText("Decimals: zoom between 2.33 and 2.34", {
      x: 0.7, y: 2.7, w: 8.6, h: 0.26, fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    addNumberLine(s, 0.9, 3.25, 8.0,
      ["2.33", "", "", "", "", "2.335", "", "", "", "", "2.34"], [5]);

    s.addText("Integers: negatives, zero, positives", {
      x: 0.7, y: 3.75, w: 8.6, h: 0.26, fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    addNumberLine(s, 0.9, 4.35, 8.0,
      ["-5", "-4", "-3", "-2", "-1", "0", "1", "2", "3", "4", "5"], [2, 7]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_C);
    runSlideDiagnostics(s, pres);
  })();

  // Slides 17-18: We Do C + reveal - order integers on a line (full width)
  const WEDOC_LINE_X = 0.9, WEDOC_LINE_W = 8.0, WEDOC_LINE_Y = 3.55;
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Place and order on the line", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 1.4, { strip: STAGE_COLORS["3"] });
      s.addText("With your partner. Place these on the line, then order them from smallest to largest:", {
        x: 0.75, y: CONTENT_TOP + 0.14, w: 8.5, h: 0.4, fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
      addTextOnShape(s, "-4      -1      3      -3      2", {
        x: 1.5, y: CONTENT_TOP + 0.62, w: 7.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });

      s.addText("Remember: further left is smaller.", {
        x: 0.7, y: 3.05, w: 8.6, h: 0.26, fontSize: 13, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", margin: 0,
      });
      addNumberLine(s, WEDOC_LINE_X, WEDOC_LINE_Y, WEDOC_LINE_W,
        ["-5", "-4", "-3", "-2", "-1", "0", "1", "2", "3", "4", "5"]);
      addFooter(s, FOOTER);
      return s;
    },
    (slide) => {
      [-4, -1, 3, -3, 2].forEach((v) => markInteger(slide, WEDOC_LINE_X, WEDOC_LINE_Y, WEDOC_LINE_W, v, C.ALERT));
      addTextOnShape(slide, "-4,  -3,  -1,  2,  3", {
        x: 0.5, y: 4.35, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_C_A);
      runSlideDiagnostics(slide, pres);
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
      { text: "compare and order the decimals.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "multiply and divide by powers of 10.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "place the integers on the line.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Compare by place value. Multiplying by 10 moves the digits, not the point.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    s.addText("Reference line:", {
      x: 0.7, y: panelY + 1.08, w: 2.2, h: 0.28,
      fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    addNumberLine(s, 1.2, panelY + 1.7, 7.6,
      ["-5", "-4", "-3", "-2", "-1", "0", "1", "2", "3", "4", "5"]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 20: Exit Ticket
  exitTicketSlide(pres,
    [
      "Order from smallest to largest: 0.4, 0.39, 0.405",
      "Work out 3.7 x 100.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 21: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: when we multiply 5.6 by 10, what happens to the digits and what happens to the point?",
      scItems: [
        "I can compare two decimals using place value.",
        "I can order decimals and multiply or divide them by powers of 10.",
        "I can place integers on a number line and explain the powers-of-10 effect.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "TermRev_Lesson1_Decimals_And_Integers.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Compare and order decimals, use powers of 10, and place integers on a line.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Line decimals up by place value and compare the biggest place first - more digits does not mean a bigger number. To multiply or divide by 10, 100 or 1000, move the digits to a new place; the point stays put.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Order 0.6, 0.45, 0.275. Write each to thousandths: 0.600, 0.450, 0.275. Compare the tenths: 6, 4, 2. Smallest to largest: 0.275, 0.45, 0.6.",
      y);

    y = addSectionHeading(doc, "Section 1 - Compare and order (for everyone)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  Circle the larger:   0.7   or   0.65", y);
    y = addWriteLine(doc, "b)  Circle the larger:   0.305   or   0.4", y);
    y = addWriteLine(doc, "c)  Order smallest to largest:  0.5,  0.55,  0.505   ->", y);

    y = addSectionHeading(doc, "Section 2 - Multiply and divide by powers of 10", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  4.7 x 10 = ______", y);
    y = addWriteLine(doc, "b)  2.6 x 100 = ______", y);
    y = addWriteLine(doc, "c)  38.0 ÷ 10 = ______", y);
    y = addWriteLine(doc, "d)  0.6 ÷ 10 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Integers on a number line", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Mark these on the line below, then order them: -3, 1, -5, 4, 0.", y);
    y = addWriteLine(doc, "Order (smallest to largest): ________________________________", y);

    y = addSectionHeading(doc, "Challenge (optional)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "Explain in place value words why 4.2 x 100 is NOT 4.200.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    addPdfFooter(doc, `Session ${SESSION} | Decimals and integers | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Session 1 review sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Compare and order", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  0.7 is larger (7 tenths beats 6 tenths).   b)  0.4 is larger (4 tenths beats 3 tenths).   c)  0.5, 0.505, 0.55.", y);

    y = addSectionHeading(doc, "Section 2 - Powers of 10", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  47        b)  260        c)  3.8        d)  0.06", y);

    y = addSectionHeading(doc, "Section 3 - Integers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Order smallest to largest:  -5,  -3,  0,  1,  4.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "4.2 x 100 moves the digits two places, giving 420. Writing 4.200 just adds zeros after the point and keeps the value 4.2 - the digits did not change place value.", y);

    y = addTipBox(doc,
      "Watch for: students who order by digit count instead of place value; students who 'add a zero' to multiply a decimal by 10; students who flip negatives (thinking -5 is bigger than -3).",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Powers of 10 and standard form (scientific notation).",
      color: C.SECONDARY,
      lessonInfo: `Session ${SESSION} | Year 8 challenge | extends Year 6 VC2M6N06`,
    });
    y = addTipBox(doc,
      "The same powers-of-10 idea grows into standard form. A number in standard form is written as a x 10^n, where a is at least 1 but less than 10. Multiplying by 10 raises the index by 1; dividing by 10 lowers it by 1.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Write 4 300 in standard form. Move the digits so there is one digit before the point: 4.3. The point moved 3 places, so 4 300 = 4.3 x 10^3.",
      y);

    y = addSectionHeading(doc, "Section 1 - Write in standard form", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  52 000 = ______", y);
    y = addWriteLine(doc, "b)  600 = ______", y);
    y = addWriteLine(doc, "c)  7 800 000 = ______", y);

    y = addSectionHeading(doc, "Section 2 - Back to an ordinary number", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3.2 x 10^4 = ______", y);
    y = addWriteLine(doc, "b)  9 x 10^2 = ______", y);
    y = addWriteLine(doc, "c)  1.05 x 10^3 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Which is larger, 6.2 x 10^4 or 8.1 x 10^3? Explain how you know.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) 5.2 x 10^4  b) 6 x 10^2  c) 7.8 x 10^6.   S2  a) 32 000  b) 900  c) 1 050.   S3  6.2 x 10^4 = 62 000 is larger than 8.1 x 10^3 = 8 100; compare the power of 10 first.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Session ${SESSION} | Year 8 Extension | Standard form`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Session 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
