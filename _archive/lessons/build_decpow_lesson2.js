"use strict";

// Decimal Operations & Place Value Mastery (Year 6 Numeracy) - Lesson 2 of 4.
// Divide decimals by 10, 100 and 1000. Place value: every digit moves to a
// place worth 10/100/1000 times LESS (digits move RIGHT). The decimal point
// does NOT move - the digits do. VC2M6N06.
// Daily Review: Advanced Number Concepts and Representation (prior).
// Fluency: vertical multiplication algorithm.
// Mirror of Lesson 1; unit variant fixed (variant 2) for cohesion.
// Catch-up: launch re-activates the whole-number divide-by-10 move; the
// worksheet has an enabling start. No session assumes the one before it.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(3)); // variant 2, fixed for the unit
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide, dailyReviewSlide,
  addPlaceValueChart, addDecimalDot,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 4;
const UNIT_TITLE = "Decimal Operations and Place Value";
const FOOTER = `Decimal Operations | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/DecPow_Lesson2_Divide_By_Powers_Of_10";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Divide decimals by 10, 100 and 1000 using the place value move.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 2 practice sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Dividing by powers of 10 in index form and combined moves.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to our decimal operations unit.
- Last lesson we multiplied decimals by 10, 100 and 1000 and the digits moved left.
- Today we do the opposite. We divide by 10, 100 and 1000, and the digits move the other way - to the right - while the point stays still.

DO:
- Have whiteboards, markers and the printed place value chart ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 2 of 4. This mirrors Lesson 1. If a student missed Lesson 1, the launch rebuilds the whole-number move, so they can still access today.

WATCH FOR:
- Students who remember 'digits move, point stays' - strong start.
- Students who look unsure - reassure, we build it together.

[General: Title | Element: Planning and Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end. The Year 8 extension is there for anyone ready for more.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards, markers and the printed place value chart ready.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student practice sheet, an answer key, and a Year 8 extension. Most teaching happens on whiteboards with the printed place value chart.

CATCH-UP NOTE:
A student who missed Lesson 1 can still do today. The launch re-activates the whole-number divide move, and Section 1 of the worksheet rebuilds it. A returner only needs the printed chart and one minute with you. Today is the mirror of Lesson 1, so the same chart and the same phrase carry both lessons.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are keeping our number knowledge sharp from earlier work.
- Read each one carefully and work it on your whiteboard.
- Think about the value of each digit.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for place value language about thousands and hundreds.

TEACHER NOTES:
Daily Review is prior learning - advanced number concepts and representation - not today's decimals. It keeps place value of larger numbers ticking over, which supports today's work.

WATCH FOR:
- Students who name the value of a digit correctly - secure.
- Students who say 'five' instead of 'fifty thousand' - prompt for the place.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 40 000 plus 600 plus 7 is 40 607.
- In 253 184 the 5 is worth 50 000, fifty thousand.
- 6 482 rounded to the nearest hundred is 6 500.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The value-of-the-digit question is the key one. A student who says 'five' has not used place value. Note them for small group.

WATCH FOR:
- Students who self-correct quickly - secure.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Vertical multiplication, now two digits times two digits.
- Set each one out vertically and multiply, keeping your columns lined up.
- Add your partial products carefully.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Scan for clean column alignment and correct partial products.

TEACHER NOTES:
Same multiplication algorithm as Lesson 1, now two digits by two digits. Next lesson we move to lattice. Lining up columns is the same place value habit we use with decimals.

WATCH FOR:
- Students who line up partial products neatly - secure.
- Students who forget the zero in the second partial product - prompt them about the tens row.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 34 times 12 is 408.
- 26 times 23 is 598.
- 45 times 14 is 630.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The usual error is the missing zero when multiplying by the tens digit. Coach the place value of that second row.

WATCH FOR:
- Students who self-correct - secure.
- Students who drop the tens-row zero - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember that 50 divided by 10 is 5.
- Look what happened: the 5 moved one place to the right, from the tens into the ones, and became ten times smaller.
- Divide by 10 again and 5 becomes 0.5. The digit keeps sliding right.
- Today we make exactly the same move with decimals.

DO:
- Point to the digit moving one place right as you say it.
- Have students chorus 'one place to the right' as you point.
- Bridge: 'the move is the same, the point will stay still'.

TEACHER NOTES:
This launch starts from whole numbers everyone can access, then connects that known move to today's decimals. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who say 'it gets ten times smaller' - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to divide decimals by 10, 100 and 1000 using place value.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Hold up the printed place value chart.

TEACHER NOTES:
The first criterion is reachable for everyone - divide by 10. The second is the core target the exit ticket checks. The third stretches to explaining why the digits move right and the point stays still.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today, the same family as last lesson.
- A power of 10 is still 10, 100 or 1000.
- Ten times smaller means each digit moves to a place worth ten times less.
- Place value means where a digit sits decides its value.

DO:
- Point to each word as you say it.
- Have students say 'the digits move, the point stays' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. The anchor phrase is the same as Lesson 1 on purpose - it carries across the whole unit and helps returning students.

WATCH FOR:
- Students who connect to last lesson - secure.
- Students who think divide means 'take a zero off' - we correct that in the I Do.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at this one together. Watch how I divide 47 by 10.
- I start with 4 tens and 7 ones. That is 47.
- When I divide by 10, every digit moves one place to the right and becomes ten times smaller.
- The 4 tens become 4 ones. The 7 ones become 7 tenths.
- So 47 divided by 10 is 4.7. Notice the point did not move - the digits did.

DO:
- Fill the 'after' chart cell by cell as you speak.
- Trace the one-place-right move with your finger for each digit.
- Say 'the digits move, the point stays' as you finish.

TEACHER NOTES:
This is the core move of the lesson, the mirror of Lesson 1. Keep saying it in place value words. Resist 'move the decimal point'.

MISCONCEPTIONS:
- Misconception: students think dividing by 10 means 'take a zero off the end'.
  Why: it looks that way for round whole numbers like 470 divided by 10.
  Impact: it fails for numbers like 47 or 6.2, where there is no zero to remove.
  Quick correction: show every digit moving one place right on the chart.

WATCH FOR:
- Students who track the digits sliding right - secure.
- Students who 'remove a zero' - re-point to the chart move.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Watch how the pattern grows.
- Divide by 10 moves every digit one place right. Divide by 100 moves them two places. Divide by 1000 moves them three places.
- So 380 divided by 100 is 3.8, and 380 divided by 1000 is 0.38.
- Look at 0.38. When the digits move past the ones place, a zero holds the ones place and we write the leading zero before the point.
- The number of zeros in the power of 10 tells me how many places each digit moves.

DO:
- Reveal the pattern rows one at a time.
- Circle the leading zero in 0.38 and name its job.
- Have students chorus 'two zeros, two places; three zeros, three places'.

TEACHER NOTES:
This connects the number of zeros to the number of places moved. The leading zero in 0.38 is the key teaching point - it shows there are no whole ones.

MISCONCEPTIONS:
- Misconception: students drop the leading zero and write .38 or 38.
  Why: they do not see the empty ones place as needing a zero.
  Impact: 0.38 and 38 get confused.
  Quick correction: the ones place is empty after the move, so a 0 holds it - show it on the chart.

WATCH FOR:
- Students who write the leading zero - secure.
- Students who lose track of places - slow down and count the move out loud.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Work out 5 divided by 100.
- Then tell your partner what happened to the 5.

DO:
- Display the prompt.
- Give 45 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 0.05, with the 5 moving from ones two places right to the hundredths, and zeros holding the ones and tenths.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students write 0.5 because they moved the 5 only one place.
- Reteach: rebuild 5 on the chart and move it two places right. The 5 lands in the hundredths. Zeros hold the ones and tenths. The answer is 0.05.
- Re-check: how many places does the 5 move, and what holds the tenths place?

TEACHER NOTES:
The trap is moving only one place. A student who writes 0.5 has divided by 10, not 100. Make them count two places.

WATCH FOR:
- Students who land the 5 in the hundredths - secure.
- Students who write 0.5 - they under-moved, reteach with the chart.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Build 6.2 divided by 10 on your whiteboard chart.
- Move each digit one place to the right.
- Whisper to your partner: where does the 6 go, and where does the 2 go?

DO:
- Display 6.2 divided by 10 above the blank chart.
- Give 75 seconds.
- Listen for 'the 6 goes to tenths, the 2 goes to hundredths'.

TEACHER NOTES:
Same move as the I Do with new numbers. Listen for place value language, not 'move the point'.

WATCH FOR:
- Pairs who land 6 in tenths and 2 in hundredths - secure.
- Pairs who forget the leading zero - remind them the ones place needs a 0.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- The 6 moves from ones to tenths. The 2 moves from tenths to hundredths.
- The ones place is empty, so a zero holds it.
- So 6.2 divided by 10 is 0.62. The point stayed still and the digits moved one place right.

DO:
- Click to reveal.
- Run the digit moves once more together.

TEACHER NOTES:
Reveal restates the place value move. The leading zero in 0.62 is the teaching point.

WATCH FOR:
- Students who self-correct - secure.
- Students who write .62 - remind them to write the leading zero.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, a tricky one.
- Build 25 divided by 1000 on your whiteboard chart.
- The 2 is in the tens and the 5 is in the ones. How many places does each move, and what fills the empty places?

DO:
- Display 25 divided by 1000 above the blank chart.
- Give 90 seconds.
- Watch for students who move the digits only one or two places.

TEACHER NOTES:
This is the hardest case in the lesson. Each digit moves three places: the 2 from tens to hundredths, the 5 from ones to thousandths. The ones and tenths are empty, so zeros hold them. The answer is 0.025.

WATCH FOR:
- Students who write 0.025 - secure.
- Students who write 0.25 or 0.0025 - they miscounted the move, reteach the count on the chart.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- The 2 moves three places right to the hundredths. The 5 moves three places right to the thousandths.
- The ones and tenths are empty, so zeros hold them, with a leading zero before the point.
- So 25 divided by 1000 is 0.025.

DO:
- Click to reveal.
- Point to the zeros holding the ones and tenths.

TEACHER NOTES:
If many wrote 0.25, do one more divide-by-1000 example before releasing to the You Do.

WATCH FOR:
- Students who place 2 in hundredths and 5 in thousandths - ready for independent work.
- Students who under-move - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 warms up with the divide by 10 move.
- Section 2 mixes divide by 10, 100 and 1000.
- Section 3 asks you to explain a move in place value words. If you finish, try the extension box.

DO:
- Distribute the practice sheet.
- Circulate and listen for 'the digits move, the point stays'.
- Cold call one or two students to explain a divide by 1000 answer.

TEACHER NOTES:
Different numbers from the We Do, same move: chart, slide the digits right, fill empty places with zeros, write the leading zero.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 is started for you with the printed chart. Do the divide by 10 row first, then build up.
- Extra Notes: Sit with these students and move the first digit together. This is also the rebuild for any returning student.
EXTENDING PROMPT:
- Task: The extension box asks students to undo a move (if the answer is 0.6, what was divided by 100) and to explain why 'take a zero off' fails for decimals.
- Extra Notes: Students who are ready can move on to the Year 8 extension sheet.

WATCH FOR:
- Students who use place value language fluently - secure.
- Students who 'take a zero off' - prompt them to point to where each digit lands.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- Work out 8 divided by 100.
- Then explain in place value words what happened to the 8. Do not say the point moved.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - divide a number by a power of 10 by moving each digit. Look for 0.08, and for an explanation that moves the 8 two places right with zeros holding the ones and tenths.

WATCH FOR:
- Students who write 0.08 and explain the two-place move - secure.
- Students who write 0.8 - they under-moved, revisit at the start of Lesson 3.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: when we divide by 100, which way do the digits move, and what stays still?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that dividing by a power of 10 moves every digit right and the point stays still. Students who can say this, and who can multiply too, are ready for multiples of powers of 10 next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on the core criterion - small group revision at the start of Lesson 3.

[General: Closing | Element: Retention and Recall]`;

// --- Helpers -----------------------------------------------------------------

function pvChart(slide, x, y, headers, values, opts) {
  const o = opts || {};
  const chart = addPlaceValueChart(slide, x, y, headers, values, {
    totalW: o.totalW || 4.0,
    valH: o.valH || 0.50,
    hdrH: o.hdrH || 0.34,
    headerColor: o.headerColor || C.PRIMARY,
  });
  if (o.dotAfter != null) {
    addDecimalDot(slide, chart, o.dotAfter, { position: "baseline" });
  }
  return chart;
}

// Before/after "digits move" panel used on the I Do and We Do slides.
function addMovePanel(slide, lg, cfg) {
  const color = cfg.color || C.PRIMARY;
  const cardH = cfg.cardH || 3.5;
  const top = lg.panelTopPadded;
  addCard(slide, lg.rightX, top, lg.rightW, cardH, { strip: color });

  slide.addText(cfg.beforeLabel, {
    x: lg.rightX + 0.15, y: top + 0.10, w: lg.rightW - 0.30, h: 0.28,
    fontSize: 14, fontFace: FONT_H, color: color, bold: true,
    align: "center", margin: 0,
  });
  pvChart(slide, lg.rightX + 0.20, top + 0.42, cfg.headers, cfg.beforeVals,
    { totalW: lg.rightW - 0.40, valH: 0.46, hdrH: 0.34, dotAfter: cfg.dotAfter });

  addTextOnShape(slide, cfg.arrow, {
    x: lg.rightX + 0.30, y: top + 1.40, w: lg.rightW - 0.60, h: 0.36, rectRadius: 0.08,
    fill: { color: C.ALERT },
  }, { fontSize: 12.5, fontFace: FONT_H, color: C.WHITE, bold: true });

  slide.addText(cfg.afterLabel, {
    x: lg.rightX + 0.15, y: top + 1.86, w: lg.rightW - 0.30, h: 0.28,
    fontSize: 14, fontFace: FONT_H, color: C.SUCCESS, bold: true,
    align: "center", margin: 0,
  });
  pvChart(slide, lg.rightX + 0.20, top + 2.18, cfg.headers, cfg.afterVals,
    { totalW: lg.rightW - 0.40, valH: 0.46, hdrH: 0.34, dotAfter: cfg.dotAfter });

  if (cfg.note) {
    slide.addText(cfg.note, {
      x: lg.rightX + 0.15, y: top + 3.06, w: lg.rightW - 0.30, h: 0.34,
      fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", margin: 0,
    });
  }
}

// Headers: Tens | Ones | Tenths | Hundredths, decimal point after Ones (index 1).
const PV_TOH = ["Tens", "Ones", "Tenths", "Hundredths"];
// Headers: Ones | Tenths | Hundredths | Thousandths, decimal point after Ones (index 0).
const PV_OTH = ["Ones", "Tenths", "Hundredths", "Thousandths"];
// Headers spanning Tens..Thousandths, decimal point after Ones (index 1).
const PV_WIDE = ["Tens", "Ones", "Tenths", "Hundredths", "Thousandths"];

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Dividing decimals by 10, 100 and 1000",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - advanced number concepts
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Number sense",
      [
        "Write 40 000 + 600 + 7 as one number.",
        "What is the value of the 5 in 253 184?",
        "Round 6 482 to the nearest hundred.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "40 607          5 is worth 50 000          6 500", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - vertical multiplication (2 x 2 digit)
  withReveal(
    () => fluencySlide(pres, "Fluency: Vertical multiplication",
      ["34 x 12", "26 x 23", "45 x 14"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "408        598        630", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - from whole numbers (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "We know this move already",
    [
      "Some of you may remember: 50 ÷ 10 = 5.",
      "The 5 moved one place RIGHT.",
      "",
      "Divide by 10 makes every digit 10 times smaller.",
      "Today: the same move, with decimals.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.ACCENT });
      slide.addText("From whole numbers", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      const chips = [["50 ÷ 10 = 5", "1 place right"], ["5 ÷ 10 = 0.5", "1 place right"]];
      const cy0 = lg.panelTopPadded + 0.58;
      chips.forEach((r, i) => {
        const cy = cy0 + i * 0.64;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.25, y: cy, w: 2.35, h: 0.50, rectRadius: 0.07,
          fill: { color: C.PRIMARY },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(r[1], {
          x: lg.rightX + 2.7, y: cy, w: lg.rightW - 2.9, h: 0.50,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
      slide.addText("The digits move. The point stays still.", {
        x: lg.rightX + 0.15, y: cy0 + 1.34, w: lg.rightW - 0.30, h: 0.34,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to divide decimals by 10, 100 and 1000 using place value.",
    [
      "I can divide a number by 10 and say the new number.",
      "I can divide a decimal by 10, 100 or 1000 by moving each digit to the correct place.",
      "I can explain why the digits move right and the decimal point stays still.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Power of 10 = 10, 100 or 1000",
      "10 times smaller = each digit moves to a place worth 10x less",
      "Place value = where a digit sits decides its value",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
      slide.addText("Powers of 10", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [["÷ 10", "1 place right"], ["÷ 100", "2 places right"], ["÷ 1000", "3 places right"]];
      const ry0 = lg.panelTopPadded + 0.55;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.56;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.35, y: ry, w: 1.9, h: 0.46, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(r[1], {
          x: lg.rightX + 2.45, y: ry, w: lg.rightW - 2.7, h: 0.46,
          fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL,
          align: "center", valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10: I Do #1 - 47 / 10 (before/after move)
  workedExSlide(pres, 2, "I Do", "Divide 47 ÷ 10",
    [
      "Start with 47.",
      "÷ 10: every digit moves ONE place RIGHT.",
      "4 tens -> 4 ones",
      "7 ones -> 7 tenths",
      "",
      "47 ÷ 10 = 4.7",
      "The digits moved. The point stayed.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addMovePanel(slide, lg, {
        color: C.PRIMARY,
        headers: PV_TOH, dotAfter: 1,
        beforeLabel: "Before: 47",
        beforeVals: ["4", "7", "", ""],
        arrow: "÷ 10  ->  move 1 place RIGHT",
        afterLabel: "After: 4.7",
        afterVals: ["", "4", "7", ""],
        note: "The point did not move - the digits did.",
      });
    }
  );

  // Slide 11: I Do #2 - / 100 and / 1000 pattern
  workedExSlide(pres, 2, "I Do", "÷ 100 and ÷ 1000: more places",
    [
      "÷ 10 -> move 1 place right.",
      "÷ 100 -> move 2 places right.",
      "÷ 1000 -> move 3 places right.",
      "",
      "Zeros count the places.",
      "Write the leading 0 before the point.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.1, { strip: C.PRIMARY });
      slide.addText("Start: 380", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [
        ["380 ÷ 10 = 38", "1 place"],
        ["380 ÷ 100 = 3.8", "2 places"],
        ["380 ÷ 1000 = 0.38", "3 places"],
      ];
      const ry0 = lg.panelTopPadded + 0.55;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.62;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.20, y: ry, w: 2.75, h: 0.52, rectRadius: 0.07,
          fill: { color: i === 2 ? C.SECONDARY : C.PRIMARY },
        }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(r[1], {
          x: lg.rightX + 3.05, y: ry, w: lg.rightW - 3.2, h: 0.52,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
      slide.addText("In 0.38 the leading 0 shows there are no whole ones.", {
        x: lg.rightX + 0.15, y: ry0 + 1.92, w: lg.rightW - 0.30, h: 0.34,
        fontSize: 11.5, fontFace: FONT_B, color: C.ALERT, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slides 12-13: CFU + reveal - 5 / 100
  withReveal(
    () => cfuSlide(pres, "CFU", "Work out 5 ÷ 100",
      { technique: "Show Me Boards",
        question: "On your whiteboard: 5 ÷ 100 = ?\n\nThen tell your partner what happened to the 5." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "5 ÷ 100 = 0.05   (the 5 moves 2 places right to the hundredths; zeros hold the ones and tenths)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 6.2 / 10
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Build 6.2 ÷ 10 together",
      [
        "With your partner.",
        "",
        "1.  Move each digit one place RIGHT.",
        "2.  Where does the 6 go? The 2?",
        "3.  Write the leading zero.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
        slide.addText("6.2 ÷ 10", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.42,
          fontSize: 28, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        pvChart(slide, lg.rightX + 0.20, lg.panelTopPadded + 0.66, PV_OTH,
          ["", "", "", ""],
          { totalW: lg.rightW - 0.40, valH: 0.80, hdrH: 0.50, dotAfter: 0 });
        slide.addText("Slide the digits right.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.06, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "6.2 ÷ 10 = 0.62   (6 -> tenths, 2 -> hundredths; a 0 holds the ones place)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - 25 / 1000 (placeholder zeros)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "A tricky one: 25 ÷ 1000",
      [
        "With your partner.",
        "",
        "2 is in the tens, 5 is in the ones.",
        "÷ 1000 -> move 3 places RIGHT.",
        "What fills the empty places?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
        slide.addText("25 ÷ 1000", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.42,
          fontSize: 26, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        pvChart(slide, lg.rightX + 0.15, lg.panelTopPadded + 0.66, PV_WIDE,
          ["2", "5", "", "", ""],
          { totalW: lg.rightW - 0.30, valH: 0.80, hdrH: 0.50, dotAfter: 1 });
        slide.addText("Count the move: 3 places.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.06, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "25 ÷ 1000 = 0.025   (2 -> hundredths, 5 -> thousandths; zeros hold the ones and tenths)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "warm up with the ÷ 10 row.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "mix ÷ 10, ÷ 100 and ÷ 1000.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "explain one move in place value words.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "The digits move RIGHT. The point stays still. Write the leading zero.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    pvChart(s, 2.5, panelY + 1.15, PV_OTH, ["", "", "", ""],
      { totalW: 5.0, valH: 0.55, hdrH: 0.38, dotAfter: 0 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Work out 8 ÷ 100.",
      "Explain in place value words what happened to the 8 (do not say the point moved).",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: when we divide by 100, which way do the digits move, and what stays still?",
      scItems: [
        "I can divide a number by 10 and say the new number.",
        "I can divide a decimal by 10, 100 or 1000 by moving each digit to the correct place.",
        "I can explain why the digits move right and the decimal point stays still.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecPow_Lesson2_Divide_By_Powers_Of_10.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Divide decimals by 10, 100 and 1000 using the place value move.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Dividing by a power of 10 moves every digit RIGHT: ÷10 one place, ÷100 two places, ÷1000 three places. The digits move, the point stays still. A zero holds any empty place, and write the leading zero before the point.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "47 ÷ 100: the 4 moves from tens to tenths, the 7 moves from ones to hundredths. There are no whole ones, so a 0 holds the ones place. 47 ÷ 100 = 0.47.",
      y);

    y = addSectionHeading(doc, "Section 1 - Warm up: divide by 10 (started for you)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  86 ÷ 10 = 8.6        (each digit moves 1 place right)", y);
    y = addWriteLine(doc, "b)  9 ÷ 10 = _______", y);
    y = addWriteLine(doc, "c)  53 ÷ 10 = _______", y);
    y = addWriteLine(doc, "d)  4.5 ÷ 10 = _______", y);

    y = addSectionHeading(doc, "Section 2 - Mix it up: ÷ 10, ÷ 100, ÷ 1000", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  380 ÷ 100 = _______", y);
    y = addWriteLine(doc, "b)  6.5 ÷ 10 = _______", y);
    y = addWriteLine(doc, "c)  25 ÷ 1000 = _______", y);
    y = addWriteLine(doc, "d)  9 ÷ 100 = _______", y);
    y = addWriteLine(doc, "e)  47.2 ÷ 10 = _______", y);

    y = addSectionHeading(doc, "Section 3 - Explain a move", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "25 ÷ 1000 = _______.  Explain in place value words where the 2 and the 5 go:", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Undo the move: a number divided by 100 gives 0.6. What was the number? __________", y);
    y = addWriteLine(doc, "Explain why 'just take a zero off' does not work for 8 ÷ 100:", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Divide by Powers of 10 | Year 6 Numeracy`);
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

    y = addSectionHeading(doc, "Section 1 - Divide by 10", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  8.6        b)  0.9        c)  5.3        d)  0.45", y);

    y = addSectionHeading(doc, "Section 2 - Mix it up", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3.8        b)  0.65        c)  0.025        d)  0.09        e)  4.72", y);
    y = addBodyText(doc, "Note: 9 ÷ 100 = 0.09 (the 9 moves two places right to the hundredths; zeros hold the ones and tenths).", y, { fontSize: 10, color: C.MUTED });

    y = addSectionHeading(doc, "Section 3 - Explain a move", y, { color: C.PRIMARY });
    y = addBodyText(doc, "25 ÷ 1000 = 0.025. The 2 moves three places right from tens to hundredths, the 5 from ones to thousandths. The ones and tenths are empty, so zeros hold them, with a leading zero before the point.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "The number was 60 (60 ÷ 100 = 0.6). 'Take a zero off' fails because 8 ÷ 100 is 0.08, not 0.8; the 8 must move two places right.", y);

    y = addTipBox(doc,
      "Watch for: students who 'take a zero off' (fails for 47 or 6.2); students who under-move on ÷ 1000; students who drop the leading zero (writing .62 instead of 0.62).",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Dividing by powers of 10 in index form and combined moves.",
      color: C.SECONDARY,
      lessonInfo: `Lesson ${SESSION} | Year 8 challenge | extends Year 6 VC2M6N06`,
    });
    y = addTipBox(doc,
      "Same place value move in reverse. 10 = 10 to the power 1, 100 = 10 squared, 1000 = 10 cubed. Dividing by 10 to the power n moves every digit n places right. Dividing by 1000 is the same as multiplying by 0.001.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "4700 ÷ 10 cubed = 4700 ÷ 1000 = 4.7. The index 3 means move every digit 3 places right.",
      y);

    y = addSectionHeading(doc, "Section 1 - Index form", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  620 ÷ 10 squared = _______", y);
    y = addWriteLine(doc, "b)  35 ÷ 10 cubed = _______", y);
    y = addWriteLine(doc, "c)  9040 ÷ 10 to the power 4 = _______", y);

    y = addSectionHeading(doc, "Section 2 - Combined moves", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  (360 ÷ 10) ÷ 100 = _______", y);
    y = addWriteLine(doc, "b)  5 ÷ 10 squared ÷ 10 = _______", y);
    y = addWriteLine(doc, "c)  Write 0.0042 as a whole number divided by a power of 10 (two different ways).", y);
    y = addWriteLine(doc, "    _________________________   and   _________________________", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Explain why dividing by 10 squared then by 10 is the same as dividing by 10 cubed.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) 6.2  b) 0.035  c) 0.904.   S2  a) 0.36  b) 0.005  c) e.g. 42 ÷ 10 000 or 4.2 ÷ 1000.   S3  The moves add: 2 places then 1 place is 3 places right, which is dividing by 10 cubed.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Year 8 Extension | Powers of 10`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
