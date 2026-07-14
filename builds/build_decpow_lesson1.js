"use strict";

// Decimal Operations & Place Value Mastery (Year 6 Numeracy) - Lesson 1 of 4.
// Multiply decimals by 10, 100 and 1000. Place value: every digit moves to a
// place worth 10/100/1000 times as much (digits move LEFT). The decimal point
// does NOT move - the digits do. VC2M6N06.
// Daily Review: Practical Problem Solving and Financial Reasoning (prior).
// Fluency: vertical multiplication algorithm.
// Unit variant fixed (variant 2) across all 4 lessons for cohesion.
// Catch-up: launch re-activates the whole-number x10 move so a returner can
// rejoin at any lesson; the worksheet has an enabling start.

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

const SESSION = 1;
const TOTAL = 4;
const UNIT_TITLE = "Decimal Operations and Place Value";
const FOOTER = `Decimal Operations | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/DecPow_Lesson1_Multiply_By_Powers_Of_10";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Multiply decimals by 10, 100 and 1000 using the place value move.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 1 practice sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Powers of 10 in index form and combined moves - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to our new unit on decimal operations and place value.
- Over four lessons we learn to multiply and divide decimals by 10, 100 and 1000, and then to estimate sensibly.
- Today we start with multiplying decimals by 10, 100 and 1000, and the big idea is that the digits move and the point stays still.

DO:
- Have whiteboards, markers and the printed place value chart ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 1 of 4. This lesson sets up the whole unit. If students leave today saying 'the digits move, the point stays', the rest of the unit runs smoothly.

WATCH FOR:
- Students who look unsure - that is expected on day one. Reassure them: if this feels new, that is okay, we build it together.

[General: Title | Element: Planning and Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end. The Year 8 extension is there for anyone who is flying.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards, markers and the printed place value chart ready.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student practice sheet, an answer key, and a Year 8 extension. Most teaching happens on whiteboards with the printed place value chart.

CATCH-UP NOTE:
A student who missed earlier sessions can still access today. The launch re-activates the whole-number times-10 move, and Section 1 of the worksheet rebuilds it. A returner only needs the printed chart and one minute with you to start. No session in this unit assumes the student saw the one before.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are warming up our financial problem solving from earlier work.
- Read each one carefully and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for how students reason about money and 'per item'.

TEACHER NOTES:
Daily Review is prior learning, not today's decimals. Money word problems keep practical problem solving and financial reasoning ticking over.

WATCH FOR:
- Students who line up dollars and cents carefully - secure.
- Students who forget the cents place - prompt them to keep two decimal places for money.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Three pencil cases at 7 dollars 50 each is 22 dollars 50.
- A 4 dollar 30 drink paid with a 10 dollar note gives 5 dollars 70 change.
- The 2 litre pack works out at 2 dollars 40 a litre, so it is the better buy.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Best-buy compares price per litre: 4.80 divided by 2 is 2.40, which beats 2.60 for one litre. Note any student who picks the larger total without comparing per litre.

WATCH FOR:
- Students who self-correct quickly - secure.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Vertical multiplication.
- Set each one out vertically on your whiteboard and multiply.
- Keep your columns lined up so your digits land in the right place.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for clean column alignment and correct carrying.

TEACHER NOTES:
Fluency this unit is the multiplication algorithm. We use vertical multiplication in Lessons 1 and 2, then lattice in Lessons 3 and 4. Lining up columns now is the same place value habit we use with decimals today.

WATCH FOR:
- Students who line up columns neatly - secure.
- Students who forget to carry - prompt: what happens when a column makes ten or more.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 47 times 6 is 282.
- 58 times 4 is 232.
- 73 times 8 is 584.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Watch the carry from the ones into the tens. A misaligned column is the usual cause of a wrong total. Keep it brisk.

WATCH FOR:
- Students who self-correct - secure.
- Students whose digits drift out of columns - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember that 5 times 10 is 50.
- Look what happened: the 5 moved one place to the left, from the ones into the tens, and became ten times bigger.
- Multiply by 10 again and 50 becomes 500. The digit keeps sliding left.
- Today we make exactly the same move, but starting with decimals.

DO:
- Point to the digit moving one place left as you say it.
- Have students chorus 'one place to the left' as you point.
- Bridge: 'the move is the same, the point will stay still'.

TEACHER NOTES:
This launch starts from whole numbers everyone can access, then connects that known move to today's decimals. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who say 'it gets ten times bigger' - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to multiply decimals by 10, 100 and 1000 using place value.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Hold up the printed place value chart.

TEACHER NOTES:
The first criterion is reachable for everyone - multiply by 10. The second is the core target the exit ticket checks. The third stretches to explaining why the digits move and the point stays still.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- A power of 10 is 10, 100 or 1000 - ten, then ten times ten, then ten times ten times ten.
- Times ten times bigger means each digit moves to a place worth ten times as much.
- Place value means where a digit sits decides its value.

DO:
- Point to each word as you say it.
- Have students say 'the digits move, the point stays' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. Keep it to these three. The phrase 'the digits move, the point stays' is the anchor for the whole unit.

WATCH FOR:
- Students who can give a power of 10 - secure.
- Students who think multiply means 'add a zero' - we correct that in the I Do.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at this one together. Watch how I multiply 2.45 by 10.
- I start with 2 ones, 4 tenths and 5 hundredths. That is 2.45.
- When I multiply by 10, every digit moves one place to the left and becomes ten times bigger.
- The 2 ones become 2 tens. The 4 tenths become 4 ones. The 5 hundredths become 5 tenths.
- So 2.45 times 10 is 24.5. Notice the point did not move - the digits did.

DO:
- Fill the 'after' chart cell by cell as you speak.
- Trace the one-place-left move with your finger for each digit.
- Say 'the digits move, the point stays' as you finish.

TEACHER NOTES:
This is the core move of the lesson. Keep saying it in place value words. Resist 'move the decimal point' - we describe the digits moving so the idea transfers to division next lesson.

MISCONCEPTIONS:
- Misconception: students think multiplying by 10 means 'add a zero'.
  Why: it works for whole numbers, so they over-generalise.
  Impact: they write 2.45 times 10 as 2.450, which is still 2.45.
  Quick correction: show the digits each moving one place left on the chart.

WATCH FOR:
- Students who track the digits sliding left - secure.
- Students who add a zero on the end - re-point to the chart move.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Watch how the pattern grows.
- Times 10 moves every digit one place left. Times 100 moves them two places. Times 1000 moves them three places.
- So 2.45 times 100 is 245, and 2.45 times 1000 is 2450.
- Look at 2450. When the digits move past the ones place, I write a zero to hold the ones place open.
- The number of zeros in the power of 10 tells me how many places each digit moves.

DO:
- Reveal the pattern rows one at a time.
- Circle the placeholder zero in 2450 and name its job.
- Have students chorus 'two zeros, two places; three zeros, three places'.

TEACHER NOTES:
This connects the number of zeros to the number of places moved. The placeholder zero is the key teaching point - it holds a place, it is not 'an extra zero we add for fun'.

MISCONCEPTIONS:
- Misconception: students think the zeros come from 'adding zeros to the end'.
  Why: the result often ends in zeros, so it looks like that.
  Impact: they mishandle cases like 0.08 times 1000.
  Quick correction: the zeros fill empty places after the digits move - show it on the chart.

WATCH FOR:
- Students who match zeros to places - secure.
- Students who lose track of places - slow down and count the move out loud.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Work out 0.7 times 100.
- Then tell your partner what happened to the 7.

DO:
- Display the prompt.
- Give 45 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 70, with the 7 moving from tenths to tens and a zero holding the ones place.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students write 0.700 or 0.70 because they 'added zeros'.
- Reteach: rebuild 0.7 on the chart and move the 7 two places left. The ones place is empty, so a zero holds it. The answer is 70, not 0.70.
- Re-check: where does the 7 land, and what holds the ones place?

TEACHER NOTES:
The trap is 'adding zeros'. A student who writes 0.700 has not moved the digit at all. Make them point to where the 7 lands.

WATCH FOR:
- Students who land the 7 in the tens - secure.
- Students who write 0.700 - the add-a-zero misconception, reteach with the chart.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Build 5.3 times 10 on your whiteboard chart.
- Move each digit one place to the left.
- Whisper to your partner: where does the 5 go, and where does the 3 go?

DO:
- Display 5.3 times 10 above the blank chart.
- Give 75 seconds.
- Listen for 'the 5 goes to tens, the 3 goes to ones'.

TEACHER NOTES:
Same move as the I Do with new numbers. Listen for place value language, not 'move the point'.

WATCH FOR:
- Pairs who land 5 in tens and 3 in ones - secure.
- Pairs who write 5.30 - the add-a-zero misconception, reteach with the chart.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- The 5 moves from ones to tens. The 3 moves from tenths to ones.
- So 5.3 times 10 is 53. The point stayed still and the digits moved one place left.

DO:
- Click to reveal.
- Run the digit moves once more together.

TEACHER NOTES:
Reveal restates the place value move. 53 has no decimal part, which surprises some students - that is correct.

WATCH FOR:
- Students who self-correct - secure.
- Students who keep a point in 5.30 - remind them 53 is a whole number now.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, a tricky one.
- Build 0.08 times 1000 on your whiteboard chart.
- The 8 starts in the hundredths. How many places does it move, and what fills the empty places?

DO:
- Display 0.08 times 1000 above the blank chart.
- Give 90 seconds.
- Watch for students who move the 8 only one or two places.

TEACHER NOTES:
This is the hardest case in the lesson. The 8 moves three places: hundredths to tenths to ones to tens. The ones place is then empty, so a zero holds it. The answer is 80.

WATCH FOR:
- Students who write 80 - secure.
- Students who write 8 or 0.8 - they did not move the full three places, reteach the count on the chart.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- The 8 moves three places to the left, landing in the tens.
- The ones place is empty, so a zero holds it.
- So 0.08 times 1000 is 80.

DO:
- Click to reveal.
- Point to the placeholder zero in the ones place.

TEACHER NOTES:
If many wrote 8 or 0.8, do one more 1000 example before releasing to the You Do.

WATCH FOR:
- Students who place the 8 in tens with a zero in ones - ready for independent work.
- Students who under-move the 8 - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 warms up with the times 10 move.
- Section 2 mixes times 10, times 100 and times 1000.
- Section 3 asks you to explain a move in place value words. If you finish, try the extension box.

DO:
- Distribute the practice sheet.
- Circulate and listen for 'the digits move, the point stays'.
- Cold call one or two students to explain a times 1000 answer.

TEACHER NOTES:
Different numbers from the We Do, same move: chart, slide the digits left, fill empty places with zeros.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 is started for you with the printed chart. Do the times 10 row first, then build up.
- Extra Notes: Sit with these students and move the first digit together. This is also the rebuild for any returning student.
EXTENDING PROMPT:
- Task: The extension box asks students to undo a move (if the answer is 360, what was multiplied by 100) and to explain why 'add a zero' fails for decimals.
- Extra Notes: Students who are ready can move on to the Year 8 extension sheet.

WATCH FOR:
- Students who use place value language fluently - secure.
- Students who 'add a zero' - prompt them to point to where each digit lands.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- Work out 6.4 times 100.
- Then explain in place value words what happened to the 6 and the 4. Do not say the point moved.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - multiply a decimal by a power of 10 by moving each digit. Look for 640, and for an explanation that moves the digits two places left with a zero holding the ones place.

WATCH FOR:
- Students who write 640 and explain the two-place move - secure.
- Students who write 6.400 - the add-a-zero misconception, revisit at the start of Lesson 2.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: when we multiply by 100, what moves and what stays still?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that multiplying by a power of 10 moves every digit left and the point stays still. Students who can say this are ready to divide by powers of 10 next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on the core criterion - small group revision at the start of Lesson 2.

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

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Multiplying decimals by 10, 100 and 1000",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - financial reasoning
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Money problems",
      [
        "Three pencil cases cost $7.50 each. What is the total?",
        "You pay for a $4.30 drink with a $10 note. How much change?",
        "Better buy: 2 L of juice for $4.80, or 1 L for $2.60?",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "$22.50          $5.70 change          2 L pack ($2.40/L)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - vertical multiplication
  withReveal(
    () => fluencySlide(pres, "Fluency: Vertical multiplication",
      ["47 x 6", "58 x 4", "73 x 8"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "282        232        584", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - from whole numbers (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "We know this move already",
    [
      "Some of you may remember: 5 x 10 = 50.",
      "The 5 moved one place LEFT.",
      "",
      "x 10 makes every digit 10 times bigger.",
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
      const chips = [["5 x 10 = 50", "1 place left"], ["50 x 10 = 500", "1 place left"]];
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
    "We are learning to multiply decimals by 10, 100 and 1000 using place value.",
    [
      "I can multiply a decimal by 10 and say the new number.",
      "I can multiply a decimal by 10, 100 or 1000 by moving each digit to the correct place.",
      "I can explain why the digits move and the decimal point stays still.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Power of 10 = 10, 100 or 1000",
      "10 times bigger = each digit moves to a place worth 10x as much",
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
      const rows = [["10", "1 zero"], ["100", "2 zeros"], ["1000", "3 zeros"]];
      const ry0 = lg.panelTopPadded + 0.55;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.56;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.35, y: ry, w: 1.9, h: 0.46, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(r[1], {
          x: lg.rightX + 2.45, y: ry, w: lg.rightW - 2.7, h: 0.46,
          fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
          align: "center", valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10: I Do #1 - 2.45 x 10 (before/after move)
  workedExSlide(pres, 2, "I Do", "Multiply 2.45 x 10",
    [
      "Start with 2.45.",
      "x 10: every digit moves ONE place LEFT.",
      "2 ones -> 2 tens",
      "4 tenths -> 4 ones",
      "5 hundredths -> 5 tenths",
      "",
      "2.45 x 10 = 24.5",
      "The digits moved. The point stayed.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addMovePanel(slide, lg, {
        color: C.PRIMARY,
        headers: PV_TOH, dotAfter: 1,
        beforeLabel: "Before: 2.45",
        beforeVals: ["", "2", "4", "5"],
        arrow: "x 10  ->  move 1 place LEFT",
        afterLabel: "After: 24.5",
        afterVals: ["2", "4", "5", ""],
        note: "The point did not move - the digits did.",
      });
    }
  );

  // Slide 11: I Do #2 - x100 and x1000 pattern
  workedExSlide(pres, 2, "I Do", "x 100 and x 1000: more places",
    [
      "x 10 -> move 1 place left.",
      "x 100 -> move 2 places left.",
      "x 1000 -> move 3 places left.",
      "",
      "Zeros count the places.",
      "Empty places get a 0 to hold them.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.1, { strip: C.PRIMARY });
      slide.addText("Start: 2.45", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [
        ["2.45 x 10 = 24.5", "1 place"],
        ["2.45 x 100 = 245", "2 places"],
        ["2.45 x 1000 = 2450", "3 places"],
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
      slide.addText("In 2450 the 0 holds the empty ones place.", {
        x: lg.rightX + 0.15, y: ry0 + 1.92, w: lg.rightW - 0.30, h: 0.34,
        fontSize: 11.5, fontFace: FONT_B, color: C.ALERT, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slides 12-13: CFU + reveal - 0.7 x 100
  withReveal(
    () => cfuSlide(pres, "CFU", "Work out 0.7 x 100",
      { technique: "Show Me Boards",
        question: "On your whiteboard: 0.7 x 100 = ?\n\nThen tell your partner what happened to the 7." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "0.7 x 100 = 70   (the 7 moves to the tens; a 0 holds the ones place)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 5.3 x 10
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Build 5.3 x 10 together",
      [
        "With your partner.",
        "",
        "1.  Move each digit one place LEFT.",
        "2.  Where does the 5 go? The 3?",
        "3.  Read the new number.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
        slide.addText("5.3 x 10", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.42,
          fontSize: 28, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        pvChart(slide, lg.rightX + 0.20, lg.panelTopPadded + 0.66, PV_TOH,
          ["", "", "", ""],
          { totalW: lg.rightW - 0.40, valH: 0.80, hdrH: 0.50, dotAfter: 1 });
        slide.addText("Slide the digits left.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.06, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "5.3 x 10 = 53   (5 -> tens, 3 -> ones; the point stays, the digits move)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - 0.08 x 1000 (placeholder zero)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "A tricky one: 0.08 x 1000",
      [
        "With your partner.",
        "",
        "The 8 starts in the hundredths.",
        "x 1000 -> move 3 places LEFT.",
        "What fills the empty places?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
        slide.addText("0.08 x 1000", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.42,
          fontSize: 26, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        pvChart(slide, lg.rightX + 0.20, lg.panelTopPadded + 0.66, PV_TOH,
          ["", "0", "0", "8"],
          { totalW: lg.rightW - 0.40, valH: 0.80, hdrH: 0.50, dotAfter: 1 });
        slide.addText("Count the move: 3 places.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.06, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "0.08 x 1000 = 80   (8 moves 3 places to the tens; a 0 holds the ones place)", {
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
      { text: "warm up with the x 10 row.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "mix x 10, x 100 and x 1000.   ", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "The digits move. The point stays still. Zeros hold the empty places.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    pvChart(s, 2.5, panelY + 1.15, PV_TOH, ["", "", "", ""],
      { totalW: 5.0, valH: 0.55, hdrH: 0.38, dotAfter: 1 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Work out 6.4 x 100.",
      "Explain in place value words what happened to the 6 and the 4 (do not say the point moved).",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: when we multiply by 100, what moves and what stays still?",
      scItems: [
        "I can multiply a decimal by 10 and say the new number.",
        "I can multiply a decimal by 10, 100 or 1000 by moving each digit to the correct place.",
        "I can explain why the digits move and the decimal point stays still.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecPow_Lesson1_Multiply_By_Powers_Of_10.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Multiply decimals by 10, 100 and 1000 using the place value move.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Multiplying by a power of 10 moves every digit LEFT: x10 one place, x100 two places, x1000 three places. The digits move, the point stays still. A zero holds any empty place.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "3.6 x 100: the 3 moves from ones to hundreds, the 6 moves from tenths to tens. The ones place is empty, so a 0 holds it. 3.6 x 100 = 360.",
      y);

    y = addSectionHeading(doc, "Section 1 - Warm up: multiply by 10 (started for you)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  4.2 x 10 = 42        (each digit moves 1 place left)", y);
    y = addWriteLine(doc, "b)  7.5 x 10 = _______", y);
    y = addWriteLine(doc, "c)  0.6 x 10 = _______", y);
    y = addWriteLine(doc, "d)  12.8 x 10 = _______", y);

    y = addSectionHeading(doc, "Section 2 - Mix it up: x 10, x 100, x 1000", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3.45 x 100 = _______", y);
    y = addWriteLine(doc, "b)  2.7 x 1000 = _______", y);
    y = addWriteLine(doc, "c)  0.09 x 100 = _______", y);
    y = addWriteLine(doc, "d)  5.04 x 1000 = _______", y);
    y = addWriteLine(doc, "e)  60.2 x 100 = _______", y);

    y = addSectionHeading(doc, "Section 3 - Explain a move", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "0.08 x 1000 = _______.  Explain in place value words where the 8 goes:", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Undo the move: a decimal multiplied by 100 gives 360. What was the decimal? __________", y);
    y = addWriteLine(doc, "Explain why 'just add a zero' does not work for 4.2 x 10:", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Multiply by Powers of 10 | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 1 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Multiply by 10", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  42        b)  75        c)  6        d)  128", y);

    y = addSectionHeading(doc, "Section 2 - Mix it up", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  345        b)  2700        c)  9        d)  5040        e)  6020", y);
    y = addBodyText(doc, "Note: 0.09 x 100 = 9 (the 9 moves from hundredths to ones). 2.7 x 1000 = 2700: the 2 lands in thousands and the 7 in hundreds, with zeros holding the empty tens and ones.", y, { fontSize: 10, color: C.MUTED });

    y = addSectionHeading(doc, "Section 3 - Explain a move", y, { color: C.PRIMARY });
    y = addBodyText(doc, "0.08 x 1000 = 80. The 8 moves three places left from hundredths to the tens. The ones place is empty, so a 0 holds it.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "The decimal was 3.6 (3.6 x 100 = 360). 'Add a zero' fails because 4.2 x 10 is 42, not 4.20; the digits must each move one place left.", y);

    y = addTipBox(doc,
      "Watch for: students who 'add a zero' (4.2 x 10 written as 4.20); students who under-move on x 1000; students who forget the placeholder zero in cases like 0.7 x 100 = 70.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Powers of 10 in index form and combined moves.",
      color: C.SECONDARY,
      lessonInfo: `Lesson ${SESSION} | Year 8 challenge | extends Year 6 VC2M6N06`,
    });
    y = addTipBox(doc,
      "Same place value move, written more powerfully. 10 = 10 to the power 1, 100 = 10 squared, 1000 = 10 cubed. The index tells you how many places each digit moves.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "4.7 x 10 cubed = 4.7 x 1000 = 4700. The index 3 means move every digit 3 places left.",
      y);

    y = addSectionHeading(doc, "Section 1 - Index form", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  6.2 x 10 squared = _______", y);
    y = addWriteLine(doc, "b)  0.35 x 10 cubed = _______", y);
    y = addWriteLine(doc, "c)  9.04 x 10 to the power 4 = _______", y);

    y = addSectionHeading(doc, "Section 2 - Combined moves", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  (3.6 x 100) x 10 = _______", y);
    y = addWriteLine(doc, "b)  0.5 x 10 squared x 10 = _______", y);
    y = addWriteLine(doc, "c)  Write 28 000 as a decimal multiplied by a power of 10 (two different ways).", y);
    y = addWriteLine(doc, "    _________________________   and   _________________________", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Explain why multiplying by 10 squared then by 10 is the same as multiplying by 10 cubed.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) 620  b) 350  c) 90 400.   S2  a) 3600  b) 500  c) e.g. 2.8 x 10 000 or 28 x 1000.   S3  Adding the indices: 2 places then 1 place is 3 places in total, which is 10 cubed.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Year 8 Extension | Powers of 10`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
