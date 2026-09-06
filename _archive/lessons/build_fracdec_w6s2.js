"use strict";

// Fractions & Decimals - Week 6 (Year 6 Numeracy) - Session 2: Compare & order common fractions on a number line
// VC2M6N03. Compare and order halves, thirds and quarters on one number line and justify the order.
// Daily Review: Fractions & Decimals. Fluency: subtraction algorithm.
// Anchor models: number line + fraction wall. Variant weekToVariant(6).

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
  addFractionStripSet, addNumberLine,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 4;
const WEEK = 6;
const UNIT_TITLE = "Fractions & Decimals";
const FOOTER = `Fractions & Decimals | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracDec_W6S2_Compare_Order";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Comparing and Ordering Fractions Practice",
  "Compare and order fractions on a number line and justify the order. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Comparing and Ordering Fractions Answer Key",
  "Worked answers for the comparing and ordering fractions practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back. Yesterday we connected percentages to fractions and decimals.
- Today we line fractions up and decide which is bigger, using one number line.

DO:
- Have whiteboards, markers and the printed number line ready.
- Settle the class before you click on.

TEACHER NOTES:
Week 6 Session 2 of 4. The number line is the anchor today. A returning student needs the printed number line and the Week 6 Catch-Up Card.

WATCH FOR:
- Students who think a bigger bottom number means a bigger fraction - we will test that today.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- Today's materials are the practice sheet and a printed number line.
- The practice sheet has an easier start and a challenge built in.

DO:
- Print one practice sheet per student and keep the answer key.
- Have the printed number line and whiteboards ready.

TEACHER NOTES:
One student practice sheet plus its answer key. Most teaching is on whiteboards and the number line.

CATCH-UP NOTE:
A returner can join here. The launch re-shows fractions on a 0-to-1 line, and the Week 6 Catch-Up Card summarises comparing and ordering with a worked example.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Fractions and decimals, not today's comparing.
- Answer each on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and check reasoning.

TEACHER NOTES:
Spaced retrieval of ordering decimals and adding fractions, both useful today.

WATCH FOR:
- Students who order decimals by place value - secure.
- Students who order by length of the number - reteach with place value.

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- In order, smallest first: 0.25, 0.3, 0.5.
- One half plus one quarter is three-quarters.
- Yes, three-sixths equals one half.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The ordering item links straight to today: same idea, decimals instead of fractions.

WATCH FOR:
- Students who self-correct - secure.
- Students who order 0.3 before 0.25 - reteach hundredths.

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Decimal subtraction, vertical.
- Line up the decimal points, subtract, regroup if you need to.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up points and regrouping.

TEACHER NOTES:
Decimal subtraction again this week. Keep it brisk.

WATCH FOR:
- Students who regroup correctly - secure.
- Students who flip a column to avoid regrouping - coach borrowing.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 7.3 minus 4.8 is 2.5.
- 10.5 minus 6.25 is 4.25.
- 8 minus 3.6 is 4.4.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
8 minus 3.6 needs 8 read as 8.0 to line up. Coach the placeholder.

WATCH FOR:
- Students who self-correct - secure.
- Students who write 8 - 3.6 = 5.6 - they did not regroup, small group focus.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Here is a number line from zero to one. Where would one half go? Right in the middle.
- Where would one quarter go? Closer to zero. Where would three-quarters go? Closer to one.
- A fraction is a point on this line, a certain distance from zero.
- Today we place fractions on one line and use their positions to put them in order.

DO:
- Point to the middle, then a quarter of the way, then three-quarters.
- Take a prediction for where one third would sit.

TEACHER NOTES:
This launch frames a fraction as a position on a line, which makes comparing about who is further along.

WATCH FOR:
- Students who place one half in the middle - strong start.
- Students who place by the size of the bottom number - we will challenge that today.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to compare and order common fractions on one number line and justify the order.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is placing a fraction. Second is comparing two on the line, the core target. Third is ordering three and explaining.

WATCH FOR:
- Students who can say further along means bigger - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- One useful idea before we start: the one-half benchmark.
- For any fraction, ask: is it more than half or less than half?
- One quarter is less than half. Three-quarters is more than half. This quick check helps us compare.

DO:
- Point to one half in the middle of the line.
- Sort a couple of fractions as more or less than half.

TEACHER NOTES:
This anchor comes after the learning intention. The half benchmark is a fast first comparison before exact work.

WATCH FOR:
- Students who can say more or less than half - secure.
- Students unsure - point to the halfway mark.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us place three quarters-fractions on the line: one quarter, one half, three-quarters.
- The line from zero to one is split into four equal quarters. One quarter is the first mark.
- One half is the same as two-quarters, the middle mark. Three-quarters is the third mark.
- Now I can read their order straight off the line: one quarter, then one half, then three-quarters.

DO:
- Mark each fraction on the quarters line.
- Read the order left to right.

TEACHER NOTES:
Start with fractions that share a denominator so position is obvious. Order is just left to right on the line.

MISCONCEPTIONS:
- Misconception: students think 1/4 is bigger than 1/2 because 4 is bigger than 2.
  Why: they compare the bottom numbers.
  Impact: they order fractions backwards.
  Quick correction: on the line, 1/4 is closer to zero, so it is smaller.

WATCH FOR:
- Students who read order left to right - secure.
- Students who order by bottom number - back to the line.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now two fractions with different bottoms: one third and one half. Which is bigger?
- I cannot compare thirds and halves directly, so I rename both to sixths.
- One third is two-sixths. One half is three-sixths. Now I place both on a sixths line.
- Two-sixths is left of three-sixths, so one third is less than one half.

DO:
- Rename both to sixths.
- Mark 2/6 and 3/6 on the sixths line and read which is further along.

TEACHER NOTES:
The core move: rename to a common denominator, then compare positions. Equivalent fractions do the heavy lifting.

MISCONCEPTIONS:
- Misconception: students say one third is bigger because thirds sounds like more pieces.
  Why: more pieces feels like more.
  Impact: wrong comparison.
  Quick correction: more pieces means each piece is smaller; rename and compare on the line.

WATCH FOR:
- Students who rename then compare - secure.
- Students who guess from the words - rename to sixths.

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Which is larger, two-thirds or three-quarters?
- Rename both to twelfths to be sure.

DO:
- Display the prompt.
- Give about 45 seconds.
- Walk and scan for the twelfths.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me the larger fraction on three, two, one, show.
- Scan for: 3/4, with 2/3 = 8/12 and 3/4 = 9/12.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students pick 2/3 because the numbers are smaller, or guess.
- Reteach: rename both to twelfths. 2/3 = 8/12, 3/4 = 9/12. Nine-twelfths is further along.
- Re-check: what are two-thirds and three-quarters as twelfths?
- Use the number line to confirm.

TEACHER NOTES:
Renaming to a common denominator is the threshold. The twelfths make the comparison certain.

WATCH FOR:
- Students who pick 3/4 with reasoning - secure.
- Students who guess - reteach renaming.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Put one quarter, one third and one half in order, smallest first.
- Use the fraction wall to compare the lengths.

DO:
- Display the three bars on the wall.
- Give about 90 seconds for partners.
- Listen for one quarter is shortest, one half is longest.

TEACHER NOTES:
The wall shows the lengths directly. Then we will justify with equivalence on the line.

WATCH FOR:
- Pairs who order 1/4, 1/3, 1/2 - secure.
- Pairs who order by bottom number - point to the bar lengths.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check on the wall.
- One quarter is the shortest bar, then one third, then one half is the longest.
- So the order smallest first is one quarter, one third, one half.

DO:
- Click to reveal.
- Run a finger along the three lengths shortest to longest.

TEACHER NOTES:
Lengths on the wall match positions on the line. Next we prove it with twelfths.

WATCH FOR:
- Students who self-correct - secure.
- Students still ordering by bottom number - small group before You Do.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, and this time we justify with equivalence on a number line.
- Order two-fifths, one half and three-fifths.
- Rename them all to tenths so they sit on the same line.

DO:
- Display the tenths line.
- Give about 90 seconds.
- Listen for 4/10, 5/10, 6/10.

TEACHER NOTES:
Renaming to tenths puts all three on one line. The order is then read left to right.

WATCH FOR:
- Pairs who rename to tenths - secure.
- Pairs who compare fifths and halves directly - rename to tenths.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Two-fifths is four-tenths, one half is five-tenths, three-fifths is six-tenths.
- On the line: four-tenths, five-tenths, six-tenths. So the order is two-fifths, one half, three-fifths.

DO:
- Click to reveal.
- Point to the three marks in order.

TEACHER NOTES:
This is the full justify-with-equivalence move. Next session we add and subtract decimals.

WATCH FOR:
- Students who justify with tenths - ready for independent work.
- Students unsure - one more ordering before You Do.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- The first section places fractions on a number line.
- The next section compares two fractions, larger or smaller.
- The last section orders three fractions and asks you to explain.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for rename to compare, and further along means bigger.
- Cold call one or two students to justify an order.

TEACHER NOTES:
Different fractions from the We Do, same method: place on the line, or rename to a common denominator, then compare positions.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed number line and the fraction wall, and do Section 1 and the first row of Section 2.
- Extra Notes: Sit with these students and place the first fraction together.
EXTENDING PROMPT:
- Task: Extension - order a mixed set such as 1/2, 2/3, 3/4, 5/6, and place a fraction greater than one, like 5/4, on a 0-to-2 line. Early finishers may start the Year 8 Extension Challenge later this week.
- Extra Notes: Push the justification with a common denominator.

WATCH FOR:
- Students who rename to compare - secure.
- Students who order by bottom number - prompt back to the line.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Put these in order, smallest first: three-quarters, one half, two-thirds.
- Then explain how you decided.

DO:
- Display the prompt.
- Give about 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - ordering and justifying. Look for renaming to twelfths: 1/2 = 6/12, 2/3 = 8/12, 3/4 = 9/12, so 1/2 < 2/3 < 3/4.

WATCH FOR:
- Students who order 1/2, 2/3, 3/4 with reasoning - secure.
- Students who order by bottom number - revisit at the start of Session 3.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: how can renaming to the same bottom number help you compare fractions?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that a fraction is a position on the line, and renaming to a common denominator lets us compare positions exactly.

WATCH FOR:
- Strong thumbs across all three - move at pace next session.
- Sideways or down on the core criterion - quick revision at the start of Session 3.

[Retention and Recall]`;

// --- Helpers -----------------------------------------------------------------

function nlCard(slide, topY, labels, marks, caption, opts) {
  const o = opts || {};
  const cardH = SAFE_BOTTOM - topY;
  addCard(slide, 0.5, topY, 9.0, cardH, { strip: o.strip || C.PRIMARY, fill: C.WHITE });
  const lineY = topY + cardH * 0.46;
  addNumberLine(slide, 1.0, lineY, 8.0, labels, marks || [], { labelFontSize: o.labelFontSize || 14 });
  if (caption) {
    slide.addText(caption, {
      x: 1.0, y: lineY + 0.6, w: 8.0, h: 0.5,
      fontSize: o.captionSize || 15, fontFace: FONT_B, color: o.captionColor || C.PRIMARY,
      bold: o.captionBold !== false, italic: Boolean(o.captionItalic),
      align: "center", margin: 0,
    });
  }
}

function wallPanel(slide, lg, headerText, strips, opts) {
  const o = opts || {};
  const cardH = o.cardH || 2.9;
  addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: o.strip || C.PRIMARY });
  slide.addText(headerText, {
    x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
    fontSize: 15, fontFace: FONT_H, color: o.strip || C.PRIMARY, bold: true,
    align: "center", margin: 0,
  });
  const wallY = lg.panelTopPadded + 0.52;
  const wallH = Math.min(cardH - 0.72, strips.length * 0.42);
  addFractionStripSet(slide, lg.rightX + 0.22, wallY, lg.rightW - 0.44, wallH, strips,
    { labelW: 0.7, labelFontSize: 12 });
}

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Week 6 Session 2: Comparing and ordering fractions",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions & decimals",
      [
        "Order smallest first: 0.3, 0.25, 0.5",
        "1/2 + 1/4 = ____",
        "Is 3/6 = 1/2?",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "0.25, 0.3, 0.5     3/4     yes", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal subtraction
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal subtraction",
      ["7.3 - 4.8", "10.5 - 6.25", "8 - 3.6"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "2.5        4.25        4.4", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - a fraction is a position
  contentSlide(pres, "Launch", C.ACCENT, "Where does each fraction sit?",
    [
      "A fraction is a point on the line.",
      "",
      "1/2 in the middle, 1/4 nearer 0,",
      "3/4 nearer 1.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.3, { strip: C.ACCENT });
      addNumberLine(slide, lg.rightX + 0.25, lg.panelTopPadded + 1.0, lg.rightW - 0.5,
        ["0", "1/4", "1/2", "3/4", "1"], [2], { labelFontSize: 13 });
      slide.addText("1/2 is exactly halfway.", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.55, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to compare and order common fractions on one number line and justify the order.",
    [
      "I can place a fraction on a 0 to 1 number line.",
      "I can compare two fractions using the number line.",
      "I can order three fractions and explain the order.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key idea - the half benchmark
  contentSlide(pres, "Key Idea", C.SECONDARY, "The one-half benchmark",
    [
      "First ask: more than half or less than half?",
      "1/4 is less than half.",
      "3/4 is more than half.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.SECONDARY });
      addNumberLine(slide, lg.rightX + 0.25, lg.panelTopPadded + 1.0, lg.rightW - 0.5,
        ["0", "1/2", "1"], [1], { labelFontSize: 14 });
      slide.addText("Less than half  |  More than half", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.5, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 10: I Do #1 - place quarters on the line (custom)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Place 1/4, 1/2, 3/4 on the line", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["2"] });
    s.addText("The line is split into four equal quarters. Mark each fraction.", {
      x: 0.75, y: CONTENT_TOP + 0.1, w: 8.5, h: 0.62,
      fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    nlCard(s, CONTENT_TOP + 1.0,
      ["0", "1/4", "1/2", "3/4", "1"], [1, 2, 3],
      "Read order left to right: 1/4, then 1/2, then 3/4.",
      { strip: STAGE_COLORS["2"], captionColor: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO1);
  })();

  // Slide 11: I Do #2 - compare 1/3 and 1/2 by renaming (custom)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Compare 1/3 and 1/2: rename to sixths", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["2"] });
    s.addText("Rename both to sixths: 1/3 = 2/6 and 1/2 = 3/6. Now compare positions.", {
      x: 0.75, y: CONTENT_TOP + 0.1, w: 8.5, h: 0.62,
      fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    nlCard(s, CONTENT_TOP + 1.0,
      ["0", "1/6", "2/6", "3/6", "4/6", "5/6", "1"], [2, 3],
      "2/6 (= 1/3) is left of 3/6 (= 1/2), so 1/3 < 1/2.",
      { strip: STAGE_COLORS["2"], captionColor: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO2);
  })();

  // Slides 12-13: CFU + reveal - 2/3 vs 3/4
  withReveal(
    () => cfuSlide(pres, "CFU", "Which is larger, 2/3 or 3/4?", "Show Me Boards",
      "Which is larger, 2/3 or 3/4?\n\nRename both to twelfths to be sure.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "3/4 is larger   (2/3 = 8/12,  3/4 = 9/12)",
        { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - order 1/4, 1/3, 1/2 on the wall
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Order 1/4, 1/3, 1/2 (smallest first)",
      [
        "With your partner.",
        "",
        "1.  Compare the bar lengths.",
        "2.  Which is shortest? Longest?",
        "3.  Write them smallest to largest.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Compare the lengths", [
          { denom: 4, shaded: 1, label: "1/4", color: C.SECONDARY },
          { denom: 3, shaded: 1, label: "1/3", color: C.PRIMARY },
          { denom: 2, shaded: 1, label: "1/2", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "1/4  <  1/3  <  1/2   (shortest to longest)", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - order 2/5, 1/2, 3/5 on tenths line (custom)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Order 2/5, 1/2, 3/5 on one line", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["3"] });
      s.addText("Rename all to tenths so they sit on the same line, then order them.", {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 8.5, h: 0.62,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      nlCard(s, CONTENT_TOP + 1.0,
        ["0", "", "2/10", "", "4/10", "5/10", "6/10", "", "8/10", "", "1"], [4, 5, 6],
        "2/5 = 4/10,  1/2 = 5/10,  3/5 = 6/10.",
        { strip: C.PRIMARY, captionColor: C.MUTED, captionBold: false, captionItalic: true });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Order: 2/5  <  1/2  <  3/5   (4/10, 5/10, 6/10)", {
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
      { text: "place fractions on the number line.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "compare two: which is larger?   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "order three and explain.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Further along the line = bigger. To compare different bottoms, rename to the same bottom first.", {
      x: 0.85, y: panelY + 0.52, w: 8.3, h: 0.5, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

    addNumberLine(s, 1.5, panelY + 1.5, 7.0,
      ["0", "1/4", "1/2", "3/4", "1"], [], { labelFontSize: 14 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Put these in order, smallest first: 3/4, 1/2, 2/3.",
      "Explain how you decided.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how can renaming to the same bottom number help you compare fractions?",
      scItems: [
        "I can place a fraction on a 0 to 1 number line.",
        "I can compare two fractions using the number line.",
        "I can order three fractions and explain the order.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDec_W6S2_Compare_Order.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Compare and order fractions on a number line, and justify the order.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "A fraction is a position on the number line - further along means bigger. To compare fractions with different bottom numbers, rename them to the same bottom (a common denominator), then compare. The one-half benchmark helps: is it more or less than half?",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Compare 2/3 and 3/4. Rename to twelfths: 2/3 = 8/12 and 3/4 = 9/12. 9/12 is further along, so 3/4 > 2/3.",
      y);

    y = addSectionHeading(doc, "Section 1 - Place on the number line (0 to 1)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Mark each fraction with an arrow.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "1/2        1/4        3/4        1/3", y);
    y = addWriteLine(doc, "0 ____________________ 1/2 ____________________ 1", y);

    y = addSectionHeading(doc, "Section 2 - Which is larger? (circle it)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2  or  1/4        b)  2/3  or  3/4        c)  1/3  or  1/2", y);
    y = addWriteLine(doc, "d)  3/5  or  1/2        e)  5/6  or  3/4", y);

    y = addSectionHeading(doc, "Section 3 - Order smallest first, then explain", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/4, 1/2, 1/3  ->  ______________________", y);
    y = addWriteLine(doc, "b)  3/4, 1/2, 2/3  ->  ______________________", y);
    y = addWriteLine(doc, "Explain one of your orders: ___________________________________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order this mixed set: 1/2, 2/3, 3/4, 5/6. Then place 5/4 on a line from 0 to 2.", y);
    y = addWriteLine(doc, "Order: ____________________________________________________", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Comparing and Ordering Fractions | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the comparing and ordering fractions practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Place on the number line", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1/4 nearest 0, then 1/3, then 1/2 in the middle, then 3/4 nearer 1.", y);

    y = addSectionHeading(doc, "Section 2 - Which is larger", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/2      b)  3/4 (8/12 vs 9/12)      c)  1/2      d)  3/5 (6/10 vs 5/10)      e)  5/6 (10/12 vs 9/12)", y);

    y = addSectionHeading(doc, "Section 3 - Order smallest first", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/4, 1/3, 1/2.   b)  1/2, 2/3, 3/4 (twelfths: 6/12, 8/12, 9/12).", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "1/2, 2/3, 3/4, 5/6 in twelfths: 6/12, 8/12, 9/12, 10/12 - already in order. 5/4 = 1 1/4, just past 1 on the 0-to-2 line.", y);

    y = addTipBox(doc,
      "Watch for: students who order by the size of the bottom number; students who compare without renaming; students who forget further along the line means larger.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Week 6 Session 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
