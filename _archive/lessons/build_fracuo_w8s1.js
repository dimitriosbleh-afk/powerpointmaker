"use strict";

// Understanding & Operating with Fractions - Week 8 (Year 6 Numeracy) - Session 1:
// Compare and order fractions with the SAME denominator, and compare UNIT fractions.
// VC2M5N03. Fraction wall + number line anchors. No renaming yet (that is Session 2).
// Daily Review: Four Process (the four operations). Fluency: decimal division (division bracket).
// Variant weekToVariant(8) = 1.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(8));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, dailyReviewSlide, fluencySlide,
  addStageBadge, addRevealAnswerBar,
  addFractionStripSet, addNumberLine,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 4;
const WEEK = 8;
const UNIT_TITLE = "Understanding & Operating with Fractions";
const FOOTER = `Fractions: Understand & Operate | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracUO_W8S1_Compare_Same_Unit";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Comparing Fractions Practice",
  "Compare and order fractions with the same bottom number, and compare unit fractions. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Comparing Fractions Answer Key",
  "Worked answers for the comparing fractions practice sheet.");
const CATCHUP_RES = makeSessionResource(SESSION,
  "Week 8 Catch Up Card",
  "One-page reference for the whole week: comparing, ordering, adding and subtracting fractions with worked examples. For any student returning after an absence.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, CATCHUP_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to our fractions week. This week we compare fractions, order them, and add and subtract them.
- Today we start by deciding which fraction is bigger, using a fraction wall and a number line.

DO:
- Have whiteboards, markers and the printed fraction wall ready.
- Settle the class before you click on.

TEACHER NOTES:
Week 8 Session 1 of 4. The fraction wall and the number line are the anchors for the whole week. Each session re-teaches its model so a returning student can rejoin quickly.

WATCH FOR:
- Students who think a bigger bottom number always means a bigger fraction. We will test that today.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- Today's materials are the comparing fractions practice sheet and the printed fraction wall.
- The practice sheet has an easier start and a challenge built in.

DO:
- Print one practice sheet per student and keep the answer key.
- Have the printed fraction wall and whiteboards ready.

TEACHER NOTES:
One student practice sheet plus its answer key. The Week 8 Catch-Up Card is the one-page summary for the whole week.

CATCH-UP NOTE:
The Week 8 Catch-Up Card included here covers comparing, ordering, adding and subtracting fractions with worked examples. Hand it to any student who has missed a session so they can rejoin at the current lesson. Each session also re-shows its model in the launch.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. The four operations, not today's fractions.
- Answer each one on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and check methods.

TEACHER NOTES:
Spaced retrieval of the four operations. Keep it brisk and low stakes.

WATCH FOR:
- Students who line up place value for the addition - secure.
- Students who muddle a times fact - flag for a quick fluency follow-up.

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 425 plus 168 is 593.
- 7 times 8 is 56.
- 84 divided by 6 is 14.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The division item links to today's fluency in a small way: equal groups. The fractions work is separate.

WATCH FOR:
- Students who self-correct - secure.
- Students stuck on 84 divided by 6 - prompt: how many sixes in 84?

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Dividing decimals, set out in the division bracket.
- Write the dividend inside the bracket, keep the decimal point lined up above, and divide as usual.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for the decimal point carried straight up.

TEACHER NOTES:
Decimal division in the division bracket this week. Call it the division bracket, not the bus stop. The key habit is keeping the point lined up.

WATCH FOR:
- Students who place the point straight up - secure.
- Students who drop the point - coach lining it up before dividing.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 4.8 divided by 4 is 1.2.
- 7.5 divided by 3 is 2.5.
- 9.6 divided by 8 is 1.2.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
In the division bracket the decimal point in the answer sits directly above the point in the dividend. That is the move to reinforce.

WATCH FOR:
- Students who self-correct - secure.
- Students who get whole-number answers - they ignored the point, small group focus.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Look at the wall. Here is one half, and here is one quarter. Which piece is bigger?
- The half is the bigger piece, even though four is a bigger number than two.
- Now here is a number line from zero to one. One half sits in the middle, one quarter sits closer to zero.
- Today we use the wall and the line to decide which fraction is bigger, and to put fractions in order.

DO:
- Point to the half bar and the quarter bar so the size difference is obvious.
- Run a finger along the line from zero to one and stop at one half.

TEACHER NOTES:
This launch sets up the big idea: the bottom number tells you the size of each piece, and more pieces means smaller pieces. It re-establishes both anchor models for any returning student.

WATCH FOR:
- Students who say the half is bigger - strong start.
- Students who pick the quarter because four is bigger - we will fix this today.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to compare and order fractions using the fraction wall and the number line.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is comparing with the same bottom number, achievable for everyone. Second is comparing unit fractions, the core target. Third is ordering a set and explaining.

WATCH FOR:
- Students who can repeat more pieces means smaller pieces - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- Two words we will use all week. The bottom number is the denominator. It tells us the size of each piece.
- The top number is the numerator. It counts how many pieces we have.
- A unit fraction has a top number of one, like one half or one fifth.

DO:
- Point to the top and bottom of a fraction as you name them.
- Have students say denominator and numerator once each.

TEACHER NOTES:
Vocabulary comes after the learning intention. Keep it light: the denominator names the size, the numerator counts the pieces.

WATCH FOR:
- Students who can point to the denominator - secure.
- Students unsure - bottom names the size, top counts.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us compare two fractions with the same bottom number: three-eighths and five-eighths.
- Both are eighths, so every piece is the same size. I only need to count the pieces.
- Five-eighths has more pieces than three-eighths, so five-eighths is bigger.
- On the line, five-eighths is further along than three-eighths. Same bottom number, just compare the tops.

DO:
- Point to 3/8 and 5/8 on the eighths line.
- Say the rule: same bottom number, compare the tops.

TEACHER NOTES:
Start with the easy case. When denominators match, comparing is just comparing the numerators.

MISCONCEPTIONS:
- Misconception: students look at the bottom number to decide size when the bottoms are equal.
  Why: they always look at the bottom.
  Impact: confusion when bottoms are the same.
  Quick correction: same bottom means same-size pieces, so count the pieces on top.

WATCH FOR:
- Students who compare the tops - secure.
- Students unsure - point to the matching pieces on the line.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now two unit fractions: one third and one sixth. Both have a top number of one.
- Watch the wall. One third is one piece out of three, so it is a big piece.
- One sixth is one piece out of six, so it is a smaller piece.
- More pieces in the whole means each piece is smaller. So one third is bigger than one sixth.

DO:
- Point to the one-third bar, then the one-sixth bar, on the wall.
- Say it together: more pieces means smaller pieces.

TEACHER NOTES:
This is the core idea for unit fractions and the most common trap: a bigger denominator gives a smaller piece.

MISCONCEPTIONS:
- Misconception: students say one sixth is bigger because six is bigger than three.
  Why: they compare the bottom numbers as if bigger means more.
  Impact: they order unit fractions backwards.
  Quick correction: cut the same whole into more pieces and the pieces get smaller.

WATCH FOR:
- Students who say one third is bigger with reasoning - secure.
- Students who pick one sixth - back to the wall lengths.

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Which is larger, one fifth or one eighth?
- Both are unit fractions, so think about the size of one piece.

DO:
- Display the prompt.
- Give about 45 seconds.
- Walk and scan the answers.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me the larger fraction on three, two, one, show.
- Scan for: 1/5, because fifths are bigger pieces than eighths.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students pick 1/8 because 8 is bigger than 5.
- Reteach: on the wall, cut one whole into 5 then into 8. The eighths are smaller pieces.
- Re-check: which is the bigger piece, one fifth or one eighth?
- Use the fraction wall to confirm.

TEACHER NOTES:
Unit fractions are the threshold for today. The bigger the bottom number, the smaller the piece.

WATCH FOR:
- Students who pick 1/5 with reasoning - secure.
- Students who pick 1/8 - reteach with the wall.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Put three-sixths, one-sixth and four-sixths in order, smallest first.
- They all have the same bottom number, so just compare the tops.

DO:
- Display the three bars on the wall.
- Give about 90 seconds for partners.
- Listen for one-sixth is smallest, four-sixths is largest.

TEACHER NOTES:
Same-denominator ordering. Position on the wall matches the size of the numerator.

WATCH FOR:
- Pairs who order 1/6, 3/6, 4/6 - secure.
- Pairs who muddle the order - point to the shaded lengths.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check on the wall.
- One-sixth is the shortest, then three-sixths, then four-sixths is the longest.
- So smallest first: one-sixth, three-sixths, four-sixths.

DO:
- Click to reveal.
- Run a finger along the three lengths shortest to longest.

TEACHER NOTES:
Same bottom number, so ordering is just ordering the numerators 1, 3, 4.

WATCH FOR:
- Students who self-correct - secure.
- Students still unsure - one more same-denominator set.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, this time unit fractions. Put one half, one quarter and one eighth in order, smallest first.
- Each has a top number of one, so think about the size of one piece.

DO:
- Display the three bars on the wall.
- Give about 90 seconds.
- Listen for one eighth is smallest, one half is largest.

TEACHER NOTES:
Unit fractions order the opposite way to the bottom numbers: the bigger the bottom, the smaller the fraction.

WATCH FOR:
- Pairs who order 1/8, 1/4, 1/2 - secure.
- Pairs who order by the bottom number - point to the bar lengths.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours on the wall.
- One eighth is the smallest piece, then one quarter, then one half is the biggest.
- So smallest first: one eighth, one quarter, one half.

DO:
- Click to reveal.
- Point out that the bottom numbers go 8, 4, 2 while the fractions go smallest to largest.

TEACHER NOTES:
A good chance to name the pattern: for unit fractions, a bigger bottom number means a smaller fraction. Next session we compare fractions with related denominators.

WATCH FOR:
- Students who explain the pattern - ready for independent work.
- Students unsure - one more ordering before You Do.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- The first section compares fractions with the same bottom number.
- The next section compares unit fractions.
- The last section orders a set and asks you to explain.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for same bottom means compare the tops, and more pieces means smaller pieces.
- Cold call one or two students to explain a comparison.

TEACHER NOTES:
Different fractions from the We Do, same ideas: compare same-denominator fractions by the tops, and unit fractions by the size of one piece.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed fraction wall and do Section 1 and the first row of Section 2. The first one is done for you.
- Extra Notes: Sit with these students and compare the first pair on the wall together.
EXTENDING PROMPT:
- Task: Extension - order a mixed set of unit fractions and explain the pattern between the bottom number and the size. Compare a same-denominator pair to a unit-fraction pair and explain the difference.
- Extra Notes: Push the language: the denominator names the size of each piece.

WATCH FOR:
- Students who compare with reasoning - secure.
- Students who order unit fractions by the bottom number - prompt back to the wall.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Which is larger, three-fifths or four-fifths? How do you know?
- Then, which is larger, one quarter or one seventh? How do you know?

DO:
- Display the prompt.
- Give about 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - comparing same-denominator fractions and unit fractions with reasoning. 4/5 > 3/5 (same bottom, more pieces). 1/4 > 1/7 (fewer pieces means a bigger piece).

WATCH FOR:
- Students who answer 4/5 and 1/4 with reasoning - secure.
- Students who pick 1/7 - revisit unit fractions at the start of Session 2.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: why does a bigger bottom number give a smaller unit fraction?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that the denominator names the size of each piece, so more pieces means each piece is smaller.

WATCH FOR:
- Strong thumbs across all three - move at pace next session.
- Sideways or down on the core criterion - quick revision at the start of Session 2.

[Retention and Recall]`;

// --- Helpers -----------------------------------------------------------------

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
      bold: o.captionBold !== false, align: "center", margin: 0,
    });
  }
}

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Week 8 Session 1: Comparing and ordering fractions",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Four Process
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: The four operations",
      [
        "Calculate: 425 + 168",
        "Calculate: 7 × 8",
        "Calculate: 84 ÷ 6",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "593      56      14", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal division (division bracket)
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal division (division bracket)",
      ["4.8 ÷ 4", "7.5 ÷ 3", "9.6 ÷ 8"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "1.2        2.5        1.2", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - which piece is bigger
  contentSlide(pres, "Launch", C.ACCENT, "Which piece is bigger?",
    [
      "Here is 1/2 and here is 1/4.",
      "",
      "The half is the bigger piece,",
      "even though 4 is more than 2.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Bigger bottom = smaller piece", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 1, label: "1/4", color: C.ACCENT },
      ], { strip: C.ACCENT, cardH: 2.1 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to compare and order fractions using the fraction wall and the number line.",
    [
      "I can compare two fractions that have the same bottom number.",
      "I can compare unit fractions (a top number of 1).",
      "I can put a set of fractions in order and explain why.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Numerator, denominator, unit fraction",
    [
      "Denominator: the bottom number. It names the size of each piece.",
      "Numerator: the top number. It counts the pieces.",
      "Unit fraction: a top number of 1 (like 1/2 or 1/5).",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "One whole, cut into pieces", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 3, shaded: 1, label: "1/3", color: C.SECONDARY },
        { denom: 4, shaded: 1, label: "1/4", color: C.ACCENT },
      ], { strip: C.SECONDARY, cardH: 2.3 });
    }
  );

  // Slide 10: I Do #1 - same denominator on the number line (custom)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Same bottom: compare the tops", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["2"] });
    s.addText("3/8 and 5/8 are both eighths, so just count the pieces on top.", {
      x: 0.75, y: CONTENT_TOP + 0.1, w: 8.5, h: 0.62,
      fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    nlCard(s, CONTENT_TOP + 1.0,
      ["0", "1/8", "2/8", "3/8", "4/8", "5/8", "6/8", "7/8", "1"], [3, 5],
      "5/8 is further along than 3/8, so 5/8 > 3/8.",
      { strip: STAGE_COLORS["2"], captionColor: C.PRIMARY, labelFontSize: 13 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO1);
  })();

  // Slide 11: I Do #2 - compare unit fractions on the wall
  workedExSlide(pres, 2, "I Do", "Unit fractions: more pieces, smaller piece",
    [
      "1/3 and 1/6 both have a top of 1.",
      "",
      "1/3 is 1 piece out of 3 - a big piece.",
      "1/6 is 1 piece out of 6 - a smaller piece.",
      "",
      "So 1/3 is bigger than 1/6.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "1/3 is a bigger piece than 1/6", [
        { denom: 3, shaded: 1, label: "1/3", color: C.PRIMARY },
        { denom: 6, shaded: 1, label: "1/6", color: C.ACCENT },
      ], { strip: C.PRIMARY, cardH: 2.0 });
    }
  );

  // Slides 12-13: CFU + reveal - 1/5 vs 1/8
  withReveal(
    () => cfuSlide(pres, "CFU", "Which is larger, 1/5 or 1/8?", "Show Me Boards",
      "Which is larger, 1/5 or 1/8?\n\nBoth are unit fractions. Think about the size of one piece.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "1/5 is larger   (fifths are bigger pieces than eighths)",
        { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - order 1/6, 3/6, 4/6 (same denom)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Order 3/6, 1/6, 4/6 (smallest first)",
      [
        "With your partner.",
        "",
        "1.  Same bottom number - just",
        "    compare the tops.",
        "2.  Write them smallest to largest.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Compare the tops: 1, 3, 4", [
          { denom: 6, shaded: 1, label: "1/6", color: C.SECONDARY },
          { denom: 6, shaded: 3, label: "3/6", color: C.PRIMARY },
          { denom: 6, shaded: 4, label: "4/6", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "1/6  <  3/6  <  4/6", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - order unit fractions 1/2, 1/4, 1/8
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Order 1/2, 1/4, 1/8 (smallest first)",
      [
        "With your partner.",
        "",
        "1.  All are unit fractions.",
        "2.  Which is the smallest piece?",
        "3.  Write them smallest to largest.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "More pieces = smaller piece", [
          { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
          { denom: 4, shaded: 1, label: "1/4", color: C.SECONDARY },
          { denom: 8, shaded: 1, label: "1/8", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "1/8  <  1/4  <  1/2   (bottoms 8, 4, 2)", { color: C.SUCCESS });
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
      { text: "compare same-bottom fractions.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "compare unit fractions.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "order a set and explain.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Same bottom? Compare the tops. Unit fractions? More pieces means a smaller piece.", {
      x: 0.85, y: panelY + 0.52, w: 8.3, h: 0.5, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFractionStripSet(s, 2.5, panelY + 1.2, 5.0, 0.85, [
      { denom: 4, shaded: 1, label: "1/4", color: C.PRIMARY },
      { denom: 8, shaded: 1, label: "1/8", color: C.ACCENT },
    ], { labelW: 0.6, labelFontSize: 12 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Which is larger, 3/5 or 4/5? How do you know?",
      "Which is larger, 1/4 or 1/7? How do you know?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why does a bigger bottom number give a smaller unit fraction?",
      scItems: [
        "I can compare two fractions that have the same bottom number.",
        "I can compare unit fractions (a top number of 1).",
        "I can put a set of fractions in order and explain why.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracUO_W8S1_Compare_Same_Unit.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Compare and order fractions with the fraction wall and number line.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Same bottom number? The pieces are the same size, so compare the top numbers. Unit fractions (a top of 1)? The more pieces the whole is cut into, the smaller each piece, so 1/8 is smaller than 1/4.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Compare 2/7 and 5/7: same bottom number, so compare the tops. 5 pieces is more than 2 pieces, so 5/7 > 2/7.",
      y);

    y = addSectionHeading(doc, "Section 1 - Same bottom number (circle the larger)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "The first one is done for you.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  3/8  or  (5/8)        b)  2/5  or  4/5        c)  7/10  or  3/10", y);
    y = addWriteLine(doc, "d)  4/6  or  5/6          e)  1/4  or  3/4", y);

    y = addSectionHeading(doc, "Section 2 - Unit fractions (circle the larger)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/3  or  1/6          b)  1/5  or  1/8          c)  1/2  or  1/10", y);
    y = addWriteLine(doc, "d)  1/4  or  1/9          e)  1/7  or  1/3", y);

    y = addSectionHeading(doc, "Section 3 - Order smallest first, then explain", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2/6, 5/6, 1/6  ->  ______________________", y);
    y = addWriteLine(doc, "b)  1/8, 1/2, 1/4  ->  ______________________", y);
    y = addWriteLine(doc, "Explain one of your orders: ___________________________________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order these unit fractions smallest first: 1/3, 1/10, 1/2, 1/5. Then explain the pattern between the bottom number and the size of the fraction.", y);
    y = addWriteLine(doc, "Order: ____________________________________________________", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Comparing Fractions | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the comparing fractions practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Same bottom number", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  5/8      b)  4/5      c)  7/10      d)  5/6      e)  3/4", y);

    y = addSectionHeading(doc, "Section 2 - Unit fractions", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/3      b)  1/5      c)  1/2      d)  1/4      e)  1/3   (a smaller bottom number means a bigger piece)", y);

    y = addSectionHeading(doc, "Section 3 - Order smallest first", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/6, 2/6, 5/6 (same bottom, order the tops).   b)  1/8, 1/4, 1/2 (more pieces means a smaller piece).", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "1/10, 1/5, 1/3, 1/2. Pattern: for unit fractions, the bigger the bottom number, the smaller the fraction.", y);

    y = addTipBox(doc,
      "Watch for: students who pick the larger fraction by looking at the bottom number; students who order unit fractions the wrong way; students who forget that a bigger bottom number means a smaller piece.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Week 8 Catch-Up Card (one-page reference for the whole week)
  await (async () => {
    const doc = createPdf({ title: CATCHUP_RES.name });
    let y = addPdfHeader(doc, CATCHUP_RES.name, {
      subtitle: "One page to help you catch up on the whole fractions week.",
      color: C.SECONDARY,
      lessonInfo: `Week ${WEEK} | Year 6 Numeracy | Keep this card with you`,
      showNameDate: false,
    });
    y = addTipBox(doc,
      "Two pictures help with every fraction this week: the fraction wall (rows of equal pieces) and the number line from 0 to 1. The bottom number names the size of each piece; the top number counts the pieces.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Comparing fractions", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Same bottom number: compare the tops. 5/8 > 3/8.", y);
    y = addBodyText(doc, "Unit fractions (top of 1): more pieces means a smaller piece. 1/3 > 1/6.", y);
    y = addBodyText(doc, "Related bottoms: rename to the same bottom, then compare. 1/2 = 4/8, so 1/2 > 3/8.", y);

    y = addSectionHeading(doc, "Equivalent (equal) fractions", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Multiply the top and bottom by the same number to rename a fraction. 1/2 = 2/4 = 3/6 = 4/8. Equal fractions sit at the same point on the number line.", y);

    y = addSectionHeading(doc, "Adding and subtracting fractions", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Same bottom number: add or subtract the tops, keep the bottom. 2/8 + 3/8 = 5/8.", y);
    y = addBodyText(doc, "Related bottoms: rename one fraction so the bottoms match, then add or subtract the tops. 1/2 + 1/4 = 2/4 + 1/4 = 3/4.", y);

    y = addSectionHeading(doc, "A worked example to copy", y, { color: C.SECONDARY });
    y = addBodyText(doc, "5/6 - 1/3:  rename 1/3 = 2/6.  Then 5/6 - 2/6 = 3/6,  which is the same as 1/2.", y);

    addPdfFooter(doc, `Week ${WEEK} | Catch-Up Card | Understanding & Operating with Fractions | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, CATCHUP_RES.fileName));
    console.log("PDF written: " + CATCHUP_RES.fileName);
  })();

  console.log("Week 8 Session 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
