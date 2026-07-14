"use strict";

// Understanding & Operating with Fractions - Week 8 (Year 6 Numeracy) - Session 2:
// Compare and order fractions with RELATED denominators by renaming, on a number line.
// VC2M5N03. Equivalent fractions sit at the same point (1/5 = 2/10). Fraction wall + number line.
// Daily Review: Four Process. Fluency: decimal division (division bracket). Variant weekToVariant(8) = 1.

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

const SESSION = 2;
const TOTAL = 4;
const WEEK = 8;
const UNIT_TITLE = "Understanding & Operating with Fractions";
const FOOTER = `Fractions: Understand & Operate | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracUO_W8S2_Compare_Related";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Comparing and Ordering Fractions Practice",
  "Compare and order fractions with related denominators by renaming. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Comparing and Ordering Fractions Answer Key",
  "Worked answers for the comparing and ordering fractions practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back. Yesterday we compared fractions with the same bottom number and unit fractions.
- Today the bottom numbers are different but related, so we rename one fraction first, then compare.

DO:
- Have whiteboards, markers and the printed fraction wall and number line ready.
- Settle the class before you click on.

TEACHER NOTES:
Week 8 Session 2 of 4. The number line is the anchor today, with the wall for renaming. A returning student needs the printed models and the Week 8 Catch-Up Card from Session 1.

WATCH FOR:
- Students who still order by the size of the bottom number. We will challenge that today.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- Today's materials are the comparing and ordering practice sheet and the printed number line.
- The practice sheet has an easier start and a challenge built in.

DO:
- Print one practice sheet per student and keep the answer key.
- Have the printed number line, the fraction wall and whiteboards ready.

TEACHER NOTES:
One student practice sheet plus its answer key. Most teaching is on whiteboards with the number line.

CATCH-UP NOTE:
A returner can join here. The launch re-shows fractions as positions on a 0-to-1 line, and the Week 8 Catch-Up Card from Session 1 summarises comparing by renaming with a worked example.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. The four operations, not today's fractions.
- Answer each one on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and check methods.

TEACHER NOTES:
Spaced retrieval of the four operations. Watch the subtraction for regrouping.

WATCH FOR:
- Students who regroup the subtraction correctly - secure.
- Students who divide 132 by 4 cleanly - secure.

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 600 take away 247 is 353.
- 9 times 6 is 54.
- 132 divided by 4 is 33.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
600 minus 247 needs regrouping across zeros. That is the common slip to scan for.

WATCH FOR:
- Students who self-correct - secure.
- Students stuck on regrouping zeros - small group with a quick model.

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Dividing decimals, set out in the division bracket.
- Dividend inside the bracket, decimal point lined up straight above, then divide.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for the decimal point carried straight up.

TEACHER NOTES:
Decimal division in the division bracket again. Call it the division bracket, not the bus stop.

WATCH FOR:
- Students who keep the point lined up - secure.
- Students who drop the point - coach lining it up first.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 6.4 divided by 4 is 1.6.
- 8.1 divided by 3 is 2.7.
- 5.5 divided by 5 is 1.1.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
8.1 divided by 3 needs a careful carry: 8 divided by 3 is 2 remainder 2, then 21 tenths divided by 3 is 7 tenths.

WATCH FOR:
- Students who self-correct - secure.
- Students who forget to carry the remainder - small group focus.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Here is a number line from zero to one. Where does one half go? Right in the middle.
- Where does one quarter go? Closer to zero. A fraction is a point a certain distance along the line.
- Some fractions land on the same point. One half and two-quarters sit in exactly the same place.
- Today we use renaming to put fractions with different bottoms onto the same line and compare them.

DO:
- Point to the middle, then a quarter of the way along.
- Show one half and two-quarters landing on the same point.

TEACHER NOTES:
This launch frames a fraction as a position and previews equivalence: equal fractions share a point. It re-establishes the number line anchor.

WATCH FOR:
- Students who place one half in the middle - strong start.
- Students who place by the size of the bottom number - we will challenge that today.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to compare and order fractions with related denominators by renaming them.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is renaming a fraction, achievable for everyone. Second is comparing two by renaming, the core target. Third is ordering three and explaining.

WATCH FOR:
- Students who can say rename to the same bottom, then compare - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- One key idea for today: related denominators.
- Related denominators means one bottom number is a multiple of the other, like 2 and 4, or 3 and 6, or 5 and 10.
- When bottoms are related, we can rename one fraction so both have the same bottom, then compare.

DO:
- Show 2 and 4, then 5 and 10, as related pairs.
- Have students say one bottom is a multiple of the other.

TEACHER NOTES:
Vocabulary after the learning intention. Related denominators are the easy case because one bottom already divides into the other.

WATCH FOR:
- Students who spot 4 is double 2, and 10 is double 5 - secure.
- Students unsure - point to the wall rows.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us compare one half and three-eighths. The bottoms are 2 and 8, and 8 is a multiple of 2, so they are related.
- I cannot compare halves and eighths directly, so I rename one half into eighths.
- One half is the same as four-eighths. Now both are eighths.
- On the line, four-eighths is further along than three-eighths, so one half is bigger than three-eighths.

DO:
- Rename 1/2 to 4/8.
- Mark 3/8 and 4/8 on the eighths line and read which is further along.

TEACHER NOTES:
The core move: rename to a common denominator, then compare positions. Equivalent fractions do the work.

MISCONCEPTIONS:
- Misconception: students say three-eighths is bigger because eighths sounds like more pieces.
  Why: more pieces feels like more.
  Impact: a wrong comparison.
  Quick correction: rename one half to four-eighths, then compare the eighths.

WATCH FOR:
- Students who rename then compare - secure.
- Students who guess from the words - rename to eighths.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now two-thirds and five-sixths. The bottoms are 3 and 6, and 6 is double 3, so they are related.
- I rename two-thirds into sixths. Two-thirds is the same as four-sixths.
- Now I compare four-sixths and five-sixths. Five-sixths has more pieces.
- So two-thirds is less than five-sixths.

DO:
- Show 2/3 on the wall, then the equal 4/6 underneath.
- Compare 4/6 and 5/6 and read which is bigger.

TEACHER NOTES:
Thirds and sixths is the other common related pair. Rename the third into sixths, then compare the tops.

MISCONCEPTIONS:
- Misconception: students compare 2/3 and 5/6 by the tops, 2 versus 5, without renaming.
  Why: they compare numerators while the bottoms differ.
  Impact: it works here by luck but fails in general.
  Quick correction: always make the bottoms match first, then compare the tops.

WATCH FOR:
- Students who rename to sixths then compare - secure.
- Students who compare without renaming - back to the wall.

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Which is larger, three-quarters or seven-eighths?
- The bottoms 4 and 8 are related. Rename three-quarters into eighths first.

DO:
- Display the prompt.
- Give about 45 seconds.
- Walk and scan for the eighths.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me the larger fraction on three, two, one, show.
- Scan for: 7/8, with 3/4 renamed to 6/8.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students pick 3/4 because the numbers look smaller, or guess.
- Reteach: rename 3/4 to 6/8. Six-eighths is less than seven-eighths.
- Re-check: what is three-quarters renamed into eighths?
- Use the number line to confirm.

TEACHER NOTES:
Renaming to a common denominator is the threshold. Eighths make the comparison certain.

WATCH FOR:
- Students who pick 7/8 with reasoning - secure.
- Students who guess - reteach renaming.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Put one half, three-quarters and five-eighths in order, smallest first.
- All three bottoms are related to eighths. Rename them all to eighths.

DO:
- Display the three bars on the wall.
- Give about 90 seconds for partners.
- Listen for 4/8, 6/8 and 5/8.

TEACHER NOTES:
Renaming all three to eighths puts them on common ground. Then order the tops 4, 5, 6.

WATCH FOR:
- Pairs who rename to eighths - secure.
- Pairs who compare without renaming - point to the wall rows.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- One half is four-eighths, five-eighths stays, three-quarters is six-eighths.
- In order: four-eighths, five-eighths, six-eighths. So one half, five-eighths, three-quarters.

DO:
- Click to reveal.
- Run a finger along the eighths smallest to largest.

TEACHER NOTES:
Once all three are eighths, ordering is just ordering 4, 5, 6. Next we use the number line to justify.

WATCH FOR:
- Students who self-correct - secure.
- Students still ordering by the bottom number - small group before You Do.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, and this time we justify on a number line.
- Order one fifth, three-tenths and one half, smallest first.
- Rename them all to tenths so they sit on the same line.

DO:
- Display the tenths line.
- Give about 90 seconds.
- Listen for 2/10, 3/10 and 5/10.

TEACHER NOTES:
Fifths and halves both rename to tenths. One fifth lands on the same point as two-tenths - the equivalence is visible on the line.

WATCH FOR:
- Pairs who rename to tenths - secure.
- Pairs who compare fifths and tenths directly - rename to tenths.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- One fifth is two-tenths, three-tenths stays, one half is five-tenths.
- On the line: two-tenths, three-tenths, five-tenths. So one fifth, three-tenths, one half.

DO:
- Click to reveal.
- Point to one fifth and two-tenths landing on the same point.

TEACHER NOTES:
This is the full justify-with-equivalence move, and it shows equal fractions sharing a point. Next session we add and subtract fractions.

WATCH FOR:
- Students who justify with tenths - ready for independent work.
- Students unsure - one more ordering before You Do.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- The first section renames a fraction to an equal fraction.
- The next section compares two fractions by renaming.
- The last section orders three fractions and asks you to explain.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for rename to the same bottom, then compare.
- Cold call one or two students to justify an order.

TEACHER NOTES:
Different fractions from the We Do, same method: rename to a common denominator, then compare positions.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed number line and fraction wall, and do Section 1 and the first row of Section 2. The first one is done for you.
- Extra Notes: Sit with these students and rename the first fraction together.
EXTENDING PROMPT:
- Task: Extension - order a mixed set such as 1/2, 5/8, 2/3, 3/4 by renaming, and place a fraction greater than one like 5/4 on a 0-to-2 line.
- Extra Notes: Push the justification with a common denominator.

WATCH FOR:
- Students who rename to compare - secure.
- Students who order by the bottom number - prompt back to the line.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Which is larger, five-sixths or two-thirds? Rename to compare.
- Then put these in order, smallest first: one half, three-eighths, three-quarters.

DO:
- Display the prompt.
- Give about 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - comparing and ordering by renaming. 2/3 = 4/6, so 5/6 is larger. In eighths: 4/8, 3/8, 6/8, so 3/8, 1/2, 3/4.

WATCH FOR:
- Students who rename then order - secure.
- Students who order by the bottom number - revisit at the start of Session 3.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: how does renaming to the same bottom number help you compare fractions?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that renaming to a common denominator puts fractions on the same line so we can compare positions exactly.

WATCH FOR:
- Strong thumbs across all three - move at pace next session.
- Sideways or down on the core criterion - quick revision at the start of Session 3.

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
      bold: o.captionBold !== false, italic: Boolean(o.captionItalic), align: "center", margin: 0,
    });
  }
}

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Week 8 Session 2: Comparing related-denominator fractions",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Four Process
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: The four operations",
      [
        "Calculate: 600 - 247",
        "Calculate: 9 × 6",
        "Calculate: 132 ÷ 4",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "353      54      33", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal division (division bracket)
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal division (division bracket)",
      ["6.4 ÷ 4", "8.1 ÷ 3", "5.5 ÷ 5"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "1.6        2.7        1.1", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - a fraction is a position; equal fractions share a point
  contentSlide(pres, "Launch", C.ACCENT, "Some fractions share a point",
    [
      "A fraction is a point on the line.",
      "",
      "1/2 and 2/4 sit in the same place.",
      "Renaming moves us onto the same line.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.3, { strip: C.ACCENT });
      addNumberLine(slide, lg.rightX + 0.25, lg.panelTopPadded + 1.0, lg.rightW - 0.5,
        ["0", "1/4", "2/4", "3/4", "1"], [2], { labelFontSize: 13 });
      slide.addText("1/2 lands on 2/4.", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.55, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to compare and order fractions with related denominators by renaming them.",
    [
      "I can rename a fraction to an equal fraction (e.g. 1/2 = 2/4).",
      "I can compare two fractions with related denominators by renaming.",
      "I can order three fractions and explain using equivalent fractions.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Related denominators",
    [
      "Related denominators: one bottom is a multiple of the other (2 and 4, 3 and 6, 5 and 10).",
      "Rename one fraction so both bottoms match, then compare.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Rename to the same bottom", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 2, label: "2/4", color: C.SECONDARY },
        { denom: 8, shaded: 4, label: "4/8", color: C.ACCENT },
      ], { strip: C.SECONDARY, cardH: 2.3 });
    }
  );

  // Slide 10: I Do #1 - compare 1/2 and 3/8 by renaming (custom number line)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Compare 1/2 and 3/8: rename to eighths", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["2"] });
    s.addText("8 is a multiple of 2, so rename: 1/2 = 4/8. Now compare on the eighths line.", {
      x: 0.75, y: CONTENT_TOP + 0.1, w: 8.5, h: 0.62,
      fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    nlCard(s, CONTENT_TOP + 1.0,
      ["0", "1/8", "2/8", "3/8", "4/8", "5/8", "6/8", "7/8", "1"], [3, 4],
      "4/8 (= 1/2) is further along than 3/8, so 1/2 > 3/8.",
      { strip: STAGE_COLORS["2"], captionColor: C.PRIMARY, labelFontSize: 13 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO1);
  })();

  // Slide 11: I Do #2 - compare 2/3 and 5/6 by renaming (wall)
  workedExSlide(pres, 2, "I Do", "Compare 2/3 and 5/6: rename to sixths",
    [
      "6 is double 3, so they are related.",
      "",
      "Rename: 2/3 = 4/6.",
      "Now compare 4/6 and 5/6.",
      "5/6 has more pieces.",
      "",
      "So 2/3 < 5/6.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Rename 2/3 to 4/6, then compare", [
        { denom: 3, shaded: 2, label: "2/3", color: C.PRIMARY },
        { denom: 6, shaded: 4, label: "4/6", color: C.SECONDARY },
        { denom: 6, shaded: 5, label: "5/6", color: C.ACCENT },
      ], { strip: C.PRIMARY, cardH: 2.6 });
    }
  );

  // Slides 12-13: CFU + reveal - 3/4 vs 7/8
  withReveal(
    () => cfuSlide(pres, "CFU", "Which is larger, 3/4 or 7/8?", "Show Me Boards",
      "Which is larger, 3/4 or 7/8?\n\nRename 3/4 into eighths first.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "7/8 is larger   (3/4 = 6/8,  and 6/8 < 7/8)",
        { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - order 1/2, 3/4, 5/8 (rename to eighths)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Order 1/2, 3/4, 5/8 (smallest first)",
      [
        "With your partner.",
        "",
        "1.  Rename them all to eighths.",
        "2.  Order the tops.",
        "3.  Write them smallest to largest.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Rename all to eighths", [
          { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
          { denom: 8, shaded: 5, label: "5/8", color: C.SECONDARY },
          { denom: 4, shaded: 3, label: "3/4", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "1/2  <  5/8  <  3/4   (4/8, 5/8, 6/8)", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - order 1/5, 3/10, 1/2 on tenths line (custom)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Order 1/5, 3/10, 1/2 on one line", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["3"] });
      s.addText("Rename all to tenths so they sit on the same line, then order them.", {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 8.5, h: 0.62,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      nlCard(s, CONTENT_TOP + 1.0,
        ["0", "", "2/10", "3/10", "", "5/10", "", "", "", "", "1"], [2, 3, 5],
        "1/5 = 2/10,  then 3/10,  then 1/2 = 5/10.",
        { strip: C.PRIMARY, captionColor: C.MUTED, captionBold: false, captionItalic: true });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Order: 1/5  <  3/10  <  1/2   (2/10, 3/10, 5/10)", {
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
      { text: "rename to an equal fraction.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "compare two by renaming.   ", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Related bottoms? Rename one fraction so the bottoms match, then compare the tops.", {
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
      "Which is larger, 5/6 or 2/3? Rename to compare.",
      "Put in order, smallest first: 1/2, 3/8, 3/4.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how does renaming to the same bottom number help you compare fractions?",
      scItems: [
        "I can rename a fraction to an equal fraction (e.g. 1/2 = 2/4).",
        "I can compare two fractions with related denominators by renaming.",
        "I can order three fractions and explain using equivalent fractions.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracUO_W8S2_Compare_Related.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Compare and order fractions with related denominators by renaming.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Related denominators means one bottom number is a multiple of the other (like 2 and 4, or 5 and 10). Rename one fraction so both bottoms match, then compare the tops. Equal fractions sit at the same point on the number line.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Compare 2/3 and 5/6. 6 is double 3, so rename 2/3 = 4/6. Now compare 4/6 and 5/6: 5/6 has more pieces, so 5/6 > 2/3.",
      y);

    y = addSectionHeading(doc, "Section 1 - Rename to an equal fraction", y, { color: C.PRIMARY });
    y = addBodyText(doc, "The first one is done for you.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  1/2 = 4/8        b)  1/2 = __/6        c)  2/3 = __/6        d)  3/5 = __/10", y);
    y = addWriteLine(doc, "e)  3/4 = __/8        f)  1/5 = __/10", y);

    y = addSectionHeading(doc, "Section 2 - Which is larger? Rename to compare (circle it)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2  or  3/8        b)  3/4  or  7/8        c)  2/3  or  5/6", y);
    y = addWriteLine(doc, "d)  2/5  or  1/2        e)  5/6  or  2/3", y);

    y = addSectionHeading(doc, "Section 3 - Order smallest first, then explain", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2, 3/4, 5/8  ->  ______________________", y);
    y = addWriteLine(doc, "b)  1/5, 1/2, 3/10  ->  ______________________", y);
    y = addWriteLine(doc, "Explain one of your orders: ___________________________________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order this set by renaming: 1/2, 5/8, 2/3, 3/4. Then place 5/4 on a line from 0 to 2.", y);
    y = addWriteLine(doc, "Order: ____________________________________________________", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Comparing and Ordering Fractions | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the comparing and ordering fractions practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Rename to an equal fraction", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4/8      b)  3/6      c)  4/6      d)  6/10      e)  6/8      f)  2/10", y);

    y = addSectionHeading(doc, "Section 2 - Which is larger", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/2 (= 4/8 > 3/8)      b)  7/8 (3/4 = 6/8)      c)  5/6 (2/3 = 4/6)      d)  1/2 (= 5/10 > 4/10)      e)  5/6 (2/3 = 4/6)", y);

    y = addSectionHeading(doc, "Section 3 - Order smallest first", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/2, 5/8, 3/4 (eighths: 4/8, 5/8, 6/8).   b)  1/5, 3/10, 1/2 (tenths: 2/10, 3/10, 5/10).", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order: 1/2, 5/8, 2/3, 3/4. Using 24ths: 1/2 = 12/24, 5/8 = 15/24, 2/3 = 16/24, 3/4 = 18/24. 5/4 = 1 1/4, just past 1 on the 0-to-2 line.", y);

    y = addTipBox(doc,
      "Watch for: students who order by the size of the bottom number; students who compare without renaming; students who forget that equal fractions sit at the same point.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Week 8 Session 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
