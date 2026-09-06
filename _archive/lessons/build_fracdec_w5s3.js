"use strict";

// Fractions & Decimals - Week 5 (Year 6 Numeracy) - Session 3: Add & subtract related-denominator fractions
// VC2M6N05. Add and subtract fractions using equivalent fractions; related denominators (2&4, 3&6, 4&8...).
// Daily Review: Metric measurement & unit conversion. Fluency: decimal addition algorithm.
// Anchor models: fraction wall + number line. Variant weekToVariant(5).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(5));
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

const SESSION = 3;
const TOTAL = 4;
const WEEK = 5;
const UNIT_TITLE = "Fractions & Decimals";
const FOOTER = `Fractions & Decimals | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracDec_W5S3_Add_Subtract_Related";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Adding and Subtracting Fractions Practice",
  "Add and subtract related-denominator fractions by renaming. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Adding and Subtracting Fractions Answer Key",
  "Worked answers for the adding and subtracting fractions practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back. We can name equivalent fractions and convert between mixed and improper fractions.
- Today we use equivalent fractions for a job: adding and subtracting fractions.

DO:
- Have whiteboards, markers and the printed fraction wall ready.
- Settle the class before you click on.

TEACHER NOTES:
Session 3 of 4. The fraction wall returns as the anchor. A returner needs the printed wall and the Week 5 Catch-Up Card.

WATCH FOR:
- Students who try to add the bottoms - we will see why that does not work.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- Today's materials are the practice sheet and the printed fraction wall.
- The practice sheet has an easier start and a challenge built in.

DO:
- Print one practice sheet per student and keep the answer key.
- Have the printed fraction wall and whiteboards ready.

TEACHER NOTES:
One student practice sheet plus its answer key. Most teaching is on whiteboards with the fraction wall.

CATCH-UP NOTE:
A returner can join here. The launch re-shows why parts must be the same size before adding, and the Week 5 Catch-Up Card summarises the renaming method with a worked example.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Measurement and unit conversion, not today's fractions.
- Answer each on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and check the conversion direction.

TEACHER NOTES:
Spaced retrieval of measurement. Watch the direction of each conversion.

WATCH FOR:
- Students who multiply or divide by the right power of ten - secure.
- Students who go the wrong way - prompt: is the answer a bigger or smaller number?

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 3.2 metres is 320 centimetres.
- 2500 grams is 2.5 kilograms.
- 0.75 litres is 750 millilitres.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Grams to kilograms divides by 1000; metres to centimetres multiplies by 100. The direction is the common slip.

WATCH FOR:
- Students who self-correct - secure.
- Students who muddle the direction - small group with a conversion chart.

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Decimal addition, vertical.
- Line up the decimal points, add, bring the point straight down.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up points.

TEACHER NOTES:
Decimal addition algorithm again. The lining-up habit transfers to renaming fractions to the same size.

WATCH FOR:
- Students who align the points - secure.
- Students who align the final digit - coach the decimal point.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 6.7 plus 2.45 is 9.15.
- 9.8 plus 3.6 is 13.4.
- 0.95 plus 4.2 is 5.15.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
6.7 plus 2.45 needs a placeholder zero in the hundredths of 6.70. Coach filling the empty place.

WATCH FOR:
- Students who self-correct - secure.
- Students who misalign - small group focus.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Look at the wall. Here is one half, and here is one quarter.
- Can we just add them to get two of something? Look closely - the parts are different sizes.
- We cannot add halves and quarters until the pieces are the same size.
- Today we make the pieces the same size first, using equivalent fractions, and then add.

DO:
- Point to the half bar and the quarter bar so the size difference is obvious.
- Ask: what could we change the half into so the pieces match the quarters?

TEACHER NOTES:
This launch creates the need for a common denominator before any procedure. Same-size parts is the whole idea.

WATCH FOR:
- Students who suggest changing the half into quarters - strong start.
- Students who add the bottoms - we will fix this in the I Do.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to add and subtract fractions with related denominators using equivalent fractions.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is same-denominator adding. Second is renaming to a common denominator, the core target. Third is full add and subtract with related denominators.

WATCH FOR:
- Students who can repeat same-size parts - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- One key idea for today.
- Related denominators means one bottom number is a multiple of the other, like 2 and 4, or 3 and 6.
- A common denominator is when both fractions have the same bottom number, the same size parts.

DO:
- Show 2 and 4, then 3 and 6, as related pairs.
- Have students say same-size parts.

TEACHER NOTES:
Vocabulary after the learning intention. Related denominators are the easiest case because one bottom already divides into the other.

WATCH FOR:
- Students who spot 4 is double 2 - secure.
- Students unsure - point to the wall rows.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us add two fractions that already have the same bottom number.
- One quarter plus two-quarters. Watch me jump on the number line.
- I start at one-quarter, then jump two more quarters. I land on three-quarters.
- When the parts are the same size, I just add the tops and keep the bottom. One plus two is three, over four.

DO:
- Jump along the quarters on the number line.
- Say the rule: same bottom, add the tops, keep the bottom.

TEACHER NOTES:
Start with the easy case so the rule is clear before renaming is needed.

MISCONCEPTIONS:
- Misconception: students add the bottoms too, writing 3/8.
  Why: they add every number they see.
  Impact: the answer is far too small.
  Quick correction: the bottom names the size; it does not change when we add same-size parts.

WATCH FOR:
- Students who keep the bottom - secure.
- Students who write 3/8 - re-point to the same-size parts.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a pair with different bottoms, but related. One half plus one quarter.
- The halves and quarters are different sizes, so first I rename the half.
- One half is the same as two-quarters. Now both are quarters.
- Two-quarters plus one-quarter is three-quarters. So one half plus one quarter is three-quarters.

DO:
- Show one half on the wall, then the equal two-quarters underneath.
- Add the quarters together and read the total.

TEACHER NOTES:
This is the core move: rename one fraction so both have the same denominator, then add the tops.

MISCONCEPTIONS:
- Misconception: students add 1/2 + 1/4 as 2/6.
  Why: they add tops and bottoms.
  Impact: a wrong, too-small answer.
  Quick correction: make the same size first; 1/2 becomes 2/4, then add.

WATCH FOR:
- Students who rename then add - secure.
- Students who add bottoms - back to the wall.

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Work out one half plus one sixth.
- First rename one half so both fractions are sixths, then add.

DO:
- Display the prompt.
- Give about 45 seconds.
- Walk and scan for the renamed half.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me your answer on three, two, one, show.
- Scan for: 1/2 renamed to 3/6, then 3/6 + 1/6 = 4/6.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students write 2/8, adding tops and bottoms.
- Reteach: rename 1/2 to 3/6 on the wall first, then add only the tops.
- Re-check: what is one half renamed into sixths?
- Use the wall to confirm.

TEACHER NOTES:
The renaming step is the threshold. Students who skip it add tops and bottoms.

WATCH FOR:
- Students who rename to sixths then add - secure.
- Students who write 2/8 - reteach renaming.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me, this time a subtraction.
- Three-quarters take away one half.
- Rename the half into quarters first, then subtract.

DO:
- Display 3/4 and 1/2 on the wall.
- Give about 90 seconds for partners.
- Listen for one half renamed to two-quarters.

TEACHER NOTES:
Subtraction with the same renaming move. The wall shows the take-away clearly.

WATCH FOR:
- Pairs who rename then subtract - secure.
- Pairs who subtract bottoms - re-point to same-size parts.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- One half is two-quarters. Three-quarters take away two-quarters is one-quarter.
- So three-quarters minus one half is one-quarter.

DO:
- Click to reveal.
- Show the take-away on the wall once more.

TEACHER NOTES:
Same renaming idea as addition. The bottom stays the same once the sizes match.

WATCH FOR:
- Students who self-correct - secure.
- Students still subtracting bottoms - small group before You Do.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together. One third plus one sixth.
- Sixths are related to thirds, because six is double three.
- Rename one third into sixths, then add.

DO:
- Display thirds and sixths on the wall.
- Give about 75 seconds.
- Listen for one third renamed to two-sixths.

TEACHER NOTES:
Thirds and sixths is the other common related pair. Rename the third, then add the tops.

WATCH FOR:
- Students who get three-sixths and simplify to a half - strong.
- Students who add bottoms - reteach with the wall.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- One third is two-sixths. Two-sixths plus one-sixth is three-sixths.
- Three-sixths is the same as one half. So one third plus one sixth is one half.

DO:
- Click to reveal.
- Point out that three-sixths simplifies to one half.

TEACHER NOTES:
A nice chance to notice the answer simplifies. Next session moves to unrelated denominators and the lowest common denominator.

WATCH FOR:
- Students who see three-sixths is a half - ready for independent work.
- Students unsure - one more example before You Do.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- The first section adds fractions that already share a bottom number.
- The next section renames one fraction, then adds.
- The last section renames, then subtracts.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for rename first, then add or subtract the tops.
- Cold call one or two students to explain a renaming.

TEACHER NOTES:
Different fractions from the We Do, same method: rename to a common denominator, then add or subtract the tops.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed fraction wall and do Section 1 and the first row of Section 2. The first one is started for you.
- Extra Notes: Sit with these students and rename the first fraction together.
EXTENDING PROMPT:
- Task: Extension - add three related fractions, and solve a missing-number problem such as 3/4 = 1/2 + ?. Early finishers may start the Year 8 Extension Challenge.
- Extra Notes: Push the language: rename to the same size, then operate on the tops.

WATCH FOR:
- Students who rename then operate - secure.
- Students who add or subtract bottoms - prompt back to same-size parts.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Work out five-sixths take away one third.
- Then work out one half plus three-eighths.

DO:
- Display the prompt.
- Give about 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - rename to a common denominator, then add or subtract. 5/6 - 1/3 uses sixths; 1/2 + 3/8 uses eighths.

WATCH FOR:
- Students who write 3/6 (or 1/2) and 7/8 - secure.
- Students who operate on bottoms - revisit at the start of Session 4.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: why do fractions need the same bottom number before we add them?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is same-size parts: rename one fraction to a common denominator, then add or subtract only the tops.

WATCH FOR:
- Strong thumbs across all three - move at pace next session.
- Sideways or down on the core criterion - quick revision at the start of Session 4.

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

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Week 5 Session 3: Adding and subtracting fractions",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Metric measurement
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Metric measurement",
      [
        "Convert: 3.2 m = ____ cm",
        "Convert: 2500 g = ____ kg",
        "How many mL in 0.75 L?",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "320 cm     2.5 kg     750 mL", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal addition",
      ["6.7 + 2.45", "9.8 + 3.6", "0.95 + 4.2"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "9.15        13.4        5.15", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - half and quarter are different sizes
  contentSlide(pres, "Launch", C.ACCENT, "Can we add these straight away?",
    [
      "Here is 1/2 and here is 1/4.",
      "",
      "The parts are different sizes.",
      "Make them the same size first.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Different sized parts", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 1, label: "1/4", color: C.ACCENT },
      ], { strip: C.ACCENT, cardH: 2.1 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to add and subtract fractions with related denominators using equivalent fractions.",
    [
      "I can add or subtract fractions that have the same denominator.",
      "I can rename one fraction so two fractions have the same denominator.",
      "I can add and subtract two fractions with related denominators.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Related & common denominators",
    [
      "Related denominators: one bottom is a multiple of the other (2 and 4, 3 and 6).",
      "Common denominator: both fractions have the same bottom (same-size parts).",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Make the parts match", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 2, label: "2/4", color: C.SECONDARY },
      ], { strip: C.SECONDARY, cardH: 1.9 });
    }
  );

  // Slide 10: I Do #1 - same denominator on the number line
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Same bottom: just add the tops", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["2"] });
    s.addText("1/4 + 2/4. Start at 1/4, then jump 2 more quarters.", {
      x: 0.75, y: CONTENT_TOP + 0.1, w: 8.5, h: 0.62,
      fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const topY = CONTENT_TOP + 1.0;
    const cardH = SAFE_BOTTOM - topY;
    addCard(s, 0.5, topY, 9.0, cardH, { strip: STAGE_COLORS["2"], fill: C.WHITE });
    const lineY = topY + cardH * 0.46;
    addNumberLine(s, 1.0, lineY, 8.0, ["0", "1/4", "2/4", "3/4", "1"], [1, 3], { labelFontSize: 16 });
    s.addText("Same size parts -> 1/4 + 2/4 = 3/4 (add tops, keep the bottom).", {
      x: 1.0, y: lineY + 0.6, w: 8.0, h: 0.5,
      fontSize: 15, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO1);
  })();

  // Slide 11: I Do #2 - rename then add
  workedExSlide(pres, 2, "I Do", "Different bottoms: rename, then add",
    [
      "1/2 + 1/4. The sizes differ.",
      "",
      "Rename: 1/2 = 2/4.",
      "Now both are quarters.",
      "2/4 + 1/4 = 3/4.",
      "",
      "So 1/2 + 1/4 = 3/4.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Rename 1/2 to 2/4, then add", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 2, label: "2/4", color: C.SECONDARY },
        { denom: 4, shaded: 3, label: "3/4", color: C.ACCENT },
      ], { strip: C.PRIMARY, cardH: 2.6 });
    }
  );

  // Slides 12-13: CFU + reveal - 1/2 + 1/6
  withReveal(
    () => cfuSlide(pres, "CFU", "Add one half and one sixth", "Show Me Boards",
      "Work out 1/2 + 1/6.\n\nRename 1/2 into sixths first, then add.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "1/2 = 3/6,  so  3/6 + 1/6 = 4/6  (= 2/3)",
        { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 3/4 - 1/2
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Subtract: 3/4 - 1/2",
      [
        "With your partner.",
        "",
        "1.  Rename 1/2 into quarters.",
        "2.  Subtract from 3/4.",
        "3.  Read the answer.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Rename 1/2 to 2/4, then take away", [
          { denom: 4, shaded: 3, label: "3/4", color: C.SECONDARY },
          { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
          { denom: 4, shaded: 1, label: "1/4", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "3/4 - 1/2 = 3/4 - 2/4 = 1/4", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - 1/3 + 1/6
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Add: 1/3 + 1/6",
      [
        "With your partner.",
        "",
        "1.  Six is double three.",
        "2.  Rename 1/3 into sixths.",
        "3.  Add, then check if it simplifies.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Rename 1/3 to 2/6, then add", [
          { denom: 3, shaded: 1, label: "1/3", color: C.SECONDARY },
          { denom: 6, shaded: 2, label: "2/6", color: C.PRIMARY },
          { denom: 6, shaded: 3, label: "3/6", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "1/3 + 1/6 = 2/6 + 1/6 = 3/6 = 1/2", { color: C.SUCCESS });
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
      { text: "add fractions with the same bottom.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "rename one fraction, then add.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "rename, then subtract.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Make the parts the same size first. Then add or subtract the tops and keep the bottom.", {
      x: 0.9, y: panelY + 0.52, w: 8.2, h: 0.5, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFractionStripSet(s, 2.5, panelY + 1.2, 5.0, 0.85, [
      { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
      { denom: 4, shaded: 2, label: "2/4", color: C.ACCENT },
    ], { labelW: 0.6, labelFontSize: 12 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Work out 5/6 - 1/3.",
      "Work out 1/2 + 3/8.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why do fractions need the same bottom number before we add them?",
      scItems: [
        "I can add or subtract fractions that have the same denominator.",
        "I can rename one fraction so two fractions have the same denominator.",
        "I can add and subtract two fractions with related denominators.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDec_W5S3_Add_Subtract_Related.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Add and subtract related-denominator fractions by renaming to the same size.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Same bottom number? Add or subtract the tops and keep the bottom. Different but related bottoms? Rename one fraction first so the parts are the same size, then add or subtract the tops.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "1/2 + 1/4: rename 1/2 = 2/4 (same size as quarters). Then 2/4 + 1/4 = 3/4.",
      y);

    y = addSectionHeading(doc, "Section 1 - Same bottom number", y, { color: C.PRIMARY });
    y = addBodyText(doc, "The first one is started for you.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  1/4 + 2/4 = 3/4        b)  2/5 + 1/5 = ______        c)  4/6 - 1/6 = ______", y);
    y = addWriteLine(doc, "d)  3/8 + 2/8 = ______      e)  5/6 - 2/6 = ______", y);

    y = addSectionHeading(doc, "Section 2 - Rename one fraction, then add", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2 + 1/4 = ______      b)  1/3 + 1/6 = ______      c)  1/2 + 3/8 = ______", y);
    y = addWriteLine(doc, "d)  1/4 + 1/8 = ______      e)  2/3 + 1/6 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Rename one fraction, then subtract", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3/4 - 1/2 = ______      b)  5/6 - 1/3 = ______      c)  7/8 - 1/2 = ______", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Fill the gap: 3/4 = 1/2 + ____. Then add three fractions: 1/4 + 1/8 + 1/2.", y);
    y = addWriteLine(doc, "3/4 = 1/2 + ______          1/4 + 1/8 + 1/2 = ______", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Adding and Subtracting Fractions | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the adding and subtracting fractions practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Same bottom number", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3/4      b)  3/5      c)  3/6 (= 1/2)      d)  5/8      e)  3/6 (= 1/2)", y);

    y = addSectionHeading(doc, "Section 2 - Rename, then add", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3/4      b)  3/6 (= 1/2)      c)  7/8      d)  3/8      e)  5/6", y);

    y = addSectionHeading(doc, "Section 3 - Rename, then subtract", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/4      b)  3/6 (= 1/2)      c)  3/8", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "3/4 = 1/2 + 1/4 (since 1/2 = 2/4). 1/4 + 1/8 + 1/2 = 2/8 + 1/8 + 4/8 = 7/8.", y);

    y = addTipBox(doc,
      "Watch for: students who add or subtract the bottom numbers; students who forget to rename; students who do not simplify (3/6 = 1/2).",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Week 5 Session 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
