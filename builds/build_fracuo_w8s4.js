"use strict";

// Understanding & Operating with Fractions - Week 8 (Year 6 Numeracy) - Session 4:
// Solve real-world problems by adding and subtracting fractions (recipes, part cups/litres).
// VC2M5N05. Rename to a common (lowest common) denominator, then operate. Fraction wall + number line.
// Daily Review: Fractions & Decimals + a metric conversion. Fluency: decimal division (division bracket).
// Ships the Year 8 Extension Challenge + answer key for the unit. Variant weekToVariant(8) = 1.

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

const SESSION = 4;
const TOTAL = 4;
const WEEK = 8;
const UNIT_TITLE = "Understanding & Operating with Fractions";
const FOOTER = `Fractions: Understand & Operate | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracUO_W8S4_Word_Problems";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Fraction Word Problems Practice",
  "Solve real-world add and subtract fraction problems by renaming. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Fraction Word Problems Answer Key",
  "Worked answers for the fraction word problems practice sheet.");
const Y8_RES = makeSessionResource(SESSION,
  "Year 8 Extension Challenge",
  "Stretch task for early finishers and advanced students: unrelated denominators with the lowest common denominator, mixed numbers, and multi-step problems.");
const Y8_KEY_RES = makeSessionResource(SESSION,
  "Year 8 Extension Answer Key",
  "Worked answers for the Year 8 Extension Challenge.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, Y8_RES, Y8_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back. We can add and subtract fractions by renaming to the same bottom number.
- Today we use that skill in real life: recipes, jugs and bottles, where amounts are fractions.

DO:
- Have whiteboards, markers and the printed fraction wall ready.
- Settle the class before you click on.

TEACHER NOTES:
Week 8 Session 4 of 4. This session applies the renaming method to word problems. A returning student needs the printed models and the Week 8 Catch-Up Card from Session 1.

WATCH FOR:
- Students who add the amounts without making the parts the same size first.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- Today's materials are the word problems practice sheet and the printed fraction wall.
- There is also a Year 8 Extension Challenge for anyone ready to push further.

DO:
- Print one practice sheet per student and keep both answer keys.
- Print a few copies of the Year 8 Extension Challenge for early finishers.

TEACHER NOTES:
One word problems sheet plus its answer key, and the Year 8 Extension Challenge plus its answer key for advanced students.

CATCH-UP NOTE:
A returner can join here. The launch re-shows why parts must be the same size before combining amounts, and the Week 8 Catch-Up Card from Session 1 summarises the method with a worked example.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Fractions, decimals and one measurement conversion, not today's word problems.
- Answer each one on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and check reasoning.

TEACHER NOTES:
Spaced retrieval across fractions and decimals and one metric conversion to keep measurement warm.

WATCH FOR:
- Students who convert metres to centimetres by multiplying by 100 - secure.
- Students who order decimals by length of the number - reteach place value.

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Seven-tenths is 0.7.
- 2.5 metres is 250 centimetres.
- In order, smallest first: 0.03, 0.3, 0.33.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Metres to centimetres multiplies by 100. Ordering by place value: compare tenths first, then hundredths.

WATCH FOR:
- Students who self-correct - secure.
- Students who place 0.3 after 0.33 - reteach hundredths.

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Dividing decimals, set out in the division bracket.
- Dividend inside the bracket, decimal point lined up straight above, then divide.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for the decimal point carried straight up.

TEACHER NOTES:
Decimal division in the division bracket, last one this week. Call it the division bracket, not the bus stop.

WATCH FOR:
- Students who keep the point lined up - secure.
- Students who stop early on 9.0 divided by 4 - they need to add a zero and keep dividing.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 12.6 divided by 6 is 2.1.
- 9.0 divided by 4 is 2.25.
- 7.8 divided by 3 is 2.6.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
9.0 divided by 4 needs an extra zero in the hundredths: 9.00 divided by 4 is 2.25. Coach adding the zero and continuing.

WATCH FOR:
- Students who self-correct - secure.
- Students who write 2 remainder 1 - show them how to keep going past the point.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Imagine a recipe. It needs half a cup of flour and a quarter of a cup of oats.
- How much dry mix is that altogether? We are adding a half and a quarter.
- Look at the wall - the half and the quarter are different sizes, so first we make them the same size.
- Today we solve real problems like this by renaming, then adding or subtracting.

DO:
- Point to the half bar and the quarter bar on the wall.
- Ask: what could we change the half into so it matches the quarters?

TEACHER NOTES:
This launch puts the renaming method into a real recipe context and re-establishes the wall anchor.

WATCH FOR:
- Students who suggest renaming the half to two-quarters - strong start.
- Students who answer two of something - we will make the parts match first.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to solve real-world problems by adding and subtracting fractions.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is pulling the fractions out of the words, achievable for everyone. Second is renaming to add or subtract, the core target. Third is a two-step problem with a sense check.

WATCH FOR:
- Students who can name the two fractions in a problem - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- One idea to keep in mind: a common denominator, and the lowest common denominator.
- A common denominator gives both fractions the same size parts so we can combine them.
- The lowest common denominator is the smallest one that works. For related bottoms, it is the larger bottom.

DO:
- Show 1/2 and 1/4 becoming 2/4 and 1/4, lowest common denominator four.
- Have students say same-size parts before we combine.

TEACHER NOTES:
Vocabulary after the learning intention. Keep it practical: rename to the same size, ideally the smallest size that fits.

WATCH FOR:
- Students who can name the common denominator - secure.
- Students unsure - point to the wall rows.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Here is the problem. A recipe uses half a cup of flour and a quarter of a cup of oats. How much dry mix altogether?
- First I pull out the fractions: one half and one quarter. The word altogether tells me to add.
- The parts are different sizes, so I rename one half into two-quarters.
- Two-quarters plus one-quarter is three-quarters. So the recipe needs three-quarters of a cup of dry mix.

DO:
- Underline the two fractions and the word altogether.
- Rename on the wall, then add and read the total with its unit, cups.

TEACHER NOTES:
Model the whole process: find the fractions, choose the operation, rename, add, and answer with units.

MISCONCEPTIONS:
- Misconception: students answer two-sixths by adding tops and bottoms.
  Why: they combine all the numbers.
  Impact: a wrong, too-small amount.
  Quick correction: make the parts the same size first; 1/2 becomes 2/4, then add the tops.

WATCH FOR:
- Students who answer 3/4 cup with the unit - secure.
- Students who drop the unit - prompt: three-quarters of what?

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Here is a subtraction problem. A jug holds five-sixths of a litre of juice. You pour out one third of a litre. How much is left?
- The fractions are five-sixths and one third, and how much is left tells me to subtract.
- I rename one third into sixths: one third is two-sixths.
- Five-sixths take away two-sixths is three-sixths, which is one half. So one half of a litre is left.

DO:
- Underline the fractions and the words how much is left.
- Rename on the wall, subtract, and simplify three-sixths to one half.

TEACHER NOTES:
Same method for subtraction in context, finishing with simplifying and stating the unit, litres.

MISCONCEPTIONS:
- Misconception: students subtract the bottom numbers as well.
  Why: they operate on every number.
  Impact: a wrong amount.
  Quick correction: make the parts the same size, then subtract only the tops.

WATCH FOR:
- Students who answer 1/2 of a litre - secure.
- Students who leave 3/6 unsimplified - prompt: can you write that more simply?

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- A jug has two-thirds of a litre of water. You pour in one sixth of a litre more. How much is in the jug now?
- Rename two-thirds into sixths first, then add.

DO:
- Display the prompt.
- Give about 60 seconds.
- Walk and scan for two-thirds renamed to four-sixths.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me how much is in the jug on three, two, one, show.
- Scan for: 2/3 = 4/6, then 4/6 + 1/6 = 5/6 of a litre.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students add to get 3/9, adding tops and bottoms.
- Reteach: rename 2/3 to 4/6 on the wall, then add only the tops.
- Re-check: what is two-thirds renamed into sixths?
- Use the wall to confirm.

TEACHER NOTES:
The renaming step in context is the threshold. Look for the unit, litres, in the answer.

WATCH FOR:
- Students who answer 5/6 of a litre - secure.
- Students who write 3/9 - reteach renaming.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. A muesli recipe needs three-quarters of a cup of oats and one eighth of a cup of seeds.
- How much is that altogether? Pull out the fractions, then rename and add.

DO:
- Display 3/4 and 1/8 on the wall.
- Give about 90 seconds for partners.
- Listen for three-quarters renamed to six-eighths.

TEACHER NOTES:
Related denominators in a recipe. Rename the quarters to eighths, then add the tops.

WATCH FOR:
- Pairs who rename then add - secure.
- Pairs who add without renaming - re-point to same-size parts.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Three-quarters is six-eighths. Six-eighths plus one-eighth is seven-eighths.
- So the recipe needs seven-eighths of a cup altogether.

DO:
- Click to reveal.
- Read the total with its unit, cups.

TEACHER NOTES:
A clean related-denominator addition in context. State the unit in the answer.

WATCH FOR:
- Students who self-correct - secure.
- Students who add bottoms - small group before You Do.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together. You have seven-eighths of a cup of flour and use one half of a cup.
- How much flour is left? Pull out the fractions, then rename and subtract.

DO:
- Display 7/8 and 1/2 on the wall.
- Give about 90 seconds.
- Listen for one half renamed to four-eighths.

TEACHER NOTES:
A two-step feel: rename, then subtract, then state the unit. The wall shows the take-away.

WATCH FOR:
- Students who rename then subtract - secure.
- Students who subtract bottoms - re-point to same-size parts.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- One half is four-eighths. Seven-eighths take away four-eighths is three-eighths.
- So three-eighths of a cup of flour is left.

DO:
- Click to reveal.
- Read the answer with its unit, cups.

TEACHER NOTES:
Finish the week by checking the answer makes sense: less than the seven-eighths you started with.

WATCH FOR:
- Students who answer 3/8 cup and check it is sensible - ready for independent work.
- Students unsure - one more before You Do.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the word problems sheet.
- For each one: pull out the fractions, decide add or subtract, rename, then solve.
- Always write the unit in your answer, like cups or litres.
- If you finish, try the extension, or ask for the Year 8 Extension Challenge.

DO:
- Distribute the word problems sheet.
- Circulate and listen for find the fractions, rename, then operate.
- Cold call one or two students to explain a problem.

TEACHER NOTES:
Different contexts from the We Do, same method: read the problem, rename to a common denominator, then add or subtract, and answer with units.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed fraction wall and do the first two problems. The fractions are underlined for you and the first one is started.
- Extra Notes: Sit with these students and find the fractions in the words together.
EXTENDING PROMPT:
- Task: Extension - a two-step problem that combines an addition and a subtraction. Then move on to the Year 8 Extension Challenge, which uses unrelated denominators and mixed numbers.
- Extra Notes: Push the sense check: is the answer a sensible amount?

WATCH FOR:
- Students who rename then operate and state units - secure.
- Students who combine amounts without renaming - prompt back to same-size parts.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task for the week, on your whiteboard.
- A bottle has five-sixths of a litre of water. You drink one half of a litre. How much is left?
- Then, a recipe uses one quarter of a cup of butter and three-eighths of a cup of sugar. How much altogether?

DO:
- Display the prompt.
- Give about 4 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - solving real problems by renaming. 5/6 - 1/2 uses sixths (1/2 = 3/6, so 2/6 = 1/3 of a litre). 1/4 + 3/8 uses eighths (1/4 = 2/8, so 5/8 of a cup).

WATCH FOR:
- Students who answer 1/3 of a litre and 5/8 of a cup with units - secure.
- Students who operate on bottoms - plan a short revision next week.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: when you read a fraction problem, what is your first step?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea for the week: rename to the same size parts, then add or subtract the tops, and answer with units.

WATCH FOR:
- Strong thumbs across all three - the unit goal is met.
- Sideways or down on the core criterion - plan a short revision early next week.

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
  titleSlide(pres, UNIT_TITLE, "Week 8 Session 4: Real-world fraction problems",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Fractions & decimals + metric
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions, decimals & measurement",
      [
        "Write 7/10 as a decimal",
        "Convert: 2.5 m = ____ cm",
        "Order smallest first: 0.3, 0.03, 0.33",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "0.7      250 cm      0.03, 0.3, 0.33", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal division (division bracket)
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal division (division bracket)",
      ["12.6 ÷ 6", "9.0 ÷ 4", "7.8 ÷ 3"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "2.1        2.25        2.6", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - a recipe needs 1/2 cup + 1/4 cup
  contentSlide(pres, "Launch", C.ACCENT, "A recipe needs 1/2 cup and 1/4 cup",
    [
      "How much dry mix altogether?",
      "",
      "The half and the quarter are",
      "different sizes - match them first.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Half a cup and a quarter cup", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 1, label: "1/4", color: C.ACCENT },
      ], { strip: C.ACCENT, cardH: 2.1 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to solve real-world problems by adding and subtracting fractions.",
    [
      "I can find the fractions in a word problem.",
      "I can rename fractions to add or subtract them.",
      "I can solve a two-step problem and check it makes sense.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Common & lowest common denominator",
    [
      "Common denominator: both fractions share a bottom (same-size parts).",
      "Lowest common denominator: the smallest one that works (for related bottoms, the larger bottom).",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Rename to the same size", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 2, label: "2/4", color: C.SECONDARY },
      ], { strip: C.SECONDARY, cardH: 1.9 });
    }
  );

  // Slide 10: I Do #1 - recipe add (1/2 + 1/4)
  workedExSlide(pres, 2, "I Do", "Recipe: 1/2 cup + 1/4 cup",
    [
      "Half a cup of flour, a quarter cup of oats. How much altogether?",
      "",
      "Fractions: 1/2 and 1/4. 'Altogether' = add.",
      "Rename: 1/2 = 2/4.",
      "2/4 + 1/4 = 3/4.",
      "",
      "Answer: 3/4 of a cup.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Rename 1/2 to 2/4, then add", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 2, label: "2/4", color: C.SECONDARY },
        { denom: 4, shaded: 3, label: "3/4", color: C.ACCENT },
      ], { strip: C.PRIMARY, cardH: 2.6 });
    }
  );

  // Slide 11: I Do #2 - juice subtract (5/6 - 1/3)
  workedExSlide(pres, 2, "I Do", "Jug: 5/6 L take away 1/3 L",
    [
      "A jug holds 5/6 L. You pour out 1/3 L. How much is left?",
      "",
      "Fractions: 5/6 and 1/3. 'Left' = subtract.",
      "Rename: 1/3 = 2/6.",
      "5/6 - 2/6 = 3/6 = 1/2.",
      "",
      "Answer: 1/2 of a litre.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Rename 1/3 to 2/6, then take away", [
        { denom: 6, shaded: 5, label: "5/6", color: C.SECONDARY },
        { denom: 3, shaded: 1, label: "1/3", color: C.PRIMARY },
        { denom: 6, shaded: 3, label: "3/6", color: C.ACCENT },
      ], { strip: C.PRIMARY, cardH: 2.6 });
    }
  );

  // Slides 12-13: CFU + reveal - jug 2/3 + 1/6
  withReveal(
    () => cfuSlide(pres, "CFU", "How much water is in the jug now?", "Show Me Boards",
      "A jug has 2/3 L. You pour in 1/6 L more.\n\nRename 2/3 into sixths, then add. How much now?",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "2/3 = 4/6,  so  4/6 + 1/6 = 5/6 of a litre",
        { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - muesli 3/4 + 1/8
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Muesli: 3/4 cup + 1/8 cup",
      [
        "With your partner.",
        "",
        "1.  Fractions: 3/4 and 1/8.",
        "2.  Rename 3/4 into eighths.",
        "3.  Add, then write the unit.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Rename 3/4 to 6/8, then add", [
          { denom: 4, shaded: 3, label: "3/4", color: C.SECONDARY },
          { denom: 8, shaded: 6, label: "6/8", color: C.PRIMARY },
          { denom: 8, shaded: 7, label: "7/8", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "3/4 + 1/8 = 6/8 + 1/8 = 7/8 of a cup", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - flour 7/8 - 1/2
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Flour: 7/8 cup take away 1/2 cup",
      [
        "With your partner.",
        "",
        "1.  Fractions: 7/8 and 1/2.",
        "2.  Rename 1/2 into eighths.",
        "3.  Subtract, then write the unit.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Rename 1/2 to 4/8, then take away", [
          { denom: 8, shaded: 7, label: "7/8", color: C.SECONDARY },
          { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
          { denom: 8, shaded: 3, label: "3/8", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "7/8 - 1/2 = 7/8 - 4/8 = 3/8 of a cup", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // Slide 18: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: word problems sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "find the fractions in the problem.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "rename, then add or subtract.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "write the unit and check it.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Make the parts the same size, add or subtract the tops, then answer with the unit (cups, litres).", {
      x: 0.85, y: panelY + 0.52, w: 8.3, h: 0.5, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFractionStripSet(s, 2.5, panelY + 1.2, 5.0, 0.85, [
      { denom: 4, shaded: 3, label: "3/4", color: C.PRIMARY },
      { denom: 8, shaded: 6, label: "6/8", color: C.ACCENT },
    ], { labelW: 0.6, labelFontSize: 12 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "A bottle has 5/6 L. You drink 1/2 L. How much is left?",
      "A recipe uses 1/4 cup butter and 3/8 cup sugar. How much altogether?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: when you read a fraction problem, what is your first step?",
      scItems: [
        "I can find the fractions in a word problem.",
        "I can rename fractions to add or subtract them.",
        "I can solve a two-step problem and check it makes sense.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracUO_W8S4_Word_Problems.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Word problems worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Solve real-world problems by adding and subtracting fractions.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "For each problem: find the two fractions, decide add or subtract, rename so the bottoms match, then add or subtract the tops. Always write the unit in your answer (cups, litres).",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "A recipe needs 1/2 cup of flour and 1/4 cup of oats. How much altogether? 1/2 = 2/4, so 2/4 + 1/4 = 3/4 cup.",
      y);

    y = addSectionHeading(doc, "Section 1 - Adding (the first is started for you)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2 cup sugar + 1/4 cup honey.  1/2 = 2/4, so 2/4 + 1/4 = ____ cup", y);
    y = addWriteLine(doc, "b)  1/3 L water + 1/6 L juice = ____ L", y);
    y = addWriteLine(doc, "c)  3/8 cup oats + 1/4 cup seeds = ____ cup", y);

    y = addSectionHeading(doc, "Section 2 - Subtracting", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "d)  3/4 cup flour, use 1/2 cup. Left = ____ cup", y);
    y = addWriteLine(doc, "e)  5/6 L milk, use 1/3 L. Left = ____ L", y);
    y = addWriteLine(doc, "f)  7/8 cup rice, use 1/4 cup. Left = ____ cup", y);

    y = addSectionHeading(doc, "Section 3 - Two-step", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "g)  A jug has 1/2 L. Add 1/4 L, then drink 1/8 L. How much now? = ____ L", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "A bottle holds 3/4 L. You drink 1/3 L, then top it up with 1/6 L. How much is in the bottle now? (Hint: rename everything to sixths.)", y);
    y = addWriteLine(doc, "Answer: ______ L", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Fraction Word Problems | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Word problems answer key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the fraction word problems practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Adding", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3/4 cup.   b)  1/3 + 1/6 = 2/6 + 1/6 = 3/6 = 1/2 L.   c)  3/8 + 1/4 = 3/8 + 2/8 = 5/8 cup.", y);

    y = addSectionHeading(doc, "Section 2 - Subtracting", y, { color: C.PRIMARY });
    y = addBodyText(doc, "d)  3/4 - 1/2 = 3/4 - 2/4 = 1/4 cup.   e)  5/6 - 1/3 = 5/6 - 2/6 = 3/6 = 1/2 L.   f)  7/8 - 1/4 = 7/8 - 2/8 = 5/8 cup.", y);

    y = addSectionHeading(doc, "Section 3 - Two-step", y, { color: C.PRIMARY });
    y = addBodyText(doc, "g)  1/2 + 1/4 = 3/4, then 3/4 - 1/8 = 6/8 - 1/8 = 5/8 L.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "3/4 - 1/3 + 1/6. Rename to sixths only where needed using twelfths: 3/4 = 9/12, 1/3 = 4/12, 1/6 = 2/12. 9/12 - 4/12 + 2/12 = 7/12 L.", y);

    y = addTipBox(doc,
      "Watch for: students who combine amounts without renaming; students who drop the unit; students who forget to simplify (3/6 = 1/2).",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension Challenge
  await (async () => {
    const doc = createPdf({ title: Y8_RES.name });
    let y = addPdfHeader(doc, Y8_RES.name, {
      subtitle: "A stretch task for students ready to go beyond Year 6 fractions.",
      color: C.SECONDARY,
      lessonInfo: `Week ${WEEK} | Year 6 Numeracy | Aimed at Year 8 level`,
    });
    y = addTipBox(doc,
      "When the bottom numbers are NOT related, find the lowest common denominator (the smallest number both bottoms divide into), rename both fractions, then add or subtract. For mixed numbers, you may rename the whole part into the fraction first.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 - Unrelated denominators (find the LCD)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2/3 + 1/4 = ______        b)  5/6 - 3/8 = ______", y);
    y = addWriteLine(doc, "c)  3/4 + 2/5 = ______        d)  7/10 - 2/3 = ______", y);

    y = addSectionHeading(doc, "Section 2 - Mixed numbers", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "e)  1 1/2 + 2 3/4 = ______        f)  3 1/3 - 1 5/6 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Multi-step problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "g)  A muffin recipe uses 2/3 cup of sugar and 3/4 cup of flour. You make a double batch.", y);
    y = addWriteLine(doc, "How much sugar in total? ______ cups        How much flour in total? ______ cups", y);
    y = addWriteLine(doc, "In one batch, how much more flour than sugar? ______ cup", y);

    y = addSectionHeading(doc, "Section 4 - Multiply (challenge)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "h)  2/3 × 3/4 = ______        i)  3/5 of 25 = ______", y);

    addPdfFooter(doc, `Week ${WEEK} | Year 8 Extension Challenge | Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, Y8_RES.fileName));
    console.log("PDF written: " + Y8_RES.fileName);
  })();

  // Year 8 Extension Answer Key
  await (async () => {
    const doc = createPdf({ title: Y8_KEY_RES.name });
    let y = addPdfHeader(doc, Y8_KEY_RES.name, {
      subtitle: "Worked answers for the Year 8 Extension Challenge.",
      color: C.SECONDARY,
      lessonInfo: `Week ${WEEK} | Year 6 Numeracy | Aimed at Year 8 level`,
    });

    y = addSectionHeading(doc, "Section 1 - Unrelated denominators", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2/3 + 1/4: LCD 12 -> 8/12 + 3/12 = 11/12.", y);
    y = addBodyText(doc, "b)  5/6 - 3/8: LCD 24 -> 20/24 - 9/24 = 11/24.", y);
    y = addBodyText(doc, "c)  3/4 + 2/5: LCD 20 -> 15/20 + 8/20 = 23/20 = 1 3/20.", y);
    y = addBodyText(doc, "d)  7/10 - 2/3: LCD 30 -> 21/30 - 20/30 = 1/30.", y);

    y = addSectionHeading(doc, "Section 2 - Mixed numbers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "e)  1 1/2 + 2 3/4 = 1 2/4 + 2 3/4 = 3 5/4 = 4 1/4.", y);
    y = addBodyText(doc, "f)  3 1/3 - 1 5/6 = 3 2/6 - 1 5/6 = 2 8/6 - 1 5/6 = 1 3/6 = 1 1/2.", y);

    y = addSectionHeading(doc, "Section 3 - Multi-step problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "g)  Sugar: 2 x 2/3 = 4/3 = 1 1/3 cups. Flour: 2 x 3/4 = 3/2 = 1 1/2 cups. More flour than sugar in one batch: 3/4 - 2/3 = 9/12 - 8/12 = 1/12 cup.", y);

    y = addSectionHeading(doc, "Section 4 - Multiply", y, { color: C.ACCENT });
    y = addBodyText(doc, "h)  2/3 × 3/4 = 6/12 = 1/2.   i)  3/5 of 25 = 15.", y);

    y = addTipBox(doc,
      "Watch for: students who use any common denominator but never the lowest; students who add the whole numbers but forget to regroup the fraction part; students who multiply tops and bottoms when they should be finding a common denominator to add.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} | Year 8 Extension Answer Key | Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, Y8_KEY_RES.fileName));
    console.log("PDF written: " + Y8_KEY_RES.fileName);
  })();

  console.log("Week 8 Session 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
