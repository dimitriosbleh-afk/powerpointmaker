"use strict";

// Fractions & Decimals - Week 5 (Year 6 Numeracy) - Session 1: Equivalent fractions & the fraction wall
// VC2M5N03 / VC2M6N03. First session: establishes the two anchor models (fraction wall + number line)
// used across the whole 2-week unit, so a returning student can re-orient quickly.
// Daily Review: Metric measurement & unit conversion. Fluency: decimal addition algorithm.
// Unit variant fixed at weekToVariant(5) across all 4 Week-5 lessons for cohesion.

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

const SESSION = 1;
const TOTAL = 4;
const WEEK = 5;
const UNIT_TITLE = "Fractions & Decimals";
const FOOTER = `Fractions & Decimals | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracDec_W5S1_Equivalent_Fractions";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Equivalent Fractions Practice",
  "Find and check equivalent fractions using the wall and the multiply rule. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Equivalent Fractions Answer Key",
  "Worked answers for the equivalent fractions practice sheet.");
const CATCHUP_RES = makeSessionResource(SESSION,
  "Week 5 Catch-Up Card",
  "One-page recap of the four Week 5 ideas with a worked example each. For a student returning after an absence.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, CATCHUP_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to our new two-week unit on fractions and decimals.
- This week we build fractions carefully, starting with the idea that the same amount can be written in more than one way.
- Today we meet our two big tools for the whole unit, the fraction wall and the number line.

DO:
- Have whiteboards, markers and the printed fraction wall ready.
- Settle the class and take a calm breath before you click on.

TEACHER NOTES:
Session 1 of 4. The fraction wall established today is the anchor for the next two weeks, so spend the time to make it feel familiar.

WATCH FOR:
- Students who look unsure - that is expected on day one. Reassure them: if this feels new, that is okay, we build it together.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end, and it has an easier start and a challenge section built in.

DO:
- Print one practice sheet per student and keep the answer key for yourself.
- Have whiteboards, markers and the printed fraction wall ready.

TEACHER NOTES:
One student practice sheet plus its answer key. Most teaching is on whiteboards with the fraction wall. The Week 5 Catch-Up Card is for absences, not for everyone.

CATCH-UP NOTE:
A student who misses sessions can rejoin at any lesson this unit. Every lesson re-builds the fraction wall in the launch, so a returner only needs the printed fraction wall, the one-page Week 5 Catch-Up Card, and one minute with you.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are thinking back to our measurement work, not today's fractions.
- Remember, 100 centimetres make 1 metre, and 1000 millilitres make 1 litre.
- Read each one carefully and write your answer on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and check that students notice the unit, not just the number.

TEACHER NOTES:
Daily Review is prior measurement learning. It keeps unit conversion fresh while today's lesson is fractions.

WATCH FOR:
- Students who convert by thinking about how many of the smaller unit fit in the larger - secure.
- Students who just move the decimal point by guessing - prompt them to say the conversion fact first.

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check together.
- 250 centimetres is 2.5 metres, because 100 centimetres make a metre.
- 1.5 litres is 1500 millilitres.
- Shortest to longest: 0.6 metres which is 60 centimetres, then 75 centimetres, then 800 millimetres which is 80 centimetres.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The ordering item is the discriminator. Students must put everything in the same unit before comparing.

WATCH FOR:
- Students who convert to a common unit before ordering - secure.
- Students who compare raw numbers across units - reteach with a common unit in small group.

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Decimal addition, set out vertically.
- Line up the decimal points so ones sit under ones and tenths under tenths.
- Add, carry where you need to, and bring the decimal point straight down.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up decimal points.

TEACHER NOTES:
Fluency this week is the decimal addition algorithm. Lining up the decimal point is the habit that makes the rest of the week accurate.

WATCH FOR:
- Students who line up the points neatly - secure.
- Students who right-align the digits instead of the points - this is the key error to coach.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 4.6 plus 3.7 is 8.3.
- 12.5 plus 8.45 is 20.95.
- 6.08 plus 2.9 is 8.98.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
6.08 plus 2.9 is the trap, because 2.9 must be read as 2.90 to line up. Coach filling the empty place with a zero.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students who add 6.08 and 2.9 as 6.08 plus 0.29 - they misaligned the points, small group focus.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Look at our fraction wall. Each row is one whole, just cut into a different number of equal parts.
- Find the one half bar. Now look down the wall. Which other bars reach exactly the same line as one half?
- Whisper to your partner what you notice.
- When bars line up to the same length, they are the same amount. We call those equivalent fractions, and that is today's learning.

DO:
- Point along the half line across the wall.
- Take two or three pairs' ideas before you reveal the word equivalent.

TEACHER NOTES:
This launch activates the wall before the learning intention, with the new word equivalent placed exactly where the lesson is going.

WATCH FOR:
- Students who spot that two-quarters and four-eighths line up with one half - strong start.
- Students unsure where to look - point to the half line and ask what lines up underneath.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to find equivalent fractions using a fraction wall and a number line.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Hold up the printed fraction wall as you read the first one.

TEACHER NOTES:
The first criterion is reachable for everyone, naming a bar. The second is the core target the exit ticket checks. The third stretches to explaining with multiplication.

WATCH FOR:
- Students who can repeat the language - tracking with you.

[Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- Equivalent means equal in value, the same amount shown a different way.
- The numerator is the top number. It counts the parts we have.
- The denominator is the bottom number. It names how many equal parts are in the whole.

DO:
- Point to the top and bottom of the fraction as you name them.
- Have students say each word once and point to the part of the fraction.

TEACHER NOTES:
Vocabulary comes after the learning intention. Keep it to these three words. Equivalent is the key idea; numerator and denominator just let us talk about it.

WATCH FOR:
- Students who can point to the numerator and denominator correctly - secure.
- Students who swap them - say top counts, bottom names the size.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at one half three ways. Watch the wall.
- One half fills exactly half of the whole.
- Now two-quarters. Two of the four parts also fills exactly half. The lengths match.
- Now three-sixths. Three of the six parts again fills exactly half.
- So one half, two-quarters and three-sixths are all the same amount. They are equivalent.

DO:
- Run your finger along each shaded bar to the same line.
- Say the sentence: same length means same amount means equivalent.

TEACHER NOTES:
This is the concrete anchor. Equivalence is sameness of amount, seen as same length on the wall, before any rule.

MISCONCEPTIONS:
- Misconception: a fraction with bigger numbers is always bigger.
  Why: students compare the digits, not the amount.
  Impact: they think three-sixths is larger than one half.
  Quick correction: line the bars up on the wall and show the lengths match.

WATCH FOR:
- Students who say the lengths are equal - secure.
- Students focused on the bigger numbers - return to the wall lengths.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now the rule that makes equivalent fractions without the wall.
- Multiply the top and the bottom by the same number.
- One half, multiply top and bottom by two, becomes two-quarters.
- One half, multiply top and bottom by three, becomes three-sixths.
- Because we multiplied the top and bottom by the same number, the amount did not change.

DO:
- Write the multiply by two and multiply by three steps as you say them.
- Point back to the wall to show the rule matches the picture.

TEACHER NOTES:
Connect the rule to the picture from the first I Do. Multiplying top and bottom by the same number is making more, smaller parts of the same whole.

MISCONCEPTIONS:
- Misconception: students multiply only the top, or add the same number to top and bottom.
  Why: they remember half a rule.
  Impact: they change the value, for example one half becoming one-quarter or two-thirds.
  Quick correction: the same number multiplies BOTH the top and the bottom.

WATCH FOR:
- Students who apply the rule and predict the wall - secure.
- Students who add instead of multiply - re-show with the wall.

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Which fraction is equivalent to two-thirds? Is it three-quarters, four-sixths, or two-sixths?
- Write the letter and be ready to say why.

DO:
- Display the three options.
- Give about 45 seconds.
- Walk and scan the boards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me your answer on three, two, one, show.
- Scan for: the letter b, four-sixths, with reasoning about multiplying by two.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students pick three-quarters because the numbers look close.
- Reteach: two-thirds, multiply top and bottom by two, gives four-sixths. Show it on the wall.
- Re-check: what do you multiply two-thirds by to make sixths?
- Use the wall to confirm before moving on.

TEACHER NOTES:
The distractor three-quarters catches students who do not check with multiplication or the wall.

WATCH FOR:
- Students who justify with multiply by two - secure.
- Students who guess from the look of the numbers - reteach with the wall.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Find a fraction equal to three-quarters.
- Look at the wall. Which eighths bar lines up exactly with three-quarters?
- Then check it with the multiply rule.

DO:
- Display three-quarters on the wall above the eighths row.
- Give about 90 seconds for partners.
- Listen for multiply top and bottom by two.

TEACHER NOTES:
Same structure as the I Do, new fraction. Students use the wall first, then confirm with multiplication.

WATCH FOR:
- Pairs who find six-eighths and justify it - secure.
- Pairs who count parts unevenly - re-point to the lined-up bars.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Three-quarters lines up exactly with six-eighths on the wall.
- And the rule agrees: three-quarters, multiply top and bottom by two, is six-eighths.

DO:
- Click to reveal.
- Run the wall and the rule once more together.

TEACHER NOTES:
Both representations agree. That agreement is the point - picture and rule give the same answer.

WATCH FOR:
- Students who self-correct - secure.
- Students who still mis-shade - small group with the wall before You Do.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, this time on the number line.
- One half and two-quarters are equivalent. Watch where they sit on the line.
- Both land on exactly the same point, halfway between zero and one.
- On the wall, equivalent means the same length. On the number line, it means the same point.

DO:
- Display the number line from zero to one.
- Ask students to predict the point before you reveal it.
- Mark the halfway point as both one half and two-quarters.

TEACHER NOTES:
This connects the wall to the number line, exactly the curriculum link between an area model and a number-line model. Same amount, two pictures.

WATCH FOR:
- Students who point to the halfway mark for both - secure.
- Students who expect two-quarters further along because four is bigger - return to the wall length.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Here it is. One half and two-quarters share the same point on the line.
- Same amount, same place. That is equivalence on a number line.

DO:
- Click to reveal the shared point.
- Restate: same length on the wall, same point on the line.

TEACHER NOTES:
Keep the two pictures side by side in students' minds. Next session uses this same number line for mixed and improper fractions.

WATCH FOR:
- Students who connect wall length to line point - ready for independent work.
- Students still surprised - one more example on the line before You Do.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- The first section uses the fraction wall to fill in the missing number.
- The next section uses the multiply rule.
- The last section is true or false - decide if two fractions are equivalent and prove it.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for multiply top and bottom by the same number.
- Cold call one or two students to justify a true-or-false answer.

TEACHER NOTES:
Different fractions from the We Do, same strategy: use the wall, then the multiply rule, then check.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed fraction wall and do Section 1 only. The first one is started for you.
- Extra Notes: Sit with these students and line up the first bar together.
EXTENDING PROMPT:
- Task: Extension - find three different fractions equivalent to three-fifths and explain the pattern in the numbers. Early finishers may start the Year 8 Extension Challenge.
- Extra Notes: Push the language: top and bottom multiplied by the same number.

WATCH FOR:
- Students who use the multiply rule fluently - secure.
- Students who add instead of multiply - prompt back to the same-number rule.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Write two fractions that are equivalent to one-third.
- Then show or explain how you know one of them is equal to one-third.

DO:
- Display the prompt.
- Give about 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - finding equivalent fractions and justifying. Look for multiply top and bottom by the same number, or a wall drawing.

WATCH FOR:
- Students who write two-sixths and three-ninths with reasoning - secure.
- Students who change only the top or only the bottom - revisit the same-number rule in Session 2.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: what does it mean for two fractions to be equivalent?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that equivalent fractions are the same amount shown differently - same length on the wall, same point on the line, made by multiplying top and bottom by the same number.

WATCH FOR:
- Strong thumbs across all three - move at pace next session.
- Sideways or down on the core criterion - quick small-group revision at the start of Session 2.

[Retention and Recall]`;

// --- Helpers -----------------------------------------------------------------

// Fraction wall drawn into a right-column panel of a worked-example/content slide.
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
  titleSlide(pres, UNIT_TITLE, "Week 5 Session 1: Equivalent fractions and the fraction wall",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Metric measurement
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Metric measurement",
      [
        "Convert: 250 cm = ____ m",
        "Convert: 1.5 L = ____ mL",
        "Order shortest to longest: 0.6 m, 75 cm, 800 mm",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "2.5 m     1500 mL     0.6 m, 75 cm, 800 mm",
        { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal addition",
      ["4.6 + 3.7", "12.5 + 8.45", "6.08 + 2.9"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "8.3        20.95        8.98", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - fraction wall
  contentSlide(pres, "Launch", C.ACCENT, "What lines up with one half?",
    [
      "Each row is one whole in equal parts.",
      "",
      "Which bars line up with 1/2?",
      "Same length = same amount = equivalent.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Fraction wall", [
        { denom: 1, shaded: 1, label: "1", color: C.PRIMARY },
        { denom: 2, shaded: 1, label: "1/2", color: C.SECONDARY },
        { denom: 4, shaded: 2, label: "2/4", color: C.ACCENT },
        { denom: 8, shaded: 4, label: "4/8", color: C.SUCCESS },
      ], { strip: C.ACCENT, cardH: 2.7 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to find equivalent fractions using a fraction wall and a number line.",
    [
      "I can name the fraction that a bar on the fraction wall shows.",
      "I can find a fraction equal to another fraction using the wall.",
      "I can explain why two fractions are equivalent by multiplying the top and bottom.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Equivalent, numerator, denominator",
    [
      "Equivalent = equal in value (same amount).",
      "Numerator = top number (parts we have).",
      "Denominator = bottom number (parts in the whole).",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
      slide.addText("1/2  =  2/4", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.18, w: lg.rightW - 0.4, h: 0.5,
        fontSize: 30, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      addFractionStripSet(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.85, lg.rightW - 0.6, 0.9, [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 2, label: "2/4", color: C.ACCENT },
      ], { labelW: 0.6, labelFontSize: 11 });
      slide.addText("Same amount, written two ways.", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.9, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 10: I Do #1 - one half three ways
  workedExSlide(pres, 2, "I Do", "One half, shown three ways",
    [
      "1/2 fills half the whole.",
      "2/4 also fills half.",
      "3/6 also fills half.",
      "",
      "Same length = same amount.",
      "So 1/2 = 2/4 = 3/6.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "All reach the same line", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 2, label: "2/4", color: C.SECONDARY },
        { denom: 6, shaded: 3, label: "3/6", color: C.ACCENT },
      ], { strip: C.PRIMARY, cardH: 2.6 });
    }
  );

  // Slide 11: I Do #2 - multiply top and bottom
  workedExSlide(pres, 2, "I Do", "Make equivalent fractions by multiplying",
    [
      "Multiply top AND bottom by the same number.",
      "",
      "1/2 = (1x2)/(2x2) = 2/4",
      "1/2 = (1x3)/(2x3) = 3/6",
      "",
      "Same number on top and bottom",
      "keeps the value the same.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.PRIMARY });
      const rows = [
        ["1/2", "x2", "2/4"],
        ["1/2", "x3", "3/6"],
      ];
      const ry0 = lg.panelTopPadded + 0.3;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.85;
        slide.addText(r[0], {
          x: lg.rightX + 0.25, y: ry, w: 1.0, h: 0.6,
          fontSize: 26, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        addTextOnShape(slide, r[1] + " top & bottom", {
          x: lg.rightX + 1.35, y: ry + 0.08, w: 1.7, h: 0.44, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
        slide.addText("= " + r[2], {
          x: lg.rightX + 3.1, y: ry, w: 1.0, h: 0.6,
          fontSize: 26, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
      slide.addText("The amount does not change.", {
        x: lg.rightX + 0.2, y: ry0 + 1.85, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slides 12-13: CFU + reveal - equivalent to 2/3
  withReveal(
    () => cfuSlide(pres, "CFU", "Which is equivalent to 2/3?", "Show Me Boards",
      "Which fraction is equivalent to 2/3?\n\na)  3/4          b)  4/6          c)  2/6",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "b)  4/6   (2/3 with top and bottom x2)",
        { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - equal to 3/4
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Find a fraction equal to 3/4",
      [
        "With your partner.",
        "",
        "1.  Find 3/4 on the wall.",
        "2.  Which eighths bar lines up?",
        "3.  Check with x2 top and bottom.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Line up 3/4 with eighths", [
          { denom: 4, shaded: 3, label: "3/4", color: C.SECONDARY },
          { denom: 8, shaded: 6, label: "6/8", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.0 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "3/4 = 6/8   (x2 top and bottom)", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - same point on the number line (full width)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Same point on the number line", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 1.0, { strip: STAGE_COLORS["3"] });
      s.addText("1/2 and 2/4 are equivalent. Predict: where do they sit between 0 and 1?", {
        x: 0.75, y: CONTENT_TOP + 0.12, w: 8.5, h: 0.76,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      const lineY = CONTENT_TOP + 1.9;
      addCard(s, 0.5, CONTENT_TOP + 1.25, 9.0, SAFE_BOTTOM - (CONTENT_TOP + 1.25),
        { strip: C.PRIMARY, fill: C.WHITE });
      addNumberLine(s, 1.0, lineY, 8.0,
        ["0", "", "1/4", "", "1/2", "", "3/4", "", "1"], [],
        { labelFontSize: 14 });
      s.addText("Halfway = 1/2 = 2/4", {
        x: 1.0, y: lineY + 0.55, w: 8.0, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1/2 and 2/4 share the SAME point - halfway between 0 and 1", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "use the wall to fill the missing number.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "use the x rule to make equivalents.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "decide true or false and prove it.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Multiply the top AND the bottom by the same number. The amount stays the same.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFractionStripSet(s, 2.5, panelY + 1.15, 5.0, 0.85, [
      { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
      { denom: 4, shaded: 2, label: "2/4", color: C.ACCENT },
    ], { labelW: 0.6, labelFontSize: 12 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Write two fractions that are equivalent to 1/3.",
      "Show or explain how you know one of them is equal to 1/3.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what does it mean for two fractions to be equivalent?",
      scItems: [
        "I can name the fraction that a bar on the fraction wall shows.",
        "I can find a fraction equal to another fraction using the wall.",
        "I can explain why two fractions are equivalent by multiplying the top and bottom.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDec_W5S1_Equivalent_Fractions.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find and check equivalent fractions using the wall and the multiply rule.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Equivalent fractions are the same amount written a different way. To make one, multiply the top AND the bottom by the same number. The value stays the same.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "1/2 = 2/4. Multiply top and bottom by 2: (1x2)/(2x2) = 2/4. On the fraction wall, 1/2 and 2/4 line up to the same length.",
      y);

    y = addSectionHeading(doc, "Section 1 - Use the fraction wall to fill the missing number", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Line up the bars. The first one is started for you.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  1/2 = 2/4          b)  1/2 = ___/6          c)  1/3 = ___/6", y);
    y = addWriteLine(doc, "d)  2/5 = ___/10        e)  3/4 = ___/8          f)  1/4 = ___/8", y);

    y = addSectionHeading(doc, "Section 2 - Multiply to make an equivalent fraction", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3/4 = ___/8   (x2)        b)  2/3 = ___/9   (x3)", y);
    y = addWriteLine(doc, "c)  1/5 = ___/10  (x2)        d)  2/3 = ___/12  (x4)", y);

    y = addSectionHeading(doc, "Section 3 - True or false? Prove it.", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2/3 = 4/6 ?  ____   Why? ___________________________________", y);
    y = addWriteLine(doc, "b)  1/2 = 3/5 ?  ____   Why? ___________________________________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Find THREE fractions equivalent to 3/5. Explain the pattern you see in the numbers.", y);
    y = addWriteLine(doc, "3/5 = ____ = ____ = ____", y);
    y = addWriteLine(doc, "Pattern: ___________________________________________________", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Equivalent Fractions | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the equivalent fractions practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Use the fraction wall", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2/4      b)  3/6      c)  2/6      d)  4/10      e)  6/8      f)  2/8", y);

    y = addSectionHeading(doc, "Section 2 - Multiply to make an equivalent", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  6/8      b)  6/9      c)  2/10      d)  8/12", y);

    y = addSectionHeading(doc, "Section 3 - True or false", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  TRUE. 2/3 multiplied top and bottom by 2 is 4/6.", y);
    y = addBodyText(doc, "b)  FALSE. 1/2 = 5/10 and 3/5 = 6/10, so they are not equal.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "3/5 = 6/10 = 9/15 = 12/20 (any correct multiples). Pattern: the top and bottom both multiply by the same number each time.", y);

    y = addTipBox(doc,
      "Watch for: students who change only the top or only the bottom; students who add the same number instead of multiplying; students who compare digit size instead of amount.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: CATCHUP_RES.name });
    let y = addPdfHeader(doc, CATCHUP_RES.name, {
      subtitle: "Missed a lesson? Here are the four big ideas of Week 5 with one worked example each.",
      color: C.SECONDARY,
      lessonInfo: `Week ${WEEK} | Year 6 Numeracy | Catch-up`,
      showNameDate: false,
    });
    y = addTipBox(doc,
      "Our two tools all unit are the FRACTION WALL (same length = same amount) and the NUMBER LINE (same point = same amount). Keep this card in your book.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Session 1 - Equivalent fractions", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Same amount, written differently. Multiply top and bottom by the same number. 1/2 = (1x2)/(2x2) = 2/4.", y);

    y = addSectionHeading(doc, "Session 2 - Mixed numbers and improper fractions", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Improper to mixed: 7/4 = 1 whole and 3/4 = 1 3/4. Mixed to improper: 2 1/3 = (2x3 + 1)/3 = 7/3.", y);

    y = addSectionHeading(doc, "Session 3 - Adding and subtracting related fractions", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Make the parts the same size first. 1/2 + 1/4 = 2/4 + 1/4 = 3/4.", y);

    y = addSectionHeading(doc, "Session 4 - Lowest common denominator", y, { color: C.PRIMARY });
    y = addBodyText(doc, "When bottoms are different, list multiples to find a common bottom. 1/2 + 1/3: common bottom 6, so 3/6 + 2/6 = 5/6.", y);

    addPdfFooter(doc, `Week ${WEEK} | Catch-Up Card | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, CATCHUP_RES.fileName));
    console.log("PDF written: " + CATCHUP_RES.fileName);
  })();

  console.log("Week 5 Session 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
