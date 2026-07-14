"use strict";

// Fractions & Decimals - Week 5 (Year 6 Numeracy) - Session 4: Lowest common denominator & real-world problems
// VC2M6N05. Find the LCD using multiples; add and subtract fractions in real-world contexts.
// Daily Review: Advanced number concepts (factors, primes, multiples). Fluency: decimal addition.
// Anchor models: fraction wall + number line. Variant weekToVariant(5). Ships the Year 8 extension for Week 5.

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

const SESSION = 4;
const TOTAL = 4;
const WEEK = 5;
const UNIT_TITLE = "Fractions & Decimals";
const FOOTER = `Fractions & Decimals | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracDec_W5S4_LCD_Problems";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lowest Common Denominator Practice",
  "Find the LCD, rename both fractions, and solve real-world problems. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lowest Common Denominator Answer Key",
  "Worked answers for the lowest common denominator practice sheet.");
const Y8_RES = makeSessionResource(SESSION,
  "Year 8 Extension Challenge",
  "Beyond Year 6 - unlike denominators, mixed-number operations, multiply and divide fractions, and reasoning.");
const Y8_KEY_RES = makeSessionResource(SESSION,
  "Year 8 Extension Answer Key",
  "Worked answers and look-fors for the Year 8 extension challenge.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, Y8_RES, Y8_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back, and the last session of Week 5.
- We can rename fractions to add and subtract with related bottoms. Today we handle any two bottoms using multiples, in real recipe and distance problems.

DO:
- Have whiteboards, markers and the printed fraction wall ready.
- Settle the class before you click on.

TEACHER NOTES:
Session 4 of 4. Pulls the week together. A returning student needs the printed wall and the Week 5 Catch-Up Card.

WATCH FOR:
- Students who freeze when bottoms are unrelated - we will give them a reliable method with multiples.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- Today there is a practice sheet, plus a Year 8 Extension Challenge for anyone who wants to push further.
- The practice sheet has an easier start and a challenge built in.

DO:
- Print one practice sheet per student and keep the answer key.
- Print a few Year 8 Extension Challenge sheets for early finishers.

TEACHER NOTES:
Two student sheets with answer keys. The Year 8 sheet is optional and aimed beyond Year 6 - unlike denominators, mixed-number operations and multiply or divide.

CATCH-UP NOTE:
A returner can still join. The launch re-shows why bottoms must match, and the Week 5 Catch-Up Card lists the LCD method with a worked example.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Factors, primes and multiples, not today's fractions.
- These ideas will actually help us today, so they are a good warm-up.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and listen for factor and multiple language.

TEACHER NOTES:
Daily Review here doubles as a bridge: multiples are exactly what we use to find a common denominator.

WATCH FOR:
- Students who list factors in pairs - secure.
- Students who confuse factors and multiples - clarify: factors divide in, multiples count up.

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- The factors of 12 are 1, 2, 3, 4, 6 and 12.
- 17 is prime - its only factors are 1 and 17.
- The first four multiples of 6 are 6, 12, 18 and 24.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Note the link to today: multiples of the two bottom numbers give us the common denominator.

WATCH FOR:
- Students who self-correct - secure.
- Students who call 17 composite - remind them a prime has exactly two factors.

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Decimal addition, vertical, last time this week.
- Line up the decimal points, add, bring the point down.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up points.

TEACHER NOTES:
Decimal addition algorithm. Next week the fluency focus becomes subtraction.

WATCH FOR:
- Students who align the points - secure.
- Students who align the last digit - coach the decimal point.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 8.4 plus 7.6 is 16.
- 5.55 plus 2.9 is 8.45.
- 12.8 plus 0.45 is 13.25.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
8.4 plus 7.6 makes a whole number, 16.0; accept 16. The others need placeholder zeros to align.

WATCH FOR:
- Students who self-correct - secure.
- Students who drop the trailing zero meaning - 16.0 is 16.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- A recipe needs one half cup of milk and one third cup of water. How much liquid in total?
- Look at the cups. Halves and thirds are different sizes, and neither bottom is a multiple of the other.
- We cannot rename just one this time. We need a size that both halves and thirds fit into.
- Today we use multiples to find that common size, the lowest common denominator.

DO:
- Hold up or point to the two different cup amounts.
- Ask: what size of piece could both a half and a third be cut into?

TEACHER NOTES:
This launch motivates the LCD: with unrelated bottoms we rename BOTH fractions to a shared size found from multiples.

WATCH FOR:
- Students who suggest sixths - strong start.
- Students who try to add halves and thirds directly - that is the need we are meeting.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to find a lowest common denominator to add and subtract fractions in real problems.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is listing multiples to find the LCD. Second is renaming both and operating, the core target. Third is applying it to a worded problem.

WATCH FOR:
- Students who can say lowest common denominator - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- One key idea today.
- A multiple is what you get when you count up in a number: multiples of 4 are 4, 8, 12.
- The lowest common denominator is the smallest number that both bottoms divide into - the first multiple they share.

DO:
- List a couple of multiples of two numbers and circle the first shared one.
- Have students say lowest common denominator.

TEACHER NOTES:
Vocabulary after the learning intention. The LCD is just the first shared multiple of the two bottoms.

WATCH FOR:
- Students who find the first shared multiple - secure.
- Students who pick any common multiple - that works too, but lowest keeps numbers small.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us add one half and one third.
- The bottoms are 2 and 3. I list multiples: 2, 4, 6 and 3, 6. The first shared one is 6.
- So 6 is my common denominator. I rename both: one half is three-sixths, one third is two-sixths.
- Three-sixths plus two-sixths is five-sixths. So one half plus one third is five-sixths.

DO:
- Write the two lists of multiples and ring the 6.
- Rename both fractions to sixths, then add the tops.

TEACHER NOTES:
The core move: find the LCD from multiples, rename BOTH fractions, then add the tops. The sixths wall confirms it.

MISCONCEPTIONS:
- Misconception: students rename only one fraction.
  Why: that worked last session with related bottoms.
  Impact: they cannot add because the sizes still differ.
  Quick correction: when neither bottom is a multiple of the other, rename BOTH to the LCD.

WATCH FOR:
- Students who rename both to sixths - secure.
- Students who change only one - point to the shared size.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a real subtraction problem. A bowl has three-quarters of a cup of flour. We use one third of a cup.
- Bottoms are 4 and 3. Multiples of 4: 4, 8, 12. Multiples of 3: 3, 6, 9, 12. The LCD is 12.
- Rename: three-quarters is nine-twelfths, one third is four-twelfths.
- Nine-twelfths take away four-twelfths is five-twelfths. Five-twelfths of a cup of flour is left.

DO:
- Show the two multiple lists and ring 12.
- Rename both to twelfths and subtract the tops.

TEACHER NOTES:
Same LCD method in a subtraction word problem. Keep the cup context so the answer means something.

MISCONCEPTIONS:
- Misconception: students subtract the bottoms, or the tops without renaming.
  Why: they skip the common-size step.
  Impact: a wrong amount of flour.
  Quick correction: rename both to twelfths first, then subtract the tops.

WATCH FOR:
- Students who get five-twelfths and say the unit - secure.
- Students who subtract bottoms - reteach the LCD step.

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Find the lowest common denominator of one quarter and one sixth.
- List multiples of 4 and of 6, and find the first shared one.

DO:
- Display the prompt.
- Give about 45 seconds.
- Walk and scan for the multiple lists.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me the LCD on three, two, one, show.
- Scan for: 12, with multiples of 4 (4, 8, 12) and 6 (6, 12).
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students multiply 4 by 6 to get 24, or pick 6 because it is bigger.
- Reteach: list the multiples of each and ring the FIRST shared one. 12 comes before 24.
- Re-check: what is the first multiple that 4 and 6 share?
- Use the lists to confirm.

TEACHER NOTES:
Multiplying the bottoms always gives a common denominator, but not always the lowest. Lowest keeps the numbers manageable.

WATCH FOR:
- Students who find 12 - secure.
- Students who write 24 - acceptable as a common denominator, but coach toward the lowest.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Two-thirds plus one quarter.
- Find the LCD from multiples of 3 and 4, rename both, then add.

DO:
- Display 2/3 and 1/4.
- Give about 90 seconds for partners.
- Listen for an LCD of 12.

TEACHER NOTES:
Unrelated bottoms, so both must be renamed. Multiples of 3 and 4 share 12.

WATCH FOR:
- Pairs who rename to twelfths and add - secure.
- Pairs who rename only one - point to the shared size.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- The LCD of 3 and 4 is 12. Two-thirds is eight-twelfths, one quarter is three-twelfths.
- Eight-twelfths plus three-twelfths is eleven-twelfths.

DO:
- Click to reveal.
- Show both renamings to twelfths.

TEACHER NOTES:
Eleven-twelfths does not simplify. A clean check of the full method.

WATCH FOR:
- Students who self-correct - secure.
- Students who rename only one - small group before You Do.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, a word problem.
- Sam walks one half of a kilometre, then two-fifths of a kilometre. How far in total?
- Find the LCD of 2 and 5, rename both, and add.

DO:
- Display the problem.
- Give about 90 seconds.
- Listen for an LCD of 10.

TEACHER NOTES:
A distance word problem. Bottoms 2 and 5 share 10. Keep the kilometre unit in the answer.

WATCH FOR:
- Pairs who get nine-tenths of a kilometre - secure.
- Pairs who add tops and bottoms - reteach the LCD step.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- The LCD of 2 and 5 is 10. One half is five-tenths, two-fifths is four-tenths.
- Five-tenths plus four-tenths is nine-tenths. Sam walked nine-tenths of a kilometre.

DO:
- Click to reveal.
- Restate the answer with its unit.

TEACHER NOTES:
Real-world answers need a unit. That is the difference between a number and an answer.

WATCH FOR:
- Students who include the kilometre - secure.
- Students who drop the unit - prompt for what the number describes.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- The first section finds the lowest common denominator.
- The next section renames both fractions and adds or subtracts.
- The last section is real-world word problems.
- If you finish, try the extension, or the Year 8 Extension Challenge.

DO:
- Distribute the practice sheet.
- Circulate and listen for find the LCD, rename both, then operate.
- Cold call one or two students to explain a word problem answer.

TEACHER NOTES:
Different numbers from the We Do, same method. Word problems need the unit in the answer.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed multiples chart and do Section 1 and the first row of Section 2. The first one is started for you.
- Extra Notes: Sit with these students and list the multiples for the first problem together.
EXTENDING PROMPT:
- Task: Extension - add three fractions with different bottoms, then try the Year 8 Extension Challenge with unlike denominators and mixed numbers.
- Extra Notes: Push the reasoning: which shared multiple keeps the numbers smallest?

WATCH FOR:
- Students who find the LCD and rename both - secure.
- Students who rename only one - prompt back to the shared size.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Work out one half plus one third.
- Then solve: a jug holds three-quarters of a litre. You pour out one sixth of a litre. How much is left?

DO:
- Display the prompt.
- Give about 4 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - find the LCD, rename both, operate, and answer a real problem with its unit. 1/2 + 1/3 uses sixths; 3/4 - 1/6 uses twelfths.

WATCH FOR:
- Students who write 5/6 and 7/12 of a litre - secure.
- Students who rename only one - revisit at the start of Week 6.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again, and think back across the whole week.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: how do you find a common size for two fractions with different bottoms?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is the LCD: list multiples, find the first shared one, rename both fractions, then add or subtract. This sets up Week 6 on percentages and decimals.

WATCH FOR:
- Strong thumbs across all three - ready for Week 6.
- Sideways or down on the core criterion - a short revision before Week 6 Session 1.

[Retention and Recall | Mastery and Application]`;

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
  titleSlide(pres, UNIT_TITLE, "Week 5 Session 4: Lowest common denominator and real problems",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Factors, primes, multiples
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Number concepts",
      [
        "List all the factors of 12",
        "Is 17 prime or composite?",
        "Write the first four multiples of 6",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "1, 2, 3, 4, 6, 12     prime     6, 12, 18, 24", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal addition",
      ["8.4 + 7.6", "5.55 + 2.9", "12.8 + 0.45"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "16        8.45        13.25", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - recipe with unrelated bottoms
  contentSlide(pres, "Launch", C.ACCENT, "Half a cup and a third of a cup",
    [
      "A recipe: 1/2 cup milk + 1/3 cup water.",
      "",
      "Different sizes, and 2 and 3 are unrelated.",
      "We need a size BOTH fit into.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Halves and thirds differ", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 3, shaded: 1, label: "1/3", color: C.ACCENT },
      ], { strip: C.ACCENT, cardH: 2.1 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to find a lowest common denominator to add and subtract fractions in real problems.",
    [
      "I can list multiples to find a common denominator for two fractions.",
      "I can rename both fractions and add or subtract them.",
      "I can solve a real-world fraction problem by adding or subtracting.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Multiple & lowest common denominator",
    [
      "Multiple: count up in a number (multiples of 4: 4, 8, 12).",
      "Lowest common denominator: the smallest number both bottoms divide into.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
      slide.addText("Find the LCD of 2 and 3", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [
        ["2:", "2,  4,  6"],
        ["3:", "3,  6"],
      ];
      const ry0 = lg.panelTopPadded + 0.55;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.55;
        slide.addText(r[0], {
          x: lg.rightX + 0.3, y: ry, w: 0.7, h: 0.45,
          fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "left", valign: "middle", margin: 0,
        });
        slide.addText(r[1], {
          x: lg.rightX + 1.0, y: ry, w: lg.rightW - 1.2, h: 0.45,
          fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL,
          align: "left", valign: "middle", margin: 0,
        });
      });
      addTextOnShape(slide, "First shared multiple = 6 = LCD", {
        x: lg.rightX + 0.25, y: ry0 + 1.2, w: lg.rightW - 0.5, h: 0.5, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10: I Do #1 - 1/2 + 1/3 with multiples
  workedExSlide(pres, 2, "I Do", "Add 1/2 + 1/3 using multiples",
    [
      "Bottoms are 2 and 3.",
      "Multiples of 2: 2, 4, 6",
      "Multiples of 3: 3, 6",
      "First shared = 6 (the LCD).",
      "",
      "1/2 = 3/6,  1/3 = 2/6",
      "3/6 + 2/6 = 5/6",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Rename both to sixths", [
        { denom: 6, shaded: 3, label: "3/6", color: C.PRIMARY },
        { denom: 6, shaded: 2, label: "2/6", color: C.SECONDARY },
        { denom: 6, shaded: 5, label: "5/6", color: C.ACCENT },
      ], { strip: C.PRIMARY, cardH: 2.6 });
    }
  );

  // Slide 11: I Do #2 - real subtraction problem with LCD 12
  workedExSlide(pres, 2, "I Do", "Flour problem: 3/4 - 1/3",
    [
      "3/4 cup of flour, use 1/3 cup.",
      "Bottoms 4 and 3 -> LCD 12.",
      "",
      "3/4 = 9/12,  1/3 = 4/12",
      "9/12 - 4/12 = 5/12",
      "",
      "5/12 of a cup of flour is left.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Rename both to twelfths", [
        { denom: 12, shaded: 9, label: "9/12", color: C.PRIMARY },
        { denom: 12, shaded: 4, label: "4/12", color: C.SECONDARY },
        { denom: 12, shaded: 5, label: "5/12", color: C.ACCENT },
      ], { strip: C.PRIMARY, cardH: 2.6 });
    }
  );

  // Slides 12-13: CFU + reveal - LCD of 1/4 and 1/6
  withReveal(
    () => cfuSlide(pres, "CFU", "Find the lowest common denominator", "Show Me Boards",
      "Find the lowest common denominator of 1/4 and 1/6.\n\nList multiples of 4 and of 6 to help.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "LCD = 12   (4: 4, 8, 12   6: 6, 12)",
        { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 2/3 + 1/4
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Add 2/3 + 1/4",
      [
        "With your partner.",
        "",
        "1.  LCD of 3 and 4?",
        "2.  Rename both fractions.",
        "3.  Add the tops.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Rename both to twelfths", [
          { denom: 12, shaded: 8, label: "8/12", color: C.SECONDARY },
          { denom: 12, shaded: 3, label: "3/12", color: C.PRIMARY },
          { denom: 12, shaded: 11, label: "11/12", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "LCD 12:  2/3 = 8/12,  1/4 = 3/12  ->  11/12", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - distance word problem
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Word problem: how far did Sam walk?", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 1.5, { strip: STAGE_COLORS["3"], fill: C.WHITE });
      s.addText("Sam walks 1/2 km, then 2/5 km. How far in total?", {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 8.5, h: 0.6,
        fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL, bold: true, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      s.addText("Find the LCD of 2 and 5, rename both, then add. Remember the unit.", {
        x: 0.75, y: CONTENT_TOP + 0.8, w: 8.5, h: 0.5,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      });

      const panelY = CONTENT_TOP + 1.7;
      addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.PRIMARY });
      addNumberLine(s, 1.5, panelY + 0.9, 7.0,
        ["0", "", "", "", "", "1/2", "", "", "", "", "1"], [], { labelFontSize: 13 });
      s.addText("Tip: tenths fit both halves and fifths.", {
        x: 0.7, y: panelY + 0.18, w: 8.6, h: 0.32,
        fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "LCD 10:  1/2 = 5/10,  2/5 = 4/10  ->  9/10 km", {
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
      { text: "find the lowest common denominator.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "rename both, then add or subtract.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "solve the word problems.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "List multiples to find the LCD. Rename BOTH fractions. Then add or subtract the tops. Word problems need a unit.", {
      x: 0.8, y: panelY + 0.52, w: 8.4, h: 0.6, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFractionStripSet(s, 2.5, panelY + 1.3, 5.0, 0.85, [
      { denom: 6, shaded: 3, label: "3/6", color: C.PRIMARY },
      { denom: 6, shaded: 2, label: "2/6", color: C.ACCENT },
    ], { labelW: 0.6, labelFontSize: 12 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Work out 1/2 + 1/3.",
      "A jug holds 3/4 L. You pour out 1/6 L. How much is left?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how do you find a common size for two fractions with different bottoms?",
      scItems: [
        "I can list multiples to find a common denominator for two fractions.",
        "I can rename both fractions and add or subtract them.",
        "I can solve a real-world fraction problem by adding or subtracting.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDec_W5S4_LCD_Problems.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find the LCD, rename both fractions, and solve real-world problems.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "List the multiples of each bottom number. The first shared multiple is the lowest common denominator (LCD). Rename BOTH fractions to the LCD, then add or subtract the tops. Word problems need a unit in the answer.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "1/2 + 1/3: multiples of 2 are 2, 4, 6; of 3 are 3, 6. LCD = 6. 1/2 = 3/6 and 1/3 = 2/6, so 3/6 + 2/6 = 5/6.",
      y);

    y = addSectionHeading(doc, "Section 1 - Find the lowest common denominator", y, { color: C.PRIMARY });
    y = addBodyText(doc, "The first one is started for you.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  1/2 and 1/3  -> LCD 6      b)  1/4 and 1/6 -> LCD ____      c)  1/3 and 1/5 -> LCD ____", y);
    y = addWriteLine(doc, "d)  1/2 and 2/5  -> LCD ____    e)  3/4 and 1/3 -> LCD ____", y);

    y = addSectionHeading(doc, "Section 2 - Rename both, then add or subtract", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2 + 1/3 = ______      b)  2/3 + 1/4 = ______      c)  1/2 - 1/5 = ______", y);
    y = addWriteLine(doc, "d)  3/4 - 1/3 = ______      e)  2/5 + 1/2 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Real-world problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Show your working and write the unit in your answer.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  A recipe uses 1/2 cup milk and 1/3 cup water. How much liquid? ______", y);
    y = addWriteLine(doc, "b)  A bottle holds 3/4 L. You drink 1/3 L. How much is left? ______", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Add three fractions: 1/2 + 1/3 + 1/4. Find a denominator all three fit into.", y);
    y = addWriteLine(doc, "1/2 + 1/3 + 1/4 = ______", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Lowest Common Denominator | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the lowest common denominator practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Find the LCD", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  6      b)  12      c)  15      d)  10      e)  12", y);

    y = addSectionHeading(doc, "Section 2 - Rename, then operate", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  5/6      b)  11/12      c)  3/10      d)  5/12      e)  9/10", y);

    y = addSectionHeading(doc, "Section 3 - Real-world", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  5/6 of a cup (3/6 + 2/6).   b)  5/12 of a litre (9/12 - 4/12).", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "1/2 + 1/3 + 1/4: LCD 12. 6/12 + 4/12 + 3/12 = 13/12 = 1 1/12.", y);

    y = addTipBox(doc,
      "Watch for: students who rename only one fraction; students who multiply the bottoms instead of finding the lowest; students who drop the unit in word problems.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension Challenge (Week 5)
  await (async () => {
    const doc = createPdf({ title: Y8_RES.name });
    let y = addPdfHeader(doc, Y8_RES.name, {
      subtitle: "Week 5 challenge - aimed beyond Year 6, for students ready to push further.",
      color: C.ASSESS,
      lessonInfo: "Week 5 | Fractions | Year 8 reach task",
    });
    y = addTipBox(doc,
      "These problems go past Year 6: unlike denominators, mixed-number addition and subtraction, multiplying and dividing fractions, and reasoning. Show all working.",
      y, { color: C.ASSESS });

    y = addSectionHeading(doc, "Part A - Unlike denominators", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2/3 + 3/5 = ______        b)  5/6 - 2/9 = ______        c)  3/4 + 5/8 = ______", y);
    y = addWriteLine(doc, "d)  7/10 - 2/15 = ______", y);

    y = addSectionHeading(doc, "Part B - Mixed numbers", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2 1/2 + 1 3/4 = ______        b)  4 1/3 - 1 5/6 = ______", y);

    y = addSectionHeading(doc, "Part C - Multiply and divide fractions", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2/3 x 3/4 = ______        b)  3/5 of 20 = ______        c)  1/2 divided by 1/4 = ______", y);

    y = addSectionHeading(doc, "Part D - Reasoning", y, { color: C.ACCENT });
    y = addBodyText(doc, "A jug is 2/3 full. After pouring out 1/4 of a litre, it is 5/12 full. How many litres does the full jug hold? Explain your reasoning.", y);
    y = addLinedAreaSafe(doc, y, 3);

    addPdfFooter(doc, "Week 5 | Year 8 Extension Challenge | Fractions");
    await writePdf(doc, path.join(OUT_DIR, Y8_RES.fileName));
    console.log("PDF written: " + Y8_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: Y8_KEY_RES.name });
    let y = addPdfHeader(doc, Y8_KEY_RES.name, {
      subtitle: "Worked answers and look-fors for the Week 5 Year 8 extension challenge.",
      color: C.ASSESS,
      lessonInfo: "Week 5 | Fractions | Year 8 reach task",
    });

    y = addSectionHeading(doc, "Part A - Unlike denominators", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  10/15 + 9/15 = 19/15 = 1 4/15.   b)  15/18 - 4/18 = 11/18.   c)  6/8 + 5/8 = 11/8 = 1 3/8.   d)  21/30 - 4/30 = 17/30.", y);

    y = addSectionHeading(doc, "Part B - Mixed numbers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2 1/2 + 1 3/4 = 4 1/4 (2/4 + 3/4 = 5/4 = 1 1/4).   b)  4 1/3 - 1 5/6 = 2 1/2 (rename to sixths: 4 2/6 - 1 5/6 = 2 3/6).", y);

    y = addSectionHeading(doc, "Part C - Multiply and divide", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2/3 x 3/4 = 6/12 = 1/2.   b)  3/5 of 20 = 12.   c)  1/2 divided by 1/4 = 1/2 x 4/1 = 2.", y);

    y = addSectionHeading(doc, "Part D - Reasoning", y, { color: C.ACCENT });
    y = addBodyText(doc, "2/3 - 5/12 = 8/12 - 5/12 = 3/12 = 1/4 of the jug equals 1/4 litre poured. So 1/4 of the jug = 1/4 L, which means the full jug holds 1 litre. Look for: equivalent fractions, correct subtraction, and the link from a fraction of the jug to litres.", y);

    addPdfFooter(doc, "Week 5 | Year 8 Extension Answer Key | Fractions");
    await writePdf(doc, path.join(OUT_DIR, Y8_KEY_RES.fileName));
    console.log("PDF written: " + Y8_KEY_RES.fileName);
  })();

  console.log("Week 5 Session 4 build complete.");
}

// Safe wrapper so a missing addLinedArea export does not break the build.
function addLinedAreaSafe(doc, y, lines) {
  try {
    const helpers = require("../themes/pdf_helpers");
    if (typeof helpers.addLinedArea === "function") {
      return helpers.addLinedArea(doc, y + 4, lines, {});
    }
  } catch (e) { /* fall through */ }
  return y;
}

build().catch((err) => { console.error(err); process.exit(1); });
