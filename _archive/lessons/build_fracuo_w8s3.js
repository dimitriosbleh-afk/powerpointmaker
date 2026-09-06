"use strict";

// Understanding & Operating with Fractions - Week 8 (Year 6 Numeracy) - Session 3:
// Add and subtract fractions with the SAME and RELATED denominators (rename one).
// VC2M5N05. Number line jumps + fraction wall. Daily Review: Fractions & Decimals.
// Fluency: decimal division (division bracket). Variant weekToVariant(8) = 1.

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

const SESSION = 3;
const TOTAL = 4;
const WEEK = 8;
const UNIT_TITLE = "Understanding & Operating with Fractions";
const FOOTER = `Fractions: Understand & Operate | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracUO_W8S3_Add_Subtract";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Adding and Subtracting Fractions Practice",
  "Add and subtract fractions with the same and related denominators by renaming. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Adding and Subtracting Fractions Answer Key",
  "Worked answers for the adding and subtracting fractions practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back. We can compare and order fractions by renaming.
- Today we use that same renaming idea for a new job: adding and subtracting fractions.

DO:
- Have whiteboards, markers and the printed fraction wall ready.
- Settle the class before you click on.

TEACHER NOTES:
Week 8 Session 3 of 4. The fraction wall and the number line return as anchors. A returning student needs the printed models and the Week 8 Catch-Up Card from Session 1.

WATCH FOR:
- Students who try to add the bottom numbers. We will see why that does not work.

[Planning | Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- Today's materials are the adding and subtracting practice sheet and the printed fraction wall.
- The practice sheet has an easier start and a challenge built in.

DO:
- Print one practice sheet per student and keep the answer key.
- Have the printed fraction wall and whiteboards ready.

TEACHER NOTES:
One student practice sheet plus its answer key. Most teaching is on whiteboards with the fraction wall.

CATCH-UP NOTE:
A returner can join here. The launch re-shows why parts must be the same size before adding, and the Week 8 Catch-Up Card from Session 1 summarises the renaming method with a worked example.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Fractions and decimals, not today's adding.
- Answer each one on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and check reasoning.

TEACHER NOTES:
Spaced retrieval of converting and ordering. The equivalence item links to today's renaming.

WATCH FOR:
- Students who order decimals by place value - secure.
- Students who order by length of the number - reteach place value.

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Three-tenths is 0.3.
- In order, smallest first: 0.25, 0.4, 0.6.
- Yes, two-quarters equals one half.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The equivalence item, 2/4 = 1/2, is exactly the renaming we use today.

WATCH FOR:
- Students who self-correct - secure.
- Students who order 0.4 before 0.25 - reteach hundredths.

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
- 9.6 divided by 6 is 1.6.
- 7.2 divided by 4 is 1.8.
- 8.5 divided by 5 is 1.7.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The decimal point in the answer sits directly above the point in the dividend. Reinforce that habit.

WATCH FOR:
- Students who self-correct - secure.
- Students with whole-number answers - they ignored the point, small group focus.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Look at the wall. Here is one half, and here is one quarter.
- Can we just add them to get two of something? Look closely - the parts are different sizes.
- We cannot add halves and quarters until the pieces are the same size.
- Today we make the pieces the same size first, using renaming, and then add or subtract.

DO:
- Point to the half bar and the quarter bar so the size difference is obvious.
- Ask: what could we change the half into so the pieces match the quarters?

TEACHER NOTES:
This launch creates the need for a common denominator before any procedure. Same-size parts is the whole idea, and it re-establishes the wall anchor.

WATCH FOR:
- Students who suggest changing the half into quarters - strong start.
- Students who add the bottoms - we will fix this in the I Do.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to add and subtract fractions with the same and related denominators.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is same-denominator adding and subtracting, achievable for everyone. Second is renaming to a common denominator, the core target. Third is full add and subtract with related denominators.

WATCH FOR:
- Students who can repeat same-size parts before adding - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- One key idea today: a common denominator.
- A common denominator is when both fractions have the same bottom number, the same size parts.
- We can only add or subtract once the parts are the same size.

DO:
- Show 1/2 and 1/4, then rename 1/2 to 2/4 so both are quarters.
- Have students say same-size parts.

TEACHER NOTES:
Vocabulary after the learning intention. A common denominator means same-size pieces, ready to add or subtract.

WATCH FOR:
- Students who can say same-size parts - secure.
- Students unsure - point to the wall rows.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us add two fractions that already have the same bottom number: two-eighths plus three-eighths.
- Watch me jump on the number line. I start at two-eighths, then jump three more eighths.
- I land on five-eighths. The parts are the same size, so I just add the tops and keep the bottom.
- Two plus three is five, over eight. Two-eighths plus three-eighths is five-eighths.

DO:
- Jump along the eighths on the number line.
- Say the rule: same bottom, add the tops, keep the bottom.

TEACHER NOTES:
Start with the easy case so the rule is clear before renaming is needed.

MISCONCEPTIONS:
- Misconception: students add the bottoms too, writing 5/16.
  Why: they add every number they see.
  Impact: the answer is far too small.
  Quick correction: the bottom names the size; it does not change when we add same-size parts.

WATCH FOR:
- Students who keep the bottom - secure.
- Students who write 5/16 - re-point to the same-size parts.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a pair with different bottoms, but related: one half plus one quarter.
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
  Quick correction: make the same size first; 1/2 becomes 2/4, then add the tops.

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
- Five-sixths take away one third.
- Rename one third into sixths first, then subtract.

DO:
- Display 5/6 and 1/3 on the wall.
- Give about 90 seconds for partners.
- Listen for one third renamed to two-sixths.

TEACHER NOTES:
Subtraction with the same renaming move. The wall shows the take-away clearly.

WATCH FOR:
- Pairs who rename then subtract - secure.
- Pairs who subtract bottoms - re-point to same-size parts.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- One third is two-sixths. Five-sixths take away two-sixths is three-sixths.
- Three-sixths is the same as one half. So five-sixths minus one third is one half.

DO:
- Click to reveal.
- Show the take-away on the wall once more.

TEACHER NOTES:
Same renaming idea as addition. The bottom stays the same once the sizes match, and 3/6 simplifies to 1/2.

WATCH FOR:
- Students who self-correct - secure.
- Students still subtracting bottoms - small group before You Do.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together. Three-quarters plus one eighth.
- Eighths are related to quarters, because eight is double four.
- Rename three-quarters into eighths, then add.

DO:
- Display quarters and eighths on the wall.
- Give about 75 seconds.
- Listen for three-quarters renamed to six-eighths.

TEACHER NOTES:
Quarters and eighths is another common related pair. Rename the quarters, then add the tops.

WATCH FOR:
- Students who get seven-eighths - strong.
- Students who add bottoms - reteach with the wall.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Three-quarters is six-eighths. Six-eighths plus one-eighth is seven-eighths.
- So three-quarters plus one eighth is seven-eighths.

DO:
- Click to reveal.
- Point to the six-eighths plus one-eighth on the wall.

TEACHER NOTES:
A clean related-denominator addition. Next session we use this in real-world recipe problems.

WATCH FOR:
- Students who explain the renaming - ready for independent work.
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
- Task: Extension - add three related fractions, and solve a missing-number problem such as 3/4 = 1/2 + ?. Then simplify any answer that can be made simpler.
- Extra Notes: Push the language: rename to the same size, then operate on the tops.

WATCH FOR:
- Students who rename then operate - secure.
- Students who add or subtract bottoms - prompt back to same-size parts.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Work out five-eighths take away one quarter.
- Then work out one third plus one sixth.

DO:
- Display the prompt.
- Give about 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - rename to a common denominator, then add or subtract. 5/8 - 1/4 uses eighths (1/4 = 2/8, so 3/8). 1/3 + 1/6 uses sixths (1/3 = 2/6, so 3/6 = 1/2).

WATCH FOR:
- Students who write 3/8 and 3/6 (or 1/2) - secure.
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
  titleSlide(pres, UNIT_TITLE, "Week 8 Session 3: Adding and subtracting fractions",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Fractions & decimals
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions & decimals",
      [
        "Write 3/10 as a decimal",
        "Order smallest first: 0.4, 0.25, 0.6",
        "Is 2/4 = 1/2?",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "0.3      0.25, 0.4, 0.6      yes", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal division (division bracket)
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal division (division bracket)",
      ["9.6 ÷ 6", "7.2 ÷ 4", "8.5 ÷ 5"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "1.6        1.8        1.7", { color: C.SUCCESS });
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
    "We are learning to add and subtract fractions with the same and related denominators.",
    [
      "I can add or subtract fractions that have the same bottom number.",
      "I can rename one fraction so two fractions share a bottom number.",
      "I can add and subtract fractions with related denominators.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Common denominator",
    [
      "Common denominator: both fractions have the same bottom (same-size parts).",
      "Add or subtract only once the parts are the same size.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      wallPanel(slide, lg, "Make the parts match", [
        { denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY },
        { denom: 4, shaded: 2, label: "2/4", color: C.SECONDARY },
      ], { strip: C.SECONDARY, cardH: 1.9 });
    }
  );

  // Slide 10: I Do #1 - same denominator on the number line (custom)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Same bottom: just add the tops", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["2"] });
    s.addText("2/8 + 3/8. Start at 2/8, then jump 3 more eighths.", {
      x: 0.75, y: CONTENT_TOP + 0.1, w: 8.5, h: 0.62,
      fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const topY = CONTENT_TOP + 1.0;
    const cardH = SAFE_BOTTOM - topY;
    addCard(s, 0.5, topY, 9.0, cardH, { strip: STAGE_COLORS["2"], fill: C.WHITE });
    const lineY = topY + cardH * 0.46;
    addNumberLine(s, 1.0, lineY, 8.0,
      ["0", "1/8", "2/8", "3/8", "4/8", "5/8", "6/8", "7/8", "1"], [2, 5], { labelFontSize: 13 });
    s.addText("Same size parts -> 2/8 + 3/8 = 5/8 (add tops, keep the bottom).", {
      x: 1.0, y: lineY + 0.6, w: 8.0, h: 0.5,
      fontSize: 15, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO1);
  })();

  // Slide 11: I Do #2 - rename then add (wall)
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

  // Slides 14-15: We Do #1 + reveal - 5/6 - 1/3
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Subtract: 5/6 - 1/3",
      [
        "With your partner.",
        "",
        "1.  Rename 1/3 into sixths.",
        "2.  Subtract from 5/6.",
        "3.  Simplify if you can.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Rename 1/3 to 2/6, then take away", [
          { denom: 6, shaded: 5, label: "5/6", color: C.SECONDARY },
          { denom: 3, shaded: 1, label: "1/3", color: C.PRIMARY },
          { denom: 6, shaded: 3, label: "3/6", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "5/6 - 1/3 = 5/6 - 2/6 = 3/6 = 1/2", { color: C.SUCCESS });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - 3/4 + 1/8
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Add: 3/4 + 1/8",
      [
        "With your partner.",
        "",
        "1.  Eight is double four.",
        "2.  Rename 3/4 into eighths.",
        "3.  Add the tops.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        wallPanel(slide, lg, "Rename 3/4 to 6/8, then add", [
          { denom: 4, shaded: 3, label: "3/4", color: C.SECONDARY },
          { denom: 8, shaded: 6, label: "6/8", color: C.PRIMARY },
          { denom: 8, shaded: 7, label: "7/8", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 2.6 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "3/4 + 1/8 = 6/8 + 1/8 = 7/8", { color: C.SUCCESS });
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
      "Work out 5/8 - 1/4.",
      "Work out 1/3 + 1/6.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why do fractions need the same bottom number before we add them?",
      scItems: [
        "I can add or subtract fractions that have the same bottom number.",
        "I can rename one fraction so two fractions share a bottom number.",
        "I can add and subtract fractions with related denominators.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracUO_W8S3_Add_Subtract.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Add and subtract fractions with the same and related denominators.",
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
    y = addWriteLine(doc, "a)  2/8 + 3/8 = 5/8        b)  2/5 + 1/5 = ______        c)  5/6 - 1/6 = ______", y);
    y = addWriteLine(doc, "d)  3/10 + 4/10 = ______      e)  7/8 - 3/8 = ______", y);

    y = addSectionHeading(doc, "Section 2 - Rename one fraction, then add", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2 + 1/4 = ______      b)  1/3 + 1/6 = ______      c)  3/4 + 1/8 = ______", y);
    y = addWriteLine(doc, "d)  1/2 + 3/10 = ______      e)  2/3 + 1/6 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Rename one fraction, then subtract", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3/4 - 1/2 = ______      b)  5/6 - 1/3 = ______      c)  7/8 - 1/4 = ______", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Fill the gap: 3/4 = 1/2 + ____. Then add three fractions: 1/4 + 1/8 + 1/2.", y);
    y = addWriteLine(doc, "3/4 = 1/2 + ______          1/4 + 1/8 + 1/2 = ______", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Adding and Subtracting Fractions | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the adding and subtracting fractions practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Same bottom number", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  5/8      b)  3/5      c)  4/6 (= 2/3)      d)  7/10      e)  4/8 (= 1/2)", y);

    y = addSectionHeading(doc, "Section 2 - Rename, then add", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3/4      b)  3/6 (= 1/2)      c)  7/8      d)  8/10 (= 4/5)      e)  5/6", y);

    y = addSectionHeading(doc, "Section 3 - Rename, then subtract", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/4      b)  3/6 (= 1/2)      c)  5/8", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "3/4 = 1/2 + 1/4 (since 1/2 = 2/4). 1/4 + 1/8 + 1/2 = 2/8 + 1/8 + 4/8 = 7/8.", y);

    y = addTipBox(doc,
      "Watch for: students who add or subtract the bottom numbers; students who forget to rename; students who do not simplify (3/6 = 1/2).",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Week 8 Session 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
