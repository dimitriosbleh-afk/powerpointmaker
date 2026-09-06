"use strict";

// Fractions & Decimals - Week 5 (Year 6 Numeracy) - Session 2: Mixed numerals & improper fractions
// VC2M5N03. Convert between mixed numerals and improper fractions and locate them on a number line.
// Daily Review: Decimal operations & place value. Fluency: decimal addition algorithm.
// Anchor models: number line (primary today) + fraction strips. Variant weekToVariant(5).

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

const SESSION = 2;
const TOTAL = 4;
const WEEK = 5;
const UNIT_TITLE = "Fractions & Decimals";
const FOOTER = `Fractions & Decimals | Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FracDec_W5S2_Mixed_And_Improper";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Mixed and Improper Fractions Practice",
  "Convert both ways and place fractions on a number line. Built-in enabling start and extension.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Mixed and Improper Fractions Answer Key",
  "Worked answers for the mixed and improper fractions practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back. Yesterday we found equivalent fractions on the wall and the number line.
- Today we use the number line again to read and write fractions that are bigger than one whole.

DO:
- Have whiteboards, markers and the printed number line ready.
- Settle the class before you click on.

TEACHER NOTES:
Session 2 of 4. The number line is the anchor today. A returning student only needs the printed number line and the Week 5 Catch-Up Card.

WATCH FOR:
- Students who think a fraction can never be bigger than one - reassure, we will see exactly how it can.

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
A returner can pick this lesson up cold. The launch re-builds the number line idea, and the Week 5 Catch-Up Card from Session 1 summarises mixed and improper fractions with a worked example.

[Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Back to decimals and place value, not today's fractions.
- Read each one carefully and answer on your whiteboard.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Walk and check place-value language.

TEACHER NOTES:
Daily Review keeps decimal place value fresh from last unit. It is prior learning, not today's content.

WATCH FOR:
- Students who name the place correctly - secure.
- Students who say point zero seven - prompt for the place name, hundredths.

[Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- In 4.07 the 7 is worth 7 hundredths, or 0.07.
- 0.6 plus 0.25 is 0.85.
- 3.48 rounded to one decimal place is 3.5, because the 8 rounds the 4 up to 5.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The rounding item is the discriminator. Students must look at the next digit to decide.

WATCH FOR:
- Students who round using the next digit - secure.
- Students who round 3.48 to 3.4 - they ignored the 8, reteach the rounding digit.

[Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Decimal addition, vertical.
- Line up the decimal points, add, and bring the point straight down.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up points.

TEACHER NOTES:
Same fluency focus all week - the decimal addition algorithm. Keep it brisk.

WATCH FOR:
- Students who line up points neatly - secure.
- Students who line up the last digit instead - coach the decimal point.

[Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 5.4 plus 2.8 is 8.2.
- 3.65 plus 4.5 is 8.15.
- 7.2 plus 0.85 is 8.05.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
3.65 plus 4.5 and 7.2 plus 0.85 both need a placeholder zero to line up. Coach filling the empty place.

WATCH FOR:
- Students who self-correct - secure.
- Students who misalign 0.85 - small group focus.

[Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Look at these bars. The first bar is one whole, cut into quarters and all four shaded.
- The second bar has three more quarters shaded.
- Altogether that is seven quarters, which we write as seven over four.
- But it is also one whole and three-quarters. The same amount, two names. That is today's learning.

DO:
- Point to the four shaded quarters, then the three more.
- Take a count of the quarters out loud with the class.

TEACHER NOTES:
This launch shows a fraction bigger than one before any rule. Seven-quarters and one-and-three-quarters are the same amount.

WATCH FOR:
- Students who count seven quarters - strong start.
- Students who think you cannot have more than four quarters - show the second bar.

[Knowledge and Memory | Attention, Focus and Regulation]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to convert between mixed numerals and improper fractions and place them on a number line.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
First criterion is reading the two forms. Second is improper to mixed, the core target. Third stretches to mixed to improper and locating on a line.

WATCH FOR:
- Students who can repeat the two words - tracking.

[Planning]`;

const NOTES_VOCAB = `SAY:
- Two words for today.
- A mixed numeral is a whole number and a fraction together, like one and three-quarters.
- An improper fraction has a top number bigger than the bottom, like seven-quarters. It is more than one whole.

DO:
- Point to the whole-number part and the fraction part of the mixed numeral.
- Have students say both words once.

TEACHER NOTES:
Vocabulary after the learning intention. Improper is not wrong - it just means top heavier than bottom.

WATCH FOR:
- Students who can tell which is which - secure.
- Students who think improper means a mistake - clarify the meaning.

[Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us change an improper fraction into a mixed number on the number line.
- Seven-quarters means seven jumps of one-quarter from zero. Watch me count them.
- Four quarters bring me to one whole. That uses up four of my seven.
- Seven take away four leaves three quarters past one.
- So seven-quarters lands at one and three-quarters. Seven over four equals one and three-quarters.

DO:
- Count the seven quarter-jumps along the line.
- Circle where one whole is, then count the leftover quarters.
- Say the result twice.

TEACHER NOTES:
This is the core move: how many whole groups fit, then what is left over. The number line makes the leftover visible.

MISCONCEPTIONS:
- Misconception: students read 7/4 as 7 and 4, or as 4 and 3.
  Why: they are unsure which number is the whole and which is the leftover.
  Impact: they write the wrong mixed number.
  Quick correction: count groups of four quarters first; the leftover is the fraction part.

WATCH FOR:
- Students who land on one and three-quarters - secure.
- Students who stop at one whole - remind them three quarters remain.

[Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now the other direction, a mixed number into an improper fraction, using a quick rule.
- Two and one-third. First, how many thirds are in two wholes? Two times three is six thirds.
- Add the extra one third. Six plus one is seven thirds.
- So two and one-third equals seven over three.
- The rule is: whole times the bottom, then add the top, and keep the same bottom.

DO:
- Show the two whole bars as three-thirds each, then the extra third.
- Write the rule as you say it: whole x bottom + top, over bottom.

TEACHER NOTES:
Connect the rule to the bars. Two wholes really are six thirds, so the rule is just counting all the thirds.

MISCONCEPTIONS:
- Misconception: students multiply the whole by the top, or forget to add the top.
  Why: they half-remember the rule.
  Impact: they get the wrong number of parts.
  Quick correction: whole times the BOTTOM, then add the top.

WATCH FOR:
- Students who get seven-thirds and can explain it - secure.
- Students who write six-thirds - they forgot to add the extra third.

[Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Change nine-quarters into a mixed number.
- Think: how many whole groups of four quarters, and what is left over?

DO:
- Display the prompt.
- Give about 45 seconds.
- Walk and scan the boards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me your mixed number on three, two, one, show.
- Scan for: 2 and 1/4, with reasoning that eight quarters make two wholes.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students write 2 and 1, or 1 and 1/4.
- Reteach: count groups of four quarters on the number line. Eight quarters is two wholes, one quarter left.
- Re-check: how many whole groups of four are in nine quarters?
- Use the number line to confirm.

TEACHER NOTES:
The leftover is the trap. Students who do not track the whole groups lose the fraction part.

WATCH FOR:
- Students who write two and one-quarter - secure.
- Students who write the leftover as a whole - back to the number line.

[Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me on the number line.
- This point is at eleven-thirds. Count along in thirds with me to find it.
- What mixed number is the same as eleven-thirds? Whisper your idea to your partner first.

DO:
- Display the marked point at eleven-thirds.
- Give about 90 seconds for partners.
- Listen for three wholes and two-thirds.

TEACHER NOTES:
Same structure as the I Do, new fraction. Students count thirds and track the whole groups.

WATCH FOR:
- Pairs who count nine thirds as three wholes, two left - secure.
- Pairs who miscount the thirds - count together slowly.

[Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Nine thirds make three wholes. Eleven take away nine leaves two thirds.
- So eleven-thirds equals three and two-thirds.

DO:
- Click to reveal.
- Count the whole groups once more together.

TEACHER NOTES:
The reveal restates groups-and-leftover. Same idea as nine-quarters in the check.

WATCH FOR:
- Students who self-correct - secure.
- Students who still miscount - small group before You Do.

[Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, the other direction.
- Change three and two-fifths into an improper fraction.
- Use the rule on your whiteboard: whole times the bottom, then add the top.

DO:
- Display three and two-fifths above the strips.
- Give about 75 seconds.
- Listen for three times five, plus two.

TEACHER NOTES:
Mixed to improper with the rule. The strips show three wholes as fifteen fifths plus two more.

WATCH FOR:
- Students who write seventeen-fifths - secure.
- Students who write six-fifths - they multiplied by the top, reteach the rule.

[Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Three times five is fifteen fifths. Add the two: seventeen fifths.
- So three and two-fifths equals seventeen over five.

DO:
- Click to reveal.
- Point to the fifteen fifths plus two on the strips.

TEACHER NOTES:
Strips and rule agree. Next session we add and subtract fractions, so secure these conversions today.

WATCH FOR:
- Students who connect strips to the rule - ready for independent work.
- Students unsure - one more example before You Do.

[Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- The first section changes improper fractions into mixed numbers.
- The next section changes mixed numbers into improper fractions.
- The last section places fractions on a number line.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for groups-and-leftover, and whole-times-bottom-plus-top.
- Cold call one or two students to explain a conversion.

TEACHER NOTES:
Different numbers from the We Do, same two strategies. The number line keeps the meaning visible.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed number line and do Section 1 only. The first one is started for you.
- Extra Notes: Sit with these students and count the first conversion together.
EXTENDING PROMPT:
- Task: Extension - convert improper fractions with bigger numbers, and order a set of mixed and improper fractions on one line. Early finishers may start the Year 8 Extension Challenge.
- Extra Notes: Push the reasoning about how many whole groups fit.

WATCH FOR:
- Students who convert both ways fluently - secure.
- Students who muddle which number is the whole - prompt back to groups-and-leftover.

[Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Change seven-thirds into a mixed number.
- Then change two and three-quarters into an improper fraction.

DO:
- Display the prompt.
- Give about 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core criterion - converting both ways. Look for groups-and-leftover and whole-times-bottom-plus-top.

WATCH FOR:
- Students who write 2 and 1/3, and 11/4 - secure.
- Students who muddle the parts - revisit at the start of Session 3.

[Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: how do you change an improper fraction into a mixed number?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is groups-and-leftover one way, and whole-times-bottom-plus-top the other way, with the number line giving meaning.

WATCH FOR:
- Strong thumbs across all three - move at pace next session.
- Sideways or down on the core criterion - quick revision at the start of Session 3.

[Retention and Recall]`;

// --- Helpers -----------------------------------------------------------------

// Full-width number line on a card, with a caption underneath.
function nlCard(slide, topY, labels, marks, caption, opts) {
  const o = opts || {};
  const cardH = SAFE_BOTTOM - topY;
  addCard(slide, 0.5, topY, 9.0, cardH, { strip: o.strip || C.PRIMARY, fill: C.WHITE });
  const lineY = topY + cardH * 0.46;
  addNumberLine(slide, 1.0, lineY, 8.0, labels, marks || [], { labelFontSize: o.labelFontSize || 13 });
  if (caption) {
    slide.addText(caption, {
      x: 1.0, y: lineY + 0.6, w: 8.0, h: 0.5,
      fontSize: o.captionSize || 15, fontFace: FONT_B, color: o.captionColor || C.CHARCOAL,
      bold: Boolean(o.captionBold), italic: o.captionItalic !== false,
      align: "center", margin: 0,
    });
  }
}

// Fraction strips drawn into a right-column panel.
function stripsPanel(slide, lg, headerText, strips, opts) {
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
  titleSlide(pres, UNIT_TITLE, "Week 5 Session 2: Mixed numerals and improper fractions",
    `Year 6 Numeracy | Week ${WEEK} Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Decimal place value
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Decimals & place value",
      [
        "What is the value of the 7 in 4.07?",
        "0.6 + 0.25 = ____",
        "Round 3.48 to one decimal place",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "7 hundredths (0.07)     0.85     3.5", { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal addition",
      ["5.4 + 2.8", "3.65 + 4.5", "7.2 + 0.85"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "8.2        8.15        8.05", { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - 7/4 as bars
  contentSlide(pres, "Launch", C.ACCENT, "Can a fraction be bigger than one?",
    [
      "One whole = four-quarters.",
      "",
      "Three more quarters shaded.",
      "Seven quarters = 1 and 3/4.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      stripsPanel(slide, lg, "Seven quarters", [
        { denom: 4, shaded: 4, label: "4/4", color: C.PRIMARY },
        { denom: 4, shaded: 3, label: "3/4", color: C.ACCENT },
      ], { strip: C.ACCENT, cardH: 2.2 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to convert between mixed numerals and improper fractions and place them on a number line.",
    [
      "I can read a mixed numeral and an improper fraction.",
      "I can change an improper fraction into a mixed number.",
      "I can change a mixed number into an improper fraction and place it on a number line.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Mixed numeral & improper fraction",
    [
      "Mixed numeral = a whole number and a fraction.",
      "Improper fraction = top number bigger than the bottom.",
      "Both can name the same amount.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
      const rows = [
        ["1 3/4", "mixed numeral"],
        ["7/4", "improper fraction"],
      ];
      const ry0 = lg.panelTopPadded + 0.35;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.9;
        slide.addText(r[0], {
          x: lg.rightX + 0.25, y: ry, w: 1.5, h: 0.65,
          fontSize: 30, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        slide.addText(r[1], {
          x: lg.rightX + 1.85, y: ry, w: lg.rightW - 2.05, h: 0.65,
          fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
          align: "left", valign: "middle", margin: 0,
        });
      });
      slide.addText("Same amount, two names.", {
        x: lg.rightX + 0.2, y: ry0 + 1.85, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 10: I Do #1 - improper -> mixed on the number line (custom full width)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Improper to mixed on the number line", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["2"] });
    s.addText("7/4 = seven quarters. Count the quarter-jumps from 0.", {
      x: 0.75, y: CONTENT_TOP + 0.1, w: 8.5, h: 0.62,
      fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    nlCard(s, CONTENT_TOP + 1.0,
      ["0", "", "", "", "1", "", "", "", "2"], [7],
      "4 quarters = 1 whole. 3 quarters left over.   So 7/4 = 1 3/4.",
      { strip: STAGE_COLORS["2"], captionBold: true, captionItalic: false, captionColor: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO1);
  })();

  // Slide 11: I Do #2 - mixed -> improper with the rule
  workedExSlide(pres, 2, "I Do", "Mixed to improper with a rule",
    [
      "2 1/3 -> how many thirds?",
      "",
      "2 wholes = 2 x 3 = 6 thirds.",
      "Add the 1 third -> 7 thirds.",
      "So 2 1/3 = 7/3.",
      "",
      "Rule: (whole x bottom) + top, over bottom.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      stripsPanel(slide, lg, "2 wholes + 1 third = 7 thirds", [
        { denom: 3, shaded: 3, label: "3/3", color: C.PRIMARY },
        { denom: 3, shaded: 3, label: "3/3", color: C.PRIMARY },
        { denom: 3, shaded: 1, label: "1/3", color: C.ACCENT },
      ], { strip: C.PRIMARY, cardH: 2.7 });
    }
  );

  // Slides 12-13: CFU + reveal - 9/4 to a mixed number
  withReveal(
    () => cfuSlide(pres, "CFU", "Change 9/4 to a mixed number", "Show Me Boards",
      "Change 9/4 into a mixed number.\n\nHow many whole groups of four quarters? What is left over?",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "9/4 = 2 1/4   (8 quarters = 2 wholes, 1 quarter left)",
        { color: C.SUCCESS, label: "Answer" });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 11/3 on the number line
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "What mixed number is this point?", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["3"] });
      s.addText("This point is at 11/3. Count along in thirds. What mixed number is it?", {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 8.5, h: 0.62,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      nlCard(s, CONTENT_TOP + 1.0,
        ["0", "", "", "1", "", "", "2", "", "", "3", "", "", "4"], [11],
        "Count the thirds. How many whole groups of 3 thirds?",
        { strip: C.PRIMARY, captionItalic: true, captionColor: C.MUTED });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "11/3 = 3 2/3   (9 thirds = 3 wholes, 2 thirds left)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - 3 2/5 to improper
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Change 3 2/5 to an improper fraction",
      [
        "With your partner.",
        "",
        "1.  How many fifths in 3 wholes?",
        "2.  Add the extra 2 fifths.",
        "3.  Use the rule to check.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        stripsPanel(slide, lg, "3 wholes + 2 fifths", [
          { denom: 5, shaded: 5, label: "5/5", color: C.SECONDARY },
          { denom: 5, shaded: 5, label: "5/5", color: C.SECONDARY },
          { denom: 5, shaded: 5, label: "5/5", color: C.SECONDARY },
          { denom: 5, shaded: 2, label: "2/5", color: C.ACCENT },
        ], { strip: C.SECONDARY, cardH: 3.0 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "3 2/5 = 17/5   (3 x 5 = 15, + 2 = 17)", { color: C.SUCCESS });
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
      { text: "change improper fractions to mixed numbers.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "change mixed numbers to improper.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "place fractions on the number line.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Improper to mixed: count whole groups, then the leftover. Mixed to improper: (whole x bottom) + top.", {
      x: 0.85, y: panelY + 0.52, w: 8.3, h: 0.55, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

    addNumberLine(s, 1.5, panelY + 1.55, 7.0,
      ["0", "", "1", "", "2", "", "3"], [], { labelFontSize: 13 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Change 7/3 into a mixed number.",
      "Change 2 3/4 into an improper fraction.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how do you change an improper fraction into a mixed number?",
      scItems: [
        "I can read a mixed numeral and an improper fraction.",
        "I can change an improper fraction into a mixed number.",
        "I can change a mixed number into an improper fraction and place it on a number line.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDec_W5S2_Mixed_And_Improper.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Convert between mixed numbers and improper fractions, and place them on a number line.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Improper to mixed: count how many whole groups fit, then the leftover (7/4 = 1 3/4). Mixed to improper: (whole x bottom) + top, over the same bottom (2 1/3 = 7/3).",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "9/4: four quarters make one whole, so 8 quarters make 2 wholes. 9 - 8 = 1 quarter left. 9/4 = 2 1/4.",
      y);

    y = addSectionHeading(doc, "Section 1 - Improper fraction to mixed number", y, { color: C.PRIMARY });
    y = addBodyText(doc, "The first one is started for you.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  7/4 = 1 3/4          b)  9/2 = ______          c)  11/4 = ______", y);
    y = addWriteLine(doc, "d)  8/3 = ______          e)  17/5 = ______         f)  13/6 = ______", y);

    y = addSectionHeading(doc, "Section 2 - Mixed number to improper fraction", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2 1/3 = ______        b)  3 1/2 = ______        c)  1 4/5 = ______", y);
    y = addWriteLine(doc, "d)  2 3/4 = ______        e)  4 2/3 = ______        f)  2 5/6 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Place on the number line (0 to 3)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Mark each fraction with an arrow on the line below.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "5/4        7/3        2 1/2", y);
    y = addWriteLine(doc, "Number line 0 ____________________ 1 ____________________ 2 ____________________ 3", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order these from smallest to largest: 2 1/4, 9/4, 1 3/4, 5/2. Explain how you decided.", y);
    y = addWriteLine(doc, "Order: ____________________________________________________", y);

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Mixed and Improper Fractions | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the mixed and improper fractions practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Week ${WEEK} Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Improper to mixed", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1 3/4      b)  4 1/2      c)  2 3/4      d)  2 2/3      e)  3 2/5      f)  2 1/6", y);

    y = addSectionHeading(doc, "Section 2 - Mixed to improper", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  7/3      b)  7/2      c)  9/5      d)  11/4      e)  14/3      f)  17/6", y);

    y = addSectionHeading(doc, "Section 3 - Place on the number line", y, { color: C.PRIMARY });
    y = addBodyText(doc, "5/4 = 1 1/4 (just past 1).   7/3 = 2 1/3 (just past 2).   2 1/2 (halfway between 2 and 3).", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "1 3/4 < 2 1/4 = 9/4 < 5/2. Note 2 1/4 and 9/4 are equal. 5/2 = 2 1/2 is largest.", y);

    y = addTipBox(doc,
      "Watch for: students who multiply the whole by the top; students who forget the leftover; students who write the leftover as a whole number.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Week ${WEEK} Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Week 5 Session 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
