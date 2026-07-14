"use strict";

// Term 3 Maths Review (Year 6 Numeracy) - Session 2 of 4.
// FRACTIONS: compare, convert, add and subtract. Reviews Week 5 (VC2M5N03 +
// VC2M6N05) and Week 8 (VC2M5N03 + VC2M5N05).
//   Block A - equivalent fractions and comparing related denominators (rename).
//   Block B - convert between mixed numerals and improper fractions.
//   Block C - add and subtract fractions with related denominators.
// Daily Review: Mastering Fractions, Decimals & Percentages (prior, Week 9).
// Fluency: subtraction vertical algorithm with decimals (unit-wide focus).
// Unit variant fixed (variant 2) across all 4 sessions for cohesion.
// Catch-up: the launch starts from a fraction wall everyone can read, and
// worksheet Section 1 re-grounds equivalence. No session assumes the one before.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 2); // variant 2, fixed for the unit
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide, dailyReviewSlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  addFractionStripSet, addNumberLine,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 4;
const UNIT_TITLE = "Term 3 Maths Review";
const FOOTER = `Term 3 Maths Review | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/TermRev_Lesson2_Fractions";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Compare related fractions, convert mixed numbers, and add and subtract with related denominators.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Session 2 review sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Adding, subtracting and multiplying fractions with unrelated denominators - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to our Term 3 review. Today is fractions day.
- We bring back three big skills: comparing fractions, swapping between mixed numbers and improper fractions, and adding and subtracting.
- The trick that ties it all together is the same one: make the pieces the same size before you compare or combine.

DO:
- Have whiteboards, markers and the review sheet ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Session 2 of 4. A review, so retrieve and firm up. The fraction wall and strips do the heavy lifting today; keep it visual.

WATCH FOR:
- Students who freeze at fractions - reassure them. The strips make it concrete again.

[General: Title | Element: Planning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The review sheet is for the You Do near the end. The Year 8 extension is there for anyone who is flying.

DO:
- Print one review sheet and one answer key per student.
- Have whiteboards and markers ready for every check.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student review sheet, an answer key, and a Year 8 extension on unlike denominators and multiplying fractions. Most of the lesson runs on whiteboards.

CATCH-UP NOTE:
A student who missed earlier work can still access today. The launch starts from a fraction wall anyone can read, and Section 1 of the review sheet rebuilds equivalent fractions from scratch. A returner needs only a whiteboard and one minute with you to start. No session in this unit assumes the student saw the one before.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Warming up our most recent work - finding a fraction or percentage of an amount.
- Read each one and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for students finding one part first, then scaling.

TEACHER NOTES:
Daily Review is prior learning. Finding a part of a quantity keeps ticking over from Week 9 and links to today's fraction work.

WATCH FOR:
- Students who find 3/4 by finding a quarter then tripling - secure.
- Students who try to do it all at once - prompt find one part first.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Half of 30 is 15.
- 20% of 50 is one fifth of 50, which is 10.
- 3/4 of 12: a quarter of 12 is 3, so three quarters is 9.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The move to reinforce: find one part, then scale. A quarter of 12 is 3, so three quarters is 9.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students stuck on 3/4 - model the find-a-quarter-then-times-3 step.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Subtraction set out vertically, with decimals.
- Set each one out in columns on your whiteboard and subtract.
- Line up the decimal points so the places sit under each other.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up columns and the point kept in line.

TEACHER NOTES:
Fluency this whole unit is the subtraction algorithm with decimals. Lining up the points is the same same-size-pieces idea we use with fractions today.

WATCH FOR:
- Students who line up the points and fill gaps with zeros - secure.
- Students who right-align the digits instead of the points - fix it in the reveal.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 5.6 take away 2.9 is 2.7.
- 9.30 take away 4.55 is 4.75.
- 6.0 take away 2.4 is 3.6. Writing 6.0 as 6.00 helps line it up.
- Tick and fix.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The 6.0 minus 2.4 item is the key one. Fill the empty place with a zero to keep the columns honest.

WATCH FOR:
- Students who self-correct - secure.
- Students whose columns drift - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember the fraction wall. Look at these two strips of the same length.
- The top strip is split into 2, and one part is shaded - that is one half. The bottom is split into 3, and one part is shaded - that is one third.
- Which shaded piece is bigger? The half. When we cut a whole into fewer pieces, each piece is larger.
- So one half is greater than one third, even though 3 is bigger than 2. That idea - look at the size of the piece, not just the number - drives today.

DO:
- Point to each shaded strip and compare the lengths.
- Ask for a thumbs up or down: is one half bigger than one third?
- Bridge: 'to compare fractions we make the pieces the same size'.

TEACHER NOTES:
This launch starts from a picture anyone can read, then connects it to today's compare-and-combine work. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who read the strip lengths confidently - strong prior knowledge.
- Students tempted to say a third is bigger because 3 is bigger - the strips settle it; we build the same-size-pieces habit this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are reviewing how to compare fractions, swap between mixed and improper fractions, and add and subtract them.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Point out that the first one is something everyone can do today.

TEACHER NOTES:
The first criterion is reachable for everyone - compare two fractions with a wall. The second is the core target the exit ticket checks. The third stretches to adding and subtracting with renaming.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- Equivalent fractions are equal in value but written differently, like one half and two quarters.
- A mixed numeral is a whole number and a fraction together, like 2 and one quarter.
- An improper fraction is top heavy - the numerator is bigger than the denominator, like nine quarters.

DO:
- Point to each word as you say it.
- Have students say 'equivalent means equal in value' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. Equivalent is the anchor word - it powers comparing, renaming, adding and subtracting.

WATCH FOR:
- Students who give another name for one half - secure.
- Students who think 2 1/4 and 9/4 are different amounts - the number line later fixes this.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO_A = `SAY:
- Let us compare two thirds and five sixths. The denominators are different, so the pieces are different sizes. We cannot compare yet.
- Watch how I rename. Thirds and sixths are related - six is a multiple of three. I cut each third in half, and two thirds becomes four sixths. Same amount, smaller pieces.
- Now both are in sixths. Four sixths against five sixths - five sixths has more pieces of the same size, so it is larger.
- So two thirds is less than five sixths. The strips show it, and the renaming proves it.

DO:
- Point to the 2/3 strip, then the renamed 4/6 strip, then the 5/6 strip.
- Trace how each third splits into two sixths.
- Have students chorus 'make the pieces the same size first'.

TEACHER NOTES:
This is the core move of Block A. Renaming to a common denominator is the same skill used to add and subtract later in the lesson.

MISCONCEPTIONS:
- Misconception: five sixths is smaller because sixths are smaller pieces.
  Why: students confuse piece size with how much is shaded.
  Impact: they order fractions back to front.
  Quick correction: once both are in sixths, just count the pieces - 5 beats 4.

WATCH FOR:
- Students who rename then count - secure.
- Students comparing the raw numbers - send them back to the strips.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_A_Q = `SAY:
- Quick check on your whiteboards.
- Which is larger: three quarters or five eighths?
- Rename first so the pieces are the same size.

DO:
- Display the two fractions.
- Give 60 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 3/4 = 6/8, so 3/4 is larger.
PROCEED: If about 80 percent have it right, click to reveal and move to Block B.
PIVOT: Most likely misconception - students compare 3 and 5 and pick five eighths.
- Reteach: quarters and eighths are related; rename 3/4 as 6/8 by splitting each quarter in half. Then 6/8 against 5/8.
- Re-check: how many eighths is three quarters?

TEACHER NOTES:
The trap is comparing the top numbers. A student who renames to eighths sees 6/8 beats 5/8.

WATCH FOR:
- Students with 3/4 larger - secure.
- Students with 5/8 larger - reteach the renaming with eighths.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_IDO_B = `SAY:
- Now we review swapping between mixed numerals and improper fractions.
- Here is 2 and one quarter on the number line. It sits a quarter past 2.
- To write it as an improper fraction, I count all the quarters from zero. Two whole ones is eight quarters, plus the extra quarter makes nine quarters.
- So 2 and one quarter equals nine quarters. On the line they are the exact same point.
- The quick way: whole times denominator, plus the numerator, all over the denominator. Two times four is eight, plus one is nine, over four.

DO:
- Point to 2 1/4 on the line and count the quarter ticks from zero.
- Write the conversion: 2 x 4 + 1 = 9, over 4.
- Have students whisper 'count all the pieces' to a partner.

TEACHER NOTES:
This is VC2M5N03. Tie the symbol rule back to counting the quarters on the line so it is not a trick.

MISCONCEPTIONS:
- Misconception: 2 1/4 becomes 3/4 by adding the whole to the numerator.
  Why: students add 2 and 1 and ignore the wholes as quarters.
  Impact: a far too small improper fraction.
  Quick correction: each whole is four quarters; count them all - 8 plus 1 is 9.

WATCH FOR:
- Students who count all the quarters - secure.
- Students who write 3/4 - back to the line, count the wholes as quarters.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_WEDO_B_Q = `SAY:
- Your turn with a partner on your whiteboards.
- One: convert 3 and two fifths to an improper fraction.
- Two: convert eleven quarters to a mixed numeral.

DO:
- Display the two prompts.
- Give 75 seconds.
- Listen for whole times denominator plus numerator, and for how many wholes fit.

TEACHER NOTES:
Same moves as the I Do with new numbers. For 11/4, ask how many fours fit into 11.

WATCH FOR:
- Pairs who convert both ways - secure.
- Pairs stuck on 11/4 - prompt how many whole quarters fit, two wholes is eight quarters.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO_B_A = `SAY:
- Let us check.
- 3 and two fifths: three times five is fifteen, plus two is seventeen, so seventeen fifths.
- Eleven quarters: two fours make eight, leaving three quarters, so 2 and three quarters.
- The number line would show each pair landing on the same point.

DO:
- Click to reveal.
- Re-say the rule for each direction.

TEACHER NOTES:
If pairs struggled, model 11/4 on a 0 to 3 quarter line before Block C.

WATCH FOR:
- Students who self-correct - secure.
- Students who miscount the wholes - small group focus.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_IDO_C = `SAY:
- Now we add fractions with related denominators. Add one half and one quarter.
- The pieces are different sizes, so I cannot add yet. Watch how I rename one half as two quarters - same amount, now in quarters.
- Now every piece is a quarter. Two quarters and one quarter make three quarters. I just count the quarters.
- The rule: make the same-size pieces first by renaming, then add the numerators and keep the denominator.

DO:
- Point to the 1/2 strip renamed as 2/4, then the 1/4 strip, then the 3/4 total.
- Write 2/4 + 1/4 = 3/4.
- Have students chorus 'same-size pieces, then add the tops'.

TEACHER NOTES:
This is VC2M5N05. Keep tying renaming back to the strips. The denominator names the piece size and does not change when we add.

MISCONCEPTIONS:
- Misconception: add the tops and the bottoms, so 1/2 + 1/4 = 2/6.
  Why: students treat numerator and denominator the same way.
  Impact: an answer smaller than one of the parts, which cannot be right.
  Quick correction: the denominator names the piece size; rename to quarters, then add only the tops.

WATCH FOR:
- Students who rename then add tops - secure.
- Students who add bottoms too - back to the strips, the piece size stays quarters.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_WEDO_C_Q = `SAY:
- Your turn with a partner. Subtract: five sixths take away one third.
- Sixths and thirds are related. Rename one third so the pieces match.
- Use the strips to check your answer makes sense.

DO:
- Display the prompt and the strips.
- Give about 90 seconds.
- Listen for renaming one third as two sixths.

TEACHER NOTES:
Subtraction works the same way as addition - rename first, then subtract the tops. One third becomes two sixths.

WATCH FOR:
- Pairs who rename and subtract the tops - secure.
- Pairs who subtract bottoms too - prompt back to the strips.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO_C_A = `SAY:
- Let us check on the strips.
- One third is the same as two sixths. So five sixths take away two sixths is three sixths.
- Three sixths is the same as one half. So the answer is three sixths, or one half.

DO:
- Click to reveal the answer.
- Point out that 3/6 simplifies to 1/2.

TEACHER NOTES:
If many subtracted the bottoms, do one more like 3/4 - 1/8 before the You Do.

WATCH FOR:
- Students who self-correct - ready for independent work.
- Students still subtracting bottoms - enabling group for the You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the review sheet.
- Section 1 warms up by comparing fractions with a wall - that is for everyone.
- Section 2 converts mixed and improper fractions. Section 3 adds and subtracts with related denominators.
- If you finish, try the challenge box, then the Year 8 extension.

DO:
- Distribute the review sheet.
- Circulate and listen for same-size-pieces language.
- Cold call one or two students to explain an answer.

TEACHER NOTES:
Different numbers from the We Do, same moves. Section 1 is the rebuild for any returning student.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 gives a fraction wall and one worked comparison. Students shade and compare related fractions.
- Extra Notes: Sit with these students and point to the strip lengths. This is also the rebuild for a returner.
EXTENDING PROMPT:
- Task: The challenge box asks students to add three related fractions and simplify the answer.
- Extra Notes: Students who are ready move on to the Year 8 extension on unlike denominators and multiplying.

WATCH FOR:
- Students who rename before combining - secure.
- Students who add or subtract the bottoms - prompt back to the wall.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- One: work out two thirds plus one sixth.
- Two: convert seven quarters to a mixed numeral.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - rename related fractions to add, and convert improper to mixed. Look for 5/6 and 1 3/4.

WATCH FOR:
- Students who rename 2/3 to 4/6 and get 5/6 - secure.
- Students who add bottoms or miscount wholes - revisit at the start of Session 3.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: why do we have to rename one half before we add it to one quarter?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that fractions must be the same-size pieces before you compare or combine them. Students who can say this are ready for percentages next session.

WATCH FOR:
- Strong thumbs up across all three - move at pace next session.
- Sideways or down on the core criterion - small group revision at the start of Session 3.

[General: Closing | Element: Retention and Recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 2: Fractions - compare, convert, add and subtract",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - finding a part of a quantity
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Find part of an amount",
      [
        "Find 1/2 of 30.",
        "Find 20% of 50.",
        "Find 3/4 of 12.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1/2 of 30 = 15       20% of 50 = 10       3/4 of 12 = 9", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - subtraction with decimals
  withReveal(
    () => fluencySlide(pres, "Fluency: Subtraction with decimals",
      ["5.6 - 2.9", "9.30 - 4.55", "6.0 - 2.4"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "2.7        4.75        3.6", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - which is bigger, 1/2 or 1/3 (fraction wall, catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "Which is bigger: one half or one third?",
    [
      "Two strips, the same length.",
      "Top split into 2; one part shaded.",
      "Bottom split into 3; one part shaded.",
      "",
      "Fewer pieces means bigger pieces.",
      "One half beats one third.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.ACCENT });
      addFractionStripSet(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.55, lg.rightW - 0.5, 1.5,
        [
          { denom: 2, shaded: 1, label: "1/2" },
          { denom: 3, shaded: 1, label: "1/3" },
        ], { labelW: 0.7 });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are reviewing how to compare fractions, convert mixed and improper fractions, and add and subtract them.",
    [
      "I can compare two fractions using a fraction wall.",
      "I can convert between mixed numerals and improper fractions.",
      "I can add and subtract fractions with related denominators.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Equivalent = equal in value (1/2 = 2/4)",
      "Mixed numeral = whole + fraction (2 1/4)",
      "Improper fraction = top heavier than bottom (9/4)",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.3, { strip: C.SECONDARY });
      slide.addText("1/2 and 2/4 are equivalent", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.32,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
      });
      addFractionStripSet(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.62, lg.rightW - 0.5, 1.3,
        [
          { denom: 2, shaded: 1, label: "1/2" },
          { denom: 4, shaded: 2, label: "2/4" },
        ], { labelW: 0.7 });
    }
  );

  // Slide 10: I Do A - compare related denominators (rename)
  workedExSlide(pres, 2, "I Do", "Compare: 2/3 and 5/6",
    [
      "Different denominators = different size pieces.",
      "Thirds and sixths are related (6 is a multiple of 3).",
      "Rename: 2/3 = 4/6 (cut each third in half).",
      "Now compare sixths: 4/6 and 5/6.",
      "",
      "5/6 has more pieces, so 2/3 < 5/6.",
    ],
    NOTES_IDO_A, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.PRIMARY });
      slide.addText("Make the pieces the same size", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addFractionStripSet(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.55, lg.rightW - 0.5, 1.9,
        [
          { denom: 3, shaded: 2, label: "2/3" },
          { denom: 6, shaded: 4, label: "4/6" },
          { denom: 6, shaded: 5, label: "5/6" },
        ], { labelW: 0.7 });
    }
  );

  // Slides 11-12: CFU A + reveal - compare related fractions
  withReveal(
    () => cfuSlide(pres, "CFU", "Which is larger?",
      { technique: "Show Me Boards",
        question: "Which is larger:\n\n3/4   or   5/8 ?\n\nRename first so the pieces are the same size." },
      NOTES_CFU_A_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3/4 = 6/8, so 3/4 is larger", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 21, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: I Do B - mixed <-> improper on a number line (full width)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Mixed numeral and improper fraction", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.5, { strip: STAGE_COLORS["2"] });
    s.addText([
      { text: "2 1/4 ", options: { fontSize: 17, color: C.PRIMARY, bold: true } },
      { text: "means 2 wholes and one quarter. Count all the quarters from zero: ", options: { fontSize: 15, color: C.CHARCOAL } },
      { text: "2 wholes = 8 quarters, + 1 = 9 quarters = 9/4.   ", options: { fontSize: 15, color: C.SECONDARY, bold: true } },
      { text: "Quick rule: whole x denominator + numerator, over the denominator.", options: { fontSize: 15, color: C.CHARCOAL, italic: true } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.14, w: 8.5, h: 1.2, fontFace: FONT_B, valign: "middle", margin: 0,
    });

    s.addText("Count the quarters: 2 1/4 and 9/4 are the same point", {
      x: 0.7, y: 3.25, w: 8.6, h: 0.28, fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
    });
    const lineX = 0.9, lineW = 8.0, lineY = 4.3;
    addNumberLine(s, lineX, lineY, lineW,
      ["0", "", "", "", "1", "", "", "", "2", "", "", "", "3"], [9]);
    // Callout above the marked point (index 9 of 12 intervals).
    const markX = lineX + (9 / 12) * lineW;
    s.addText("2 1/4 = 9/4", {
      x: markX - 0.9, y: lineY - 0.5, w: 1.8, h: 0.3,
      fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_B);
    runSlideDiagnostics(s, pres);
  })();

  // Slides 14-15: We Do B + reveal - convert both ways
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your turn: convert",
      [
        "With your partner.",
        "",
        "1.  Write 3 2/5 as an improper fraction.",
        "2.  Write 11/4 as a mixed numeral.",
        "",
        "Count all the pieces.",
      ],
      NOTES_WEDO_B_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
        slide.addText("Whole x denominator\n+ numerator", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.2, w: lg.rightW - 0.4, h: 0.8,
          fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        addTextOnShape(slide, "How many wholes\nfit into 11/4?", {
          x: lg.rightX + 0.5, y: lg.panelTopPadded + 1.2, w: lg.rightW - 1.0, h: 0.9, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      }),
    (slide) => {
      addTextOnShape(slide, "3 2/5 = 17/5        11/4 = 2 3/4", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_B_A);
    }
  );

  // Slide 16: I Do C - add with related denominators (rename, then add)
  workedExSlide(pres, 2, "I Do", "Add: 1/2 + 1/4",
    [
      "Different size pieces - rename first.",
      "1/2 is the same as 2/4.",
      "Now every piece is a quarter.",
      "2/4 + 1/4 = 3/4.",
      "",
      "Same-size pieces, then add the tops.",
    ],
    NOTES_IDO_C, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.PRIMARY });
      slide.addText("Rename 1/2 as 2/4, then add", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addFractionStripSet(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.55, lg.rightW - 0.5, 1.9,
        [
          { denom: 4, shaded: 2, label: "1/2" },
          { denom: 4, shaded: 1, label: "1/4" },
          { denom: 4, shaded: 3, label: "3/4" },
        ], { labelW: 0.7 });
    }
  );

  // Slides 17-18: We Do C + reveal - subtract with related denominators
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Subtract: 5/6 - 1/3",
      [
        "With your partner.",
        "",
        "Sixths and thirds are related.",
        "Rename 1/3 so the pieces match.",
        "Then subtract the tops.",
      ],
      NOTES_WEDO_C_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.SECONDARY });
        slide.addText("Rename 1/3 as 2/6", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        addFractionStripSet(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.55, lg.rightW - 0.5, 1.85,
          [
            { denom: 6, shaded: 5, label: "5/6" },
            { denom: 6, shaded: 2, label: "2/6" },
            { denom: 6, shaded: 3, label: "3/6" },
          ], { labelW: 0.7 });
      }),
    (slide) => {
      addTextOnShape(slide, "1/3 = 2/6, so 5/6 - 2/6 = 3/6 = 1/2", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_C_A);
    }
  );

  // Slide 19: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: review sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "compare the fractions with the wall.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "convert mixed and improper fractions.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "add and subtract with related denominators.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.4;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember", {
      x: 0.7, y: panelY + 0.13, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
    });
    addTextOnShape(s, "Make the pieces the same size first. Then compare, add or subtract the tops.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    s.addText("Reference wall:", {
      x: 0.7, y: panelY + 1.06, w: 2.4, h: 0.28,
      fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    addFractionStripSet(s, 3.0, panelY + 1.04, 4.0, 0.9,
      [
        { denom: 2, shaded: 1, label: "1/2" },
        { denom: 4, shaded: 2, label: "2/4" },
      ], { labelW: 0.7 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 20: Exit Ticket
  exitTicketSlide(pres,
    [
      "Work out 2/3 + 1/6.",
      "Convert 7/4 to a mixed numeral.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 21: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why do we have to rename 1/2 before we add it to 1/4?",
      scItems: [
        "I can compare two fractions using a fraction wall.",
        "I can convert between mixed numerals and improper fractions.",
        "I can add and subtract fractions with related denominators.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "TermRev_Lesson2_Fractions.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Compare related fractions, convert mixed numbers, and add and subtract.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Make the pieces the same size before you compare or combine fractions. Rename to a common denominator using equivalent fractions, then compare, add or subtract the numerators. The denominator names the piece size and does not change.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Add 1/2 + 1/4. Rename 1/2 as 2/4 so the pieces match. Then 2/4 + 1/4 = 3/4.",
      y);

    y = addSectionHeading(doc, "Section 1 - Compare (for everyone)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  Circle the larger:   2/3   or   3/4   (rename to twelfths if it helps)", y);
    y = addWriteLine(doc, "b)  Circle the larger:   5/6   or   2/3", y);
    y = addWriteLine(doc, "c)  Order smallest to largest:   1/2,  3/8,  3/4   ->", y);

    y = addSectionHeading(doc, "Section 2 - Convert mixed and improper", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2 3/5  as an improper fraction = ______", y);
    y = addWriteLine(doc, "b)  3 1/2  as an improper fraction = ______", y);
    y = addWriteLine(doc, "c)  13/4  as a mixed numeral = ______", y);
    y = addWriteLine(doc, "d)  9/2  as a mixed numeral = ______", y);

    y = addSectionHeading(doc, "Section 3 - Add and subtract (related denominators)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/3 + 1/6 = ______", y);
    y = addWriteLine(doc, "b)  3/4 - 1/8 = ______", y);
    y = addWriteLine(doc, "c)  2/5 + 3/10 = ______", y);

    y = addSectionHeading(doc, "Challenge (optional)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "Work out 1/2 + 1/4 + 1/8, then simplify if you can.  ______", y);

    addPdfFooter(doc, `Session ${SESSION} | Fractions | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Session 2 review sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Compare", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3/4 is larger (3/4 = 9/12, 2/3 = 8/12).   b)  5/6 is larger (2/3 = 4/6).   c)  3/8, 1/2, 3/4.", y);

    y = addSectionHeading(doc, "Section 2 - Convert", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  13/5        b)  7/2        c)  3 1/4        d)  4 1/2", y);

    y = addSectionHeading(doc, "Section 3 - Add and subtract", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/3 + 1/6 = 2/6 + 1/6 = 3/6 = 1/2.   b)  3/4 - 1/8 = 6/8 - 1/8 = 5/8.   c)  2/5 + 3/10 = 4/10 + 3/10 = 7/10.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "1/2 + 1/4 + 1/8 = 4/8 + 2/8 + 1/8 = 7/8.", y);

    y = addTipBox(doc,
      "Watch for: students who compare the top numbers only; students who add or subtract the denominators; students who add the whole to the numerator when converting (2 1/4 is 9/4, not 3/4).",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Unlike denominators and multiplying fractions.",
      color: C.SECONDARY,
      lessonInfo: `Session ${SESSION} | Year 8 challenge | extends Year 6 VC2M5N05`,
    });
    y = addTipBox(doc,
      "When denominators are NOT related, find a common denominator by multiplying. For 1/3 + 1/4, twelfths work for both: 1/3 = 4/12, 1/4 = 3/12. To multiply fractions, multiply the tops and multiply the bottoms.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Add 1/3 + 1/4. A common denominator is 12. 1/3 = 4/12 and 1/4 = 3/12, so 4/12 + 3/12 = 7/12.",
      y);

    y = addSectionHeading(doc, "Section 1 - Unlike denominators", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2 + 1/3 = ______", y);
    y = addWriteLine(doc, "b)  3/4 - 2/3 = ______", y);
    y = addWriteLine(doc, "c)  2/5 + 1/2 = ______", y);

    y = addSectionHeading(doc, "Section 2 - Multiply fractions", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/2 x 1/3 = ______", y);
    y = addWriteLine(doc, "b)  2/3 x 3/4 = ______  (simplify)", y);
    y = addWriteLine(doc, "c)  3/5 x 1/2 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "When you multiply two fractions less than 1, is the answer larger or smaller than both? Explain.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) 5/6  b) 1/12  c) 9/10.   S2  a) 1/6  b) 6/12 = 1/2  c) 3/10.   S3  Smaller than both - taking a fraction OF a fraction gives a smaller part.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Session ${SESSION} | Year 8 Extension | Operating with fractions`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Session 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
