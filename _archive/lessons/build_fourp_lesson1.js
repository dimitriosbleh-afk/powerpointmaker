"use strict";

// Four Processes (Year 6 Numeracy) - Lesson 1 of 5.
// Multiplication and division are linked. One array tells a FAMILY of facts
// (2 multiplication + 2 division). The inverse (multiplication undoes division)
// lets us solve missing-number problems and check our work. VC2M5A01.
// Daily Review: Fractions & Decimals (prior). Fluency: decimal x 1-digit algorithm.
// Unit variant fixed (variant 3, Ocean Logic) across all 5 lessons for cohesion.
// Catch-up: the launch rebuilds the idea from a tiny 2x5 array everyone can see,
// and worksheet Section 1 is an enabling rebuild. No session assumes the one before.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 3); // variant 3 (Ocean Logic), fixed for the unit
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide, dailyReviewSlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 5;
const UNIT_TITLE = "Four Processes";
const FOOTER = `Four Processes | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FourP_Lesson1_Multiplication_Division_Links";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Write families of facts from arrays and use the inverse to find missing numbers.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 1 practice sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Using inverse operations to solve one- and two-step equations - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Small reusable array helper (a clean grid of equal squares) -------------
// Used on the launch, I Do, CFU and We Do panels so the array visual is
// consistent across the lesson. roundRect renders reliably in LibreOffice.
function drawArray(slide, x, y, rows, cols, cell, opts) {
  const o = opts || {};
  const fill = o.fill || C.SECONDARY;
  const gap = o.gap != null ? o.gap : 0.05;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      slide.addShape("roundRect", {
        x: x + c * (cell + gap), y: y + r * (cell + gap),
        w: cell, h: cell, rectRadius: cell * 0.18,
        fill: { color: fill }, line: { color: C.WHITE, width: 1 },
      });
    }
  }
  return { w: cols * (cell + gap) - gap, h: rows * (cell + gap) - gap };
}

// Fact-triangle style panel: the PRODUCT (whole) sits on top; the two FACTORS
// below multiply UP to it, and it divides DOWN by either factor. Makes the
// inverse visible.
function addFactTriangle(slide, x, y, w, whole, a, b, color) {
  const cx = x + w / 2;
  const chipW = 1.05;
  const symW = 0.5;
  addTextOnShape(slide, String(whole), {
    x: cx - 0.7, y: y, w: 1.4, h: 0.55, rectRadius: 0.08,
    fill: { color: color },
  }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  slide.addText("Multiply up   |   Divide down", {
    x: x, y: y + 0.62, w: w, h: 0.26,
    fontSize: 11.5, fontFace: FONT_B, color: C.MUTED, italic: true,
    align: "center", margin: 0,
  });
  addTextOnShape(slide, String(a), {
    x: x, y: y + 0.95, w: chipW, h: 0.55, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
  slide.addText("x", {
    x: cx - symW / 2, y: y + 0.95, w: symW, h: 0.55,
    fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  addTextOnShape(slide, String(b), {
    x: x + w - chipW, y: y + 0.95, w: chipW, h: 0.55, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to our new unit, Four Processes. Over the week we get smarter at the four operations - adding, subtracting, multiplying and dividing.
- Today we start with a big idea: multiplication and division are linked. If you know one, you can work out the other.
- By the end you will be able to look at one array and read four number facts from it.

DO:
- Have whiteboards, markers and counters ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 1 of 5. This lesson sets up the inverse relationship that the whole unit leans on. Keep it concrete with arrays today.

WATCH FOR:
- Students who look unsure - that is expected on day one. Reassure them: if this feels new, that is okay, we build it together.

[General: Title | Element: Planning and Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end. The Year 8 extension is there for anyone who is flying.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards, markers and a tub of counters ready for building arrays.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student practice sheet, an answer key, and a Year 8 extension. Most teaching happens on whiteboards and with counters.

CATCH-UP NOTE:
A student who missed earlier sessions can still access today. The launch rebuilds the idea from a small 2 by 5 array everyone can see, and Section 1 of the worksheet rebuilds it again. A returner only needs counters and one minute with you to start. No session in this unit assumes the student saw the one before.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are warming up our fractions and decimals from earlier work.
- Read each one carefully and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for how students compare decimal sizes.

TEACHER NOTES:
Daily Review is prior learning, not today's new content. Fractions and decimals keep ticking over because we use them later this week.

WATCH FOR:
- Students who line up place value when comparing decimals - secure.
- Students who think 0.45 is bigger than 0.6 because 45 is bigger than 6 - common error, address it in the reveal.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 0.7 is the same as seven tenths, or 7 over 10.
- 0.6 is larger than 0.45. Six tenths beats four tenths, even though 45 looks like a big number.
- Smallest to largest: 0.05, then 0.5, then 0.55.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The 0.6 versus 0.45 item is the key one. Line up the tenths place: 6 tenths against 4 tenths. The number of digits after the point does not decide the size.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students who still trust digit count over place value - small group focus.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Multiplying a decimal by a single digit using the algorithm.
- Set each one out vertically on your whiteboard and multiply.
- Keep your columns lined up and remember where the decimal point sits.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for clean column alignment and the point kept in line.

TEACHER NOTES:
Fluency this unit is the multiplication algorithm with decimals. Today is decimal times one digit. Lining up columns now is the same place value habit we use all week.

WATCH FOR:
- Students who keep the decimal point lined up - secure.
- Students who lose the point - prompt: the answer has the same number of decimal places as the number you multiplied.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 4.6 times 3 is 13.8.
- 2.7 times 6 is 16.2.
- 3.5 times 4 is 14.0, which we write as 14.
- Tick and fix.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
One decimal place in, one decimal place out. 3.5 times 4 lands on 14.0; accept 14. Keep it brisk.

WATCH FOR:
- Students who self-correct - secure.
- Students whose point drifts - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember building arrays with counters. Look at this one: 2 rows of 5.
- I can read it across: 2 rows of 5 is 10. I can read it down: 5 columns of 2 is 10. Same picture, both true.
- Now think backwards. 10 shared into 2 rows gives 5 in each row. 10 shared into 5 columns gives 2 in each.
- One little array just told us four facts. Today we use bigger numbers and let that idea do the heavy lifting.

DO:
- Point to the rows, then the columns, then trace the sharing back.
- Have students chorus the four facts with you.
- Bridge: 'multiplication and division come from the same picture'.

TEACHER NOTES:
This launch starts from a tiny array everyone can access, then connects it to today's fact families and the inverse. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who can give all four facts - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning that multiplication and division are linked, so we can use one to work out the other.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Hold up an array of counters as a reminder.

TEACHER NOTES:
The first criterion is reachable for everyone - write the facts from an array. The second is the core target the exit ticket checks. The third stretches to explaining why the inverse helps.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- An array is equal rows and columns, like counters in a neat grid.
- Inverse means the operation that undoes another. Multiplying and dividing are inverses.
- A family of facts is the two times facts and two divide facts that come from one array.

DO:
- Point to each word as you say it.
- Have students say 'multiply and divide are inverses' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. Keep it to these three. 'Inverse' is the anchor word for the whole unit.

WATCH FOR:
- Students who can name the inverse of multiply - secure.
- Students who think inverse means opposite number - clarify: it means the undoing operation.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at this one together. Here is an array of 4 rows with 6 in each row.
- Watch how I read it. Across the rows: 4 sixes is 24. Down the columns: 6 fours is 24. So 4 times 6 and 6 times 4 both make 24.
- Now I think backwards. If 24 counters make 4 equal rows, each row has 6, so 24 divided by 4 is 6. If they make 6 equal columns, each has 4, so 24 divided by 6 is 4.
- One array, four facts. That is a family of facts.

DO:
- Point to the rows, then the columns, then trace the division back.
- Write the four facts one at a time as you say them.
- Have students chorus 'one array, four facts'.

TEACHER NOTES:
This is the core move of the lesson. Always tie each fact back to the picture so division is sharing the array, not a separate mystery.

MISCONCEPTIONS:
- Misconception: students think 24 divided by 4 needs a brand new strategy.
  Why: division is often taught separately from multiplication.
  Impact: they get stuck on division facts they could read straight off the array.
  Quick correction: point to the rows - 4 rows of 6 means 24 divided by 4 is 6.

WATCH FOR:
- Students who read all four facts from the picture - secure.
- Students who only see the two multiplication facts - prompt the sharing-back step.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now watch how the inverse helps me with a harder one: 240 divided by 20.
- Instead of long division, I ask the multiplication question: 20 times what gives 240?
- I know 20 times 12 is 240, so 240 divided by 20 must be 12. The multiplication did the work.
- It runs the other way too. If I am stuck on 17 times something equals 221, I can divide: 221 divided by 17 is 13.
- When one is tricky, swap to its inverse.

DO:
- Reveal the two chips one at a time.
- Trace the arrow from the times fact to the divide fact and back.
- Have students whisper 'swap to the inverse' to a partner.

TEACHER NOTES:
This connects the fact-family idea to a real strategy: turn a hard division into an easier multiplication, or a hard missing factor into a division. This is the heart of VC2M5A01.

MISCONCEPTIONS:
- Misconception: students believe division must always be done by an algorithm.
  Why: they have not connected it to known multiplication facts.
  Impact: slow, error-prone work on facts they actually know.
  Quick correction: model 'what times the divisor gives this number'.

WATCH FOR:
- Students who reach for a known times fact - secure.
- Students who start long division on 240 divided by 20 - prompt the inverse question.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards. Here is an array of 3 rows of 5.
- Write the whole family of facts: two times facts and two divide facts.

DO:
- Display the array.
- Give 60 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 3 x 5 = 15, 5 x 3 = 15, 15 divided by 3 = 5, 15 divided by 5 = 3.
PROCEED: If about 80 percent have all four facts, click to reveal and move to We Do.
PIVOT: Most likely misconception - students write the two multiplication facts but miss the divisions.
- Reteach: point to the 3 rows. 15 counters in 3 equal rows means 15 divided by 3 is 5. Do the same for the columns.
- Re-check: where does each division come from in the picture?

TEACHER NOTES:
The trap is stopping at two facts. A student who writes only multiplication has not yet seen division as sharing the same array.

WATCH FOR:
- Students with all four facts - secure.
- Students with only two - reteach the sharing-back step with the array.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. 8 times something equals 56. Find the missing number.
- Do not guess and check. Use the inverse: what division undoes 8 times?
- Whisper to your partner the division fact you would use.

DO:
- Display 8 x ? = 56 with the fact triangle.
- Give 60 seconds.
- Listen for '56 divided by 8'.

TEACHER NOTES:
Same inverse move as the I Do with new numbers. Listen for division language, not guessing.

WATCH FOR:
- Pairs who say 56 divided by 8 is 7 - secure.
- Pairs who count up in 8s - accept it, then show the inverse as the faster route.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- 8 times something equals 56. The inverse is 56 divided by 8, which is 7.
- So the missing number is 7, and 8 times 7 is 56. The division undid the multiplication.

DO:
- Click to reveal.
- Run the inverse once more, both directions, on the triangle.

TEACHER NOTES:
Reveal restates the inverse. Point out that we can check instantly: 8 times 7 is 56, so it fits.

WATCH FOR:
- Students who self-correct - secure.
- Students who still guess - steer them to the division fact.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, a bigger one. 180 divided by 30.
- Swap to the inverse: 30 times what gives 180?
- Build it on your whiteboard and tell your partner the times fact you used.

DO:
- Display 180 divided by 30 with the fact triangle.
- Give 75 seconds.
- Watch for students who try long division instead of the inverse.

TEACHER NOTES:
This shows the inverse turning a chunky division into a known times fact. 30 times 6 is 180, so 180 divided by 30 is 6.

WATCH FOR:
- Students who use 30 times 6 - secure.
- Students who start long division - prompt the inverse question.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- 30 times 6 is 180, so 180 divided by 30 is 6.
- The multiplication fact did the hard work. The division just read off the answer.

DO:
- Click to reveal.
- Point to the times fact that solved it.

TEACHER NOTES:
If many tried long division, do one more like 240 divided by 40 before releasing to the You Do.

WATCH FOR:
- Students who use the inverse fluently - ready for independent work.
- Students who still default to long division - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 warms up by reading fact families from arrays.
- Section 2 uses the inverse to find missing numbers.
- Section 3 asks you to explain how a times fact helps a divide. If you finish, try the extension box.

DO:
- Distribute the practice sheet.
- Circulate and listen for inverse language.
- Cold call one or two students to explain a missing-number answer.

TEACHER NOTES:
Different numbers from the We Do, same moves: read the family from the array, then use the inverse for missing numbers.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 has the array drawn and the first fact written. Read the rows, then the columns, then share back.
- Extra Notes: Sit with these students and point to the picture for each fact. This is also the rebuild for any returning student.
EXTENDING PROMPT:
- Task: The extension box asks students to use the inverse to check a division and to write their own missing-number puzzle for a partner.
- Extra Notes: Students who are ready can move on to the Year 8 extension sheet on solving equations with inverse operations.

WATCH FOR:
- Students who use the inverse fluently - secure.
- Students who guess and check - prompt them back to the matching times or divide fact.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- Solve 270 divided by 30 by thinking about multiplication.
- Write the multiplication fact you used and the answer.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - use the inverse to solve a division. Look for 30 times 9 is 270, so 270 divided by 30 is 9.

WATCH FOR:
- Students who write the times fact and the answer 9 - secure.
- Students who attempt long division and stall - revisit the inverse at the start of Lesson 2.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: how does knowing 6 times 7 is 42 help you work out 42 divided by 6?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that multiplication and division are inverses, so one helps with the other. Students who can say this are ready for the properties of operations next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on the core criterion - small group revision at the start of Lesson 2.

[General: Closing | Element: Retention and Recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Multiplication and division are linked",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - fractions & decimals
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions & decimals",
      [
        "Write 0.7 as a fraction in tenths.",
        "Which is larger: 0.6 or 0.45?",
        "Order from smallest: 0.5, 0.05, 0.55",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "7/10          0.6 is larger          0.05, 0.5, 0.55", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal x 1-digit
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal x single digit",
      ["4.6 x 3", "2.7 x 6", "3.5 x 4"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "13.8        16.2        14.0", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - one small array, four facts (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "One array, four facts",
    [
      "2 rows of 5 counters.",
      "Across: 2 x 5 = 10.",
      "Down: 5 x 2 = 10.",
      "Share back: 10 ÷ 2 = 5, 10 ÷ 5 = 2.",
      "",
      "Same picture. Four facts.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.ACCENT });
      slide.addText("2 rows of 5", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      drawArray(slide, lg.rightX + 1.15, lg.panelTopPadded + 0.55, 2, 5, 0.40);
      slide.addText("2 x 5 = 10        5 x 2 = 10", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.75, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      slide.addText("10 ÷ 2 = 5        10 ÷ 5 = 2", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.10, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning that multiplication and division are linked, so we can use one to work out the other.",
    [
      "I can write the family of facts from an array.",
      "I can use the inverse to find a missing number in a multiplication or division.",
      "I can explain how using the inverse makes a calculation easier.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Array = equal rows and columns",
      "Inverse = the operation that undoes another",
      "Family of facts = the 2 x and 2 ÷ facts from one array",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.SECONDARY });
      slide.addText("x  and  ÷  are inverses", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.32,
        fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [["3 x 4 = 12", C.PRIMARY], ["12 ÷ 4 = 3", C.SECONDARY]];
      const ry0 = lg.panelTopPadded + 0.62;
      rows.forEach((r, i) => {
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.7, y: ry0 + i * 0.66, w: lg.rightW - 1.4, h: 0.52, rectRadius: 0.07,
          fill: { color: r[1] },
        }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
      slide.addText("One undoes the other.", {
        x: lg.rightX + 0.15, y: ry0 + 1.42, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 10: I Do #1 - array 4x6 -> four facts
  workedExSlide(pres, 2, "I Do", "One array, four facts: 4 x 6",
    [
      "4 rows of 6.",
      "Across the rows: 4 x 6 = 24.",
      "Down the columns: 6 x 4 = 24.",
      "Share back into 4 rows: 24 ÷ 4 = 6.",
      "Share back into 6 columns: 24 ÷ 6 = 4.",
      "",
      "One array. Four facts.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.PRIMARY });
      slide.addText("4 rows of 6 = 24", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      drawArray(slide, lg.rightX + 0.75, lg.panelTopPadded + 0.52, 4, 6, 0.33);
      const facts = [
        ["4 x 6 = 24", C.PRIMARY], ["6 x 4 = 24", C.PRIMARY],
        ["24 ÷ 4 = 6", C.SECONDARY], ["24 ÷ 6 = 4", C.SECONDARY],
      ];
      const fy = lg.panelTopPadded + 2.28;
      facts.forEach((f, i) => {
        const col = i % 2, row = Math.floor(i / 2);
        addTextOnShape(slide, f[0], {
          x: lg.rightX + 0.20 + col * 1.95, y: fy + row * 0.52, w: 1.8, h: 0.44, rectRadius: 0.06,
          fill: { color: f[1] },
        }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slide 11: I Do #2 - the inverse makes division easier
  workedExSlide(pres, 2, "I Do", "Use the inverse for harder ones",
    [
      "Stuck on 240 ÷ 20?",
      "Ask the times question: 20 x ? = 240.",
      "20 x 12 = 240, so 240 ÷ 20 = 12.",
      "",
      "Stuck on 17 x ? = 221?",
      "Divide instead: 221 ÷ 17 = 13.",
      "",
      "When one is tricky, swap to its inverse.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("Swap to the inverse", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      const pairs = [
        ["240 ÷ 20 = ?", "20 x 12 = 240", "so 240 ÷ 20 = 12"],
        ["17 x ? = 221", "221 ÷ 17 = 13", "so the ? is 13"],
      ];
      const py0 = lg.panelTopPadded + 0.55;
      pairs.forEach((p, i) => {
        const py = py0 + i * 1.30;
        addTextOnShape(slide, p[0], {
          x: lg.rightX + 0.30, y: py, w: lg.rightW - 0.60, h: 0.46, rectRadius: 0.07,
          fill: { color: C.ALERT },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
        addTextOnShape(slide, p[1], {
          x: lg.rightX + 0.30, y: py + 0.52, w: lg.rightW - 0.60, h: 0.42, rectRadius: 0.07,
          fill: { color: C.SECONDARY },
        }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(p[2], {
          x: lg.rightX + 0.30, y: py + 0.96, w: lg.rightW - 0.60, h: 0.28,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      });
    }
  );

  // Slides 12-13: CFU + reveal - family of facts from 3x5
  withReveal(
    () => cfuSlide(pres, "CFU", "Write the family of facts",
      { technique: "Show Me Boards",
        question: "Here is an array of 3 rows of 5.\n\nWrite all four facts: two x facts and two ÷ facts." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3 x 5 = 15    5 x 3 = 15    15 ÷ 3 = 5    15 ÷ 5 = 3", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 8 x ? = 56
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Find the missing number: 8 x ? = 56",
      [
        "With your partner.",
        "",
        "1.  Do not guess and check.",
        "2.  Use the inverse: 56 ÷ 8.",
        "3.  Read the missing number.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
        slide.addText("8 x ? = 56", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.40,
          fontSize: 26, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addFactTriangle(slide, lg.rightX + 0.45, lg.panelTopPadded + 0.62, lg.rightW - 0.9,
          56, 8, "?", C.SECONDARY);
      }),
    (slide) => {
      addTextOnShape(slide, "56 ÷ 8 = 7, so 8 x 7 = 56. The missing number is 7.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - 180 ÷ 30
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "A bigger one: 180 ÷ 30",
      [
        "With your partner.",
        "",
        "Swap to the inverse.",
        "30 x ? = 180.",
        "Which times fact gives 180?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
        slide.addText("180 ÷ 30", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.40,
          fontSize: 28, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addFactTriangle(slide, lg.rightX + 0.45, lg.panelTopPadded + 0.62, lg.rightW - 0.9,
          180, 30, "?", C.SECONDARY);
      }),
    (slide) => {
      addTextOnShape(slide, "30 x 6 = 180, so 180 ÷ 30 = 6.", {
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
      { text: "read the fact families from the arrays.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "use the inverse to find missing numbers.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "explain how a times fact helps a divide.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Multiply and divide are inverses. When one is tricky, swap to the other.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    // A small reference array so the page is anchored, not text-only.
    s.addText("Reference array:  3 x 5 = 15", {
      x: 3.0, y: panelY + 1.02, w: 4.0, h: 0.28,
      fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });
    drawArray(s, 4.15, panelY + 1.34, 3, 5, 0.30);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Solve 270 ÷ 30 by thinking about multiplication.",
      "Write the multiplication fact you used and the answer.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how does knowing 6 x 7 = 42 help you work out 42 ÷ 6?",
      scItems: [
        "I can write the family of facts from an array.",
        "I can use the inverse to find a missing number in a multiplication or division.",
        "I can explain how using the inverse makes a calculation easier.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FourP_Lesson1_Multiplication_Division_Links.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Write families of facts from arrays and use the inverse to find missing numbers.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "One array tells a family of four facts: two multiplication and two division. Multiply and divide are inverses - when a division is tricky, ask the matching times question.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "An array of 3 rows of 4 makes 12. Across: 3 x 4 = 12. Down: 4 x 3 = 12. Share back: 12 ÷ 3 = 4 and 12 ÷ 4 = 3. One array, four facts.",
      y);

    y = addSectionHeading(doc, "Section 1 - Read the fact family (started for you)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Array A: 2 rows of 5 (10 counters).", y);
    y = addWriteLine(doc, "2 x 5 = 10        5 x 2 = ____        10 ÷ 2 = ____        10 ÷ 5 = ____", y);
    y = addBodyText(doc, "Array B: 4 rows of 3 (12 counters).", y);
    y = addWriteLine(doc, "____ x ____ = ____   ____ x ____ = ____   ____ ÷ ____ = ____   ____ ÷ ____ = ____", y);

    y = addSectionHeading(doc, "Section 2 - Find the missing number (use the inverse)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  7 x ____ = 63", y);
    y = addWriteLine(doc, "b)  ____ x 6 = 54", y);
    y = addWriteLine(doc, "c)  96 ÷ ____ = 12", y);
    y = addWriteLine(doc, "d)  240 ÷ 20 = ____", y);

    y = addSectionHeading(doc, "Section 3 - Explain", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "How does knowing 8 x 9 = 72 help you work out 72 ÷ 8?", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "Use the inverse to CHECK: is 144 ÷ 12 = 12 correct? Show the times fact. __________", y);
    y = addWriteLine(doc, "Write your own missing-number puzzle for a partner: ____ x ____ = ____", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Multiplication and division are linked | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 1 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Fact families", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Array A: 5 x 2 = 10, 10 ÷ 2 = 5, 10 ÷ 5 = 2.", y);
    y = addBodyText(doc, "Array B: 4 x 3 = 12, 3 x 4 = 12, 12 ÷ 4 = 3, 12 ÷ 3 = 4.", y);

    y = addSectionHeading(doc, "Section 2 - Missing numbers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  9   (63 ÷ 7 = 9)        b)  9   (54 ÷ 6 = 9)        c)  8   (96 ÷ 8 = 12)        d)  12   (20 x 12 = 240)", y);

    y = addSectionHeading(doc, "Section 3 - Explain", y, { color: C.PRIMARY });
    y = addBodyText(doc, "8 x 9 = 72 is the inverse of 72 ÷ 8. Because 8 nines make 72, sharing 72 into 8 groups gives 9. So 72 ÷ 8 = 9.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "144 ÷ 12 = 12 is correct because 12 x 12 = 144. Puzzles vary; check the multiplication balances.", y);

    y = addTipBox(doc,
      "Watch for: students who write only the two multiplication facts; students who guess and check instead of using the inverse; students who start long division on 240 ÷ 20 instead of asking 20 x ? = 240.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Using inverse operations to solve equations.",
      color: C.SECONDARY,
      lessonInfo: `Lesson ${SESSION} | Year 8 challenge | extends Year 6 VC2M5A01`,
    });
    y = addTipBox(doc,
      "The same inverse idea solves equations. To undo an operation, do its inverse to both sides. Multiplying is undone by dividing; adding is undone by subtracting.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Solve 3n = 51. The 3 multiplies n, so divide both sides by 3: n = 51 ÷ 3 = 17. Check: 3 x 17 = 51.",
      y);

    y = addSectionHeading(doc, "Section 1 - One step", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  7n = 91,  n = ______", y);
    y = addWriteLine(doc, "b)  n ÷ 4 = 12,  n = ______", y);
    y = addWriteLine(doc, "c)  n + 38 = 100,  n = ______", y);

    y = addSectionHeading(doc, "Section 2 - Two steps", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2n + 5 = 31,  n = ______   (undo the + 5 first, then the x 2)", y);
    y = addWriteLine(doc, "b)  3n - 7 = 26,  n = ______", y);
    y = addWriteLine(doc, "c)  (n ÷ 5) + 4 = 13,  n = ______", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Explain why you undo the + 5 before the x 2 when solving 2n + 5 = 31.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) 13  b) 48  c) 62.   S2  a) 13  b) 11  c) 45.   S3  Work backwards through the order of operations: the last thing done to n was + 5, so undo it first, then undo the x 2.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Year 8 Extension | Inverse operations`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
