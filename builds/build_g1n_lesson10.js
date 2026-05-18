"use strict";

// Grade 1 Numeracy — Lesson 10: Money Story Problems
// AC9M1N05 — modelling simple money problems involving addition and subtraction
//            using whole dollar amounts; e.g. "I had $14 and was given $15 for my
//            birthday".
// Daily Review: Counting and place value (addition story, subtraction story, partition).
// Fluency:      Skip counting, addition and patterns (count by 5s, +10).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade1", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  workedExSlide, exitTicketSlide,
  addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  addInstructionCard,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 10;
const TOTAL = 10;
const UNIT_TITLE = "Numbers to 120";
const LESSON_TITLE = "Money Stories: Shop Problems";
const FOOTER = `${UNIT_TITLE} | Lesson ${SESSION} of ${TOTAL} | Year 1 Numeracy`;
const OUT_DIR = `output/G1N_Lesson${SESSION}_Money_Stories`;
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 10 Shop Stories",
  "Solve money stories using the Think Board. Whole dollars only.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 10 Answer Key",
  "Teacher reference for the shop stories.");
const EXTENDING_RES = makeSessionResource(SESSION, "Lesson 10 Extension",
  "Open a class shop - design your own price tags and solve customer problems.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

function drawNumeralCard(slide, x, y, w, h, numeral, opts) {
  const o = opts || {};
  const fill = o.fill || C.WHITE;
  const stroke = o.stroke || C.PRIMARY;
  const textColor = o.color || C.PRIMARY;
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.14,
    fill: { color: fill },
    line: { color: stroke, width: 2.5 },
  });
  slide.addText(String(numeral), {
    x, y, w, h,
    fontSize: o.fontSize || 60, fontFace: FONT_H, color: textColor,
    bold: true, align: "center", valign: "middle", margin: 0,
  });
}

// Price tag — a tag-shape with $N inside
function drawPriceTag(slide, x, y, w, h, amount, opts) {
  const o = opts || {};
  const fill = o.fill || C.ALERT;
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.15,
    fill: { color: fill },
    line: { color: C.CHARCOAL, width: 2 },
  });
  // small hole/circle on the left
  slide.addShape("roundRect", {
    x: x + 0.1, y: y + h / 2 - 0.08, w: 0.16, h: 0.16, rectRadius: 0.08,
    fill: { color: C.WHITE },
  });
  slide.addText(`$${amount}`, {
    x: x + 0.25, y, w: w - 0.25, h,
    fontSize: o.fontSize || 38, fontFace: FONT_H, color: C.WHITE,
    bold: true, align: "center", valign: "middle", margin: 0,
  });
}

function drawThinkBoard(slide, x, y, w, h, opts) {
  const o = opts || {};
  const cellW = w / 2;
  const cellH = h / 2;
  const cells = [
    { x: x, y: y, label: "Story", colour: C.PRIMARY },
    { x: x + cellW, y: y, label: "Picture", colour: C.SECONDARY },
    { x: x, y: y + cellH, label: "Counters or coins", colour: C.ACCENT },
    { x: x + cellW, y: y + cellH, label: "Number sentence", colour: C.SUCCESS },
  ];
  cells.forEach((cell, idx) => {
    slide.addShape("rect", {
      x: cell.x, y: cell.y, w: cellW, h: cellH,
      fill: { color: C.WHITE },
      line: { color: C.CHARCOAL, width: 1.5 },
    });
    slide.addShape("rect", {
      x: cell.x, y: cell.y, w: cellW, h: 0.32,
      fill: { color: cell.colour },
    });
    slide.addText(cell.label, {
      x: cell.x + 0.08, y: cell.y, w: cellW - 0.16, h: 0.32,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true, valign: "middle", margin: 0,
    });
    if (o.contents && o.contents[idx]) {
      slide.addText(o.contents[idx], {
        x: cell.x + 0.12, y: cell.y + 0.38, w: cellW - 0.24, h: cellH - 0.46,
        fontSize: o.fontSize || 14, fontFace: FONT_B, color: C.CHARCOAL,
        margin: 0, valign: "top", fit: "shrink", shrinkText: true,
      });
    }
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we open a class shop!
- We use real maths to buy and sell things
- Money problems use the SAME Think Board we know

DO:
- Display the title
- Set up a small shop area at the front with 3-4 items priced
- Have play money in whole dollar notes

TEACHER NOTES:
Lesson 10 of 10 - the unit close. Money brings everything together. Whole dollar amounts only.

WATCH FOR:
- Excitement about the shop - lean in

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today we use play money - dollar notes
- Shop items at the front
- Think Board worksheet again

DO:
- Set up shop items with price tags
- Hand out play money to each table

TEACHER NOTES:
The role play is essential. Even a simple setup with 4 items boosts engagement.

WATCH FOR:
- Materials ready

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_LAUNCH = `SAY:
- Birthday story - I had $5. My grandma gave me $3
- How much money do I have now?

DO:
- Display the launch story with $5 and $3 note pictures
- Partner talk
- Cold call

TEACHER NOTES:
The birthday context is universal. Answer is $8.

WATCH FOR:
- Students who answer $8 - they have it
- Students unsure - reassure, money is just like counting

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_DR1_Q = `SAY:
- Addition. 5 + 4. How many altogether?

DO:
- Show me

TEACHER NOTES:
Retrieves Lesson 8.

WATCH FOR:
- 9

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- 9

DO:
- Reveal 9

TEACHER NOTES:
Quick.

WATCH FOR:
- Confident

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- Subtraction story - 10 take away 3. How many left?

DO:
- Show me

TEACHER NOTES:
Retrieves Lesson 9.

WATCH FOR:
- 7

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- 7. 10 - 3 = 7

DO:
- Reveal 7

TEACHER NOTES:
Brief.

WATCH FOR:
- Confident

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- Partition - 25. How many tens?

DO:
- Show me

TEACHER NOTES:
Retrieves Lesson 6.

WATCH FOR:
- 2

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- 2 tens, 5 ones

DO:
- Reveal 2

TEACHER NOTES:
Brief.

WATCH FOR:
- Chorus

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1 = `SAY:
- Count by 5s. Money often comes in 5s
- 5, 10, 15, 20, 25, 30

DO:
- Lead choral count of 5s to 30

TEACHER NOTES:
Counting by 5s connects money (5 dollar notes) to skip counting.

WATCH FOR:
- Hesitations at 25-30

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL2 = `SAY:
- 10 more quick fire. We have practised this all unit
- 14 - say it - 24. 27 - say it - 37

DO:
- Quick fire: 14, 27, 33, 41, 58
- Class +10 each

TEACHER NOTES:
+10 is the most-used fluency in the unit.

WATCH FOR:
- Confident chorus

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- A new symbol - the dollar sign
- We write $ before the number
- $5 means five dollars
- $14 means fourteen dollars

DO:
- Display $5 large
- Write $14 underneath

TEACHER NOTES:
The $ sign is the new symbol. Students should write $ before the numeral, not after.

WATCH FOR:
- Students who write 5$ - reteach

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_LI_SC = `SAY:
- Read with me - We are learning to solve money stories
- I can read a money story
- I can show the dollars with counters or pictures
- I can write a number sentence with the dollar sign

DO:
- Choral read

TEACHER NOTES:
SC3 introduces the $ sign in the number sentence.

WATCH FOR:
- Engaged class - we are close to the unit close

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_BIRTHDAY = `SAY:
- Watch me. STORY: I had $14. Grandma gave me $15 for my birthday
- PICTURE: $14 note and $15 note
- COUNTERS: 14 + 15
- NUMBER SENTENCE: $14 + $15 = $29
- I have $29 altogether

DO:
- Build Think Board with money story
- Show $14 + $15 with play money or drawings
- Write $14 + $15 = $29
- Underline $29 - "this is the answer"

TEACHER NOTES:
This is the curriculum example. Take time. Whole-dollar addition is the lesson's main skill.

WATCH FOR:
- Students who follow the count - good

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_SHOP = `SAY:
- Shop story. I have $10. A teddy costs $4. I buy it
- How much money do I have LEFT?
- $10 minus $4 equals $6
- I have $6 left

DO:
- Show $10 and a teddy with $4 price tag
- Take 4 away from 10 with counters
- Write $10 - $4 = $6

TEACHER NOTES:
Subtraction with money - paying for an item. Same Think Board, same number sense.

WATCH FOR:
- Students who recognise "left" from Lesson 9 - praise

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. STORY: A book costs $6. A pencil costs $2
- How much do they cost altogether?
- Write the answer with the $ sign

DO:
- Display the two prices
- 12 seconds, show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Two prices. How much altogether? Write the answer with the dollar sign. Show me"
- Scan for: $8
PROCEED: If 80%+ show $8, move to We Do.
PIVOT: Most likely misconception - students write 8 without the $ sign. Reteach with the symbol. Re-check with "$3 + $5 = ?"

TEACHER NOTES:
This CFU also tests the $ symbol usage.

WATCH FOR:
- $8 = secure
- 8 only = needs $ symbol practice

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- $8 altogether. $6 + $2 = $8

DO:
- Reveal $8

TEACHER NOTES:
Brief.

WATCH FOR:
- Confident class

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn. STORY: I have $20. A toy car costs $7. I buy it
- How much do I have LEFT?
- Use the Think Board

DO:
- Display story
- Allow 60 seconds
- Cold call one pair

TEACHER NOTES:
Subtraction with the answer crossing back below 20. Use the count-back strategy or known facts.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use real play money. Count out $20. Pay $7. Count what is left.
- Extra Notes: Small group at front with the teacher and a $1 note set.
EXTENDING PROMPT:
- Task: Write a similar story with bigger numbers - $25 - $8 = ?
- Extra Notes: Whole-class extension.

WATCH FOR:
- Students who can count back from 20 - they have the strategy
- Students who count up from 7 to 20 - alternative valid strategy

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- $13 left. $20 - $7 = $13

DO:
- Reveal $13

TEACHER NOTES:
Confirm with the count.

WATCH FOR:
- Confident class

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge check. I have $5 saved. Mum gives me $10. Dad gives me $5
- How much do I have altogether? Write with the $ sign

DO:
- Display the story
- 20 seconds, show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Add up the dollars. Show me with the $ sign"
- Scan for: $20
PROCEED: If 80%+ show $20, move to You Do.
PIVOT: Most likely misconception - students forget the saved $5 (answer $15) or write 20 without $. Reteach with three piles of money. Re-check with "$6 + $4 + $5".

TEACHER NOTES:
This hinge has THREE numbers - the unit's highest demand.

WATCH FOR:
- $20 = unit success
- $15 = missed one number; small group review

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- $20 altogether. $5 + $10 + $5 = $20

DO:
- Reveal $20

TEACHER NOTES:
Final hinge of the whole unit.

WATCH FOR:
- Strong response = whole unit landed

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Last independent work of the unit!
- Take the worksheet - Shop Stories
- Read each story. Use the Think Board. Write the answer with the $ sign

DO:
- Distribute worksheet
- Circulate - watch for missing $ signs

TEACHER NOTES:
Worksheet has 2 shop stories. One addition, one subtraction.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use real play money. Build each story with notes alongside the worksheet.
- Extra Notes: Small group at back.
EXTENDING PROMPT:
- Task: Lesson 10 Extension - open a class shop with your own price tags.
- Extra Notes: Distribute Lesson 10 Extension PDF.

WATCH FOR:
- Students missing the $ sign - reteach
- Students who can solve both stories - they have closed the unit well

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Final exit ticket of the unit!
- STORY: A book costs $8. A pencil costs $3. How much altogether?
- Write your answer with the dollar sign

DO:
- Students write
- Collect

TEACHER NOTES:
Final unit exit ticket. Should land at $11. Sort into $11, 11 (no $), other.

WATCH FOR:
- $11 = unit complete

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Last lesson - check our success criteria
- I can read a money story - thumbs
- I can show with counters or pictures - thumbs
- I can write a number sentence with the $ sign - thumbs
- Turn and tell - what is the biggest number we have used in this unit?

DO:
- Thumbs check
- Partner talk - reflect on the whole unit

TEACHER NOTES:
The closing also closes the unit. Students should remember the journey from numerals to 120 to money problems.

WATCH FOR:
- Students who name 120 or higher - they have the full unit

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, LESSON_TITLE, "Lesson 10 of 10 - Numbers to 120", "Year 1 Numeracy | AC9M1N05", NOTES_TITLE);
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Launch — birthday $14 + $15
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "Birthday money story", { color: STAGE_COLORS["1"] });

    addInstructionCard(s, [
      { role: "header", text: "Story" },
      { role: "body", text: "I had $5." },
      { role: "body", text: "My grandma gave me $3 for my birthday." },
      { role: "emphasis", text: "How much money do I have now?" },
    ], { x: 1.0, y: 1.8, w: 8, h: 2.7, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // DR1 — 5 + 4
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Addition", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 1.6, 2.0, "5", { fontSize: 100 });
      s.addText("+", { x: 4.2, y: 1.9, w: 0.7, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, 4.9, 1.9, 1.6, 2.0, "4", { fontSize: 100 });
      s.addText("= ?", { x: 6.6, y: 1.9, w: 1.8, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "9", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // DR2 — 10 - 3
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Subtraction", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 1.6, 2.0, "10", { fontSize: 90 });
      s.addText("-", { x: 4.2, y: 1.9, w: 0.7, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, 4.9, 1.9, 1.6, 2.0, "3", { fontSize: 100 });
      s.addText("= ?", { x: 6.6, y: 1.9, w: 1.8, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "7", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR2_A);
    }
  );

  // DR3 — Tens in 25
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many tens in 25?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.5, "25", { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "2 tens", {
        x: 3.0, y: 4.55, w: 4.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // Fluency 1 — Count by 5s
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Count by 5s", { color: STAGE_COLORS["1"] });

    const fives = [5, 10, 15, 20, 25, 30];
    fives.forEach((n, i) => {
      const x = 0.5 + i * 1.55;
      drawNumeralCard(s, x, 2.0, 1.4, 1.4, String(n), { fontSize: 44, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });
    });

    addInstructionCard(s, [
      { role: "header", text: "Count with me - 5, 10, 15..." },
    ], { x: 0.5, y: 4.0, w: 9, h: 1.1, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL1);
  })();

  // Fluency 2 — +10 quick fire
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "+10 quick fire", { color: STAGE_COLORS["1"] });

    const examples = [
      { from: "14", to: "24" },
      { from: "27", to: "37" },
      { from: "33", to: "43" },
      { from: "58", to: "68" },
    ];
    examples.forEach((ex, i) => {
      const x = 0.6 + (i % 2) * 4.5;
      const y = 1.8 + Math.floor(i / 2) * 1.3;
      drawNumeralCard(s, x, y, 1.4, 1.1, ex.from, { fontSize: 48 });
      s.addText("+10", { x: x + 1.5, y: y, w: 0.7, h: 1.1, fontSize: 20, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, x + 2.3, y, 1.4, 1.1, ex.to, { fontSize: 48, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });
    });

    addInstructionCard(s, [
      { role: "header", text: "Add 10 each time" },
    ], { x: 0.5, y: 4.5, w: 9, h: 0.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL2);
  })();

  // Vocabulary — Dollar sign
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Symbol", { color: C.PRIMARY });
    addTitle(s, "The dollar sign  $", { color: C.PRIMARY });

    drawPriceTag(s, 1.5, 1.9, 2.0, 1.2, "5", { fontSize: 56 });
    drawPriceTag(s, 4.0, 1.9, 2.0, 1.2, "14", { fontSize: 56, fill: C.SECONDARY });
    drawPriceTag(s, 6.5, 1.9, 2.0, 1.2, "29", { fontSize: 56, fill: C.SUCCESS });

    addInstructionCard(s, [
      { role: "header", text: "The $ goes BEFORE the number" },
      { role: "body", text: "$5 means five dollars." },
      { role: "body", text: "$14 means fourteen dollars." },
    ], { x: 0.5, y: 3.6, w: 9, h: 1.5, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // LI/SC
  liSlide(pres,
    ["We are learning to solve money stories."],
    [
      "I can read a money story.",
      "I can show the dollars with counters or pictures.",
      "I can write a number sentence with the dollar sign.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do — Birthday addition $14 + $15
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Birthday: $14 + $15", { color: STAGE_COLORS["2"] });

    drawThinkBoard(s, 0.5, 1.7, 9.0, 3.0, {
      contents: [
        "I had $14. Grandma gave me $15 for my birthday. How much altogether?",
        "[Draw $14 note + $15 note]",
        "Count out 14 + 15 in $1 notes",
        "$14 + $15 = $29",
      ],
      fontSize: 12,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_BIRTHDAY);
  })();

  // I Do — Shop subtraction $10 - $4
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Shop: $10 buy a $4 teddy", { color: STAGE_COLORS["2"] });

    drawThinkBoard(s, 0.5, 1.7, 9.0, 3.0, {
      contents: [
        "I have $10. A teddy costs $4. I buy it. How much LEFT?",
        "[Draw $10 - cross out $4 worth]",
        "10 counters - take 4 away",
        "$10 - $4 = $6",
      ],
      fontSize: 12,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_SHOP);
  })();

  // CFU — $6 + $2
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Book and pencil - total?", { color: C.ALERT });

      drawPriceTag(s, 2.0, 1.9, 2.0, 1.4, "6", { fontSize: 60 });
      drawPriceTag(s, 6.0, 1.9, 2.0, 1.4, "2", { fontSize: 60, fill: C.SECONDARY });
      s.addText("book", { x: 2.0, y: 3.4, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });
      s.addText("pencil", { x: 6.0, y: 3.4, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });

      addInstructionCard(s, [
        { role: "header", text: "Total cost?" },
        { role: "body", text: "Write the answer with the $ sign." },
      ], { x: 0.5, y: 4.0, w: 9, h: 1.1, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "$8 - $6 + $2 = $8!", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // We Do — $20 - $7 toy car
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "I have $20, toy car is $7", { color: STAGE_COLORS["3"] });

      drawThinkBoard(s, 0.5, 1.65, 9.0, 3.1, {
        contents: [
          "I have $20. A toy car costs $7. I buy it. How much LEFT?",
          "",
          "",
          "",
        ],
        fontSize: 12,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "$13 - $20 - $7 = $13!", {
        x: 2.0, y: 4.85, w: 6.0, h: 0.25, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge — $5 + $10 + $5
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "How much altogether?", { color: C.ALERT });

      addInstructionCard(s, [
        { role: "header", text: "Story" },
        { role: "body", text: "I have $5 saved." },
        { role: "body", text: "Mum gives me $10. Dad gives me $5." },
        { role: "emphasis", text: "How much do I have altogether?" },
      ], { x: 1.0, y: 1.85, w: 8, h: 2.5, strip: C.ALERT });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer with the $ sign." },
      ], { x: 0.5, y: 4.4, w: 9, h: 0.7, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "$20 - $5 + $10 + $5 = $20!", {
        x: 1.0, y: 4.55, w: 8.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Shop Stories",
    [
      "Read each shop story.",
      "Use the Think Board.",
      "Draw the money or count out coins.",
      "Write the answer with the $ sign.",
      "",
      "Two stories to solve!",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.ALERT });
      slide.addText("Remember", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "$ goes before the number.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Read the story twice.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Altogether = +, Left = -", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Exit Ticket — final
  exitTicketSlide(pres,
    ["A book costs $8. A pencil costs $3. How much altogether? Use the $ sign."],
    NOTES_EXIT, FOOTER);

  // Closing
  closingSlide(pres, {
    reflectionPrompt: "Turn and tell: what is the biggest number we have used this unit?",
    scItems: [
      "I can read a money story.",
      "I can show the dollars with counters or pictures.",
      "I can write a number sentence with the dollar sign.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1N_Lesson10_Money_Stories.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  function hex(c) { return c.startsWith("#") ? c : "#" + c; }

  function drawPdfThinkBoard(doc, x, y, w, h, story) {
    const cellW = w / 2;
    const cellH = h / 2;
    const cells = [
      { label: "Story", color: C.PRIMARY, content: story || "" },
      { label: "Picture", color: C.SECONDARY, content: "" },
      { label: "Counters or coins", color: C.ACCENT, content: "" },
      { label: "Number sentence", color: C.SUCCESS, content: "" },
    ];
    cells.forEach((cell, idx) => {
      const cx = x + (idx % 2) * cellW;
      const cy = y + Math.floor(idx / 2) * cellH;
      doc.rect(cx, cy, cellW, cellH).lineWidth(1.5).strokeColor("#444").stroke();
      doc.save();
      doc.rect(cx, cy, cellW, 18).fill(hex(cell.color));
      doc.restore();
      doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
      doc.text(cell.label, cx + 6, cy + 4);
      if (cell.content) {
        doc.fontSize(10).font("Sans").fillColor("#000");
        doc.text(cell.content, cx + 6, cy + 26, { width: cellW - 12 });
      }
    });
  }

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Shop stories with money",
      color: C.NAVY,
      lessonInfo: "Lesson 10 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Read each money story. Use the Think Board. Use the dollar sign $ in your number sentence.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Shop Story 1: Adding up", y, { color: C.NAVY });
    drawPdfThinkBoard(doc, 60, y, 480, 280, "I have $7 saved. Nan gives me $4 for helping. How much money altogether? Use the dollar sign $.");
    y += 295;

    y = addSectionHeading(doc, "Shop Story 2: Buying", y, { color: C.NAVY });
    drawPdfThinkBoard(doc, 60, y, 480, 280, "I have $15. A book costs $6. I buy it. How much money is LEFT?");

    addPdfFooter(doc, "Lesson 10 | Numbers to 120 | Year 1");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Shop Stories - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Lesson 10 of 10 | Year 1 Numeracy",
    });
    y = addSectionHeading(doc, "Story 1: $7 + $4", y, { color: C.NAVY });
    y = addBodyText(doc, "Picture: $7 + $4 (any reasonable money drawing)", y);
    y = addBodyText(doc, "Counters: 7 + 4 dollar notes", y);
    y = addBodyText(doc, "Number sentence: $7 + $4 = $11", y);
    y = addBodyText(doc, "Answer: $11 altogether", y);
    y = addSectionHeading(doc, "Story 2: $15 - $6", y, { color: C.NAVY });
    y = addBodyText(doc, "Picture: $15 with $6 crossed out", y);
    y = addBodyText(doc, "Counters: 15 - 6", y);
    y = addBodyText(doc, "Number sentence: $15 - $6 = $9", y);
    y = addBodyText(doc, "Answer: $9 left", y);
    y = addBodyText(doc, "Watch for: missing $ sign in the number sentence. Reteach the symbol position.", y);
    addPdfFooter(doc, "Lesson 10 | Answer Key | Year 1");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension - class shop with own tags
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Open your own class shop",
      color: C.TEAL,
      lessonInfo: "Lesson 10 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Design your own price tags. Then write 2 shop stories using your prices.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Step 1 - Design 4 price tags", y, { color: C.NAVY });
    // Draw 4 tag boxes for students to fill in
    const tagX = 60;
    const tagW = 110;
    const tagH = 60;
    const tagGap = 20;
    for (let i = 0; i < 4; i += 1) {
      const cx = tagX + i * (tagW + tagGap);
      doc.rect(cx, y, tagW, tagH).lineWidth(1.5).strokeColor("#999").stroke();
      doc.fontSize(10).font("Sans").fillColor("#777");
      doc.text("Item:", cx + 6, y + 6);
      doc.text("$ _____", cx + 6, y + 30);
    }
    y += tagH + 25;

    y = addSectionHeading(doc, "Step 2 - Write a shop story (addition)", y, { color: C.NAVY });
    doc.rect(60, y, 480, 70).lineWidth(1).strokeColor("#999").stroke();
    y += 80;

    y = addSectionHeading(doc, "Step 3 - Write a shop story (take-away)", y, { color: C.NAVY });
    doc.rect(60, y, 480, 70).lineWidth(1).strokeColor("#999").stroke();
    y += 80;

    addPdfFooter(doc, "Lesson 10 | Extension | Year 1");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
