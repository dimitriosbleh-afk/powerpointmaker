"use strict";

// Grade 1 Numeracy — Lesson 8: Think Board - Addition Story Problems
// AC9M1N05 — mathematical modelling of addition stories using a Think Board.
// Daily Review: Counting and place value (pairs to 10, tens-ones).
// Fluency:      Skip counting, addition and patterns (doubles, +1 add).

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
  addTensFrame,
  STAGE_COLORS,
} = T;

const SESSION = 8;
const TOTAL = 10;
const UNIT_TITLE = "Numbers to 120";
const LESSON_TITLE = "Think Board: Addition Stories";
const FOOTER = `${UNIT_TITLE} | Lesson ${SESSION} of ${TOTAL} | Year 1 Numeracy`;
const OUT_DIR = `output/G1N_Lesson${SESSION}_Think_Board_Addition`;
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 8 Addition Think Board",
  "Solve two addition stories using the four-square Think Board.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 8 Answer Key",
  "Teacher reference with completed Think Boards.");
const EXTENDING_RES = makeSessionResource(SESSION, "Lesson 8 Extension",
  "Write your own addition story and solve it on a Think Board.");
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

// Think Board: 4 quadrants - Story, Picture, Sentence, Materials (or labels)
function drawThinkBoard(slide, x, y, w, h, opts) {
  const o = opts || {};
  const cellW = w / 2;
  const cellH = h / 2;
  const cells = [
    { x: x, y: y, label: "Story", colour: C.PRIMARY },
    { x: x + cellW, y: y, label: "Picture", colour: C.SECONDARY },
    { x: x, y: y + cellH, label: "Counters", colour: C.ACCENT },
    { x: x + cellW, y: y + cellH, label: "Number sentence", colour: C.SUCCESS },
  ];
  cells.forEach((cell, idx) => {
    slide.addShape("rect", {
      x: cell.x, y: cell.y, w: cellW, h: cellH,
      fill: { color: C.WHITE },
      line: { color: C.CHARCOAL, width: 1.5 },
    });
    // Label strip at top of cell
    slide.addShape("rect", {
      x: cell.x, y: cell.y, w: cellW, h: 0.32,
      fill: { color: cell.colour },
    });
    slide.addText(cell.label, {
      x: cell.x + 0.08, y: cell.y, w: cellW - 0.16, h: 0.32,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true, valign: "middle", margin: 0,
    });
    // Optional content per cell
    if (o.contents && o.contents[idx]) {
      const content = o.contents[idx];
      slide.addText(content, {
        x: cell.x + 0.12, y: cell.y + 0.38, w: cellW - 0.24, h: cellH - 0.46,
        fontSize: o.fontSize || 14, fontFace: FONT_B, color: C.CHARCOAL,
        margin: 0, valign: "top",
        fit: "shrink", shrinkText: true,
      });
    }
  });
  // Also let caller draw into cells
  if (o.drawCell) {
    cells.forEach((cell, idx) => {
      o.drawCell(slide, idx, cell.x + 0.1, cell.y + 0.4, cellW - 0.2, cellH - 0.5);
    });
  }
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we become story solvers!
- Real-life stories with maths inside them
- We use a Think Board - four squares to help us think

DO:
- Display title
- Draw a Think Board on a large sheet at the front
- Have counters and number tracks ready

TEACHER NOTES:
Lesson 8 of 10. Word-problem work is challenging for Year 1 - the language plus the maths together. The Think Board chunks the thinking.

WATCH FOR:
- Students excited by the Think Board structure

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- We need counters and number tracks today
- The Think Board worksheet has 4 boxes for us to fill in

DO:
- Have counters at each table
- Show the Think Board template at the front

TEACHER NOTES:
The Think Board is essential. Print one large copy for class display.

WATCH FOR:
- All materials ready

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_LAUNCH = `SAY:
- Story time - on the bus 6 children sit down
- 3 more children get on
- How many children on the bus now? Tell your partner

DO:
- Display a simple bus picture or sketch on the board
- Partner talk
- Cold call

TEACHER NOTES:
A simple story to warm up the addition thinking. The answer is 9.

WATCH FOR:
- Students who use fingers to count - good strategy
- Students who say 9 instantly - they have the fact

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_DR1_Q = `SAY:
- Pair to 10. 7 plus what?

DO:
- Display 7 + ? = 10
- Show me

TEACHER NOTES:
Retrieves Lesson 5.

WATCH FOR:
- 3

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- 3. 7 + 3 = 10

DO:
- Reveal 3

TEACHER NOTES:
Quick.

WATCH FOR:
- Confident

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- How many tens in 48?

DO:
- Display 48
- Show me

TEACHER NOTES:
Retrieves Lesson 6.

WATCH FOR:
- 4

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- 4 tens, 8 ones

DO:
- Reveal 4

TEACHER NOTES:
Brief.

WATCH FOR:
- Chorus

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- 10 more than 47?

DO:
- 8 seconds, show me

TEACHER NOTES:
Retrieves +10.

WATCH FOR:
- 57

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- 57. 4 tens then 5 tens, 7 ones same

DO:
- Reveal 57

TEACHER NOTES:
Brief.

WATCH FOR:
- Class confident

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1 = `SAY:
- Fluency. Doubles
- 1 + 1, 2 + 2, 3 + 3, 4 + 4, 5 + 5

DO:
- Lead choral count of doubles to 10 + 10
- Use fingers or counters as anchor

TEACHER NOTES:
Doubles fluency is foundational addition. Today's stories use small additions.

WATCH FOR:
- Confident chorus on small doubles
- Hesitation past 5 + 5

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL2 = `SAY:
- Add 1. I say 7, you say 8

DO:
- Quick fire: 7, 9, 12, 18, 24
- Class +1 each

TEACHER NOTES:
+1 is the simplest addition. Useful for the count-on strategy in stories.

WATCH FOR:
- Quick and confident chorus

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- A new word for today - altogether
- Altogether means how many in total
- When we add, we find altogether

DO:
- Show 4 red counters and 3 blue counters
- Say "4 and 3 - how many altogether? 7"

TEACHER NOTES:
"Altogether" is the addition signal word. Students should recognise it.

WATCH FOR:
- Students who repeat the word

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_LI_SC = `SAY:
- Read with me - We are learning to solve addition stories with a Think Board
- I can read an addition story
- I can show the story with counters or a picture
- I can write the number sentence for the story

DO:
- Choral read

TEACHER NOTES:
SC1, SC2, SC3 each map to one quadrant of the Think Board.

WATCH FOR:
- Engaged class

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_BOARD = `SAY:
- I will show you the Think Board
- Four boxes - Story, Picture, Counters, Number sentence
- We start by reading the STORY
- Then we draw a PICTURE
- Then we use COUNTERS to act it out
- Last we write the NUMBER SENTENCE

DO:
- Display the empty Think Board
- Point to each quadrant as you name it

TEACHER NOTES:
This slide introduces the routine. Students need to internalise the four-quadrant structure.

WATCH FOR:
- Students nodding through the four steps

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_STORY1 = `SAY:
- Watch me solve a story
- STORY: 5 ducks are in the pond. 3 more swim in. How many ducks altogether?
- PICTURE: I draw 5 ducks then 3 more
- COUNTERS: 5 red counters, 3 blue counters
- NUMBER SENTENCE: 5 + 3 = 8
- Answer: 8 ducks altogether!

DO:
- Build the Think Board on the camera as you talk
- Read the story slowly
- Draw simple stick-figure ducks
- Place 5 red + 3 blue counters
- Write 5 + 3 = 8

TEACHER NOTES:
This is the lesson's master example. Take your time and fill each quadrant in front of the class.

WATCH FOR:
- Students saying "8" before you do - they have the sum
- Students who watch the whole sequence - good engagement

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. The story is - I have 4 apples. My friend gives me 2 more
- How many apples altogether? Write on your whiteboard

DO:
- Display the story
- 10 seconds, show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Read the story. Work out the answer. Show me"
- Scan for: 6
PROCEED: If 80%+ show 6, move to We Do.
PIVOT: Most likely misconception - students miss the second number or get distracted by the words. Reteach by drawing each apple as you read. Re-check with "3 dogs, 4 more dogs, how many?"

TEACHER NOTES:
First independent CFU. Tests whether students can extract the numbers and add.

WATCH FOR:
- 6 = secure
- 4 or 2 = missed the second number; needs picture support

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- 6 apples altogether. 4 + 2 = 6

DO:
- Reveal 6

TEACHER NOTES:
Brief.

WATCH FOR:
- Confident class

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Read the story together
- 7 birds are on the tree. 4 more land
- Fill in the Think Board with me

DO:
- Display the empty Think Board with the story
- Class reads story together
- Choral fill: picture (draw 7 birds, 4 more), counters (7 + 4), number sentence (7 + 4 = 11)

TEACHER NOTES:
A bigger sum (across 10). Students may need to count on - 8, 9, 10, 11.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work with real counters. Build 7 and add 4 one at a time, counting "8, 9, 10, 11".
- Extra Notes: Small group at the front with the teacher and a number track.
EXTENDING PROMPT:
- Task: Write a similar story with different numbers - 8 birds, 5 more.
- Extra Notes: Whole-class extension.

WATCH FOR:
- Students who count all 11 from 1 - acceptable
- Students who count on from 7 - more efficient; praise this

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- 11 birds. 7 + 4 = 11

DO:
- Reveal 11 on the Think Board
- Show counter representation (a ten frame full + 1 extra)

TEACHER NOTES:
Crossing 10 is a Year 1 skill. Don't skip this.

WATCH FOR:
- Confident class

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge check. Story - on the bus 8 children. 5 more get on. How many altogether?
- Write the answer on your whiteboard

DO:
- 15 seconds, show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Read the story. Write the answer. Show me"
- Scan for: 13
PROCEED: If 80%+ show 13, move to You Do.
PIVOT: Most likely misconception - students write 3 (ignoring the 8) or 5 (the second number alone). Reteach by drawing each child as you read. Re-check with "6 cats, 4 more, how many?"

TEACHER NOTES:
Bigger sum across 10. Use the counter-on strategy.

WATCH FOR:
- 13 = secure
- 3 or 5 = need explicit reread of the story

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- 13 children altogether. 8 + 5 = 13

DO:
- Reveal 13
- Show the count on - 8 then 9, 10, 11, 12, 13

TEACHER NOTES:
Final hinge.

WATCH FOR:
- Confident class = ready

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent work - Think Board time
- Take the worksheet
- Read each story carefully
- Fill in all four boxes - picture, counters, number sentence, answer

DO:
- Distribute worksheet
- Circulate to read stories aloud for any students who need help
- Hand out Extension when ready

TEACHER NOTES:
The worksheet has 2 story problems. Students fill the full Think Board for each.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work with the teacher. The teacher reads the story aloud. The student draws and counts.
- Extra Notes: Small group at back.
EXTENDING PROMPT:
- Task: Lesson 8 Extension - write your own story and solve it.
- Extra Notes: Distribute Lesson 8 Extension PDF.

WATCH FOR:
- Students who skip the picture box - prompt them back
- Students who do all four boxes well - praise and offer Extension

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Story - I have 6 grapes. I get 4 more. How many grapes altogether?

DO:
- Students write answer on slip
- Collect

TEACHER NOTES:
Exit assesses SC3 (number sentence). Sort into 10 correct, partial, other.

WATCH FOR:
- 10 correct = ready for Lesson 9 subtraction

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Check our success criteria
- I can read an addition story - thumbs
- I can show it with counters or a picture - thumbs
- I can write the number sentence - thumbs
- Tell your partner the four boxes on the Think Board

DO:
- Thumbs check
- Partner talk

TEACHER NOTES:
The reflection retrieves the Think Board structure.

WATCH FOR:
- Story, picture, counters, number sentence answers

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, LESSON_TITLE, "Lesson 8 of 10 - Numbers to 120", "Year 1 Numeracy | AC9M1N05", NOTES_TITLE);
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Launch — bus story
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "How many children on the bus?", { color: STAGE_COLORS["1"] });

    addInstructionCard(s, [
      { role: "header", text: "Story" },
      { role: "body", text: "On the bus, 6 children sit down." },
      { role: "body", text: "3 more children get on." },
      { role: "emphasis", text: "How many altogether?" },
    ], { x: 1.0, y: 1.8, w: 8, h: 2.7, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // DR1 — 7 + ? = 10
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Pair to make 10?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 1.6, 2.0, "7", { fontSize: 100 });
      s.addText("+", { x: 4.2, y: 1.9, w: 0.7, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, 4.9, 1.9, 1.6, 2.0, "?", { fontSize: 100, color: C.ALERT, stroke: C.ALERT });
      s.addText("= 10", { x: 6.6, y: 1.9, w: 1.8, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });

      addInstructionCard(s, [
        { role: "body", text: "Write the missing number." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "3", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // DR2 — tens in 48
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many tens in 48?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.5, "48", { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "4 tens", {
        x: 3.0, y: 4.55, w: 4.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR2_A);
    }
  );

  // DR3 — 10 more than 47
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "10 more than 47?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 2.2, 2.0, "47", { fontSize: 120 });
      s.addText("+ 10", { x: 4.8, y: 1.9, w: 1.0, h: 2.0, fontSize: 36, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, 5.9, 1.9, 2.2, 2.0, "?", { fontSize: 120, color: C.ALERT, stroke: C.ALERT });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer. Show me!" },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "57", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // Fluency 1 — Doubles
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Doubles", { color: STAGE_COLORS["1"] });

    const doubles = [
      { from: "1+1", to: "2" },
      { from: "2+2", to: "4" },
      { from: "3+3", to: "6" },
      { from: "4+4", to: "8" },
      { from: "5+5", to: "10" },
    ];
    doubles.forEach((d, i) => {
      const x = 0.55 + i * 1.8;
      drawNumeralCard(s, x, 1.85, 1.55, 1.4, d.from, { fontSize: 28 });
      s.addText("=", { x, y: 3.3, w: 1.55, h: 0.4, fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });
      drawNumeralCard(s, x + 0.15, 3.55, 1.25, 0.8, d.to, { fontSize: 32, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });
    });

    addInstructionCard(s, [
      { role: "header", text: "Say doubles with me" },
    ], { x: 0.5, y: 4.5, w: 9, h: 0.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL1);
  })();

  // Fluency 2 — +1 quick fire
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "+1 - count up", { color: STAGE_COLORS["1"] });

    const examples = [
      { from: "7", to: "8" },
      { from: "9", to: "10" },
      { from: "12", to: "13" },
      { from: "18", to: "19" },
    ];
    examples.forEach((ex, i) => {
      const x = 0.6 + (i % 2) * 4.5;
      const y = 1.8 + Math.floor(i / 2) * 1.3;
      drawNumeralCard(s, x, y, 1.4, 1.1, ex.from, { fontSize: 48 });
      s.addText("+1", { x: x + 1.5, y: y, w: 0.6, h: 1.1, fontSize: 22, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, x + 2.2, y, 1.4, 1.1, ex.to, { fontSize: 48, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });
    });

    addInstructionCard(s, [
      { role: "header", text: "I say. You add 1." },
    ], { x: 0.5, y: 4.5, w: 9, h: 0.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL2);
  })();

  // Vocabulary — altogether
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Word", { color: C.PRIMARY });
    addTitle(s, "Altogether", { color: C.PRIMARY });

    // 4 red + 3 blue counters
    addTensFrame(s, 2.5, 1.9, 5.0, 0, { cellH: 1.0 });
    // Draw 4 red dots in cells 1-4, 3 blue in cells 5-7
    for (let i = 0; i < 4; i += 1) {
      const cx = 2.5 + (i % 5) * 1.0 + 0.18;
      const cy = 1.9 + Math.floor(i / 5) * 1.0 + 0.18;
      s.addShape("roundRect", {
        x: cx, y: cy, w: 0.64, h: 0.64, rectRadius: 0.32,
        fill: { color: C.ALERT },
      });
    }
    for (let i = 4; i < 7; i += 1) {
      const cx = 2.5 + (i % 5) * 1.0 + 0.18;
      const cy = 1.9 + Math.floor(i / 5) * 1.0 + 0.18;
      s.addShape("roundRect", {
        x: cx, y: cy, w: 0.64, h: 0.64, rectRadius: 0.32,
        fill: { color: C.PRIMARY },
      });
    }

    addInstructionCard(s, [
      { role: "header", text: "Altogether means TOTAL" },
      { role: "body", text: "4 red + 3 blue = 7 altogether." },
    ], { x: 0.5, y: 4.1, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // LI/SC
  liSlide(pres,
    ["We are learning to solve addition stories with a Think Board."],
    [
      "I can read an addition story.",
      "I can show the story with counters or a picture.",
      "I can write the number sentence for the story.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do — Introduce Think Board
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "The Think Board", { color: STAGE_COLORS["2"] });

    drawThinkBoard(s, 0.5, 1.7, 9.0, 3.0, {});

    addInstructionCard(s, [
      { role: "header", text: "Four boxes to help us think" },
      { role: "body", text: "Story  -  Picture  -  Counters  -  Number sentence" },
    ], { x: 0.5, y: 4.85, w: 9, h: 0.25, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_BOARD);
  })();

  // I Do — Duck story Think Board completed
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "5 ducks + 3 more ducks", { color: STAGE_COLORS["2"] });

    drawThinkBoard(s, 0.5, 1.7, 9.0, 3.0, {
      contents: [
        "5 ducks are in the pond. 3 more swim in. How many ducks altogether?",
        "[Draw 5 ducks then 3 more]",
        "5 red counters and 3 blue counters",
        "5 + 3 = 8 ducks!",
      ],
      fontSize: 12,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_STORY1);
  })();

  // CFU — Apple story
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How many apples altogether?", { color: C.ALERT });

      addInstructionCard(s, [
        { role: "header", text: "Story" },
        { role: "body", text: "I have 4 apples. My friend gives me 2 more." },
        { role: "emphasis", text: "How many altogether?" },
      ], { x: 1.0, y: 1.85, w: 8, h: 2.5, strip: C.ALERT });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer. Show me!" },
      ], { x: 0.5, y: 4.4, w: 9, h: 0.7, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "6 apples - 4 + 2 = 6", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // We Do — bird story Think Board
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "7 birds + 4 more birds", { color: STAGE_COLORS["3"] });

      drawThinkBoard(s, 0.5, 1.65, 9.0, 3.1, {
        contents: [
          "7 birds are on the tree. 4 more birds land. How many altogether?",
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
      // Reveal a full Think Board on top via a darkened panel
      addTextOnShape(slide, "11 birds - 7 + 4 = 11!", {
        x: 1.5, y: 4.85, w: 7.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge CFU — Bus story
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "How many children on the bus?", { color: C.ALERT });

      addInstructionCard(s, [
        { role: "header", text: "Story" },
        { role: "body", text: "On the bus there are 8 children." },
        { role: "body", text: "5 more children get on." },
        { role: "emphasis", text: "How many altogether?" },
      ], { x: 1.0, y: 1.85, w: 8, h: 2.5, strip: C.ALERT });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer on your whiteboard." },
      ], { x: 0.5, y: 4.4, w: 9, h: 0.7, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "13 children - 8 + 5 = 13!", {
        x: 1.5, y: 4.55, w: 7.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Solve with Think Board",
    [
      "Read the story.",
      "Draw the picture.",
      "Show with counters.",
      "Write the number sentence.",
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
        { text: "Read the story twice.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Use the four boxes.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "'Altogether' means add!", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Exit Ticket
  exitTicketSlide(pres,
    ["I have 6 grapes. I get 4 more. How many grapes altogether?"],
    NOTES_EXIT, FOOTER);

  // Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner the four boxes on the Think Board.",
    scItems: [
      "I can read an addition story.",
      "I can show the story with counters or a picture.",
      "I can write the number sentence for the story.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1N_Lesson8_Think_Board_Addition.pptx");
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
      { label: "Counters", color: C.ACCENT, content: "" },
      { label: "Number sentence", color: C.SUCCESS, content: "" },
    ];
    cells.forEach((cell, idx) => {
      const cx = x + (idx % 2) * cellW;
      const cy = y + Math.floor(idx / 2) * cellH;
      doc.rect(cx, cy, cellW, cellH).lineWidth(1.5).strokeColor("#444").stroke();
      // Label bar
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
      subtitle: "Addition Think Board",
      color: C.NAVY,
      lessonInfo: "Lesson 8 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Read each story. Fill in all four boxes - picture, counters, number sentence, answer.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Story 1", y, { color: C.NAVY });
    drawPdfThinkBoard(doc, 60, y, 480, 280, "I have 5 stickers. My friend gives me 4 more stickers. How many stickers altogether?");
    y += 295;

    y = addSectionHeading(doc, "Story 2", y, { color: C.NAVY });
    drawPdfThinkBoard(doc, 60, y, 480, 280, "There are 6 fish in the bowl. Dad puts 3 more fish in. How many fish altogether?");
    y += 295;

    addPdfFooter(doc, "Lesson 8 | Numbers to 120 | Year 1");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Addition Think Board - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Lesson 8 of 10 | Year 1 Numeracy",
    });
    y = addSectionHeading(doc, "Story 1: 5 stickers + 4 more", y, { color: C.NAVY });
    y = addBodyText(doc, "Picture: 5 stickers + 4 stickers (any reasonable drawing)", y);
    y = addBodyText(doc, "Counters: 5 + 4 (any colour mix)", y);
    y = addBodyText(doc, "Number sentence: 5 + 4 = 9", y);
    y = addBodyText(doc, "Answer: 9 stickers altogether", y);
    y = addSectionHeading(doc, "Story 2: 6 fish + 3 more", y, { color: C.NAVY });
    y = addBodyText(doc, "Picture: 6 fish + 3 fish (any reasonable drawing)", y);
    y = addBodyText(doc, "Counters: 6 + 3 (any colour mix)", y);
    y = addBodyText(doc, "Number sentence: 6 + 3 = 9", y);
    y = addBodyText(doc, "Answer: 9 fish altogether", y);
    y = addBodyText(doc, "Watch for: students skipping the picture box. The picture is part of the thinking, not an extra.", y);
    addPdfFooter(doc, "Lesson 8 | Answer Key | Year 1");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension - write your own story
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Write your own story",
      color: C.TEAL,
      lessonInfo: "Lesson 8 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Write your own addition story. Solve it with the Think Board.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "My story", y, { color: C.NAVY });
    drawPdfThinkBoard(doc, 60, y, 480, 280, "Write your own story here:");
    y += 295;

    y = addSectionHeading(doc, "A bigger story (numbers past 10)", y, { color: C.NAVY });
    drawPdfThinkBoard(doc, 60, y, 480, 280, "Write a story with bigger numbers:");

    addPdfFooter(doc, "Lesson 8 | Extension | Year 1");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
