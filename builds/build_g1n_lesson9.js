"use strict";

// Grade 1 Numeracy — Lesson 9: Think Board - Subtraction Story Problems
// AC9M1N05 — modelling subtraction stories using a Think Board.
// Daily Review: Counting and place value (tens-ones, pair to 10).
// Fluency:      Skip counting, addition and patterns (counting back, -1).

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

const SESSION = 9;
const TOTAL = 10;
const UNIT_TITLE = "Numbers to 120";
const LESSON_TITLE = "Think Board: Take-Away Stories";
const FOOTER = `${UNIT_TITLE} | Lesson ${SESSION} of ${TOTAL} | Year 1 Numeracy`;
const OUT_DIR = `output/G1N_Lesson${SESSION}_Think_Board_Subtraction`;
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 9 Take Away Think Board",
  "Solve two take-away stories using the four-square Think Board.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 9 Answer Key",
  "Teacher reference with completed take-away Think Boards.");
const EXTENDING_RES = makeSessionResource(SESSION, "Lesson 9 Extension",
  "Write your own take-away story and solve it on a Think Board.");
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
- Today we solve take-away stories
- Things that go away, get eaten, get lost - that is subtraction
- Same Think Board - new word to look for

DO:
- Display title
- Have counters and number tracks ready
- Have a Think Board ready at the front

TEACHER NOTES:
Lesson 9 of 10. Subtraction stories are harder than addition for Year 1 because two strategies emerge - take away and count back.

WATCH FOR:
- Students who remember the Think Board from yesterday - praise

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Same materials as yesterday - counters and Think Board
- Today our number sentence uses a minus sign

DO:
- Show counters at each table
- Show the minus sign on the board

TEACHER NOTES:
The minus sign is new. Introduce it briefly.

WATCH FOR:
- Materials ready

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_LAUNCH = `SAY:
- Story - I have 8 lollies. I eat 3 of them
- How many lollies do I have left? Tell your partner

DO:
- Display 8 lollies image or sketch
- Partner talk
- Cold call

TEACHER NOTES:
A take-away story to activate the subtraction idea. Answer is 5.

WATCH FOR:
- Students who count back 8, 7, 6, 5 - they have a strategy
- Students who count fingers down - also valid

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_DR1_Q = `SAY:
- Addition story - 4 + 3. How many altogether?

DO:
- Display 4 + 3
- Show me

TEACHER NOTES:
Retrieves Lesson 8 addition.

WATCH FOR:
- 7

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- 7. 4 + 3 = 7

DO:
- Reveal 7

TEACHER NOTES:
Quick.

WATCH FOR:
- Confident

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- Tens and ones - 73. How many tens? Write it

DO:
- 8 seconds, show me

TEACHER NOTES:
Retrieves Lesson 6.

WATCH FOR:
- 7

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- 7 tens. 73 has 7 tens and 3 ones

DO:
- Reveal 7

TEACHER NOTES:
Brief.

WATCH FOR:
- Chorus

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- Pair to 10. 9 + ?

DO:
- Show me

TEACHER NOTES:
Retrieves Lesson 5.

WATCH FOR:
- 1

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- 1. 9 + 1 = 10

DO:
- Reveal 1

TEACHER NOTES:
Brief.

WATCH FOR:
- Confident

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1 = `SAY:
- Fluency. Count back from 10
- 10, 9, 8, 7, 6, 5, 4, 3, 2, 1

DO:
- Lead choral count back

TEACHER NOTES:
Counting back is the prerequisite for subtraction.

WATCH FOR:
- Hesitations - mark for support

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL2 = `SAY:
- Minus 1. I say a number. You say one less
- 10 - say it - 9. 7 - say it - 6

DO:
- Quick fire: 10, 7, 12, 18, 25
- Class -1 each

TEACHER NOTES:
The "one less" strategy is the bridge to subtraction.

WATCH FOR:
- Quick chorus = ready

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- A new word - left
- Left means what is still there after some go away
- I had 6, I ate 2, I have 4 LEFT

DO:
- Show 6 counters
- Take 2 away
- Say "4 left"

TEACHER NOTES:
"Left" is the subtraction signal word. "Take away" is another. Both signal subtraction.

WATCH FOR:
- Students who use the word "left" in stories

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_LI_SC = `SAY:
- Read with me - We are learning to solve take-away stories with a Think Board
- I can read a take-away story
- I can show the story with counters or a picture
- I can write a number sentence with a minus sign

DO:
- Choral read

TEACHER NOTES:
SC3 introduces the minus sign in writing.

WATCH FOR:
- Engaged class

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_LOLLIES = `SAY:
- Watch me. STORY: I have 7 lollies. I eat 2. How many left?
- PICTURE: Draw 7 lollies, cross out 2
- COUNTERS: Take 7 counters. Move 2 away. How many left? 5
- NUMBER SENTENCE: 7 minus 2 equals 5. 7 - 2 = 5
- Answer: 5 lollies left

DO:
- Build Think Board on camera
- Read story
- Draw 7 lollies, cross out 2
- Use counters, move 2 away
- Write 7 - 2 = 5
- Point to minus sign and name it

TEACHER NOTES:
The cross-out is essential. Students see the take-away action.

WATCH FOR:
- Students who say "5" before the count - they have the fact

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_FISH = `SAY:
- Another - 10 fish in the bowl. Dad takes out 4. How many left?
- 10 counters. Take 4 away
- 10 minus 4 equals 6
- 6 fish left

DO:
- Build Think Board
- Read story
- Counters down, 4 away
- Write 10 - 4 = 6

TEACHER NOTES:
A bigger sum. Uses a known fact (rainbow 6 + 4 = 10) for those who notice.

WATCH FOR:
- Students who say 6 instantly - they have the pair fact

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. STORY: I have 8 grapes. I eat 3. How many left?
- Write on your whiteboard

DO:
- 10 seconds, show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Read the story. Work out how many left. Show me"
- Scan for: 5
PROCEED: If 80%+ show 5, move to We Do.
PIVOT: Most likely misconception - students add the numbers (11). Reteach by drawing the grapes and crossing out the eaten ones. Stress the word "eat" or "left". Re-check with "9 birds, 2 fly away, how many left?"

TEACHER NOTES:
First CFU. Tests whether students recognise the subtraction signal.

WATCH FOR:
- 5 = secure
- 11 = adding by mistake; needs the take-away visual

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- 5 grapes left. 8 - 3 = 5

DO:
- Reveal 5

TEACHER NOTES:
Brief.

WATCH FOR:
- Confident

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn. STORY: 12 ducks in the pond. 4 fly away
- How many ducks left? Use the Think Board

DO:
- Display the empty Think Board with the story
- Class works the story together
- Cold call

TEACHER NOTES:
12 - 4 = 8. Crossing back over 10 is tricky.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use real counters and a number track. Place 12 counters. Take 4 away. Count what is left.
- Extra Notes: Small group at the front.
EXTENDING PROMPT:
- Task: Write a similar story with different numbers - 11 ducks, 5 fly away.
- Extra Notes: Whole-class extension.

WATCH FOR:
- Students who count back accurately - they have the strategy
- Students who lose count - support with the counters

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- 8 ducks left. 12 - 4 = 8

DO:
- Reveal 8

TEACHER NOTES:
Confirm with the count back: 12, 11, 10, 9, 8.

WATCH FOR:
- Confident class

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge check. STORY: I have 15 stickers. I give 6 to my friend. How many left?
- Write on your whiteboard

DO:
- 15 seconds, show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Story - 15 stickers, give away 6. How many left? Show me"
- Scan for: 9
PROCEED: If 80%+ show 9, move to You Do.
PIVOT: Most likely misconception - students add (21) or count back wrong. Reteach by drawing 15 squares, crossing out 6. Re-check with "14 - 5 = ?"

TEACHER NOTES:
Bigger numbers across 10. Uses count-back or known facts.

WATCH FOR:
- 9 = SC2 secure
- 21 = read story but ignored "give away"

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- 9 stickers left. 15 - 6 = 9

DO:
- Reveal 9

TEACHER NOTES:
Final hinge.

WATCH FOR:
- Confident class = ready

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent work - take-away Think Board time
- Read each story
- Fill all four boxes - picture, counters, number sentence
- Use the minus sign in the number sentence

DO:
- Distribute worksheet
- Circulate
- Hand out Extension when ready

TEACHER NOTES:
Worksheet has 2 take-away stories.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Teacher reads story aloud. Student draws and crosses out. Use counters too.
- Extra Notes: Small group at back.
EXTENDING PROMPT:
- Task: Lesson 9 Extension - write your own take-away story.
- Extra Notes: Distribute Lesson 9 Extension PDF.

WATCH FOR:
- Students using + when they should use - - reteach the "left" word
- Quick finishers - hand out Extension

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- 9 birds on the tree. 3 fly away. How many left?

DO:
- Students write answer
- Collect

TEACHER NOTES:
Exit assesses subtraction story-to-answer. Sort into 6, other.

WATCH FOR:
- 6 = ready
- 12 (added) = needs more take-away work

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Check our success criteria
- I can read a take-away story - thumbs
- I can show with counters or a picture - thumbs
- I can write the number sentence with minus - thumbs
- Tell your partner one word that tells you to take away

DO:
- Thumbs check
- Partner talk

TEACHER NOTES:
"Left" or "take away" or "eat" or "give away" are all valid.

WATCH FOR:
- Variety of signal words

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, LESSON_TITLE, "Lesson 9 of 10 - Numbers to 120", "Year 1 Numeracy | AC9M1N05", NOTES_TITLE);
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Launch — 8 lollies, eat 3
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "How many lollies left?", { color: STAGE_COLORS["1"] });

    addInstructionCard(s, [
      { role: "header", text: "Story" },
      { role: "body", text: "I have 8 lollies." },
      { role: "body", text: "I eat 3 of them." },
      { role: "emphasis", text: "How many lollies LEFT?" },
    ], { x: 1.0, y: 1.8, w: 8, h: 2.7, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // DR1 — 4 + 3
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many altogether?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 1.6, 2.0, "4", { fontSize: 100 });
      s.addText("+", { x: 4.2, y: 1.9, w: 0.7, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, 4.9, 1.9, 1.6, 2.0, "3", { fontSize: 100 });
      s.addText("=", { x: 6.6, y: 1.9, w: 0.7, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, 7.3, 1.9, 1.6, 2.0, "?", { fontSize: 100, color: C.ALERT, stroke: C.ALERT });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "7", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // DR2 — tens in 73
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many tens in 73?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.5, "73", { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "7 tens", {
        x: 3.0, y: 4.55, w: 4.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR2_A);
    }
  );

  // DR3 — 9 + ? = 10
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Pair to make 10?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 1.6, 2.0, "9", { fontSize: 100 });
      s.addText("+", { x: 4.2, y: 1.9, w: 0.7, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, 4.9, 1.9, 1.6, 2.0, "?", { fontSize: 100, color: C.ALERT, stroke: C.ALERT });
      s.addText("= 10", { x: 6.6, y: 1.9, w: 1.8, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });

      addInstructionCard(s, [
        { role: "body", text: "Write the missing number." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // Fluency 1 — Count back from 10
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Count back from 10", { color: STAGE_COLORS["1"] });

    const nums = [10, 9, 8, 7, 6, 5, 4, 3, 2, 1];
    nums.forEach((n, i) => {
      const x = 0.5 + (i % 5) * 1.85;
      const y = 1.8 + Math.floor(i / 5) * 1.2;
      drawNumeralCard(s, x, y, 1.5, 1.0, String(n), { fontSize: 36, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });
    });

    addInstructionCard(s, [
      { role: "header", text: "Count BACK with me" },
    ], { x: 0.5, y: 4.4, w: 9, h: 0.7, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL1);
  })();

  // Fluency 2 — Minus 1
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Minus 1 - one less", { color: STAGE_COLORS["1"] });

    const examples = [
      { from: "10", to: "9" },
      { from: "7", to: "6" },
      { from: "12", to: "11" },
      { from: "18", to: "17" },
    ];
    examples.forEach((ex, i) => {
      const x = 0.6 + (i % 2) * 4.5;
      const y = 1.8 + Math.floor(i / 2) * 1.3;
      drawNumeralCard(s, x, y, 1.4, 1.1, ex.from, { fontSize: 48 });
      s.addText("-1", { x: x + 1.5, y: y, w: 0.6, h: 1.1, fontSize: 22, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, x + 2.2, y, 1.4, 1.1, ex.to, { fontSize: 48, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });
    });

    addInstructionCard(s, [
      { role: "header", text: "I say. You say one LESS." },
    ], { x: 0.5, y: 4.5, w: 9, h: 0.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL2);
  })();

  // Vocabulary — left
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Word", { color: C.PRIMARY });
    addTitle(s, "Left", { color: C.PRIMARY });

    // Tens frame with 4 counters showing 2 crossed out
    addTensFrame(s, 2.5, 1.85, 5.0, 6, { cellH: 1.0, fillColor: C.PRIMARY });
    // Draw "X" over the last 2
    [4, 5].forEach((idx) => {
      const c = idx % 5;
      const r = Math.floor(idx / 5);
      const cx = 2.5 + c * 1.0;
      const cy = 1.85 + r * 1.0;
      s.addShape("line", {
        x: cx + 0.2, y: cy + 0.2, w: 0.6, h: 0.6,
        line: { color: C.ALERT, width: 4 },
      });
      s.addShape("line", {
        x: cx + 0.8, y: cy + 0.2, w: -0.6, h: 0.6,
        line: { color: C.ALERT, width: 4 },
      });
    });

    addInstructionCard(s, [
      { role: "header", text: "Left = what is still there" },
      { role: "body", text: "I had 6. I took 2 away. 4 are LEFT." },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // LI/SC
  liSlide(pres,
    ["We are learning to solve take-away stories with a Think Board."],
    [
      "I can read a take-away story.",
      "I can show the story with counters or a picture.",
      "I can write a number sentence with a minus sign.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do — 7 lollies, eat 2
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "7 lollies, eat 2", { color: STAGE_COLORS["2"] });

    drawThinkBoard(s, 0.5, 1.7, 9.0, 3.0, {
      contents: [
        "I have 7 lollies. I eat 2 of them. How many lollies left?",
        "[Draw 7 lollies, cross out 2]",
        "7 counters - move 2 away",
        "7 - 2 = 5 lollies",
      ],
      fontSize: 12,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_LOLLIES);
  })();

  // I Do — 10 fish, take 4
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "10 fish, take 4", { color: STAGE_COLORS["2"] });

    drawThinkBoard(s, 0.5, 1.7, 9.0, 3.0, {
      contents: [
        "10 fish in the bowl. Dad takes 4 out. How many left?",
        "[Draw 10 fish, cross out 4]",
        "10 counters - take 4 away",
        "10 - 4 = 6 fish",
      ],
      fontSize: 12,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_FISH);
  })();

  // CFU — 8 grapes, eat 3
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How many grapes left?", { color: C.ALERT });

      addInstructionCard(s, [
        { role: "header", text: "Story" },
        { role: "body", text: "I have 8 grapes. I eat 3 of them." },
        { role: "emphasis", text: "How many grapes LEFT?" },
      ], { x: 1.0, y: 1.85, w: 8, h: 2.5, strip: C.ALERT });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer. Show me!" },
      ], { x: 0.5, y: 4.4, w: 9, h: 0.7, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "5 grapes - 8 - 3 = 5", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // We Do — Ducks
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "12 ducks, 4 fly away", { color: STAGE_COLORS["3"] });

      drawThinkBoard(s, 0.5, 1.65, 9.0, 3.1, {
        contents: [
          "12 ducks in the pond. 4 fly away. How many left?",
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
      addTextOnShape(slide, "8 ducks - 12 - 4 = 8!", {
        x: 2.0, y: 4.85, w: 6.0, h: 0.25, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge CFU — 15 stickers, give 6
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "How many stickers left?", { color: C.ALERT });

      addInstructionCard(s, [
        { role: "header", text: "Story" },
        { role: "body", text: "I have 15 stickers." },
        { role: "body", text: "I give 6 to my friend." },
        { role: "emphasis", text: "How many stickers LEFT?" },
      ], { x: 1.0, y: 1.85, w: 8, h: 2.5, strip: C.ALERT });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer on your whiteboard." },
      ], { x: 0.5, y: 4.4, w: 9, h: 0.7, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "9 stickers - 15 - 6 = 9!", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Solve with Think Board",
    [
      "Read the story.",
      "Draw the picture - cross out what goes away.",
      "Show with counters.",
      "Write the number sentence with -.",
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
        { text: "Cross out what goes away.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Count what is LEFT.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Use the minus sign!", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Exit Ticket
  exitTicketSlide(pres,
    ["9 birds on the tree. 3 fly away. How many left?"],
    NOTES_EXIT, FOOTER);

  // Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner one word that tells you to take away.",
    scItems: [
      "I can read a take-away story.",
      "I can show with counters or a picture.",
      "I can write a number sentence with a minus sign.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1N_Lesson9_Think_Board_Subtraction.pptx");
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
      subtitle: "Take-away Think Board",
      color: C.NAVY,
      lessonInfo: "Lesson 9 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Read each story. Fill in all four boxes. Cross out what goes away.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Story 1", y, { color: C.NAVY });
    drawPdfThinkBoard(doc, 60, y, 480, 280, "I have 7 cookies. I eat 3 cookies. How many cookies are left?");
    y += 295;

    y = addSectionHeading(doc, "Story 2", y, { color: C.NAVY });
    drawPdfThinkBoard(doc, 60, y, 480, 280, "There are 10 ducks in the pond. 4 ducks fly away. How many ducks are left?");

    addPdfFooter(doc, "Lesson 9 | Numbers to 120 | Year 1");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Take-away Think Board - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Lesson 9 of 10 | Year 1 Numeracy",
    });
    y = addSectionHeading(doc, "Story 1: 7 cookies - 3 eaten", y, { color: C.NAVY });
    y = addBodyText(doc, "Picture: 7 cookies with 3 crossed out", y);
    y = addBodyText(doc, "Counters: 7 with 3 moved away", y);
    y = addBodyText(doc, "Number sentence: 7 - 3 = 4", y);
    y = addBodyText(doc, "Answer: 4 cookies left", y);
    y = addSectionHeading(doc, "Story 2: 10 ducks - 4 fly away", y, { color: C.NAVY });
    y = addBodyText(doc, "Picture: 10 ducks with 4 crossed out (or with arrows)", y);
    y = addBodyText(doc, "Counters: 10 with 4 moved away", y);
    y = addBodyText(doc, "Number sentence: 10 - 4 = 6", y);
    y = addBodyText(doc, "Answer: 6 ducks left", y);
    y = addBodyText(doc, "Watch for: students using + instead of - - reteach the 'take away' or 'left' signal.", y);
    addPdfFooter(doc, "Lesson 9 | Answer Key | Year 1");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension - write your own story
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Write your own take-away story",
      color: C.TEAL,
      lessonInfo: "Lesson 9 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Write your own take-away story. Use words like eat, take, give away, fly away.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "My take-away story", y, { color: C.NAVY });
    drawPdfThinkBoard(doc, 60, y, 480, 280, "Write your own take-away story here:");
    y += 295;

    y = addSectionHeading(doc, "A bigger take-away story", y, { color: C.NAVY });
    drawPdfThinkBoard(doc, 60, y, 480, 280, "Try numbers past 10 in your story:");

    addPdfFooter(doc, "Lesson 9 | Extension | Year 1");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
