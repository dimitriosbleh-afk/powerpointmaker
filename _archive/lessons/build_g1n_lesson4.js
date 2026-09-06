"use strict";

// Grade 1 Numeracy — Lesson 4: Extending Numbers to 120
// AC9M1N01 — recognise, represent and order numbers to at least 120.
// Daily Review: Counting and place value (find on chart, count by 10s).
// Fluency:      Skip counting, addition and patterns (count on past 100).

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
  addNumberTrack, addBaseTenBlocks,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const TOTAL = 10;
const UNIT_TITLE = "Numbers to 120";
const LESSON_TITLE = "Extending the Count to 120";
const FOOTER = `${UNIT_TITLE} | Lesson ${SESSION} of ${TOTAL} | Year 1 Numeracy`;
const OUT_DIR = `output/G1N_Lesson${SESSION}_Extending_To_120`;
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 4 Numbers Past 100",
  "Count and write numbers from 100 to 120 with ten-rod pictures.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 4 Answer Key",
  "Teacher reference with answers for the past-100 sheet.");
const EXTENDING_RES = makeSessionResource(SESSION, "Lesson 4 Extension",
  "Order numbers across 99/100 and find the missing numbers to 120.");
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

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we go past 100
- We will count up to 120 - that is more than ever
- We use what we know about tens to help us

DO:
- Display the title slide
- Have base-10 blocks (MAB) ready
- Have a number track 0-120 ready

TEACHER NOTES:
Lesson 4 of 10. Going past 100 is exciting for Year 1. The key idea - one hundred and one, one hundred and two - keep adding ones.

WATCH FOR:
- Students who think 100 is the biggest number - reassure them numbers keep going

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today we use MAB blocks - tens and ones
- The worksheet shows numbers past 100 with block pictures

DO:
- Show the MAB rods at the front
- Distribute base-10 sets to each table

TEACHER NOTES:
Each table needs at least 12 tens rods and 12 ones cubes.

WATCH FOR:
- Missing manipulatives

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_LAUNCH = `SAY:
- Look. 10 tens rods. How many is that?
- Tell your partner

DO:
- Show 10 tens rods on the board
- Allow 30 seconds partner talk
- Cold call

TEACHER NOTES:
Activates "10 tens makes 100" understanding. The bridge to today's content.

WATCH FOR:
- Students who answer 10 - they may be counting rods, not the ones inside
- Students who answer 100 - they have the strategy

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_DR1_Q = `SAY:
- Daily Review. I am in the 70s row. I end in 4
- Write the number on your whiteboard

DO:
- Allow 10 seconds
- Show me

TEACHER NOTES:
Retrieves hundreds-chart skill from Lesson 3.

WATCH FOR:
- Students who write 74 quickly - they have the chart strategy

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- 74. 7 tens and 4 ones

DO:
- Reveal 74

TEACHER NOTES:
Quick reveal.

WATCH FOR:
- Whole-class chorus

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- Count by 10s. 20, 30, blank, 50
- What is missing?

DO:
- Display the sequence
- 8 seconds, show me

TEACHER NOTES:
Retrieves the 10s count from Lesson 3.

WATCH FOR:
- Students who write 40

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- 40 is the missing number

DO:
- Reveal 40

TEACHER NOTES:
Brief reveal.

WATCH FOR:
- Confident class

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- Last review. What is 10 more than 80? Write it

DO:
- 8 seconds, show me

TEACHER NOTES:
Tests "ten more" with bigger numbers - previews today's go-past-100 work.

WATCH FOR:
- Students who write 90 - they have ten more
- Students who freeze at 80 - need support

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- 10 more than 80 is 90
- 8 tens then 9 tens

DO:
- Reveal 90

TEACHER NOTES:
Naming "8 tens then 9 tens" plants language for the go-past-100 leap (9 tens then 10 tens).

WATCH FOR:
- Confident class = ready

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1 = `SAY:
- Fluency. Count from 90 to 100 with me
- 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100

DO:
- Display the 90-100 number track
- Lead choral count
- Pause at 100 - "the next number after 99 is 100"

TEACHER NOTES:
The 99 to 100 crossing is a key fluency. Slow at 99 to 100.

WATCH FOR:
- Students who hesitate at 100 - reteach with the count

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL2 = `SAY:
- Now we count past 100. 100, 101, 102
- Read with me

DO:
- Display 100, 101, 102, 103, 104, 105
- Lead a choral read

TEACHER NOTES:
"One hundred and one, one hundred and two" - the new spoken pattern.

WATCH FOR:
- Students who say "ten one" or "tenty-one" - reteach the spoken form

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- A new word for the biggest number - hundred
- One hundred is 10 tens
- 10 tens makes 100

DO:
- Show 10 tens rods on the board
- Say "ten tens makes one hundred"

TEACHER NOTES:
"Hundred" is the new place. Use the rods as the concrete anchor.

WATCH FOR:
- Students looking at the 10 rods and seeing 100

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_LI_SC = `SAY:
- Read with me - We are learning to count and write numbers past 100, all the way to 120
- I can count from 90 to 110
- I can write numbers past 100
- I can show a number past 100 with tens and ones blocks

DO:
- Choral read

TEACHER NOTES:
SC1 - the counting bridge. SC2 - the writing target. SC3 - using blocks.

WATCH FOR:
- Engaged class

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_PAST100 = `SAY:
- Watch how I count past 100
- I show 10 tens. That is 100
- I add 1 cube. Now I have 101
- I add another cube. Now I have 102
- I write it - 1, 0, 2 - one hundred and two

DO:
- Use MAB blocks on the document camera
- Add ones one at a time, naming each
- Write 102 on the board

TEACHER NOTES:
The concrete blocks are essential. Year 1 students need to see the ones being added.

WATCH FOR:
- Students who count along - good engagement

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_115 = `SAY:
- Now a trickier one. 115
- I need 1 hundred. So 10 tens rods
- I need 1 more ten. So 1 more rod. That makes 110
- I need 5 ones. 5 cubes
- 1 hundred, 1 ten, 5 ones - 115

DO:
- Build 115 with blocks under the camera
- Count the tens out loud as you place them
- Place the 5 cubes

TEACHER NOTES:
115 is the more demanding example. It uses 11 tens (10 for the hundred, 1 more) and 5 ones.

WATCH FOR:
- Students who can count 11 rods - they have the bridge

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. The next number after 109 is what?
- Write it on your whiteboard

DO:
- 8 seconds
- Show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Write the next number after 109. Show me"
- Scan for: 110
PROCEED: If 80%+ show 110, move to We Do.
PIVOT: Most likely misconception - students write 200 or 1010. Reteach with the block model: 109 = 1 hundred + 9 ones. Add 1 more cube. That is 10 ones = 1 ten. So 1 hundred + 1 ten + 0 ones = 110. Re-check with "after 119".

TEACHER NOTES:
The 109 to 110 crossing tests whether students can apply place value past 100.

WATCH FOR:
- Students who write 1010 - they need block work

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- 110. After 109 comes 110

DO:
- Reveal 110

TEACHER NOTES:
Brief reveal.

WATCH FOR:
- Hands up with 110 = ready

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Look at the picture - 1 hundred, 1 ten, 3 ones
- What number is that? Write it on your whiteboard

DO:
- Show the picture of 1 hundred flat + 1 ten rod + 3 ones cubes
- 15 seconds
- Show me

TEACHER NOTES:
This We Do practises reading the block picture as a numeral. Bridge to writing.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use real MAB blocks. Build 113 alongside the student. Count tens and ones aloud.
- Extra Notes: Small group at the front table with real blocks.
EXTENDING PROMPT:
- Task: Add the next two numbers - 114 and 115 - by drawing one cube at a time.
- Extra Notes: Whole-class extension; can be done from seats.

WATCH FOR:
- Students who write 113 confidently - they have block-to-numeral
- Students who write 1113 - they are not yet seeing the hundred as 100; reteach with blocks

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- 113. One hundred, one ten, three ones
- 1, 1, 3 - one hundred and thirteen

DO:
- Reveal 113
- Tap each digit

TEACHER NOTES:
Naming the digits in order plants the writing routine.

WATCH FOR:
- Confident class

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge check. Write the missing number - 119, blank, 121

DO:
- Display the sequence
- 10 seconds, show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Missing number - write it. Show me"
- Scan for: 120
PROCEED: If 80%+ show 120, move to You Do.
PIVOT: Most likely misconception - students write 200, or 110, or 119. Reteach with the count - 119, then one more, then one more. Re-check with "before 120, after 120".

TEACHER NOTES:
This hinge tests reading and writing across the 120 boundary - the end-of-unit target.

WATCH FOR:
- Students who write 120 confidently - they have full SC2

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- 120. After 119 comes 120

DO:
- Reveal 120

TEACHER NOTES:
Final hinge before independent work.

WATCH FOR:
- Confident class = ready

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent work
- Take the worksheet
- First - count the blocks
- Next - write the number
- Then - fill in the missing numbers on the number track

DO:
- Distribute worksheet
- Circulate to spot students writing 1015 instead of 115
- Hand out Extension when ready

TEACHER NOTES:
The worksheet has 4 block-picture problems (Section A) and a missing-number track from 100 to 120 (Section B).

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work with the teacher and real MAB blocks. Build each number from the worksheet picture, then write.
- Extra Notes: Small group at back table.
EXTENDING PROMPT:
- Task: Lesson 4 Extension - order numbers across 99 and 100, find missing numbers from 95 to 120.
- Extra Notes: Distribute Lesson 4 Extension PDF.

WATCH FOR:
- Students who write 1015 for 115 - they need block work
- Students who finish quickly - hand out Extension

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Write the number that comes right after 109

DO:
- Students write on slip
- Collect

TEACHER NOTES:
Exit ticket targets SC2 (write past 100). Sort into 110, other.

WATCH FOR:
- 110 correct = ready for partitioning lessons (5-7)
- 200 or other = small group support tomorrow

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Check our success criteria
- I can count from 90 to 110 - thumbs
- I can write numbers past 100 - thumbs
- I can show a number with blocks - thumbs
- Tell your partner the biggest number you can now write

DO:
- Run thumbs check
- Partner talk

TEACHER NOTES:
This is the unit's place-value peak. Lessons 5 to 7 focus inward on partitioning.

WATCH FOR:
- Confidence rising - the past-100 fear is gone

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, LESSON_TITLE, "Lesson 4 of 10 - Numbers to 120", "Year 1 Numeracy | AC9M1N01", NOTES_TITLE);
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch — 10 tens rods
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "How many is 10 tens?", { color: STAGE_COLORS["1"] });

    // Draw 10 tens rods using base-10 helper
    addBaseTenBlocks(s, 1.7, 2.0, 0, 10, 0, { unit: 0.18, color: C.PRIMARY });

    addInstructionCard(s, [
      { role: "header", text: "Tell your partner" },
      { role: "body", text: "How many are in 10 tens rods altogether?" },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // DR1 — 70s row, ends in 4
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "What is my number?", { color: STAGE_COLORS["1"] });

      addInstructionCard(s, [
        { role: "header", text: "I am in the 70s row." },
        { role: "emphasis", text: "I end in 4." },
        { role: "body", text: "Write me on your whiteboard." },
      ], { x: 1.5, y: 1.9, w: 7, h: 2.5, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "74", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // DR2 — Missing 10s pattern 20, 30, ?, 50
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Missing number?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 0.8, 2.0, 1.7, 1.6, "20", { fontSize: 70 });
      drawNumeralCard(s, 2.8, 2.0, 1.7, 1.6, "30", { fontSize: 70 });
      drawNumeralCard(s, 4.8, 2.0, 1.7, 1.6, "?", { fontSize: 70, color: C.ALERT, stroke: C.ALERT });
      drawNumeralCard(s, 6.8, 2.0, 1.7, 1.6, "50", { fontSize: 70 });

      addInstructionCard(s, [
        { role: "body", text: "Write the missing number on your whiteboard." },
      ], { x: 0.5, y: 3.85, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "40", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR2_A);
    }
  );

  // DR3 — 10 more than 80
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "10 more than 80?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 2.2, 2.0, "80", { fontSize: 120 });
      s.addText("+ 10", {
        x: 4.8, y: 1.9, w: 1.0, h: 2.0,
        fontSize: 36, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0,
      });
      drawNumeralCard(s, 5.9, 1.9, 2.2, 2.0, "?", { fontSize: 120, color: C.ALERT, stroke: C.ALERT });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer on your whiteboard." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "90", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // Fluency 1 — Count 90 to 100
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Count 90 to 100", { color: STAGE_COLORS["1"] });

    addNumberTrack(s, 0.5, 2.2, 9.0, 90, 100, [100], { cellH: 0.85, fontSize: 22 });

    addInstructionCard(s, [
      { role: "header", text: "Count with me - 90, 91, 92..." },
      { role: "body", text: "Pause at 100 - the next number after 99 is 100." },
    ], { x: 0.5, y: 3.4, w: 9, h: 1.65, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL1);
  })();

  // Fluency 2 — Past 100 (100, 101, 102, 103, 104, 105)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Past 100!", { color: STAGE_COLORS["1"] });

    addNumberTrack(s, 0.5, 2.2, 9.0, 100, 110, [100], { cellH: 0.85, fontSize: 22 });

    addInstructionCard(s, [
      { role: "header", text: "Read with me" },
      { role: "body", text: "one hundred, one hundred and one, one hundred and two..." },
    ], { x: 0.5, y: 3.4, w: 9, h: 1.65, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL2);
  })();

  // Vocabulary — hundred
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Word", { color: C.PRIMARY });
    addTitle(s, "Hundred", { color: C.PRIMARY });

    // Show 10 tens rods
    addBaseTenBlocks(s, 1.7, 1.9, 0, 10, 0, { unit: 0.18, color: C.PRIMARY });

    addInstructionCard(s, [
      { role: "header", text: "10 tens make 1 hundred." },
      { role: "body", text: "100 is the new place." },
      { role: "emphasis", text: "We write it 1-0-0." },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // LI/SC
  liSlide(pres,
    ["We are learning to count and write numbers past 100, all the way to 120."],
    [
      "I can count from 90 to 110.",
      "I can write numbers past 100.",
      "I can show a number past 100 with tens and ones blocks.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do — 100 + 2 = 102
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Count past 100", { color: STAGE_COLORS["2"] });

    // Show 10 tens + 2 ones
    addBaseTenBlocks(s, 1.2, 1.95, 0, 10, 2, { unit: 0.18, color: C.PRIMARY });

    // Big numeral 102
    drawNumeralCard(s, 6.5, 2.0, 2.5, 1.8, "102",
      { fontSize: 100, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });

    addInstructionCard(s, [
      { role: "header", text: "10 tens = 100" },
      { role: "body", text: "Add 2 ones = 102" },
      { role: "emphasis", text: "one hundred and two" },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_PAST100);
  })();

  // I Do — 115
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Show 115 with blocks", { color: STAGE_COLORS["2"] });

    // 11 tens + 5 ones
    addBaseTenBlocks(s, 0.7, 1.95, 0, 11, 5, { unit: 0.16, color: C.PRIMARY });

    drawNumeralCard(s, 6.5, 2.0, 2.5, 1.8, "115",
      { fontSize: 100, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });

    addInstructionCard(s, [
      { role: "header", text: "11 tens + 5 ones = 115" },
      { role: "emphasis", text: "one hundred and fifteen" },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_115);
  })();

  // CFU — After 109?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "What comes after 109?", { color: C.ALERT });

      drawNumeralCard(s, 2.5, 1.9, 2.2, 2.0, "109", { fontSize: 90 });
      s.addText("then", {
        x: 4.7, y: 1.9, w: 1.0, h: 2.0,
        fontSize: 26, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0,
      });
      drawNumeralCard(s, 5.8, 1.9, 2.2, 2.0, "?", { fontSize: 90, color: C.ALERT, stroke: C.ALERT });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer on your whiteboard. Show me!" },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "110!", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // We Do — What does this picture show? (1 hundred + 1 ten + 3 ones)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "What number is this?", { color: STAGE_COLORS["3"] });

      // 11 tens + 3 ones for 113
      addBaseTenBlocks(s, 0.7, 1.95, 0, 11, 3, { unit: 0.16, color: C.PRIMARY });

      addInstructionCard(s, [
        { role: "header", text: "Read the picture" },
        { role: "body", text: "Write the number on your whiteboard." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.SECONDARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "113 - one hundred and thirteen!", {
        x: 1.5, y: 4.55, w: 7.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge CFU — missing 119, ?, 121
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "Missing number?", { color: C.ALERT });

      drawNumeralCard(s, 1.5, 1.9, 2.2, 2.0, "119", { fontSize: 90 });
      drawNumeralCard(s, 3.9, 1.9, 2.2, 2.0, "?", { fontSize: 90, color: C.ALERT, stroke: C.ALERT });
      drawNumeralCard(s, 6.3, 1.9, 2.2, 2.0, "121", { fontSize: 90 });

      addInstructionCard(s, [
        { role: "body", text: "Write the missing number on your whiteboard." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "120!", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Count past 100",
    [
      "First - count the blocks.",
      "Next - write the number.",
      "Then - fill in the missing numbers on the track.",
      "",
      "You have 10 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.ALERT });
      slide.addText("Remember", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "10 tens = 100.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Count tens first, then ones.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "After 109 comes 110!", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Exit Ticket
  exitTicketSlide(pres,
    ["Write the number that comes right after 109."],
    NOTES_EXIT, FOOTER);

  // Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner the biggest number you can now write.",
    scItems: [
      "I can count from 90 to 110.",
      "I can write numbers past 100.",
      "I can show a number with tens and ones blocks.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1N_Lesson4_Extending_To_120.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  function hex(c) { return c.startsWith("#") ? c : "#" + c; }

  function drawBlockPicture(doc, x, y, tens, ones, color) {
    const unit = 4;
    let cx = x;
    // Hundreds flat (10 tens = 100). For numbers >100, draw one 10x10 flat
    if (tens >= 10) {
      doc.save();
      for (let r = 0; r < 10; r += 1) {
        for (let c = 0; c < 10; c += 1) {
          doc.rect(cx + c * unit, y + r * unit, unit, unit)
            .fillAndStroke(hex(color), "#FFFFFF");
        }
      }
      doc.restore();
      cx += unit * 10 + 12;
      const extraTens = tens - 10;
      for (let t = 0; t < extraTens; t += 1) {
        doc.save();
        for (let r = 0; r < 10; r += 1) {
          doc.rect(cx + t * (unit + 1), y + r * unit, unit, unit)
            .fillAndStroke(hex(color), "#FFFFFF");
        }
        doc.restore();
      }
      cx += extraTens * (unit + 1) + 10;
    } else {
      for (let t = 0; t < tens; t += 1) {
        doc.save();
        for (let r = 0; r < 10; r += 1) {
          doc.rect(cx + t * (unit + 1), y + r * unit, unit, unit)
            .fillAndStroke(hex(color), "#FFFFFF");
        }
        doc.restore();
      }
      cx += tens * (unit + 1) + 10;
    }
    // Ones cubes
    for (let o = 0; o < ones; o += 1) {
      doc.save();
      doc.rect(cx + o * (unit + 1), y, unit, unit)
        .fillAndStroke(hex(color), "#FFFFFF");
      doc.restore();
    }
  }

  // Worksheet — block pictures + missing-number track
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Numbers past 100",
      color: C.NAVY,
      lessonInfo: "Lesson 4 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Count the tens and ones in the picture. Write the number on the line.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Count the blocks and write the number", y, { color: C.NAVY });
    function drawProblem(doc, num, tens, ones, y) {
      const x = 60;
      doc.fontSize(12).font("Sans-Bold").fillColor("#000");
      doc.text(num + ".", x, y);
      drawBlockPicture(doc, x + 20, y, tens, ones, C.PRIMARY);
      // Write line
      doc.fontSize(11).font("Sans-Bold").fillColor("#000");
      doc.text("Write:", x + 360, y + 12);
      doc.moveTo(x + 405, y + 32).lineTo(x + 535, y + 32).lineWidth(1.2).strokeColor("#000").stroke();
      return y + 55;
    }
    y = drawProblem(doc, "1", 10, 3, y);
    y = drawProblem(doc, "2", 10, 7, y);
    y = drawProblem(doc, "3", 11, 2, y);
    y = drawProblem(doc, "4", 11, 5, y);

    y = addSectionHeading(doc, "Fill in the missing numbers", y, { color: C.NAVY });

    // Draw a number track from 100 to 120 with some missing
    const trackX = 60;
    const cellW = 45;
    const cellH = 35;
    const trackNums = [100, null, 102, null, null, 105, 106, null, 108, null, 110];
    trackNums.forEach((n, i) => {
      const cx = trackX + (i % 11) * cellW;
      doc.rect(cx, y, cellW, cellH).lineWidth(1).strokeColor(hex(C.NAVY)).stroke();
      if (n != null) {
        doc.fontSize(14).font("Sans-Bold").fillColor(hex(C.NAVY));
        doc.text(String(n), cx, y + 10, { width: cellW, align: "center" });
      }
    });
    y += cellH + 12;

    const trackNums2 = [111, null, null, 114, null, 116, null, 118, null, 120];
    trackNums2.forEach((n, i) => {
      const cx = trackX + i * cellW;
      doc.rect(cx, y, cellW, cellH).lineWidth(1).strokeColor(hex(C.NAVY)).stroke();
      if (n != null) {
        doc.fontSize(14).font("Sans-Bold").fillColor(hex(C.NAVY));
        doc.text(String(n), cx, y + 10, { width: cellW, align: "center" });
      }
    });

    addPdfFooter(doc, "Lesson 4 | Numbers to 120 | Year 1");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Numbers past 100 - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Lesson 4 of 10 | Year 1 Numeracy",
    });
    y = addSectionHeading(doc, "Block pictures", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 10 tens + 3 ones = 103 (one hundred and three)", y);
    y = addBodyText(doc, "2. 10 tens + 7 ones = 107 (one hundred and seven)", y);
    y = addBodyText(doc, "3. 11 tens + 2 ones = 112 (one hundred and twelve)", y);
    y = addBodyText(doc, "4. 11 tens + 5 ones = 115 (one hundred and fifteen)", y);
    y = addSectionHeading(doc, "Number track", y, { color: C.NAVY });
    y = addBodyText(doc, "Row 1 missing: 101, 103, 104, 107, 109", y);
    y = addBodyText(doc, "Row 2 missing: 112, 113, 115, 117, 119", y);
    y = addBodyText(doc, "Watch for: students writing 1013 for 113 - reteach with blocks.", y);
    addPdfFooter(doc, "Lesson 4 | Answer Key | Year 1");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension - cross-100 ordering
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Numbers across 100 - order and find",
      color: C.TEAL,
      lessonInfo: "Lesson 4 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Order each set. Find the missing numbers. Use the count: 99 then 100 then 101!", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Order each set from smallest to biggest", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  98 , 101 , 95 , 103 , 99   ->   _____ , _____ , _____ , _____ , _____", y);
    y = addBodyText(doc, "2.  110 , 99 , 105 , 119 , 102", y);
    y = addBodyText(doc, "    ->   _____ , _____ , _____ , _____ , _____", y);
    y = addBodyText(doc, "3.  120 , 89 , 100 , 111 , 95", y);
    y = addBodyText(doc, "    ->   _____ , _____ , _____ , _____ , _____", y);

    y = addSectionHeading(doc, "Find the missing numbers", y, { color: C.NAVY });
    y = addBodyText(doc, "4.   99 , _____ , 101", y);
    y = addBodyText(doc, "5.   108 , _____ , 110", y);
    y = addBodyText(doc, "6.   118 , _____ , 120", y);
    y = addBodyText(doc, "7.   104 , _____ , _____ , 107", y);

    addPdfFooter(doc, "Lesson 4 | Extension | Year 1");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
