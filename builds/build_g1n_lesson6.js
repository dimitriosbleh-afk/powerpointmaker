"use strict";

// Grade 1 Numeracy — Lesson 6: Partitioning Two-Digit Numbers into Tens & Ones
// AC9M1N02 — using physical and virtual materials to partition numbers into
//            counts of tens and ones; e.g. recognise 35 as 3 tens + 5 ones.
// Daily Review: Counting and place value (pairs to 10, ten more, read).
// Fluency:      Skip counting, addition and patterns (10 more, count by 10s).

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
  addBaseTenBlocks, addPlaceValueChart,
  STAGE_COLORS,
} = T;

const SESSION = 6;
const TOTAL = 10;
const UNIT_TITLE = "Numbers to 120";
const LESSON_TITLE = "Tens and Ones";
const FOOTER = `${UNIT_TITLE} | Lesson ${SESSION} of ${TOTAL} | Year 1 Numeracy`;
const OUT_DIR = `output/G1N_Lesson${SESSION}_Tens_And_Ones`;
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 6 Tens and Ones",
  "Match the blocks to the number and complete the tens-ones chart.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 6 Answer Key",
  "Teacher reference for the tens-and-ones practice sheet.");
const EXTENDING_RES = makeSessionResource(SESSION, "Lesson 6 Extension",
  "Build big numbers - partition 47, 63, 82 with blocks and number sentences.");
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
- Today we break numbers into tens and ones
- Like LEGO bricks - some are big tens, some are small ones
- We will see 35 is really 3 tens and 5 ones

DO:
- Display title
- Have MAB blocks ready
- Have a place-value chart at the front

TEACHER NOTES:
Lesson 6 of 10. Place value is the most important Year 1 idea. Build the visual model first - blocks then chart then numeral.

WATCH FOR:
- Students excited by the breaking-apart framing

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today we use tens rods and ones cubes
- Plus a place-value chart - tens on the left, ones on the right

DO:
- Show MAB rods
- Point to a place-value chart on the wall

TEACHER NOTES:
Each table needs a small MAB set. Place-value chart can be drawn on whiteboards if no printed version.

WATCH FOR:
- Tables ready

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_LAUNCH = `SAY:
- I have 3 tens rods and 5 ones cubes. How many altogether?
- Tell your partner

DO:
- Show 3 rods and 5 cubes on the board
- Partner talk

TEACHER NOTES:
Concrete starter for partition. Students should answer 35.

WATCH FOR:
- Students who count by 10s, 10s, 10s, then ones - they have the strategy
- Students who count by 1s - acceptable but slow

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_DR1_Q = `SAY:
- Pairs to 10. 4 plus what equals 10?

DO:
- Display "4 + ? = 10"
- Show me

TEACHER NOTES:
Retrieves Lesson 5.

WATCH FOR:
- 6

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- 6. 4 + 6 = 10

DO:
- Reveal 6

TEACHER NOTES:
Quick.

WATCH FOR:
- Confident

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- 10 more than 60? Write it

DO:
- Display "60 + 10"
- Show me

TEACHER NOTES:
Retrieves Lesson 4 ten-more skill.

WATCH FOR:
- 70

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- 70. 10 more than 60

DO:
- Reveal 70

TEACHER NOTES:
Brief.

WATCH FOR:
- Class confident

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- Last review. Read this numeral when I point

DO:
- Display 46
- Point, chorus

TEACHER NOTES:
Retrieves Lesson 1 reading.

WATCH FOR:
- "Forty-six"

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- Forty-six

DO:
- Reveal

TEACHER NOTES:
Brief.

WATCH FOR:
- All in

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1 = `SAY:
- Count by 10s. We start at 0
- 10, 20, 30, 40, 50, 60, 70, 80, 90, 100

DO:
- Lead choral count

TEACHER NOTES:
Refreshes the 10s count.

WATCH FOR:
- Confident chorus

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL2 = `SAY:
- Now 10 more. I say a number. You add 10
- 25 - say it - 35

DO:
- Quick fire 5 numbers - 25, 41, 17, 63, 8
- Class chorus +10 each

TEACHER NOTES:
"10 more" fluency previews the tens-digit-changes-by-1 pattern.

WATCH FOR:
- Students who add 10 without finger counting - they have it

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- Two important words today - tens and ones
- A TEN is a stick of 10
- A ONE is a single cube
- 10 ones make a ten

DO:
- Show a tens rod and a ones cube side by side
- Build a tens rod from 10 ones cubes

TEACHER NOTES:
This is the central vocabulary for Year 1 place value.

WATCH FOR:
- Students who use the words correctly later

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_LI_SC = `SAY:
- Read with me - We are learning to break numbers into tens and ones
- I can show a number with tens and ones blocks
- I can say how many tens and how many ones
- I can write the number when I see the blocks

DO:
- Choral read

TEACHER NOTES:
SC1 is the build. SC2 is the name. SC3 is the write.

WATCH FOR:
- All engaged

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_35 = `SAY:
- Watch how I break 35 into parts
- I take 3 tens rods. That is 30
- I take 5 ones cubes. That is 5
- 30 plus 5 equals 35
- 35 has 3 tens and 5 ones

DO:
- Build 35 with MAB on the camera
- Place 3 rods, count "10, 20, 30"
- Place 5 cubes, count "31, 32, 33, 34, 35"
- Write 3 in the tens column, 5 in the ones column

TEACHER NOTES:
The count-on from 30 (not from 1) is key. It is the bridge from the concrete to the digits.

WATCH FOR:
- Students who follow the count - good

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_72 = `SAY:
- Now 72. 7 tens and 2 ones
- 7 rods - 10, 20, 30, 40, 50, 60, 70
- 2 cubes - 71, 72
- 72 has 7 tens and 2 ones

DO:
- Build 72 with blocks
- Trace each digit as you count

TEACHER NOTES:
A bigger number. The strategy is the same.

WATCH FOR:
- Students who count rods by tens easily - they have the model

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. The picture shows 4 rods and 6 cubes
- What number is it? Write it on your whiteboard

DO:
- Display 4 rods + 6 cubes
- 8 seconds, show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Count the tens, count the ones. Write the number. Show me"
- Scan for: 46
PROCEED: If 80%+ show 46, move to We Do.
PIVOT: Most likely misconception - students write 64 (reversed) or 10 (rods + cubes added). Reteach the count - 10, 20, 30, 40 then 41, 42... Re-check with 5 rods + 3 cubes.

TEACHER NOTES:
This CFU tests both partition reading and tens-first ordering.

WATCH FOR:
- 46 = secure
- 64 = digit reversal

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- 46. 4 tens, 6 ones

DO:
- Reveal 46

TEACHER NOTES:
Quick.

WATCH FOR:
- Confident

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn. The number is 58
- Tell me how many tens and how many ones
- Write the number sentence - tens plus ones equals 58

DO:
- Display the numeral 58 large
- Allow 30 seconds
- Cold call

TEACHER NOTES:
This is the reverse direction - numeral to parts. Equally important.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work with real blocks. Build 58 with the teacher - count out 5 rods, then 8 cubes. Then write the answer.
- Extra Notes: Small group at front with teacher.
EXTENDING PROMPT:
- Task: After writing 58 = 50 + 8, write a second way - 58 = 40 + 18 (using 4 tens and 18 ones).
- Extra Notes: Whole-class extension; previews Lesson 7.

WATCH FOR:
- Students who write 5 tens, 8 ones - they have it
- Students who write 8 tens, 5 ones - they need tens-first reminder

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- 5 tens and 8 ones. 50 + 8 = 58

DO:
- Reveal 5 tens, 8 ones, 50 + 8 = 58

TEACHER NOTES:
Show all three forms.

WATCH FOR:
- Confident

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge check. Which number has 6 tens and 3 ones?
- A: 36
- B: 63
- C: 9
- Hold up 1, 2 or 3 fingers

DO:
- Display the options
- Show me

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- "Which number has 6 tens and 3 ones? Show me"
- Scan for: 2 fingers (B - 63)
PROCEED: If 80%+ show 2, move to You Do.
PIVOT: Most likely misconception - students choose A (36). Reteach with blocks: count out 6 rods THEN 3 cubes. Re-check with "4 tens and 2 ones".

TEACHER NOTES:
The hinge tests tens-first ordering.

WATCH FOR:
- B confident - place value secure

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- B - 63. 6 tens and 3 ones
- 36 has 3 tens and 6 ones - different

DO:
- Reveal B - 63

TEACHER NOTES:
Final hinge.

WATCH FOR:
- Class ready

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent work
- Take the worksheet
- First - count the tens and the ones
- Next - write how many tens and how many ones
- Then - write the number

DO:
- Distribute worksheet
- Circulate
- Hand out Extension when ready

TEACHER NOTES:
Worksheet has 5 block pictures plus 3 numeral-to-parts items.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work with real blocks. Build each item alongside the teacher.
- Extra Notes: Back table small group.
EXTENDING PROMPT:
- Task: Lesson 6 Extension - partition 47, 63, 82 with both blocks and number sentences.
- Extra Notes: Distribute Lesson 6 Extension PDF.

WATCH FOR:
- Students writing rods + cubes as one number (e.g. 7 for 4 rods + 3 cubes) - reteach with the count "10, 20, 30, 40 then 41, 42, 43"
- Quick finishers - hand out Extension

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- The number is 47
- How many tens? How many ones?

DO:
- Collect slips

TEACHER NOTES:
Exit ticket assesses SC2 (saying tens and ones). Sort into 4 tens + 7 ones (correct), other.

WATCH FOR:
- Correct = ready for Lesson 7

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Check our success criteria
- I can show a number with blocks - thumbs
- I can say tens and ones - thumbs
- I can write the number - thumbs
- Tell your partner how many tens and ones are in 73

DO:
- Thumbs check
- Partner talk

TEACHER NOTES:
The reflection retrieves the partition for a fresh number.

WATCH FOR:
- "7 tens, 3 ones" answers

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, LESSON_TITLE, "Lesson 6 of 10 - Numbers to 120", "Year 1 Numeracy | AC9M1N02", NOTES_TITLE);
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Launch — 3 rods and 5 cubes
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "How many altogether?", { color: STAGE_COLORS["1"] });

    addBaseTenBlocks(s, 1.5, 2.0, 0, 3, 5, { unit: 0.2, color: C.PRIMARY });

    addInstructionCard(s, [
      { role: "header", text: "Tell your partner" },
      { role: "body", text: "3 tens rods and 5 ones cubes. How many?" },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // DR1 — 4 + ? = 10
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Pair to make 10?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 1.6, 2.0, "4", { fontSize: 100 });
      s.addText("+", { x: 4.2, y: 1.9, w: 0.7, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, 4.9, 1.9, 1.6, 2.0, "?", { fontSize: 100, color: C.ALERT, stroke: C.ALERT });
      s.addText("= 10", { x: 6.6, y: 1.9, w: 1.8, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });

      addInstructionCard(s, [
        { role: "body", text: "Write the missing number on your whiteboard." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "6", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // DR2 — 60 + 10
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "10 more than 60?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 2.2, 2.0, "60", { fontSize: 120 });
      s.addText("+ 10", { x: 4.8, y: 1.9, w: 1.0, h: 2.0, fontSize: 36, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, 5.9, 1.9, 2.2, 2.0, "?", { fontSize: 120, color: C.ALERT, stroke: C.ALERT });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer on your whiteboard." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "70", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR2_A);
    }
  );

  // DR3 — Read 46
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Read this numeral", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.5, "46", { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "forty-six", {
        x: 2.5, y: 4.55, w: 5.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // Fluency 1 — Count by 10s
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Count by 10s", { color: STAGE_COLORS["1"] });

    const tens = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100];
    tens.forEach((n, i) => {
      const x = 0.5 + (i % 5) * 1.85;
      const y = 1.8 + Math.floor(i / 5) * 1.2;
      drawNumeralCard(s, x, y, 1.5, 1.0, String(n), { fontSize: 36, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });
    });

    addInstructionCard(s, [
      { role: "header", text: "Read them with me" },
    ], { x: 0.5, y: 4.4, w: 9, h: 0.7, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL1);
  })();

  // Fluency 2 — 10 more quick fire
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "10 more, 10 more!", { color: STAGE_COLORS["1"] });

    const examples = [
      { from: "25", to: "35" },
      { from: "41", to: "51" },
      { from: "17", to: "27" },
      { from: "63", to: "73" },
    ];
    examples.forEach((ex, i) => {
      const x = 0.6 + (i % 2) * 4.5;
      const y = 1.8 + Math.floor(i / 2) * 1.3;
      drawNumeralCard(s, x, y, 1.4, 1.1, ex.from, { fontSize: 48 });
      s.addText("+10", { x: x + 1.5, y: y, w: 0.7, h: 1.1, fontSize: 20, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, x + 2.3, y, 1.4, 1.1, ex.to, { fontSize: 48, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });
    });

    addInstructionCard(s, [
      { role: "header", text: "I say a number. You add 10." },
    ], { x: 0.5, y: 4.5, w: 9, h: 0.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL2);
  })();

  // Vocabulary — Tens and Ones
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Words", { color: C.PRIMARY });
    addTitle(s, "Tens and Ones", { color: C.PRIMARY });

    addBaseTenBlocks(s, 2.5, 1.85, 0, 1, 1, { unit: 0.25, color: C.PRIMARY });

    s.addText("ten", { x: 2.0, y: 4.0, w: 1.0, h: 0.4, fontSize: 18, fontFace: FONT_B, color: C.PRIMARY, bold: true, align: "center", margin: 0 });
    s.addText("one", { x: 3.2, y: 4.0, w: 1.0, h: 0.4, fontSize: 18, fontFace: FONT_B, color: C.PRIMARY, bold: true, align: "center", margin: 0 });

    addInstructionCard(s, [
      { role: "header", text: "A ten is a stick of 10." },
      { role: "body", text: "A one is a single cube." },
      { role: "emphasis", text: "10 ones make 1 ten!" },
    ], { x: 4.6, y: 1.8, w: 4.9, h: 2.7, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // LI/SC
  liSlide(pres,
    ["We are learning to break numbers into tens and ones."],
    [
      "I can show a number with tens and ones blocks.",
      "I can say how many tens and how many ones.",
      "I can write the number when I see the blocks.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do — 35 = 3 tens + 5 ones
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Break 35 into parts", { color: STAGE_COLORS["2"] });

    addBaseTenBlocks(s, 0.7, 1.85, 0, 3, 5, { unit: 0.2, color: C.PRIMARY });

    // Place-value chart on the right
    addPlaceValueChart(s, 6.3, 2.0, ["Tens", "Ones"], ["3", "5"], { cellW: 1.2 });

    drawNumeralCard(s, 6.3, 3.4, 2.4, 1.0, "35", { fontSize: 60, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });

    addInstructionCard(s, [
      { role: "header", text: "3 tens + 5 ones = 35" },
      { role: "emphasis", text: "30 + 5 = 35" },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_35);
  })();

  // I Do — 72 = 7 tens + 2 ones
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Break 72 into parts", { color: STAGE_COLORS["2"] });

    addBaseTenBlocks(s, 0.7, 1.85, 0, 7, 2, { unit: 0.18, color: C.PRIMARY });

    addPlaceValueChart(s, 6.5, 2.0, ["Tens", "Ones"], ["7", "2"], { cellW: 1.2 });
    drawNumeralCard(s, 6.5, 3.4, 2.4, 1.0, "72", { fontSize: 60, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });

    addInstructionCard(s, [
      { role: "header", text: "7 tens + 2 ones = 72" },
      { role: "emphasis", text: "70 + 2 = 72" },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_72);
  })();

  // CFU — 4 rods + 6 cubes
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "What number?", { color: C.ALERT });

      addBaseTenBlocks(s, 1.5, 1.95, 0, 4, 6, { unit: 0.2, color: C.PRIMARY });

      addInstructionCard(s, [
        { role: "body", text: "Count the tens and the ones. Write the number on your whiteboard." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "46 - 4 tens, 6 ones", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // We Do — 58 - reverse direction (numeral to parts)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "How many tens and ones?", { color: STAGE_COLORS["3"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.0, "58", { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addInstructionCard(s, [
        { role: "header", text: "On your whiteboard" },
        { role: "body", text: "How many tens? How many ones?" },
        { role: "body", text: "Write the number sentence: ___ + ___ = 58" },
      ], { x: 0.5, y: 3.9, w: 9, h: 1.2, strip: C.SECONDARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "5 tens + 8 ones - 50 + 8 = 58!", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge CFU — Which has 6 tens and 3 ones?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "Which has 6 tens and 3 ones?", { color: C.ALERT });

      drawNumeralCard(s, 1.0, 1.9, 2.4, 2.0, "36", { fontSize: 100 });
      drawNumeralCard(s, 3.8, 1.9, 2.4, 2.0, "63", { fontSize: 100 });
      drawNumeralCard(s, 6.6, 1.9, 2.4, 2.0, "9", { fontSize: 100 });

      s.addText("A", { x: 1.0, y: 4.0, w: 2.4, h: 0.3, fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });
      s.addText("B", { x: 3.8, y: 4.0, w: 2.4, h: 0.3, fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });
      s.addText("C", { x: 6.6, y: 4.0, w: 2.4, h: 0.3, fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });

      addInstructionCard(s, [
        { role: "body", text: "Hold up 1, 2 or 3 fingers. Show me!" },
      ], { x: 0.5, y: 4.35, w: 9, h: 0.75, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "B - 63 - 6 tens, 3 ones!", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Tens and Ones",
    [
      "First - count the tens.",
      "Next - count the ones.",
      "Then - write the number.",
      "",
      "Show your counting!",
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
        { text: "Tens are sticks of 10.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Count tens first - 10, 20, 30...", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Tens digit first when writing!", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Exit Ticket
  exitTicketSlide(pres,
    ["The number is 47. How many tens? How many ones?"],
    NOTES_EXIT, FOOTER);

  // Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner how many tens and ones are in 73.",
    scItems: [
      "I can show a number with tens and ones blocks.",
      "I can say how many tens and how many ones.",
      "I can write the number when I see the blocks.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1N_Lesson6_Tens_And_Ones.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  function hex(c) { return c.startsWith("#") ? c : "#" + c; }

  function drawPdfBlocks(doc, x, y, tens, ones, color) {
    const unit = 5;
    let cx = x;
    for (let t = 0; t < tens; t += 1) {
      for (let r = 0; r < 10; r += 1) {
        doc.rect(cx + t * (unit + 1), y + r * unit, unit, unit)
          .fillAndStroke(hex(color), "#FFFFFF");
      }
    }
    cx += tens * (unit + 1) + 12;
    for (let o = 0; o < ones; o += 1) {
      doc.rect(cx + o * (unit + 1), y, unit, unit)
        .fillAndStroke(hex(color), "#FFFFFF");
    }
  }

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Tens and ones",
      color: C.NAVY,
      lessonInfo: "Lesson 6 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Count the tens and the ones. Write each part. Write the whole number.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Blocks to number", y, { color: C.NAVY });
    function drawBlocksProblem(doc, num, tens, ones, y) {
      const x = 60;
      doc.fontSize(12).font("Sans-Bold").fillColor("#000");
      doc.text(num + ".", x, y);
      drawPdfBlocks(doc, x + 22, y, tens, ones, C.PRIMARY);
      doc.fontSize(11).font("Sans-Bold").fillColor("#000");
      doc.text("_____ tens   _____ ones   =   _____", x + 270, y + 15);
      return y + 65;
    }
    y = drawBlocksProblem(doc, "1", 2, 4, y);
    y = drawBlocksProblem(doc, "2", 5, 3, y);
    y = drawBlocksProblem(doc, "3", 3, 7, y);

    y = addSectionHeading(doc, "Section B: Number to parts", y, { color: C.NAVY });
    function drawNumProblem(doc, num, number, y) {
      const x = 60;
      doc.fontSize(12).font("Sans-Bold").fillColor("#000");
      doc.text(num + ".  The number is " + number + ".", x, y);
      y += 22;
      doc.fontSize(11).font("Sans-Bold").fillColor("#000");
      doc.text("_____ tens   +   _____ ones   =   " + number, x + 30, y);
      return y + 28;
    }
    y = drawNumProblem(doc, "4", 28, y);
    y = drawNumProblem(doc, "5", 65, y);
    y = drawNumProblem(doc, "6", 81, y);

    addPdfFooter(doc, "Lesson 6 | Numbers to 120 | Year 1");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Tens and ones - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Lesson 6 of 10 | Year 1 Numeracy",
    });
    y = addSectionHeading(doc, "Section A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 2 tens + 4 ones = 24", y);
    y = addBodyText(doc, "2. 5 tens + 3 ones = 53", y);
    y = addBodyText(doc, "3. 3 tens + 7 ones = 37", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.NAVY });
    y = addBodyText(doc, "4. 2 tens + 8 ones = 28", y);
    y = addBodyText(doc, "5. 6 tens + 5 ones = 65", y);
    y = addBodyText(doc, "6. 8 tens + 1 one = 81", y);
    y = addBodyText(doc, "Watch for: digit reversal (e.g. 42 for 24). Reteach with blocks.", y);
    addPdfFooter(doc, "Lesson 6 | Answer Key | Year 1");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Build big numbers",
      color: C.TEAL,
      lessonInfo: "Lesson 6 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Show each number two ways - draw the blocks and write the number sentence.", y, { color: C.TEAL });

    function drawBlock(doc, n, y) {
      doc.fontSize(13).font("Sans-Bold").fillColor("#000");
      doc.text("Number: " + n, 60, y);
      y += 20;
      doc.fontSize(11).font("Sans");
      doc.text("Draw the blocks here:", 60, y);
      doc.rect(60, y + 15, 480, 60).lineWidth(1).strokeColor("#999").stroke();
      y += 80;
      doc.fontSize(11).font("Sans-Bold").fillColor("#000");
      doc.text("Write:   _____ tens   +   _____ ones   =   " + n, 60, y);
      return y + 28;
    }
    y = drawBlock(doc, 47, y);
    y = drawBlock(doc, 63, y);
    y = drawBlock(doc, 82, y);

    addPdfFooter(doc, "Lesson 6 | Extension | Year 1");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
