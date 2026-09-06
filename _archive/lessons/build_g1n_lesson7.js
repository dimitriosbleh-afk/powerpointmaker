"use strict";

// Grade 1 Numeracy — Lesson 7: Partitioning Numbers in Different Ways
// AC9M1N02 — using part-part-whole reasoning to represent 24 in different
//            ways; e.g. 10+10+4 = 24 or 10+14 = 24.
// Daily Review: Counting and place value (tens-ones, read past 100).
// Fluency:      Skip counting, addition and patterns (rainbow facts, 10 more).

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
  addBaseTenBlocks,
  STAGE_COLORS,
} = T;

const SESSION = 7;
const TOTAL = 10;
const UNIT_TITLE = "Numbers to 120";
const LESSON_TITLE = "Different Ways to Show a Number";
const FOOTER = `${UNIT_TITLE} | Lesson ${SESSION} of ${TOTAL} | Year 1 Numeracy`;
const OUT_DIR = `output/G1N_Lesson${SESSION}_Different_Ways`;
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 7 Different Ways",
  "Show 24 in different ways using tens and ones blocks.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 7 Answer Key",
  "Teacher reference with the different partitions of 24, 26, 35.");
const EXTENDING_RES = makeSessionResource(SESSION, "Lesson 7 Extension",
  "Find all the ways to partition 32 and 45 using tens and ones.");
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
- Today we discover a big idea
- A number can be broken in MORE THAN ONE way
- 24 can be 2 tens and 4 ones - or 1 ten and 14 ones!

DO:
- Display title
- Have MAB blocks ready (lots of ones cubes for the unpacking)

TEACHER NOTES:
Lesson 7 of 10. This is the trickiest place-value idea so far. Students may resist "two ways to make the same number". Keep it concrete.

WATCH FOR:
- Students surprised by the multiple ways - lean into the discovery

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Same tens rods and ones cubes today
- Some extra ones cubes - we will need lots

DO:
- Hand out generous MAB sets

TEACHER NOTES:
For 24 as 1 ten + 14 ones, you need 14+ ones cubes. Have spares.

WATCH FOR:
- Enough cubes per table

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_LAUNCH = `SAY:
- I have 2 tens rods and 4 ones cubes - that is 24
- Now watch - I trade ONE tens rod for 10 ones cubes
- I still have the same number. Do I?

DO:
- Build 24 on the camera (2 rods + 4 cubes)
- Replace 1 rod with 10 cubes - now 1 rod + 14 cubes
- Ask partners "still 24?"

TEACHER NOTES:
The trade is the key insight. The whole stays the same. The parts can change.

WATCH FOR:
- Students surprised - that means the idea is landing
- Students who say "yes" confidently - they have conservation

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_DR1_Q = `SAY:
- Daily Review. The number is 56. How many tens?

DO:
- Display 56
- Show me

TEACHER NOTES:
Retrieves Lesson 6.

WATCH FOR:
- 5

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- 5 tens. The tens digit is 5

DO:
- Reveal 5

TEACHER NOTES:
Quick.

WATCH FOR:
- Chorus

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- Pair to 10. 8 plus what equals 10?

DO:
- Display 8 + ? = 10
- Show me

TEACHER NOTES:
Retrieves Lesson 5.

WATCH FOR:
- 2

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- 2. 8 + 2 = 10

DO:
- Reveal 2

TEACHER NOTES:
Brief.

WATCH FOR:
- Confident

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- Read the number after 99

DO:
- Display 99
- Show me

TEACHER NOTES:
Retrieves Lesson 4.

WATCH FOR:
- 100

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- 100. After 99 comes 100

DO:
- Reveal 100

TEACHER NOTES:
Brief.

WATCH FOR:
- Confident

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1 = `SAY:
- Fluency. Pairs to 10 - quick fire
- 3 plus what? 4 plus what? 1 plus what?

DO:
- Quick fire: 3, 4, 1, 7, 9
- Class chorus the partner each time

TEACHER NOTES:
Retrieves rainbow facts from Lesson 5.

WATCH FOR:
- Confident chorus

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL2 = `SAY:
- 10 more - one round
- I say 32 - you say 42. 18 - you say 28

DO:
- Quick fire: 32, 18, 47, 65, 9
- Class +10 each

TEACHER NOTES:
Refreshes the +10 pattern.

WATCH FOR:
- Quick and confident

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- A new word - partition
- Partition means to break or split apart
- We can partition 24 in different ways

DO:
- Show 24 partitioned three ways
- Repeat "partition - break apart"

TEACHER NOTES:
"Partition" is a Year 1 word but kept brief.

WATCH FOR:
- Students using the word

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_LI_SC = `SAY:
- Read with me - We are learning to show one number in different ways
- I can show 24 as 2 tens and 4 ones
- I can show 24 as 1 ten and 14 ones
- I can write a number sentence for each way

DO:
- Choral read

TEACHER NOTES:
The unusual SC2 (1 ten and 14 ones) is the lesson's challenge.

WATCH FOR:
- Confused looks at SC2 - reassure, we will build it together

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_WAY1 = `SAY:
- Watch me show 24 the first way
- 2 tens rods, 4 ones cubes
- I write 24 = 20 + 4
- That is the normal way

DO:
- Build 2 rods + 4 cubes
- Write 24 = 20 + 4

TEACHER NOTES:
Way 1 is the default partition from Lesson 6.

WATCH FOR:
- Students confident with this

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_WAY2 = `SAY:
- Now watch - I trade 1 ten for 10 ones
- I have 1 rod left. I have 4 + 10 = 14 cubes
- The number is still 24!
- I write 24 = 10 + 14

DO:
- Physically replace 1 rod with 10 cubes under the camera
- Show the new arrangement
- Write 24 = 10 + 14

TEACHER NOTES:
The trade IS the lesson. Do it slowly. Year 1 students need to see the swap.

WATCH FOR:
- Students who say "still 24" - they got it

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_WAY3 = `SAY:
- One more way - all ones cubes
- I trade BOTH rods. 24 single cubes
- 24 = 24 ones

DO:
- Show 24 single cubes
- Mention "but the rods help us count, so we usually use rods"

TEACHER NOTES:
Way 3 reinforces conservation. Some students will love this.

WATCH FOR:
- Students delighted by the visual

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. Look at the picture - 1 ten and 14 ones
- What is the number? Write it on your whiteboard

DO:
- Display 1 rod + 14 cubes
- 10 seconds, show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Count the tens AND the ones. Write the number. Show me"
- Scan for: 24
PROCEED: If 80%+ show 24, move to We Do.
PIVOT: Most likely misconception - students write 15 (1 ten + 14 - they count tens digit + ones digit). Reteach by counting cubes one at a time after the rod count of 10. Re-check with 2 rods + 14 cubes (the answer is 34).

TEACHER NOTES:
The unusual case - more than 9 ones - is the tricky bit.

WATCH FOR:
- 24 correct = ready
- 15 wrong = needs reteach with the count

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- 24. 1 ten is 10. Plus 14 ones is 24

DO:
- Reveal 24
- Show 10 + 14 = 24

TEACHER NOTES:
Confirm the count.

WATCH FOR:
- Aha moments

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn. Show me 26 two different ways on your whiteboard
- Way 1 - the normal way with tens and ones
- Way 2 - using only one ten and more ones

DO:
- 60 seconds
- Cold call two students

TEACHER NOTES:
Two ways are required. Students may default to one. Prompt for the second.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work with real blocks. Build 26 the normal way, then trade 1 rod for 10 cubes.
- Extra Notes: Small group at the front.
EXTENDING PROMPT:
- Task: Find a THIRD way - 26 = 0 tens + 26 ones (all cubes).
- Extra Notes: Whole-class extension.

WATCH FOR:
- Students who write both ways - they have it
- Students who write the same way twice - prompt for the trade

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]

`;

const NOTES_WEDO_A = `SAY:
- 26 - two ways
- Way 1: 2 tens and 6 ones - 20 + 6 = 26
- Way 2: 1 ten and 16 ones - 10 + 16 = 26
- The number is the same!

DO:
- Reveal both ways

TEACHER NOTES:
The reveal includes both number sentences.

WATCH FOR:
- Confident class

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge check. Which is the SAME as 35?
- A: 3 tens + 5 ones
- B: 2 tens + 15 ones
- C: BOTH A and B
- Hold up 1, 2 or 3 fingers

DO:
- Display the three options
- 12 seconds, show me

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- "Show me 1, 2 or 3"
- Scan for: 3 fingers (C - both)
PROCEED: If 80%+ show 3, move to You Do.
PIVOT: Most likely misconception - students choose A only. Reteach: build 35 both ways. Re-check with "Which is the same as 42 - 4 tens 2 ones, 3 tens 12 ones, both?"

TEACHER NOTES:
This is the highest-order hinge so far. It tests whether students accept multiple valid partitions.

WATCH FOR:
- C confident = key insight secure
- A only = need more block work

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- C - both! 3 tens + 5 ones is 35. 2 tens + 15 ones is also 35
- The whole stays the same

DO:
- Reveal C - both
- Show both partitions

TEACHER NOTES:
Final hinge.

WATCH FOR:
- Confident class = ready

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent work
- Take the worksheet
- For each number, show it TWO ways with the blocks
- Then write the two number sentences

DO:
- Distribute worksheet
- Circulate - prompt students who only show one way

TEACHER NOTES:
Worksheet has 3 numbers - 24, 31, 47 - each needs two partitions.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work with real blocks. Build 24 with the teacher both ways before writing.
- Extra Notes: Small group at back.
EXTENDING PROMPT:
- Task: Lesson 7 Extension - find all the ways to partition 32 and 45.
- Extra Notes: Distribute Lesson 7 Extension PDF.

WATCH FOR:
- Students who write the same partition twice - reteach the trade
- Quick finishers - hand out Extension

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- Show me 25 in TWO ways. Draw the tens and ones, or write the number sentence

DO:
- Collect slips

TEACHER NOTES:
Exit ticket assesses SC2 and SC3. Sort into "two valid ways", "one valid way", "neither".

WATCH FOR:
- Two valid ways = SC2 + SC3 secure
- One valid way = SC2 only; small group on the trade

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Check our success criteria
- I can show 24 as 2 tens 4 ones - thumbs
- I can show 24 as 1 ten 14 ones - thumbs
- I can write a number sentence - thumbs
- Tell your partner why the number stays the same when we trade

DO:
- Thumbs check
- Partner talk

TEACHER NOTES:
The reflection retrieves the conservation insight.

WATCH FOR:
- Students who explain "the whole stays the same"

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, LESSON_TITLE, "Lesson 7 of 10 - Numbers to 120", "Year 1 Numeracy | AC9M1N02", NOTES_TITLE);
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Launch — 24 two ways
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "Is it still 24?", { color: STAGE_COLORS["1"] });

    addCard(s, 0.5, 1.7, 4.4, 2.5, { strip: C.PRIMARY });
    s.addText("Way 1", { x: 0.7, y: 1.8, w: 4.0, h: 0.4, fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0 });
    addBaseTenBlocks(s, 0.7, 2.4, 0, 2, 4, { unit: 0.2, color: C.PRIMARY });

    addCard(s, 5.1, 1.7, 4.4, 2.5, { strip: C.SECONDARY });
    s.addText("Way 2", { x: 5.3, y: 1.8, w: 4.0, h: 0.4, fontSize: 18, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0 });
    addBaseTenBlocks(s, 5.3, 2.4, 0, 1, 14, { unit: 0.16, color: C.SECONDARY });

    addInstructionCard(s, [
      { role: "header", text: "Both pictures show 24" },
      { role: "body", text: "How can that be? Tell your partner." },
    ], { x: 0.5, y: 4.3, w: 9, h: 0.8, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // DR1 — Tens digit of 56
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many tens in 56?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.5, "56", { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "5 tens", {
        x: 3.0, y: 4.55, w: 4.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // DR2 — 8 + ? = 10
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Pair to make 10?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 1.6, 2.0, "8", { fontSize: 100 });
      s.addText("+", { x: 4.2, y: 1.9, w: 0.7, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, 4.9, 1.9, 1.6, 2.0, "?", { fontSize: 100, color: C.ALERT, stroke: C.ALERT });
      s.addText("= 10", { x: 6.6, y: 1.9, w: 1.8, h: 2.0, fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });

      addInstructionCard(s, [
        { role: "body", text: "Write the missing number." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "2", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR2_A);
    }
  );

  // DR3 — after 99
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "What comes after 99?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.0, 1.9, 2.4, 2.0, "99", { fontSize: 110 });
      s.addText("?", { x: 5.6, y: 1.9, w: 2.4, h: 2.0, fontSize: 110, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer. Show me!" },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "100", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // Fluency 1 — Rainbow facts quick fire
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Pairs to 10 - quick fire", { color: STAGE_COLORS["1"] });

    const pairs = [
      { from: "3", to: "7" },
      { from: "4", to: "6" },
      { from: "1", to: "9" },
      { from: "7", to: "3" },
      { from: "9", to: "1" },
    ];
    pairs.forEach((p, i) => {
      const x = 0.55 + i * 1.8;
      const y = 2.0;
      drawNumeralCard(s, x, y, 0.7, 0.9, p.from, { fontSize: 30 });
      s.addText("+", { x: x + 0.75, y: y, w: 0.3, h: 0.9, fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, x + 1.05, y, 0.65, 0.9, p.to, { fontSize: 30, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });
    });

    addInstructionCard(s, [
      { role: "header", text: "I say the first number. You say the partner." },
    ], { x: 0.5, y: 4.0, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL1);
  })();

  // Fluency 2 — 10 more quick fire
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "10 more - quick fire", { color: STAGE_COLORS["1"] });

    const examples = [
      { from: "32", to: "42" },
      { from: "18", to: "28" },
      { from: "47", to: "57" },
      { from: "65", to: "75" },
    ];
    examples.forEach((ex, i) => {
      const x = 0.6 + (i % 2) * 4.5;
      const y = 1.8 + Math.floor(i / 2) * 1.3;
      drawNumeralCard(s, x, y, 1.4, 1.1, ex.from, { fontSize: 48 });
      s.addText("+10", { x: x + 1.5, y: y, w: 0.7, h: 1.1, fontSize: 20, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, x + 2.3, y, 1.4, 1.1, ex.to, { fontSize: 48, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });
    });

    addInstructionCard(s, [
      { role: "header", text: "+ 10 each time" },
    ], { x: 0.5, y: 4.5, w: 9, h: 0.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL2);
  })();

  // Vocabulary — partition
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Word", { color: C.PRIMARY });
    addTitle(s, "Partition", { color: C.PRIMARY });

    // Show 24 partitioned three ways
    addCard(s, 0.5, 1.8, 2.85, 2.6, { strip: C.PRIMARY });
    s.addText("Way 1", { x: 0.7, y: 1.9, w: 2.5, h: 0.35, fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0 });
    addBaseTenBlocks(s, 0.7, 2.4, 0, 2, 4, { unit: 0.14, color: C.PRIMARY });
    s.addText("20 + 4", { x: 0.7, y: 3.9, w: 2.5, h: 0.4, fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });

    addCard(s, 3.6, 1.8, 2.85, 2.6, { strip: C.SECONDARY });
    s.addText("Way 2", { x: 3.8, y: 1.9, w: 2.5, h: 0.35, fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0 });
    addBaseTenBlocks(s, 3.8, 2.4, 0, 1, 14, { unit: 0.11, color: C.SECONDARY });
    s.addText("10 + 14", { x: 3.6, y: 3.9, w: 2.85, h: 0.4, fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });

    addCard(s, 6.7, 1.8, 2.85, 2.6, { strip: C.ACCENT });
    s.addText("Way 3", { x: 6.9, y: 1.9, w: 2.5, h: 0.35, fontSize: 14, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0 });
    addBaseTenBlocks(s, 6.9, 2.4, 0, 0, 24, { unit: 0.07, color: C.ACCENT });
    s.addText("0 + 24", { x: 6.7, y: 3.9, w: 2.85, h: 0.4, fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0 });

    addInstructionCard(s, [
      { role: "header", text: "All three ways show 24" },
    ], { x: 0.5, y: 4.5, w: 9, h: 0.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // LI/SC
  liSlide(pres,
    ["We are learning to show one number in different ways."],
    [
      "I can show 24 as 2 tens and 4 ones.",
      "I can show 24 as 1 ten and 14 ones.",
      "I can write a number sentence for each way.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do Way 1 — 24 = 20 + 4
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Way 1: the normal way", { color: STAGE_COLORS["2"] });

    addBaseTenBlocks(s, 1.5, 1.95, 0, 2, 4, { unit: 0.25, color: C.PRIMARY });

    addTextOnShape(s, "24 = 20 + 4", {
      x: 0.5, y: 4.05, w: 9.0, h: 1.0, rectRadius: 0.1,
      fill: { color: C.SUCCESS },
    }, { fontSize: 32, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_WAY1);
  })();

  // I Do Way 2 — 24 = 10 + 14
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Way 2: trade a ten for ones", { color: STAGE_COLORS["2"] });

    addBaseTenBlocks(s, 0.7, 1.95, 0, 1, 14, { unit: 0.18, color: C.PRIMARY });

    addTextOnShape(s, "24 = 10 + 14", {
      x: 0.5, y: 4.05, w: 9.0, h: 1.0, rectRadius: 0.1,
      fill: { color: C.SUCCESS },
    }, { fontSize: 32, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_WAY2);
  })();

  // I Do Way 3 — all ones (brief)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Way 3: all ones", { color: STAGE_COLORS["2"] });

    addBaseTenBlocks(s, 0.7, 2.1, 0, 0, 24, { unit: 0.12, color: C.PRIMARY });

    addTextOnShape(s, "24 = 24 ones", {
      x: 0.5, y: 4.05, w: 9.0, h: 1.0, rectRadius: 0.1,
      fill: { color: C.SUCCESS },
    }, { fontSize: 32, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_WAY3);
  })();

  // CFU — 1 ten + 14 ones = ?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "What number is this?", { color: C.ALERT });

      addBaseTenBlocks(s, 1.0, 1.95, 0, 1, 14, { unit: 0.2, color: C.PRIMARY });

      addInstructionCard(s, [
        { role: "body", text: "Count the tens AND the ones. Write the number on your whiteboard." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "24 - 1 ten + 14 ones!", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // We Do — 26 two ways
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Show 26 two ways", { color: STAGE_COLORS["3"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.0, "26", { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addInstructionCard(s, [
        { role: "header", text: "On your whiteboard" },
        { role: "body", text: "Way 1: ___ tens + ___ ones" },
        { role: "body", text: "Way 2: ___ tens + ___ ones (use only 1 ten!)" },
      ], { x: 0.5, y: 3.85, w: 9, h: 1.25, strip: C.SECONDARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Way 1: 2 tens + 6 ones  -  Way 2: 1 ten + 16 ones", {
        x: 0.3, y: 4.55, w: 9.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge CFU — Which is the same as 35?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "Which is the SAME as 35?", { color: C.ALERT });

      const opts = [
        { x: 0.5, label: "A", text: "3 tens + 5 ones" },
        { x: 3.6, label: "B", text: "2 tens + 15 ones" },
        { x: 6.7, label: "C", text: "Both A and B" },
      ];
      opts.forEach((o) => {
        addCard(s, o.x, 1.9, 2.8, 2.1, { strip: C.ALERT });
        s.addText(o.label, {
          x: o.x + 0.15, y: 1.95, w: 0.5, h: 0.4,
          fontSize: 24, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", margin: 0,
        });
        s.addText(o.text, {
          x: o.x + 0.15, y: 2.5, w: 2.5, h: 1.3,
          fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0,
        });
      });

      addInstructionCard(s, [
        { role: "body", text: "Hold up 1, 2 or 3 fingers. Show me!" },
      ], { x: 0.5, y: 4.15, w: 9, h: 1.0, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "C - BOTH! Same whole, different parts.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Show two ways",
    [
      "First - show the normal way (tens and ones).",
      "Next - trade 1 ten for 10 ones.",
      "Then - write the number sentence for each way.",
      "",
      "You have 10 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW,
        Math.min(2.5, SAFE_BOTTOM - lg.panelTopPadded), { strip: C.ALERT });
      slide.addText("Remember", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "The number stays the same.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Trade 1 ten for 10 ones.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Two ways - same whole!", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Exit Ticket
  exitTicketSlide(pres,
    ["Show 25 in two different ways. Draw the blocks OR write the number sentences."],
    NOTES_EXIT, FOOTER);

  // Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner why the number stays the same when we trade a ten for ten ones.",
    scItems: [
      "I can show 24 as 2 tens and 4 ones.",
      "I can show 24 as 1 ten and 14 ones.",
      "I can write a number sentence for each way.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1N_Lesson7_Different_Ways.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  function hex(c) { return c.startsWith("#") ? c : "#" + c; }

  // Worksheet — Show three numbers two ways each
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Show each number two ways",
      color: C.NAVY,
      lessonInfo: "Lesson 7 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Show each number the normal way AND by trading 1 ten for 10 ones. Write both number sentences.", y, { color: C.TEAL });

    function drawTwoWaysFor(doc, n, y) {
      const x = 60;
      doc.fontSize(14).font("Sans-Bold").fillColor("#000");
      doc.text("The number is " + n, x, y);
      y += 22;
      // Way 1 box
      doc.fontSize(11).font("Sans-Bold");
      doc.text("Way 1:", x, y);
      doc.rect(x + 50, y - 5, 200, 50).lineWidth(1).strokeColor("#999").stroke();
      doc.fontSize(10).font("Sans");
      doc.text("Draw the blocks", x + 60, y + 50);
      doc.fontSize(11).font("Sans-Bold");
      doc.text("_____ + _____ = " + n, x + 300, y + 15);
      y += 75;
      // Way 2 box
      doc.fontSize(11).font("Sans-Bold");
      doc.text("Way 2:", x, y);
      doc.rect(x + 50, y - 5, 200, 50).lineWidth(1).strokeColor("#999").stroke();
      doc.fontSize(10).font("Sans");
      doc.text("Trade 1 ten for ones", x + 60, y + 50);
      doc.fontSize(11).font("Sans-Bold");
      doc.text("_____ + _____ = " + n, x + 300, y + 15);
      return y + 75;
    }
    y = drawTwoWaysFor(doc, 24, y);
    y = drawTwoWaysFor(doc, 31, y);

    addPdfFooter(doc, "Lesson 7 | Numbers to 120 | Year 1");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Two ways - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Lesson 7 of 10 | Year 1 Numeracy",
    });
    y = addSectionHeading(doc, "24", y, { color: C.NAVY });
    y = addBodyText(doc, "Way 1: 2 tens + 4 ones = 24  (20 + 4)", y);
    y = addBodyText(doc, "Way 2: 1 ten + 14 ones = 24  (10 + 14)", y);
    y = addSectionHeading(doc, "31", y, { color: C.NAVY });
    y = addBodyText(doc, "Way 1: 3 tens + 1 one = 31  (30 + 1)", y);
    y = addBodyText(doc, "Way 2: 2 tens + 11 ones = 31  (20 + 11)", y);
    y = addBodyText(doc, "Watch for: students writing the SAME way twice. Prompt for the trade.", y);
    addPdfFooter(doc, "Lesson 7 | Answer Key | Year 1");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Find all the ways",
      color: C.TEAL,
      lessonInfo: "Lesson 7 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Find as many different ways as you can to partition each number. Use tens and ones.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Partition 32", y, { color: C.NAVY });
    y = addBodyText(doc, "_____ tens + _____ ones = 32", y);
    y = addBodyText(doc, "_____ tens + _____ ones = 32", y);
    y = addBodyText(doc, "_____ tens + _____ ones = 32", y);
    y = addBodyText(doc, "_____ tens + _____ ones = 32", y);

    y = addSectionHeading(doc, "Partition 45", y, { color: C.NAVY });
    y = addBodyText(doc, "_____ tens + _____ ones = 45", y);
    y = addBodyText(doc, "_____ tens + _____ ones = 45", y);
    y = addBodyText(doc, "_____ tens + _____ ones = 45", y);
    y = addBodyText(doc, "_____ tens + _____ ones = 45", y);
    y = addBodyText(doc, "_____ tens + _____ ones = 45", y);

    addPdfFooter(doc, "Lesson 7 | Extension | Year 1");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
