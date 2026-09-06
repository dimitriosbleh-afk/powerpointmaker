"use strict";

// Grade 1 Numeracy — Lesson 5: Part-Part-Whole Facts to 10
// AC9M1N02 — building part-part-whole facts to 10 using ten frames; finding
//            pairs of numbers that combine to make 10.
// Daily Review: Counting and place value (read past 100, fill missing, order).
// Fluency:      Skip counting, addition and patterns (rainbow facts to 10).

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

const SESSION = 5;
const TOTAL = 10;
const UNIT_TITLE = "Numbers to 120";
const LESSON_TITLE = "Pairs that Make 10";
const FOOTER = `${UNIT_TITLE} | Lesson ${SESSION} of ${TOTAL} | Year 1 Numeracy`;
const OUT_DIR = `output/G1N_Lesson${SESSION}_Pairs_That_Make_10`;
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 5 Pairs That Make 10",
  "Complete the ten-frame pairs and write the matching number sentence.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 5 Answer Key",
  "Teacher reference with all rainbow facts to 10.");
const EXTENDING_RES = makeSessionResource(SESSION, "Lesson 5 Extension",
  "Investigate part-part-whole for numbers 7, 8 and 9.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

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

// Tens frame with two-colour fill: redCount cells red, blueCount cells blue.
function addTwoColourTensFrame(slide, x, y, w, redCount, blueCount, opts) {
  const o = opts || {};
  const cellW = w / 5;
  const cellH = o.cellH || cellW;
  const redColor = o.redColor || C.ALERT;
  const blueColor = o.blueColor || C.PRIMARY;
  const borderColor = o.borderColor || C.CHARCOAL;
  for (let r = 0; r < 2; r += 1) {
    for (let c = 0; c < 5; c += 1) {
      const idx = r * 5 + c;
      const cx = x + c * cellW;
      const cy = y + r * cellH;
      slide.addShape("rect", {
        x: cx, y: cy, w: cellW, h: cellH,
        fill: { color: C.WHITE },
        line: { color: borderColor, width: 1.5 },
      });
      let fill = null;
      if (idx < redCount) fill = redColor;
      else if (idx < redCount + blueCount) fill = blueColor;
      if (fill) {
        const dotW = cellW * 0.62;
        slide.addShape("roundRect", {
          x: cx + (cellW - dotW) / 2,
          y: cy + (cellH - dotW) / 2,
          w: dotW, h: dotW,
          rectRadius: dotW / 2,
          fill: { color: fill },
        });
      }
    }
  }
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we hunt for pairs of numbers that make 10
- 10 is a special number - it is the top of our ten frame
- We will find ALL the ways to make 10

DO:
- Display title
- Have ten frames and red/blue counters ready
- Have rainbow facts poster ready

TEACHER NOTES:
Lesson 5 of 10. The pairs to 10 are foundational for the rest of Year 1 maths and beyond. Use one ten frame and two counter colours.

WATCH FOR:
- Students excited by the hunt - lean into the discovery framing

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today we use red and blue counters and a ten frame
- The worksheet shows pairs to 10 with the matching number sentence

DO:
- Show the red/blue counters at the front
- Have ten frames at each table

TEACHER NOTES:
Counter colours matter today - the two parts visually contrast. Use red and blue, or any two strong colours.

WATCH FOR:
- Missing counters or ten frames

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_LAUNCH = `SAY:
- I have 4 red counters on the ten frame
- How many more do I need to make 10? Whisper to your partner

DO:
- Display the ten frame with 4 red counters
- Allow 20 seconds partner talk
- Cold call

TEACHER NOTES:
Activates the "how many more to 10" question that drives the lesson.

WATCH FOR:
- Students who count empty cells - they have the strategy
- Students who count up from 4 to 10 - also valid

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_DR1_Q = `SAY:
- Daily Review. Write the number that comes after 109

DO:
- Display 109
- Show me

TEACHER NOTES:
Retrieves Lesson 4 (past 100).

WATCH FOR:
- Quick 110

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- 110. After 109 comes 110

DO:
- Reveal 110

TEACHER NOTES:
Quick.

WATCH FOR:
- Chorus

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- Skip count by 10s. 30, blank, 50, 60

DO:
- 8 seconds, show me

TEACHER NOTES:
Retrieves Lesson 3.

WATCH FOR:
- 40

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- 40

DO:
- Reveal 40

TEACHER NOTES:
Brief.

WATCH FOR:
- Confident class

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- Read the number - point and chorus

DO:
- Display 87
- Point, chorus "eighty-seven"

TEACHER NOTES:
Retrieves number reading.

WATCH FOR:
- Choral chorus

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- Eighty-seven

DO:
- Reveal "eighty-seven"

TEACHER NOTES:
Quick.

WATCH FOR:
- Class confident

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1 = `SAY:
- Fluency. Pairs that make 10 - the rainbow facts
- Listen first - 1 and 9, 2 and 8, 3 and 7, 4 and 6, 5 and 5
- Say them with me

DO:
- Display the rainbow facts chart
- Lead a choral say of each pair

TEACHER NOTES:
Today's lesson IS the fluency. This early echo plants the rhythm.

WATCH FOR:
- Confident voices on "5 and 5" - the easiest fact
- Hesitation on "3 and 7" - reteach

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL2_Q = `SAY:
- Quick fire - what makes 10 with 6?
- Write on your whiteboard

DO:
- Display "6 and ? make 10"
- 5 seconds, show me

TEACHER NOTES:
First quick-fire pair.

WATCH FOR:
- 4

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_FL2_A = `SAY:
- 4. 6 and 4 make 10

DO:
- Reveal 4

TEACHER NOTES:
Brief.

WATCH FOR:
- Chorus

[Stage 1: Fluency 2 Answer | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- Two new words - parts and whole
- The WHOLE is the total - 10
- The PARTS are the two numbers that make the whole

DO:
- Display a part-part-whole diagram with 10 at the top, 6 and 4 below
- Tap the labels

TEACHER NOTES:
Part-part-whole language is the heart of this lesson and Lessons 6-7.

WATCH FOR:
- Students who use the words back to you - praise

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_LI_SC = `SAY:
- Read with me - We are learning to find pairs of numbers that make 10
- I can use a ten frame to show two parts that make 10
- I can find the missing part when I see one part
- I can write a number sentence for pairs that make 10

DO:
- Choral read

TEACHER NOTES:
SC1 is reachable for everyone. SC2 is the target. SC3 introduces number sentences (10 = 6 + 4).

WATCH FOR:
- Engaged class

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_3AND7 = `SAY:
- Watch me find a pair that makes 10
- I put 3 red counters in the ten frame
- Then I put blue counters in the empty spaces. 1, 2, 3, 4, 5, 6, 7 blue
- 3 red and 7 blue. 3 plus 7 equals 10
- I write the number sentence - 3 + 7 = 10

DO:
- Use real counters on the document camera
- Count each blue counter as you place it
- Write 3 + 7 = 10 underneath

TEACHER NOTES:
Modelling the count-up-to-10 strategy. The ten frame makes the second part visible.

WATCH FOR:
- Students mouthing the count - good

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_8AND2 = `SAY:
- Another one. 8 red counters
- How many blue spaces are left? 1, 2 blue
- 8 plus 2 equals 10
- 8 + 2 = 10

DO:
- Display ten frame with 8 red
- Add 2 blue, counting aloud
- Write 8 + 2 = 10

TEACHER NOTES:
Variation - more of one colour, fewer of the other. Reinforces that the parts can be uneven.

WATCH FOR:
- Students who say "2" before you finish counting - they are subitising; praise

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. The ten frame has 5 red. How many more to make 10?
- Write on your whiteboard

DO:
- Display ten frame with 5 red
- 8 seconds, show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "How many MORE to make 10? Write it. Show me"
- Scan for: 5
PROCEED: If 80%+ show 5, move to We Do.
PIVOT: Most likely misconception - students write 10 (total) or 4 (off by one). Reteach with real counters - count the empty cells aloud. Re-check with 6 red.

TEACHER NOTES:
The "5 and 5" fact is the easiest. If students miss this, the rest will be hard.

WATCH FOR:
- 5 written = good
- 4 or 6 = small-group support

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- 5. 5 and 5 make 10

DO:
- Reveal 5

TEACHER NOTES:
Easy reveal.

WATCH FOR:
- All in

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. The ten frame shows 7 red
- How many blue counters do I need to make 10? Write it
- Then write the number sentence

DO:
- Display ten frame with 7 red
- 30 seconds
- Cold call

TEACHER NOTES:
Two-part response - the missing part AND the number sentence.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use real counters and a paper ten frame. Place 7 red, count empty cells, place 7 - actually 3 - blue.
- Extra Notes: Small group at the front with the teacher and physical materials.
EXTENDING PROMPT:
- Task: After you find the answer, write the matching subtraction - 10 - 7 = 3.
- Extra Notes: Whole-class extension; can be done from seats.

WATCH FOR:
- Students writing "10 = 7 + 3" or "7 + 3 = 10" - both are correct
- Students who only write the missing number, not the sentence - prompt for the sentence

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- 3 blue. 7 + 3 = 10

DO:
- Reveal the answer with both the visual (3 blue) and the sentence

TEACHER NOTES:
Show the sentence form clearly.

WATCH FOR:
- Confident class

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge check. Which two numbers make 10?
- A: 4 and 7
- B: 6 and 4
- C: 5 and 3
- Hold up A, B or C on your fingers

DO:
- Display the three options
- 10 seconds, show me

CFU CHECKPOINT:
Technique: Finger Voting (1 for A, 2 for B, 3 for C)
Script:
- "1 finger for A, 2 for B, 3 for C. Show me"
- Scan for: 2 fingers (B is correct - 6 + 4 = 10)
PROCEED: If 80%+ show 2, move to You Do.
PIVOT: Most likely misconception - students choose A (4 + 7 = 11). Reteach by building each pair on a ten frame and checking if it fills exactly. Re-check with "9 and 1" vs "8 and 3".

TEACHER NOTES:
This hinge tests whether students can recognise a correct pair without building first.

WATCH FOR:
- Students who pick B confidently - they have the facts
- Students choosing A or C - need more ten-frame work

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- B - 6 and 4 make 10. 4 + 7 = 11, 5 + 3 = 8

DO:
- Reveal B with the matching ten frame

TEACHER NOTES:
Final hinge.

WATCH FOR:
- Confident class = ready

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent work
- Take the worksheet
- First - look at the ten frame
- Next - work out the missing part
- Then - write the number sentence

DO:
- Distribute worksheet
- Circulate
- Hand out Extension when ready

TEACHER NOTES:
Worksheet has 6 ten-frame pairs and a blank ten-frame at the bottom for the student's own pair.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work with the teacher using real counters and a ten frame. Build each pair before writing.
- Extra Notes: Small group at the back.
EXTENDING PROMPT:
- Task: Lesson 5 Extension - find all the pairs for 7, 8 and 9 (not just 10).
- Extra Notes: Distribute Lesson 5 Extension PDF.

WATCH FOR:
- Students writing 4 + 7 = 10 - they need ten-frame work
- Students completing quickly - hand out Extension

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- 7 red counters are in the ten frame
- How many MORE counters to make 10? Write it

DO:
- Collect slips

TEACHER NOTES:
Exit ticket assesses SC2. Sort into 3 correct, other.

WATCH FOR:
- Slips with 3 = ready for Lesson 6 (tens-and-ones partition)

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Check our success criteria
- I can use a ten frame - thumbs
- I can find the missing part - thumbs
- I can write a number sentence - thumbs
- Tell your partner one pair that makes 10

DO:
- Run thumbs check
- Partner talk

TEACHER NOTES:
The reflection retrieves a fact - the rainbow facts will keep coming back.

WATCH FOR:
- Variety of pairs shared

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, LESSON_TITLE, "Lesson 5 of 10 - Numbers to 120", "Year 1 Numeracy | AC9M1N02", NOTES_TITLE);
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch — 4 red, how many more to 10?
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "How many more to make 10?", { color: STAGE_COLORS["1"] });

    addTwoColourTensFrame(s, 2.5, 1.9, 5.0, 4, 0, { cellH: 1.0 });

    addInstructionCard(s, [
      { role: "header", text: "Look - 4 red counters" },
      { role: "body", text: "Whisper to your partner: how many more to make 10?" },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // DR1 — after 109
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "What comes after 109?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.0, 1.9, 2.5, 2.0, "109", { fontSize: 100 });
      s.addText("?", {
        x: 5.8, y: 1.9, w: 2.0, h: 2.0,
        fontSize: 100, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0,
      });

      addInstructionCard(s, [
        { role: "body", text: "Write the answer on your whiteboard." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "110", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // DR2 — Skip count by 10s
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Missing 10s number?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 0.8, 2.0, 1.7, 1.6, "30", { fontSize: 70 });
      drawNumeralCard(s, 2.8, 2.0, 1.7, 1.6, "?", { fontSize: 70, color: C.ALERT, stroke: C.ALERT });
      drawNumeralCard(s, 4.8, 2.0, 1.7, 1.6, "50", { fontSize: 70 });
      drawNumeralCard(s, 6.8, 2.0, 1.7, 1.6, "60", { fontSize: 70 });

      addInstructionCard(s, [
        { role: "body", text: "Write the missing number. Show me!" },
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

  // DR3 — Read 87
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Read this numeral", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.5, "87", { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "eighty-seven", {
        x: 2.5, y: 4.55, w: 5.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // Fluency 1 — Rainbow facts overview
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Rainbow facts: pairs to 10", { color: STAGE_COLORS["1"] });

    // 5 pairs of small ten frames in a row
    const pairs = [
      { red: 1, blue: 9 }, { red: 2, blue: 8 }, { red: 3, blue: 7 },
      { red: 4, blue: 6 }, { red: 5, blue: 5 },
    ];
    pairs.forEach((p, i) => {
      const x = 0.55 + i * 1.8;
      addTwoColourTensFrame(s, x, 1.85, 1.55, p.red, p.blue, { cellH: 0.31 });
      s.addText(`${p.red} + ${p.blue}`, {
        x, y: 3.2, w: 1.55, h: 0.4,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
    });

    addInstructionCard(s, [
      { role: "header", text: "All the pairs that make 10" },
      { role: "body", text: "Say them with me: 1 and 9, 2 and 8, 3 and 7, 4 and 6, 5 and 5" },
    ], { x: 0.5, y: 3.85, w: 9, h: 1.2, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL1);
  })();

  // Fluency 2 — Quick fire 6 + ? = 10
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "What makes 10 with 6?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.5, 1.9, 1.6, 2.0, "6", { fontSize: 100 });
      s.addText("+", {
        x: 4.2, y: 1.9, w: 0.7, h: 2.0,
        fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0,
      });
      drawNumeralCard(s, 4.9, 1.9, 1.6, 2.0, "?", { fontSize: 100, color: C.ALERT, stroke: C.ALERT });
      s.addText("= 10", {
        x: 6.6, y: 1.9, w: 1.8, h: 2.0,
        fontSize: 60, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0,
      });

      addInstructionCard(s, [
        { role: "body", text: "Write the missing number on your whiteboard." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FL2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "4 - so 6 + 4 = 10", {
        x: 2.5, y: 4.55, w: 5.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FL2_A);
    }
  );

  // Vocabulary — parts and whole
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Words", { color: C.PRIMARY });
    addTitle(s, "Parts and Whole", { color: C.PRIMARY });

    // Part-Part-Whole diagram
    addTextOnShape(s, "10", {
      x: 4.0, y: 1.8, w: 2.0, h: 0.9, rectRadius: 0.1,
      fill: { color: C.PRIMARY },
    }, { fontSize: 36, fontFace: FONT_H, color: C.WHITE, bold: true });
    s.addText("whole", { x: 6.2, y: 2.0, w: 1.5, h: 0.5, fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, valign: "middle", margin: 0 });

    addTextOnShape(s, "6", {
      x: 2.5, y: 3.0, w: 1.6, h: 0.9, rectRadius: 0.1,
      fill: { color: C.SECONDARY },
    }, { fontSize: 36, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "4", {
      x: 5.9, y: 3.0, w: 1.6, h: 0.9, rectRadius: 0.1,
      fill: { color: C.ACCENT },
    }, { fontSize: 36, fontFace: FONT_H, color: C.WHITE, bold: true });
    s.addText("part", { x: 1.0, y: 3.2, w: 1.5, h: 0.5, fontSize: 16, fontFace: FONT_B, color: C.SECONDARY, bold: true, valign: "middle", margin: 0 });
    s.addText("part", { x: 7.7, y: 3.2, w: 1.5, h: 0.5, fontSize: 16, fontFace: FONT_B, color: C.ACCENT, bold: true, valign: "middle", margin: 0 });

    addInstructionCard(s, [
      { role: "header", text: "Two parts make the whole" },
      { role: "body", text: "6 + 4 = 10" },
    ], { x: 0.5, y: 4.1, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // LI/SC
  liSlide(pres,
    ["We are learning to find pairs of numbers that make 10."],
    [
      "I can use a ten frame to show two parts that make 10.",
      "I can find the missing part when I see one part.",
      "I can write a number sentence for pairs that make 10.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do — 3 and 7 make 10
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "3 and 7 make 10", { color: STAGE_COLORS["2"] });

    addTwoColourTensFrame(s, 0.7, 1.8, 5.0, 3, 7, { cellH: 1.0 });

    addTextOnShape(s, "3 + 7 = 10", {
      x: 6.3, y: 2.4, w: 3.0, h: 0.9, rectRadius: 0.1,
      fill: { color: C.SUCCESS },
    }, { fontSize: 32, fontFace: FONT_H, color: C.WHITE, bold: true });

    addInstructionCard(s, [
      { role: "header", text: "3 red counters, 7 blue counters" },
      { role: "body", text: "The ten frame is full - 10 in total!" },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_3AND7);
  })();

  // I Do — 8 and 2 make 10
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "8 and 2 make 10", { color: STAGE_COLORS["2"] });

    addTwoColourTensFrame(s, 0.7, 1.8, 5.0, 8, 2, { cellH: 1.0 });

    addTextOnShape(s, "8 + 2 = 10", {
      x: 6.3, y: 2.4, w: 3.0, h: 0.9, rectRadius: 0.1,
      fill: { color: C.SUCCESS },
    }, { fontSize: 32, fontFace: FONT_H, color: C.WHITE, bold: true });

    addInstructionCard(s, [
      { role: "header", text: "8 red, only 2 blue" },
      { role: "body", text: "Bigger and smaller parts. Still makes 10!" },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_8AND2);
  })();

  // CFU — 5 red, how many more
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How many more to make 10?", { color: C.ALERT });

      addTwoColourTensFrame(s, 2.5, 1.85, 5.0, 5, 0, { cellH: 1.0 });

      addInstructionCard(s, [
        { role: "body", text: "Write the missing number on your whiteboard. Show me!" },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "5 - so 5 + 5 = 10!", {
        x: 2.5, y: 4.55, w: 5.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // We Do — 7 red, find blue
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Find the missing part", { color: STAGE_COLORS["3"] });

      addTwoColourTensFrame(s, 2.5, 1.85, 5.0, 7, 0, { cellH: 1.0 });

      addInstructionCard(s, [
        { role: "header", text: "How many blue counters?" },
        { role: "body", text: "Write the number AND the number sentence." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.SECONDARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "3 blue   7 + 3 = 10", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge CFU — Which pair makes 10?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "Which pair makes 10?", { color: C.ALERT });

      // Three options A, B, C
      const opts = [
        { x: 0.5, label: "A", a: 4, b: 7 },
        { x: 3.6, label: "B", a: 6, b: 4 },
        { x: 6.7, label: "C", a: 5, b: 3 },
      ];
      opts.forEach((o) => {
        addCard(s, o.x, 1.8, 2.8, 2.0, { strip: C.ALERT });
        s.addText(o.label, {
          x: o.x + 0.1, y: 1.85, w: 0.4, h: 0.4,
          fontSize: 20, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", margin: 0,
        });
        s.addText(`${o.a} + ${o.b}`, {
          x: o.x, y: 2.4, w: 2.8, h: 0.6,
          fontSize: 28, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
        });
        addTwoColourTensFrame(s, o.x + 0.4, 3.05, 2.0, o.a, o.b, { cellH: 0.4 });
      });

      addInstructionCard(s, [
        { role: "body", text: "Hold up 1 for A, 2 for B, 3 for C. Show me!" },
      ], { x: 0.5, y: 4.1, w: 9, h: 1.0, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "B - 6 + 4 = 10!", {
        x: 3.0, y: 4.55, w: 4.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Find pairs to make 10",
    [
      "First - look at the ten frame.",
      "Next - work out the missing part.",
      "Then - write the number sentence.",
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
        { text: "Count the empty cells.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Two parts make the whole.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Rainbow facts to 10!", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Exit Ticket
  exitTicketSlide(pres,
    ["7 red counters are in the ten frame. How many MORE to make 10?"],
    NOTES_EXIT, FOOTER);

  // Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner one pair of numbers that makes 10.",
    scItems: [
      "I can use a ten frame to show two parts that make 10.",
      "I can find the missing part when I see one part.",
      "I can write a number sentence for pairs that make 10.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1N_Lesson5_Pairs_That_Make_10.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  function hex(c) { return c.startsWith("#") ? c : "#" + c; }

  function drawPdfTwoColourFrame(doc, x, y, cellSize, redCount, blueCount) {
    for (let r = 0; r < 2; r += 1) {
      for (let c = 0; c < 5; c += 1) {
        const idx = r * 5 + c;
        const cx = x + c * cellSize;
        const cy = y + r * cellSize;
        doc.rect(cx, cy, cellSize, cellSize).lineWidth(1).strokeColor("#444").stroke();
        if (idx < redCount) {
          doc.circle(cx + cellSize / 2, cy + cellSize / 2, cellSize * 0.32)
            .fill(hex(C.ALERT));
        } else if (idx < redCount + blueCount) {
          doc.circle(cx + cellSize / 2, cy + cellSize / 2, cellSize * 0.32)
            .fill(hex(C.PRIMARY));
        }
      }
    }
  }

  // Worksheet — 6 ten-frame pairs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Pairs that make 10",
      color: C.NAVY,
      lessonInfo: "Lesson 5 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Look at the red counters. Write the missing number. Write the number sentence.", y, { color: C.TEAL });

    function drawPair(doc, num, redCount, y) {
      const x = 60;
      doc.fontSize(12).font("Sans-Bold").fillColor("#000");
      doc.text(num + ".", x, y);
      drawPdfTwoColourFrame(doc, x + 25, y, 28, redCount, 0);
      // "+ ___ = 10" prompt
      doc.fontSize(16).font("Sans-Bold").fillColor("#000");
      doc.text(redCount + "  +  _____  =  10", x + 200, y + 14);
      return y + 70;
    }
    y = drawPair(doc, "1", 2, y);
    y = drawPair(doc, "2", 4, y);
    y = drawPair(doc, "3", 6, y);
    y = drawPair(doc, "4", 8, y);
    y = drawPair(doc, "5", 1, y);
    y = drawPair(doc, "6", 9, y);

    addPdfFooter(doc, "Lesson 5 | Numbers to 120 | Year 1");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Pairs that make 10 - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Lesson 5 of 10 | Year 1 Numeracy",
    });
    y = addSectionHeading(doc, "Rainbow facts to 10", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  2 + 8 = 10", y);
    y = addBodyText(doc, "2.  4 + 6 = 10", y);
    y = addBodyText(doc, "3.  6 + 4 = 10", y);
    y = addBodyText(doc, "4.  8 + 2 = 10", y);
    y = addBodyText(doc, "5.  1 + 9 = 10", y);
    y = addBodyText(doc, "6.  9 + 1 = 10", y);
    y = addBodyText(doc, "Note: students may also write the answer as 10 = X + Y - both are correct.", y);
    y = addBodyText(doc, "Watch for: 4 + 7 type errors - reteach with real ten frames and counters.", y);
    addPdfFooter(doc, "Lesson 5 | Answer Key | Year 1");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension - pairs for 7, 8, 9
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Find all the parts for 7, 8 and 9",
      color: C.TEAL,
      lessonInfo: "Lesson 5 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Find as many pairs as you can for each whole number.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Pairs that make 7", y, { color: C.NAVY });
    y = addBodyText(doc, "_____ + _____ = 7        _____ + _____ = 7", y);
    y = addBodyText(doc, "_____ + _____ = 7        _____ + _____ = 7", y);

    y = addSectionHeading(doc, "Pairs that make 8", y, { color: C.NAVY });
    y = addBodyText(doc, "_____ + _____ = 8        _____ + _____ = 8", y);
    y = addBodyText(doc, "_____ + _____ = 8        _____ + _____ = 8", y);

    y = addSectionHeading(doc, "Pairs that make 9", y, { color: C.NAVY });
    y = addBodyText(doc, "_____ + _____ = 9        _____ + _____ = 9", y);
    y = addBodyText(doc, "_____ + _____ = 9        _____ + _____ = 9", y);

    addPdfFooter(doc, "Lesson 5 | Extension | Year 1");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
