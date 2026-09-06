"use strict";

// Grade 1 Numeracy — Lesson 3: Hundreds Chart - Tens and Ones
// AC9M1N01 — using hundreds charts to build understanding and fluency with
//            numbers; colour-code tens and ones to see the pattern.
// Daily Review: Counting and place value (read numerals, between).
// Fluency:      Skip counting, addition and patterns (count by 10s to 100).

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
  addTensFrame, addNumberTrack,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 10;
const UNIT_TITLE = "Numbers to 120";
const LESSON_TITLE = "Hundreds Chart: Tens & Ones Patterns";
const FOOTER = `${UNIT_TITLE} | Lesson ${SESSION} of ${TOTAL} | Year 1 Numeracy`;
const OUT_DIR = `output/G1N_Lesson${SESSION}_Hundreds_Chart`;
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 3 Colour the Hundreds Chart",
  "Colour the tens column and find patterns on the hundreds chart.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 3 Answer Key",
  "Teacher reference with completed hundreds chart patterns.");
const EXTENDING_RES = makeSessionResource(SESSION, "Lesson 3 Extension",
  "Find my number - hundreds chart puzzles using row and column clues.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Hundreds chart helper ──────────────────────────────────────────────────

// Draw a 10x10 hundreds chart (1-100 or 0-99). cellW determines size.
function drawHundredsChart(slide, x, y, cellW, opts) {
  const o = opts || {};
  const startFrom = o.startFromZero ? 0 : 1;
  const highlight = o.highlight || []; // array of values to colour
  const highlightColor = o.highlightColor || C.PRIMARY;
  const highlightTextColor = o.highlightTextColor || C.WHITE;
  const baseColor = o.baseColor || C.WHITE;
  const baseTextColor = o.baseTextColor || C.CHARCOAL;
  const borderColor = o.borderColor || C.CHARCOAL;
  const fontSize = o.fontSize || Math.max(8, Math.min(14, Math.round(cellW * 26)));

  for (let r = 0; r < 10; r += 1) {
    for (let c = 0; c < 10; c += 1) {
      const value = startFrom + r * 10 + c;
      if (value > 99 + startFrom) continue;
      const cx = x + c * cellW;
      const cy = y + r * cellW;
      const hi = highlight.indexOf(value) !== -1;
      slide.addShape("rect", {
        x: cx, y: cy, w: cellW, h: cellW,
        fill: { color: hi ? highlightColor : baseColor },
        line: { color: borderColor, width: 0.7 },
      });
      slide.addText(String(value), {
        x: cx, y: cy, w: cellW, h: cellW,
        fontSize, fontFace: FONT_B, color: hi ? highlightTextColor : baseTextColor,
        bold: hi, align: "center", valign: "middle", margin: 0,
      });
    }
  }
}

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
- Today we meet a powerful new tool - the hundreds chart
- A grid with all the numbers from 1 to 100
- We will find patterns by colouring tens and ones

DO:
- Display the title
- Have crayons or coloured pencils ready
- Have a printed wall hundreds chart at the front

TEACHER NOTES:
Lesson 3 of 10. The hundreds chart is the central tool for the rest of the unit. Today is the first careful look.

WATCH FOR:
- Students who feel overwhelmed - the chart is big at first; build with rows

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Materials today - crayons, the hundreds chart worksheet, the wall chart
- Look up at the wall chart - we will use it together

DO:
- Point to wall chart
- Have worksheets stacked at the door

TEACHER NOTES:
The wall chart is essential for collective reference. If a wall chart is unavailable, project the chart slide.

WATCH FOR:
- Crayons sharpened and stacked

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_LAUNCH = `SAY:
- A grown-up has hidden 7 in a chart - that is the question for today
- Look at the chart. Can you see 7? Point to it
- Now point to 17. Now 27
- What do you notice?

DO:
- Project a hundreds chart with 7, 17, 27 highlighted faintly
- Cold call partners to share

TEACHER NOTES:
The "look down the column" insight is the heart of today. Students will notice they all end in 7.

WATCH FOR:
- Students who say "they all have a 7" - celebrate; that is the pattern

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_DR1_Q = `SAY:
- Daily Review - read this number out loud when I point

DO:
- Display 32
- Point, listen for chorus "thirty-two"

TEACHER NOTES:
Retrieves numeral reading from Lessons 1 and 2.

WATCH FOR:
- Quick chorus = ready

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- Thirty-two. 3 tens and 2 ones

DO:
- Reveal "thirty-two"

TEACHER NOTES:
Brief reveal.

WATCH FOR:
- Confident chorus

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- Find the missing number. 49, blank, 51

DO:
- Display 49, ?, 51
- Wait 8 seconds, signal "show me"

TEACHER NOTES:
Retrieves missing-number strategy from Lesson 2 and previews chart row work.

WATCH FOR:
- Students who write 50 - they have the count-on strategy across the decade

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- The answer is 50. After 49 comes 50

DO:
- Reveal "50"

TEACHER NOTES:
The 49 to 50 decade transition trips some Year 1 students. Note who pauses.

WATCH FOR:
- Students who pause - small-group support

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- Last one. Put these in order, smallest first - 27, 19, 32

DO:
- Display the three numerals
- Allow 30 seconds
- Cold call

TEACHER NOTES:
Cross-decade ordering retrieval. Students should answer 19, 27, 32.

WATCH FOR:
- Students who order by tens digit confidently - they have place value

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- 19, 27, 32. 19 is in the teens. 27 is in the twenties. 32 is in the thirties

DO:
- Reveal the order

TEACHER NOTES:
Naming the decades plants language for today's chart work.

WATCH FOR:
- Ordered class = ready

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1 = `SAY:
- Fluency. We count by 10s today
- 10, 20, 30, 40 - keep going

DO:
- Display the 10s pattern with highlighted cells
- Lead a choral count 10, 20, 30, 40, 50, 60, 70, 80, 90, 100

TEACHER NOTES:
Counting by 10s is the most important fluency for today. Connect to the highlighted column on the chart.

WATCH FOR:
- Students who chant 10, 20, 30, 40, 50 confidently - they have the tens names

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL2_Q = `SAY:
- Pattern check. 30, 40, blank, 60
- What is the missing number? Write it on your whiteboard

DO:
- Display the sequence
- 8 seconds, show me

TEACHER NOTES:
A 10s pattern with a missing number. Students should count up by 10 from 40.

WATCH FOR:
- Students who write 50 - they have the count-by-10 strategy

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_FL2_A = `SAY:
- The missing number is 50. 30, 40, 50, 60

DO:
- Reveal 50

TEACHER NOTES:
Quick reveal.

WATCH FOR:
- Hands up with correct answer = ready

[Stage 1: Fluency 2 Answer | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- A new word - column
- A column goes UP and DOWN
- This column has 6, 16, 26, 36, all the numbers ending in 6

DO:
- Display a small chart with the "6 column" highlighted
- Point your hand UP and DOWN to show direction

TEACHER NOTES:
"Column" is essential for today's chart work. Pair with "row" (which goes across) only if time allows.

WATCH FOR:
- Students who confuse row and column - use the hand gesture

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_LI_SC = `SAY:
- Read with me - We are learning to use a hundreds chart to find tens and ones
- I can find a number on the hundreds chart
- I can colour a column to show numbers ending in the same ones digit
- I can use the chart to count by 10s

DO:
- Choral read

TEACHER NOTES:
SC2 is the core - colour a column. SC3 connects to the 10s pattern.

WATCH FOR:
- All students engaging

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_FIND = `SAY:
- Watch me find a number on the chart - I look for 27
- 27 is in the twenties row. I look at the third row
- 27 ends in 7. So it is in the 7 column
- Row 3, column 7 - there it is

DO:
- Display the chart with 27 highlighted
- Trace from the side (rows) and the top (columns) with your finger

TEACHER NOTES:
The "row tells me the tens, column tells me the ones" strategy is the lesson's key idea.

WATCH FOR:
- Students nodding - good
- Students still staring at the chart - sit beside in We Do

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_COLOUR = `SAY:
- Watch me colour the chart
- I colour the WHOLE column for 4
- 4, 14, 24, 34, 44 - all the numbers ending in 4
- See the pattern? They all end in 4

DO:
- Display a chart with the 4 column highlighted top-to-bottom
- Trace down the column with your finger

TEACHER NOTES:
The colouring task makes the ones pattern visible. This is the most important visual moment in the lesson.

WATCH FOR:
- Students who say "they all end in 4" - they have the insight

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. Find 35 on the chart
- Point to where 35 should be when I signal
- 1, 2, point

DO:
- Display blank-style chart (no highlight)
- Cue point
- Scan for where students are pointing

CFU CHECKPOINT:
Technique: Point and Show
Script:
- "Point to where 35 should be. Show me on three"
- Scan for: students pointing to row 4 (the 30s row), column 5 (the 5s column)
PROCEED: If 80%+ point to the right place, move to We Do.
PIVOT: Most likely misconception - students point to "any number with a 3 in it" (e.g. 53). Reteach: "Find the row that starts with 30. Then count across to 35". Re-check with 42.

TEACHER NOTES:
Point-and-show gives immediate visible response from every student.

WATCH FOR:
- Students who point to 53 - they are seeing the digits, not the position

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- 35 is here - third decade row, fifth column
- 3 tens and 5 ones

DO:
- Reveal 35 highlighted on the chart

TEACHER NOTES:
Reinforce the row-column strategy.

WATCH FOR:
- Visible "ah" moment = success

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. We will colour the 6 column together
- On the chart - 6, 16, 26, all the way down
- I have set up the chart - tell me the next number to colour

DO:
- Display the chart with 6 already coloured
- Class chorus next number 16
- Continue 26, 36, 46, 56, 66, 76, 86, 96
- Reveal full column coloured

TEACHER NOTES:
Co-constructed colouring. The class drives the pattern. This is the lesson's high-engagement moment.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Colour the column for 1 only on a 0-30 hundreds chart strip. Smaller chart and a single column.
- Extra Notes: Sit with the small group; use the same task at a slower pace.
EXTENDING PROMPT:
- Task: Once the 6 column is coloured, colour the 9 column. Then say what is the same and different about both columns.
- Extra Notes: Whole-class extension; can be done from the same chart.

WATCH FOR:
- Students who can predict the next number - they have the pattern
- Students who guess randomly - sit beside them; trace the column with them

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- We coloured the whole 6 column
- 6, 16, 26, 36, 46, 56, 66, 76, 86, 96
- They ALL end in 6. The pattern is "ten more each time"

DO:
- Reveal the complete coloured column
- Tap each number from top to bottom

TEACHER NOTES:
Naming "ten more each time" plants the count-by-10 anchor.

WATCH FOR:
- Visible recognition of the pattern

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge check. I am thinking of a number
- It is in the 50s row. It ends in 8
- What number is it? Write it on your whiteboard

DO:
- Allow 10 seconds
- Signal show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Write the number on your whiteboard. Show me"
- Scan for: 58
PROCEED: If 80%+ show 58, move to You Do.
PIVOT: Most likely misconception - students write 85 (reversed). Reteach with the chart - show the 50s row, count across to the 8 column. Re-check with "30s row, ends in 4".

TEACHER NOTES:
This hinge tests whether students can build a number from row-column clues.

WATCH FOR:
- Students who write 85 - they need the chart in front of them, not in their head

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- The answer is 58. 5 tens and 8 ones

DO:
- Reveal 58 highlighted on the chart

TEACHER NOTES:
Final hinge before independent work.

WATCH FOR:
- Confident class = ready

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent work
- Take your worksheet
- First - colour the column for 3 (3, 13, 23, all the way down)
- Next - colour the column for 7 in a different colour
- Then - write what you notice

DO:
- Distribute worksheet
- Circulate to watch column-tracing technique
- Hand out Extension when ready

TEACHER NOTES:
The 3 and 7 columns are visually similar; students should notice both columns are vertical and ten apart each time.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Colour the 3 column only. Use the 0-30 strip on the worksheet so the work is shorter.
- Extra Notes: Sit with this small group; trace the column together.
EXTENDING PROMPT:
- Task: Lesson 3 Extension - find my number puzzles with row and column clues.
- Extra Notes: Distribute Lesson 3 Extension PDF.

WATCH FOR:
- Students who jump down a row instead of column - prompt to point and trace
- Students who finish quickly - hand out Extension

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- I am thinking of a number in the 40s row that ends in 6
- Write the number on your slip

DO:
- Display the prompt
- Collect slips

TEACHER NOTES:
Exit ticket assesses SC2 (column reading). Sort into correct (46), incorrect, or reversed (64).

WATCH FOR:
- Slips with 46 = ready for Lesson 4 (extending to 120)
- Slips with 64 = reversal; small group support

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Check our success criteria
- I can find a number on the chart - thumbs
- I can colour a column - thumbs
- I can count by 10s - thumbs
- Tell your partner one thing you noticed about a column

DO:
- Run thumbs check
- Partner talk

TEACHER NOTES:
The reflection cues students to articulate the pattern they found.

WATCH FOR:
- Students explaining "they all end in the same number"

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, LESSON_TITLE, "Lesson 3 of 10 - Numbers to 120", "Year 1 Numeracy | AC9M1N01", NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch — point to 7, 17, 27 on small chart
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "What do you notice?", { color: STAGE_COLORS["1"] });

    // Small hundreds chart, 0.32" cells = 3.2" wide, fits left half
    drawHundredsChart(s, 0.7, 1.7, 0.32, {
      highlight: [7, 17, 27],
      highlightColor: C.ACCENT,
      highlightTextColor: C.WHITE,
    });

    addInstructionCard(s, [
      { role: "header", text: "Point to 7, 17 and 27" },
      { role: "body", text: "Tell your partner what you notice." },
    ], { x: 4.6, y: 1.7, w: 4.9, h: 2.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 4-5: DR1 — read 32
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Read this numeral", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.5, "32", { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "thirty-two", {
        x: 2.5, y: 4.55, w: 5.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // Slide 6-7: DR2 — missing 49, ?, 51
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Missing number?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.0, 1.9, 2.0, 1.8, "49", { fontSize: 90 });
      drawNumeralCard(s, 4.3, 1.9, 2.0, 1.8, "?", { fontSize: 90, color: C.ALERT, stroke: C.ALERT });
      drawNumeralCard(s, 6.6, 1.9, 2.0, 1.8, "51", { fontSize: 90 });

      addInstructionCard(s, [
        { role: "header", text: "On your whiteboard" },
        { role: "body", text: "Write the missing number. Show me!" },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "50", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR2_A);
    }
  );

  // Slide 8-9: DR3 — order 27, 19, 32
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Smallest to biggest", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 1.5, 1.9, 2.0, 1.8, "27", { fontSize: 90 });
      drawNumeralCard(s, 4.0, 1.9, 2.0, 1.8, "19", { fontSize: 90 });
      drawNumeralCard(s, 6.5, 1.9, 2.0, 1.8, "32", { fontSize: 90 });

      addInstructionCard(s, [
        { role: "header", text: "Write them in order on your whiteboard" },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "19 - 27 - 32", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // Slide 10: Fluency 1 — count by 10s
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Count by 10s to 100", { color: STAGE_COLORS["1"] });

    // Hundreds chart highlighting 10, 20, 30, 40, 50, 60, 70, 80, 90, (100 not in chart, show 100 as separate badge)
    drawHundredsChart(s, 0.7, 1.7, 0.32, {
      highlight: [10, 20, 30, 40, 50, 60, 70, 80, 90],
      highlightColor: C.PRIMARY,
      highlightTextColor: C.WHITE,
    });

    addInstructionCard(s, [
      { role: "header", text: "Count the highlighted numbers" },
      { role: "body", text: "10, 20, 30, 40 ... 100!" },
    ], { x: 4.6, y: 1.7, w: 4.9, h: 2.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL1);
  })();

  // Slide 11-12: Fluency 2 — Pattern (30, 40, ?, 60)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "Count by 10s: missing", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 0.8, 2.0, 1.7, 1.6, "30", { fontSize: 70 });
      drawNumeralCard(s, 2.8, 2.0, 1.7, 1.6, "40", { fontSize: 70 });
      drawNumeralCard(s, 4.8, 2.0, 1.7, 1.6, "?", { fontSize: 70, color: C.ALERT, stroke: C.ALERT });
      drawNumeralCard(s, 6.8, 2.0, 1.7, 1.6, "60", { fontSize: 70 });

      addInstructionCard(s, [
        { role: "header", text: "What is the missing number?" },
        { role: "body", text: "Write it on your whiteboard." },
      ], { x: 0.5, y: 3.85, w: 9, h: 1.2, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FL2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "50", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FL2_A);
    }
  );

  // Slide 13: Vocabulary — column
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Word", { color: C.PRIMARY });
    addTitle(s, "Column", { color: C.PRIMARY });

    drawHundredsChart(s, 0.7, 1.7, 0.32, {
      highlight: [6, 16, 26, 36, 46, 56, 66, 76, 86, 96],
      highlightColor: C.PRIMARY,
      highlightTextColor: C.WHITE,
    });

    addInstructionCard(s, [
      { role: "header", text: "A column goes UP and DOWN" },
      { role: "body", text: "This column has 6, 16, 26, 36..." },
      { role: "body", text: "They all end in 6." },
    ], { x: 4.6, y: 1.7, w: 4.9, h: 2.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 14: LI/SC
  liSlide(pres,
    ["We are learning to use a hundreds chart to find tens and ones."],
    [
      "I can find a number on the hundreds chart.",
      "I can colour a column to show numbers ending the same way.",
      "I can use the chart to count by 10s.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 15: I Do — Find 27 on the chart
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Find a number on the chart", { color: STAGE_COLORS["2"] });

    drawHundredsChart(s, 0.7, 1.65, 0.32, {
      highlight: [27],
      highlightColor: C.SUCCESS,
      highlightTextColor: C.WHITE,
    });

    addInstructionCard(s, [
      { role: "header", text: "Find 27" },
      { role: "body", text: "Row: the twenties (3rd row)." },
      { role: "body", text: "Column: the 7s." },
      { role: "body", text: "Row + column = 27!" },
    ], { x: 4.6, y: 1.65, w: 4.9, h: 3.0, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_FIND);
  })();

  // Slide 16: I Do — Colour the 4 column
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Colour a whole column", { color: STAGE_COLORS["2"] });

    drawHundredsChart(s, 0.7, 1.65, 0.32, {
      highlight: [4, 14, 24, 34, 44, 54, 64, 74, 84, 94],
      highlightColor: C.SECONDARY,
      highlightTextColor: C.WHITE,
    });

    addInstructionCard(s, [
      { role: "header", text: "The 4 column" },
      { role: "body", text: "4, 14, 24, 34, 44..." },
      { role: "body", text: "All end in 4." },
      { role: "body", text: "Each one is ten more." },
    ], { x: 4.6, y: 1.65, w: 4.9, h: 3.0, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_COLOUR);
  })();

  // Slide 17-18: CFU — point to where 35 should be
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Point to 35", { color: C.ALERT });

      drawHundredsChart(s, 0.7, 1.65, 0.32, {});

      addInstructionCard(s, [
        { role: "header", text: "Where is 35?" },
        { role: "body", text: "Point on the screen when I signal." },
      ], { x: 4.6, y: 1.65, w: 4.9, h: 2.4, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Row 4, column 5 - 35!", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // Slide 19-20: We Do — Colour the 6 column together
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Colour the 6 column together", { color: STAGE_COLORS["3"] });

      drawHundredsChart(s, 0.7, 1.65, 0.32, {
        highlight: [6],
        highlightColor: C.SECONDARY,
        highlightTextColor: C.WHITE,
      });

      addInstructionCard(s, [
        { role: "header", text: "I have started" },
        { role: "body", text: "Say the next number to colour." },
        { role: "body", text: "Keep going down the column!" },
      ], { x: 4.6, y: 1.65, w: 4.9, h: 2.6, strip: C.SECONDARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "6, 16, 26, 36, 46, 56, 66, 76, 86, 96 - ten more each time!", {
        x: 0.3, y: 4.55, w: 9.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 21-22: Hinge CFU — 50s row, ends in 8
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "What is my number?", { color: C.ALERT });

      drawHundredsChart(s, 0.7, 1.65, 0.32, {});

      addInstructionCard(s, [
        { role: "header", text: "It is in the 50s row" },
        { role: "body", text: "It ends in 8." },
        { role: "body", text: "Write the number on your whiteboard." },
      ], { x: 4.6, y: 1.65, w: 4.9, h: 2.6, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "58 - 5 tens and 8 ones!", {
        x: 1.5, y: 4.55, w: 7.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // Slide 23: You Do
  workedExSlide(pres, 4, "You Do", "Colour the 3 and 7 columns",
    [
      "First - colour the 3 column.",
      "Next - colour the 7 column.",
      "Then - write what you notice.",
      "",
      "Take your time!",
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
        { text: "Use two different colours.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Trace down the column.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Ten more each step!", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 24: Exit Ticket
  exitTicketSlide(pres,
    ["My number is in the 40s row. It ends in 6. Write the number."],
    NOTES_EXIT, FOOTER);

  // Slide 25: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner one thing you noticed about a column on the chart.",
    scItems: [
      "I can find a number on the hundreds chart.",
      "I can colour a column to show numbers ending the same way.",
      "I can use the chart to count by 10s.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1N_Lesson3_Hundreds_Chart.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  function hex(c) { return c.startsWith("#") ? c : "#" + c; }

  function drawPdfHundredsChart(doc, x, y, cellSize, highlightCols, highlightColor) {
    // Draw 10x10 chart 1-100, with the highlight column(s) coloured
    const startFrom = 1;
    for (let r = 0; r < 10; r += 1) {
      for (let c = 0; c < 10; c += 1) {
        const value = startFrom + r * 10 + c;
        if (value > 100) continue;
        const cx = x + c * cellSize;
        const cy = y + r * cellSize;
        const onesDigit = value % 10;
        const isHighlight = (highlightCols || []).indexOf(onesDigit) !== -1;
        doc.rect(cx, cy, cellSize, cellSize)
          .lineWidth(0.5).strokeColor("#666").stroke();
        if (isHighlight) {
          doc.save();
          doc.rect(cx, cy, cellSize, cellSize).fill(hex(highlightColor));
          doc.restore();
        }
        doc.fontSize(7).font("Sans").fillColor(isHighlight ? "#FFFFFF" : "#222");
        doc.text(String(value), cx + 1, cy + cellSize / 2 - 3, { width: cellSize - 2, align: "center" });
      }
    }
  }

  // Worksheet — colour 3 column and 7 column
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Colour the hundreds chart",
      color: C.NAVY,
      lessonInfo: "Lesson 3 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Colour the 3 column in one colour. Colour the 7 column in another colour. Look down each column - what do you notice?", y, { color: C.TEAL });

    // Empty hundreds chart for student colouring (no pre-highlight)
    y = addSectionHeading(doc, "Hundreds chart - colour the 3 and 7 columns", y, { color: C.NAVY });
    const cellSize = 38;
    drawPdfHundredsChart(doc, 60, y, cellSize, [], C.NAVY);
    y += cellSize * 10 + 20;

    // Notice questions
    y = addSectionHeading(doc, "What do you notice?", y, { color: C.NAVY });
    y = addBodyText(doc, "The 3 column has: __________________________________", y);
    y = addBodyText(doc, "The 7 column has: __________________________________", y);
    y = addBodyText(doc, "Both columns: ______________________________________", y);

    addPdfFooter(doc, "Lesson 3 | Numbers to 120 | Year 1");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Hundreds chart - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Lesson 3 of 10 | Year 1 Numeracy",
    });
    y = addSectionHeading(doc, "3 column", y, { color: C.NAVY });
    y = addBodyText(doc, "3, 13, 23, 33, 43, 53, 63, 73, 83, 93 - all end in 3, each is 10 more.", y);
    y = addSectionHeading(doc, "7 column", y, { color: C.NAVY });
    y = addBodyText(doc, "7, 17, 27, 37, 47, 57, 67, 77, 87, 97 - all end in 7, each is 10 more.", y);
    y = addSectionHeading(doc, "Both columns", y, { color: C.NAVY });
    y = addBodyText(doc, "Both go straight down. Both add 10 each step. They are 4 columns apart.", y);
    y = addBodyText(doc, "Sample chart with the 7 column highlighted:", y);

    drawPdfHundredsChart(doc, 60, y, 38, [7], C.PRIMARY);
    addPdfFooter(doc, "Lesson 3 | Answer Key | Year 1");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension - Find my number puzzles
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Find my number puzzles",
      color: C.TEAL,
      lessonInfo: "Lesson 3 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Use the hundreds chart to find each number. Use the row and column clues.", y, { color: C.TEAL });

    // Mini hundreds chart for reference
    drawPdfHundredsChart(doc, 60, y, 30, [], C.NAVY);
    y += 30 * 10 + 10;

    y = addSectionHeading(doc, "Find each number", y, { color: C.NAVY });
    y = addBodyText(doc, "1. I am in the 60s row. I end in 5. I am _____.", y);
    y = addBodyText(doc, "2. I am one less than 70. I am _____.", y);
    y = addBodyText(doc, "3. I am ten more than 36. I am _____.", y);
    y = addBodyText(doc, "4. I am in the 80s row. I end in 1. I am _____.", y);
    y = addBodyText(doc, "5. I am ten less than 50. I am _____.", y);

    addPdfFooter(doc, "Lesson 3 | Extension | Year 1");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
