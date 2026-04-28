"use strict";

// Subtraction Strategies (Year 2 Numeracy) — Lesson 5: Two-Digit Subtraction by Partitioning
// VC2M2N04 / VC2M2N06 — partition tens and ones to subtract; apply to a money problem.
//   Example: 45 - 23 by tens-and-ones: 40-20=20, 5-3=2, total 22.
// Daily Review: Place value (2- and 3-digit) — the importance of zero.
// Fluency: Place value — 10 more / 10 less, partitioning, expanding numbers.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Same theme as Lessons 1-4 (variant 0) — unit cohesion.
const T = createTheme("numeracy", "grade2", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  addBaseTenBlocks,
  STAGE_COLORS,
} = T;

const SESSION = 5;
const TOTAL = 5;
const UNIT_TITLE = "Subtraction Strategies";
const FOOTER = `Subtraction Strategies | Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`;
const OUT_DIR = "output/Subs_Lesson5_Two_Digit_Money";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 5 Money Problem Practice",
  "Solve two-digit take away problems and a money problem using partitioning.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 5 Answer Key",
  "Worked answers for the money problem practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a partitioning panel: shows the tens and ones split for a number.
function drawPartitionCard(slide, x, y, w, h, num, opts) {
  const o = opts || {};
  const tens = Math.floor(num / 10);
  const ones = num % 10;
  const tensColor = o.tensColor || C.PRIMARY;
  const onesColor = o.onesColor || C.SECONDARY;

  // Card frame
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.10,
    fill: { color: C.WHITE },
    line: { color: tensColor, width: 1.5 },
  });

  // Header
  slide.addText(String(num), {
    x, y: y + 0.05, w, h: 0.50,
    fontSize: 28, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Tens chip
  const chipY = y + 0.65;
  const chipW = (w - 0.40) / 2;
  slide.addShape("roundRect", {
    x: x + 0.15, y: chipY, w: chipW, h: 0.50, rectRadius: 0.08,
    fill: { color: tensColor },
  });
  slide.addText(String(tens * 10), {
    x: x + 0.15, y: chipY, w: chipW, h: 0.50,
    fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  slide.addText(tens + " tens", {
    x: x + 0.15, y: chipY + 0.55, w: chipW, h: 0.30,
    fontSize: 12, fontFace: FONT_B, color: C.MUTED,
    align: "center", valign: "middle", margin: 0,
  });

  // Ones chip
  slide.addShape("roundRect", {
    x: x + 0.25 + chipW, y: chipY, w: chipW, h: 0.50, rectRadius: 0.08,
    fill: { color: onesColor },
  });
  slide.addText(String(ones), {
    x: x + 0.25 + chipW, y: chipY, w: chipW, h: 0.50,
    fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  slide.addText(ones + " ones", {
    x: x + 0.25 + chipW, y: chipY + 0.55, w: chipW, h: 0.30,
    fontSize: 12, fontFace: FONT_B, color: C.MUTED,
    align: "center", valign: "middle", margin: 0,
  });
}

// Draw a 3-cell place value chart with header H/T/O.
function drawPvChart(slide, x, y, w, h, hundreds, tens, ones) {
  const cellW = w / 3;
  const hdrH = h * 0.40;
  const valH = h * 0.60;

  // Headers
  ["H", "T", "O"].forEach((label, i) => {
    slide.addShape("rect", {
      x: x + i * cellW, y, w: cellW, h: hdrH,
      fill: { color: C.PRIMARY },
      line: { color: C.WHITE, width: 1 },
    });
    slide.addText(label, {
      x: x + i * cellW, y, w: cellW, h: hdrH,
      fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  // Values
  const values = [hundreds, tens, ones];
  values.forEach((v, i) => {
    slide.addShape("rect", {
      x: x + i * cellW, y: y + hdrH, w: cellW, h: valH,
      fill: { color: C.WHITE },
      line: { color: C.PRIMARY, width: 1 },
    });
    if (v != null) {
      slide.addText(String(v), {
        x: x + i * cellW, y: y + hdrH, w: cellW, h: valH,
        fontSize: 30, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome team. Today is the last lesson of our subtraction unit.
- We are going to use everything we know to solve a real problem about money.
- Today's strategy: partition the tens and ones.

DO:
- Have whiteboards, markers and play money or counters ready.

TEACHER NOTES:
Lesson 5 of 5. Capstone lesson. Students apply partitioning to a two-digit subtraction and to a money problem from elaboration VC2M2N04. Bridges to multi-digit subtraction in Year 3.

WATCH FOR:
- Students who default to counting back - acknowledge the strategy and name when partitioning is faster.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the two numbers on the slide.
- One is 50. One is 5.
- Whisper: which is bigger? What does the zero in 50 do?

DO:
- Display the two numbers in place value charts.
- Allow 30 seconds for partner whisper.

TEACHER NOTES:
Daily Review revisits the importance of zero. The zero in 50 holds the ones place so we know there are 5 tens and zero ones.

WATCH FOR:
- Students who say "the 0 means nothing" - prompt them: it means zero ones, and it makes 50 ten times bigger than 5.
- Students who think 50 and 5 are the same value.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- 50 is bigger.
- 50 has 5 tens and 0 ones. The zero holds the ones place.
- Without the zero, 50 would be just 5.

DO:
- Click to reveal.
- Tick on whiteboards.

TEACHER NOTES:
Tick and fix. The zero is a placeholder, not nothing.

WATCH FOR:
- Students who self-correct on the reveal.
- Students who still think 50 = 5 - small group focus.

[Stage 1: Daily Review Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Place value warm up.
- I will say a number. Tell me 10 more. Then 10 less.
- 34. 10 more? 10 less?
- 60. 10 more? 10 less?

DO:
- Brisk choral response. 4 to 5 numbers.
- Then quick partition: "Show me 47 as tens and ones."

TEACHER NOTES:
Today's strategy needs partitioning fluency. Drill it briefly here so the I Do feels familiar.

WATCH FOR:
- Students who flip 10 more and 10 less - quick correction.
- Students who can't partition without blocks - prompt with the place value chart.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to subtract two-digit numbers by partitioning the tens and the ones.
- Read the success criteria together.

DO:
- Choral read the LI.
- Choral read each I can statement.

TEACHER NOTES:
The first I can is achievable for everyone (just splitting into tens and ones). The second is the core target. The third stretches students into a real money problem.

WATCH FOR:
- Students who say tens and ones fluently - secure.
- Students unsure - the I Do uses base 10 blocks.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. The problem is 45 take away 23.
- I will partition. 45 is 4 tens and 5 ones. 23 is 2 tens and 3 ones.
- I take the tens first. 4 tens take 2 tens is 2 tens.
- Then I take the ones. 5 ones take 3 ones is 2 ones.
- 2 tens and 2 ones is 22.
- So 45 take away 23 is 22.

DO:
- Use base 10 blocks on the document camera.
- Show 4 rods and 5 cubes for 45.
- Take 2 rods (saying "take 2 tens") and 3 cubes (saying "take 3 ones").
- Touch the remaining 2 rods and 2 cubes as you say the answer.

TEACHER NOTES:
First model. The two operations (tens take tens, ones take ones) are independent. Note: this works because there is no regrouping today. Avoid 45 - 27 type problems for now.

MISCONCEPTIONS:
- Misconception: Students take ones from tens (e.g. 4 tens take 3 ones).
  Why: They line up the digits visually.
  Impact: They get a wrong answer.
  Quick correction: "Tens take tens. Ones take ones. Stay in your place."

WATCH FOR:
- Students who whisper the steps - tracking.
- Students who get confused with the order - reteach with blocks.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. The problem is 36 take away 14.
- First, partition 14. How many tens? How many ones?

DO:
- Display 36 - 14 with the partition card on screen.
- Allow 5 seconds.
- Say: "Show me on your fingers. Tens first."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "How many tens in 14? Show me."
- Scan for: 1 finger held up.
PROCEED: If 80% show 1 finger, click to reveal and continue.
PIVOT: Most likely misconception - students count digits not values.
- Reteach: "14 is 1 ten and 4 ones."
- Re-check: "How many tens?"

TEACHER NOTES:
Just the partition step for now. We Do completes the subtraction.

WATCH FOR:
- Students who hold up 1 - secure.
- Students who hold up 14 - they read the digits; reteach.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. The problem is 56 take away 23.
- Partition: 56 is 5 tens and 6 ones. 23 is 2 tens and 3 ones.
- Take the tens. Then take the ones.
- Whisper your answer to your partner. Then write it.

DO:
- Display the partition cards: 56 above 23.
- Allow 30 seconds.
- Walk and scan.

TEACHER NOTES:
We Do uses different numbers from I Do but the same strategy. Listen for "5 tens take 2 tens is 3 tens. 6 ones take 3 ones is 3 ones."

WATCH FOR:
- Pairs who calculate tens then ones - secure.
- Pairs who calculate digit by digit - prompt them: "Stay in tens. Then ones."

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 5 tens take 2 tens is 3 tens.
- 6 ones take 3 ones is 3 ones.
- 3 tens and 3 ones is 33.
- So 56 take away 23 is 33.

DO:
- Click to reveal.
- Repeat the language pattern with the class.

TEACHER NOTES:
Highlight that we did two simple subtractions, then added the parts.

WATCH FOR:
- Students who said 33 first - secure.
- Students who said 32 or 34 - they slipped on one of the partial subtractions.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2_Q = `SAY:
- Hinge question. Look at this work.
- The problem was 47 take away 25.
- I said 4 tens take 2 tens is 2 tens.
- I said 7 ones take 5 ones is 2 ones.
- I said the answer is 22.
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Display the worked example.
- Wait 5 seconds.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "47 take away 25 is 22. Thumbs up if yes."
- Scan for: thumbs UP. The reasoning is correct.
PROCEED: If 80% show thumbs up, click to reveal and confirm.
PIVOT: Most likely misconception - students think the answer must be bigger.
- Reteach: "We took 2 tens and 5 ones. So we took 25. The answer is the rest."
- Re-check: "What is 47 take 25?"

TEACHER NOTES:
This hinge probes whether students believe the partition method.

WATCH FOR:
- Confident thumbs up - they trust the strategy.
- Slow thumbs - they are unsure; reteach with blocks.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn. A real money problem.
- Eli has $48. He buys a book for $25. How much does Eli have left?
- First: write the take away. Second: partition. Third: solve.

DO:
- Display the word problem with a coin or note picture.
- Allow 2-3 minutes.
- Circulate. Listen for "I split the tens and ones".
- Cold call 1-2 students.

TEACHER NOTES:
Mathematical modelling (VC2M2N06). Students choose subtraction because they know the total and one part. The numbers are designed so partitioning works without regrouping.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use $35 - $14 instead.
- Extra Notes: Sit with these students; play money available.
EXTENDING PROMPT:
- Task: After solving, what if Eli also bought a $12 pencil case? How much would he have left?
- Extra Notes: Two-step problem; encourage them to write the second take away.

WATCH FOR:
- Students who write 48 - 25 = 23 - secure.
- Students who write 48 + 25 - they chose the wrong operation; remind them: total and a part means take away.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Eve has 64 stickers. She gives 32 away. How many does Eve have left?
- Partition the tens and ones. Show your thinking.

DO:
- Display the prompt.
- Allow 90 seconds.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses the core target - using partitioning for two-digit subtraction. The answer is 32. Splits: 60-30=30, 4-2=2, total 32.

WATCH FOR:
- Students who write 32 with reasoning - on track.
- Students who add by mistake - reteach the part-whole story.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today's strategy: partition the tens and ones.
- This is the last lesson of our unit. We have five strategies now: counting back, think addition, doubles, bridging through 10, and partitioning.
- Show me thumbs: which strategy do you like best?
- Turn and tell your partner: when do you choose partitioning?

DO:
- Read the success criteria with the class.
- Use thumbs up sideways down.
- Cold call 2 students for the partner share.

TEACHER NOTES:
Capstone reflection. Encourage students to choose a strategy that fits the numbers. Keep the unit's strategies visible on the wall.

WATCH FOR:
- Students who name a favourite strategy - secure with the unit.
- Students who only know one strategy - small group revisit.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- These are the materials for today.
- The money problem sheet is for table time.

DO:
- Print the money problem practice sheet (one per student) only if you want a written record.
- Have whiteboards, markers, base 10 blocks, and play money ready.

TEACHER NOTES:
One printed student resource for this lesson.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 5: Two-Digit Take Away with Money",
    `Year 2 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slides 2-3: Daily Review with reveal — importance of zero
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Which is bigger: 50 or 5?", { color: C.ACCENT });

      // Left card: instruction
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Which is bigger?", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "What does the 0 do?", options: { fontSize: 22, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right card: two place value charts
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      drawPvChart(s, 5.55, CONTENT_TOP + 0.45, 3.60, 1.10, null, 5, 0);
      s.addText("50", {
        x: 5.55, y: CONTENT_TOP + 0.10, w: 3.60, h: 0.30,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      drawPvChart(s, 5.55, CONTENT_TOP + 2.05, 3.60, 1.10, null, null, 5);
      s.addText("5", {
        x: 5.55, y: CONTENT_TOP + 1.70, w: 3.60, h: 0.30,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "50 is bigger.   The 0 holds the ones place.   50 = 5 tens, 0 ones.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — place value warm up (10 more / 10 less, partition)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.ACCENT);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Place value warm up", { color: C.ACCENT });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

    // Two prompt cards side by side
    addCard(s, 0.85, CONTENT_TOP + 0.30, 4.10, 2.50, { strip: C.PRIMARY });
    s.addText("34", {
      x: 0.85, y: CONTENT_TOP + 0.55, w: 4.10, h: 1.10,
      fontSize: 84, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("10 more?      10 less?", {
      x: 0.85, y: CONTENT_TOP + 1.85, w: 4.10, h: 0.50,
      fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addCard(s, 5.05, CONTENT_TOP + 0.30, 4.10, 2.50, { strip: C.SECONDARY });
    s.addText("47", {
      x: 5.05, y: CONTENT_TOP + 0.55, w: 4.10, h: 1.10,
      fontSize: 84, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("Tens and ones?", {
      x: 5.05, y: CONTENT_TOP + 1.85, w: 4.10, h: 0.50,
      fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("Whisper to your partner. Then write on your whiteboard.", {
      x: 0.5, y: CONTENT_TOP + 3.05, w: 9.0, h: 0.40,
      fontSize: 18, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FLUENCY);
  })();

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to subtract two-digit numbers by partitioning the tens and ones.",
    [
      "I can split a two-digit number into tens and ones.",
      "I can take the tens then take the ones to find the answer.",
      "I can solve a money take away problem using partitioning.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — 45 - 23 = 22
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "45 take away 23", { color: STAGE_COLORS["2"] });

    // Top fact strip
    addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["2"] });
    s.addText("45 - 23 = ?", {
      x: 0.5, y: CONTENT_TOP + 0.10, w: 9.0, h: 0.65,
      fontSize: 28, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Two partition cards side by side: 45 and 23
    drawPartitionCard(s, 0.85, CONTENT_TOP + 1.00, 3.20, 1.45, 45);
    s.addText("take", {
      x: 4.10, y: CONTENT_TOP + 1.55, w: 1.10, h: 0.40,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    drawPartitionCard(s, 5.25, CONTENT_TOP + 1.00, 3.20, 1.45, 23);

    // Bottom answer strip — single string keeps the XML simple
    addCard(s, 0.5, CONTENT_TOP + 2.65, 9.0, 1.10, { strip: C.SECONDARY });
    s.addText("Tens:  40 - 20 = 20.   Ones:  5 - 3 = 2.   Total:  20 + 2 = 22.", {
      x: 0.75, y: CONTENT_TOP + 2.78, w: 8.5, h: 0.85,
      fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
  })();

  // Slides 7-8: CFU 1 with reveal — partition 14
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Partition 14", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Big numeral
      s.addText("14", {
        x: 0.5, y: CONTENT_TOP + 0.30, w: 9.0, h: 1.40,
        fontSize: 110, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Question
      s.addText("How many TENS in 14?   Show me on your fingers.", {
        x: 0.5, y: CONTENT_TOP + 2.10, w: 9.0, h: 0.55,
        fontSize: 24, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "14 = 1 ten and 4 ones.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 9-10: We Do with reveal — 56 - 23
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "56 take away 23", { color: STAGE_COLORS["3"] });

      // Top reasoning strip
      addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["3"] });
      s.addText("56 - 23 = ?", {
        x: 0.5, y: CONTENT_TOP + 0.10, w: 9.0, h: 0.65,
        fontSize: 26, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Two partition cards
      drawPartitionCard(s, 0.85, CONTENT_TOP + 1.00, 3.20, 1.45, 56);
      s.addText("take", {
        x: 4.10, y: CONTENT_TOP + 1.55, w: 1.10, h: 0.40,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      drawPartitionCard(s, 5.25, CONTENT_TOP + 1.00, 3.20, 1.45, 23);

      // On whiteboard prompt
      addCard(s, 0.5, CONTENT_TOP + 2.65, 9.0, 1.10, { strip: C.SECONDARY });
      s.addText("Tens: ____ - ____ = ____    Ones: ____ - ____ = ____    Total = ____", {
        x: 0.75, y: CONTENT_TOP + 2.85, w: 8.5, h: 0.65,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Tens: 50 - 20 = 30.   Ones: 6 - 3 = 3.   Total = 33.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 11-12: CFU Hinge with reveal — Is 47 - 25 = 22 right?
  withReveal(
    () => cfuSlide(pres, "CFU", "Is the answer 22?", "Thumbs Up or Thumbs Down",
      "47 take away 25.\n\n4 tens take 2 tens = 2 tens.\n7 ones take 5 ones = 2 ones.\n\nAnswer: 22.\n\nIs that right?",
      NOTES_CFU2_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs UP.   20 + 2 = 22.   The partition method works.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: You Do — money problem
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Eli's money", { color: STAGE_COLORS["4"] });

    // Top story card — single string
    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.20, { strip: STAGE_COLORS["4"] });
    s.addText("Eli has $48.   He buys a book for $25.   How much does Eli have left?", {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.85,
      fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Steps panel — three discrete bullets keep the XML simple
    const panelY = CONTENT_TOP + 1.35;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText([
      { text: "Step 1:  Write the take away.   $48 - $25 = ____", options: { fontSize: 20, color: C.CHARCOAL, breakLine: true } },
      { text: "Step 2:  Partition the tens and ones.", options: { fontSize: 20, color: C.CHARCOAL, breakLine: true } },
      { text: "Step 3:  Take the tens. Take the ones. Find the total.", options: { fontSize: 20, color: C.CHARCOAL } },
    ], {
      x: 0.85, y: panelY + 0.20, w: 8.30, h: 1.85,
      fontFace: FONT_B, valign: "top", margin: 0,
      paraSpaceAfter: 8,
    });

    s.addText("On your whiteboard. Show your thinking.", {
      x: 0.5, y: panelY + 2.05, w: 9.0, h: 0.36,
      fontSize: 16, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 14: Exit Ticket
  exitTicketSlide(pres,
    [
      "Eve has 64 stickers. She gives 32 away. How many does Eve have left? Show your tens-and-ones thinking.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 15: Closing — capstone, name the 5 strategies
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: which subtraction strategy do you like best?",
      scItems: [
        "I can split a two-digit number into tens and ones.",
        "I can take the tens then take the ones to find the answer.",
        "I can solve a money take away problem using partitioning.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 16: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Subs_Lesson5_Two_Digit_Money.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Money problem practice sheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Solve by partitioning the tens and the ones.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addTipBox(doc,
      "Tens take tens. Ones take ones. Add the parts back together.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Practice", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  47 - 24 = ____      Tens: ____   Ones: ____   Total: ____", y);
    y = addBodyText(doc, "b)  68 - 35 = ____      Tens: ____   Ones: ____   Total: ____", y);
    y = addSectionHeading(doc, "Money problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "c)  Maya has $59. She spends $34 on a board game. How much does she have left?", y);
    y += 8;
    y = addBodyText(doc, "    My take away:  ____ - ____ = ____", y);
    y = addBodyText(doc, "    Tens:  ____ - ____ = ____", y);
    y = addBodyText(doc, "    Ones:  ____ - ____ = ____", y);
    y = addBodyText(doc, "    Total left:  $____", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Money Problem Practice | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the money problem practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addBodyText(doc, "Each answer partitions the tens and ones, subtracts each place separately, then adds the parts.", y);
    y = addSectionHeading(doc, "Answers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  47 - 24:  Tens 40 - 20 = 20.   Ones 7 - 4 = 3.   Total = 23.", y);
    y = addBodyText(doc, "b)  68 - 35:  Tens 60 - 30 = 30.   Ones 8 - 5 = 3.   Total = 33.", y);
    y = addBodyText(doc, "c)  $59 - $34:  Tens $50 - $30 = $20.   Ones $9 - $4 = $5.   Total = $25.   Maya has $25 left.", y);
    y = addTipBox(doc,
      "If a student gets a wrong total, ask them to show the partition. Most errors come from mixing up the place value.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 5 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
