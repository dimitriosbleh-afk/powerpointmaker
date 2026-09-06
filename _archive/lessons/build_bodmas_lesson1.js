"use strict";

// BODMAS Unit (Year 5/6 Numeracy) - Lesson 1: Why we need an order of operations.
// First time looking at this content this year, so introduce the idea carefully.
// Daily Review: 3D shapes (faces, edges, vertices).
// Number Fluency: Adding 2-digit numbers.
// Notes: lots of manipulatives. Students are weaker on additive thinking,
// stronger on subtraction - so additive expressions use friendly numbers.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Theme: Year 5/6 numeracy. Variant fixed across all 5 lessons for cohesion.
const T = createTheme("numeracy", "grade56", weekToVariant(7));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 5;
const UNIT_TITLE = "BODMAS: Order of Operations";
const FOOTER = `BODMAS | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/BODMAS_Lesson1_Why_We_Need_Rules";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PRACTICE_RES = makeSessionResource(SESSION,
  "Lesson 1 BODMAS Reference Card",
  "A small reference card students cut and stick into their maths books.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 1 Answer Key",
  "Teacher reference with the worked answers for today's practice questions.");
const RESOURCE_ITEMS = [PRACTICE_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes ----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to a brand new unit, BODMAS.
- BODMAS is the rule mathematicians around the world use when a number sentence has more than one step.
- This is the first time we are looking at it this year. If it feels new, that is okay. We will build it together.

DO:
- Settle students before clicking on.
- Have whiteboards and counters out before you start.

TEACHER NOTES:
Lesson 1 of 5. Today is about WHY mathematicians need an agreed order. Save the full acronym work for the rule slide later in the lesson.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are our materials for today.
- You will use your whiteboard, counters and the small BODMAS reference card.

DO:
- Have a pot of 20 counters per pair on each desk.
- Print one BODMAS reference card per student. Cut along the dotted line, then glue into maths books at the end of the lesson.
- Print the answer key for yourself.

TEACHER NOTES:
Only one small student resource today. The lesson runs on counters, whiteboards and partner talk.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Today we are revising 3D shapes.
- Look at the shape. On your whiteboard, write its name and how many faces it has.

DO:
- Show the prompt.
- 60 seconds.
- Walk and scan.

TEACHER NOTES:
3D shapes prior learning. The cube has 6 faces, 12 edges, 8 vertices. Watch for the classic mix-up between vertices and edges.

WATCH FOR:
- Students who confuse face with side - quick correction with "a face is a whole flat part".

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check yours.
- This shape is a cube. It has 6 faces, 12 edges and 8 vertices.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Note any student who still writes 4 or 8 faces for follow-up small group.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Adding two-digit numbers.
- Whisper the answer to your partner, then write it on your board.

DO:
- Show the three prompts.
- 45 seconds.
- Walk and scan.

TEACHER NOTES:
The cohort is weaker on additive thinking. Numbers chosen are friendly. Encourage strategies like "split into tens and ones" or "add the tens first".

WATCH FOR:
- Students who count on slowly - prompt with "use the tens first".
- Students who answer instantly - secure, push to a harder strategy in small group.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 23 + 14 = 37.
- 35 + 22 = 57.
- 46 + 18 = 64.
- Tick and fix.

DO:
- Click to reveal.
- Note any student still finger-counting.

TEACHER NOTES:
Look for students who add ones first and forget to regroup on 46 + 18. Quick small group support tomorrow.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Look at the number sentence on the board: 5 + 3 x 2.
- Tia worked it out and got 16.
- Sam worked it out and got 11.
- Same numbers. Same signs. Two different answers. How can that be?
- Talk to your partner for 30 seconds about how each one got their answer.

DO:
- Display the disagreement.
- 30 seconds partner talk.
- Cold call 2 students to share what they noticed.

TEACHER NOTES:
This is the hook. Students do not need to know the rule yet. The goal is to feel the problem - if everyone chooses their own order, we all get different answers. That is why mathematicians needed a shared rule.

WATCH FOR:
- Students who spot that Tia did the times first and Sam did the plus first - tracking.
- Students who think one of them is wrong because of bad maths - prompt with "they both did real maths, just in a different order".

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning that we need an agreed order for working out a number sentence with more than one step.
- Now the success criteria. Read each one with me.

DO:
- Choral read each statement.
- Point to the LI as you read.

TEACHER NOTES:
SC1 is recall of WHY the rule exists. Almost every student can reach this today. SC2 is the core target - applying the rule to a simple "multiply first, then add" expression. SC3 is the stretch - explaining what would happen without the rule.

WATCH FOR:
- Students who can already point to the answer from the launch - secure.
- Students who still look unsure - that is fine, we will build it together.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- A quick look at three words we will use today.
- An OPERATION is a maths action. Plus, minus, times and divide are all operations.
- ORDER means which one we do first.
- A RULE is something everyone agrees to follow.
- Whisper to your partner: what is one example of a rule we follow at school?

DO:
- Point to each word.
- 30 seconds partner talk for the school rule example.

TEACHER NOTES:
Three small words land the concept of "agreed order". The school rule analogy bridges into mathematicians agreeing on BODMAS.

WATCH FOR:
- Students who give a clear example like "left side of the path" - secure.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_IDO1 = `SAY:
- Watch me. Same number sentence: 5 + 3 x 2.
- Tia did the plus first. 5 + 3 is 8. Then 8 x 2 is 16.
- Sam did the times first. 3 x 2 is 6. Then 5 + 6 is 11.
- Both did real maths. They just chose a different order.
- The trouble is, mathematicians need ONE answer everyone agrees on.

DO:
- Build both paths on the board as you talk.
- Use two different colours: one for each path.
- Underline the operation done first in each.

TEACHER NOTES:
Modelling the two paths makes the WHY visible. Do not introduce the BODMAS rule yet. The point of this slide is the problem, not the solution.

MISCONCEPTIONS:
- Misconception: Students think one of the two answers is wrong because of mis-calculation.
  Why: They are used to right-or-wrong arithmetic.
  Impact: They miss the bigger idea about needing a shared order.
  Quick correction: "Both answers are real maths. The problem is that we cannot have two answers to the same sum."

WATCH FOR:
- Students who track both paths with you - secure for the rule reveal.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Mathematicians around the world agreed on one order. We call it BODMAS.
- Read each letter with me.
- B is brackets. Anything inside brackets is done first.
- O is orders, things like powers. We will not focus on this one in Year 5 and 6.
- D and M are division and multiplication. We do these next.
- A and S are addition and subtraction. We do these last.
- That means for 5 + 3 x 2, the times comes BEFORE the plus. The answer is 11.
- Sam was right. Tia did real maths, but in the wrong order.

DO:
- Point to each letter as you read.
- Circle the M and the times sign on the board.
- Write the answer 11.

TEACHER NOTES:
This is the rule reveal. Keep brackets and orders simple - we will teach brackets properly in Lesson 3. For Year 5/6, orders is rare in the curriculum and can be safely parked.

WATCH FOR:
- Students who repeat "times before plus" - tracking.
- Students who still look confused - reassure: "We will practise this together."

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Look at the sentence 4 + 2 x 3.
- Which operation do we do FIRST?
- On your whiteboard, write the operation: plus, or times.
- Then write the answer.

DO:
- Show the prompt.
- 45 seconds.
- Walk and scan boards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your boards on three, two, one, show."
- Scan for: "times" and answer "10".
PROCEED: If 80% write times and 10, click to reveal and move on.
PIVOT: Most likely misconception - students do the plus first because it appears first.
- Reteach: Walk back to the BODMAS letters. "Which letter comes earlier - M or A?"
- Re-check: "Try 6 + 5 x 2 on your board now. Which one first?"

TEACHER NOTES:
This single CFU decides whether we go to the We Do or pivot back to the rule. Walk and scan boards before you click reveal.

WATCH FOR:
- Boards showing 10 with "times first" - secure.
- Boards showing 18 - they did plus first, pivot back to the rule.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- Check yours.
- Times comes before plus, so 2 x 3 is 6. Then 4 + 6 is 10.
- Tick and fix.

DO:
- Click to reveal.
- Allow 15 seconds for self-correction.

TEACHER NOTES:
If more than a few students wrote 18, reteach using the counters before moving to the We Do.

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me.
- 6 + 4 x 2.
- Step one - which operation comes first?
- Tell your partner.
- Step two - work out the times.
- Step three - finish with the plus.
- On your board.

DO:
- Display the expression.
- Partner talk for 20 seconds.
- Build it with counters at the front: 4 groups of 2, then add 6 more.
- Cold call one pair to share Step 1.
- Reveal on the next click.

TEACHER NOTES:
Use the counter demo to make the times visible BEFORE you write the answer. This is the additive thinking scaffold the cohort needs - they should see 8 counters, then add 6 more to get 14.

WATCH FOR:
- Pairs who say "times first" - secure.
- Pairs still adding first - prompt with the BODMAS card on their desk.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together.
- 4 x 2 is 8.
- 6 + 8 is 14.
- The answer is 14.
- Tick or fix.

DO:
- Click to reveal.
- Build 4 groups of 2 with counters at the front, then add 6 more counters to show 14.

TEACHER NOTES:
The counter build doubles as the additive scaffold. After 4 x 2 they see 8 counters - the addition is now a friendly "6 + 8" rather than abstract.

[Stage 3: We Do Answer | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- One number sentence at a time on your whiteboard.
- For each one, first underline the operation that comes first, then work it out step by step.
- I will check in after each one.

DO:
- Display one prompt at a time.
- 1 minute per prompt.
- Build with counters at the front for any student who needs the visual.
- Cold call after each one.

TEACHER NOTES:
Three problems. All use the "times then plus" or "times then minus" pattern - no brackets yet, no division yet. The additive scaffold is built in via friendly numbers.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work alongside the teacher with counters for each problem. Build the times group, then add or subtract.
- Extra Notes: Sit at the front. Build it before you write it.
EXTENDING PROMPT:
- Task: After all three, write a number sentence of your own where the times comes second. Show what the answer would be if you broke BODMAS, and what it is if you follow BODMAS.
- Extra Notes: Encourage students to use friendly numbers.

WATCH FOR:
- Students who underline the times sign first - secure.
- Students who underline the plus - prompt back to the BODMAS card.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Work out 7 + 3 x 4. Show me which operation you did first by underlining it.
- One short sentence: why does it matter which operation we do first?

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards or photograph.

TEACHER NOTES:
Exit ticket assesses SC2 - applying the rule. The "why" prompt touches SC1 (rule exists) and SC3 (explain). Expected answer: 19, because times comes before plus.

WATCH FOR:
- Students who write 19 and explain the rule - secure.
- Students who write 40 - they added first. Plan a small group recap tomorrow.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria again.
- Thumbs up, sideways, or down for each one.
- Turn and tell your partner: what is BODMAS, in your own words?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call 1-2 students.

TEACHER NOTES:
Threshold idea is "times and divide before plus and minus". Students who can say this are ready for Lesson 2.

WATCH FOR:
- Strong thumbs up across the board - move at pace tomorrow.
- Sideways or down on SC2 - revisit with counters in the small group at the start of Lesson 2.

[General: Closing | VTLM 2.0: Reflection]`;

// --- Helpers ---------------------------------------------------------------

// Draw a simple cube net or cube shape with face/edge/vertex highlights
function drawCube(slide, x, y, size) {
  const offset = size * 0.30;
  const front = {
    x, y: y + offset, w: size, h: size,
  };
  // Back face polygon points (rough isometric)
  slide.addShape("rect", {
    x: x + offset, y, w: size, h: size,
    fill: { color: C.BG_LIGHT || "F6F6F6" },
    line: { color: C.CHARCOAL, width: 1.5 },
  });
  // Connection lines
  slide.addShape("line", {
    x: x, y: y + offset, w: offset, h: -offset,
    line: { color: C.CHARCOAL, width: 1.2 },
  });
  slide.addShape("line", {
    x: x + size, y: y + offset, w: offset, h: -offset,
    line: { color: C.CHARCOAL, width: 1.2 },
  });
  slide.addShape("line", {
    x: x + size, y: y + size + offset, w: offset, h: -offset,
    line: { color: C.CHARCOAL, width: 1.2 },
  });
  slide.addShape("line", {
    x: x, y: y + size + offset, w: offset, h: -offset,
    line: { color: C.CHARCOAL, width: 1.2 },
  });
  // Front face on top
  slide.addShape("rect", {
    ...front,
    fill: { color: C.PRIMARY },
    line: { color: C.CHARCOAL, width: 1.5 },
  });
}

// Draw the BODMAS acronym as a row of colour-coded letter cards with meaning underneath.
function drawBodmasRow(slide, x, y, w, h) {
  const letters = [
    { letter: "B", meaning: "Brackets",      color: C.PRIMARY,   active: true },
    { letter: "O", meaning: "Orders",        color: C.SECONDARY, active: false },
    { letter: "D", meaning: "Division",      color: C.ACCENT,    active: true },
    { letter: "M", meaning: "Multiplication", color: C.ACCENT,    active: true },
    { letter: "A", meaning: "Addition",      color: C.ALERT,     active: true },
    { letter: "S", meaning: "Subtraction",   color: C.ALERT,     active: true },
  ];
  const gap = 0.10;
  const cellW = (w - gap * (letters.length - 1)) / letters.length;
  letters.forEach((it, i) => {
    const cx = x + i * (cellW + gap);
    addTextOnShape(slide, it.letter, {
      x: cx, y, w: cellW, h: h * 0.55, rectRadius: 0.10,
      fill: { color: it.color },
    }, { fontSize: 32, fontFace: FONT_H, color: C.WHITE, bold: true });

    slide.addText(it.meaning, {
      x: cx, y: y + h * 0.58, w: cellW, h: h * 0.42,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL,
      bold: true, align: "center", valign: "top", margin: 0,
    });
  });
}

// Draw a single large expression with two highlighted parts (used in launch hook)
function drawExpression(slide, x, y, w, h, exprText, opts) {
  const o = opts || {};
  addTextOnShape(slide, exprText, {
    x, y, w, h, rectRadius: 0.12,
    fill: { color: o.fill || C.BG_LIGHT || "F6F6F6" },
    line: { color: o.line || C.CHARCOAL, width: 1.4 },
  }, {
    fontSize: o.fontSize || 56, fontFace: FONT_H,
    color: o.color || C.CHARCOAL, bold: true,
  });
}

// Draw a row of counters (as small dots). Used to scaffold additive thinking.
function drawCounters(slide, x, y, count, opts) {
  const o = opts || {};
  const dotSize = o.dotSize || 0.22;
  const gap = o.gap || 0.06;
  const color = o.color || C.PRIMARY;
  for (let i = 0; i < count; i += 1) {
    const cx = x + i * (dotSize + gap);
    slide.addShape("roundRect", {
      x: cx, y, w: dotSize, h: dotSize,
      rectRadius: dotSize / 2,
      fill: { color },
      line: { color, width: 0.2 },
    });
  }
}

// --- Build ------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Why we need rules",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - 3D shapes (cube)
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Name this 3D shape", [
      "Name this shape.",
      "How many faces?",
    ], NOTES_DR_Q, FOOTER, (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.ACCENT });
      slide.addText("This shape", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.34,
        fontSize: 16, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      drawCube(slide, lg.rightX + 0.95, lg.panelTopPadded + 0.65, 1.85);
    }),
    (slide) => {
      addTextOnShape(slide, "Cube  |  6 faces, 12 edges, 8 vertices", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    },
  );

  // Slides 5-6: Fluency + reveal - 2-digit addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Adding 2-digit numbers",
      ["23 + 14", "35 + 22", "46 + 18"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "37      57      64", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    },
  );

  // Slide 7: Launch - same numbers, two answers
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Launch", { color: C.PRIMARY });
    addTitle(s, "Two students. Same sum. Two answers.");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    // Hero expression
    drawExpression(s, 2.0, CONTENT_TOP + 0.25, 6.0, 0.85, "5  +  3  x  2");

    // Two student boxes side by side
    const studentY = CONTENT_TOP + 1.30;
    const studentH = 1.55;
    const colW = 4.10;

    addTextOnShape(s, "Tia: 16", {
      x: 0.85, y: studentY, w: colW, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    s.addText("She did the + first.\n5 + 3 = 8, then 8 x 2 = 16.", {
      x: 0.85, y: studentY + 0.52, w: colW, h: studentH - 0.55,
      fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, "Sam: 11", {
      x: 5.05, y: studentY, w: colW, h: 0.45, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    s.addText("He did the x first.\n3 x 2 = 6, then 5 + 6 = 11.", {
      x: 5.05, y: studentY + 0.52, w: colW, h: studentH - 0.55,
      fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("Turn and tell: how did each one work it out?", {
      x: 0.7, y: SAFE_BOTTOM - 0.45, w: 8.6, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning that we need an agreed order when a number sentence has more than one step.",
    [
      "I can say WHY we need an order of operations.",
      "I can work out a number sentence with x and + using the right order.",
      "I can explain what would happen if everyone chose their own order.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key vocabulary (after LI/SC per §0a)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Three words for today");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const wordY = CONTENT_TOP + 0.35;
    const cellW = 2.85;
    const gap = 0.20;
    const startX = (10 - (cellW * 3 + gap * 2)) / 2;

    const words = [
      { word: "operation", def: "a maths action like\n+  -  x  ÷" },
      { word: "order",     def: "which one we do\nfirst, then next" },
      { word: "rule",      def: "something everyone\nagrees to follow" },
    ];

    words.forEach((it, i) => {
      const cx = startX + i * (cellW + gap);
      addTextOnShape(s, it.word, {
        x: cx, y: wordY, w: cellW, h: 0.55, rectRadius: 0.10,
        fill: { color: C.PRIMARY },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(it.def, {
        x: cx, y: wordY + 0.65, w: cellW, h: 1.40,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0,
      });
    });

    s.addText("Whisper to your partner: what is one school rule we follow?", {
      x: 0.7, y: SAFE_BOTTOM - 0.50, w: 8.6, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 10: I Do (1) - show two paths
  workedExSlide(pres, 2, "I Do", "Two paths from the same sum",
    [
      "5 + 3 x 2",
      "",
      "Tia: + first",
      "5 + 3 = 8",
      "8 x 2 = 16",
      "",
      "Sam: x first",
      "3 x 2 = 6",
      "5 + 6 = 11",
      "",
      "Same numbers. Two answers.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("Why this matters", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.32,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      addTextOnShape(slide, "Mathematicians needed ONE shared order.", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.40, h: 0.95,
        rectRadius: 0.10,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText(
        "Without a shared order, two people would always disagree on simple sums.",
        {
          x: lg.rightX + 0.20, y: lg.panelTopPadded + 1.70, w: lg.rightW - 0.40, h: 1.20,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
          italic: true, align: "center", valign: "top", margin: 0,
        },
      );
    });

  // Slide 11: I Do (2) - the BODMAS rule
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "The rule mathematicians agreed on", { color: STAGE_COLORS["2"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["2"] });

    drawBodmasRow(s, 0.85, CONTENT_TOP + 0.30, 8.3, 1.55);

    // Worked sum below
    s.addText("So for 5 + 3 x 2, the x comes BEFORE the +.", {
      x: 0.7, y: CONTENT_TOP + 2.10, w: 8.6, h: 0.36,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
      bold: true, align: "center", margin: 0,
    });

    addTextOnShape(s, "5 + 3 x 2  =  5 + 6  =  11", {
      x: 1.5, y: CONTENT_TOP + 2.60, w: 7.0, h: 0.55, rectRadius: 0.10,
      fill: { color: C.SUCCESS },
    }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO2);
  })();

  // Slides 12-13: CFU + reveal - which one first?
  withReveal(
    () => cfuSlide(pres, "CFU", "Which operation first?", "Show Me Boards",
      "4 + 2 x 3\n\nOn your whiteboard:\n\n1.  Write which operation comes FIRST.\n2.  Write the answer.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "x first.  2 x 3 = 6.  Then 4 + 6 = 10.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    },
  );

  // Slides 14-15: We Do + reveal - 6 + 4 x 2 with counter visual
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Together: 6 + 4 x 2", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      // Big expression
      drawExpression(s, 2.0, CONTENT_TOP + 0.25, 6.0, 0.80, "6  +  4  x  2");

      // Step prompts
      s.addText("Step 1.  Which one first?", {
        x: 0.85, y: CONTENT_TOP + 1.30, w: 4.10, h: 0.36,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        bold: true, align: "left", margin: 0,
      });
      addTextOnShape(s, "x", {
        x: 4.10, y: CONTENT_TOP + 1.25, w: 0.60, h: 0.50, rectRadius: 0.10,
        fill: { color: C.ACCENT },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

      s.addText("Step 2.  4 groups of 2.", {
        x: 0.85, y: CONTENT_TOP + 1.85, w: 8.30, h: 0.32,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
        align: "left", margin: 0,
      });

      // Four groups of 2 counters
      const grpY = CONTENT_TOP + 2.25;
      const grpX0 = 0.95;
      for (let g = 0; g < 4; g += 1) {
        const gx = grpX0 + g * 0.90;
        s.addShape("roundRect", {
          x: gx, y: grpY, w: 0.78, h: 0.50, rectRadius: 0.08,
          fill: { color: C.BG_LIGHT || "F6F6F6" },
          line: { color: C.ACCENT, width: 1.2 },
        });
        drawCounters(s, gx + 0.10, grpY + 0.14, 2, { dotSize: 0.22, color: C.ACCENT });
      }
      s.addText("8 counters", {
        x: 4.80, y: grpY + 0.06, w: 1.40, h: 0.38,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
        bold: true, align: "left", valign: "middle", margin: 0,
      });

      // Step 3 row
      s.addText("Step 3.  Add the 6.", {
        x: 6.30, y: CONTENT_TOP + 2.30, w: 2.80, h: 0.32,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
        align: "left", margin: 0,
      });
      s.addText("6 + 8 = ?", {
        x: 6.30, y: CONTENT_TOP + 2.65, w: 2.80, h: 0.42,
        fontSize: 20, fontFace: FONT_H, color: C.ACCENT,
        bold: true, align: "left", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "4 x 2 = 8.  6 + 8 = 14.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    },
  );

  // Slide 16: You Do - three problems
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: underline the first step", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.20, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "underline the operation that comes first.   ", options: { fontSize: 16, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "do that step.   ", options: { fontSize: 16, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "finish the sentence.", options: { fontSize: 16, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.15, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const probY = CONTENT_TOP + 1.40;
    const probH = (SAFE_BOTTOM - probY - 0.20) / 3 - 0.10;
    const problems = ["8 + 3 x 2", "5 x 4 + 9", "7 + 2 x 5"];
    problems.forEach((p, i) => {
      const y = probY + i * (probH + 0.10);
      addTextOnShape(s, p, {
        x: 1.5, y, w: 7.0, h: probH, rectRadius: 0.10,
        fill: { color: C.WHITE },
        line: { color: C.PRIMARY, width: 1.6 },
      }, { fontSize: 30, fontFace: FONT_H, color: C.CHARCOAL, bold: true });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Work out  7 + 3 x 4.  Underline the step you did first.",
      "Why does it matter which operation we do first?  Answer in one short sentence.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is BODMAS, in your own words?",
      scItems: [
        "I can say WHY we need an order of operations.",
        "I can work out a number sentence with x and + using the right order.",
        "I can explain what would happen if everyone chose their own order.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BODMAS_Lesson1_Why_We_Need_Rules.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDF: BODMAS Reference Card -------------------------------------------

  await (async () => {
    const doc = createPdf({ title: PRACTICE_RES.name });
    let y = addPdfHeader(doc, PRACTICE_RES.name, {
      subtitle: "Cut along the dotted line and stick the card into your maths book.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addTipBox(doc,
      "BODMAS is the rule mathematicians agreed on for the ORDER we work out a number sentence with more than one step.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "The BODMAS order", y, { color: C.PRIMARY });
    y = addBodyText(doc, "B — Brackets.  Do anything in brackets FIRST.", y);
    y = addBodyText(doc, "O — Orders.  Powers (we will not focus on this in Year 5 and 6).", y);
    y = addBodyText(doc, "D — Division.  Do these next (left to right).", y);
    y = addBodyText(doc, "M — Multiplication.  Same level as Division (left to right).", y);
    y = addBodyText(doc, "A — Addition.  Do these last (left to right).", y);
    y = addBodyText(doc, "S — Subtraction.  Same level as Addition (left to right).", y);

    y = addSectionHeading(doc, "Worked example", y, { color: C.PRIMARY });
    y = addBodyText(doc, "5 + 3 x 2  =  5 + 6  =  11", y);
    y = addBodyText(doc, "We did the x first because M comes before A in BODMAS.", y);

    y = addSectionHeading(doc, "Quick practice (try one on a whiteboard)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "a)  4 + 2 x 3   =   _____", y);
    y = addWriteLine(doc, "b)  6 + 4 x 2   =   _____", y);
    y = addWriteLine(doc, "c)  8 + 3 x 2   =   _____", y);

    y = addTipBox(doc,
      "Underline the operation you do FIRST. That is the M or D step. Then finish with the A or S step.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | BODMAS Reference Card | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, PRACTICE_RES.fileName));
    console.log("PDF written: " + PRACTICE_RES.fileName);
  })();

  // --- PDF: Answer Key ------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 1 quick practice.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Quick practice answers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4 + 2 x 3  =  4 + 6  =  10.", y);
    y = addBodyText(doc, "b)  6 + 4 x 2  =  6 + 8  =  14.", y);
    y = addBodyText(doc, "c)  8 + 3 x 2  =  8 + 6  =  14.", y);

    y = addSectionHeading(doc, "Exit ticket answers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "7 + 3 x 4  =  7 + 12  =  19.", y);
    y = addBodyText(doc, "The x is done first because M comes before A in BODMAS.", y);

    y = addTipBox(doc,
      "Watch for: students who get 40 — they added first. Bring them back to BODMAS in tomorrow's I Do.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
