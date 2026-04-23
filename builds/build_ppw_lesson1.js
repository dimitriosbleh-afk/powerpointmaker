"use strict";

// Part Part Whole Unit — Session 1: Shake and Spill
// Foundation Numeracy | Term 2 Week 4 | Variant 3
// DR: Part Part Whole 1-5 with 2-sided counters
// Fluency: Counting 1-20
// VC2MFN04 — partition and combine collections up to 10 using part-part-whole relationships and subitising

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", weekToVariant(4));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = "Part, Part, Whole | Session 1 of 4 | Foundation Numeracy";
const OUT_DIR = "output/PPW_Session1_Shake_and_Spill";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Shake and Record", "Recording sheet for Shake and Spill with counters.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Sample part-part-whole splits for Shake and Record.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-drawn counter dots for numbers 3, 4, and 5.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Shake and Spill with 10 counters, draw all the ways.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a row of two-sided counters: red for "part A", yellow for "part B"
function drawTwoColourCounters(slide, x, y, total, partA, opts) {
  const o = opts || {};
  const size = o.size || 0.42;
  const gap = o.gap || 0.08;
  const colorA = o.colorA || "D64545"; // red
  const colorB = o.colorB || "F4C430"; // yellow
  const borderColor = C.CHARCOAL;
  for (let i = 0; i < total; i++) {
    const cx = x + i * (size + gap);
    slide.addShape("roundRect", {
      x: cx, y, w: size, h: size, rectRadius: size / 2,
      fill: { color: i < partA ? colorA : colorB },
      line: { color: borderColor, width: 1 },
    });
  }
}

// Draw a labelled PPW diagram (whole on top, two parts below)
function drawPpwDiagram(slide, x, y, w, h, whole, partA, partB, opts) {
  const o = opts || {};
  const wholeH = h * 0.38;
  const partH = h * 0.44;
  const partW = w * 0.46;
  const gap = w - 2 * partW;

  // Whole box (top, full width) — big number only
  slide.addShape("roundRect", {
    x, y, w, h: wholeH, rectRadius: 0.08,
    fill: { color: o.wholeColor || C.PRIMARY },
    line: { color: C.PRIMARY, width: 2 },
  });
  slide.addText(String(whole), {
    x, y, w, h: wholeH,
    fontSize: 40, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Part A (bottom left) — big number only
  slide.addShape("roundRect", {
    x, y: y + wholeH + 0.10, w: partW, h: partH, rectRadius: 0.08,
    fill: { color: o.partAColor || "D64545" },
    line: { color: o.partAColor || "D64545", width: 2 },
  });
  slide.addText(String(partA), {
    x, y: y + wholeH + 0.10, w: partW, h: partH,
    fontSize: 36, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Part B (bottom right) — big number only
  slide.addShape("roundRect", {
    x: x + partW + gap, y: y + wholeH + 0.10, w: partW, h: partH, rectRadius: 0.08,
    fill: { color: o.partBColor || "F4C430" },
    line: { color: o.partBColor || "F4C430", width: 2 },
  });
  slide.addText(String(partB), {
    x: x + partW + gap, y: y + wholeH + 0.10, w: partW, h: partH,
    fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Good morning! Today we start a new maths adventure
- This week we are learning about part, part, whole
- Every number is made of smaller parts
- We will use counters, cubes and tens frames to find the parts

DO:
- Have a plastic cup and 10 two-sided counters ready for your demo
- Display title slide as students settle on the mat

TEACHER NOTES:
Session 1 of 4 for the Part-Part-Whole unit. This session uses Shake and Spill to introduce the PPW concept with numbers to 5 first, then up to 10. Keep the language simple: whole, parts, and together.

WATCH FOR:
- Students who can already say "5 is 3 and 2" - prior knowledge is present
- Students who seem new to the idea - that is okay, the visuals will build it

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up time. Look at the counters on the slide
- Count the red ones. Now count the yellow ones
- Whisper to your partner what the two parts are

DO:
- Read the first counter picture with the class
- Allow 10-15 seconds of partner whispering per question
- Students can use fingers to show the counts

TEACHER NOTES:
Daily Review on PPW 1-5 with two-sided counters primes today's lesson. Keep it quick and playful - this is retrieval, not teaching.

WATCH FOR:
- Students who count each colour separately and then add - they already see the two parts
- Students who count every counter from 1 - they are still seeing the whole as one group

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let's check our answers together
- For each one, say: The whole is ___. The parts are ___ and ___
- Tick your answer if it was right. Fix it if you need to.

DO:
- Click to reveal the answers
- Repeat the whole-and-parts language pattern with the class
- Note who needed help with the smaller numbers - they will need support today

TEACHER NOTES:
Tick-and-fix gives immediate feedback. The language pattern "The whole is ___. The parts are ___ and ___." is the anchor for today's lesson.

WATCH FOR:
- Students who say the full pattern without prompting - they are ready for larger numbers
- Students who only say one number - model the full pattern for them

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time! Let's count from 1 all the way to 20
- Stand up. Clap on every number
- Ready? 1, 2, 3...
- Now let's do it again, this time a bit faster

DO:
- Lead choral count 1-20 with claps
- Repeat 2-3 times, increasing pace
- Have students sit back down before moving on

TEACHER NOTES:
Quick counting fluency. This is about rhythm and automaticity, not new teaching. Keeps energy up before the lesson focus.

WATCH FOR:
- Students who lose track around 13-14 - note for extra support later
- Students who count fluently - counting is secure

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- We are learning to partition numbers to 10 into two parts using materials
- Partition means to split into parts
- Let's read the success criteria together

DO:
- Choral read the LI and each SC
- Ask: What is a part? What is a whole?
- Accept student ideas - do not correct yet

TEACHER NOTES:
SC1 is the floor - all students should reach this with support. SC2 is the core target. SC3 extends to explaining their thinking. Partition is a new word - keep it simple: "splitting a number into parts".

WATCH FOR:
- Students who can already show a number as parts - they will lead their table
- Students who look unsure - the visuals in I Do will build the concept

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. I have 5 counters in my cup
- 5 is the whole
- I shake them and spill them onto the table
- Look - some landed red, some landed yellow
- Count the red: 3. Count the yellow: 2. That is 2 parts.
- The whole is 5. The parts are 3 and 2.

DO:
- Hold up a real cup with 5 counters for students to see
- Shake and spill on a table or under the document camera
- Point to the visual on the slide as you say each part
- Repeat the sentence: "The whole is 5. The parts are 3 and 2."

TEACHER NOTES:
This I Do makes the PPW concept concrete. The spill shows partition happens naturally - we do not choose, the counters land that way. Students see the whole stays 5 even when the parts change.

MISCONCEPTIONS:
- Misconception: Students think the parts can be different numbers, so the whole is not 5 anymore
  Why: At Foundation, "5" can feel fixed to "5 counters all together" rather than 5 as a quantity that can be split
  Impact: They cannot see 2 and 3, or 4 and 1, as both ways to make 5
  Quick correction: "Look - I did not take any counters away. The parts are different but the whole is still 5."

WATCH FOR:
- Students who nod and say the pattern with you - they are picking up the language
- Students who look confused - use the cup and counters again with a slower pace

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Let's see another way to make 5
- Same cup, same 5 counters
- I shake again... and spill
- This time, 4 landed red and 1 landed yellow
- The whole is still 5. The parts are 4 and 1.
- Wait - before we had 3 and 2. Now we have 4 and 1. Both make 5!

DO:
- Shake and spill again under the document camera (or use the slide visual)
- Point to the two different ways on the slide
- Emphasise: "Same whole, different parts"

TEACHER NOTES:
The key teaching move here is showing the same whole (5) can be made with different parts. This is the foundation of addition and subtraction. Say "same whole, different parts" many times today.

WATCH FOR:
- Students who say "same whole, different parts" back to you - the concept is forming
- Students who focus only on one way - they need more exposures

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. I will show you a picture of counters
- Count the red ones. Count the yellow ones.
- Show me on your fingers how many yellow
- Then tell me the whole and the parts

DO:
- Display the picture (whole of 4: 3 red, 1 yellow)
- Scan fingers held up for the yellow count
- Cold call 2 students for the parts

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "Look at the counters. How many yellow? Show me on your fingers. Ready... show me."
- Scan for: 1 finger. Students who show 1 are reading the picture correctly.
PROCEED: If 80%+ show 1, move to the We Do.
PIVOT: Most likely misconception - students count all counters instead of just the yellow part. Reteach: "The part is just one colour. Point to only the yellow ones. How many did you point to?"

TEACHER NOTES:
Fingers are universal for Foundation. Quick and every student responds. Use a real cup and counters if some students need a second example.

WATCH FOR:
- Students who show 1 quickly - they can separate the parts
- Students who show 4 - they counted everything; reteach the "just one colour" idea

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. I have a mystery spill on the slide
- With your partner, look at the picture
- Whisper: What is the whole? What are the parts?
- Hold up your whiteboard when you are ready

DO:
- Display the whole-of-6 spill: 4 red, 2 yellow
- Partners whisper for 20 seconds
- Students write on whiteboards: "6 = 4 and 2"

TEACHER NOTES:
We Do uses a bigger whole (6) to stretch from the I Do (5). Same routine, one number bigger.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 1 Enabling Scaffold. Pre-drawn dot pictures show the red and yellow counters already grouped. Students only need to count each group and say the parts.
- Extra Notes: Distribute before the We Do begins.
EXTENDING PROMPT:
- Task: After answering, students make their own spill with 8 counters and record both parts on their whiteboard.
- Extra Notes: Have 10-counter cups ready on your extender table.

WATCH FOR:
- Students who write "6 = 4 and 2" - they have the concept and the language
- Students who write only one number - remind them a spill always has two parts (even if one is zero)

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. The whole is 6. The parts are 4 and 2.
- Say it with me: 6 is 4 and 2
- Ask: Could we make 6 a different way? [5 and 1, 3 and 3, 6 and 0]

DO:
- Reveal the answer
- Repeat the pattern: "6 is 4 and 2"
- Accept 2-3 other ways and record on the class board

TEACHER NOTES:
Asking for other ways opens the door to Lesson 2 (finding all ways). Keep it brief - just affirm that 6 has many ways.

WATCH FOR:
- Students who call out other ways like "5 and 1" - they are generalising
- Students who repeat "4 and 2" only - still working on the one example

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Thumbs up if I am right. Thumbs down if I am wrong.
- The whole is 5. The parts are 2 and 3. Is that right?
- Show me... now.

DO:
- Thumbs up is correct (2 + 3 = 5)
- Then try a wrong one: "The whole is 5. The parts are 2 and 4. Is that right?"
- Thumbs down: 2 + 4 = 6, not 5
- Cold call: "How do you know?"

CFU CHECKPOINT:
Technique: Thumbs Up/Down
Script:
- Say: "Whole is 5, parts are 2 and 3. Is that right? Think... show me."
- Scan for: thumbs up. Then: "Whole is 5, parts are 2 and 4. Is that right?" Scan for: thumbs down.
PROCEED: If 80%+ get both right, move to You Do.
PIVOT: Most likely misconception - students focus on the parts and forget to check the whole. Reteach: "We check by adding the parts. 2 and 4 is... 6. Not 5. So this one is wrong."

TEACHER NOTES:
Two-part hinge - positive then negative - confirms students are actually checking, not just agreeing.

WATCH FOR:
- Thumbs down with confidence on the wrong one - they are checking the whole
- Students who hesitate - they are still trusting the teacher instead of the numbers

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own
- First: Put 5 counters in your cup
- Next: Shake and spill on your table
- Then: Draw what you see on your sheet. Write the parts.
- You have 8 minutes. Go.

DO:
- Hand out the Session 1 Shake and Record sheet
- Each student needs a cup and 5 two-sided counters
- Circulate: start with students who needed help in We Do
- Enabling students get the Session 1 Enabling Scaffold (pre-drawn dots)
- Extending students get the Session 1 Extension (use 10 counters)

TEACHER NOTES:
You Do mirrors the class demo but every student does it independently. The Shake and Record sheet has 3 spill boxes so students can do it multiple times and see different parts each time.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 1 Enabling Scaffold has the counters pre-drawn. Students circle the red group, circle the yellow group, and write the two parts.
- Extra Notes: Numbers 3, 4, and 5 only. One group per page.
EXTENDING PROMPT:
- Task: Session 1 Extension asks students to shake 10 counters, then try to find 3 different ways to make 10. Draw each way.
- Extra Notes: Self-contained - dots and prompts provided.

WATCH FOR:
- Students who draw and label independently - they understand
- Students who need help counting - work with them at the enabler table
- Readiness signal: filling all 3 spill boxes with different parts

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Work on your own.
- Look at each picture. Write the whole and the parts.
- You have 3 minutes.

DO:
- Display the exit ticket
- Students respond on whiteboards or in their workbook
- Collect or scan to sort into secure, developing, beginning

TEACHER NOTES:
Exit ticket assesses SC1 (Q1 - identify parts of a small number) and SC2 (Q2 - identify parts of a larger number). Q3 touches SC3 (explain the PPW language).

WATCH FOR:
- Students who complete all three correctly - ready for Lesson 2 (finding all the ways)
- Students who only get Q1 - they need more concrete practice in Lesson 2

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our learning. Look at the success criteria.
- Thumbs up, sideways, or down for each one
- Turn and tell your partner: What is a part? What is a whole?

DO:
- Display success criteria
- Run thumbs check for each SC
- 30 seconds turn and talk
- Cold call 2 students to share

TEACHER NOTES:
Closing connects back to today's language. Students who can say "a part is one piece, a whole is all of it together" have grasped the core idea.

WATCH FOR:
- Students saying "part" and "whole" correctly in their own words - the concept is forming
- Students who cannot yet explain - they will get more practice in Lessons 2-4

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- These are the resources for today
- The Shake and Record sheet is the main worksheet
- Enabler and extension are on your differentiation shelf

DO:
- Point out each resource

TEACHER NOTES:
Resource slide for teacher reference. Materials needed: cups, two-sided counters (5 per student), whiteboards.

WATCH FOR:
- N/A - teacher reference only

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 1: Shake and Spill",
    "Foundation Numeracy | Session 1 of 4 | Term 2 Week 4", NOTES_TITLE);

  // Slide 2-3: Daily Review (withReveal) — PPW 1-5 with 2-sided counters
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Two-Sided Counters", { y: 0.65, fontSize: 24, color: STAGE_COLORS["1"] });

      // Left: instruction card
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText([
        { text: "Look at each spill.", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "How many RED?", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "How many YELLOW?", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "What is the whole?", options: { bullet: true, fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: three counter spills
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });

      const spills = [
        { label: "A.", total: 3, red: 2, yellow: 1 },
        { label: "B.", total: 4, red: 1, yellow: 3 },
        { label: "C.", total: 5, red: 3, yellow: 2 },
      ];
      spills.forEach((sp, i) => {
        const rowY = CONTENT_TOP + 0.25 + i * 1.1;
        s.addText(sp.label, {
          x: 5.35, y: rowY, w: 0.4, h: 0.5,
          fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
          align: "center", valign: "middle",
        });
        drawTwoColourCounters(s, 5.9, rowY + 0.05, sp.total, sp.red, { size: 0.42, gap: 0.10 });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A) whole 3, parts 2 + 1     B) whole 4, parts 1 + 3     C) whole 5, parts 3 + 2", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Counting 1-20
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Count from 1 to 20", { y: 0.65, fontSize: 24, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFluency.addText([
    { text: "Stand up!", options: { fontSize: 20, bold: true, color: C.ALERT, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Clap on every number.", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
    { text: "Count from 1 to 20.", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
    { text: "Now faster!", options: { bullet: true, fontSize: 16, color: C.CHARCOAL } },
  ], {
    x: 0.75, y: CONTENT_TOP + 0.2, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  // Right: numbers 1-20 in a 4x5 grid
  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
  const gridStartX = 5.45;
  const gridStartY = CONTENT_TOP + 0.15;
  const gridCellW = 0.78;
  const gridCellH = 0.62;
  for (let i = 0; i < 20; i++) {
    const row = Math.floor(i / 5);
    const col = i % 5;
    const num = i + 1;
    sFluency.addShape("roundRect", {
      x: gridStartX + col * gridCellW,
      y: gridStartY + row * gridCellH,
      w: gridCellW - 0.08,
      h: gridCellH - 0.08,
      rectRadius: 0.08,
      fill: { color: C.PRIMARY },
      line: { color: C.PRIMARY, width: 1 },
    });
    sFluency.addText(String(num), {
      x: gridStartX + col * gridCellW,
      y: gridStartY + row * gridCellH,
      w: gridCellW - 0.08,
      h: gridCellH - 0.08,
      fontSize: 20, fontFace: FONT_H, color: C.WHITE,
      align: "center", valign: "middle", bold: true, margin: 0,
    });
  }

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to partition numbers to 10 into two parts using materials"],
    [
      "I can show a number as two parts",
      "I can find different ways to make the same number",
      "I can use pictures, objects, or numbers to explain my thinking",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — What are parts and whole?
  workedExSlide(pres, 2, "I Do", "Parts and Whole",
    [
      "I have 5 counters",
      "5 is the WHOLE",
      "",
      "I shake and spill",
      "3 land red. 2 land yellow.",
      "",
      "3 and 2 are the PARTS",
      "The whole is still 5",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });

      // PPW diagram: whole 5, parts 3 and 2
      drawPpwDiagram(slide, lg.rightX + 0.4, lg.panelTopPadded + 0.18, lg.rightW - 0.8, 2.1, 5, 3, 2);

      // Counter row below diagram
      drawTwoColourCounters(slide, lg.rightX + 0.5, lg.panelTopPadded + 2.50, 5, 3, { size: 0.42, gap: 0.09 });
    }
  );

  // Slide 7: I Do #2 — Same whole, different parts
  workedExSlide(pres, 2, "I Do", "Same Whole, Different Parts",
    [
      "Same cup. Same 5 counters.",
      "",
      "I shake again",
      "",
      "This time:",
      "4 land red. 1 lands yellow.",
      "",
      "Whole is still 5!",
      "Parts are now 4 and 1.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.5, { strip: C.SECONDARY });

      slide.addText("Way 1", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "left",
      });
      drawTwoColourCounters(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.5, 5, 3, { size: 0.38, gap: 0.08 });
      slide.addText("5 = 3 and 2", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.0, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0, align: "center",
      });

      slide.addText("Way 2", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.55, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "left",
      });
      drawTwoColourCounters(slide, lg.rightX + 0.3, lg.panelTopPadded + 1.95, 5, 4, { size: 0.38, gap: 0.08 });
      slide.addText("5 = 4 and 1", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.45, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });

      slide.addText("Same whole. Different parts.", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.95, w: lg.rightW - 0.3, h: 0.35,
        fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, italic: true, margin: 0, align: "center",
      });
    }
  );

  // Slide 8-9: CFU 1 (withReveal) — Finger Voting with counters picture
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["2"]);
      addStageBadge(s, 2, "CFU");
      addTitle(s, "How Many Yellow?", { y: 0.65, fontSize: 24, color: STAGE_COLORS["2"] });

      // Left: question card
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["2"] });
      s.addText([
        { text: "Show me on your fingers:", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "How many YELLOW?", options: { fontSize: 22, bold: true, color: C.ALERT, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Then say the parts:", options: { fontSize: 14, color: C.MUTED, breakLine: true } },
        { text: "Whole is ___. Parts are ___ and ___.", options: { fontSize: 14, italic: true, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.2, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: spill picture - whole 4, 3 red 1 yellow
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText("Look at this spill:", {
        x: 5.35, y: CONTENT_TOP + 0.2, w: 4, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });
      // Big counter picture
      drawTwoColourCounters(s, 5.7, CONTENT_TOP + 0.9, 4, 3, { size: 0.7, gap: 0.15 });
      s.addText("Whole = 4", {
        x: 5.35, y: CONTENT_TOP + 1.9, w: 4, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0, align: "center",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1 yellow!   Whole is 4.  Parts are 3 and 1.", {
        x: 1.0, y: 4.1, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do (withReveal) — whole of 6 spill
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "What Are the Parts?", { y: 0.65, fontSize: 24, color: STAGE_COLORS["3"] });

      // Left: instructions
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "With your partner:", options: { fontSize: 18, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Count the red.", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "Count the yellow.", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "What is the whole?", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Write on your whiteboard:", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "6 = ___ and ___", options: { fontSize: 18, bold: true, color: C.PRIMARY } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: spill 6 = 4 red + 2 yellow
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText("The Spill", {
        x: 5.35, y: CONTENT_TOP + 0.2, w: 4, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });
      // 6 counters in two rows of 3
      drawTwoColourCounters(s, 5.7, CONTENT_TOP + 0.95, 3, 3, { size: 0.55, gap: 0.12 });
      drawTwoColourCounters(s, 5.7, CONTENT_TOP + 1.7, 3, 1, { size: 0.55, gap: 0.12 });
      s.addText("Whole = 6", {
        x: 5.35, y: CONTENT_TOP + 2.5, w: 4, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0, align: "center",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "6 = 4 and 2     (whole is 6, parts are 4 and 2)", {
        x: 0.8, y: 4.15, w: 8.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Is This Right?", "Thumbs Up/Down",
      "Whole = 5\n\nParts = 2 and 4\n\nIs that right?\n\nThumbs UP = YES    Thumbs DOWN = NO",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN!  2 and 4 make 6, not 5", {
        x: 1.0, y: 4.1, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do — Shake and Spill on your own
  workedExSlide(pres, 4, "You Do", "Your Turn: Shake and Spill",
    [
      "First: Put 5 counters in your cup",
      "",
      "Next: Shake and spill",
      "",
      "Then: Draw what you see.",
      "Write the parts.",
      "",
      "Do it 3 times.",
      "You have 8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.ALERT });
      slide.addText("You need:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "a cup", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "5 two-sided counters", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "your recording sheet", options: { bullet: true, fontSize: 15, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.5,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Tip card
      addCard(slide, lg.rightX, lg.panelTopPadded + 2.4, lg.rightW, 0.9, { strip: C.PRIMARY });
      slide.addText("Tip: 5 = ___ red and ___ yellow", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.5, w: lg.rightW - 0.4, h: 0.7,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
        align: "center", valign: "middle",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Look at 4 red and 1 yellow counter. What is the whole?",
      "The whole is 6. Draw one way to make 6.",
      "Tell your teacher what 'part' means.",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "Turn and tell: What is a part? What is a whole?",
    [
      "I can show a number as two parts",
      "I can find different ways to make the same number",
      "I can use pictures, objects, or numbers to explain my thinking",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPW_Session1_Shake_and_Spill.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Worksheet: Shake and Record
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Shake your cup. Spill the counters. Draw what you see.",
      color: C.PRIMARY,
      lessonInfo: "Session 1 of 4 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Put 5 counters in your cup. Shake and spill. Colour the dots to show what landed red and what landed yellow. Write the parts.", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Shake 1", y, { color: C.PRIMARY });
    // Draw 5 empty circles for students to colour
    const drawCircleRow = (ty, count) => {
      const size = 28;
      const gap = 10;
      const startX = 60;
      for (let i = 0; i < count; i++) {
        doc.circle(startX + i * (size + gap) + size / 2, ty + size / 2, size / 2).lineWidth(1.5).stroke("#333333");
      }
      return ty + size + 10;
    };
    y = drawCircleRow(y, 5);
    doc.fontSize(13).fillColor("#333333").text("Whole = 5.   Parts = ____ and ____", 60, y);
    y += 34;

    y = addSectionHeading(doc, "Shake 2", y, { color: C.PRIMARY });
    y = drawCircleRow(y, 5);
    doc.fontSize(13).fillColor("#333333").text("Whole = 5.   Parts = ____ and ____", 60, y);
    y += 34;

    y = addSectionHeading(doc, "Shake 3", y, { color: C.PRIMARY });
    y = drawCircleRow(y, 5);
    doc.fontSize(13).fillColor("#333333").text("Whole = 5.   Parts = ____ and ____", 60, y);
    y += 34;

    y = addTipBox(doc, "Did you get the same parts each time? Or were they different?", y, { color: C.SECONDARY });

    addPdfFooter(doc, "Session 1 | Shake and Record | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Shake and Record - sample responses",
      color: C.PRIMARY,
      lessonInfo: "Session 1 of 4 | Foundation Numeracy",
    });
    y = addBodyText(doc, "Every shake will be different. These are sample outcomes for a whole of 5:", y);
    y = addSectionHeading(doc, "Possible PPW pairs for whole = 5", y, { color: C.PRIMARY });
    y = addBodyText(doc, "5 and 0  (all red, no yellow)", y);
    y = addBodyText(doc, "4 and 1  (four red, one yellow)", y);
    y = addBodyText(doc, "3 and 2  (three red, two yellow)", y);
    y = addBodyText(doc, "2 and 3  (two red, three yellow)", y);
    y = addBodyText(doc, "1 and 4  (one red, four yellow)", y);
    y = addBodyText(doc, "0 and 5  (no red, all yellow)", y);
    y = addTipBox(doc, "Any spill that adds to 5 is correct. Look for students who write the same pair each time - they may be recording what they remember, not what they see.", y, { color: C.ACCENT });
    addPdfFooter(doc, "Session 1 | Answer Key | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — smaller numbers, grouped counters pre-drawn
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Smaller wholes. Counters are drawn for you.",
      color: C.ACCENT,
      lessonInfo: "Session 1 of 4 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Look at the red group and the yellow group. Count each group. Write the two parts.", y, { color: C.ACCENT });

    // Helper: draw coloured filled circles
    const drawColouredRow = (ty, redCount, yellowCount) => {
      const size = 28;
      const gap = 10;
      const startX = 60;
      for (let i = 0; i < redCount; i++) {
        doc.circle(startX + i * (size + gap) + size / 2, ty + size / 2, size / 2)
          .lineWidth(1.2).fillAndStroke("#D64545", "#333333");
      }
      for (let i = 0; i < yellowCount; i++) {
        const cx = startX + (redCount + i) * (size + gap) + size / 2;
        doc.circle(cx, ty + size / 2, size / 2)
          .lineWidth(1.2).fillAndStroke("#F4C430", "#333333");
      }
      doc.fillColor("#333333");
      return ty + size + 14;
    };

    y = addSectionHeading(doc, "Whole = 3", y, { color: C.PRIMARY });
    y = drawColouredRow(y, 2, 1);
    doc.fontSize(13).fillColor("#333333").text("Parts = ____ red and ____ yellow", 60, y);
    y += 34;

    y = addSectionHeading(doc, "Whole = 4", y, { color: C.PRIMARY });
    y = drawColouredRow(y, 3, 1);
    doc.fontSize(13).fillColor("#333333").text("Parts = ____ red and ____ yellow", 60, y);
    y += 34;

    y = addSectionHeading(doc, "Whole = 5", y, { color: C.PRIMARY });
    y = drawColouredRow(y, 3, 2);
    doc.fontSize(13).fillColor("#333333").text("Parts = ____ red and ____ yellow", 60, y);
    y += 34;

    y = addTipBox(doc, "Ask your teacher: Could 3 be made any other way?", y, { color: C.SECONDARY });

    addPdfFooter(doc, "Session 1 | Enabling Scaffold | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — 10 counters, find different ways
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Shake and Spill with 10 counters",
      color: C.PRIMARY,
      lessonInfo: "Session 1 of 4 | Foundation Numeracy",
    });
    y = addBodyText(doc, "You are a Part-Part-Whole explorer! Put 10 counters in your cup.", y);
    y = addBodyText(doc, "Shake and spill. Draw what you see. Try to find 3 different ways to make 10.", y);

    const drawEmptyRow = (ty, count) => {
      const size = 24;
      const gap = 8;
      const startX = 60;
      for (let i = 0; i < count; i++) {
        doc.circle(startX + i * (size + gap) + size / 2, ty + size / 2, size / 2)
          .lineWidth(1.2).stroke("#333333");
      }
      return ty + size + 12;
    };

    y = addSectionHeading(doc, "Way 1: 10 = ____ and ____", y, { color: C.PRIMARY });
    y = drawEmptyRow(y, 10);
    y += 6;

    y = addSectionHeading(doc, "Way 2: 10 = ____ and ____", y, { color: C.PRIMARY });
    y = drawEmptyRow(y, 10);
    y += 6;

    y = addSectionHeading(doc, "Way 3: 10 = ____ and ____", y, { color: C.PRIMARY });
    y = drawEmptyRow(y, 10);
    y += 6;

    y = addTipBox(doc, "Challenge: Can you find more than 3 ways to make 10? How many ways are there in total?", y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 1 | Extension | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 1 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
