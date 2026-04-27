"use strict";

// Part, Part, Whole Introduction Unit — Lesson 1: What is a Part? What is a Whole?
// Foundation Numeracy | Term 1 Week 1 | Variant 0
// VC2MFN04 - partition and combine collections up to 10 using PPW + subitising
// Daily Review: Digit Formations — write the teen numbers
// Fluency: Counting from 1 to 20

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics,
  addTensFrame, addFiveFrame, addDotCard, addNumberTrack,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 5;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = `Part, Part, Whole | Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`;
const OUT_DIR = "output/PPWi_Lesson1_Parts_and_Whole";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 1 Cut and Paste", "Cut, sort and paste counters to show parts of 3 and 4.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 1 Answer Key", "Sample part splits for wholes of 3 and 4.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a labelled PPW mat: whole on top, two part-cells underneath.
function drawPpwMat(slide, x, y, w, h, opts) {
  const o = opts || {};
  const wholeH = h * 0.36;
  const partH  = h * 0.56;
  const partW  = (w - 0.16) / 2;

  // Whole bar (top, full width) — shows "Whole = N" or just "Whole".
  slide.addShape("roundRect", {
    x, y, w, h: wholeH, rectRadius: 0.10,
    fill: { color: o.wholeFill || C.PRIMARY },
    line: { color: o.wholeFill || C.PRIMARY, width: 1 },
  });
  const wholeText = o.wholeNumber != null
    ? "Whole = " + o.wholeNumber
    : (o.wholeLabel || "Whole");
  slide.addText(wholeText, {
    x, y, w, h: wholeH,
    fontSize: o.wholeNumber != null ? 32 : 24,
    fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  // Part A (bottom left)
  const partY = y + wholeH + 0.08;
  slide.addShape("roundRect", {
    x, y: partY, w: partW, h: partH, rectRadius: 0.10,
    fill: { color: C.WHITE },
    line: { color: o.partLine || C.SECONDARY, width: 2.5 },
  });
  slide.addText("Part", {
    x: x + 0.10, y: partY + 0.04, w: partW - 0.20, h: 0.30,
    fontSize: 13, fontFace: FONT_B, color: o.partLine || C.SECONDARY, bold: true,
    align: "left", valign: "top", margin: 0,
  });

  // Part B (bottom right)
  slide.addShape("roundRect", {
    x: x + partW + 0.16, y: partY, w: partW, h: partH, rectRadius: 0.10,
    fill: { color: C.WHITE },
    line: { color: o.partLine || C.SECONDARY, width: 2.5 },
  });
  slide.addText("Part", {
    x: x + partW + 0.26, y: partY + 0.04, w: partW - 0.20, h: 0.30,
    fontSize: 13, fontFace: FONT_B, color: o.partLine || C.SECONDARY, bold: true,
    align: "left", valign: "top", margin: 0,
  });

  // Place dots in each part if dots given
  function placeDots(boxX, boxY, count, color) {
    if (!count) return;
    const dotSize = 0.34;
    const gap = 0.10;
    const totalW = count * dotSize + (count - 1) * gap;
    const startX = boxX + (partW - totalW) / 2;
    const cy = boxY + partH * 0.62 - dotSize / 2;
    for (let i = 0; i < count; i += 1) {
      slide.addShape("roundRect", {
        x: startX + i * (dotSize + gap),
        y: cy,
        w: dotSize, h: dotSize,
        rectRadius: dotSize / 2,
        fill: { color },
        line: { color: C.CHARCOAL, width: 0.8 },
      });
    }
  }
  if (o.partA != null) placeDots(x, partY, o.partA, o.partAColor || "D64545");
  if (o.partB != null) placeDots(x + partW + 0.16, partY, o.partB, o.partBColor || "F4C430");
}

// Draw a row of two-coloured counters (red/yellow).
function drawCounterRow(slide, x, y, total, redCount, opts) {
  const o = opts || {};
  const size = o.size || 0.55;
  const gap = o.gap || 0.12;
  const colorA = o.colorA || "D64545";
  const colorB = o.colorB || "F4C430";
  for (let i = 0; i < total; i += 1) {
    slide.addShape("roundRect", {
      x: x + i * (size + gap), y,
      w: size, h: size,
      rectRadius: size / 2,
      fill: { color: i < redCount ? colorA : colorB },
      line: { color: C.CHARCOAL, width: 1 },
    });
  }
}

// Draw a numeral inside a coloured chip (used for friendly labels).
function drawNumChip(slide, x, y, w, h, num, color) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.10,
    fill: { color: color || C.PRIMARY },
  });
  slide.addText(String(num), {
    x, y, w, h,
    fontSize: 36, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Good morning everyone. Today we start a brand new maths idea.
- It is called part, part, whole.
- This is new for everyone, so we will go slowly.
- We will use counters to find the parts that make a number.

DO:
- Have 4 two-sided counters and a small cup ready on the floor mat.
- Settle students on the mat before clicking past the title.

TEACHER NOTES:
Lesson 1 of 5. Most students are seeing PPW for the first time. Treat every part of this lesson as new. Keep language concrete: whole, parts, the same.

WATCH FOR:
- Students who shrug or look unsure - that is normal at the start.
- Students who say "I know this" - listen for the language whole and parts to confirm.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up time. Let us write a teen number on our whiteboards.
- Look at the number on the slide.
- Whisper its name to your partner.
- Now write it on your whiteboard. Top to bottom. 1 first, then the next digit.

DO:
- Display the number 13.
- Allow 20-30 seconds for whisper and write.
- Walk and check the strokes.

TEACHER NOTES:
Daily Review: digit formations for teen numbers. Today we focus on 13. The 1 first, then the next digit. Foundation students often reverse teen digits.

WATCH FOR:
- Students who write 31 instead of 13 - swap order is the most common error.
- Students who form the 3 with two open curves the wrong way - model on the board.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check together. The number is thirteen.
- Trace 13 with your finger in the air.
- Tick if you wrote 13. Fix if you wrote 31.

DO:
- Click to reveal 13 written large.
- Lead a finger trace in the air.

TEACHER NOTES:
Tick and fix. Note any students who reversed the digits.

WATCH FOR:
- Students who self-correct - they noticed the error.
- Students who keep the wrong number - small group support tomorrow.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time! Stand up.
- Let us count from 1 to 20 together.
- Clap on every number. Ready... 1, 2, 3...
- One more time, a little bit faster.

DO:
- Lead choral count 1-20 with claps.
- Repeat 2-3 times, increasing pace each time.
- Sit students back down before moving on.

TEACHER NOTES:
This is automaticity, not new teaching. Keep it brisk and joyful.

WATCH FOR:
- Students who lose track at 13-15 - common stumbling spot.
- Students who count smoothly - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to find the parts that make a whole.
- A whole is the big number. The parts are the smaller groups inside it.
- Let us read the success criteria together.

DO:
- Choral read the LI.
- Choral read each success criterion.
- Hold up a cup of 4 counters as you say "whole".

TEACHER NOTES:
The first SC is achievable for all - just naming the whole. SC2 is the core target. SC3 stretches students to share their thinking with words.

WATCH FOR:
- Students who repeat back "whole" and "parts" - language is forming.
- Students who look unsure - the I Do will build the picture.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. I have 3 counters. 3 is the whole.
- I am going to put them onto the mat.
- 2 go in this part. 1 goes in this part.
- The whole is 3. The parts are 2 and 1.
- Say with me: whole 3, parts 2 and 1.

DO:
- Use the document camera or hold up a real PPW mat with 3 counters.
- Move counters into the two part-cells one at a time.
- Touch each cell as you say "part".
- Repeat the language: "whole 3, parts 2 and 1".

TEACHER NOTES:
First exposure to the PPW mat. Move slowly. The mat is the new tool. Say the language pattern at least 3 times. Concrete language is critical for Foundation students - "whole" and "parts" are new words.

MISCONCEPTIONS:
- Misconception: Students think the parts have to be the same.
  Why: Sharing fairly is a routine they may already know.
  Impact: They may not see 3 = 2 and 1 as a valid split.
  Quick correction: "Parts can be different sizes. The whole stays the same."

WATCH FOR:
- Students who repeat the language back - the pattern is forming.
- Students who count 1, 2, 3, 4 - they are counting all dots, not seeing parts. Re-show with a colour difference.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now watch this. I have 4 counters. 4 is the whole.
- I put 2 in this part. I put 2 in this part.
- The whole is 4. The parts are 2 and 2.
- Same parts this time. The parts can be the same.

DO:
- Hold up the mat with 4 counters split 2 and 2.
- Touch each part as you name it.
- Say "whole 4, parts 2 and 2" three times.

TEACHER NOTES:
Second I Do shows the parts can also be the same. Avoid showing too many splits today - we are building the language only. Future lessons will explore other ways to make the same whole.

WATCH FOR:
- Students who say "same parts" or "different parts" - they are noticing the pattern.
- Students who look confused - use real counters and move slowly.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Look at the counters on the slide.
- Count the red. Count the yellow.
- Show me on your fingers how many yellow.

DO:
- Display the picture: 3 red and 1 yellow counter.
- Wait 5 seconds.
- Say: "Show me on your fingers... ready... show."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "How many yellow? Show me on your fingers... show."
- Scan for: 1 finger held up.
PROCEED: If 80% show 1 finger, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count all the counters, not just yellow.
- Reteach: "Just one colour. Point only to yellow. How many is that?"
- Re-check: Show a different picture (3 yellow, 1 red) and ask "How many yellow now?"

TEACHER NOTES:
Fingers are universal at Foundation. Quick, every-student response.

WATCH FOR:
- Students who hold up 1 quickly - secure with parts.
- Students who hold up 4 - they counted everything; reteach colour focus.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. I have 4 cubes on the slide.
- Some are blue. Some are green.
- With your partner, whisper: How many blue? How many green?
- Show your answer on your whiteboard like this: parts are ___ and ___.

DO:
- Display 4 cubes: 2 blue, 2 green.
- Allow 30 seconds for partner whisper and whiteboard write.
- Walk and scan.

TEACHER NOTES:
We Do mirrors I Do but uses cubes instead of counters - same idea, slightly different visual to build flexibility. The whole-of-4 split 2 and 2 connects directly to I Do 2.

WATCH FOR:
- Pairs who say the language pattern aloud - secure.
- Pairs who count all 4 without separating colours - prompt: "Just blue first."

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together. The whole is 4.
- The parts are 2 and 2.
- Same parts this time! 2 and 2.

DO:
- Click to reveal.
- Repeat the language pattern with the class.

TEACHER NOTES:
Same parts is a key idea. Students who said "two and two" can lead a turn-and-tell.

WATCH FOR:
- Students who said the answer first - confident.
- Students who needed help - small group at table time.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. Look at the counters.
- I am saying: "The whole is 5. The parts are 2 and 1."
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Display 3 counters total (2 red, 1 yellow).
- Say the false claim aloud while pointing to the slide.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Whole is 5, parts are 2 and 1. Thumbs up if yes, thumbs down if no."
- Scan for: thumbs DOWN. The whole is 3, not 5.
PROCEED: If 80% show thumbs down, click to reveal correct version.
PIVOT: Most likely misconception - students agree because they trust the teacher.
- Reteach: "Count the dots all together. How many?"
- Re-check: "Is the whole 5 or 3?"

TEACHER NOTES:
This hinge probes whether students know the whole is the total of the parts. A "no" answer means they are tracking both parts and whole together.

WATCH FOR:
- Confident thumbs down - students see the mismatch.
- Slow thumbs - they are unsure; reteach with concrete counters.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- You need 4 cubes and your PPW mat.
- Put the 4 cubes on the mat. Put some in this part, some in this part.
- Tell your partner: whole 4, parts ___ and ___.

DO:
- Hand out 4 unifix cubes per child and a printed PPW mat.
- Circulate. Listen for the language pattern.
- Cold call 2 students to share their split.

TEACHER NOTES:
The first solo go. Students will put the cubes on the mat in many ways. All correct splits are valid: 0 and 4, 1 and 3, 2 and 2, 3 and 1, 4 and 0. The goal today is using the language, not finding all ways.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 3 cubes instead of 4. The whole is 3.
EXTENDING PROMPT:
- Task: After making one split, slide one cube across to the other part. Tell your partner the new parts.

WATCH FOR:
- Students who say the full pattern - secure.
- Students who put all cubes in one part - that is still valid (4 and 0). Accept and name it.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. Look at the picture.
- On your whiteboard, write the whole and the parts.

DO:
- Display: 3 red and 1 yellow counter.
- Allow 60 seconds.
- Collect whiteboards or take a quick photo for your records.

TEACHER NOTES:
Exit ticket assesses the core target - naming the whole and the parts from a picture. The whole is 4, parts are 3 and 1.

WATCH FOR:
- Students who write whole 4, parts 3 and 1 - on track.
- Students who write only one number - they are still attending to one group only.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- We learned a new idea today: part, part, whole.
- A whole is the big number. The parts are the smaller groups inside it.
- Show me thumbs up if you can name a whole and parts.
- Turn and tell your partner: what is a whole?

DO:
- Read the success criteria with the class.
- Use thumbs up sideways down.
- Cold call 1-2 students for the partner share.

TEACHER NOTES:
Self-assessment data informs Lesson 2 grouping. Most students should reach SC1 today; SC2 and SC3 build over the unit.

WATCH FOR:
- Strong thumbs up - they can use the language.
- Sideways or down - small group focus tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- These are the materials for today.
- The cut-and-paste sheet is for table time.

DO:
- Print the cut and paste sheet (one per student).
- Have unifix cubes, two-sided counters, and printed PPW mats ready.

TEACHER NOTES:
Only one printed student resource for this lesson. Most of the work happens with cubes, counters and the printed PPW mat.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: What is a Part? What is a Whole?",
    `Foundation Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slides 2-3: Daily Review with reveal — Digit formation: 13
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Write This Number", { color: C.ACCENT });

      // One large hero numeral on the right; instruction card on left
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 26, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Look at the number.", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "Write it.", options: { fontSize: 22, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Big numeral on right
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("13", {
        x: 5.2, y: CONTENT_TOP + 0.20, w: 4.3, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontSize: 220, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "thirteen   1 first, then 3", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Counting 1-20 with number track
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Count 1 to 20", { color: C.ACCENT });

  // Number track 1-20 on top; clap reminder underneath
  const trackY = CONTENT_TOP + 0.30;
  // Two rows of 10 cells for readability
  addNumberTrack(sFluency, 0.5, trackY, 9.0, 1, 10, [], { cellH: 0.85, fontSize: 32 });
  addNumberTrack(sFluency, 0.5, trackY + 1.05, 9.0, 11, 20, [], { cellH: 0.85, fontSize: 32 });

  // Stand up and clap card at bottom
  addCard(sFluency, 1.5, trackY + 2.20, 7.0, 0.85, { strip: C.PRIMARY });
  sFluency.addText("Stand up. Clap on every number.", {
    x: 1.7, y: trackY + 2.22, w: 6.6, h: 0.80,
    fontSize: 24, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to find the parts that make a whole.",
    [
      "I can point to the whole.",
      "I can name the two parts.",
      "I can use a mat to show parts and a whole.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do 1 — Whole of 3
  workedExSlide(pres, 2, "I Do", "What is a Part? What is a Whole?",
    [
      "I have 3 counters.",
      "",
      "3 is the whole.",
      "",
      "I put 2 here.",
      "I put 1 here.",
      "",
      "Whole 3.",
      "Parts 2 and 1.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      // PPW mat showing whole 3, parts 2 and 1
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 3,
        partA: 2, partB: 1,
        partAColor: "D64545",
        partBColor: "F4C430",
      });
    }
  );

  // Slide 7: I Do 2 — Whole of 4 (parts can be the same)
  workedExSlide(pres, 2, "I Do", "Same Parts. Same Whole.",
    [
      "I have 4 counters.",
      "",
      "4 is the whole.",
      "",
      "I put 2 here.",
      "I put 2 here.",
      "",
      "Whole 4.",
      "Parts 2 and 2.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 4,
        partA: 2, partB: 2,
        partAColor: "D64545",
        partBColor: "F4C430",
      });
    }
  );

  // Slides 8-9: CFU 1 with reveal — How many yellow?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How Many Yellow?", { color: C.ALERT });

      // CHECK stamp
      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      // Hero card with counters
      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Big counter row 4 counters: 3 red, 1 yellow
      drawCounterRow(s, 1.6, CONTENT_TOP + 0.45, 4, 3, { size: 1.10, gap: 0.30 });

      // Prompt below
      s.addText("Show me on your fingers.", {
        x: 0.7, y: CONTENT_TOP + 2.30, w: 8.6, h: 0.55,
        fontSize: 30, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1 yellow.   Whole 4.   Parts 3 and 1.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 10-11: We Do with reveal — whole of 4 with cubes
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "What Are the Parts?", { color: STAGE_COLORS["3"] });

      // Left card: cubes visual
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      // 4 cubes: 2 blue, 2 green
      const cubeY = CONTENT_TOP + 0.50;
      function drawCube(slide, x, y, size, color) {
        slide.addShape("roundRect", {
          x, y, w: size, h: size, rectRadius: 0.10,
          fill: { color },
          line: { color: C.CHARCOAL, width: 1.2 },
        });
      }
      drawCube(s, 0.85, cubeY, 0.85, "1976D2");
      drawCube(s, 1.85, cubeY, 0.85, "1976D2");
      drawCube(s, 2.85, cubeY, 0.85, "2E7D32");
      drawCube(s, 3.85, cubeY, 0.85, "2E7D32");

      s.addText("4 cubes", {
        x: 0.6, y: cubeY + 1.15, w: 4.3, h: 0.50,
        fontSize: 26, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      s.addText("How many BLUE? How many GREEN?", {
        x: 0.6, y: cubeY + 1.80, w: 4.3, h: 0.55,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
      });

      // Right card: prompt
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Whole = ____", options: { fontSize: 26, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Parts = ____ and ____", options: { fontSize: 24, bold: true, color: C.CHARCOAL } },
      ], {
        x: 5.4, y: CONTENT_TOP + 0.30, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.55,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Whole 4.   Parts 2 and 2.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 12-13: CFU Hinge with reveal — Is this right?
  withReveal(
    () => cfuSlide(pres, "CFU", "Is This Right?", "Thumbs Up or Thumbs Down",
      "Look at the counters.\n\nWhole is 5.\nParts are 2 and 1.\n\nIs that right?",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN!  The whole is 3, not 5.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do — 4 cubes on the mat
  workedExSlide(pres, 4, "You Do", "Your Turn: Make 4",
    [
      "First: Get 4 cubes.",
      "",
      "Next: Put some on each part of the mat.",
      "",
      "Then: Tell your partner.",
      "",
      "Whole 4.",
      "Parts ___ and ___.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      // Materials card top
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.30, { strip: C.ALERT });
      slide.addText("You need:", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.40, h: 0.36,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "4 cubes", options: { bullet: true, fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "your PPW mat", options: { bullet: true, fontSize: 18, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.40, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.50, h: 0.78,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // PPW mat below (empty)
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 1.45, lg.rightW, 2.30, {
        wholeNumber: 4,
        partA: null, partB: null,
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Whole 4, parts 3 and 1. Draw it on your whiteboard.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 16: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: what is a whole?",
      scItems: [
        "I can point to the whole.",
        "I can name the two parts.",
        "I can use a mat to show parts and a whole.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 17: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPWi_Lesson1_Parts_and_Whole.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Cut and paste worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Cut the dots. Paste them in the parts. Show the whole.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addTipBox(doc, "Cut along the dotted line. Stick the dots into the two parts. Then say: whole ___, parts ___ and ___.", y, { color: C.ACCENT });

    // PPW mat for whole = 3
    function drawPdfPpwMat(yTop, wholeLabel) {
      const matX = 70;
      const matW = 470;
      const wholeH = 60;
      const partH = 100;
      // Whole bar
      doc.lineWidth(2).rect(matX, yTop, matW, wholeH).fillAndStroke("#" + C.PRIMARY, "#333333");
      doc.fillColor("#FFFFFF").fontSize(28).font("Sans-Bold").text("Whole = " + wholeLabel, matX, yTop + 16, { width: matW, align: "center" });
      // Two part boxes
      const partGap = 10;
      const partW = (matW - partGap) / 2;
      doc.lineWidth(2).rect(matX, yTop + wholeH + 10, partW, partH).fillAndStroke("#FFFFFF", "#333333");
      doc.lineWidth(2).rect(matX + partW + partGap, yTop + wholeH + 10, partW, partH).fillAndStroke("#FFFFFF", "#333333");
      doc.fillColor("#333333").fontSize(14).font("Sans-Bold").text("Part", matX + 10, yTop + wholeH + 18);
      doc.fillColor("#333333").fontSize(14).font("Sans-Bold").text("Part", matX + partW + partGap + 10, yTop + wholeH + 18);
      doc.fillColor("#333333").font("Sans");
      return yTop + wholeH + partH + 20;
    }

    y = addSectionHeading(doc, "Whole = 3", y, { color: C.PRIMARY });
    y = drawPdfPpwMat(y, "3");
    doc.fontSize(13).fillColor("#333333").text("Whole 3.   Parts ____ and ____", 70, y);
    y += 30;

    y = addSectionHeading(doc, "Whole = 4", y, { color: C.PRIMARY });
    y = drawPdfPpwMat(y, "4");
    doc.fontSize(13).fillColor("#333333").text("Whole 4.   Parts ____ and ____", 70, y);
    y += 30;

    y = addTipBox(doc, "Cut these dots out for the mats above:", y, { color: C.SECONDARY });
    // Draw 7 large filled circles for cutting
    const cutY = y + 4;
    for (let i = 0; i < 7; i += 1) {
      doc.circle(85 + i * 60, cutY + 20, 18).lineWidth(1.4).fillAndStroke("#D64545", "#333333");
    }
    doc.fillColor("#333333");

    addPdfFooter(doc, `Lesson ${SESSION} | Lesson 1 Cut and Paste | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Sample part splits.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addBodyText(doc, "Children may show any of these splits. All are correct as long as the parts add to the whole.", y);
    y = addSectionHeading(doc, "Whole = 3", y, { color: C.PRIMARY });
    y = addBodyText(doc, "3 and 0   |   2 and 1   |   1 and 2   |   0 and 3", y);
    y = addSectionHeading(doc, "Whole = 4", y, { color: C.PRIMARY });
    y = addBodyText(doc, "4 and 0   |   3 and 1   |   2 and 2   |   1 and 3   |   0 and 4", y);
    y = addTipBox(doc, "Today's focus is the language pattern (whole ___, parts ___ and ___). All splits are valid. The next lesson explores parts of 5.", y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
