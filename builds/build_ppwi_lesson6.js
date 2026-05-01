"use strict";

// Part, Part, Whole Introduction Unit — Lesson 6: All the Ways to Make 5
// Foundation Numeracy | Lesson 6 of 10 | Variant 0
// VC2MFN04 - partition and combine collections up to 10 using PPW + subitising
// Daily Review: Digit Formations - write the teen numbers
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
  titleSlide, liSlide, cfuSlide, closingSlide, contentSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  addNumberTrack,
  STAGE_COLORS,
} = T;

const SESSION = 6;
const TOTAL = 10;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = `Part, Part, Whole | Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`;
const OUT_DIR = "output/PPWi_Lesson6_All_Ways_Make_5";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 6 All the Ways to Make 5", "Show every way to make 5 with two parts.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 6 Answer Key", "All the ways to make 5.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

function drawPpwMat(slide, x, y, w, h, opts) {
  const o = opts || {};
  const wholeH = h * 0.36;
  const partH  = h * 0.56;
  const partW  = (w - 0.16) / 2;

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

  function placeDots(boxX, boxY, count, color) {
    if (!count) return;
    const dotSize = 0.30;
    const gap = 0.08;
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

// Draw a small mini-mat showing one way to make a whole. Used for "all the ways" gallery.
function drawMiniWayCard(slide, x, y, w, h, partA, partB, opts) {
  const o = opts || {};
  // Card background
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.10,
    fill: { color: C.WHITE },
    line: { color: o.borderColor || C.SECONDARY, width: 2 },
  });
  // Equation strip at top
  const eqH = 0.42;
  slide.addText(`${partA} and ${partB}`, {
    x, y: y + 0.05, w, h: eqH,
    fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  // Counter row centred under the equation
  const total = partA + partB;
  const dotSize = 0.26;
  const gap = 0.06;
  const totalW = total * dotSize + (total - 1) * gap;
  const startX = x + (w - totalW) / 2;
  const cy = y + 0.65;
  for (let i = 0; i < total; i += 1) {
    slide.addShape("roundRect", {
      x: startX + i * (dotSize + gap),
      y: cy,
      w: dotSize, h: dotSize,
      rectRadius: dotSize / 2,
      fill: { color: i < partA ? "D64545" : "F4C430" },
      line: { color: C.CHARCOAL, width: 0.7 },
    });
  }
}

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

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back to part, part, whole.
- Last week we found parts of small numbers.
- Today we will show every way to make 5.

DO:
- Have 5 two-sided counters and a PPW mat ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 6 of 10. Students have met whole, parts, and the PPW mat. Today's new step is finding more than one way to make the same whole and recording each way.

WATCH FOR:
- Students who recall the words whole and parts - good.
- Students who look unsure - we will rebuild the language together.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are the materials for today.
- We will use the All Ways to Make 5 sheet at the table.

DO:
- Print the All Ways to Make 5 sheet (one per child).
- Have 5 unifix cubes or two-sided counters per child and a PPW mat.

TEACHER NOTES:
The sheet records every way two parts can make 5. It mirrors the PPW mats students have used.

[General: Resources]`;

const NOTES_LAUNCH = `SAY:
- Look at this. I have 5 counters in my cup.
- Watch what happens when I tip them out.
- Some land here. Some land here.
- Now watch again.
- Same 5 counters. Different parts.
- Today we are going to find every way to make 5.

DO:
- Tip 5 two-sided counters out of a cup.
- Land them with a different split each time (4 and 1, then 2 and 3).
- Point to the whole each time and say still 5.

TEACHER NOTES:
This launch reminds students that the whole stays the same while the parts can change. Keep it short - 60 to 90 seconds. The new idea today is finding ALL the ways.

WATCH FOR:
- Students who say still 5 - the whole idea is secure.
- Students who recount each time - normal at this stage.

[General: Launch | VTLM 2.0: Activate Prior Knowledge]`;

const NOTES_DR_Q = `SAY:
- Warm up time. Whiteboards out.
- Look at the number on the slide.
- Whisper its name to your partner.
- Now write it. 1 first, then the next digit.

DO:
- Display the number 14.
- Allow 20 to 30 seconds.
- Walk and check the strokes.

TEACHER NOTES:
Daily Review: digit formations for teen numbers. Today is 14. Watch for reversed digits (41).

WATCH FOR:
- Students who write 41 instead of 14 - swap order is the most common error.
- Students who form the 4 with three strokes correctly.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check together. The number is fourteen.
- Trace 14 in the air with your finger.
- Tick if you wrote 14. Fix if you wrote 41.

DO:
- Click to reveal 14.
- Lead a finger trace.

TEACHER NOTES:
Tick and fix. Note any students who reversed the digits.

WATCH FOR:
- Students who self-correct - they noticed the error.
- Students who keep the wrong number - small group support.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Stand up tall.
- We are counting from 1 to 20 with claps.
- Ready... 1, 2, 3...
- One more time, a little faster.

DO:
- Lead choral count 1-20 with claps.
- Repeat 2 to 3 times, increasing pace.
- Sit students back down before moving on.

TEACHER NOTES:
This is automaticity, not new teaching. Keep it brisk and joyful.

WATCH FOR:
- Students who lose track around 13 to 15 - common stumbling spot.
- Students who count smoothly - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to find different ways to make 5.
- Read the success criteria with me.

DO:
- Choral read the LI.
- Choral read each criterion.
- Hold up 5 counters as you say five.

TEACHER NOTES:
The first SC is achievable for all - just pointing to the whole. SC2 is the core target. SC3 stretches students to record more than one way.

WATCH FOR:
- Students who chime in - language is forming.
- Students who look quiet - I Do will rebuild the picture.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. I have 5 counters. 5 is the whole.
- I put 4 in this part. I put 1 in this part.
- Whole 5. Parts 4 and 1.
- Say with me: whole 5, parts 4 and 1.

DO:
- Hold up the PPW mat with 5 counters.
- Move 4 counters to the left part, 1 to the right.
- Point to each part as you name it.
- Repeat the language pattern at least 3 times.

TEACHER NOTES:
First way today. Move slowly. The new step is naming this as one way out of several.

WATCH FOR:
- Students who echo whole 5, parts 4 and 1 - language secure.
- Students who count all 5 dots without separating - re-show with colour difference.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now watch this. Same 5 counters. I move them.
- I put 3 here. I put 2 here.
- Whole 5. Parts 3 and 2.
- Same whole. Different parts.

DO:
- Slide one counter from the left part to the right part.
- Now show 3 and 2.
- Point to the whole and say still 5.
- Add a tally or tick mark on the board to show way one and way two.

TEACHER NOTES:
This is the key insight - we can find another way and another way. Today we record TWO ways together. We Do and You Do will explore more.

WATCH FOR:
- Students who say new parts - they are noticing the shift.
- Students who think we now have 10 - they think the parts are extra. Re-show: same 5 counters, just moved.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Look at the counters on the slide.
- Count the red. Count the yellow.
- Show me on your fingers how many yellow.

DO:
- Display 5 counters: 1 red, 4 yellow.
- Wait 5 seconds.
- Say: Show me on your fingers... ready... show.

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: How many yellow? Show me on your fingers... show.
- Scan for: 4 fingers held up.
PROCEED: If 80% show 4 fingers, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count all 5 counters, not just yellow.
- Reteach: Just one colour. Point only to yellow. How many is that?
- Re-check: Show 5 counters with 2 red, 3 yellow and ask How many yellow now?

TEACHER NOTES:
Fingers are universal at Foundation. Quick, every-student response.

WATCH FOR:
- 4 fingers held up quickly - secure with parts.
- 5 fingers - they counted everything; reteach colour focus.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Look at the mini cards on the slide.
- We have shown two ways to make 5 already.
- With your partner, whisper: what other ways can we find?
- On your whiteboard, draw one new way for whole 5.

DO:
- Display the gallery with 4 and 1 and 3 and 2 already filled.
- Allow 60 seconds for partner whisper and whiteboard sketch.
- Walk and scan.

TEACHER NOTES:
We Do invites students to find a NEW way that is not already on the board. Most will offer 5 and 0 or 0 and 5. Some will repeat one shown. Accept and discuss.

WATCH FOR:
- Pairs offering 5 and 0 - secure.
- Pairs repeating 4 and 1 - prompt: that is one we have. Find a different way.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together. Here are more ways to make 5.
- 5 and 0 - all in one part.
- 0 and 5 - all in the other part.
- We have lots of ways now.

DO:
- Click to reveal the rest of the gallery.
- Repeat the language pattern with the class.

TEACHER NOTES:
Five ways: 0&5, 1&4, 2&3, 3&2, 4&1, 5&0. For Foundation introduction we accept students naming any unique way.

WATCH FOR:
- Students who try to find more - secure.
- Students who needed prompting - small group at table time.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. Look at the counters.
- I am saying: Whole is 5. Parts are 3 and 3.
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Display 5 counters total (3 red, 2 yellow).
- Say the false claim aloud while pointing to the slide.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: Whole is 5. Parts are 3 and 3. Thumbs up if yes, thumbs down if no.
- Scan for: thumbs DOWN. The parts are 3 and 2, not 3 and 3.
PROCEED: If 80% show thumbs down, click to reveal correct version.
PIVOT: Most likely misconception - students do not check the parts add to the whole.
- Reteach: Count the red. Count the yellow. Add them up. Does it match the whole?
- Re-check: Show parts 4 and 1 and ask: is the whole 5?

TEACHER NOTES:
This hinge probes whether students check both parts AND the whole. A no answer means they are tracking both.

WATCH FOR:
- Confident thumbs down - students see the mismatch.
- Slow thumbs - they are unsure; reteach with concrete counters.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- You need 5 cubes and your PPW mat.
- Find at least two different ways to make 5.
- Tell your partner: whole 5, parts ___ and ___.

DO:
- Hand out 5 unifix cubes and a printed PPW mat.
- Circulate. Listen for the language pattern and watch for repeat splits.
- Cold call 2 students to share their two ways.

TEACHER NOTES:
The first systematic try. All correct splits are valid. The goal today is finding more than one way and naming each one.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 4 cubes instead. Find two ways to make 4.
EXTENDING PROMPT:
- Task: Find every way to make 5. Record on your whiteboard.

WATCH FOR:
- Students who find 2 unique splits and use the language - secure.
- Students who keep one split - prompt: now move one cube. What are the new parts?

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard, draw one way to make 5.
- Then write the parts.

DO:
- Allow 60 seconds.
- Collect whiteboards or take a photo.

TEACHER NOTES:
Exit ticket assesses the core target - showing one way to make 5 and naming the parts. Any valid pair (0&5, 1&4, 2&3, 3&2, 4&1, 5&0) is correct.

WATCH FOR:
- Students who show parts that add to 5 - on track.
- Students who only draw 5 dots in one part with no second part - prompt: where is the other part?

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today we found different ways to make 5.
- The whole stayed the same. The parts changed.
- Show me thumbs up if you found two ways.
- Turn and tell: what is one way to make 5?

DO:
- Read the success criteria with the class.
- Use thumbs up sideways down.
- Cold call 1 to 2 students for the partner share.

TEACHER NOTES:
Self-assessment data informs Lesson 7. Most students should reach SC1 today.

WATCH FOR:
- Strong thumbs up - language and skill forming.
- Sideways or down - small group focus.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 6: All the Ways to Make 5",
    `Foundation Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch — Same 5, different parts
  contentSlide(pres, "Launch", C.SUCCESS, "Same 5, Different Parts",
    [
      "5 counters in my cup.",
      "I tip them out.",
      "I move them. Same 5.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      const visualX = lg.rightX + 0.10;
      const visualY = lg.panelTopPadded + 0.15;

      slide.addText("Way 1", {
        x: visualX, y: visualY, w: lg.rightW - 0.20, h: 0.36,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      drawCounterRow(slide, visualX + 0.20, visualY + 0.45, 4, 4,
        { size: 0.50, gap: 0.10 });
      drawCounterRow(slide, visualX + 2.85, visualY + 0.45, 1, 0,
        { size: 0.50, gap: 0.10 });

      slide.addShape("line", {
        x: visualX + 0.20, y: visualY + 1.30, w: lg.rightW - 0.60, h: 0,
        line: { color: C.MUTED || C.CHARCOAL, width: 1, dashType: "dash" },
      });

      slide.addText("Way 2", {
        x: visualX, y: visualY + 1.45, w: lg.rightW - 0.20, h: 0.36,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      drawCounterRow(slide, visualX + 0.30, visualY + 1.95, 2, 2,
        { size: 0.50, gap: 0.10 });
      drawCounterRow(slide, visualX + 1.85, visualY + 1.95, 3, 0,
        { size: 0.50, gap: 0.10 });

      slide.addText("Still 5!", {
        x: visualX, y: visualY + 2.75, w: lg.rightW - 0.20, h: 0.42,
        fontSize: 22, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slides 4-5: Daily Review with reveal — Digit formation: 14
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Write This Number", { color: C.ACCENT });

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

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("14", {
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
      addTextOnShape(slide, "fourteen   1 first, then 4", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 6: Fluency — Counting 1-20 with number track
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Count 1 to 20", { color: C.ACCENT });

  const trackY = CONTENT_TOP + 0.30;
  addNumberTrack(sFluency, 0.5, trackY, 9.0, 1, 10, [], { cellH: 0.85, fontSize: 32 });
  addNumberTrack(sFluency, 0.5, trackY + 1.05, 9.0, 11, 20, [], { cellH: 0.85, fontSize: 32 });

  addCard(sFluency, 1.5, trackY + 2.20, 7.0, 0.85, { strip: C.PRIMARY });
  sFluency.addText("Stand up. Clap on every number.", {
    x: 1.7, y: trackY + 2.22, w: 6.6, h: 0.80,
    fontSize: 24, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 7: LI/SC
  liSlide(pres,
    "We are learning to find different ways to make 5.",
    [
      "I can point to the whole.",
      "I can show one way to make 5.",
      "I can find more than one way to make 5.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 8: I Do 1 — Whole 5, parts 4 and 1
  workedExSlide(pres, 2, "I Do", "One Way to Make 5",
    [
      "I have 5 counters.",
      "",
      "5 is the whole.",
      "",
      "I put 4 here.",
      "I put 1 here.",
      "",
      "Whole 5.",
      "Parts 4 and 1.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 5,
        partA: 4, partB: 1,
        partAColor: "D64545",
        partBColor: "F4C430",
      });
    }
  );

  // Slide 9: I Do 2 — Whole 5, parts 3 and 2
  workedExSlide(pres, 2, "I Do", "Another Way to Make 5",
    [
      "Same 5 counters.",
      "",
      "I move them.",
      "",
      "I put 3 here.",
      "I put 2 here.",
      "",
      "Whole 5.",
      "Parts 3 and 2.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 5,
        partA: 3, partB: 2,
        partAColor: "D64545",
        partBColor: "F4C430",
      });
    }
  );

  // Slides 10-11: CFU 1 with reveal — How many yellow? (1 red, 4 yellow)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How Many Yellow?", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      drawCounterRow(s, 1.6, CONTENT_TOP + 0.45, 5, 1, { size: 1.05, gap: 0.25 });

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
      addTextOnShape(slide, "4 yellow.   Whole 5.   Parts 1 and 4.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: We Do — Find more ways gallery (reveal full gallery)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "What Other Ways?", { color: STAGE_COLORS["3"] });

      // Two known ways shown on left
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText("Already shown:", {
        x: 0.6, y: CONTENT_TOP + 0.10, w: 4.3, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      drawMiniWayCard(s, 0.7, CONTENT_TOP + 0.55, 1.95, 1.20, 4, 1);
      drawMiniWayCard(s, 2.85, CONTENT_TOP + 0.55, 1.95, 1.20, 3, 2);

      s.addText("On your whiteboard:", {
        x: 0.6, y: CONTENT_TOP + 1.95, w: 4.3, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("Draw a NEW way to make 5.", {
        x: 0.6, y: CONTENT_TOP + 2.30, w: 4.3, h: 0.50,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
      });

      // Right card: blank PPW mat for class build
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      drawPpwMat(s, 5.45, CONTENT_TOP + 0.20, 3.8, 3.20, {
        wholeNumber: 5,
        partA: null, partB: null,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      // Reveal: gallery banner showing more ways at bottom
      addCard(slide, 0.5, 4.20, 9.0, 0.85, { strip: C.SUCCESS });
      slide.addText("More ways: 5 and 0   |   0 and 5   |   2 and 3   |   1 and 4", {
        x: 0.6, y: 4.25, w: 8.8, h: 0.78,
        fontSize: 19, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 14-15: CFU Hinge — Is this right? (3 red 2 yellow, claim 3 and 3)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Is This Right?", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.5, 3.0, { strip: C.ALERT });
      drawCounterRow(s, 0.85, CONTENT_TOP + 0.55, 5, 3, { size: 0.75, gap: 0.18 });
      s.addText("These counters", {
        x: 0.6, y: CONTENT_TOP + 1.85, w: 4.3, h: 0.40,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 3.0, { strip: C.ALERT });
      s.addText([
        { text: "I say:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Whole is 5.", options: { fontSize: 26, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "Parts are 3 and 3.", options: { fontSize: 26, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Is that right?", options: { fontSize: 22, italic: true, color: C.ALERT } },
      ], {
        x: 5.4, y: CONTENT_TOP + 0.30, w: 3.9, h: 2.55,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU2);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN!  Parts are 3 and 2.  3 and 3 makes 6.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 16: You Do — Find at least two ways to make 5
  workedExSlide(pres, 4, "You Do", "Find Two Ways to Make 5",
    [
      "First: Get 5 cubes.",
      "",
      "Next: Show one way on the mat.",
      "",
      "Then: Move cubes. Show a different way.",
      "",
      "Tell your partner each way.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.30, { strip: C.ALERT });
      slide.addText("You need:", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.40, h: 0.36,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "5 cubes", options: { bullet: true, fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "your PPW mat", options: { bullet: true, fontSize: 18, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.40, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.50, h: 0.78,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 1.45, lg.rightW, 2.30, {
        wholeNumber: 5,
        partA: null, partB: null,
      });
    }
  );

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Whole 5. Draw one way to show parts. Write the parts.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: what is one way to make 5?",
      scItems: [
        "I can point to the whole.",
        "I can show one way to make 5.",
        "I can find more than one way to make 5.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPWi_Lesson6_All_Ways_Make_5.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // All Ways to Make 5 worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find every way to make 5 with two parts.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addTipBox(doc, "Use 5 counters. Move them. Draw what you see. Say: whole 5, parts ___ and ___.", y, { color: C.ACCENT });

    function drawPdfPpwMat(yTop, wholeLabel) {
      const matX = 70;
      const matW = 470;
      const wholeH = 50;
      const partH = 80;
      doc.lineWidth(2).rect(matX, yTop, matW, wholeH).fillAndStroke("#" + C.PRIMARY, "#333333");
      doc.fillColor("#FFFFFF").fontSize(24).font("Sans-Bold").text("Whole = " + wholeLabel, matX, yTop + 14, { width: matW, align: "center" });
      const partGap = 10;
      const partW = (matW - partGap) / 2;
      doc.lineWidth(2).rect(matX, yTop + wholeH + 8, partW, partH).fillAndStroke("#FFFFFF", "#333333");
      doc.lineWidth(2).rect(matX + partW + partGap, yTop + wholeH + 8, partW, partH).fillAndStroke("#FFFFFF", "#333333");
      doc.fillColor("#333333").fontSize(13).font("Sans-Bold").text("Part", matX + 10, yTop + wholeH + 14);
      doc.fillColor("#333333").fontSize(13).font("Sans-Bold").text("Part", matX + partW + partGap + 10, yTop + wholeH + 14);
      doc.fillColor("#333333").font("Sans");
      return yTop + wholeH + partH + 16;
    }

    y = addSectionHeading(doc, "Way 1", y, { color: C.PRIMARY });
    y = drawPdfPpwMat(y, "5");
    doc.fontSize(13).fillColor("#333333").text("Whole 5.   Parts ____ and ____", 70, y);
    y += 24;

    y = addSectionHeading(doc, "Way 2", y, { color: C.PRIMARY });
    y = drawPdfPpwMat(y, "5");
    doc.fontSize(13).fillColor("#333333").text("Whole 5.   Parts ____ and ____", 70, y);
    y += 24;

    y = addSectionHeading(doc, "Way 3", y, { color: C.PRIMARY });
    y = drawPdfPpwMat(y, "5");
    doc.fontSize(13).fillColor("#333333").text("Whole 5.   Parts ____ and ____", 70, y);
    y += 24;

    addPdfFooter(doc, `Lesson ${SESSION} | Lesson 6 All the Ways to Make 5 | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "All the ways to make 5.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addBodyText(doc, "Children may show any of these splits. Order does not matter for Foundation - 4 and 1 is the same idea as 1 and 4.", y);
    y = addSectionHeading(doc, "Whole = 5", y, { color: C.PRIMARY });
    y = addBodyText(doc, "0 and 5   |   1 and 4   |   2 and 3   |   3 and 2   |   4 and 1   |   5 and 0", y);
    y = addTipBox(doc, "Today's focus is finding MORE THAN ONE way. Most Foundation students will find 2 or 3 ways. That is on track.", y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 6 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
