"use strict";

// Part, Part, Whole Introduction Unit — Lesson 9: One Part is Hiding
// Foundation Numeracy | Lesson 9 of 10 | Variant 0
// VC2MFN04 - partition and combine collections up to 10 using PPW + subitising
// Daily Review: Part/Part Whole - different ways to partition numbers 1-5
// Fluency: Teen Numbers - before and after

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
  STAGE_COLORS,
} = T;

const SESSION = 9;
const TOTAL = 10;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = `Part, Part, Whole | Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`;
const OUT_DIR = "output/PPWi_Lesson9_One_Part_Hiding";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 9 One Part is Hiding", "Find the hidden part. Whole and one part are shown.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 9 Answer Key", "Hidden parts answers.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// PPW mat with optional "hidden" part (a cup-style cover instead of dots).
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
  // Part A box
  slide.addShape("roundRect", {
    x, y: partY, w: partW, h: partH, rectRadius: 0.10,
    fill: { color: C.WHITE },
    line: { color: o.partLine || C.SECONDARY, width: 2.5 },
  });
  if (!o.partAHidden) {
    slide.addText("Part", {
      x: x + 0.10, y: partY + 0.04, w: partW - 0.20, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: o.partLine || C.SECONDARY, bold: true,
      align: "left", valign: "top", margin: 0,
    });
  }

  // Part B box
  slide.addShape("roundRect", {
    x: x + partW + 0.16, y: partY, w: partW, h: partH, rectRadius: 0.10,
    fill: { color: C.WHITE },
    line: { color: o.partLine || C.SECONDARY, width: 2.5 },
  });
  if (!o.partBHidden) {
    slide.addText("Part", {
      x: x + partW + 0.26, y: partY + 0.04, w: partW - 0.20, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: o.partLine || C.SECONDARY, bold: true,
      align: "left", valign: "top", margin: 0,
    });
  }

  function placeDots(boxX, boxY, count, color) {
    if (!count) return;
    const dotSize = 0.32;
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

  function placeCover(boxX, boxY) {
    // A cup-style overlay: rounded rect with question mark
    const coverW = partW - 0.40;
    const coverH = partH - 0.50;
    const coverX = boxX + 0.20;
    const coverY = boxY + 0.30;
    slide.addShape("roundRect", {
      x: coverX, y: coverY, w: coverW, h: coverH, rectRadius: 0.20,
      fill: { color: C.ALERT },
      line: { color: C.CHARCOAL, width: 1.5 },
    });
    slide.addText("?", {
      x: coverX, y: coverY, w: coverW, h: coverH,
      fontSize: 60, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }

  // Render parts: number, hidden, or null
  if (o.partAHidden) {
    placeCover(x, partY);
  } else if (o.partA != null) {
    placeDots(x, partY, o.partA, o.partAColor || "D64545");
  }
  if (o.partBHidden) {
    placeCover(x + partW + 0.16, partY);
  } else if (o.partB != null) {
    placeDots(x + partW + 0.16, partY, o.partB, o.partBColor || "F4C430");
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
- Welcome back to part, part, whole.
- This week we made 5, 6, and 7 in lots of ways.
- Today is something new. One part will hide.
- We will use what we know to find the hidden part.

DO:
- Have 5 counters and a small paper cup ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 9 of 10. New step today: missing part. The whole and one part are shown. Students must find the other part. Keep numbers small (3 to 5). Cup or paper cover is the hiding tool.

WATCH FOR:
- Students who look excited - they enjoy mystery tasks.
- Students who shrug - we will model carefully.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are the materials for today.
- We will use the One Part is Hiding sheet at the table.

DO:
- Print the One Part is Hiding sheet (one per child).
- Have 5 unifix cubes per child and a small paper cup.

TEACHER NOTES:
The cup is the hiding tool. The sheet shows the whole and one part for several wholes. Students find the hidden part.

[General: Resources]`;

const NOTES_LAUNCH = `SAY:
- Watch carefully. I have 5 counters. The whole is 5.
- I put 3 counters here. I see 3.
- Now watch. I cover the rest with my cup.
- The whole is still 5. I can see 3.
- How many are hiding under the cup?

DO:
- Show 5 counters on the floor mat.
- Move 3 to one side. Cover the other 2 with a paper cup.
- Pause. Let students think.
- Lift the cup to reveal 2.
- Repeat the language: 5 in all. 3 I can see. 2 hiding.

TEACHER NOTES:
This launch introduces the new idea: when one part is hidden we can use the whole and the visible part to find the hidden part. Keep numbers small (5 or less today). 60 to 90 seconds.

WATCH FOR:
- Students who shout 2 before you lift the cup - they used the whole and seen part.
- Students who recount when you lift the cup - that is fine for now.

[General: Launch | VTLM 2.0: Activate Prior Knowledge]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the whole on the slide.
- Look at one part.
- On your whiteboard, draw or write the other part.

DO:
- Display whole 4, one part shown as 1.
- Allow 30 seconds.
- Walk and check.

TEACHER NOTES:
Daily Review: partitioning 1 to 5. This DR previews today's I Do — different ways to partition small numbers, with one part shown.

WATCH FOR:
- Students who write 3 - secure.
- Students who write 4 (the whole) - reteach: 4 is the WHOLE, 1 is one PART.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check together.
- Whole 4. One part is 1. The other part is 3.
- Whole 4. Parts 1 and 3.

DO:
- Click to reveal the answer.
- Choral say whole 4, parts 1 and 3.

TEACHER NOTES:
Tick and fix. Some students will need a number track or counters.

WATCH FOR:
- Students who said 3 - on track.
- Students who said another number - small group at table time.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Look at the cards on the slide.
- I will say a number. You say what comes BEFORE and AFTER.
- For 14: 13, 14, 15.

DO:
- Display the prompt with 14 in the middle.
- Choral practise: 13 14 15. Then try 16, then 11.
- Whisper to partner: what comes before 17? after 19?

TEACHER NOTES:
Fluency: before and after. Builds awareness of order.

WATCH FOR:
- Students who say the right neighbours - secure.
- Students who say a random number - prompt: count along the number track.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to find the part that is hiding.
- Read the success criteria with me.

DO:
- Choral read the LI.
- Choral read each criterion.
- Hold up a cup as you say hiding.

TEACHER NOTES:
SC1: name the whole. SC2: find the hidden part for whole 4 or 5. SC3: explain how you knew.

WATCH FOR:
- Students who chime in confidently - language secure.
- Students who look quiet - I Do will rebuild the picture.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. The whole is 4.
- I see 3 counters in this part.
- The other part is hiding under the cup.
- I think about it: 3 and how many makes 4?
- The hidden part is 1.

DO:
- Show the PPW mat: whole 4, part shown 3, part hidden under cup.
- Think aloud slowly.
- Lift the cup at the end to reveal 1.
- Repeat: whole 4, part 3, hidden part 1.

TEACHER NOTES:
First exposure to missing-part. The whole is the total. The seen part is what we can see. The hidden part is what is left. Use fingers or a number track if students need extra support.

MISCONCEPTIONS:
- Misconception: Students count the visible part as the answer.
  Why: They have learned to count what they see.
  Impact: They give 3 instead of 1.
  Quick correction: Point to the cup. The hidden part is what is missing from the whole.

WATCH FOR:
- Students who say 1 before lifting the cup - secure.
- Students who say 3 - reteach: 3 is what we SEE. The hidden part makes the whole 4.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Watch this one. The whole is 5.
- I see 2 counters.
- The other part is hiding.
- I think: 2 and how many makes 5?
- The hidden part is 3.

DO:
- Show the PPW mat: whole 5, part shown 2, part hidden under cup.
- Think aloud slowly.
- Lift the cup at the end to reveal 3.

TEACHER NOTES:
Second exposure. Same routine, slightly bigger whole. Move slowly. Foundation students will need many examples.

WATCH FOR:
- Students who say 3 confidently - secure.
- Students who say 5 (the whole) - reteach: 5 is the whole, 2 is one part.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Look at the picture.
- The whole is 4. I see 1 counter.
- Show me on your fingers how many are hiding.

DO:
- Display: whole 4, one part visible as 1, other part hidden.
- Wait 5 seconds.
- Say: Show me on your fingers... ready... show.

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: How many hiding? Show me on your fingers... show.
- Scan for: 3 fingers held up.
PROCEED: If 80% show 3 fingers, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count the seen part again.
- Reteach: 1 is what we SEE. The whole is 4. What makes 4? 1 and 3.
- Re-check: Whole 4, see 2, how many hiding?

TEACHER NOTES:
Fingers are universal at Foundation. Quick, every-student response.

WATCH FOR:
- 3 fingers held up - secure with missing part.
- 1 finger - they returned the visible part. Reteach.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Look at the slide.
- Whole 5. One part is 4. One part is hiding.
- With your partner, whisper: how many are hiding?
- Show me on your fingers when I count to 3.

DO:
- Display: whole 5, part shown 4, part hidden.
- Allow 20 seconds for whisper.
- Count to 3, then say show.

TEACHER NOTES:
We Do mirrors I Do. Foundation students may need to count up from 4: 4, 5. That's 1 more. The hidden part is 1.

WATCH FOR:
- Pairs offering 1 - secure.
- Pairs offering 4 - they returned the seen part. Prompt: 4 is what we SEE. How many MORE for 5?

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together. Whole 5. Part 4. Hidden part is 1.
- 4 and 1 makes 5.

DO:
- Click to reveal.
- Choral say whole 5, parts 4 and 1.

TEACHER NOTES:
Strong moment to lock the language pattern.

WATCH FOR:
- Students who said 1 first - confident.
- Students who needed help - small group at table time.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. Look at the picture.
- Whole is 5. One part shown is 3.
- I am saying: the hidden part is 4.
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Display whole 5, one part shown 3, other part hidden.
- Say the false claim aloud while pointing to the slide.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: Whole 5, part 3, hidden part 4. Thumbs up if yes, thumbs down if no.
- Scan for: thumbs DOWN. The hidden part should be 2, not 4.
PROCEED: If 80% show thumbs down, click to reveal correct version.
PIVOT: Most likely misconception - students agree without checking.
- Reteach: 3 and 4 makes 7, not 5. Use fingers: 3 fingers up, then keep going to 5. 3, 4, 5. That's 2 more.
- Re-check: Whole 5, part 1, hidden part?

TEACHER NOTES:
This hinge probes whether students check the parts make the whole. A no answer means they are checking.

WATCH FOR:
- Confident thumbs down - students see the mismatch.
- Slow thumbs - they are unsure; reteach with concrete counters.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- You need 5 cubes, your PPW mat, and a paper cup.
- Your partner makes a whole with the cubes.
- Your partner covers some with the cup.
- You find the hidden part. Then swap.

DO:
- Hand out 5 cubes, a PPW mat, and a paper cup per pair.
- Demonstrate one round at the front.
- Circulate. Listen for the language pattern.

TEACHER NOTES:
The first solo go. Pairs take turns hiding and finding. Use wholes 3 to 5.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 3 cubes only. Whole is 3.
EXTENDING PROMPT:
- Task: Try with whole 6. Use 6 cubes.

WATCH FOR:
- Students who use the whole-and-seen logic - secure.
- Students who lift the cup before guessing - prompt: think first, lift after.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. Look at the slide.
- Whole 4. One part is 2. How many are hiding?
- Write the answer on your whiteboard.

DO:
- Display whole 4, part 2, part hidden.
- Allow 60 seconds.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses the core target - finding the hidden part. Answer is 2.

WATCH FOR:
- Students who write 2 - on track.
- Students who write 4 (the whole) or 2 (returned visible) - small group support.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today we found the hidden part.
- We used the whole and the part we could see.
- Show me thumbs up if you found a hidden part today.
- Turn and tell: what helped you find the hidden part?

DO:
- Read the success criteria with the class.
- Use thumbs up sideways down.
- Cold call 1 to 2 students for the partner share.

TEACHER NOTES:
Self-assessment data informs Lesson 10 grouping.

WATCH FOR:
- Strong thumbs up - language and skill forming.
- Sideways or down - small group focus.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 9: One Part is Hiding",
    `Foundation Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch — Cup hides part of 5
  contentSlide(pres, "Launch", C.SUCCESS, "Where is the Other Part?",
    [
      "Whole is 5.",
      "I see 3 counters.",
      "Some are hiding!",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      const visualX = lg.rightX + 0.10;
      const visualY = lg.panelTopPadded + 0.15;

      // Whole label
      slide.addText("Whole = 5", {
        x: visualX, y: visualY, w: lg.rightW - 0.20, h: 0.42,
        fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Three visible counters
      slide.addText("I see:", {
        x: visualX, y: visualY + 0.55, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
      });
      drawCounterRow(slide, visualX + 0.50, visualY + 0.95, 3, 3,
        { size: 0.50, gap: 0.12 });

      // Cup with question mark
      slide.addText("Hiding:", {
        x: visualX, y: visualY + 1.75, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addShape("roundRect", {
        x: visualX + 1.05, y: visualY + 2.10, w: 1.30, h: 0.95, rectRadius: 0.30,
        fill: { color: C.ALERT },
        line: { color: C.CHARCOAL, width: 1.5 },
      });
      slide.addText("?", {
        x: visualX + 1.05, y: visualY + 2.10, w: 1.30, h: 0.95,
        fontSize: 56, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slides 4-5: Daily Review with reveal — Partition 4: whole 4, part 1, hidden ?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "What is the Other Part?", { color: C.ACCENT });

      // Left card: instructions
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 24, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Whole 4.", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "One part is 1.", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "What is the other part?", options: { fontSize: 22, color: C.CHARCOAL, bold: true } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right card: PPW mat with one part shown, other hidden
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      drawPpwMat(s, 5.45, CONTENT_TOP + 0.20, 3.8, 3.20, {
        wholeNumber: 4,
        partA: 1,
        partBHidden: true,
        partAColor: "D64545",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "The other part is 3.   Whole 4. Parts 1 and 3.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 6: Fluency — Before and after teen numbers
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Before and After", { color: C.ACCENT });

  // Three chips: ?, 14, ?
  const chipW = 1.65;
  const chipH = 1.65;
  const chipY = CONTENT_TOP + 0.50;
  const gap = 0.30;
  const totalW = 3 * chipW + 2 * gap;
  const startX = (10 - totalW) / 2;

  // Before chip - ?
  sFluency.addShape("roundRect", {
    x: startX, y: chipY, w: chipW, h: chipH, rectRadius: 0.15,
    fill: { color: C.WHITE },
    line: { color: C.PRIMARY, width: 3, dashType: "dash" },
  });
  sFluency.addText("?", {
    x: startX, y: chipY, w: chipW, h: chipH,
    fontSize: 70, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  sFluency.addText("before", {
    x: startX, y: chipY + chipH + 0.10, w: chipW, h: 0.35,
    fontSize: 18, fontFace: FONT_B, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Middle chip - 14
  drawNumChip(sFluency, startX + chipW + gap, chipY, chipW, chipH, 14, C.PRIMARY);
  sFluency.addText("today's number", {
    x: startX + chipW + gap, y: chipY + chipH + 0.10, w: chipW, h: 0.35,
    fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
    align: "center", valign: "middle", margin: 0,
  });

  // After chip - ?
  sFluency.addShape("roundRect", {
    x: startX + 2 * (chipW + gap), y: chipY, w: chipW, h: chipH, rectRadius: 0.15,
    fill: { color: C.WHITE },
    line: { color: C.PRIMARY, width: 3, dashType: "dash" },
  });
  sFluency.addText("?", {
    x: startX + 2 * (chipW + gap), y: chipY, w: chipW, h: chipH,
    fontSize: 70, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  sFluency.addText("after", {
    x: startX + 2 * (chipW + gap), y: chipY + chipH + 0.10, w: chipW, h: 0.35,
    fontSize: 18, fontFace: FONT_B, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addCard(sFluency, 1.0, chipY + chipH + 0.65, 8.0, 0.55, { strip: C.SECONDARY });
  sFluency.addText("Whisper the missing numbers.", {
    x: 1.2, y: chipY + chipH + 0.66, w: 7.6, h: 0.50,
    fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 7: LI/SC
  liSlide(pres,
    "We are learning to find the part that is hiding.",
    [
      "I can point to the whole.",
      "I can find the hidden part.",
      "I can say how I knew.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 8: I Do 1 — Whole 4, part 3, hidden 1
  workedExSlide(pres, 2, "I Do", "Find the Hidden Part",
    [
      "Whole is 4.",
      "",
      "I see 3 counters.",
      "",
      "I think:",
      "3 and how many makes 4?",
      "",
      "The hidden part is 1.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 4,
        partA: 3,
        partBHidden: true,
        partAColor: "D64545",
      });
    }
  );

  // Slide 9: I Do 2 — Whole 5, part 2, hidden 3
  workedExSlide(pres, 2, "I Do", "Find the Hidden Part",
    [
      "Whole is 5.",
      "",
      "I see 2 counters.",
      "",
      "I think:",
      "2 and how many makes 5?",
      "",
      "The hidden part is 3.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 5,
        partA: 2,
        partBHidden: true,
        partAColor: "D64545",
      });
    }
  );

  // Slides 10-11: CFU 1 — Whole 4, see 1, hidden ?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How Many Are Hiding?", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      // Centred PPW mat: whole 4, see 1, hide other
      addCard(s, 2.3, CONTENT_TOP, 5.4, SAFE_BOTTOM - CONTENT_TOP - 0.05, { strip: C.ALERT });
      drawPpwMat(s, 2.55, CONTENT_TOP + 0.20, 4.9, 3.20, {
        wholeNumber: 4,
        partA: 1,
        partBHidden: true,
        partAColor: "D64545",
      });

      // Prompt
      s.addText("Show me on your fingers.", {
        x: 0.5, y: CONTENT_TOP + 3.55, w: 9.0, h: 0.40,
        fontSize: 22, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "3 hiding.   Whole 4. Parts 1 and 3.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: We Do — Whole 5, see 4, hidden ?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "How Many Are Hiding?", { color: STAGE_COLORS["3"] });

      // Left card: PPW mat
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      drawPpwMat(s, 0.65, CONTENT_TOP + 0.20, 4.2, 3.20, {
        wholeNumber: 5,
        partA: 4,
        partBHidden: true,
        partAColor: "D64545",
      });

      // Right card: prompt
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Whole = 5", options: { fontSize: 24, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "I see 4.", options: { fontSize: 24, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Hiding = ____", options: { fontSize: 24, bold: true, color: C.ALERT } },
      ], {
        x: 5.4, y: CONTENT_TOP + 0.30, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.55,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1 hiding.   Whole 5. Parts 4 and 1.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 14-15: CFU Hinge — claim hidden part is 4 when correct is 2 (whole 5, see 3)
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
      drawPpwMat(s, 0.65, CONTENT_TOP + 0.10, 4.2, 2.85, {
        wholeNumber: 5,
        partA: 3,
        partBHidden: true,
        partAColor: "D64545",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 3.0, { strip: C.ALERT });
      s.addText([
        { text: "I say:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Whole 5. See 3.", options: { fontSize: 22, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "Hiding is 4.", options: { fontSize: 26, bold: true, color: C.CHARCOAL, breakLine: true } },
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
      addTextOnShape(slide, "Thumbs DOWN!  Hiding is 2.  3 and 4 makes 7, not 5.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 16: You Do — Pair hide and find
  workedExSlide(pres, 4, "You Do", "Hide and Find",
    [
      "First: Get 5 cubes and a cup.",
      "",
      "Next: Partner makes a whole.",
      "Partner covers some with the cup.",
      "",
      "Then: You find the hidden part.",
      "",
      "Swap and try again.",
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
        { text: "a paper cup", options: { bullet: true, fontSize: 18, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.40, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.50, h: 0.78,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 1.45, lg.rightW, 2.30, {
        wholeNumber: 5,
        partA: null,
        partBHidden: true,
      });
    }
  );

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Whole 4. I see 2 counters. How many are hiding? Write the answer.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: what helped you find the hidden part?",
      scItems: [
        "I can point to the whole.",
        "I can find the hidden part.",
        "I can say how I knew.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPWi_Lesson9_One_Part_Hiding.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find the hidden part. The whole and one part are shown.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addTipBox(doc, "The whole is on top. One part has dots. The other part has a question mark. Write the hidden part.", y, { color: C.ACCENT });

    function drawPdfPpwMatHidden(yTop, wholeLabel, seenCount) {
      const matX = 70;
      const matW = 470;
      const wholeH = 50;
      const partH = 90;
      doc.lineWidth(2).rect(matX, yTop, matW, wholeH).fillAndStroke("#" + C.PRIMARY, "#333333");
      doc.fillColor("#FFFFFF").fontSize(24).font("Sans-Bold").text("Whole = " + wholeLabel, matX, yTop + 14, { width: matW, align: "center" });
      const partGap = 10;
      const partW = (matW - partGap) / 2;
      // Visible part with dots
      doc.lineWidth(2).rect(matX, yTop + wholeH + 8, partW, partH).fillAndStroke("#FFFFFF", "#333333");
      // Draw dots
      const dotSize = 14;
      const dotGap = 8;
      const dotsTotalW = seenCount * dotSize + (seenCount - 1) * dotGap;
      const dotsStartX = matX + (partW - dotsTotalW) / 2;
      const dotsCY = yTop + wholeH + 8 + partH / 2 - dotSize / 2;
      for (let i = 0; i < seenCount; i += 1) {
        doc.circle(dotsStartX + i * (dotSize + dotGap) + dotSize / 2, dotsCY + dotSize / 2, dotSize / 2)
           .fillAndStroke("#D64545", "#333333");
      }
      // Hidden part with question mark
      doc.lineWidth(2).rect(matX + partW + partGap, yTop + wholeH + 8, partW, partH).fillAndStroke("#FFFFFF", "#333333");
      doc.fillColor("#" + C.ALERT).fontSize(60).font("Sans-Bold").text("?", matX + partW + partGap, yTop + wholeH + 8 + 12, { width: partW, align: "center" });
      doc.fillColor("#333333").font("Sans");
      return yTop + wholeH + partH + 16;
    }

    y = addSectionHeading(doc, "Whole = 3", y, { color: C.PRIMARY });
    y = drawPdfPpwMatHidden(y, "3", 1);
    doc.fontSize(13).fillColor("#333333").text("I see 1.   Hidden part = ____", 70, y);
    y += 22;

    y = addSectionHeading(doc, "Whole = 4", y, { color: C.PRIMARY });
    y = drawPdfPpwMatHidden(y, "4", 3);
    doc.fontSize(13).fillColor("#333333").text("I see 3.   Hidden part = ____", 70, y);
    y += 22;

    y = addSectionHeading(doc, "Whole = 5", y, { color: C.PRIMARY });
    y = drawPdfPpwMatHidden(y, "5", 2);
    doc.fontSize(13).fillColor("#333333").text("I see 2.   Hidden part = ____", 70, y);
    y += 22;

    addPdfFooter(doc, `Lesson ${SESSION} | Lesson 9 One Part is Hiding | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Hidden part answers.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addBodyText(doc, "Each answer is the part that makes the whole when added to what is seen.", y);
    y = addSectionHeading(doc, "Answers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Whole 3. See 1. Hidden = 2.", y);
    y = addBodyText(doc, "Whole 4. See 3. Hidden = 1.", y);
    y = addBodyText(doc, "Whole 5. See 2. Hidden = 3.", y);
    y = addTipBox(doc, "Today's focus is using the whole and the seen part to find the hidden part. Foundation students may count up from the seen number.", y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 9 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
