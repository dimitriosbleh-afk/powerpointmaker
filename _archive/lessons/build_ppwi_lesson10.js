"use strict";

// Part, Part, Whole Introduction Unit — Lesson 10: Show What We Know
// Foundation Numeracy | Lesson 10 of 10 | Variant 0
// VC2MFN04 - partition and combine collections up to 10 using PPW + subitising
// Daily Review: Part/Part Whole - different ways to partition numbers 6-10
// Fluency: Counting to 20 from various starting points

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

const SESSION = 10;
const TOTAL = 10;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = `Part, Part, Whole | Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`;
const OUT_DIR = "output/PPWi_Lesson10_Show_What_We_Know";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 10 Show What We Know", "Show ways to make 8 and find a hidden part.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 10 Answer Key", "Sample answers for the celebration sheet.");

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
  if (!o.partAHidden) {
    slide.addText("Part", {
      x: x + 0.10, y: partY + 0.04, w: partW - 0.20, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: o.partLine || C.SECONDARY, bold: true,
      align: "left", valign: "top", margin: 0,
    });
  }

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
    const dotSize = 0.26;
    const gap = 0.06;
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
      fontSize: 50, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }

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

function drawMiniWayCard(slide, x, y, w, h, partA, partB, opts) {
  const o = opts || {};
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.10,
    fill: { color: C.WHITE },
    line: { color: o.borderColor || C.SECONDARY, width: 2 },
  });
  slide.addText(`${partA} and ${partB}`, {
    x, y: y + 0.05, w, h: 0.40,
    fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  const total = partA + partB;
  const dotSize = 0.18;
  const gap = 0.04;
  const totalW = total * dotSize + (total - 1) * gap;
  const startX = x + (w - totalW) / 2;
  const cy = y + 0.58;
  for (let i = 0; i < total; i += 1) {
    slide.addShape("roundRect", {
      x: startX + i * (dotSize + gap),
      y: cy,
      w: dotSize, h: dotSize,
      rectRadius: dotSize / 2,
      fill: { color: i < partA ? "D64545" : "F4C430" },
      line: { color: C.CHARCOAL, width: 0.6 },
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
- Welcome to our last Part, Part, Whole lesson.
- Today is a celebration. We will show what we know.
- We will find ways to make 8. We will find a hidden part.

DO:
- Have 8 counters and a small paper cup ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 10 of 10. Today consolidates the unit: language (whole, parts), finding ways to make a number, and finding a hidden part. Use celebratory tone.

WATCH FOR:
- Students who are excited to share what they know - good.
- Students who are unsure - sit them near a partner who is confident.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are the materials for today.
- We will use the Show What We Know sheet at the table.

DO:
- Print the Show What We Know sheet (one per child).
- Have 8 unifix cubes per child, a PPW mat, and a paper cup.

TEACHER NOTES:
The celebration sheet revisits both skills from the unit: ways to make a whole, and finding a hidden part.

[General: Resources]`;

const NOTES_LAUNCH = `SAY:
- Look at all the things we have learned.
- We learned the words whole and parts.
- We made 5, 6, 7, 8, 9 and even 10 in lots of ways.
- We found a hidden part.
- Today we put it all together.

DO:
- Show the gallery on the slide.
- Point to each card as you name it.
- Hold up 8 counters at the end.

TEACHER NOTES:
This launch is a quick recap of the whole unit. Keep it under 90 seconds. The visual reminds students how much they have learned.

WATCH FOR:
- Students who say something we did - secure with the routine.
- Students who recall the language - language secure.

[General: Launch | VTLM 2.0: Activate Prior Knowledge]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the whole on the slide.
- Look at one part.
- On your whiteboard, write the other part.

DO:
- Display whole 8, one part 5.
- Allow 30 seconds.
- Walk and check.

TEACHER NOTES:
Daily Review: partitioning 6 to 10. Today is whole 8, part 5. Students should write 3.

WATCH FOR:
- Students who write 3 - secure.
- Students who write 8 (the whole) or 5 (the seen part) - reteach with counters.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check together.
- Whole 8. One part is 5. The other part is 3.
- Whole 8. Parts 5 and 3.

DO:
- Click to reveal the answer.
- Choral say whole 8, parts 5 and 3.

TEACHER NOTES:
Tick and fix.

WATCH FOR:
- Students who said 3 - on track.
- Students who said another number - small group at table time.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Stand up.
- We will start counting from a different number each time.
- I will say start at 5. You count from 5 up to 20.
- Ready... 5, 6, 7...

DO:
- Lead choral count from 5 to 20.
- Then start at 12 and count to 20.
- Then start at 8 and count to 20.

TEACHER NOTES:
Fluency: counting from various starting points. Builds flexibility.

WATCH FOR:
- Students who start counting from 1 instead of the start number - common error.
- Students who count smoothly from any start - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are showing what we know about parts and wholes.
- Read the success criteria with me.

DO:
- Choral read the LI.
- Choral read each criterion.
- Hold up a PPW mat as you say show what we know.

TEACHER NOTES:
SC1: name the whole. SC2: show a way to make a number. SC3: find a hidden part.

WATCH FOR:
- Students who chime in confidently - language secure.
- Students who look quiet - I Do will rebuild the picture.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. The whole is 8.
- I have 8 counters.
- I put 5 here. I put 3 here.
- Whole 8. Parts 5 and 3.
- Say with me: whole 8, parts 5 and 3.

DO:
- Show the PPW mat with 8 counters.
- Move 5 to the left part, 3 to the right.
- Repeat the language pattern at least 3 times.

TEACHER NOTES:
Reminder of the routine: whole, parts. Whole 8 may be a stretch for some students. Use fingers and counters together.

WATCH FOR:
- Students who echo whole 8, parts 5 and 3 - language secure.
- Students who count all 8 dots without separating - re-show with colour difference.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now watch this. The whole is 8.
- I see 6 counters.
- The other part is hiding.
- I think: 6 and how many makes 8?
- The hidden part is 2.

DO:
- Show the PPW mat: whole 8, part 6, other part covered.
- Think aloud slowly.
- Lift the cup at the end to reveal 2.

TEACHER NOTES:
Reminder of the missing part routine from Lesson 9. Move slowly.

WATCH FOR:
- Students who say 2 confidently - secure.
- Students who say 8 (the whole) or 6 (the seen part) - reteach.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Look at the picture.
- The whole is 8. I see 4 counters.
- Show me on your fingers how many are hiding.

DO:
- Display: whole 8, one part visible as 4, other part hidden.
- Wait 5 seconds.
- Say: Show me on your fingers... ready... show.

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: How many hiding? Show me on your fingers... show.
- Scan for: 4 fingers held up.
PROCEED: If 80% show 4 fingers, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count the seen part again.
- Reteach: 4 is what we SEE. The whole is 8. What makes 8? 4 and 4.
- Re-check: Whole 8, see 5, how many hiding?

TEACHER NOTES:
Fingers are universal at Foundation. Quick, every-student response.

WATCH FOR:
- 4 fingers held up - secure with missing part.
- 4 (returning the seen) is also confusable here. Listen to language.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Look at the slide.
- The whole is 7. With your partner, find two ways to make 7.
- Draw both ways on your whiteboard.

DO:
- Display the prompt: whole 7, find two ways.
- Allow 90 seconds.
- Walk and scan.

TEACHER NOTES:
We Do is a celebration of finding ways. Most pairs will offer 6 and 1, 5 and 2, 4 and 3, or 7 and 0.

WATCH FOR:
- Pairs offering 2 unique splits - secure.
- Pairs offering only one - prompt: now move one cube. Show a different way.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together. Here are some ways to make 7.
- 6 and 1. 5 and 2. 4 and 3. 7 and 0.
- Did you get any of these?

DO:
- Click to reveal the gallery.
- Cold call 2 students to share their two ways.

TEACHER NOTES:
Strong moment to celebrate. Students see their ways match.

WATCH FOR:
- Students who match their ways to the gallery - secure.
- Students who tried something different - validate any valid pair.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. Look at the picture.
- Whole is 7. One part shown is 4.
- I am saying: the hidden part is 4.
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Display whole 7, one part 4, other part hidden.
- Say the false claim aloud while pointing to the slide.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: Whole 7, part 4, hidden part 4. Thumbs up if yes, thumbs down if no.
- Scan for: thumbs DOWN. The hidden part should be 3, not 4.
PROCEED: If 80% show thumbs down, click to reveal correct version.
PIVOT: Most likely misconception - students see one part as the hidden too.
- Reteach: 4 and 4 makes 8, not 7. Use fingers: 4 fingers up, count to 7. 4, 5, 6, 7. That is 3 more.
- Re-check: Whole 7, part 5, hidden part?

TEACHER NOTES:
This hinge probes whether students check the parts make the whole.

WATCH FOR:
- Confident thumbs down - students see the mismatch.
- Slow thumbs - they are unsure; reteach with concrete counters.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- You need 8 cubes, your PPW mat, and a paper cup.
- Find two ways to make 8.
- Then have your partner cover one part. Find the hidden part.

DO:
- Hand out 8 cubes, a PPW mat, and a paper cup per pair.
- Demonstrate one round at the front.
- Circulate. Listen for the language pattern.

TEACHER NOTES:
The celebration solo go. Students show both skills. Wholes 6 to 10 are all valid.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 5 cubes. Find one way to make 5.
EXTENDING PROMPT:
- Task: Find 3 ways to make 8 and have your partner hide one part for each.

WATCH FOR:
- Students using both skills - language and concept secure.
- Students stuck on one skill - prompt with cubes.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. Look at the slide.
- The whole is 6. I see 4 counters.
- How many are hiding? Write the answer.

DO:
- Display whole 6, part 4, part hidden.
- Allow 60 seconds.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses the unit's core target. Answer is 2.

WATCH FOR:
- Students who write 2 - on track.
- Students who write 6 or 4 - small group support next term.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- We finished our part, part, whole unit!
- We learned the words whole and parts.
- We made numbers in lots of ways.
- We found hidden parts.
- Show me thumbs up for what we did today.

DO:
- Read the success criteria with the class.
- Use thumbs up sideways down.
- Cold call 1 to 2 students to share what they remember most.
- Celebrate with a clap.

TEACHER NOTES:
Self-assessment data informs Term planning. Strong end-of-unit reflection.

WATCH FOR:
- Strong thumbs up - language and skill secure.
- Sideways or down - small group focus.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 10: Show What We Know",
    `Foundation Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch — Look how much we have learned!
  contentSlide(pres, "Launch", C.SUCCESS, "Look What We Learned!",
    [
      "Whole and parts.",
      "Lots of ways to make a number.",
      "How to find a hidden part.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      const visualX = lg.rightX + 0.10;
      const visualY = lg.panelTopPadded + 0.10;

      // Three small concept cards stacked
      // Card 1: whole and parts
      slide.addShape("roundRect", {
        x: visualX, y: visualY, w: lg.rightW - 0.20, h: 0.95, rectRadius: 0.10,
        fill: { color: C.WHITE },
        line: { color: C.PRIMARY, width: 2 },
      });
      slide.addText("Whole and parts", {
        x: visualX + 0.10, y: visualY + 0.10, w: lg.rightW - 0.40, h: 0.40,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "left", valign: "middle", margin: 0,
      });
      drawCounterRow(slide, visualX + 0.20, visualY + 0.55, 4, 2,
        { size: 0.30, gap: 0.06 });

      // Card 2: ways to make
      slide.addShape("roundRect", {
        x: visualX, y: visualY + 1.15, w: lg.rightW - 0.20, h: 0.95, rectRadius: 0.10,
        fill: { color: C.WHITE },
        line: { color: C.SECONDARY, width: 2 },
      });
      slide.addText("Ways to make a number", {
        x: visualX + 0.10, y: visualY + 1.25, w: lg.rightW - 0.40, h: 0.40,
        fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "left", valign: "middle", margin: 0,
      });
      slide.addText("3 + 2.   4 + 1.   2 + 3.", {
        x: visualX + 0.20, y: visualY + 1.65, w: lg.rightW - 0.40, h: 0.35,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "left", valign: "middle", margin: 0,
      });

      // Card 3: hidden part
      slide.addShape("roundRect", {
        x: visualX, y: visualY + 2.30, w: lg.rightW - 0.20, h: 0.95, rectRadius: 0.10,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 2 },
      });
      slide.addText("Find the hidden part", {
        x: visualX + 0.10, y: visualY + 2.40, w: lg.rightW - 0.40, h: 0.40,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "left", valign: "middle", margin: 0,
      });
      // Tiny cup icon
      slide.addShape("roundRect", {
        x: visualX + 0.20, y: visualY + 2.78, w: 0.50, h: 0.40, rectRadius: 0.12,
        fill: { color: C.ALERT },
        line: { color: C.CHARCOAL, width: 1 },
      });
      slide.addText("?", {
        x: visualX + 0.20, y: visualY + 2.78, w: 0.50, h: 0.40,
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slides 4-5: Daily Review with reveal — Partition 8: whole 8, part 5, hidden ?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "What is the Other Part?", { color: C.ACCENT });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 24, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Whole 8.", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "One part is 5.", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "What is the other part?", options: { fontSize: 22, color: C.CHARCOAL, bold: true } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      drawPpwMat(s, 5.45, CONTENT_TOP + 0.20, 3.8, 3.20, {
        wholeNumber: 8,
        partA: 5,
        partBHidden: true,
        partAColor: "D64545",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "The other part is 3.   Whole 8. Parts 5 and 3.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 6: Fluency — Counting from various starts
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Start Here. Count to 20.", { color: C.ACCENT });

  // Three start chips with arrows
  const starts = [5, 12, 8];
  const chipW = 1.65;
  const chipH = 1.65;
  const chipY = CONTENT_TOP + 0.50;
  const cgap = 0.50;
  const totalW = 3 * chipW + 2 * cgap;
  const startX = (10 - totalW) / 2;

  for (let i = 0; i < starts.length; i += 1) {
    const cx = startX + i * (chipW + cgap);
    drawNumChip(sFluency, cx, chipY, chipW, chipH, starts[i], C.PRIMARY);
    sFluency.addText("count to 20", {
      x: cx, y: chipY + chipH + 0.10, w: chipW, h: 0.40,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "middle", margin: 0,
    });
  }

  addCard(sFluency, 1.0, chipY + chipH + 0.65, 8.0, 0.55, { strip: C.SECONDARY });
  sFluency.addText("Stand up. Start at each number. Count to 20.", {
    x: 1.2, y: chipY + chipH + 0.66, w: 7.6, h: 0.50,
    fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 7: LI/SC
  liSlide(pres,
    "We are showing what we know about parts and wholes.",
    [
      "I can point to the whole.",
      "I can show a way to make a number.",
      "I can find a hidden part.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 8: I Do 1 — Whole 8, parts 5 and 3
  workedExSlide(pres, 2, "I Do", "One Way to Make 8",
    [
      "I have 8 counters.",
      "",
      "8 is the whole.",
      "",
      "I put 5 here.",
      "I put 3 here.",
      "",
      "Whole 8.",
      "Parts 5 and 3.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 8,
        partA: 5, partB: 3,
        partAColor: "D64545",
        partBColor: "F4C430",
      });
    }
  );

  // Slide 9: I Do 2 — Whole 8, part 6, hidden 2
  workedExSlide(pres, 2, "I Do", "Find the Hidden Part",
    [
      "Whole is 8.",
      "",
      "I see 6 counters.",
      "",
      "I think:",
      "6 and how many makes 8?",
      "",
      "The hidden part is 2.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 8,
        partA: 6,
        partBHidden: true,
        partAColor: "D64545",
      });
    }
  );

  // Slides 10-11: CFU 1 — Whole 8, see 4, hidden ?
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

      addCard(s, 2.3, CONTENT_TOP, 5.4, SAFE_BOTTOM - CONTENT_TOP - 0.05, { strip: C.ALERT });
      drawPpwMat(s, 2.55, CONTENT_TOP + 0.20, 4.9, 3.20, {
        wholeNumber: 8,
        partA: 4,
        partBHidden: true,
        partAColor: "D64545",
      });

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
      addTextOnShape(slide, "4 hiding.   Whole 8. Parts 4 and 4.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: We Do — Find two ways to make 7
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Show Two Ways to Make 7", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: STAGE_COLORS["3"], breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Whole = 7", options: { fontSize: 24, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Show TWO ways.", options: { fontSize: 22, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Tell your partner each way.", options: { fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 0.7, y: CONTENT_TOP + 0.20, w: 4.1, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      drawPpwMat(s, 5.45, CONTENT_TOP + 0.20, 3.8, 3.20, {
        wholeNumber: 7,
        partA: null, partB: null,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      // Reveal: gallery of ways to make 7
      addCard(slide, 0.5, 4.20, 9.0, 0.85, { strip: C.SUCCESS });
      slide.addText("Some ways: 6 and 1   |   5 and 2   |   4 and 3   |   7 and 0", {
        x: 0.6, y: 4.25, w: 8.8, h: 0.78,
        fontSize: 19, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 14-15: CFU Hinge — claim hidden part is 4 when correct is 3 (whole 7, see 4)
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
        wholeNumber: 7,
        partA: 4,
        partBHidden: true,
        partAColor: "D64545",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 3.0, { strip: C.ALERT });
      s.addText([
        { text: "I say:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Whole 7. See 4.", options: { fontSize: 22, bold: true, color: C.CHARCOAL, breakLine: true } },
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
      addTextOnShape(slide, "Thumbs DOWN!  Hiding is 3.  4 and 4 makes 8, not 7.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 16: You Do — Show what we know
  workedExSlide(pres, 4, "You Do", "Show What You Know",
    [
      "First: Get 8 cubes and a cup.",
      "",
      "Next: Find two ways to make 8.",
      "",
      "Then: Partner covers one part.",
      "Find the hidden part.",
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
        { text: "8 cubes", options: { bullet: true, fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "your PPW mat and cup", options: { bullet: true, fontSize: 18, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.40, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.50, h: 0.78,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 1.45, lg.rightW, 2.30, {
        wholeNumber: 8,
        partA: null, partB: null,
      });
    }
  );

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Whole 6. I see 4. How many are hiding? Write the answer.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: what is the best thing you learned?",
      scItems: [
        "I can point to the whole.",
        "I can show a way to make a number.",
        "I can find a hidden part.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPWi_Lesson10_Show_What_We_Know.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Show two ways to make 8 and find a hidden part.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addTipBox(doc, "Use 8 counters. Move them. Draw what you see. Then find the hidden part below.", y, { color: C.ACCENT });

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
      return yTop + wholeH + partH + 14;
    }

    function drawPdfPpwMatHidden(yTop, wholeLabel, seenCount) {
      const matX = 70;
      const matW = 470;
      const wholeH = 50;
      const partH = 80;
      doc.lineWidth(2).rect(matX, yTop, matW, wholeH).fillAndStroke("#" + C.PRIMARY, "#333333");
      doc.fillColor("#FFFFFF").fontSize(24).font("Sans-Bold").text("Whole = " + wholeLabel, matX, yTop + 14, { width: matW, align: "center" });
      const partGap = 10;
      const partW = (matW - partGap) / 2;
      doc.lineWidth(2).rect(matX, yTop + wholeH + 8, partW, partH).fillAndStroke("#FFFFFF", "#333333");
      const dotSize = 13;
      const dotGap = 7;
      const dotsTotalW = seenCount * dotSize + (seenCount - 1) * dotGap;
      const dotsStartX = matX + (partW - dotsTotalW) / 2;
      const dotsCY = yTop + wholeH + 8 + partH / 2 - dotSize / 2;
      for (let i = 0; i < seenCount; i += 1) {
        doc.circle(dotsStartX + i * (dotSize + dotGap) + dotSize / 2, dotsCY + dotSize / 2, dotSize / 2)
           .fillAndStroke("#D64545", "#333333");
      }
      doc.lineWidth(2).rect(matX + partW + partGap, yTop + wholeH + 8, partW, partH).fillAndStroke("#FFFFFF", "#333333");
      doc.fillColor("#" + C.ALERT).fontSize(54).font("Sans-Bold").text("?", matX + partW + partGap, yTop + wholeH + 8 + 12, { width: partW, align: "center" });
      doc.fillColor("#333333").font("Sans");
      return yTop + wholeH + partH + 14;
    }

    y = addSectionHeading(doc, "Way 1 to Make 8", y, { color: C.PRIMARY });
    y = drawPdfPpwMat(y, "8");
    doc.fontSize(13).fillColor("#333333").text("Whole 8.   Parts ____ and ____", 70, y);
    y += 22;

    y = addSectionHeading(doc, "Way 2 to Make 8", y, { color: C.PRIMARY });
    y = drawPdfPpwMat(y, "8");
    doc.fontSize(13).fillColor("#333333").text("Whole 8.   Parts ____ and ____", 70, y);
    y += 22;

    y = addSectionHeading(doc, "Find the Hidden Part", y, { color: C.ALERT });
    y = drawPdfPpwMatHidden(y, "6", 4);
    doc.fontSize(13).fillColor("#333333").text("I see 4.   Hidden part = ____", 70, y);
    y += 22;

    addPdfFooter(doc, `Lesson ${SESSION} | Lesson 10 Show What We Know | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Sample answers for the celebration sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addBodyText(doc, "Children may show any valid splits for whole 8.", y);
    y = addSectionHeading(doc, "Whole = 8", y, { color: C.PRIMARY });
    y = addBodyText(doc, "0 and 8   |   1 and 7   |   2 and 6   |   3 and 5   |   4 and 4   |   5 and 3   |   6 and 2   |   7 and 1   |   8 and 0", y);
    y = addSectionHeading(doc, "Hidden Part", y, { color: C.ALERT });
    y = addBodyText(doc, "Whole 6. See 4. Hidden = 2.", y);
    y = addTipBox(doc, "End-of-unit celebration. Foundation students should show two valid ways for 8 and find the hidden part of 6.", y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 10 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
