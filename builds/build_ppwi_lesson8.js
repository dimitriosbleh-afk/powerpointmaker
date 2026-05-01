"use strict";

// Part, Part, Whole Introduction Unit — Lesson 8: All the Ways to Make 7
// Foundation Numeracy | Lesson 8 of 10 | Variant 0
// VC2MFN04 - partition and combine collections up to 10 using PPW + subitising
// Daily Review: Subitising 11-20
// Fluency: Teen Numbers - order numbers to 20

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
  addTensFrame,
  STAGE_COLORS,
} = T;

const SESSION = 8;
const TOTAL = 10;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = `Part, Part, Whole | Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`;
const OUT_DIR = "output/PPWi_Lesson8_All_Ways_Make_7";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 8 All the Ways to Make 7", "Show different ways to make 7 with two parts.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 8 Answer Key", "All the ways to make 7.");

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
    const dotSize = 0.26;
    const gap = 0.07;
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

function drawMiniWayCard(slide, x, y, w, h, partA, partB, opts) {
  const o = opts || {};
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.10,
    fill: { color: C.WHITE },
    line: { color: o.borderColor || C.SECONDARY, width: 2 },
  });
  slide.addText(`${partA} and ${partB}`, {
    x, y: y + 0.05, w, h: 0.40,
    fontSize: 20, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  const total = partA + partB;
  const dotSize = 0.20;
  const gap = 0.05;
  const totalW = total * dotSize + (total - 1) * gap;
  const startX = x + (w - totalW) / 2;
  const cy = y + 0.62;
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
- Yesterday we made 6 in different ways.
- Today we will find different ways to make 7.

DO:
- Have 7 two-sided counters and a PPW mat ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 8 of 10. Whole is now 7. Same routine, slightly bigger number. Foundation students should feel confident with the routine.

WATCH FOR:
- Students who confidently say whole and parts - language secure.
- Students who flag the bigger number - reassure them.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are the materials for today.
- We will use the All Ways to Make 7 sheet at the table.

DO:
- Print the All Ways to Make 7 sheet (one per child).
- Have 7 unifix cubes per child and a PPW mat.

TEACHER NOTES:
Same sheet style as yesterday but with whole 7. Routine is now familiar.

[General: Resources]`;

const NOTES_LAUNCH = `SAY:
- Yesterday we made 6 in different ways.
- Today we will make 7.
- Watch. I have 7 counters in my cup.
- I tip them out. Some land here. Some land here.
- Same 7 counters. Different parts.

DO:
- Tip 7 two-sided counters out of a cup.
- Show 6 and 1, then 4 and 3.
- Point to the whole and say still 7.

TEACHER NOTES:
Same launch routine as Lesson 7. Brisk - 60 to 90 seconds.

WATCH FOR:
- Students who say still 7 - the whole idea is secure.
- Students who confidently subitise the colour groups - good prior knowledge.

[General: Launch | VTLM 2.0: Activate Prior Knowledge]`;

const NOTES_DR_Q = `SAY:
- Warm up time. Look quickly.
- I will show a tens frame. Do not count one by one.
- Say the number when I say show.

DO:
- Display a tens frame showing 13 (10 on top row, 3 on bottom row).
- Hide for 2 seconds, then say show.
- Allow students to call out together.

TEACHER NOTES:
Daily Review: subitising teen numbers. The aim is recognition without one-by-one counting. Hide-and-show helps subitising.

WATCH FOR:
- Students who say 13 quickly - subitising secure.
- Students who count dots one by one - prompt see the 10 row first.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check together. The number is thirteen.
- 10 on top, 3 more underneath. 13.
- Say it with me: thirteen.

DO:
- Click to reveal 13.
- Lead a choral say of thirteen.

TEACHER NOTES:
Tick and fix. Note any students who counted one by one - small group at table time.

WATCH FOR:
- Students who self-correct - they noticed.
- Students who keep guessing - small group with tens frames.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Look at the cards on the slide.
- They are mixed up.
- With your partner, whisper them in order.
- Smallest first.

DO:
- Display number cards 11, 14, 17, 20, 13.
- Allow 30 seconds.
- Cold call a pair to share.

TEACHER NOTES:
Fluency: ordering teen numbers. Builds awareness of relative size.

WATCH FOR:
- Pairs who order them correctly - secure.
- Pairs who put 17 before 14 - common confusion.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to find different ways to make 7.
- Read the success criteria with me.

DO:
- Choral read the LI.
- Choral read each criterion.
- Hold up 7 counters as you say seven.

TEACHER NOTES:
SC1 is achievable for all. SC2 is the core target. SC3 stretches students to find more than one way.

WATCH FOR:
- Students who chime in - language secure.
- Students who look quiet - I Do will rebuild the picture.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. I have 7 counters. 7 is the whole.
- I put 6 in this part. I put 1 in this part.
- Whole 7. Parts 6 and 1.
- Say with me: whole 7, parts 6 and 1.

DO:
- Show the PPW mat with 7 counters.
- Move 6 to the left part, 1 to the right.
- Repeat the language pattern at least 3 times.

TEACHER NOTES:
First way today. Move slowly.

WATCH FOR:
- Students who echo whole 7, parts 6 and 1 - language secure.
- Students who count all 7 dots - re-show with colour difference.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now watch this. Same 7 counters. I move them.
- I put 4 here. I put 3 here.
- Whole 7. Parts 4 and 3.
- Same whole. Different parts.

DO:
- Slide counters to make 4 and 3.
- Point to the whole and say still 7.
- Say whole 7, parts 4 and 3 three times.

TEACHER NOTES:
4 and 3 is a balanced split close to half. Foundation students enjoy noticing close-to-fair splits.

WATCH FOR:
- Students who say closer to fair - they are noticing.
- Students who say it's still 7 - the whole concept is secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Look at the counters on the slide.
- Count the red. Count the yellow.
- Show me on your fingers how many yellow.

DO:
- Display 7 counters: 5 red, 2 yellow.
- Wait 5 seconds.
- Say: Show me on your fingers... ready... show.

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: How many yellow? Show me on your fingers... show.
- Scan for: 2 fingers held up.
PROCEED: If 80% show 2 fingers, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count all 7 counters.
- Reteach: Just one colour. Point only to yellow. How many is that?
- Re-check: Show 7 counters with 3 red, 4 yellow and ask How many yellow?

TEACHER NOTES:
Fingers are universal at Foundation. Quick, every-student response.

WATCH FOR:
- 2 fingers held up - secure with parts.
- 7 fingers - they counted everything; reteach colour focus.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Look at the mini cards.
- We have shown two ways to make 7 already.
- With your partner, whisper: what other ways can we find?
- On your whiteboard, draw one new way for whole 7.

DO:
- Display the gallery with 6 and 1 and 4 and 3 already filled.
- Allow 60 seconds.
- Walk and scan.

TEACHER NOTES:
We Do invites students to find a NEW way that is not already on the board.

WATCH FOR:
- Pairs offering 5 and 2 or 7 and 0 - secure.
- Pairs repeating one shown - prompt: that is one we have. Find a different way.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together. Here are more ways to make 7.
- 5 and 2.
- 2 and 5.
- 7 and 0.
- We have lots of ways now.

DO:
- Click to reveal the rest of the gallery.
- Repeat the language pattern.

TEACHER NOTES:
For 7: 0&7, 1&6, 2&5, 3&4, 4&3, 5&2, 6&1, 7&0.

WATCH FOR:
- Students who try to find more - secure.
- Students who needed prompting - small group.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. Look at the counters.
- I am saying: Whole is 7. Parts are 5 and 3.
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Display 7 counters total (5 red, 2 yellow).
- Say the false claim aloud while pointing to the slide.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: Whole is 7. Parts are 5 and 3. Thumbs up if yes, thumbs down if no.
- Scan for: thumbs DOWN. The parts are 5 and 2, not 5 and 3.
PROCEED: If 80% show thumbs down, click to reveal correct version.
PIVOT: Most likely misconception - students see one part correctly and trust the second.
- Reteach: Count yellow. Is it 3 or 2?
- Re-check: Show parts 6 and 1 and ask: is the whole 7?

TEACHER NOTES:
This hinge probes whether students check both parts. A no answer means they are tracking each part separately.

WATCH FOR:
- Confident thumbs down - students see the mismatch.
- Slow thumbs - they are unsure; reteach with concrete counters.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- You need 7 cubes and your PPW mat.
- Find at least two different ways to make 7.
- Tell your partner: whole 7, parts ___ and ___.

DO:
- Hand out 7 unifix cubes and a printed PPW mat.
- Circulate. Listen for the language pattern.
- Cold call 2 students to share their two ways.

TEACHER NOTES:
All correct splits are valid. The goal today is finding more than one way and naming each one.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 6 cubes instead. Find two ways to make 6.
EXTENDING PROMPT:
- Task: Find every way to make 7. Record on your whiteboard.

WATCH FOR:
- Students who find 2 unique splits - secure.
- Students who keep one split - prompt: now move one cube.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard, draw one way to make 7.
- Then write the parts.

DO:
- Allow 60 seconds.
- Collect whiteboards or take a photo.

TEACHER NOTES:
Exit ticket assesses the core target. Any valid pair is correct.

WATCH FOR:
- Students who show parts that add to 7 - on track.
- Students who only draw 7 dots in one part - prompt: where is the other part?

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today we found different ways to make 7.
- The whole stayed the same. The parts changed.
- Show me thumbs up if you found two ways.
- Turn and tell: what is one way to make 7?

DO:
- Read the success criteria with the class.
- Use thumbs up sideways down.
- Cold call 1 to 2 students for the partner share.

TEACHER NOTES:
Self-assessment data informs Lesson 9 grouping.

WATCH FOR:
- Strong thumbs up - language and skill forming.
- Sideways or down - small group focus.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 8: All the Ways to Make 7",
    `Foundation Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch
  contentSlide(pres, "Launch", C.SUCCESS, "Yesterday 6. Today 7.",
    [
      "7 counters in my cup.",
      "I tip them out.",
      "I move them. Same 7.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      const visualX = lg.rightX + 0.10;
      const visualY = lg.panelTopPadded + 0.15;

      slide.addText("Way 1: 6 and 1", {
        x: visualX, y: visualY, w: lg.rightW - 0.20, h: 0.36,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      drawCounterRow(slide, visualX + 0.05, visualY + 0.45, 6, 6,
        { size: 0.38, gap: 0.07 });
      drawCounterRow(slide, visualX + 2.95, visualY + 0.45, 1, 0,
        { size: 0.38, gap: 0.07 });

      slide.addShape("line", {
        x: visualX + 0.20, y: visualY + 1.30, w: lg.rightW - 0.60, h: 0,
        line: { color: C.MUTED || C.CHARCOAL, width: 1, dashType: "dash" },
      });

      slide.addText("Way 2: 4 and 3", {
        x: visualX, y: visualY + 1.45, w: lg.rightW - 0.20, h: 0.36,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      drawCounterRow(slide, visualX + 0.20, visualY + 1.95, 4, 4,
        { size: 0.38, gap: 0.07 });
      drawCounterRow(slide, visualX + 2.10, visualY + 1.95, 3, 0,
        { size: 0.38, gap: 0.07 });

      slide.addText("Still 7!", {
        x: visualX, y: visualY + 2.75, w: lg.rightW - 0.20, h: 0.42,
        fontSize: 22, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slides 4-5: Daily Review with reveal — Subitising 13 from tens frame
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "What Number Is It?", { color: C.ACCENT });

      // Left card: tens frame showing 13 (10 + 3)
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      // Show "Look fast" at top
      s.addText("Look fast.", {
        x: 0.6, y: CONTENT_TOP + 0.15, w: 4.3, h: 0.40,
        fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      // Tens frame (top): 10 filled
      addTensFrame(s, 0.95, CONTENT_TOP + 0.70, 3.6, 10, { cellH: 0.55 });
      // Tens frame (bottom): 3 filled
      addTensFrame(s, 0.95, CONTENT_TOP + 1.95, 3.6, 3, { cellH: 0.55 });

      // Right card: prompt
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText([
        { text: "What number?", options: { fontSize: 28, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 12, breakLine: true } },
        { text: "Whisper to your partner.", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 12, breakLine: true } },
        { text: "Say it with me on 3.", options: { fontSize: 22, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.30, w: 3.85, h: SAFE_BOTTOM - CONTENT_TOP - 0.55,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "13   (10 and 3 more)", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 6: Fluency — Order numbers up to 20 (mixed cards)
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Put Them in Order", { color: C.ACCENT });

  // 5 number cards mixed up across the slide
  const cards = [11, 14, 17, 20, 13];
  const cardW = 1.35;
  const cardH = 1.35;
  const totalCardsW = cards.length * cardW + (cards.length - 1) * 0.30;
  const startX = (10 - totalCardsW) / 2;
  const cardY = CONTENT_TOP + 0.50;
  for (let i = 0; i < cards.length; i += 1) {
    sFluency.addShape("roundRect", {
      x: startX + i * (cardW + 0.30), y: cardY,
      w: cardW, h: cardH,
      rectRadius: 0.12,
      fill: { color: C.PRIMARY },
      line: { color: C.PRIMARY, width: 1 },
    });
    sFluency.addText(String(cards[i]), {
      x: startX + i * (cardW + 0.30), y: cardY,
      w: cardW, h: cardH,
      fontSize: 56, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }

  addCard(sFluency, 1.0, cardY + cardH + 0.30, 8.0, 0.85, { strip: C.SECONDARY });
  sFluency.addText("Smallest first. Whisper to your partner.", {
    x: 1.2, y: cardY + cardH + 0.32, w: 7.6, h: 0.80,
    fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 7: LI/SC
  liSlide(pres,
    "We are learning to find different ways to make 7.",
    [
      "I can point to the whole.",
      "I can show one way to make 7.",
      "I can find more than one way to make 7.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 8: I Do 1 — Whole 7, parts 6 and 1
  workedExSlide(pres, 2, "I Do", "One Way to Make 7",
    [
      "I have 7 counters.",
      "",
      "7 is the whole.",
      "",
      "I put 6 here.",
      "I put 1 here.",
      "",
      "Whole 7.",
      "Parts 6 and 1.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 7,
        partA: 6, partB: 1,
        partAColor: "D64545",
        partBColor: "F4C430",
      });
    }
  );

  // Slide 9: I Do 2 — Whole 7, parts 4 and 3
  workedExSlide(pres, 2, "I Do", "Another Way to Make 7",
    [
      "Same 7 counters.",
      "",
      "I move them.",
      "",
      "I put 4 here.",
      "I put 3 here.",
      "",
      "Whole 7.",
      "Parts 4 and 3.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 7,
        partA: 4, partB: 3,
        partAColor: "D64545",
        partBColor: "F4C430",
      });
    }
  );

  // Slides 10-11: CFU 1 — How many yellow? (5 red, 2 yellow)
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

      drawCounterRow(s, 0.85, CONTENT_TOP + 0.45, 7, 5, { size: 0.85, gap: 0.18 });

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
      addTextOnShape(slide, "2 yellow.   Whole 7.   Parts 5 and 2.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: We Do — Find more ways to make 7
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "What Other Ways?", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText("Already shown:", {
        x: 0.6, y: CONTENT_TOP + 0.10, w: 4.3, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      drawMiniWayCard(s, 0.7, CONTENT_TOP + 0.55, 1.95, 1.20, 6, 1);
      drawMiniWayCard(s, 2.85, CONTENT_TOP + 0.55, 1.95, 1.20, 4, 3);

      s.addText("On your whiteboard:", {
        x: 0.6, y: CONTENT_TOP + 1.95, w: 4.3, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("Draw a NEW way to make 7.", {
        x: 0.6, y: CONTENT_TOP + 2.30, w: 4.3, h: 0.50,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
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
      addCard(slide, 0.5, 4.20, 9.0, 0.85, { strip: C.SUCCESS });
      slide.addText("More ways: 5 and 2   |   2 and 5   |   7 and 0   |   3 and 4", {
        x: 0.6, y: 4.25, w: 8.8, h: 0.78,
        fontSize: 19, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 14-15: CFU Hinge — false claim 5 and 3 with 5 and 2 picture
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
      drawCounterRow(s, 0.65, CONTENT_TOP + 0.55, 7, 5, { size: 0.55, gap: 0.13 });
      s.addText("These counters", {
        x: 0.6, y: CONTENT_TOP + 1.85, w: 4.3, h: 0.40,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 3.0, { strip: C.ALERT });
      s.addText([
        { text: "I say:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Whole is 7.", options: { fontSize: 26, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "Parts are 5 and 3.", options: { fontSize: 26, bold: true, color: C.CHARCOAL, breakLine: true } },
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
      addTextOnShape(slide, "Thumbs DOWN!  Parts are 5 and 2.  5 and 3 makes 8.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 16: You Do
  workedExSlide(pres, 4, "You Do", "Find Two Ways to Make 7",
    [
      "First: Get 7 cubes.",
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
        { text: "7 cubes", options: { bullet: true, fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "your PPW mat", options: { bullet: true, fontSize: 18, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.40, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.50, h: 0.78,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 1.45, lg.rightW, 2.30, {
        wholeNumber: 7,
        partA: null, partB: null,
      });
    }
  );

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Whole 7. Draw one way to show parts. Write the parts.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: what is one way to make 7?",
      scItems: [
        "I can point to the whole.",
        "I can show one way to make 7.",
        "I can find more than one way to make 7.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPWi_Lesson8_All_Ways_Make_7.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find different ways to make 7 with two parts.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addTipBox(doc, "Use 7 counters. Move them. Draw what you see. Say: whole 7, parts ___ and ___.", y, { color: C.ACCENT });

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
    y = drawPdfPpwMat(y, "7");
    doc.fontSize(13).fillColor("#333333").text("Whole 7.   Parts ____ and ____", 70, y);
    y += 24;

    y = addSectionHeading(doc, "Way 2", y, { color: C.PRIMARY });
    y = drawPdfPpwMat(y, "7");
    doc.fontSize(13).fillColor("#333333").text("Whole 7.   Parts ____ and ____", 70, y);
    y += 24;

    y = addSectionHeading(doc, "Way 3", y, { color: C.PRIMARY });
    y = drawPdfPpwMat(y, "7");
    doc.fontSize(13).fillColor("#333333").text("Whole 7.   Parts ____ and ____", 70, y);
    y += 24;

    addPdfFooter(doc, `Lesson ${SESSION} | Lesson 8 All the Ways to Make 7 | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "All the ways to make 7.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addBodyText(doc, "Children may show any of these splits. Order does not matter for Foundation - 6 and 1 is the same idea as 1 and 6.", y);
    y = addSectionHeading(doc, "Whole = 7", y, { color: C.PRIMARY });
    y = addBodyText(doc, "0 and 7   |   1 and 6   |   2 and 5   |   3 and 4   |   4 and 3   |   5 and 2   |   6 and 1   |   7 and 0", y);
    y = addTipBox(doc, "Today's focus is finding more than one way. Most Foundation students will find 2 to 3 ways.", y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 8 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
