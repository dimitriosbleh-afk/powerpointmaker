"use strict";

// Part, Part, Whole Introduction Unit — Lesson 3: Make 6 and 7
// Foundation Numeracy | Lesson 3 of 10 | Variant 0
// VC2MFN04
// Daily Review: Subitising 0-10
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
  addDotCard,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 10;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = `Part, Part, Whole | Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`;
const OUT_DIR = "output/PPWi_Lesson3_Make_6_and_7";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 3 Ways to Make 6 and 7", "Find and record the parts that make 6 and 7.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 3 Answer Key", "Sample part splits for 6 and 7.");

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
  const wholeText = o.wholeNumber != null ? "Whole = " + o.wholeNumber : (o.wholeLabel || "Whole");
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
    // 2-row layout when count > 4
    const rows = count > 4 ? 2 : 1;
    const perRow = Math.ceil(count / rows);
    const dotSize = 0.30;
    const gap = 0.08;
    const totalW = perRow * dotSize + (perRow - 1) * gap;
    const startX = boxX + (partW - totalW) / 2;
    const totalRowsH = rows * dotSize + (rows - 1) * gap;
    const rowStartY = boxY + (partH * 0.62) - totalRowsH / 2;
    let placed = 0;
    for (let r = 0; r < rows && placed < count; r += 1) {
      const inRow = Math.min(perRow, count - placed);
      const rowW = inRow * dotSize + (inRow - 1) * gap;
      const rx = boxX + (partW - rowW) / 2;
      for (let i = 0; i < inRow; i += 1) {
        slide.addShape("roundRect", {
          x: rx + i * (dotSize + gap),
          y: rowStartY + r * (dotSize + gap),
          w: dotSize, h: dotSize,
          rectRadius: dotSize / 2,
          fill: { color },
          line: { color: C.CHARCOAL, width: 0.8 },
        });
        placed += 1;
      }
    }
  }
  if (o.partA != null) placeDots(x, partY, o.partA, o.partAColor || "D64545");
  if (o.partB != null) placeDots(x + partW + 0.16, partY, o.partB, o.partBColor || "F4C430");
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
- Today our wholes are bigger.
- We will make 6 and we will make 7.
- Same idea: whole and parts.

DO:
- Have 7 two-sided counters and a cup ready.
- Show the title slide as students settle.

TEACHER NOTES:
Lesson 3 of 5. Numbers grow to 6 and 7. The PPW language is now familiar.

WATCH FOR:
- Students saying "I know this" - language is sticking.
- Students still unsure - more concrete repetitions today.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the dots.
- Do NOT count one by one.
- Say the number all at once.

DO:
- Display dot card showing 5 in a dice-5 pattern.
- Wait 3 seconds.
- Cold call 2-3 students.

TEACHER NOTES:
Subitising 0-10. Today's pattern is 5 (dice-5). Show, hide quickly so students recognise the whole pattern, not count individuals.

WATCH FOR:
- Students who say "5" instantly - secure subitiser.
- Students who count dots one-by-one - still building.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- 5! Five dots in a dice pattern.
- Five fingers on one hand.
- Tick if you said 5.

DO:
- Click to reveal 5.
- Hold up one hand: 5 fingers.

TEACHER NOTES:
Tick and fix.

WATCH FOR:
- Students who match dots to fingers - secure.
- Students still counting - revisit subitising at small group.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency. The teen numbers got mixed up!
- Help me put them in order from smallest to biggest.
- 11 is first. What comes next?

DO:
- Display teen numbers out of order: 14, 11, 17, 13, 20.
- Lead a class chant putting them in order.
- Students show fingers for each number as it is named.

TEACHER NOTES:
Teen-number ordering. Do this orally. The slide just shows the numbers.

WATCH FOR:
- Students who order quickly - secure.
- Students who confuse 13 and 30 - common; emphasise the "teen" sound.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Today our wholes are 6 and 7.
- We are learning to find the parts that make 6 and 7.

DO:
- Choral read LI and SC.
- Hold up 6 fingers. Hold up 7 fingers.

TEACHER NOTES:
SC1: at least one way for whole 6 or 7. SC2: ways for both. SC3: explain that the whole stays the same.

WATCH FOR:
- Students who already named splits - secure.
- Students still hesitating - small-group at table time.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- I have 6 counters. 6 is the whole.
- I shake and spill.
- 4 are red. 2 are yellow.
- Whole 6. Parts 4 and 2.

DO:
- Use a cup with 6 two-sided counters.
- Shake and spill on document camera.
- Move counters into the part-cells of the mat.
- Repeat: "whole 6, parts 4 and 2".

TEACHER NOTES:
Same routine, bigger whole. The mat is now familiar.

WATCH FOR:
- Students who chant the pattern with you - secure.
- Students who count 1, 2, 3, 4, 5, 6 each time - prompt: "Count just one colour."

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now whole 7. I have 7 counters.
- Shake and spill.
- 5 are red. 2 are yellow.
- Whole 7. Parts 5 and 2.

DO:
- Use a cup with 7 counters.
- Shake and spill.
- Move counters into the part-cells.

TEACHER NOTES:
7 is the biggest whole today. The big idea is the same: same whole, different parts.

WATCH FOR:
- Students who say "I see 5 quickly" - using subitising.
- Students who lose count - manage their workload, prompt one colour at a time.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Look at the picture.
- Show me on your fingers how many YELLOW.

DO:
- Display 6 counters: 4 red, 2 yellow.
- Wait 5 seconds.
- "Show me... ready... show."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "How many yellow? Show me."
- Scan for: 2 fingers.
PROCEED: If 80% show 2, click to reveal.
PIVOT: Most likely misconception - students count all 6.
- Reteach: "Just yellow. Point to only yellow."
- Re-check: Same picture, ask "How many red?"

TEACHER NOTES:
Same finger-voting routine as previous lessons.

WATCH FOR:
- 2 fingers up quickly - secure.
- 6 fingers up - they counted everything.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Look at the spill.
- With your partner, whisper: how many RED? How many YELLOW?
- Write on your whiteboard: 7 = ___ and ___.

DO:
- Display 7 counters: 3 red, 4 yellow.
- Allow 30 seconds.
- Walk and scan whiteboards.

TEACHER NOTES:
The We Do uses a different split for 7. The whole stays 7.

WATCH FOR:
- Pairs who write 7 = 3 and 4 - secure.
- Pairs who write 7 = 7 - prompt: "Each colour."

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Whole 7. Parts 3 and 4.
- Same whole, different parts again.

DO:
- Click to reveal.
- Repeat the pattern with the class.

TEACHER NOTES:
Same whole, different parts is the through-line of the unit.

WATCH FOR:
- Confident response - secure.
- Slow response - small group.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge time. Listen carefully.
- I say: "6 is 3 and 4."
- Is that right? Thumbs up or thumbs down.

DO:
- Say the false claim aloud.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "6 is 3 and 4. Thumbs up if yes, thumbs down if no."
- Scan for: thumbs DOWN. 3 and 4 makes 7, not 6.
PROCEED: If 80% show thumbs down, click to reveal.
PIVOT: Most likely misconception - students agree without checking.
- Reteach: "Hold up 3 fingers. Now 4 more. Count them. How many?"
- Re-check: "Is the whole 6 or 7?"

TEACHER NOTES:
The hinge probes part-whole tracking with a near-miss.

WATCH FOR:
- Confident thumbs down - secure.
- Mixed thumbs - reteach with fingers.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn.
- Get 6 counters and your mat. Find one way.
- Then get 7 counters. Find one way.

DO:
- Distribute 7 two-sided counters and a printed mat per child.
- Circulate. Listen for the language pattern.
- Cold call 2 students for each whole.

TEACHER NOTES:
Today students work with both 6 and 7. Acceptable splits include any pair that adds to the whole.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 5 counters. Find one way to make 5 (revision).
EXTENDING PROMPT:
- Task: Find at least 3 ways to make 7. Draw each one.

WATCH FOR:
- Students who use the language without prompting - secure.
- Students who show one way only - prompt: "What if one moved across?"

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. Look at the picture.
- Write the whole and the parts on your whiteboard.

DO:
- Display: 5 red, 1 yellow.
- Allow 60 seconds.
- Collect or photograph whiteboards.

TEACHER NOTES:
Exit ticket: whole 6, parts 5 and 1.

WATCH FOR:
- Students who write whole 6, parts 5 and 1 - secure.
- Students who write only 6 - reteach part-whole language.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today our wholes were 6 and 7.
- The whole stayed the same. The parts changed.
- Show me thumbs: can you find one way to make 6?

DO:
- Read SC.
- Use thumbs up sideways down.
- Quick partner share.

TEACHER NOTES:
Most students should reach SC2.

WATCH FOR:
- Strong thumbs up - secure.
- Sideways/down - small group tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Today's recording sheet has space for 6 and 7.

DO:
- Print one sheet per student.
- Have 7 two-sided counters and a cup at each table.

TEACHER NOTES:
One core student resource today.

[General: Resources]`;

const NOTES_LAUNCH = `SAY:
- Yesterday we found ways to make 5.
- Look at this mat. The whole is 5. The parts are 3 and 2.
- Whisper to your partner: what is the whole? What are the parts?
- Today our wholes are bigger - 6 and 7.

DO:
- Display the PPW mat showing whole 5, parts 3 and 2.
- Allow 15 seconds for partner whisper.
- Cold call 1 student.
- Hold up 6 counters and say "Today we will make 6, then 7."

TEACHER NOTES:
Quick recall of yesterday's "Make 5" plus a clear bridge to today's bigger wholes. Keep it under 90 seconds.

WATCH FOR:
- Pairs who use whole and parts language confidently - secure.
- Pairs who only point - prompt: "What is the big number called?"

[General: Launch | VTLM 2.0: Activate Prior Knowledge]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 3: Make 6 and 7",
    `Foundation Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch — recall make 5 and connect to today's bigger wholes
  contentSlide(pres, "Launch", C.SUCCESS, "Yesterday: Make 5",
    [
      "Yesterday we made 5 in lots of ways.",
      "What is the whole? What are the parts?",
      "Today our wholes are 6 and 7.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 5,
        partA: 3, partB: 2,
        partAColor: "D64545",
        partBColor: "F4C430",
      });
    }
  );

  // Slides 2-3: Daily Review with reveal — Subitising
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How Many Dots?", { color: C.ACCENT });

      addCard(s, 0.5, CONTENT_TOP, 9.0, 3.0, { strip: C.ACCENT });

      // One large dot card showing 5 (sized to leave clear gap above reveal bar at y=4.45)
      addDotCard(s, 3.95, CONTENT_TOP + 0.10, 2.1, 5, {
        dotColor: C.PRIMARY,
        cardFill: C.WHITE,
        borderColor: C.CHARCOAL,
      });

      // Subitise routine cue (kept compact above reveal bar)
      s.addText("Say it. Don't count.", {
        x: 0.7, y: CONTENT_TOP + 2.65, w: 8.6, h: 0.30,
        fontSize: 18, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "5 dots!  (dice-5 pattern)", {
        x: 1.5, y: 4.45, w: 7.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Teen Numbers ordering
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Order the Teen Numbers", { color: C.ACCENT });

  // Five large numeral chips: 14, 11, 17, 13, 20 (out of order)
  const teens = [14, 11, 17, 13, 20];
  const chipY = CONTENT_TOP + 0.50;
  const chipW = 1.5;
  const chipH = 1.5;
  const totalW = teens.length * chipW + (teens.length - 1) * 0.20;
  const chipStartX = (10 - totalW) / 2;
  teens.forEach((n, i) => {
    const cx = chipStartX + i * (chipW + 0.20);
    sFluency.addShape("roundRect", {
      x: cx, y: chipY, w: chipW, h: chipH, rectRadius: 0.12,
      fill: { color: C.PRIMARY },
    });
    sFluency.addText(String(n), {
      x: cx, y: chipY, w: chipW, h: chipH,
      fontSize: 56, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  addCard(sFluency, 1.0, chipY + chipH + 0.40, 8.0, 0.85, { strip: C.PRIMARY });
  sFluency.addText("Smallest to biggest!", {
    x: 1.2, y: chipY + chipH + 0.42, w: 7.6, h: 0.80,
    fontSize: 26, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to find the parts that make 6 and 7.",
    [
      "I can find one way to make 6.",
      "I can find one way to make 7.",
      "I can say the whole stays the same when the parts change.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do 1 — Make 6 (4 and 2)
  workedExSlide(pres, 2, "I Do", "Make 6",
    [
      "I have 6 counters.",
      "6 is the whole.",
      "",
      "Shake and spill.",
      "",
      "4 are red.",
      "2 are yellow.",
      "",
      "Whole 6.",
      "Parts 4 and 2.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 6,
        partA: 4, partB: 2,
      });
    }
  );

  // Slide 7: I Do 2 — Make 7 (5 and 2)
  workedExSlide(pres, 2, "I Do", "Make 7",
    [
      "I have 7 counters.",
      "7 is the whole.",
      "",
      "Shake and spill.",
      "",
      "5 are red.",
      "2 are yellow.",
      "",
      "Whole 7.",
      "Parts 5 and 2.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 7,
        partA: 5, partB: 2,
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

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // 6 counters: 4 red, 2 yellow
      drawCounterRow(s, 1.4, CONTENT_TOP + 0.55, 6, 4, { size: 0.85, gap: 0.18 });

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
      addTextOnShape(slide, "2 yellow.   Whole 6.   Parts 4 and 2.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 10-11: We Do with reveal — 7 = 3 and 4
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Find the Parts of 7", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, 3.0, { strip: STAGE_COLORS["3"] });

      // 7 counters in a single clean row: 3 red + 4 yellow
      // 7 * 0.50 + 6 * 0.10 = 4.10 wide; centred in 4.5" card from x=0.5
      drawCounterRow(s, 0.70, CONTENT_TOP + 0.55, 7, 3, { size: 0.50, gap: 0.10 });

      s.addText("Whole = 7", {
        x: 0.6, y: CONTENT_TOP + 1.40, w: 4.3, h: 0.50,
        fontSize: 26, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      s.addText("How many RED? How many YELLOW?", {
        x: 0.6, y: CONTENT_TOP + 2.00, w: 4.3, h: 0.50,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 3.0, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "7 = ____ and ____", options: { fontSize: 28, bold: true, color: C.CHARCOAL } },
      ], {
        x: 5.4, y: CONTENT_TOP + 0.35, w: 3.9, h: 2.55,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Whole 7.   Parts 3 and 4.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 12-13: CFU Hinge with reveal — 6 is 3 and 4 (false; that makes 7)
  withReveal(
    () => cfuSlide(pres, "CFU", "Is This Right?", "Thumbs Up or Thumbs Down",
      "I say:\n6 is 3 and 4.\n\nIs that right?",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN!  3 and 4 make 7, not 6.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Your Turn: 6 and 7",
    [
      "First: 6 counters.",
      "Find one way to make 6.",
      "",
      "Next: 7 counters.",
      "Find one way to make 7.",
      "",
      "Tell your partner:",
      "Whole ___, parts ___ and ___.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.20, { strip: C.ALERT });
      slide.addText("You need:", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.40, h: 0.36,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "7 two-sided counters", options: { bullet: true, fontSize: 17, color: C.CHARCOAL, breakLine: true } },
        { text: "your PPW mat", options: { bullet: true, fontSize: 17, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.40, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.50, h: 0.70,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 1.35, lg.rightW, 2.30, {
        wholeNumber: 7,
        partA: null, partB: null,
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Whole 6, parts 5 and 1. Draw it on your whiteboard.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 16: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Tell your partner one way to make 6 or 7.",
      scItems: [
        "I can find one way to make 6.",
        "I can find one way to make 7.",
        "I can say the whole stays the same when the parts change.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPWi_Lesson3_Make_6_and_7.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find ways to make 6 and 7.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addTipBox(doc, "Shake your cup. Spill the counters. Colour the dots. Write the parts.", y, { color: C.ACCENT });

    function drawShakeRow(yTop, total, label) {
      doc.fontSize(16).font("Sans-Bold").fillColor("#" + C.PRIMARY).text(label, 60, yTop);
      const dotY = yTop + 28;
      const size = 30;
      const gap = 10;
      for (let i = 0; i < total; i += 1) {
        doc.circle(80 + i * (size + gap) + size / 2, dotY + size / 2, size / 2)
          .lineWidth(1.5).stroke("#333333");
      }
      doc.fontSize(14).font("Sans").fillColor("#333333").text(`Whole ${total}.   Parts ____ and ____`, 60, dotY + size + 12);
      return dotY + size + 46;
    }

    y = addSectionHeading(doc, "Whole = 6", y, { color: C.PRIMARY });
    y = drawShakeRow(y, 6, "Shake 1");
    y = drawShakeRow(y, 6, "Shake 2");

    y = addSectionHeading(doc, "Whole = 7", y, { color: C.PRIMARY });
    y = drawShakeRow(y, 7, "Shake 1");
    y = drawShakeRow(y, 7, "Shake 2");

    y = addTipBox(doc, "Did your two shakes give the same parts? Or different parts?", y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Lesson 3 Ways to Make 6 and 7 | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Sample part splits.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addBodyText(doc, "All splits below are correct.", y);
    y = addSectionHeading(doc, "Whole = 6", y, { color: C.PRIMARY });
    y = addBodyText(doc, "6 and 0   |   5 and 1   |   4 and 2   |   3 and 3   |   2 and 4   |   1 and 5   |   0 and 6", y);
    y = addSectionHeading(doc, "Whole = 7", y, { color: C.PRIMARY });
    y = addBodyText(doc, "7 and 0   |   6 and 1   |   5 and 2   |   4 and 3   |   3 and 4   |   2 and 5   |   1 and 6   |   0 and 7", y);
    y = addTipBox(doc, "Children may use the same split with red and yellow swapped. Accept all splits that add to the whole.", y, { color: C.ACCENT });
    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
