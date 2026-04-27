"use strict";

// Part, Part, Whole Introduction Unit — Lesson 2: Make 5
// Foundation Numeracy | Lesson 2 of 5 | Variant 0
// VC2MFN04 - partition and combine collections up to 10 using PPW + subitising
// Daily Review: Counting up to 20 fluently by 1s
// Fluency: Backwards counting from 20

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
  withReveal,
  addNumberTrack,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 5;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = `Part, Part, Whole | Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`;
const OUT_DIR = "output/PPWi_Lesson2_Make_5";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 2 Ways to Make 5", "Find and record the parts that make 5.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 2 Answer Key", "All the ways to make 5.");

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
- Welcome back. Yesterday we learned about parts and a whole.
- Today we will find lots of ways to make 5.
- 5 is the whole. We will find different parts that make 5.

DO:
- Hold up a cup of 5 two-sided counters as students settle.
- Show the cover slide. Read the title.

TEACHER NOTES:
Lesson 2 of 5. Build directly on Lesson 1 language. Today's focus: 5 has many splits.

WATCH FOR:
- Students who say "I remember whole and parts" - Lesson 1 stuck.
- Students who look unsure - they will get more chances today.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the number track.
- Whisper count from 1 to 20 with your partner.
- Then tell me: which number comes after 13?

DO:
- Display the 1-20 number track.
- Allow 30 seconds for partner whisper count.
- Hold up fingers ready for the after-13 question.

TEACHER NOTES:
Daily Review: counting fluency to 20. The "what comes after" question primes the teen-number sequence work.

WATCH FOR:
- Students who count smoothly - secure.
- Students who skip 14, 16 or stumble at the teens - common.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- 14 comes after 13.
- Tick if you got 14.
- Now let us count from 1 to 20 together one more time.

DO:
- Click to reveal 14 highlighted on the track.
- Lead a quick choral count 1-20.

TEACHER NOTES:
Tick and fix. The choral count seals the teen sequence.

WATCH FOR:
- Students who self-correct - on track.
- Students still stuck - small group at table time.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time! Now we count BACKWARDS from 20.
- This is harder. Listen first.
- 20, 19, 18, 17, 16... your turn.
- Stand and count down. We are counting down to blast off!

DO:
- Model a backwards count once for the class.
- Lead choral backwards count 20-1.
- Repeat 2-3 times.

TEACHER NOTES:
Backwards counting from 20. This builds number-sequence strength and prepares for take-away later in the year.

WATCH FOR:
- Students who can count back smoothly - secure.
- Students who get stuck - common from 14 down. Pause and prompt.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to find different ways to make 5.
- 5 is the whole. The parts can change.
- Read the success criteria with me.

DO:
- Choral read the LI and each SC.
- Hold up 5 counters as you say "5 is the whole".

TEACHER NOTES:
SC1 names the whole and finds one way. SC2 finds at least two different ways. SC3 explains that the whole stays 5.

WATCH FOR:
- Students who can already say "5 is 4 and 1" - secure.
- Students who only know one way - the I Do will build more.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. The whole is 5.
- I shake the cup and spill the counters.
- Look - 4 are red, 1 is yellow.
- Whole 5. Parts 4 and 1.

DO:
- Use a real cup with 5 two-sided counters.
- Shake and spill on the document camera or floor mat.
- Place counters into the part-cells of the mat.
- Repeat: "whole 5, parts 4 and 1".

TEACHER NOTES:
First way to make 5. Use the same language pattern as Lesson 1. The split happens randomly when you spill - this is the magic of "Shake and Spill".

WATCH FOR:
- Students who say the pattern back - secure.
- Students who count one colour twice - reteach colour focus.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- I shake again. Same 5 counters in the cup.
- Spill! This time, 3 are red, 2 are yellow.
- Whole 5. Parts 3 and 2.
- Same whole. Different parts!

DO:
- Shake and spill again.
- Move counters into the part-cells.
- Point to both ways on the slide and emphasise: "same whole, different parts".

TEACHER NOTES:
This is the key idea of the lesson. Same whole, different parts. Say this phrase many times today.

MISCONCEPTIONS:
- Misconception: If the parts change, the whole must change.
  Why: Foundation students often link the parts and the whole as fixed.
  Impact: They will struggle with addition fact families later.
  Quick correction: "I did not take any counters away. The whole is still 5. The parts just moved."

WATCH FOR:
- Students who echo "same whole, different parts" - they get it.
- Students who look confused - re-show with fingers held up.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Look at the picture.
- Show me on your fingers how many YELLOW.

DO:
- Display the picture: 5 counters - 2 red and 3 yellow.
- Wait 5 seconds.
- Say: "Show me... ready... show."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "How many yellow? Show me on your fingers."
- Scan for: 3 fingers held up.
PROCEED: If 80% show 3, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count all 5 counters.
- Reteach: "Just yellow. Point to only the yellow counters."
- Re-check: Same picture, ask "How many red?" Expected: 2.

TEACHER NOTES:
Finger voting works for any class size. Quick scan for readiness.

WATCH FOR:
- 3 fingers up quickly - secure.
- 5 fingers up - they counted everything.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Look at this spill.
- With your partner, whisper: how many RED? How many YELLOW?
- Then write on your whiteboard: 5 = ___ and ___.

DO:
- Display 5 counters: 1 red, 4 yellow.
- Allow 30 seconds for partner whisper and whiteboard write.
- Walk and scan whiteboards.

TEACHER NOTES:
This is a different split from the I Do. We Do uses 1 and 4 to extend the same pattern. The whole stays 5.

WATCH FOR:
- Pairs who write 5 = 1 and 4 - secure.
- Pairs who write 5 = 5 (counted total only) - prompt: "What about each colour?"

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check. Whole 5. Parts 1 and 4.
- Same whole. Different parts again!

DO:
- Click to reveal.
- Repeat the pattern with the class.

TEACHER NOTES:
Three ways shown so far: 4 and 1, 3 and 2, 1 and 4. Notice the connections in students' eyes.

WATCH FOR:
- Students who say "I see a pattern" - extend with: "Can you find another way?"
- Students who needed prompting - small group focus.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge time. Listen carefully.
- I say: "5 is 2 and 4."
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Say the false claim aloud.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "5 is 2 and 4. Thumbs up if yes, thumbs down if no."
- Scan for: thumbs DOWN. 2 and 4 makes 6, not 5.
PROCEED: If 80% show thumbs down, click to reveal.
PIVOT: Most likely misconception - students agree without checking.
- Reteach: "Hold up 2 fingers. Now hold up 4 more. Count them all. How many?"
- Re-check: "Is the whole 5 or 6?"

TEACHER NOTES:
This hinge probes whether students check the whole when given parts. A correct thumbs-down means they are tracking both.

WATCH FOR:
- Confident thumbs down with quick fingers - secure.
- Slow or mixed thumbs - reteach and re-check.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Get 5 counters and your mat.
- Make a split. Tell your partner.
- Then make a different split. Tell your partner again.

DO:
- Distribute 5 two-sided counters per child and a printed mat.
- Circulate. Listen for "whole 5, parts ___ and ___".
- Cold call 2 students for their splits.

TEACHER NOTES:
First time finding multiple ways independently. Acceptable splits for whole = 5: 5 and 0, 4 and 1, 3 and 2, 2 and 3, 1 and 4, 0 and 5. Order matters less than naming both parts and the whole.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 4 counters first. Find one split for whole 4.
EXTENDING PROMPT:
- Task: Find ALL the ways to make 5. How many ways are there?

WATCH FOR:
- Students who find 2 splits and use the language - on track.
- Students who repeat the same split - prompt: "What if one moved?"

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. Look at the picture.
- On your whiteboard, write the whole and the parts.

DO:
- Display: 3 red and 2 yellow counters.
- Allow 60 seconds.
- Collect or photograph whiteboards.

TEACHER NOTES:
Exit ticket assesses SC2 - the core target. Whole 5, parts 3 and 2.

WATCH FOR:
- Students who write whole 5, parts 3 and 2 - on track.
- Students who write only 5 - prompt language for next lesson's reteach.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today we found different ways to make 5.
- The whole stayed 5. The parts changed.
- Show me thumbs: can you find one way to make 5?

DO:
- Read the success criteria.
- Use thumbs up sideways down.
- Quick partner share for the reflection.

TEACHER NOTES:
Most students should reach SC2 today. SC3 (explain the whole stays the same) is the stretch.

WATCH FOR:
- Strong thumbs up - secure.
- Sideways/down - small group tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Today's recording sheet is for your table time.
- You will draw 3 ways to make 5.

DO:
- Print one sheet per student.
- Have 5 two-sided counters and a cup at each table.

TEACHER NOTES:
One core student resource today. The answer key is for the teacher.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Make 5",
    `Foundation Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slides 2-3: Daily Review with reveal — Counting fluency
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "What Comes After 13?", { color: C.ACCENT });

      // Number track 1-20
      const trackY = CONTENT_TOP + 0.30;
      addNumberTrack(s, 0.5, trackY, 9.0, 1, 10, [], { cellH: 0.85, fontSize: 32 });
      addNumberTrack(s, 0.5, trackY + 1.05, 9.0, 11, 20, [13], { cellH: 0.85, fontSize: 32 });

      addCard(s, 1.5, trackY + 2.20, 7.0, 0.85, { strip: C.PRIMARY });
      s.addText("Whisper count 1 to 20. What comes after 13?", {
        x: 1.7, y: trackY + 2.22, w: 6.6, h: 0.80,
        fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "14 comes after 13.", {
        x: 2.0, y: 4.45, w: 6.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Backwards counting from 20
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Count Backwards From 20", { color: C.ACCENT });

  const trackY2 = CONTENT_TOP + 0.30;
  addNumberTrack(sFluency, 0.5, trackY2, 9.0, 11, 20, [], { cellH: 0.85, fontSize: 32 });
  addNumberTrack(sFluency, 0.5, trackY2 + 1.05, 9.0, 1, 10, [], { cellH: 0.85, fontSize: 32 });

  addCard(sFluency, 1.5, trackY2 + 2.20, 7.0, 0.85, { strip: C.PRIMARY });
  sFluency.addText("Stand up. Count down: 20, 19, 18...", {
    x: 1.7, y: trackY2 + 2.22, w: 6.6, h: 0.80,
    fontSize: 24, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to find different ways to make 5.",
    [
      "I can make 5 with two parts.",
      "I can find different ways to make 5.",
      "I can say the whole stays the same when the parts change.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do 1 — One way to make 5 (4 and 1)
  workedExSlide(pres, 2, "I Do", "One Way to Make 5",
    [
      "I have 5 counters.",
      "5 is the whole.",
      "",
      "I shake and spill.",
      "",
      "4 are red.",
      "1 is yellow.",
      "",
      "Whole 5.",
      "Parts 4 and 1.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 5,
        partA: 4, partB: 1,
      });
    }
  );

  // Slide 7: I Do 2 — Same whole, different parts (3 and 2)
  workedExSlide(pres, 2, "I Do", "Same Whole. Different Parts.",
    [
      "Same 5 counters.",
      "Whole is still 5.",
      "",
      "Shake again.",
      "",
      "3 are red.",
      "2 are yellow.",
      "",
      "Whole 5.",
      "Parts 3 and 2.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 5,
        partA: 3, partB: 2,
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

      // 5 counters: 2 red, 3 yellow
      drawCounterRow(s, 1.4, CONTENT_TOP + 0.45, 5, 2, { size: 1.05, gap: 0.25 });

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
      addTextOnShape(slide, "3 yellow.   Whole 5.   Parts 2 and 3.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 10-11: We Do with reveal — 5 = 1 and 4
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Find the Parts", { color: STAGE_COLORS["3"] });

      // Left card: counter visual
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      drawCounterRow(s, 0.85, CONTENT_TOP + 0.55, 5, 1, { size: 0.65, gap: 0.10 });

      s.addText("Whole = 5", {
        x: 0.6, y: CONTENT_TOP + 1.45, w: 4.3, h: 0.50,
        fontSize: 28, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      s.addText("How many RED? How many YELLOW?", {
        x: 0.6, y: CONTENT_TOP + 2.10, w: 4.3, h: 0.50,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
      });

      // Right card: prompt
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "5 = ____ and ____", options: { fontSize: 28, bold: true, color: C.CHARCOAL } },
      ], {
        x: 5.4, y: CONTENT_TOP + 0.35, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.55,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Whole 5.   Parts 1 and 4.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 12-13: CFU Hinge with reveal — 5 is 2 and 4 (false)
  withReveal(
    () => cfuSlide(pres, "CFU", "Is This Right?", "Thumbs Up or Thumbs Down",
      "I say:\n5 is 2 and 4.\n\nIs that right?",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN!  2 and 4 make 6, not 5.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Find Two Ways to Make 5",
    [
      "First: Get 5 counters.",
      "",
      "Next: Make a split on the mat.",
      "",
      "Then: Find a different way.",
      "",
      "Whole 5.",
      "Parts ___ and ___.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.20, { strip: C.ALERT });
      slide.addText("You need:", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.40, h: 0.36,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "5 two-sided counters", options: { bullet: true, fontSize: 17, color: C.CHARCOAL, breakLine: true } },
        { text: "your PPW mat", options: { bullet: true, fontSize: 17, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.40, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.50, h: 0.70,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      drawPpwMat(slide, lg.rightX, lg.panelTopPadded + 1.35, lg.rightW, 2.30, {
        wholeNumber: 5,
        partA: null, partB: null,
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Whole 5, parts 3 and 2. Draw it on your whiteboard.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 16: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Tell your partner one way to make 5.",
      scItems: [
        "I can make 5 with two parts.",
        "I can find different ways to make 5.",
        "I can say the whole stays the same when the parts change.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 17: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPWi_Lesson2_Make_5.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find three ways to make 5.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addTipBox(doc, "Shake your cup. Spill the counters. Colour the dots. Write the parts.", y, { color: C.ACCENT });

    function drawShakeRow(yTop, label) {
      doc.fontSize(16).font("Sans-Bold").fillColor("#" + C.PRIMARY).text(label, 60, yTop);
      const dotY = yTop + 28;
      // 5 empty circles
      const size = 32;
      const gap = 12;
      for (let i = 0; i < 5; i += 1) {
        doc.circle(80 + i * (size + gap) + size / 2, dotY + size / 2, size / 2)
          .lineWidth(1.5).stroke("#333333");
      }
      doc.fontSize(14).font("Sans").fillColor("#333333").text("Whole 5.   Parts ____ and ____", 60, dotY + size + 14);
      return dotY + size + 50;
    }

    y = drawShakeRow(y, "Shake 1");
    y = drawShakeRow(y, "Shake 2");
    y = drawShakeRow(y, "Shake 3");

    y = addTipBox(doc, "Are any two shakes the same? Are they all different?", y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Lesson 2 Ways to Make 5 | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "All the ways to make 5.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addBodyText(doc, "There are six ways to make 5 with two parts:", y);
    y = addSectionHeading(doc, "Whole = 5", y, { color: C.PRIMARY });
    y = addBodyText(doc, "5 and 0   |   4 and 1   |   3 and 2   |   2 and 3   |   1 and 4   |   0 and 5", y);
    y = addTipBox(doc, "Order is not the focus today. Children may swap red and yellow without realising. Accept any spill that adds to 5.", y, { color: C.ACCENT });
    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
