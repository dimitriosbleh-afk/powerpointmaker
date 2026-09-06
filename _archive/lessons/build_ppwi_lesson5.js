"use strict";

// Part, Part, Whole Introduction Unit — Lesson 5: Friends of 10
// Foundation Numeracy | Lesson 5 of 10 | Variant 0
// VC2MFN04
// Daily Review: Teen Numbers - name 16-20 from tens frames
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
  addTensFrame,
  STAGE_COLORS,
} = T;

const SESSION = 5;
const TOTAL = 10;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = `Part, Part, Whole | Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`;
const OUT_DIR = "output/PPWi_Lesson5_Friends_of_10";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 5 Friends of 10", "Find the friends of 10 - all the ways to make 10.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 5 Answer Key", "All friends of 10.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

function drawTwoColourTensFrame(slide, x, y, w, redCount, yellowCount) {
  const cellW = w / 5;
  const cellH = cellW;
  const dotRatio = 0.62;
  const total = redCount + yellowCount;
  for (let row = 0; row < 2; row += 1) {
    for (let col = 0; col < 5; col += 1) {
      const idx = row * 5 + col;
      const cx = x + col * cellW;
      const cy = y + row * cellH;
      slide.addShape("rect", {
        x: cx, y: cy, w: cellW, h: cellH,
        fill: { color: C.WHITE },
        line: { color: C.CHARCOAL, width: 1.5 },
      });
      if (idx < total) {
        const dotW = cellW * dotRatio;
        const dotH = cellH * dotRatio;
        const color = idx < redCount ? "D64545" : "F4C430";
        slide.addShape("roundRect", {
          x: cx + (cellW - dotW) / 2,
          y: cy + (cellH - dotH) / 2,
          w: dotW, h: dotH,
          rectRadius: Math.min(dotW, dotH) / 2,
          fill: { color },
          line: { color: C.CHARCOAL, width: 0.6 },
        });
      }
    }
  }
}

function drawPpwMatWithTensFrame(slide, x, y, w, h, opts) {
  const o = opts || {};
  const wholeH = h * 0.32;

  slide.addShape("roundRect", {
    x, y, w, h: wholeH, rectRadius: 0.10,
    fill: { color: o.wholeFill || C.PRIMARY },
    line: { color: o.wholeFill || C.PRIMARY, width: 1 },
  });
  slide.addText("Whole = " + o.wholeNumber, {
    x, y, w, h: wholeH,
    fontSize: 30, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
    fit: "shrink", shrinkText: true,
  });

  const frameY = y + wholeH + 0.20;
  const frameW = Math.min(w, 3.6);
  const frameX = x + (w - frameW) / 2;
  drawTwoColourTensFrame(slide, frameX, frameY, frameW, o.partA || 0, o.partB || 0);

  const labelY = frameY + (frameW / 5) * 2 + 0.18;
  slide.addText(`Parts ${o.partA} and ${o.partB}`, {
    x, y: labelY, w, h: 0.40,
    fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today is a special day. We meet the friends of 10.
- The friends of 10 are pairs of numbers that make 10 together.
- 10 is the whole.

DO:
- Have a magnetic ten frame and 10 counters ready.
- Show the title slide.

TEACHER NOTES:
Lesson 5 of 5. The grand finale of the unit. "Friends of 10" is the key knowledge for later addition and subtraction strategies.

WATCH FOR:
- Students excited about 10 - the round number is engaging.
- Students worried about a bigger number - the ten frame structures it.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Look at the ten frame.
- Count the dots quietly.
- Show me the number on your fingers.

DO:
- Display ten frame showing 17 (top row 10, bottom row 7).
- Wait 5 seconds.
- "Show me... ready... show."

TEACHER NOTES:
Daily Review: name 16-20 from a ten frame. Today is 17.

WATCH FOR:
- Students who say "10 and 7" - secure structure.
- Students still counting one-by-one - reteach.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- 17! Ten and seven.
- The full top row is 10. Seven more.
- Tick if you said 17.

DO:
- Click to reveal 17.
- Point: full row is 10. Then 7 more.

TEACHER NOTES:
Tick and fix.

WATCH FOR:
- Students who match dots to fingers - secure.
- Students still counting - small group.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency. Today we count from different starting numbers.
- Start at 7. Count to 20.
- 7, 8, 9...

DO:
- Lead choral count from 7 to 20.
- Repeat from 13 to 20.
- Repeat from 16 to 20.

TEACHER NOTES:
Counting from various starting points - especially from a teen number to 20. This builds flexible number-sequence knowledge.

WATCH FOR:
- Students who count smoothly from any start - very secure.
- Students who restart from 1 each time - prompt: "Start from 7. Just go up."

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Today we meet the friends of 10.
- Friends of 10 are two numbers that make 10 together.
- Read the success criteria with me.

DO:
- Choral read LI and SC.
- Hold up 10 fingers.
- Hide 3 fingers behind your back. Show 7. Say: "7 and 3 are friends of 10!"

TEACHER NOTES:
SC1: at least one friend pair. SC2: more than one. SC3: explain that the friends always add to 10.

WATCH FOR:
- Students saying "5 and 5" or "9 and 1" - secure.
- Students hesitating - we will build them through I Do.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. The whole is 10.
- I put 5 red counters in the top row of the ten frame.
- I put 5 yellow counters in the bottom row.
- 5 and 5 is 10!
- 5 and 5 are friends of 10.

DO:
- Use a magnetic ten frame.
- Place 5 red across the top row.
- Place 5 yellow across the bottom row.
- Both rows full.

TEACHER NOTES:
The cleanest friend of 10. Use this as the anchor.

WATCH FOR:
- Students who say "5 and 5 makes 10" - secure.
- Students who count from 1 to 10 - prompt: "Each row is 5."

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now another friend pair.
- I put 7 red counters first. The top row is full plus 2 more in the bottom.
- I put 3 yellow counters in the rest of the bottom row.
- 7 and 3 are friends of 10.
- 7 and 3 is 10.

DO:
- Place 7 red counters in the ten frame.
- Place 3 yellow counters in the rest of the bottom row.
- All 10 cells filled.

TEACHER NOTES:
Show another friend pair: 7 and 3. The ten frame is now full. The structure helps students see how big each part is.

MISCONCEPTIONS:
- Misconception: Students think the friends of 10 must be 5 and 5.
  Why: 5 and 5 is the most familiar pair.
  Impact: They miss other friend pairs.
  Quick correction: "There are many friend pairs. 7 and 3 also make 10."

WATCH FOR:
- Students who say more pairs without prompting - secure.
- Students who only know 5 and 5 - we will find more in We Do.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Look at the ten frame.
- One part is red. The empty cells are the other part.
- How many EMPTY cells? Show me on your fingers.

DO:
- Display ten frame: 6 red, 0 yellow (4 cells empty).
- Wait 5 seconds.
- "Show me... ready... show."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "How many empty cells? Show me."
- Scan for: 4 fingers.
PROCEED: If 80% show 4, click to reveal.
PIVOT: Most likely misconception - students count the red dots.
- Reteach: "Count just the empty cells. Point to each one."
- Re-check: "How many empty?"

TEACHER NOTES:
This builds the friends-of-10 idea: red part + empty part = 10. The empty cells "are" the missing part.

WATCH FOR:
- 4 fingers up - secure.
- 6 fingers up - they counted the red.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Look at the ten frame.
- 8 red dots. How many MORE to make 10?
- Whisper to your partner.
- Show me on your fingers.

DO:
- Display ten frame: 8 red, 2 empty cells.
- Allow 30 seconds for partner whisper and finger show.
- Walk and scan.

TEACHER NOTES:
We Do extends to: how many more to make 10? This is the foundation of "addition to 10" later.

WATCH FOR:
- Pairs who say "2 more" - secure.
- Pairs who count all 10 - prompt: "Just count the empty cells."

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- 2 more! 8 and 2 make 10.
- 8 and 2 are friends of 10.

DO:
- Click to reveal.
- Repeat: "8 and 2 are friends of 10."

TEACHER NOTES:
Now students have seen 5 and 5, 7 and 3, and 8 and 2.

WATCH FOR:
- Students who notice the pattern - extend with: "Can you find another friend pair?"
- Students still hesitant - small group.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge time. Listen carefully.
- I say: "6 and 3 are friends of 10."
- Is that right? Thumbs up or thumbs down.

DO:
- Say the false claim aloud.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "6 and 3 are friends of 10. Thumbs up if yes, thumbs down if no."
- Scan for: thumbs DOWN. 6 and 3 makes 9, not 10.
PROCEED: If 80% show thumbs down, click to reveal.
PIVOT: Most likely misconception - students agree without checking.
- Reteach: "Hold up 6 fingers. Now 3 more. Count. How many?"
- Re-check: "Is 6 and 3 a friend of 10?"

TEACHER NOTES:
The hinge probes whether students check that the parts add to 10. A "no" answer means they are tracking the whole.

WATCH FOR:
- Confident thumbs down - secure.
- Mixed thumbs - reteach with fingers and a ten frame.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn.
- Get 10 counters and your ten frame.
- Find a friend pair. Tell your partner.
- Then find a different friend pair.

DO:
- Distribute 10 two-sided counters and a printed ten frame per child.
- Circulate. Listen for "___ and ___ are friends of 10".
- Cold call 2 students.

TEACHER NOTES:
Today the goal is at least two friend pairs. The bigger goal is recognising the structure.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 7 counters. Find one way to make 7 (revision).
EXTENDING PROMPT:
- Task: Find ALL the friends of 10. List them in order.

WATCH FOR:
- Students who find more than 2 - secure.
- Students who repeat the same pair - prompt: "What if one moved to the other colour?"

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. The ten frame has some red dots.
- Write the friend pair on your whiteboard.
- Whole 10. Parts ___ and ___.

DO:
- Display ten frame: 7 red, 3 empty cells.
- Allow 60 seconds.
- Collect or photograph whiteboards.

TEACHER NOTES:
Exit ticket: friend pair 7 and 3. Assesses whether students can name a friend pair from a ten frame.

WATCH FOR:
- Students who write whole 10, parts 7 and 3 - secure.
- Students stuck - reteach with fingers next lesson.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today we met the friends of 10.
- The friends of 10 are pairs that make 10 together.
- Show me thumbs: can you name one friend of 10?

DO:
- Read SC.
- Use thumbs up sideways down.
- Quick partner share.

TEACHER NOTES:
End-of-unit reflection. Most students should reach SC2. The friends of 10 will return many times this year.

WATCH FOR:
- Strong thumbs up - secure foundations.
- Sideways/down - revisit with small group next week.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Today's recording sheet has space for friend pairs.

DO:
- Print one sheet per student.
- Have 10 two-sided counters and a printed ten frame at each table.

TEACHER NOTES:
End-of-unit recording sheet. Students record the friend pairs they find.

[General: Resources]`;

const NOTES_LAUNCH = `SAY:
- Yesterday we made 8 and 9 on the ten frame.
- Look at this ten frame. It has 9 counters. The whole is 9. One cell is empty.
- Today our whole is 10. The ten frame is full.
- Friends of 10 are two parts that fill the whole ten frame.

DO:
- Display a ten frame with 9 counters.
- Point to the 9 filled cells, then to the 1 empty cell.
- Add the 10th counter on the camera or with finger trace.
- Say "10. The whole is 10. The ten frame is full."

TEACHER NOTES:
The launch bridges from yesterday (9) to today (10) and names the special idea: friends of 10. Keep it under 90 seconds. Today is the final lesson of the unit, so the bridge is also a celebration of progress.

WATCH FOR:
- Students who say "1 more makes 10" - the friend-pair language is forming.
- Students who count all 10 - that is fine, gently echo "the whole is 10".

[General: Launch | VTLM 2.0: Activate Prior Knowledge]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 5: Friends of 10",
    `Foundation Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch — bridge from yesterday's 9 to today's whole 10
  contentSlide(pres, "Launch", C.SUCCESS, "From 9 to 10",
    [
      "Yesterday: we made 8 and 9.",
      "This ten frame shows 9. One cell is empty.",
      "Today our whole is 10. The ten frame is full.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      // Ten frame with 9 counters — visible empty cell signals the bridge to 10
      addTensFrame(slide, lg.rightX + 0.10, lg.panelTopPadded + 0.55,
        lg.rightW - 0.20, 9, { fillColor: "D64545" });
      slide.addText("Whole = 9", {
        x: lg.rightX, y: lg.panelTopPadded + 2.50, w: lg.rightW, h: 0.42,
        fontSize: 22, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slides 2-3: Daily Review with reveal — Teen number from ten frame (17)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "What Number?", { color: C.ACCENT });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

      // Full ten frame on left (10) + 7 dots on right
      addTensFrame(s, 1.0, CONTENT_TOP + 0.45, 4.0, 10, { fillColor: C.PRIMARY });

      // 7 dots arranged 4-3 to the right
      const extraX = 6.0;
      const extraY = CONTENT_TOP + 0.45;
      const dotSize = 0.55;
      const dotGap = 0.18;
      // row 1: 4 dots
      for (let i = 0; i < 4; i += 1) {
        s.addShape("roundRect", {
          x: extraX + i * (dotSize + dotGap),
          y: extraY + 0.20,
          w: dotSize, h: dotSize, rectRadius: dotSize / 2,
          fill: { color: C.PRIMARY },
          line: { color: C.CHARCOAL, width: 0.8 },
        });
      }
      // row 2: 3 dots
      for (let i = 0; i < 3; i += 1) {
        s.addShape("roundRect", {
          x: extraX + i * (dotSize + dotGap) + (dotSize + dotGap) / 2,
          y: extraY + 0.20 + dotSize + dotGap,
          w: dotSize, h: dotSize, rectRadius: dotSize / 2,
          fill: { color: C.PRIMARY },
          line: { color: C.CHARCOAL, width: 0.8 },
        });
      }

      s.addText("Show me on your fingers.", {
        x: 0.7, y: CONTENT_TOP + 2.50, w: 8.6, h: 0.40,
        fontSize: 22, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "17!  10 and 7 more.", {
        x: 1.5, y: 4.45, w: 7.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Counting from various starting points
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Count From Here to 20", { color: C.ACCENT });

  // Three big start-number cards
  const startNums = [7, 13, 16];
  const sCardY = CONTENT_TOP + 0.40;
  const sCardW = 2.5;
  const sCardH = 2.3;
  const sGap = 0.40;
  const sTotalW = sCardW * 3 + sGap * 2;
  const sStartX = (10 - sTotalW) / 2;
  startNums.forEach((n, i) => {
    const cx = sStartX + i * (sCardW + sGap);
    sFluency.addShape("roundRect", {
      x: cx, y: sCardY, w: sCardW, h: sCardH, rectRadius: 0.12,
      fill: { color: C.PRIMARY },
    });
    sFluency.addText(String(n), {
      x: cx, y: sCardY, w: sCardW, h: sCardH,
      fontSize: 90, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    sFluency.addText("Start here", {
      x: cx, y: sCardY + sCardH + 0.05, w: sCardW, h: 0.35,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      align: "center", valign: "top", margin: 0,
    });
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning the friends of 10.",
    [
      "I can find one friend pair of 10.",
      "I can find more than one friend pair of 10.",
      "I can show friend pairs in a ten frame.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do 1 — 5 and 5
  workedExSlide(pres, 2, "I Do", "5 and 5 are Friends of 10",
    [
      "Whole is 10.",
      "",
      "5 red in the top row.",
      "",
      "5 yellow in the bottom row.",
      "",
      "5 and 5 is 10.",
      "",
      "5 and 5 are friends of 10.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      drawPpwMatWithTensFrame(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 10,
        partA: 5, partB: 5,
      });
    }
  );

  // Slide 7: I Do 2 — 7 and 3
  workedExSlide(pres, 2, "I Do", "7 and 3 are Friends of 10",
    [
      "Whole is 10.",
      "",
      "7 red.",
      "(Top row full plus 2 more.)",
      "",
      "3 yellow.",
      "",
      "7 and 3 is 10.",
      "",
      "7 and 3 are friends of 10.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      drawPpwMatWithTensFrame(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 10,
        partA: 7, partB: 3,
      });
    }
  );

  // Slides 8-9: CFU 1 with reveal — How many empty cells?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How Many Empty Cells?", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Ten frame: 6 red, 4 empty
      drawTwoColourTensFrame(s, 2.5, CONTENT_TOP + 0.35, 5.0, 6, 0);

      s.addText("Show me on your fingers.", {
        x: 0.7, y: CONTENT_TOP + 2.40, w: 8.6, h: 0.45,
        fontSize: 24, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "4 empty.   6 and 4 are friends of 10.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 10-11: We Do with reveal — 8 and ___ make 10
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "How Many More to Make 10?", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      // Ten frame: 8 red
      drawTwoColourTensFrame(s, 0.85, CONTENT_TOP + 0.50, 3.8, 8, 0);

      s.addText("8 dots so far", {
        x: 0.6, y: CONTENT_TOP + 2.30, w: 4.3, h: 0.50,
        fontSize: 26, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "How many MORE", options: { fontSize: 24, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "to make 10?", options: { fontSize: 24, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 12, breakLine: true } },
        { text: "Show me on your fingers.", options: { fontSize: 18, color: C.CHARCOAL } },
      ], {
        x: 5.4, y: CONTENT_TOP + 0.40, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.55,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "2 more!   8 and 2 are friends of 10.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 12-13: CFU Hinge with reveal — 6 and 3 are friends of 10? (false)
  withReveal(
    () => cfuSlide(pres, "CFU", "Are They Friends of 10?", "Thumbs Up or Thumbs Down",
      "I say:\n6 and 3 are friends of 10.\n\nIs that right?",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN!  6 and 3 make 9, not 10.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Find Friend Pairs of 10",
    [
      "First: Get 10 counters.",
      "",
      "Next: Make a friend pair.",
      "Tell your partner.",
      "",
      "Then: Find a different friend pair.",
      "",
      "___ and ___ are friends of 10.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.20, { strip: C.ALERT });
      slide.addText("You need:", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.40, h: 0.36,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "10 two-sided counters", options: { bullet: true, fontSize: 17, color: C.CHARCOAL, breakLine: true } },
        { text: "your ten frame", options: { bullet: true, fontSize: 17, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.40, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.50, h: 0.70,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Empty ten frame
      drawTwoColourTensFrame(slide, lg.rightX + 0.20, lg.panelTopPadded + 1.65, lg.rightW - 0.40, 0, 0);
      slide.addText("Whole = 10", {
        x: lg.rightX, y: lg.panelTopPadded + 3.10, w: lg.rightW, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "The ten frame has 7 red dots. Write the friend pair: ___ and ___ are friends of 10.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 16: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Tell your partner one friend of 10.",
      scItems: [
        "I can find one friend pair of 10.",
        "I can find more than one friend pair of 10.",
        "I can show friend pairs in a ten frame.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPWi_Lesson5_Friends_of_10.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find friend pairs of 10.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addTipBox(doc, "Colour the dots in two colours. Write the friend pair below.", y, { color: C.ACCENT });

    function drawTenFramePdf(yTop) {
      const x0 = 90;
      const cell = 50;
      for (let r = 0; r < 2; r += 1) {
        for (let c = 0; c < 5; c += 1) {
          doc.lineWidth(1.5).rect(x0 + c * cell, yTop + r * cell, cell, cell).stroke("#333333");
          doc.circle(x0 + c * cell + cell / 2, yTop + r * cell + cell / 2, cell * 0.32)
            .lineWidth(1).stroke("#999999");
        }
      }
      return yTop + 2 * cell;
    }

    y = addSectionHeading(doc, "Friend Pair 1", y, { color: C.PRIMARY });
    y = drawTenFramePdf(y) + 14;
    doc.fontSize(14).font("Sans").fillColor("#333333").text("____ and ____ are friends of 10.", 60, y);
    y += 36;

    y = addSectionHeading(doc, "Friend Pair 2", y, { color: C.PRIMARY });
    y = drawTenFramePdf(y) + 14;
    doc.fontSize(14).font("Sans").fillColor("#333333").text("____ and ____ are friends of 10.", 60, y);
    y += 36;

    y = addSectionHeading(doc, "Friend Pair 3", y, { color: C.PRIMARY });
    y = drawTenFramePdf(y) + 14;
    doc.fontSize(14).font("Sans").fillColor("#333333").text("____ and ____ are friends of 10.", 60, y);
    y += 36;

    addPdfFooter(doc, `Lesson ${SESSION} | Lesson 5 Friends of 10 | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "All friends of 10.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addBodyText(doc, "There are eleven pairs that make 10:", y);
    y = addSectionHeading(doc, "Friends of 10", y, { color: C.PRIMARY });
    y = addBodyText(doc, "10 and 0", y);
    y = addBodyText(doc, "9 and 1", y);
    y = addBodyText(doc, "8 and 2", y);
    y = addBodyText(doc, "7 and 3", y);
    y = addBodyText(doc, "6 and 4", y);
    y = addBodyText(doc, "5 and 5", y);
    y = addBodyText(doc, "4 and 6", y);
    y = addBodyText(doc, "3 and 7", y);
    y = addBodyText(doc, "2 and 8", y);
    y = addBodyText(doc, "1 and 9", y);
    y = addBodyText(doc, "0 and 10", y);
    y = addTipBox(doc, "Friends of 10 will return many times this year. Strong recall of these pairs underpins later addition and subtraction.", y, { color: C.ACCENT });
    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 5 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
