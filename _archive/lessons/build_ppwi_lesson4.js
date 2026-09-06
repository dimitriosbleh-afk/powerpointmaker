"use strict";

// Part, Part, Whole Introduction Unit — Lesson 4: Make 8 and 9 (Ten Frame)
// Foundation Numeracy | Lesson 4 of 10 | Variant 0
// VC2MFN04
// Daily Review: Teen Numbers - name 11-15 from tens frames
// Fluency: What number comes before and after

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

const SESSION = 4;
const TOTAL = 10;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = `Part, Part, Whole | Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`;
const OUT_DIR = "output/PPWi_Lesson4_Make_8_and_9";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 4 Ten Frame Parts", "Find the parts of 8 and 9 using a ten frame.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 4 Answer Key", "Sample part splits for 8 and 9.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Two-colour ten frame: first `redCount` cells are red, next `yellowCount` are yellow,
// remainder are empty. Used to show parts within a whole.
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

  // Two-colour ten frame underneath
  const frameY = y + wholeH + 0.20;
  const frameW = Math.min(w, 3.6);
  const frameX = x + (w - frameW) / 2;
  drawTwoColourTensFrame(slide, frameX, frameY, frameW, o.partA || 0, o.partB || 0);

  // Parts label below
  const labelY = frameY + (frameW / 5) * 2 + 0.18;
  slide.addText(`Parts ${o.partA} and ${o.partB}`, {
    x, y: labelY, w, h: 0.40,
    fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today our wholes are 8 and 9.
- We will use the ten frame to show the parts.
- A ten frame has 10 spaces. We can see 8 or 9 inside it.

DO:
- Have a magnetic ten frame and 9 counters ready on the board.
- Show the title slide.

TEACHER NOTES:
Lesson 4 of 5. Today the ten frame becomes the main visual. The PPW mat continues but the ten frame is now embedded.

WATCH FOR:
- Students who already know the ten frame from earlier work - secure.
- Students new to the ten frame - move slowly with the visual.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Look at the ten frame.
- Count the dots quietly to yourself.
- Show me the number with your fingers.

DO:
- Display tens frame showing 13 (top row full, bottom row 3).
- Wait 5 seconds.
- "Show me on your fingers... show."

TEACHER NOTES:
Daily Review: name a teen number from a ten frame plus extras. Most students recognise the full top row as 10.

WATCH FOR:
- Students who say "10 and 3" - they see the structure.
- Students who count every dot - still building.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- 13! Ten and three.
- The full top row is 10.
- Three more in the bottom row.
- 10 and 3 is 13.

DO:
- Click to reveal 13.
- Point to the full top row, then the 3 dots.

TEACHER NOTES:
This builds the place-value foundation: teen numbers are 10 plus some more.

WATCH FOR:
- Students who say "10 and 3" - secure place-value start.
- Students still counting one-by-one - reteach the full row as 10.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. I will say a number.
- You tell me the number that comes BEFORE.
- Then the number that comes AFTER.

DO:
- Display the number 15 large.
- Cold call 2 students for "before" (14).
- Cold call 2 students for "after" (16).

TEACHER NOTES:
What comes before, what comes after. The number 15 sits between 14 and 16.

WATCH FOR:
- Students who answer instantly - secure number-line knowledge.
- Students who hesitate - draw a small number track on the board.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Today our wholes are 8 and 9.
- We will use the ten frame to find the parts.
- Read the success criteria with me.

DO:
- Choral read LI and SC.
- Hold up 8 fingers, then 9 fingers.

TEACHER NOTES:
SC1: at least one way for whole 8 or 9. SC2: ways for both. SC3: explain that the whole stays the same when parts change.

WATCH FOR:
- Students who already know one way - secure.
- Students new to the bigger numbers - the ten frame helps them see the structure.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. The whole is 8.
- I put 5 red counters in the top row of the ten frame.
- I put 3 yellow counters in the bottom row.
- 5 red and 3 yellow is 8.
- Whole 8. Parts 5 and 3.

DO:
- Use a magnetic ten frame on the board.
- Place 5 red counters in the top row first.
- Then 3 yellow counters in the bottom row.
- Touch each part as you say its number.

TEACHER NOTES:
The full top row is 5. Adding 3 more is 8. This connects subitising (5 = full top row) to PPW.

MISCONCEPTIONS:
- Misconception: Students count every dot from 1 each time.
  Why: They are still building cardinality.
  Impact: They miss the structure of the ten frame.
  Quick correction: "The full top row is 5. We do not need to count it. 5 and 3 more is 8."

WATCH FOR:
- Students who say "5 and 3 is 8" without counting - secure.
- Students who count one-by-one - prompt to use the row of 5.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now whole 9. I have 9 counters.
- I put 4 red in the top row.
- I put 5 yellow.
- The yellow goes 1 in the top row, then 4 in the bottom row.
- 4 red and 5 yellow is 9.
- Whole 9. Parts 4 and 5.

DO:
- Place 4 red counters in the top row.
- Place 5 yellow counters - 1 in the top row to fill it, then 4 in the bottom row.
- Point: "Top row is full now - that is 5. Then 4 more in the bottom."

TEACHER NOTES:
9 in the ten frame shows one empty cell. Students see how close 9 is to 10.

WATCH FOR:
- Students who say "1 less than 10" - very secure.
- Students who try to count every dot - prompt: "The top row is 5."

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Look at the ten frame.
- How many YELLOW dots? Show me on your fingers.

DO:
- Display ten frame with 5 red, 3 yellow (whole 8).
- Wait 5 seconds.
- "Show me... ready... show."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "How many yellow? Show me."
- Scan for: 3 fingers.
PROCEED: If 80% show 3, click to reveal.
PIVOT: Most likely misconception - students count all dots.
- Reteach: "Just yellow. Point to only yellow."
- Re-check: Same picture, ask "How many red?"

TEACHER NOTES:
Standard finger-voting routine.

WATCH FOR:
- 3 fingers up - secure.
- 8 fingers up - they counted everything.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Look at the ten frame.
- With your partner, whisper: how many RED? How many YELLOW?
- Write on your whiteboard: 9 = ___ and ___.

DO:
- Display ten frame: 6 red, 3 yellow (whole 9).
- Allow 30 seconds.
- Walk and scan whiteboards.

TEACHER NOTES:
We Do uses 6 and 3 as the split for whole 9.

WATCH FOR:
- Pairs who write 9 = 6 and 3 - secure.
- Pairs stuck - prompt: "Top row first."

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Whole 9. Parts 6 and 3.
- Same whole, different parts.

DO:
- Click to reveal.
- Repeat the pattern with the class.

TEACHER NOTES:
Build the connection: 9 has many ways - we just saw 4 and 5, and now 6 and 3.

WATCH FOR:
- Confident response - secure.
- Mixed response - small group focus.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge time. Listen carefully.
- I say: "8 is 5 and 4."
- Is that right? Thumbs up or thumbs down.

DO:
- Say the false claim aloud.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "8 is 5 and 4. Thumbs up if yes, thumbs down if no."
- Scan for: thumbs DOWN. 5 and 4 makes 9, not 8.
PROCEED: If 80% show thumbs down, click to reveal.
PIVOT: Most likely misconception - students agree without checking.
- Reteach: "Hold up 5 fingers. Now 4 more. Count them. How many?"
- Re-check: "Is the whole 8 or 9?"

TEACHER NOTES:
Hinge probes part-whole tracking with a near-miss (one off).

WATCH FOR:
- Confident thumbs down - secure.
- Mixed - reteach with fingers.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn.
- Get 8 counters and a printed ten frame.
- Find one way to make 8.
- Then find one way to make 9.

DO:
- Distribute 9 two-sided counters and a printed ten frame per child.
- Circulate. Listen for the language pattern.
- Cold call 2 students for each whole.

TEACHER NOTES:
Today the ten frame is the main tool. Students work with both 8 and 9.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 6 counters first. Find one way to make 6 (revision from yesterday).
EXTENDING PROMPT:
- Task: Find ALL the ways to make 8. Draw each one.

WATCH FOR:
- Students who fill the top row first - using structure.
- Students who scatter counters - prompt: "Fill the top row first."

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. Look at the ten frame.
- Write the whole and the parts on your whiteboard.

DO:
- Display ten frame: 5 red, 3 yellow (whole 8).
- Allow 60 seconds.
- Collect or photograph whiteboards.

TEACHER NOTES:
Exit ticket: whole 8, parts 5 and 3.

WATCH FOR:
- Students who write whole 8, parts 5 and 3 - secure.
- Students stuck - small group at start of Lesson 5.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today our wholes were 8 and 9.
- We used a ten frame to see the parts.
- Show me thumbs: can you find one way to make 8?

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
- Today's recording sheet has ten frames for 8 and 9.

DO:
- Print one sheet per student.
- Have 9 two-sided counters and a printed ten frame at each table.

TEACHER NOTES:
One core student resource today.

[General: Resources]`;

const NOTES_LAUNCH = `SAY:
- Yesterday we made 6 and 7. We used the PPW mat.
- Today we have a new tool: a ten frame.
- Look at this ten frame. How many counters?
- Whisper to your partner.

DO:
- Display a ten frame with 7 counters in standard order.
- Allow 10 seconds for whisper.
- Cold call: "How many?" Expect "7".
- Say: "Today we will make 8 and 9 on a ten frame."

TEACHER NOTES:
The launch introduces the ten frame as a new tool. Many students will recognise it from Daily Review. Keep it brisk - 60 seconds. Today the ten frame replaces the PPW mat for visual clarity with bigger numbers.

WATCH FOR:
- Quick "7" responses - subitising on the ten frame is forming.
- Slow counting - prompt: "Top row first - that's 5. How many more?"

[General: Launch | VTLM 2.0: Activate Prior Knowledge]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 4: Make 8 and 9",
    `Foundation Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch — introduce today's new tool, the ten frame
  contentSlide(pres, "Launch", C.SUCCESS, "A New Tool: Ten Frame",
    [
      "Yesterday we made 6 and 7.",
      "Today we use a new tool.",
      "How many counters on this ten frame?",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      // Ten frame with 7 counters — preview today's new tool
      addTensFrame(slide, lg.rightX + 0.10, lg.panelTopPadded + 0.55,
        lg.rightW - 0.20, 7, { fillColor: "D64545" });
      slide.addText("How many?", {
        x: lg.rightX, y: lg.panelTopPadded + 2.50, w: lg.rightW, h: 0.42,
        fontSize: 22, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slides 4-5: Daily Review with reveal — Teen number from ten frame
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "What Number?", { color: C.ACCENT });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

      // One full ten frame (10) plus three more dots in a small extra group
      addTensFrame(s, 1.6, CONTENT_TOP + 0.45, 4.0, 10, { fillColor: C.PRIMARY });

      // 3 extra dots to the right (showing "10 and 3 more")
      const extraX = 6.4;
      const extraY = CONTENT_TOP + 0.45;
      const dotSize = 0.65;
      for (let i = 0; i < 3; i += 1) {
        s.addShape("roundRect", {
          x: extraX + (i % 3) * (dotSize + 0.18),
          y: extraY + 0.65,
          w: dotSize, h: dotSize,
          rectRadius: dotSize / 2,
          fill: { color: C.PRIMARY },
          line: { color: C.CHARCOAL, width: 1 },
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
      addTextOnShape(slide, "13!  10 and 3 more.", {
        x: 1.5, y: 4.45, w: 7.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Before and after 15
  const sFluency = pres.addSlide();
  addTopBar(sFluency, C.ACCENT);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Before and After", { color: C.ACCENT });

  // Three big number panels: ___, 15, ___
  const panelY = CONTENT_TOP + 0.30;
  const panelW = 2.6;
  const panelH = 2.3;
  const gap = 0.40;
  const totalPW = panelW * 3 + gap * 2;
  const startX = (10 - totalPW) / 2;

  // BEFORE box (empty)
  sFluency.addShape("roundRect", {
    x: startX, y: panelY, w: panelW, h: panelH, rectRadius: 0.12,
    fill: { color: C.WHITE },
    line: { color: C.SECONDARY, width: 3, dashType: "dash" },
  });
  sFluency.addText("?", {
    x: startX, y: panelY, w: panelW, h: panelH,
    fontSize: 100, fontFace: FONT_H, color: C.SECONDARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  sFluency.addText("Before", {
    x: startX, y: panelY + panelH + 0.05, w: panelW, h: 0.35,
    fontSize: 18, fontFace: FONT_B, color: C.SECONDARY, bold: true,
    align: "center", valign: "top", margin: 0,
  });

  // 15 in the middle
  const midX = startX + panelW + gap;
  sFluency.addShape("roundRect", {
    x: midX, y: panelY, w: panelW, h: panelH, rectRadius: 0.12,
    fill: { color: C.PRIMARY },
  });
  sFluency.addText("15", {
    x: midX, y: panelY, w: panelW, h: panelH,
    fontSize: 100, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // AFTER box (empty)
  const afterX = midX + panelW + gap;
  sFluency.addShape("roundRect", {
    x: afterX, y: panelY, w: panelW, h: panelH, rectRadius: 0.12,
    fill: { color: C.WHITE },
    line: { color: C.SECONDARY, width: 3, dashType: "dash" },
  });
  sFluency.addText("?", {
    x: afterX, y: panelY, w: panelW, h: panelH,
    fontSize: 100, fontFace: FONT_H, color: C.SECONDARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  sFluency.addText("After", {
    x: afterX, y: panelY + panelH + 0.05, w: panelW, h: 0.35,
    fontSize: 18, fontFace: FONT_B, color: C.SECONDARY, bold: true,
    align: "center", valign: "top", margin: 0,
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to find the parts that make 8 and 9 using a ten frame.",
    [
      "I can find one way to make 8.",
      "I can find one way to make 9.",
      "I can show parts of 8 and 9 in a ten frame.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do 1 — Make 8 with ten frame
  workedExSlide(pres, 2, "I Do", "Make 8 in the Ten Frame",
    [
      "Whole is 8.",
      "",
      "5 red in the top row.",
      "",
      "3 yellow in the bottom row.",
      "",
      "5 and 3 more is 8.",
      "",
      "Whole 8.",
      "Parts 5 and 3.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      drawPpwMatWithTensFrame(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 8,
        partA: 5, partB: 3,
      });
    }
  );

  // Slide 7: I Do 2 — Make 9 with ten frame
  workedExSlide(pres, 2, "I Do", "Make 9 in the Ten Frame",
    [
      "Whole is 9.",
      "",
      "4 red.",
      "",
      "5 yellow.",
      "(Top row is full.)",
      "",
      "4 and 5 more is 9.",
      "",
      "Whole 9.",
      "Parts 4 and 5.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      drawPpwMatWithTensFrame(slide, lg.rightX, lg.panelTopPadded + 0.10, lg.rightW, 3.55, {
        wholeNumber: 9,
        partA: 4, partB: 5,
      });
    }
  );

  // Slides 8-9: CFU 1 with reveal — How many yellow in ten frame?
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

      // Ten frame: 5 red top, 3 yellow bottom
      drawTwoColourTensFrame(s, 2.5, CONTENT_TOP + 0.35, 5.0, 5, 3);

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
      addTextOnShape(slide, "3 yellow.   Whole 8.   Parts 5 and 3.", {
        x: 0.8, y: 4.45, w: 8.4, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 10-11: We Do with reveal — 9 = 6 and 3
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Find the Parts of 9", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      // Ten frame: 6 red, 3 yellow
      drawTwoColourTensFrame(s, 0.85, CONTENT_TOP + 0.50, 3.8, 6, 3);

      s.addText("Whole = 9", {
        x: 0.6, y: CONTENT_TOP + 2.30, w: 4.3, h: 0.50,
        fontSize: 26, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "9 = ____ and ____", options: { fontSize: 28, bold: true, color: C.CHARCOAL } },
      ], {
        x: 5.4, y: CONTENT_TOP + 0.35, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.55,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Whole 9.   Parts 6 and 3.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 12-13: CFU Hinge with reveal — 8 is 5 and 4 (false)
  withReveal(
    () => cfuSlide(pres, "CFU", "Is This Right?", "Thumbs Up or Thumbs Down",
      "I say:\n8 is 5 and 4.\n\nIs that right?",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN!  5 and 4 make 9, not 8.", {
        x: 1.0, y: 4.45, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Your Turn: 8 and 9",
    [
      "First: 8 counters.",
      "Fill the top row first.",
      "Find one way to make 8.",
      "",
      "Next: 9 counters.",
      "Find one way to make 9.",
      "",
      "Tell your partner.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.20, { strip: C.ALERT });
      slide.addText("You need:", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.40, h: 0.36,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "9 two-sided counters", options: { bullet: true, fontSize: 17, color: C.CHARCOAL, breakLine: true } },
        { text: "your ten frame", options: { bullet: true, fontSize: 17, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.40, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.50, h: 0.70,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Empty ten frame for reference
      drawTwoColourTensFrame(slide, lg.rightX + 0.20, lg.panelTopPadded + 1.65, lg.rightW - 0.40, 0, 0);
      slide.addText("Whole 8 or 9", {
        x: lg.rightX, y: lg.panelTopPadded + 3.10, w: lg.rightW, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Whole 8, parts 5 and 3. Draw the ten frame on your whiteboard.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 16: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Tell your partner one way to make 8 or 9.",
      scItems: [
        "I can find one way to make 8.",
        "I can find one way to make 9.",
        "I can show parts of 8 and 9 in a ten frame.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPWi_Lesson4_Make_8_and_9.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find ways to make 8 and 9 in a ten frame.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Foundation Numeracy`,
    });
    y = addTipBox(doc, "Colour the dots. Use 2 colours. Write the parts.", y, { color: C.ACCENT });

    function drawTenFramePdf(yTop) {
      const x0 = 90;
      const cell = 50;
      // 2 rows of 5 cells
      for (let r = 0; r < 2; r += 1) {
        for (let c = 0; c < 5; c += 1) {
          doc.lineWidth(1.5).rect(x0 + c * cell, yTop + r * cell, cell, cell).stroke("#333333");
          // Empty circle in each cell
          doc.circle(x0 + c * cell + cell / 2, yTop + r * cell + cell / 2, cell * 0.32)
            .lineWidth(1).stroke("#999999");
        }
      }
      return yTop + 2 * cell;
    }

    y = addSectionHeading(doc, "Whole = 8", y, { color: C.PRIMARY });
    y = drawTenFramePdf(y) + 14;
    doc.fontSize(14).font("Sans").fillColor("#333333").text("Whole 8.   Parts ____ and ____", 60, y);
    y += 36;

    y = addSectionHeading(doc, "Whole = 9", y, { color: C.PRIMARY });
    y = drawTenFramePdf(y) + 14;
    doc.fontSize(14).font("Sans").fillColor("#333333").text("Whole 9.   Parts ____ and ____", 60, y);
    y += 36;

    y = addTipBox(doc, "Tip: fill the top row of dots first.", y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Lesson 4 Ten Frame Parts | Foundation Numeracy`);
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
    y = addSectionHeading(doc, "Whole = 8", y, { color: C.PRIMARY });
    y = addBodyText(doc, "8 and 0   |   7 and 1   |   6 and 2   |   5 and 3   |   4 and 4   |   3 and 5   |   2 and 6   |   1 and 7   |   0 and 8", y);
    y = addSectionHeading(doc, "Whole = 9", y, { color: C.PRIMARY });
    y = addBodyText(doc, "9 and 0   |   8 and 1   |   7 and 2   |   6 and 3   |   5 and 4   |   4 and 5   |   3 and 6   |   2 and 7   |   1 and 8   |   0 and 9", y);
    y = addTipBox(doc, "Children fluent at filling the top row first will spot the structure: parts of 8 are 5 and some more (like 5 and 3).", y, { color: C.ACCENT });
    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Foundation Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
