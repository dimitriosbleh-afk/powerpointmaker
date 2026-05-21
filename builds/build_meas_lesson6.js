"use strict";

// Measurement Unit (Year 5/6 Numeracy) — Lesson 6: Elapsed time.
// Year 6 content. Number line / jump strategy. Hand-held clocks support.
// Daily Review: solving equations with x/÷.
// Fluency: subtraction.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(5));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 6;
const TOTAL = 7;
const UNIT_TITLE = "Measurement: Units, Time and Timetables";
const FOOTER = `Measurement | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/Meas_Lesson6_Elapsed_Time";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 6 Elapsed Time Practice",
  "Calculate elapsed time using a number line and jumps.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 6 Answer Key",
  "Worked answers for the Elapsed Time Practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back. Yesterday we read 12 and 24 hour time.
- Today we calculate elapsed time - how long something takes.

DO:
- Hand out hand-held analog clocks (one per pair).

TEACHER NOTES:
Lesson 6 of 7. Year 6 content. Number line + jumps is the strategy of choice. Hand-held clocks support hand position checks.

WATCH FOR:
- Students who try to subtract in vertical algorithm form - that often fails for time (eg. 4:15 - 1:45). Redirect to number lines.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- Each pair: one hand-held analog clock.
- One Elapsed Time Practice sheet per student.

DO:
- Pass out clocks.
- Print one sheet per student.
- Print one answer key.

TEACHER NOTES:
One student resource plus answer key. Clocks support hand position checks during the You Do.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Solve each equation. The unknown is x.
- Write what x equals on your whiteboard.

DO:
- Display the three prompts.
- 2 minutes.

TEACHER NOTES:
Continued from Lesson 4. Inverse operations.

WATCH FOR:
- Students who undo with the inverse - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 8 x x = 56. So x = 56 ÷ 8 = 7.
- x ÷ 4 = 9. So x = 9 x 4 = 36.
- 7 x x = 63. So x = 63 ÷ 7 = 9.

DO:
- Click to reveal.

TEACHER NOTES:
Same strategy as Lesson 4. Note students who still guess.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Subtraction facts.
- Quick whisper-answer then write.

DO:
- Display the three prompts.
- 30 seconds.

TEACHER NOTES:
Subtraction fluency continued.

WATCH FOR:
- Students who use known facts - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 18 - 9 = 9.
- 15 - 8 = 7.
- 17 - 9 = 8.

DO:
- Click to reveal.

TEACHER NOTES:
Look for the "think addition" strategy.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- One key idea today.
- Elapsed time is the time BETWEEN two clock times.
- It tells us how long something took.

DO:
- Point at the classroom clock and say "9 AM to 11 AM is 2 hours of elapsed time".

TEACHER NOTES:
Anchor the meaning of elapsed time. It is the gap between two times. We will use a number line and jumps.

WATCH FOR:
- Students who can repeat "elapsed = the gap" - tracking.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to calculate elapsed time using a number line.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 is whole-hour elapsed time. SC2 is mixed hours and minutes. SC3 stretches to elapsed time across noon or midnight.

WATCH FOR:
- Students who can repeat the SC - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. We will calculate the elapsed time from 9:35 AM to 11:20 AM.
- I draw a number line. Start time on the left. End time on the right.
- Now I make friendly jumps.
- 9:35 to 10:00 - that is 25 minutes (the jump to the next hour).
- 10:00 to 11:00 - that is 1 hour.
- 11:00 to 11:20 - that is 20 minutes.
- Add the jumps: 25 + 60 + 20 = 105 minutes.
- 105 minutes = 1 hour 45 minutes.

DO:
- Draw the number line on the board.
- Mark 9:35 - 10:00 - 11:00 - 11:20.
- Label each jump.
- Add them out loud.

TEACHER NOTES:
The number line + jumps strategy is the most reliable for elapsed time. Always jump to the next friendly time first (next hour), then full hours, then the final minutes.

MISCONCEPTIONS:
- Misconception: Students subtract in column form (11:20 - 9:35 = 2:-15).
  Why: They treat time as base 10.
  Impact: Wildly wrong answers because minutes are base 60.
  Quick correction: "Time is not base 10. Use jumps."

WATCH FOR:
- Students who jump to the next hour first - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Watch me with a second example: 7:50 PM to 9:15 PM.
- 7:50 to 8:00 - that is 10 minutes (next hour).
- 8:00 to 9:00 - that is 1 hour.
- 9:00 to 9:15 - that is 15 minutes.
- Add: 10 + 60 + 15 = 85 minutes = 1 hour 25 minutes.

DO:
- Draw the number line.
- Mark the jumps.
- Add them out loud.

TEACHER NOTES:
Second worked example. Same strategy. Smaller numbers to build confidence before students try alone.

WATCH FOR:
- Students who set up the number line - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- Calculate the elapsed time from 8:40 AM to 10:25 AM.
- Use a number line and jumps. Write the answer in hours and minutes.

DO:
- Display the prompt.
- 90 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your boards on three, two, one, show."
- Scan for: 1 hour 45 minutes (or 105 minutes).
PROCEED: If 80% have the right answer, click to reveal and move to We Do.
PIVOT: Most likely misconception - students subtract column-style and get 1:85 or 2:-15.
- Reteach: Use the number line. "Jump to the next hour first: 20 minutes."
- Re-check: "9:50 to 11:10 - elapsed time?"

TEACHER NOTES:
Probe the number line strategy. 8:40 to 9:00 = 20, 9:00 to 10:00 = 60, 10:00 to 10:25 = 25. Total 105 = 1h 45.

WATCH FOR:
- Students who set up jumps - secure.
- Students who try to subtract column-style - reteach.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me.
- Four elapsed time problems.
- Partner talk for each: where are the jumps?

DO:
- Display the four problems.
- 4 minutes total.
- Cold call after each.

TEACHER NOTES:
Mix of small and larger elapsed times. Stay on the number line strategy.

WATCH FOR:
- Pairs who draw jumps - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 10:15 to 11:00 = 45 minutes.
- 2:30 PM to 4:00 PM = 1 hour 30 minutes.
- 6:45 AM to 9:20 AM = 2 hours 35 minutes (15 + 120 + 20).
- 11:50 to 1:30 PM = 1 hour 40 minutes (10 + 60 + 30; crosses noon).

DO:
- Click to reveal.
- Run the jumps for each.

TEACHER NOTES:
Reveal confirms the jump structure. The last one crosses noon - it is a stretch question.

WATCH FOR:
- Pairs who got three of four - secure.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Two students disagree.
- Mei found the elapsed time from 9:45 to 11:15 was 1 hour 30 minutes.
- Jay found it was 2 hours 30 minutes.
- Thumbs UP for Mei. Thumbs DOWN for Jay.

DO:
- Display the disagreement.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP for Mei. DOWN for Jay."
- Scan for: thumbs UP. Mei is right.
PROCEED: If 80% agree with Mei, click to reveal and confirm.
PIVOT: Most likely misconception - Jay added one extra hour (off-by-one when crossing 11:00).
- Reteach: "Jump 9:45 to 10:00 = 15. 10:00 to 11:00 = 60. 11:00 to 11:15 = 15. Total 90 = 1h 30."
- Re-check: "8:50 to 10:20 - elapsed time?"

TEACHER NOTES:
The hinge tests off-by-one errors when counting hours. Use the number line every time.

WATCH FOR:
- Confident thumbs UP for Mei - they used jumps.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the Elapsed Time Practice sheet.
- For each one: draw the number line, make the jumps, add them up.
- Use your hand-held clock if it helps.

DO:
- Distribute the sheet.
- Circulate. Stop and ask "where is your first jump?"

TEACHER NOTES:
Different times from the We Do. Same number line strategy.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use Part A only. Sit with these students. Pre-draw the number line skeleton for the first one.
- Extra Notes: Use the hand-held clock to model each jump.
EXTENDING PROMPT:
- Task: Part D (problems crossing noon or midnight) - calculate elapsed time when the AM/PM changes.
- Extra Notes: Encourage students to convert to 24 hour time first if it helps.

WATCH FOR:
- Students who draw jumps and add - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Problem 1: Calculate the elapsed time from 7:25 AM to 9:10 AM.
- Problem 2: A movie starts at 6:50 PM and ends at 8:35 PM. How long is the movie?

DO:
- Display the prompt.
- 4 minutes.

TEACHER NOTES:
Exit ticket assesses SC2. Look for 1 h 45 min (both).

WATCH FOR:
- Students who get both right with number lines - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria.
- Show me thumbs up, sideways, or down for each.
- Turn and tell your partner: what is the first jump on a number line for elapsed time?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea: first jump to the next hour. Then full hours. Then the final minutes. Tomorrow we read timetables and plan a trip.

WATCH FOR:
- Strong thumbs up - move at pace tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

// Reuse the analog clock helper from Lesson 5.
function drawAnalogClock(slide, cx, cy, r, hour, minute, opts) {
  const o = opts || {};
  const ringColor = o.ringColor || C.CHARCOAL;
  const faceColor = o.faceColor || C.WHITE;
  const hourHandColor = o.hourHandColor || C.PRIMARY;
  const minuteHandColor = o.minuteHandColor || C.ALERT;

  slide.addShape("roundRect", {
    x: cx - r, y: cy - r, w: r * 2, h: r * 2, rectRadius: r,
    fill: { color: ringColor },
  });
  slide.addShape("roundRect", {
    x: cx - r * 0.92, y: cy - r * 0.92, w: r * 1.84, h: r * 1.84,
    rectRadius: r * 0.92,
    fill: { color: faceColor },
  });
  const numbers = [
    { text: "12", x: cx, y: cy - r * 0.78 },
    { text: "3",  x: cx + r * 0.72, y: cy },
    { text: "6",  x: cx, y: cy + r * 0.72 },
    { text: "9",  x: cx - r * 0.72, y: cy },
  ];
  numbers.forEach((n) => {
    slide.addText(n.text, {
      x: n.x - 0.20, y: n.y - 0.12, w: 0.40, h: 0.24,
      fontSize: 12, fontFace: FONT_H, color: C.CHARCOAL,
      bold: true, align: "center", valign: "middle", margin: 0,
    });
  });
  for (let h = 1; h <= 12; h++) {
    const angle = (h / 12) * 2 * Math.PI - Math.PI / 2;
    const tx1 = cx + Math.cos(angle) * r * 0.85;
    const ty1 = cy + Math.sin(angle) * r * 0.85;
    const tx2 = cx + Math.cos(angle) * r * 0.92;
    const ty2 = cy + Math.sin(angle) * r * 0.92;
    slide.addShape("line", {
      x: Math.min(tx1, tx2), y: Math.min(ty1, ty2),
      w: Math.abs(tx2 - tx1), h: Math.abs(ty2 - ty1),
      flipH: tx2 < tx1, flipV: ty2 < ty1,
      line: { color: C.CHARCOAL, width: 1 },
    });
  }
  const hourAngle = ((hour % 12) + minute / 60) / 12 * 2 * Math.PI - Math.PI / 2;
  const hourLen = r * 0.50;
  const hx = cx + Math.cos(hourAngle) * hourLen;
  const hy = cy + Math.sin(hourAngle) * hourLen;
  slide.addShape("line", {
    x: Math.min(cx, hx), y: Math.min(cy, hy),
    w: Math.abs(hx - cx), h: Math.abs(hy - cy),
    flipH: hx < cx, flipV: hy < cy,
    line: { color: hourHandColor, width: 4 },
  });
  const minAngle = (minute / 60) * 2 * Math.PI - Math.PI / 2;
  const minLen = r * 0.78;
  const mx = cx + Math.cos(minAngle) * minLen;
  const my = cy + Math.sin(minAngle) * minLen;
  slide.addShape("line", {
    x: Math.min(cx, mx), y: Math.min(cy, my),
    w: Math.abs(mx - cx), h: Math.abs(my - cy),
    flipH: mx < cx, flipV: my < cy,
    line: { color: minuteHandColor, width: 3 },
  });
  slide.addShape("roundRect", {
    x: cx - 0.06, y: cy - 0.06, w: 0.12, h: 0.12, rectRadius: 0.06,
    fill: { color: C.CHARCOAL },
  });
}

// Draw a number line with custom marks and arc-style jumps labelled above each.
// marks: [{ label, position }] where position is 0..1
// jumps: [{ from, to, label }] each refers to a mark index
function drawElapsedNumberLine(slide, x, y, w, marks, jumps) {
  const lineY = y + 0.40;
  // Main line
  slide.addShape("line", {
    x, y: lineY, w, h: 0,
    line: { color: C.CHARCOAL, width: 2 },
  });
  // Arrows
  slide.addShape("line", {
    x: x - 0.15, y: lineY - 0.10, w: 0.15, h: 0.10,
    line: { color: C.CHARCOAL, width: 1.5 },
  });
  slide.addShape("line", {
    x: x + w, y: lineY - 0.10, w: 0.15, h: 0.10,
    line: { color: C.CHARCOAL, width: 1.5 },
  });
  // Marks + labels. Label widths are clamped so adjacent labels don't overlap.
  marks.forEach((m, i) => {
    const mx = x + m.position * w;
    slide.addShape("line", {
      x: mx, y: lineY - 0.07, w: 0, h: 0.14,
      line: { color: C.CHARCOAL, width: 1.5 },
    });
    // Compute label width: at most the half-distance to each neighbour mark, minus a small gap.
    const prevPos = i > 0 ? marks[i - 1].position : m.position - 0.20;
    const nextPos = i < marks.length - 1 ? marks[i + 1].position : m.position + 0.20;
    const leftHalf = (m.position - prevPos) * w / 2;
    const rightHalf = (nextPos - m.position) * w / 2;
    const labelW = Math.max(0.40, Math.min(1.00, 2 * Math.min(leftHalf, rightHalf) - 0.04));
    // Shrink font for very narrow labels so they still fit.
    const fontSize = labelW < 0.55 ? 9 : 11;
    slide.addText(m.label, {
      x: mx - labelW / 2, y: lineY + 0.10, w: labelW, h: 0.26,
      fontSize, fontFace: FONT_B, color: C.CHARCOAL,
      bold: true, align: "center", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  });
  // Jumps (drawn as arcs above the line using rounded rectangles + labels)
  jumps.forEach((j, i) => {
    const startMark = marks[j.from];
    const endMark = marks[j.to];
    const sx = x + startMark.position * w;
    const ex = x + endMark.position * w;
    const arcW = ex - sx;
    const arcY = lineY - 0.55;
    // Bar (rounded rect to suggest the jump)
    slide.addShape("roundRect", {
      x: sx + 0.04, y: arcY, w: arcW - 0.08, h: 0.30, rectRadius: 0.06,
      fill: { color: i % 2 === 0 ? C.PRIMARY : C.ACCENT },
    });
    slide.addText(j.label, {
      x: sx, y: arcY, w: arcW, h: 0.30,
      fontSize: 11, fontFace: FONT_B, color: C.WHITE,
      bold: true, align: "center", valign: "middle", margin: 0,
    });
  });
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 6: Elapsed time",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal — solving equations
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Solving equations",
      ["8 × x = 56", "x ÷ 4 = 9", "7 × x = 63"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "x = 7         x = 36         x = 9", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal — subtraction
  withReveal(
    () => fluencySlide(pres, "Fluency: Subtraction",
      ["18 - 9", "15 - 8", "17 - 9"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "9     7     8", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Key vocabulary — elapsed time
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Elapsed time = the gap between two times");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    // Two clocks with an arrow between them
    drawAnalogClock(s, 2.5, CONTENT_TOP + 1.30, 0.95, 9, 0);
    s.addText("9:00 AM", {
      x: 1.5, y: CONTENT_TOP + 2.55, w: 2.0, h: 0.30,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY,
      bold: true, align: "center", margin: 0,
    });

    addTextOnShape(s, "elapsed time = 2 hours", {
      x: 3.7, y: CONTENT_TOP + 1.10, w: 2.6, h: 0.55, rectRadius: 0.10,
      fill: { color: C.ACCENT },
    }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    s.addText("->", {
      x: 3.7, y: CONTENT_TOP + 1.75, w: 2.6, h: 0.50,
      fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL,
      bold: true, align: "center", margin: 0,
    });

    drawAnalogClock(s, 7.5, CONTENT_TOP + 1.30, 0.95, 11, 0);
    s.addText("11:00 AM", {
      x: 6.5, y: CONTENT_TOP + 2.55, w: 2.0, h: 0.30,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY,
      bold: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to calculate elapsed time using a number line.",
    [
      "I can find the elapsed time when the times are whole hours (e.g. 9 AM to 11 AM).",
      "I can find the elapsed time using number-line jumps (next hour, full hours, then minutes).",
      "I can find the elapsed time across noon or midnight (e.g. 11:50 AM to 1:30 PM).",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: I Do (1) — 9:35 to 11:20 worked example
  workedExSlide(pres, 2, "I Do", "Elapsed time: 9:35 AM to 11:20 AM",
    [
      "Strategy: number line + jumps.",
      "",
      "Jump 1: 9:35 -> 10:00 = 25 min.",
      "Jump 2: 10:00 -> 11:00 = 60 min.",
      "Jump 3: 11:00 -> 11:20 = 20 min.",
      "",
      "Total: 25 + 60 + 20 = 105 min.",
      "",
      "= 1 hour 45 minutes.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("Number line + jumps", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });

      drawElapsedNumberLine(slide,
        lg.rightX + 0.30, lg.panelTopPadded + 1.10, lg.rightW - 0.60,
        [
          { label: "9:35",  position: 0 },
          { label: "10:00", position: 0.28 },
          { label: "11:00", position: 0.78 },
          { label: "11:20", position: 1 },
        ],
        [
          { from: 0, to: 1, label: "25 min" },
          { from: 1, to: 2, label: "60 min" },
          { from: 2, to: 3, label: "20 min" },
        ]
      );

      addTextOnShape(slide, "Total: 1 hr 45 min", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.40, w: lg.rightW - 0.40, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10: I Do (2) — 7:50 PM to 9:15 PM
  workedExSlide(pres, 2, "I Do", "Elapsed time: 7:50 PM to 9:15 PM",
    [
      "Same strategy.",
      "",
      "Jump 1: 7:50 -> 8:00 = 10 min.",
      "Jump 2: 8:00 -> 9:00 = 60 min.",
      "Jump 3: 9:00 -> 9:15 = 15 min.",
      "",
      "Total: 10 + 60 + 15 = 85 min.",
      "",
      "= 1 hour 25 minutes.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("Number line + jumps", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      drawElapsedNumberLine(slide,
        lg.rightX + 0.30, lg.panelTopPadded + 1.10, lg.rightW - 0.60,
        [
          { label: "7:50", position: 0 },
          { label: "8:00", position: 0.12 },
          { label: "9:00", position: 0.82 },
          { label: "9:15", position: 1 },
        ],
        [
          { from: 0, to: 1, label: "10 min" },
          { from: 1, to: 2, label: "60 min" },
          { from: 2, to: 3, label: "15 min" },
        ]
      );

      addTextOnShape(slide, "Total: 1 hr 25 min", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.40, w: lg.rightW - 0.40, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 11-12: CFU + reveal — 8:40 AM to 10:25 AM
  withReveal(
    () => cfuSlide(pres, "CFU", "Elapsed time: 8:40 AM to 10:25 AM", "Show Me Boards",
      "On your whiteboard:\n\nDraw a number line and make the jumps.\n\nWrite the answer in hours and minutes.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1 hour 45 minutes  (20 + 60 + 25 = 105 min)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 13-14: We Do + reveal — four elapsed-time problems
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Find each elapsed time with your partner", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      const items = [
        { from: "10:15", to: "11:00",     note: "AM" },
        { from: "2:30 PM", to: "4:00 PM", note: "" },
        { from: "6:45 AM", to: "9:20 AM", note: "" },
        { from: "11:50",  to: "1:30 PM",  note: "crosses noon" },
      ];

      const cellW = 2.0;
      const gap = 0.15;
      const totalW = cellW * 4 + gap * 3;
      const startX = (10 - totalW) / 2;
      const cardY = CONTENT_TOP + 0.30;
      const cardH = 1.80;

      items.forEach((it, i) => {
        const cx = startX + i * (cellW + gap);
        addCard(s, cx, cardY, cellW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
        s.addText("From", {
          x: cx, y: cardY + 0.10, w: cellW, h: 0.24,
          fontSize: 11, fontFace: FONT_B, color: C.MUTED,
          align: "center", margin: 0,
        });
        s.addText(it.from, {
          x: cx, y: cardY + 0.35, w: cellW, h: 0.40,
          fontSize: 18, fontFace: FONT_H, color: C.PRIMARY,
          bold: true, align: "center", margin: 0,
        });
        s.addText("to", {
          x: cx, y: cardY + 0.85, w: cellW, h: 0.24,
          fontSize: 11, fontFace: FONT_B, color: C.MUTED,
          italic: true, align: "center", margin: 0,
        });
        s.addText(it.to, {
          x: cx, y: cardY + 1.10, w: cellW, h: 0.40,
          fontSize: 18, fontFace: FONT_H, color: C.SECONDARY,
          bold: true, align: "center", margin: 0,
        });
        if (it.note) {
          s.addText(it.note, {
            x: cx, y: cardY + 1.55, w: cellW, h: 0.22,
            fontSize: 10, fontFace: FONT_B, color: C.ALERT,
            italic: true, align: "center", margin: 0,
          });
        }
      });

      s.addText("Partner talk: where is the first jump? Add: jump 1 + full hours + jump 3.", {
        x: 0.7, y: SAFE_BOTTOM - 0.55, w: 8.6, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "45 min   |   1 h 30 min   |   2 h 35 min   |   1 h 40 min", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 15-16: CFU hinge — Mei vs Jay
  withReveal(
    () => cfuSlide(pres, "CFU", "Who is right?", "Thumbs Up or Thumbs Down",
      "Elapsed time from 9:45 to 11:15.\n\nMei: 1 hour 30 minutes.\n\nJay: 2 hours 30 minutes.\n\nThumbs UP for Mei.   Thumbs DOWN for Jay.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Mei is right.  9:45 -> 10:00 = 15.  10:00 -> 11:00 = 60.  11:00 -> 11:15 = 15.  Total 90 = 1 h 30.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: Elapsed Time Practice", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "draw your number line.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "make the jumps.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "add them up. Convert minutes if needed.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Remember the 3-jump pattern", {
      x: 0.7, y: panelY + 0.10, w: 8.6, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    // Show the three jump labels in a chip row
    const chipY = panelY + 0.55;
    const chips = [
      { label: "Jump 1: to the next hour", color: C.PRIMARY },
      { label: "Jump 2: full hours",        color: C.ACCENT },
      { label: "Jump 3: extra minutes",     color: C.ALERT },
    ];
    const chipW = 2.7;
    const chipGap = 0.15;
    const chipTotalW = chipW * 3 + chipGap * 2;
    const chipStartX = (10 - chipTotalW) / 2;
    chips.forEach((c, i) => {
      addTextOnShape(s, c.label, {
        x: chipStartX + i * (chipW + chipGap), y: chipY, w: chipW, h: 0.55, rectRadius: 0.10,
        fill: { color: c.color },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    });

    addTextOnShape(s, "Then ADD the three jumps.", {
      x: 2.0, y: chipY + 0.75, w: 6.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 18: Exit Ticket
  exitTicketSlide(pres,
    [
      "Calculate the elapsed time from 7:25 AM to 9:10 AM. Show your jumps.",
      "A movie starts at 6:50 PM and ends at 8:35 PM. How long is the movie?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 19: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the first jump on a number line for elapsed time?",
      scItems: [
        "I can find the elapsed time when the times are whole hours (e.g. 9 AM to 11 AM).",
        "I can find the elapsed time using number-line jumps (next hour, full hours, then minutes).",
        "I can find the elapsed time across noon or midnight (e.g. 11:50 AM to 1:30 PM).",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Meas_Lesson6_Elapsed_Time.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Draw a number line. Jump to the next hour, full hours, then the final minutes.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "3 jumps. Jump 1: from start time to the next hour (minutes). Jump 2: full hours. Jump 3: extra minutes at the end. Add the three. Convert 60 minutes into 1 hour at the end.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Part A — Whole hours (warm-up)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  9:00 AM to 12:00 PM = ___", y);
    y = addWriteLine(doc, "b)  1:00 PM to 5:00 PM = ___", y);
    y = addWriteLine(doc, "c)  8:00 AM to 3:00 PM = ___", y);

    y = addSectionHeading(doc, "Part B — Mixed hours and minutes", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  9:25 AM to 11:00 AM = ___ hr ___ min", y);
    y = addWriteLine(doc, "b)  3:15 PM to 5:50 PM = ___ hr ___ min", y);
    y = addWriteLine(doc, "c)  6:40 AM to 9:05 AM = ___ hr ___ min", y);
    y = addWriteLine(doc, "d)  10:10 AM to 1:35 PM = ___ hr ___ min", y);

    y = addSectionHeading(doc, "Part C — Word problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Use a number line for each.", y);
    y = addWriteLine(doc, "a)  A swim lesson runs from 4:25 PM to 5:10 PM. How long is the lesson? ___", y);
    y = addWriteLine(doc, "b)  Recess starts at 10:50 AM and ends at 11:20 AM. How long is recess? ___", y);
    y = addWriteLine(doc, "c)  A bus leaves the station at 7:45 AM and arrives at 9:30 AM. How long is the trip? ___", y);

    y = addSectionHeading(doc, "Part D — Extension: crosses noon or midnight", y, { color: C.ACCENT });
    y = addWriteLine(doc, "a)  11:35 AM to 1:20 PM = ___ hr ___ min", y);
    y = addWriteLine(doc, "b)  10:50 PM to 1:15 AM = ___ hr ___ min", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Elapsed Time Practice | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Elapsed Time Practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Part A — Whole hours", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  9 AM to 12 PM = 3 hours.", y);
    y = addBodyText(doc, "b)  1 PM to 5 PM = 4 hours.", y);
    y = addBodyText(doc, "c)  8 AM to 3 PM = 7 hours.", y);

    y = addSectionHeading(doc, "Part B — Mixed", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  9:25 AM to 11:00 AM = 1 hr 35 min  (35 + 60 = 95 min).", y);
    y = addBodyText(doc, "b)  3:15 PM to 5:50 PM = 2 hr 35 min  (45 + 60 + 50 = 155 min).", y);
    y = addBodyText(doc, "c)  6:40 AM to 9:05 AM = 2 hr 25 min  (20 + 120 + 5 = 145 min).", y);
    y = addBodyText(doc, "d)  10:10 AM to 1:35 PM = 3 hr 25 min  (50 + 120 + 35 = 205 min).", y);

    y = addSectionHeading(doc, "Part C — Word problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4:25 PM to 5:10 PM = 45 min.", y);
    y = addBodyText(doc, "b)  10:50 AM to 11:20 AM = 30 min.", y);
    y = addBodyText(doc, "c)  7:45 AM to 9:30 AM = 1 hr 45 min.", y);

    y = addSectionHeading(doc, "Part D — Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  11:35 AM to 1:20 PM = 1 hr 45 min  (25 + 60 + 20 = 105 min).", y);
    y = addBodyText(doc, "b)  10:50 PM to 1:15 AM = 2 hr 25 min  (10 + 60 + 60 + 15 = 145 min).", y);

    y = addTipBox(doc,
      "Watch for: students who subtract column-style (4:15 - 1:45 will not work); off-by-one errors when crossing the hour; students who forget to convert 90 min to 1 h 30.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 6 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
