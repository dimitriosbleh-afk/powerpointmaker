"use strict";

// Geometry and Data Unit (Year 5/6 Numeracy) - Lesson 7: Interpreting line graphs.
// Year 5 VC2M5ST02 - interpret line graphs representing change over time.
// Year 6 VC2M6ST01 connections - comparative displays.
// Daily Review: Solving equations (retrieval mix).
// Fluency: Mixed multiplication and division (less scaffold).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 2);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 7;
const TOTAL = 8;
const UNIT_TITLE = "Geometry and Data: Shapes, Tessellations, Data";
const FOOTER = `Geometry and Data | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/GD_Lesson7_Line_Graphs";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 7 Line Graph Reading",
  "Read and interpret three line graphs. Describe change over time.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 7 Answer Key",
  "Worked answers for the Line Graph Reading sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// Temperature data for I Do (one day, morning to afternoon)
const TEMP_DATA = [
  { x: "8 am",  y: 14 },
  { x: "10 am", y: 18 },
  { x: "12 pm", y: 22 },
  { x: "2 pm",  y: 26 },
  { x: "4 pm",  y: 23 },
  { x: "6 pm",  y: 19 },
];

// Plant height data for CFU
const PLANT_DATA = [
  { x: "Wk 1", y: 2 },
  { x: "Wk 2", y: 4 },
  { x: "Wk 3", y: 7 },
  { x: "Wk 4", y: 11 },
  { x: "Wk 5", y: 14 },
  { x: "Wk 6", y: 16 },
];

// Library borrowing (rising and falling) for We Do
const LIBRARY_DATA = [
  { x: "Mon", y: 25 },
  { x: "Tue", y: 30 },
  { x: "Wed", y: 18 },
  { x: "Thu", y: 22 },
  { x: "Fri", y: 35 },
];

// =====================================================================
// Teacher notes
// =====================================================================

const NOTES_TITLE = `SAY:
- Yesterday we used a column graph to show categories.
- Today we look at LINE graphs - graphs that show CHANGE OVER TIME.
- The line goes up when something increases. Down when it decreases. Flat when it stays the same.

DO:
- Settle.
- Whiteboards ready.

TEACHER NOTES:
Lesson 7 of 8. Today's focus: read line graphs and describe what changes. Tomorrow is the final critical-thinking lesson on media data.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- Whiteboards on desks.
- One Line Graph Reading sheet per student.

DO:
- Print one student sheet per student.
- Print one answer key.

TEACHER NOTES:
The student sheet has three line graphs to read. They are simple - the focus is interpretation, not drawing from scratch.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Solving equations.
- Work each on your whiteboard.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Mixed retrieval of solving equations. Touches all three earlier days' content.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 4 x n = 28, so n = 7.
- n divided by 5 = 8, so n = 40.
- 6 x n - 5 = 31, so 6n = 36, n = 6.

DO:
- Click to reveal.

TEACHER NOTES:
Quick reveal.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Mixed multiplication and division.
- Whisper-answer then write.

DO:
- Display the three prompts.
- 25 seconds.

TEACHER NOTES:
Day seven. Mixed mult/div. Most release of the week - students should not need scaffolding language.

WATCH FOR:
- Students who recall instantly - secure for end-of-unit fluency check.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 7 x 12 = 84.
- 96 divided by 8 = 12.
- 11 x 11 = 121.

DO:
- Click to reveal.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Quick prompt. Stand up.
- I will say a time. You "draw" the line with your arm.
- Morning - arm low. Lunch time - arm middle. 2 pm - arm high. Dinner time - arm middle again. Bed - arm low.
- That is the shape of a line graph: heights at different times.

DO:
- Call out the times. Class mimics.

TEACHER NOTES:
A short kinaesthetic launch that turns "change over time" into a body movement. The line graph is just this curve, drawn.

WATCH FOR:
- Students who can describe the arm shape as "up then down" - secure.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to interpret line graphs that show change over time.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 - read one value off a line graph. SC2 - core target - describe when the value rose, fell or stayed the same. SC3 - draw conclusions about the situation behind the graph.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Three key words today.
- Axis - the labelled side of the graph (X for time, Y for value).
- Trend - the overall direction of the line.
- Conclusion - what we can say about the situation from the graph.

DO:
- Point at each visual.

TEACHER NOTES:
Trend is the key idea. "Rising trend" - going up over time. "Falling trend" - going down. "No clear trend" - jumps around.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_IDO_READ = `SAY:
- Watch me. This is a temperature graph for one day.
- Time goes along the bottom (X axis). Temperature goes up the side (Y axis).
- At 12 pm I follow up to the line, then across to read the temperature. 22 degrees.
- At 4 pm I do the same. 23 degrees. The hottest point was actually 2 pm, at 26 degrees.

DO:
- Use your finger on the on-screen graph.
- Trace up from 12 pm, then across.

TEACHER NOTES:
Anchor the read-off move: up from the time, across to the temperature. Many students try to read along the line directly.

WATCH FOR:
- Students who trace up first, then across - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_TREND = `SAY:
- Now I describe the trend.
- From 8 am to 2 pm, the line goes UP. Temperature rose.
- From 2 pm onwards, the line goes DOWN. Temperature fell.
- Overall trend - rise in the morning, peak in the afternoon, then fall.

DO:
- Sweep your finger along the rising part.
- Sweep along the falling part.

TEACHER NOTES:
Pair the direction of the line with the word: rising, falling, peak. These are the words students should use in their conclusions.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Plant height graph.
- Between week 3 and week 4, how much did the plant grow?
- Write the difference in cm on your board.

DO:
- Display the prompt.
- 45 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me on three, two, one, show."
- Scan for: 4 cm.
PROCEED: If 80% have 4, click to reveal and move to We Do.
PIVOT: Most likely misconception - students read 11 (the height at week 4) or 7 (the height at week 3), not the difference.
- Reteach: "Growth is the CHANGE - week 4 height MINUS week 3 height."
- Re-check: "Between week 5 and week 6, how much did it grow?"

TEACHER NOTES:
Plant data: Wk 1=2, Wk 2=4, Wk 3=7, Wk 4=11, Wk 5=14, Wk 6=16. Growth Wk 3 to Wk 4 = 11 - 7 = 4 cm.

WATCH FOR:
- Students who write 4 - secure.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO = `SAY:
- With me. Library borrowing graph.
- On which day did borrowing DROP compared to the day before? Why might that be?
- On which day was borrowing HIGHEST? Why might that be?

DO:
- Look at the graph together.
- Ask the class for each answer.

TEACHER NOTES:
Library data: Mon=25, Tue=30, Wed=18, Thu=22, Fri=35.
Drop = Wednesday (down from 30 to 18). Possible reasons - assembly, sports day, library closed.
Highest = Friday (35). Possible reasons - kids borrow for the weekend.

WATCH FOR:
- Students who offer a real-world reason - secure SC3.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Drop = Wednesday (30 down to 18).
- Highest = Friday (35).
- Possible reasons - assembly on Wednesday, students borrow for the weekend on Friday.

DO:
- Click to reveal.

TEACHER NOTES:
Highlight: the GRAPH does not tell us WHY. The conclusion is a sensible guess based on the data.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn. Take the Line Graph Reading sheet.
- Three short graphs - school rainfall, screen time and class attendance.
- For each: read off a value, describe the trend, write a conclusion.

DO:
- Distribute the sheet.
- Circulate.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Do Graph 1 only. Focus on reading values and describing direction.
- Extra Notes: Sit with these students. Demonstrate the up-then-across read.
EXTENDING PROMPT:
- Task: For Graph 3 (class attendance), suggest TWO different explanations and pick which is most likely.
- Extra Notes: Push for evidence-based reasoning, not guessing.

TEACHER NOTES:
Different content from We Do. Same skills: read, describe, conclude.

WATCH FOR:
- Students who use "rose" "fell" "stayed the same" - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Look at the line graph on the slide.
- Q1: What was the temperature at 2 pm?
- Q2: When did the temperature start to fall?

DO:
- Display the prompt.
- 90 seconds.

TEACHER NOTES:
Exit ticket assesses SC2 - reading a value and describing change over time.

WATCH FOR:
- Students who answer 26 degrees and "after 2 pm" or "from 2 pm" - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria for today.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: what does a line going UP mean?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea - a line going up shows an increase over time. Tomorrow we look at how graphs in the news can be misleading.

[General: Closing | VTLM 2.0: Reflection]`;

// =====================================================================
// Helpers
// =====================================================================

// Draw a line graph in the plot area (x,y,w,h).
function drawLineGraph(slide, x, y, w, h, data, opts) {
  const o = opts || {};
  const padL = 0.55;
  const padB = 0.50;
  const plotX = x + padL;
  const plotY = y;
  const plotW = w - padL;
  const plotH = h - padB;

  const yVals = data.map(d => d.y);
  const yMin = 0;
  const yMaxRaw = Math.max(...yVals);
  const yMax = Math.ceil((yMaxRaw + 2) / 5) * 5;

  // Axes
  slide.addShape("line", {
    x: plotX, y: plotY, w: 0, h: plotH,
    line: { color: C.CHARCOAL, width: 1.5 },
  });
  slide.addShape("line", {
    x: plotX, y: plotY + plotH, w: plotW, h: 0,
    line: { color: C.CHARCOAL, width: 1.5 },
  });

  // Y ticks
  const step = yMax <= 10 ? 2 : (yMax <= 20 ? 4 : (yMax <= 40 ? 5 : 10));
  for (let v = 0; v <= yMax; v += step) {
    const ty = plotY + plotH - (v / yMax) * plotH;
    slide.addShape("line", {
      x: plotX - 0.06, y: ty, w: 0.06, h: 0,
      line: { color: C.CHARCOAL, width: 1.2 },
    });
    slide.addText(String(v), {
      x: plotX - 0.50, y: ty - 0.12, w: 0.42, h: 0.24,
      fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL,
      align: "right", valign: "middle", margin: 0,
    });
  }

  // X categories
  const stepX = plotW / (data.length - 1);
  data.forEach((d, i) => {
    const px = plotX + i * stepX;
    slide.addShape("line", {
      x: px, y: plotY + plotH, w: 0, h: 0.06,
      line: { color: C.CHARCOAL, width: 1.2 },
    });
    slide.addText(d.x, {
      x: px - 0.45, y: plotY + plotH + 0.05, w: 0.90, h: 0.30,
      fontSize: 10, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "top", margin: 0,
    });
  });

  // Line segments
  const lineColor = o.lineColor || C.PRIMARY;
  for (let i = 0; i < data.length - 1; i++) {
    const x1 = plotX + i * stepX;
    const y1 = plotY + plotH - (data[i].y / yMax) * plotH;
    const x2 = plotX + (i + 1) * stepX;
    const y2 = plotY + plotH - (data[i + 1].y / yMax) * plotH;
    const lineX = Math.min(x1, x2);
    const lineY = Math.min(y1, y2);
    slide.addShape("line", {
      x: lineX, y: lineY,
      w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
      line: { color: lineColor, width: 2.5 },
      flipV: (x1 < x2) === (y1 < y2) ? false : true,
    });
  }
  // Points
  data.forEach((d, i) => {
    const px = plotX + i * stepX;
    const py = plotY + plotH - (d.y / yMax) * plotH;
    slide.addShape("roundRect", {
      x: px - 0.08, y: py - 0.08, w: 0.16, h: 0.16, rectRadius: 0.08,
      fill: { color: lineColor }, line: { color: lineColor, width: 1 },
    });
  });

  // Axis titles
  if (o.xLabel) {
    slide.addText(o.xLabel, {
      x: plotX, y: plotY + plotH + 0.36, w: plotW, h: 0.25,
      fontSize: 11, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", margin: 0,
    });
  }
  if (o.yLabel) {
    slide.addText(o.yLabel, {
      x: x - 0.10, y: plotY - 0.35, w: 2.40, h: 0.25,
      fontSize: 10, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "left", margin: 0,
    });
  }
}

// =====================================================================
// Build
// =====================================================================

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 7: Interpreting line graphs",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Solving equations",
      ["4 x n = 28", "n / 5 = 8", "6 x n - 5 = 31"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "n = 7         n = 40         n = 6", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal
  withReveal(
    () => fluencySlide(pres, "Fluency: Mixed multiplication and division",
      ["7 x 12", "96 div 8", "11 x 11"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "84       12       121", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Launch", { color: C.PRIMARY });
    addTitle(s, "Draw the day with your arm");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    s.addText([
      { text: "Stand up. As the teacher says a time, raise your arm to the right height.", options: { fontSize: 16, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 12, breakLine: true } },
      { text: "Morning -> arm low.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "Lunchtime -> arm middle.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "2 pm -> arm high.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "Dinner time -> arm middle.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "Bedtime -> arm low.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 12, breakLine: true } },
      { text: "That shape is a line graph.", options: { fontSize: 16, color: C.ALERT, bold: true } },
    ], {
      x: 1.0, y: CONTENT_TOP + 0.30, w: 8.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.60,
      fontFace: FONT_B, valign: "top", margin: 0, align: "center",
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to interpret line graphs that show change over time.",
    [
      "I can read one value off a line graph.",
      "I can describe when the value rose, fell or stayed the same.",
      "I can draw a sensible conclusion about the situation behind the graph.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key vocabulary
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Axis, trend, conclusion");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const items = [
      { word: "Axis",       simple: "the side",  detail: "X axis = time (bottom). Y axis = value (side).", color: C.PRIMARY },
      { word: "Trend",      simple: "direction", detail: "Overall direction of the line: up, down, or flat.", color: C.ACCENT },
      { word: "Conclusion", simple: "what we learn", detail: "A sensible statement about what the graph shows.", color: C.ALERT },
    ];
    const cellW = 2.7, gap = 0.20;
    const totalW = cellW * 3 + gap * 2;
    const startX = (10 - totalW) / 2;
    const cardY = CONTENT_TOP + 0.30;
    const cardH = 2.50;

    items.forEach((it, i) => {
      const cx = startX + i * (cellW + gap);
      addCard(s, cx, cardY, cellW, cardH, { strip: it.color, fill: C.WHITE });
      addTextOnShape(s, it.word, {
        x: cx + 0.15, y: cardY + 0.20, w: cellW - 0.30, h: 0.60, rectRadius: 0.10,
        fill: { color: it.color },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(it.simple, {
        x: cx, y: cardY + 0.95, w: cellW, h: 0.40,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
      s.addText(it.detail, {
        x: cx + 0.15, y: cardY + 1.50, w: cellW - 0.30, h: 0.85,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 10: I Do - read the temperature graph
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Temperature for one day", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    drawLineGraph(s,
      0.85, CONTENT_TOP + 0.45,
      5.30, 2.55,
      TEMP_DATA, {
        lineColor: C.ALERT,
        xLabel: "Time of day",
        yLabel: "Temperature (°C)",
      }
    );

    // Right side: read-off cue
    addCard(s, 6.40, CONTENT_TOP + 0.30, 3.10, 2.80, { strip: stageColor, fill: C.WHITE });
    s.addText([
      { text: "Read 12 pm:", options: { fontSize: 14, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "1. Up from 12 pm.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
      { text: "2. Across to the Y axis.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
      { text: "Answer = 22 degrees.", options: { fontSize: 15, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Peak = 2 pm, 26 degrees.", options: { fontSize: 13, color: C.CHARCOAL, italic: true } },
    ], {
      x: 6.55, y: CONTENT_TOP + 0.45, w: 2.85, h: 2.50,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_READ);
  })();

  // Slide 11: I Do - describe the trend
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Describe the trend", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    drawLineGraph(s,
      0.85, CONTENT_TOP + 0.45,
      5.30, 2.55,
      TEMP_DATA, {
        lineColor: C.ALERT,
        xLabel: "Time of day",
        yLabel: "Temperature (°C)",
      }
    );

    addCard(s, 6.40, CONTENT_TOP + 0.30, 3.10, 2.80, { strip: C.SUCCESS, fill: C.WHITE });
    s.addText([
      { text: "Trend:", options: { fontSize: 14, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "Up (8 am -> 2 pm)", options: { fontSize: 13, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "Down (2 pm -> 6 pm)", options: { fontSize: 13, color: C.ALERT, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Conclusion:", options: { fontSize: 14, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "Warmed up in the morning. Peaked at 2 pm. Cooled in the evening.", options: { fontSize: 11, color: C.CHARCOAL, italic: true } },
    ], {
      x: 6.55, y: CONTENT_TOP + 0.45, w: 2.85, h: 2.50,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_TREND);
  })();

  // Slides 12-13: CFU + reveal — plant growth between weeks
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Growth from week 3 to week 4 = ?", { color: C.ALERT });
      addTextOnShape(s, "CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      drawLineGraph(s,
        0.85, CONTENT_TOP + 0.40,
        5.30, 2.55,
        PLANT_DATA, {
          lineColor: C.SUCCESS,
          xLabel: "Week",
          yLabel: "Plant height (cm)",
        }
      );

      addCard(s, 6.40, CONTENT_TOP + 0.30, 3.10, 2.80, { strip: C.ALERT, fill: C.WHITE });
      s.addText([
        { text: "Plant heights:", options: { fontSize: 14, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "Week 3 = 7 cm", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Week 4 = 11 cm", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "On your board:", options: { fontSize: 14, color: C.ALERT, bold: true, breakLine: true } },
        { text: "Growth = ___ cm", options: { fontSize: 15, color: C.CHARCOAL, italic: true } },
      ], {
        x: 6.55, y: CONTENT_TOP + 0.45, w: 2.85, h: 2.50,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Growth = 11 - 7 = 4 cm.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do — library borrowing graph
  withReveal(
    () => {
      const s = pres.addSlide();
      const stageColor = STAGE_COLORS["3"];
      addTopBar(s, stageColor);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Library borrowing this week", { color: stageColor });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

      drawLineGraph(s,
        0.85, CONTENT_TOP + 0.40,
        5.30, 2.55,
        LIBRARY_DATA, {
          lineColor: C.SECONDARY,
          xLabel: "Day",
          yLabel: "Books borrowed",
        }
      );

      addCard(s, 6.40, CONTENT_TOP + 0.30, 3.10, 2.80, { strip: stageColor, fill: C.WHITE });
      s.addText([
        { text: "Together:", options: { fontSize: 14, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "1. Which day DROPPED most?", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "2. Which day was HIGHEST?", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "3. Why might these be?", options: { fontSize: 13, color: C.CHARCOAL, italic: true } },
      ], {
        x: 6.55, y: CONTENT_TOP + 0.45, w: 2.85, h: 2.50,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Drop = Wed (30 to 18). Highest = Fri (35). Reasons: assembly Wed, weekend borrowing Fri.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // Slide 16: You Do
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["4"];
    addTopBar(s, stageColor);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: three graphs to read", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    s.addText([
      { text: "First: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "read off the value asked.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Next: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "describe the trend (up/down/flat).", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Then: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "write one conclusion in your own words.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Extension: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "explain Graph 3 with two possible reasons.", options: { fontSize: 15, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    // Right preview: a small line graph thumbnail
    const px = 5.30, py = CONTENT_TOP + 0.20;
    const pw = 4.10, ph = 3.10;
    addCard(s, px, py, pw, ph, { strip: stageColor, fill: C.WHITE });
    drawLineGraph(s,
      px + 0.20, py + 0.40,
      pw - 0.40, ph - 1.0,
      TEMP_DATA, {
        lineColor: C.ALERT,
        xLabel: "Time",
      }
    );
    s.addText("Use the Line Graph Reading sheet.", {
      x: px, y: SAFE_BOTTOM - 0.40, w: pw, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 17: Exit Ticket - look at a small graph
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "Exit Ticket", { color: C.ALERT });
    addTextOnShape(s, "Assesses SC2", {
      x: 2.65, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.2 },
    }, { fontSize: 12, fontFace: FONT_B, color: C.ALERT, bold: true });
    addTitle(s, "Look at the temperature graph", { color: C.ALERT });

    drawLineGraph(s,
      0.85, CONTENT_TOP + 0.20,
      5.30, 2.55,
      TEMP_DATA, {
        lineColor: C.ALERT,
        xLabel: "Time of day",
        yLabel: "Temperature (°C)",
      }
    );

    addCard(s, 6.40, CONTENT_TOP + 0.10, 3.10, 2.80, { strip: C.ALERT, fill: C.WHITE });
    s.addText([
      { text: "On your board:", options: { fontSize: 14, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 10, breakLine: true } },
      { text: "Q1. What was the temperature at 2 pm?", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 10, breakLine: true } },
      { text: "Q2. When did the temperature start to fall?", options: { fontSize: 13, color: C.CHARCOAL } },
    ], {
      x: 6.55, y: CONTENT_TOP + 0.25, w: 2.85, h: 2.50,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_EXIT);
  })();

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what does a line going UP on a line graph mean?",
      scItems: [
        "I can read one value off a line graph.",
        "I can describe when the value rose, fell or stayed the same.",
        "I can draw a sensible conclusion about the situation behind the graph.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "GD_Lesson7_Line_Graphs.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Read three line graphs. Describe the trend. Write a conclusion.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "To read a value: go UP from the time on the X axis, then ACROSS to the Y axis. " +
      "Describe the trend with words like 'rose', 'fell', or 'stayed the same'. A conclusion is a sensible statement about the situation.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Graph 1 — School rainfall (mm) over a week", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Data: Mon 2 mm, Tue 5 mm, Wed 12 mm, Thu 8 mm, Fri 1 mm, Sat 0 mm, Sun 4 mm.", y);
    y = addWriteLine(doc, "a)  On which day was rainfall the highest?  ____________", y);
    y = addWriteLine(doc, "b)  Describe the trend from Wed to Sat.    ________________________________________", y);
    y = addWriteLine(doc, "c)  Conclusion: ________________________________________________________________", y);

    y = addSectionHeading(doc, "Graph 2 — Daily screen time (min) over a week", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Data: Mon 45, Tue 50, Wed 40, Thu 55, Fri 80, Sat 120, Sun 100.", y);
    y = addWriteLine(doc, "a)  Which days had the most screen time?  ____________", y);
    y = addWriteLine(doc, "b)  Describe the trend from weekday to weekend.  ____________________________________", y);
    y = addWriteLine(doc, "c)  Conclusion: ________________________________________________________________", y);

    y = addSectionHeading(doc, "Graph 3 — Class attendance over a term", y, { color: C.ACCENT });
    y = addBodyText(doc, "Data: Wk1 24, Wk2 25, Wk3 22, Wk4 18, Wk5 21, Wk6 25, Wk7 24, Wk8 23.", y);
    y = addWriteLine(doc, "a)  In which week did the most students miss school?  ____________", y);
    y = addWriteLine(doc, "b)  Describe the trend across the term.  ____________________________________", y);
    y = addBodyText(doc, "c)  Stretch: write TWO possible reasons for the drop in week 4, and pick which seems more likely.", y);
    y = addWriteLine(doc, "_____________________________________________________________", y);
    y = addWriteLine(doc, "_____________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Line Graph Reading | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Sample answers for the Line Graph Reading sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Graph 1 — School rainfall", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Wednesday (12 mm).", y);
    y = addBodyText(doc, "b)  The line FELL from Wednesday to Saturday - rainfall decreased to almost zero.", y);
    y = addBodyText(doc, "c)  Sample: The middle of the week was the rainiest, with the rain easing off into the weekend.", y);

    y = addSectionHeading(doc, "Graph 2 — Daily screen time", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Saturday (120 min) and Sunday (100 min).", y);
    y = addBodyText(doc, "b)  Screen time ROSE across the week and was highest on the weekend.", y);
    y = addBodyText(doc, "c)  Sample: Students used screens more on the weekend than on school days.", y);

    y = addSectionHeading(doc, "Graph 3 — Class attendance", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  Week 4 - only 18 students attended.", y);
    y = addBodyText(doc, "b)  Attendance was mostly steady around 22-25, with a drop in week 4 and a recovery after.", y);
    y = addBodyText(doc, "c)  Sample possible reasons: a flu outbreak, a public holiday day, or a school camp absence. The flu outbreak fits because there is a recovery rather than a single missing day.", y);

    y = addTipBox(doc,
      "Watch for: students who report the height of the line instead of the change; students who use 'goes up and down' without naming when. Push for time language - 'between Wednesday and Friday', 'after week 4'.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 7 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
