"use strict";

// Geometry and Data Unit (Year 5/6 Numeracy) - Lesson 6: Build the column graph.
// Year 5 VC2M5ST01 + VC2M5ST03 and Year 6 VC2M6ST01.
// Students use yesterday's car-colour tally to build a column graph and analyse mode/range/shape.
// Daily Review: Time.
// Fluency: Division by 10, 100, 1000.

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

const SESSION = 6;
const TOTAL = 8;
const UNIT_TITLE = "Geometry and Data: Shapes, Tessellations, Data";
const FOOTER = `Geometry and Data | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/GD_Lesson6_Column_Graph";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 6 Graph Template",
  "Blank column graph and analysis questions for the car-colour data.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 6 Answer Key",
  "Sample graph and analysis answers using example car-colour data.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// Example data we use for I Do/We Do illustrations. Teachers will substitute the actual class data.
const SAMPLE_DATA = [
  { colour: "White",  count: 12, hex: "EFEFEF" },
  { colour: "Black",  count: 8,  hex: "303030" },
  { colour: "Grey",   count: 10, hex: "808080" },
  { colour: "Silver", count: 5,  hex: "BFBFBF" },
  { colour: "Red",    count: 4,  hex: "B22222" },
  { colour: "Blue",   count: 6,  hex: "1B6FD8" },
  { colour: "Other",  count: 3,  hex: "F2C14E" },
];

// =====================================================================
// Teacher notes
// =====================================================================

const NOTES_TITLE = `SAY:
- Yesterday we collected real data at the boundary.
- Today we turn that data into a graph and read what the graph tells us.

DO:
- Have yesterday's Tally Sheets out per student.
- Whiteboards ready.

TEACHER NOTES:
Lesson 6 of 8. Today is the analytical payoff for yesterday's collection. Students build a column graph, then describe the shape using mode and range.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- Yesterday's Tally Sheets.
- One Graph Template per student.
- Coloured pencils for the column bars.

DO:
- Confirm everyone has their Tally Sheet.
- Have a spare sheet ready if any went missing.

TEACHER NOTES:
The example graph on the slides uses sample data. Students use THEIR actual data on the worksheet.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Time.
- Work out each on your whiteboard.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Revision of time - last week's unit. Mixed: 12 vs 24 hour, elapsed time, and conversion.

WATCH FOR:
- Students who use the number-line jump strategy for elapsed time - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 14:30 in 12-hour time = 2:30 pm.
- 8:15 am to 9:50 am = 1 hour 35 minutes.
- 2 hours 15 minutes = 135 minutes.

DO:
- Click to reveal.

TEACHER NOTES:
Quick reveal. Note any student who needed extra time on the elapsed time.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Dividing by 10, 100, 1000.
- Whisper-answer then write.

DO:
- Display the three prompts.
- 25 seconds.

TEACHER NOTES:
Day six. Today the digits shift RIGHT. The mirror of yesterday's place value strategy.

WATCH FOR:
- Students who name the place value shift - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 720 divided by 10 = 72.
- 4500 divided by 100 = 45.
- 6.8 divided by 10 = 0.68.

DO:
- Click to reveal.

TEACHER NOTES:
The third one stretches - dividing a decimal by 10 shifts each digit one place right and a new digit appears in the tenths.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Quick prediction. Look at your Tally Sheet.
- Without working out the totals exactly - which colour had the longest tally?
- Tell your partner.

DO:
- 30 seconds.

TEACHER NOTES:
This launch primes the prediction step before the graph. A column graph is essentially a tally with the bars upright - the longest tally becomes the tallest column.

WATCH FOR:
- Students who can predict the mode - secure.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to display categorical data using a column graph and describe its mode and range.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 - draw the bars from a frequency table. SC2 - core target - label the graph correctly and read the mode. SC3 - describe the shape and the range and compare distributions.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Three key words today.
- Mode - the value that appears most often. The tallest column.
- Range - the difference between the biggest and the smallest count.
- Frequency - how many times each category appears.

DO:
- Point at each visual.

TEACHER NOTES:
Anchor MODE with the visual of the tallest column. Range is a number, not a column. Frequency is the total in each row.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_IDO_GRAPH = `SAY:
- Watch me build the column graph from a sample tally.
- The categories go along the bottom - the X axis.
- The frequency goes up the side - the Y axis.
- Each category gets one column. The column's height shows the frequency.

DO:
- Build the graph on the slide live or use the on-screen version.
- Label each axis.

TEACHER NOTES:
The labels matter. Without "Colour" on the X axis and "Number of cars" on the Y axis, the graph is incomplete.

WATCH FOR:
- Students who can name both axis labels - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_MODE_RANGE = `SAY:
- Now we read the graph.
- The tallest column is white. So the MODE is white.
- The biggest frequency is 12. The smallest is 3. RANGE = 12 - 3 = 9.
- The shape - mostly tall on the left, smaller on the right.

DO:
- Point at the tallest column.
- Point at the smallest column.
- Subtract to get the range.

TEACHER NOTES:
Mode = the WORD that wins. Range = a NUMBER (a difference). Students often confuse the two.

MISCONCEPTIONS:
- Misconception: Mode is the height of the tallest column.
  Why: They see "biggest number" and report it.
  Impact: They say "the mode is 12".
  Quick correction: Mode is the CATEGORY name. The height is the frequency.

WATCH FOR:
- Students who say "white" not "12" - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- Look at the sample graph.
- What is the MODE? Write a colour, not a number.

DO:
- Display the prompt.
- 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me on three, two, one, show."
- Scan for: White (or "white").
PROCEED: If 80% have White, click to reveal and move to We Do.
PIVOT: Most likely misconception - students write 12.
- Reteach: "Mode is the WORD - the name of the category. 12 is the FREQUENCY."
- Re-check: "Which colour has the second-highest frequency?"

TEACHER NOTES:
Sample data here: White=12, Black=8, Grey=10, Silver=5, Red=4, Blue=6, Other=3. Mode = White.

WATCH FOR:
- Students who write "White" - secure.
- Students who write "12" - need the category vs height correction.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO = `SAY:
- With me. We will calculate the RANGE of the sample data.
- Step 1: find the highest column - White = 12.
- Step 2: find the lowest column - Other = 3.
- Step 3: subtract - 12 minus 3 = 9.
- The range is 9 cars.

DO:
- Walk through the three steps on the slide.
- Ask the class for each step.

TEACHER NOTES:
Highlight that the range is a number with units (cars). The mode is a category. They look different on the page.

WATCH FOR:
- Students who include the unit "cars" - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Range = 12 minus 3 = 9 cars.

DO:
- Click to reveal.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn. Take the Graph Template.
- Step 1: write your totals in the table.
- Step 2: draw your column graph. Label both axes. Give it a title.
- Step 3: write the mode and the range.

DO:
- Distribute the template.
- Circulate. Confirm the axes are labelled before students draw bars.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the pre-labelled axes on Part B of the template. Fill in bars only.
- Extra Notes: Sit with these students. Help with the scale.
EXTENDING PROMPT:
- Task: Combine your data with two other students' data. Build a comparison column graph (theirs vs yours vs combined). Discuss whether the mode is the same.
- Extra Notes: This is the Year 6 ST01 stretch.

TEACHER NOTES:
Different content from We Do - students use their OWN data. Same skill of drawing the graph.

WATCH FOR:
- Students who label both axes BEFORE drawing bars - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Q1: What is the MODE of your data? Write a colour.
- Q2: What is the RANGE of your data? Show the subtraction.

DO:
- Display the prompt.
- 2 minutes.

TEACHER NOTES:
Exit ticket assesses SC2 directly using each student's own data.

WATCH FOR:
- Students who write a colour for Q1 and a number with subtraction for Q2 - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria for today.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: was the mode in your data what you predicted?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea - a column graph turns counts into bar heights so we can read mode, range and shape at a glance. Tomorrow: line graphs.

[General: Closing | VTLM 2.0: Reflection]`;

// =====================================================================
// Helpers
// =====================================================================

// Draw a column graph centred at (x,y) with width w and height h.
function drawColumnGraph(slide, x, y, w, h, data, opts) {
  const o = opts || {};
  const padL = 0.5;
  const padB = 0.5;
  const plotX = x + padL;
  const plotY = y;
  const plotW = w - padL;
  const plotH = h - padB;

  const maxValue = Math.max(...data.map(d => d.count));
  const yMax = Math.ceil(maxValue / 2) * 2 + 2;
  const barW = plotW / data.length * 0.72;
  const colW = plotW / data.length;

  // Y axis
  slide.addShape("line", {
    x: plotX, y: plotY, w: 0, h: plotH,
    line: { color: C.CHARCOAL, width: 1.5 },
  });
  // X axis
  slide.addShape("line", {
    x: plotX, y: plotY + plotH, w: plotW, h: 0,
    line: { color: C.CHARCOAL, width: 1.5 },
  });

  // Y axis ticks/labels (0 to yMax in steps of 2)
  const ticks = [];
  for (let v = 0; v <= yMax; v += 2) ticks.push(v);
  ticks.forEach((v) => {
    const ty = plotY + plotH - (v / yMax) * plotH;
    slide.addShape("line", {
      x: plotX - 0.06, y: ty, w: 0.06, h: 0,
      line: { color: C.CHARCOAL, width: 1.2 },
    });
    slide.addText(String(v), {
      x: plotX - 0.45, y: ty - 0.12, w: 0.36, h: 0.24,
      fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL,
      align: "right", valign: "middle", margin: 0,
    });
  });

  // Bars
  data.forEach((d, i) => {
    const bx = plotX + i * colW + (colW - barW) / 2;
    const bh = (d.count / yMax) * plotH;
    const by = plotY + plotH - bh;
    slide.addShape("rect", {
      x: bx, y: by, w: barW, h: bh,
      fill: { color: d.hex || C.PRIMARY },
      line: { color: C.CHARCOAL, width: 1 },
    });
    // Category label
    slide.addText(d.colour, {
      x: plotX + i * colW, y: plotY + plotH + 0.05, w: colW, h: 0.30,
      fontSize: 10, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "top", margin: 0,
    });
    // Frequency above bar
    if (o.showValues) {
      slide.addText(String(d.count), {
        x: bx - 0.10, y: by - 0.25, w: barW + 0.20, h: 0.22,
        fontSize: 9, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
    }
  });

  // Axis titles
  if (o.xLabel) {
    slide.addText(o.xLabel, {
      x: plotX, y: plotY + plotH + 0.35, w: plotW, h: 0.25,
      fontSize: 11, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", margin: 0,
    });
  }
  if (o.yLabel) {
    // Vertical-ish y label (just place above the y axis)
    slide.addText(o.yLabel, {
      x: x - 0.10, y: plotY - 0.35, w: 1.80, h: 0.25,
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
  titleSlide(pres, UNIT_TITLE, "Lesson 6: Build the column graph",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Time",
      ["14:30 in 12-hour =", "8:15 am to 9:50 am =", "2 hr 15 min = ___ min"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "2:30 pm        1 hr 35 min        135 min", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal
  withReveal(
    () => fluencySlide(pres, "Fluency: Dividing by 10, 100, 1000",
      ["720 div 10", "4500 div 100", "6.8 div 10"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "72       45       0.68", {
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
    addTitle(s, "Predict before you graph");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    s.addText([
      { text: "Without working out exact totals -", options: { fontSize: 18, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 10, breakLine: true } },
      { text: "Which colour had the LONGEST tally on your sheet?", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 12, breakLine: true } },
      { text: "Tell your partner.", options: { fontSize: 16, color: C.ALERT, bold: true } },
    ], {
      x: 1.0, y: CONTENT_TOP + 0.50, w: 8.0, h: 2.50,
      fontFace: FONT_B, valign: "top", margin: 0, align: "center",
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to display categorical data using a column graph and describe its mode and range.",
    [
      "I can draw the bars on a column graph from a frequency table.",
      "I can label both axes and read the mode from the column graph.",
      "I can describe the shape of the data and calculate the range.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key vocabulary
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Mode, range, frequency");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const items = [
      { word: "Mode",      simple: "most common",  detail: "The category that appears most often. The tallest column.", color: C.PRIMARY },
      { word: "Range",     simple: "biggest - smallest", detail: "A NUMBER. The difference between the largest and smallest counts.", color: C.ACCENT },
      { word: "Frequency", simple: "how many",     detail: "The count in one category. The height of one column.", color: C.ALERT },
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
        x: cx + 0.15, y: cardY + 1.50, w: cellW - 0.30, h: 0.95,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 10: I Do - Build the column graph
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "From tally to column graph", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Graph on left
    drawColumnGraph(s,
      0.85, CONTENT_TOP + 0.45,
      5.30, 2.55,
      SAMPLE_DATA, {
        showValues: true,
        xLabel: "Colour",
        yLabel: "Number of cars",
      }
    );

    // Right side checklist
    addCard(s, 6.40, CONTENT_TOP + 0.30, 3.10, 2.80, { strip: C.ACCENT, fill: C.WHITE });
    s.addText([
      { text: "Build checklist:", options: { fontSize: 14, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "+ Title at the top", options: { fontSize: 12, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "+ X axis label (Colour)", options: { fontSize: 12, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "+ Y axis label (Number)", options: { fontSize: 12, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "+ Scale starts at 0", options: { fontSize: 12, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "+ Bars are equal width", options: { fontSize: 12, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "+ Bars have gaps", options: { fontSize: 12, color: C.SUCCESS, bold: true } },
    ], {
      x: 6.55, y: CONTENT_TOP + 0.45, w: 2.85, h: 2.50,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_GRAPH);
  })();

  // Slide 11: I Do - read mode and range from the graph
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Read the mode and range", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    drawColumnGraph(s,
      0.85, CONTENT_TOP + 0.45,
      5.30, 2.55,
      SAMPLE_DATA, {
        showValues: true,
        xLabel: "Colour",
        yLabel: "Number of cars",
      }
    );

    addCard(s, 6.40, CONTENT_TOP + 0.30, 3.10, 2.80, { strip: C.SUCCESS, fill: C.WHITE });
    s.addText([
      { text: "Mode = White", options: { fontSize: 18, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "(tallest column)", options: { fontSize: 11, color: C.MUTED, italic: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Range = 12 - 3", options: { fontSize: 18, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "= 9 cars", options: { fontSize: 18, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Shape: bigger on the cool/neutral colours", options: { fontSize: 11, color: C.CHARCOAL, italic: true } },
    ], {
      x: 6.55, y: CONTENT_TOP + 0.45, w: 2.85, h: 2.50,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_MODE_RANGE);
  })();

  // Slides 12-13: CFU + reveal - what is the MODE?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "What is the MODE? (a colour, not a number)", { color: C.ALERT });
      addTextOnShape(s, "CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      drawColumnGraph(s,
        0.85, CONTENT_TOP + 0.40,
        5.30, 2.55,
        SAMPLE_DATA, {
          showValues: true,
          xLabel: "Colour",
          yLabel: "Number of cars",
        }
      );

      addCard(s, 6.40, CONTENT_TOP + 0.30, 3.10, 2.80, { strip: C.ALERT, fill: C.WHITE });
      s.addText([
        { text: "On your board:", options: { fontSize: 16, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 12, breakLine: true } },
        { text: "What is the MODE?", options: { fontSize: 16, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Write a colour, not a number.", options: { fontSize: 13, color: C.CHARCOAL, italic: true } },
      ], {
        x: 6.55, y: CONTENT_TOP + 0.50, w: 2.85, h: 2.50,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Mode = White. The tallest column wins.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do — calculate the range
  withReveal(
    () => {
      const s = pres.addSlide();
      const stageColor = STAGE_COLORS["3"];
      addTopBar(s, stageColor);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Together: calculate the range", { color: stageColor });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

      drawColumnGraph(s,
        0.85, CONTENT_TOP + 0.40,
        5.30, 2.55,
        SAMPLE_DATA, {
          showValues: true,
          xLabel: "Colour",
          yLabel: "Number of cars",
        }
      );

      addCard(s, 6.40, CONTENT_TOP + 0.30, 3.10, 2.80, { strip: stageColor, fill: C.WHITE });
      s.addText([
        { text: "Three steps together:", options: { fontSize: 14, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "1. Highest count = ___", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "2. Lowest count = ___", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "3. Range = ___ - ___ = ___", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 6.55, y: CONTENT_TOP + 0.50, w: 2.85, h: 2.50,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Range = 12 - 3 = 9 cars.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // Slide 16: You Do
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["4"];
    addTopBar(s, stageColor);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: graph your own data", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    s.addText([
      { text: "First: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "write your totals in the table on the sheet.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Next: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "draw your column graph. Label both axes.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Then: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "write the mode and the range.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Extension: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "compare with two other students' data.", options: { fontSize: 15, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    // Right preview: blank graph axes
    const px = 5.30, py = CONTENT_TOP + 0.20;
    const pw = 4.10, ph = 3.10;
    addCard(s, px, py, pw, ph, { strip: stageColor, fill: C.WHITE });
    // Axes
    s.addShape("line", {
      x: px + 0.50, y: py + 0.40, w: 0, h: ph - 1.0,
      line: { color: C.CHARCOAL, width: 1.5 },
    });
    s.addShape("line", {
      x: px + 0.50, y: py + ph - 0.60, w: pw - 0.80, h: 0,
      line: { color: C.CHARCOAL, width: 1.5 },
    });
    s.addText("Colour", {
      x: px + 0.50, y: py + ph - 0.40, w: pw - 0.80, h: 0.20,
      fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", margin: 0,
    });
    s.addText("Number", {
      x: px + 0.10, y: py + 0.20, w: 0.80, h: 0.20,
      fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "left", margin: 0,
    });

    s.addText("Use the Graph Template.", {
      x: px, y: SAFE_BOTTOM - 0.40, w: pw, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "What is the MODE of your data? Write a colour.",
      "What is the RANGE of your data? Show your subtraction.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: was the mode in your data what you predicted at the start?",
      scItems: [
        "I can draw the bars on a column graph from a frequency table.",
        "I can label both axes and read the mode from the column graph.",
        "I can describe the shape of the data and calculate the range.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "GD_Lesson6_Column_Graph.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Use your tally totals. Draw the column graph. Read mode and range.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Title at the top. Both axes labelled. Scale starts at 0 with equal steps. Bars equal width with small gaps between them.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Part A — Your totals", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Fill in the total for each colour from your Tally Sheet.", y);
    y = addBodyText(doc, "Colour       |  Total", y, { fontSize: 11 });
    y = addBodyText(doc, "White        |", y, { fontSize: 11 });
    y = addBodyText(doc, "Black        |", y, { fontSize: 11 });
    y = addBodyText(doc, "Grey         |", y, { fontSize: 11 });
    y = addBodyText(doc, "Silver       |", y, { fontSize: 11 });
    y = addBodyText(doc, "Red          |", y, { fontSize: 11 });
    y = addBodyText(doc, "Blue         |", y, { fontSize: 11 });
    y = addBodyText(doc, "Other        |", y, { fontSize: 11 });

    y = addSectionHeading(doc, "Part B — Draw your column graph", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Draw your graph on grid paper or in the box on the next page. Remember the labels and title.", y);
    y = addBodyText(doc, "Title: ____________________________________________________________________", y);
    y = addBodyText(doc, "Y axis (vertical): __________________________________", y);
    y = addBodyText(doc, "X axis (horizontal): __________________________________", y);

    y = addSectionHeading(doc, "Part C — Mode and range", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Mode (colour): ____________", y);
    y = addWriteLine(doc, "Highest count: ___    Lowest count: ___    Range = ___ - ___ = ___", y);

    y = addSectionHeading(doc, "Part D — Describe the shape", y, { color: C.ACCENT });
    y = addBodyText(doc, "In one or two sentences, what does the shape of your graph tell us?", y);
    y = addWriteLine(doc, "_____________________________________________________________", y);
    y = addWriteLine(doc, "_____________________________________________________________", y);

    y = addSectionHeading(doc, "Stretch — compare", y, { color: C.ACCENT });
    y = addBodyText(doc, "Find two other students. Compare your modes. Are they the same? Why might they be different?", y);
    y = addWriteLine(doc, "_____________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Graph Template | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Sample graph and analysis using example data.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Example totals", y, { color: C.PRIMARY });
    y = addBodyText(doc, "White = 12, Black = 8, Grey = 10, Silver = 5, Red = 4, Blue = 6, Other = 3.", y);
    y = addBodyText(doc, "Total cars observed = 48.", y);

    y = addSectionHeading(doc, "Example mode and range", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Mode (most common colour) = WHITE.", y);
    y = addBodyText(doc, "Highest count = 12 (White). Lowest count = 3 (Other).", y);
    y = addBodyText(doc, "Range = 12 - 3 = 9 cars.", y);

    y = addSectionHeading(doc, "Example shape description", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Most cars were neutral colours (white, black, grey). Bright colours (red, blue, other) were less common.", y);

    y = addSectionHeading(doc, "Mark allocation", y, { color: C.ACCENT });
    y = addBodyText(doc, "SC1 - student has drawn bars. Award if all 7 columns are present and proportional to the totals.", y);
    y = addBodyText(doc, "SC2 - student labelled both axes. Award if 'Colour' on X and 'Number of cars' (or equivalent) on Y.", y);
    y = addBodyText(doc, "SC3 - student calculated range with subtraction and described the shape.", y);

    y = addTipBox(doc,
      "Common errors: writing the mode as a number instead of a colour; forgetting axis labels; starting the Y axis above 0; making the bars different widths.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 6 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
