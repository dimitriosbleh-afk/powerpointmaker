"use strict";

// Geometry and Data Unit (Year 5/6 Numeracy) - Lesson 8: Critiquing media data.
// Year 6 VC2M6ST02 - identify statistically informed arguments in traditional and digital media;
// discuss and critique methods, data representations and conclusions.
// Final lesson of the unit.
// Daily Review: Adding and subtracting fractions and decimals (mixed retrieval).
// Fluency: Multi-digit multiplication / division puzzles.

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

const SESSION = 8;
const TOTAL = 8;
const UNIT_TITLE = "Geometry and Data: Shapes, Tessellations, Data";
const FOOTER = `Geometry and Data | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/GD_Lesson8_Critiquing_Media_Data";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 8 Spot the Trick",
  "Critique three media-style data claims. Decide fair or misleading.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 8 Answer Key",
  "Worked critiques for the Spot the Trick sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// =====================================================================
// Teacher notes
// =====================================================================

const NOTES_TITLE = `SAY:
- Last lesson of our unit.
- Today we look at graphs the way news websites and ads use them.
- Sometimes the same data can be drawn in a fair way OR a misleading way.
- Our job - spot the trick.

DO:
- Settle.
- Whiteboards ready.

TEACHER NOTES:
Lesson 8 of 8. This is the critical-thinking payoff. Year 6 ST02 explicitly calls for students to critique methods, data representations and conclusions. Build curiosity, not cynicism.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- Whiteboards on desks.
- One Spot the Trick sheet per student.
- A few real news graphs printed off if you can - any will do.

DO:
- Print one student sheet per student.
- Optional: bring a recent news graph from your phone or a printed page.

TEACHER NOTES:
The lesson works without real news graphs but is richer with them. Look for a graph with a truncated Y axis or a small sample size - they are very common.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review.
- Add or subtract on your whiteboard.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Final retrieval of fractions and decimals from the week. Mixed difficulty.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 5/6 - 1/3 = 5/6 - 2/6 = 3/6 = 1/2.
- 3.45 + 1.8 = 5.25.
- 4 - 0.65 = 3.35.

DO:
- Click to reveal.

TEACHER NOTES:
The first one needs a common denominator. The third one needs a borrow / regroup. Note any student who got all three.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Multi-digit puzzles.
- Last fluency of the unit.

DO:
- Display the three prompts.
- 30 seconds.

TEACHER NOTES:
Day eight. Stretch fluency - multi-digit. Mix mult and div including a missing-factor.

WATCH FOR:
- Students who decompose 15 x 12 as 15 x 10 + 15 x 2 - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 15 x 12 = 180.
- 132 divided by 11 = 12.
- 8 x ___ = 96  ->  12.

DO:
- Click to reveal.

TEACHER NOTES:
Final reveal. Congratulate the class on the week of fluency work.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Quick prompt. When you see a news headline like "Crime up 50%", what does that actually mean?
- Tell your partner what you would want to know before believing it.

DO:
- 30 seconds.
- Take a few answers.

TEACHER NOTES:
Surface ideas - up from what? Where? How was it measured? Who collected the data? These are exactly the critical-thinking moves we teach today.

WATCH FOR:
- Students who ask "from when?" or "where?" - secure.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to critique data and graphs we see in the media.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 - notice when a Y axis does not start at zero. SC2 - core target - identify at least one issue with a media claim. SC3 - explain in words why the trick changes the picture.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Three key words today.
- Misleading - true facts presented in a way that gives the wrong impression.
- Sample size - how many people or items the data is based on.
- Source - who collected the data and where it came from.

DO:
- Point at each visual.

TEACHER NOTES:
"Misleading" is not "lying". The data can be technically correct but PRESENTED in a way that misleads. That difference matters.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_IDO_AXIS = `SAY:
- Watch me. Same data, two graphs.
- Both show class test scores rising from 75% in March to 80% in June.
- Graph A starts the Y axis at 0%. Looks like a small rise.
- Graph B starts the Y axis at 70%. Looks like a HUGE rise.
- Same data. Different impression. Graph B is misleading because the Y axis is TRUNCATED.

DO:
- Point at the truncated axis on Graph B.
- Highlight the cut symbol or low starting value.

TEACHER NOTES:
This is the most common media trick. Once students see it, they spot it everywhere. The data is not wrong - the presentation is misleading.

MISCONCEPTIONS:
- Misconception: If the data is correct, the graph is fair.
  Why: Students assume only false numbers can lie.
  Impact: They accept a misleading visual at face value.
  Quick correction: A graph can be technically correct but draw a misleading picture if the axis is truncated or the scale is uneven.

WATCH FOR:
- Students who say "the bars look bigger but the difference is the same" - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_SAMPLE = `SAY:
- Watch me. A second trick - small sample.
- Headline says "8 out of 10 students prefer Brand X".
- Then in tiny print - "Sample size: 10 students."
- That is a TINY sample. Just 10 different students could give a totally different result.
- We should not trust this claim to apply to everyone.

DO:
- Highlight "8 out of 10" and "Sample size: 10".

TEACHER NOTES:
Sample size affects how much we trust the result. 10 students is too few. Even 100 might be too few if students were not chosen fairly.

WATCH FOR:
- Students who can name another small-sample claim - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- A graph shows screen time going from 60 minutes to 65 minutes per day. The Y axis starts at 58 minutes.
- What is the trick? Write the name.
- A: cherry-picking
- B: truncated Y axis
- C: tiny sample size

DO:
- Display the prompt.
- 45 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me on three, two, one, show."
- Scan for: B.
PROCEED: If 80% have B, click to reveal and move to We Do.
PIVOT: Most likely misconception - students choose C because the numbers feel small.
- Reteach: "Sample size is about how many people. The trick here is the Y AXIS - it does not start at zero, so a 5 minute rise looks huge."
- Re-check: Show another graph with a Y axis starting at 50%. What is the trick?

TEACHER NOTES:
B - the Y axis starting at 58 minutes makes a small change look big.

WATCH FOR:
- Students who write B - secure.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO = `SAY:
- With me. A news headline.
- "Tap Water Less Popular Than Bottled Water - Survey Reveals Truth."
- The article says 6 out of 10 people prefer bottled water.
- The survey - 10 people, all leaving a bottled water shop.
- What is wrong here?

DO:
- Read the headline aloud.
- Ask the class for issues.

TEACHER NOTES:
Two issues stack here - tiny sample (10) AND biased sample (people leaving a bottled water shop are more likely to prefer it). Students may spot one or both.

WATCH FOR:
- Students who name the bias - secure SC3.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Two issues.
- 1. Sample size is tiny (only 10 people).
- 2. Sample is BIASED - they only asked people leaving a bottled water shop. Those people already chose bottled water.

DO:
- Click to reveal.

TEACHER NOTES:
The conclusion of the article does not follow from the data. This is what Year 6 ST02 asks students to do - critique conclusions.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn. Take the Spot the Trick sheet.
- Three short claims, each with a small picture or numbers.
- For each: identify the trick and explain in your own words why it is misleading.

DO:
- Distribute the sheet.
- Circulate. Ask students to point at the part that gives it away.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Stick with Claim 1. Use the prompt - is the Y axis at zero?
- Extra Notes: Sit with these students. Compare the bars side by side.
EXTENDING PROMPT:
- Task: Find a real graph in a news app or print article. Identify any issues. Sketch a fair version.
- Extra Notes: Year 6 ST02 explicitly asks for this.

TEACHER NOTES:
Different content from We Do. Same three critical-thinking moves: spot, name, explain.

WATCH FOR:
- Students who use words like "truncated", "tiny sample", "biased" - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Q1: A graph shows fitness improvement of 2 points but the Y axis starts at 60 not 0. Name the trick.
- Q2: A claim says "90% love this drink" but only 12 people were asked. Name the trick.

DO:
- Display the prompt.
- 90 seconds.

TEACHER NOTES:
Exit ticket assesses SC2. Look for: truncated Y axis; small / tiny sample size.

WATCH FOR:
- Students who answer both correctly - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- This is the end of our unit.
- Look at the success criteria for today.
- Show me thumbs up, sideways, or down.
- Now - across the WHOLE unit - turn and tell your partner ONE thing you can do now that you could not do at the start of last week.

DO:
- Read each I can statement.
- Use thumbs.
- Take 60 seconds for partner reflection.

TEACHER NOTES:
This is also the end of the unit. Celebrate: transformations, symmetry, nets, tessellations, real data collection, column graphs, line graphs, critical thinking.

[General: Closing | VTLM 2.0: Reflection]`;

// =====================================================================
// Helpers - two-bar comparison graphs
// =====================================================================

// Draw a simple two-bar graph used for the truncated-axis demo.
function drawTwoBar(slide, x, y, w, h, yMin, yMax, valueA, valueB, opts) {
  const o = opts || {};
  const padL = 0.50;
  const padB = 0.50;
  const plotX = x + padL;
  const plotY = y;
  const plotW = w - padL;
  const plotH = h - padB;

  // Axes
  slide.addShape("line", { x: plotX, y: plotY, w: 0, h: plotH, line: { color: C.CHARCOAL, width: 1.5 } });
  slide.addShape("line", { x: plotX, y: plotY + plotH, w: plotW, h: 0, line: { color: C.CHARCOAL, width: 1.5 } });

  // Y axis labels (yMin, mid, yMax)
  const ticks = [yMin, Math.round((yMin + yMax) / 2), yMax];
  ticks.forEach((v) => {
    const ty = plotY + plotH - ((v - yMin) / (yMax - yMin)) * plotH;
    slide.addShape("line", { x: plotX - 0.06, y: ty, w: 0.06, h: 0, line: { color: C.CHARCOAL, width: 1.0 } });
    slide.addText(String(v) + "%", {
      x: plotX - 0.50, y: ty - 0.12, w: 0.45, h: 0.24,
      fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL,
      align: "right", valign: "middle", margin: 0,
    });
  });

  // Two bars (March and June)
  const colW = plotW / 2;
  const barW = colW * 0.55;
  const xA = plotX + colW / 2 - barW / 2;
  const xB = plotX + 1.5 * colW - barW / 2;

  const heightA = ((valueA - yMin) / (yMax - yMin)) * plotH;
  const heightB = ((valueB - yMin) / (yMax - yMin)) * plotH;

  slide.addShape("rect", {
    x: xA, y: plotY + plotH - heightA, w: barW, h: heightA,
    fill: { color: o.colorA || C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
  });
  slide.addShape("rect", {
    x: xB, y: plotY + plotH - heightB, w: barW, h: heightB,
    fill: { color: o.colorB || C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
  });

  // Bar values above bars
  slide.addText(String(valueA) + "%", {
    x: xA - 0.20, y: plotY + plotH - heightA - 0.30, w: barW + 0.40, h: 0.25,
    fontSize: 10, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", margin: 0,
  });
  slide.addText(String(valueB) + "%", {
    x: xB - 0.20, y: plotY + plotH - heightB - 0.30, w: barW + 0.40, h: 0.25,
    fontSize: 10, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", margin: 0,
  });

  // X labels
  slide.addText(o.labelA || "A", {
    x: plotX, y: plotY + plotH + 0.05, w: colW, h: 0.30,
    fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
    align: "center", valign: "top", margin: 0,
  });
  slide.addText(o.labelB || "B", {
    x: plotX + colW, y: plotY + plotH + 0.05, w: colW, h: 0.30,
    fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
    align: "center", valign: "top", margin: 0,
  });
}

// =====================================================================
// Build
// =====================================================================

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 8: Spot the trick (critiquing media data)",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions and decimals",
      ["5/6 - 1/3", "3.45 + 1.8", "4 - 0.65"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1/2         5.25         3.35", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal
  withReveal(
    () => fluencySlide(pres, "Fluency: Multi-digit puzzles",
      ["15 x 12", "132 div 11", "8 x ? = 96"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "180       12       12", {
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
    addTitle(s, "\"Crime up 50%!\" — what would you ask?");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    // Big headline-style quote
    addTextOnShape(s, "Crime up 50%!", {
      x: 1.5, y: CONTENT_TOP + 0.45, w: 7.0, h: 0.90, rectRadius: 0.10,
      fill: { color: C.BG_DARK || C.PRIMARY },
    }, { fontSize: 32, fontFace: FONT_H, color: C.WHITE, bold: true });

    s.addText([
      { text: "Tell your partner: what would you want to know before believing this?", options: { fontSize: 16, color: C.CHARCOAL, italic: true, breakLine: true } },
      { text: "", options: { fontSize: 10, breakLine: true } },
      { text: "(Up from what? Where? How many crimes? Over what time period? Who counted?)", options: { fontSize: 12, color: C.MUTED, italic: true } },
    ], {
      x: 1.0, y: CONTENT_TOP + 1.60, w: 8.0, h: 1.40,
      fontFace: FONT_B, valign: "top", margin: 0, align: "center",
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to critique data and graphs we see in the media.",
    [
      "I can notice when a Y axis does not start at zero.",
      "I can identify at least one issue with a media claim - the axis, the sample, or the conclusion.",
      "I can explain in my own words why the trick changes the picture the graph gives.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key vocabulary
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Misleading, sample size, source");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const items = [
      { word: "Misleading",  simple: "wrong impression", detail: "True facts presented to give the wrong picture.", color: C.PRIMARY },
      { word: "Sample size", simple: "how many",         detail: "Number of people or items in the data.",         color: C.ACCENT },
      { word: "Source",      simple: "who collected it", detail: "Where the data came from and who reported it.",   color: C.ALERT },
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
        fontSize: 15, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
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

  // Slide 10: I Do — Same data, two graphs (truncated axis)
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Same data. Two graphs.", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Graph A — Y axis from 0
    const gA_x = 0.85;
    const gA_y = CONTENT_TOP + 0.30;
    const gA_w = 4.10;
    const gA_h = 2.50;
    drawTwoBar(s, gA_x, gA_y, gA_w, gA_h, 0, 100, 75, 80, {
      colorA: C.PRIMARY, colorB: C.PRIMARY, labelA: "March", labelB: "June",
    });
    s.addText("Graph A — Y axis from 0%", {
      x: gA_x, y: gA_y + gA_h + 0.05, w: gA_w, h: 0.30,
      fontSize: 12, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });
    s.addText("Small change. Looks fair.", {
      x: gA_x, y: gA_y + gA_h + 0.35, w: gA_w, h: 0.25,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", margin: 0,
    });

    // Graph B — Y axis from 70
    const gB_x = 5.05;
    const gB_y = CONTENT_TOP + 0.30;
    const gB_w = 4.10;
    const gB_h = 2.50;
    drawTwoBar(s, gB_x, gB_y, gB_w, gB_h, 70, 82, 75, 80, {
      colorA: C.ALERT, colorB: C.ALERT, labelA: "March", labelB: "June",
    });
    s.addText("Graph B — Y axis from 70%", {
      x: gB_x, y: gB_y + gB_h + 0.05, w: gB_w, h: 0.30,
      fontSize: 12, fontFace: FONT_H, color: C.ALERT, bold: true,
      align: "center", margin: 0,
    });
    s.addText("Looks huge. Same data. TRUNCATED.", {
      x: gB_x, y: gB_y + gB_h + 0.35, w: gB_w, h: 0.25,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, italic: true,
      align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_AXIS);
  })();

  // Slide 11: I Do — Small sample claim
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "\"8 out of 10 students prefer Brand X\"", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    addTextOnShape(s, "8 out of 10 students prefer Brand X", {
      x: 1.0, y: CONTENT_TOP + 0.40, w: 8.0, h: 0.80, rectRadius: 0.10,
      fill: { color: C.BG_LIGHT }, line: { color: stageColor, width: 1.5 },
    }, { fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center" });

    s.addText("Sample size: 10 students.", {
      x: 1.0, y: CONTENT_TOP + 1.40, w: 8.0, h: 0.40,
      fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", margin: 0,
    });

    addCard(s, 1.5, CONTENT_TOP + 1.95, 7.0, 1.20, { strip: C.ALERT, fill: C.WHITE });
    s.addText([
      { text: "Why this is misleading:", options: { fontSize: 14, color: C.ALERT, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "Only 10 students were asked. Just 10 different students could give a very different result. We cannot trust this claim to apply to everyone.", options: { fontSize: 12, color: C.CHARCOAL, italic: true } },
    ], {
      x: 1.65, y: CONTENT_TOP + 2.10, w: 6.70, h: 1.00,
      fontFace: FONT_B, valign: "top", margin: 0, align: "left",
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_SAMPLE);
  })();

  // Slides 12-13: CFU + reveal — name the trick
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Name the trick", { color: C.ALERT });
      addTextOnShape(s, "CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      s.addText([
        { text: "A graph shows screen time going from 60 to 65 minutes per day.", options: { fontSize: 16, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "The Y axis starts at 58 minutes.", options: { fontSize: 16, color: C.CHARCOAL, italic: true, breakLine: true } },
        { text: "", options: { fontSize: 14, breakLine: true } },
        { text: "What is the trick?", options: { fontSize: 16, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "A.  cherry-picking", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "B.  truncated Y axis", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "C.  tiny sample size", options: { fontSize: 15, color: C.CHARCOAL } },
      ], {
        x: 1.0, y: CONTENT_TOP + 0.30, w: 8.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.60,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "B. Truncated Y axis. A 5-min change looks huge because the axis starts at 58.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do — Headline critique
  withReveal(
    () => {
      const s = pres.addSlide();
      const stageColor = STAGE_COLORS["3"];
      addTopBar(s, stageColor);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Critique this headline together", { color: stageColor });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

      // Headline panel
      addTextOnShape(s, "Tap Water Less Popular Than Bottled Water!", {
        x: 0.85, y: CONTENT_TOP + 0.30, w: 8.30, h: 0.70, rectRadius: 0.10,
        fill: { color: C.BG_DARK || stageColor },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true, align: "center" });

      // Body text
      s.addText([
        { text: "The article says: 6 out of 10 people prefer bottled water.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "The survey: 10 people, all leaving a bottled water shop.", options: { fontSize: 14, color: C.CHARCOAL, italic: true, breakLine: true } },
        { text: "", options: { fontSize: 12, breakLine: true } },
        { text: "What is wrong here? Find TWO issues with your partner.", options: { fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: 1.0, y: CONTENT_TOP + 1.20, w: 8.0, h: 2.0,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1. Sample too small (10 people). 2. Sample BIASED - bottled water shoppers.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // Slide 16: You Do
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["4"];
    addTopBar(s, stageColor);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: three claims to critique", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    s.addText([
      { text: "First: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "look at each claim on the sheet.", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Next: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "decide: fair or misleading?", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Then: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "name the trick (or say why it is fair).", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Then: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "write one sentence to explain.", options: { fontSize: 15, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    // Right preview: three short claim cards
    const px = 5.30, py = CONTENT_TOP + 0.20;
    const pw = 4.10, ph = 3.10;
    addCard(s, px, py, pw, ph, { strip: stageColor, fill: C.WHITE });
    const claimY = py + 0.20;
    const claimH = (ph - 0.40) / 3 - 0.05;
    [
      { label: "1. Y axis trick", text: "Lunch line shorter by 1 min. Y axis starts at 5 min." },
      { label: "2. Tiny sample", text: "\"75% love it!\" Sample size: 4 people." },
      { label: "3. Cherry-picked time", text: "\"Sales up 200%!\" Compared 2-day sale week to a quiet week." },
    ].forEach((c, i) => {
      const cy = claimY + i * (claimH + 0.05);
      s.addShape("rect", {
        x: px + 0.20, y: cy, w: pw - 0.40, h: claimH,
        fill: { color: C.BG_LIGHT }, line: { color: stageColor, width: 1 },
      });
      s.addText(c.label, {
        x: px + 0.30, y: cy + 0.05, w: pw - 0.60, h: 0.25,
        fontSize: 11, fontFace: FONT_H, color: stageColor, bold: true,
        align: "left", margin: 0,
      });
      s.addText(c.text, {
        x: px + 0.30, y: cy + 0.30, w: pw - 0.60, h: claimH - 0.35,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "left", valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "A fitness graph shows a 2 point rise but the Y axis starts at 60, not 0. Name the trick.",
      "A claim says \"90% love this drink\" but only 12 people were asked. Name the trick.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing — unit reflection
  closingSlide(pres,
    {
      reflectionPrompt: "Across the WHOLE unit: turn and tell your partner ONE thing you can do now that you could not at the start of last week.",
      scItems: [
        "I can notice when a Y axis does not start at zero.",
        "I can identify at least one issue with a media claim - the axis, the sample, or the conclusion.",
        "I can explain in my own words why the trick changes the picture the graph gives.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "GD_Lesson8_Critiquing_Media_Data.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Three claims. Spot the trick. Write a short critique.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Three common tricks to look for: " +
      "(1) Y axis that does NOT start at zero. " +
      "(2) Sample size too small. " +
      "(3) A time period or group chosen unfairly to make the change look bigger.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Claim 1 — School lunch line", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Headline: \"Lunch lines SLASHED in half!\" A graph shows wait time falling from 6 minutes to 5 minutes. The Y axis on the graph starts at 5 minutes.", y);
    y = addWriteLine(doc, "Fair or misleading?  ____________", y);
    y = addWriteLine(doc, "Name the trick:      ____________________________", y);
    y = addWriteLine(doc, "Why is it misleading? ___________________________________________________", y);

    y = addSectionHeading(doc, "Claim 2 — Energy drink survey", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Headline: \"75% love new energy drink!\" The small print on the ad says: sample size = 4 people.", y);
    y = addWriteLine(doc, "Fair or misleading?  ____________", y);
    y = addWriteLine(doc, "Name the trick:      ____________________________", y);
    y = addWriteLine(doc, "Why is it misleading? ___________________________________________________", y);

    y = addSectionHeading(doc, "Claim 3 — Shop sales report", y, { color: C.ACCENT });
    y = addBodyText(doc, "Headline: \"Sales UP 200% in one week!\" The fine print: the shop compared a two-day stocktake sale week to a normal quiet week.", y);
    y = addWriteLine(doc, "Fair or misleading?  ____________", y);
    y = addWriteLine(doc, "Name the trick:      ____________________________", y);
    y = addWriteLine(doc, "Why is it misleading? ___________________________________________________", y);

    y = addSectionHeading(doc, "Stretch — find your own", y, { color: C.ACCENT });
    y = addBodyText(doc, "Find a graph in a news app or print article. Stick or sketch it here. Identify any issues. Then sketch a fair version.", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Spot the Trick | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked critiques for the Spot the Trick sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Claim 1 — Lunch line", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Misleading.", y);
    y = addBodyText(doc, "Trick: truncated Y axis.", y);
    y = addBodyText(doc, "Why: the Y axis starts at 5 minutes, so a 1 minute drop fills almost the whole graph. The wait time only went from 6 to 5 - that is one minute, not 'slashed in half'.", y);

    y = addSectionHeading(doc, "Claim 2 — Energy drink", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Misleading.", y);
    y = addBodyText(doc, "Trick: tiny sample size.", y);
    y = addBodyText(doc, "Why: 4 people is far too few. 75% of 4 is just 3 people. We cannot draw a conclusion about 'everyone' from 4 people.", y);

    y = addSectionHeading(doc, "Claim 3 — Shop sales", y, { color: C.ACCENT });
    y = addBodyText(doc, "Misleading.", y);
    y = addBodyText(doc, "Trick: unfair comparison (cherry-picked time period).", y);
    y = addBodyText(doc, "Why: comparing a stocktake sale week to a normal quiet week. The shop made the sale week look like a huge result. A fair comparison would use two similar weeks.", y);

    y = addSectionHeading(doc, "Mark allocation", y, { color: C.ACCENT });
    y = addBodyText(doc, "SC1 - student spots that the Y axis does not start at zero in Claim 1.", y);
    y = addBodyText(doc, "SC2 - student identifies at least one issue in each claim.", y);
    y = addBodyText(doc, "SC3 - student explains in their own words why the issue changes the picture.", y);

    y = addTipBox(doc,
      "Push students to use the specific names: 'truncated Y axis', 'tiny sample', 'unfair comparison'. Naming the trick is the difference between vague suspicion and statistical critique.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 8 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
