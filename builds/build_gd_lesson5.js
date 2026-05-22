"use strict";

// Geometry and Data Unit (Year 5/6 Numeracy) - Lesson 5: Plan the car-colour data investigation.
// Year 5 VC2M5ST01 + VC2M5ST03 and Year 6 VC2M6ST03.
// Students plan and run a 10-minute boundary count of car colours.
// Daily Review: Converting measurement units.
// Fluency: x10, x100 (place value strategy).

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

const SESSION = 5;
const TOTAL = 8;
const UNIT_TITLE = "Geometry and Data: Shapes, Tessellations, Data";
const FOOTER = `Geometry and Data | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/GD_Lesson5_Data_Collection_Plan";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 5 Tally Sheet",
  "Tally chart for the 10-minute car colour observation at the school boundary.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 5 Teacher Notes",
  "Safety, logistics and discussion prompts for the boundary observation.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// =====================================================================
// Teacher notes
// =====================================================================

const NOTES_TITLE = `SAY:
- Last week we looked at shapes. Now we move into DATA.
- For three sessions we are statistical investigators. We will collect real data, build a graph, and read graphs from the news.
- Today we plan our investigation. Then we head out to the boundary to count.

DO:
- Settle.
- Have clipboards or hardback exercise books ready (to collect data outside).

TEACHER NOTES:
Lesson 5 of 8. This kicks off the data block. The investigation is car-colour counting at the school boundary - a real categorical-data activity. Avoid the trap of "favourite animal" style data - we want students to count, not to ask boring questions.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- Whiteboards on desks for the start of the lesson.
- Clipboards or hardback books for outside work.
- One Tally Sheet per student.
- Pencils.

DO:
- Confirm safe boundary location and supervision before going outside.

TEACHER NOTES:
Choose a safe, visible boundary point with good driver visibility - footpath, fence line or nature strip with no road crossing. Plan the 10-minute window carefully. Carry the Tally Sheets and a backup.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review.
- Convert each measurement.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Converting measurement units - last week's content. Quick spaced review.

WATCH FOR:
- Students who use the place-value chart strategy - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 2.5 metres = 250 cm.
- 1500 grams = 1.5 kg.
- 3.2 litres = 3200 mL.

DO:
- Click to reveal.

TEACHER NOTES:
Watch for students who multiply or divide by the wrong factor.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Multiplying by 10 and 100.
- Whisper-answer then write.

DO:
- Display the three prompts.
- 25 seconds.

TEACHER NOTES:
Day five of mult/div fluency. Today is place value strategy - shift digits, not multiply.

WATCH FOR:
- Students who shift the digit and explain it - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 47 x 10 = 470.
- 3.5 x 100 = 350.
- 28 x 1000 = 28000.

DO:
- Click to reveal.

TEACHER NOTES:
Anchor the "digits shift left" rule for x10, x100, x1000.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- A good investigation needs a question that has a USEFUL answer.
- What do you think makes a great data question? Tell your partner one example and one bad example.

DO:
- 30 seconds of partner talk.
- Take a couple of answers.

TEACHER NOTES:
Use this to surface the contrast. Bad: "What is your favourite animal?" (boring, no use, results are mixed). Good: "What colour are most of the cars passing the school?" (useful, observable, can be counted).

WATCH FOR:
- Students who say "we need to be able to count it" - secure.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to plan and run a statistical investigation using categorical data.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 - identify a good investigation question. SC2 - the core target - tally categorical data accurately. SC3 - explain why our sample is fair or not.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Three key words today.
- Categorical data - data that can be sorted into groups (like colours).
- Sample - the part of the data we actually collect (not every car ever).
- Tally - a quick mark to count one at a time.

DO:
- Point at each visual.
- Have students whisper each word back.

TEACHER NOTES:
Make the distinction clear: categorical (colour, brand) versus numerical (number of wheels, speed). Today we focus on a categorical variable.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_IDO_QUESTION = `SAY:
- Watch me design the question.
- Question: "What colour are most of the cars passing our school during morning drop-off?"
- It is countable. It is useful (maybe car safety, maybe road interest). It does not ask people - we observe.
- Note - we are not asking "favourite". We are counting what we SEE.

DO:
- Write the question on the board.
- Highlight the action words "passing our school" and "during morning drop-off".

TEACHER NOTES:
Anchor the difference between a survey question and an observation question. Today is OBSERVATION.

MISCONCEPTIONS:
- Misconception: Any data collection question is a good question.
  Why: Students think popularity surveys are statistical investigations.
  Impact: They generate "favourite ice cream" data that has no use.
  Quick correction: A good question has a USEFUL answer and can be observed or measured, not guessed.

WATCH FOR:
- Students who can rewrite a bad question into a better one - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_TALLY = `SAY:
- Watch me set up the tally chart.
- One row for each colour we expect: white, black, grey, silver, red, blue, other.
- I keep "other" small. If too many go to "other", we add a category.
- One car = one tally mark. Every fifth tally crosses the previous four. Easy to count.

DO:
- Demonstrate four tallies then the fifth as a crossbar.
- Sketch a partial chart on the board.

TEACHER NOTES:
The fifth-bar tally is a classic counting trick. Make sure students see it before they go out.

WATCH FOR:
- Students who can demonstrate the crossbar - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_SAMPLE = `SAY:
- We will watch for 10 minutes at the boundary.
- That is our SAMPLE. We did not see every car ever - just the ones in our window.
- A bigger window means better data. But 10 minutes is a fair start.
- We will compare classes if we do it again later.

DO:
- Show the planned location on a map of the school.
- Confirm safe positioning.

TEACHER NOTES:
This introduces the idea that data depends on the sample. Year 6 students should notice that 10 minutes might not be representative of all day.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- Which question is the BEST statistical investigation question for Year 5/6?
- A: What is your favourite colour?
- B: What colour are most cars passing the school in 10 minutes?
- C: How fast are the cars going?
- Write A, B or C.

DO:
- Display the prompt.
- 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me on three, two, one, show."
- Scan for: B.
PROCEED: If 80% have B, click to reveal and move to We Do.
PIVOT: Most likely misconception - students think a "favourite" survey is a real investigation.
- Reteach: B can be COUNTED. A depends on opinion. C is hard to measure without equipment.
- Re-check: "Which is countable in 10 minutes at the gate?"

TEACHER NOTES:
B is the correct answer. A is opinion only. C is hard to measure.

WATCH FOR:
- Students who write B - secure.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO = `SAY:
- With me. Let's set up our tally chart together.
- Categories - white, black, grey, silver, red, blue, other.
- One row per colour. Tally column. Total column at the end.
- We will fill the totals after our 10 minutes outside.

DO:
- Demonstrate the chart layout on the board.
- Ask the class which colours to include.

TEACHER NOTES:
Use the class to co-construct the chart. The Tally Sheet matches this so the transfer outside is clean.

WATCH FOR:
- Students who pick reasonable colour categories - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Our tally chart is ready.
- Now we head outside for 10 minutes.

DO:
- Click to reveal the finished tally template.

TEACHER NOTES:
This is the transition into the practical part of the lesson.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Outside time.
- Stand at the boundary point. Face the road safely from behind the fence.
- Watch for 10 minutes. Tally each car by colour.
- When the time is up, count your totals and bring them back inside.

DO:
- Walk to the boundary as a class.
- Set the timer.
- Walk between students to check tally accuracy.
- Return to class.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Partner up. One student watches and calls colours; the other tallies.
- Extra Notes: Sit these pairs at the front of the observation group.
EXTENDING PROMPT:
- Task: Record TIME OF ARRIVAL for the first 10 cars (numerical data) as well as colour. Note any pattern.
- Extra Notes: Bring this back to discuss in Lesson 6.

TEACHER NOTES:
Safety first. Position the class behind the fence or boundary. Confirm no driveway crosses the observation zone.

WATCH FOR:
- Students who use the fifth-bar tally cleanly - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Q1: Why was 10 minutes our sample, not every car all day?
- Q2: Name one categorical variable from today's data.

DO:
- Display the prompt.
- 90 seconds.

TEACHER NOTES:
Exit ticket assesses SC2 and SC3. Q1 - we cannot observe every car; this is our sample. Q2 - colour.

WATCH FOR:
- Students who use "sample" - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria for today.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: which colour was most common?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
Hold onto the tally sheets. Tomorrow we use them to build the column graph.

[General: Closing | VTLM 2.0: Reflection]`;

// =====================================================================
// Helpers
// =====================================================================

// Draw a tally chart with given categories and pre-filled marks (for I Do/We Do illustrations).
function drawTallyChart(slide, x, y, w, rows, opts) {
  const o = opts || {};
  const headerColor = o.headerColor || C.PRIMARY;
  const rowH = o.rowH || 0.36;
  const colWidths = [w * 0.30, w * 0.55, w * 0.15];

  // Header
  let cx = x;
  ["Colour", "Tally", "Total"].forEach((h, i) => {
    slide.addShape("rect", {
      x: cx, y, w: colWidths[i], h: rowH,
      fill: { color: headerColor }, line: { color: C.WHITE, width: 1 },
    });
    slide.addText(h, {
      x: cx, y, w: colWidths[i], h: rowH,
      fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    cx += colWidths[i];
  });

  // Rows
  rows.forEach((row, i) => {
    const ry = y + (i + 1) * rowH;
    cx = x;
    [row.colour || "", row.tally || "", row.total != null ? String(row.total) : ""].forEach((cell, ci) => {
      slide.addShape("rect", {
        x: cx, y: ry, w: colWidths[ci], h: rowH,
        fill: { color: i % 2 ? C.BG_LIGHT : C.WHITE }, line: { color: C.MUTED, width: 0.5 },
      });
      slide.addText(cell, {
        x: cx + 0.10, y: ry, w: colWidths[ci] - 0.20, h: rowH,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL,
        align: ci === 0 ? "left" : "center", valign: "middle", margin: 0,
      });
      cx += colWidths[ci];
    });
  });
}

// =====================================================================
// Build
// =====================================================================

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 5: Plan the car-colour investigation",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Converting units",
      ["2.5 m = ___ cm", "1500 g = ___ kg", "3.2 L = ___ mL"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "250 cm        1.5 kg        3200 mL", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal
  withReveal(
    () => fluencySlide(pres, "Fluency: Multiplying by 10, 100, 1000",
      ["47 x 10", "3.5 x 100", "28 x 1000"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "470       350       28000", {
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
    addTitle(s, "Good question vs boring question");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    // Two columns: Bad question / Good question
    const colY = CONTENT_TOP + 0.45;
    const colH = 2.55;
    addCard(s, 0.85, colY, 4.20, colH, { strip: C.ALERT, fill: C.WHITE });
    addCard(s, 5.15, colY, 4.00, colH, { strip: C.SUCCESS, fill: C.WHITE });

    // Bad column
    addTextOnShape(s, "Boring", {
      x: 1.05, y: colY + 0.20, w: 3.80, h: 0.55, rectRadius: 0.10,
      fill: { color: C.ALERT },
    }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    s.addText([
      { text: "\"What is your favourite animal?\"", options: { fontSize: 14, color: C.CHARCOAL, italic: true, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "- everyone gives a different answer", options: { fontSize: 12, color: C.CHARCOAL, breakLine: true } },
      { text: "- no useful conclusion", options: { fontSize: 12, color: C.CHARCOAL, breakLine: true } },
      { text: "- nothing to count or compare", options: { fontSize: 12, color: C.CHARCOAL } },
    ], {
      x: 1.10, y: colY + 0.90, w: 4.00, h: 1.50,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    // Good column
    addTextOnShape(s, "Useful", {
      x: 5.35, y: colY + 0.20, w: 3.60, h: 0.55, rectRadius: 0.10,
      fill: { color: C.SUCCESS },
    }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    s.addText([
      { text: "\"What colour are most cars passing the school?\"", options: { fontSize: 14, color: C.CHARCOAL, italic: true, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "- countable", options: { fontSize: 12, color: C.CHARCOAL, breakLine: true } },
      { text: "- categorical (groups)", options: { fontSize: 12, color: C.CHARCOAL, breakLine: true } },
      { text: "- has a useful answer", options: { fontSize: 12, color: C.CHARCOAL } },
    ], {
      x: 5.40, y: colY + 0.90, w: 3.80, h: 1.50,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to plan and run a statistical investigation using categorical data.",
    [
      "I can identify a question that gives us useful, countable data.",
      "I can use a tally chart to record categorical data accurately during the 10-minute observation.",
      "I can explain why our sample of 10 minutes is or is not representative of all day.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key vocabulary
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Categorical, sample, tally");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const items = [
      { word: "Categorical", simple: "in groups",  detail: "Data that sorts into named groups (e.g. colour).", color: C.PRIMARY },
      { word: "Sample",      simple: "a portion",  detail: "The part of all possible data we actually look at.", color: C.ACCENT },
      { word: "Tally",       simple: "quick count", detail: "One mark per item. Crossbar at the fifth.",         color: C.ALERT },
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

  // Slide 10: I Do - The question
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "The question we are investigating", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Big quote
    addTextOnShape(s, "What colour are most of the cars passing our school during morning drop-off?", {
      x: 1.0, y: CONTENT_TOP + 0.50, w: 8.0, h: 1.30, rectRadius: 0.10,
      fill: { color: C.BG_LIGHT },
      line: { color: stageColor, width: 1.5 },
    }, { fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center" });

    // Why this question works
    s.addText([
      { text: "Why this works:", options: { fontSize: 14, color: stageColor, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "- We OBSERVE - we are not asking people.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
      { text: "- Colours sort into groups (categorical).", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
      { text: "- A 10-minute window gives us enough cars to count.", options: { fontSize: 13, color: C.CHARCOAL } },
    ], {
      x: 1.0, y: CONTENT_TOP + 2.05, w: 8.0, h: 1.20,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_QUESTION);
  })();

  // Slide 11: I Do - Tally chart structure
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Set up the tally chart", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    drawTallyChart(s, 1.0, CONTENT_TOP + 0.30, 8.0, [
      { colour: "White",  tally: "|||| | |", total: 7 },
      { colour: "Black",  tally: "|||| |",   total: 6 },
      { colour: "Grey",   tally: "|||| ||",  total: 7 },
      { colour: "Silver", tally: "|||",      total: 3 },
      { colour: "Red",    tally: "||",       total: 2 },
      { colour: "Blue",   tally: "|||",      total: 3 },
      { colour: "Other",  tally: "|",        total: 1 },
    ], { headerColor: stageColor, rowH: 0.36 });

    s.addText("Each tally = one car. Every fifth mark crosses the previous four.", {
      x: 0.7, y: SAFE_BOTTOM - 0.40, w: 8.6, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_TALLY);
  })();

  // Slide 12: I Do - Sample
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Our sample: 10 minutes at the boundary", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Clock-like illustration
    const ccx = 2.8;
    const ccy = CONTENT_TOP + 1.60;
    const ccr = 1.10;
    s.addShape("ellipse", {
      x: ccx - ccr, y: ccy - ccr, w: 2 * ccr, h: 2 * ccr,
      fill: { color: C.BG_LIGHT },
      line: { color: stageColor, width: 2 },
    });
    s.addText("10 min", {
      x: ccx - 0.8, y: ccy - 0.30, w: 1.6, h: 0.50,
      fontSize: 22, fontFace: FONT_H, color: stageColor, bold: true,
      align: "center", margin: 0,
    });

    // Right: explanation
    addCard(s, 5.20, CONTENT_TOP + 0.25, 4.30, 2.90, { strip: C.ACCENT, fill: C.WHITE });
    s.addText([
      { text: "Why a sample?", options: { fontSize: 16, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "We cannot watch every car all day.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
      { text: "We take a 10-minute window during morning drop-off.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Stretch question:", options: { fontSize: 14, color: C.ALERT, bold: true, breakLine: true } },
      { text: "Would our results change if we counted at 3 pm pick-up?", options: { fontSize: 13, color: C.CHARCOAL, italic: true } },
    ], {
      x: 5.40, y: CONTENT_TOP + 0.45, w: 4.00, h: 2.60,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_SAMPLE);
  })();

  // Slides 13-14: CFU + reveal — which is the best question?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Which is the best investigation question?", { color: C.ALERT });
      addTextOnShape(s, "CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      s.addText([
        { text: "A.   What is your favourite colour?", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "B.   What colour are most cars passing the school in 10 minutes?", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "C.   How fast are the cars going past the school?", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 14, breakLine: true } },
        { text: "Write A, B, or C on your board.", options: { fontSize: 16, color: C.ALERT, bold: true } },
      ], {
        x: 1.0, y: CONTENT_TOP + 0.40, w: 8.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.80,
        fontFace: FONT_B, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "B. We can observe and count it. A is opinion. C is hard to measure.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 15-16: We Do — co-construct tally chart
  withReveal(
    () => {
      const s = pres.addSlide();
      const stageColor = STAGE_COLORS["3"];
      addTopBar(s, stageColor);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Together: design our tally chart", { color: stageColor });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

      drawTallyChart(s, 1.0, CONTENT_TOP + 0.40, 8.0, [
        { colour: "White",  tally: "", total: "" },
        { colour: "Black",  tally: "", total: "" },
        { colour: "Grey",   tally: "", total: "" },
        { colour: "Silver", tally: "", total: "" },
        { colour: "Red",    tally: "", total: "" },
        { colour: "Blue",   tally: "", total: "" },
        { colour: "Other",  tally: "", total: "" },
      ], { headerColor: stageColor, rowH: 0.32 });

      s.addText("Together: any colours to add? Anything to merge into Other?", {
        x: 0.7, y: SAFE_BOTTOM - 0.40, w: 8.6, h: 0.28,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Chart is ready. Take the Tally Sheet and let's go outside.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // Slide 17: You Do - outside
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["4"];
    addTopBar(s, stageColor);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Outside: 10-minute count", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    s.addText([
      { text: "First: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "stand safely at the boundary point.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 10, breakLine: true } },
      { text: "Next: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "tally each car by colour. One mark per car.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 10, breakLine: true } },
      { text: "Then: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "count your totals. Bring the sheet back.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 10, breakLine: true } },
      { text: "Safety: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "stay behind the fence. Listen to the teacher.", options: { fontSize: 14, color: C.CHARCOAL, italic: true } },
    ], {
      x: 1.0, y: CONTENT_TOP + 0.30, w: 8.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.60,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 18: Exit Ticket
  exitTicketSlide(pres,
    [
      "Why was 10 minutes our sample, not every car all day?",
      "Name one categorical variable from today's data.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 19: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: which colour was most common in your data?",
      scItems: [
        "I can identify a question that gives us useful, countable data.",
        "I can use a tally chart to record categorical data accurately during the 10-minute observation.",
        "I can explain why our sample of 10 minutes is or is not representative of all day.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "GD_Lesson5_Data_Collection_Plan.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Tally chart for the 10-minute car-colour count at the boundary.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "One car = one tally mark. Every fifth mark crosses the previous four (||||  with a crossbar). " +
      "If you see a colour that does not match a row, put it in OTHER.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Tally Sheet", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Question: What colour are most cars passing our school during the 10-minute window?", y);
    y = addBodyText(doc, "Date: ____________     Time started: ____________     Time finished: ____________", y);
    y = addBodyText(doc, "Observer: ____________________________   Location at the boundary: ____________________________", y);

    y = addBodyText(doc, " ", y);
    y = addBodyText(doc, "Colour       |  Tally                                |  Total", y, { fontSize: 11 });
    y = addBodyText(doc, "White        |                                       |", y, { fontSize: 11 });
    y = addBodyText(doc, "Black        |                                       |", y, { fontSize: 11 });
    y = addBodyText(doc, "Grey         |                                       |", y, { fontSize: 11 });
    y = addBodyText(doc, "Silver       |                                       |", y, { fontSize: 11 });
    y = addBodyText(doc, "Red          |                                       |", y, { fontSize: 11 });
    y = addBodyText(doc, "Blue         |                                       |", y, { fontSize: 11 });
    y = addBodyText(doc, "Other        |                                       |", y, { fontSize: 11 });

    y = addSectionHeading(doc, "After the 10 minutes", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Total cars observed: ___", y);
    y = addWriteLine(doc, "Which colour was most common?  ____________", y);
    y = addWriteLine(doc, "Which colour was least common? ____________", y);

    y = addSectionHeading(doc, "Stretch — reflection", y, { color: C.ACCENT });
    y = addBodyText(doc, "Do you think your 10-minute sample would look the same at 3 pm pick-up? Why or why not?", y);
    y = addWriteLine(doc, "_____________________________________________________________", y);
    y = addWriteLine(doc, "_____________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Tally Sheet | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Safety, logistics and assessment notes for the boundary observation.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Before going outside", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1. Confirm a safe boundary location with good driver visibility. Avoid driveways and gates that swing across the observation area.", y);
    y = addBodyText(doc, "2. Position the class behind the fence or on the school side of the footpath.", y);
    y = addBodyText(doc, "3. Set the timer for 10 minutes once everyone is in place.", y);

    y = addSectionHeading(doc, "What to watch for during the count", y, { color: C.PRIMARY });
    y = addBodyText(doc, "- Students who switch which side they are watching mid-count (drift).", y);
    y = addBodyText(doc, "- Students who guess if they miss a car (skip).", y);
    y = addBodyText(doc, "- Students whose tally is hard to read (mark per car, fifth-bar crossbar).", y);

    y = addSectionHeading(doc, "Discussion prompts back inside", y, { color: C.PRIMARY });
    y = addBodyText(doc, "- Did your total match your partner's? Why or why not?", y);
    y = addBodyText(doc, "- Which colour was the mode (most common)?", y);
    y = addBodyText(doc, "- If we counted at a different time of day, would the result change?", y);

    y = addSectionHeading(doc, "Assessment look-fors (SC2 and SC3)", y, { color: C.ACCENT });
    y = addBodyText(doc, "SC2 - student's tally is one mark per car, totals add to a reasonable number for 10 minutes.", y);
    y = addBodyText(doc, "SC3 - student can explain that a 10-minute window is a sample, and that a different time of day might give different results.", y);

    y = addTipBox(doc,
      "Hold all Tally Sheets to use in Lesson 6. Take a quick photo of one or two charts in case any sheet goes missing overnight.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Teacher Notes | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 5 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
