"use strict";

// Measurement Unit (Year 5/6 Numeracy) — Lesson 7: Reading timetables and planning a trip.
// Year 6 content. Bus timetable used as the source. Brings all prior skills together.
// Daily Review: adding fractions.
// Fluency: addition.

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

const SESSION = 7;
const TOTAL = 7;
const UNIT_TITLE = "Measurement: Units, Time and Timetables";
const FOOTER = `Measurement | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/Meas_Lesson7_Timetables_and_Trips";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 7 Read the Timetable",
  "Read a bus timetable. Answer questions and plan a short trip.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 7 Answer Key",
  "Worked answers for the Read the Timetable sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// Sample bus timetable used across the lesson. Times are in 24 hour format.
// Columns: Riverside, City Centre, Hospital, Westmall, Beachpark
const TIMETABLE_STOPS = ["Riverside", "City Centre", "Hospital", "Westmall", "Beachpark"];
const TIMETABLE_SERVICES = [
  ["07:15", "07:35", "07:50", "08:05", "08:25"],
  ["08:00", "08:20", "08:35", "08:50", "09:10"],
  ["08:45", "09:05", "09:20", "09:35", "09:55"],
  ["09:30", "09:50", "10:05", "10:20", "10:40"],
];

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome to the last lesson in our measurement unit.
- Today we bring everything together. We will read a timetable and plan a short trip.

DO:
- Settle students before clicking past the title.
- Have whiteboards ready.

TEACHER NOTES:
Lesson 7 of 7. Year 6 content. The lesson uses 24 hour time, elapsed time, and choosing units indirectly. All prior skills apply.

WATCH FOR:
- Students who treat the timetable like a calendar - it is a grid of stops and times.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- One Read the Timetable sheet per student.
- Whiteboards on desks.

DO:
- Print one sheet per student.
- Print one answer key.

TEACHER NOTES:
One student resource plus answer key. The timetable is also on the slide so students can refer to it.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Adding fractions.
- Solve each on your whiteboard.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Adding fractions with like denominators. Last review of the week.

WATCH FOR:
- Students who add numerators only - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 2/5 + 1/5 = 3/5.
- 4/7 + 2/7 = 6/7.
- 3/10 + 5/10 = 8/10 = 4/5.
- Tick or fix.

DO:
- Click to reveal.

TEACHER NOTES:
The third one simplifies. Note students who notice and simplify.

WATCH FOR:
- Students who simplify - extension.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Addition.
- Whisper-answer then write.

DO:
- Display the three prompts.
- 30 seconds.

TEACHER NOTES:
Last addition fluency of the week.

WATCH FOR:
- Students who recall - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 7 + 5 = 12.
- 8 + 6 = 14.
- 9 + 7 = 16.

DO:
- Click to reveal.

TEACHER NOTES:
Brisk last fluency.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- Three key words today.
- Timetable - a grid that shows when something departs and arrives.
- Itinerary - a planned list of stops and times for a trip.
- Service - one trip in the timetable (one column of times).

DO:
- Point at the example timetable on the next slide.
- Run your finger along one column to show one service.

TEACHER NOTES:
Anchor the vocabulary on the actual timetable. Students who can name a "service" and a "stop" will navigate the table much faster.

WATCH FOR:
- Students who can point at a service - tracking.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to read a timetable and plan a short trip.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 is reading a single time off the timetable. SC2 is the core target. SC3 stretches to planning the trip with the latest possible departure.

WATCH FOR:
- Students who can repeat the SC - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. This is a bus timetable.
- The first column lists the STOPS.
- Each other column is a SERVICE - one bus trip.
- To find the time the second service leaves Riverside, I find the Riverside row and the second column. 08:00.
- To find when that same service arrives at Beachpark, I stay in the same column and go down to Beachpark. 09:10.

DO:
- Use your finger or laser pointer to trace down the second column.
- Highlight Riverside 08:00 and Beachpark 09:10.

TEACHER NOTES:
Anchor the navigation move. Many students try to jump around the table. The disciplined move is: find the stop (row), then move down the same column (service).

MISCONCEPTIONS:
- Misconception: Students read across rows instead of down columns.
  Why: They confuse the structure with a calendar.
  Impact: They mix up services.
  Quick correction: "One service = one column. Stay in the column to follow that bus."

WATCH FOR:
- Students who trace down a column - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Watch me work out the elapsed time for the second service.
- Riverside 08:00 to Beachpark 09:10.
- Jump 1: 08:00 -> 09:00 = 60 min.
- Jump 2: 09:00 -> 09:10 = 10 min.
- Total: 70 min = 1 hour 10 minutes.
- That is the full trip time for this bus.

DO:
- Write the jumps on the board.
- Add them out loud.

TEACHER NOTES:
This brings yesterday's elapsed time strategy into the timetable. Reading the time off the timetable is step one; the elapsed time calculation is step two. Keep both steps explicit.

WATCH FOR:
- Students who use the number line jumps - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- The third service leaves Riverside at 08:45.
- What time does that service arrive at Hospital? Use the timetable.

DO:
- Display the prompt.
- 45 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your boards on three, two, one, show."
- Scan for: 09:20.
PROCEED: If 80% have 09:20, click to reveal and move to We Do.
PIVOT: Most likely misconception - students read the wrong column.
- Reteach: "Riverside 08:45 is the third column. Stay in that column down to the Hospital row."
- Re-check: "Same service, what time does it arrive at Westmall?"

TEACHER NOTES:
Probe the column-discipline. 08:45 is the third service.

WATCH FOR:
- Students who write 09:20 - secure.
- Students who write 09:35 (Westmall) or 09:05 (City Centre) - column drift.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Four questions about the same timetable.
- Partner talk for each. We will check together.

DO:
- Display the four questions.
- 4 minutes total.
- Cold call after each.

TEACHER NOTES:
Mix of reading and elapsed time. Some require column discipline; others require jumps.

WATCH FOR:
- Pairs who use a finger to trace columns - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 1: First service Westmall arrival = 08:05.
- 2: How long the first service takes Riverside -> Beachpark = 1 hour 10 minutes (07:15 to 08:25).
- 3: The latest service that arrives at Hospital before 10:00 is the third (09:20).
- 4: If a student catches the second service at City Centre (08:20) and gets off at Westmall (08:50), trip time = 30 minutes.

DO:
- Click to reveal.
- Run the trace for each.

TEACHER NOTES:
Reveal models the trace step explicitly each time.

WATCH FOR:
- Students who got three of four - secure.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. You need to be at Hospital by 09:30 for an appointment.
- Theo says you should catch the bus that LEAVES Riverside at 08:45.
- Mira says you should catch the bus that ARRIVES at Hospital at 09:35.
- Thumbs UP for Theo. Thumbs DOWN for Mira.

DO:
- Display the disagreement.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP for Theo. DOWN for Mira."
- Scan for: thumbs UP. Theo is right.
PROCEED: If 80% agree with Theo, click to reveal and confirm.
PIVOT: Most likely misconception - students count any service that touches Hospital, not the one that arrives in time.
- Reteach: "I need to ARRIVE by 09:30. The third service arrives 09:20. The fourth service arrives 09:35 - too late."
- Re-check: "Which service should you catch to be at Westmall by 08:50?"

TEACHER NOTES:
The hinge tests "latest service that arrives in time". The arrival must be at or before the deadline.

WATCH FOR:
- Confident thumbs UP for Theo - they read arrival times.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the Read the Timetable sheet.
- The same bus timetable is printed on it.
- Answer Parts A and B. The extension asks you to plan a short trip.

DO:
- Distribute the sheet.
- Circulate. Stop and ask "which service? Show me the column."

TEACHER NOTES:
Different questions from the We Do. Same timetable.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use Part A only. Sit with these students. Trace the column with a ruler.
- Extra Notes: Encourage students to highlight the row before reading.
EXTENDING PROMPT:
- Task: Part C - plan a complete trip with arrival deadline. Choose the latest service and justify.
- Extra Notes: Use elapsed time language.

WATCH FOR:
- Students who trace columns - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task of the unit. On your whiteboard.
- Look at the timetable.
- Q1: What time does the second service arrive at Hospital?
- Q2: How long does the first service take from City Centre to Beachpark?

DO:
- Display the prompt.
- 4 minutes.

TEACHER NOTES:
Exit ticket assesses SC2. Look for 08:35 and 50 minutes (07:35 to 08:25).

WATCH FOR:
- Students who get both - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria for today.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: ONE thing you learned this unit that you did NOT know before.

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea: a column is a service; stay in the column to follow the bus. This is also the close of the whole unit. Take a moment to celebrate the week.

WATCH FOR:
- Strong thumbs up - secure for end-of-unit assessment.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw the bus timetable grid. Rows are stops; columns are services.
function drawTimetable(slide, x, y, w, h, stops, services, opts) {
  const o = opts || {};
  const headerColor = o.headerColor || C.PRIMARY;
  const stopColor = o.stopColor || C.ACCENT;
  const altRow = o.altRow !== false;
  const numColumns = 1 + services.length; // stop column + service columns
  const numRows = 1 + stops.length;       // header row + stop rows
  // Make the stop column wider than the service columns.
  const stopColW = Math.min(1.8, w * 0.30);
  const serviceColW = (w - stopColW) / services.length;
  const headerH = Math.min(0.50, h * 0.16);
  const rowH = (h - headerH) / stops.length;

  // Header row: "Stop" + "Service 1" .. "Service N"
  slide.addShape("rect", {
    x, y, w: stopColW, h: headerH,
    fill: { color: headerColor }, line: { color: C.WHITE, width: 1 },
  });
  slide.addText("Stop", {
    x, y, w: stopColW, h: headerH,
    fontSize: 12, fontFace: FONT_H, color: C.WHITE,
    bold: true, align: "center", valign: "middle", margin: 0,
  });
  services.forEach((_, j) => {
    const cx = x + stopColW + j * serviceColW;
    slide.addShape("rect", {
      x: cx, y, w: serviceColW, h: headerH,
      fill: { color: headerColor }, line: { color: C.WHITE, width: 1 },
    });
    slide.addText(`Service ${j + 1}`, {
      x: cx, y, w: serviceColW, h: headerH,
      fontSize: 11, fontFace: FONT_H, color: C.WHITE,
      bold: true, align: "center", valign: "middle", margin: 0,
    });
  });

  // Stop rows
  stops.forEach((stop, i) => {
    const ry = y + headerH + i * rowH;
    const rowFill = altRow && i % 2 === 1 ? (C.BG_LIGHT || "F6F6F6") : C.WHITE;

    // Stop column (highlighted)
    slide.addShape("rect", {
      x, y: ry, w: stopColW, h: rowH,
      fill: { color: stopColor }, line: { color: C.WHITE, width: 1 },
    });
    slide.addText(stop, {
      x: x + 0.05, y: ry, w: stopColW - 0.10, h: rowH,
      fontSize: 11, fontFace: FONT_H, color: C.WHITE,
      bold: true, align: "left", valign: "middle", margin: 0,
    });

    // Service cells
    services.forEach((svc, j) => {
      const cx = x + stopColW + j * serviceColW;
      slide.addShape("rect", {
        x: cx, y: ry, w: serviceColW, h: rowH,
        fill: { color: rowFill }, line: { color: C.MUTED, width: 0.5 },
      });
      slide.addText(svc[i], {
        x: cx, y: ry, w: serviceColW, h: rowH,
        fontSize: 12, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
    });
  });
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 7: Reading timetables and planning a trip",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal — adding fractions
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Adding fractions",
      ["2/5 + 1/5 =", "4/7 + 2/7 =", "3/10 + 5/10 ="],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3/5         6/7         8/10 = 4/5", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal — addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Addition",
      ["7 + 5", "8 + 6", "9 + 7"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "12     14     16", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Key vocabulary
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Timetable, itinerary, service");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const cellW = 2.7;
    const gap = 0.20;
    const totalW = cellW * 3 + gap * 2;
    const startX = (10 - totalW) / 2;
    const cardY = CONTENT_TOP + 0.30;
    const cardH = 2.50;

    const items = [
      { label: "Timetable", desc: "a grid of times",        detail: "rows = stops, cols = services", color: C.PRIMARY },
      { label: "Itinerary", desc: "a planned trip list",    detail: "your chosen stops and times",   color: C.ACCENT },
      { label: "Service",   desc: "one trip",               detail: "one column of times",           color: C.ALERT },
    ];

    items.forEach((it, i) => {
      const cx = startX + i * (cellW + gap);
      addCard(s, cx, cardY, cellW, cardH, { strip: it.color, fill: C.WHITE });
      addTextOnShape(s, it.label, {
        x: cx + 0.20, y: cardY + 0.20, w: cellW - 0.40, h: 0.60, rectRadius: 0.10,
        fill: { color: it.color },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(it.desc, {
        x: cx, y: cardY + 0.95, w: cellW, h: 0.40,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", margin: 0,
      });
      s.addText(it.detail, {
        x: cx, y: cardY + 1.45, w: cellW, h: 0.60,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to read a timetable and plan a short trip.",
    [
      "I can find a single time on a timetable by tracing the correct row and column.",
      "I can use a timetable to find when a bus arrives or departs and work out the trip time.",
      "I can plan a trip with an arrival deadline by choosing the latest service that arrives in time.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: I Do (1) — show the timetable and trace one service
  (() => {
    const s = pres.addSlide();
    const stageColor = STAGE_COLORS["2"];
    addTopBar(s, stageColor);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Reading the bus timetable", { color: stageColor });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: stageColor });

    // Use the wide single card layout for the timetable
    drawTimetable(s,
      0.85, CONTENT_TOP + 0.25, 8.3, 2.85,
      TIMETABLE_STOPS, TIMETABLE_SERVICES);

    s.addText("One column = one service. Stay in the column to follow that bus.", {
      x: 0.7, y: SAFE_BOTTOM - 0.45, w: 8.6, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO1);
  })();

  // Slide 10: I Do (2) — elapsed time for service 2
  workedExSlide(pres, 2, "I Do", "How long does Service 2 take?",
    [
      "Service 2:",
      "Riverside  08:00.",
      "Beachpark  09:10.",
      "",
      "Jump 1: 08:00 -> 09:00 = 60 min.",
      "Jump 2: 09:00 -> 09:10 = 10 min.",
      "",
      "Total: 70 min = 1 hr 10 min.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("Service 2 trip", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      addTextOnShape(slide, "Riverside  08:00", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.40, h: 0.50, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("▼", {
        x: lg.rightX, y: lg.panelTopPadded + 1.10, w: lg.rightW, h: 0.40,
        fontSize: 24, fontFace: FONT_H, color: C.SECONDARY,
        bold: true, align: "center", margin: 0,
      });

      addTextOnShape(slide, "Beachpark  09:10", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 1.55, w: lg.rightW - 0.40, h: 0.50, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(slide, "Trip = 1 hr 10 min", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.40, w: lg.rightW - 0.40, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 11: Bring back the timetable as a CFU prompt
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU", { color: C.ALERT });
    addTitle(s, "Service 3 arrives at Hospital at...?", { color: C.ALERT });

    addTextOnShape(s, "✓ CHECK", {
      x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
      fill: { color: C.WHITE },
      line: { color: C.ALERT, width: 1.5 },
    }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

    drawTimetable(s,
      0.85, CONTENT_TOP + 0.25, 8.3, 2.85,
      TIMETABLE_STOPS, TIMETABLE_SERVICES);

    s.addText("On your board: trace the third column down to Hospital. Write the time.", {
      x: 0.7, y: SAFE_BOTTOM - 0.45, w: 8.6, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU_Q);
  })();

  // Slide 12: CFU reveal
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU Reveal", { color: C.ALERT });
    addTitle(s, "Service 3 arrives at Hospital at...?", { color: C.ALERT });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

    drawTimetable(s,
      0.85, CONTENT_TOP + 0.25, 8.3, 2.50,
      TIMETABLE_STOPS, TIMETABLE_SERVICES);

    addTextOnShape(s, "Service 3 arrives at Hospital at 09:20.", {
      x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
      fill: { color: C.SUCCESS },
    }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
  })();

  // Slide 13: We Do — four questions about the timetable
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["3"]);
    addStageBadge(s, 3, "We Do");
    addTitle(s, "Answer with your partner", { color: STAGE_COLORS["3"] });

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

    s.addText([
      { text: "Use the timetable.", options: { fontSize: 16, color: C.CHARCOAL, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "1.  When does Service 1 arrive at Westmall?",       options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "2.  How long does Service 1 take Riverside -> Beachpark?", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "3.  Latest service that arrives at Hospital BEFORE 10:00?", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "4.  Catch Service 2 at City Centre. Get off at Westmall. Trip time?", options: { fontSize: 14, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    // Right column: timetable
    drawTimetable(s,
      5.20, CONTENT_TOP + 0.10, 4.30, 3.10,
      TIMETABLE_STOPS, TIMETABLE_SERVICES);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO_Q);
  })();

  // Slide 14: We Do reveal
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["3"]);
    addStageBadge(s, 3, "We Do Reveal");
    addTitle(s, "Check your answers", { color: STAGE_COLORS["3"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

    const answers = [
      "1.  Service 1 arrives at Westmall at 08:05.",
      "2.  Service 1: Riverside 07:15 -> Beachpark 08:25 = 1 hr 10 min.",
      "3.  Service 3 arrives at Hospital 09:20 - the latest before 10:00.",
      "4.  Service 2: City Centre 08:20 -> Westmall 08:50 = 30 min.",
    ];
    answers.forEach((a, i) => {
      addTextOnShape(s, a, {
        x: 0.8, y: CONTENT_TOP + 0.20 + i * 0.85, w: 8.4, h: 0.65, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true, align: "left" });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO_A);
  })();

  // Slides 15-16: CFU hinge + reveal — Theo vs Mira
  withReveal(
    () => cfuSlide(pres, "CFU", "Who is right?", "Thumbs Up or Thumbs Down",
      "You must be at Hospital by 09:30 for an appointment.\n\nTheo: catch the bus that LEAVES Riverside at 08:45.\n\nMira: catch the bus that ARRIVES at Hospital at 09:35.\n\nThumbs UP for Theo.   Thumbs DOWN for Mira.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Theo is right. Service 3 arrives 09:20 - in time. Service 4 arrives 09:35 - too late.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: You Do — instructions + timetable for reference
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: Read the Timetable", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["4"] });

    s.addText([
      { text: "First: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "trace the column.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Next: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "read the time.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Then: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "calculate the elapsed time.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "Extension: ", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: false } },
      { text: "plan a trip with an arrival deadline.", options: { fontSize: 16, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
      fontFace: FONT_B, valign: "top", margin: 0,
    });

    // Right side: timetable
    drawTimetable(s,
      5.20, CONTENT_TOP + 0.10, 4.30, 3.10,
      TIMETABLE_STOPS, TIMETABLE_SERVICES);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 18: Exit Ticket
  exitTicketSlide(pres,
    [
      "Use the timetable. What time does Service 2 arrive at Hospital?",
      "Use the timetable. How long does Service 1 take from City Centre to Beachpark?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 19: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: ONE thing you learned this unit that you did NOT know before.",
      scItems: [
        "I can find a single time on a timetable by tracing the correct row and column.",
        "I can use a timetable to find when a bus arrives or departs and work out the trip time.",
        "I can plan a trip with an arrival deadline by choosing the latest service that arrives in time.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Meas_Lesson7_Timetables_and_Trips.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Use the bus timetable. Trace each column. Calculate elapsed time.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "One column = one service. Stay in the column to follow the bus. Use number-line jumps for elapsed time. Read arrival times carefully when planning a trip with a deadline.",
      y, { color: C.ACCENT });

    // Render the timetable in the PDF as a simple text grid.
    y = addSectionHeading(doc, "Bus Timetable", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Stop          | Service 1 | Service 2 | Service 3 | Service 4", y, { fontSize: 11 });
    y = addBodyText(doc, "Riverside     |   07:15   |   08:00   |   08:45   |   09:30", y, { fontSize: 11 });
    y = addBodyText(doc, "City Centre   |   07:35   |   08:20   |   09:05   |   09:50", y, { fontSize: 11 });
    y = addBodyText(doc, "Hospital      |   07:50   |   08:35   |   09:20   |   10:05", y, { fontSize: 11 });
    y = addBodyText(doc, "Westmall      |   08:05   |   08:50   |   09:35   |   10:20", y, { fontSize: 11 });
    y = addBodyText(doc, "Beachpark     |   08:25   |   09:10   |   09:55   |   10:40", y, { fontSize: 11 });

    y = addSectionHeading(doc, "Part A — Read the timetable", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  What time does Service 2 leave Riverside? ___", y);
    y = addWriteLine(doc, "b)  What time does Service 3 arrive at Westmall? ___", y);
    y = addWriteLine(doc, "c)  What time does Service 4 arrive at Beachpark? ___", y);
    y = addWriteLine(doc, "d)  Which service leaves Riverside at 07:15? Service ___", y);

    y = addSectionHeading(doc, "Part B — Elapsed time", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Use number-line jumps.", y);
    y = addWriteLine(doc, "a)  Service 2 Riverside to Beachpark = ___ hr ___ min", y);
    y = addWriteLine(doc, "b)  Service 1 City Centre to Westmall = ___ min", y);
    y = addWriteLine(doc, "c)  Service 4 Hospital to Beachpark = ___ min", y);

    y = addSectionHeading(doc, "Part C — Plan the trip (extension)", y, { color: C.ACCENT });
    y = addBodyText(doc,
      "You need to be at Beachpark by 10:00 to meet a friend. You will catch the bus at City Centre.", y);
    y = addWriteLine(doc, "a)  Which is the latest service you can catch? Service ___", y);
    y = addWriteLine(doc, "b)  What time does it leave City Centre? ___", y);
    y = addWriteLine(doc, "c)  What time does it arrive at Beachpark? ___", y);
    y = addWriteLine(doc, "d)  How many minutes spare do you have before 10:00? ___ min", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Read the Timetable | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Read the Timetable sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Part A — Read the timetable", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Service 2 leaves Riverside at 08:00.", y);
    y = addBodyText(doc, "b)  Service 3 arrives at Westmall at 09:35.", y);
    y = addBodyText(doc, "c)  Service 4 arrives at Beachpark at 10:40.", y);
    y = addBodyText(doc, "d)  Service 1 leaves Riverside at 07:15.", y);

    y = addSectionHeading(doc, "Part B — Elapsed time", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Service 2 Riverside (08:00) to Beachpark (09:10) = 1 hr 10 min.", y);
    y = addBodyText(doc, "b)  Service 1 City Centre (07:35) to Westmall (08:05) = 30 min.", y);
    y = addBodyText(doc, "c)  Service 4 Hospital (10:05) to Beachpark (10:40) = 35 min.", y);

    y = addSectionHeading(doc, "Part C — Plan the trip", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  Service 3 is the latest that arrives at Beachpark before 10:00 (arrives 09:55).", y);
    y = addBodyText(doc, "b)  Service 3 leaves City Centre at 09:05.", y);
    y = addBodyText(doc, "c)  Arrives at Beachpark at 09:55.", y);
    y = addBodyText(doc, "d)  5 minutes spare before 10:00.", y);
    y = addBodyText(doc,
      "(If a student chooses Service 4 - leaves City Centre 09:50 - it arrives Beachpark 10:40, which is AFTER 10:00. Not valid.)", y);

    y = addTipBox(doc,
      "Watch for: students who confuse departure and arrival; students who jump across rows instead of staying in the column; students who pick any service that touches the destination instead of the latest one that arrives in time.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 7 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
