"use strict";

// Measurement Unit (Year 5/6 Numeracy) — Lesson 5: 12-hour and 24-hour time.
// Year 5 content. Hand-held analog clocks used as manipulatives throughout.
// Daily Review: subtracting decimals.
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

const SESSION = 5;
const TOTAL = 7;
const UNIT_TITLE = "Measurement: Units, Time and Timetables";
const FOOTER = `Measurement | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/Meas_Lesson5_12_and_24_Hour_Time";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 5 12 and 24 Hour Time",
  "Convert between 12 and 24 hour time. Use hand-held clocks to check.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 5 Answer Key",
  "Worked answers for the 12 and 24 Hour Time sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back. We finish the metric units this week and we move into time.
- Today is about reading both 12 hour and 24 hour time.
- You will use your hand-held clocks.

DO:
- Settle students before clicking past the title.
- Hand out one hand-held analog clock per pair.

TEACHER NOTES:
Lesson 5 of 7. Year 5 content. Hand-held clocks are critical as the manipulative throughout the lesson. Set them up before students enter if possible.

WATCH FOR:
- Students who know analog time - they will pick up 24 hour quickly.
- Students who are unsure of analog - pair them with a confident partner.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- Each pair needs one hand-held analog clock.
- You also need your whiteboard and the 12 and 24 Hour Time sheet.

DO:
- Pass out hand-held clocks (one per pair).
- Print one sheet per student.
- Print one answer key.

TEACHER NOTES:
Hand-held clocks are central. Print one Time sheet per student plus answer key. The classroom analog clock above the board doubles as a class anchor.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Subtracting decimals.
- Line up the decimal points first.
- Solve each on your whiteboard.

DO:
- Display the three prompts.
- 2 minutes.

TEACHER NOTES:
Subtracting decimals reviews last unit. Watch for borrowing errors and missing zeros (eg. 2.5 - 0.45 needs students to think of 2.5 as 2.50).

WATCH FOR:
- Students who line up and add a zero where needed - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 2.5 - 0.45 = 2.05.
- 6.3 - 1.7 = 4.6.
- 4.05 - 1.20 = 2.85.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The key idea: think of 2.5 as 2.50 to line up the hundredths.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Addition facts.
- Whisper-answer then write.

DO:
- Display the three prompts.
- 30 seconds.

TEACHER NOTES:
Back to addition. Today: bridging facts.

WATCH FOR:
- Students who use bridging - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 9 + 4 = 13.
- 6 + 7 = 13.
- 8 + 5 = 13.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
All three make 13 - a nice teachable reveal that doubles and near-doubles agree.

WATCH FOR:
- Students who notice they all make 13 - tracking.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- Three key words today.
- AM is the morning. From midnight to noon.
- PM is the afternoon and evening. From noon to midnight.
- 24 hour time uses numbers from 00:00 to 23:59. No AM or PM needed.

DO:
- Point at the classroom clock as you say AM and PM.
- Write 00:00, 12:00, and 23:59 on the board.

TEACHER NOTES:
AM and PM are essential vocabulary. Many students confuse "12 PM" with "12 AM" - we will pin them down explicitly.

WATCH FOR:
- Students who can say AM = morning, PM = afternoon/evening - tracking.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to read and convert between 12 hour and 24 hour time.
- Now the success criteria.

DO:
- Choral read.
- Hold up the hand-held clock.

TEACHER NOTES:
SC1 reads a 24 hour time and identifies AM/PM. SC2 is the conversion. SC3 stretches to setting the time on the analog clock.

WATCH FOR:
- Students who can repeat the SC - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. Here is a 24 hour day on a number line.
- The day starts at 00:00 - midnight.
- 12:00 is noon (midday).
- The day ends at 23:59 - just before midnight.
- Anything BEFORE noon is AM.
- Anything AFTER noon is PM.
- 24 hour time keeps counting past 12. So 13:00 is 1 PM. 14:00 is 2 PM. 23:00 is 11 PM.

DO:
- Trace the number line with your finger.
- Point at 12:00 noon and say "AM ends, PM begins".
- Hold up the hand-held clock at 3 o'clock and ask: "AM or PM?".

TEACHER NOTES:
The number line is the anchor. The trick for converting 24 hour PM times: if the hour is bigger than 12, subtract 12.

MISCONCEPTIONS:
- Misconception: Students think 12 AM is noon (or that 12 PM is midnight).
  Why: The 12 swap is genuinely confusing.
  Impact: They convert times wrong by 12 hours.
  Quick correction: "12 PM = noon (the P is in the middle, like the sun overhead). 12 AM = midnight."

WATCH FOR:
- Students who answer "AM or PM?" correctly for the analog clock - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Watch me convert 14:30 into 12 hour time.
- 14 is bigger than 12, so it must be PM.
- I subtract 12 from the hour: 14 - 12 = 2.
- The minutes stay the same: 30.
- So 14:30 = 2:30 PM.

DO:
- Write "14:30" on the board.
- Circle the "14" and say "bigger than 12 = PM".
- Write "14 - 12 = 2".
- Write "= 2:30 PM".
- Set the hand-held clock to 2:30 to show what the analog looks like.

TEACHER NOTES:
Second worked example. This is the most common conversion students will need. The rule: if the hour is 13 or more, subtract 12 and add PM. If the hour is 00, that is 12 AM (midnight). If the hour is 12, that is 12 PM (noon).

MISCONCEPTIONS:
- Misconception: Students leave PM off ("14:30 = 2:30").
  Why: They forget the AM/PM tag.
  Impact: The time is ambiguous.
  Quick correction: "12 hour time MUST have AM or PM. Otherwise we do not know which one it is."

WATCH FOR:
- Students who include AM/PM - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- Convert 18:45 into 12 hour time.
- On your whiteboard. Then set your hand-held clock to that time.

DO:
- Display the prompt.
- 60 seconds.
- Walk and scan boards and clocks.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your boards on three, two, one, show."
- Scan for: 6:45 PM.
PROCEED: If 80% have 6:45 PM, click to reveal and move to We Do.
PIVOT: Most likely misconception - students forget to subtract 12.
- Reteach: "18 is bigger than 12. PM. Then 18 - 12 = 6."
- Re-check: "Convert 19:20."

TEACHER NOTES:
Probe the subtract-12 move. The hand-held clock check (set to 6:45) is the second layer of evidence.

WATCH FOR:
- Students who write 6:45 PM AND set the clock correctly - secure.
- Students who write 18:45 PM (no subtraction) - reteach.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me.
- Four times. Convert each to the other system.
- Two are 24 hour to 12 hour. Two are 12 hour to 24 hour.
- Partner talk for each.

DO:
- Display the four times.
- 3 minutes total.
- Cold call after each.

TEACHER NOTES:
Mixed directions. The reverse (12 hour PM to 24 hour) requires ADDING 12.

WATCH FOR:
- Pairs who explain the direction - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 09:15 = 9:15 AM (under 12, so AM, same hour).
- 7:50 PM = 19:50 (PM, add 12, 7 + 12 = 19).
- 21:05 = 9:05 PM (over 12, subtract 12).
- 11:30 AM = 11:30 (AM, under 12, same hour with leading zero stays).

DO:
- Click to reveal.
- Run each rule.

TEACHER NOTES:
Reveal confirms the four cases. Watch for students who add or subtract correctly but lose the leading zero on small hours (eg. write 9:15 instead of 09:15 for 24 hour).

WATCH FOR:
- Students who got three of four - secure.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Two students disagree about 12:00 PM.
- Holly says 12:00 PM = 00:00 in 24 hour time.
- Ari says 12:00 PM = 12:00 in 24 hour time.
- Thumbs UP for Holly. Thumbs DOWN for Ari.

DO:
- Display the disagreement.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP for Holly. DOWN for Ari."
- Scan for: thumbs DOWN. Ari is right.
PROCEED: If 80% agree with Ari, click to reveal and confirm.
PIVOT: Most likely misconception - the 12 swap.
- Reteach: "12 PM is NOON. In 24 hour time noon is 12:00, NOT 00:00. 00:00 is MIDNIGHT, which is 12 AM."
- Re-check: "What is 12:00 AM in 24 hour time?"

TEACHER NOTES:
The hinge tests the noon/midnight trap. Many students freeze at 12. Pin it down with the visual: noon is exactly halfway through the day, and a clock shows 12.

WATCH FOR:
- Confident thumbs DOWN for Holly - they see noon = 12:00.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the 12 and 24 Hour Time sheet.
- Use your hand-held clock to check each one.

DO:
- Distribute the sheet.
- Circulate. Stop and ask "AM or PM? How do you know?"

TEACHER NOTES:
Different times from the We Do. Same rules.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use Part A only. Sit with these students. Show the day-line for each time before converting.
- Extra Notes: Hold the hand-held clock alongside.
EXTENDING PROMPT:
- Task: Part D (real schedule scenarios) - read a flight time and say what time it lands locally.
- Extra Notes: Encourage them to write both 24 hour and 12 hour versions.

WATCH FOR:
- Students who set the analog clock correctly - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Convert 16:20 to 12 hour time.
- Convert 8:45 AM to 24 hour time.

DO:
- Display the prompt.
- 3 minutes.

TEACHER NOTES:
Exit ticket assesses SC2. Look for 4:20 PM and 08:45.

WATCH FOR:
- Students who get both with the leading zero - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria.
- Show me thumbs up, sideways, or down for each.
- Turn and tell your partner: what is the trick for converting PM 24 hour times?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea: bigger than 12 = PM, subtract 12. Tomorrow we calculate elapsed time using the same clocks.

WATCH FOR:
- Strong thumbs up - move at pace tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a simple analog clock with hour and minute hands.
// hour: 1-12, minute: 0-59.
function drawAnalogClock(slide, cx, cy, r, hour, minute, opts) {
  const o = opts || {};
  const faceColor = o.faceColor || (C.WHITE);
  const ringColor = o.ringColor || C.CHARCOAL;
  const hourHandColor = o.hourHandColor || C.PRIMARY;
  const minuteHandColor = o.minuteHandColor || C.ALERT;

  // Outer ring
  slide.addShape("roundRect", {
    x: cx - r, y: cy - r, w: r * 2, h: r * 2, rectRadius: r,
    fill: { color: ringColor },
  });
  // Face
  slide.addShape("roundRect", {
    x: cx - r * 0.92, y: cy - r * 0.92, w: r * 1.84, h: r * 1.84,
    rectRadius: r * 0.92,
    fill: { color: faceColor },
  });
  // 12, 3, 6, 9 numerals (no other hour marks to keep it clean)
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
  // Tick marks at each hour
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
  // Hour hand
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
  // Minute hand
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
  // Centre dot
  slide.addShape("roundRect", {
    x: cx - 0.06, y: cy - 0.06, w: 0.12, h: 0.12, rectRadius: 0.06,
    fill: { color: C.CHARCOAL },
  });
}

// Draw a horizontal "day line" showing the 24-hour day broken at noon.
function drawDayLine(slide, x, y, w) {
  const h = 0.18;
  // Main bar
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: C.BG_LIGHT || "F6F6F6" },
    line: { color: C.CHARCOAL, width: 1 },
  });
  // AM half
  slide.addShape("rect", {
    x, y, w: w / 2, h,
    fill: { color: C.PRIMARY },
  });
  // PM half
  slide.addShape("rect", {
    x: x + w / 2, y, w: w / 2, h,
    fill: { color: C.ALERT },
  });
  // Hour labels
  const labels = ["00:00", "06:00", "12:00", "18:00", "23:59"];
  const positions = [0, 0.25, 0.5, 0.75, 1];
  labels.forEach((l, i) => {
    const lx = x + positions[i] * w - 0.45;
    slide.addText(l, {
      x: lx, y: y + h + 0.04, w: 0.90, h: 0.26,
      fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
      bold: true, align: "center", margin: 0,
    });
    // Tick mark
    slide.addShape("line", {
      x: x + positions[i] * w, y: y - 0.05, w: 0, h: h + 0.05,
      line: { color: C.CHARCOAL, width: 1.5 },
    });
  });

  // AM / PM labels
  slide.addText("AM (midnight to noon)", {
    x, y: y - 0.30, w: w / 2, h: 0.24,
    fontSize: 11, fontFace: FONT_B, color: C.PRIMARY,
    bold: true, align: "center", margin: 0,
  });
  slide.addText("PM (noon to midnight)", {
    x: x + w / 2, y: y - 0.30, w: w / 2, h: 0.24,
    fontSize: 11, fontFace: FONT_B, color: C.ALERT,
    bold: true, align: "center", margin: 0,
  });
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 5: 12 hour and 24 hour time",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal — subtracting decimals
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Subtracting decimals",
      ["2.5 - 0.45 =", "6.3 - 1.7 =", "4.05 - 1.20 ="],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "2.05         4.6         2.85", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal — addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Addition",
      ["9 + 4", "6 + 7", "8 + 5"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "13     13     13", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Key vocabulary — AM, PM, 24 hour
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "AM, PM, and 24 hour time");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const cellW = 2.7;
    const gap = 0.20;
    const totalW = cellW * 3 + gap * 2;
    const startX = (10 - totalW) / 2;
    const cardY = CONTENT_TOP + 0.30;
    const cardH = 2.50;

    const items = [
      { label: "AM",       desc: "morning",          detail: "midnight to noon", color: C.PRIMARY },
      { label: "PM",       desc: "afternoon",        detail: "noon to midnight", color: C.ALERT },
      { label: "24 hour",  desc: "no AM or PM",      detail: "00:00 to 23:59",   color: C.ACCENT },
    ];

    items.forEach((it, i) => {
      const cx = startX + i * (cellW + gap);
      addCard(s, cx, cardY, cellW, cardH, { strip: it.color, fill: C.WHITE });
      addTextOnShape(s, it.label, {
        x: cx + 0.20, y: cardY + 0.20, w: cellW - 0.40, h: 0.70, rectRadius: 0.10,
        fill: { color: it.color },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(it.desc, {
        x: cx, y: cardY + 1.05, w: cellW, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", margin: 0,
      });
      s.addText(it.detail, {
        x: cx, y: cardY + 1.55, w: cellW, h: 0.40,
        fontSize: 14, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to read and convert between 12 hour and 24 hour time.",
    [
      "I can read a 24 hour time and say whether it is AM or PM.",
      "I can convert a 24 hour time into a 12 hour time (and back).",
      "I can set the time on a hand-held analog clock to match a digital time.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: I Do (1) — day line
  workedExSlide(pres, 2, "I Do", "The 24 hour day",
    [
      "The day starts at 00:00 (midnight).",
      "12:00 is noon.",
      "The day ends at 23:59.",
      "",
      "Before noon = AM.",
      "After noon = PM.",
      "",
      "Bigger than 12? PM.",
      "13:00 = 1 PM.  23:00 = 11 PM.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("The day at a glance", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });

      drawDayLine(slide, lg.rightX + 0.20, lg.panelTopPadded + 0.95, lg.rightW - 0.40);

      // A small mapping table underneath
      addTextOnShape(slide, "13:00 = 1 PM", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 1.85, w: lg.rightW - 0.40, h: 0.35, rectRadius: 0.06,
        fill: { color: C.SECONDARY },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "18:00 = 6 PM", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.25, w: lg.rightW - 0.40, h: 0.35, rectRadius: 0.06,
        fill: { color: C.ACCENT },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "23:00 = 11 PM", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.65, w: lg.rightW - 0.40, h: 0.35, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10: I Do (2) — convert 14:30
  workedExSlide(pres, 2, "I Do", "Convert 14:30 to 12 hour time",
    [
      "14 is bigger than 12.",
      "Bigger than 12 = PM.",
      "",
      "Subtract 12 from the hour.",
      "14 - 12 = 2.",
      "",
      "Minutes stay the same.",
      "",
      "14:30 = 2:30 PM.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("14:30 = 2:30 PM", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.15, w: lg.rightW - 0.20, h: 0.45,
        fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Show analog clock at 2:30
      const ccx = lg.rightX + lg.rightW / 2;
      const ccy = lg.panelTopPadded + 1.80;
      drawAnalogClock(slide, ccx, ccy, 0.85, 2, 30);

      slide.addText("Hand-held clock check", {
        x: lg.rightX, y: lg.panelTopPadded + 2.80, w: lg.rightW, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slides 11-12: CFU + reveal — 18:45
  withReveal(
    () => cfuSlide(pres, "CFU", "Convert 18:45 to 12 hour time", "Show Me Boards + Clocks",
      "On your whiteboard:\n\nConvert 18:45 to 12 hour time.\nSet your hand-held clock to that time.\n\nAM or PM?  Hour?  Minutes?",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "18:45 = 6:45 PM   (bigger than 12 = PM, 18 - 12 = 6)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 13-14: We Do + reveal — four mixed conversions
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Convert each time with your partner", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      const items = [
        { from: "09:15",    arrow: "->", to: "? AM/PM" },
        { from: "7:50 PM",  arrow: "->", to: "? (24h)" },
        { from: "21:05",    arrow: "->", to: "? AM/PM" },
        { from: "11:30 AM", arrow: "->", to: "? (24h)" },
      ];

      const cellW = 2.0;
      const gap = 0.15;
      const totalW = cellW * 4 + gap * 3;
      const startX = (10 - totalW) / 2;
      const cardY = CONTENT_TOP + 0.30;
      const cardH = 1.60;

      items.forEach((it, i) => {
        const cx = startX + i * (cellW + gap);
        addCard(s, cx, cardY, cellW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
        s.addText(it.from, {
          x: cx, y: cardY + 0.20, w: cellW, h: 0.40,
          fontSize: 18, fontFace: FONT_H, color: C.PRIMARY,
          bold: true, align: "center", margin: 0,
        });
        s.addText(it.arrow, {
          x: cx, y: cardY + 0.70, w: cellW, h: 0.30,
          fontSize: 16, fontFace: FONT_B, color: C.MUTED,
          align: "center", margin: 0,
        });
        s.addText(it.to, {
          x: cx, y: cardY + 1.05, w: cellW, h: 0.40,
          fontSize: 16, fontFace: FONT_H, color: C.SECONDARY,
          bold: true, align: "center", margin: 0,
        });
      });

      s.addText("Partner talk: AM or PM? Hour change? Then set the clock.", {
        x: 0.7, y: SAFE_BOTTOM - 0.55, w: 8.6, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "09:15 = 9:15 AM   |   7:50 PM = 19:50   |   21:05 = 9:05 PM   |   11:30 AM = 11:30", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 15-16: CFU hinge — Holly vs Ari
  withReveal(
    () => cfuSlide(pres, "CFU", "Who is right?", "Thumbs Up or Thumbs Down",
      "What is 12:00 PM in 24 hour time?\n\nHolly says 12:00 PM = 00:00.\n\nAri says 12:00 PM = 12:00.\n\nThumbs UP for Holly.   Thumbs DOWN for Ari.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Ari is right. 12:00 PM is NOON, which is 12:00. (00:00 is midnight = 12 AM.)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: 12 and 24 Hour Time", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "decide AM or PM.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "convert the hour.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "check on your hand-held clock.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Remember", {
      x: 0.7, y: panelY + 0.10, w: 8.6, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "Bigger than 12 = PM. Subtract 12.", {
      x: 1.0, y: panelY + 0.45, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Mini analog clock + day line
    drawAnalogClock(s, 2.5, panelY + 1.60, 0.65, 3, 45);
    drawDayLine(s, 4.0, panelY + 1.85, 5.0);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 18: Exit Ticket
  exitTicketSlide(pres,
    [
      "Convert 16:20 to 12 hour time. AM or PM? Hour? Minutes?",
      "Convert 8:45 AM to 24 hour time. Use the leading zero.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 19: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the trick for converting PM 24 hour times?",
      scItems: [
        "I can read a 24 hour time and say whether it is AM or PM.",
        "I can convert a 24 hour time into a 12 hour time (and back).",
        "I can set the time on a hand-held analog clock to match a digital time.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Meas_Lesson5_12_and_24_Hour_Time.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Use your hand-held clock to check each time.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "AM = morning (midnight to noon). PM = afternoon and evening (noon to midnight). 24 hour: bigger than 12 means PM - subtract 12 to get the 12 hour version.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Part A — 24 hour to 12 hour", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  08:30 = ___ AM/PM", y);
    y = addWriteLine(doc, "b)  13:15 = ___ AM/PM", y);
    y = addWriteLine(doc, "c)  17:45 = ___ AM/PM", y);
    y = addWriteLine(doc, "d)  22:10 = ___ AM/PM", y);

    y = addSectionHeading(doc, "Part B — 12 hour to 24 hour", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  6:20 AM = ___", y);
    y = addWriteLine(doc, "b)  3:55 PM = ___", y);
    y = addWriteLine(doc, "c)  9:30 PM = ___", y);
    y = addWriteLine(doc, "d)  12:00 PM = ___   (noon)", y);

    y = addSectionHeading(doc, "Part C — Set the clock", y, { color: C.PRIMARY });
    y = addBodyText(doc, "On your hand-held clock, set the time and have your partner check.", y);
    y = addWriteLine(doc, "a)  15:25 = ___ AM/PM   (set hand-held clock and check)", y);
    y = addWriteLine(doc, "b)  06:40 = ___ AM/PM   (set hand-held clock and check)", y);

    y = addSectionHeading(doc, "Part D — Extension: real schedules", y, { color: C.ACCENT });
    y = addBodyText(doc, "A plane is scheduled to depart at 18:30.", y);
    y = addWriteLine(doc, "a)  Write the time in 12 hour time. ___", y);
    y = addWriteLine(doc, "b)  Is the flight in the morning or evening? ___", y);
    y = addBodyText(doc, "A school day starts at 9:00 AM and finishes at 3:30 PM.", y);
    y = addWriteLine(doc, "c)  Write both times in 24 hour time. ___ and ___", y);

    addPdfFooter(doc, `Lesson ${SESSION} | 12 and 24 Hour Time | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the 12 and 24 Hour Time sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Part A — 24 hour to 12 hour", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  08:30 = 8:30 AM    (less than 12)", y);
    y = addBodyText(doc, "b)  13:15 = 1:15 PM    (13 - 12 = 1)", y);
    y = addBodyText(doc, "c)  17:45 = 5:45 PM    (17 - 12 = 5)", y);
    y = addBodyText(doc, "d)  22:10 = 10:10 PM    (22 - 12 = 10)", y);

    y = addSectionHeading(doc, "Part B — 12 hour to 24 hour", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  6:20 AM = 06:20    (AM, leading zero)", y);
    y = addBodyText(doc, "b)  3:55 PM = 15:55    (PM, 3 + 12 = 15)", y);
    y = addBodyText(doc, "c)  9:30 PM = 21:30    (PM, 9 + 12 = 21)", y);
    y = addBodyText(doc, "d)  12:00 PM = 12:00    (noon - stays as 12:00)", y);

    y = addSectionHeading(doc, "Part C — Set the clock", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  15:25 = 3:25 PM    (clock hour hand between 3 and 4, minute hand on 5)", y);
    y = addBodyText(doc, "b)  06:40 = 6:40 AM    (clock hour hand between 6 and 7, minute hand on 8)", y);

    y = addSectionHeading(doc, "Part D — Real schedules", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  18:30 = 6:30 PM.", y);
    y = addBodyText(doc, "b)  Evening (PM).", y);
    y = addBodyText(doc, "c)  9:00 AM = 09:00. 3:30 PM = 15:30.", y);

    y = addTipBox(doc,
      "Watch for: students who forget AM/PM when converting from 24 hour; students who add 12 (instead of subtracting) for 24h to 12h; students who get 12:00 AM/PM swapped; students who drop the leading zero in 06:20.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 5 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
