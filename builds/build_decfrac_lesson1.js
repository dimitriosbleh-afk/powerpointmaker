"use strict";

// Decimals & Fractions Unit (Year 5/6 Numeracy) — Lesson 1: Place value with decimals
// VC2M6N02 — apply knowledge of place value to add and subtract decimals.
// Daily Review: Solving equations with multiplication, division, and operations.
// Fluency: Adding and subtracting decimals.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addPvChartPdf, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Theme: Year 5/6 numeracy. Variant 0 fixed across all 6 lessons of this unit
// (CLAUDE.md theme cohesion rule: same palette across a unit).
const T = createTheme("numeracy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addPlaceValueChart, addDecimalDot, addNumberLine,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 6;
const UNIT_TITLE = "Decimals and Fractions";
const FOOTER = `Decimals & Fractions | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecFrac_Lesson1_Place_Value_With_Decimals";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 1 Place Value Practice",
  "Place value charts, decimal naming, and decimal comparison.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 1 Answer Key",
  "Worked answers for the place value practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome to a new unit. This is the first time we are looking at decimals and fractions for the year.
- Over the next six lessons we will build up to adding and subtracting decimals and fractions.
- Today we are going back to place value so the rest of the unit makes sense.

DO:
- Have whiteboards, markers, and the place value reference charts ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 1 of 6. Place value is the foundation for the whole unit. If students are not sure how the digits in 3.456 line up, decimal addition later in the week will collapse. Treat this as a careful refresh.

WATCH FOR:
- Students who look uncertain about thousandths - that is okay. We build it together.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The place value practice sheet is for the You Do section and table work.

DO:
- Print one copy of the practice sheet and answer key.
- Have whiteboards, markers, and the printed place value reference cards ready.
- Optional: physical MAB blocks or place value counters for support students.

TEACHER NOTES:
One student resource (practice sheet) plus answer key. Most work is on whiteboards. Print MAB or place value reference cards if you have a small group that needs the concrete model.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Warm up time. These come from last week.
- Look at each equation. Whisper the missing number to your partner.
- Then write the answer on your whiteboard.

DO:
- Display the three equations.
- Allow 90 seconds.
- Walk and listen for the inverse-operation language.

TEACHER NOTES:
Daily Review retrieves equation-solving with multiplication and division. This is prior learning, not today's focus. Listen for "I divided" or "I used the inverse" - that is the language we want to retrieve.

WATCH FOR:
- Students who guess without using the inverse - prompt "What is the opposite of multiply?"
- Students who get all three quickly - secure on inverse operations.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- For 6 times n equals 48, n is 8. We divide 48 by 6.
- For n divided by 5 equals 9, n is 45. We multiply 9 by 5.
- For 12 plus n equals 30, n is 18. We subtract 12 from 30.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick-and-fix.

TEACHER NOTES:
Inverse operations are the strategy. Note any student still guessing - small group focus this week.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students who said 8 for the second one - they multiplied instead of using the inverse.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency time. Adding and subtracting decimals.
- Whisper your answer for each one, then write it on your board.
- Stack the digits so the decimal points line up.

DO:
- Display the three prompts.
- 60 seconds working time.
- Use whiteboards.

TEACHER NOTES:
Fluency starts today even though formal addition is Lesson 3. These are tenths only, so students can rely on what they already know. The aim is to surface place value habits we will use later.

WATCH FOR:
- Students who line up the right-hand digit instead of the decimal point - flag for I Do.
- Students who get all three correct - secure with tenths.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 0.4 plus 0.3 is 0.7.
- 0.9 minus 0.5 is 0.4.
- 1.2 plus 0.6 is 1.8.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Tick and fix. The point of the reveal is the decimal-point lining-up habit. If any student wrote 0.07 for the first, they treated the dot as a separator instead of a place value marker - that is the misconception we tackle in I Do.

WATCH FOR:
- Students who wrote 0.07 for 0.4 + 0.3 - place value misconception.
- Students who lined up the decimal points without prompting - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to use place value to read, write and compare decimals to thousandths.
- Now the success criteria. Read each one together.

DO:
- Choral read.
- Hold up the printed place value chart.

TEACHER NOTES:
SC1 is achievable for everyone today - just naming the place. SC2 is the core target the exit ticket assesses. SC3 stretches to comparison and explanation.

WATCH FOR:
- Students who can name tenths but not thousandths - that is expected pre-teach.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. The number is 3.456.
- I am going to put each digit into a place value chart.
- The 3 sits in the ones place. That is 3 ones.
- The 4 sits in the tenths place. That is 4 tenths, or four out of ten.
- The 5 sits in the hundredths place. That is 5 hundredths.
- The 6 sits in the thousandths place. That is 6 thousandths.
- I read it: "three and four hundred fifty-six thousandths."

DO:
- Point to each cell as you say the name.
- Trace the decimal point as a marker between whole numbers and parts.
- Repeat the read-it sentence twice.

TEACHER NOTES:
First model. The decimal point is a marker between the ones and the tenths, not a separator. Say the place name every time you point. Many Year 5/6 students miss that thousandths is one-tenth of a hundredth.

MISCONCEPTIONS:
- Misconception: Students read 3.456 as "three point four hundred fifty-six."
  Why: They treat the decimal point as a divider.
  Impact: They cannot compare decimals with different lengths later.
  Quick correction: "Each digit has a place name. Read the smallest place at the end."

WATCH FOR:
- Students who repeat the place names - tracking.
- Students who look confused - we use the chart again in the CFU.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Watch me compare two decimals.
- Which is bigger: 0.7 or 0.07?
- I put each one into the chart.
- 0.7 has 7 tenths. That is 7 out of 10.
- 0.07 has 0 tenths and 7 hundredths. That is 7 out of 100.
- 7 tenths is much bigger than 7 hundredths.
- So 0.7 is bigger than 0.07.

DO:
- Display both numbers in side-by-side charts.
- Underline the tenths place in both numbers.
- Say "compare the same place" three times.

TEACHER NOTES:
This addresses the "longer means bigger" misconception. 0.07 looks longer but is actually smaller. The rule is: compare the same place, starting from the largest place value, not the length of the number.

MISCONCEPTIONS:
- Misconception: 0.07 is bigger than 0.7 because it has more digits.
  Why: Students transfer whole-number "longer is bigger" thinking.
  Impact: They order decimals wrongly later this week.
  Quick correction: Use the place value chart. Compare the same place first.

WATCH FOR:
- Students who nod when you say "compare the same place" - tracking.
- Students who still want 0.07 to win - revisit with the chart and 100s grid.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. The number is 4.205.
- On your whiteboard, write each digit under the right place name.
- Ones, tenths, hundredths, thousandths.

DO:
- Display the prompt.
- Allow 30 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your chart on three, two, one, show."
- Scan for: 4 in ones, 2 in tenths, 0 in hundredths, 5 in thousandths.
PROCEED: If 80% have the correct chart, click to reveal and move to We Do.
PIVOT: Most likely misconception - students skip the zero and write 4, 2, 5.
- Reteach: "Each place has a digit. The zero holds the hundredths place. We do not skip it."
- Re-check: "Where does the zero go in 4.205?"

TEACHER NOTES:
Place value zero is the trap. Students who skip the zero will later misread 4.205 as 4.25.

WATCH FOR:
- Students who hold the zero in hundredths - secure.
- Students who skip the zero - place value misconception.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. We are going to compare 0.8 and 0.75.
- Which is bigger?
- Put each number into the chart on your whiteboard.
- Then whisper your answer to your partner.

DO:
- Display both numbers, blank charts beside each.
- 60 seconds working time.
- Walk and listen for "compare the same place."

TEACHER NOTES:
Same structure as the I Do but new numbers. Listen for students who notice 0.8 has 8 tenths while 0.75 only has 7 tenths. That is the relational move.

WATCH FOR:
- Pairs who say "compare the tenths first" - secure.
- Pairs who say 0.75 because it has more digits - reteach using the chart.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 0.8 has 8 tenths. 0.75 has 7 tenths.
- 8 tenths is bigger than 7 tenths.
- So 0.8 is bigger than 0.75.
- Even though 0.75 looks longer, the tenths place tells us the truth.

DO:
- Click to reveal.
- Underline the tenths digit in both numbers as you speak.

TEACHER NOTES:
The reveal restates the rule: compare the same place, starting from the largest. Students who landed on 0.75 are still using whole-number thinking.

WATCH FOR:
- Students who self-correct on the reveal - they see the trap now.
- Students who still want 0.75 - small group focus tomorrow.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_HINGE_Q = `SAY:
- Hinge question. Two students disagree.
- Sam says 0.4 is bigger than 0.36 because 4 is bigger than 36 in the same place.
- Alex says 0.36 is bigger because it has more digits.
- Thumbs up for Sam. Thumbs down for Alex.

DO:
- Display the disagreement.
- Wait 10 seconds for thinking.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs up if you agree with Sam. Thumbs down if you agree with Alex."
- Scan for: thumbs UP. Sam is right.
PROCEED: If 80% agree with Sam, click to reveal and confirm the rule.
PIVOT: Most likely misconception - students agree with Alex (more digits = bigger).
- Reteach: Use a tenths strip. 0.4 = 4 out of 10. 0.36 = 36 out of 100, which is 3 tenths and 6 hundredths.
- Re-check: "Which has more tenths?"

TEACHER NOTES:
This hinge probes whether students apply the "compare the same place" rule. If most agree with Alex, we need another tenths-strip model before You Do.

WATCH FOR:
- Confident thumbs up for Sam - they see the place value.
- Thumbs down for Alex or sideways - they need the tenths strip again.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the practice sheet from the table.
- Section 1: Write the place value of each underlined digit.
- Section 2: Compare two decimals. Circle the bigger one and explain.

DO:
- Distribute the practice sheet.
- Circulate, listen for "compare the same place."
- Cold call 1-2 students to share their explanation.

TEACHER NOTES:
Different numbers from the We Do. Same strategy: place value chart, compare the same place.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use only Section 1. Use the printed place value reference card beside you.
- Extra Notes: Sit with these students for the first one. Build the chart together.
EXTENDING PROMPT:
- Task: Section 3 - Build the largest and smallest decimal from the digits 3, 5, 0, 8. Explain why.
- Extra Notes: Encourage them to write a one-sentence reason using place value language.

WATCH FOR:
- Students who write the place name fluently - secure.
- Students who write "point seven" instead of "seven tenths" - prompt them to use the place name.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Question 1: Write 5.073 in the place value chart.
- Question 2: Which is bigger, 0.6 or 0.59? Circle the bigger one and explain.

DO:
- Display the prompt.
- Allow 2 minutes.
- Collect whiteboards or take a quick photo.

TEACHER NOTES:
Exit ticket assesses SC2 - reading, writing and comparing decimals using place value. Look for the zero in hundredths and the "compare the same place" reasoning.

WATCH FOR:
- Students who write 5 in ones, 0 in tenths, 7 in hundredths, 3 in thousandths - secure on place value.
- Students who circle 0.59 - still using "more digits is bigger" thinking. Small group focus tomorrow.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: what does the decimal point really mean?

DO:
- Read each I can statement.
- Use thumbs to self-assess.
- Cold call 1-2 students for the partner share.

TEACHER NOTES:
Self-assessment data informs Lesson 2 grouping. The threshold idea is "the decimal point separates the whole from the parts, and each place has a value." Students who can say this are ready for rounding tomorrow.

WATCH FOR:
- Strong thumbs up across all three SC - move at pace tomorrow.
- Sideways or down on SC2 or SC3 - small group revision in Lesson 2.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a labelled place value chart on a slide and optionally add a decimal dot
// between two columns. Returns the chart geometry.
function placePvChart(slide, x, y, headers, values, opts) {
  const o = opts || {};
  const totalW = o.totalW || 4.5;
  const valH = o.valH || 0.85;
  const hdrH = o.hdrH || 0.50;
  const chart = addPlaceValueChart(slide, x, y, headers, values, {
    totalW, valH, hdrH,
    headerColor: o.headerColor || C.PRIMARY,
  });
  if (o.dotAfter != null) {
    addDecimalDot(slide, chart, o.dotAfter, { position: "baseline" });
  }
  return chart;
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Place value with decimals",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources (immediately after title per megaprompt v10.3 §44)
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal — solving equations with multiplication/division
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Find the missing number",
      [
        "6 × n = 48",
        "n ÷ 5 = 9",
        "12 + n = 30",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "n = 8     n = 45     n = 18", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal — adding and subtracting decimals (tenths)
  withReveal(
    () => fluencySlide(pres, "Fluency: Adding & Subtracting Decimals",
      ["0.4 + 0.3", "0.9 − 0.5", "1.2 + 0.6"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "0.7     0.4     1.8", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: LI/SC
  liSlide(pres,
    "We are learning to use place value to read, write and compare decimals to thousandths.",
    [
      "I can name the place of each digit in a decimal (ones, tenths, hundredths, thousandths).",
      "I can read, write and place a decimal into a place value chart.",
      "I can compare two decimals by lining up the place values and explain which is bigger.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 8: I Do — place value chart for 3.456
  workedExSlide(pres, 2, "I Do", "Place value chart: 3.456",
    [
      "Each digit has a place name.",
      "3 sits in the ones place.",
      "4 sits in the tenths place.",
      "5 sits in the hundredths place.",
      "6 sits in the thousandths place.",
      "",
      "Read it: three and four hundred fifty-six thousandths.",
      "",
      "The decimal point is the marker.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      // Place value chart card
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.PRIMARY });
      slide.addText("3.456", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.45,
        fontSize: 28, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      const chartY = lg.panelTopPadded + 0.70;
      placePvChart(slide, lg.rightX + 0.10, chartY,
        ["Ones", "Tenths", "Hundredths", "Thousandths"],
        ["3", "4", "5", "6"],
        { totalW: lg.rightW - 0.20, valH: 0.70, hdrH: 0.42, dotAfter: 0 });

      // Read-aloud label below the chart
      slide.addText("Read: three and 456 thousandths", {
        x: lg.rightX + 0.10, y: chartY + 1.30, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 9: I Do — comparing 0.7 and 0.07 using the chart
  workedExSlide(pres, 2, "I Do", "Compare 0.7 and 0.07",
    [
      "Longer does NOT mean bigger.",
      "Use the place value chart.",
      "",
      "0.7 has 7 tenths.",
      "0.07 has 0 tenths and 7 hundredths.",
      "",
      "7 tenths > 7 hundredths.",
      "",
      "So 0.7 > 0.07.",
      "",
      "Rule: compare the same place,",
      "starting from the largest place.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      const cardH = 3.0;
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.SECONDARY });
      slide.addText("Same place, side by side", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Chart 1: 0.7
      slide.addText("0.7", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      placePvChart(slide, lg.rightX + 0.10, lg.panelTopPadded + 0.72,
        ["Ones", "Tenths", "Hundredths"],
        ["0", "7", "0"],
        { totalW: lg.rightW - 0.20, valH: 0.55, hdrH: 0.34, dotAfter: 0 });

      // Chart 2: 0.07
      slide.addText("0.07", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 1.70, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      placePvChart(slide, lg.rightX + 0.10, lg.panelTopPadded + 2.00,
        ["Ones", "Tenths", "Hundredths"],
        ["0", "0", "7"],
        { totalW: lg.rightW - 0.20, valH: 0.55, hdrH: 0.34, dotAfter: 0 });
    }
  );

  // Slides 10-11: CFU 1 + reveal — write 4.205 in the chart
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Place 4.205 in the chart", { color: C.ALERT });

      // CHECK stamp top-right
      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      // Left card: the prompt
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Write each digit of 4.205", options: { fontSize: 17, color: C.CHARCOAL, breakLine: true } },
        { text: "under the correct place name.", options: { fontSize: 17, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Watch the zero!", options: { fontSize: 18, color: C.ALERT, bold: true } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right card: empty chart, large
      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("4.205", {
        x: 5.3, y: CONTENT_TOP + 0.15, w: 4.2, h: 0.45,
        fontSize: 32, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      placePvChart(s, 5.45, CONTENT_TOP + 0.95,
        ["Ones", "Tenths", "Hundredths", "Thousandths"],
        ["", "", "", ""],
        { totalW: 3.90, valH: 0.85, hdrH: 0.42, dotAfter: 0 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Ones 4   |   Tenths 2   |   Hundredths 0   |   Thousandths 5", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: We Do + reveal — compare 0.8 and 0.75
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Compare 0.8 and 0.75",
      [
        "With your partner.",
        "",
        "Put each number into a chart.",
        "Compare the tenths first.",
        "Which is bigger?",
        "",
        "Be ready to say why using",
        "the words tenths and hundredths.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        const cardH = 3.0;
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.SECONDARY });
        slide.addText("Build both charts", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });

        slide.addText("0.8", {
          x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.20, h: 0.30,
          fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
          align: "center", margin: 0,
        });
        placePvChart(slide, lg.rightX + 0.10, lg.panelTopPadded + 0.72,
          ["Ones", "Tenths", "Hundredths"],
          ["", "", ""],
          { totalW: lg.rightW - 0.20, valH: 0.55, hdrH: 0.34, dotAfter: 0 });

        slide.addText("0.75", {
          x: lg.rightX + 0.10, y: lg.panelTopPadded + 1.70, w: lg.rightW - 0.20, h: 0.30,
          fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
          align: "center", margin: 0,
        });
        placePvChart(slide, lg.rightX + 0.10, lg.panelTopPadded + 2.00,
          ["Ones", "Tenths", "Hundredths"],
          ["", "", ""],
          { totalW: lg.rightW - 0.20, valH: 0.55, hdrH: 0.34, dotAfter: 0 });
      }),
    (slide) => {
      addTextOnShape(slide, "0.8 > 0.75   (8 tenths > 7 tenths)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 14-15: CFU hinge + reveal — Sam vs Alex
  withReveal(
    () => cfuSlide(pres, "CFU", "Who is right?", "Thumbs Up or Thumbs Down",
      "Sam:  0.4 is bigger than 0.36.\n     (4 tenths vs 3 tenths.)\n\nAlex: 0.36 is bigger because it has more digits.\n\nThumbs UP for Sam.   Thumbs DOWN for Alex.",
      NOTES_CFU_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Sam is right.   0.4 has 4 tenths.   0.36 has 3 tenths.   0.4 > 0.36.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 16: You Do — practice sheet from the table
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    // Top steps strip
    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 1 — name the place of each underlined digit.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 2 — compare and circle the bigger decimal.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Write one sentence to explain.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Bottom panel: visual hint + reminder
    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Remember the rule", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "Compare the SAME place, starting with the largest place.", {
      x: 1.5, y: panelY + 0.55, w: 7.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Mini reference chart
    placePvChart(s, 2.5, panelY + 1.15,
      ["Ones", "Tenths", "Hundredths", "Thousandths"],
      [".", "", "", ""],
      { totalW: 5.0, valH: 0.55, hdrH: 0.34, dotAfter: 0 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Write 5.073 into a place value chart. Name each digit's place.",
      "Which is bigger: 0.6 or 0.59? Circle it. Explain using the words tenths and hundredths.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what does the decimal point really mean?",
      scItems: [
        "I can name the place of each digit in a decimal (ones, tenths, hundredths, thousandths).",
        "I can read, write and place a decimal into a place value chart.",
        "I can compare two decimals by lining up the place values and explain which is bigger.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecFrac_Lesson1_Place_Value_With_Decimals.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Practice worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Place value, naming digits, and comparing decimals.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Compare the SAME place, starting from the largest place. The decimal point is the marker between ones and tenths.",
      y, { color: C.ACCENT });

    // Section 1: name the place
    y = addSectionHeading(doc, "Section 1 — Name the place of the underlined digit", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  4.5_2_     The 2 is in the _________________ place.", y);
    y = addWriteLine(doc, "b)  3._7_06    The 7 is in the _________________ place.", y);
    y = addWriteLine(doc, "c)  0.20_8_    The 8 is in the _________________ place.", y);
    y = addWriteLine(doc, "d)  _6_.345    The 6 is in the _________________ place.", y);

    // Section 2: place into chart (simple write lines)
    y = addSectionHeading(doc, "Section 2 — Write each decimal into the place value chart", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Headers: Ones | Tenths | Hundredths | Thousandths", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  2.45       Ones: ___  Tenths: ___  Hundredths: ___  Thousandths: ___", y);
    y = addWriteLine(doc, "b)  3.108      Ones: ___  Tenths: ___  Hundredths: ___  Thousandths: ___", y);
    y = addWriteLine(doc, "c)  0.072      Ones: ___  Tenths: ___  Hundredths: ___  Thousandths: ___", y);

    // Section 3: compare and explain
    y = addSectionHeading(doc, "Section 3 — Circle the bigger decimal and explain", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   0.7    or    0.65    Bigger: ______   Why: __________________________", y);
    y = addWriteLine(doc, "b)   0.3    or    0.29    Bigger: ______   Why: __________________________", y);
    y = addWriteLine(doc, "c)   2.40   or    2.4     Bigger: ______   Why: __________________________", y);

    // Optional extension
    y = addSectionHeading(doc, "Extension (optional) — Build the largest and smallest", y, { color: C.ACCENT });
    y = addBodyText(doc, "Use each of the digits 3, 5, 0, 8 exactly once. Place a decimal point anywhere.", y);
    y = addWriteLine(doc, "Largest decimal you can build: ___________________________________________", y);
    y = addWriteLine(doc, "Smallest decimal you can build: __________________________________________", y);
    y = addWriteLine(doc, "Explain (one sentence): __________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Place Value Practice | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the place value practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Name the place", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  hundredths        b)  tenths        c)  thousandths        d)  ones", y);

    y = addSectionHeading(doc, "Section 2 — Place value chart", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2.45     Ones 2  | Tenths 4  | Hundredths 5  | Thousandths 0", y);
    y = addBodyText(doc, "b)  3.108    Ones 3  | Tenths 1  | Hundredths 0  | Thousandths 8", y);
    y = addBodyText(doc, "c)  0.072    Ones 0  | Tenths 0  | Hundredths 7  | Thousandths 2", y);

    y = addSectionHeading(doc, "Section 3 — Compare and explain", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  0.7 is bigger.   0.7 has 7 tenths; 0.65 has 6 tenths.", y);
    y = addBodyText(doc, "b)  0.3 is bigger.   0.3 has 3 tenths; 0.29 has 2 tenths.", y);
    y = addBodyText(doc, "c)  Equal.   2.40 = 2.4 (the extra zero in hundredths adds nothing).", y);

    y = addSectionHeading(doc, "Extension — Largest and smallest", y, { color: C.ACCENT });
    y = addBodyText(doc, "Largest:  8530 (no decimal) or 853.0   Smallest:  0.358 or 0.0358 (if leading zero allowed).", y);
    y = addBodyText(doc, "Accept reasoning that places the largest digits in the largest place values, and vice versa.", y);

    y = addTipBox(doc,
      "Common errors to scan for: students who skip a zero in hundredths (3.108 written as 3.18); students who pick 0.65 as bigger than 0.7 because of more digits; students who do not recognise 2.40 = 2.4.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
