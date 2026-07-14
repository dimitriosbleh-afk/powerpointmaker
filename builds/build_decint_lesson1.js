"use strict";

// Decimals & Integers Unit (Year 6 Numeracy) - Lesson 1: Decimal place value to thousandths
// VC2M5N01. First explicit teaching of thousandths and the x10 relationship between places.
// Daily Review: Tessellations. Fluency: vertical addition algorithm.
// Unit variant fixed across all 4 lessons for cohesion.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide,
  addPlaceValueChart, addDecimalDot,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 4;
const UNIT_TITLE = "Decimals and Integers";
const FOOTER = `Decimals & Integers | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/DecInt_Lesson1_Decimal_Place_Value";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Decimal Place Value to Thousandths",
  "Name the place, build decimals to thousandths in a chart, and match words to digits.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the decimal place value practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to our new unit on decimals and integers.
- Over the next four lessons we build decimals carefully and then meet a brand new kind of number called integers.
- Today we are zooming in on place value, all the way to thousandths.

DO:
- Have whiteboards, markers, and printed place value charts ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 1 of 4. This is the foundation for the whole unit. Secure tenths, hundredths and thousandths today and comparing decimals next lesson runs smoothly.

WATCH FOR:
- Students who look unsure - that is expected on day one. Reassure: if this feels new, that is okay, we build it together.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do section near the end.

DO:
- Print one practice sheet and answer key per student.
- Have whiteboards, markers and printed place value charts ready.
- Optional: MAB blocks or a metre tape for small group support.

TEACHER NOTES:
One student practice sheet plus answer key. Most teaching is on whiteboards with the printed place value chart. The sheet has a built-in enabling start and an extension section for students who are ready.

CATCH-UP NOTE:
A student who is away this week can rejoin at any lesson. Each lesson re-activates the place value chart in the launch, so a returner only needs the printed chart and one minute with you.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are thinking about tessellations from our geometry work.
- A shape tessellates when copies of it cover a surface with no gaps and no overlaps.
- For each shape, write yes or no on your whiteboard.

DO:
- Display the three shapes.
- 60 seconds.
- Walk and listen for reasoning about gaps.

TEACHER NOTES:
Daily Review is prior geometry, not today's decimals. Tessellations connect to angles meeting at a point. Squares and equilateral triangles tessellate; regular pentagons leave gaps.

WATCH FOR:
- Students who explain using gaps - secure.
- Students guessing - prompt them to picture the shapes fitting together.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Shape A, the equilateral triangle - yes, it tessellates. Six fit around a point.
- Shape B, the square - yes, four fit around a point.
- Shape C, the regular pentagon - no, it leaves gaps.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick-and-fix.

TEACHER NOTES:
The key idea is that angles must meet exactly around a point. Note any student who thinks every shape tessellates - revisit in small group.

WATCH FOR:
- Students who self-correct quickly - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Vertical addition.
- Set each one out vertically on your whiteboard, ones under ones, tens under tens.
- Add, carry where you need to, and write the total.

DO:
- Display the three prompts.
- 90 seconds.
- Scan for clean column alignment.

TEACHER NOTES:
Fluency this whole unit is the vertical addition algorithm. Lining up columns now is the same habit that lines up decimal points later in the unit. Keep it brisk.

WATCH FOR:
- Students who line up columns neatly - secure.
- Students who forget to carry - prompt: what happens when a column makes ten or more.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 348 plus 275 is 623.
- 1246 plus 583 is 1829.
- 609 plus 487 is 1096.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Look for the carry from the ones into the tens, and tens into hundreds. Students who misalign columns get the wrong total - that is the thing to coach.

WATCH FOR:
- Students who self-correct - secure.
- Students whose digits drift out of columns - small group focus.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember tenths and hundredths from earlier work.
- Look at the chart. We have ones, then tenths, then hundredths.
- Whisper to your partner: what do you think the next place to the right is called?
- Today we zoom in one more step, to thousandths.

DO:
- Point to each column heading as you read it.
- Take two or three predictions for the empty column.
- Reveal the word thousandths.

TEACHER NOTES:
This launch bridges known tenths and hundredths to the new place, thousandths. It activates prior knowledge before the learning intention, exactly where the new word is going.

WATCH FOR:
- Students who predict thousandths or a tenth-of-a-hundredth idea - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | VTLM 2.0: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to build and name decimals to thousandths using place value.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the LI and the three I can statements.
- Hold up the printed place value chart.

TEACHER NOTES:
SC1 is achievable for everyone - it is naming tenths and hundredths. SC2 is the core target the exit ticket checks. SC3 stretches to explaining the ten-times relationship.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- A tenth is one whole split into ten equal parts. We write it 0.1.
- A hundredth is one tenth split into ten again. We write it 0.01.
- A thousandth is one hundredth split into ten once more. We write it 0.001.
- Each step to the right is ten times smaller.

DO:
- Point to each row as you say it.
- Have students say each word and show the size with their hands getting smaller.

TEACHER NOTES:
Vocabulary comes after the learning intention. Keep it to these three words. The big idea is the relationship - each place is ten times smaller than the one to its left.

WATCH FOR:
- Students who can order the three sizes - secure.
- Students who think more digits means smaller pieces only - clarify with the chart.

[General: Key Vocabulary | VTLM 2.0: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Watch me build a decimal one place at a time.
- I start with 2 ones. I put the 2 in the ones place.
- The decimal point is the marker between whole and parts.
- Now 4 tenths goes in the tenths place. The number is 2.4.
- Now 5 hundredths in the hundredths place. The number is 2.45.
- Now 6 thousandths in the thousandths place. The number is 2.456.
- I read the smallest place at the end: two and four hundred fifty-six thousandths.

DO:
- Fill the chart cell by cell as you speak.
- Point to the decimal point and call it the marker.
- Say the read-it sentence twice.

TEACHER NOTES:
This is the core build move. Reading the smallest place at the end keeps place value language alive. Push two and four hundred fifty-six thousandths, not two point four five six.

MISCONCEPTIONS:
- Misconception: students read 2.456 as two point four five six.
  Why: they treat the point as a divider.
  Impact: they cannot compare or order decimals next lesson.
  Quick correction: name the smallest place. Read it as thousandths.

WATCH FOR:
- Students who echo the place names - tracking.
- Students who say two point four five six - prompt for the place name.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Watch how the places relate to each other.
- Look at the arrows. As I move one place to the right, the value gets ten times smaller.
- One tenth is the same as ten hundredths. One hundredth is the same as ten thousandths.
- That means thousandths are ten times smaller than hundredths.
- Moving left is the opposite - ten times bigger each step.

DO:
- Trace the ten-times-smaller arrow to the right.
- Trace the ten-times-bigger arrow to the left.
- Have students chorus ten times smaller as you point right.

TEACHER NOTES:
This is the multiplicative relationship from the curriculum. It also sets up renaming, for example 0.6 divided by 10 becomes 0.06, six tenths renamed as sixty hundredths then divided. Keep that as a stretch idea in the extension, not on the slide.

MISCONCEPTIONS:
- Misconception: students think each place is just one less, not ten times smaller.
  Why: they count places like whole-number columns without the size idea.
  Impact: they struggle to rename decimals for mental computation.
  Quick correction: one tenth equals ten hundredths - show it on the chart.

WATCH FOR:
- Students who say ten times smaller without prompting - secure.
- Students who say a bit smaller - re-point to the arrows.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- The number is 3.072.
- Write each digit under the correct place name: ones, tenths, hundredths, thousandths.
- Then circle the digit in the thousandths place.

DO:
- Display the prompt.
- 45 seconds.
- Walk and scan the charts.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me your chart on three, two, one, show.
- Scan for: 3 ones, 0 tenths, 7 hundredths, 2 thousandths, with 2 circled.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students drop the zero and write 3, 7, 2.
- Reteach: the zero holds the tenths place so every digit keeps its value. Rebuild 3.072 cell by cell.
- Re-check: where does the zero go in 3.072, and what is its place?

TEACHER NOTES:
The placeholder zero is the trap. A student who drops it reads 3.072 as 3.72 and loses a whole place.

WATCH FOR:
- Students who hold the zero in tenths - secure.
- Students who skip the zero - placeholder misconception.

[Stage 2: CFU | VTLM 2.0: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Build 1.305 in the chart on your whiteboard.
- Whisper to your partner: which place has the zero, and why does it matter?
- Then read the number using place names.

DO:
- Display 1.305 above a blank chart.
- 90 seconds.
- Listen for one and three hundred five thousandths.

TEACHER NOTES:
Same structure as the I Do with new numbers and a placeholder zero in the hundredths place. Listen for the place-value reading.

WATCH FOR:
- Pairs who keep the zero in hundredths - secure.
- Pairs who write 1.35 - the placeholder zero is missing, reteach with the chart.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- 1 one. Decimal point. 3 tenths. 0 hundredths. 5 thousandths.
- We read it: one and three hundred five thousandths.
- The zero holds the hundredths place so the 5 stays in thousandths.

DO:
- Click to reveal.
- Run the place names once more together.

TEACHER NOTES:
Reveal restates the place value. The zero is the teaching point - without it the 5 would slide into hundredths.

WATCH FOR:
- Students who self-correct - secure.
- Students who still drop the zero - small group focus before You Do.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, the other way around.
- I say the words: five and eighty-three thousandths.
- On your whiteboard, write that as a decimal.
- Think carefully about how many places you need.

DO:
- Display the words.
- 75 seconds.
- Watch for the placeholder zero in the tenths place.

TEACHER NOTES:
Words to digits is harder because eighty-three thousandths needs a zero in the tenths place to push 83 into the hundredths and thousandths. This is the most common error, so reveal slowly.

WATCH FOR:
- Students who write 5.083 - secure.
- Students who write 5.83 - they forgot the placeholder zero, reteach with the chart.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Five and eighty-three thousandths is 5.083.
- The zero in tenths makes room so 83 sits in the hundredths and thousandths places.

DO:
- Click to reveal.
- Point to the placeholder zero.

TEACHER NOTES:
If many wrote 5.83, do one more words-to-digits example before releasing to the You Do.

WATCH FOR:
- Students who placed the zero correctly - ready for independent work.
- Students who wrote 5.83 - enabling group for You Do.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1: name the place of the underlined digit.
- Section 2: write each decimal into the chart.
- Section 3: match the words to the decimal.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for place-name language.
- Cold call one or two students to explain a thousandths answer.

TEACHER NOTES:
Different numbers from the We Do, same strategy: chart, name the place, watch the placeholder zero.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: The first chart row is started for you. Use the printed place value chart and do Sections 1 and 2 only.
- Extra Notes: Sit with these students and build the first one together.
EXTENDING PROMPT:
- Task: Extension section - rename decimals by dividing by 10 (for example 0.6 divided by 10) and order four decimals to thousandths, then explain how you decided.
- Extra Notes: Push place-value language in the explanation. Strong students can preview the Year 8 challenge sheet in the unit.

WATCH FOR:
- Students who use place names fluently - secure.
- Students who write point seven instead of seven tenths - prompt the place name.

[Stage 4: You Do | VTLM 2.0: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Write 6.209 into a place value chart and name each digit's place.
- Then read 6.209 aloud to your partner using place names.

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses SC2 - writing a decimal to thousandths into a chart and reading it. Look for the placeholder zero in the hundredths place and place-value reading.

WATCH FOR:
- Students who place 6 ones, 2 tenths, 0 hundredths, 9 thousandths - secure.
- Students who drop the zero - placeholder misconception, revisit in Lesson 2.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: what does each step to the right do to the size of the place?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that each place is ten times smaller as you move right, and the placeholder zero keeps every digit in its place. Students who can say this are ready to compare and order decimals next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on SC2 - small group revision at the start of Lesson 2.

[General: Closing | VTLM 2.0: Reflection]`;

// --- Helpers -----------------------------------------------------------------

function placePvChart(slide, x, y, headers, values, opts) {
  const o = opts || {};
  const totalW = o.totalW || 4.0;
  const valH = o.valH || 0.55;
  const hdrH = o.hdrH || 0.34;
  const chart = addPlaceValueChart(slide, x, y, headers, values, {
    totalW, valH, hdrH,
    headerColor: o.headerColor || C.PRIMARY,
  });
  if (o.dotAfter != null) {
    addDecimalDot(slide, chart, o.dotAfter, { position: "baseline" });
  }
  return chart;
}

const PV4 = ["Ones", "Tenths", "Hundredths", "Thousandths"];

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Decimal place value to thousandths",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Tessellations
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Daily Review: Which shapes tessellate?", { color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      const shapeY = CONTENT_TOP + 0.45;
      // A - equilateral triangle
      s.addShape("triangle", {
        x: 1.35, y: shapeY, w: 1.7, h: 1.6,
        fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("A", {
        x: 1.35, y: shapeY + 1.7, w: 1.7, h: 0.4,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
      // B - square
      s.addShape("rect", {
        x: 4.2, y: shapeY + 0.05, w: 1.5, h: 1.5,
        fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("B", {
        x: 4.05, y: shapeY + 1.7, w: 1.8, h: 0.4,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
      // C - regular pentagon
      s.addShape("pentagon", {
        x: 6.95, y: shapeY, w: 1.7, h: 1.6,
        fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("C", {
        x: 6.95, y: shapeY + 1.7, w: 1.7, h: 0.4,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });

      s.addText("For each shape write yes or no: does it tessellate on its own?", {
        x: 0.7, y: SAFE_BOTTOM - 0.5, w: 8.6, h: 0.35,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A triangle = yes     B square = yes     C pentagon = no (gaps)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - vertical addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Vertical addition",
      ["348 + 275", "1246 + 583", "609 + 487"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "623        1829        1096", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - place value zoom
  contentSlide(pres, "Launch", C.ACCENT, "Zooming in on place value",
    [
      "We know tenths and hundredths.",
      "",
      "What comes next, to the right?",
      "Each step right is 10 times smaller.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.ACCENT });
      slide.addText("Predict the next place", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      const chartY = lg.panelTopPadded + 0.55;
      placePvChart(slide, lg.rightX + 0.10, chartY,
        ["Ones", "Tenths", "Hundredths", "?"],
        ["", "", "", "?"],
        { totalW: lg.rightW - 0.20, valH: 0.70, hdrH: 0.50, dotAfter: 0 });
      slide.addText("Each step to the right: 10 times smaller", {
        x: lg.rightX + 0.10, y: chartY + 1.30, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to build and name decimals to thousandths using place value.",
    [
      "I can name the tenths and hundredths place in a decimal.",
      "I can write a decimal with thousandths into a place value chart and read it.",
      "I can explain why each place is 10 times smaller than the place to its left.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Tenths, hundredths, thousandths",
    [
      "Tenth = 0.1 (one whole split into 10)",
      "Hundredth = 0.01 (a tenth split into 10)",
      "Thousandth = 0.001 (a hundredth split into 10)",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
      slide.addText("Each step: 10 times smaller", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [
        ["1 tenth", "0.1"],
        ["1 hundredth", "0.01"],
        ["1 thousandth", "0.001"],
      ];
      const ry0 = lg.panelTopPadded + 0.55;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.58;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.25, y: ry, w: 2.1, h: 0.48, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true });
        slide.addText(r[1], {
          x: lg.rightX + 2.5, y: ry, w: lg.rightW - 2.7, h: 0.48,
          fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10: I Do #1 - build 2.456 in the chart
  workedExSlide(pres, 2, "I Do", "Build a decimal one place at a time",
    [
      "Start with 2 ones.",
      "Add 4 tenths -> 2.4",
      "Add 5 hundredths -> 2.45",
      "Add 6 thousandths -> 2.456",
      "",
      "Read the smallest place last:",
      "two and four hundred fifty-six thousandths.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.PRIMARY });
      slide.addText("2.456", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.45,
        fontSize: 30, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      const chartY = lg.panelTopPadded + 0.66;
      placePvChart(slide, lg.rightX + 0.10, chartY, PV4,
        ["2", "4", "5", "6"],
        { totalW: lg.rightW - 0.20, valH: 0.72, hdrH: 0.50, dotAfter: 0 });
      slide.addText("Read: two and four hundred fifty-six thousandths", {
        x: lg.rightX + 0.10, y: chartY + 1.32, w: lg.rightW - 0.20, h: 0.34,
        fontSize: 11.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 11: I Do #2 - the ten-times relationship
  workedExSlide(pres, 2, "I Do", "Each place is 10 times smaller",
    [
      "Move RIGHT: 10 times smaller.",
      "Move LEFT: 10 times bigger.",
      "",
      "1 tenth = 10 hundredths",
      "1 hundredth = 10 thousandths",
      "",
      "So thousandths are 10 times",
      "smaller than hundredths.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.PRIMARY });
      const chartY = lg.panelTopPadded + 0.55;
      const chart = placePvChart(slide, lg.rightX + 0.10, chartY, PV4,
        ["1", "0", "0", "0"],
        { totalW: lg.rightW - 0.20, valH: 0.66, hdrH: 0.50, dotAfter: 0 });
      // x10 smaller arrow (to the right)
      slide.addText(">>  10 times smaller  >>", {
        x: lg.rightX + 0.10, y: chartY + 1.28, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 13, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", margin: 0,
      });
      slide.addText("<<  10 times bigger  <<", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      void chart;
    }
  );

  // Slides 12-13: CFU + reveal - place 3.072
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Place 3.072 in the chart", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 8.0, y: 0.20, w: 1.4, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 19, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 9, breakLine: true } },
        { text: "Write each digit of 3.072", options: { fontSize: 17, color: C.CHARCOAL, breakLine: true } },
        { text: "under the right place name.", options: { fontSize: 17, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 9, breakLine: true } },
        { text: "Circle the thousandths digit.", options: { fontSize: 18, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 9, breakLine: true } },
        { text: "Watch the zero!", options: { fontSize: 18, color: C.ALERT, bold: true } },
      ], {
        x: 0.72, y: CONTENT_TOP + 0.18, w: 3.85, h: SAFE_BOTTOM - CONTENT_TOP - 0.34,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.1, CONTENT_TOP, 4.4, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("3.072", {
        x: 5.1, y: CONTENT_TOP + 0.15, w: 4.4, h: 0.50,
        fontSize: 32, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      placePvChart(s, 5.25, CONTENT_TOP + 1.0, PV4,
        ["", "", "", ""],
        { totalW: 4.10, valH: 0.95, hdrH: 0.55, dotAfter: 0 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Ones 3  |  Tenths 0  |  Hundredths 7  |  Thousandths (2)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - build 1.305
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Build 1.305 together",
      [
        "With your partner.",
        "",
        "1.  Write 1.305 into the chart.",
        "2.  Which place holds the zero?",
        "3.  Read it using place names.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
        slide.addText("1.305", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.42,
          fontSize: 28, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        placePvChart(slide, lg.rightX + 0.10, lg.panelTopPadded + 0.62, PV4,
          ["", "", "", ""],
          { totalW: lg.rightW - 0.20, valH: 0.80, hdrH: 0.55, dotAfter: 0 });
      }),
    (slide) => {
      addTextOnShape(slide, "1 . 3 tenths . 0 hundredths . 5 thousandths  =  one and three hundred five thousandths", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - words to digits
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Words to a decimal",
      [
        "I say the words.",
        "You write the decimal.",
        "",
        "five and eighty-three",
        "thousandths",
        "",
        "How many places do you need?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.3, { strip: C.SECONDARY });
        slide.addText("five and eighty-three thousandths", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.15, w: lg.rightW - 0.4, h: 0.7,
          fontSize: 17, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        slide.addText("Write it as a decimal on your whiteboard.", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.0, w: lg.rightW - 0.4, h: 0.5,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", valign: "middle", margin: 0,
        });
        placePvChart(slide, lg.rightX + 0.30, lg.panelTopPadded + 1.55, PV4,
          ["", "", "", ""],
          { totalW: lg.rightW - 0.60, valH: 0.55, hdrH: 0.40, dotAfter: 0 });
      }),
    (slide) => {
      addTextOnShape(slide, "5.083   (the zero in tenths makes room for 83 in hundredths and thousandths)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // Slide 18: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "name the place of each underlined digit.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "write each decimal into the chart.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "match the words to the decimal.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.4;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember", {
      x: 0.7, y: panelY + 0.13, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "Read the smallest place last. The zero is a placeholder - do not drop it.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    placePvChart(s, 2.5, panelY + 1.15, PV4, [".", "", "", ""],
      { totalW: 5.0, valH: 0.55, hdrH: 0.38, dotAfter: 0 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Write 6.209 into a place value chart and name each digit's place.",
      "Read 6.209 aloud using place names (not point two zero nine).",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what does each step to the right do to the size of the place?",
      scItems: [
        "I can name the tenths and hundredths place in a decimal.",
        "I can write a decimal with thousandths into a place value chart and read it.",
        "I can explain why each place is 10 times smaller than the place to its left.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecInt_Lesson1_Decimal_Place_Value.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Name the place, build decimals to thousandths, and match words to digits.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Each place is 10 times smaller as you move right: ones, tenths, hundredths, thousandths. Read the smallest place last. A zero is a placeholder - keep it.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "In 3.4_5_6 the underlined 6 is in the THOUSANDTHS place. We read 3.456 as 'three and four hundred fifty-six thousandths'.",
      y);

    y = addSectionHeading(doc, "Section 1 - Name the place of the underlined digit", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2.4_5_6     The 5 is in the _________________ place.", y);
    y = addWriteLine(doc, "b)  0.00_7_     The 7 is in the _________________ place.", y);
    y = addWriteLine(doc, "c)  5._1_28     The 1 is in the _________________ place.", y);
    y = addWriteLine(doc, "d)  9.30_4_     The 4 is in the _________________ place.", y);

    y = addSectionHeading(doc, "Section 2 - Write each decimal into the place value chart", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Headers: Ones | Tenths | Hundredths | Thousandths. The first one is started for you.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  2.456   Ones: 2    Tenths: 4    Hundredths: ___    Thousandths: ___", y);
    y = addWriteLine(doc, "b)  1.305   Ones: ___  Tenths: ___  Hundredths: ___    Thousandths: ___", y);
    y = addWriteLine(doc, "c)  0.087   Ones: ___  Tenths: ___  Hundredths: ___    Thousandths: ___", y);

    y = addSectionHeading(doc, "Section 3 - Match the words to the decimal", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  six and twenty-five thousandths       = __________", y);
    y = addWriteLine(doc, "b)  four and seven hundredths             = __________", y);
    y = addWriteLine(doc, "c)  five and eighty-three thousandths     = __________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Rename by dividing by 10: 0.6 divided by 10 = __________   (Hint: every digit shifts one place right.)", y);
    y = addWriteLine(doc, "Order from smallest to largest:  0.5    0.405    0.45    0.045", y);
    y = addWriteLine(doc, "Order: ____________________________________________________________", y);
    y = addWriteLine(doc, "Explain how you decided: ___________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Decimal Place Value to Thousandths | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the decimal place value practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Name the place", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  hundredths      b)  thousandths      c)  tenths      d)  thousandths", y);

    y = addSectionHeading(doc, "Section 2 - Place value chart", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2.456   Ones 2 | Tenths 4 | Hundredths 5 | Thousandths 6", y);
    y = addBodyText(doc, "b)  1.305   Ones 1 | Tenths 3 | Hundredths 0 | Thousandths 5   (zero holds the hundredths place)", y);
    y = addBodyText(doc, "c)  0.087   Ones 0 | Tenths 0 | Hundredths 8 | Thousandths 7", y);

    y = addSectionHeading(doc, "Section 3 - Words to a decimal", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  6.025      b)  4.07      c)  5.083", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "0.6 divided by 10 = 0.06. Order: 0.045 < 0.405 < 0.45 < 0.5.", y);
    y = addBodyText(doc, "Why: compare the tenths first (0 vs 4 vs 4 vs 5), then hundredths to separate 0.405 and 0.45.", y);

    y = addTipBox(doc,
      "Watch for: students who drop the placeholder zero (1.305 read as 1.35); students who write 5.83 for five and eighty-three thousandths; students who think more digits always means bigger.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
