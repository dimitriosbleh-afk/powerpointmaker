"use strict";

// Decimals & Integers Unit (Year 6 Numeracy) - Lesson 2: Comparing and ordering decimals
// VC2M5N01. Compare and order decimals to 3+ places; locate decimals on a number line.
// Daily Review: Tessellations. Fluency: vertical addition algorithm (4-digit).
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
  addPlaceValueChart, addDecimalDot, addNumberLine,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 4;
const UNIT_TITLE = "Decimals and Integers";
const FOOTER = `Decimals & Integers | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/DecInt_Lesson2_Comparing_Ordering_Decimals";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Comparing and Ordering Decimals",
  "Compare decimals with the place value rule, order sets, and locate decimals on a number line.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the comparing and ordering decimals practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 2 of our decimals and integers unit.
- Last lesson we built decimals to thousandths. Today we compare them and put them in order.
- We will also place decimals exactly on a number line.

DO:
- Have whiteboards, markers and printed place value charts ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 2 of 4. Comparing and ordering builds directly on place value from Lesson 1. The launch re-activates that so a returning student can rejoin here.

WATCH FOR:
- Students who relied on more-digits-is-bigger last lesson - today fixes that idea for good.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do section.

DO:
- Print one practice sheet and answer key per student.
- Have whiteboards, markers and printed place value charts ready.
- Optional: a printed blank number line for small group support.

TEACHER NOTES:
One student practice sheet plus answer key. The sheet has a built-in enabling start (decimals lined up for you) and an extension that orders longer sets and reasons about a midpoint.

CATCH-UP NOTE:
A student who missed Lesson 1 can still access today. The launch rebuilds the place value chart, and the enabling section lines the decimals up so the comparison is visible. Spend one minute with returners on the place value chart first.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Back to tessellations.
- Some shapes always tessellate, no matter how you draw them.
- For each one, write yes or no on your whiteboard.

DO:
- Display the three shapes.
- 60 seconds.
- Listen for reasoning about angles meeting.

TEACHER NOTES:
Prior geometry, not today's decimals. Any triangle tessellates, any quadrilateral tessellates, and the regular hexagon tessellates. This contrasts with the pentagon from Lesson 1.

WATCH FOR:
- Students who remember the pentagon did not tessellate - secure.
- Students unsure - prompt them to picture copies fitting with no gaps.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Shape A, a triangle - yes. Any triangle tessellates.
- Shape B, a quadrilateral - yes. Any quadrilateral tessellates.
- Shape C, a regular hexagon - yes. Three meet at each point.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick-and-fix.

TEACHER NOTES:
The big idea is angles meeting exactly around a point. Note students still unsure for small group geometry support.

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Vertical addition again, larger numbers today.
- Set each one out vertically, line up the columns, carry where you need to.

DO:
- Display the three prompts.
- 2 minutes.
- Scan for clean column alignment and carries.

TEACHER NOTES:
Same algorithm as Lesson 1 with four-digit numbers and more carrying. Lining up columns is the exact habit we use to line up decimal points - name that link.

WATCH FOR:
- Students who line up columns and carry cleanly - secure.
- Students who lose a carry across several columns - slow them down.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 3487 plus 2956 is 6443.
- 5064 plus 1938 is 7002.
- 7209 plus 894 is 8103.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The middle one carries into the thousands and makes a zero in the tens - a common slip. The third lines a 3-digit under a 4-digit number, so column alignment matters.

WATCH FOR:
- Students who self-correct - secure.
- Students who misalign the shorter number - reteach right-align of the ones.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember our place value chart from last lesson.
- Here are two decimals: 0.45 and 0.5.
- On your whiteboard, write which one you think is bigger.
- Hold it up. Do not change it yet - we will test our thinking together.

DO:
- Display 0.45 and 0.5.
- Take a quick show of boards.
- Note the split between the two answers without confirming yet.

TEACHER NOTES:
This launch surfaces the more-digits-is-bigger misconception before the learning intention. Many will pick 0.45 because it looks longer. Hold the answer - we resolve it in the I Do.

WATCH FOR:
- Students who pick 0.5 and can say five tenths beats four tenths - strong place value.
- Students who pick 0.45 - exactly the thinking we fix today.

[Stage: Launch | VTLM 2.0: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to compare and order decimals, and place them on a number line.
- Now the three I can statements. Read them with me.

DO:
- Choral read the LI and SC.
- Hold up the printed place value chart and a blank number line.

TEACHER NOTES:
SC1 is achievable for everyone - comparing two decimals with the same number of places. SC2 is the core target the exit ticket checks. SC3 stretches to ordering longer sets and justifying a number line position.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- Two ideas for today.
- Equivalent decimals are equal in value: 0.4 equals 0.40 equals 0.400. Extra zeros on the end do not change the size.
- Greater than and less than. The open mouth always faces the bigger number.

DO:
- Point to 0.4 = 0.40 = 0.400 as you say it.
- Air-write the greater-than and less-than symbols with the class.

TEACHER NOTES:
Equivalent decimals let us line numbers up by adding trailing zeros so every number has the same number of places. That makes comparing a place-by-place job.

WATCH FOR:
- Students who see 0.4 and 0.40 as equal - secure.
- Students who think 0.40 is bigger because it is longer - revisit with tenths.

[General: Key Vocabulary | VTLM 2.0: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Watch me compare 0.45 and 0.5.
- First I line them up by place. I write 0.50 so both have two places. The extra zero does not change the value.
- Now I compare from the left. Tenths first: 5 tenths against 4 tenths.
- 5 tenths is bigger, so 0.5 is bigger than 0.45.
- Longer is not bigger. Same place, biggest digit, wins.

DO:
- Write 0.45 and 0.50 lined up in the chart.
- Point to the tenths column and compare.
- Say the rule: line up, then compare from the left.

TEACHER NOTES:
This is the core rule for the unit. Lining up with a trailing zero turns comparing into a left-to-right place check. It directly resolves the launch.

MISCONCEPTIONS:
- Misconception: 0.45 is bigger than 0.5 because it has more digits.
  Why: students transfer whole-number length thinking.
  Impact: every comparison and ordering task goes wrong.
  Quick correction: line up the places, compare tenths first. 5 tenths beats 4 tenths.

WATCH FOR:
- Students who compare tenths first - secure.
- Students still counting digits - re-point to the tenths column.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Watch me place 2.335 on a number line.
- This line is zoomed in between 2.33 and 2.34.
- I split that gap into ten equal steps - those are thousandths.
- 2.335 is exactly halfway, five steps along.
- So 2.33 is less than 2.335 is less than 2.34. It sits right in the middle.

DO:
- Point to the two endpoints first.
- Count the ten steps with the class.
- Mark 2.335 at the midpoint and say it is between the two.

TEACHER NOTES:
This models locating a decimal by zooming in. The endpoints define the place we are working in - thousandths between hundredths. Halfway is the clearest first case.

MISCONCEPTIONS:
- Misconception: students think 2.335 sits near 2.3 or near 2.4.
  Why: they ignore the zoomed scale.
  Impact: they cannot justify number line positions.
  Quick correction: read the endpoints. We are between 2.33 and 2.34, not 2.3 and 2.4.

WATCH FOR:
- Students who find the midpoint quickly - secure.
- Students who count steps wrongly - count together from the left endpoint.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Which is bigger: 3.06 or 3.6?
- Line them up first, then compare from the left. Write the bigger one and a reason.

DO:
- Display the prompt.
- 45 seconds.
- Walk and scan for lined-up working.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me the bigger number on three, two, one, show.
- Scan for: 3.6, with reasoning about 6 tenths against 0 tenths.
PROCEED: If about 80 percent write 3.6, click to reveal and move to We Do.
PIVOT: Most likely misconception - students pick 3.06 because it has more digits.
- Reteach: line up as 3.60 and 3.06. Compare tenths: 6 tenths beats 0 tenths.
- Re-check: which has more tenths, and which is bigger?

TEACHER NOTES:
This is the launch misconception in a new outfit. 3.06 looks longer to some students. Lining up exposes the empty tenths place in 3.06.

WATCH FOR:
- Students who line up and compare tenths - secure.
- Students who pick 3.06 - more-digits misconception, reteach with the lined-up version.

[Stage 2: CFU | VTLM 2.0: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Order these from smallest to largest: 0.5, 0.405, 0.45.
- First line them up with trailing zeros so each has three places.
- Then compare from the left, place by place.

DO:
- Display the three decimals.
- 90 seconds.
- Listen for line-up-then-compare reasoning.

TEACHER NOTES:
Lining up gives 0.500, 0.405, 0.450. Compare tenths first: 4, 4, 5. Two tie at 4 tenths, so go to hundredths to separate 0.405 and 0.450.

WATCH FOR:
- Pairs who add trailing zeros to line up - secure.
- Pairs who order by length - reteach the line-up step.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Lined up: 0.405, 0.450, 0.500.
- Tenths: 4, 4, 5. The 5 is biggest, so 0.5 is largest.
- The two 4-tenths numbers split at the hundredths: 0 against 5. So 0.405 is smallest.
- Order: 0.405, then 0.45, then 0.5.

DO:
- Click to reveal.
- Trace tenths, then hundredths, with your finger.

TEACHER NOTES:
The reveal shows that ties at one place are broken at the next place to the right. That is the whole ordering method.

WATCH FOR:
- Students who self-correct - secure.
- Students who still rank by length - small group before You Do.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together on the number line.
- This line is zoomed between 5.28 and 5.29.
- Place 5.283. Whisper to your partner: is it closer to 5.28 or to 5.29?

DO:
- Display the zoomed number line.
- 75 seconds.
- Listen for closer-to-5.28 reasoning.

TEACHER NOTES:
5.283 is three steps of a thousandth past 5.28, so it sits below halfway and is closer to 5.28. This builds the rounding intuition for later years without naming rounding yet.

WATCH FOR:
- Pairs who say closer to 5.28 with a reason - secure.
- Pairs who guess the middle - count the steps from 5.28 together.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- 5.283 is three small steps past 5.28.
- That is below halfway, so it is closer to 5.28.
- 5.28 is less than 5.283 is less than 5.29.

DO:
- Click to reveal the marked point.
- Point to how far it sits from each end.

TEACHER NOTES:
If pairs struggled, count the ten steps again from the left endpoint before You Do.

WATCH FOR:
- Students who justify closer-to-5.28 - ready for independent work.
- Students who placed it past halfway - enabling group for You Do.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1: compare two decimals using greater than or less than.
- Section 2: order each set from smallest to largest.
- Section 3: place each decimal on the number line.
- If you finish, try the extension.

DO:
- Distribute the practice sheet.
- Circulate and listen for line-up-then-compare language.
- Cold call one or two students to justify a number line position.

TEACHER NOTES:
Different numbers from the We Do, same method: line up, compare from the left, break ties at the next place.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: In Section 1 the decimals are already lined up with trailing zeros. Do Sections 1 and 2 only, using the place value chart.
- Extra Notes: Sit with these students and do the first comparison together.
EXTENDING PROMPT:
- Task: Extension - order a set of five decimals to thousandths, and decide which of two decimals a given value is closer to on the number line, with a reason.
- Extra Notes: Push place-value justification. Strong students can preview the Year 8 challenge sheet in the unit.

WATCH FOR:
- Students who line up before comparing - secure.
- Students who order by length - prompt them to add trailing zeros first.

[Stage 4: You Do | VTLM 2.0: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard.
- Put these in order from smallest to largest: 0.6, 0.06, 0.66, 0.606.
- Then mark roughly where 0.66 would sit between 0.6 and 0.7 on the line.

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses SC2 - ordering decimals and placing one on a number line. Look for trailing-zero line-up (0.600, 0.060, 0.660, 0.606) and left-to-right comparison.

WATCH FOR:
- Students who order 0.06, 0.6, 0.606, 0.66 - secure.
- Students who put 0.606 or 0.66 first because they are longer - more-digits misconception, revisit next lesson.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: what is the first thing you do before comparing two decimals?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is line up the places first, then compare from the left, breaking ties at the next place. Students who can say this are ready for integers next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on SC2 - small group revision at the start of Lesson 3.

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
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Comparing and ordering decimals",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Tessellations
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Daily Review: Do these tessellate?", { color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      const shapeY = CONTENT_TOP + 0.45;
      // A - triangle (any triangle)
      s.addShape("triangle", {
        x: 1.35, y: shapeY, w: 1.8, h: 1.6,
        fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("A", {
        x: 1.35, y: shapeY + 1.7, w: 1.8, h: 0.4,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      // B - quadrilateral (trapezoid)
      s.addShape("trapezoid", {
        x: 4.05, y: shapeY + 0.25, w: 1.9, h: 1.2,
        fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("B", {
        x: 4.05, y: shapeY + 1.7, w: 1.9, h: 0.4,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      // C - regular hexagon
      s.addShape("hexagon", {
        x: 6.95, y: shapeY, w: 1.8, h: 1.6,
        fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("C", {
        x: 6.95, y: shapeY + 1.7, w: 1.8, h: 0.4,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });

      s.addText("For each shape write yes or no: does it tessellate on its own?", {
        x: 0.7, y: SAFE_BOTTOM - 0.5, w: 8.6, h: 0.35,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A triangle = yes     B quadrilateral = yes     C hexagon = yes", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - vertical addition (4-digit)
  withReveal(
    () => fluencySlide(pres, "Fluency: Vertical addition",
      ["3487 + 2956", "5064 + 1938", "7209 + 894"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "6443        7002        8103", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - which is bigger?
  contentSlide(pres, "Launch", C.ACCENT, "Which one is bigger?",
    [
      "Write your first answer.",
      "Hold it up - do not change it yet.",
      "",
      "We will test our thinking together.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.ACCENT });
      const cy = lg.panelTopPadded + 0.45;
      addTextOnShape(slide, "0.45", {
        x: lg.rightX + 0.5, y: cy, w: 1.5, h: 1.0, rectRadius: 0.10,
        fill: { color: C.PRIMARY },
      }, { fontSize: 40, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("or", {
        x: lg.rightX + 2.0, y: cy, w: 0.8, h: 1.0,
        fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      addTextOnShape(slide, "0.5", {
        x: lg.rightX + 2.8, y: cy, w: 1.4, h: 1.0, rectRadius: 0.10,
        fill: { color: C.SECONDARY },
      }, { fontSize: 40, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("Write the bigger one on your whiteboard.", {
        x: lg.rightX + 0.2, y: cy + 1.15, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to compare and order decimals and place them on a number line.",
    [
      "I can compare two decimals that have the same number of places.",
      "I can order decimals by lining up the place values and compare from the left.",
      "I can place a decimal on a number line and justify its position.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Equivalent decimals and comparing",
    [
      "Equivalent decimals are equal in value.",
      "0.4 = 0.40 = 0.400",
      "Trailing zeros do not change the size.",
      "The open mouth faces the bigger number: > and <.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
      slide.addText("Same value, lined up", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const eqY = lg.panelTopPadded + 0.6;
      addTextOnShape(slide, "0.4  =  0.40  =  0.400", {
        x: lg.rightX + 0.35, y: eqY, w: lg.rightW - 0.7, h: 0.7, rectRadius: 0.10,
        fill: { color: C.PRIMARY },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("Add zeros on the end so both decimals have the same number of places.", {
        x: lg.rightX + 0.2, y: eqY + 0.85, w: lg.rightW - 0.4, h: 0.6,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", valign: "top", margin: 0,
      });
    }
  );

  // Slide 10: I Do #1 - compare 0.45 and 0.5
  workedExSlide(pres, 2, "I Do", "Line up, then compare from the left",
    [
      "Compare 0.45 and 0.5.",
      "",
      "Line up: write 0.50 so both",
      "have two places.",
      "",
      "Compare tenths first:",
      "5 tenths beats 4 tenths.",
      "",
      "So 0.5 > 0.45.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.55, { strip: C.PRIMARY });
      slide.addText("Line them up (same places)", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      const ry0 = lg.panelTopPadded + 0.55;
      addTextOnShape(slide, "0.50", {
        x: lg.rightX + 1.05, y: ry0, w: 2.1, h: 0.62, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 30, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "0.45", {
        x: lg.rightX + 1.05, y: ry0 + 0.74, w: 2.1, h: 0.62, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 30, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("Compare the tenths: 5 beats 4, so 0.5 is bigger.", {
        x: lg.rightX + 0.15, y: ry0 + 1.52, w: lg.rightW - 0.3, h: 0.40,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slide 11: I Do #2 - locate 2.335 on a number line
  workedExSlide(pres, 2, "I Do", "Place 2.335 on a number line",
    [
      "Zoom between 2.33 and 2.34.",
      "",
      "Split the gap into 10 steps.",
      "Those steps are thousandths.",
      "",
      "2.335 is halfway - 5 steps along.",
      "",
      "2.33 < 2.335 < 2.34.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.PRIMARY });
      slide.addText("Zoomed between 2.33 and 2.34", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      const nlY = lg.panelTopPadded + 1.35;
      addNumberLine(slide, lg.rightX + 0.35, nlY, lg.rightW - 0.7,
        ["2.33", "", "", "", "", "2.335", "", "", "", "", "2.34"], [5]);
      slide.addText("2.335 sits right in the middle.", {
        x: lg.rightX + 0.2, y: nlY + 0.5, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slides 12-13: CFU + reveal - compare 3.06 and 3.6
  withReveal(
    () => cfuSlide(pres, "CFU", "Which is bigger?", "Show Me Boards",
      "Compare 3.06 and 3.6.\n\nLine them up first, then compare from the left.\n\nWrite the bigger number and one reason.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3.6 is bigger.   Lined up: 3.60 vs 3.06.   6 tenths beats 0 tenths.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - order three decimals
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Order from smallest to largest",
      [
        "With your partner.",
        "",
        "Order: 0.5, 0.405, 0.45",
        "",
        "1.  Add trailing zeros to line up.",
        "2.  Compare tenths, then hundredths.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
        slide.addText("Line them up", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        const rows = ["0.500", "0.405", "0.450"];
        const ry0 = lg.panelTopPadded + 0.55;
        rows.forEach((r, i) => {
          addTextOnShape(slide, r, {
            x: lg.rightX + 1.1, y: ry0 + i * 0.58, w: 2.0, h: 0.48, rectRadius: 0.06,
            fill: { color: C.PRIMARY },
          }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
        });
        slide.addText("Compare the tenths column first.", {
          x: lg.rightX + 0.2, y: ry0 + 1.82, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "Smallest to largest:   0.405   <   0.45   <   0.5", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - locate 5.283
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Closer to which end?",
      [
        "With your partner.",
        "",
        "Place 5.283 on the line",
        "between 5.28 and 5.29.",
        "",
        "Is it closer to 5.28 or 5.29?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
        slide.addText("Zoomed between 5.28 and 5.29", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        const nlY = lg.panelTopPadded + 1.35;
        addNumberLine(slide, lg.rightX + 0.35, nlY, lg.rightW - 0.7,
          ["5.28", "", "", "", "", "", "", "", "", "", "5.29"], []);
        slide.addText("Count the steps from 5.28.", {
          x: lg.rightX + 0.2, y: nlY + 0.5, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "5.283 is 3 steps past 5.28 - below halfway, so closer to 5.28.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "compare two decimals with > or <.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "order each set, smallest to largest.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "place each decimal on the number line.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.4;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember", {
      x: 0.7, y: panelY + 0.13, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
    });
    addTextOnShape(s, "Line up the places first. Then compare from the left. Break ties at the next place.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    addNumberLine(s, 2.0, panelY + 1.55, 6.0,
      ["0.4", "", "", "", "", "0.45", "", "", "", "", "0.5"], []);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Order from smallest to largest: 0.6, 0.06, 0.66, 0.606.",
      "Mark roughly where 0.66 sits between 0.6 and 0.7 on a number line.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the first thing you do before comparing two decimals?",
      scItems: [
        "I can compare two decimals that have the same number of places.",
        "I can order decimals by lining up the place values and compare from the left.",
        "I can place a decimal on a number line and justify its position.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecInt_Lesson2_Comparing_Ordering_Decimals.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Compare, order, and place decimals on a number line.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Line up the places first (add trailing zeros so each decimal has the same number of places). Then compare from the LEFT. Break ties at the next place to the right.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Compare 0.45 and 0.5. Line up: 0.45 and 0.50. Tenths: 4 vs 5. So 0.45 < 0.50, which means 0.5 is bigger.",
      y);

    y = addSectionHeading(doc, "Section 1 - Compare with > or < (lined up for you)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  0.70 ____ 0.68", y);
    y = addWriteLine(doc, "b)  3.06 ____ 3.60", y);
    y = addWriteLine(doc, "c)  1.250 ____ 1.205", y);

    y = addSectionHeading(doc, "Section 2 - Order each set from smallest to largest", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  0.3   0.33   0.303   ->  ____  ,  ____  ,  ____", y);
    y = addWriteLine(doc, "b)  2.5   2.05   2.55   ->  ____  ,  ____  ,  ____", y);
    y = addWriteLine(doc, "c)  0.9   0.09   0.909   ->  ____  ,  ____  ,  ____", y);

    y = addSectionHeading(doc, "Section 3 - Place each decimal on the number line", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Draw an arrow to roughly the right spot. The line runs from 4.20 to 4.30 in steps of 0.01.", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  4.25     b)  4.23     c)  4.285", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order from smallest to largest:  1.4   1.04   1.404   1.044   1.44", y);
    y = addWriteLine(doc, "Order: ____________________________________________________________", y);
    y = addBodyText(doc, "Is 7.236 closer to 7.23 or 7.24 on the number line? Explain.", y);
    y = addWriteLine(doc, "Answer: ___________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Comparing and Ordering Decimals | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the comparing and ordering decimals practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Compare", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  0.70 > 0.68      b)  3.06 < 3.60      c)  1.250 > 1.205", y);

    y = addSectionHeading(doc, "Section 2 - Order (smallest to largest)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  0.3, 0.303, 0.33     (lined up: 0.300, 0.303, 0.330)", y);
    y = addBodyText(doc, "b)  2.05, 2.5, 2.55      (lined up: 2.05, 2.50, 2.55)", y);
    y = addBodyText(doc, "c)  0.09, 0.9, 0.909     (lined up: 0.090, 0.900, 0.909)", y);

    y = addSectionHeading(doc, "Section 3 - Number line (4.20 to 4.30)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4.25 sits halfway.   b)  4.23 is 3 steps in.   c)  4.285 is between 4.28 and 4.29, closer to 4.29.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order: 1.04, 1.044, 1.4, 1.404, 1.44   (lined up: 1.040, 1.044, 1.400, 1.404, 1.440).", y);
    y = addBodyText(doc, "7.236 is closer to 7.24: it is 6 steps of a thousandth past 7.23, which is past halfway.", y);

    y = addTipBox(doc,
      "Watch for: students who order by length (0.606 before 0.66); students who forget trailing zeros when lining up; students who ignore the zoomed scale on the number line.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
