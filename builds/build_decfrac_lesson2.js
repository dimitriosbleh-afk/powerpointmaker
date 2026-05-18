"use strict";

// Decimals & Fractions Unit (Year 5/6 Numeracy) — Lesson 2: Estimation & rounding with decimals
// VC2M6N02 — apply place value to add/subtract decimals; use estimation & rounding
//            to check the reasonableness of answers.
// Daily Review: Solving equations with multiplication, division, and operations.
// Fluency: Adding and subtracting decimals.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addNumberLine,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 6;
const UNIT_TITLE = "Decimals and Fractions";
const FOOTER = `Decimals & Fractions | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecFrac_Lesson2_Estimation_And_Rounding";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 2 Rounding and Estimation Practice",
  "Round decimals, estimate sums, and check reasonableness.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 2 Answer Key",
  "Worked answers for the rounding and estimation sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we are going to round decimals and use estimation to check whether an answer makes sense.
- Estimation is a quick reasonableness check before we calculate, and again after.

DO:
- Have whiteboards and the printed number lines ready.

TEACHER NOTES:
Lesson 2 of 6. Rounding sits on top of yesterday's place value work. The big idea is reasonableness - knowing roughly what an answer should be before we calculate.

WATCH FOR:
- Students who think rounding means "chopping off" - we untangle that in I Do.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The rounding practice sheet is for the You Do section.

DO:
- Print one copy of the practice sheet and answer key.
- Have whiteboards, markers, and printed number lines ready.

TEACHER NOTES:
One student resource (practice sheet) plus answer key. The number line is the main visual model today.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Warm up. Find the missing number in each equation.
- Whisper the answer to your partner, then write it on your board.

DO:
- Display the three equations.
- Allow 90 seconds.
- Listen for the inverse-operation language.

TEACHER NOTES:
Daily Review continues equation-solving from Lesson 1. Same retrieval focus. This is prior learning.

WATCH FOR:
- Students who use the inverse fluently - secure.
- Students who guess - prompt "What is the opposite operation?"

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 7 times n equals 56. n is 8 (because 56 divided by 7 is 8).
- n divided by 4 equals 12. n is 48 (because 12 times 4 is 48).
- 9 plus n equals 21. n is 12 (because 21 minus 9 is 12).

DO:
- Click to reveal.
- Pause for tick-and-fix.

TEACHER NOTES:
Same retrieval as yesterday. Note any student still guessing - small group focus this week.

WATCH FOR:
- Strong self-correction - secure.
- Students who keep multiplying instead of dividing - reteach with the inverse.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Adding and subtracting decimals.
- Whisper your answer for each, then write it.
- Line up the decimal point.

DO:
- Display the three prompts.
- 60 seconds.
- Whiteboards.

TEACHER NOTES:
Fluency continues from Lesson 1. Tenths still. We want the line-up-the-dot habit forming.

WATCH FOR:
- Students who line up the decimal point without prompting - secure.
- Students who line up the right-hand digit - place value misconception.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 0.6 plus 0.4 is 1.0 (or just 1).
- 1.5 minus 0.7 is 0.8.
- 2.3 plus 1.6 is 3.9.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The first one matters - 0.6 plus 0.4 makes a whole. Students who write 0.10 are still treating the dot as a separator.

WATCH FOR:
- Students who wrote 0.10 for the first - reteach place value bundling.
- Students who self-correct on the reveal - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to round decimals and use estimation to check that our answers make sense.
- Now the success criteria.

DO:
- Choral read.
- Hold up the rounding number line.

TEACHER NOTES:
SC1 is achievable today - rounding to the nearest whole. SC2 is the core target - rounding to a stated place. SC3 stretches to using rounding for reasonableness.

WATCH FOR:
- Students who repeat the language - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me round 4.7 to the nearest whole number.
- I draw a number line from 4 to 5.
- I find 4.7 on the line. It is closer to 5 than to 4.
- The half-way point is 4.5. Anything 4.5 or above goes up to 5. Anything below 4.5 goes down to 4.
- 4.7 rounds to 5.

DO:
- Display the number line with 4.7 marked.
- Point to 4, 5, the half-way mark, and 4.7.
- Say the rule: "Find the nearest. Half-way and above rounds up."

TEACHER NOTES:
The number line makes "nearest" visible. The "5 or more, round up" rule sits on top of the picture, not in place of it.

MISCONCEPTIONS:
- Misconception: Round always means "chop the digits."
  Why: Students learn a shortcut without the meaning.
  Impact: They cannot round 4.7 because chopping gives 4.
  Quick correction: "Round means nearest. Find it on a number line first."

WATCH FOR:
- Students who repeat the rule - tracking.
- Students who say "chop the 7" - place value misconception.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now I round 3.46 to the nearest tenth.
- I look at the tenths digit. That is 4.
- I look at the digit right after it - the hundredths. That is 6.
- 6 is 5 or more, so I round up. The 4 becomes a 5.
- 3.46 rounds to 3.5.
- I draw the line from 3.4 to 3.5 to check. 3.46 sits past the half-way mark, so 3.5 wins.

DO:
- Display the number line from 3.4 to 3.5 with 3.46 marked.
- Underline the hundredths digit as you say "look at the next place."

TEACHER NOTES:
The "look at the next place" rule is the rounding heuristic. The number line confirms it. We connect the rule and the picture so students can choose either.

MISCONCEPTIONS:
- Misconception: Students round the tenths and the hundredths separately.
  Why: They apply the rule digit by digit instead of using the next place.
  Impact: 3.46 rounded becomes 3.5 instead of 3.5 - sometimes 3.47.
  Quick correction: "Look at the next place ONLY. Decide once."

WATCH FOR:
- Students who underline the hundredths quickly - secure.
- Students who try to round each digit - reteach.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Round 6.83 to the nearest whole number.
- On your whiteboard. Show me.

DO:
- Display the number line from 6 to 7.
- Allow 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your answer on three, two, one, show."
- Scan for: 7.
PROCEED: If 80% have 7, click to reveal and move on.
PIVOT: Most likely misconception - students wrote 6 because they ignored the tenths.
- Reteach: "6.83 sits past the half-way mark (6.5). It is closer to 7."
- Re-check: "Where does 6.83 sit on the line?"

TEACHER NOTES:
Quick check on the rule. The number line backs it up.

WATCH FOR:
- Students who hold up 7 - secure.
- Students who hold up 6 - they ignored the tenths place.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Round 8.27 to the nearest tenth.
- Step 1: look at the tenths place. What digit is there?
- Step 2: look at the next place - the hundredths. What digit?
- Step 3: 5 or more rounds up; less than 5 rounds down.
- Write your answer on your whiteboard.

DO:
- Display the prompt with the steps visible.
- 60 seconds.
- Walk and scan.

TEACHER NOTES:
Same structure as I Do, different number. The steps stay visible so students can self-cue.

WATCH FOR:
- Pairs who underline the hundredths digit - secure.
- Pairs who change every digit - reteach.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- The tenths digit is 2. The next place (hundredths) is 7. 7 is 5 or more, so we round up.
- 8.27 rounds to 8.3.

DO:
- Click to reveal.
- Run the rule once more aloud.

TEACHER NOTES:
The reveal restates the rule and the answer. Look for students who wrote 8.2 - they probably looked at the wrong place.

WATCH FOR:
- Students who self-correct - secure.
- Students who keep 8.2 - they did not look at the hundredths.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Use estimation to check this answer.
- The calculation: 6.8 + 3.1 = 19.9
- Is this answer reasonable?
- Round each number to the nearest whole first.
- Thumbs UP if reasonable. Thumbs DOWN if not.

DO:
- Display the calculation.
- 15 seconds for thinking.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP if reasonable. Thumbs DOWN if not."
- Scan for: thumbs DOWN. 7 plus 3 is 10, not 20.
PROCEED: If 80% thumbs down, click to reveal.
PIVOT: Most likely misconception - students try to calculate exactly and get stuck.
- Reteach: "Estimation is a quick check. Round each number first. 6.8 rounds to 7. 3.1 rounds to 3. 7 + 3 = 10."
- Re-check: "Is 19.9 close to 10?"

TEACHER NOTES:
The point is using estimation as a reasonableness check, not as a perfect calculation. Students who try to calculate the exact answer are not yet using estimation as a tool.

WATCH FOR:
- Confident thumbs down - they used estimation.
- Students who try to calculate exactly - reteach the rounding-first habit.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own. Take the practice sheet.
- Section 1: round to the nearest whole.
- Section 2: round to the nearest tenth.
- Section 3: estimate to check whether an answer is reasonable.

DO:
- Distribute the practice sheet.
- Circulate.
- Watch for students who can use estimation as a check.

TEACHER NOTES:
The practice sheet sequences the three skills: round to whole, round to tenth, then use rounding to check reasonableness. Section 3 is the threshold concept.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed number line. Mark the decimal first, then choose the nearest.
- Extra Notes: Sit with this group. Do the first one together.
EXTENDING PROMPT:
- Task: Round to the nearest hundredth in Section 4. Write one sentence describing the rule.
- Extra Notes: Encourage them to write the rule in their own words.

WATCH FOR:
- Students who underline the next place before rounding - secure.
- Students who chop digits - reteach with the number line.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task.
- Question 1: Round 4.62 to the nearest tenth.
- Question 2: Use estimation to decide. The calculation 3.7 + 2.4 = 11.1. Is this reasonable? Why or why not?

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses SC2 (rounding to a stated place) and SC3 (using estimation for reasonableness).

WATCH FOR:
- Students who write 4.6 with the right reasoning - secure.
- Students who write "not reasonable - 4 + 2 = 6" - they used estimation.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria.
- Show me thumbs for each.
- Turn and tell your partner: why do we estimate before we calculate?

DO:
- Read each I can statement.
- Use thumbs to self-assess.

TEACHER NOTES:
Estimation as a reasonableness check is the big idea. Students who can articulate "to know roughly what to expect" have the relational understanding.

WATCH FOR:
- Students who say "to check the answer" - secure.
- Students unsure - revisit in Lesson 3.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a number line that spans a decimal range. Labels only the endpoints;
// intermediate ticks render as unlabelled ticks. We mark the value of interest
// with addLineMarker() separately so labels never overlap.
function placeDecimalNumberLine(slide, x, y, w, start, end, step, marked, opts) {
  const o = opts || {};
  const count = Math.round((end - start) / step);
  const decimals = step >= 1 ? 0 : (step >= 0.1 ? 1 : 2);
  const labels = [];
  for (let i = 0; i <= count; i += 1) {
    if (i === 0) labels.push(start.toFixed(decimals));
    else if (i === count) labels.push(end.toFixed(decimals));
    else labels.push("");
  }
  const markedIdx = [];
  (marked || []).forEach((m) => {
    const idx = Math.round((m - start) / step);
    if (idx >= 0 && idx <= count) markedIdx.push(idx);
  });
  return addNumberLine(slide, x, y, w, labels, markedIdx, {
    labelFontSize: o.labelFontSize || 14,
  });
}

// Place a free-floating marked decimal (e.g. 4.7 on a 4-to-5 line)
function addLineMarker(slide, lineGeo, value, lineStart, lineEnd, opts) {
  const o = opts || {};
  const ratio = (value - lineStart) / (lineEnd - lineStart);
  const mx = lineGeo.x + ratio * lineGeo.w;
  // Dot
  slide.addShape("roundRect", {
    x: mx - 0.09, y: lineGeo.y - 0.09, w: 0.18, h: 0.18, rectRadius: 0.09,
    fill: { color: o.color || C.ALERT },
  });
  // Label
  slide.addText(String(value), {
    x: mx - 0.4, y: lineGeo.y - 0.50, w: 0.8, h: 0.28,
    fontSize: 14, fontFace: FONT_H, color: o.color || C.ALERT, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Estimation & rounding with decimals",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Find the missing number",
      [
        "7 × n = 56",
        "n ÷ 4 = 12",
        "9 + n = 21",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "n = 8     n = 48     n = 12", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal
  withReveal(
    () => fluencySlide(pres, "Fluency: Adding & Subtracting Decimals",
      ["0.6 + 0.4", "1.5 − 0.7", "2.3 + 1.6"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1.0     0.8     3.9", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: LI/SC
  liSlide(pres,
    "We are learning to round decimals and use estimation to check that our answers make sense.",
    [
      "I can round a decimal to the nearest whole number using a number line.",
      "I can round a decimal to the nearest tenth or hundredth.",
      "I can use estimation to check whether a calculation answer is reasonable.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 8: I Do — round 4.7 to nearest whole
  workedExSlide(pres, 2, "I Do", "Round 4.7 to the nearest whole",
    [
      "Find 4.7 on a number line.",
      "Look at the half-way mark.",
      "",
      "Half-way for 4 to 5 is 4.5.",
      "4.7 sits past 4.5.",
      "",
      "So 4.7 is closer to 5.",
      "4.7 rounds to 5.",
      "",
      "Rule: 5 or more, round up.",
      "Less than 5, round down.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.PRIMARY });
      slide.addText("4.7 on the number line", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      const lineGeo = placeDecimalNumberLine(slide,
        lg.rightX + 0.30, lg.panelTopPadded + 1.20, lg.rightW - 0.60,
        4, 5, 0.1, []);
      // Manually mark half-way (4.5) with a soft tick line
      const halfX = lineGeo.x + 0.5 * lineGeo.w;
      slide.addShape("line", {
        x: halfX, y: lineGeo.y - 0.20, w: 0, h: 0.40,
        line: { color: C.MUTED, width: 1.5, dashType: "dash" },
      });
      slide.addText("4.5", {
        x: halfX - 0.3, y: lineGeo.y + 0.30, w: 0.6, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      addLineMarker(slide, lineGeo, 4.7, 4, 5);

      // Arrow toward 5
      slide.addText("→ rounds to 5", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.20, w: lg.rightW - 0.40, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slide 9: I Do — round 3.46 to nearest tenth
  workedExSlide(pres, 2, "I Do", "Round 3.46 to the nearest tenth",
    [
      "Look at the tenths place — that is the 4.",
      "Look at the NEXT place (hundredths) — that is 6.",
      "",
      "6 is 5 or more, so round UP.",
      "The 4 becomes a 5.",
      "",
      "3.46 rounds to 3.5.",
      "",
      "Rule: look at the next place only.",
      "Decide once.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.SECONDARY });
      slide.addText("3.46 between 3.4 and 3.5", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const lineGeo = placeDecimalNumberLine(slide,
        lg.rightX + 0.30, lg.panelTopPadded + 1.20, lg.rightW - 0.60,
        3.4, 3.5, 0.01, []);
      // Mark half-way 3.45
      const halfX = lineGeo.x + 0.5 * lineGeo.w;
      slide.addShape("line", {
        x: halfX, y: lineGeo.y - 0.20, w: 0, h: 0.40,
        line: { color: C.MUTED, width: 1.5, dashType: "dash" },
      });
      slide.addText("3.45", {
        x: halfX - 0.35, y: lineGeo.y + 0.30, w: 0.7, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      addLineMarker(slide, lineGeo, 3.46, 3.4, 3.5);

      slide.addText("→ rounds to 3.5", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.20, w: lg.rightW - 0.40, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slides 10-11: CFU 1 + reveal — round 6.83 to nearest whole
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Round 6.83 to the nearest whole", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      // Left card: prompt
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Round 6.83 to the", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "nearest whole number.", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Use the number line.", options: { fontSize: 16, color: C.MUTED, italic: true } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right card: number line from 6 to 7
      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("6 to 7", {
        x: 5.3, y: CONTENT_TOP + 0.15, w: 4.2, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      const lineGeo = placeDecimalNumberLine(s,
        5.45, CONTENT_TOP + 1.50, 3.90, 6, 7, 0.1, []);
      const halfX = lineGeo.x + 0.5 * lineGeo.w;
      s.addShape("line", {
        x: halfX, y: lineGeo.y - 0.20, w: 0, h: 0.40,
        line: { color: C.MUTED, width: 1.5, dashType: "dash" },
      });
      s.addText("6.5", {
        x: halfX - 0.3, y: lineGeo.y + 0.30, w: 0.6, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      addLineMarker(s, lineGeo, 6.83, 6, 7);

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "6.83 rounds to 7   (past the 6.5 half-way mark)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: We Do + reveal — round 8.27 to nearest tenth
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Round 8.27 to the nearest tenth",
      [
        "With your partner.",
        "",
        "Step 1: which digit is in the tenths place?",
        "Step 2: which digit is in the next place?",
        "Step 3: 5 or more → round up.",
        "         Less than 5 → round down.",
        "",
        "Write your answer on your whiteboard.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.SECONDARY });
        slide.addText("Number line: 8.2 to 8.3", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        const lineGeo = placeDecimalNumberLine(slide,
          lg.rightX + 0.30, lg.panelTopPadded + 1.20, lg.rightW - 0.60,
          8.2, 8.3, 0.01, []);
        const halfX = lineGeo.x + 0.5 * lineGeo.w;
        slide.addShape("line", {
          x: halfX, y: lineGeo.y - 0.20, w: 0, h: 0.40,
          line: { color: C.MUTED, width: 1.5, dashType: "dash" },
        });
        slide.addText("8.25", {
          x: halfX - 0.35, y: lineGeo.y + 0.30, w: 0.7, h: 0.25,
          fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
          align: "center", margin: 0,
        });
        addLineMarker(slide, lineGeo, 8.27, 8.2, 8.3);

        slide.addText("Your answer: __________", {
          x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.25, w: lg.rightW - 0.40, h: 0.40,
          fontSize: 16, fontFace: FONT_B, color: C.MUTED, italic: true,
          align: "center", valign: "middle", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "8.27 rounds to 8.3   (7 is 5 or more → round up)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 14-15: CFU hinge + reveal — reasonableness of 6.8 + 3.1 = 19.9
  withReveal(
    () => cfuSlide(pres, "CFU", "Reasonable or not?", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n6.8 + 3.1 = 19.9\n\nIs this answer reasonable?\n\nRound each number first.\nThen check.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "NOT reasonable.   7 + 3 = 10, not ~20.   The student moved a decimal point.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 16: You Do — practice sheet
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 1 — round each decimal to the nearest whole.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 2 — round to the nearest tenth.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 3 — estimate to check reasonableness.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Two rules to remember", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "1.  Look at the NEXT place. 5 or more → round up.", {
      x: 1.0, y: panelY + 0.55, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "2.  Estimate by rounding FIRST. Then check the exact answer.", {
      x: 1.0, y: panelY + 1.10, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Round 4.62 to the nearest tenth. Show your reasoning.",
      "A student wrote 3.7 + 2.4 = 11.1. Use estimation. Is this answer reasonable? Why or why not?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why do we estimate before we calculate?",
      scItems: [
        "I can round a decimal to the nearest whole number using a number line.",
        "I can round a decimal to the nearest tenth or hundredth.",
        "I can use estimation to check whether a calculation answer is reasonable.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecFrac_Lesson2_Estimation_And_Rounding.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Round decimals and use estimation to check reasonableness.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Rule: look at the NEXT place after the one you are rounding to. 5 or more rounds up. Less than 5 rounds down.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Round to the nearest WHOLE number", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   4.7    →    _______", y);
    y = addWriteLine(doc, "b)   12.3   →    _______", y);
    y = addWriteLine(doc, "c)   8.51   →    _______", y);
    y = addWriteLine(doc, "d)   0.49   →    _______", y);

    y = addSectionHeading(doc, "Section 2 — Round to the nearest TENTH", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   3.46   →    _______", y);
    y = addWriteLine(doc, "b)   5.82   →    _______", y);
    y = addWriteLine(doc, "c)   0.95   →    _______", y);
    y = addWriteLine(doc, "d)   2.349  →    _______", y);

    y = addSectionHeading(doc, "Section 3 — Use estimation to check reasonableness", y, { color: C.PRIMARY });
    y = addBodyText(doc, "For each, round both numbers to the nearest whole. Estimate the answer. Then decide if the given answer is reasonable.", y);
    y = addWriteLine(doc, "a)  6.8 + 3.1 = 19.9.   Estimate: ___  Reasonable?  YES / NO   Why: ____________", y);
    y = addWriteLine(doc, "b)  9.6 − 2.4 = 7.2.    Estimate: ___  Reasonable?  YES / NO   Why: ____________", y);
    y = addWriteLine(doc, "c)  4.5 + 5.5 = 10.0.   Estimate: ___  Reasonable?  YES / NO   Why: ____________", y);

    y = addSectionHeading(doc, "Section 4 — Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Round to the nearest HUNDREDTH. Then explain the rounding rule in your own words.", y);
    y = addWriteLine(doc, "a)   2.347   →   _______      b)   0.4956   →   _______", y);
    y = addWriteLine(doc, "Rule (in your own words): _______________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Rounding & Estimation Practice | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the rounding and estimation practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Nearest whole", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  5     b)  12     c)  9     d)  0", y);

    y = addSectionHeading(doc, "Section 2 — Nearest tenth", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3.5     b)  5.8     c)  1.0 (or 1)     d)  2.3", y);

    y = addSectionHeading(doc, "Section 3 — Reasonableness", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Estimate ~10 (7 + 3).   19.9 is NOT reasonable - the student moved a decimal point.", y);
    y = addBodyText(doc, "b)  Estimate ~7 (10 - 2).   7.2 IS reasonable.", y);
    y = addBodyText(doc, "c)  Estimate ~10 (5 + 6 or 4 + 6).   10.0 IS reasonable - 4.5 + 5.5 = 10 exactly.", y);

    y = addSectionHeading(doc, "Section 4 — Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  2.35      b)  0.50 (or 0.5)", y);
    y = addBodyText(doc, "Rule: look at the next place after the one you are rounding to. 5 or more rounds up.", y);

    y = addTipBox(doc,
      "Watch for: students who write 4.6 for 4.62 nearest tenth - they need to look at the hundredths to decide. Students who round digit-by-digit on Section 4 should be redirected to look at the next place ONLY.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
