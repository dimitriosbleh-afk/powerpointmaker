"use strict";

// Subtraction Strategies (Year 2 Numeracy) — Lesson 3: Doubles and Near Doubles
// VC2M2N04 / VC2M2A02 — use known doubles to subtract
//   (e.g. 7 + 7 = 14 so 14 - 7 = 7; near double 8 + 7 = 15 so 15 - 7 = 8).
// Daily Review: Measurement - Length (which object is longer? estimate cm).
// Fluency: Addition facts within 10.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Same theme as Lesson 1/2 (variant 0) — unit cohesion.
const T = createTheme("numeracy", "grade2", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  exitTicketSlide, addStageBadge,
  fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  addTensFrame,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 5;
const UNIT_TITLE = "Subtraction Strategies";
const FOOTER = `Subtraction Strategies | Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`;
const OUT_DIR = "output/Subs_Lesson3_Doubles_Near_Doubles";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 3 Doubles Practice",
  "Use a double to find a take away fact.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 3 Answer Key",
  "Worked answers for the doubles practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a double-tens-frame visual to show 7 + 7 = 14 etc.
function drawDoubleFrame(slide, x, y, frameW, top, bottom, opts) {
  const o = opts || {};
  const colorTop = o.colorTop || C.PRIMARY;
  const colorBottom = o.colorBottom || C.SECONDARY;
  const cellW = frameW / 5;
  const cellH = cellW;
  // Top frame with `top` filled
  addTensFrame(slide, x, y,                    frameW, top,    { fillColor: colorTop, cellH });
  // Bottom frame with `bottom` filled
  addTensFrame(slide, x, y + cellH * 2 + 0.10, frameW, bottom, { fillColor: colorBottom, cellH });
}

// Draw a simple ruler-style length comparison.
function drawRulerBar(slide, x, y, w, label, color) {
  slide.addShape("roundRect", {
    x, y, w, h: 0.30, rectRadius: 0.05,
    fill: { color },
    line: { color: C.CHARCOAL, width: 1 },
  });
  // Tick marks
  for (let i = 1; i < Math.round(w / 0.4); i += 1) {
    slide.addShape("line", {
      x: x + i * 0.40, y: y - 0.04, w: 0, h: 0.08,
      line: { color: C.CHARCOAL, width: 1 },
    });
  }
  slide.addText(label, {
    x: x + w + 0.10, y: y - 0.06, w: 1.6, h: 0.42,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
    align: "left", valign: "middle", margin: 0,
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back team. Yesterday we used add facts to take away.
- Today we use a special set of add facts: doubles.
- Doubles are facts where the parts are the same.

DO:
- Have whiteboards, counters, and tens frames ready.

TEACHER NOTES:
Lesson 3 of 5. Doubles are foundational fluency facts. Once students see 7 + 7 = 14, they can use it to find 14 - 7 = 7 instantly. Near doubles extends the idea by one.

WATCH FOR:
- Students who do not yet have doubles facts to 10 - tens frames will help.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the two ribbons on the slide.
- Whisper: which ribbon is longer? About how many centimetres?

DO:
- Display the two ribbon bars (red 12 cm, blue 7 cm).
- Allow 30 seconds for partner whisper.

TEACHER NOTES:
Daily Review revisits length from earlier work. Today we compare two lengths and estimate.

WATCH FOR:
- Students who say the longer one is 10 - prompt them to count tick marks.
- Students who confuse longer and shorter - point at the longer ribbon and gesture wider.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- The red ribbon is longer.
- The red ribbon is about 12 centimetres. The blue ribbon is about 7.
- The red ribbon is about 5 centimetres longer.

DO:
- Click to reveal.
- Tick on whiteboards.

TEACHER NOTES:
Tick and fix. Note any students who guessed the wrong one.

WATCH FOR:
- Students who self-correct on the reveal.
- Students who muddled longer and shorter.

[Stage 1: Daily Review Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency. Add facts to 10.
- I will say a fact. You whisper the answer.
- 6 + 3? 4 + 5? 7 + 2?

DO:
- Brisk choral response. 5 to 7 facts.
- Then doubles to 10: 1+1, 2+2, 3+3, 4+4, 5+5.

TEACHER NOTES:
Today's strategy needs doubles to 10. Drill them briefly here so the I Do feels familiar.

WATCH FOR:
- Students who pause on doubles - they will need extra support today.
- Students who answer instantly - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to use doubles to subtract.
- Read the success criteria together.

DO:
- Choral read the LI.
- Choral read each I can statement.

TEACHER NOTES:
The first I can is achievable for everyone (recalling a known double). The second is the core target. The third stretches students to use a near-double.

WATCH FOR:
- Students who say doubles back fluently - secure.
- Students who pause - allow tens frames as a scaffold today.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. I know that 7 plus 7 is 14.
- That is a double. Both parts are 7.
- Look at the tens frames. Top has 7. Bottom has 7. Total is 14.
- So 14 take away 7 must be 7.

DO:
- Use a double tens frame on the board.
- Place 7 counters on top, 7 on bottom.
- Touch each frame as you say the double.
- Repeat the language at least twice.

TEACHER NOTES:
First model. The double is the anchor. If you know 7 + 7, you know 14 - 7. Use the tens-frame visual so students see the equal parts.

MISCONCEPTIONS:
- Misconception: Students think doubles are only for adding.
  Why: Doubles are usually first taught as an addition fact.
  Impact: They miss the connection to subtraction.
  Quick correction: "Doubles work both ways. Add or take away. The whole stays the same."

WATCH FOR:
- Students who say the double quickly - secure.
- Students who hesitate - count counters together.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. The double is 6 plus 6 is 12.
- What is 12 take away 6?
- Show me on your fingers.

DO:
- Display the double frame: 6 on top, 6 on bottom.
- Allow 10 seconds.
- Say: "Show me... ready... show."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "12 take away 6? Show me on your fingers."
- Scan for: 6 fingers held up.
PROCEED: If 80% show 6, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count back from 12 and lose track.
- Reteach: "Doubles. Both parts are equal. If 6 + 6 is 12, then 12 take 6 is 6."
- Re-check: "Show me 12 take away 6 again."

TEACHER NOTES:
Fingers are universal. Quick, every-student response.

WATCH FOR:
- Students who hold up 6 - using doubles fluently.
- Students who count back - redirect to the doubles fact.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. The near double is 8 plus 7 is 15.
- Notice: 8 plus 7 is one more than 7 plus 7.
- What is 15 take away 7?
- Whisper to your partner. Then write it.

DO:
- Display the double frame: 8 on top, 7 on bottom.
- Allow 30 seconds.
- Walk and scan.

TEACHER NOTES:
We Do steps up to a near double (one more than a double). Listen for "7 + 7 is 14, so 7 + 8 is 15, so 15 take 7 is 8".

WATCH FOR:
- Pairs who use the doubles trick - secure.
- Pairs who count back from 15 - prompt them: "What is 7 + 7? Now plus one more."

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 7 + 7 is 14. So 8 + 7 is 15.
- 15 take away 7 is 8.

DO:
- Click to reveal.
- Repeat the doubles connection with the class.

TEACHER NOTES:
The near double is one more than the double. Make this connection visible on the tens frame.

WATCH FOR:
- Students who said 8 first - secure.
- Students who said 7 - they used the double directly without adjusting; reteach.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2_Q = `SAY:
- Hinge question. Look at this claim.
- I know 5 + 5 = 10. So 10 take away 5 is 5.
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Display the claim.
- Wait 5 seconds.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "10 take away 5 is 5. Thumbs up if yes."
- Scan for: thumbs UP. The doubles reasoning is correct.
PROCEED: If 80% show thumbs up, click to reveal and confirm.
PIVOT: Most likely misconception - students think doubles cannot be reversed.
- Reteach: "5 + 5 is 10. The whole is 10. Take one part of 5, the other part of 5 is left."
- Re-check: "What is 10 take 5?"

TEACHER NOTES:
This hinge probes whether students apply doubles to subtraction.

WATCH FOR:
- Confident thumbs up - using doubles.
- Slow or sideways - they are unsure of the reverse move.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own. The double is 9 + 9 = 18.
- Find 18 take away 9.
- Then try: 9 + 8 = 17. Find 17 take away 9.

DO:
- Hand out the doubles practice sheet (1 page) or use whiteboards.
- Circulate. Listen for "I used the double".
- Cold call 1-2 students to share.

TEACHER NOTES:
Different numbers from We Do. Goal: students apply the doubles strategy independently.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 4 + 4 = 8. Find 8 take away 4.
- Extra Notes: Sit with these students; double tens frame available.
EXTENDING PROMPT:
- Task: Use 10 + 10 = 20 to find 20 take away 9. (Hint: not a double.)
- Extra Notes: Encourage them to explain "20 take 10 is 10, plus 1 is 11".

WATCH FOR:
- Students who write 9 and 8 - secure.
- Students who write the wrong answer for the near double - reteach with the tens frame.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Use the double 8 + 8 = 16. Find 16 take away 8.

DO:
- Display the prompt.
- Allow 60 seconds.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses the core target - using doubles to subtract. The answer is 8.

WATCH FOR:
- Students who write 8 - on track.
- Students who count back - they did not use the double; reteach.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today's strategy: doubles and near doubles.
- Show me thumbs: did you use a doubles fact today?
- Turn and tell your partner: what is your favourite doubles fact?

DO:
- Read the success criteria with the class.
- Use thumbs up sideways down.
- Cold call 1-2 students.

TEACHER NOTES:
Self-assessment data informs Lesson 4 grouping.

WATCH FOR:
- Strong thumbs up - confident.
- Sideways or down - small group focus tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- These are the materials for today.

DO:
- Print the doubles practice sheet (one per student) only if you want a written record.
- Have whiteboards, markers, and double tens frames ready.

TEACHER NOTES:
One printed student resource for this lesson.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 3: Doubles and Near Doubles",
    `Year 2 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 2-3: Daily Review with reveal — length compare
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Which ribbon is longer?", { color: C.ACCENT });

      // Left card: instruction
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Which ribbon is longer?", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "About how many cm?", options: { fontSize: 22, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right card: ruler bars
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      drawRulerBar(s, 5.40, CONTENT_TOP + 1.20, 2.40, "blue ribbon",  C.SECONDARY);
      drawRulerBar(s, 5.40, CONTENT_TOP + 2.10, 3.50, "red ribbon",   C.ALERT);

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Red is longer.   Blue ~ 7 cm.   Red ~ 12 cm.   Difference ~ 5 cm.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — addition facts to 10
  fluencySlide(pres, "Fluency: Add facts to 10",
    ["6 + 3", "4 + 5", "7 + 2"],
    NOTES_FLUENCY, FOOTER);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to use a double to find a take away fact.",
    [
      "I can say a doubles fact (like 5 + 5 = 10).",
      "I can use a double to subtract (14 - 7 = 7).",
      "I can use a near double to subtract (15 - 7 = 8).",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — 7 + 7 = 14, so 14 - 7 = 7
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "If I know 7 + 7 = 14...", { color: STAGE_COLORS["2"] });

    // Top fact strip
    addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["2"] });
    s.addText("7 + 7 = 14   (a double)", {
      x: 0.5, y: CONTENT_TOP + 0.10, w: 9.0, h: 0.65,
      fontSize: 28, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Double tens frame visual on left
    drawDoubleFrame(s, 0.7, CONTENT_TOP + 1.10, 3.20, 7, 7);

    // Right derived facts
    addCard(s, 4.6, CONTENT_TOP + 1.10, 4.9, 2.40, { strip: C.SECONDARY });
    s.addText([
      { text: "Top has 7. Bottom has 7. Total is 14.", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "14 - 7 = 7", options: { fontSize: 30, color: C.SECONDARY, bold: true } },
    ], {
      x: 4.85, y: CONTENT_TOP + 1.30, w: 4.5, h: 2.05,
      fontFace: FONT_B, valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
  })();

  // Slides 7-8: CFU 1 with reveal — 6 + 6 = 12, find 12 - 6
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Use 6 + 6 = 12", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Double tens frame
      drawDoubleFrame(s, 0.9, CONTENT_TOP + 0.30, 3.20, 6, 6);

      s.addText("12 take away 6 = ?", {
        x: 4.6, y: CONTENT_TOP + 0.80, w: 4.9, h: 1.05,
        fontSize: 38, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      s.addText("Show me on your fingers.", {
        x: 4.6, y: CONTENT_TOP + 1.95, w: 4.9, h: 0.50,
        fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "12 take away 6 is 6.   (because 6 + 6 = 12)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 9-10: We Do with reveal — near double 8 + 7 = 15, find 15 - 7
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Near double: 8 + 7 = 15", { color: STAGE_COLORS["3"] });

      // Top reasoning strip
      addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["3"] });
      s.addText("7 + 7 = 14, so 8 + 7 = 15.", {
        x: 0.5, y: CONTENT_TOP + 0.10, w: 9.0, h: 0.65,
        fontSize: 22, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Double frame: top 8, bottom 7
      drawDoubleFrame(s, 0.9, CONTENT_TOP + 1.05, 3.20, 8, 7);

      // Right side — question
      addCard(s, 4.6, CONTENT_TOP + 1.05, 4.9, 2.40, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 18, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "15 take away 7 = ____", options: { fontSize: 26, color: C.CHARCOAL, bold: true } },
      ], {
        x: 4.85, y: CONTENT_TOP + 1.30, w: 4.5, h: 2.00,
        fontFace: FONT_B, valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "15 take away 7 is 8.   (8 + 7 = 15, so 15 - 7 = 8)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 11-12: CFU Hinge with reveal — Is the doubles reasoning right?
  withReveal(
    () => cfuSlide(pres, "CFU", "Is this right?", "Thumbs Up or Thumbs Down",
      "I know 5 + 5 = 10.\n\nSo 10 take away 5 is 5.\n\nIs that right?",
      NOTES_CFU2_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs UP.   Doubles work both ways.   10 - 5 = 5.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: You Do — 9 + 9 = 18 (double) and 9 + 8 = 17 (near double)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn", { color: STAGE_COLORS["4"] });

    // Top steps strip
    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.15, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 22, color: C.ALERT, bold: true } },
      { text: "Use 9 + 9 = 18 to find 18 take away 9.   ", options: { fontSize: 22, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 22, color: C.ALERT, bold: true } },
      { text: "Use the near double 9 + 8 = 17 to find 17 take away 9.", options: { fontSize: 22, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.85,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Lower panel with two double frames side by side
    const panelY = CONTENT_TOP + 1.30;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    drawDoubleFrame(s, 0.85, panelY + 0.40, 3.40, 9, 9);
    s.addText("9 + 9 = 18", {
      x: 0.5, y: panelY + 1.85, w: 4.0, h: 0.36,
      fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    drawDoubleFrame(s, 5.50, panelY + 0.40, 3.40, 9, 8);
    s.addText("9 + 8 = 17  (near double)", {
      x: 5.20, y: panelY + 1.85, w: 4.30, h: 0.36,
      fontSize: 18, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("On your whiteboard. Write both take away facts.", {
      x: 0.5, y: panelY + 2.30, w: 9.0, h: 0.32,
      fontSize: 16, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 14: Exit Ticket
  exitTicketSlide(pres,
    [
      "Use 8 + 8 = 16. Find 16 take away 8.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 15: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: what is your favourite doubles fact?",
      scItems: [
        "I can say a doubles fact (like 5 + 5 = 10).",
        "I can use a double to subtract (14 - 7 = 7).",
        "I can use a near double to subtract (15 - 7 = 8).",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 16: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Subs_Lesson3_Doubles_Near_Doubles.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Doubles practice sheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Use the double to find a take away fact.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addTipBox(doc,
      "If you know a double like 7 + 7 = 14, then 14 take away 7 is 7. The whole stays the same.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Doubles", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Use 5 + 5 = 10.   10 take away 5 = ____", y);
    y = addBodyText(doc, "b)  Use 6 + 6 = 12.   12 take away 6 = ____", y);
    y = addBodyText(doc, "c)  Use 8 + 8 = 16.   16 take away 8 = ____", y);
    y = addSectionHeading(doc, "Near doubles", y, { color: C.PRIMARY });
    y = addBodyText(doc, "d)  Use 6 + 7 = 13.   13 take away 6 = ____", y);
    y = addBodyText(doc, "e)  Use 9 + 8 = 17.   17 take away 9 = ____", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Doubles Practice | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the doubles practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addBodyText(doc, "Doubles work both ways. Add or take away. The whole stays the same.", y);
    y = addSectionHeading(doc, "Doubles", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  10 take away 5 = 5", y);
    y = addBodyText(doc, "b)  12 take away 6 = 6", y);
    y = addBodyText(doc, "c)  16 take away 8 = 8", y);
    y = addSectionHeading(doc, "Near doubles", y, { color: C.PRIMARY });
    y = addBodyText(doc, "d)  13 take away 6 = 7   (because 6 + 7 = 13)", y);
    y = addBodyText(doc, "e)  17 take away 9 = 8   (because 9 + 8 = 17)", y);
    y = addTipBox(doc,
      "If a student writes the wrong answer for a near double, ask: \"What is the matching double? How is this fact different?\"",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
