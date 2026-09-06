"use strict";

// Subtraction Strategies (Year 2 Numeracy) — Lesson 1: Counting Back to Subtract
// VC2M2N04 / VC2M2A02 — count back on a number line / track to subtract within 20.
// Daily Review: 2D Shapes — name shapes by sides and corners.
// Fluency: Patterns — find the next number in a +2 pattern.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Theme: Year 2 numeracy. Variant 0 fixed across all 5 lessons of this unit
// (CLAUDE.md theme cohesion rule: same palette across a unit).
const T = createTheme("numeracy", "grade2", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  addNumberLine, addNumberTrack, addTensFrame,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 5;
const UNIT_TITLE = "Subtraction Strategies";
const FOOTER = `Subtraction Strategies | Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`;
const OUT_DIR = "output/Subs_Lesson1_Counting_Back";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 1 Counting Back Practice",
  "Number line jumps for subtraction within 20.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 1 Answer Key",
  "Worked answers for the counting-back practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a number line jump arc from start -> end with a labelled "-N" cue.
function drawJumpArc(slide, lineGeo, fromValue, toValue, opts) {
  const o = opts || {};
  const intervalW = lineGeo.intervalW;
  const x0 = lineGeo.x + (fromValue - lineGeo.startVal) * intervalW;
  const x1 = lineGeo.x + (toValue - lineGeo.startVal) * intervalW;
  const arcY = lineGeo.y - 0.55;
  const color = o.color || C.ALERT;

  // Curved line approximation via two short segments + a label
  slide.addShape("line", {
    x: x0, y: lineGeo.y, w: 0, h: -0.45,
    line: { color, width: 2.4 },
  });
  slide.addShape("line", {
    x: x0, y: arcY, w: x1 - x0, h: 0,
    line: { color, width: 2.4 },
  });
  slide.addShape("line", {
    x: x1, y: arcY, w: 0, h: 0.45,
    line: { color, width: 2.4 },
  });
  // Arrow tip on the landing tick
  slide.addShape("line", {
    x: x1 - 0.08, y: lineGeo.y - 0.08, w: 0.08, h: 0.08,
    line: { color, width: 2 },
  });
  slide.addShape("line", {
    x: x1 + 0.08, y: lineGeo.y - 0.08, w: -0.08, h: 0.08,
    line: { color, width: 2 },
  });

  // -N label
  const labelText = (o.label != null) ? o.label : `- ${Math.abs(fromValue - toValue)}`;
  const labelW = 0.7;
  const labelX = (x0 + x1) / 2 - labelW / 2;
  slide.addText(labelText, {
    x: labelX, y: arcY - 0.36, w: labelW, h: 0.30,
    fontSize: 16, fontFace: FONT_H, color, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// Number line factory that returns geometry including startVal so we can place jumps.
function placeNumberLine(slide, x, y, w, start, end, marked, opts) {
  const o = opts || {};
  const labels = [];
  for (let v = start; v <= end; v += 1) labels.push(String(v));
  const markedIdx = (marked || []).map((v) => v - start);
  const geo = addNumberLine(slide, x, y, w, labels, markedIdx, {
    labelFontSize: o.labelFontSize || 14,
  });
  return Object.assign({}, geo, { startVal: start, endVal: end });
}

// Full-width number-track with arc jumps — used for I Do, We Do, You Do.
// Returns geometry so callers can layer jumps above the track.
function placeNumberTrack(slide, x, y, w, start, end, highlights, opts) {
  const o = opts || {};
  const count = end - start + 1;
  const cellW = w / count;
  const cellH = o.cellH || 0.55;
  addNumberTrack(slide, x, y, w, start, end, highlights || [], {
    cellH,
    fontSize: o.fontSize || Math.max(14, Math.min(28, Math.round(cellW * 26))),
  });
  return { x, y, w, cellW, cellH, start, end };
}

// Draw a jump arc above a number-track from cell `fromValue` to cell `toValue`.
function drawTrackArc(slide, geo, fromValue, toValue, opts) {
  const o = opts || {};
  const color = o.color || C.ALERT;
  const fromX = geo.x + (fromValue - geo.start + 0.5) * geo.cellW;
  const toX   = geo.x + (toValue   - geo.start + 0.5) * geo.cellW;
  const baseY = geo.y;        // top of track
  const arcY  = baseY - 0.40; // arc peak
  // Vertical from cell-top to arc level on the start side
  slide.addShape("line", { x: fromX, y: baseY, w: 0, h: -0.40, line: { color, width: 2.4 } });
  // Horizontal arc top
  slide.addShape("line", { x: fromX, y: arcY, w: toX - fromX, h: 0, line: { color, width: 2.4 } });
  // Vertical down to landing cell
  slide.addShape("line", { x: toX, y: arcY, w: 0, h: 0.40, line: { color, width: 2.4 } });
  // Tiny arrowhead on landing tick
  slide.addShape("line", { x: toX - 0.06, y: baseY - 0.06, w: 0.06, h: 0.06, line: { color, width: 2 } });
  slide.addShape("line", { x: toX + 0.06, y: baseY - 0.06, w: -0.06, h: 0.06, line: { color, width: 2 } });
  // Label
  if (o.label) {
    const labelW = 0.70;
    const labelX = (fromX + toX) / 2 - labelW / 2;
    slide.addText(String(o.label), {
      x: labelX, y: arcY - 0.34, w: labelW, h: 0.30,
      fontSize: 14, fontFace: FONT_H, color, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }
}

// Draw a 2D shape mockup (square / triangle / rectangle / circle) with a label.
function drawShape(slide, kind, x, y, size, opts) {
  const o = opts || {};
  const color = o.color || C.PRIMARY;
  if (kind === "circle") {
    slide.addShape("roundRect", {
      x, y, w: size, h: size, rectRadius: size / 2,
      fill: { color }, line: { color: C.CHARCOAL, width: 1.2 },
    });
  } else if (kind === "rectangle") {
    slide.addShape("rect", {
      x, y, w: size * 1.5, h: size * 0.8,
      fill: { color }, line: { color: C.CHARCOAL, width: 1.2 },
    });
  } else if (kind === "triangle") {
    slide.addShape("triangle", {
      x, y, w: size, h: size,
      fill: { color }, line: { color: C.CHARCOAL, width: 1.2 },
    });
  } else {
    // square
    slide.addShape("rect", {
      x, y, w: size, h: size,
      fill: { color }, line: { color: C.CHARCOAL, width: 1.2 },
    });
  }
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Good morning team. Today we start a brand new unit.
- We are learning to subtract using clever strategies.
- Today's strategy is counting back.

DO:
- Have student whiteboards, markers and a number line track ready on the floor.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 1 of 5 in the Subtraction Strategies unit. Counting back is the foundation strategy. Every other strategy will build on this idea.

WATCH FOR:
- Students unsure of what subtract means - revisit "take away" briefly in I Do.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up time. Look at the shape on the slide.
- Whisper to your partner: how many sides? How many corners?

DO:
- Display the four shapes.
- Allow 30 seconds for partner whisper.
- Walk and listen to language.

TEACHER NOTES:
Daily Review revisits 2D shapes from earlier in Term 1. Listen for the words sides and corners. This is prior learning, not today's focus.

WATCH FOR:
- Students who count edges twice - prompt them to touch each side once.
- Students who confuse sides and corners - point to a side, then a corner.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- A square has 4 sides and 4 corners.
- A triangle has 3 sides and 3 corners.
- A rectangle has 4 sides and 4 corners.
- A circle has 0 corners.

DO:
- Click to reveal the answers.
- Tick on whiteboards for any students who wrote them.

TEACHER NOTES:
Tick and fix. Note any students who are still mixing up sides and corners for a small group later.

WATCH FOR:
- Students who say a square has 4 corners - secure.
- Students who give 8 for a square - they doubled by counting corners and sides.

[Stage 1: Daily Review Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Look at the pattern.
- The numbers go up by 2 each time.
- 2... 4... 6... 8... what comes next?
- Whisper the next number.

DO:
- Display 2, 4, 6, 8, ___ pattern.
- Choral skip count from 2 to 20 with the class twice.
- Brisk pace, low stakes.

TEACHER NOTES:
This is automaticity in additive patterns - prior learning that supports today's counting back work. Do NOT new-teach the strategy here.

WATCH FOR:
- Students who say 9 instead of 10 - they reverted to counting by 1s.
- Students who count by 2s smoothly - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to subtract by counting back on a number line.
- Let us read the success criteria together.

DO:
- Choral read the LI.
- Choral read each I can statement.
- Hold up the number line and point at the start tick.

TEACHER NOTES:
The first I can is achievable for everyone today - just placing a finger on the start number. The second is the core target. The third stretches students to explain the process aloud.

WATCH FOR:
- Students who repeat back the language - pattern is forming.
- Students who look unsure - the I Do will build the picture.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. The problem is 8 take away 3.
- I start at 8. I put my finger on 8.
- Now I count back 3 jumps. 7... 6... 5.
- I land on 5. So 8 take away 3 is 5.
- Say it with me: 8 take away 3 is 5.

DO:
- Use a giant number line on the floor or board.
- Point at 8.
- Count back 3 with a clear gesture for each jump.
- Repeat the language pattern at least twice.

TEACHER NOTES:
First model. Move slowly. The number line is the new tool. Say the language pattern at least three times. Use the words start, jump back, land.

MISCONCEPTIONS:
- Misconception: Students count the start number as the first jump.
  Why: Counting all is their default strategy.
  Impact: They land on 6 instead of 5 for 8 take away 3.
  Quick correction: "Start at 8. Do not count 8. Jump first to 7."

WATCH FOR:
- Students who whisper the count along with me - tracking.
- Students who do not move their fingers - cue them to follow on their printed line.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. The problem is 7 take away 2.
- Start at 7. Count back 2 jumps.
- Show me on your fingers where you land.

DO:
- Display the number line marked at 7.
- Allow 10 seconds for thinking.
- Say: "Show me on your fingers... ready... show."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "Where do you land? Show me on your fingers... show."
- Scan for: 5 fingers held up.
PROCEED: If 80% show 5 fingers, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count 7 as the first jump and land on 6.
- Reteach: "Touch 7. Do not count 7. Jump first to 6."
- Re-check: "Now where do you land for 7 take away 2?"

TEACHER NOTES:
Fingers are universal in Year 2. Quick, every-student response.

WATCH FOR:
- Students who hold up 5 quickly - secure.
- Students who hold up 6 - counted the start as a jump.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. The problem is 9 take away 4.
- Start at 9. Count back 4 jumps.
- Whisper to your partner: where do you land?
- Now write your answer on your whiteboard.

DO:
- Display the number line at 9 with no jumps shown.
- Allow 30 seconds for partner whisper and whiteboard write.
- Walk and scan.

TEACHER NOTES:
We Do mirrors I Do but uses different numbers - same strategy. Listen for the language pattern start, jump, land.

WATCH FOR:
- Pairs who say the language pattern aloud - secure.
- Pairs who land on 6 instead of 5 - they counted the start; redirect.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together.
- 9 take away 4 is 5. We start at 9 and jump back 4. We land on 5.
- Tick on your whiteboard if you got 5.

DO:
- Click to reveal the jumps and the answer.
- Repeat the language pattern with the class.

TEACHER NOTES:
Some students will land on 6. That tells us they counted the start. Redirect on the spot if more than two students show 6.

WATCH FOR:
- Students who self-correct on the reveal - they noticed the error.
- Students who keep 6 - small group support tomorrow.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2_Q = `SAY:
- Hinge question. Look at this picture.
- I started at 6 and counted back 2 jumps.
- I landed on 5.
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Display the number line showing a 6 -> 5 -> 4 jump path but with the answer claim "I landed on 5".
- Wait 5 seconds.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "I said I landed on 5. Thumbs up if yes. Thumbs down if no."
- Scan for: thumbs DOWN. The correct answer is 4.
PROCEED: If 80% show thumbs down, click to reveal and confirm.
PIVOT: Most likely misconception - students count the start tick as one jump.
- Reteach: "Start at 6. Do not count 6. First jump goes to 5. Second jump goes to 4."
- Re-check: "Where do you really land?"

TEACHER NOTES:
This hinge probes whether students count the start tick as a jump. A "no" answer means they are tracking the strategy correctly.

WATCH FOR:
- Confident thumbs down - students see the off-by-one mistake.
- Slow or sideways thumbs - they need another model.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- The problem is 10 take away 4.
- Start at 10. Count back 4 jumps. Show your jumps on the line.
- Tell your partner the answer.

DO:
- Hand out a printed mini number line and whiteboard markers.
- Circulate and listen for the language start, jump, land.
- Cold call 1-2 students to share their answer.

TEACHER NOTES:
Different numbers from We Do but the same strategy. Goal: students show the jumps clearly and say the language.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 8 take away 2 instead. The number line still starts at 0.
- Extra Notes: Sit with these students; do the first jump together.
EXTENDING PROMPT:
- Task: Try 10 take away 7. How many jumps?
- Extra Notes: Encourage them to predict the answer before drawing.

WATCH FOR:
- Students who land on 6 - secure.
- Students who land on 7 - counted the start; reteach with the giant number line.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Solve 8 take away 5. Show your jumps on the number line.

DO:
- Display the prompt.
- Allow 60 seconds.
- Collect whiteboards or take a quick photo.

TEACHER NOTES:
Exit ticket assesses the core target - using counting back to subtract. The answer is 3.

WATCH FOR:
- Students who write 3 with clear jumps - on track.
- Students who write 4 - counted the start; small group focus tomorrow.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- We learned a clever strategy today: counting back.
- Show me thumbs up if you can count back to subtract.
- Turn and tell your partner: where do we put our finger first?

DO:
- Read the success criteria with the class.
- Use thumbs up sideways down.
- Cold call 1-2 students for the partner share.

TEACHER NOTES:
Self-assessment data informs Lesson 2 grouping. Most students should reach the first I can today; the strategy will be used again in Lessons 2-5.

WATCH FOR:
- Strong thumbs up - they can use the strategy.
- Sideways or down - small group focus tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- These are the materials for today.
- The counting back sheet is for table time after the lesson.

DO:
- Print the counting back sheet (one per student).
- Have whiteboards, markers, and a giant floor number line ready.

TEACHER NOTES:
One printed student resource for this lesson. Most of the work happens on whiteboards and the floor number line.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Counting Back to Subtract",
    `Year 2 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 2-3: Daily Review with reveal — 2D Shapes (sides and corners)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Sides and Corners", { color: C.ACCENT });

      // Left card: instruction
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "How many sides?", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "How many corners?", options: { fontSize: 22, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right card: four shapes
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      const shapeY = CONTENT_TOP + 0.40;
      drawShape(s, "square",    5.55, shapeY,        1.10, { color: C.PRIMARY });
      drawShape(s, "triangle",  7.15, shapeY,        1.10, { color: C.SECONDARY });
      drawShape(s, "rectangle", 5.55, shapeY + 1.55, 1.10, { color: C.ACCENT });
      drawShape(s, "circle",    7.40, shapeY + 1.55, 1.10, { color: C.ALERT });

      // Shape labels
      s.addText("square",    { x: 5.40, y: shapeY + 1.20, w: 1.40, h: 0.30, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0, bold: true });
      s.addText("triangle",  { x: 7.00, y: shapeY + 1.20, w: 1.40, h: 0.30, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0, bold: true });
      s.addText("rectangle", { x: 5.30, y: shapeY + 2.50, w: 1.65, h: 0.30, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0, bold: true });
      s.addText("circle",    { x: 7.20, y: shapeY + 2.50, w: 1.40, h: 0.30, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0, bold: true });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "square 4 + 4   triangle 3 + 3   rectangle 4 + 4   circle 0 corners", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Patterns (counting by 2s)
  fluencySlide(pres, "Fluency: What comes next?",
    ["2, 4, 6, 8, ?", "10, 12, 14, ?"],
    NOTES_FLUENCY, FOOTER);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to subtract by counting back on a number line.",
    [
      "I can find the start number on the number line.",
      "I can count back the right number of jumps.",
      "I can say the answer using the words start, jump, land.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — 8 take away 3 = 5 (stacked layout: steps top, full-width track bottom)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "8 take away 3", { color: STAGE_COLORS["2"] });

    // Top steps card — narrative the teacher uses while modelling
    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["2"] });
    s.addText([
      { text: "Start at 8.   ", options: { fontSize: 26, color: C.PRIMARY, bold: true } },
      { text: "Count back 3 jumps:  7... 6... 5.   ", options: { fontSize: 22, color: C.CHARCOAL } },
      { text: "Land on 5.", options: { fontSize: 26, color: C.PRIMARY, bold: true } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.20, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Bottom track panel — full slide width
    const panelY = CONTENT_TOP + 1.45;
    const panelH = SAFE_BOTTOM - panelY;
    addCard(s, 0.5, panelY, 9.0, panelH, { strip: C.PRIMARY });

    // Number track 0-10 inside the panel; arcs sit above (within panel)
    const trackY = panelY + 0.95;
    const trackGeo = placeNumberTrack(s, 0.80, trackY, 8.40, 0, 10, [5, 8], { cellH: 0.55 });

    // Three back-jumps: 8 -> 7, 7 -> 6, 6 -> 5
    drawTrackArc(s, trackGeo, 8, 7, { label: "-1" });
    drawTrackArc(s, trackGeo, 7, 6, { label: "-1" });
    drawTrackArc(s, trackGeo, 6, 5, { label: "-1" });

    // Answer line below the track
    s.addText("8 take away 3 = 5", {
      x: 0.5, y: trackY + 0.70, w: 9.0, h: 0.40,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
  })();

  // Slides 7-8: CFU 1 with reveal — 7 take away 2 = 5
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "7 take away 2", { color: C.ALERT });

      // CHECK stamp top-right
      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      // Hero card
      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // Full-width number track 0-10 marked at 7
      const trackGeo = placeNumberTrack(s, 0.80, CONTENT_TOP + 1.10, 8.40, 0, 10, [7], { cellH: 0.55 });

      s.addText("Start at 7. Count back 2.   Show me on your fingers where you land.", {
        x: 0.7, y: trackGeo.y + 0.85, w: 8.6, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "7 take away 2 is 5.", {
        x: 1.0, y: 4.55, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 9-10: We Do with reveal — 9 take away 4 = 5
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "9 take away 4", { color: STAGE_COLORS["3"] });

      // Top prompt card
      addCard(s, 0.5, CONTENT_TOP, 9.0, 1.20, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "Start at 9.   ", options: { fontSize: 24, color: C.PRIMARY, bold: true } },
        { text: "Count back 4 jumps.   ", options: { fontSize: 22, color: C.CHARCOAL } },
        { text: "Land on ____", options: { fontSize: 26, color: C.CHARCOAL, bold: true } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.85,
        fontFace: FONT_B, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      // Bottom track panel
      const panelY = CONTENT_TOP + 1.35;
      addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
      placeNumberTrack(s, 0.80, panelY + 0.95, 8.40, 0, 10, [9], { cellH: 0.55 });

      // Whiteboard cue
      s.addText("On your whiteboard.", {
        x: 0.5, y: panelY + 1.85, w: 9.0, h: 0.36,
        fontSize: 16, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", valign: "middle", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "9 take away 4 is 5.   Jumps: 9 -> 8 -> 7 -> 6 -> 5.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 11-12: CFU Hinge with reveal — Did I land on 5?
  withReveal(
    () => cfuSlide(pres, "CFU", "Did I land on 5?", "Thumbs Up or Thumbs Down",
      "I started at 6.\nI counted back 2 jumps.\n\nI said I landed on 5.\n\nIs that right?",
      NOTES_CFU2_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN.   6 take away 2 is 4.", {
        x: 1.0, y: 4.55, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: You Do — 10 take away 4 (different content from We Do)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: 10 take away 4", { color: STAGE_COLORS["4"] });

    // Top steps strip
    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.20, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 22, color: C.ALERT, bold: true } },
      { text: "Find 10 on the line.   ", options: { fontSize: 22, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 22, color: C.ALERT, bold: true } },
      { text: "Count back 4 jumps.   ", options: { fontSize: 22, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 22, color: C.ALERT, bold: true } },
      { text: "Tell your partner.", options: { fontSize: 22, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.85,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Bottom track panel
    const panelY = CONTENT_TOP + 1.35;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    placeNumberTrack(s, 0.80, panelY + 0.85, 8.40, 0, 10, [10], { cellH: 0.55 });

    // Whiteboard cue at bottom of panel
    s.addText("On your whiteboard. Show your jumps.", {
      x: 0.5, y: panelY + 1.75, w: 9.0, h: 0.36,
      fontSize: 16, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 14: Exit Ticket
  exitTicketSlide(pres,
    [
      "Solve 8 take away 5. Show your jumps on the number line.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 15: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: where do we put our finger first?",
      scItems: [
        "I can find the start number on the number line.",
        "I can count back the right number of jumps.",
        "I can say the answer using the words start, jump, land.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 16: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Subs_Lesson1_Counting_Back.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Counting back practice worksheet (DOCX would be primary but PDF is what we have)
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Solve each one by counting back on the number line. Show your jumps.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addTipBox(doc,
      "Start at the first number. Count back. Land on the answer. Use the words start, jump, land.",
      y, { color: C.ACCENT });

    // Helper to draw a number line in PDF
    function drawPdfNumberLine(x, yTop, w, start, end, marked) {
      const count = end - start;
      const stepW = w / count;
      doc.lineWidth(2).moveTo(x, yTop).lineTo(x + w, yTop).strokeColor("#333333").stroke();
      // Arrows
      doc.moveTo(x - 8, yTop - 6).lineTo(x, yTop).lineTo(x - 8, yTop + 6).stroke();
      doc.moveTo(x + w + 8, yTop - 6).lineTo(x + w, yTop).lineTo(x + w + 8, yTop + 6).stroke();
      for (let i = 0; i <= count; i += 1) {
        const tx = x + i * stepW;
        const isMarked = (marked || []).indexOf(start + i) >= 0;
        doc.lineWidth(2).moveTo(tx, yTop - 7).lineTo(tx, yTop + 7).stroke();
        if (isMarked) {
          doc.circle(tx, yTop, 5).fillColor("#" + C.ALERT).fill();
        }
        doc.fillColor("#333333").fontSize(11).font("Sans-Bold").text(String(start + i), tx - 6, yTop + 12, { width: 14, align: "center" });
      }
      doc.font("Sans");
    }

    function addProblem(label, start, mark, yPos) {
      doc.fontSize(14).fillColor("#" + C.PRIMARY).font("Sans-Bold").text(label, 70, yPos);
      drawPdfNumberLine(80, yPos + 36, 460, 0, 10, [mark]);
      doc.fillColor("#333333").font("Sans").fontSize(12).text("Answer: ____", 70, yPos + 70);
      return yPos + 110;
    }

    y = addSectionHeading(doc, "Practice", y, { color: C.PRIMARY });
    y = addProblem("a)  9 take away 3 = ?",  9, 9, y);
    y = addProblem("b)  7 take away 4 = ?",  7, 7, y);
    y = addProblem("c)  10 take away 6 = ?", 10, 10, y);

    addPdfFooter(doc, `Lesson ${SESSION} | Counting Back Practice | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the counting back practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addBodyText(doc, "Each problem is solved by starting at the first number and counting back the right number of jumps.", y);
    y = addSectionHeading(doc, "Answers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  9 take away 3 = 6     (jumps: 9 -> 8 -> 7 -> 6)", y);
    y = addBodyText(doc, "b)  7 take away 4 = 3     (jumps: 7 -> 6 -> 5 -> 4 -> 3)", y);
    y = addBodyText(doc, "c)  10 take away 6 = 4    (jumps: 10 -> 9 -> 8 -> 7 -> 6 -> 5 -> 4)", y);
    y = addTipBox(doc,
      "If a student counts the start number as a jump, they will land one number too high. Reteach: start at the first number, do not count it, jump first.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
