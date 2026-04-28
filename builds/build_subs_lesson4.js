"use strict";

// Subtraction Strategies (Year 2 Numeracy) — Lesson 4: Bridging Through 10
// VC2M2N04 / VC2M2A02 — split the second number to land on 10 first.
//   Example: 13 - 5  =>  13 - 3 = 10, then 10 - 2 = 8.
// Daily Review: Measurement - Length (compare two pencils with cm).
// Fluency: Number formation — write numerals from 0-9 carefully.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Same theme as Lessons 1-3 (variant 0) — unit cohesion.
const T = createTheme("numeracy", "grade2", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  exitTicketSlide, addStageBadge,
  fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  addNumberTrack,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const TOTAL = 5;
const UNIT_TITLE = "Subtraction Strategies";
const FOOTER = `Subtraction Strategies | Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`;
const OUT_DIR = "output/Subs_Lesson4_Bridging_Through_Ten";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 4 Bridging Through Ten Practice",
  "Land on 10 first, then take the rest.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 4 Answer Key",
  "Worked answers for the bridging through ten practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Number track 0-20 with two coloured arc jumps.
function placeNumberTrack(slide, x, y, w, start, end, highlights, opts) {
  const o = opts || {};
  const count = end - start + 1;
  const cellW = w / count;
  const cellH = o.cellH || 0.50;
  addNumberTrack(slide, x, y, w, start, end, highlights || [], {
    cellH,
    fontSize: o.fontSize || Math.max(11, Math.min(20, Math.round(cellW * 26))),
  });
  return { x, y, w, cellW, cellH, start, end };
}

// Draw a jump arc from cell `fromValue` to cell `toValue`.
function drawTrackArc(slide, geo, fromValue, toValue, opts) {
  const o = opts || {};
  const color = o.color || C.ALERT;
  const fromX = geo.x + (fromValue - geo.start + 0.5) * geo.cellW;
  const toX   = geo.x + (toValue   - geo.start + 0.5) * geo.cellW;
  const baseY = geo.y;
  const arcY  = baseY - (o.arcHeight || 0.35);
  slide.addShape("line", { x: fromX, y: baseY, w: 0, h: arcY - baseY, line: { color, width: 2.4 } });
  slide.addShape("line", { x: fromX, y: arcY, w: toX - fromX, h: 0, line: { color, width: 2.4 } });
  slide.addShape("line", { x: toX, y: arcY, w: 0, h: baseY - arcY, line: { color, width: 2.4 } });
  slide.addShape("line", { x: toX - 0.06, y: baseY - 0.06, w: 0.06, h: 0.06, line: { color, width: 2 } });
  slide.addShape("line", { x: toX + 0.06, y: baseY - 0.06, w: -0.06, h: 0.06, line: { color, width: 2 } });
  if (o.label) {
    const labelW = 0.85;
    const labelX = (fromX + toX) / 2 - labelW / 2;
    slide.addText(String(o.label), {
      x: labelX, y: arcY - 0.32, w: labelW, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }
}

// Draw two pencil-style bars for the Daily Review length compare.
function drawPencilBar(slide, x, y, w, label, color) {
  // Body
  slide.addShape("rect", {
    x, y, w: w - 0.20, h: 0.30,
    fill: { color },
    line: { color: C.CHARCOAL, width: 1 },
  });
  // Tip (triangle)
  slide.addShape("triangle", {
    x: x + w - 0.20, y, w: 0.20, h: 0.30,
    fill: { color: C.CHARCOAL },
    line: { color: C.CHARCOAL, width: 1 },
  });
  slide.addText(label, {
    x: x + w + 0.10, y: y - 0.06, w: 1.6, h: 0.42,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
    align: "left", valign: "middle", margin: 0,
  });
}

// Draw a single big numeral with an arrow showing stroke order — for Fluency.
function drawNumeralCard(slide, x, y, size, num, color) {
  slide.addShape("roundRect", {
    x, y, w: size, h: size, rectRadius: 0.10,
    fill: { color: C.WHITE },
    line: { color, width: 2 },
  });
  slide.addText(String(num), {
    x, y, w: size, h: size,
    fontSize: Math.round(size * 70), fontFace: FONT_H,
    color, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back team. We have learned counting back, think addition and doubles.
- Today's strategy is bridging through 10.
- Land on 10 first, then take the rest.

DO:
- Have whiteboards, counters, and a 0-20 number track ready.

TEACHER NOTES:
Lesson 4 of 5. Bridging through 10 is powerful for harder facts (like 13 - 5). It uses the friendly anchor of 10 to break a big jump into two smaller jumps.

WATCH FOR:
- Students who say "just count back" - prompt them to try the new strategy at least once today.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the two pencils on the slide.
- Whisper: which is longer? About how many cm?

DO:
- Display the two pencil bars.
- Allow 30 seconds for partner whisper.

TEACHER NOTES:
Daily Review continues length. Today we estimate centimetres.

WATCH FOR:
- Students who say the bigger one is "lots" - prompt them to estimate using fingers (1 cm wide).
- Students who confuse longer and shorter.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- The blue pencil is longer.
- The green pencil is about 9 cm. The blue pencil is about 14 cm.
- Difference is about 5 cm.

DO:
- Click to reveal.
- Tick on whiteboards.

TEACHER NOTES:
Tick and fix. Note any students who guessed the wrong one.

WATCH FOR:
- Students who self-correct on the reveal.
- Students who struggle with cm estimation - small group focus tomorrow.

[Stage 1: Daily Review Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Today we focus on number formation.
- Watch the screen. Trace each numeral in the air.
- 6, 7, 8, 9. Top to bottom. Smooth strokes.

DO:
- Display 6 7 8 9 large.
- Lead a finger trace in the air.
- Then ask students to write 6, 7, 8, 9 on whiteboards. Walk and check pencil grip.

TEACHER NOTES:
Number formation is automaticity. Year 2 students often reverse 6 and 9. Quick correction during this routine.

WATCH FOR:
- Students who reverse 6 and 9 - small group support.
- Students who form 7 with two strokes the wrong way - model on the board.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to subtract by bridging through 10.
- Read the success criteria together.

DO:
- Choral read the LI.
- Choral read each I can statement.

TEACHER NOTES:
The first I can is achievable for everyone (just landing on 10). The second is the core target. The third asks students to choose a strategy across the unit.

WATCH FOR:
- Students who repeat back "land on 10" - secure with the language.
- Students unsure - model with the giant number track.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. The problem is 13 take away 5.
- I want to land on 10 first. From 13 I jump back 3. I land on 10.
- I have 5 to take away. I already took 3. I have 2 more to take.
- 10 take away 2 is 8. So 13 take away 5 is 8.

DO:
- Use a giant 0-20 number track on the floor or board.
- Draw the first jump 13 -> 10 (back 3).
- Draw the second jump 10 -> 8 (back 2).
- Repeat the language: jump to 10, then take the rest.

TEACHER NOTES:
First model. The big idea: split 5 into 3 + 2 so we can land on 10. The two arcs make this visible. Use the words land, jump, rest.

MISCONCEPTIONS:
- Misconception: Students take 5 in one jump and land on 8 anyway.
  Why: They are using the counting back strategy from Lesson 1.
  Impact: They miss the bridging strategy.
  Quick correction: "We are practising landing on 10 first. The answer is the same; the path is different."

WATCH FOR:
- Students who whisper the splits - tracking.
- Students who jump 5 in one go - ask them to land on 10 first.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. The problem is 14 take away 5.
- First, jump from 14 to 10. How many back? Show me on your fingers.

DO:
- Display the number track marked at 14.
- Allow 5 seconds.
- Say: "Show me the first jump... ready... show."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "From 14 to 10. How many back? Show me."
- Scan for: 4 fingers held up.
PROCEED: If 80% show 4, click to reveal and confirm.
PIVOT: Most likely misconception - students jump back 5 in one go.
- Reteach: "We land on 10 first. From 14 to 10 is 4 back."
- Re-check: "How many back from 14 to 10?"

TEACHER NOTES:
Just the first jump for now. We Do completes the second jump.

WATCH FOR:
- Students who hold up 4 - secure.
- Students who hold up 5 - they jumped past 10; reteach.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. The problem is 12 take away 5.
- First jump: from 12 to 10. That is 2 back.
- We took 2. We need to take 3 more.
- 10 take away 3 is ____?
- Whisper to your partner. Then write your answer.

DO:
- Display the number track marked at 12 with first jump shown to 10.
- Allow 30 seconds.
- Walk and scan.

TEACHER NOTES:
We Do shares the first jump and asks students to complete the second. This builds confidence with the strategy without overloading them.

WATCH FOR:
- Pairs who say "10 take 3 is 7" - secure.
- Pairs who count back from 10 by ones - acceptable here, but prompt the connection to known facts.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 12 take 2 lands on 10. Then 10 take 3 is 7.
- So 12 take away 5 is 7.

DO:
- Click to reveal both jumps and the final answer.
- Repeat the bridging language.

TEACHER NOTES:
Highlight the 2 + 3 split. The 5 became 2 + 3 because we wanted to land on 10 first.

WATCH FOR:
- Students who said 7 first - secure.
- Students who said 8 - they over-jumped; reteach.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2_Q = `SAY:
- Hinge question. Look at the claim.
- I split 6 into 4 and 2. I jumped from 16 back 4 to land on 10. Then I jumped back 2 more.
- I said the answer is 8.
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Display the work shown.
- Wait 5 seconds.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "16 take away 6 is 8 with the bridging strategy. Thumbs up if yes."
- Scan for: thumbs UP. The reasoning is correct (16 - 6 = 10, 10 - 0 = 10... actually 16 - 4 = 12, then 12 - 2 = 10. Hmm, recheck.)
PROCEED: If 80% show thumbs DOWN, click to reveal.
PIVOT: Most likely misconception - students do not check the splits add to 6.
- Reteach: "The split was 4 + 2 = 6. From 16 back 4 is 12, not 10. The reasoning was wrong."
- Re-check: "What is 16 take away 6 using bridging? Land on 10 first."

TEACHER NOTES:
This hinge probes whether students check that their splits add to the original number. The claim is wrong: 16 - 4 = 12, not 10. Students who think carefully will catch the error.

WATCH FOR:
- Confident thumbs down - they checked the first jump.
- Thumbs up without checking - prompt them to verify the first jump.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own. The problem is 15 take away 7.
- Step 1: Jump from 15 to 10. How many back?
- Step 2: How many more do you need to take?
- Step 3: Take them. Where do you land?

DO:
- Hand out the bridging through ten practice sheet (1 page) or use whiteboards.
- Circulate. Listen for "land on 10".
- Cold call 1-2 students.

TEACHER NOTES:
Different numbers from We Do. Goal: students do both jumps and check that the splits add to 7.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 12 take away 4. (Smaller second jump.)
- Extra Notes: Sit with these students; track available.
EXTENDING PROMPT:
- Task: Use 14 take away 8. Find the splits yourself.
- Extra Notes: Encourage them to record both splits clearly.

WATCH FOR:
- Students who write the splits clearly - secure.
- Students who jump 7 in one go - acceptable answer (8) but prompt them to try the new strategy.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Use bridging through 10. Solve 13 take away 6.

DO:
- Display the prompt.
- Allow 90 seconds.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses the core target - using bridging to subtract. The answer is 7. Splits: 13 - 3 = 10, 10 - 3 = 7.

WATCH FOR:
- Students who show both jumps and write 7 - on track.
- Students who write 7 without showing the bridging - acceptable but ask for the splits next time.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today's strategy: bridging through 10.
- Show me thumbs: did you land on 10 first today?
- Turn and tell your partner: when is bridging useful?

DO:
- Read the success criteria with the class.
- Use thumbs up sideways down.
- Cold call 1-2 students.

TEACHER NOTES:
Self-assessment data informs Lesson 5 grouping.

WATCH FOR:
- Strong thumbs up - confident.
- Sideways or down - small group focus tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- These are the materials for today.

DO:
- Print the bridging through ten practice sheet (one per student) only if you want a written record.
- Have whiteboards, markers, and a 0-20 number track ready.

TEACHER NOTES:
One printed student resource for this lesson.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 4: Bridging Through Ten",
    `Year 2 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slides 2-3: Daily Review with reveal — pencil length compare
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Which pencil is longer?", { color: C.ACCENT });

      // Left card: instruction
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Which pencil is longer?", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "About how many cm?", options: { fontSize: 22, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right card: pencil bars
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      drawPencilBar(s, 5.40, CONTENT_TOP + 1.20, 2.30, "green", C.SUCCESS || C.SECONDARY);
      drawPencilBar(s, 5.40, CONTENT_TOP + 2.10, 3.50, "blue",  C.PRIMARY);

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Blue is longer.   Green ~ 9 cm.   Blue ~ 14 cm.   Difference ~ 5 cm.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — number formation 6 7 8 9
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.ACCENT);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Number formation: 6 7 8 9", { color: C.ACCENT });

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
    const cardY = CONTENT_TOP + 0.30;
    const size = 1.55;
    const gap = 0.40;
    const totalW = size * 4 + gap * 3;
    const startX = (10 - totalW) / 2;
    drawNumeralCard(s, startX + 0 * (size + gap), cardY, size, 6, C.PRIMARY);
    drawNumeralCard(s, startX + 1 * (size + gap), cardY, size, 7, C.SECONDARY);
    drawNumeralCard(s, startX + 2 * (size + gap), cardY, size, 8, C.ACCENT);
    drawNumeralCard(s, startX + 3 * (size + gap), cardY, size, 9, C.ALERT);

    s.addText("Trace each numeral in the air. Then write 6, 7, 8, 9 on your whiteboard.", {
      x: 0.5, y: cardY + size + 0.30, w: 9.0, h: 0.50,
      fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FLUENCY);
  })();

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to subtract by bridging through 10.",
    [
      "I can find the jump that lands on 10.",
      "I can take the rest from 10 to find the answer.",
      "I can choose bridging when it is faster than counting back.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — 13 - 5 by bridging through 10
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "13 take away 5", { color: STAGE_COLORS["2"] });

    // Top reasoning strip
    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["2"] });
    s.addText([
      { text: "Step 1: ", options: { fontSize: 22, color: C.PRIMARY, bold: true } },
      { text: "13 take 3 = 10.   ", options: { fontSize: 22, color: C.CHARCOAL } },
      { text: "Step 2: ", options: { fontSize: 22, color: C.PRIMARY, bold: true } },
      { text: "10 take 2 = 8.", options: { fontSize: 22, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Full-width 0-20 track with two arcs
    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.PRIMARY });
    const trackGeo = placeNumberTrack(s, 0.60, panelY + 1.05, 8.80, 0, 20, [10, 8, 13], { cellH: 0.45 });
    drawTrackArc(s, trackGeo, 13, 10, { label: "-3", arcHeight: 0.55 });
    drawTrackArc(s, trackGeo, 10, 8,  { label: "-2", arcHeight: 0.30 });

    s.addText("13 take away 5 = 8", {
      x: 0.5, y: panelY + 1.85, w: 9.0, h: 0.40,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
  })();

  // Slides 7-8: CFU 1 with reveal — first jump from 14 to 10
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "From 14 to 10", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      placeNumberTrack(s, 0.60, CONTENT_TOP + 1.10, 8.80, 0, 20, [10, 14], { cellH: 0.45 });

      s.addText("How many back from 14 to 10?   Show me on your fingers.", {
        x: 0.5, y: CONTENT_TOP + 1.95, w: 9.0, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "4 back.   14 take 4 = 10.", {
        x: 1.0, y: 4.55, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 9-10: We Do with reveal — 12 - 5 (first jump shown, second jump asked)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "12 take away 5", { color: STAGE_COLORS["3"] });

      // Top reasoning strip — first jump done, asking the second
      addCard(s, 0.5, CONTENT_TOP, 9.0, 1.10, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "Step 1 done: ", options: { fontSize: 20, color: STAGE_COLORS["3"], bold: true } },
        { text: "12 take 2 = 10.   ", options: { fontSize: 20, color: C.CHARCOAL } },
        { text: "Step 2: ", options: { fontSize: 20, color: STAGE_COLORS["3"], bold: true } },
        { text: "10 take 3 = ____", options: { fontSize: 22, color: C.CHARCOAL, bold: true } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.80,
        fontFace: FONT_B, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      // Track with first jump shown
      const panelY = CONTENT_TOP + 1.25;
      addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
      const trackGeo = placeNumberTrack(s, 0.60, panelY + 1.05, 8.80, 0, 20, [10, 12], { cellH: 0.45 });
      drawTrackArc(s, trackGeo, 12, 10, { label: "-2", arcHeight: 0.45 });

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
      addTextOnShape(slide, "10 take 3 = 7.   So 12 take away 5 = 7.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 11-12: CFU Hinge with reveal — Did the splits add to 6?
  withReveal(
    () => cfuSlide(pres, "CFU", "Are the jumps right?", "Thumbs Up or Thumbs Down",
      "I split 6 into 4 and 2.\n\nFrom 16 I jumped back 4 to land on 10.\n\nIs the first jump right?",
      NOTES_CFU2_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN.   16 take 4 = 12, not 10.   Splits must add to 6.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: You Do — 15 - 7 by bridging
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: 15 take away 7", { color: STAGE_COLORS["4"] });

    // Steps strip
    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "Step 1: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Jump back to land on 10.   ", options: { fontSize: 20, color: C.CHARCOAL } },
      { text: "Step 2: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Take the rest from 10.   ", options: { fontSize: 20, color: C.CHARCOAL } },
      { text: "Step 3: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Tell your partner.", options: { fontSize: 20, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Empty track for student work
    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    placeNumberTrack(s, 0.60, panelY + 1.05, 8.80, 0, 20, [10, 15], { cellH: 0.45 });

    s.addText("On your whiteboard. Show both jumps.", {
      x: 0.5, y: panelY + 1.80, w: 9.0, h: 0.36,
      fontSize: 16, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 14: Exit Ticket
  exitTicketSlide(pres,
    [
      "Use bridging through 10. Solve 13 take away 6. Show both jumps.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 15: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: when is bridging through 10 useful?",
      scItems: [
        "I can find the jump that lands on 10.",
        "I can take the rest from 10 to find the answer.",
        "I can choose bridging when it is faster than counting back.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 16: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Subs_Lesson4_Bridging_Through_Ten.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Bridging practice sheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Land on 10 first. Then take the rest.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addTipBox(doc,
      "Step 1: Jump back to 10.   Step 2: Take the rest from 10.   The two jumps add up to the take away.",
      y, { color: C.ACCENT });

    function pdfTrack(yTop) {
      const x = 60, w = 480, count = 21, cellW = w / count, cellH = 26;
      doc.fontSize(11).font("Sans-Bold").fillColor("#333333");
      for (let i = 0; i < count; i += 1) {
        const cx = x + i * cellW;
        doc.lineWidth(1).rect(cx, yTop, cellW, cellH).stroke();
        doc.text(String(i), cx, yTop + 7, { width: cellW, align: "center" });
      }
      doc.font("Sans");
      return yTop + cellH + 6;
    }

    function addQuestion(label, yPos) {
      doc.fontSize(14).fillColor("#" + C.PRIMARY).font("Sans-Bold").text(label, 70, yPos);
      let yy = pdfTrack(yPos + 22);
      doc.fillColor("#333333").font("Sans").fontSize(13)
        .text("Step 1:  ____ - ____ = 10", 90, yy)
        .text("Step 2:  10 - ____ = ____", 90, yy + 22);
      return yy + 56;
    }

    y = addSectionHeading(doc, "Bridging through 10", y, { color: C.PRIMARY });
    y = addQuestion("a)  14 take away 5",  y);
    y = addQuestion("b)  13 take away 6",  y);

    addPdfFooter(doc, `Lesson ${SESSION} | Bridging Through Ten Practice | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the bridging through ten practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addBodyText(doc, "Each question splits the take away into two jumps that add to the original number.", y);
    y = addSectionHeading(doc, "Answers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  14 - 4 = 10, then 10 - 1 = 9.   So 14 take away 5 = 9.   (split 5 = 4 + 1)", y);
    y = addBodyText(doc, "b)  13 - 3 = 10, then 10 - 3 = 7.   So 13 take away 6 = 7.   (split 6 = 3 + 3)", y);
    y = addTipBox(doc,
      "If a student's splits do not add to the take away amount, the answer will be wrong. Always check: do the two jumps add to the original take away?",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
