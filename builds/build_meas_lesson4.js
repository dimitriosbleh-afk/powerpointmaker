"use strict";

// Measurement Unit (Year 5/6 Numeracy) — Lesson 4: Converting mass and capacity.
// Year 6 content. Same staircase idea as Lesson 3, applied to mass and capacity.
// Daily Review: solving equations with x/÷.
// Fluency: subtraction.

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

const SESSION = 4;
const TOTAL = 7;
const UNIT_TITLE = "Measurement: Units, Time and Timetables";
const FOOTER = `Measurement | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/Meas_Lesson4_Converting_Mass_Capacity";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 4 Mass and Capacity Conversions",
  "Convert between mg, g, kg, t and mL, L, kL. Use decimals where needed.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 4 Answer Key",
  "Worked answers for the Mass and Capacity Conversions sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back. Yesterday we converted between length units.
- Today we use the same rules for mass and capacity.

DO:
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 4 of 7. Year 6 content. Same staircase logic as Lesson 3, but each jump is x1000.

WATCH FOR:
- Students who recall yesterday's "up = multiply, down = divide" - great carryover.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- You will need your whiteboard and the Mass and Capacity Conversions sheet.

DO:
- Print one sheet per student.
- Print one answer key.

TEACHER NOTES:
One student resource plus answer key. The staircase visual is on the slides.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Solve each equation.
- The unknown is x.
- On your whiteboard, write what x equals.

DO:
- Display the three prompts.
- 2 minutes.

TEACHER NOTES:
Solving equations with multiplication and division reviews prior learning. The strategy is inverse operations.

WATCH FOR:
- Students who use the inverse - secure.
- Students who guess and check - prompt them to "undo" the operation.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 4 x x = 32. So x = 32 ÷ 4 = 8.
- x ÷ 5 = 7. So x = 7 x 5 = 35.
- 6 x x = 54. So x = 54 ÷ 6 = 9.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The key idea: do the inverse operation to both sides. Note students who still guess.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Subtraction facts.
- Quick whisper-answer then write on your board.

DO:
- Display the three prompts.
- 30 seconds.

TEACHER NOTES:
Continuing subtraction fluency. Watch for students who use known facts vs counting back.

WATCH FOR:
- Students who recall - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 16 - 9 = 7.
- 12 - 5 = 7.
- 14 - 8 = 6.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Look for the "think addition" strategy - 16 - 9: what plus 9 makes 16?

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- A quick reminder. Same rule today: convert means write the same amount in a different unit.
- 2 kg and 2000 g are the same mass.
- 1.5 L and 1500 mL are the same capacity.

DO:
- Hold up a 1 L water bottle. Say "1 L = 1000 mL".
- Hold up a 1 kg textbook. Say "1 kg = 1000 g".

TEACHER NOTES:
Reinforce the equivalence. The number changes, the amount does not.

WATCH FOR:
- Students who can say "the amount is the same" - tracking.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to convert between metric units of mass and capacity.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 is single-step conversions in the easier direction. SC2 is both directions. SC3 stretches to decimals (1.5 L = 1500 mL).

WATCH FOR:
- Students who can repeat the SC - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. Mass and capacity each have their own staircase.
- The rule is the same as yesterday.
- UP the stairs (big unit to small unit): MULTIPLY.
- DOWN the stairs (small unit to big unit): DIVIDE.
- The only difference: each jump here is x1000.
- 1 g = 1000 mg.   1 kg = 1000 g.   1 t = 1000 kg.
- 1 L = 1000 mL.   1 kL = 1000 L.

DO:
- Point at each step on the staircase.
- Trace "up" and "down" with your hand.
- Repeat "x1000" three times.

TEACHER NOTES:
Same staircase idea, bigger jumps. Remind students that x1000 means the decimal moves THREE places (not two).

MISCONCEPTIONS:
- Misconception: Students use x100 for kg to g (transfer from length).
  Why: They mix the two staircases.
  Impact: Answers off by a factor of 10.
  Quick correction: "Mass and capacity jumps are x1000, not x100. Three zeros."

WATCH FOR:
- Students who nod at "three places" - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Watch me convert 1.5 L into mL.
- Direction: big to small (L to mL).
- Going up the stairs - MULTIPLY by 1000.
- 1.5 x 1000 = 1500. The decimal moves three places to the right.
- 1.5 L = 1500 mL.
- Check: more mL than L? Yes. Makes sense.

DO:
- Write "1.5 L" on the board.
- Trace the up arrow with your finger.
- Move the decimal point three places as you say "x 1000".
- Write "= 1500 mL".

TEACHER NOTES:
Second worked example. Hammer the decimal movement - x1000 = three places.

WATCH FOR:
- Students who move the decimal three places - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- Convert 2.3 kg into g.
- On your whiteboard. Direction? Multiply or divide? By how much?

DO:
- Display the prompt.
- 45 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your boards on three, two, one, show."
- Scan for: 2300 g.
PROCEED: If 80% have 2300, click to reveal and move to We Do.
PIVOT: Most likely misconception - students use x100 instead of x1000.
- Reteach: "kg to g is one jump, x1000. Decimal moves three places. 2.300 becomes 2300."
- Re-check: "Convert 4.5 kg to g."

TEACHER NOTES:
Probe whether students apply x1000 (not x100).

WATCH FOR:
- Students who write 2300 - secure.
- Students who write 230 - used x100. Reteach.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me.
- Four conversions. Two mass, two capacity. Both directions.
- Partner talk for each: direction, multiplier, answer.

DO:
- Display the four conversions.
- 3 minutes total.
- Cold call after each.

TEACHER NOTES:
Mix of directions. Watch for x100 errors (length transfer).

WATCH FOR:
- Pairs who say "up - multiply by 1000" - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together.
- 3 kg = 3000 g (up, x1000).
- 4500 mL = 4.5 L (down, ÷1000).
- 2 L = 2000 mL (up, x1000).
- 6000 g = 6 kg (down, ÷1000).

DO:
- Click to reveal.
- Run direction + operation for each.

TEACHER NOTES:
Reveal confirms each move.

WATCH FOR:
- Students who got three of four - secure.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Two students disagree about 750 g in kg.
- Sam says 750 g = 7.5 kg. He divided by 100.
- Eli says 750 g = 0.75 kg. She divided by 1000.
- Thumbs UP for Sam. Thumbs DOWN for Eli.

DO:
- Display the disagreement.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP for Sam. DOWN for Eli."
- Scan for: thumbs DOWN. Eli is right.
PROCEED: If 80% agree with Eli, click to reveal and confirm.
PIVOT: Most likely misconception - students use the length multiplier (x100) for mass.
- Reteach: "Mass jumps are x1000. 750 g must be less than 1 kg because 1 kg = 1000 g."
- Re-check: "Convert 450 g to kg."

TEACHER NOTES:
The hinge tests the "x1000 for mass" rule. Check the SENSE - 750 g is less than 1 kg.

WATCH FOR:
- Confident thumbs DOWN for Sam - they applied x1000.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the Mass and Capacity Conversions sheet.
- Use the staircases. Check: more or fewer?

DO:
- Distribute the sheet.
- Circulate.

TEACHER NOTES:
Different numbers from the We Do. Same rule.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use Part A only. Sit with these students. Talk through direction first.
- Extra Notes: Keep the staircase visible.
EXTENDING PROMPT:
- Task: Part D (mixed-unit problems: 1 kg 250 g into g; 2 L 350 mL into mL).
- Extra Notes: Encourage students to convert the larger unit first, then add.

WATCH FOR:
- Students who use direction language - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Convert 3.2 kg into g.
- Convert 2400 mL into L.

DO:
- Display the prompt.
- 3 minutes.

TEACHER NOTES:
Exit ticket assesses SC2. Look for 3200 g and 2.4 L.

WATCH FOR:
- Students who get both - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: how is mass/capacity different from length when converting?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea is "each mass/capacity jump is x1000". Tomorrow we move to 12 and 24 hour time.

WATCH FOR:
- Strong thumbs up - move at pace tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

// Mass/capacity staircase. Each jump is x1000.
// Label and multiplier heights are sized from stepH so they fit in compact staircases.
function drawStaircase(slide, x, y, w, h, units, multiplier) {
  const stepCount = units.length;
  const stepW = w / stepCount;
  const stepH = h / stepCount;
  // Label fills ~70% of step height; multiplier sits below.
  const labelH = Math.max(0.18, Math.min(0.40, stepH * 0.62));
  const labelFont = Math.max(10, Math.min(22, stepH * 36));

  units.forEach((u, i) => {
    const sx = x + i * stepW;
    const sy = y + (stepCount - 1 - i) * stepH;
    slide.addShape("rect", {
      x: sx, y: sy, w: stepW, h: h - (stepCount - 1 - i) * stepH,
      fill: { color: u.color }, line: { color: C.CHARCOAL, width: 1 },
    });
    slide.addText(u.code, {
      x: sx, y: sy + 0.02, w: stepW, h: labelH,
      fontSize: labelFont, fontFace: FONT_H, color: C.WHITE,
      bold: true, align: "center", margin: 0,
    });
  });

  // Multipliers inside each step (below the code label, top portion only).
  for (let i = 0; i < stepCount - 1; i++) {
    const jx = x + (i + 0.5) * stepW;
    const jy = y + (stepCount - 1 - i) * stepH + labelH + 0.02;
    slide.addText(multiplier, {
      x: jx - 0.40, y: jy, w: 0.80, h: 0.20,
      fontSize: 10, fontFace: FONT_B, color: C.WHITE,
      bold: true, align: "center", margin: 0,
    });
  }
}

const MASS = [
  { code: "mg", color: C.PRIMARY },
  { code: "g",  color: C.SECONDARY },
  { code: "kg", color: C.ACCENT },
  { code: "t",  color: C.ALERT },
];
const CAP = [
  { code: "mL", color: C.PRIMARY },
  { code: "L",  color: C.ACCENT },
  { code: "kL", color: C.ALERT },
];

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 4: Converting mass and capacity",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal — solving equations
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Solving equations",
      ["4 × x = 32", "x ÷ 5 = 7", "6 × x = 54"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "x = 8         x = 35         x = 9", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal — subtraction
  withReveal(
    () => fluencySlide(pres, "Fluency: Subtraction",
      ["16 - 9", "12 - 5", "14 - 8"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "7     7     6", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Key vocabulary — same amount, different unit
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Same amount, different unit");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    addTextOnShape(s, "1 kg  =  1000 g", {
      x: 1.0, y: CONTENT_TOP + 0.30, w: 8.0, h: 0.80, rectRadius: 0.10,
      fill: { color: C.PRIMARY },
    }, { fontSize: 34, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "1.5 L  =  1500 mL", {
      x: 1.0, y: CONTENT_TOP + 1.25, w: 8.0, h: 0.80, rectRadius: 0.10,
      fill: { color: C.SECONDARY },
    }, { fontSize: 34, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "2 t  =  2000 kg", {
      x: 1.0, y: CONTENT_TOP + 2.20, w: 8.0, h: 0.80, rectRadius: 0.10,
      fill: { color: C.ACCENT },
    }, { fontSize: 34, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to convert between metric units of mass and capacity.",
    [
      "I can multiply by 1000 to convert from a bigger unit to a smaller unit (e.g. kg to g).",
      "I can divide by 1000 to convert from a smaller unit to a bigger unit (e.g. mL to L).",
      "I can convert decimals (e.g. 1.5 L = 1500 mL) and check the answer makes sense.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: I Do (1) — mass staircase
  workedExSlide(pres, 2, "I Do", "Mass and capacity staircase",
    [
      "Same as length:",
      "UP = MULTIPLY.",
      "DOWN = DIVIDE.",
      "",
      "BUT each jump is x1000.",
      "",
      "1000 mg = 1 g.",
      "1000 g = 1 kg.",
      "1000 mL = 1 L.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("Mass staircase", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      drawStaircase(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.45,
        lg.rightW - 0.50, 1.20, MASS, "x1000");

      slide.addText("Capacity staircase", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.78, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      drawStaircase(slide, lg.rightX + 0.65, lg.panelTopPadded + 2.15,
        lg.rightW - 1.30, 0.85, CAP, "x1000");
    }
  );

  // Slide 10: I Do (2) — 1.5 L to mL
  workedExSlide(pres, 2, "I Do", "Convert 1.5 L into mL",
    [
      "Direction: big to small (L to mL).",
      "Going UP - MULTIPLY by 1000.",
      "",
      "1.5 x 1000 = 1500.",
      "Move the decimal THREE places right.",
      "",
      "1.5 L = 1500 mL.",
      "",
      "Check: more mL than L? Yes.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("1.5 L  =  1500 mL", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.20, w: lg.rightW - 0.20, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      const moveY = lg.panelTopPadded + 0.95;
      slide.addText("1 . 5 0 0", {
        x: lg.rightX + 0.10, y: moveY, w: lg.rightW - 0.20, h: 0.45,
        fontSize: 26, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", margin: 0,
      });
      slide.addText("x 1000  ->  decimal moves 3 right", {
        x: lg.rightX + 0.10, y: moveY + 0.50, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", margin: 0,
      });
      slide.addText("1 5 0 0 .", {
        x: lg.rightX + 0.10, y: moveY + 0.90, w: lg.rightW - 0.20, h: 0.45,
        fontSize: 26, fontFace: FONT_H, color: C.SECONDARY,
        bold: true, align: "center", margin: 0,
      });

      addTextOnShape(slide, "x1000 = THREE places.", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.55, w: lg.rightW - 0.30, h: 0.45, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 11-12: CFU + reveal — 2.3 kg to g
  withReveal(
    () => cfuSlide(pres, "CFU", "Convert 2.3 kg into g", "Show Me Boards",
      "On your whiteboard:\n\nWrite the conversion of 2.3 kg into g.\n\nDirection?  Multiply or divide?  By how much?",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "2.3 kg = 2300 g   (UP, x1000, decimal moves 3 right)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 13-14: We Do + reveal — four conversions (mass + capacity)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Convert each one with your partner", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      const items = [
        { from: "3 kg",      arrow: "->", to: "? g",  type: "M" },
        { from: "4500 mL",   arrow: "->", to: "? L",  type: "C" },
        { from: "2 L",       arrow: "->", to: "? mL", type: "C" },
        { from: "6000 g",    arrow: "->", to: "? kg", type: "M" },
      ];

      const cellW = 2.0;
      const gap = 0.15;
      const totalW = cellW * 4 + gap * 3;
      const startX = (10 - totalW) / 2;
      const cardY = CONTENT_TOP + 0.30;
      const cardH = 1.60;

      items.forEach((it, i) => {
        const cx = startX + i * (cellW + gap);
        addCard(s, cx, cardY, cellW, cardH, { strip: it.type === "M" ? C.PRIMARY : C.SECONDARY, fill: C.WHITE });
        // M or C tag
        addTextOnShape(s, it.type === "M" ? "MASS" : "CAPACITY", {
          x: cx + 0.2, y: cardY + 0.10, w: cellW - 0.4, h: 0.28, rectRadius: 0.05,
          fill: { color: it.type === "M" ? C.PRIMARY : C.SECONDARY },
        }, { fontSize: 10, fontFace: FONT_B, color: C.WHITE, bold: true });

        s.addText(it.from, {
          x: cx, y: cardY + 0.45, w: cellW, h: 0.40,
          fontSize: 20, fontFace: FONT_H, color: C.PRIMARY,
          bold: true, align: "center", margin: 0,
        });
        s.addText(it.arrow, {
          x: cx, y: cardY + 0.90, w: cellW, h: 0.28,
          fontSize: 14, fontFace: FONT_B, color: C.MUTED,
          align: "center", margin: 0,
        });
        s.addText(it.to, {
          x: cx, y: cardY + 1.18, w: cellW, h: 0.40,
          fontSize: 20, fontFace: FONT_H, color: C.SECONDARY,
          bold: true, align: "center", margin: 0,
        });
      });

      s.addText("Partner talk: direction, x1000 or ÷1000, then the answer.", {
        x: 0.7, y: SAFE_BOTTOM - 0.55, w: 8.6, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "3 kg = 3000 g   |   4500 mL = 4.5 L   |   2 L = 2000 mL   |   6000 g = 6 kg", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 15-16: CFU hinge — Sam vs Eli
  withReveal(
    () => cfuSlide(pres, "CFU", "Who is right?", "Thumbs Up or Thumbs Down",
      "Convert 750 g into kg.\n\nSam says 750 g = 7.5 kg. He divided by 100.\n\nEli says 750 g = 0.75 kg. She divided by 1000.\n\nThumbs UP for Sam.   Thumbs DOWN for Eli.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Eli is right. Mass jumps are x1000. 750 g is less than 1 kg.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: Mass & Capacity Conversions", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "use the staircase to choose multiply or divide.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "do the conversion (x1000 or ÷1000).   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "check: more or fewer?", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("UP the staircase = MULTIPLY by 1000     |     DOWN = DIVIDE by 1000", {
      x: 0.7, y: panelY + 0.12, w: 8.6, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    // Two mini staircases side by side
    s.addText("Mass", {
      x: 0.7, y: panelY + 0.50, w: 4.0, h: 0.25,
      fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });
    drawStaircase(s, 0.85, panelY + 0.80, 3.6, 1.10, MASS, "x1000");

    s.addText("Capacity", {
      x: 5.3, y: panelY + 0.50, w: 4.0, h: 0.25,
      fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });
    drawStaircase(s, 5.95, panelY + 0.80, 2.7, 1.10, CAP, "x1000");

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 18: Exit Ticket
  exitTicketSlide(pres,
    [
      "Convert 3.2 kg into g. Show your direction and the multiplier.",
      "Convert 2400 mL into L. Show your direction and the divider.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 19: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how is mass/capacity different from length when converting?",
      scItems: [
        "I can multiply by 1000 to convert from a bigger unit to a smaller unit (e.g. kg to g).",
        "I can divide by 1000 to convert from a smaller unit to a bigger unit (e.g. mL to L).",
        "I can convert decimals (e.g. 1.5 L = 1500 mL) and check the answer makes sense.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Meas_Lesson4_Converting_Mass_Capacity.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Each jump on the staircase is x1000.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "1000 mg = 1 g.  1000 g = 1 kg.  1000 kg = 1 t.  1000 mL = 1 L.  1000 L = 1 kL. Big to small = MULTIPLY by 1000 (decimal moves 3 right). Small to big = DIVIDE by 1000 (decimal moves 3 left).",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Part A — Mass (multiply or divide?)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  4 kg = ___ g", y);
    y = addWriteLine(doc, "b)  5000 g = ___ kg", y);
    y = addWriteLine(doc, "c)  2.5 kg = ___ g", y);
    y = addWriteLine(doc, "d)  850 g = ___ kg", y);

    y = addSectionHeading(doc, "Part B — Capacity (multiply or divide?)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3 L = ___ mL", y);
    y = addWriteLine(doc, "b)  6500 mL = ___ L", y);
    y = addWriteLine(doc, "c)  0.5 L = ___ mL", y);
    y = addWriteLine(doc, "d)  250 mL = ___ L", y);

    y = addSectionHeading(doc, "Part C — Mix", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2.75 kg = ___ g", y);
    y = addWriteLine(doc, "b)  3500 mL = ___ L", y);
    y = addWriteLine(doc, "c)  1.2 t = ___ kg", y);
    y = addWriteLine(doc, "d)  4 kL = ___ L", y);

    y = addSectionHeading(doc, "Part D — Extension: mixed units", y, { color: C.ACCENT });
    y = addBodyText(doc, "Convert each one to a single smaller unit.", y);
    y = addWriteLine(doc, "a)  1 kg 250 g = ___ g", y);
    y = addWriteLine(doc, "b)  2 L 350 mL = ___ mL", y);
    y = addWriteLine(doc, "c)  3 kg 75 g = ___ g", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Mass and Capacity Conversions | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Mass and Capacity Conversions sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Part A — Mass", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4 kg = 4000 g    (x1000)", y);
    y = addBodyText(doc, "b)  5000 g = 5 kg    (÷1000)", y);
    y = addBodyText(doc, "c)  2.5 kg = 2500 g    (x1000; decimal moves 3 right)", y);
    y = addBodyText(doc, "d)  850 g = 0.85 kg    (÷1000; decimal moves 3 left)", y);

    y = addSectionHeading(doc, "Part B — Capacity", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3 L = 3000 mL    (x1000)", y);
    y = addBodyText(doc, "b)  6500 mL = 6.5 L    (÷1000)", y);
    y = addBodyText(doc, "c)  0.5 L = 500 mL    (x1000)", y);
    y = addBodyText(doc, "d)  250 mL = 0.25 L    (÷1000)", y);

    y = addSectionHeading(doc, "Part C — Mix", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2.75 kg = 2750 g    (x1000)", y);
    y = addBodyText(doc, "b)  3500 mL = 3.5 L    (÷1000)", y);
    y = addBodyText(doc, "c)  1.2 t = 1200 kg    (x1000)", y);
    y = addBodyText(doc, "d)  4 kL = 4000 L    (x1000)", y);

    y = addSectionHeading(doc, "Part D — Mixed units", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  1 kg 250 g = 1000 g + 250 g = 1250 g.", y);
    y = addBodyText(doc, "b)  2 L 350 mL = 2000 mL + 350 mL = 2350 mL.", y);
    y = addBodyText(doc, "c)  3 kg 75 g = 3000 g + 75 g = 3075 g.", y);

    y = addTipBox(doc,
      "Watch for: students who use x100 from the length lesson (mass and capacity jumps are x1000); students who put the decimal in the wrong place (count three places, not two); students who skip the 'more or fewer' check.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
