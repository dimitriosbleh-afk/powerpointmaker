"use strict";

// Measurement Unit (Year 5/6 Numeracy) — Lesson 2: Choosing metric units for mass and capacity.
// Year 5 content. Same lesson pattern as Lesson 1, swapping length for mass and capacity.
// Daily Review: adding decimals.
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

const SESSION = 2;
const TOTAL = 7;
const UNIT_TITLE = "Measurement: Units, Time and Timetables";
const FOOTER = `Measurement | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/Meas_Lesson2_Choosing_Units_Mass_Capacity";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 2 Mass and Capacity Unit Match",
  "Choose the best mass or capacity unit and justify the choice.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 2 Answer Key",
  "Worked answers for the Mass and Capacity Unit Match sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back. Yesterday we chose the best unit for length.
- Today we use the same rule for mass (how heavy) and capacity (how much it holds).

DO:
- Settle students before clicking past the title.
- Have a small classroom object, a backpack, and an empty water bottle visible.

TEACHER NOTES:
Lesson 2 of 7. Same lesson pattern as Lesson 1, swapping the attribute. The rule from yesterday transfers exactly: pick the unit that gives a small whole number.

WATCH FOR:
- Students who recall the small whole number rule - good carryover from yesterday.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- You will need your whiteboard and the Mass and Capacity Unit Match sheet.

DO:
- Print one sheet per student.
- Have a classroom object, a backpack, and a water bottle ready.
- Print one answer key for yourself.

TEACHER NOTES:
One student resource plus answer key. Real classroom objects strongly support estimation.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Adding decimals.
- Line up the decimal points first, then add.
- Solve each one on your whiteboard.

DO:
- Display the three prompts.
- 2 minutes.
- Walk and scan for line-up errors.

TEACHER NOTES:
Adding decimals reviews last unit. The trap is misaligned decimal points. Watch for students adding 1.2 + 0.45 as 1.65 by misaligning.

WATCH FOR:
- Students who line up correctly - secure.
- Students who right-align like whole numbers - line them up by the decimal point.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 1.2 + 0.45 = 1.65.
- 3.6 + 2.7 = 6.3.
- 0.8 + 0.45 = 1.25.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The key idea: line up the decimal points so the same place values add together.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Addition facts.
- Quick whisper-answer then write on your board.

DO:
- Display the three prompts.
- 30 seconds.

TEACHER NOTES:
Continuing addition fluency this week. Today: making 10 and bridging facts.

WATCH FOR:
- Students who use known facts - secure.
- Students who count on - prompt with doubles.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 8 + 5 = 13.
- 7 + 8 = 15.
- 6 + 9 = 15.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Look for students still counting one by one.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- Quick vocabulary check. Today we have two sets.
- Mass tells us how heavy something is.
- Capacity tells us how much something holds.
- Whisper each unit name to your partner.

DO:
- Point to each unit.
- Listen for the correct names.

TEACHER NOTES:
Mass units: mg (milligram), g (gram), kg (kilogram), t (tonne). Capacity units: mL (millilitre), L (litre), kL (kilolitre). These are the units Year 5 students need today. Megalitre is in Year 6 - we will meet it in Lesson 4.

WATCH FOR:
- Students who can order them small to big - secure.
- Students who say "litre" for both - prompt the difference.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to choose the best metric unit when we measure mass or capacity.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 is recall of unit names. SC2 is the core choosing-the-unit move. SC3 stretches to comparing two units for the same object.

WATCH FOR:
- Students who can repeat the SC - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. Mass units, smallest to largest.
- mg - milligram. The mass of a grain of salt.
- g - gram. The mass of a paperclip.
- kg - kilogram. The mass of a textbook.
- t - tonne. The mass of a car or a small elephant.
- 1000 mg = 1 g.   1000 g = 1 kg.   1000 kg = 1 t.

DO:
- Hold up a paperclip.
- Hold up a textbook.
- Point at the door and say "imagine a car".
- Repeat the size order chant.

TEACHER NOTES:
Mass jumps in thousands at each step. That is different from length (where mm to cm is x10). Flag this early - students often expect tens at every step.

MISCONCEPTIONS:
- Misconception: Students mix mass and capacity (eg. "a kilolitre of bricks").
  Why: They hear "litre" and "lo" sounds in both.
  Impact: They pick wrong units on the worksheet.
  Quick correction: "Litre is what something HOLDS. Gram is what something WEIGHS."

WATCH FOR:
- Students who nod at the size order - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Capacity units, smallest to largest.
- mL - millilitre. A small medicine dose.
- L - litre. A drink bottle.
- kL - kilolitre. A backyard pool or a big rain water tank.
- 1000 mL = 1 L.   1000 L = 1 kL.

DO:
- Hold up a teaspoon (about 5 mL).
- Hold up a drink bottle (about 600 mL).
- Point outside and say "a backyard pool".

TEACHER NOTES:
Same x1000 jumps as mass. The pool example often surprises students - many guess "litres" for a pool, but the rule says "small whole number wins" and a pool is about 30 to 50 kL.

WATCH FOR:
- Students who guess litres for a pool - reteach with the rule.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. The mass of YOU.
- Which unit would you pick: mg, g, kg, or t?
- On your whiteboard, write the unit and one short reason.

DO:
- Display the prompt.
- 45 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your boards on three, two, one, show."
- Scan for: "kg" with a sensible reason.
PROCEED: If 80% choose kg, click to reveal and move to We Do.
PIVOT: Most likely misconception - students pick grams ("more accurate").
- Reteach: "A Year 5/6 student is about 35 to 50 kg. In grams that is 35 000 to 50 000. Ugly. Kg wins."
- Re-check: "What about a school bag full of books?"

TEACHER NOTES:
Probe whether students apply the small whole number rule.

WATCH FOR:
- Students who write kg with a clear reason - secure.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me.
- We have four items. Two mass and two capacity.
- Partner talk for each: which unit and why?

DO:
- Display the four items.
- 2 minutes total.
- Cold call after each.

TEACHER NOTES:
Same structure as the I Do. Listen for small-whole-number reasoning.

WATCH FOR:
- Pairs who use the rule - secure.
- Pairs who default to kg or L for everything - prompt with the size.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- An apple = g. About 150 grams.
- A truck = t. About 10 tonnes.
- A teaspoon = mL. About 5 mL.
- A swimming pool = kL. About 30 to 50 kL.

DO:
- Click to reveal.
- Run the reason for each.

TEACHER NOTES:
Reveal confirms the rule. Students who got three of four are secure.

WATCH FOR:
- Students who got 2 or fewer - small group focus during You Do.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Two students disagree.
- Mason says the mass of a slice of bread is best in milligrams.
- Lina says the mass of a slice of bread is best in grams.
- Thumbs UP for Mason. Thumbs DOWN for Lina.

DO:
- Display the disagreement.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP for Mason. DOWN for Lina."
- Scan for: thumbs DOWN. Lina is right.
PROCEED: If 80% agree with Lina, click to reveal and confirm.
PIVOT: Most likely misconception - students pick mg ("smaller is more accurate").
- Reteach: "A slice of bread is about 30 g. In mg that is 30 000. Big ugly number. g wins."
- Re-check: "What about a vitamin tablet?"

TEACHER NOTES:
The disagreement probes the rule. Both technically work, but g gives the cleaner number.

WATCH FOR:
- Confident thumbs DOWN for Mason - they see the rule.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the Mass and Capacity Unit Match sheet.
- For each item, circle the best unit and write one short reason.
- The extension is at the bottom.

DO:
- Distribute the sheet.
- Circulate. Listen for small-whole-number reasoning.
- Cold call 1-2 students.

TEACHER NOTES:
Different items from the We Do. Same strategy.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use only Part A (first four items) with the unit anchor visible on the board.
- Extra Notes: Sit with these students. Do the first one together.
EXTENDING PROMPT:
- Task: Pick two items. Write the mass or capacity in TWO units (eg. 2 L AND 2000 mL). Which is easier to read?
- Extra Notes: Encourage explicit small-whole-number reasoning.

WATCH FOR:
- Students who write unit + reason - secure.
- Students who pick a unit but skip the reason - prompt with the rule.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- For each item, write the best unit and one short reason.

DO:
- Display the prompt.
- 3 minutes.

TEACHER NOTES:
Exit ticket assesses SC2. Look for unit + reason on both.

WATCH FOR:
- Students who write unit + clear reason - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria again.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: what is the rule for picking the best unit?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
Same threshold idea as Lesson 1, applied to mass and capacity. Tomorrow we convert between units.

WATCH FOR:
- Strong thumbs up - move at pace tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

// Reuse the same unit ladder pattern. units array drives the cells.
function drawUnitLadder(slide, x, y, w, h, units, colors) {
  const gap = 0.10;
  const cellW = (w - gap * (units.length - 1)) / units.length;

  units.forEach((u, i) => {
    const cx = x + i * (cellW + gap);
    addTextOnShape(slide, u.code, {
      x: cx, y, w: cellW, h: h * 0.45, rectRadius: 0.08,
      fill: { color: colors[i] },
    }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });

    slide.addText(u.name, {
      x: cx, y: y + h * 0.48, w: cellW, h: h * 0.22,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL,
      bold: true, align: "center", valign: "middle", margin: 0,
    });
    slide.addText(u.desc, {
      x: cx, y: y + h * 0.70, w: cellW, h: h * 0.28,
      fontSize: 10, fontFace: FONT_B, color: C.MUTED,
      italic: true, align: "center", valign: "top", margin: 0,
    });
  });

  const arrowY = y + h + 0.08;
  slide.addShape("line", {
    x: x + 0.10, y: arrowY, w: w - 0.20, h: 0,
    line: { color: C.CHARCOAL, width: 2 },
  });
  slide.addText("smaller", {
    x, y: arrowY + 0.05, w: 1.0, h: 0.22,
    fontSize: 10, fontFace: FONT_B, color: C.MUTED,
    italic: true, align: "left", margin: 0,
  });
  slide.addText("larger", {
    x: x + w - 1.0, y: arrowY + 0.05, w: 1.0, h: 0.22,
    fontSize: 10, fontFace: FONT_B, color: C.MUTED,
    italic: true, align: "right", margin: 0,
  });
}

const MASS_UNITS = [
  { code: "mg", name: "milligram", desc: "a grain of salt" },
  { code: "g",  name: "gram",      desc: "a paperclip" },
  { code: "kg", name: "kilogram",  desc: "a textbook" },
  { code: "t",  name: "tonne",     desc: "a small car" },
];
const CAP_UNITS = [
  { code: "mL", name: "millilitre", desc: "a medicine spoon" },
  { code: "L",  name: "litre",      desc: "a drink bottle" },
  { code: "kL", name: "kilolitre",  desc: "a backyard pool" },
];

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Choosing units for mass and capacity",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal — adding decimals
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Adding decimals",
      ["1.2 + 0.45 =", "3.6 + 2.7 =", "0.8 + 0.45 ="],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1.65         6.3         1.25", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal — addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Addition",
      ["8 + 5", "7 + 8", "6 + 9"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "13     15     15", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Key vocabulary — mass and capacity unit ladders
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Mass and capacity: smallest to largest");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    s.addText("Mass (how heavy)", {
      x: 0.7, y: CONTENT_TOP + 0.10, w: 8.6, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });
    drawUnitLadder(s, 0.85, CONTENT_TOP + 0.45, 8.3, 1.10,
      MASS_UNITS, [C.PRIMARY, C.SECONDARY, C.ACCENT, C.ALERT]);

    s.addText("Capacity (how much it holds)", {
      x: 0.7, y: CONTENT_TOP + 1.85, w: 8.6, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });
    drawUnitLadder(s, 1.65, CONTENT_TOP + 2.20, 6.7, 1.10,
      CAP_UNITS, [C.PRIMARY, C.ACCENT, C.ALERT]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to choose the best metric unit when we measure mass or capacity.",
    [
      "I can name the units of mass (mg, g, kg, t) and capacity (mL, L, kL) in order.",
      "I can choose the best unit for a real object and explain why.",
      "I can give the same measurement in two different units (e.g. 2 L and 2000 mL) and say which is easier to read.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: I Do (1) — mass units with examples
  workedExSlide(pres, 2, "I Do", "Mass units, smallest to largest",
    [
      "mg = milligram. A grain of salt.",
      "g = gram. A paperclip (about 1 g).",
      "kg = kilogram. A textbook (about 1 kg).",
      "t = tonne. A small car (about 1 t).",
      "",
      "Each jump is x1000.",
      "1000 mg = 1 g.",
      "1000 g = 1 kg.",
      "1000 kg = 1 t.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("Mass examples", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });

      const examples = [
        { unit: "mg", obj: "grain of salt", color: C.PRIMARY },
        { unit: "g",  obj: "paperclip",     color: C.SECONDARY },
        { unit: "kg", obj: "textbook",      color: C.ACCENT },
        { unit: "t",  obj: "small car",     color: C.ALERT },
      ];
      const exH = 0.50;
      examples.forEach((ex, i) => {
        const exY = lg.panelTopPadded + 0.55 + i * (exH + 0.08);
        addTextOnShape(slide, ex.unit, {
          x: lg.rightX + 0.15, y: exY, w: 0.85, h: exH, rectRadius: 0.06,
          fill: { color: ex.color },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(ex.obj, {
          x: lg.rightX + 1.10, y: exY, w: lg.rightW - 1.25, h: exH,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
          bold: true, align: "left", valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10: I Do (2) — capacity units
  workedExSlide(pres, 2, "I Do", "Capacity units, smallest to largest",
    [
      "mL = millilitre. A teaspoon (about 5 mL).",
      "L = litre. A drink bottle (about 600 mL).",
      "kL = kilolitre. A backyard pool.",
      "",
      "Each jump is x1000.",
      "1000 mL = 1 L.",
      "1000 L = 1 kL.",
      "",
      "Same rule: small whole number wins.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("Capacity examples", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      const examples = [
        { unit: "mL", obj: "teaspoon (5 mL)",        color: C.PRIMARY },
        { unit: "L",  obj: "drink bottle (600 mL)",  color: C.ACCENT },
        { unit: "kL", obj: "backyard pool (40 kL)",  color: C.ALERT },
      ];
      const exH = 0.65;
      examples.forEach((ex, i) => {
        const exY = lg.panelTopPadded + 0.55 + i * (exH + 0.12);
        addTextOnShape(slide, ex.unit, {
          x: lg.rightX + 0.15, y: exY, w: 0.95, h: exH, rectRadius: 0.06,
          fill: { color: ex.color },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(ex.obj, {
          x: lg.rightX + 1.20, y: exY, w: lg.rightW - 1.35, h: exH,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
          bold: true, align: "left", valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slides 11-12: CFU + reveal — mass of YOU
  withReveal(
    () => cfuSlide(pres, "CFU", "Which unit for the mass of YOU?", "Show Me Boards",
      "On your whiteboard:\n\nWrite the BEST unit for the mass of a Year 5/6 student.\nAdd one short reason.\n\nmg     g     kg     t",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "kilograms   |   you are about 35 to 50 kg", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 13-14: We Do + reveal — four mixed items
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Pick the best unit for each", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      const items = [
        { kind: "mass",     label: "An apple",        options: ["mg", "g", "kg", "t"] },
        { kind: "mass",     label: "A truck",         options: ["mg", "g", "kg", "t"] },
        { kind: "capacity", label: "A teaspoon",      options: ["mL", "L", "kL", ""] },
        { kind: "capacity", label: "A swimming pool", options: ["mL", "L", "kL", ""] },
      ];

      const cellW = 1.95;
      const gap = 0.15;
      const totalW = cellW * 4 + gap * 3;
      const startX = (10 - totalW) / 2;
      const objY = CONTENT_TOP + 0.20;

      items.forEach((it, i) => {
        const cx = startX + i * (cellW + gap);
        // simple visual: coloured circle with letter inside
        const colorPick = it.kind === "mass" ? C.PRIMARY : C.SECONDARY;
        addTextOnShape(s, it.kind === "mass" ? "M" : "C", {
          x: cx + 0.55, y: objY + 0.10, w: 0.85, h: 0.85, rectRadius: 0.42,
          fill: { color: colorPick },
        }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });

        s.addText(it.label, {
          x: cx, y: objY + 1.05, w: cellW, h: 0.35,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
          bold: true, align: "center", margin: 0,
        });

        const oW = cellW / 4 - 0.05;
        it.options.forEach((opt, j) => {
          if (!opt) return;
          addTextOnShape(s, opt, {
            x: cx + j * (oW + 0.04), y: objY + 1.55, w: oW, h: 0.32, rectRadius: 0.05,
            fill: { color: C.BG_LIGHT || "F6F6F6" },
            line: { color: C.MUTED, width: 0.75 },
          }, { fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true });
        });
      });

      s.addText("M = mass.  C = capacity.  Partner talk: which unit and why?", {
        x: 0.7, y: SAFE_BOTTOM - 0.55, w: 8.6, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Apple = g   |   Truck = t   |   Teaspoon = mL   |   Pool = kL", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 15-16: CFU hinge + reveal — Mason vs Lina (bread)
  withReveal(
    () => cfuSlide(pres, "CFU", "Who is right?", "Thumbs Up or Thumbs Down",
      "Mason: a slice of bread is best in milligrams.\n\nLina: a slice of bread is best in grams.\n\nThumbs UP for Mason.   Thumbs DOWN for Lina.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Lina is right. A slice is about 30 g. Cleaner than 30 000 mg.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: You Do — practice sheet
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: Mass and Capacity Unit Match", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "circle the BEST unit for each item.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "write ONE short reason.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "try the extension.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Remember", {
      x: 0.7, y: panelY + 0.10, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "Pick the unit that gives a small whole number.", {
      x: 1.0, y: panelY + 0.45, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Two mini ladders side by side
    s.addText("Mass: mg, g, kg, t.   Capacity: mL, L, kL.", {
      x: 0.7, y: panelY + 1.05, w: 8.6, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
      bold: true, align: "center", margin: 0,
    });

    addTextOnShape(s, "Each jump is x1000.", {
      x: 2.5, y: panelY + 1.45, w: 5.0, h: 0.40, rectRadius: 0.06,
      fill: { color: C.ACCENT },
    }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 18: Exit Ticket
  exitTicketSlide(pres,
    [
      "Choose the best unit for the MASS of a CAT. Give one short reason.",
      "Choose the best unit for the CAPACITY of a CUP of tea. Give one short reason.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 19: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the rule for picking the best unit?",
      scItems: [
        "I can name the units of mass (mg, g, kg, t) and capacity (mL, L, kL) in order.",
        "I can choose the best unit for a real object and explain why.",
        "I can give the same measurement in two different units (e.g. 2 L and 2000 mL) and say which is easier to read.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Meas_Lesson2_Choosing_Units_Mass_Capacity.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "For each item, circle the best unit and write one short reason.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Rule: pick the unit that gives a small whole number. Mass: mg, g, kg, t (each jump x1000). Capacity: mL, L, kL (each jump x1000).",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Part A — Mass: circle the best unit", y, { color: C.PRIMARY });
    const massItems = [
      "A vitamin tablet",
      "A loaf of bread",
      "A school bag full of books",
      "A blue whale",
    ];
    massItems.forEach((label) => {
      y = addWriteLine(doc,
        `${label}     mg   /   g   /   kg   /   t     Why? ____________________________`,
        y);
    });

    y = addSectionHeading(doc, "Part B — Capacity: circle the best unit", y, { color: C.PRIMARY });
    const capItems = [
      "A medicine cup",
      "A drink bottle",
      "A bath full of water",
      "An eye drop",
    ];
    capItems.forEach((label) => {
      y = addWriteLine(doc,
        `${label}     mL   /   L   /   kL     Why? ____________________________________`,
        y);
    });

    y = addSectionHeading(doc, "Part C — Same measurement, two units", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  A drink bottle holds 2 L. In mL that is ___ mL.", y);
    y = addWriteLine(doc, "b)  A bag of flour is 2 kg. In grams that is ___ g.", y);
    y = addWriteLine(doc, "c)  Which version is easier to read for each? Why? __________________________", y);

    y = addSectionHeading(doc, "Part D — Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc,
      "Pick one classroom item. Estimate its mass OR capacity in two different units. Which is easier to read?", y);
    y = addWriteLine(doc, "Item: _______________________________________________", y);
    y = addWriteLine(doc, "Estimate 1: _________________________________________", y);
    y = addWriteLine(doc, "Estimate 2: _________________________________________", y);
    y = addWriteLine(doc, "Easier to read: ___________   Why? ___________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Mass and Capacity Unit Match | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Mass and Capacity Unit Match sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Part A — Mass", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Vitamin tablet = mg    (about 500 mg).", y);
    y = addBodyText(doc, "Loaf of bread = g    (about 700 g).", y);
    y = addBodyText(doc, "School bag of books = kg    (about 5 to 8 kg).", y);
    y = addBodyText(doc, "Blue whale = t    (about 100 to 150 tonnes).", y);

    y = addSectionHeading(doc, "Part B — Capacity", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Medicine cup = mL    (about 20 to 30 mL).", y);
    y = addBodyText(doc, "Drink bottle = mL or L    (a small bottle is best in mL; a 1 L bottle is best in L).", y);
    y = addBodyText(doc, "Bath full of water = L    (about 150 L).", y);
    y = addBodyText(doc, "An eye drop = mL    (less than 1 mL each).", y);

    y = addSectionHeading(doc, "Part C — Two units", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2 L = 2000 mL.", y);
    y = addBodyText(doc, "b)  2 kg = 2000 g.", y);
    y = addBodyText(doc, "c)  In each case, the smaller whole number is easier to read. L for the bottle, kg for the bag.", y);

    y = addSectionHeading(doc, "Part D — Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Accept any reasonable estimate. Look for smaller whole number choice and clear reasoning.", y);

    y = addTipBox(doc,
      "Watch for: students who pick the smallest unit ('mg sounds accurate'); students who confuse mass and capacity ('kL for a bag of bricks'); students who skip the reason.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
