"use strict";

// Measurement Unit (Year 5/6 Numeracy) — Lesson 1: Choosing the right metric unit for length.
// First time looking at this content for the year.
// Daily Review: adding/subtracting fractions and decimals.
// Fluency: addition (within 20).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Theme: Year 5/6 numeracy. Variant fixed across all 7 lessons for cohesion.
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

const SESSION = 1;
const TOTAL = 7;
const UNIT_TITLE = "Measurement: Units, Time and Timetables";
const FOOTER = `Measurement | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/Meas_Lesson1_Choosing_Metric_Units_Length";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 1 Choose the Best Unit",
  "Match the object to the best unit and justify the choice.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 1 Answer Key",
  "Worked answers for the Choose the Best Unit sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome to a new unit on measurement.
- This is the first time we are looking at this content this year, so we will build it carefully.
- Today is all about choosing the right unit when we measure length.

DO:
- Have rulers, a metre stick, and whiteboards ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 1 of 7. The unit covers metric units, 12/24 hour time, elapsed time, and timetables. Today teaches the Year 5 idea of choosing an appropriate unit for the job.

WATCH FOR:
- Students who look unsure - reassure: "If this feels new, that is okay. We are building it together."

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- You will need a ruler, a metre stick on the board, and your whiteboard.

DO:
- Print one Choose the Best Unit sheet per student.
- Place a metre stick or measuring tape across the front of the room.
- Have rulers ready on desks.
- Print one answer key for yourself.

TEACHER NOTES:
One student resource (matching sheet) plus answer key. Most early work is on whiteboards and through partner talk.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are recalling adding fractions with the same denominator.
- Solve each one on your whiteboard.
- Whisper to your partner what the answer simplifies to, if it does.

DO:
- Display the three prompts.
- 90 seconds.
- Walk and scan for line-up errors.

TEACHER NOTES:
Adding fractions with like denominators is prior learning for this cohort. Watch for students who add the denominators - that is the most common slip.

WATCH FOR:
- Students who add numerators only and keep the denominator - secure.
- Students who add both numerator and denominator - quick correction.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 1/4 + 2/4 = 3/4.
- 3/8 + 4/8 = 7/8.
- 2/5 + 1/5 = 3/5.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The key idea: same denominator means the parts are the same size, so we just add the parts. Note any student still doubling the denominator for small group focus.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Addition facts.
- Whisper-answer each one, then write it on your board.

DO:
- Display the three prompts.
- 30 seconds.

TEACHER NOTES:
Fluency this week is addition. We start brisk and easy so every student gets a quick success on day one.

WATCH FOR:
- Students who answer instantly - secure.
- Students who count on - prompt them to use a known fact (eg. doubles or near doubles).

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 3 + 8 = 11.
- 7 + 6 = 13.
- 9 + 5 = 14.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Look for students still counting on rather than recalling - small group focus.

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- Quick vocabulary check.
- These four units are all units of length.
- Whisper to your partner what each one stands for and a real object that uses it.

DO:
- Point to each unit.
- Listen for "millimetre, centimetre, metre, kilometre".
- Use the chant: small to big.

TEACHER NOTES:
The four units we use today: mm, cm, m, km. Year 5 students may know these but rarely connect them to "which unit do I pick?". Locking the size order first makes the choice easier.

WATCH FOR:
- Students who can order them small to big - secure.
- Students who say "cm is the smallest" - prompt with the millimetre marks on the ruler.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to choose the best metric unit when we measure length.
- Now the success criteria.

DO:
- Choral read.
- Hold up the ruler and the metre stick as you read SC2.

TEACHER NOTES:
SC1 is achievable for everyone today - it is recall of unit size. SC2 is the core target. SC3 stretches to combining units like 1 m 35 cm.

WATCH FOR:
- Students who can repeat the success criteria - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. Here are the four units we use for length.
- Smallest to largest: millimetre, centimetre, metre, kilometre.
- 10 millimetres make 1 centimetre.
- 100 centimetres make 1 metre.
- 1000 metres make 1 kilometre.
- The bigger the thing, the bigger the unit I usually pick.

DO:
- Point to a millimetre mark on the ruler.
- Point to a centimetre mark on the ruler.
- Hold up the metre stick.
- Point to the door, then say "from here to the next town is kilometres".

TEACHER NOTES:
This is the size-order anchor for the lesson. Without it, students cannot judge which unit suits the task. The visual stays simple on purpose.

MISCONCEPTIONS:
- Misconception: Students think "more digits is more accurate".
  Why: They confuse precision with unit choice.
  Impact: They measure a desk in millimetres and lose track of the number.
  Quick correction: "We want accurate AND easy to read. Pick a unit that gives a small, clean number."

WATCH FOR:
- Students who nod at the size order - tracking.
- Students who look unsure on the jumps (10, 100, 1000) - we will revisit in Lesson 3.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Watch me think aloud.
- I want to measure the length of my pencil.
- Kilometres? Way too big. The pencil would be 0.00018 km. Ugly number.
- Metres? Still too big. The pencil would be 0.18 m. A decimal is harder to read.
- Centimetres? 18 cm. Clean and easy. I pick centimetres.
- Millimetres? 180 mm. Accurate, but the number is bigger than it needs to be.
- The best unit gives a small whole number with no awkward decimals.

DO:
- Hold up a pencil.
- Write each option on the board as you say it.
- Circle "centimetres".

TEACHER NOTES:
Model the thinking, not just the answer. The trick students need is: "pick the unit that gives a small whole number". Write this rule beside the model.

WATCH FOR:
- Students who repeat "small whole number" - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. The school oval.
- Which unit would you pick: millimetres, centimetres, metres, or kilometres?
- On your whiteboard, write the unit and one short reason.

DO:
- Display the prompt.
- 45 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your boards on three, two, one, show."
- Scan for: "metres" with a sensible reason.
PROCEED: If 80% choose metres with a reason about size, click to reveal and move to We Do.
PIVOT: Most likely misconception - students pick kilometres because "the oval is long".
- Reteach: "Kilometres is for things between towns. The oval is not that long. Most school ovals are about 100 metres."
- Re-check: "What about the long jump pit? Metres or centimetres?"

TEACHER NOTES:
Probe whether students apply the "small whole number" rule. A school oval is 90 to 100 metres, so metres is the clean answer.

WATCH FOR:
- Students who write metres with a clear reason - secure.
- Students who write kilometres - reteach with a 100 m visual.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me.
- We have four objects. We will choose the best unit for each one together.
- Partner talk for each: which unit and why?
- Then I will cold call one pair.

DO:
- Display the four objects.
- 2 minutes total - about 30 seconds per object.
- Cold call after each.

TEACHER NOTES:
Same structure as the I Do, but four guided examples. Listen for "small whole number" reasoning.

WATCH FOR:
- Pairs who use the rule - secure.
- Pairs who default to "centimetres for everything" - prompt with the size of the object.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together.
- Your fingernail: millimetres. Small object, small unit.
- A textbook: centimetres. Bigger than a fingernail, smaller than a metre.
- A classroom wall: metres. About 6 metres long.
- The drive home from school: kilometres. Long distance.

DO:
- Click to reveal.
- Run through each answer with the reason.

TEACHER NOTES:
Reveal confirms the "small whole number" rule for each object. Students who got three of four are secure for the You Do.

WATCH FOR:
- Students who self correct - secure.
- Students who got 2 or fewer - small group focus during You Do.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Two students disagree.
- Tia says the height of a doorway is best measured in centimetres.
- Sam says the height of a doorway is best measured in metres.
- Thumbs UP for Tia. Thumbs DOWN for Sam.

DO:
- Display the disagreement.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP for Tia. DOWN for Sam."
- Scan for: thumbs DOWN. Sam is right.
PROCEED: If 80% agree with Sam, click to reveal and confirm.
PIVOT: Most likely misconception - students pick the unit they know best (cm).
- Reteach: "A doorway is about 2 metres tall. 200 cm works too, but metres gives a smaller, cleaner number."
- Re-check: "What about the height of a kitchen bench? Metres or centimetres?"

TEACHER NOTES:
The disagreement probes the "small whole number" rule. Both units are technically correct, but metres gives the cleaner number for a doorway.

WATCH FOR:
- Confident thumbs DOWN for Tia - they see the rule.
- Thumbs UP for Tia - they need the "small whole number" prompt again.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the Choose the Best Unit sheet.
- For each object, circle the best unit and write one short reason.
- The extension is at the bottom if you finish.

DO:
- Distribute the sheet.
- Circulate. Listen for "small whole number" reasoning.
- Cold call 1-2 students to share their thinking.

TEACHER NOTES:
Different objects from the We Do. Same strategy: pick the unit that gives a small whole number.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the first four items only with the size-order anchor on the board.
- Extra Notes: Sit with these students. Do the first one together.
EXTENDING PROMPT:
- Task: For two items, give a measurement in TWO different units (eg. 2 m AND 200 cm). Explain which is easier to read.
- Extra Notes: Encourage "small whole number" reasoning.

WATCH FOR:
- Students who write a short clear reason - secure.
- Students who pick a unit but leave the reason blank - prompt with "use the rule".

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- For each object, write the best unit and one short reason.

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards or photograph.

TEACHER NOTES:
Exit ticket assesses SC2 (choosing the best unit). Look for unit + reason on both.

WATCH FOR:
- Students who write unit + clear reason - secure.
- Students who write a unit but no reason - prompt with the rule tomorrow.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: what is the rule for picking the best unit?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call 1-2 students.

TEACHER NOTES:
The threshold idea is "pick the unit that gives a small whole number". Students who can say this are ready for mass and capacity tomorrow.

WATCH FOR:
- Strong thumbs up across all three - move at pace tomorrow.
- Sideways or down on SC2 - small group revision in Lesson 2.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a horizontal "size ladder" of the four length units.
function drawUnitLadder(slide, x, y, w, h) {
  const units = [
    { code: "mm", name: "millimetre", desc: "marks on a ruler" },
    { code: "cm", name: "centimetre", desc: "width of a finger" },
    { code: "m",  name: "metre",      desc: "tall ruler / door" },
    { code: "km", name: "kilometre",  desc: "between two towns" },
  ];
  const colors = [C.PRIMARY, C.SECONDARY, C.ACCENT, C.ALERT];
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

  // Arrow underneath: small → big
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

// Draw a small ruler representation with mm and cm marks
function drawMiniRuler(slide, x, y, w) {
  const h = 0.45;
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: C.BG_LIGHT || "F6F6F6" },
    line: { color: C.CHARCOAL, width: 1.2 },
  });
  // 10 cm marks
  for (let i = 0; i <= 10; i++) {
    const tx = x + (w / 10) * i;
    slide.addShape("line", {
      x: tx, y, w: 0, h: h * 0.55,
      line: { color: C.CHARCOAL, width: 1.2 },
    });
    if (i % 5 === 0 || i === 10) {
      slide.addText(String(i), {
        x: tx - 0.12, y: y + h * 0.55, w: 0.24, h: h * 0.45,
        fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0,
      });
    }
  }
  // mm marks (smaller)
  for (let i = 1; i < 100; i++) {
    if (i % 10 === 0) continue;
    const tx = x + (w / 100) * i;
    slide.addShape("line", {
      x: tx, y, w: 0, h: h * 0.25,
      line: { color: C.CHARCOAL, width: 0.5 },
    });
  }
}

// Draw a simple object pictogram. Type drives shape.
function drawObjectIcon(slide, x, y, w, h, type, label) {
  const padBottom = 0.30;
  const drawH = h - padBottom;
  if (type === "oval") {
    addTextOnShape(slide, "", {
      x: x + 0.12, y: y + 0.06, w: w - 0.24, h: drawH - 0.12,
      rectRadius: (drawH - 0.12) / 2,
      fill: { color: C.SECONDARY },
    }, {});
  } else if (type === "house") {
    // Simple roof + box
    slide.addShape("triangle", {
      x: x + 0.30, y: y + 0.06, w: w - 0.60, h: drawH * 0.42,
      fill: { color: C.ALERT }, line: { color: C.CHARCOAL, width: 1 },
    });
    slide.addShape("rect", {
      x: x + 0.40, y: y + drawH * 0.42, w: w - 0.80, h: drawH * 0.50,
      fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
    });
  } else if (type === "pencil") {
    slide.addShape("rect", {
      x: x + 0.10, y: y + drawH * 0.40, w: w - 0.20, h: drawH * 0.20,
      fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
    });
    slide.addShape("triangle", {
      x: x + w - 0.25, y: y + drawH * 0.40, w: 0.25, h: drawH * 0.20,
      fill: { color: C.ALERT }, line: { color: C.CHARCOAL, width: 1 },
    });
  } else if (type === "road") {
    slide.addShape("rect", {
      x: x + 0.05, y: y + drawH * 0.40, w: w - 0.10, h: drawH * 0.20,
      fill: { color: C.CHARCOAL || "2D3142" },
    });
    slide.addShape("rect", {
      x: x + 0.20, y: y + drawH * 0.48, w: w - 0.40, h: drawH * 0.04,
      fill: { color: C.WHITE },
    });
  } else {
    // default: a small filled circle for "object"
    slide.addShape("roundRect", {
      x: x + 0.20, y: y + 0.10, w: w - 0.40, h: drawH - 0.20,
      rectRadius: (drawH - 0.20) / 2,
      fill: { color: C.PRIMARY },
    });
  }

  slide.addText(label, {
    x, y: y + h - padBottom, w, h: padBottom,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL,
    bold: true, align: "center", valign: "top", margin: 0,
  });
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Choosing the best unit for length",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal — adding fractions
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Adding fractions",
      ["1/4 + 2/4 =", "3/8 + 4/8 =", "2/5 + 1/5 ="],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3/4         7/8         3/5", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal — addition within 20
  withReveal(
    () => fluencySlide(pres, "Fluency: Addition",
      ["3 + 8", "7 + 6", "9 + 5"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "11     13     14", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Key vocabulary — the four length units on a size ladder
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Units of length: smallest to largest");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
    drawUnitLadder(s, 0.85, CONTENT_TOP + 0.35, 8.3, 2.2);

    s.addText("Say each one with your partner. Order them small to big.", {
      x: 0.7, y: SAFE_BOTTOM - 0.45, w: 8.6, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to choose the best metric unit when we measure length.",
    [
      "I can name the four units of length (mm, cm, m, km) in order from smallest to largest.",
      "I can choose the best unit for a real object and explain why.",
      "I can give the same length in two different units (e.g. 2 m and 200 cm) and say which is easier to read.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: I Do (1) — the unit ladder with a ruler reference
  workedExSlide(pres, 2, "I Do", "The four units, in order",
    [
      "mm = millimetre. Tiny marks on a ruler.",
      "cm = centimetre. About a fingernail wide.",
      "m = metre. A tall ruler.",
      "km = kilometre. Distance between two towns.",
      "",
      "10 mm = 1 cm.",
      "100 cm = 1 m.",
      "1000 m = 1 km.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("On a ruler", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });

      // Mini ruler showing mm + cm marks
      drawMiniRuler(slide, lg.rightX + 0.20, lg.panelTopPadded + 0.55, lg.rightW - 0.40);

      slide.addText("Big marks = cm.   Small marks = mm.", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 1.20, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      // metre + kilometre comparison underneath
      addTextOnShape(slide, "1 metre = 100 cm", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 1.60, w: lg.rightW - 0.40, h: 0.40, rectRadius: 0.06,
        fill: { color: C.ACCENT },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(slide, "1 kilometre = 1000 m", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.10, w: lg.rightW - 0.40, h: 0.40, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10: I Do (2) — pencil thinking aloud
  workedExSlide(pres, 2, "I Do", "Which unit suits this pencil?",
    [
      "0.00018 km. Ugly. Too big a unit.",
      "0.18 m. Decimal. A bit awkward.",
      "18 cm. Clean. Easy to read.",
      "180 mm. Accurate, but bigger number than needed.",
      "",
      "Best unit: cm.",
      "",
      "Rule: pick the unit",
      "that gives a small whole number.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("Pencil = 18 cm", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.40,
        fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Pencil shape
      const pY = lg.panelTopPadded + 0.65;
      slide.addShape("rect", {
        x: lg.rightX + 0.40, y: pY + 0.20, w: lg.rightW - 1.10, h: 0.30,
        fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      slide.addShape("triangle", {
        x: lg.rightX + lg.rightW - 0.70, y: pY + 0.20, w: 0.30, h: 0.30,
        fill: { color: C.ALERT }, line: { color: C.CHARCOAL, width: 1 },
      });
      // tip
      slide.addShape("rect", {
        x: lg.rightX + 0.20, y: pY + 0.20, w: 0.20, h: 0.30,
        fill: { color: C.MUTED }, line: { color: C.CHARCOAL, width: 1 },
      });

      addTextOnShape(slide, "Rule: small whole number wins.", {
        x: lg.rightX + 0.20, y: lg.panelTopPadded + 1.80, w: lg.rightW - 0.40, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS || C.ACCENT },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("Easy to say. Easy to read.", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 2.50, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slides 11-12: CFU + reveal — school oval
  withReveal(
    () => cfuSlide(pres, "CFU", "Which unit for the school oval?", "Show Me Boards",
      "On your whiteboard:\n\nWrite the BEST unit for the length of the school oval.\nAdd one short reason.\n\nmm     cm     m     km",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "metres   |   the oval is about 100 m long", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 13-14: We Do + reveal — four objects, side by side
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Pick the best unit for each object", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      // Top row of four objects
      const items = [
        { type: "default",  label: "Fingernail" },
        { type: "default",  label: "Textbook" },
        { type: "house",    label: "Classroom wall" },
        { type: "road",     label: "Drive home" },
      ];
      const cellW = 1.95;
      const gap = 0.15;
      const totalW = cellW * 4 + gap * 3;
      const startX = (10 - totalW) / 2;
      const objY = CONTENT_TOP + 0.20;

      items.forEach((it, i) => {
        const cx = startX + i * (cellW + gap);
        // small icon area on top
        if (it.type === "default" && it.label === "Fingernail") {
          // little curved rectangle
          addTextOnShape(s, "", {
            x: cx + 0.55, y: objY + 0.15, w: 0.85, h: 0.55, rectRadius: 0.25,
            fill: { color: C.MUTED },
          }, {});
        } else if (it.type === "default" && it.label === "Textbook") {
          s.addShape("rect", {
            x: cx + 0.35, y: objY + 0.10, w: 1.25, h: 0.85,
            fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
          });
          s.addShape("rect", {
            x: cx + 0.35, y: objY + 0.10, w: 1.25, h: 0.15,
            fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 0.5 },
          });
        } else if (it.type === "house") {
          s.addShape("rect", {
            x: cx + 0.20, y: objY + 0.20, w: 1.55, h: 0.85,
            fill: { color: C.BG_LIGHT || "F6F6F6" }, line: { color: C.CHARCOAL, width: 1 },
          });
          s.addShape("rect", {
            x: cx + 0.80, y: objY + 0.60, w: 0.35, h: 0.45,
            fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 0.5 },
          });
        } else if (it.type === "road") {
          s.addShape("rect", {
            x: cx + 0.15, y: objY + 0.40, w: 1.65, h: 0.25,
            fill: { color: C.CHARCOAL || "2D3142" },
          });
          s.addShape("rect", {
            x: cx + 0.30, y: objY + 0.50, w: 1.35, h: 0.04,
            fill: { color: C.WHITE },
          });
          s.addShape("rect", {
            x: cx + 1.20, y: objY + 0.20, w: 0.50, h: 0.30,
            fill: { color: C.ALERT },
          });
        }

        s.addText(it.label, {
          x: cx, y: objY + 1.15, w: cellW, h: 0.30,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
          bold: true, align: "center", margin: 0,
        });

        // four option chips
        const opts = ["mm", "cm", "m", "km"];
        const oW = cellW / 4 - 0.05;
        opts.forEach((opt, j) => {
          addTextOnShape(s, opt, {
            x: cx + j * (oW + 0.04), y: objY + 1.55, w: oW, h: 0.32, rectRadius: 0.05,
            fill: { color: C.BG_LIGHT || "F6F6F6" },
            line: { color: C.MUTED, width: 0.75 },
          }, { fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true });
        });
      });

      s.addText("Partner talk: which unit and why?  We will check together.", {
        x: 0.7, y: SAFE_BOTTOM - 0.55, w: 8.6, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Fingernail = mm   |   Textbook = cm   |   Wall = m   |   Drive home = km", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 15-16: CFU hinge + reveal — Tia vs Sam (doorway)
  withReveal(
    () => cfuSlide(pres, "CFU", "Who is right?", "Thumbs Up or Thumbs Down",
      "Tia: a doorway is best measured in centimetres.\n\nSam: a doorway is best measured in metres.\n\nThumbs UP for Tia.   Thumbs DOWN for Sam.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Sam is right. A doorway is about 2 m. Cleaner than 200 cm.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: You Do — pointing at the practice sheet
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: Choose the Best Unit", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "circle the BEST unit for each object.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "write ONE short reason.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "try the extension if you finish.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Remember", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "Pick the unit that gives a small whole number.", {
      x: 1.0, y: panelY + 0.55, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

    drawUnitLadder(s, 1.0, panelY + 1.20, 8.0, 0.95);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 18: Exit Ticket
  exitTicketSlide(pres,
    [
      "Choose the best unit for the LENGTH of a SWIMMING POOL. Give one short reason.",
      "Choose the best unit for the LENGTH of an ANT. Give one short reason.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 19: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the rule for picking the best unit?",
      scItems: [
        "I can name the four units of length (mm, cm, m, km) in order from smallest to largest.",
        "I can choose the best unit for a real object and explain why.",
        "I can give the same length in two different units (e.g. 2 m and 200 cm) and say which is easier to read.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Meas_Lesson1_Choosing_Metric_Units_Length.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "For each object, circle the best unit and write one short reason.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Rule: pick the unit that gives a small whole number. mm = tiny marks. cm = fingernail width. m = a tall ruler. km = between two towns.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Part A — Circle the best unit for each object", y, { color: C.PRIMARY });
    const items = [
      "A school pencil",
      "The length of your shoe",
      "The length of a soccer field",
      "The height of a tall tree",
      "The thickness of a 5 cent coin",
      "The drive from school to the city",
      "The length of an eraser",
      "The length of a bus",
    ];
    items.forEach((label) => {
      y = addWriteLine(doc,
        `${label}     mm   /   cm   /   m   /   km     Why? ____________________________`,
        y);
    });

    y = addSectionHeading(doc, "Part B — Same length, two units", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "Some lengths can be written in two units. Write each one and say which is easier to read.", y);
    y = addWriteLine(doc, "a)  A doorway is about 2 metres tall. In centimetres that is ___ cm.", y);
    y = addWriteLine(doc, "b)  A pencil is 18 centimetres long. In millimetres that is ___ mm.", y);
    y = addWriteLine(doc, "c)  Which version is easier to read for each one? Why? __________________________", y);

    y = addSectionHeading(doc, "Part C — Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc,
      "Pick one classroom object. Measure it twice using TWO different units. Which gives the smaller, cleaner number?", y);
    y = addWriteLine(doc, "Object: ______________________________________________", y);
    y = addWriteLine(doc, "Measurement 1: ______________________________________", y);
    y = addWriteLine(doc, "Measurement 2: ______________________________________", y);
    y = addWriteLine(doc, "Easier to read: _____________   Why? __________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Choose the Best Unit | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Choose the Best Unit sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Part A — Best unit", y, { color: C.PRIMARY });
    y = addBodyText(doc, "A school pencil = cm    (about 18 cm; small whole number).", y);
    y = addBodyText(doc, "Length of your shoe = cm    (about 22 cm).", y);
    y = addBodyText(doc, "Soccer field = m    (about 100 m; metres gives the cleanest number).", y);
    y = addBodyText(doc, "Height of a tall tree = m    (about 20 m).", y);
    y = addBodyText(doc, "Thickness of a 5c coin = mm    (about 1 to 2 mm).", y);
    y = addBodyText(doc, "Drive from school to the city = km    (most cities are 5+ km away).", y);
    y = addBodyText(doc, "Length of an eraser = cm    (about 4 to 5 cm).", y);
    y = addBodyText(doc, "Length of a bus = m    (about 12 m).", y);

    y = addSectionHeading(doc, "Part B — Same length, two units", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2 metres = 200 cm.", y);
    y = addBodyText(doc, "b)  18 cm = 180 mm.", y);
    y = addBodyText(doc, "c)  In each case, the SMALLER number is easier to read. Metres for the doorway, centimetres for the pencil.", y);

    y = addSectionHeading(doc, "Part C — Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Look for students who measure twice with different units, then pick the unit that gives the smaller whole number.", y);

    y = addTipBox(doc,
      "Watch for: students who default to centimetres for everything (prompt the size of the object); students who jump to kilometres for anything 'long' (a soccer field is only about 100 m, not 1 km); students who pick a unit but give no reason.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
