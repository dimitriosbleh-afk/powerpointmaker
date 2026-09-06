"use strict";

// Metric Measurement & Conversion — WEEK 4 REVIEW (Year 5/6 composite Numeracy)
// A single consolidation session across the whole of Week 4 Term 3:
//   S1 choose & order units | S2 convert length | S3 convert mass
//   S4 convert capacity | S5 choose, convert & compare
//
// Victorian Curriculum 2.0: Mathematics, Measurement, Year 5/6 — choosing and
// converting between metric units; comparing measurements by converting to a
// common unit first.
//
// UNIT ANCHOR (locked from the week, megaprompt section 5b — the teacher's own
// wording, never swapped for the formal term):
//   "Set the stairs. Count the steps. Multiply or divide."
//   Down to a smaller unit -> multiply.  Up to a bigger unit -> divide.
//   "Make the units match" before comparing or calculating.
//   Fluency is the DIVISION BRACKET (never "bus stop").
//
// Lesson shape: consolidation with an interleaved practice block. Conversions
// deliberately alternate attribute and direction rather than being taught in
// blocked attribute order as the week did, so students must DECIDE rather than
// copy (megaprompt section 77).
//
// Daily Review: multiply/divide by powers of 10 (Week 3 decimals) — the engine
// of every conversion, so retrieval feeds directly into today.
// Fluency: division bracket, as in every session this week.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  PAGE, hex,
} = require("../themes/pdf_helpers");

// Same unit, same palette: Week 4 shipped on variant 0 and this review belongs
// to that unit. Variant 0 also keeps ACCENT contrast-safe for text on fills.
const UNIT_VARIANT = 0;
const T = createTheme("numeracy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide, keyWordSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide, addRevealAnswerBar,
  addPlaceValueChart,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  clickBuild, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

// The review sits after Sessions 1-5 in the same weekly sequence, so its
// resources are numbered 6 and stay unique alongside the week's own PDFs.
const SESSION = 6;
const FOOTER = "Metric Measurement | Week 4 Review | Year 5/6 Numeracy";
const OUT_DIR = "output/MMC_Week4_Review_Session";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Review Practice",
  "Convert, compare and solve across length, mass and capacity.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the review practice sheet.");
const SCAFFOLD_RES = makeSessionResource(SESSION,
  "Staircase Helper",
  "The three staircases drawn, with the first conversion in each already worked. Glue into books.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, SCAFFOLD_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher notes (Glance Format, megaprompt sections 45-47) ────────────────

const NOTES_TITLE =
  "Open the review. Name the week in one line, then move to Teacher Resources.";

const NOTES_RESOURCES =
  "Prep slide. Print the review practice sheet, the answer key, and the Staircase Helper " +
  "for anyone who needs it. Have whiteboards, markers and a metre ruler ready.\n" +
  "CATCH-UP: missed a session this week? Hand them the Staircase Helper, point to the " +
  "recap slide, say the three moves, and start them at Section 1 of the practice sheet - " +
  "it rebuilds the routine from scratch and needs no earlier session.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "This session reviews all five Week 4 sessions in one lesson. It re-teaches the one " +
  "routine that ran the whole week and then interleaves the attributes deliberately, so " +
  "students decide the direction and the factor rather than repeating a blocked pattern.\n" +
  "Decision points: the boards check after the convert block, the hinge after the compare " +
  "block, and the exit ticket. Between them keep the pace brisk.\n" +
  "Flex: the multi-step We Do can be cut if the compare block runs long.\n" +
  "SOURCES: Numeracy Week 4 Term 3 deck (Sessions 1-5), supplied by the teacher.";

const NOTES_DR = composeGlanceNotes({
  answer: "470, 3.5 and 1250.",
  beats: [
    "POINT to the place value chart. SAY: Digits move. The point stays still.",
    [
      "ASK: Write all three answers.",
      "60 sec. Cue: Write it... chin it... show me.",
      "EXPECT: 470, 3.5, 1250.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> reveal, tick and fix, straight into fluency.",
      "Less -> place 4.7 on the chart, slide each digit two left, re-ask.",
    ],
  ],
  trap: [
    "adding a zero, so 4.7 times 100 is written 4.70.",
    "Fix: move each digit two columns, student rewrites it.",
  ],
  prep: "Retrieval from Week 3 decimals. Multiplying and dividing by powers of 10 IS the engine of every conversion today. Whole block under 5 minutes.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_FLUENCY = composeGlanceNotes({
  answer: "32, 39 and 19.",
  beats: [
    "SAY: Fluency. Set each one out with the division bracket, work from the left.",
    [
      "ASK: What are the three answers?",
      "80 sec. Boards up on cue.",
      "EXPECT: 32, 39, 19.",
    ],
    [
      "SCAN boards.",
      "80%+ -> reveal, tick and fix.",
      "Less -> rebuild 224 divided by 7 in the bracket, carry the 1, re-ask.",
    ],
  ],
  trap: [
    "dropping the carry, so 152 divided by 8 lands near 10.",
    "Fix: redo the carry step in the bracket, student rewrites it.",
  ],
  prep: "The same fluency focus held across every session this week: the division bracket. Brisk retrieval, under 4 minutes, not new teaching.",
  tag: "[Stage 1 | Fluency | Retention and recall | HITS 6]",
});

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - listen for the idea that all three break into equal smaller parts.",
  beats: [
    "POINT to the three facts. SAY: One metre, one kilogram, one litre.",
    [
      "ASK: What is the SAME about all three?",
      "40 sec. Turn and tell. Partner A first.",
      "EXPECT: each one splits into equal smaller parts.",
    ],
    [
      "SCAN the room as pairs talk.",
      "80%+ talking about the splitting -> reveal, then next slide.",
      "Less -> hold up a metre ruler, count the 100 centimetres, re-ask.",
    ],
    "REVEAL the shared idea after partner talk. SAY: One routine covers all three.",
  ],
  trap: [
    "thinking each attribute needs its own separate method.",
    "Fix: name the one routine, students say the three moves back.",
  ],
  prep: "Low-coupling launch: answerable from 1 m = 100 cm alone, so a student who missed sessions can join from here.",
  tag: "[Launch | Knowledge and memory | HITS 2, 6]",
});

const NOTES_LI_SC = composeGlanceNotes({
  beats: [
    "POINT to the learning intention. SAY: One routine, three kinds of measurement.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone can do today - say which unit is bigger.",
  ],
  prep: "SC1 reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches to comparing and explaining. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "equivalent means equal in value - one amount, written two ways.",
  beats: [
    "SAY: Equivalent means equal in value. Same amount, different unit.",
    [
      "POINT to 1.25 m and 125 cm.",
      "SAY: These are not two lengths. They are one length, written twice.",
    ],
    [
      "ASK: Is 1250 g the same amount as 1.25 kg?",
      "10 sec. Thumbs only, voices off. Show me... now.",
      "EXPECT: thumbs up.",
    ],
  ],
  trap: [
    "thinking 125 cm is longer because 125 beats 1.25.",
    "Fix: hold the metre ruler, student says both name the same length.",
  ],
  prep: "The one word that carries the whole week. Students who can explain equivalence stop distrusting their own decimal answers.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_RECAP = composeGlanceNotes({
  answer: "down to a smaller unit multiply, up to a bigger unit divide.",
  beats: [
    "POINT across the staircases. SAY: Set the stairs. Count the steps. Multiply or divide.",
    [
      "SAY: Mass and capacity are kinder - every step is times one thousand.",
      "SAY: Length steps change: 1000, then 100, then 10.",
    ],
    [
      "ASK: Say the three moves back to me.",
      "5 sec. Choral response: everyone, together, on three.",
      "EXPECT: set the stairs, count the steps, multiply or divide.",
    ],
  ],
  trap: [
    "dividing to reach a smaller unit because the unit sounds small.",
    "Fix: ask more or fewer, student re-decides.",
  ],
  stretch: "which two staircases share the same step size?",
  help: "Staircase Helper card, student traces one step with a finger.",
  prep: "The catch-up re-entry slide. A student who missed Sessions 2 to 4 can convert from this slide alone.",
  tag: "[Recap | Knowledge and memory | HITS 3, 6]",
});

const NOTES_IDO1 = composeGlanceNotes({
  answer: "2400 m. One step down, so times 1000.",
  beats: [
    "SAY: Watch how I decide. Same three moves, every single time.",
    "SAY: One step from kilometres to metres, and that step is a thousand.",
    [
      "SAY: Down to a smaller unit means I need MORE of them.",
      "SAY: So I multiply. 2.4 times 1000.",
    ],
    "REVEAL 2400 m after you have said the working.",
  ],
  trap: [
    "dividing because metres sound smaller.",
    "Fix: ask more or fewer, student re-decides.",
  ],
  stretch: "convert 2.4 km into centimetres. Two steps.",
  help: "Staircase Helper, finger on km, one step down.",
  prep: "First model. The anchor is restated in full so a returning student can follow from this slide alone.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_IDO2 = composeGlanceNotes({
  answer: "0.75 kg. One step up, so divided by 1000.",
  beats: [
    "SAY: New attribute, same routine. Mass this time.",
    "SAY: Grams up to kilograms is one step, and every mass step is a thousand.",
    [
      "SAY: Up to a bigger unit means FEWER of them.",
      "SAY: So I divide. 750 divided by 1000.",
    ],
    "REVEAL 0.75 kg after the divide is said.",
  ],
  trap: [
    "writing 7.5 kg from a mis-counted place value move.",
    "Fix: rebuild it on the chart, student rewrites.",
  ],
  stretch: "how many milligrams in 750 g?",
  help: "number line 0 to 1 kg marked every 250 g.",
  prep: "Second model, deliberately a different attribute AND the divide direction. Interleaved so students decide, not copy.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_CFU1 = composeGlanceNotes({
  answer: "3500 mL. One step down, so times 1000.",
  beats: [
    "SAY: Check on your boards. Capacity this time, same routine.",
    [
      "ASK: Convert 3.5 litres into millilitres.",
      "40 sec. Cue: Write it... chin it... show me.",
      "EXPECT: 3500 mL.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one board: How did you decide to multiply?",
      "Less -> set the capacity stairs, count one step, re-ask 2.5 L.",
    ],
    "REVEAL only after every board is up.",
  ],
  trap: [
    "writing 350 mL from a single place value move.",
    "Fix: three places, not one, student recounts.",
  ],
  prep: "Decision point one. Checks the routine transfers to a third attribute before guided practice.",
  tag: "[Stage 2 | CFU | Supported application | HITS 7, 8]",
});

const NOTES_WEDO1 = composeGlanceNotes({
  answer: "1750 g. One step down, so times 1000.",
  beats: [
    "SAY: Your turn with me. Run the three moves with your partner.",
    [
      "ASK: Convert 1.75 kilograms into grams.",
      "60 sec. Boards up on cue.",
      "EXPECT: 1750 g.",
    ],
    [
      "CIRCULATE. Look for the stairs drawn before the answer.",
      "80%+ -> reveal, then straight on.",
      "Less -> build 1 kg as 1000 g, add the 0.75, re-ask.",
    ],
  ],
  trap: [
    "answering 175 g by moving two places, not three.",
    "Fix: count the three moves aloud, student redoes.",
  ],
  stretch: "write 1.75 kg in milligrams.",
  help: "part-filled chart with 1000 g already under 1 kg.",
  prep: "Guided practice, decimal times 1000. Faded from the I Do: no worked steps left on screen.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_WEDO2 = composeGlanceNotes({
  answer: "4.2 L. One step up, so divided by 1000.",
  beats: [
    "SAY: Now the other direction, and back to capacity.",
    [
      "ASK: Convert 4200 millilitres into litres.",
      "60 sec. Boards up on cue.",
      "EXPECT: 4.2 L.",
    ],
    [
      "CIRCULATE. Check the direction decision before the arithmetic.",
      "80%+ -> reveal, then on to comparing.",
      "Less -> ask more litres or fewer, re-ask.",
    ],
  ],
  trap: [
    "deleting the zeros to get 42 L.",
    "Fix: three places, so 4200 becomes 4.2, student rewrites.",
  ],
  stretch: "4200 mL plus 800 mL, written in litres.",
  help: "Staircase Helper, finger up one step from mL to L.",
  prep: "Same routine, opposite move, so the direction has to be decided rather than repeated.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_IDO3 = composeGlanceNotes({
  answer: "850 g is heavier. 0.8 kg is only 800 g.",
  beats: [
    "SAY: Now the golden rule. Make the units match before you compare.",
    "SAY: I cannot compare 0.8 with 850. The units are different.",
    [
      "SAY: I convert the kilograms down into grams.",
      "SAY: 0.8 times 1000 is 800 grams.",
    ],
    "REVEAL after the conversion is said. SAY: 800 is less than 850.",
  ],
  trap: [
    "picking 0.8 kg because kilograms sound bigger.",
    "Fix: convert first, student names the heavier one.",
  ],
  stretch: "which is heavier, 0.08 kg or 85 g?",
  help: "both already in grams, student only compares.",
  prep: "Session 5's move, re-modelled. Note: the original Session 5 slide wrote 0.5 kg as 500 mL - it is 500 g.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_WEDO3 = composeGlanceNotes({
  answer: "2.5 m is longer. It is 250 cm, and 250 beats 240.",
  beats: [
    "SAY: Your turn. Same golden rule, this time with length.",
    [
      "ASK: Which is longer, 2.5 metres or 240 centimetres?",
      "60 sec. Convert first. Boards up on cue.",
      "EXPECT: 2.5 m, shown as 250 cm.",
    ],
    [
      "CIRCULATE. Look for a conversion written before the choice.",
      "80%+ -> reveal, then bounce: Do you agree? Add one thing.",
      "Less -> convert 2.5 m together, re-ask 3 m or 280 cm.",
    ],
  ],
  trap: [
    "choosing 240 cm because 240 is the bigger number.",
    "Fix: convert both to centimetres, student re-decides.",
  ],
  stretch: "order 1.4 m, 130 cm and 1450 mm.",
  help: "2.5 m already shown as 250 cm, student compares.",
  prep: "Guided compare. The number-beats-unit trap is the week's most persistent error.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_WEDO4 = composeGlanceNotes({
  answer: "1350 mL left. 2 L is 2000 mL, then take away 650.",
  beats: [
    "SAY: Last one together. Two moves now: convert, then subtract.",
    [
      "ASK: A 2 litre bottle. You drink 650 mL. How much is left?",
      "90 sec with your partner. Boards up on cue.",
      "EXPECT: 1350 mL.",
    ],
    [
      "CIRCULATE. Check they converted BEFORE subtracting.",
      "80%+ -> reveal, then straight to the check.",
      "Less -> write 2 L as 2000 mL first, re-ask.",
    ],
  ],
  trap: [
    "subtracting across units, such as 2 take away 650.",
    "Fix: convert first, student redoes both steps.",
  ],
  stretch: "how much more to finish the bottle, in litres?",
  help: "the 2000 mL step already written, student subtracts.",
  prep: "Flex slide - cut this if the compare block ran long. The exit ticket still reaches multi-step work.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_CFU2 = composeGlanceNotes({
  answer: "C. 2.5 kg is 2500 g, because every mass step is 1000.",
  beats: [
    "SAY: Hinge question. A student wrote this. Is it right?",
    [
      "ASK: Is it A, B or C?",
      "30 sec, voices off. Cue: Write it... chin it... show me.",
      "EXPECT: C.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one C: Convince us. Then reveal.",
      "Less -> set the mass stairs, count the step, re-ask 1.5 kg.",
    ],
    "REVEAL only after every board is up.",
  ],
  trap: [
    "choosing B, using the length step of 100 on mass.",
    "Fix: point to the mass stairs, student redoes.",
  ],
  prep: "Decision point two, the hinge of the week. B means right routine, wrong factor. A means no conversion at all.",
  tag: "[Stage 2 | CFU hinge | Supported application | HITS 7, 8]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - listen for the stairs drawn before any arithmetic.",
  beats: [
    "SAY: On your own now. Section 1 rebuilds the routine, then it builds up.",
    "SAY: Every question, same three moves. Match the units before comparing.",
    [
      "CIRCULATE. Sit with the catch-up group first.",
      "COLLECT: check Section 1 before anyone starts Section 3.",
    ],
    "TIME: about 10 minutes, then stop for the exit ticket.",
  ],
  stretch: "the Challenge box - order three measurements in three units.",
  help: "the Staircase Helper, and Section 1 worked with the teacher.",
  prep: "Section 1 is the re-grounding task: doable from the recap slide alone, with no earlier session needed.",
  tag: "[Stage 4 | You Do | Mastery and application | HITS 4, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "2.4 kg. And 1.6 m is longer, because 1.6 m is 160 cm.",
  beats: [
    "SAY: On your own, on your whiteboard. No partner talk for this one.",
    [
      "SAY: One, convert 2400 grams into kilograms.",
      "SAY: Two, decide which is longer and explain how you know.",
    ],
    [
      "TIME: about 3 minutes.",
      "COLLECT boards, or photograph them for your records.",
    ],
  ],
  trap: [
    "choosing 150 cm because 150 beats 1.6.",
    "Fix: note the name, reteach convert-before-compare next session.",
  ],
  prep: "Assesses the core target (converting with the right operation) and reaches into comparing. Keep the SC number off the slide.",
  tag: "[Stage 5 | Exit Ticket | Mastery and application | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - listen for make the units match, and the multiply or divide rule.",
  beats: [
    "SAY: Look back at the three I can statements. Thumbs for each one.",
    [
      "ASK: What is the one rule you will never forget about converting?",
      "30 sec. Turn and tell. Partner B first.",
      "EXPECT: match the units, or smaller unit means multiply.",
    ],
    "COLLECT: note who showed need more practice, for a short revisit.",
  ],
  prep: "Whole-week close. The keeper: a smaller unit needs a bigger number, and units must match before comparing.",
  tag: "[Closing | Retention and recall | HITS 9]",
});

// ─── Visual helpers (build-script local) ─────────────────────────────────────

// One metric staircase drawn at an explicit position: units largest at the top,
// smallest at the bottom, with the multiply/divide factor between each step.
// The unit's hero visual and its catch-up anchor.
function staircase(slide, x, y, w, bottom, opts) {
  const o = opts || {};
  const color = o.color || C.PRIMARY;
  const steps = o.steps;

  addCard(slide, x, y, w, bottom - y, { strip: color });
  slide.addText(o.title, {
    x: x + 0.12, y: y + 0.10, w: w - 0.24, h: 0.30,
    fontSize: 15, fontFace: FONT_H, color, bold: true, align: "center", margin: 0,
  });

  // Fixed pill and gap sizes across every staircase so three cards side by side
  // read as one diagram; a shorter staircase centres rather than stretching.
  // Four pills plus three gaps must clear the card, or the last unit overhangs.
  const pillH = 0.34;
  const gap = 0.38;
  const innerTop = y + 0.48;
  const region = bottom - innerTop - 0.10;
  const stackH = steps.length * pillH + (steps.length - 1) * gap;
  const pillW = w - 0.6;
  const pillX = x + (w - pillW) / 2;
  let cy = innerTop + Math.max(0, (region - stackH) / 2);

  steps.forEach((st, i) => {
    if (i > 0) {
      const halfW = pillW / 2 - 0.04;
      slide.addText("× " + st.f, {
        x: pillX, y: cy - gap + 0.05, w: halfW, h: gap - 0.10,
        fontSize: 12.5, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "right", valign: "middle", margin: 0,
      });
      slide.addText("÷ " + st.f, {
        x: pillX + pillW / 2 + 0.04, y: cy - gap + 0.05, w: halfW, h: gap - 0.10,
        fontSize: 12.5, fontFace: FONT_B, color: C.ALERT, bold: true,
        align: "left", valign: "middle", margin: 0,
      });
    }
    addTextOnShape(slide, st.u, {
      x: pillX, y: cy, w: pillW, h: pillH, rectRadius: 0.08,
      fill: { color: i === 0 ? color : C.SECONDARY },
    }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    cy += pillH + gap;
  });
}

// A worked conversion box: start value, the operation pill, and the result.
// Returns a revealEnd() so the result can be held back for a click build while
// the start value and the operation stay on screen throughout.
function convertPanel(slide, lg, opts) {
  const o = opts || {};
  const x = lg.rightX;
  const w = lg.rightW;
  const y0 = lg.panelTopPadded;
  const bottom = o.bottom != null ? o.bottom : lg.safeBottom;
  const color = o.color || C.PRIMARY;

  addCard(slide, x, y0, w, bottom - y0, { strip: color });
  slide.addText(o.title || "Convert", {
    x: x + 0.18, y: y0 + 0.10, w: w - 0.36, h: 0.28,
    fontSize: 14.5, fontFace: FONT_H, color, bold: true, align: "center", margin: 0,
  });

  const innerX = x + 0.45;
  const innerW = w - 0.9;
  const startY = y0 + 0.50;
  const captionReserve = o.caption ? 0.60 : 0.12;
  const region = bottom - startY - captionReserve;
  const boxH = region * 0.30;
  const opH = region * 0.22;
  const gap = (region - 2 * boxH - opH) / 2;

  addTextOnShape(slide, o.start, {
    x: innerX, y: startY, w: innerW, h: boxH, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 23, fontFace: FONT_H, color: C.WHITE, bold: true });

  const opColor = o.divide ? C.ALERT : C.SUCCESS;
  addTextOnShape(slide, (o.divide ? "÷ " : "× ") + o.factor, {
    x: innerX + innerW * 0.18, y: startY + boxH + gap, w: innerW * 0.64, h: opH, rectRadius: 0.08,
    fill: { color: opColor },
  }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });

  const endY = startY + boxH + gap + opH + gap;
  const capY = startY + region + 0.12;

  return {
    revealEnd() {
      addTextOnShape(slide, o.end, {
        x: innerX, y: endY, w: innerW, h: boxH, rectRadius: 0.08,
        fill: { color: o.endFill || C.SUCCESS },
      }, { fontSize: 23, fontFace: FONT_H, color: C.WHITE, bold: true });
      if (o.caption) {
        slide.addText(o.caption, {
          x: x + 0.22, y: capY, w: w - 0.44, h: bottom - capY - 0.06,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
          align: "center", valign: "top", margin: 0,
        });
      }
    },
  };
}

// A check-for-understanding slide with room reserved for a click-revealed
// answer bar. Built locally because the shared cfuSlide grows its question card
// into the space a reveal bar needs.
function checkSlide(pres, cfg) {
  const s = pres.addSlide();
  addTopBar(s, C.ALERT);
  addBadge(s, "CFU", { color: C.ALERT });
  addTitle(s, cfg.title, { color: C.ALERT });

  addTextOnShape(s, "✓  CHECK", {
    x: 8.2, y: 0.20, w: 1.3, h: 0.32, rectRadius: 0.08,
    fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
  }, {
    fontSize: 11, fontFace: FONT_B, color: C.ALERT,
    bold: true, align: "center", valign: "middle", margin: 0,
  });

  addTextOnShape(s, cfg.technique, {
    x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.44, rectRadius: 0.08,
    fill: { color: C.ALERT },
  }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });

  const qY = CONTENT_TOP + 0.58;
  const options = cfg.options || [];
  // Without options the question card grows to fill the space above the
  // revealed answer bar, so the slide never ends on a dead band.
  const qH = options.length ? 1.15 : 2.25;
  addCard(s, 0.5, qY, 9, qH, { strip: C.ALERT });
  if (cfg.lead) {
    s.addText(cfg.lead, {
      x: 0.75, y: qY + 0.10, w: 8.5, h: 0.26,
      fontSize: 13.5, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
    });
  }
  s.addText(cfg.question, {
    x: 0.75, y: qY + (cfg.lead ? 0.38 : 0.14), w: 8.5, h: qH - (cfg.lead ? 0.50 : 0.28),
    fontSize: cfg.questionSize || 30, fontFace: FONT_B, color: C.CHARCOAL,
    bold: Boolean(cfg.questionBold), align: "center", valign: "middle", margin: 0,
  });

  if (options.length) {
    const optY = qY + qH + 0.18;
    const optGap = 0.16;
    const optW = (9 - optGap * (options.length - 1)) / options.length;
    options.forEach((opt, i) => {
      addTextOnShape(s, opt, {
        x: 0.5 + i * (optW + optGap), y: optY, w: optW, h: 0.62, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.5 },
      }, {
        fontSize: 15, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      });
    });
  }

  addFooter(s, FOOTER);
  s.addNotes(cfg.notes);
  return s;
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "Metric Measurement & Conversion",
    "Week 4 Review: one routine, three measurements",
    "Year 5/6 Numeracy | Review Session", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Teacher-facing overview
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.MUTED);
    addBadge(s, "For the teacher", { color: C.MUTED, w: 2.4 });
    addTitle(s, "Week 4 review at a glance", { color: C.MUTED });

    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("What the week taught", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "S1  Choose and order units", options: { bold: true, breakLine: true } },
      { text: "S2  Convert length (km, m, cm, mm)", options: { bold: true, breakLine: true } },
      { text: "S3  Convert mass (t, kg, g, mg)", options: { bold: true, breakLine: true } },
      { text: "S4  Convert capacity (kL, L, mL)", options: { bold: true, breakLine: true } },
      { text: "S5  Choose, convert and compare", options: { bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The routine that ran all week:", options: { color: C.MUTED, breakLine: true } },
      { text: "Set the stairs. Count the steps.", options: { color: C.PRIMARY, breakLine: true } },
      { text: "Multiply or divide.", options: { color: C.PRIMARY, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Down to a smaller unit, multiply.", options: { color: C.SUCCESS, breakLine: true } },
      { text: "Up to a bigger unit, divide.", options: { color: C.ALERT } },
    ], {
      x: 0.7, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addCard(s, 5.1, cardY, 4.4, cardH, { strip: C.ACCENT });
    s.addText("How today works", {
      x: 5.3, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText([
      { text: "Shape: consolidation. Light re-modelling, then interleaved practice.", options: { breakLine: true } },
      { text: "Attributes and directions alternate on purpose, so students decide rather than copy.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "boards check after converting, hinge after comparing, exit ticket.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Catch-up:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "the launch and recap re-ground the whole routine, so a missed session is no barrier.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Flex: cut the multi-step We Do if the compare block runs long.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // 4. Daily Review — multiply and divide by powers of 10
  (() => {
    const s = dailyReviewSlide(pres, "Daily Review: Digits move, the point stays",
      ["4.7 × 100", "350 ÷ 100", "1.25 × 1000"],
      NOTES_DR, FOOTER,
      (slide, lg) => {
        // Card stops at 4.15 so the click-revealed answer bar (top 4.30) keeps
        // its clearance, and every chip inside stays within the card.
        const y0 = lg.panelTopPadded;
        const cardBottom = 4.15;
        addCard(slide, lg.rightX, y0, lg.rightW, cardBottom - y0, { strip: C.ACCENT });
        slide.addText("Place value chart", {
          x: lg.rightX + 0.15, y: y0 + 0.10, w: lg.rightW - 0.30, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
        });
        addPlaceValueChart(slide, lg.rightX + 0.20, y0 + 0.48,
          ["hundreds", "tens", "ones", "tenths", "hundredths"],
          [null, null, null, null, null],
          { w: lg.rightW - 0.40, hdrH: 0.38, valH: 0.56, headerColor: C.PRIMARY });
        slide.addText("Digits move. The point stays.", {
          x: lg.rightX + 0.15, y: y0 + 1.50, w: lg.rightW - 0.30, h: 0.28,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
        });
        addTextOnShape(slide, "× 10  moves left", {
          x: lg.rightX + 0.20, y: y0 + 1.86, w: lg.rightW - 0.40, h: 0.34, rectRadius: 0.07,
          fill: { color: C.SUCCESS },
        }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
        addTextOnShape(slide, "÷ 10  moves right", {
          x: lg.rightX + 0.20, y: y0 + 2.26, w: lg.rightW - 0.40, h: 0.34, rectRadius: 0.07,
          fill: { color: C.ALERT },
        }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
      });
    clickBuild(s, [
      () => { addRevealAnswerBar(s, ["470", "3.5", "1250"], { color: C.SUCCESS }); },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 5. Fluency — division bracket
  (() => {
    const s = fluencySlide(pres, "Fluency: Division bracket",
      ["224 ÷ 7", "195 ÷ 5", "152 ÷ 8"],
      NOTES_FLUENCY, FOOTER);
    clickBuild(s, [
      () => { addRevealAnswerBar(s, ["32", "39", "19"], { color: C.SUCCESS }); },
    ]);
  })();

  // 6. Launch — what is the same about all three?
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Launch", { color: C.SECONDARY });
    addTitle(s, "What is the same about all three?", { color: C.SECONDARY });

    const facts = [
      { label: "LENGTH", fact: "1 m = 100 cm" },
      { label: "MASS", fact: "1 kg = 1000 g" },
      { label: "CAPACITY", fact: "1 L = 1000 mL" },
    ];
    facts.forEach((f, i) => {
      const x = 0.5 + i * 3.05;
      addCard(s, x, CONTENT_TOP, 2.9, 2.05, { strip: C.SECONDARY });
      s.addText(f.label, {
        x: x + 0.12, y: CONTENT_TOP + 0.14, w: 2.66, h: 0.28,
        fontSize: 13, fontFace: FONT_B, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      s.addText(f.fact, {
        x: x + 0.12, y: CONTENT_TOP + 0.52, w: 2.66, h: 1.38,
        fontSize: 24, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    addTextOnShape(s, "Turn and tell your partner. What is the same?", {
      x: 0.5, y: 3.52, w: 9, h: 0.46, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, { fontSize: 17, fontFace: FONT_B, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    clickBuild(s, [
      () => {
        addTextOnShape(s, "One routine works for all three: set the stairs, count the steps, multiply or divide.", {
          x: 0.5, y: 4.16, w: 9, h: 0.88, rectRadius: 0.1,
          fill: { color: C.SUCCESS },
        }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 7. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning to use one routine to convert and compare measurements of length, mass and capacity.",
    [
      "I can set the stairs and say which unit is bigger.",
      "I can convert between units by multiplying or dividing.",
      "I can make the units match to compare two measurements and explain my thinking.",
    ],
    NOTES_LI_SC, FOOTER);

  // 8. Key vocabulary — the one word that carries the week
  keyWordSlide(pres, {
    word: "equivalent",
    meaning: "Equal in value. The same amount, written in a different unit.",
    example: "1.25 m and 125 cm are equivalent. One length, written two ways.",
    routine: ["Say it", "Show it", "Use it"],
    color: C.PRIMARY,
  }, NOTES_VOCAB, FOOTER);

  // 9. Recap — the routine on all three staircases (catch-up anchor)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Recap", { color: C.PRIMARY });
    addTitle(s, "Our routine, on every staircase", { color: C.PRIMARY });

    const moves = ["1.  Set the stairs", "2.  Count the steps", "3.  Multiply or divide"];
    moves.forEach((m, i) => {
      addTextOnShape(s, m, {
        x: 0.5 + i * 3.05, y: CONTENT_TOP, w: 2.9, h: 0.48, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    });

    const stairTop = CONTENT_TOP + 0.62;
    staircase(s, 0.5, stairTop, 2.9, SAFE_BOTTOM, {
      title: "Length", color: C.PRIMARY,
      steps: [{ u: "km" }, { u: "m", f: 1000 }, { u: "cm", f: 100 }, { u: "mm", f: 10 }],
    });
    staircase(s, 3.55, stairTop, 2.9, SAFE_BOTTOM, {
      title: "Mass", color: C.SECONDARY,
      steps: [{ u: "tonne" }, { u: "kg", f: 1000 }, { u: "g", f: 1000 }, { u: "mg", f: 1000 }],
    });
    staircase(s, 6.6, stairTop, 2.9, SAFE_BOTTOM, {
      title: "Capacity", color: C.ACCENT,
      steps: [{ u: "kL" }, { u: "L", f: 1000 }, { u: "mL", f: 1000 }],
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_RECAP);
    runSlideDiagnostics(s, pres);
  })();

  // 10. I Do — convert down (length)
  (() => {
    let panel;
    const s = workedExSlide(pres, 2, "I Do", "Convert 2.4 km into metres",
      [
        "Set the stairs: km, m, cm, mm.",
        "Count the steps: km to m is one step.",
        "That step is × 1000.",
        "Going DOWN to a smaller unit, so multiply.",
        "2.4 × 1000 = ?",
      ],
      NOTES_IDO1, FOOTER,
      (slide, lg) => {
        panel = convertPanel(slide, lg, {
          title: "Down one step", start: "2.4 km", factor: "1000",
          end: "2400 m", caption: "Same length, new unit.",
        });
      });
    clickBuild(s, [() => panel.revealEnd()]);
    runSlideDiagnostics(s, pres);
  })();

  // 11. I Do — convert up (mass)
  (() => {
    let panel;
    const s = workedExSlide(pres, 2, "I Do", "Convert 750 g into kilograms",
      [
        "Set the stairs: tonne, kg, g, mg.",
        "Count the steps: g to kg is one step.",
        "Every mass step is × 1000.",
        "Going UP to a bigger unit, so divide.",
        "750 ÷ 1000 = ?",
      ],
      NOTES_IDO2, FOOTER,
      (slide, lg) => {
        panel = convertPanel(slide, lg, {
          title: "Up one step", start: "750 g", factor: "1000", divide: true,
          end: "0.75 kg", caption: "Fewer kilograms than grams, so the number shrank.",
        });
      });
    clickBuild(s, [() => panel.revealEnd()]);
    runSlideDiagnostics(s, pres);
  })();

  // 12. CFU — decision point one
  (() => {
    const s = checkSlide(pres, {
      title: "Convert it",
      technique: "Show Me Boards",
      question: "Convert 3.5 L into millilitres.",
      questionBold: true,
      notes: NOTES_CFU1,
    });
    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, "3500 mL.  One step down, so × 1000.",
          { color: C.SUCCESS, label: "Answer", fontSize: 22, y: 4.25, h: 0.80 });
      },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 13. We Do — convert down (mass, decimal)
  (() => {
    let panel;
    const s = workedExSlide(pres, 3, "We Do", "Convert 1.75 kg into grams",
      [
        "With your partner.",
        "",
        "Set the stairs.",
        "Count the steps: kg to g is one step.",
        "Bigger unit to smaller unit, so which operation?",
      ],
      NOTES_WEDO1, FOOTER,
      (slide, lg) => {
        panel = convertPanel(slide, lg, {
          title: "Your turn", start: "1.75 kg", factor: "1000",
          end: "1750 g", caption: "Every digit moves three places.",
        });
      });
    clickBuild(s, [() => panel.revealEnd()]);
    runSlideDiagnostics(s, pres);
  })();

  // 14. We Do — convert up (capacity)
  (() => {
    let panel;
    const s = workedExSlide(pres, 3, "We Do", "Convert 4200 mL into litres",
      [
        "With your partner.",
        "",
        "Set the stairs.",
        "Count the steps: mL to L is one step.",
        "Smaller unit to bigger unit, so which operation?",
      ],
      NOTES_WEDO2, FOOTER,
      (slide, lg) => {
        panel = convertPanel(slide, lg, {
          title: "Your turn", start: "4200 mL", factor: "1000", divide: true,
          end: "4.2 L", caption: "Not 42. Three places, not two.",
        });
      });
    clickBuild(s, [() => panel.revealEnd()]);
    runSlideDiagnostics(s, pres);
  })();

  // 15. I Do — make the units match, then compare
  (() => {
    let panel;
    const s = workedExSlide(pres, 2, "I Do", "Which is heavier: 0.8 kg or 850 g?",
      [
        "Different units, so I cannot compare yet.",
        "Make the units match first.",
        "",
        "Convert 0.8 kg down into grams.",
        "0.8 × 1000 = ?",
      ],
      NOTES_IDO3, FOOTER,
      (slide, lg) => {
        panel = convertPanel(slide, lg, {
          title: "Make the units match", start: "0.8 kg", factor: "1000",
          end: "800 g", caption: "800 g is less than 850 g, so 850 g is heavier.",
        });
      });
    clickBuild(s, [() => panel.revealEnd()]);
    runSlideDiagnostics(s, pres);
  })();

  // 16. We Do — compare (length)
  (() => {
    let panel;
    const s = workedExSlide(pres, 3, "We Do", "Which is longer: 2.5 m or 240 cm?",
      [
        "With your partner.",
        "",
        "Step 1: convert 2.5 m into centimetres.",
        "Step 2: compare the two lengths.",
        "Step 3: write which one is longer.",
      ],
      NOTES_WEDO3, FOOTER,
      (slide, lg) => {
        panel = convertPanel(slide, lg, {
          title: "Make the units match", start: "2.5 m", factor: "100",
          end: "250 cm", caption: "250 cm is more than 240 cm, so 2.5 m is longer.",
        });
      });
    clickBuild(s, [() => panel.revealEnd()]);
    runSlideDiagnostics(s, pres);
  })();

  // 17. We Do — multi-step (capacity). Flex slide.
  (() => {
    let panel;
    const s = workedExSlide(pres, 3, "We Do", "2 L bottle, drink 650 mL. How much is left?",
      [
        "With your partner.",
        "",
        "Step 1: convert 2 L into millilitres.",
        "Step 2: subtract 650.",
        "Step 3: write the amount left, in mL.",
      ],
      NOTES_WEDO4, FOOTER,
      (slide, lg) => {
        panel = convertPanel(slide, lg, {
          title: "Convert, then subtract", start: "2 L", factor: "1000",
          end: "2000 mL", caption: "2000 take away 650 leaves 1350 mL.",
        });
      });
    clickBuild(s, [() => panel.revealEnd()]);
    runSlideDiagnostics(s, pres);
  })();

  // 18. CFU hinge — decision point two
  (() => {
    const s = checkSlide(pres, {
      title: "Spot the error",
      technique: "Show Me Boards",
      lead: "A student wrote:",
      question: "2.5 kg = 25 g",
      questionBold: true,
      questionSize: 34,
      options: [
        "A. Correct",
        "B. Wrong: it is 250 g",
        "C. Wrong: it is 2500 g",
      ],
      notes: NOTES_CFU2,
    });
    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, "C.  2.5 × 1000 = 2500 g",
          { color: C.ALERT, label: "Check", fontSize: 24, y: 4.25, h: 0.80 });
      },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 19. You Do — review practice sheet
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: review practice", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.15, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 1, set the stairs and convert.  ", options: { fontSize: 16.5, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 2, compare by matching the units.  ", options: { fontSize: 16.5, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 3, solve the problems.", options: { fontSize: 16.5, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.14, w: 8.5, h: 0.87,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.32;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Every question, the same three moves", {
      x: 0.7, y: panelY + 0.14, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "1. Set the stairs.  2. Count the steps.  3. Multiply or divide.", {
      x: 1.0, y: panelY + 0.60, w: 8.0, h: 0.46, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "Comparing? Make the units MATCH first.", {
      x: 1.0, y: panelY + 1.18, w: 8.0, h: 0.46, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "Finished? Try the Challenge box at the end.", {
      x: 1.0, y: panelY + 1.76, w: 8.0, h: 0.46, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // 20. Exit Ticket
  exitTicketSlide(pres,
    [
      "Convert 2400 g into kilograms.",
      "Which is longer: 1.6 m or 150 cm? Convert first, then explain how you know.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 21. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the one rule you will never forget about converting units?",
      scItems: [
        "I can set the stairs and say which unit is bigger.",
        "I can convert between units by multiplying or dividing.",
        "I can make the units match to compare two measurements and explain my thinking.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MMC Week 4 Review.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  await generatePdfs();
  console.log("Week 4 Review build complete.");
}

// ─── PDFs ───────────────────────────────────────────────────────────────────

// Draw one metric staircase on paper: unit boxes largest at the top, with the
// multiply and divide factor printed between each step. The slide anchor's twin,
// so the scaffold shows the model rather than describing it.
function drawStaircasePdf(doc, x, y, w, opts) {
  const o = opts || {};
  const steps = o.steps;
  const color = o.color;
  const boxH = 26;
  const gap = 26;

  doc.fontSize(11).font("Sans-Bold").fillColor(hex(color));
  doc.text(o.title, x, y, { width: w, align: "center" });

  let cy = y + 18;
  steps.forEach((st, i) => {
    if (i > 0) {
      doc.fontSize(8.5).font("Sans-Bold").fillColor(hex("227545"));
      doc.text("x " + st.f, x, cy - gap + 8, { width: w / 2 - 2, align: "right" });
      doc.fontSize(8.5).font("Sans-Bold").fillColor(hex("A52A2A"));
      doc.text("div " + st.f, x + w / 2 + 2, cy - gap + 8, { width: w / 2 - 2, align: "left" });
    }
    doc.save();
    doc.roundedRect(x, cy, w, boxH, 4).fill(hex(i === 0 ? color : "2F6A88"));
    doc.restore();
    doc.fontSize(12).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(st.u, x, cy + 7, { width: w, align: "center" });
    cy += boxH + gap;
  });

  return cy - gap;
}

async function generatePdfs() {
  const STAIR_TIP =
    "Set the stairs. Count the steps. Multiply or divide. " +
    "DOWN to a smaller unit, multiply. UP to a bigger unit, divide. " +
    "Length steps: x1000, x100, x10. Mass and capacity: every step x1000.";

  // ── Review practice sheet ────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Convert, compare and solve across length, mass and capacity.",
      color: C.PRIMARY,
      lessonInfo: "Week 4 Review | Year 5/6 Numeracy",
    });

    y = addTipBox(doc, STAIR_TIP, y, { color: C.ACCENT });

    // Each item is a label followed by one ruled writing line, never a label
    // with its own underscores AND a rule - that prints two answer spaces.
    const writeItem = (label) => { y = addWriteLine(doc, label, y) + 3; };

    y = addSectionHeading(doc, "Section 1 - Set the stairs and convert", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "Worked for you:   5 m = 500 cm     (m to cm is one step down, so 5 x 100 = 500)",
      y, { italic: true, color: "5F6E7A" });
    writeItem("a)   Convert 4 km into metres:");
    writeItem("b)   Convert 2500 g into kilograms:");
    writeItem("c)   Convert 1.4 L into millilitres:");
    writeItem("d)   Convert 0.6 kg into grams:");

    y = addSectionHeading(doc, "Section 2 - Make the units match, then compare", y, { color: C.PRIMARY });
    writeItem("a)   Which is longer, 3 m or 280 cm?   Longer:");
    writeItem("b)   Which is heavier, 0.4 kg or 450 g?   Heavier:");
    writeItem("c)   Which is more, 1.2 L or 1150 mL?   More:");

    y = addSectionHeading(doc, "Section 3 - Solve the problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  A ribbon is 4 m long. You cut off 90 cm. How many centimetres are left?", y);
    writeItem("      Working:");
    writeItem("      Answer:");
    y = addBodyText(doc, "b)  A jug holds 1.5 L. You pour out 350 mL. How many millilitres are left?", y);
    writeItem("      Working:");
    writeItem("      Answer:");

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Order these from smallest to largest. Convert them all to one unit first:   1.3 m,   125 cm,   1400 mm", y);
    writeItem("      Smallest to largest:");

    addPdfFooter(doc, "Week 4 Review | Metric Measurement | Year 5/6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // ── Answer key ───────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Week 4 review practice sheet.",
      color: C.PRIMARY,
      lessonInfo: "Week 4 Review | Year 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "Section 1 - Set the stairs and convert", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4000 m   (km to m is one step down, 4 x 1000)", y);
    y = addBodyText(doc, "b)  2.5 kg   (g to kg is one step up, 2500 div 1000)", y);
    y = addBodyText(doc, "c)  1400 mL   (L to mL is one step down, 1.4 x 1000)", y);
    y = addBodyText(doc, "d)  600 g   (kg to g is one step down, 0.6 x 1000)", y);

    y = addSectionHeading(doc, "Section 2 - Compare", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3 m is longer.   3 m = 300 cm, and 300 is more than 280.", y);
    y = addBodyText(doc, "b)  450 g is heavier.   0.4 kg = 400 g, and 400 is less than 450.", y);
    y = addBodyText(doc, "c)  1.2 L is more.   1.2 L = 1200 mL, and 1200 is more than 1150.", y);

    y = addSectionHeading(doc, "Section 3 - Solve", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4 m = 400 cm, then 400 - 90 = 310 cm left.", y);
    y = addBodyText(doc, "b)  1.5 L = 1500 mL, then 1500 - 350 = 1150 mL left.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Convert to centimetres: 1.3 m = 130 cm, 125 cm stays, 1400 mm = 140 cm. " +
      "Smallest to largest: 125 cm, then 1.3 m, then 1400 mm.", y);

    y = addTipBox(doc,
      "Watch for: comparing the raw numbers without converting (Section 2 is where this shows up); " +
      "using the length step of 100 on a mass or capacity conversion; and subtracting before converting in Section 3. " +
      "Section 1 is the re-grounding task - a student who missed sessions should be able to do it from the staircase alone.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, "Week 4 Review | Answer Key | Year 5/6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // ── Staircase Helper (enabling scaffold + glue-in reference) ─────────────
  await (async () => {
    const doc = createPdf({ title: SCAFFOLD_RES.name });
    let y = addPdfHeader(doc, SCAFFOLD_RES.name, {
      subtitle: "Set the stairs. Count the steps. Multiply or divide.",
      color: C.SECONDARY,
      lessonInfo: "Week 4 Review | Year 5/6 Numeracy",
    });

    const colW = (PAGE.CONTENT_W - 40) / 3;
    const stairY = y + 6;
    const bottoms = [
      drawStaircasePdf(doc, PAGE.MARGIN, stairY, colW, {
        title: "LENGTH", color: C.PRIMARY,
        steps: [{ u: "km" }, { u: "m", f: 1000 }, { u: "cm", f: 100 }, { u: "mm", f: 10 }],
      }),
      drawStaircasePdf(doc, PAGE.MARGIN + colW + 20, stairY, colW, {
        title: "MASS", color: C.SECONDARY,
        steps: [{ u: "tonne" }, { u: "kg", f: 1000 }, { u: "g", f: 1000 }, { u: "mg", f: 1000 }],
      }),
      drawStaircasePdf(doc, PAGE.MARGIN + 2 * (colW + 20), stairY, colW, {
        title: "CAPACITY", color: C.ACCENT,
        steps: [{ u: "kL" }, { u: "L", f: 1000 }, { u: "mL", f: 1000 }],
      }),
    ];
    y = Math.max.apply(null, bottoms) + 16;

    y = addTipBox(doc,
      "Going DOWN the stairs to a smaller unit, you need MORE of them, so the number gets bigger: MULTIPLY. " +
      "Going UP the stairs to a bigger unit, you need FEWER of them, so the number gets smaller: DIVIDE.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "One worked step for each staircase", y, { color: C.SECONDARY });
    y = addBodyText(doc, "LENGTH      3 m into cm.   One step down, x 100.   3 x 100 = 300 cm.", y);
    y = addBodyText(doc, "MASS        2 kg into g.    One step down, x 1000.  2 x 1000 = 2000 g.", y);
    y = addBodyText(doc, "CAPACITY    5000 mL into L. One step up, div 1000.  5000 div 1000 = 5 L.", y);

    y = addSectionHeading(doc, "Now your turn - the stairs are drawn for you above", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "For each one: how many steps, multiply or divide, then the answer.", y, { italic: true, color: "5F6E7A" });
    y = addWriteLine(doc, "a)   Convert 6 m into centimetres:", y) + 8;
    y = addWriteLine(doc, "b)   Convert 4 kg into grams:", y) + 8;
    y = addWriteLine(doc, "c)   Convert 3000 mL into litres:", y) + 8;

    y = addTipBox(doc,
      "Teacher answers: a) 6 x 100 = 600 cm.  b) 4 x 1000 = 4000 g.  c) 3000 div 1000 = 3 L. " +
      "Glue this card into the front of the maths book so the staircases are there every session.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, "Week 4 Review | Staircase Helper | Year 5/6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, SCAFFOLD_RES.fileName));
    console.log("PDF written: " + SCAFFOLD_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
