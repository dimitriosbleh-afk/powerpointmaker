"use strict";

// Part Part Whole Unit — Session 4: Lolly Jars
// Foundation Numeracy | Term 2 Week 4 | Variant 3
// DR: Digit formation — teen numbers 11-20
// Fluency: Counting forwards from various starting points
// VC2MFN04 — partition and combine collections up to 10 using part-part-whole relationships
// This session is the unit assessment — the Lolly Jars activity is the primary exit evidence.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", weekToVariant(4));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = "Part, Part, Whole | Session 4 of 4 | Foundation Numeracy";
const OUT_DIR = "output/PPW_Session4_Lolly_Jars";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lolly Jars Worksheet", "A3 Lolly Jars assessment — how many ways can you make 10?");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "All 11 ways to split 10 lollies across two jars.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Lolly Jars for a whole of 5, with some lollies pre-drawn.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Three-jar Lolly Shop: split 10 lollies across 3 jars.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a lolly jar with `lollies` circular lollies inside. Simple 2D jar shape.
function drawLollyJar(slide, x, y, w, h, lollies, opts) {
  const o = opts || {};
  const neckH = h * 0.12;
  const lidH = h * 0.07;
  const bodyH = h - neckH - lidH;
  const bodyY = y + neckH + lidH;
  const jarColor = o.jarColor || "C4E6F5";
  const rimColor = o.rimColor || "6FA3D2";
  const lollyColor = o.lollyColor || "E63946";

  // Lid
  slide.addShape("roundRect", {
    x: x + w * 0.15, y: y, w: w * 0.70, h: lidH, rectRadius: 0.04,
    fill: { color: rimColor },
    line: { color: rimColor, width: 1 },
  });
  // Neck
  slide.addShape("rect", {
    x: x + w * 0.20, y: y + lidH, w: w * 0.60, h: neckH,
    fill: { color: jarColor },
    line: { color: rimColor, width: 1 },
  });
  // Body (rounded rect)
  slide.addShape("roundRect", {
    x, y: bodyY, w, h: bodyH, rectRadius: 0.12,
    fill: { color: jarColor },
    line: { color: rimColor, width: 1.5 },
  });

  // Lollies: arrange in 2 rows up to 5 per row
  const cols = 5;
  const rows = 2;
  const lollySize = Math.min((w * 0.75) / cols, bodyH * 0.4);
  const lollyX0 = x + (w - cols * lollySize) / 2;
  const lollyY0 = bodyY + bodyH * 0.12;
  const gapX = 0.04;
  const gapY = 0.04;
  for (let i = 0; i < Math.min(lollies, 10); i++) {
    const r = Math.floor(i / cols);
    const c = i % cols;
    const cx = lollyX0 + c * (lollySize + gapX) + lollySize / 2;
    const cy = lollyY0 + r * (lollySize + gapY) + lollySize / 2;
    slide.addShape("roundRect", {
      x: cx - lollySize / 2, y: cy - lollySize / 2,
      w: lollySize - gapX, h: lollySize - gapY, rectRadius: (lollySize - gapX) / 2,
      fill: { color: lollyColor },
      line: { color: C.CHARCOAL, width: 0.8 },
    });
  }
}

// Draw a big number card
function drawNumberCard(slide, x, y, w, h, num, color) {
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.15,
    fill: { color: C.WHITE },
    line: { color: color || C.ACCENT, width: 4 },
  });
  const fontSize = h > 1.5 ? 110 : 70;
  slide.addText(String(num), {
    x, y, w, h,
    fontSize, fontFace: FONT_H, color: color || C.ACCENT, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome to our last Part Part Whole lesson!
- Today you have a very special job
- You are the owner of a lolly shop
- Your job is to split 10 lollies across two lolly jars - as many different ways as you can

DO:
- Have A3 Lolly Jars worksheets ready (one per student)
- Have counters or unifix cubes in 3 colours available
- Display title slide as students settle

TEACHER NOTES:
Session 4 of 4. This is the unit assessment. The Lolly Jars activity pulls together everything from Sessions 1-3. Keep the "lolly shop owner" framing fun and active.

WATCH FOR:
- Students who are excited - they are bringing confidence from the unit
- Students who seem unsure - reassure that they have all the tools they need

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up our pencils! Watch the slide.
- I will show you a teen number. You write it on your whiteboard.
- Neat numbers, please!

DO:
- Display each teen number one at a time
- Students write on whiteboards
- Check formation: 1s straight, teens go tens-then-ones

TEACHER NOTES:
Daily Review on teen digit formation. Foundation students may form their teen numbers incorrectly (writing "41" for "14" because they say "four-teen"). Watch the order.

WATCH FOR:
- Students writing "41" for "14" - classic reversal, reteach "ten first, then extras"
- Students forming the digits correctly - writing fluency is secure

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your writing
- Did you write the tens digit first? That is the 1 on the left.
- Then the ones digit? That is the extra.
- Tick if correct. Fix it if you need to.

DO:
- Reveal each teen number in large print
- Students compare their whiteboard to the slide
- Note students who reversed the digits

TEACHER NOTES:
Tick-and-fix builds self-correction habits. Reversed teens (41 for 14) is a normal Foundation error - it means they are saying the number in their head but writing left-to-right from the ones word.

WATCH FOR:
- Students self-correcting - writing habits are building
- Students who fix without prompting - strong metacognition

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Counting fluency! Today we count from different starting numbers.
- Start at 5. Count to 15.
- Now start at 8. Count to 20.
- Now start at 13. Count to 20.

DO:
- Call each starting number clearly
- Lead choral count to the target
- Keep rhythm with claps or finger tapping

TEACHER NOTES:
Counting from different starting points builds flexibility. Foundation students often know 1-20 only from 1. Starting mid-sequence is genuinely harder and reveals whether the count is automated.

WATCH FOR:
- Students who join mid-count fluently - counting is secure
- Students who silently count up to the start first - they need more practice
- Students who stumble after 10 - reteach the decade shift

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Our learning intention is the same as all week
- Today we show what we learned
- We partition numbers to 10 into two parts
- Today's parts are LOLLIES in JARS!

DO:
- Choral read LI and SC
- Remind: this is the last day; today they show what they know

TEACHER NOTES:
Same LI and SC. Students now meeting or exceeding SC2 across the unit have solid PPW understanding. The Lolly Jars activity today generates strong evidence of SC2 and SC3.

WATCH FOR:
- Students who read along confidently - unit vocabulary is secure
- Students who know "parts" and "whole" in their own words - concept mastery

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. I have a lolly shop. My whole is 10 lollies.
- I need to put them into 2 jars.
- Watch. I put 6 lollies in Jar A. I put 4 lollies in Jar B.
- Count: 6 and 4. That is 10.
- Whole = 10. Parts = 6 and 4.

DO:
- Use counters or unifix cubes in 2 colours under the document camera
- Place 6 counters on one side, 4 on the other
- Point to the slide visual as you talk

TEACHER NOTES:
The lolly jar frame is a new context but the PPW concept is the same. Lollies are the counters. Jars are the "groups" or "parts". Students who mastered tens frames should transfer easily.

MISCONCEPTIONS:
- Misconception: Students add lollies in each jar but then add an extra somewhere else
  Why: They are trying to make each jar feel "fair" or similar size
  Impact: The whole changes; they lose track of 10
  Quick correction: "Let's count the total. How many in Jar A? [6] How many in Jar B? [4]. 6 and 4 is 10. That matches our whole."

WATCH FOR:
- Students who say "6 and 4 is 10" - they are applying addition fluently
- Students who want to count every lolly from 1 - let them; it confirms the total

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Let's split the 10 lollies a DIFFERENT way
- Same 10 lollies. Same two jars.
- Watch. 7 in Jar A. 3 in Jar B.
- Count: 7 and 3. Still 10.
- Same whole. Different parts.

DO:
- Rearrange: 7 on one side, 3 on the other
- Write on the board: "10 = 6 + 4" and "10 = 7 + 3"
- Emphasise: The whole does not change

TEACHER NOTES:
Second I Do shows flexibility. By showing TWO ways, you invite the lolly shop owner thinking: "there are many ways to do this job." That sets up the You Do perfectly.

WATCH FOR:
- Students who call out more ways before you ask - generalisation is happening
- Students who look relieved when they see repetition - building confidence

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Look at the two jars
- Jar A has some lollies. Jar B has some lollies.
- The whole is 10.
- How many lollies are in Jar B?
- Show me on your fingers.

DO:
- Display: Jar A with 8 lollies, Jar B with 2
- Scan fingers
- Cold call 2 students for the parts

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "Jar A has 8. The whole is 10. How many in Jar B? Show me."
- Scan for: 2 fingers
PROCEED: If 80%+ show 2, move to We Do
PIVOT: Most likely misconception - students count both jars together and give 10. Reteach: "Count ONLY Jar B. Do not count Jar A. How many dots in JUST Jar B?"

TEACHER NOTES:
This CFU tests missing-part thinking: given the whole and one part, find the other. This is the foundation of subtraction. Foundation students at this point should be able to do it with a visual.

WATCH FOR:
- Students showing 2 fingers quickly - strong PPW reasoning
- Students showing 10 fingers - they are counting the whole; reteach

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. I have filled Jar A with 5 lollies.
- The whole is 10. How many lollies in Jar B?
- Work it out with your partner on your whiteboard.
- Write: 10 = 5 and ___

DO:
- Show Jar A with 5 lollies, Jar B empty
- Partners discuss and write
- Allow 30 seconds
- Cold call 3 pairs for their answer

TEACHER NOTES:
We Do is missing-addend work using the jar metaphor. Partners support each other. Keep the pace quick.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 4 Enabling Scaffold - Lolly Jars for a whole of 5 (not 10). Simpler numbers to build confidence.
- Extra Notes: Some lollies pre-drawn; students complete to 5.
EXTENDING PROMPT:
- Task: Session 4 Extension - 3 jars for a whole of 10. Can you split 10 lollies across 3 jars?
- Extra Notes: Open-ended; many correct answers.

WATCH FOR:
- Students who write "5 and 5" confidently - strong PPW
- Students who need help - offer counters or pair them with a stronger partner

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. Jar A has 5. Whole is 10.
- So Jar B needs 5 more.
- 10 = 5 and 5.
- Ask: Could Jar A have 5 AND Jar B have 6? Why or why not?

DO:
- Reveal the answer
- Accept that 5+5 is the one right answer for THIS jar setup
- Emphasise: Once one part is set, the other part has only one answer

TEACHER NOTES:
Missing-addend work has exactly one right answer once the whole and one part are fixed. This is a subtle but important idea.

WATCH FOR:
- Students who argue "6 would be too many" - they understand the constraint
- Students who say "both" - redirect to the whole

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. I say: Jar A has 4 lollies. Jar B has 7 lollies. The whole is 10.
- Thumbs up if I am right. Thumbs down if I am wrong.

DO:
- Scan thumbs
- Follow up: "How do you know?" Cold call 2 students
- Correct answer: thumbs down (4 + 7 = 11, not 10)

CFU CHECKPOINT:
Technique: Thumbs Up/Down
Script:
- Say: "Jar A has 4, Jar B has 7, whole is 10. Right or wrong?"
- Scan for: thumbs down
PROCEED: If 80%+ show thumbs down, move to You Do
PIVOT: Most likely misconception - students add 4 and 7 and get 11 but say "close enough to 10." Reteach: "The whole is 10 EXACTLY. 4 and 7 is 11. That is one too many. So this is wrong."

TEACHER NOTES:
This negative-example hinge tests whether students can reject an incorrect PPW even when it "looks reasonable". By this point in the unit, most students should be able to check the sum.

WATCH FOR:
- Confident thumbs down - students are checking the sum
- Thumbs up - students are trusting the teacher instead of the numbers

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- You are the lolly shop owner!
- Your job: take 10 lollies. Split them across 2 jars.
- Find as many different ways as you can.
- Draw each way on your big worksheet.
- You have 15 minutes. Get ready to shop!

DO:
- Hand out the Session 4 Lolly Jars A3 worksheet
- Each student has 10 counters or unifix cubes
- Circulate: start with students who needed support in We Do
- Enabler students get the Session 4 Enabling Scaffold (whole of 5)
- Extender students get the Session 4 Extension (3 jars)
- Take photos or notes on who finds how many ways

TEACHER NOTES:
This is the unit assessment. You are collecting evidence of SC2 (finding different ways) and SC3 (explaining parts and whole). Aim for 3+ ways as a core target. 5+ ways is strong. 11 ways is mastery.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 4 Enabling Scaffold - 2 jars, whole of 5. Some lollies pre-drawn. Students complete each jar pair to 5.
- Extra Notes: Lower cognitive load, same routine.
EXTENDING PROMPT:
- Task: Session 4 Extension - 3 jars, whole of 10. How many different ways can you split 10 across 3 jars?
- Extra Notes: Open exploration; record every combination found.

WATCH FOR:
- Students recording more than 3 ways - strong application
- Students repeating the same pair in a different order (e.g. 5+5 and 5+5) - prompt: "Is that a different way?"
- Students using the order (3+7 and 7+3) as different ways - Foundation-appropriate; accept either way

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. This is our last check!
- Two quick questions.
- Work on your own. 2 minutes.

DO:
- Display the exit ticket
- Students respond on whiteboards or workbook
- Collect or scan - this data supports unit reporting

TEACHER NOTES:
Exit ticket assesses the unit. Q1 checks SC1 (identify parts). Q2 checks SC2 (create a way). Q3 checks SC3 (explain). Combined with the Lolly Jars worksheet, this is rich assessment evidence.

WATCH FOR:
- Students who write confidently and quickly - ready to move to bonds to 10 in the next unit
- Students who still need prompting - note for Term 3 intervention

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's look back at our week.
- Thumbs up, sideways, or down for each criterion.
- Tell your partner: What is your favourite way to show parts and whole?
- Is it counters? Towers? Tens frames? Lolly jars?

DO:
- Display SC
- Thumbs check each
- 30 seconds Turn and Talk
- Cold call 3 students to share - record their preferences

TEACHER NOTES:
Closing the unit. Students reflect on the tools they used. Their preferences tell you something about how they think - some like the discrete feel of counters, some like the structure of tens frames, some like the visual of jars.

WATCH FOR:
- Students who can compare tools and explain their preference - deep metacognition
- Students who pick one favourite - that is great; they are building identity as a learner

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Today's resources
- The Lolly Jars Worksheet is the main A3 assessment
- Enabler and extension are on the differentiation shelf

DO:
- Point out each resource

TEACHER NOTES:
Materials needed: 10 counters or unifix cubes per student, A3 Lolly Jars worksheet, whiteboards. Print the Lolly Jars Worksheet A3 for impact.

WATCH FOR:
- N/A - teacher reference only

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 4: Lolly Jars Shop",
    "Foundation Numeracy | Session 4 of 4 | Term 2 Week 4", NOTES_TITLE);

  // Slide 2-3: Daily Review — teen digit formation (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Writing Teen Numbers", { y: 0.65, fontSize: 24, color: STAGE_COLORS["1"] });

      // Left: instructions
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Write each teen number", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Tens digit FIRST", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Ones digit next", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Listen carefully!", options: { fontSize: 15, bold: true, color: C.ALERT } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: three words to write
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      const words = [
        { label: "A.", word: "fourteen" },
        { label: "B.", word: "seventeen" },
        { label: "C.", word: "twenty" },
      ];
      words.forEach((w, i) => {
        const rowY = CONTENT_TOP + 0.2 + i * 1.0;
        s.addShape("roundRect", {
          x: 5.4, y: rowY, w: 3.9, h: 0.8, rectRadius: 0.08,
          fill: { color: C.BG_CARD },
          line: { color: C.SECONDARY, width: 2 },
        });
        s.addText(w.label + "  " + w.word, {
          x: 5.4, y: rowY, w: 3.9, h: 0.8,
          fontSize: 24, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A) 14     B) 17     C) 20", {
        x: 0.8, y: 4.55, w: 8.4, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — count from various starting points
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Count from Different Starts", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFluency.addText([
    { text: "Stand up!", options: { fontSize: 20, bold: true, color: C.ALERT, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Start at the number I say", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
    { text: "Count up together", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
    { text: "Stop when I clap!", options: { bullet: true, fontSize: 15, color: C.CHARCOAL } },
  ], {
    x: 0.75, y: CONTENT_TOP + 0.2, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  // Right: three start-stop cards
  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
  const starts = [
    { from: 5, to: 15 },
    { from: 8, to: 20 },
    { from: 13, to: 20 },
  ];
  starts.forEach((st, i) => {
    const rowY = CONTENT_TOP + 0.2 + i * 1.0;
    sFluency.addShape("roundRect", {
      x: 5.4, y: rowY, w: 3.9, h: 0.8, rectRadius: 0.08,
      fill: { color: C.WHITE },
      line: { color: C.SECONDARY, width: 2 },
    });
    sFluency.addText(String(st.from), {
      x: 5.5, y: rowY, w: 0.9, h: 0.8,
      fontSize: 28, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    sFluency.addText("to", {
      x: 6.4, y: rowY, w: 0.9, h: 0.8,
      fontSize: 16, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });
    sFluency.addText(String(st.to), {
      x: 7.3, y: rowY, w: 0.9, h: 0.8,
      fontSize: 28, fontFace: FONT_H, color: C.ALERT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    // Arrow visual on right
    sFluency.addText(">", {
      x: 8.3, y: rowY, w: 0.9, h: 0.8,
      fontSize: 28, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to partition numbers to 10 into two parts using materials"],
    [
      "I can show a number as two parts",
      "I can find different ways to make the same number",
      "I can use pictures, objects, or numbers to explain my thinking",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Lolly Jars, first way
  workedExSlide(pres, 2, "I Do", "The Lolly Shop: 10 Lollies",
    [
      "I have 10 lollies. That is my WHOLE.",
      "",
      "I put them in 2 jars",
      "",
      "Jar A: 6 lollies",
      "Jar B: 4 lollies",
      "",
      "6 and 4 make 10",
      "",
      "10 = 6 + 4",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.5, { strip: C.PRIMARY });

      // Header
      slide.addText("Whole = 10 lollies", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Two jars side by side
      drawLollyJar(slide, lg.rightX + 0.15, lg.panelTopPadded + 0.5, 1.75, 2.2, 6, { lollyColor: "E63946" });
      drawLollyJar(slide, lg.rightX + 2.10, lg.panelTopPadded + 0.5, 1.75, 2.2, 4, { lollyColor: "F4C430" });

      slide.addText("Jar A = 6", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.75, w: 1.75, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: "E63946", bold: true, margin: 0, align: "center",
      });
      slide.addText("Jar B = 4", {
        x: lg.rightX + 2.10, y: lg.panelTopPadded + 2.75, w: 1.75, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: "C89010", bold: true, margin: 0, align: "center",
      });

      slide.addText("10 = 6 + 4", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 3.10, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });
    }
  );

  // Slide 7: I Do #2 — A different way
  workedExSlide(pres, 2, "I Do", "Same 10, Different Way",
    [
      "Same 10 lollies. Same 2 jars.",
      "",
      "This time:",
      "Jar A: 7 lollies",
      "Jar B: 3 lollies",
      "",
      "7 and 3 make 10",
      "",
      "Same whole!",
      "Different parts!",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.5, { strip: C.SECONDARY });

      slide.addText("10 = 7 + 3", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });

      // Two jars
      drawLollyJar(slide, lg.rightX + 0.15, lg.panelTopPadded + 0.5, 1.75, 2.2, 7, { lollyColor: "E63946" });
      drawLollyJar(slide, lg.rightX + 2.10, lg.panelTopPadded + 0.5, 1.75, 2.2, 3, { lollyColor: "F4C430" });

      slide.addText("Jar A = 7", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.75, w: 1.75, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: "E63946", bold: true, margin: 0, align: "center",
      });
      slide.addText("Jar B = 3", {
        x: lg.rightX + 2.10, y: lg.panelTopPadded + 2.75, w: 1.75, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: "C89010", bold: true, margin: 0, align: "center",
      });

      slide.addText("Same whole. Different parts.", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 3.10, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, italic: true, bold: true, margin: 0, align: "center",
      });
    }
  );

  // Slide 8-9: CFU — How many in Jar B? (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["2"]);
      addStageBadge(s, 2, "CFU");
      addTitle(s, "How Many in Jar B?", { y: 0.65, fontSize: 24, color: STAGE_COLORS["2"] });

      // Left: prompt
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["2"] });
      s.addText([
        { text: "Whole = 10", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Show me on your fingers:", options: { fontSize: 16, bold: true, color: C.ALERT, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "How many are in Jar B?", options: { fontSize: 18, bold: true, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.2, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: two jars - Jar A has 8, Jar B has 2 (but B is shown empty-looking or with a ? mark)
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

      drawLollyJar(s, 5.4, CONTENT_TOP + 0.4, 1.95, 2.4, 8, { lollyColor: "E63946" });
      drawLollyJar(s, 7.55, CONTENT_TOP + 0.4, 1.95, 2.4, 2, { lollyColor: "F4C430" });

      s.addText("Jar A = 8", {
        x: 5.4, y: CONTENT_TOP + 2.85, w: 1.95, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: "E63946", bold: true, margin: 0, align: "center",
      });
      s.addText("Jar B = ?", {
        x: 7.55, y: CONTENT_TOP + 2.85, w: 1.95, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Jar B = 2!   10 = 8 + 2", {
        x: 1.0, y: 4.1, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do — find the missing part (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Jar A has 5. What is Jar B?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      // Left: prompt
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "With your partner:", options: { fontSize: 18, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Whole = 10", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "Jar A = 5", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "Jar B = ?", options: { bullet: true, fontSize: 16, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "On your whiteboard:", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "10 = 5 and ___", options: { fontSize: 18, bold: true, color: C.PRIMARY } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: two jars - Jar A has 5, Jar B empty
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      drawLollyJar(s, 5.4, CONTENT_TOP + 0.4, 1.95, 2.4, 5, { lollyColor: "E63946" });
      drawLollyJar(s, 7.55, CONTENT_TOP + 0.4, 1.95, 2.4, 0, { lollyColor: "F4C430" });

      // Big question mark in empty jar
      s.addText("?", {
        x: 7.55, y: CONTENT_TOP + 0.8, w: 1.95, h: 1.5,
        fontSize: 90, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      s.addText("Jar A = 5", {
        x: 5.4, y: CONTENT_TOP + 2.85, w: 1.95, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: "E63946", bold: true, margin: 0, align: "center",
      });
      s.addText("Jar B = ?", {
        x: 7.55, y: CONTENT_TOP + 2.85, w: 1.95, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Jar B = 5!   10 = 5 + 5", {
        x: 1.0, y: 4.1, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU Hinge — Is this right? (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Is This Right?", "Thumbs Up/Down",
      "Whole = 10\n\nJar A = 4    Jar B = 7\n\nIs that right?\n\nThumbs UP = YES    Thumbs DOWN = NO",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN!  4 + 7 = 11, not 10", {
        x: 1.0, y: 4.1, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do — Lolly Jars A3 activity
  workedExSlide(pres, 4, "You Do", "Lolly Shop Owner",
    [
      "You are the owner!",
      "",
      "First: Take 10 lollies (counters)",
      "",
      "Next: Put them in 2 jars.",
      "",
      "Then: Draw it on your Lolly Jars sheet.",
      "",
      "Find as MANY WAYS as you can!",
      "",
      "You have 15 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.ALERT });

      slide.addText("Your challenge:", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.3, h: 0.35,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });

      // Draw two jars with ?
      drawLollyJar(slide, lg.rightX + 0.15, lg.panelTopPadded + 0.55, 1.85, 2.0, 0, { lollyColor: "E63946" });
      drawLollyJar(slide, lg.rightX + 2.15, lg.panelTopPadded + 0.55, 1.85, 2.0, 0, { lollyColor: "F4C430" });

      slide.addText("?", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.95, w: 1.85, h: 1.4,
        fontSize: 70, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText("?", {
        x: lg.rightX + 2.15, y: lg.panelTopPadded + 0.95, w: 1.85, h: 1.4,
        fontSize: 70, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("10 = ? + ?", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 2.65, w: lg.rightW - 0.2, h: 0.4,
        fontSize: 18, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      slide.addText("How many ways?", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 3.05, w: lg.rightW - 0.2, h: 0.25,
        fontSize: 13, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0, align: "center",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Look at the jars: Jar A has 6, Jar B has 4. What is the whole?",
      "The whole is 10. Draw a different way to split 10 lollies.",
      "Tell your teacher: What word do we use for the TOTAL number of lollies?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "What is your favourite way to show parts and whole? Counters? Towers? Tens frames? Lolly jars?",
    [
      "I can show a number as two parts",
      "I can find different ways to make the same number",
      "I can use pictures, objects, or numbers to explain my thinking",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPW_Session4_Lolly_Jars.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Helper: draw a PDF lolly jar
  function drawPdfLollyJar(doc, x, y, w, h, preFilledCount, lollyColor) {
    const lidH = h * 0.08;
    const neckH = h * 0.10;
    const bodyY = y + lidH + neckH;
    const bodyH = h - lidH - neckH;
    // Lid
    doc.rect(x + w * 0.18, y, w * 0.64, lidH).fillAndStroke("#6FA3D2", "#4080B0");
    // Neck
    doc.rect(x + w * 0.22, y + lidH, w * 0.56, neckH).fillAndStroke("#D3ECF7", "#6FA3D2");
    // Body
    doc.roundedRect(x, bodyY, w, bodyH, 12).fillAndStroke("#D3ECF7", "#6FA3D2");
    doc.fillColor("#333333");
    // Lollies inside
    if (preFilledCount > 0) {
      const cols = 5;
      const lollyR = Math.min((w * 0.65) / (cols * 2), bodyH * 0.15);
      const lollyX0 = x + (w - (cols * 2 * lollyR + (cols - 1) * 4)) / 2;
      const lollyY0 = bodyY + bodyH * 0.20;
      for (let i = 0; i < preFilledCount; i++) {
        const r = Math.floor(i / cols);
        const c = i % cols;
        const cx = lollyX0 + c * (2 * lollyR + 4) + lollyR;
        const cy = lollyY0 + r * (2 * lollyR + 4) + lollyR;
        doc.circle(cx, cy, lollyR).lineWidth(0.8).fillAndStroke(lollyColor || "#E63946", "#333333");
      }
      doc.fillColor("#333333");
    }
  }

  // Worksheet: Lolly Jars A3-ish (still A4 portrait but with big jars)
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "You are the lolly shop owner! Split 10 lollies across 2 jars.",
      color: C.PRIMARY,
      lessonInfo: "Session 4 of 4 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Find as many different ways as you can. Colour the lollies. Write the parts.", y, { color: C.ACCENT });

    const drawJarRow = (ty, way) => {
      doc.fontSize(13).fillColor("#333333").text("Way " + way + ":  10 = _____ + _____", 60, ty);
      // Two empty jars
      drawPdfLollyJar(doc, 60, ty + 22, 130, 110, 0);
      drawPdfLollyJar(doc, 220, ty + 22, 130, 110, 0);
      doc.fontSize(11).fillColor("#333333")
        .text("Jar A: _____", 90, ty + 142)
        .text("Jar B: _____", 250, ty + 142);
      return ty + 170;
    };

    y = drawJarRow(y, 1);
    y = drawJarRow(y, 2);
    y = drawJarRow(y, 3);

    y = addTipBox(doc, "Challenge: Can you find more than 3 ways? Use the back of this sheet.", y, { color: C.SECONDARY });

    addPdfFooter(doc, "Session 4 | Lolly Jars Worksheet | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "All 11 ways to split 10 lollies across 2 jars",
      color: C.PRIMARY,
      lessonInfo: "Session 4 of 4 | Foundation Numeracy",
    });
    y = addBodyText(doc, "For a whole of 10 split into 2 jars, there are 11 ways:", y);
    y = addSectionHeading(doc, "All Ways", y, { color: C.PRIMARY });
    y = addBodyText(doc, "10 + 0    0 + 10", y);
    y = addBodyText(doc, "9 + 1      1 + 9", y);
    y = addBodyText(doc, "8 + 2      2 + 8", y);
    y = addBodyText(doc, "7 + 3      3 + 7", y);
    y = addBodyText(doc, "6 + 4      4 + 6", y);
    y = addBodyText(doc, "5 + 5", y);

    y = addTipBox(doc, "Pattern to notice: the numbers in Jar A go 10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0 while the numbers in Jar B go 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10. Each pair adds to 10.", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Assessment Guidance", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Core target: 3 or more different pairs.", y);
    y = addBodyText(doc, "Strong: 5 or more different pairs.", y);
    y = addBodyText(doc, "Mastery: 7+ different pairs or all 11 ways found.", y);

    addPdfFooter(doc, "Session 4 | Answer Key | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — whole of 5, some lollies pre-drawn
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "5 lollies in 2 jars. Some are drawn for you.",
      color: C.ACCENT,
      lessonInfo: "Session 4 of 4 | Foundation Numeracy",
    });
    y = addTipBox(doc, "The jars are drawn. Count the lollies that are there. Then draw the rest so both jars have 5 lollies altogether.", y, { color: C.ACCENT });

    const drawScaffoldRow = (ty, way, aCount, bCount) => {
      doc.fontSize(13).fillColor("#333333").text("Way " + way + ":  Whole = 5", 60, ty);
      drawPdfLollyJar(doc, 60, ty + 22, 130, 110, aCount);
      drawPdfLollyJar(doc, 220, ty + 22, 130, 110, bCount);
      doc.fontSize(11).fillColor("#333333")
        .text("Jar A: " + aCount + " lollies", 85, ty + 142)
        .text("Jar B: " + bCount + " lollies", 245, ty + 142);
      doc.fontSize(11).fillColor("#333333")
        .text("5 = " + aCount + " + " + bCount, 170, ty + 160);
      return ty + 180;
    };

    y = drawScaffoldRow(y, 1, 4, 1);
    y = drawScaffoldRow(y, 2, 3, 2);
    y = drawScaffoldRow(y, 3, 2, 3);

    addPdfFooter(doc, "Session 4 | Enabling Scaffold | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — 3 jars for 10
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Challenge: split 10 lollies across THREE jars",
      color: C.PRIMARY,
      lessonInfo: "Session 4 of 4 | Foundation Numeracy",
    });
    y = addBodyText(doc, "Big challenge! You have 10 lollies and 3 jars. How many different ways can you fill them?", y);

    const drawTripleJarRow = (ty, way) => {
      doc.fontSize(13).fillColor("#333333").text("Way " + way + ":  10 = _____ + _____ + _____", 60, ty);
      drawPdfLollyJar(doc, 60, ty + 20, 110, 95, 0);
      drawPdfLollyJar(doc, 195, ty + 20, 110, 95, 0);
      drawPdfLollyJar(doc, 330, ty + 20, 110, 95, 0);
      doc.fontSize(10).fillColor("#333333")
        .text("Jar A:", 90, ty + 125)
        .text("Jar B:", 225, ty + 125)
        .text("Jar C:", 360, ty + 125);
      return ty + 150;
    };

    y = drawTripleJarRow(y, 1);
    y = drawTripleJarRow(y, 2);
    y = drawTripleJarRow(y, 3);

    y = addTipBox(doc, "Super challenge: there are more ways than you think. Try: 5+3+2, 4+4+2, 6+2+2... how many can you find? Use the back of this sheet to record.", y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 4 | Extension | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 4 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
