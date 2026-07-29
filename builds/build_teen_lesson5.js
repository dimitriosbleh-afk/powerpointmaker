"use strict";

// Teen Numbers Unit — Session 5: Representing Teen Numbers
// Term 2 Week 1, Foundation Numeracy, Variant 0
// DR: Data Collection — I can sort and classify objects into categories
// Fluency: Counting to 20 from various starting points
// VC2MFN01 — name, represent and order numbers including zero to at least 20

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  addTwoColumnOrganiser,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", 0);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 5;
const UNIT_TITLE = "Teen Numbers";
const FOOTER = "Teen Numbers | Session 5 of 5 | Foundation Numeracy";
const OUT_DIR = "output/Teen_Session5_Representing_Teen_Numbers";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Multiple representations practice for teen numbers.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference for representations.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Guided representation matching with visual support.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Create a teen numbers poster with all representations.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Ten frame helper ────────────────────────────────────────────────────────

function drawTenFrame(slide, x, y, w, h, filled, opts) {
  const o = opts || {};
  const cols = 5;
  const rows = 2;
  const cellW = w / cols;
  const cellH = h / rows;

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const idx = r * cols + c;
      const cx = x + c * cellW;
      const cy = y + r * cellH;
      slide.addShape("rect", {
        x: cx, y: cy, w: cellW, h: cellH,
        fill: { color: idx < filled ? (o.fillColor || C.PRIMARY) : (o.emptyColor || C.WHITE) },
        line: { color: o.borderColor || C.PRIMARY, width: 1.5 },
      });
      if (idx < filled) {
        const dotSize = Math.min(cellW, cellH) * 0.55;
        slide.addShape("roundRect", {
          x: cx + (cellW - dotSize) / 2, y: cy + (cellH - dotSize) / 2,
          w: dotSize, h: dotSize, rectRadius: dotSize / 2,
          fill: { color: o.dotColor || C.ACCENT },
        });
      }
    }
  }
}

function drawTenFrameWithExtras(slide, x, y, w, h, total, opts) {
  const o = opts || {};
  drawTenFrame(slide, x, y, w, h, Math.min(total, 10), o);
  const extras = total - 10;
  if (extras > 0) {
    const cellW = w / 5;
    const cellH = h / 2;
    const dotSize = Math.min(cellW, cellH) * 0.55;
    const extraY = y + h + 0.06;
    for (let i = 0; i < extras; i++) {
      const cx = x + i * cellW;
      slide.addShape("rect", {
        x: cx, y: extraY, w: cellW, h: cellH,
        fill: { color: C.WHITE },
        line: { color: C.SECONDARY, width: 1.5 },
      });
      slide.addShape("roundRect", {
        x: cx + (cellW - dotSize) / 2, y: extraY + (cellH - dotSize) / 2,
        w: dotSize, h: dotSize, rectRadius: dotSize / 2,
        fill: { color: o.extraDotColor || C.ALERT },
      });
    }
  }
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- This is our last lesson on teen numbers this week
- Today we are going to show teen numbers in DIFFERENT ways
- You already know how to say them, write them, and use ten frames
- Now we will put it all together

DO:
- Display title slide
- Have whiteboards, counters, and ten frames ready

TEACHER NOTES:
Session 5 of 5. This consolidation lesson pulls together naming, writing, ten frame representation, and number words. Multiple representations deepen understanding -- students who can show 14 as a numeral, a word, a ten frame, and as objects have flexible number sense.

WATCH FOR:
- Students who seem confident -- they have consolidated well across the week
- Students who are still shaky on one mode (e.g., can say but not write) -- note for future planning

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Let's warm up with some sorting
- Look at these objects. Sort them into two groups.
- Group 1: things that are red. Group 2: things that are NOT red.
- Then try the second sort: things with wheels, things without wheels.

DO:
- Display the sorting challenge (pictures of objects)
- Students call out which group each object belongs in
- Use whiteboards for the second sort

TEACHER NOTES:
Daily Review practises data collection (sorting and classifying). This is a Term 1 concept being retrieved. Sorting requires the same categorical thinking used when classifying number representations.

WATCH FOR:
- Students who sort quickly and confidently -- classification skills are strong
- Students who put objects in the wrong group -- they may need clearer criteria

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check: the apple is red, the car is red, the leaf is NOT red, the ball is red, the book is NOT red
- Wheels: car and bike have wheels. Apple, book, ball do not.
- Sorting means putting things into groups that belong together

DO:
- Click to reveal sorted groups
- Students check their answers
- Briefly connect: "Today we will sort different WAYS to show numbers"

TEACHER NOTES:
The bridge to today's lesson: sorting objects into groups is the same skill as matching different representations of the same number.

WATCH FOR:
- Students who see the connection to numbers -- strong transfer thinking

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Today we are going to count from different starting points.
- Do not always start at 1! Sometimes we start in the middle.
- Ready? Count from 5 to 15 with me. 5, 6, 7, 8...
- Now count from 11 to 20. 11, 12, 13...
- Now count from 8 to 18. 8, 9, 10, 11...

DO:
- Lead choral counts from different starting points
- 3 rounds: 5-15, 11-20, 8-18
- Keep it brisk
- Students count chorally, getting louder on teen numbers

TEACHER NOTES:
Counting from various starting points builds flexible number sense. Students who can only count from 1 have a rigid chain -- starting from 5 or 11 requires genuine understanding of the sequence.

WATCH FOR:
- Students who hesitate when starting from 8 or 11 -- they rely on counting from 1
- Students who jump in confidently from any starting point -- flexible counting is developing

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to show teen numbers in different ways"
- Our success criteria: read each one.
- Ask: What ways do we already know to show a teen number? [Say it, write it, ten frame]

DO:
- Choral read LI and SC
- Quick brainstorm of representation types already learned

TEACHER NOTES:
This lesson consolidates all representation modes from the week. SC1 is the floor (show with objects/drawings). SC2 is the core (ten frame). SC3 extends to matching across all modes.

WATCH FOR:
- Students who name multiple representations -- they have been tracking the week's learning
- Students who only remember one way -- the lesson will expand their repertoire

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me show the number 14 in FOUR different ways
- Way 1: I say it. Fourteen.
- Way 2: I write it. 1-4. Fourteen.
- Way 3: I draw it on a ten frame. Full ten plus 4 extras.
- Way 4: I use objects. 14 counters -- 10 in a line, 4 more.
- All four ways show the SAME number: 14.

DO:
- Point to each representation on screen as you model it
- Emphasise "same number, different ways"
- Have physical counters to demonstrate Way 4

TEACHER NOTES:
Multiple representations build flexible number sense. The key insight for students is that 14, fourteen, a ten frame with 4 extras, and 14 counters are all the same thing shown differently. This is an early equivalence concept.

WATCH FOR:
- Students who nod along -- they see the connection between representations
- Students who look surprised that they are all the same -- the connection is new and important

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Let me show you another number: 17
- I say it: seventeen
- I write it: 1-7
- Ten frame: full ten plus 7 extras
- Now let me show you 20 in four ways
- I say it: twenty. I write it: 2-0
- Ten frame: two full ten frames
- Objects: 20 counters in two groups of 10

DO:
- Model each representation for 17 and 20
- For 20, emphasise the difference (2 tens, not 1 ten plus extras)

TEACHER NOTES:
Modelling two more numbers embeds the pattern. Including 20 reminds students of the special case where both tens are full.

WATCH FOR:
- Students who can predict the ten frame before you draw it -- they are transferring the structure
- Students who remember that 20 is two tens -- Session 2 learning has stuck

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. I will show you a teen number in ONE way. You show me another way.
- I show you a ten frame with a full ten and 6 extras. What number is this?
- Write it on your whiteboard!

DO:
- Display ten frame showing 16
- Students write 16 on whiteboards
- Follow up: "Now say it to your partner" (oral representation)

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "I am showing you a ten frame. Full ten, 6 extras. Write the NUMBER on your whiteboard. Show me!"
- Scan for: 16. Students who write 16 are converting between ten frame and numeral representations.
PROCEED: If 80%+ write 16, move to We Do.
PIVOT: Most likely misconception -- students write 6 (reading only the extras, not including the ten). Reteach: "The full ten frame means 10. Plus 6 extras. 10 and 6 is 16. Always include the 10!"

TEACHER NOTES:
This CFU checks cross-representation transfer -- from visual (ten frame) to symbolic (numeral). This is the core skill for the lesson.

WATCH FOR:
- Students who convert instantly -- representation fluency is strong
- Readiness signal: 16 written quickly and correctly

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Partner time. I will show you a number. You need to show it THREE ways.
- Way 1: Write the numeral
- Way 2: Draw a ten frame
- Way 3: Say it to your partner
- First number: 13!

DO:
- Display the number 13 (written large)
- Partners work together: write numeral, draw ten frame, say it
- Then repeat for 18 and 15
- Circulate and check ten frame drawings especially

TEACHER NOTES:
We Do asks students to produce three representations from a given number. This is the reverse of I Do (where the teacher produced them). Different numbers from I Do ensures genuine practice.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 5 Enabling Scaffold with guided matching. Students draw a line from a numeral to its matching ten frame picture. Pre-drawn visuals reduce production demands.
- Extra Notes: Distribute the Session 5 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Use the Session 5 Extension to create a teen numbers poster page showing each number 11-20 in all four representations (numeral, word, ten frame, tally marks).
- Extra Notes: Self-contained creative task.

WATCH FOR:
- Students who draw accurate ten frames quickly -- visual representation is secure
- Students who struggle with the ten frame drawing but write the numeral correctly -- drawing needs more practice

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. 13: numeral is 1-3. Say it: thirteen. Ten frame: full ten plus 3 extras.
- 18: numeral is 1-8. Say it: eighteen. Ten frame: full ten plus 8 extras.
- 15: numeral is 1-5. Say it: fifteen. Ten frame: full ten plus 5 extras.

DO:
- Reveal representations for each number
- Students compare their ten frame drawings
- Celebrate accurate representations

TEACHER NOTES:
Confirming all three representations side by side reinforces the equivalence -- same number, different forms.

WATCH FOR:
- Students who self-correct their ten frames -- they are learning from the model
- Students whose three representations all match -- flexible number sense is developing

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. I am going to show you two things. Do they show the SAME number?
- A ten frame with full ten and 5 extras... and the numeral 15. Same number?
- Thumbs up for YES, thumbs down for NO.

DO:
- Students show thumbs (should be up -- both show 15)
- Then: "A ten frame with full ten and 3 extras... and the numeral 14. Same number?"
- Thumbs down (ten frame shows 13, not 14)

CFU CHECKPOINT:
Technique: Thumbs Up/Down
Script:
- Say: "Thumbs up if they match, thumbs down if they do not. Ten frame: full ten plus 5 extras. Numeral: 15. Same number? Show me."
- Scan for: thumbs up. Then: "Ten frame: full ten plus 3 extras. Numeral: 14. Same number? Show me."
- Scan for: thumbs down (ten frame = 13, not 14).
PROCEED: If 80%+ get both correct, move to You Do.
PIVOT: Most likely misconception -- students show thumbs up for both because they assume the teacher always shows matching pairs. Reteach: "Count the extras carefully. 1, 2, 3. That is 3 extras. 10 and 3 is 13. But the numeral says 14. 13 and 14 are NOT the same number."

TEACHER NOTES:
The mismatch example is critical -- it confirms students are actually checking, not just agreeing.

WATCH FOR:
- Students who catch the mismatch instantly -- strong cross-representation checking
- Readiness signal: confident thumbs down with explanation that "ten frame shows 13, not 14"

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice. Your worksheet has three sections.
- First: Match the numeral to the ten frame. Next: Draw ten frames for given numbers. Then: Write the numeral and draw the ten frame for each number word.
- You have 7 minutes.

DO:
- Distribute Session 5 Worksheet
- Students work independently
- Circulate: check ten frame drawings and numeral-to-representation matching
- Distribute scaffolds/extensions as needed

TEACHER NOTES:
You Do requires all three representation conversions: numeral to ten frame, ten frame to numeral, and word to both. Different content from We Do ensures independent application.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete the enabling scaffold with pre-drawn ten frames to match. Students draw lines from numerals to the correct ten frame picture. No drawing required.
- Extra Notes: Distribute the Session 5 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Create a teen numbers poster using the Session 5 Extension. Show each number 11-20 with all four representations (numeral, word, ten frame, tally marks).
- Extra Notes: Self-contained creative investigation.

WATCH FOR:
- Students who complete all three sections accurately -- representation fluency is strong
- Students who match correctly but struggle to draw ten frames -- production is harder than recognition
- Readiness signal: completing 8+ items correctly in 5 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Final exit ticket of the week! Three questions.
- Work on your own. 2 minutes.

DO:
- Display exit ticket
- Students respond on whiteboards
- Sort for end-of-unit assessment data

TEACHER NOTES:
Exit ticket assesses all three SC across the week. Q1 checks representation (drawing). Q2 checks cross-representation matching. Q3 checks understanding of structure. This data informs whether the unit needs revision in future weeks.

WATCH FOR:
- Students who get all three -- teen numbers unit is secure
- Students who struggle with any question -- note the specific gap for future planning

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- What a week! Let's check our success criteria one last time.
- SC1: I can show a teen number using objects or drawings. Thumbs?
- SC2: I can use a ten frame to represent teen numbers 11-20. Thumbs?
- SC3: I can match a teen number to its ten frame, numeral, and word. Thumbs?
- Turn and talk: How many ways can you show the number 16? Tell your partner all the ways!

DO:
- Display SC
- Thumbs check
- Turn and Talk
- Cold call 2-3 students -- celebrate multiple representations

TEACHER NOTES:
The final closing celebrates a full week of learning. The Turn and Talk asks students to demonstrate the core skill: showing a number in multiple ways. Students who can name 3+ representations (numeral, word, ten frame, objects) have flexible teen number knowledge.

WATCH FOR:
- Students who name 3-4 ways -- strong consolidation
- Students who name 1-2 ways -- they have a foundation to build on next week

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Here are this week's resources for the final session

DO:
- Point out resources

TEACHER NOTES:
Resource slide for Session 5.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Session 5: Representing Teen Numbers",
    "Foundation Numeracy | Session 5 of 5 | Term 2 Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // DR (withReveal) — Data: sorting objects
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Sort It Out!", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText([
        { text: "Sort these into two groups:", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "apple, car, leaf, ball, book", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Sort 1: Red / Not Red", options: { fontSize: 15, bold: true, color: C.ALERT, breakLine: true } },
        { text: "Which ones are red?", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Sort 2: Has Wheels / No Wheels", options: { fontSize: 15, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Which ones have wheels?", options: { fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Write R for red, N for not red.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Then write W for wheels,", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "NW for no wheels.", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Red: apple, car, ball   |   Wheels: car   |   Not red: leaf, book", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency — Counting from various starting points
  contentSlide(pres, "Fluency", STAGE_COLORS["1"], "Count From Here!",
    [
      "Count together from each starting point:",
      "",
      "Round 1: Start at 5, count to 15",
      "Round 2: Start at 11, count to 20",
      "Round 3: Start at 8, count to 18",
      "",
      "Get LOUDER on the teen numbers!",
    ],
    NOTES_FLUENCY, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.ACCENT });
      slide.addText("Start Anywhere!", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });

      // Three starting point badges
      const starts = [
        { from: "5", to: "15", color: C.PRIMARY },
        { from: "11", to: "20", color: C.SECONDARY },
        { from: "8", to: "18", color: C.ALERT },
      ];
      starts.forEach((st, i) => {
        const sy = lg.panelTopPadded + 0.5 + i * 0.6;
        addTextOnShape(slide, st.from + " -> " + st.to, {
          x: lg.rightX + 0.4, y: sy, w: lg.rightW - 0.8, h: 0.45, rectRadius: 0.08,
          fill: { color: st.color },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // LI/SC
  liSlide(pres,
    ["We are learning to show teen numbers in different ways"],
    [
      "I can show a teen number using objects or drawings",
      "I can use a ten frame to represent teen numbers 11 to 20",
      "I can match a teen number to its ten frame, numeral, and word",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do — Four representations of 14
  workedExSlide(pres, 2, "I Do", "Four Ways to Show 14",
    [
      "Same number, different ways!",
      "",
      "Way 1: SAY it -- fourteen",
      "Way 2: WRITE it -- 14",
      "Way 3: TEN FRAME -- full ten + 4",
      "Way 4: OBJECTS -- 14 counters",
      "",
      "All four ways show 14!",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.PRIMARY });
      slide.addText("14", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.45,
        fontSize: 36, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Ten frame for 14
      drawTenFrameWithExtras(slide, lg.rightX + 0.5, lg.panelTopPadded + 0.6, 1.3, 0.4, 14);

      // Word
      addTextOnShape(slide, "fourteen", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.5, w: lg.rightW - 0.6, h: 0.35, rectRadius: 0.06,
        fill: { color: C.SECONDARY },
      }, { fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true });

      // Objects visual (dot grid)
      slide.addText("Objects: 14 counters", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.0, w: lg.rightW - 0.3, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, align: "center",
      });
      for (let i = 0; i < 14; i++) {
        const col = i % 7;
        const row = Math.floor(i / 7);
        slide.addShape("roundRect", {
          x: lg.rightX + 0.4 + col * 0.45,
          y: lg.panelTopPadded + 2.3 + row * 0.4,
          w: 0.3, h: 0.3, rectRadius: 0.15,
          fill: { color: i < 10 ? C.PRIMARY : C.ALERT },
        });
      }
    }
  );

  // I Do — 17 and 20
  workedExSlide(pres, 2, "I Do", "More Examples: 17 and 20",
    [
      "17: say seventeen, write 17",
      "   Ten frame: full ten + 7 extras",
      "",
      "20: say twenty, write 20",
      "   20 is special! Two full tens!",
      "   No extras left over",
      "",
      "Every teen number can be shown",
      "in all four ways.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.SECONDARY });
      // 17 section
      slide.addText("17 = seventeen", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.5, lg.panelTopPadded + 0.42, 1.2, 0.36, 17);

      // 20 section
      slide.addText("20 = twenty (two tens!)", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.35, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });
      drawTenFrame(slide, lg.rightX + 0.3, lg.panelTopPadded + 1.7, 1.2, 0.36, 10, { dotColor: C.ACCENT });
      drawTenFrame(slide, lg.rightX + 1.7, lg.panelTopPadded + 1.7, 1.2, 0.36, 10, { dotColor: C.ALERT });

      addTextOnShape(slide, "10 + 10 = 20", {
        x: lg.rightX + 0.5, y: lg.panelTopPadded + 2.25, w: 2.2, h: 0.3, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // CFU (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Ten Frame to Numeral", "Show Me Boards",
      "I show you a ten frame:\nFull ten + 6 extras.\n\nWrite the NUMBER on your whiteboard!\n\nThen SAY it to your partner.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "16!  Sixteen.  (10 + 6 = 16)", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Show It Three Ways!", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "With your partner:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Show each number THREE ways:", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "1. Write the numeral", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "2. Draw a ten frame", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "3. Say it to your partner", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Number A: 13", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true, paraSpaceAfter: 8 } },
        { text: "Number B: 18", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true, paraSpaceAfter: 8 } },
        { text: "Number C: 15", options: { fontSize: 18, bold: true, color: C.PRIMARY } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 2.0, { strip: C.ACCENT });
      s.addText([
        { text: "Three Ways:", options: { fontSize: 15, bold: true, color: C.ACCENT, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
      // Three way badges
      const ways = [
        { text: "WRITE the numeral", color: C.PRIMARY },
        { text: "DRAW a ten frame", color: C.SECONDARY },
        { text: "SAY it aloud", color: C.ALERT },
      ];
      ways.forEach((w, i) => {
        addTextOnShape(s, w.text, {
          x: 5.5, y: CONTENT_TOP + 0.55 + i * 0.45, w: 3.6, h: 0.38, rectRadius: 0.06,
          fill: { color: w.color },
        }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A) 13 = thirteen, ten+3   B) 18 = eighteen, ten+8   C) 15 = fifteen, ten+5", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Do They Match?", "Thumbs Up/Down",
      "Ten frame: full ten + 5 extras AND numeral 15.\nSame number? Thumbs up or down!\n\nThen: ten frame: full ten + 3 extras AND numeral 14.\nSame number?",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "First: YES (both = 15)    Second: NO (ten frame = 13, not 14!)", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Show Teen Numbers Your Way",
    [
      "First: Match the numeral to its ten frame.",
      "Next: Draw ten frames for the given numbers.",
      "Then: Write the numeral AND draw the ten frame",
      "for each number word.",
      "",
      "Use all three sections on your worksheet.",
      "",
      "You have 7 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.ALERT });
      slide.addText("This Week You Learned:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Say teen numbers", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Write teen numbers (1 first!)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Order teen numbers", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Show teen numbers with ten frames", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 1.1,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Exit Ticket
  exitTicketSlide(pres,
    [
      "Draw a ten frame to show the number 16.",
      "A ten frame shows a full ten and 2 extras. What numeral is this?",
      "Why do teen numbers always start with 1? (Hint: the ten!)",
    ],
    NOTES_EXIT, FOOTER);

  // Closing
  closingSlide(pres,
    "How many ways can you show the number 16? Tell your partner ALL the ways!",
    [
      "I can show a teen number using objects or drawings",
      "I can use a ten frame to represent teen numbers 11 to 20",
      "I can match a teen number to its ten frame, numeral, and word",
    ],
    NOTES_CLOSING);

  // Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Teen_Session5_Representing_Teen_Numbers.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Representing Teen Numbers", color: C.PRIMARY, lessonInfo: "Session 5 of 5 | Foundation Numeracy" });
    y = addTipBox(doc, "Show teen numbers in different ways: numeral, ten frame, and number word!", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section A: Match the numeral to its description", y, { color: C.PRIMARY });
    y = addProblem(doc, 1, "14  ->  ten and ___ extras", y, { color: C.PRIMARY });
    y = addProblem(doc, 2, "17  ->  ten and ___ extras", y, { color: C.PRIMARY });
    y = addProblem(doc, 3, "11  ->  ten and ___ extra", y, { color: C.PRIMARY });
    y = addProblem(doc, 4, "20  ->  ___ full tens", y, { color: C.PRIMARY });
    y = addProblem(doc, 5, "19  ->  ten and ___ extras", y, { color: C.PRIMARY });

    y = addSectionHeading(doc, "Section B: Draw a ten frame for each number", y, { color: C.PRIMARY });
    y = addProblem(doc, 6, "Draw a ten frame showing 15 (shade the ten, draw extras)", y, { color: C.PRIMARY, writeLines: [{ label: "" }] });
    y = addProblem(doc, 7, "Draw a ten frame showing 18", y, { color: C.PRIMARY, writeLines: [{ label: "" }] });
    y = addProblem(doc, 8, "Draw TWO ten frames showing 20", y, { color: C.PRIMARY, writeLines: [{ label: "" }] });

    y = addSectionHeading(doc, "Section C: Word to numeral and ten frame", y, { color: C.PRIMARY });
    y = addProblem(doc, 9, "thirteen: write the numeral ___ and draw the ten frame", y, { color: C.PRIMARY, writeLines: [{ label: "" }] });
    y = addProblem(doc, 10, "sixteen: write the numeral ___ and draw the ten frame", y, { color: C.PRIMARY, writeLines: [{ label: "" }] });

    addPdfFooter(doc, "Session 5 | Teen Numbers | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Representing Teen Numbers", color: C.PRIMARY, lessonInfo: "Session 5 of 5 | Foundation Numeracy" });
    y = addSectionHeading(doc, "Section A", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1. 14 = ten and 4 extras   2. 17 = ten and 7 extras   3. 11 = ten and 1 extra", y);
    y = addBodyText(doc, "4. 20 = 2 full tens   5. 19 = ten and 9 extras", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.PRIMARY });
    y = addBodyText(doc, "6. 15: full ten frame + 5 extras shaded", y);
    y = addBodyText(doc, "7. 18: full ten frame + 8 extras shaded", y);
    y = addBodyText(doc, "8. 20: two full ten frames (all cells shaded)", y);
    y = addSectionHeading(doc, "Section C", y, { color: C.PRIMARY });
    y = addBodyText(doc, "9. thirteen = 13, ten frame: full ten + 3 extras", y);
    y = addBodyText(doc, "10. sixteen = 16, ten frame: full ten + 6 extras", y);
    addPdfFooter(doc, "Session 5 | Answer Key | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Match the Representations", color: C.ACCENT, lessonInfo: "Session 5 of 5 | Foundation Numeracy" });
    y = addTipBox(doc, "Draw a line from each number to the correct description. No drawing needed!", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Match each number to its ten frame description", y, { color: C.PRIMARY });
    y = addProblem(doc, 1, "12  ->  ten and ?  (write: 1, 2, 3, 4, or 5 extras)", y, { color: C.PRIMARY });
    y = addProblem(doc, 2, "14  ->  ten and ?  extras", y, { color: C.PRIMARY });
    y = addProblem(doc, 3, "11  ->  ten and ?  extra", y, { color: C.PRIMARY });
    y = addProblem(doc, 4, "15  ->  ten and ?  extras", y, { color: C.PRIMARY });
    y = addProblem(doc, 5, "13  ->  ten and ?  extras", y, { color: C.PRIMARY });

    addPdfFooter(doc, "Session 5 | Enabling Scaffold | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF: " + ENABLING_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "Teen Numbers Poster Page", color: C.PRIMARY, lessonInfo: "Session 5 of 5 | Foundation Numeracy" });
    y = addSectionHeading(doc, "Create Your Teen Numbers Poster!", y, { color: C.PRIMARY });
    y = addBodyText(doc, "For each number below, fill in ALL four ways to show it:", y);
    y = addBodyText(doc, "1. Write the NUMERAL    2. Write the WORD    3. Draw a TEN FRAME    4. Draw TALLY MARKS", y);
    y += 5;

    const posterNums = [11, 13, 16, 19, 20];
    posterNums.forEach((n, i) => {
      y = addSectionHeading(doc, "Number: " + n, y, { color: C.PRIMARY, fontSize: 12 });
      y = addWriteLine(doc, "Numeral:", y);
      y = addWriteLine(doc, "Word:", y);
      y = addWriteLine(doc, "Ten frame (draw):", y);
      y = addWriteLine(doc, "Tally marks (draw):", y);
      y += 8;
      if (y > 720) { doc.addPage(); y = 50; }
    });

    y = addTipBox(doc, "Tally marks: draw a line for each one. Cross every group of 5. For 13: IIII IIII III (that is 5 + 5 + 3 = 13)", y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 5 | Extension | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 5 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
