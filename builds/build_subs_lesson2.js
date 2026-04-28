"use strict";

// Subtraction Strategies (Year 2 Numeracy) — Lesson 2: Think Addition (Part-Part-Whole)
// VC2M2N04 / VC2M2A02 — use a known addition fact to find a subtraction fact
//   (e.g. 8 + 7 = 15  =>  15 - 7 = 8 and 15 - 8 = 7).
// Daily Review: 2D Shapes — sort shapes by sides.
// Fluency: Skip counting in 5s and 10s.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Same theme as Lesson 1 (variant 0) — unit cohesion.
const T = createTheme("numeracy", "grade2", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 5;
const UNIT_TITLE = "Subtraction Strategies";
const FOOTER = `Subtraction Strategies | Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`;
const OUT_DIR = "output/Subs_Lesson2_Think_Addition";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 2 Fact Family Practice",
  "Use a known addition to find two subtraction facts.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 2 Answer Key",
  "Worked answers for the fact family practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a Part-Part-Whole bar model with the whole on top and two parts under.
// `whole`, `partA`, `partB` are numbers; pass null to leave the cell blank.
function drawPpwBar(slide, x, y, w, h, whole, partA, partB, opts) {
  const o = opts || {};
  const wholeH = h * 0.40;
  const partH  = h * 0.55;
  const partGap = 0.10;
  const partW  = (w - partGap) / 2;
  const wholeColor = o.wholeColor || C.PRIMARY;
  const partColor  = o.partColor  || C.SECONDARY;

  // Whole bar
  slide.addShape("roundRect", {
    x, y, w, h: wholeH, rectRadius: 0.10,
    fill: { color: wholeColor },
  });
  slide.addText(whole == null ? "Whole" : "Whole = " + whole, {
    x, y, w, h: wholeH,
    fontSize: whole == null ? 22 : 30, fontFace: FONT_H,
    color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Part A
  const partsY = y + wholeH + 0.05;
  slide.addShape("roundRect", {
    x, y: partsY, w: partW, h: partH, rectRadius: 0.10,
    fill: { color: partColor },
  });
  slide.addText(partA == null ? "?" : String(partA), {
    x, y: partsY, w: partW, h: partH,
    fontSize: partA == null ? 44 : 36, fontFace: FONT_H,
    color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  // Part B
  slide.addShape("roundRect", {
    x: x + partW + partGap, y: partsY, w: partW, h: partH, rectRadius: 0.10,
    fill: { color: partColor },
  });
  slide.addText(partB == null ? "?" : String(partB), {
    x: x + partW + partGap, y: partsY, w: partW, h: partH,
    fontSize: partB == null ? 44 : 36, fontFace: FONT_H,
    color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// Draw a single shape outline at given coords.
function drawShape(slide, kind, x, y, size, color) {
  if (kind === "circle") {
    slide.addShape("roundRect", {
      x, y, w: size, h: size, rectRadius: size / 2,
      fill: { color }, line: { color: C.CHARCOAL, width: 1.2 },
    });
  } else if (kind === "triangle") {
    slide.addShape("triangle", {
      x, y, w: size, h: size,
      fill: { color }, line: { color: C.CHARCOAL, width: 1.2 },
    });
  } else if (kind === "pentagon") {
    slide.addShape("pentagon", {
      x, y, w: size, h: size,
      fill: { color }, line: { color: C.CHARCOAL, width: 1.2 },
    });
  } else if (kind === "hexagon") {
    slide.addShape("hexagon", {
      x, y, w: size, h: size,
      fill: { color }, line: { color: C.CHARCOAL, width: 1.2 },
    });
  } else {
    slide.addShape("rect", {
      x, y, w: size, h: size,
      fill: { color }, line: { color: C.CHARCOAL, width: 1.2 },
    });
  }
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back team. Yesterday we counted back to subtract.
- Today we use a clever shortcut: think addition.
- If you know an addition fact, you already know two subtraction facts.

DO:
- Have whiteboards, markers and counters ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 2 of 5. Today's strategy connects to Lesson 1: students still find a missing part, but they use a known addition fact to do it. The Part-Part-Whole model anchors this idea.

WATCH FOR:
- Students who shrug at the word strategy - reframe as "a clever shortcut".

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Look at the shapes on the slide.
- Whisper to your partner: which one has 5 sides?

DO:
- Display the four shapes (square, triangle, pentagon, hexagon).
- Allow 30 seconds for partner whisper.
- Walk and listen.

TEACHER NOTES:
Daily Review continues 2D shapes from earlier in the term. Today's focus is naming sides for less familiar shapes (pentagon, hexagon).

WATCH FOR:
- Students who guess hexagon - count sides aloud together.
- Students who confuse pentagon with triangle - point at corners to slow them down.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- The pentagon has 5 sides.
- The hexagon has 6 sides.
- The square has 4. The triangle has 3.

DO:
- Click to reveal the answers.
- Tick on whiteboards if students wrote pentagon.

TEACHER NOTES:
Tick and fix. Note any students who muddled pentagon and hexagon for a small group.

WATCH FOR:
- Students who self-correct on the reveal - secure.
- Students who keep the wrong shape - small group focus.

[Stage 1: Daily Review Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. We are skip counting in 5s.
- Stand up. 5, 10, 15, 20, 25, 30...
- Now in 10s. 10, 20, 30, 40, 50...

DO:
- Lead choral skip count by 5s to 50, then by 10s to 100.
- Repeat the 5s count once more, brisk.
- Sit students back down.

TEACHER NOTES:
Skip counting builds the part-whole flexibility students will use today. This is automaticity, not new teaching.

WATCH FOR:
- Students who lose track at 25 - the typical stumble. Pause and restart.
- Students who skip count smoothly - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to use a known addition fact to find a subtraction fact.
- Read the success criteria together.

DO:
- Choral read the LI.
- Choral read each I can statement.
- Hold up two fingers and three fingers and say "two and three is five, so five take two is three".

TEACHER NOTES:
The first I can is achievable for everyone. The second is the core target. The third stretches students to explain the connection in their own words.

WATCH FOR:
- Students who copy the language back - pattern is forming.
- Students who look unsure - the I Do uses counters to make it concrete.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. I know that 8 + 7 is 15.
- I have 15 counters. The whole is 15.
- One part is 8. The other part is 7.
- So 15 take away 7 must be 8.
- And 15 take away 8 must be 7.

DO:
- Use a Part-Part-Whole mat or board diagram.
- Place 15 counters in two parts: 8 and 7.
- Touch each part as you say the language pattern.
- Repeat the matched fact family at least twice.

TEACHER NOTES:
First model. The whole is unchanged; only the missing part changes. Say the language pattern at least three times. Use the words whole, part, missing part.

MISCONCEPTIONS:
- Misconception: Students think they need to count back from the whole.
  Why: They learned counting back yesterday and apply it everywhere.
  Impact: They miss the faster strategy.
  Quick correction: "We can count back. Today we use addition we already know to skip the counting."

WATCH FOR:
- Students who say "I knew it was 8" - they are using the addition fact.
- Students who count counters one by one - reteach with smaller numbers (4 + 3 = 7).

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Use the addition fact 6 + 4 = 10.
- What is 10 take away 4?
- Show me on your fingers.

DO:
- Display the bar model: whole 10, part 4, part ?
- Allow 10 seconds.
- Say: "Show me... ready... show."

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "10 take away 4? Show me on your fingers."
- Scan for: 6 fingers held up.
PROCEED: If 80% show 6, click to reveal and move to We Do.
PIVOT: Most likely misconception - students count back 4 from 10 and lose track.
- Reteach: "Look at the addition fact. 6 + 4 is 10. The missing part is 6."
- Re-check: "Show me 10 take away 4 again."

TEACHER NOTES:
Fingers are universal. Quick, every-student response.

WATCH FOR:
- Students who hold up 6 quickly - using addition.
- Students who count down on fingers - they are using yesterday's strategy; redirect.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. The fact is 5 + 4 = 9.
- What is 9 take away 5?
- Whisper to your partner. Then write your answer on your whiteboard.

DO:
- Display the bar model: whole 9, part 5, part ?
- Allow 30 seconds for partner talk and whiteboard write.
- Walk and scan.

TEACHER NOTES:
We Do uses different numbers from I Do but the same strategy. Listen for "5 plus 4 is 9, so 9 take 5 is 4".

WATCH FOR:
- Pairs who connect to addition - secure.
- Pairs who count back - prompt them: "Look at the fact. What is the missing part?"

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together.
- 5 plus 4 is 9. The whole is 9. One part is 5. The other part is 4.
- So 9 take away 5 is 4.

DO:
- Click to reveal.
- Repeat the language pattern with the class.

TEACHER NOTES:
Highlight that the whole stayed at 9. Only the missing part changed.

WATCH FOR:
- Students who say "I knew because of the addition" - secure.
- Students who keep miscounting - small group support tomorrow.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2_Q = `SAY:
- Hinge question. I say:
- I know 7 + 6 = 13. So 13 take away 6 must be 7.
- Is that right? Thumbs up for yes. Thumbs down for no.

DO:
- Display the claim.
- Wait 5 seconds.
- Scan thumbs.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "13 take away 6 is 7. Thumbs up if yes. Thumbs down if no."
- Scan for: thumbs UP. The reasoning is correct.
PROCEED: If 80% show thumbs up, click to reveal and confirm.
PIVOT: Most likely misconception - students think the parts cannot be swapped.
- Reteach: "The addition fact gives us both parts. Either part can be the missing part."
- Re-check: "What is 13 take away 7?"

TEACHER NOTES:
This hinge probes whether students see that ONE addition fact gives TWO subtraction facts.

WATCH FOR:
- Confident thumbs up - students see the connection.
- Slow thumbs - they are unsure of the part swap.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own. The fact is 9 + 6 = 15.
- Find 15 take away 6.
- Then find 15 take away 9.
- Whisper your answers to your partner.

DO:
- Hand out the fact family practice sheet (1 page) or use whiteboards.
- Circulate. Listen for "the addition gave me the answer".
- Cold call 1-2 students to share.

TEACHER NOTES:
Different numbers from We Do. Goal: students use the addition fact to find both subtraction facts.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use 5 + 5 = 10. What is 10 take away 5?
- Extra Notes: Sit with students; counters available.
EXTENDING PROMPT:
- Task: Use 8 + 8 = 16. Make 4 facts (2 add, 2 take away) from the doubles fact.
- Extra Notes: Encourage them to write their own missing-part question for a friend.

WATCH FOR:
- Students who write 9 and 6 - secure.
- Students who write only one answer - prompt for the second part swap.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Use the fact 7 + 5 = 12. What is 12 take away 5?

DO:
- Display the prompt.
- Allow 60 seconds.
- Collect whiteboards or take a quick photo.

TEACHER NOTES:
Exit ticket assesses the core target - using a known addition fact to find a subtraction fact. The answer is 7.

WATCH FOR:
- Students who write 7 with reasoning - on track.
- Students who write 5 - they swapped the parts incorrectly. Reteach.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Today's clever strategy is think addition.
- If we know an add fact, we know two take away facts.
- Show me thumbs: did you use this strategy today?

DO:
- Read the success criteria with the class.
- Use thumbs up sideways down.
- Cold call 1-2 students for the partner share.

TEACHER NOTES:
Self-assessment data informs Lesson 3 grouping. Students who reach the second I can are ready to extend.

WATCH FOR:
- Strong thumbs up - confident.
- Sideways or down - small group focus tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- These are the materials for today.
- The fact family sheet is for table time after the lesson.

DO:
- Print the fact family practice sheet (one per student) only if you want a written record.
- Have whiteboards, markers, and counters ready.

TEACHER NOTES:
One printed student resource for this lesson. Most of the work happens on whiteboards and the Part-Part-Whole mat.

[General: Resources]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Think Addition",
    `Year 2 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slides 2-3: Daily Review with reveal — 2D shapes (sides)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ACCENT);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Which shape has 5 sides?", { color: C.ACCENT });

      // Left card: instruction
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 22, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Find the shape", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "with 5 sides.", options: { fontSize: 22, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right card: four shapes
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      const baseY = CONTENT_TOP + 0.40;
      drawShape(s, "square",   5.55, baseY,        1.10, C.PRIMARY);
      drawShape(s, "triangle", 7.15, baseY,        1.10, C.SECONDARY);
      drawShape(s, "pentagon", 5.55, baseY + 1.55, 1.10, C.ACCENT);
      drawShape(s, "hexagon",  7.40, baseY + 1.55, 1.10, C.ALERT);

      s.addText("square",   { x: 5.40, y: baseY + 1.20, w: 1.40, h: 0.30, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0, bold: true });
      s.addText("triangle", { x: 7.00, y: baseY + 1.20, w: 1.40, h: 0.30, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0, bold: true });
      s.addText("pentagon", { x: 5.30, y: baseY + 2.50, w: 1.65, h: 0.30, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0, bold: true });
      s.addText("hexagon",  { x: 7.20, y: baseY + 2.50, w: 1.40, h: 0.30, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0, bold: true });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Pentagon = 5 sides.   Hexagon = 6 sides.   Square = 4. Triangle = 3.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — skip counting in 5s
  fluencySlide(pres, "Fluency: Skip counting",
    ["5, 10, 15, 20, ?", "10, 20, 30, ?"],
    NOTES_FLUENCY, FOOTER);

  // Slide 5: LI/SC
  liSlide(pres,
    "We are learning to use a known addition fact to find a subtraction fact.",
    [
      "I can find the whole and the parts on a Part-Part-Whole mat.",
      "I can use an addition fact to find a missing part.",
      "I can say two take away facts from one addition fact.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — 8 + 7 = 15  =>  15 - 7 = 8 and 15 - 8 = 7
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "If I know 8 + 7 = 15...", { color: STAGE_COLORS["2"] });

    // Top fact strip
    addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["2"] });
    s.addText("8 + 7 = 15", {
      x: 0.5, y: CONTENT_TOP + 0.10, w: 9.0, h: 0.65,
      fontSize: 32, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // PPW bar in the middle
    drawPpwBar(s, 3.0, CONTENT_TOP + 1.05, 4.0, 1.65, 15, 8, 7);

    // Two derived facts on the bottom
    addCard(s, 0.5, CONTENT_TOP + 2.85, 4.30, 0.85, { strip: C.SECONDARY });
    s.addText("15 - 7 = 8", {
      x: 0.5, y: CONTENT_TOP + 2.95, w: 4.30, h: 0.65,
      fontSize: 26, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addCard(s, 5.20, CONTENT_TOP + 2.85, 4.30, 0.85, { strip: C.SECONDARY });
    s.addText("15 - 8 = 7", {
      x: 5.20, y: CONTENT_TOP + 2.95, w: 4.30, h: 0.65,
      fontSize: 26, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
  })();

  // Slides 7-8: CFU 1 with reveal — Use 6 + 4 = 10 to find 10 - 4
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Use 6 + 4 = 10", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.7, y: 0.20, w: 1.7, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });

      // PPW bar with whole 10, part 4, part ?
      drawPpwBar(s, 2.8, CONTENT_TOP + 0.40, 4.4, 1.95, 10, null, 4);

      s.addText("What is 10 take away 4?   Show me on your fingers.", {
        x: 0.5, y: CONTENT_TOP + 2.55, w: 9.0, h: 0.55,
        fontSize: 24, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "10 take away 4 is 6.   (because 6 + 4 = 10)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 9-10: We Do with reveal — Use 5 + 4 = 9 to find 9 - 5
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Use 5 + 4 = 9", { color: STAGE_COLORS["3"] });

      // Top fact strip
      addCard(s, 0.5, CONTENT_TOP, 9.0, 0.85, { strip: STAGE_COLORS["3"] });
      s.addText("5 + 4 = 9", {
        x: 0.5, y: CONTENT_TOP + 0.10, w: 9.0, h: 0.65,
        fontSize: 28, fontFace: FONT_H, color: STAGE_COLORS["3"], bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // PPW bar with whole 9, part 5, part ?
      drawPpwBar(s, 2.8, CONTENT_TOP + 1.05, 4.4, 1.95, 9, 5, null);

      s.addText("9 take away 5 = ____   (Whisper to your partner. Then write it.)", {
        x: 0.5, y: CONTENT_TOP + 3.20, w: 9.0, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "9 take away 5 is 4.   (because 5 + 4 = 9)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 11-12: CFU Hinge with reveal — Is the reasoning right?
  withReveal(
    () => cfuSlide(pres, "CFU", "Is this right?", "Thumbs Up or Thumbs Down",
      "I know 7 + 6 = 13.\n\nSo 13 take away 6 is 7.\n\nIs that right?",
      NOTES_CFU2_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs UP.   13 take away 6 is 7.   13 take away 7 is 6.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: You Do — Use 9 + 6 = 15
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: 9 + 6 = 15", { color: STAGE_COLORS["4"] });

    // Top steps strip
    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.15, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 22, color: C.ALERT, bold: true } },
      { text: "Find 15 take away 6.   ", options: { fontSize: 22, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 22, color: C.ALERT, bold: true } },
      { text: "Find 15 take away 9.   ", options: { fontSize: 22, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 22, color: C.ALERT, bold: true } },
      { text: "Tell your partner.", options: { fontSize: 22, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.85,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // PPW bar in lower half (whole 15, parts 9 and 6 shown)
    const panelY = CONTENT_TOP + 1.30;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    // Show the bar with one missing-part question highlighted as ?
    drawPpwBar(s, 3.0, panelY + 0.30, 4.0, 1.65, 15, 9, 6);

    s.addText("On your whiteboard. Write both take away facts.", {
      x: 0.5, y: panelY + 2.10, w: 9.0, h: 0.36,
      fontSize: 16, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 14: Exit Ticket
  exitTicketSlide(pres,
    [
      "Use 7 + 5 = 12. Find 12 take away 5. Show your thinking.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 15: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell: how does an add fact help us subtract?",
      scItems: [
        "I can find the whole and the parts on a Part-Part-Whole mat.",
        "I can use an addition fact to find a missing part.",
        "I can say two take away facts from one addition fact.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Slide 16: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Subs_Lesson2_Think_Addition.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Fact family practice sheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Use the addition fact to find two take away facts.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addTipBox(doc,
      "If you know an add fact, you know two take away facts. Whole stays the same. Parts can swap.",
      y, { color: C.ACCENT });

    // Helper to draw a PPW bar in the PDF
    function pdfPpwBar(yTop, whole) {
      const x = 90;
      const w = 420;
      const wholeH = 44;
      const partH  = 60;
      doc.lineWidth(2).rect(x, yTop, w, wholeH).fillAndStroke("#" + C.PRIMARY, "#333333");
      doc.fillColor("#FFFFFF").fontSize(20).font("Sans-Bold").text("Whole = " + whole, x, yTop + 12, { width: w, align: "center" });
      const partsY = yTop + wholeH + 8;
      const partGap = 10;
      const partW = (w - partGap) / 2;
      doc.lineWidth(2).rect(x, partsY, partW, partH).fillAndStroke("#FFFFFF", "#333333");
      doc.lineWidth(2).rect(x + partW + partGap, partsY, partW, partH).fillAndStroke("#FFFFFF", "#333333");
      doc.fillColor("#333333").font("Sans-Bold").fontSize(13).text("Part", x + 8, partsY + 6);
      doc.fillColor("#333333").font("Sans-Bold").fontSize(13).text("Part", x + partW + partGap + 8, partsY + 6);
      doc.font("Sans");
      return yTop + wholeH + partH + 16;
    }

    function addQuestion(label, factText, whole, yPos) {
      doc.fontSize(14).fillColor("#" + C.PRIMARY).font("Sans-Bold").text(label + "    " + factText, 70, yPos);
      let yy = pdfPpwBar(yPos + 22, whole);
      doc.fillColor("#333333").font("Sans").fontSize(13)
        .text("Take away facts:   _______ - _______ = _______", 90, yy)
        .text("                    _______ - _______ = _______", 90, yy + 22);
      return yy + 56;
    }

    y = addSectionHeading(doc, "Practice", y, { color: C.PRIMARY });
    y = addQuestion("a)", "Use 6 + 4 = 10",  10, y);
    y = addQuestion("b)", "Use 8 + 5 = 13",  13, y);

    addPdfFooter(doc, `Lesson ${SESSION} | Fact Family Practice | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the fact family practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 2 Numeracy`,
    });
    y = addBodyText(doc, "Each addition fact gives two take away facts. The parts can swap.", y);
    y = addSectionHeading(doc, "Answers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  6 + 4 = 10    =>    10 - 4 = 6     and     10 - 6 = 4", y);
    y = addBodyText(doc, "b)  8 + 5 = 13    =>    13 - 5 = 8     and     13 - 8 = 5", y);
    y = addTipBox(doc,
      "If a student writes only one fact, prompt: \"What if the parts swap?\" Both subtraction facts come from the one addition fact.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 2 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
