"use strict";

// Part Part Whole Unit — Session 2: Unifix Tower Ways
// Foundation Numeracy | Term 2 Week 4 | Variant 3
// DR: Part Part Whole 6-10 with fingers
// Fluency: Counting backwards from 20
// VC2MFN04 — partition and combine collections up to 10 using part-part-whole relationships

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

const SESSION = 2;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = "Part, Part, Whole | Session 2 of 4 | Foundation Numeracy";
const OUT_DIR = "output/PPW_Session2_Unifix_Tower_Ways";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Tower Recording Sheet", "Record different ways to build a number with unifix towers.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "All the ways to make numbers 5-10 with two-colour towers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-drawn tower outlines for numbers 4-6.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Find every way to build a tower of 10.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a vertical unifix tower: bottom-up, first redCount cubes red, then blue
function drawUnifixTower(slide, x, bottomY, cubeH, cubeW, redCount, blueCount, opts) {
  const o = opts || {};
  const colorA = o.colorA || "D64545"; // red
  const colorB = o.colorB || "3D6CB4"; // blue
  const border = C.CHARCOAL;
  const total = redCount + blueCount;
  for (let i = 0; i < total; i++) {
    const y = bottomY - (i + 1) * cubeH;
    slide.addShape("rect", {
      x, y, w: cubeW, h: cubeH,
      fill: { color: i < redCount ? colorA : colorB },
      line: { color: border, width: 1 },
    });
    // Small stud circle on top of each cube
    slide.addShape("roundRect", {
      x: x + cubeW * 0.38, y: y - cubeH * 0.08, w: cubeW * 0.24, h: cubeH * 0.16, rectRadius: cubeH * 0.08,
      fill: { color: i < redCount ? colorA : colorB },
      line: { color: border, width: 0.6 },
    });
  }
}

// Draw hand showing fingers: simple representation - circles in a row
function drawFingerRow(slide, x, y, upCount, totalCount, opts) {
  const o = opts || {};
  const size = o.size || 0.36;
  const gap = o.gap || 0.06;
  for (let i = 0; i < totalCount; i++) {
    const cx = x + i * (size + gap);
    slide.addShape("roundRect", {
      x: cx, y: i < upCount ? y : y + size * 0.35, w: size, h: i < upCount ? size * 1.3 : size * 0.9, rectRadius: size * 0.22,
      fill: { color: i < upCount ? "FFD9A3" : "E8D0B0" },
      line: { color: "A06040", width: 1 },
    });
  }
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Good morning! Yesterday we shook and spilled counters to find parts
- Today we use unifix blocks to build towers
- We will find different ways to make the same number

DO:
- Have unifix cubes in two colours ready on each table
- Display title slide as students settle

TEACHER NOTES:
Session 2 of 4. Builds on Session 1 by moving from flat counter spills to vertical towers. Towers make "same height, different parts" very visible when students line them up side by side.

WATCH FOR:
- Students eager to start building - they remember yesterday
- Students who seem to have forgotten PPW - use the Daily Review to retrieve it

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up your brain! Look at the hands on the slide
- How many fingers are up in total? That is the whole.
- How many fingers on the LEFT hand? That is one part.
- How many on the RIGHT hand? That is the other part.

DO:
- Read the first hand picture with the class
- Partners whisper their answers
- Allow 15 seconds per picture

TEACHER NOTES:
Daily Review on PPW 6-10. Uses hands as the representation because fingers are something every student carries. Keeps yesterday's whole-and-parts language warm.

WATCH FOR:
- Students who say whole and parts in full sentences - strong retention from yesterday
- Students who say just numbers - model the sentence pattern

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers
- For each hand picture, say: The whole is ___. The parts are ___ and ___.
- Tick if you got it. Fix it if you need to.

DO:
- Click to reveal answers
- Choral say the PPW pattern

TEACHER NOTES:
Tick-and-fix on PPW 6-10 extends the number range from yesterday (1-5) into today's focus (6-10).

WATCH FOR:
- Students who confidently say "8 is 5 and 3" - they see the bigger numbers as splittable
- Students who miscount total fingers - work with them during You Do

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Counting fluency. Today we count BACKWARDS.
- Start at 20. Count down all the way to 0.
- 20, 19, 18... keep going!
- Now we try from 10. Ready? 10, 9, 8...

DO:
- Lead choral countdown from 20 to 0
- Then countdown from 10
- Use fingers or an arm movement to mark the rhythm

TEACHER NOTES:
Backwards counting is harder than forward counting. Do not rush. If students get stuck in the teens, that is normal - pause and model the tricky parts.

WATCH FOR:
- Students who stumble at 13-12-11 - teen numbers are the common sticking point
- Students who count down fluently - backwards counting is automated

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- We are learning to partition numbers to 10 into two parts using materials
- Our materials today are unifix blocks
- Let's read the success criteria

DO:
- Choral read LI and SC
- Remind: partition means to split into parts

TEACHER NOTES:
Same LI and SC as Session 1 - the unit runs across 4 lessons. Repetition of the criteria helps Foundation students know what they are working towards.

WATCH FOR:
- Students reading along confidently - familiarity with the language is building
- Students who look at the ground - make eye contact and choral read again

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me build a tower
- I will make the whole of 6
- I will use red cubes and blue cubes
- Red cubes first: 1, 2, 3, 4
- Now blue cubes: 5, 6
- I have 4 red and 2 blue. The whole is 6.

DO:
- Build a real tower of 6 cubes in front of the students
- 4 red on the bottom, 2 blue on top
- Say each cube aloud as you click it on
- Point to the visual on the slide

TEACHER NOTES:
This I Do shows a tower as a PPW model. The tower height equals the whole. The two colours equal the two parts. Say "whole" as you point to the full tower and "parts" as you point to each colour section.

MISCONCEPTIONS:
- Misconception: Students think each cube is a separate number, so a tower of 6 is "six ones"
  Why: Counting habit is number-by-number; seeing chunks takes experience
  Impact: They cannot say 6 as 4 and 2 - they only see 1,2,3,4,5,6
  Quick correction: "Look at the red chunk. How many red? [4] And the blue chunk? [2] So the whole tower is 4 and 2. That makes 6."

WATCH FOR:
- Students who say "4 and 2 makes 6" with you - they see the chunks
- Students who count every cube - help them see the colour groups

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Let's make 6 a different way
- This time: 3 red cubes and 3 blue cubes
- Count them: 1, 2, 3 red. 1, 2, 3 blue.
- The whole is still 6. The parts are now 3 and 3.
- Put the two towers side by side. Same height!
- Same whole. Different parts.

DO:
- Build a second tower: 3 red + 3 blue
- Stand it next to the first tower (4 red + 2 blue)
- Point to the tops - both reach the same level
- Repeat: "Same whole. Different parts."

TEACHER NOTES:
The side-by-side tower comparison makes "same whole, different parts" visible. The heights match because both totals are 6. This is the conceptual core of the lesson.

WATCH FOR:
- Students who notice the heights match - they see the whole is invariant
- Students pointing at the colour difference - they see the parts are different
- Students who think the second tower is a different number - reteach with physical cubes

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Look at the tower on the slide
- Count the red cubes. Count the blue cubes.
- Show me on your fingers how many blue
- Turn and tell your partner: What is the whole? What are the parts?

DO:
- Display tower: 5 red + 1 blue (whole of 6)
- Scan fingers for blue count
- Cold call 2 students

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "How many blue cubes? Show me on your fingers."
- Scan for: 1 finger
PROCEED: If 80%+ show 1, move to the We Do
PIVOT: Most likely misconception - students count every cube from 1. Reteach: "Just count the BLUE part. Ignore the red for now. Point at only blue. How many?"

TEACHER NOTES:
Fingers work for every student. This check is about whether students can isolate one colour group (one part) from the whole tower.

WATCH FOR:
- Students showing 1 immediately - they can isolate the parts
- Students showing 5 or 6 - they are still counting the whole

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me
- I rolled a dice and got 7
- Let's find TWO different ways to build 7 with our tower
- With your partner, write two ways on your whiteboard
- 7 = ___ and ___     7 = ___ and ___

DO:
- Show the dice face: 7
- Give partners 60 seconds to discuss and write
- Cold call 3-4 pairs to share their ways

TEACHER NOTES:
We Do uses 7 to stretch students past the Daily Review range. Finding TWO different ways requires flexibility. Some pairs will write the same pair twice - prompt them to find a different way.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 2 Enabling Scaffold has pre-drawn tower outlines for the number 5. Students colour each tower with two colours to show one way at a time.
- Extra Notes: Hand out to students who struggled in the Daily Review.
EXTENDING PROMPT:
- Task: Find ALL the ways to make 7. How many are there? Record every pair.
- Extra Notes: Expect: 7+0, 6+1, 5+2, 4+3, 3+4, 2+5, 1+6, 0+7 = 8 ways.

WATCH FOR:
- Students who write two different pairs - they see multiple ways
- Students who write the same pair twice - ask "Is there another way?"

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's share ways to make 7
- 7 = 3 and 4   or   7 = 5 and 2   or   7 = 6 and 1...
- All of these are right! Same whole, different parts.

DO:
- Reveal the answer panel
- Record on the class board any other ways students found
- Celebrate flexibility - there is no single right answer

TEACHER NOTES:
The reveal is a list of correct pairs, not a single answer. This normalises that numbers have many part pairs.

WATCH FOR:
- Students calling out pairs not on the slide - they are generalising
- Students who cling to one pair - they may need Session 3 before the idea clicks

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Thumbs up if I am right, thumbs down if I am wrong.
- The whole is 7. I say the parts are 4 and 3. Is that right?
- Think... show me.
- Now try this one. The whole is 7. Parts are 5 and 3. Is that right?

DO:
- Thumbs up for 4 and 3 (correct - equals 7)
- Thumbs down for 5 and 3 (wrong - equals 8)
- Cold call: "How do you know it is wrong?"

CFU CHECKPOINT:
Technique: Thumbs Up/Down
Script:
- Say: "Whole is 7, parts are 4 and 3. Right or wrong? Show me."
- Scan for: thumbs up
- Then: "Whole is 7, parts are 5 and 3. Right or wrong?"
- Scan for: thumbs down
PROCEED: If 80%+ get both, move to You Do
PIVOT: Most likely misconception - students check the parts but not the sum. Reteach: "Let's count the parts together. 5 and 3. 1, 2, 3, 4, 5... 6, 7, 8. That is 8, not 7. So the parts are wrong for a whole of 7."

TEACHER NOTES:
Thumbs lets every student respond. The wrong example is critical - it tests whether students are actually checking, not just agreeing.

WATCH FOR:
- Thumbs down with confidence on the wrong one - students are checking the sum
- Hesitation or copying - students are guessing rather than reasoning

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own (or with your partner)
- First: Roll the dice. Read the number.
- Next: Build that number with your unifix tower in two colours.
- Then: Record your tower on your sheet. Then build it a different way.
- You have 10 minutes.

DO:
- Hand out the Session 2 Tower Recording Sheet
- Each pair needs: unifix cubes in 2 colours, one 10-sided dice
- Circulate: start with students who needed support
- Enabler students get the Session 2 Enabling Scaffold
- Extender students get the Session 2 Extension

TEACHER NOTES:
You Do applies the class routine independently. Partners share materials but each records on their own sheet.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 2 Enabling Scaffold has pre-drawn tower outlines for 4, 5, and 6. Students colour one tower at a time with two colours.
- Extra Notes: Numbers 4-6 only. One way per tower.
EXTENDING PROMPT:
- Task: Session 2 Extension - find every way to build a tower of 10. Record each way. How many ways are there?
- Extra Notes: Self-contained with tower outlines.

WATCH FOR:
- Students building and recording independently - target learning is happening
- Students who need help reading the dice - that is a counting issue, not a PPW issue

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Work on your own.
- Look at each question. Write or draw your answer.
- You have 3 minutes.

DO:
- Display the exit ticket slide
- Students respond in workbook or on whiteboard
- Collect or scan to sort students

TEACHER NOTES:
Exit ticket assesses SC1 (Q1 - build one way), SC2 (Q2 - find a different way), and SC3 (Q3 - explain in words).

WATCH FOR:
- Students who draw two different towers for the same whole - they have mastered the idea
- Students who draw the same tower twice - they need Session 3 practice

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's review our success criteria
- Thumbs up, sideways, or down for each
- Turn and tell: What is the best thing you learned today about making numbers?

DO:
- Show the success criteria
- Thumbs check for each
- 30 seconds Turn and Talk
- Cold call 2 students to share

TEACHER NOTES:
Closing connects today's tower work back to the overall PPW goal. Students should recognise that the tower makes the whole and parts visible.

WATCH FOR:
- Students saying "different ways" or "same whole" in their share - core concepts forming
- Students naming materials (cubes, dice) rather than concepts - normal, more practice coming

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Here are today's resources
- The Tower Recording Sheet is the main one
- Enabler and Extension are on your differentiation shelf

DO:
- Point out each resource

TEACHER NOTES:
Materials needed: unifix cubes (two colours per student/pair), one 10-sided dice per pair, whiteboards, the Tower Recording Sheet.

WATCH FOR:
- N/A - teacher reference only

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 2: Unifix Tower Ways",
    "Foundation Numeracy | Session 2 of 4 | Term 2 Week 4", NOTES_TITLE);

  // Slide 2-3: Daily Review (PPW 6-10 with fingers) withReveal
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Hand Parts", { y: 0.65, fontSize: 24, color: STAGE_COLORS["1"] });

      // Left: instruction card
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText([
        { text: "Look at each pair of hands.", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "How many LEFT?", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "How many RIGHT?", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "What is the whole?", options: { bullet: true, fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: three finger pairs
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      const pairs = [
        { label: "A.", left: 4, right: 2 },
        { label: "B.", left: 3, right: 4 },
        { label: "C.", left: 5, right: 5 },
      ];
      pairs.forEach((p, i) => {
        const rowY = CONTENT_TOP + 0.25 + i * 1.05;
        s.addText(p.label, {
          x: 5.3, y: rowY, w: 0.35, h: 0.6,
          fontSize: 20, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
          align: "center", valign: "middle",
        });
        // Left hand
        drawFingerRow(s, 5.7, rowY, p.left, 5, { size: 0.22, gap: 0.04 });
        // Gap
        // Right hand
        drawFingerRow(s, 7.55, rowY, p.right, 5, { size: 0.22, gap: 0.04 });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A) 6 = 4 + 2    B) 7 = 3 + 4    C) 10 = 5 + 5", {
        x: 0.8, y: 4.55, w: 8.4, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Count backwards from 20
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Count Backwards from 20", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFluency.addText([
    { text: "Stand up!", options: { fontSize: 20, bold: true, color: C.ALERT, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Start at 20.", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
    { text: "Count down to 0.", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
    { text: "Now try from 10!", options: { bullet: true, fontSize: 16, color: C.CHARCOAL } },
  ], {
    x: 0.75, y: CONTENT_TOP + 0.2, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  // Right: countdown arrow visual 20 → 0
  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
  const countdownNums = [20, 19, 18, 17, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0];
  // Show as a descending ladder of the first 11
  const ladderX = 5.4;
  const ladderY = CONTENT_TOP + 0.2;
  const stepH = 0.25;
  for (let i = 0; i < 11; i++) {
    const num = countdownNums[i];
    sFluency.addShape("roundRect", {
      x: ladderX, y: ladderY + i * stepH, w: 1.3, h: stepH - 0.04, rectRadius: 0.05,
      fill: { color: i === 0 ? C.ALERT : C.PRIMARY },
      line: { color: C.PRIMARY, width: 0.8 },
    });
    sFluency.addText(String(num), {
      x: ladderX, y: ladderY + i * stepH, w: 1.3, h: stepH - 0.04,
      fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }
  sFluency.addText("...keep going to 0!", {
    x: 6.9, y: ladderY + 0.8, w: 2.3, h: 0.4,
    fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, italic: true,
    align: "left", valign: "middle", margin: 0,
  });
  sFluency.addText("Starts at", {
    x: 6.9, y: ladderY + 0.05, w: 2.3, h: 0.3,
    fontSize: 14, fontFace: FONT_B, color: C.MUTED, margin: 0,
    align: "left", valign: "middle", italic: true,
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

  // Slide 6: I Do — Build a tower of 6
  workedExSlide(pres, 2, "I Do", "Build a Tower of 6",
    [
      "I need 6 cubes",
      "",
      "4 red cubes at the bottom",
      "2 blue cubes on top",
      "",
      "Red part = 4",
      "Blue part = 2",
      "",
      "Whole = 6",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.5, { strip: C.PRIMARY });

      // Draw unifix tower of 6: 4 red + 2 blue
      const cubeW = 0.85;
      const cubeH = 0.45;
      const towerX = lg.rightX + (lg.rightW - cubeW) / 2;
      const towerBottom = lg.panelTopPadded + 3.3;
      drawUnifixTower(slide, towerX, towerBottom, cubeH, cubeW, 4, 2);

      slide.addText("Tower = 6", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.2, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
    }
  );

  // Slide 7: I Do #2 — Same whole, different parts (side-by-side)
  workedExSlide(pres, 2, "I Do", "Same Whole, Different Parts",
    [
      "Still making 6",
      "",
      "Tower 1: 4 red, 2 blue",
      "Tower 2: 3 red, 3 blue",
      "",
      "Both are 6 cubes tall",
      "",
      "Same whole!",
      "Different parts!",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.5, { strip: C.SECONDARY });

      const cubeW = 0.55;
      const cubeH = 0.36;
      const towerBottom = lg.panelTopPadded + 3.2;

      // Tower 1: 4 red + 2 blue
      const x1 = lg.rightX + 0.5;
      drawUnifixTower(slide, x1, towerBottom, cubeH, cubeW, 4, 2);
      slide.addText("4 + 2", {
        x: x1 - 0.1, y: towerBottom + 0.02, w: 0.75, h: 0.28,
        fontSize: 12, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0, align: "center",
      });

      // Tower 2: 3 red + 3 blue
      const x2 = lg.rightX + lg.rightW - cubeW - 0.5;
      drawUnifixTower(slide, x2, towerBottom, cubeH, cubeW, 3, 3);
      slide.addText("3 + 3", {
        x: x2 - 0.1, y: towerBottom + 0.02, w: 0.75, h: 0.28,
        fontSize: 12, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0, align: "center",
      });

      // Header
      slide.addText("Two towers of 6", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
    }
  );

  // Slide 8-9: CFU 1 — Tower of 6 (5 red + 1 blue)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["2"]);
      addStageBadge(s, 2, "CFU");
      addTitle(s, "How Many Blue?", { y: 0.65, fontSize: 24, color: STAGE_COLORS["2"] });

      // Left: prompt
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["2"] });
      s.addText([
        { text: "Show me on your fingers:", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "How many BLUE cubes?", options: { fontSize: 22, bold: true, color: C.ALERT, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Then say the parts:", options: { fontSize: 14, color: C.MUTED, breakLine: true } },
        { text: "Whole is ___. Parts are ___ and ___.", options: { fontSize: 14, italic: true, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.2, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: tower 5 red + 1 blue
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText("This tower:", {
        x: 5.35, y: CONTENT_TOP + 0.15, w: 4, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });
      const cubeW = 0.75;
      const cubeH = 0.42;
      const towerX = 5.2 + (4.3 - cubeW) / 2;
      const towerBottom = CONTENT_TOP + 3.2;
      drawUnifixTower(s, towerX, towerBottom, cubeH, cubeW, 5, 1);

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1 blue!   Whole is 6.  Parts are 5 and 1.", {
        x: 1.0, y: 4.1, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do — Roll 7, find 2 ways
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "The Dice Rolled 7", { y: 0.65, fontSize: 24, color: STAGE_COLORS["3"] });

      // Left: prompt
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "With your partner:", options: { fontSize: 18, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Find TWO ways to make 7", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "Different parts each time", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Write on your whiteboard:", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "7 = ___ and ___", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "7 = ___ and ___", options: { fontSize: 18, bold: true, color: C.PRIMARY } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: dice face + prompt
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      // Big dice with 7
      s.addShape("roundRect", {
        x: 6.2, y: CONTENT_TOP + 0.35, w: 2.3, h: 1.8, rectRadius: 0.18,
        fill: { color: C.WHITE },
        line: { color: C.ACCENT, width: 4 },
      });
      s.addText("7", {
        x: 6.2, y: CONTENT_TOP + 0.35, w: 2.3, h: 1.8,
        fontSize: 110, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("Build a tower of 7", {
        x: 5.35, y: CONTENT_TOP + 2.3, w: 4, h: 0.4,
        fontSize: 16, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });
      s.addText("Two different ways!", {
        x: 5.35, y: CONTENT_TOP + 2.8, w: 4, h: 0.35,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, margin: 0, align: "center",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "7 = 6+1, 5+2, 4+3, 3+4, 2+5, 1+6, 7+0 or 0+7", {
        x: 0.6, y: 4.55, w: 8.8, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU Hinge
  withReveal(
    () => cfuSlide(pres, "CFU", "Is This a Way to Make 7?", "Thumbs Up/Down",
      "Whole = 7\n\nParts = 5 and 3\n\nIs that right?\n\nThumbs UP = YES    Thumbs DOWN = NO",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN!  5 and 3 make 8, not 7", {
        x: 1.0, y: 4.1, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do — Roll and build
  workedExSlide(pres, 4, "You Do", "Your Turn: Roll and Build",
    [
      "First: Roll the 10-sided dice",
      "",
      "Next: Build the tower with 2 colours",
      "",
      "Then: Draw your tower on the sheet",
      "",
      "Build it a DIFFERENT way",
      "",
      "You have 10 minutes",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.ALERT });
      slide.addText("You need:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "2 colours of cubes", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "a 10-sided dice", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "your recording sheet", options: { bullet: true, fontSize: 15, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Tip card
      addCard(slide, lg.rightX, lg.panelTopPadded + 2.2, lg.rightW, 1.1, { strip: C.PRIMARY });
      slide.addText("Check: Are your two towers the same height?", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.35, w: lg.rightW - 0.3, h: 0.7,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, italic: true, margin: 0,
        align: "center", valign: "middle",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Build a tower of 5. Draw it on your whiteboard with 2 colours.",
      "Show a DIFFERENT way to make 5. Draw it.",
      "Tell your teacher why both towers are the same height.",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "Turn and tell: What is one way to make the number 8?",
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
  const pptxPath = path.join(OUT_DIR, "PPW_Session2_Unifix_Tower_Ways.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Tower Recording Sheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Roll the dice. Build the tower. Draw it. Then build it a different way.",
      color: C.PRIMARY,
      lessonInfo: "Session 2 of 4 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Roll the dice. Build a tower with the number you rolled using 2 colours. Colour the cubes to match your tower. Then build a DIFFERENT way.", y, { color: C.ACCENT });

    const drawTowerOutline = (ty, cubes) => {
      const cubeW = 30;
      const cubeH = 28;
      const startX = 80;
      // Draw from bottom up
      for (let i = 0; i < cubes; i++) {
        const cubeY = ty - (i + 1) * cubeH;
        doc.rect(startX, cubeY, cubeW, cubeH).lineWidth(1.2).stroke("#333333");
      }
      return ty + 10;
    };

    y = addSectionHeading(doc, "Round 1:  I rolled ______", y, { color: C.PRIMARY });
    doc.fontSize(12).fillColor("#333333").text("Way 1:", 60, y);
    const a1Top = y + 20 + 7 * 28;
    drawTowerOutline(a1Top, 7);
    doc.fontSize(12).fillColor("#333333").text("______ + ______ = ______", 180, a1Top - 120);
    doc.fontSize(12).fillColor("#333333").text("Way 2:", 280, y);
    const a2Top = y + 20 + 7 * 28;
    drawTowerOutline(a2Top, 7);
    // Shift the second tower to the right - redraw
    // Simpler: stack vertically
    y = a1Top + 30;

    y = addSectionHeading(doc, "Round 2:  I rolled ______", y, { color: C.PRIMARY });
    doc.fontSize(12).fillColor("#333333").text("Way 1: colour the cubes", 60, y);
    const b1Top = y + 20 + 7 * 28;
    drawTowerOutline(b1Top, 7);
    doc.fontSize(12).fillColor("#333333").text("______ + ______ = ______", 180, b1Top - 120);
    y = b1Top + 30;

    y = addSectionHeading(doc, "Round 3:  I rolled ______", y, { color: C.PRIMARY });
    doc.fontSize(12).fillColor("#333333").text("Way 1: colour the cubes", 60, y);
    const c1Top = y + 20 + 7 * 28;
    drawTowerOutline(c1Top, 7);
    doc.fontSize(12).fillColor("#333333").text("______ + ______ = ______", 180, c1Top - 120);

    addPdfFooter(doc, "Session 2 | Tower Recording Sheet | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "All the ways to make 5 through 10",
      color: C.PRIMARY,
      lessonInfo: "Session 2 of 4 | Foundation Numeracy",
    });
    y = addBodyText(doc, "Students can record any correct pair of parts. These lists show all the ways for each whole.", y);

    y = addSectionHeading(doc, "Ways to make 5", y, { color: C.PRIMARY });
    y = addBodyText(doc, "5 + 0, 4 + 1, 3 + 2, 2 + 3, 1 + 4, 0 + 5  (6 ways)", y);

    y = addSectionHeading(doc, "Ways to make 6", y, { color: C.PRIMARY });
    y = addBodyText(doc, "6 + 0, 5 + 1, 4 + 2, 3 + 3, 2 + 4, 1 + 5, 0 + 6  (7 ways)", y);

    y = addSectionHeading(doc, "Ways to make 7", y, { color: C.PRIMARY });
    y = addBodyText(doc, "7 + 0, 6 + 1, 5 + 2, 4 + 3, 3 + 4, 2 + 5, 1 + 6, 0 + 7  (8 ways)", y);

    y = addSectionHeading(doc, "Ways to make 8", y, { color: C.PRIMARY });
    y = addBodyText(doc, "8 + 0, 7 + 1, 6 + 2, 5 + 3, 4 + 4, 3 + 5, 2 + 6, 1 + 7, 0 + 8  (9 ways)", y);

    y = addSectionHeading(doc, "Ways to make 9", y, { color: C.PRIMARY });
    y = addBodyText(doc, "9 + 0, 8 + 1, 7 + 2, 6 + 3, 5 + 4, 4 + 5, 3 + 6, 2 + 7, 1 + 8, 0 + 9  (10 ways)", y);

    y = addSectionHeading(doc, "Ways to make 10", y, { color: C.PRIMARY });
    y = addBodyText(doc, "10 + 0, 9 + 1, 8 + 2, 7 + 3, 6 + 4, 5 + 5, 4 + 6, 3 + 7, 2 + 8, 1 + 9, 0 + 10  (11 ways)", y);

    y = addTipBox(doc, "Pattern to notice: a whole of N has N + 1 ways to be made (including 0 + N).", y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 2 | Answer Key | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — towers of 4, 5, 6
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Colour each tower with 2 colours",
      color: C.ACCENT,
      lessonInfo: "Session 2 of 4 | Foundation Numeracy",
    });
    y = addTipBox(doc, "The towers are already drawn. Colour some cubes red and some cubes blue. Then write the parts.", y, { color: C.ACCENT });

    const drawCubes = (ty, count) => {
      const cubeW = 36;
      const cubeH = 32;
      const startX = 80;
      for (let i = 0; i < count; i++) {
        const cubeY = ty - (i + 1) * cubeH;
        doc.rect(startX, cubeY, cubeW, cubeH).lineWidth(1.2).stroke("#333333");
      }
      return ty + 10;
    };

    y = addSectionHeading(doc, "Tower of 4", y, { color: C.PRIMARY });
    const t1 = y + 20 + 4 * 32;
    drawCubes(t1, 4);
    doc.fontSize(13).fillColor("#333333").text("Parts:  ______ red and ______ blue", 150, t1 - 60);
    y = t1 + 20;

    y = addSectionHeading(doc, "Tower of 5", y, { color: C.PRIMARY });
    const t2 = y + 20 + 5 * 32;
    drawCubes(t2, 5);
    doc.fontSize(13).fillColor("#333333").text("Parts:  ______ red and ______ blue", 150, t2 - 60);
    y = t2 + 20;

    y = addSectionHeading(doc, "Tower of 6", y, { color: C.PRIMARY });
    const t3 = y + 20 + 6 * 32;
    drawCubes(t3, 6);
    doc.fontSize(13).fillColor("#333333").text("Parts:  ______ red and ______ blue", 150, t3 - 60);
    y = t3 + 20;

    addPdfFooter(doc, "Session 2 | Enabling Scaffold | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — all ways to make 10
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Find every way to build a tower of 10",
      color: C.PRIMARY,
      lessonInfo: "Session 2 of 4 | Foundation Numeracy",
    });
    y = addBodyText(doc, "Your challenge: build a tower of 10 using 2 colours. Find every different way. How many ways are there?", y);

    const drawCubes = (ty, count) => {
      const cubeW = 24;
      const cubeH = 24;
      const startX = 60;
      for (let i = 0; i < count; i++) {
        const cubeY = ty - (i + 1) * cubeH;
        doc.rect(startX + Math.floor(count / 10) * 30, cubeY, cubeW, cubeH).lineWidth(1).stroke("#333333");
      }
    };

    const drawTowersRow = (ty, ways) => {
      const cubeW = 22;
      const cubeH = 22;
      const gap = 60;
      const startX = 60;
      ways.forEach((count, idx) => {
        const towerX = startX + idx * gap;
        for (let i = 0; i < count; i++) {
          const cubeY = ty - (i + 1) * cubeH;
          doc.rect(towerX, cubeY, cubeW, cubeH).lineWidth(1).stroke("#333333");
        }
      });
    };

    y = addSectionHeading(doc, "Way 1:  10 = ______ + ______", y, { color: C.PRIMARY });
    const w1Top = y + 20 + 10 * 22;
    drawTowersRow(w1Top, [10]);
    y = w1Top + 20;

    y = addSectionHeading(doc, "Way 2:  10 = ______ + ______", y, { color: C.PRIMARY });
    const w2Top = y + 20 + 10 * 22;
    drawTowersRow(w2Top, [10]);
    y = w2Top + 20;

    y = addSectionHeading(doc, "Way 3:  10 = ______ + ______", y, { color: C.PRIMARY });
    const w3Top = y + 20 + 10 * 22;
    drawTowersRow(w3Top, [10]);
    y = w3Top + 20;

    y = addSectionHeading(doc, "Way 4:  10 = ______ + ______", y, { color: C.PRIMARY });
    const w4Top = y + 20 + 10 * 22;
    drawTowersRow(w4Top, [10]);
    y = w4Top + 30;

    y = addTipBox(doc, "Challenge: Can you find more than 4 ways? Write them on the back of this sheet. How many are there in total?", y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 2 | Extension | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 2 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
