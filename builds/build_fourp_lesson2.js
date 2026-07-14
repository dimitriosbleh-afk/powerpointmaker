"use strict";

// Four Processes (Year 6 Numeracy) - Lesson 2 of 5.
// The properties of multiplication: swap (commutative), regroup (associative)
// and split (distributive). Use them to calculate smartly and to balance
// equivalent number sentences. Division does NOT behave the same way. VC2M5A02.
// Daily Review: Fractions & Decimals (prior). Fluency: decimal x 1-digit algorithm.
// Unit variant fixed (variant 3, Ocean Logic) across all 5 lessons for cohesion.
// Catch-up: the launch rebuilds the idea from a rotated array everyone can see,
// and worksheet Section 1 is an enabling rebuild. No session assumes the one before.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 3); // variant 3 (Ocean Logic), fixed for the unit
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide, dailyReviewSlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 5;
const UNIT_TITLE = "Four Processes";
const FOOTER = `Four Processes | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FourP_Lesson2_Properties_Of_Operations";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Split multiplications into easier parts and balance number sentences.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 2 practice sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "The distributive property with algebra tiles and expanding brackets - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Reusable array helpers --------------------------------------------------
function drawArray(slide, x, y, rows, cols, cell, opts) {
  const o = opts || {};
  const fill = o.fill || C.SECONDARY;
  const gap = o.gap != null ? o.gap : 0.05;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      slide.addShape("roundRect", {
        x: x + c * (cell + gap), y: y + r * (cell + gap),
        w: cell, h: cell, rectRadius: cell * 0.18,
        fill: { color: fill }, line: { color: C.WHITE, width: 1 },
      });
    }
  }
  return { w: cols * (cell + gap) - gap, h: rows * (cell + gap) - gap };
}

// Split array for the distributive property: columns before splitCol use
// colorA, the rest use colorB, with a small visible gap between the two blocks.
function drawSplitArray(slide, x, y, rows, cols, splitCol, cell, opts) {
  const o = opts || {};
  const colorA = o.colorA || C.PRIMARY;
  const colorB = o.colorB || C.ACCENT;
  const gap = o.gap != null ? o.gap : 0.03;
  const splitGap = 0.16;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const extra = c >= splitCol ? splitGap : 0;
      slide.addShape("roundRect", {
        x: x + c * (cell + gap) + extra, y: y + r * (cell + gap),
        w: cell, h: cell, rectRadius: cell * 0.16,
        fill: { color: c < splitCol ? colorA : colorB }, line: { color: C.WHITE, width: 0.75 },
      });
    }
  }
  return { w: cols * (cell + gap) - gap + splitGap, h: rows * (cell + gap) - gap };
}

// Balance panel: two number sentences that must be equal, with = in between.
function addBalance(slide, x, y, w, leftStr, rightStr, leftColor, rightColor) {
  const chipW = (w - 0.6) / 2;
  addTextOnShape(slide, leftStr, {
    x: x, y: y, w: chipW, h: 0.62, rectRadius: 0.08,
    fill: { color: leftColor },
  }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
  slide.addText("=", {
    x: x + chipW, y: y, w: 0.6, h: 0.62,
    fontSize: 26, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  addTextOnShape(slide, rightStr, {
    x: x + chipW + 0.6, y: y, w: chipW, h: 0.62, rectRadius: 0.08,
    fill: { color: rightColor },
  }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to Four Processes. Yesterday we saw that multiplication and division are linked.
- Today we get smarter at multiplication. It has some helpful properties: we can swap, regroup and split.
- We will use those properties to make hard multiplications easy, and to keep number sentences balanced.

DO:
- Have whiteboards, markers and counters ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 2 of 5. The hero idea today is the distributive property - splitting a factor to make multiplication easier. Keep it concrete with split arrays.

WATCH FOR:
- Students who look unsure - that is expected. Reassure them: if this feels new, that is okay, we build it together.

[General: Title | Element: Planning and Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end. The Year 8 extension is there for anyone who is flying.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards, markers and counters ready for building arrays.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student practice sheet, an answer key, and a Year 8 extension. Most teaching happens on whiteboards and with split arrays.

CATCH-UP NOTE:
A student who missed earlier sessions can still access today. The launch rebuilds the idea from a rotated array everyone can see, and Section 1 of the worksheet rebuilds the split. A returner only needs counters and one minute with you to start. No session in this unit assumes the student saw the one before.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are warming up our fractions and decimals from earlier work.
- Read each one carefully and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for how students convert between fractions and decimals.

TEACHER NOTES:
Daily Review is prior learning, not today's new content. Fractions and decimals keep ticking over because we use them later this week.

WATCH FOR:
- Students who know 1/4 is 0.25 - secure.
- Students who add 0.3 and 0.4 and get 0.07 - place value slip, address it in the reveal.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 1/4 is 0.25.
- 0.3 plus 0.4 is 0.7. Three tenths plus four tenths is seven tenths.
- 3/10 is 0.3, which is larger than 0.25.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The 0.3 + 0.4 item checks tenths place value. The 3/10 versus 0.25 item rewards converting to a common form before comparing.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students who compare 3/10 and 0.25 by digit count - reteach converting first.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Multiplying a decimal by a single digit using the algorithm.
- Set each one out vertically on your whiteboard and multiply.
- Keep your columns lined up and remember where the decimal point sits.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for clean column alignment and the point kept in line.

TEACHER NOTES:
Fluency this unit is the multiplication algorithm with decimals. Today is decimal times one digit again, building automaticity before we multiply by larger numbers later in the week.

WATCH FOR:
- Students who keep the decimal point lined up - secure.
- Students who lose the point - prompt: one decimal place in, one decimal place out.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 5.3 times 4 is 21.2.
- 6.2 times 3 is 18.6.
- 4.8 times 5 is 24.0, which we write as 24.
- Tick and fix.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
4.8 times 5 lands on 24.0; accept 24. Keep it brisk.

WATCH FOR:
- Students who self-correct - secure.
- Students whose point drifts - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember building arrays. Here is 3 rows of 4. That is 3 times 4, which is 12.
- Watch what happens when I turn the same array on its side. Now it is 4 rows of 3. Still 12.
- The order swapped but the answer did not change. That is a property of multiplication: order can swap.
- Hold that thought. Today we will use properties like this to make hard multiplications easy. But careful - division does not always play by the same rules.

DO:
- Show the array, then describe rotating it.
- Have students chorus '3 times 4 and 4 times 3 both make 12'.
- Bridge: 'multiplication has helpful properties we can use'.

TEACHER NOTES:
This launch starts from a rotated array everyone can access (the commutative property), then sets up today's smart-calculating work. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who see order does not matter for times - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to use the properties of multiplication to calculate smartly and to keep number sentences balanced.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Hold up an array of counters as a reminder.

TEACHER NOTES:
The first criterion is reachable for everyone - swapping the order. The second is the core target the exit ticket checks - splitting to multiply. The third stretches to balancing sentences and explaining why division is different.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three property words for today.
- Commutative means order can swap: 3 times 4 equals 4 times 3.
- Associative means grouping can change: 2 times 2 times 3 gives 12 however you group it.
- Distributive means you can split a factor: 4 times 13 becomes 4 times 10 plus 4 times 3.

DO:
- Point to each word and its example.
- Have students say 'split it to make it easier' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. Students do not need to spell these words; they need to use the moves. 'Distributive' is the anchor for today.

WATCH FOR:
- Students who can give an example of swapping - secure.
- Students who confuse the three - focus on distributive, the one we use most today.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at this one together. 4 times 13 is awkward to do in one jump.
- Watch how I split it. I break the 13 into 10 and 3, because tens are easy.
- 4 times 10 is 40. 4 times 3 is 12. Look at the array - the blue block is 40 and the gold block is 12.
- Now I add the two parts: 40 plus 12 is 52. So 4 times 13 is 52.
- Splitting a number into friendly parts is the distributive property, and it is the smart way to multiply.

DO:
- Point to the two coloured blocks of the array as you name each part.
- Write 4 x 10 = 40 and 4 x 3 = 12, then add.
- Have students chorus 'split, multiply each part, add'.

TEACHER NOTES:
This is the core move of the lesson. Always tie the split back to the picture so students see why 4 x 13 equals 4 x 10 plus 4 x 3, not just memorise it.

MISCONCEPTIONS:
- Misconception: students split the 13 but forget to multiply BOTH parts by 4.
  Why: they multiply 4 x 10 and then just add the 3.
  Impact: they get 43 instead of 52.
  Quick correction: point to the gold block - it is 4 rows of 3, so 4 x 3 = 12, not 3.

WATCH FOR:
- Students who multiply both parts and add - secure.
- Students who add the leftover instead of multiplying it - re-point to the array block.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a different use of properties: keeping a number sentence balanced.
- Look at 3 times 5 equals 30 divided by a missing number. The equals sign means both sides have the same value.
- I work out the side I can: 3 times 5 is 15. So the right side must also be 15.
- Now I ask: 30 divided by what gives 15? That is 30 divided by 2. So the missing number is 2.
- The trick is always the same: find the side you can, then make the other side match.

DO:
- Cover the missing number and work out the left side first.
- Point to the equals sign and say 'both sides, same value'.
- Reveal the missing number once students predict it.

TEACHER NOTES:
This is relational thinking - reading the equals sign as balance, not as 'write the answer'. It is the heart of VC2M5A02. Keep modelling 'work the side you can, then match'.

MISCONCEPTIONS:
- Misconception: students think the equals sign means 'the answer comes next'.
  Why: that is how it is often used in early arithmetic.
  Impact: they try to write 15 in the box instead of finding what balances.
  Quick correction: cover both sides and say the sentence is a balance - both sides must weigh the same.

WATCH FOR:
- Students who balance both sides - secure.
- Students who put the total in the box - reteach the equals sign as balance.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards. Use the split to work out 6 times 12.
- Break the 12 into 10 and 2. Show both parts and the total.

DO:
- Display the prompt.
- Give 60 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 6 x 10 = 60, 6 x 2 = 12, total 72.
PROCEED: If about 80 percent reach 72 with both parts shown, click to reveal and move to We Do.
PIVOT: Most likely misconception - students do 6 x 10 = 60 then add the 2 to get 62.
- Reteach: draw the split array 6 by 12. The second block is 6 rows of 2, so it is 12, not 2.
- Re-check: what is the second block worth, 2 or 12?

TEACHER NOTES:
The trap is multiplying only the first part. A student who writes 62 has not multiplied the second block by 6.

WATCH FOR:
- Students who reach 72 - secure.
- Students who reach 62 - reteach the second block with the array.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Balance this sentence: 4 times 6 equals a missing number times 3.
- Work out the side you can first. Then ask what makes the other side equal.
- Whisper your missing number to your partner.

DO:
- Display 4 x 6 = ? x 3.
- Give 60 seconds.
- Listen for 'both sides must be 24'.

TEACHER NOTES:
Same balancing move as the I Do with new numbers. Listen for relational language, not guessing.

WATCH FOR:
- Pairs who find both sides equal 24, so the box is 8 - secure.
- Pairs who write 24 in the box - reteach the equals sign as balance.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- 4 times 6 is 24, so the right side must also be 24.
- A missing number times 3 equals 24 means 24 divided by 3, which is 8.
- So the missing number is 8, and 8 times 3 is 24. Both sides balance.

DO:
- Click to reveal.
- Run the balance once more: 24 on the left, 24 on the right.

TEACHER NOTES:
Reveal restates the balance. Point out we can check instantly: 8 times 3 is 24, which matches 4 times 6.

WATCH FOR:
- Students who self-correct - secure.
- Students who still guess - steer them to work each side.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, and this one is a warning about division.
- Is 8 divided by 2 divided by 2 the same as 8 divided by 2 in brackets divided by 2 the other way? True or false?
- Work out both on your whiteboard and decide.

DO:
- Display the two versions.
- Give 75 seconds.
- Watch for students who assume division regroups like multiplication.

TEACHER NOTES:
This shows division is NOT associative. Left to right: 8 divided by 2 is 4, divided by 2 is 2. But 8 divided by the bracket 2 divided by 2 is 8 divided by 1, which is 8. Different answers, so it is false.

WATCH FOR:
- Students who get 2 and 8 and say false - secure.
- Students who assume both give 2 - show the bracket changes the meaning.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Left to right: 8 divided by 2 is 4, then divided by 2 is 2.
- With the bracket: 2 divided by 2 is 1, then 8 divided by 1 is 8.
- 2 does not equal 8, so it is false. Division does not regroup the way multiplication does.

DO:
- Click to reveal.
- Underline that the grouping changed the answer for division.

TEACHER NOTES:
The big takeaway: multiplication is friendly - you can swap and regroup. Division is fussy - the order and grouping matter.

WATCH FOR:
- Students who explain why it is false - ready for independent work.
- Students who still expect division to regroup - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 warms up with splitting to multiply.
- Section 2 balances number sentences.
- Section 3 asks you to decide true or false about division properties. If you finish, try the extension box.

DO:
- Distribute the practice sheet.
- Circulate and listen for 'split, multiply each part, add' and 'both sides equal'.
- Cold call one or two students to explain a balanced sentence.

TEACHER NOTES:
Different numbers from the We Do, same moves: split to multiply, balance both sides, and remember division is fussy.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 has the split array drawn and the first part filled in. Multiply each block, then add.
- Extra Notes: Sit with these students and point to each coloured block. This is also the rebuild for any returning student.
EXTENDING PROMPT:
- Task: The extension box asks students to split a three-digit multiplication (e.g. 6 x 24) and to write their own balanced number sentence with a missing number.
- Extra Notes: Students who are ready can move on to the Year 8 extension sheet on the distributive property with brackets.

WATCH FOR:
- Students who split and balance fluently - secure.
- Students who multiply only one part - prompt them back to the array blocks.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- Use the split to work out 7 times 14.
- Show the two parts and the total.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - split a multiplication into easier parts. Look for 7 x 10 = 70, 7 x 4 = 28, total 98.

WATCH FOR:
- Students who show 70, 28 and 98 - secure.
- Students who write 7 x 10 = 70 then add 4 to get 74 - revisit the second block at the start of Lesson 3.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: why is splitting the 13 into 10 and 3 a smart way to do 4 times 13?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that multiplication can be split and rearranged to make it easier, while division cannot. Students who can split confidently are ready to multiply decimals next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on the core criterion - small group revision at the start of Lesson 3.

[General: Closing | Element: Retention and Recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 2: Smart calculating with the properties",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - fractions & decimals
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions & decimals",
      [
        "Write 1/4 as a decimal.",
        "What is 0.3 + 0.4?",
        "Which is larger: 3/10 or 0.25?",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "0.25          0.7          3/10 (= 0.3)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal x 1-digit
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal x single digit",
      ["5.3 x 4", "6.2 x 3", "4.8 x 5"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "21.2        18.6        24.0", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - order can swap (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "Turn the array, same answer",
    [
      "3 rows of 4 = 12.",
      "Turn it on its side: 4 rows of 3 = 12.",
      "",
      "The order swapped. The answer did not.",
      "Multiplication has helpful properties.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.85, { strip: C.ACCENT });
      slide.addText("3 x 4 = 4 x 3 = 12", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      drawArray(slide, lg.rightX + 0.55, lg.panelTopPadded + 0.55, 3, 4, 0.34);
      drawArray(slide, lg.rightX + 2.55, lg.panelTopPadded + 0.55, 4, 3, 0.34, { fill: C.PRIMARY });
      slide.addText("3 rows of 4          4 rows of 3", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.28, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to use the properties of multiplication to calculate smartly and to keep number sentences balanced.",
    [
      "I can swap the order of a multiplication and still get the same answer.",
      "I can split a multiplication into easier parts to work it out.",
      "I can find a missing number that keeps both sides equal, and explain why division is different.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Three properties",
    [
      "Commutative = order can swap",
      "Associative = grouping can change",
      "Distributive = split a factor",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.85, { strip: C.SECONDARY });
      const rows = [
        ["Swap", "3 x 4 = 4 x 3", C.PRIMARY],
        ["Regroup", "(2 x 2) x 3 = 2 x (2 x 3)", C.SECONDARY],
        ["Split", "4 x 13 = 4 x 10 + 4 x 3", C.ACCENT],
      ];
      const ry0 = lg.panelTopPadded + 0.18;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.84;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.25, y: ry, w: 1.25, h: 0.56, rectRadius: 0.07,
          fill: { color: r[2] },
        }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(r[1], {
          x: lg.rightX + 1.6, y: ry, w: lg.rightW - 1.8, h: 0.56,
          fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
          valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10: I Do #1 - distributive split 4 x 13
  workedExSlide(pres, 2, "I Do", "Split to multiply: 4 x 13",
    [
      "4 x 13 is awkward in one step.",
      "Split the 13 into 10 and 3.",
      "4 x 10 = 40.",
      "4 x 3 = 12.",
      "40 + 12 = 52.",
      "",
      "So 4 x 13 = 52.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.PRIMARY });
      slide.addText("4 rows of 13, split into 10 and 3", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 13.5, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      drawSplitArray(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.55, 4, 13, 10, 0.205,
        { colorA: C.PRIMARY, colorB: C.ACCENT });
      addTextOnShape(slide, "4 x 10 = 40", {
        x: lg.rightX + 0.30, y: lg.panelTopPadded + 1.95, w: 1.85, h: 0.50, rectRadius: 0.07,
        fill: { color: C.PRIMARY },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "4 x 3 = 12", {
        x: lg.rightX + 2.25, y: lg.panelTopPadded + 1.95, w: 1.6, h: 0.50, rectRadius: 0.07,
        fill: { color: C.ACCENT },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "40 + 12 = 52", {
        x: lg.rightX + 0.85, y: lg.panelTopPadded + 2.55, w: 2.5, h: 0.52, rectRadius: 0.07,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 11: I Do #2 - balance the sentence 3 x 5 = 30 ÷ ?
  workedExSlide(pres, 2, "I Do", "Keep both sides equal",
    [
      "3 x 5 = 30 ÷ ?",
      "The equals sign means both sides match.",
      "Work the side you can: 3 x 5 = 15.",
      "So the right side must be 15 too.",
      "30 ÷ ? = 15, so ? = 2.",
      "",
      "Find the side you can, then match.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Both sides, same value", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addBalance(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.62, lg.rightW - 0.5,
        "3 x 5", "30 ÷ ?", C.PRIMARY, C.ALERT);
      slide.addText("Left side = 15", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.45, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", margin: 0,
      });
      slide.addText("So 30 ÷ ? must be 15", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.80, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", margin: 0,
      });
      addTextOnShape(slide, "? = 2   (30 ÷ 2 = 15)", {
        x: lg.rightX + 0.65, y: lg.panelTopPadded + 2.25, w: lg.rightW - 1.3, h: 0.52, rectRadius: 0.07,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: CFU + reveal - split 6 x 12
  withReveal(
    () => cfuSlide(pres, "CFU", "Split to work out 6 x 12",
      { technique: "Show Me Boards",
        question: "On your whiteboard: split the 12 into 10 and 2.\n\nShow 6 x 10, 6 x 2, and the total." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "6 x 10 = 60,  6 x 2 = 12,  60 + 12 = 72", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - balance 4 x 6 = ? x 3
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Balance it: 4 x 6 = ? x 3",
      [
        "With your partner.",
        "",
        "1.  Work the side you can.",
        "2.  4 x 6 = ?",
        "3.  What times 3 gives the same?",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
        slide.addText("Both sides, same value", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.16, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addBalance(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.62, lg.rightW - 0.5,
          "4 x 6", "? x 3", C.SECONDARY, C.ALERT);
        slide.addText("Make both sides equal.", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.48, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "4 x 6 = 24, so ? x 3 = 24, which means ? = 8.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - division not associative (true/false)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "True or false? A division warning",
      [
        "With your partner.",
        "",
        "Is 8 ÷ 2 ÷ 2 the same as 8 ÷ (2 ÷ 2)?",
        "Work out BOTH.",
        "Decide: true or false?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.SECONDARY });
        slide.addText("Work out both", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addTextOnShape(slide, "8 ÷ 2 ÷ 2 = ?", {
          x: lg.rightX + 0.45, y: lg.panelTopPadded + 0.60, w: lg.rightW - 0.9, h: 0.55, rectRadius: 0.07,
          fill: { color: C.PRIMARY },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText("left to right", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.18, w: lg.rightW - 0.4, h: 0.26,
          fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
          align: "center", margin: 0,
        });
        addTextOnShape(slide, "8 ÷ (2 ÷ 2) = ?", {
          x: lg.rightX + 0.45, y: lg.panelTopPadded + 1.50, w: lg.rightW - 0.9, h: 0.55, rectRadius: 0.07,
          fill: { color: C.ALERT },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText("group the last two first", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.08, w: lg.rightW - 0.4, h: 0.26,
          fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "8 ÷ 2 ÷ 2 = 2   but   8 ÷ (2 ÷ 2) = 8.   FALSE - division is fussy.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // Slide 18: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "split to multiply.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "balance the number sentences.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "decide true or false about division.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.4;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember", {
      x: 0.7, y: panelY + 0.13, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "Split, multiply each part, add. Keep both sides of the equals sign equal.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Reference split array so the page is anchored, not text-only.
    s.addText("Reference:  5 x 12 = 5 x 10 + 5 x 2", {
      x: 3.0, y: panelY + 1.00, w: 4.0, h: 0.26,
      fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });
    drawSplitArray(s, 3.70, panelY + 1.30, 5, 12, 10, 0.17, { colorA: C.PRIMARY, colorB: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Use the split to work out 7 x 14.",
      "Show the two parts and the total.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why is splitting 13 into 10 and 3 a smart way to do 4 x 13?",
      scItems: [
        "I can swap the order of a multiplication and still get the same answer.",
        "I can split a multiplication into easier parts to work it out.",
        "I can find a missing number that keeps both sides equal, and explain why division is different.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FourP_Lesson2_Properties_Of_Operations.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Split multiplications into easier parts and balance number sentences.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Split a factor into friendly parts: 4 x 13 = 4 x 10 + 4 x 3 = 40 + 12 = 52. The equals sign means both sides match - work the side you can, then make the other side equal. Division is fussy: order and grouping matter.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "6 x 14: split the 14 into 10 and 4. 6 x 10 = 60. 6 x 4 = 24. 60 + 24 = 84. So 6 x 14 = 84.",
      y);

    y = addSectionHeading(doc, "Section 1 - Split to multiply (started for you)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Split A: 4 x 12 = 4 x 10 + 4 x 2", y);
    y = addWriteLine(doc, "4 x 10 = 40        4 x 2 = ____        Total = ____", y);
    y = addBodyText(doc, "Split B: 5 x 13", y);
    y = addWriteLine(doc, "5 x ____ = ____        5 x ____ = ____        Total = ____", y);
    y = addWriteLine(doc, "Split C: 8 x 16 = ____", y);

    y = addSectionHeading(doc, "Section 2 - Balance the sentence (find the missing number)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2 x 6 = 24 ÷ ____", y);
    y = addWriteLine(doc, "b)  5 x 4 = ____ x 2", y);
    y = addWriteLine(doc, "c)  ____ x 3 = 6 x 4", y);

    y = addSectionHeading(doc, "Section 3 - True or false (division)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  6 ÷ 3 = 3 ÷ 6     True / False", y);
    y = addWriteLine(doc, "b)  12 ÷ 2 ÷ 3 = 12 ÷ (2 ÷ 3)     True / False", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "Split to work out 6 x 24 = ____   (split the 24 into 20 and 4)", y);
    y = addWriteLine(doc, "Write your own balanced sentence with a missing number: ____ x ____ = ____ ÷ ____", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Properties of operations | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 2 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Split to multiply", y, { color: C.PRIMARY });
    y = addBodyText(doc, "A: 4 x 2 = 8, total 48.   B: 5 x 10 = 50, 5 x 3 = 15, total 65.   C: 8 x 16 = 8 x 10 + 8 x 6 = 80 + 48 = 128.", y);

    y = addSectionHeading(doc, "Section 2 - Balance the sentence", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2   (24 ÷ 2 = 12 = 2 x 6).        b)  10   (5 x 4 = 20 = 10 x 2).        c)  8   (8 x 3 = 24 = 6 x 4).", y);

    y = addSectionHeading(doc, "Section 3 - True or false", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  False. 6 ÷ 3 = 2 but 3 ÷ 6 = 0.5; division is not commutative.", y);
    y = addBodyText(doc, "b)  False. 12 ÷ 2 ÷ 3 = 2 but 12 ÷ (2 ÷ 3) = 18; division is not associative.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "6 x 24 = 6 x 20 + 6 x 4 = 120 + 24 = 144. Balanced sentences vary; check both sides are equal.", y);

    y = addTipBox(doc,
      "Watch for: students who multiply only the first part of a split (e.g. 6 x 12 = 62); students who write the total in the box instead of balancing; students who assume division regroups like multiplication.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "The distributive property with brackets.",
      color: C.SECONDARY,
      lessonInfo: `Lesson ${SESSION} | Year 8 challenge | extends Year 6 VC2M5A02`,
    });
    y = addTipBox(doc,
      "The same split you used for 4 x 13 is how algebra expands brackets. a(b + c) = ab + ac. Multiply the outside number by EVERY term inside.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Expand 4(x + 3). Multiply 4 by each term: 4 times x is 4x, 4 times 3 is 12. So 4(x + 3) = 4x + 12. It is exactly the split 4 x 13 = 4 x 10 + 4 x 3, but with a letter.",
      y);

    y = addSectionHeading(doc, "Section 1 - Expand", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3(x + 5) = ______", y);
    y = addWriteLine(doc, "b)  6(2x + 1) = ______", y);
    y = addWriteLine(doc, "c)  5(2a + 3b) = ______", y);

    y = addSectionHeading(doc, "Section 2 - Use the split to multiply mentally", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  7 x 28 = 7(20 + 8) = ______", y);
    y = addWriteLine(doc, "b)  12 x 15 = 12(10 + 5) = ______", y);
    y = addWriteLine(doc, "c)  Show how to do 99 x 6 by splitting 99 into 100 - 1: ______", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Explain why 4(x + 3) is NOT equal to 4x + 3.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) 3x + 15  b) 12x + 6  c) 10a + 15b.   S2  a) 196  b) 180  c) 99 x 6 = (100 - 1) x 6 = 600 - 6 = 594.   S3  The 4 multiplies the whole bracket, so it must multiply the 3 as well: 4 times 3 is 12, not 3.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Year 8 Extension | Distributive property`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
