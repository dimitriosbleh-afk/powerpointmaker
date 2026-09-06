"use strict";

// Mastering Fractions, Decimals and Percentages (Year 6 Numeracy) - Lesson 1 of 4.
// Find a familiar fraction OF a quantity. Unit anchor: "Find one part (divide
// by the denominator), then scale (multiply by the numerator)." From 1/3 of a
// quantity, build 2/3 and 4/3. Contexts: money, length, duration. VC2M6N07.
// Daily Review: understanding fractions (prior). Fluency: decimal column
// addition (unit-wide fluency focus).
// Unit variant fixed (variant 1, "Forest Calculation") across all 4 lessons.
// CATCH-UP: every lesson's launch re-grounds "share a quantity into equal
// parts", and worksheet Section 1 rebuilds it, so a student who missed earlier
// sessions can rejoin at any lesson. No session assumes the one before it.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(2)); // variant 1, fixed for the unit
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide, dailyReviewSlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics, getContrastColor,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 4;
const UNIT_TITLE = "Mastering Fractions, Decimals & Percentages";
const FOOTER = `Fractions, Decimals & Percentages | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FDP_Lesson1_Fractions_Of_A_Quantity";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Find a familiar fraction of a quantity by finding one part, then scaling.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 1 practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Shared bar-model helper (unit visual anchor) ---------------------------
// A whole bar split into `parts` equal segments. The first `shaded` segments
// are highlighted to show the fraction we want. Optional per-segment labels
// show the value of one part. Kept identical across the four lessons so the
// representation stays consistent for catch-up students.
function partBar(slide, x, y, w, h, parts, shaded, opts) {
  const o = opts || {};
  const segW = w / parts;
  const fill = o.fill || C.SUCCESS;
  for (let i = 0; i < parts; i++) {
    const isShaded = i < shaded;
    slide.addShape("rect", {
      x: x + i * segW, y, w: segW, h,
      fill: { color: isShaded ? fill : C.WHITE },
      line: { color: C.PRIMARY, width: 1.25 },
    });
    if (o.segLabels && o.segLabels[i] != null) {
      slide.addText(String(o.segLabels[i]), {
        x: x + i * segW, y, w: segW, h,
        fontSize: o.segFontSize || 14, fontFace: FONT_B,
        color: isShaded ? getContrastColor(fill) : C.CHARCOAL,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
    }
  }
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to our new unit on fractions, decimals and percentages.
- Over four lessons we learn to find a fraction, decimal or percentage of an amount, including shopping discounts, and to estimate sensibly.
- Today we start with the big move that runs through the whole unit: find one part first, then scale up to the part you need.

DO:
- Have whiteboards and markers ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 1 of 4. The phrase "find one part, then scale" is the anchor for the entire unit. If students leave today able to find one part and multiply, percentages later become easy.

WATCH FOR:
- Students who look unsure - that is expected on day one. Reassure them: if this feels new, that is okay, we build it together.

[General: Title | Element: Attention, focus and regulation]`;

const NOTES_OVERVIEW = `SAY:
- This slide is for you, the teacher, not the students. Click past it quickly on the day.
- It shows how the four lessons fit together and how a student who misses a session can still catch up.

DO:
- Read the four-lesson map once before teaching.
- Note the catch-up design so you can reassure returning students.

TEACHER NOTES:
Unit map: L1 fraction of a quantity, L2 connect fractions, decimals and percentages, L3 percentage of a quantity and discounts, L4 estimating to check. Each lesson re-teaches "find one part, then scale", so the lessons are deliberately low-coupling. A student who misses one or two sessions is NOT locked out: every launch rebuilds the core move from scratch, and Section 1 of every worksheet is the same re-grounding task. Progress still happens - later lessons go further - but no lesson depends on having seen the lesson before it.

WATCH FOR:
- Returning students - point them at worksheet Section 1 and one minute of your time to rejoin.

[General: Planning | Element: Planning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end. Everything else is whiteboards and partner talk.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards and markers ready.

TEACHER NOTES:
One student practice sheet and an answer key. Most teaching happens on whiteboards with a drawn bar model.

CATCH-UP NOTE:
A student who missed earlier sessions can still access today. The launch re-grounds sharing a quantity into equal parts, and Section 1 of the worksheet rebuilds it. A returner needs only the worksheet and one minute with you to start. No session in this unit assumes the student saw the one before.

[General: Resources | Element: Enabling Learning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are warming up our fraction understanding from earlier work.
- Read each one and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for how students reason about equal parts.

TEACHER NOTES:
Daily Review is prior learning, not today's "fraction of a quantity". Equivalent fractions and comparing fractions keep fraction sense ticking over.

WATCH FOR:
- Students who explain equal value clearly - secure.
- Students who compare only numerators - prompt them to think about the size of the parts.

[Stage 1: Daily Review | Element: Retention and recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- A fraction equal to one half with denominator ten is five tenths.
- Three quarters is larger than two thirds.
- One quarter plus one quarter is two quarters, which is one half.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
For comparing, a quick bar sketch settles 3/4 versus 2/3. Note any student who thinks 2/3 is larger because 3 is bigger than 4 in the denominator.

WATCH FOR:
- Students who self-correct quickly - secure.

[Stage 1: Daily Review Answers | Element: Retention and recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Adding decimals with the vertical algorithm.
- Set each one out vertically and line up the decimal points.
- Keep your columns straight so tenths sit under tenths and ones under ones.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up decimal points.

TEACHER NOTES:
Fluency for the whole unit is decimal column addition. Lining up the decimal point is the same place value habit we use with money and percentages later. Keep it brisk.

WATCH FOR:
- Students who line up the points neatly - secure.
- Students who right-align the digits instead of the points - reteach: points under points.

[Stage 1: Fluency | Element: Retention and recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 4.5 plus 2.3 is 6.8.
- 12.6 plus 3.8 is 16.4.
- 5.45 plus 1.27 is 6.72.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Watch the carry from hundredths into tenths in the third one. A misaligned point is the usual cause of a wrong total.

WATCH FOR:
- Students who self-correct - secure.
- Students whose points drift - small group focus.

[Stage 1: Fluency Answers | Element: Retention and recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember sharing things equally. If 12 stickers are shared between 3 people, how many does each person get?
- Each person gets 4. We just found one third of 12 without even saying the word fraction.
- Finding one third of an amount is the same as sharing it into 3 equal groups, which is dividing by 3.
- Today we use that one idea to find any familiar fraction of an amount.

DO:
- Draw 12 dots and circle them into 3 groups of 4 as you talk.
- Have students chorus "one third means share into 3 equal parts".
- Bridge: "find one part first, then we can scale up".

TEACHER NOTES:
This launch starts from equal sharing, which every student can access, then names it as finding a unit fraction. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who say "divide by 3" straight away - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | Element: Knowledge and memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to find a familiar fraction of a quantity by finding one part, then scaling.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Point to the words "one part" and "scale".

TEACHER NOTES:
The first criterion is reachable for everyone - find a unit fraction by dividing. The second is the core target the exit ticket checks. The third stretches to fractions beyond the whole, like 4/3.

WATCH FOR:
- Students who can repeat "divide, then multiply" - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- A quantity is just an amount, like 18 dollars, 30 metres or 60 minutes.
- The denominator is the bottom number. It tells us how many equal parts to share into.
- The numerator is the top number. It tells us how many of those parts we want.

DO:
- Point to each word as you say it.
- Have students say "divide by the denominator, multiply by the numerator" once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. Keep it to these three. Connect them straight to the action: denominator means divide, numerator means multiply.

WATCH FOR:
- Students who can name top and bottom correctly - secure.
- Students who mix them up - anchor with "Down = Divide" for the denominator.

[General: Key Vocabulary | Element: Knowledge and memory]`;

const NOTES_IDO1 = `SAY:
- Let us work through this one together. Watch how I find one third of 18 dollars.
- The denominator is 3, so I share the 18 dollars into 3 equal parts. 18 divided by 3 is 6.
- So each part is worth 6 dollars. I will write 6 dollars inside each part of my bar.
- One third is just one of those parts, so one third of 18 dollars is 6 dollars.
- Notice what I did: I found one part first. That one part is the key to everything else.

DO:
- Draw the bar, split it into 3 equal parts, write $6 in each as you speak.
- Shade one part and label it "1/3 = $6".
- Say "find one part first" as you finish.

TEACHER NOTES:
This is the core move of the lesson and the unit. Keep saying "find one part". Resist jumping to the answer - the value of one part is what makes 2/3 and 4/3 easy next.

MISCONCEPTIONS:
- Misconception: students multiply by the denominator instead of dividing.
  Why: they see the number 3 and reach for times.
  Impact: they get 54 instead of 6.
  Quick correction: the denominator shares the whole into parts, so it divides. Show the 3 equal parts on the bar.

WATCH FOR:
- Students who divide 18 by 3 confidently - secure.
- Students who multiply - re-point to the 3 equal parts on the bar.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now watch how one part does all the work.
- One part is 6 dollars. Two thirds is just two of those parts, so 2 times 6 is 12 dollars.
- Here is the powerful bit. Four thirds is four parts. Even though four thirds is more than a whole, I can still do 4 times 6, which is 24 dollars.
- So once I know one part is 6 dollars, I can scale to any number of thirds: two thirds is 12, four thirds is 24.

DO:
- Reuse the same bar, each part worth $6.
- Shade two parts for 2/3, then add an extra part beyond the whole for 4/3.
- Write "2/3 = 2 x $6 = $12" and "4/3 = 4 x $6 = $24".

TEACHER NOTES:
This is the "scale" half of the anchor. Four thirds shows the method does not break when the fraction is more than one whole - it is still parts times the value of a part.

MISCONCEPTIONS:
- Misconception: students think a fraction must be less than one, so 4/3 "cannot exist".
  Why: most early fraction work stays under a whole.
  Impact: they stall on 4/3 or change it to 3/4.
  Quick correction: 4/3 is just 4 parts, and each part is 6 dollars, so it is one whole and one more part.

WATCH FOR:
- Students who scale by multiplying - secure.
- Students who re-divide for 2/3 - remind them: one part is already found, now just multiply.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Find one quarter of 20.
- Tell your partner the one part first, then the answer.

DO:
- Display the prompt.
- Give 45 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 20 divided by 4 is 5, so one quarter of 20 is 5.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students multiply 20 by 4 and write 80.
- Reteach: draw a bar of 20 split into 4 equal parts. Each part is 5. One quarter is one part, so 5.
- Re-check: how many equal parts, and what is one part worth?

TEACHER NOTES:
The trap is multiplying by the denominator. A student who writes 80 has not shared into parts. Make them point to one part on the bar.

WATCH FOR:
- Students who write 5 - secure.
- Students who write 80 - the multiply-the-denominator error, reteach with the bar.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. We are finding two fifths of 30 metres.
- First, find one part. The denominator is 5, so share 30 into 5 equal parts.
- Whisper to your partner: what is one fifth of 30?

DO:
- Display 2/5 of 30 m above the blank bar.
- Give 75 seconds.
- Listen for "30 divided by 5 is 6".

TEACHER NOTES:
Same move as the I Do with new numbers and a length context. Hold students at "find one part" before they rush to the answer.

WATCH FOR:
- Pairs who find one fifth is 6 - secure.
- Pairs who jump straight to multiplying - prompt: one part first.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- One fifth of 30 metres is 6 metres. That is our one part.
- Two fifths is two parts, so 2 times 6 is 12 metres.
- Find one part, then scale. Same move every time.

DO:
- Click to reveal.
- Run the two steps once more together: divide, then multiply.

TEACHER NOTES:
Reveal restates the anchor. The context changed to metres but the method did not.

WATCH FOR:
- Students who self-correct - secure.
- Students who wrote 6 only - remind them two fifths needs two parts.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, this time with time.
- We want three quarters of 60 minutes.
- First find one part: share 60 minutes into 4 equal parts. Then scale to three parts.

DO:
- Display 3/4 of 60 min above the blank bar.
- Give 90 seconds.
- Watch for students who find one part but forget to scale to three.

TEACHER NOTES:
60 divided by 4 is 15, so one quarter is 15 minutes. Three quarters is 3 times 15, which is 45 minutes. A familiar real context - three quarters of an hour.

WATCH FOR:
- Students who write 45 - secure.
- Students who write 15 - they stopped at one part, prompt them to scale to three.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- One quarter of 60 minutes is 15 minutes - that is one part.
- Three quarters is three parts, so 3 times 15 is 45 minutes.
- Three quarters of an hour is 45 minutes. That matches what we know about time.

DO:
- Click to reveal.
- Connect to a clock: 45 minutes is three quarters of the way round.

TEACHER NOTES:
If many wrote 15, do one more quick "find one part then scale" before releasing to the You Do.

WATCH FOR:
- Students who scale correctly - ready for independent work.
- Students who stop at one part - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 warms up by finding one part - the unit fraction.
- Section 2 mixes fractions of money, length and time.
- Section 3 asks you to explain your thinking. If you finish, try the challenge box with a fraction beyond a whole.

DO:
- Distribute the practice sheet.
- Circulate and listen for "find one part, then scale".
- Cold call one or two students to explain a two-step answer.

TEACHER NOTES:
Different numbers and contexts from the We Do, same move: divide to find one part, multiply to scale.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 is started with a drawn bar. Find one part first, then move on. This is also the rebuild for any returning student.
- Extra Notes: Sit with these students and find the first "one part" together.
EXTENDING PROMPT:
- Task: The challenge box asks for a fraction beyond a whole, like 5/4 of 20, and to explain why the method still works.
- Extra Notes: Push them to justify with the bar - parts beyond the whole are still parts.

WATCH FOR:
- Students who use the two-step method fluently - secure.
- Students who multiply by the denominator - re-anchor with one part on the bar.

[Stage 4: You Do | Element: Mastery and application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- Find two thirds of 27 dollars.
- Show how you found one part first.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - find a familiar fraction of a quantity by finding one part then scaling. Look for 27 divided by 3 is 9, then 2 times 9 is 18 dollars. The SC target is SC2.

WATCH FOR:
- Students who show $9 as one part then $18 - secure.
- Students who write $9 only - they stopped at one part, revisit at the start of Lesson 2.

[Stage 5: Exit Ticket | Element: Mastery and application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: what are the two steps for finding a fraction of an amount?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is "find one part, then scale". Students who can say the two steps are ready to connect fractions to decimals and percentages next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on the core criterion - small group revision at the start of Lesson 2.

[General: Closing | Element: Retention and recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Finding a fraction of a quantity",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Unit map + catch-up design (teacher-facing, Session 1 only)
  contentSlide(pres, "Teacher", C.PRIMARY, "Unit map & catch-up design",
    [
      "L1: Find a fraction of a quantity",
      "L2: Connect fractions, decimals & percentages",
      "L3: Percentage of a quantity & discounts",
      "L4: Estimate to check answers",
      "",
      "Anchor every lesson: find one part, then scale.",
    ],
    NOTES_OVERVIEW, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
      slide.addText("Catch-up by design", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      slide.addText([
        { text: "Lessons are low-coupling.", options: { bullet: true, breakLine: true, bold: true } },
        { text: "Every launch rebuilds the core move.", options: { bullet: true, breakLine: true } },
        { text: "Worksheet Section 1 = the same re-grounding task.", options: { bullet: true, breakLine: true } },
        { text: "A returner needs 1 minute with you to rejoin.", options: { bullet: true, breakLine: true } },
        { text: "Progress still happens - later goes further.", options: { bullet: true, breakLine: false } },
      ], {
        x: lg.rightX + 0.25, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.45, h: 2.35,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
        paraSpaceAfter: 5,
      });
    }
  );

  // Slides 4-5: Daily Review + reveal - understanding fractions
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions",
      [
        "Write a fraction equal to 1/2 with a denominator of 10.",
        "Which is larger: 2/3 or 3/4?",
        "1/4 + 1/4 = ?",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "5/10          3/4 is larger          2/4 = 1/2", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 6-7: Fluency + reveal - decimal column addition
  withReveal(
    () => fluencySlide(pres, "Fluency: Adding decimals",
      ["4.5 + 2.3", "12.6 + 3.8", "5.45 + 1.27"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "6.8        16.4        6.72", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 8: Launch - equal sharing (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "We can already share equally",
    [
      "12 stickers shared between 3 people.",
      "How many each?",
      "",
      "One third of 12 = share into 3 = divide by 3.",
      "Today: find one part, then scale.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.ACCENT });
      slide.addText("Share 12 into 3 groups", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      // Three groups of four dots
      const gy = lg.panelTopPadded + 0.62;
      for (let g = 0; g < 3; g++) {
        const gx = lg.rightX + 0.35 + g * 1.18;
        addCard(slide, gx, gy, 1.0, 1.0, { fill: C.BG_LIGHT, strip: null });
        for (let d = 0; d < 4; d++) {
          const dx = gx + 0.18 + (d % 2) * 0.46;
          const dy = gy + 0.18 + Math.floor(d / 2) * 0.46;
          slide.addShape("roundRect", {
            x: dx, y: dy, w: 0.26, h: 0.26, rectRadius: 0.13,
            fill: { color: C.SECONDARY },
          });
        }
      }
      addTextOnShape(slide, "12 ÷ 3 = 4 in each group", {
        x: lg.rightX + 0.30, y: gy + 1.20, w: lg.rightW - 0.60, h: 0.46, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 9: LI/SC
  liSlide(pres,
    "We are learning to find a familiar fraction of a quantity by finding one part, then scaling.",
    [
      "I can find a unit fraction of an amount by dividing into equal parts.",
      "I can find a familiar fraction of an amount by finding one part, then multiplying.",
      "I can find a fraction beyond a whole, like 4/3, and explain why the method still works.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 10: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Quantity = an amount (18 dollars, 30 m, 60 min)",
      "Denominator = bottom number = how many equal parts",
      "Numerator = top number = how many parts we want",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.3, { strip: C.SECONDARY });
      slide.addText("The two-step move", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [["Denominator", "DIVIDE"], ["Numerator", "MULTIPLY"]];
      const ry0 = lg.panelTopPadded + 0.58;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.74;
        slide.addText(r[0], {
          x: lg.rightX + 0.25, y: ry, w: 1.95, h: 0.56,
          fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        addTextOnShape(slide, r[1], {
          x: lg.rightX + 2.25, y: ry, w: 1.55, h: 0.56, rectRadius: 0.07,
          fill: { color: i === 0 ? C.PRIMARY : C.SUCCESS },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slide 11: I Do #1 - 1/3 of $18
  workedExSlide(pres, 2, "I Do", "Find 1/3 of $18",
    [
      "Denominator is 3 -> DIVIDE.",
      "Share $18 into 3 equal parts.",
      "18 ÷ 3 = 6, so one part = $6.",
      "",
      "1/3 of $18 = $6.",
      "Find one part first.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.PRIMARY });
      slide.addText("Whole = $18", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      partBar(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.58, lg.rightW - 0.50, 0.78,
        3, 1, { segLabels: ["$6", "$6", "$6"], fill: C.SUCCESS, segFontSize: 16 });
      addTextOnShape(slide, "18 ÷ 3 = $6 in each part", {
        x: lg.rightX + 0.30, y: lg.panelTopPadded + 1.55, w: lg.rightW - 0.60, h: 0.46, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "1/3 of $18 = $6", {
        x: lg.rightX + 0.30, y: lg.panelTopPadded + 2.20, w: lg.rightW - 0.60, h: 0.54, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 12: I Do #2 - 2/3 and 4/3 of $18 (scale the part)
  workedExSlide(pres, 2, "I Do", "Now scale: 2/3 and 4/3 of $18",
    [
      "One part = $6 (we already found it).",
      "Numerator tells us how many parts.",
      "",
      "2/3 = 2 parts = 2 x $6 = $12.",
      "4/3 = 4 parts = 4 x $6 = $24.",
      "Parts beyond a whole still work.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.PRIMARY });
      slide.addText("Each part = $6", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      // 2/3 bar (3 parts, 2 shaded)
      partBar(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.50, lg.rightW - 0.50, 0.62,
        3, 2, { segLabels: ["$6", "$6", "$6"], fill: C.SECONDARY, segFontSize: 14 });
      const rows = [
        ["2/3 = 2 x $6 = $12", C.SECONDARY],
        ["4/3 = 4 x $6 = $24", C.SUCCESS],
      ];
      const ry0 = lg.panelTopPadded + 1.40;
      rows.forEach((r, i) => {
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.30, y: ry0 + i * 0.66, w: lg.rightW - 0.60, h: 0.52, rectRadius: 0.08,
          fill: { color: r[1] },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
      slide.addText("4/3 is one whole and one more part.", {
        x: lg.rightX + 0.15, y: ry0 + 1.36, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 11.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slides 13-14: CFU + reveal - 1/4 of 20
  withReveal(
    () => cfuSlide(pres, "CFU", "Find 1/4 of 20",
      { technique: "Show Me Boards",
        question: "On your whiteboard: 1/4 of 20 = ?\n\nTell your partner the one part first, then the answer." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1/4 of 20 = 5   (20 ÷ 4 = 5; one quarter is one part)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 15-16: We Do #1 + reveal - 2/5 of 30 m
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Find 2/5 of 30 m together",
      [
        "With your partner.",
        "",
        "1.  Find one part: 30 ÷ 5 = ?",
        "2.  Scale: 2 parts.",
        "3.  Read the answer in metres.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.SECONDARY });
        slide.addText("2/5 of 30 m", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.42,
          fontSize: 24, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        slide.addText("Whole = 30 m, share into 5 parts", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.58, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
          align: "center", margin: 0,
        });
        partBar(slide, lg.rightX + 0.20, lg.panelTopPadded + 0.94, lg.rightW - 0.40, 0.72,
          5, 0, { fill: C.SECONDARY });
        slide.addText("Find one part, then scale to 2.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.80, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "2/5 of 30 m = 12 m   (30 ÷ 5 = 6; 2 x 6 = 12)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 17-18: We Do #2 + reveal - 3/4 of 60 min
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Find 3/4 of 60 minutes",
      [
        "With your partner.",
        "",
        "1.  Find one part: 60 ÷ 4 = ?",
        "2.  Scale: 3 parts.",
        "3.  Read the answer in minutes.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.SECONDARY });
        slide.addText("3/4 of 60 min", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.42,
          fontSize: 24, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        slide.addText("Whole = 60 min, share into 4 parts", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.58, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
          align: "center", margin: 0,
        });
        partBar(slide, lg.rightX + 0.20, lg.panelTopPadded + 0.94, lg.rightW - 0.40, 0.72,
          4, 0, { fill: C.SECONDARY });
        slide.addText("Find one part, then scale to 3.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.80, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "3/4 of 60 min = 45 min   (60 ÷ 4 = 15; 3 x 15 = 45)", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // Slide 19: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "find one part (divide).   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "scale to the parts you need (multiply).   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "explain your thinking in words.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Find one part, then scale.  Denominator divides, numerator multiplies.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Blank bar template (8 parts) the teacher can reference
    slide_blankBar(s, 2.5, panelY + 1.15, 5.0, 0.55);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);

    function slide_blankBar(sl, x, y, w, h) {
      partBar(sl, x, y, w, h, 4, 0, { fill: C.SECONDARY });
      sl.addText("Whole", {
        x: x, y: y + h + 0.04, w: w, h: 0.24,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED,
        align: "center", margin: 0,
      });
    }
  })();

  // Slide 20: Exit Ticket
  exitTicketSlide(pres,
    [
      "Find 2/3 of $27.",
      "Show how you found one part first.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 21: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what are the two steps for finding a fraction of an amount?",
      scItems: [
        "I can find a unit fraction of an amount by dividing into equal parts.",
        "I can find a familiar fraction of an amount by finding one part, then multiplying.",
        "I can find a fraction beyond a whole, like 4/3, and explain why the method still works.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FDP_Lesson1_Fractions_Of_A_Quantity.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find a familiar fraction of a quantity by finding one part, then scaling.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Two steps every time. Find ONE part: divide the amount by the denominator (the bottom number). Then SCALE: multiply by the numerator (the top number).",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Find 2/5 of 20. One part: 20 divided by 5 is 4. Scale: 2 parts, so 2 x 4 = 8. So 2/5 of 20 = 8.",
      y);

    y = addSectionHeading(doc, "Section 1 - Find one part (started for you)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/3 of 12:  12 ÷ 3 = 4   (one third of 12 is 4)", y);
    y = addWriteLine(doc, "b)  1/4 of 24:  24 ÷ 4 = _______", y);
    y = addWriteLine(doc, "c)  1/5 of 35:  35 ÷ 5 = _______", y);
    y = addWriteLine(doc, "d)  1/6 of 30:  30 ÷ 6 = _______", y);

    y = addSectionHeading(doc, "Section 2 - Find one part, then scale", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2/3 of $18 = _______", y);
    y = addWriteLine(doc, "b)  3/4 of 40 m = _______", y);
    y = addWriteLine(doc, "c)  2/5 of 45 min = _______", y);
    y = addWriteLine(doc, "d)  3/5 of $25 = _______", y);
    y = addWriteLine(doc, "e)  5/6 of 60 kg = _______", y);

    y = addSectionHeading(doc, "Section 3 - Explain your thinking", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "2/3 of $30 = _______.  Explain the one part and the scaling step:", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addSectionHeading(doc, "Challenge (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "5/4 of 20 = _______.  This fraction is more than a whole. Explain why the method still works.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Fraction of a Quantity | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 1 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Find one part", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4        b)  6        c)  7        d)  5", y);

    y = addSectionHeading(doc, "Section 2 - Find one part, then scale", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  $12  (18 ÷ 3 = 6; 2 x 6)        b)  30 m  (40 ÷ 4 = 10; 3 x 10)        c)  18 min  (45 ÷ 5 = 9; 2 x 9)", y);
    y = addBodyText(doc, "d)  $15  (25 ÷ 5 = 5; 3 x 5)        e)  50 kg  (60 ÷ 6 = 10; 5 x 10)", y);

    y = addSectionHeading(doc, "Section 3 - Explain your thinking", y, { color: C.PRIMARY });
    y = addBodyText(doc, "2/3 of $30 = $20. One part: 30 ÷ 3 = 10. Scale: 2 x 10 = 20.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "5/4 of 20 = 25. One part: 20 ÷ 4 = 5. Scale: 5 x 5 = 25. The method works because 5/4 is just 5 equal parts, even though that is more than one whole (one whole is 4 parts = 20, plus one more part = 5).", y);

    y = addTipBox(doc,
      "Watch for: students who multiply by the denominator (e.g. 1/4 of 20 = 80); students who stop after finding one part and forget to scale; students who think a fraction cannot be more than a whole.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
