"use strict";

// Four Processes (Year 6 Numeracy) - Lesson 4 of 5.
// Fractions and percentages of a quantity. The unifying idea: divide to make
// ONE part, then multiply to scale. 1/3 of an amount by dividing by 3, then 2/3
// and 4/3 by scaling; percentages via 10% (e.g. 30% = 10% x 3); 25% = 1/4.
// VC2M5N07, VC2M6N09.
// Daily Review: Fractions & Decimals (prior). Fluency: decimal x 1-digit algorithm.
// Unit variant fixed (variant 3, Ocean Logic) across all 5 lessons for cohesion.
// Catch-up: the launch rebuilds the idea from sharing 12 lollies everyone can do,
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

const SESSION = 4;
const TOTAL = 5;
const UNIT_TITLE = "Four Processes";
const FOOTER = `Four Processes | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FourP_Lesson4_Fractions_Percentages_Of_Quantity";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Find fractions and percentages of amounts by finding one part and scaling.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 4 practice sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Percentage change with multipliers and reverse percentages - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Helpers -----------------------------------------------------------------
// A partitioned bar: `parts` equal segments, first `shaded` filled.
function drawBarModel(slide, x, y, w, h, parts, shaded, opts) {
  const o = opts || {};
  const fill = o.fill || C.SECONDARY;
  const empty = o.empty || C.WHITE;
  const seg = w / parts;
  for (let i = 0; i < parts; i++) {
    slide.addShape("rect", {
      x: x + i * seg, y, w: seg, h,
      fill: { color: i < shaded ? fill : empty },
      line: { color: C.PRIMARY, width: 1 },
    });
    if (o.segLabel) {
      slide.addText(String(o.segLabel), {
        x: x + i * seg, y, w: seg, h,
        fontSize: o.segFont || 12, fontFace: FONT_B,
        color: i < shaded ? C.WHITE : C.MUTED, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  }
  return { seg };
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to Four Processes. We have linked the operations, used the properties, and estimated with decimals.
- Today we find fractions and percentages of amounts - the maths behind sale prices and sharing.
- There is one big idea that does both jobs: divide to make one part, then multiply to scale.

DO:
- Have whiteboards, markers and counters ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 4 of 5. The hero idea is 'one part, then scale'. It unifies fractions of a quantity and percentages, including discounts.

WATCH FOR:
- Students who look unsure - that is expected. Reassure them: if this feels new, that is okay, we build it together.

[General: Title | Element: Planning and Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end. The Year 8 extension is there for anyone who is flying.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards, markers and counters ready.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student practice sheet, an answer key, and a Year 8 extension. Most teaching happens on whiteboards with bar models.

CATCH-UP NOTE:
A student who missed earlier sessions can still access today. The launch rebuilds the idea from sharing 12 lollies into 3 groups, and Section 1 of the worksheet rebuilds finding one part. A returner only needs counters and one minute with you to start. No session in this unit assumes the student saw the one before.

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
Daily Review is prior learning, not today's new content. Converting between fractions and decimals sets students up to link them to percentages today.

WATCH FOR:
- Students who know 3/4 is 0.75 - secure.
- Students who cannot order mixed forms - prompt them to convert to decimals first.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 3/4 is 0.75.
- 0.5 is the same as 1/2.
- Smallest to largest: 0.4, then 1/2, then 3/4.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Converting to decimals makes the ordering easy: 0.4, 0.5, 0.75. This links straight into percentages today.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students who order without converting - reteach converting first.

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
Fluency this unit is the multiplication algorithm with decimals. Today's scaling work uses multiplication, so this warm-up feeds straight in.

WATCH FOR:
- Students who keep the decimal point lined up - secure.
- Students who lose the point - prompt: one decimal place in, one decimal place out.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 7.2 times 4 is 28.8.
- 3.9 times 6 is 23.4.
- 5.4 times 5 is 27.0, which we write as 27.
- Tick and fix.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
5.4 times 5 lands on 27.0; accept 27. Keep it brisk.

WATCH FOR:
- Students who self-correct - secure.
- Students whose point drifts - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember sharing things equally. Imagine 12 lollies shared between 3 friends.
- Each friend gets 12 divided by 3, which is 4. That is one third of the lollies.
- Now, what if I want two thirds? I just take two of those groups: 4 and 4 is 8.
- That is the whole idea for today. Find one part by dividing, then take as many parts as you need.

DO:
- Show 12 counters, share into 3 groups of 4.
- Have students chorus 'one third is 4, two thirds is 8'.
- Bridge: 'find one part, then scale - it works for fractions and percentages'.

TEACHER NOTES:
This launch starts from equal sharing everyone can access, then names today's strategy. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who find two thirds by doubling one third - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to find a fraction or a percentage of an amount by finding one part and then scaling.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Hold up a bar model drawn on the board.

TEACHER NOTES:
The first criterion is reachable for everyone - find one part by dividing. The second is the core target the exit ticket checks - find a fraction or percentage by scaling. The third stretches to a full discount problem with reasoning.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- Percent means out of 100. 50% is 0.5, 25% is 0.25, 10% is 0.1.
- A fraction of an amount means split into equal parts and take some.
- A discount is the amount taken off a price in a sale.

DO:
- Point to each word and its meaning.
- Have students say '10% is one tenth, divide by 10' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. Connect 50% to a half, 25% to a quarter, 10% to a tenth so percentages feel like familiar fractions.

WATCH FOR:
- Students who link 25% with a quarter - secure.
- Students who think 25% means 25 dollars - clarify percent is out of 100.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at this one together. I want one third of 24 dollars.
- One third means I share into 3 equal parts: 24 divided by 3 is 8. So one third of 24 is 8.
- Look at the bar - it is split into 3 equal parts of 8 dollars each.
- Now two thirds is just two of those parts: 8 times 2 is 16. And four thirds is four parts: 8 times 4 is 32, which is more than the whole.
- Find one part by dividing, then multiply to take as many as you need.

DO:
- Point to each part of the bar as you name it.
- Write 24 div 3 = 8, then 8 x 2 = 16 and 8 x 4 = 32.
- Have students chorus 'one part, then scale'.

TEACHER NOTES:
This is the core move of the lesson. The bar makes the parts visible so students see that two thirds is two of the three parts, not a separate calculation.

MISCONCEPTIONS:
- Misconception: students try to divide 24 by 2 to get two thirds.
  Why: they see the 2 in two thirds and divide by it.
  Impact: they get 12 instead of 16.
  Quick correction: the bottom number says how many parts to cut, the top says how many to take. Cut into 3, take 2.

WATCH FOR:
- Students who find one part then scale - secure.
- Students who divide by the top number - re-point to the bar parts.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now percentages, the same way. A 40 dollar toy is 30% off. How much comes off?
- 30% is hard in one go, but 10% is easy: 10% is one tenth, so 40 divided by 10 is 4.
- Now scale. 30% is three lots of 10%, so 4 times 3 is 12. The discount is 12 dollars.
- So I pay 40 take away 12, which is 28 dollars.
- Find 10% first - that is the one part. Then multiply to reach the percentage you need.

DO:
- Shade three of the ten strips and name each as 10%.
- Write 10% = 4, then 30% = 12, then 40 - 12 = 28.
- Have students chorus 'find 10%, then scale'.

TEACHER NOTES:
This is the elaboration in the curriculum: find 30% by using 10% and multiplying by 3. The strip shows each tenth is 10%, so three tenths is 30%.

MISCONCEPTIONS:
- Misconception: students give the discount as the final answer and forget to subtract.
  Why: they work out 12 and stop.
  Impact: they say the toy costs 12 dollars instead of 28.
  Quick correction: the 12 is what comes OFF; the price is 40 minus 12.

WATCH FOR:
- Students who find 10% then scale, then subtract - secure.
- Students who stop at the discount - prompt: is that the saving or the new price?

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards. Find 20% of 50 dollars.
- Find 10% first, then scale.

DO:
- Display the prompt.
- Give 45 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 10% of 50 is 5, then 5 x 2 = 10. So 20% is 10 dollars.
PROCEED: If about 80 percent reach 10 dollars, click to reveal and move to We Do.
PIVOT: Most likely misconception - students divide 50 by 20.
- Reteach: 20% is two lots of 10%. Find 10% first: 50 divided by 10 is 5. Then double it.
- Re-check: what is 10% of 50, and how many tens are in 20%?

TEACHER NOTES:
The trap is dividing by the percentage number. Always find 10% first, then scale.

WATCH FOR:
- Students who reach 10 dollars - secure.
- Students who divide 50 by 20 - reteach finding 10% first.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Find three quarters of 20 dollars.
- Find one quarter first by dividing. Then scale to three quarters.
- Whisper your one part to your partner.

DO:
- Display 3/4 of $20 with a bar split into 4 parts.
- Give 60 seconds.
- Listen for 'one quarter is 5'.

TEACHER NOTES:
Same one-part-then-scale move as the I Do with a different fraction. One quarter of 20 is 5, three quarters is 15.

WATCH FOR:
- Pairs who find 5 then scale to 15 - secure.
- Pairs who divide by 3 - reteach: quarters means cut into 4.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- One quarter of 20 is 20 divided by 4, which is 5.
- Three quarters is three of those parts: 5 times 3 is 15.
- So three quarters of 20 dollars is 15 dollars.

DO:
- Click to reveal.
- Point to the three shaded parts of the bar.

TEACHER NOTES:
Reveal restates the move. Note that 1/4 is the same as 25%, linking fractions and percentages.

WATCH FOR:
- Students who self-correct - secure.
- Students who took only one part - remind them to scale to three.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, a real sale. A 60 dollar jacket has 25% off.
- Remember 25% is the same as one quarter. Find one quarter, then work out the sale price.
- Build it on your whiteboard.

DO:
- Display the jacket problem.
- Give 75 seconds.
- Watch for students who forget to subtract the discount.

TEACHER NOTES:
This links 25% with a quarter and asks for the sale price, not just the discount. One quarter of 60 is 15 off, so the price is 45 dollars.

WATCH FOR:
- Students who find 15 off and pay 45 - secure.
- Students who say the jacket costs 15 - prompt: is that the saving or the new price?

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- 25% is one quarter. One quarter of 60 is 60 divided by 4, which is 15.
- The discount is 15 dollars, so the sale price is 60 take away 15, which is 45 dollars.
- Find one part, scale if needed, then subtract for a sale price.

DO:
- Click to reveal.
- Separate the saving (15) from the new price (45).

TEACHER NOTES:
The big takeaway: a discount question has two answers in it - how much comes off, and how much you pay. Read which one is asked.

WATCH FOR:
- Students who give both the saving and the price - ready for independent work.
- Students who confuse the two - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 warms up by finding one part.
- Section 2 finds fractions and percentages of amounts.
- Section 3 works out sale prices. If you finish, try the extension box.

DO:
- Distribute the practice sheet.
- Circulate and listen for 'one part, then scale'.
- Cold call one or two students to explain a sale price.

TEACHER NOTES:
Different numbers from the We Do, same move: divide to make one part, multiply to scale, subtract for a sale price.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 has the bar drawn and the first part filled in. Divide to find one part, then count the shaded parts.
- Extra Notes: Sit with these students and find the one part together. This is also the rebuild for any returning student.
EXTENDING PROMPT:
- Task: The extension box asks students to compare two sale offers (e.g. 25% off vs 20 dollars off a 70 dollar item) and to justify the better deal.
- Extra Notes: Students who are ready can move on to the Year 8 extension sheet on percentage change with multipliers.

WATCH FOR:
- Students who find one part then scale - secure.
- Students who divide by the top number of a fraction - prompt them back to the bar.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- A 30 dollar game has 40% off. How much do you save?
- Find 10% first, then scale.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - find a percentage of an amount by finding one part and scaling. Look for 10% of 30 is 3, times 4 is 12. Save 12 dollars.

WATCH FOR:
- Students who find 10% then scale to 40% - secure.
- Students who divide 30 by 40 - revisit finding 10% first at the start of Lesson 5.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: how is finding 30% of an amount like finding two thirds of an amount?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is 'find one part, then scale', which works for both fractions and percentages. Students who can do this are ready to choose operations and budget next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on the core criterion - small group revision at the start of Lesson 5.

[General: Closing | Element: Retention and Recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 4: Fractions and percentages of an amount",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - fractions & decimals
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Fractions & decimals",
      [
        "Write 3/4 as a decimal.",
        "Write 0.5 as a fraction.",
        "Order from smallest: 1/2, 0.4, 3/4",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "0.75          1/2          0.4, 1/2, 3/4", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal x 1-digit
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal x single digit",
      ["7.2 x 4", "3.9 x 6", "5.4 x 5"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "28.8        23.4        27.0", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - sharing lollies (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "One part, then scale",
    [
      "12 lollies shared between 3 friends.",
      "Each gets 12 ÷ 3 = 4. That is one third.",
      "",
      "Two thirds? Take two groups: 4 + 4 = 8.",
      "Find one part, then take what you need.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.ACCENT });
      slide.addText("12 shared into 3", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      drawBarModel(slide, lg.rightX + 0.35, lg.panelTopPadded + 0.55, lg.rightW - 0.70, 0.62, 3, 2,
        { segLabel: "4", segFont: 18 });
      slide.addText("1/3 = 4        2/3 = 8", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.40, w: lg.rightW - 0.30, h: 0.32,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      slide.addText("Each part is one third.", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.82, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to find a fraction or a percentage of an amount by finding one part and then scaling.",
    [
      "I can find one part of an amount by dividing.",
      "I can find a fraction or percentage of an amount by finding one part and scaling.",
      "I can work out a sale price after a discount and explain my steps.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Percent = out of 100 (50% = 0.5, 25% = 0.25)",
      "Fraction of an amount = split, then take some",
      "Discount = the amount taken off the price",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.SECONDARY });
      slide.addText("Same value, three ways", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.32,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [
        ["50%", "= 1/2 = 0.5", C.PRIMARY],
        ["25%", "= 1/4 = 0.25", C.SECONDARY],
        ["10%", "= 1/10 = 0.1", C.ACCENT],
      ];
      const ry0 = lg.panelTopPadded + 0.60;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.60;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.40, y: ry, w: 1.2, h: 0.48, rectRadius: 0.06,
          fill: { color: r[2] },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(r[1], {
          x: lg.rightX + 1.75, y: ry, w: lg.rightW - 2.0, h: 0.48,
          fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
          valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10: I Do #1 - fraction of a quantity (1/3, 2/3, 4/3 of 24)
  workedExSlide(pres, 2, "I Do", "Fraction of an amount: thirds of $24",
    [
      "1/3 of $24: share into 3 parts.",
      "24 ÷ 3 = 8. So 1/3 = $8.",
      "2/3 = two parts: 8 x 2 = $16.",
      "4/3 = four parts: 8 x 4 = $32.",
      "",
      "One part, then scale.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.05, { strip: C.PRIMARY });
      slide.addText("$24 split into 3 equal parts", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      drawBarModel(slide, lg.rightX + 0.35, lg.panelTopPadded + 0.55, lg.rightW - 0.70, 0.62, 3, 2,
        { segLabel: "$8", segFont: 16 });
      slide.addText("2 parts shaded = 2/3", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.26, w: lg.rightW - 0.30, h: 0.26,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      const chips = [["1/3 = $8", C.PRIMARY], ["2/3 = $16", C.SECONDARY], ["4/3 = $32", C.ACCENT]];
      const cy = lg.panelTopPadded + 1.62;
      chips.forEach((ch, i) => {
        addTextOnShape(slide, ch[0], {
          x: lg.rightX + 0.20 + i * 1.27, y: cy, w: 1.18, h: 0.52, rectRadius: 0.06,
          fill: { color: ch[1] },
        }, { fontSize: 13.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
      slide.addText("4/3 is more than the whole.", {
        x: lg.rightX + 0.15, y: cy + 0.62, w: lg.rightW - 0.30, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 11: I Do #2 - percentage discount via 10% (30% off $40)
  workedExSlide(pres, 2, "I Do", "Percentage off: 30% off $40",
    [
      "30% is tricky, but 10% is easy.",
      "10% of $40: 40 ÷ 10 = $4.",
      "30% = three lots of 10%: 4 x 3 = $12.",
      "Discount = $12.",
      "Pay: 40 - 12 = $28.",
      "",
      "Find 10%, then scale.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.05, { strip: C.PRIMARY });
      slide.addText("$40 in ten 10% parts", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      drawBarModel(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.55, lg.rightW - 0.60, 0.55, 10, 3,
        { fill: C.ALERT });
      slide.addText("3 parts shaded = 30%   (each part = $4)", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.18, w: lg.rightW - 0.30, h: 0.26,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      const chips = [["10% = $4", C.PRIMARY], ["30% = $12", C.ALERT]];
      const cy = lg.panelTopPadded + 1.54;
      chips.forEach((ch, i) => {
        addTextOnShape(slide, ch[0], {
          x: lg.rightX + 0.40 + i * 1.75, y: cy, w: 1.6, h: 0.52, rectRadius: 0.06,
          fill: { color: ch[1] },
        }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
      addTextOnShape(slide, "Pay $28", {
        x: lg.rightX + 1.15, y: cy + 0.62, w: 1.7, h: 0.50, rectRadius: 0.07,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: CFU + reveal - 20% of $50
  withReveal(
    () => cfuSlide(pres, "CFU", "Find 20% of $50",
      { technique: "Show Me Boards",
        question: "On your whiteboard: find 10% of $50 first, then scale to 20%." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "10% of $50 = $5,  so 20% = 5 x 2 = $10", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 3/4 of $20
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Find 3/4 of $20",
      [
        "With your partner.",
        "",
        "1.  Find one quarter: 20 ÷ 4.",
        "2.  Scale to three quarters.",
        "3.  Read the answer.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.45, { strip: C.SECONDARY });
        slide.addText("$20 split into 4 parts", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        drawBarModel(slide, lg.rightX + 0.35, lg.panelTopPadded + 0.60, lg.rightW - 0.70, 0.62, 4, 3,
          { segLabel: "?", segFont: 16 });
        slide.addText("One quarter first, then scale.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.40, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "1/4 of $20 = $5, so 3/4 = 5 x 3 = $15.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - 25% off a $60 jacket
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "A real sale: 25% off a $60 jacket",
      [
        "With your partner.",
        "",
        "25% is the same as one quarter.",
        "Find one quarter of $60.",
        "What is the sale price?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.55, { strip: C.SECONDARY });
        slide.addText("25% = one quarter", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        drawBarModel(slide, lg.rightX + 0.35, lg.panelTopPadded + 0.60, lg.rightW - 0.70, 0.58, 4, 1,
          { fill: C.ALERT, segLabel: "?", segFont: 15 });
        slide.addText("Shaded part = the 25% discount", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.34, w: lg.rightW - 0.30, h: 0.26,
          fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
          align: "center", margin: 0,
        });
        slide.addText("Discount first, then sale price.", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.70, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "1/4 of $60 = $15 off, so the sale price is 60 - 15 = $45.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "find one part by dividing.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "find fractions and percentages.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "work out the sale prices.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Divide to make one part. Multiply to scale. Subtract for a sale price.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Reference bar so the page is anchored, not text-only.
    s.addText("Reference:  10% of $50 = $5, so 30% = $15", {
      x: 2.5, y: panelY + 1.02, w: 5.0, h: 0.28,
      fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });
    drawBarModel(s, 3.0, panelY + 1.34, 4.0, 0.48, 10, 3, { fill: C.ALERT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "A $30 game has 40% off. How much do you save?",
      "Find 10% first, then scale to 40%.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how is finding 30% of an amount like finding two thirds of an amount?",
      scItems: [
        "I can find one part of an amount by dividing.",
        "I can find a fraction or percentage of an amount by finding one part and scaling.",
        "I can work out a sale price after a discount and explain my steps.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FourP_Lesson4_Fractions_Percentages_Of_Quantity.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find fractions and percentages of amounts by finding one part and scaling.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Find one part by dividing, then multiply to scale. For a fraction: divide by the bottom number, multiply by the top. For percentages: find 10% (divide by 10), then scale - or use 50% = a half, 25% = a quarter. For a sale price, subtract the discount.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "2/3 of $18: one third is 18 div 3 = 6. Two thirds is 6 x 2 = 12. So 2/3 of $18 is $12.",
      y);

    y = addSectionHeading(doc, "Section 1 - Find one part (started for you)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  1/4 of 20 = 20 ÷ 4 = 5        b)  1/3 of 27 = _______        c)  10% of 80 = _______", y);

    y = addSectionHeading(doc, "Section 2 - Find the fraction or percentage", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3/4 of 24 = _______", y);
    y = addWriteLine(doc, "b)  2/3 of 30 = _______", y);
    y = addWriteLine(doc, "c)  20% of 45 = _______", y);
    y = addWriteLine(doc, "d)  30% of 60 = _______", y);

    y = addSectionHeading(doc, "Section 3 - Sale prices", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  A $40 toy has 25% off. Discount: ______   Sale price: ______", y);
    y = addWriteLine(doc, "b)  A $50 jacket has 10% off. Discount: ______   Sale price: ______", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "Which is the better deal on a $70 item: 25% off, or $20 off? Show your working.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Fractions and percentages of an amount | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 4 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Find one part", y, { color: C.PRIMARY });
    y = addBodyText(doc, "b)  9   (27 div 3)        c)  8   (80 div 10)", y);

    y = addSectionHeading(doc, "Section 2 - Find the fraction or percentage", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  18   (1/4 = 6, x 3).        b)  20   (1/3 = 10, x 2).        c)  9   (10% = 4.5, x 2).        d)  18   (10% = 6, x 3).", y);

    y = addSectionHeading(doc, "Section 3 - Sale prices", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Discount $10 (1/4 of 40), sale price $30.        b)  Discount $5 (10% of 50), sale price $45.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "25% of $70 is $17.50 off, leaving $52.50. $20 off leaves $50. The $20 off is the better deal by $2.50.", y);

    y = addTipBox(doc,
      "Watch for: students who divide by the top number of a fraction; students who divide by the percentage number instead of finding 10% first; students who give the discount as the final price instead of subtracting it.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Percentage change with multipliers.",
      color: C.SECONDARY,
      lessonInfo: `Lesson ${SESSION} | Year 8 challenge | extends Year 6 VC2M5N07, VC2M6N09`,
    });
    y = addTipBox(doc,
      "A multiplier does a percentage change in one step. 20% off means you pay 80%, so multiply by 0.8. A 15% increase means multiply by 1.15. To reverse a change, divide by the multiplier.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "A $50 jacket has 20% off. You pay 80%, so multiply: 0.8 x 50 = $40. One step, no separate subtraction.",
      y);

    y = addSectionHeading(doc, "Section 1 - Use a multiplier", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  30% off $80 (pay 70%): 0.7 x 80 = ______", y);
    y = addWriteLine(doc, "b)  15% off $60: ______", y);
    y = addWriteLine(doc, "c)  Add 10% GST to $45 (multiply by 1.1): ______", y);

    y = addSectionHeading(doc, "Section 2 - Trickier percentages", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  12.5% of $80 = ______        b)  33 1/3% of $90 = ______", y);

    y = addSectionHeading(doc, "Section 3 - Reverse percentages", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "A jacket costs $45 after 25% off. What was the original price?  ______", y);
    y = addWriteLine(doc, "Explain how you used the multiplier to work backwards.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) 56  b) 51  c) 49.50.   S2  a) 10 (12.5% = 1/8)  b) 30 (33 1/3% = 1/3).   S3  $45 is 75% of the original, so original = 45 div 0.75 = $60. Check: 0.75 x 60 = 45.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Year 8 Extension | Percentage change`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
