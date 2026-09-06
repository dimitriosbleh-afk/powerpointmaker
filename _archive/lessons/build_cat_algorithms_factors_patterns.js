"use strict";

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf,
  addPdfHeader, addSectionHeading, addBodyText, addProblem, addTipBox,
  addPdfFooter, addWriteLine, addLinedArea,
  addResourceSlide, makeSessionResource, formatSessionResourceFileName,
  PAGE, hex,
} = require("../themes/pdf_helpers");

// ── Theme ──
const T = createTheme("numeracy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide,
  addStageBadge, STAGE_COLORS,
  withReveal,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape, addInstructionCard,
  CONTENT_TOP, SAFE_BOTTOM, SLIDE_W,
  runSlideDiagnostics,
} = T;

// ── Output paths ──
const UNIT = "CAT_Algorithms_Factors_Patterns";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "CAT Review - Algorithms Factors and Patterns.pptx";
const FOOTER = "CAT Review | Algorithms, Factors & Patterns | Grade 5/6";
const SESSION = 1;

async function build() {
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });

  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // ================================================================
  // SLIDE 1: Title
  // ================================================================
  titleSlide(
    pres,
    "Algorithms, Factors\n& Patterns",
    "CAT Review Session",
    "Grade 5/6 Numeracy",
    [
      "SAY:",
      "- We have our Common Assessment Task coming up, so today we are going to refresh the key ideas you will need.",
      "- This is a quick review session. We will go through the big concepts together, then you will have time to complete the CAT independently.",
      "",
      "DO:",
      "- Display slide as students settle.",
      "- Distribute CAT papers face down.",
      "",
      "TEACHER NOTES:",
      "This is a 30-minute revision session. Roughly 10-12 minutes of guided review, then students complete the CAT independently. The review targets the concept families in the assessment: number patterns, factors and multiples, divisibility, prime and square numbers, and conditional algorithms.",
      "",
      "WATCH FOR:",
      "- Students who seem anxious about the assessment. Reassure them this review will help them feel ready.",
    ].join("\n")
  );

  // ================================================================
  // SLIDE 2: LI / SC
  // ================================================================
  liSlide(
    pres,
    ["We are learning to recall and apply our knowledge of number patterns, factors, and algorithms so we are ready for our assessment"],
    [
      "I can identify and continue a number pattern by finding the rule",
      "I can find all the factors of a number and identify prime and square numbers",
      "I can follow a step-by-step algorithm to solve a number problem",
    ],
    [
      "SAY:",
      "- Here is what we are working towards today. Read through these three success criteria with me.",
      "- Read from slide: I can identify and continue a number pattern by finding the rule.",
      "- Read from slide: I can find all the factors of a number and identify prime and square numbers.",
      "- Read from slide: I can follow a step-by-step algorithm to solve a number problem.",
      "- These are the three big ideas on your assessment. Our quick review will touch on each one.",
      "",
      "DO:",
      "- Point to each SC as you read it aloud.",
      "- Leave the slide visible for 15-20 seconds so students can read independently.",
      "",
      "TEACHER NOTES:",
      "The three SC map directly to the three concept families in the CAT: patterns (Q1,2,7), factors/primes/squares (Q3,4,6,9,11,12), and algorithms (Q5,8,10). This gives students a mental map of what is coming.",
      "",
      "WATCH FOR:",
      "- Students who cannot read or understand the SC. Paraphrase if needed: patterns, factors, step-by-step rules.",
    ].join("\n"),
    FOOTER
  );

  // ================================================================
  // SLIDE 3: Review - Number Patterns (Question)
  // ================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Review");
    addTitle(s, "Number Patterns: Find the Rule");

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      const items = [
        { text: "What is the rule? What comes next?", role: "header" },
        { text: "", role: "spacer" },
        { text: "Pattern A:   8,  16,  24,  32,  ___,  ___", role: "body" },
        { text: "", role: "spacer" },
        { text: "Pattern B:   5,  9,  13,  17,  ___", role: "body" },
        { text: "", role: "spacer" },
        { text: "Pattern C:   2,  6,  18,  54,  ...", role: "body" },
      ];

      const textRuns = items.map((item, i) => {
        if (item.role === "spacer") return { text: "", options: { fontSize: 6, breakLine: true } };
        const isHeader = item.role === "header";
        return {
          text: item.text,
          options: {
            fontSize: isHeader ? 17 : 16,
            fontFace: isHeader ? FONT_H : FONT_B,
            color: isHeader ? C.PRIMARY : C.CHARCOAL,
            bold: isHeader,
            breakLine: i < items.length - 1,
          },
        };
      });

      s.addText(textRuns, {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 8.5, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes([
        "SAY:",
        "- Let's start with patterns. Look at Pattern A. What is happening each time? [Adding 8]",
        "- On your whiteboards, write the next two numbers for Pattern A.",
        "- Now Pattern B. What is the starting number and what is the rule? [Start at 5, add 4]",
        "- Pattern C is different. The jumps are getting bigger. What is the rule here? [Multiply by 3]",
        "",
        "DO:",
        "- Give students 30 seconds to write answers for all three patterns on whiteboards.",
        "- Cold call for each pattern before revealing answers.",
        "",
        "CFU CHECKPOINT:",
        "Technique: Show Me Boards",
        "Script:",
        "- Say: Write your answers for all three patterns on your whiteboard. You have 30 seconds. Hold up on my signal.",
        "- Scan for: Pattern A = 40, 48. Pattern B = 21. Pattern C = multiply by 3.",
        "PROCEED: If 80%+ correct, click to reveal and move on.",
        "PIVOT: If Pattern C causes confusion, clarify: The rule is not adding - the gap between numbers is growing. Ask: What do I multiply 2 by to get 6? [3] What about 6 to get 18? [3]. So the rule is multiply by 3 each time.",
        "",
        "TEACHER NOTES:",
        "Covers CAT Q1 (add 8), Q2 (start at 5, add 4), Q7 (multiply by 3). The key distinction is additive vs multiplicative patterns - most errors come from assuming all patterns are additive.",
        "",
        "WATCH FOR:",
        "- Students who write 36, 40 for Pattern A (adding inconsistent amounts). Prompt: What is 32 + 8?",
        "- Students who write 'add 12' for Pattern C. They are finding the first gap only, not checking if it holds.",
      ].join("\n"));
  }

  // ================================================================
  // SLIDE 3b: Review - Number Patterns (Reveal)
  // ================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Review");
    addTitle(s, "Number Patterns: Find the Rule");

    addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

    const items = [
      { text: "What is the rule? What comes next?", role: "header" },
      { text: "", role: "spacer" },
      { text: "Pattern A:   8,  16,  24,  32,  40,  48", role: "body" },
      { text: "Rule: Add 8 each time", role: "answer" },
      { text: "", role: "spacer" },
      { text: "Pattern B:   5,  9,  13,  17,  21", role: "body" },
      { text: "Rule: Start at 5, add 4 each time", role: "answer" },
      { text: "", role: "spacer" },
      { text: "Pattern C:   2,  6,  18,  54,  162", role: "body" },
      { text: "Rule: Multiply by 3 each time", role: "answer" },
    ];

    const textRuns = items.map((item, i) => {
      if (item.role === "spacer") return { text: "", options: { fontSize: 6, breakLine: true } };
      const isHeader = item.role === "header";
      const isAnswer = item.role === "answer";
      return {
        text: item.text,
        options: {
          fontSize: isHeader ? 17 : isAnswer ? 14 : 16,
          fontFace: isHeader ? FONT_H : FONT_B,
          color: isHeader ? C.PRIMARY : isAnswer ? C.SUCCESS : C.CHARCOAL,
          bold: isHeader || isAnswer,
          breakLine: i < items.length - 1,
        },
      };
    });

    s.addText(textRuns, {
      x: 0.75, y: CONTENT_TOP + 0.15, w: 8.5, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
      fontFace: FONT_B, margin: 0, valign: "top",
    });

    addFooter(s, FOOTER);
    s.addNotes([
      "SAY:",
      "- Check your answers. Tick if correct, fix if not.",
      "- Pattern A adds 8 each time. Pattern B starts at 5, adds 4. Pattern C multiplies by 3.",
      "- The big takeaway: always check whether the pattern is adding or multiplying.",
      "",
      "DO:",
      "- Click to reveal. Students tick correct, fix errors.",
      "- Emphasise the additive vs multiplicative distinction.",
      "",
      "TEACHER NOTES:",
      "The tick-and-fix cycle gives students immediate feedback on their retrieval attempt. The additive vs multiplicative distinction is the highest-yield takeaway for the CAT.",
      "",
      "WATCH FOR:",
      "- Students who got Pattern C wrong. Note these students for monitoring during the CAT.",
    ].join("\n"));
  }

  // ================================================================
  // SLIDE 4: Review - Factors, Multiples & Divisibility (with reveal)
  // ================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Review");
    addTitle(s, "Factors, Multiples & Divisibility");

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const leftItems = [
        { text: "Quick Recall", role: "header" },
        { text: "", role: "spacer" },
        { text: "A factor is a whole number that divides evenly into another number.", role: "body" },
        { text: "", role: "spacer" },
        { text: "A multiple is the result of multiplying a number by a whole number.", role: "body" },
        { text: "", role: "spacer" },
        { text: "Divisible means it divides evenly with no remainder.", role: "body" },
      ];

      const leftRuns = leftItems.map((item, i) => {
        if (item.role === "spacer") return { text: "", options: { fontSize: 5, breakLine: true } };
        const isHeader = item.role === "header";
        return {
          text: item.text,
          options: {
            fontSize: isHeader ? 16 : 13,
            fontFace: isHeader ? FONT_H : FONT_B,
            color: isHeader ? C.PRIMARY : C.CHARCOAL,
            bold: isHeader,
            breakLine: i < leftItems.length - 1,
          },
        };
      });

      s.addText(leftRuns, {
        x: 0.75, y: CONTENT_TOP + 0.12, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right column - challenges
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

      const rightItems = [
        { text: "On Your Whiteboards", role: "header" },
        { text: "", role: "spacer" },
        { text: "1. List ALL factors of 24", role: "body" },
        { text: "", role: "spacer" },
        { text: "2. Which number is a multiple of BOTH 4 and 6?  8 / 18 / 24 / 16", role: "body" },
        { text: "", role: "spacer" },
        { text: "3. Which can be divided evenly by 9?  54 / 52 / 56 / 58", role: "body" },
      ];

      const rightRuns = rightItems.map((item, i) => {
        if (item.role === "spacer") return { text: "", options: { fontSize: 5, breakLine: true } };
        const isHeader = item.role === "header";
        return {
          text: item.text,
          options: {
            fontSize: isHeader ? 16 : 13,
            fontFace: isHeader ? FONT_H : FONT_B,
            color: isHeader ? C.ACCENT : C.CHARCOAL,
            bold: isHeader,
            breakLine: i < rightItems.length - 1,
          },
        };
      });

      s.addText(rightRuns, {
        x: 5.45, y: CONTENT_TOP + 0.12, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes([
        "SAY:",
        "- Let's refresh factors, multiples, and divisibility. Read the definitions on the left.",
        "- Now try the three challenges on the right. Write your answers on your whiteboards.",
        "- For question 1, remember to work in pairs: 1 and 24, 2 and 12, and so on.",
        "",
        "DO:",
        "- Give students 60 seconds for all three questions.",
        "- Circulate and check factor pair strategy for Q1.",
        "",
        "CFU CHECKPOINT:",
        "Technique: Show Me Boards",
        "Script:",
        "- Say: Hold up your whiteboards on my signal. I want to see all your factors of 24.",
        "- Scan for: 1, 2, 3, 4, 6, 8, 12, 24 (8 factors). Q2 = 24. Q3 = 54.",
        "PROCEED: If 80%+ correct on all three, reveal and move on.",
        "PIVOT: If students miss factors of 24, model the systematic factor pair approach on the board: start at 1 x 24, then 2 x 12, then 3 x 8, then 4 x 6. When the pairs meet, you have found them all.",
        "",
        "TEACHER NOTES:",
        "Covers CAT Q3 (common multiples), Q4 (divisibility by 9), Q6 (factors of 24). The factor pair strategy is the highest-yield move here - it prevents students from listing random factors and missing some.",
        "",
        "WATCH FOR:",
        "- Students who list only small factors (1, 2, 3, 4) and forget 6, 8, 12, 24. They are not using pairs.",
        "- Students who confuse factors and multiples. Quick clarifier: factors go INTO, multiples come OUT.",
      ].join("\n"));
  }

  // ================================================================
  // SLIDE 4b: Review - Factors, Multiples & Divisibility (Reveal)
  // ================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Review");
    addTitle(s, "Factors, Multiples & Divisibility");

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    const leftRuns = [
        { text: "Quick Recall", options: { fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 5, breakLine: true } },
        { text: "A factor is a whole number that divides evenly into another number.", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 5, breakLine: true } },
        { text: "A multiple is the result of multiplying a number by a whole number.", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 5, breakLine: true } },
        { text: "Divisible means it divides evenly with no remainder.", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
      ];

      s.addText(leftRuns, {
        x: 0.75, y: CONTENT_TOP + 0.12, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right column - answers
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SUCCESS });

      const rightRuns = [
        { text: "Answers", options: { fontSize: 16, fontFace: FONT_H, color: C.SUCCESS, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 5, breakLine: true } },
        { text: "1. Factors of 24:", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "   1, 2, 3, 4, 6, 8, 12, 24  (8 factors)", options: { fontSize: 13, fontFace: FONT_B, color: C.SUCCESS, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 5, breakLine: true } },
        { text: "2. Multiple of both 4 and 6:", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "   24  (4 x 6 = 24, and 6 x 4 = 24)", options: { fontSize: 13, fontFace: FONT_B, color: C.SUCCESS, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 5, breakLine: true } },
        { text: "3. Divisible by 9:", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "   54  (54 / 9 = 6, no remainder)", options: { fontSize: 13, fontFace: FONT_B, color: C.SUCCESS, bold: true, breakLine: true } },
      ];

      s.addText(rightRuns, {
        x: 5.45, y: CONTENT_TOP + 0.12, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes([
        "SAY:",
        "- Check your answers. Tick and fix.",
        "- For factors of 24, the trick is working in pairs: 1 and 24, 2 and 12, 3 and 8, 4 and 6. That gives us 8 factors.",
        "- For common multiples, 24 is in both the 4 times table and the 6 times table.",
        "- For divisibility, 54 divided by 9 is exactly 6. The others leave remainders.",
        "",
        "DO:",
        "- Click to reveal answers. Students tick and fix.",
        "- Briefly model the factor pair strategy on the board if students struggled.",
        "",
        "TEACHER NOTES:",
        "The factor pair method is the single most important strategy for the factors questions on the CAT. Emphasise working systematically from 1 upwards.",
        "",
        "WATCH FOR:",
        "- Students who got fewer than 8 factors. They likely skipped pairs - note for monitoring during CAT.",
      ].join("\n"));
  }

  // ================================================================
  // SLIDE 5: Review - Prime & Square Numbers (Question)
  // ================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Review");
    addTitle(s, "Prime Numbers & Square Numbers");

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });

      const leftRuns = [
        { text: "Definitions", options: { fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 5, breakLine: true } },
        { text: "Prime number:", options: { fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, breakLine: true } },
        { text: "Has exactly 2 factors: 1 and itself.", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
        { text: "e.g. 7 = 1 x 7 (only two factors)", options: { fontSize: 12, fontFace: FONT_B, color: C.MUTED, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Square number:", options: { fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, breakLine: true } },
        { text: "Made by multiplying a whole number by itself.", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
        { text: "e.g. 9 = 3 x 3", options: { fontSize: 12, fontFace: FONT_B, color: C.MUTED, breakLine: true } },
      ];

      s.addText(leftRuns, {
        x: 0.75, y: CONTENT_TOP + 0.12, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right column - challenges
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

      const rightRuns = [
        { text: "Challenge", options: { fontSize: 16, fontFace: FONT_H, color: C.ACCENT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 5, breakLine: true } },
        { text: "1. List all prime numbers between 20 and 40", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "2. List all square numbers less than 50", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Tip: For primes, test each number - can anything other than 1 and itself divide into it?", options: { fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, breakLine: true } },
      ];

      s.addText(rightRuns, {
        x: 5.45, y: CONTENT_TOP + 0.12, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes([
        "SAY:",
        "- Two special number types you will see on the assessment: prime numbers and square numbers.",
        "- Read the definitions on the left. A prime number has exactly two factors. A square number is a number times itself.",
        "- Try the two challenges on the right. Use your whiteboards.",
        "- For primes, go through each number: 21 - can anything divide into it? 3 x 7, so not prime. Keep going.",
        "",
        "DO:",
        "- Give students 60-90 seconds for both challenges.",
        "- Circulate and check for common errors (including 1 as prime, missing 37).",
        "",
        "CFU CHECKPOINT:",
        "Technique: Finger Voting + Cold Call",
        "Script:",
        "- Say: How many prime numbers did you find between 20 and 40? Show me on your fingers.",
        "- Scan for: 4 fingers (23, 29, 31, 37). Cold call one student to read their list.",
        "PROCEED: If 80%+ found all 4 primes, reveal and move on.",
        "PIVOT: If students include non-primes (e.g. 21, 27, 33, 35), ask: What are the factors of 21? [1, 3, 7, 21] - that is four factors, so not prime. A prime has EXACTLY two.",
        "",
        "TEACHER NOTES:",
        "Covers CAT Q11 (primes 20-40) and Q12 (squares < 50). Common errors: including 1 as prime (1 has only one factor, not two), and forgetting 37 or including 27/33.",
        "",
        "MISCONCEPTIONS:",
        "- Misconception: 1 is a prime number.",
        "  Why: Students learn 'only divisible by 1 and itself' and assume 1 fits.",
        "  Impact: Incorrect answers on any prime-identification task.",
        "  Quick correction: A prime has EXACTLY two different factors. 1 only has one factor (itself). So 1 is not prime.",
        "",
        "WATCH FOR:",
        "- Students listing 21, 27, 33, or 35 as prime. They are not testing divisibility.",
        "- Students forgetting 1 x 1 = 1 is a square number. Prompt: What is 1 x 1?",
      ].join("\n"));
  }

  // ================================================================
  // SLIDE 5b: Review - Prime & Square Numbers (Reveal)
  // ================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Review");
    addTitle(s, "Prime Numbers & Square Numbers");

    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });

    const leftRuns = [
      { text: "Definitions", options: { fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 5, breakLine: true } },
        { text: "Prime number:", options: { fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, breakLine: true } },
        { text: "Has exactly 2 factors: 1 and itself.", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
        { text: "e.g. 7 = 1 x 7 (only two factors)", options: { fontSize: 12, fontFace: FONT_B, color: C.MUTED, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Square number:", options: { fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, breakLine: true } },
        { text: "Made by multiplying a whole number by itself.", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
        { text: "e.g. 9 = 3 x 3", options: { fontSize: 12, fontFace: FONT_B, color: C.MUTED, breakLine: true } },
      ];

      s.addText(leftRuns, {
        x: 0.75, y: CONTENT_TOP + 0.12, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right column - answers
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SUCCESS });

      const rightRuns = [
        { text: "Answers", options: { fontSize: 16, fontFace: FONT_H, color: C.SUCCESS, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 5, breakLine: true } },
        { text: "Primes between 20 and 40:", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "23,  29,  31,  37", options: { fontSize: 15, fontFace: FONT_B, color: C.SUCCESS, bold: true, breakLine: true } },
        { text: "(4 prime numbers)", options: { fontSize: 11, fontFace: FONT_B, color: C.MUTED, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Square numbers less than 50:", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "1, 4, 9, 16, 25, 36, 49", options: { fontSize: 15, fontFace: FONT_B, color: C.SUCCESS, bold: true, breakLine: true } },
        { text: "(1x1, 2x2, 3x3, 4x4, 5x5, 6x6, 7x7)", options: { fontSize: 11, fontFace: FONT_B, color: C.MUTED, breakLine: true } },
      ];

      s.addText(rightRuns, {
        x: 5.45, y: CONTENT_TOP + 0.12, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes([
        "SAY:",
        "- Check your answers. There are exactly 4 primes between 20 and 40: 23, 29, 31, 37.",
        "- For square numbers, start from 1 x 1 and work up: 1, 4, 9, 16, 25, 36, 49. Seven square numbers.",
        "- Remember: 1 is a square number but NOT a prime number.",
        "",
        "DO:",
        "- Click to reveal. Students tick and fix.",
        "- If students missed primes, briefly show the testing method: 21 = 3 x 7 (not prime), 23 = only 1 x 23 (prime).",
        "",
        "TEACHER NOTES:",
        "The distinction between 1 as a square but not a prime is a common assessment trap. Make this explicit.",
        "",
        "WATCH FOR:",
        "- Students who missed 37. They may have stopped checking after 31.",
      ].join("\n"));
  }

  // ================================================================
  // SLIDE 6: Review - Algorithms / Step-by-Step Rules
  // ================================================================
  contentSlide(
    pres,
    "Review",
    STAGE_COLORS["1"],
    "Algorithms: Follow the Steps",
    [
      "An algorithm is a set of step-by-step rules.",
      "Read each step carefully. Do exactly what it says.",
      "Example: Start with 12.",
      "  Step 1: 12 is even, so halve it -> 6",
      "  Step 2: 6 is even, so halve it -> 3",
      "  Step 3: 3 is odd, so add 1 -> 4",
      "Key: Check the condition FIRST, then do the action.",
    ],
    [
      "SAY:",
      "- The last concept family is algorithms - step-by-step rules. You will see a few of these on the assessment.",
      "- The trick is to read each step carefully and check the condition before doing anything.",
      "- Look at this example. We start with 12. Step 1: is 12 even or odd? [Even] So we halve it - 6.",
      "- Step 2: is 6 even or odd? [Even] Halve it again - 3. Step 3: is 3 even or odd? [Odd] So we add 1 - the answer is 4.",
      "- The condition changes what you do. Do not just halve every time.",
      "",
      "DO:",
      "- Walk through the example step by step, pausing at each condition check.",
      "- Point to each step on the slide as you narrate.",
      "",
      "CFU CHECKPOINT:",
      "Technique: Choral Response",
      "Script:",
      "- Say: Step 1 - is 12 even or odd? [Class responds: Even!] So what do we do? [Halve it!]",
      "- Say: Step 2 - is 6 even or odd? [Even!] What do we do? [Halve it!]",
      "- Say: Step 3 - is 3 even or odd? [Odd!] What do we do? [Add 1!]",
      "PROCEED: If class responds correctly and confidently, move to the CAT launch.",
      "PIVOT: If students hesitate on the condition, re-model: Every step has an IF. IF even, do this. IF odd, do that. Circle the IF on the slide.",
      "",
      "TEACHER NOTES:",
      "Covers CAT Q5, Q8, Q10. The conditional logic (if/then) is the hardest concept for some students. The key error is ignoring the condition and applying the same operation every time. The choral response keeps all students engaged and makes hesitation visible.",
      "",
      "WATCH FOR:",
      "- Students who halve 3 instead of adding 1. They are ignoring the odd/even condition.",
      "- Students who are confident here. They are ready for the CAT.",
    ].join("\n"),
    FOOTER
  );

  // ================================================================
  // SLIDE 7: CAT Launch (You Do)
  // ================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Common Assessment Task", { color: STAGE_COLORS["4"] });

    // Left card - instructions
    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["4"] });

    const instrRuns = [
      { text: "Your Assessment", options: { fontSize: 17, fontFace: FONT_H, color: STAGE_COLORS["4"], bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "First: Write your name and date on the paper.", options: { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 4, breakLine: true } },
      { text: "Next: Read each question carefully. Check if it is multiple choice or short answer.", options: { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 4, breakLine: true } },
      { text: "Then: Work through the questions in order. Show your working where asked.", options: { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
      { text: "", options: { fontSize: 4, breakLine: true } },
      { text: "Finally: Check your answers before handing in.", options: { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true } },
    ];

    s.addText(instrRuns, {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
      fontFace: FONT_B, margin: 0, valign: "top",
    });

    // Right card - reminders
    addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SUCCESS });

    const reminderRuns = [
      { text: "Remember", options: { fontSize: 17, fontFace: FONT_H, color: C.SUCCESS, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "Patterns: Check if it adds or multiplies", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true, bullet: true } },
      { text: "", options: { fontSize: 4, breakLine: true } },
      { text: "Factors: Work in pairs (1 x ?, 2 x ?, ...)", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true, bullet: true } },
      { text: "", options: { fontSize: 4, breakLine: true } },
      { text: "Primes: Exactly 2 factors only", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true, bullet: true } },
      { text: "", options: { fontSize: 4, breakLine: true } },
      { text: "Squares: A number times itself", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true, bullet: true } },
      { text: "", options: { fontSize: 4, breakLine: true } },
      { text: "Algorithms: Read the condition FIRST", options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, breakLine: true, bullet: true } },
    ];

    s.addText(reminderRuns, {
      x: 5.45, y: CONTENT_TOP + 0.12, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
      fontFace: FONT_B, margin: 0, valign: "top",
    });

    addFooter(s, FOOTER);
    s.addNotes([
      "SAY:",
      "- Right, you are ready. Turn your CAT paper over.",
      "- First: name and date. Next: read each question carefully - some are multiple choice, some are short answer.",
      "- Work through in order. Show your working on the Working Out sheet where it asks.",
      "- Use the reminders on the right if you get stuck. These are the key strategies we just reviewed.",
      "- When you finish, go back and check your answers.",
      "",
      "DO:",
      "- Signal students to turn over CAT papers.",
      "- Leave this slide displayed throughout the assessment as a reference.",
      "- Circulate silently. Note which questions cause the most difficulty for post-assessment analysis.",
      "- Do not provide content help during the assessment - only clarify instructions if needed.",
      "",
      "ENABLING & EXTENDING:",
      "ENABLING PROMPT:",
      "- Task: Students who are stuck on factors questions can draw arrays or use counters to find factor pairs. Provide counters if available.",
      "- Extra Notes: Some students may need the algorithm questions read aloud. This is acceptable for the assessment if it is their usual accommodation.",
      "EXTENDING PROMPT:",
      "- Task: Students who finish early should check all working, especially the short-answer questions (Q8, Q9, Q10, Q11, Q12). Then attempt to write their own algorithm question for a partner.",
      "",
      "TEACHER NOTES:",
      "This is the You Do / independent assessment phase. The slide stays up as a reference. Students work independently. Circulate and note patterns in errors for post-assessment planning.",
      "",
      "WATCH FOR:",
      "- Students who are stuck on Q5 or Q10 (algorithms). They may be ignoring conditions. A quiet prompt: 'Read each step. What does it tell you to check first?' is acceptable.",
      "- Students rushing through multiple choice without checking. A quiet tap on the desk and a gesture to re-read can help.",
    ].join("\n"));
  }

  // ================================================================
  // SLIDE 8: Closing
  // ================================================================
  closingSlide(
    pres,
    "Look at our three success criteria. Give me a thumbs up, sideways, or down for each one. Which one felt strongest? Which one do you want more practice with?",
    [
      "I can identify and continue a number pattern by finding the rule",
      "I can find all the factors of a number and identify prime and square numbers",
      "I can follow a step-by-step algorithm to solve a number problem",
    ],
    [
      "SAY:",
      "- Pens down. Let's come back to our success criteria.",
      "- Read from slide: I can identify and continue a number pattern by finding the rule. Thumbs up, sideways, or down.",
      "- Read from slide: I can find all the factors of a number and identify prime and square numbers. Thumbs.",
      "- Read from slide: I can follow a step-by-step algorithm to solve a number problem. Thumbs.",
      "- Whatever you showed, that is honest self-assessment and that is valuable.",
      "",
      "DO:",
      "- Read each SC aloud. Pause for thumbs on each one.",
      "- Scan the room and note which SC gets the most sideways/down.",
      "- Collect CAT papers.",
      "",
      "TEACHER NOTES:",
      "The self-assessment data tells you which concept family needs the most follow-up. If algorithms got the most thumbs-down, plan a reteach. Use the CAT results to confirm or adjust.",
      "",
      "WATCH FOR:",
      "- Students who show thumbs-down on all three. Check in privately - they may need confidence support as much as content support.",
    ].join("\n")
  );

  // ================================================================
  // Write PPTX
  // ================================================================
  const pptxPath = path.join(LESSON_FOLDER, PPTX_NAME);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to", pptxPath);

  // ================================================================
  // No companion PDFs needed for this session - the CAT itself is the
  // student resource (pre-printed and distributed by the teacher).
  // ================================================================

  console.log("Build complete:", UNIT);
}

build().catch((err) => {
  console.error("Build failed:", err);
  process.exit(1);
});
