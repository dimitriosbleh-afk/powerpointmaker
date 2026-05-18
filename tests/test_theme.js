"use strict";

/**
 * Test harness — generates a sample PPTX for any (subject, yearLevel, variant) combo.
 *
 * Usage:
 *   node tests/test_theme.js <subject> <yearLevel> [variant]
 *
 * Examples:
 *   node tests/test_theme.js literacy grade56 0
 *   node tests/test_theme.js numeracy grade34 2
 *   node tests/test_theme.js science foundation
 *
 * Generates: output/test_<subject>_<yearLevel>_v<variant>.pptx
 */

const pptxgen = require("pptxgenjs");
const fs      = require("fs");
const { createTheme, weekToVariant } = require("../themes/factory");

const subject   = process.argv[2] || "literacy";
const yearLevel = process.argv[3] || "grade34";
const variant   = parseInt(process.argv[4]) || 0;

console.log(`Creating test deck: ${subject} / ${yearLevel} / variant ${variant}`);

const T = createTheme(subject, yearLevel, variant);
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  withReveal, addCard, addFooter, addTextOnShape,
} = T;

const FOOTER = `Test  |  ${subject} / ${yearLevel} / v${variant}  |  ${FONT_H} / ${FONT_B}`;

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Theme Test";
  pres.title  = `Test: ${subject} ${yearLevel} v${variant}`;

  // 1. Title slide
  titleSlide(pres,
    `${subject.charAt(0).toUpperCase() + subject.slice(1)} — Week ${variant + 1}`,
    `Year Level: ${yearLevel}  |  Variant: ${variant}`,
    `Fonts: ${FONT_H} / ${FONT_B}  |  Palette: ${T._paletteName}`,
    "Test title slide"
  );

  // 2. LI slide
  liSlide(pres,
    ["Understand the new theme system", "Apply colour palettes to slide builders"],
    ["I can identify the PRIMARY, SECONDARY, and ACCENT colours",
     "I can describe how the factory pattern works",
     "I can generate a test deck for any subject/level combo"],
    "Test LI slide",
    FOOTER
  );

  // 3. Content slide (no drawRight)
  contentSlide(pres,
    "Stage 1", C.PRIMARY,
    "Content Slide — Full Width",
    [
      "This is a full-width content slide with no right column.",
      "PRIMARY colour: #" + C.PRIMARY,
      "SECONDARY colour: #" + C.SECONDARY,
      "ACCENT colour: #" + C.ACCENT,
      "ALERT colour: #" + C.ALERT,
      "Font pair: " + FONT_H + " / " + FONT_B,
    ],
    "Test content slide",
    FOOTER
  );

  // 4. Content slide (with drawRight)
  contentSlide(pres,
    "Stage 2", C.SECONDARY,
    "Content Slide — Two Column",
    [
      "This slide has a right-column visual.",
      "The drawRight callback adds custom content.",
      "Card strips use the badge colour.",
    ],
    "Test content slide with drawRight",
    FOOTER,
    (s) => {
      // Right-side colour swatch cards
      const swatches = [
        { label: "PRIMARY",   color: C.PRIMARY },
        { label: "SECONDARY", color: C.SECONDARY },
        { label: "ACCENT",    color: C.ACCENT },
        { label: "ALERT",     color: C.ALERT },
        { label: "SUCCESS",   color: C.SUCCESS },
      ];
      swatches.forEach((sw, i) => {
        const y = 1.3 + i * 0.72;
        addTextOnShape(s, sw.label, {
          x: 6.2, y, w: 3.2, h: 0.55, rectRadius: 0.08,
          fill: { color: sw.color },
        }, {
          fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
        });
      });
    }
  );

  // 5. CFU slide
  cfuSlide(pres,
    "Check", "Quick Check",
    "Think-Pair-Share",
    "Can you explain the difference between PRIMARY and SECONDARY colours in this theme?",
    "Test CFU slide",
    FOOTER
  );

  // 6. withReveal on CFU
  withReveal(
    () => cfuSlide(pres,
      "CFU", "Click-to-Reveal Test",
      "Show Me Boards",
      "What colour is used for the ACCENT in this theme?",
      "Test withReveal — question slide",
      FOOTER
    ),
    (s) => {
      addTextOnShape(s, "Answer: #" + C.ACCENT, {
        x: 3.0, y: 4.0, w: 4.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.ACCENT },
      }, {
        fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true,
      });
    }
  );

  // 7. Subject-specific slides
  if (subject === "literacy") {
    if (T.vocabSlide) {
      T.vocabSlide(pres, "Theme", "noun",
        "A unifying visual design applied consistently across presentation slides.",
        "The midnight scholar theme gave the slides a sophisticated, literary feel.",
        "Test vocab slide", FOOTER);
    }
    if (T.quoteSlide) {
      T.quoteSlide(pres, "Read Aloud", "Chapter 1",
        "The factory pattern binds colours to builders, creating a seamless visual identity.",
        "p. 42",
        "How does this design pattern help maintain consistency?",
        "Test quote slide", FOOTER);
    }
    if (T.modellingSlide) {
      T.modellingSlide(pres, "I Do — Watch Me", "Teacher Modelling Example",
        "Left column: key terms, definitions, and structure notes for the teacher.",
        "Right column: worked example text showing the modelled writing in italic font.",
        "Test modelling slide", FOOTER);
    }
    if (T.pairShareSlide) {
      T.pairShareSlide(pres, "Discuss with Your Partner",
        ["What makes a good colour palette?",
         "How many variants does each subject have?",
         "Why do we use semantic colour names?"],
        "Test pair-share slide", FOOTER);
    }
  }

  if (subject === "numeracy") {
    if (T.workedExSlide) {
      T.workedExSlide(pres, 2, "I Do", "Worked Example",
        ["Step 1: Identify the place value of each digit.",
         "Step 2: Partition the number into thousands, hundreds, tens, ones.",
         "Step 3: Represent using base-10 blocks.",
         "Step 4: Check by re-composing."],
        "Test worked example slide", FOOTER);
    }
    if (T.exitTicketSlide) {
      T.exitTicketSlide(pres,
        ["Write 4 567 in expanded form.",
         "What is the value of the 5 in 3 508?",
         "Draw a place value chart for 12.34."],
        "Test exit ticket slide", FOOTER);
    }
  }

  if (subject === "inquiry") {
    if (T.investigationSlide) {
      T.investigationSlide(pres, "Investigate", "Our Big Question",
        "How does water cycle through our local environment?",
        "We think most water in our school yard evaporates rather than soaking into the ground.",
        ["Observe the school yard after rain", "Measure puddle sizes over 3 days", "Record findings"],
        "Test investigation slide", FOOTER);
    }
    if (T.findingsSlide) {
      T.findingsSlide(pres, "Findings", "What We Discovered",
        ["Puddles in concrete areas evaporated within 2 hours",
         "Puddles on grass areas soaked in within 30 minutes",
         "Shaded areas kept water longest"],
        "Different surfaces affect water absorption and evaporation differently.",
        "Test findings slide", FOOTER);
    }
  }

  if (subject === "wellbeing") {
    if (T.scenarioSlide) {
      T.scenarioSlide(pres, "Scenario", "What Would You Do?",
        "You notice a classmate sitting alone at lunch every day this week. They look sad but haven't said anything to anyone.",
        ["What might this person be feeling?",
         "What could you do to help?",
         "Who else could you involve?"],
        "Test scenario slide", FOOTER);
    }
    if (T.reflectionSlide) {
      T.reflectionSlide(pres, "My Reflection",
        ["One thing I learned about myself today...",
         "Something I want to try this week...",
         "A person I want to thank and why..."],
        "Test reflection slide", FOOTER);
    }
  }

  if (subject === "science") {
    if (T.experimentSlide) {
      T.experimentSlide(pres, "Experiment", "Testing Absorbency",
        "Paper towels will absorb more water than newspaper.",
        ["3 paper towels", "3 newspaper sheets", "Water", "Measuring cup", "Timer"],
        ["Place each material flat on the table",
         "Pour 50mL of water onto each",
         "Wait 30 seconds",
         "Measure remaining water"],
        "Test experiment slide", FOOTER);
    }
    if (T.observationSlide) {
      T.observationSlide(pres, "Observe", "Record Your Observations",
        ["What happened to the paper towel?",
         "What happened to the newspaper?",
         "Which material absorbed more water?"],
        "Test observation slide", FOOTER);
    }
    if (T.conclusionSlide) {
      T.conclusionSlide(pres, "Conclude", "Drawing Conclusions",
        "Which material absorbs the most water?",
        ["Paper towels absorbed 45mL out of 50mL",
         "Newspaper absorbed only 20mL out of 50mL",
         "Paper towels left almost no puddle"],
        "Paper towels are much more absorbent than newspaper because of their thicker, more porous fibres.",
        "Test conclusion slide", FOOTER);
    }
  }

  // 8. Closing slide
  closingSlide(pres,
    "What are three things you noticed about this theme's visual design?",
    ["Consistent colour usage across all slide types",
     "Dark title/closing slides, light content slides",
     "Badge and card strip colours match the pedagogical stage"],
    "Test closing slide"
  );

  // Write output
  const outDir  = "output";
  const outFile = `test_${subject}_${yearLevel}_v${variant}.pptx`;
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  await pres.writeFile({ fileName: `${outDir}/${outFile}` });
  console.log(`Written: ${outDir}/${outFile}`);
}

build().catch((err) => {
  console.error("Build failed:", err);
  process.exit(1);
});
