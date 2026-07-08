"use strict";

// Visual helper catalogue - internal QA reference deck, NOT a lesson.
// Renders every shared visual-anchor helper across the three grade bands
// (foundation / grade2 / grade56) so theme changes can be eyeballed in one
// pass. Rebuild after ANY change to themes/core/manipulatives.js,
// themes/builders/numeracy.js visual helpers, or grade-band sizing:
//   node scripts/build_and_check.js builds/build_visual_catalogue.js
//   python3 scripts/pptx_to_images.py output/Visual_Catalogue/*.pptx

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme } = require("../themes/factory");

const OUT_DIR = "output/Visual_Catalogue";
const FOOTER = "Visual helper catalogue | internal QA reference";

const BANDS = [
  { level: "foundation", label: "Foundation band" },
  { level: "grade2", label: "Year 1-2 band" },
  { level: "grade56", label: "Year 3-6 band" },
];

function addHelperLabel(T, s, x, y, text, w) {
  s.addText(text, {
    x, y, w: Math.min(w || 2.8, 9.5 - x), h: 0.26,
    fontSize: 11, fontFace: T.FONT_B, color: T.C.MUTED, italic: true, margin: 0,
  });
}

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  const T0 = createTheme("numeracy", "grade56", 0);
  T0.titleSlide(pres, "Visual Helper Catalogue",
    "Every shared visual anchor, rendered per grade band",
    "Internal QA reference | rebuild after theme changes",
    "QA reference deck. Compare bands for size, spacing and clarity. Not for classroom use.");

  BANDS.forEach((band) => {
    const T = createTheme("numeracy", band.level, 0);
    const { C, FONT_H, FONT_B, CONTENT_TOP } = T;

    // ── Slide 1: frames, counters, part-part-whole ──
    (() => {
      const s = pres.addSlide();
      T.addTopBar(s, C.PRIMARY);
      T.addBadge(s, band.label, { color: C.PRIMARY, w: 2.2 });
      T.addTitle(s, "Frames, counters, part-part-whole");

      addHelperLabel(T, s, 0.5, CONTENT_TOP + 0.05, "addTensFrame (filled 7)", 2.7);
      T.addTensFrame(s, 0.5, CONTENT_TOP + 0.38, 2.5, 7);

      addHelperLabel(T, s, 3.45, CONTENT_TOP + 0.05, "addFiveFrame (filled 3)", 2.7);
      T.addFiveFrame(s, 3.45, CONTENT_TOP + 0.38, 2.5, 3);

      addHelperLabel(T, s, 6.45, CONTENT_TOP + 0.05, "addDotCard (count 8)", 2.6);
      T.addDotCard(s, 6.45, CONTENT_TOP + 0.38, 1.15, 8);

      addHelperLabel(T, s, 0.5, CONTENT_TOP + 1.85, "addGroupedCounters (3 groups of 2)", 4.0);
      T.addGroupedCounters(s, 0.55, CONTENT_TOP + 2.3, 3, 2);

      addHelperLabel(T, s, 4.9, CONTENT_TOP + 1.85, "addPartPartWholeMat (7 = 4 + 3)", 3.4);
      T.addPartPartWholeMat(s, 4.9, CONTENT_TOP + 2.18, 2.7, 1.5, { whole: 7, partA: 4, partB: 3 });

      T.addFooter(s, FOOTER);
      s.addNotes("QA reference: " + band.label + " frames, counters and PPW mat.");
      T.runSlideDiagnostics(s, pres);
    })();

    // ── Slide 2: tracks, lines, chips ──
    (() => {
      const s = pres.addSlide();
      T.addTopBar(s, C.SECONDARY);
      T.addBadge(s, band.label, { color: C.SECONDARY, w: 2.2 });
      T.addTitle(s, "Number track, number line, chips");

      addHelperLabel(T, s, 0.5, CONTENT_TOP + 0.05, "addNumberTrack (1-10, highlight 7)", 4.2);
      T.addNumberTrack(s, 0.5, CONTENT_TOP + 0.38, 5.8, 1, 10, [7]);

      addHelperLabel(T, s, 6.7, CONTENT_TOP + 0.05, "addChipRow ([1/2, 3/4])", 2.7);
      T.addChipRow(s, 6.7, CONTENT_TOP + 0.38, 2.6, ["1/2", "3/4"]);

      addHelperLabel(T, s, 0.5, CONTENT_TOP + 1.75, "addNumberLine (0 to 2 in thirds, mark at 1)", 4.8);
      T.addNumberLine(s, 0.8, CONTENT_TOP + 2.55, 6.2,
        ["0", "", "", "1", "", "", "2"], [3]);

      T.addFooter(s, FOOTER);
      s.addNotes("QA reference: " + band.label + " track, line and chip row.");
      T.runSlideDiagnostics(s, pres);
    })();

    // ── Slide 3: strips, array, MAB, answer bar ──
    (() => {
      const s = pres.addSlide();
      T.addTopBar(s, C.ACCENT);
      T.addBadge(s, band.label, { color: C.ACCENT, w: 2.2 });
      T.addTitle(s, "Strips, array, MAB, answer bar");

      addHelperLabel(T, s, 0.5, CONTENT_TOP + 0.05, "addFractionStripSet (3 wholes in quarters)", 3.9);
      T.addFractionStripSet(s, 0.5, CONTENT_TOP + 0.38, 3.4, 1.55,
        [
          { denom: 4, shaded: 0, label: "1 whole", color: C.SECONDARY },
          { denom: 4, shaded: 0, label: "1 whole", color: C.SECONDARY },
          { denom: 4, shaded: 0, label: "1 whole", color: C.SECONDARY },
        ], { labelW: 0.8 });

      addHelperLabel(T, s, 4.6, CONTENT_TOP + 0.05, "addArray (3 x 4)", 2.2);
      T.addArray(s, 4.7, CONTENT_TOP + 0.45, 3, 4, { cellSize: 0.38 });

      addHelperLabel(T, s, 7.1, CONTENT_TOP + 0.05, "addBaseTenBlocks (1, 2, 3)", 2.35);
      T.addBaseTenBlocks(s, 7.1, CONTENT_TOP + 0.45, 1, 2, 3);

      addHelperLabel(T, s, 0.5, CONTENT_TOP + 2.15, "addRevealAnswerBar (two answers)", 3.7);
      T.addRevealAnswerBar(s, ["3 divided by 1/4 = 12", "twelve quarters fit into 3"],
        { y: CONTENT_TOP + 2.5, h: 0.85, fontSize: 20 });

      T.addFooter(s, FOOTER);
      s.addNotes("QA reference: " + band.label + " strips, array, MAB and answer bar.");
      T.runSlideDiagnostics(s, pres);
    })();

    // ── Slide 4: keyWordSlide (visual word card) ──
    T.keyWordSlide(pres, {
      word: "equivalent",
      meaning: "The same value, even when it looks different.",
      example: "One half is equivalent to two quarters.",
    }, "QA reference: " + band.label + " keyWordSlide word card.", FOOTER);

    // ── Slide 5: sparse CFU (density-aware hero sizing regression) ──
    T.cfuSlide(pres, "CFU", "Quick check",
      { technique: "Show Me Boards", question: "Which part turns liquid water into vapour?" },
      "QA reference: " + band.label + " sparse CFU - question must be hero-sized and centred, no dead bottom half.",
      FOOTER);
  });

  // ── Photo placeholder regression (grade56 annotatedModelSlide) ──
  {
    const T = createTheme("literacy", "grade56", 0);
    T.annotatedModelSlide(pres, "I Do", "Poster mockup with image region",
      [
        { text: "QA reference", role: "header" },
        { text: "The image region must read as a picture (sun and mountains), never grey bars.", role: "body" },
      ],
      "Save Water poster",
      [
        { label: "Heading", detail: "Big bold words carry the message." },
        { label: "Image", detail: "A picture shows the problem." },
        { label: "Colour", detail: "Cool blue fits the message." },
      ],
      "QA reference: photo placeholder and feature-card annotation floors.",
      FOOTER,
      {
        previewSpec: {
          components: [
            { kind: "masthead", text: "SAVE EVERY DROP" },
            { kind: "image" },
            { kind: "cta", text: "Turn off the tap" },
          ],
        },
      });
  }

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Visual Helper Catalogue.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
