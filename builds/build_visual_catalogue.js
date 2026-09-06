"use strict";

// Visual helper catalogue - internal QA reference deck, NOT a lesson.
// Renders every shared visual-anchor helper across the three grade bands
// (foundation / grade2 / grade56), every built-in pictogram, and the
// pattern builders (heroVisualSlide, choiceSlide, youDoSlide,
// textExtractSlide, cycle/process with icons) so theme changes can be
// eyeballed in one pass. Rebuild after ANY change to
// themes/core/manipulatives.js, themes/core/pictograms.js,
// themes/core/visualSpec.js, the builders, or grade-band sizing:
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

    // ── Slide 4: keyWordSlide (visual word card, with its pictogram) ──
    T.keyWordSlide(pres, {
      word: "equivalent",
      meaning: "The same value, even when it looks different.",
      example: "One half is equivalent to two quarters.",
      pictogram: "scales",
    }, "QA reference: " + band.label + " keyWordSlide word card with pictogram.", FOOTER);

    // ── Slide 5: sparse CFU (density-aware hero sizing regression) ──
    T.cfuSlide(pres, "CFU", "Quick check",
      { technique: "Show Me Boards", question: "Which part turns liquid water into vapour?" },
      "QA reference: " + band.label + " sparse CFU - question must be hero-sized and centred, no dead bottom half.",
      FOOTER);

    // ── Slide 6: heroVisualSlide (visual-only teaching slide) ──
    T.heroVisualSlide(pres, "I Do", "How many counters?",
      { type: "tensFrame", filled: 7 },
      "QA reference: " + band.label + " heroVisualSlide - the representation fills the panel.",
      FOOTER, { label: "Tens frame", prompt: "Show me on your fingers" });

    // ── Slide 7: choiceSlide with a marked answer ──
    (() => {
      const s = T.choiceSlide(pres, "CFU", "Which shows 7?", "Show me A, B or C on your board", [
        { visual: { type: "tensFrame", filled: 7 } },
        { visual: { type: "groupedCounters", groups: 2, per: 3 } },
        { visual: { type: "dotCard", count: 6 } },
      ], "QA reference: " + band.label + " choiceSlide with markChoice on A.", FOOTER, { badgeColor: C.ALERT });
      T.markChoice(s, 0);
    })();

    // ── Slide 8: youDoSlide with a mini model ──
    T.youDoSlide(pres, "Make 10 on your own", "Draw counters to make 10.",
      ["Look at the frame", "Draw the counters", "Write how many more"],
      "QA reference: " + band.label + " youDoSlide - task is the hero, steps are chips.", FOOTER,
      { where: "On your worksheet", visual: { type: "tensFrame", filled: 6 }, frame: "___ and ___ make 10", visualLabel: "6 in the frame" });

    // ── Slide 9: contentSlide sparse hero mode + spec right column ──
    T.contentSlide(pres, "Launch", C.PRIMARY, "Where does rain come from?",
      ["Turn and tell your partner: where do you think rain comes from?", "A puddle dries up on a sunny day. Where does that water go?"],
      "QA reference: " + band.label + " contentSlide with two short prompts renders as a hero panel.", FOOTER);
    T.contentSlide(pres, "I Do", C.PRIMARY, "Three groups of four",
      ["3 groups.", "4 in each group.", "12 altogether."],
      "QA reference: " + band.label + " contentSlide with a visual spec in the right column.", FOOTER,
      { type: "groupedCounters", groups: 3, per: 4 });
  });

  // ── Pattern builders and pictograms at Year 3-6 sizing ──
  {
    const T = createTheme("literacy", "grade56", 3);
    T.titleSlide(pres, "Persuasive Posters", "Headings, images and colour that make you act",
      "Year 4 Literacy | Persuasive Texts", "QA reference: literacy title slide with the subject glyph.");
    T.textExtractSlide(pres, "Read", "The storm",
      "The wind howled all night. Joey pressed his gnarled hands against the stable door as the rain crept under it, cold and silent.",
      "QA reference: textExtractSlide with highlighted phrases, source line and prompt bar.", FOOTER,
      { highlights: ["gnarled hands", "crept"], source: "Sample text, chapter 3", prompt: "Which word shows Joey is afraid?" });
    T.heroVisualSlide(pres, "We Do", "Which animals lay eggs?", {
      type: "table",
      rows: [["Animal", "Legs", "Lays eggs?"], ["Frog", "4", "Yes"], ["Dog", "4", "No"], ["Bird", "2", "Yes"]],
    }, "QA reference: table visual via heroVisualSlide.", FOOTER, { prompt: "Which row surprised you?" });
  }
  {
    const T = createTheme("numeracy", "grade34", 1);
    T.titleSlide(pres, "Equivalent Fractions", "Same amount, different name", "Year 3/4 Numeracy | Fractions",
      "QA reference: title slide carrying the lesson's own visual anchor.",
      { visual: { type: "fractionStrips", strips: [{ denom: 2, shaded: 1 }, { denom: 4, shaded: 2 }] } });
  }
  {
    const T = createTheme("science", "grade34", 2);
    T.cycleDiagramSlide(pres, "I Do", "The Water Cycle", "Watch and listen",
      ["I will name each part and how the water moves.", "Follow the arrows around the cycle."],
      "Water Cycle", [
        { label: "Evaporation", detail: "Sun heats the water", icon: "sun" },
        { label: "Condensation", detail: "Vapour cools into cloud", icon: "cloud" },
        { label: "Precipitation", detail: "Rain or snow falls", icon: "rain" },
        { label: "Collection", detail: "Gathers in rivers and sea", icon: "waves" },
      ], "QA reference: cycleDiagramSlide with stage icons.", FOOTER);
    T.processFlowSlide(pres, "I Do", "Where does food go?", "Think together", ["Where does food go first?"], [
      { label: "Mouth", detail: "Teeth chew, saliva softens", icon: "tooth" },
      { label: "Stomach", detail: "Acid breaks food down" },
      { label: "Intestines", detail: "Nutrients are absorbed" },
    ], "QA reference: processFlowSlide with an icon on the first chip.", FOOTER);
    T.heroVisualSlide(pres, "Launch", "How are you feeling today?",
      { type: "pictograms", items: ["happy", "calm", "worried", "sad", "angry"] },
      "QA reference: pictogram row as the hero visual.", FOOTER, { prompt: "Point to one. Turn and tell your partner." });
  }

  // ── Pictogram catalogue: every name, Year 3-6 sizing ──
  {
    const T = createTheme("inquiry", "grade56", 0);
    const names = T.listPictograms();
    const cols = 7;
    const rows = 4;
    const perSlide = cols * rows;
    for (let start = 0; start < names.length; start += perSlide) {
      const s = pres.addSlide();
      T.addTopBar(s, T.C.PRIMARY);
      T.addBadge(s, "Pictograms", { color: T.C.PRIMARY });
      T.addTitle(s, `Built-in pictograms ${start + 1}-${Math.min(start + perSlide, names.length)} of ${names.length}`);
      const colW = 9 / cols;
      const rowH = (5.1 - 1.3) / rows;
      const d = 0.56;
      names.slice(start, start + perSlide).forEach((name, i) => {
        const r = Math.floor(i / cols);
        const c = i % cols;
        const x = 0.5 + c * colW + (colW - d) / 2;
        const y = 1.3 + r * rowH + 0.06;
        T.addPictogram(s, name, x, y, d, { label: name, labelFontSize: 10, labelW: colW - 0.06 });
      });
      T.addFooter(s, FOOTER);
      s.addNotes(`QA reference: pictograms ${start + 1} to ${Math.min(start + perSlide, names.length)} - names as accepted by addPictogram, keyWordSlide { pictogram } and visual specs.`);
    }
  }

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
