"use strict";

// Shared data + visual anchors for the Year 5/6 Information Report unit
// (Term 3 Weeks 9 and 10, eight sessions). Every session imports from here so
// the unit anchor, the model texts and the checklists stay byte-identical
// across the fortnight (megaprompt section 79, unit anchor consistency).
//
// UNIT ANCHOR (locked, megaprompt sections 5b and 79 - built from the
// teacher's OWN planning template wording: "Introduction: (General
// classification)", "Aspect 1 / 2 / 3", "Conclusion: (Summary/interesting
// fact)". Never swapped for other textbook labels mid-unit):
//
//   "Classify it. Describe it, one aspect at a time. Wrap it up.
//    Facts all the way through."
//
// SOURCES. The teacher supplied the unit plan, the class topic (Green Tree
// Frog), the planning template, the editing checklist and the Two Stars and a
// Wish routine. No mentor text was supplied, so every model report below is
// ORIGINAL, written for this unit and fact-checked. Nothing here is quoted
// from, or attributed to, a published source (megaprompt sections 5 and 5a).
// If the teacher has real mentor texts, they swap straight in: the structure
// taught is the same.

// --- The anchor -------------------------------------------------------------

const ANCHOR_PHRASE =
  "Classify it. Describe it, one aspect at a time. Wrap it up. Facts all the way through.";

// Short form for a slide face where the full phrase would not fit.
const ANCHOR_SHORT = "Classify. Describe. Wrap up. Facts all the way through.";

const MAP_BANDS = {
  classify: { key: "classify", label: "Classify", sub: "What is it? What group is it in?" },
  describe: { key: "describe", label: "Describe", sub: "One aspect per paragraph" },
  wrap: { key: "wrap", label: "Wrap up", sub: "Sum up, or one last interesting fact" },
};

const GENERIC_ASPECTS = ["Aspect 1", "Aspect 2", "Aspect 3"];
const FROG_ASPECTS = ["Appearance", "Habitat", "Diet"];

const FACTS_BAR = "FACTS all the way through";

// The class topic for every piece of modelling, both weeks (teacher-supplied).
const CLASS_TOPIC = "The Green Tree Frog";

// --- Model report 1: the class topic ----------------------------------------

const FROG = {
  title: "The Green Tree Frog",
  classify:
    "The green tree frog is an amphibian that lives in the warm, wet parts of northern and eastern Australia. It is one of the largest tree frogs in the country.",
  aspects: [
    {
      heading: "Appearance",
      topicSentence: "Green tree frogs have smooth, bright green skin and a creamy white belly.",
      text: "Green tree frogs have smooth, bright green skin and a creamy white belly. An adult can grow to about ten centimetres long. Wide, round pads on its toes help it grip and climb.",
    },
    {
      heading: "Habitat",
      topicSentence: "Green tree frogs shelter in trees, reeds and rocks near still water.",
      text: "Green tree frogs shelter in trees, reeds and rocks near still water. They also live in cool, damp places around houses, such as downpipes, water tanks and letterboxes.",
    },
    {
      heading: "Diet",
      topicSentence: "Green tree frogs hunt at night.",
      text: "Green tree frogs hunt at night. They eat insects such as moths, crickets and cockroaches. A sticky tongue flicks out and traps the prey.",
    },
  ],
  wrap:
    "The green tree frog is a common but remarkable Australian animal. Scientists study the chemicals in its skin because some of them kill germs.",
  caption:
    "Figure 1: The toe pads of a green tree frog. Wide, sticky pads let it climb smooth surfaces.",
};

// --- Model report 2: a place (used where the You Do must differ) -------------

const REEF = {
  title: "The Great Barrier Reef",
  classify:
    "The Great Barrier Reef is a coral reef system off the coast of Queensland, Australia. It is the largest reef system on Earth.",
  aspects: [
    {
      heading: "Size",
      topicSentence: "The reef stretches for more than two thousand kilometres along the Queensland coast.",
      text: "The reef stretches for more than two thousand kilometres along the Queensland coast. It is made of thousands of separate reefs and hundreds of islands.",
    },
    {
      heading: "Living things",
      topicSentence: "The reef is home to more than fifteen hundred species of fish.",
      text: "The reef is home to more than fifteen hundred species of fish. Turtles, dolphins, sharks and dugongs also feed and shelter there.",
    },
    {
      heading: "Threats",
      topicSentence: "Warm ocean water is the biggest danger to the coral.",
      text: "Warm ocean water is the biggest danger to the coral. When the water heats up the coral turns white, which is called coral bleaching. Pollution from the land can damage the reef as well.",
    },
  ],
  wrap:
    "The Great Barrier Reef is one of the natural wonders of the world. It is so large that it can be seen from space.",
  caption:
    "Figure 1: A map of the Great Barrier Reef along the Queensland coast.",
};

// --- Model report 3: a phenomenon -------------------------------------------

const THUNDER = {
  title: "Thunder",
  classify:
    "Thunder is the loud sound made by lightning. It can be heard during storms all over the world.",
  aspects: [
    {
      heading: "How it forms",
      topicSentence: "A flash of lightning heats the air around it to a very high temperature.",
      text: "A flash of lightning heats the air around it to a very high temperature. The hot air expands so quickly that it crashes into the cooler air beside it. That crash is the sound we call thunder.",
    },
    {
      heading: "Why it comes later",
      topicSentence: "Light travels much faster than sound.",
      text: "Light travels much faster than sound. This is why lightning is seen before thunder is heard, even though the two happen at the same moment.",
    },
    {
      heading: "Counting the distance",
      topicSentence: "Thunder can be used to work out how far away a storm is.",
      text: "Thunder can be used to work out how far away a storm is. Sound travels about one kilometre every three seconds, so counting the seconds between the flash and the rumble gives a rough distance.",
    },
  ],
  wrap:
    "Thunder cannot hurt anyone on its own, but it is a useful warning. It tells people that lightning is close.",
  caption:
    "Figure 1: Lightning heats the air, and the air crashes back together to make thunder.",
};

// --- Features of an information report (Session 1) --------------------------

const REPORT_FEATURES = [
  { label: "Facts", detail: "Everything can be checked. No opinions." },
  { label: "Present tense", detail: "Frogs hunt at night. Not hunted." },
  { label: "Headings", detail: "One aspect under each heading." },
  { label: "Visual + caption", detail: "The picture teaches as well." },
];

// --- Language features (Session 4) ------------------------------------------

const FACT_OPINION = [
  { text: "Green tree frogs can grow to about ten centimetres long.", fact: true },
  { text: "Green tree frogs are the best pets in Australia.", fact: false },
  { text: "The green tree frog eats moths and crickets.", fact: true },
  { text: "Everybody should love frogs.", fact: false },
  { text: "Green tree frogs hunt at night.", fact: true },
  { text: "Frogs are much nicer than lizards.", fact: false },
];

const TENSE_FIXES = [
  { before: "The green tree frog lived near still water.", after: "The green tree frog lives near still water." },
  { before: "Yesterday the frog ate a moth.", after: "Green tree frogs eat moths." },
  { before: "Frogs will hunt after dark.", after: "Frogs hunt after dark." },
];

const TECHNICAL_WORDS = [
  { word: "amphibian", meaning: "an animal that can live in water and on land" },
  { word: "nocturnal", meaning: "active at night" },
  { word: "habitat", meaning: "the place where an animal lives" },
  { word: "predator", meaning: "an animal that hunts other animals" },
  { word: "adaptation", meaning: "a body part or habit that helps an animal survive" },
];

const NOUN_GROUPS = [
  { plain: "the frog", grown: "the large, bright green tree frog" },
  { plain: "its tongue", grown: "its long, sticky tongue" },
  { plain: "the pads", grown: "the wide, round pads on its toes" },
  { plain: "a call", grown: "a low, barking call" },
];

// --- Research and notes (Session 3) -----------------------------------------

// The teacher's own wondering-to-question example, kept in their words.
const RESEARCH_QUESTIONS = [
  { wondering: "What does it look like?", aspect: "Appearance" },
  { wondering: "Where does it live?", aspect: "Habitat" },
  { wondering: "What does it eat?", aspect: "Diet" },
  { wondering: "How does it survive?", aspect: "Adaptations" },
];

// Two short texts on the SAME aspect, written for this unit. The teaching
// point is two texts saying overlapping things in different words. If real
// books are available, swap them in - the routine is unchanged.
const DIET_SOURCES = [
  {
    label: "Source 1",
    text: "Green tree frogs are night hunters. They wait very still, then flick out a sticky tongue to catch moths, crickets and cockroaches.",
    notes: ["hunt at night", "wait still, then strike", "sticky tongue", "moths, crickets, cockroaches"],
  },
  {
    label: "Source 2",
    text: "The diet of the green tree frog is mostly insects. It feeds after dark, when moths and other insects are active. Large green tree frogs have also been recorded eating spiders and small frogs.",
    notes: ["mostly insects", "feeds after dark", "big frogs eat spiders and small frogs"],
  },
];

const DIET_PARAPHRASE =
  "Green tree frogs feed after dark. They mostly eat insects such as moths and crickets, which they catch with a sticky tongue.";

const CALL_SOURCES = [
  {
    label: "Source 1",
    text: "The call of the green tree frog is a low, barking crawk. Males call from trees and gutters, and they call more often before rain.",
    notes: ["low barking call", "males call", "call more before rain"],
  },
  {
    label: "Source 2",
    text: "Green tree frogs are noisy in the wet season. The deep barking call of the male repeats over and over, and it grows louder when a storm is coming.",
    notes: ["noisy in the wet season", "deep call, repeats", "louder before a storm"],
  },
];

const CALL_PARAPHRASE =
  "Male green tree frogs make a low, barking call. They call more often before rain and through the wet season.";

// The sentence modelled in the I Do and re-used in the Session 3 hinge, so
// students judge notes taken from a sentence they have just seen worked.
const NOTE_SOURCE_SENTENCE =
  "The green tree frog is an amphibian that lives in the warm, wet parts of northern Australia.";

const NOTE_KEY_WORDS = ["amphibian", "warm, wet", "northern Australia"];

const NOTE_MODEL_POINTS = ["amphibian", "warm, wet places", "northern Australia"];

// Session 3 hinge: three sets of notes from that same sentence.
const NOTE_OPTIONS = [
  { key: "A", text: NOTE_SOURCE_SENTENCE, verdict: "copied whole" },
  { key: "B", text: "amphibian / warm, wet places / northern Australia", verdict: "best" },
  { key: "C", text: "frogs", verdict: "too little" },
];

// --- Editing, captions and publishing (Week 10) -----------------------------

// Teacher-supplied checklist, wording locked (megaprompt section 5b).
const EDIT_CHECKLIST = [
  "Clear introduction that classifies the topic?",
  "Body paragraphs, each on one aspect, with a topic sentence?",
  "Factual language, with no opinions?",
  "Present tense used correctly?",
  "Technical words used where they help?",
  "Noun groups that describe?",
  "Spelling and punctuation: capitals, full stops, commas?",
];

// The teacher's own peer-feedback routine, wording locked.
const STARS_AND_WISH = "Two Stars (two things done well) and a Wish (one area for improvement)";

const VISUAL_TYPES = [
  { label: "Labelled diagram", use: "Shows the parts of something" },
  { label: "Map", use: "Shows where something is" },
  { label: "Photograph", use: "Shows what it really looks like" },
  { label: "Table", use: "Compares numbers or groups" },
];

const CAPTION_OPTIONS = [
  { key: "A", text: "A green tree frog." },
  { key: "B", text: "Figure 2: A green tree frog on a window. Its toe pads grip the smooth glass." },
  { key: "C", text: "I found this picture on the internet." },
];

const PUBLISH_RULES = [
  "One clear title at the top",
  "Subheadings that match your aspects",
  "One font and one size for the body text",
  "The visual next to the part it belongs to",
  "The caption directly under the visual",
];

const PUBLISH_FORMATS = [
  { label: "Typed page", detail: "Title, subheadings, image with text wrapped around it." },
  { label: "Infographic", detail: "Facts in boxes, icons, one strong diagram." },
  { label: "Slide deck", detail: "One aspect per slide, few words, big visual." },
];

// --- Topic choice (Week 10 Session 1) ---------------------------------------

// The teacher's own broad themes, quoted from the planner.
const CLASS_THEMES = ["Extreme Weather", "Renewable Energy", "Space Exploration", "Animals"];

const TOPIC_TESTS = [
  { topic: "Animals", verdict: "too big", why: "A whole library could not cover it." },
  { topic: "My dog Rex", verdict: "too small", why: "One animal is a story, not a report." },
  { topic: "Emperor penguins", verdict: "just right", why: "A whole group, easy to split into aspects." },
];

// The teacher's planning template, wording kept, dashes normalised to ASCII.
const PLAN_ROWS = [
  { label: "Topic", hint: "" },
  { label: "Introduction", hint: "General classification" },
  { label: "Aspect 1", hint: "for example Appearance - key facts and notes" },
  { label: "Aspect 2", hint: "for example Habitat - key facts and notes" },
  { label: "Aspect 3", hint: "for example Diet - key facts and notes" },
  { label: "Conclusion", hint: "Summary or interesting fact" },
];

const SOURCE_LINE_UNIT =
  "Teacher-supplied unit plan, planning template, editing checklist and class topic (Green Tree Frog).";
const SOURCE_LINE_MODEL =
  "Model report written for this unit and fact-checked; not quoted from a published text.";

// --- Page mockup specs ------------------------------------------------------

/**
 * A tidy published information-report page, as a structured mockup spec.
 *
 * This is the ONE page mockup for the unit. Every builder that renders a
 * report page (annotatedModelSlide, compareVisualSlide) takes a spec from
 * here, so I Do and We Do never differ in visual fidelity (CLAUDE.md,
 * previewSpec consistency rule).
 *
 * @param {object} C       theme palette
 * @param {object} report  FROG / REEF / THUNDER
 * @param {object} [o]     { captionText }
 */
function reportPageSpec(C, report, o) {
  const opts = o || {};
  return {
    accent: C.PRIMARY,
    pageFill: "FFFFFF",
    components: [
      { kind: "masthead", text: report.title, scale: 0.62 },
      { kind: "textBlock", count: 2, scale: 0.55 },
      { kind: "heading", text: report.aspects[0].heading, scale: 0.42 },
      // The image region needs real height or the wireframe scene squashes
      // into a grey wedge and stops reading as a picture.
      { kind: "photo", scale: 1.55 },
      { kind: "caption", text: opts.captionText || report.caption, scale: 0.42 },
    ],
  };
}

/**
 * The same report with no visual at all: words only. The counter-example for
 * "what does a diagram and its caption actually add?"
 */
function noVisualPageSpec(C, report) {
  return {
    accent: C.SECONDARY,
    pageFill: "FFFFFF",
    components: [
      { kind: "masthead", text: report.title, scale: 0.62 },
      { kind: "textBlock", count: 3, scale: 1.0 },
      { kind: "heading", text: report.aspects[0].heading, scale: 0.42 },
      { kind: "textBlock", count: 4, scale: 1.2 },
    ],
  };
}

/**
 * The same content laid out badly: no title bar, headings missing, the visual
 * stranded at the bottom away from the text it belongs to, no caption. Used
 * only as the counter-example beside reportPageSpec.
 */
function messyPageSpec(C, report) {
  return {
    accent: C.MUTED,
    pageFill: "FFFFFF",
    components: [
      { kind: "textBlock", count: 3, scale: 0.9 },
      { kind: "textBlock", count: 3, scale: 0.9 },
      { kind: "photo", scale: 1.0, overlayText: "picture, no caption" },
    ],
  };
}

// --- Shared slide chrome ----------------------------------------------------

// Every click-revealed answer bar in this unit sits at the same place, so the
// content floor above it is identical on every slide. Anything drawn below
// CONTENT_FLOOR would be covered by the reveal, and the layout diagnostics
// only see the finished stack, so keeping the floor is on the author.
const REVEAL_Y = 4.28;
const REVEAL_H = 0.78;
const CONTENT_FLOOR = 4.13;

/** Slide chrome for a custom layout: top bar, badge, title. Returns the slide. */
function customSlide(pres, T, badgeText, badgeColor, title, opts) {
  const o = opts || {};
  const s = pres.addSlide();
  T.addTopBar(s, badgeColor);
  T.addBadge(s, badgeText, { color: badgeColor, w: o.badgeW || 1.95 });
  T.addTitle(s, title, o.titleOpts);
  return s;
}

/**
 * Draw the unit anchor: the Report Map.
 *
 * Three stacked bands (Classify / Describe / Wrap up) with the Describe band
 * carrying one chip per aspect, and a FACTS bar underneath. This is the ONE
 * representation for the fortnight - never hand-draw a variant in a session
 * script, extend this instead (megaprompt section 79).
 *
 * @param {object} slide  PptxGenJS slide
 * @param {object} T      theme object from createTheme
 * @param {object} o      { x, y, w, h, focus, aspects, showFacts, compact, parts }
 *   focus: "classify" | "describe" | "wrap" | "facts" | null - the band being
 *          taught today; other bands render as pale outline cards.
 *   parts: subset of ["classify","describe","wrap","facts"] to draw. Positions
 *          are always computed for the FULL layout, so calling this repeatedly
 *          with different parts inside clickBuild() steps builds the map up one
 *          band per click without anything shifting.
 * @returns {number} the y coordinate just below the map
 */
function drawReportMap(slide, T, o) {
  const { C, FONT_H, FONT_B } = T;
  const x = Number(o.x);
  const y = Number(o.y);
  const w = Number(o.w);
  const h = Number(o.h);
  const focus = o.focus || null;
  const compact = Boolean(o.compact);
  const showFacts = o.showFacts !== false;
  const aspects = (Array.isArray(o.aspects) && o.aspects.length === 3) ? o.aspects : GENERIC_ASPECTS;
  const parts = Array.isArray(o.parts) ? o.parts : null;
  const wants = (key) => !parts || parts.indexOf(key) !== -1;

  const gap = compact ? 0.08 : 0.10;
  const factsH = showFacts ? (compact ? 0.38 : 0.46) : 0;
  const bandsH = h - factsH - (showFacts ? gap : 0) - gap * 2;
  // Describe carries three chips, so it needs roughly twice the height.
  const unit = bandsH / 4.2;
  const hClassify = unit;
  const hDescribe = unit * 2.2;
  const hWrap = unit;

  const labelSize = compact ? 14 : 17;
  const subSize = compact ? 10 : 12;
  const chipSize = compact ? 10.5 : 12.5;

  const bandFills = { classify: C.PRIMARY, describe: C.SECONDARY, wrap: C.ASSESS };

  function bandShell(key, by, bh) {
    const isFocus = !focus || focus === key;
    const fill = isFocus ? bandFills[key] : C.WHITE;
    const textColor = isFocus ? C.WHITE : C.MUTED;
    slide.addShape("roundRect", {
      x, y: by, w, h: bh, rectRadius: 0.08,
      fill: { color: fill },
      line: isFocus ? null : { color: C.MUTED, width: 1 },
    });
    return { isFocus, fill, textColor };
  }

  const yClassify = y;
  const yDescribe = yClassify + hClassify + gap;
  const yWrap = yDescribe + hDescribe + gap;
  const yFacts = yWrap + hWrap + gap;

  if (wants("classify")) {
    const st = bandShell("classify", yClassify, hClassify);
    slide.addText([
      { text: MAP_BANDS.classify.label, options: { fontSize: labelSize, bold: true, color: st.textColor, fontFace: FONT_H, breakLine: true } },
      { text: MAP_BANDS.classify.sub, options: { fontSize: subSize, color: st.textColor, fontFace: FONT_B } },
    ], {
      x: x + 0.18, y: yClassify + 0.03, w: w - 0.36, h: hClassify - 0.06,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  }

  if (wants("describe")) {
    const st = bandShell("describe", yDescribe, hDescribe);
    const headH = Math.min(0.52, hDescribe * 0.40);
    slide.addText([
      { text: MAP_BANDS.describe.label, options: { fontSize: labelSize, bold: true, color: st.textColor, fontFace: FONT_H, breakLine: true } },
      { text: MAP_BANDS.describe.sub, options: { fontSize: subSize, color: st.textColor, fontFace: FONT_B } },
    ], {
      x: x + 0.18, y: yDescribe + 0.03, w: w - 0.36, h: headH,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    const chipGap = 0.08;
    const chipY = yDescribe + headH + 0.04;
    const chipH = Math.max(0.24, hDescribe - headH - 0.12);
    const chipW = (w - 0.36 - chipGap * 2) / 3;
    aspects.forEach((label, i) => {
      const cx = x + 0.18 + i * (chipW + chipGap);
      slide.addShape("roundRect", {
        x: cx, y: chipY, w: chipW, h: chipH, rectRadius: 0.06,
        fill: { color: C.WHITE },
        line: { color: st.isFocus ? C.WHITE : C.MUTED, width: 1 },
      });
      slide.addText(String(label), {
        x: cx + 0.04, y: chipY, w: chipW - 0.08, h: chipH,
        fontSize: chipSize, fontFace: FONT_B, bold: true,
        color: st.isFocus ? C.SECONDARY : C.MUTED,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });
  }

  if (wants("wrap")) {
    const st = bandShell("wrap", yWrap, hWrap);
    slide.addText([
      { text: MAP_BANDS.wrap.label, options: { fontSize: labelSize, bold: true, color: st.textColor, fontFace: FONT_H, breakLine: true } },
      { text: MAP_BANDS.wrap.sub, options: { fontSize: subSize, color: st.textColor, fontFace: FONT_B } },
    ], {
      x: x + 0.18, y: yWrap + 0.03, w: w - 0.36, h: hWrap - 0.06,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  }

  if (showFacts && wants("facts")) {
    const isFocus = !focus || focus === "facts";
    slide.addShape("roundRect", {
      x, y: yFacts, w, h: factsH, rectRadius: 0.08,
      fill: { color: isFocus ? C.ACCENT : C.WHITE },
      line: isFocus ? null : { color: C.MUTED, width: 1 },
    });
    slide.addText(String(o.factsText || FACTS_BAR), {
      x: x + 0.18, y: yFacts, w: w - 0.36, h: factsH,
      fontSize: compact ? 11.5 : 13.5, bold: true,
      color: isFocus ? C.WHITE : C.MUTED, fontFace: FONT_B,
      align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
    return yFacts + factsH;
  }

  return yWrap + hWrap;
}

/**
 * Draw a report extract as a page-like card with real, readable sentences.
 *
 * This is the close-reading anchor: a structured mockup shows what a report
 * page LOOKS like, but students cannot read a mockup. Sections are stacked and
 * given height in proportion to their text, so a long paragraph does not
 * crowd a short one.
 *
 * @param {object} o { x, y, w, h, title, sections, bodySize, labelColor,
 *                     titleColor, showLabels, fill }
 *   sections: [{ label, text, highlight, color, textColor }]
 * @returns {number} y below the card
 */
function drawReportExtract(slide, T, o) {
  const { C, FONT_H, FONT_B } = T;
  const x = Number(o.x);
  const y = Number(o.y);
  const w = Number(o.w);
  const h = Number(o.h);
  const sections = Array.isArray(o.sections) ? o.sections : [];
  const pad = o.pad != null ? Number(o.pad) : 0.14;
  const showLabels = o.showLabels !== false;
  const bodySize = o.bodySize || 13;
  const labelSize = o.labelSize || 11;
  const gap = 0.08;

  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: o.fill || C.WHITE },
    line: { color: C.MUTED, width: 1 },
  });

  let cursor = y + pad;
  const innerW = w - pad * 2;

  if (o.title) {
    const titleH = o.titleH || 0.40;
    slide.addText(String(o.title), {
      x: x + pad, y: cursor, w: innerW, h: titleH,
      fontSize: o.titleSize || 20, fontFace: FONT_H, bold: true,
      color: o.titleColor || C.PRIMARY,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
    cursor += titleH + gap;
  }

  const avail = (y + h - pad) - cursor - gap * Math.max(sections.length - 1, 0);
  const labelH = 0.24;
  const labelTotal = showLabels
    ? sections.filter((sec) => sec && sec.label).length * (labelH + 0.03)
    : 0;
  // Weight by ESTIMATED LINES, not raw characters. Four one-line sentences
  // must get four equal boxes; weighting by length would give the longest
  // sentence a taller box holding the same single line, which renders as
  // random-looking vertical gaps.
  const charsPerLine = Math.max(20, innerW / (bodySize * 0.0072));
  const weights = Array.isArray(o.weights) && o.weights.length === sections.length
    ? o.weights.map(Number)
    : sections.map((sec) => Math.max(1, Math.ceil(String((sec && sec.text) || "").length / charsPerLine)));
  const weightTotal = weights.reduce((a, b) => a + b, 0) || 1;
  const textTotal = Math.max(avail - labelTotal, 0.4);

  sections.forEach((sec, i) => {
    const accent = (sec && sec.color) || o.labelColor || C.SECONDARY;
    if (showLabels && sec && sec.label) {
      const lw = Math.min(innerW, o.labelW || 1.55);
      slide.addShape("roundRect", {
        x: x + pad, y: cursor, w: lw, h: labelH, rectRadius: 0.05,
        fill: { color: accent },
      });
      slide.addText(String(sec.label), {
        x: x + pad, y: cursor, w: lw, h: labelH,
        fontSize: labelSize, fontFace: FONT_B, bold: true, color: C.WHITE,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
      cursor += labelH + 0.03;
    }
    const textH = textTotal * (weights[i] / weightTotal);
    if (sec && sec.highlight) {
      slide.addShape("roundRect", {
        x: x + pad - 0.05, y: cursor - 0.02, w: innerW + 0.10, h: textH + 0.04, rectRadius: 0.05,
        fill: { color: C.BG_LIGHT },
        line: { color: accent, width: 1 },
      });
    }
    slide.addText(String((sec && sec.text) || ""), {
      x: x + pad, y: cursor, w: innerW, h: textH,
      fontSize: bodySize, fontFace: FONT_B,
      color: (sec && sec.textColor) || C.CHARCOAL,
      valign: "top", margin: 0, fit: "shrink", shrinkText: true,
    });
    cursor += textH + gap;
  });

  return y + h;
}

/**
 * A/B/C option stack for a hinge question, sized to stop above the reveal bar.
 *
 * @param {object} o { x, y, w, optionH, gap, color, fontSize }
 * @returns {number} y below the last option
 */
function drawOptionStack(slide, T, options, o) {
  const { C, FONT_H, FONT_B } = T;
  const x = o.x != null ? Number(o.x) : 0.5;
  const y = Number(o.y);
  const w = o.w != null ? Number(o.w) : 9;
  const oh = o.optionH != null ? Number(o.optionH) : 0.66;
  const gap = o.gap != null ? Number(o.gap) : 0.12;
  const accent = o.color || C.PRIMARY;
  const fontSize = o.fontSize || 16;

  options.forEach((opt, i) => {
    const oy = y + i * (oh + gap);
    slide.addShape("roundRect", {
      x, y: oy, w, h: oh, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: accent, width: 1.5 },
    });
    slide.addShape("roundRect", {
      x: x + 0.18, y: oy + (oh - 0.44) / 2, w: 0.44, h: 0.44, rectRadius: 0.22,
      fill: { color: accent },
    });
    slide.addText(String(opt.key), {
      x: x + 0.18, y: oy + (oh - 0.44) / 2, w: 0.44, h: 0.44,
      fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    slide.addText(String(opt.text), {
      x: x + 0.78, y: oy, w: w - 1.05, h: oh,
      fontSize, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  });

  return y + options.length * oh + Math.max(options.length - 1, 0) * gap;
}

/**
 * Dot-point note card: the paper note-taking box, drawn on screen so the model
 * matches the worksheet exactly.
 *
 * @param {object} o { x, y, w, h, heading, points, color, fontSize, headSize }
 * @returns {number} y below the card
 */
function drawNoteCard(slide, T, o) {
  const { C, FONT_B } = T;
  const x = Number(o.x);
  const y = Number(o.y);
  const w = Number(o.w);
  const h = Number(o.h);
  const accent = o.color || C.SUCCESS;
  const points = Array.isArray(o.points) ? o.points : [];
  const headH = 0.34;

  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: C.WHITE }, line: { color: accent, width: 1.5 },
  });
  slide.addShape("roundRect", {
    x, y, w, h: headH, rectRadius: 0.08, fill: { color: accent },
  });
  slide.addText(String(o.heading || "Notes"), {
    x: x + 0.12, y, w: w - 0.24, h: headH,
    fontSize: o.headSize || 12.5, fontFace: FONT_B, bold: true, color: C.WHITE,
    valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
  });

  const bodyY = y + headH + 0.08;
  const bodyH = h - headH - 0.16;
  const rowH = points.length ? bodyH / points.length : bodyH;
  points.forEach((point, i) => {
    const py = bodyY + i * rowH;
    slide.addShape("roundRect", {
      x: x + 0.16, y: py + rowH / 2 - 0.045, w: 0.09, h: 0.09, rectRadius: 0.045,
      fill: { color: accent },
    });
    slide.addText(String(point), {
      x: x + 0.34, y: py, w: w - 0.50, h: rowH,
      fontSize: o.fontSize || 13.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  });

  return y + h;
}

/**
 * Draw the teacher's planning template as labelled rows.
 *
 * Row positions are always computed for the FULL row list, so calling this
 * repeatedly with different `parts` inside clickBuild() steps fills the plan
 * one row per click without anything shifting.
 *
 * @param {object} o { x, y, w, h, rows, parts, labelW, fontSize, labelSize }
 *   rows: [{ label, text, hint, color }] - text null renders an empty row.
 * @returns {number} y below the block
 */
function drawPlanRows(slide, T, o) {
  const { C, FONT_B } = T;
  const x = Number(o.x);
  const y = Number(o.y);
  const w = Number(o.w);
  const h = Number(o.h);
  const rows = Array.isArray(o.rows) ? o.rows : [];
  const parts = Array.isArray(o.parts) ? o.parts : null;
  const wants = (i) => !parts || parts.indexOf(i) !== -1;
  const labelW = o.labelW != null ? Number(o.labelW) : 1.55;
  const fontSize = o.fontSize || 12.5;
  const labelSize = o.labelSize || 11.5;
  const gap = o.gap != null ? Number(o.gap) : 0.09;
  const rowH = (h - gap * Math.max(rows.length - 1, 0)) / Math.max(rows.length, 1);

  rows.forEach((row, i) => {
    if (!wants(i)) return;
    const ry = y + i * (rowH + gap);
    const accent = row.color || C.PRIMARY;
    slide.addShape("roundRect", {
      x, y: ry, w, h: rowH, rectRadius: 0.06,
      fill: { color: C.WHITE }, line: { color: accent, width: 1.2 },
    });
    slide.addShape("roundRect", {
      x, y: ry, w: labelW, h: rowH, rectRadius: 0.06,
      fill: { color: accent },
    });
    slide.addText(String(row.label), {
      x: x + 0.04, y: ry, w: labelW - 0.08, h: rowH,
      fontSize: labelSize, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
    const body = row.text != null ? String(row.text) : (row.hint ? String(row.hint) : "");
    if (body) {
      slide.addText(body, {
        x: x + labelW + 0.14, y: ry, w: w - labelW - 0.28, h: rowH,
        fontSize, fontFace: FONT_B,
        color: row.text != null ? C.CHARCOAL : C.MUTED,
        italic: row.text == null,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    }
  });

  return y + h;
}

/**
 * Exit ticket built as a hero panel rather than the default builder card.
 *
 * The shared exitTicketSlide caps a single prompt card for Y3-6 and offsets it
 * high, which leaves a dead bottom third on a short prompt (megaprompt 15h).
 * Rather than fork a builder used by 250+ decks, this unit uses one consistent
 * hero panel across all eight sessions. The SC target stays in the notes only,
 * never on the face (section 0a item 18).
 */
function exitTicketPanel(pres, T, cfg) {
  const { C, FONT_H, FONT_B } = T;
  const s = pres.addSlide();
  const strip = C.ASSESS;

  s.background = { color: C.BG_CARD };
  s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: strip } });
  T.addBadge(s, "Exit Ticket", { color: strip, x: 0.5, y: 0.2, w: 1.85, h: 0.38 });
  s.addText(cfg.title || "Show what you know", {
    x: 0.5, y: 0.68, w: 9, h: 0.60,
    fontSize: 30, fontFace: FONT_H, color: strip, bold: true, margin: 0,
    fit: "shrink", shrinkText: true,
  });

  let y = 1.40;
  if (cfg.topic) {
    T.addTextOnShape(s, String(cfg.topic), {
      x: 0.5, y, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: strip },
    }, {
      fontSize: 19, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0.1,
    });
    y += 0.86;
  }

  const cueH = cfg.cue ? 0.60 : 0;
  const taskBottom = cfg.cue ? 4.20 : 4.88;
  const taskH = taskBottom - y;
  T.addCard(s, 0.5, y, 9, taskH, { strip, fill: C.WHITE });
  s.addText(String(cfg.task || ""), {
    x: 0.85, y: y + 0.10, w: 8.3, h: taskH - 0.20,
    fontSize: cfg.taskSize || 28, fontFace: FONT_B, color: C.CHARCOAL,
    align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
  });

  if (cfg.cue) {
    T.addTextOnShape(s, String(cfg.cue), {
      x: 0.5, y: 4.30, w: 9, h: cueH, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: strip, width: 1.5 },
    }, {
      fontSize: 16, fontFace: FONT_B, color: strip, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }

  if (cfg.footer) T.addFooter(s, cfg.footer);
  if (cfg.notes) s.addNotes(cfg.notes);
  return s;
}

// --- PDF twins --------------------------------------------------------------

/**
 * The Report Map on paper. Same three bands, same aspect chips, same FACTS bar
 * as the slide anchor, so the printed card and the screen match exactly.
 *
 * @param {PDFDocument} doc
 * @param {number} y
 * @param {object} o { colors:{classify,describe,wrap,facts}, aspects, left,
 *                     width, bandH, describeH }
 * @returns {number} y below the map
 */
function addReportMapPdf(doc, y, o) {
  const opts = o || {};
  const left = opts.left != null ? opts.left : 50;
  const width = opts.width != null ? opts.width : 495.28;
  const colors = opts.colors || {};
  const aspects = (Array.isArray(opts.aspects) && opts.aspects.length === 3) ? opts.aspects : GENERIC_ASPECTS;
  // Floors, not defaults: below these the label and its subtitle no longer fit
  // inside the band and the text clips against the fill edge.
  const bandH = Math.max(34, opts.bandH || 34);
  const describeH = Math.max(66, opts.describeH || 66);
  const gap = 8;

  function band(label, sub, color, top, height) {
    doc.save();
    doc.roundedRect(left, top, width, height, 4).fill(color);
    doc.fontSize(13).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(label, left + 10, top + 5, { width: width - 20, lineBreak: false });
    doc.fontSize(9.5).font("Sans").fillColor("#FFFFFF");
    doc.text(sub, left + 10, top + 20, { width: width - 20, lineBreak: false });
    doc.restore();
  }

  band(MAP_BANDS.classify.label, MAP_BANDS.classify.sub, colors.classify || "#3A4F2D", y, bandH);
  y += bandH + gap;

  band(MAP_BANDS.describe.label, MAP_BANDS.describe.sub, colors.describe || "#7B2D3A", y, describeH);
  // Chips sit inside the band, 5pt clear of its bottom edge.
  const chipY = y + describeH - 27;
  const chipGap = 8;
  const chipW = (width - 20 - chipGap * 2) / 3;
  aspects.forEach((label, i) => {
    const cx = left + 10 + i * (chipW + chipGap);
    doc.save();
    doc.roundedRect(cx, chipY, chipW, 22, 3).fill("#FFFFFF");
    doc.fontSize(10).font("Sans-Bold").fillColor(colors.describe || "#7B2D3A");
    doc.text(String(label), cx, chipY + 6, { width: chipW, align: "center", lineBreak: false });
    doc.restore();
  });
  y += describeH + gap;

  band(MAP_BANDS.wrap.label, MAP_BANDS.wrap.sub, colors.wrap || "#6B3A7B", y, bandH);
  y += bandH + gap;

  doc.save();
  doc.roundedRect(left, y, width, 24, 4).fill(colors.facts || "#8A6E1E");
  doc.fontSize(11).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text(FACTS_BAR, left, y + 7, { width, align: "center", lineBreak: false });
  doc.restore();

  return y + 24 + 8;
}

/**
 * A bordered source box: heading strip plus the source paragraph. Used wherever
 * a worksheet asks students to read a short text and take notes from it.
 *
 * @returns {number} y below the box
 */
function addSourceBoxPdf(doc, y, heading, text, o) {
  const opts = o || {};
  const left = opts.left != null ? opts.left : 50;
  const width = opts.width != null ? opts.width : 495.28;
  const color = opts.color || "#3A4F2D";
  const fontSize = opts.fontSize || 11;

  const textW = width - 20;
  const textH = doc.fontSize(fontSize).font("Sans").heightOfString(String(text), { width: textW });
  const boxH = 27 + textH + 10;

  const contentBottom = (opts.pageHeight || 841.89) - (opts.pageMargin || 50) - 20;
  if (y + boxH > contentBottom) {
    doc.addPage();
    y = opts.pageMargin || 50;
  }

  doc.save();
  doc.roundedRect(left, y, width, boxH, 4).lineWidth(1).strokeColor(color).stroke();
  doc.roundedRect(left, y, width, 20, 4).fill(color);
  doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text(String(heading), left + 8, y + 6, { width: textW, lineBreak: false });
  doc.fontSize(fontSize).font("Sans").fillColor("#000000");
  doc.text(String(text), left + 10, y + 27, { width: textW });
  doc.restore();

  return y + boxH + 10;
}

/**
 * Dot-point note box: a labelled box with dot markers and ruled lines, the
 * paper twin of drawNoteCard. Pass `answers` to render the answer-key version.
 *
 * @returns {number} y below the box
 */
function addNoteBoxPdf(doc, y, heading, o) {
  const opts = o || {};
  const left = opts.left != null ? opts.left : 50;
  const width = opts.width != null ? opts.width : 495.28;
  const color = opts.color || "#2D6B4A";
  const lines = opts.lines || 4;
  const spacing = opts.spacing || 24;
  const answers = Array.isArray(opts.answers) ? opts.answers : [];
  const boxH = 22 + lines * spacing + 8;

  const contentBottom = (opts.pageHeight || 841.89) - (opts.pageMargin || 50) - 20;
  if (y + boxH > contentBottom) {
    doc.addPage();
    y = opts.pageMargin || 50;
  }

  doc.save();
  doc.roundedRect(left, y, width, boxH, 4).lineWidth(1).strokeColor(color).stroke();
  doc.roundedRect(left, y, width, 20, 4).fill(color);
  doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text(String(heading), left + 8, y + 6, { width: width - 16, lineBreak: false });
  // A heading that asks the student to name something needs somewhere to
  // write it: a white rule across the rest of the header bar.
  if (opts.headingLine) {
    const startX = left + 12 + doc.widthOfString(String(heading));
    doc.moveTo(startX, y + 15).lineTo(left + width - 10, y + 15)
      .lineWidth(0.9).strokeColor("#FFFFFF").stroke();
  }
  doc.restore();

  const dots = opts.dots !== false;
  for (let i = 0; i < lines; i += 1) {
    const ly = y + 22 + i * spacing + spacing - 8;
    const lineStart = dots ? left + 24 : left + 14;
    doc.save();
    if (dots) doc.circle(left + 14, ly - 4, 2).fill(color);
    doc.moveTo(lineStart, ly).lineTo(left + width - 14, ly)
      .lineWidth(0.9).strokeColor("#000000").stroke();
    doc.restore();
    if (answers[i]) {
      doc.fontSize(10).font("Sans-Italic").fillColor(color);
      doc.text(String(answers[i]), lineStart + 4, ly - 13, { width: width - 46, lineBreak: false });
    }
  }

  return y + boxH + 10;
}

/**
 * Print a model report on paper with a label box beside every part, so the
 * deconstruction task is done ON the text rather than beside it.
 *
 * Pass `labels` to render the answer-key version (the boxes come out filled).
 *
 * @param {PDFDocument} doc
 * @param {number} y
 * @param {object} report  FROG / REEF / THUNDER
 * @param {object} o { labels, color, left, width, labelW, fontSize, showCaption }
 * @returns {number} y below the block
 */
function addReportTextPdf(doc, y, report, o) {
  const opts = o || {};
  const left = opts.left != null ? opts.left : 50;
  const width = opts.width != null ? opts.width : 495.28;
  const color = opts.color || "#3A4F2D";
  const labelW = opts.labelW || 96;
  const fontSize = opts.fontSize || 10.5;
  const labels = Array.isArray(opts.labels) ? opts.labels : [];
  const gapX = 10;
  const textX = left + labelW + gapX;
  const textW = width - labelW - gapX;
  const pageBottom = (opts.pageHeight || 841.89) - (opts.pageMargin || 50) - 20;

  doc.fontSize(15).font("Sans-Bold").fillColor(color);
  doc.text(report.title, left, y, { width });
  y = doc.y + 8;

  const parts = [{ text: report.classify }]
    .concat(report.aspects.map((a) => ({ text: a.text })))
    .concat([{ text: report.wrap }]);

  // Students underline inside this text, so it is set with a line gap: a
  // pencil line under a sentence must not strike through the line below.
  const lineGap = opts.lineGap != null ? opts.lineGap : 4;

  parts.forEach((part, i) => {
    const textH = doc.fontSize(fontSize).font("Sans").heightOfString(part.text, { width: textW, lineGap });
    const rowH = Math.max(textH, 30) + 10;
    if (y + rowH > pageBottom) {
      doc.addPage();
      y = opts.pageMargin || 50;
    }
    doc.save();
    if (labels[i]) {
      doc.roundedRect(left, y, labelW, 22, 3).fill(color);
      doc.fontSize(9.5).font("Sans-Bold").fillColor("#FFFFFF");
      doc.text(String(labels[i]), left, y + 6, { width: labelW, align: "center", lineBreak: false });
    } else {
      doc.roundedRect(left, y, labelW, 22, 3).lineWidth(1).strokeColor(color).stroke();
    }
    doc.restore();
    doc.fontSize(fontSize).font("Sans").fillColor("#000000");
    doc.text(part.text, textX, y, { width: textW, lineGap });
    y += rowH;
  });

  if (opts.showCaption !== false && report.caption) {
    if (y + 40 > pageBottom) {
      doc.addPage();
      y = opts.pageMargin || 50;
    }
    doc.save();
    doc.rect(textX, y, textW, 34).lineWidth(0.8).strokeColor("#9CA3AF").dash(3, { space: 2 }).stroke();
    doc.undash();
    doc.restore();
    doc.fontSize(9.5).font("Sans-Italic").fillColor("#4B5563");
    doc.text(report.caption, textX + 8, y + 10, { width: textW - 16 });
    y += 44;
  }

  return y + 4;
}

/**
 * Checkbox list: one empty square per item. Used for the editing checklist and
 * the publishing rules, so students tick as they go.
 *
 * @returns {number} y below the list
 */
function addChecklistPdf(doc, y, items, o) {
  const opts = o || {};
  const left = opts.left != null ? opts.left : 50;
  const width = opts.width != null ? opts.width : 495.28;
  const color = opts.color || "#3A4F2D";
  const boxSize = opts.boxSize || 13;
  const fontSize = opts.fontSize || 11;
  const rowGap = opts.rowGap || 8;

  items.forEach((item) => {
    const textW = width - boxSize - 14;
    const textH = doc.fontSize(fontSize).font("Sans").heightOfString(String(item), { width: textW });
    const rowH = Math.max(boxSize, textH) + rowGap;
    const contentBottom = (opts.pageHeight || 841.89) - (opts.pageMargin || 50) - 20;
    if (y + rowH > contentBottom) {
      doc.addPage();
      y = opts.pageMargin || 50;
    }
    doc.save();
    doc.roundedRect(left, y, boxSize, boxSize, 2).lineWidth(1).strokeColor(color).stroke();
    doc.restore();
    doc.fontSize(fontSize).font("Sans").fillColor("#000000");
    doc.text(String(item), left + boxSize + 10, y, { width: textW });
    y += rowH;
  });

  return y + 4;
}

module.exports = {
  // anchor
  ANCHOR_PHRASE, ANCHOR_SHORT, MAP_BANDS, GENERIC_ASPECTS, FROG_ASPECTS,
  FACTS_BAR, CLASS_TOPIC,
  // model texts
  FROG, REEF, THUNDER, REPORT_FEATURES,
  // language features
  FACT_OPINION, TENSE_FIXES, TECHNICAL_WORDS, NOUN_GROUPS,
  // research
  RESEARCH_QUESTIONS, DIET_SOURCES, DIET_PARAPHRASE,
  CALL_SOURCES, CALL_PARAPHRASE, NOTE_OPTIONS,
  NOTE_SOURCE_SENTENCE, NOTE_KEY_WORDS, NOTE_MODEL_POINTS,
  // week 10
  EDIT_CHECKLIST, STARS_AND_WISH, VISUAL_TYPES, CAPTION_OPTIONS,
  PUBLISH_RULES, PUBLISH_FORMATS, CLASS_THEMES, TOPIC_TESTS, PLAN_ROWS,
  // sources
  SOURCE_LINE_UNIT, SOURCE_LINE_MODEL,
  // page mockups
  reportPageSpec, messyPageSpec, noVisualPageSpec,
  // chrome + anchors
  REVEAL_Y, REVEAL_H, CONTENT_FLOOR,
  customSlide, drawReportMap, drawReportExtract, drawOptionStack, drawNoteCard,
  drawPlanRows, exitTicketPanel,
  // pdf twins
  addReportMapPdf, addSourceBoxPdf, addNoteBoxPdf, addChecklistPdf,
  addReportTextPdf,
};
