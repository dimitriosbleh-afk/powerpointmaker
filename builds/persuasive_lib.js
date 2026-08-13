"use strict";

// Shared data + visual anchor for the Year 5/6 Persuasive Writing fortnight
// (Term 3 Weeks 6 and 7). Every session imports from here so the unit anchor,
// the word banks and the source excerpts stay byte-identical across nine decks
// (megaprompt section 79, unit anchor consistency).
//
// UNIT ANCHOR (locked, megaprompt section 5b - the SCHOOL's own wording from
// its Persuasive Structure chart, never swapped for the textbook terms
// "introduction / body / conclusion"):
//
//   "Bold Beginning. Mighty Middle. Excellent Ending. Voice all the way through."
//
// Week 7 extends the same anchor rather than replacing it:
//
//   "Same belief. Same reasons. Different voice."
//
// SOURCES, all supplied by the teacher (Stacey Gannon, 6 Aug 2026):
//  - Persuasive Structure anchor chart + "Why Donkey Should Leave Shrek's
//    Swamp", a 2019 DCE Grade 6 student exemplar (screenshot in the planner).
//  - Persuasive word banks (emotive language, rhetorical question starters,
//    bossy verbs, exaggeration and promises, precise words, modality).
//  - ARC "We must do more to stop cyberbullying", Persuasive writing Level 5-6.
//
// Local image assets live in assets/persuasive/ and are GITIGNORED: the
// exemplar is a named cohort's student work collaged over copyrighted film
// stills, and this repository has a public remote.

const path = require("path");

const ASSET_DIR = path.join(__dirname, "..", "assets", "persuasive");

const IMAGES = {
  structureChart: path.join(ASSET_DIR, "dce_persuasive_structure_chart.png"),
  shrekExemplar: path.join(ASSET_DIR, "dce_shrek_exemplar.png"),
  sourceFull: path.join(ASSET_DIR, "dce_persuasive_source_full.png"),
};

// ─── The anchor ──────────────────────────────────────────────────────────────

const ANCHOR_PHRASE = "Bold Beginning. Mighty Middle. Excellent Ending. Voice all the way through.";
const ANCHOR_PHRASE_W7 = "Same belief. Same reasons. Different voice.";

// Wording taken from the school's own chart, including the bracketed glosses.
const STRUCTURE = [
  { key: "BB", label: "Bold Beginning", sub: "Introduce the topic and your perspective" },
  { key: "MM", label: "Mighty Middle", sub: "Reasons backed up with evidence" },
  { key: "EE", label: "Excellent Ending", sub: "Finish strong, one final time to convince" },
];

// The four Voice tools taught across Week 6, in the order they are taught.
const VOICE_TOOLS = [
  { key: "emotive", label: "Emotive words", cue: "make them FEEL it" },
  { key: "precise", label: "Precise words", cue: "swap good for outstanding" },
  { key: "question", label: "Rhetorical question", cue: "make them answer in their head" },
  { key: "modality", label: "High modality", cue: "must, never, definitely" },
];

// The teacher-modelled topic for the whole of Week 7. Chosen from the
// teacher's own funny-topic list in Week 6 Session 1 so the modelling has
// continuity, and because it works genuinely for all three audiences.
const MODEL_TOPIC = "Every classroom should have a scrap paper bin";

// ─── Word banks (supplied by the teacher, verbatim except where noted) ───────

// "startling" appeared twice in the supplied list; the duplicate is dropped.
const EMOTIVE_WORDS = [
  "agony", "cruel", "honest", "panic", "courageous", "unbelievable",
  "agree", "tremendous", "damaging", "inferior", "repulsive", "demand",
  "incredible", "corrupt", "urge", "disaster", "innocent", "riot",
  "unmissable", "outrage", "harsh", "appalling", "disgust", "magic",
  "secret", "astonishing", "outrageous", "forbidden", "vulnerable", "beg",
  "dreadful", "magnificent", "tragic", "threat", "extraordinary", "shame",
  "wicked", "miracle", "startling", "phenomenal", "terrified", "freedom",
  "must", "you",
];

const RHETORICAL_STARTERS = [
  "Do you think that...?",
  "Don't you think that...?",
  "Isn't it time...?",
  "Have you ever thought about...?",
  "Why not...?",
  "Need a...?",
  "Fancy...?",
  "Need to...?",
  "Fed up with...?",
  "Bored by...?",
  "Haven't you always wanted...?",
  "Worried about...?",
];

const BOSSY_VERBS = [
  "Don't...",
  "Go on...",
  "Try a...",
  "Enjoy the...",
  "Imagine...",
  "Buy...",
];

// "You'll never to ... Again" in the supplied list is a typo; corrected here.
const EXAGGERATION_PROMISES = [
  "You will be...",
  "It will...",
  "You'll never... again",
  "Just think how...",
  "Now you can...",
  "For the rest of your life...",
  "You'll always...",
];

// "Mavellous" in the supplied list is a typo; corrected to "Marvellous".
const PRECISE_WORDS = [
  { flat: "HAPPY", better: ["Cheerful", "Delighted", "Pleased", "Content", "Amused", "Thrilled"] },
  { flat: "SAD", better: ["Miserable", "Frustrated", "Distraught", "On edge", "Gloomy", "Devastated"] },
  { flat: "NICE", better: ["Courteous", "Likeable", "Gracious", "Considerate", "Approachable", "Charming"] },
  { flat: "GOOD", better: ["Excellent", "Amazing", "Sensational", "Marvellous", "Terrific", "Splendid", "Outstanding"] },
  { flat: "BAD", better: ["Awful", "Appalling", "Mean", "Dreadful", "Wicked", "Outrageous", "Nasty"] },
  { flat: "PRETTY", better: ["Gorgeous", "Stunning", "Exquisite", "Dazzling", "Appealing", "Delicious", "Mesmerising", "Adorable", "Beautiful", "Elegant", "Perfect"] },
  { flat: "SCARED", better: ["Tense", "Concerned", "Unnerved", "Petrified", "Alarmed"] },
  { flat: "ANGRY", better: ["Irate", "Enraged", "Mad", "Cross", "Infuriated"] },
  { flat: "FUNNY", better: ["Amusing", "Hilarious"] },
  { flat: "SHOCKED", better: ["Taken aback", "Lost for words", "Outraged", "Astounded", "Astonished", "Speechless", "Stunned", "Appalled", "Unbelievable"] },
  { flat: "INTERESTING", better: ["Exciting", "Captivating", "Engaging", "Thrilling", "Fascinating", "Mind-blowing", "Invigorating", "Electrifying"] },
  { flat: "USEFUL", better: ["Beneficial", "Rewarding", "Fulfilling", "Of great use", "Efficient"] },
  { flat: "BIG", better: ["Massive", "Gigantic", "Enormous", "Large", "Colossal", "Immense", "Tremendous"] },
  { flat: "SMALL", better: ["Petite", "Minute", "Mini", "Teeny", "Miniature", "Microscopic", "Miniscule"] },
  { flat: "OTHER", better: ["Trustworthy", "Helpful", "Safe", "Honest", "Loyal", "Powerful", "Popular"] },
];

// "must" appeared twice in the supplied High column; the duplicate is dropped.
// "would" appears in BOTH the supplied Medium and High columns and is left as
// supplied - that is a teacher content decision, not a transcription slip.
const MODALITY = {
  low: ["could", "couldn't", "doubtful", "may", "maybe", "might", "occasionally",
    "perhaps", "possibly", "potentially", "sometimes", "rarely"],
  medium: ["should", "shouldn't", "would", "wouldn't", "apparently", "frequently",
    "probably", "likely", "often"],
  high: ["would", "is", "must", "absolutely", "always", "certainly", "clearly",
    "definitely", "mustn't", "never", "obviously"],
};

// The teacher's own funny persuasive topics, quoted from the planner.
const FUNNY_TOPICS = [
  "Why we must have a scrap paper bin in every classroom",
  "Why a broccoli should be elected the best vegetable",
  "Why a snake should be scared of you, and not you of them",
];

// ─── Source excerpts (exact, megaprompt section 5a) ──────────────────────────

// ARC, "We must do more to stop cyberbullying", Persuasive writing Level 5-6.
const ARC = {
  title: "We must do more to stop cyberbullying",
  hook: "Imagine opening your device and finding hurtful messages waiting for you.",
  rhetorical: "Should any student be made to feel unsafe every time they go online?",
  position: "Schools, families and online communities must do more to prevent cyberbullying because it harms wellbeing, affects learning and can make students feel unsafe.",
  body1: "Cyberbullying can seriously affect a student's wellbeing because hurtful comments, rumours and exclusion can cause sadness, stress and anxiety. According to surveys, many young people report experiencing some form of online bullying during their school years. This means that students will be worried and on edge when checking messages, never being completely relaxed. No child should have to feel afraid every time they check their messages. Therefore, it is important that adults and students work together to create safer online spaces.",
  body1Reason: "Cyberbullying can seriously affect a student's wellbeing because hurtful comments, rumours and exclusion can cause sadness, stress and anxiety.",
  body1Evidence: "According to surveys, many young people report experiencing some form of online bullying during their school years.",
  body1Explain: "This means that students will be worried and on edge when checking messages, never being completely relaxed.",
  body1Link: "Therefore, it is important that adults and students work together to create safer online spaces.",
  callToAction: "Stand up against cyberbullying and help create a kinder online world today.",
  repetition: "Everyone deserves to feel safe. Everyone deserves to feel respected. Everyone deserves to feel included.",
  counterOpening: "Some people believe that cyberbullying is not a serious problem because it can be avoided.",
  counterTurn: "However, cyberbullying is not always easy to escape.",
};

const ARC_SOURCE_LINE =
  "ARC, Persuasive writing Levels 5-6, sample text \"We must do more to stop cyberbullying\", supplied by the teacher.";
const SCHOOL_SOURCE_LINE =
  "DCE Persuasive Structure chart and 2019 DCE Grade 6 exemplar, supplied by the teacher.";

// The Mighty Middle sentence frame, held identical across both weeks.
const MM_FRAME = [
  { step: "Reason", frame: "[Position] because [reason].", color: "PRIMARY" },
  { step: "Evidence", frame: "According to... / For example...", color: "SECONDARY" },
  { step: "Explain", frame: "This means that...", color: "ASSESS" },
  { step: "Link", frame: "Therefore, ...", color: "SUCCESS" },
];

// ─── The anchor, drawn ───────────────────────────────────────────────────────

/**
 * Draw the unit anchor: the school's three-part persuasive structure as
 * stacked bands with the Voice bar running underneath.
 *
 * This is the ONE representation for the fortnight. Never hand-draw a variant
 * of it in a session script; extend this function instead.
 *
 * @param {object} slide  PptxGenJS slide
 * @param {object} T      theme object from createTheme
 * @param {object} o      { x, y, w, h, focus, showVoice, voiceText, compact, parts }
 *   focus: "BB" | "MM" | "EE" | null  - the band being taught today; the other
 *          bands render as pale outline cards so the focus reads at a glance.
 *   parts: subset of ["BB","MM","EE","voice"] to draw. Positions are always
 *          computed for the FULL layout, so calling this repeatedly with
 *          different parts inside clickBuild() steps builds the anchor up one
 *          band per click without anything shifting.
 * @returns {number} the y coordinate just below the anchor
 */
function drawStructureStrip(slide, T, o) {
  const { C, FONT_H, FONT_B } = T;
  const x = Number(o.x);
  const y = Number(o.y);
  const w = Number(o.w);
  const h = Number(o.h);
  const focus = o.focus || null;
  const showVoice = o.showVoice !== false;
  const compact = Boolean(o.compact);
  const parts = Array.isArray(o.parts) ? o.parts : null;
  const wants = (key) => !parts || parts.indexOf(key) !== -1;

  const gap = compact ? 0.08 : 0.10;
  const voiceH = showVoice ? (compact ? 0.42 : 0.52) : 0;
  const bandsH = h - voiceH - (showVoice ? gap : 0) - gap * 2;
  const bandH = bandsH / 3;

  const bandFills = [C.PRIMARY, C.SECONDARY, C.ASSESS];
  const labelSize = compact ? 15 : Math.min(20, Math.max(14, bandH * 22));
  const subSize = compact ? 10.5 : Math.min(13, Math.max(9.5, bandH * 15));

  STRUCTURE.forEach((band, i) => {
    if (!wants(band.key)) return;
    const by = y + i * (bandH + gap);
    const isFocus = !focus || focus === band.key;
    const fill = isFocus ? bandFills[i] : C.WHITE;
    const textColor = isFocus ? C.WHITE : C.MUTED;

    slide.addShape("roundRect", {
      x, y: by, w, h: bandH, rectRadius: 0.08,
      fill: { color: fill },
      line: isFocus ? null : { color: C.MUTED, width: 1 },
    });
    slide.addText([
      {
        text: band.label,
        options: { fontSize: labelSize, bold: true, color: textColor, fontFace: FONT_H, breakLine: true },
      },
      {
        text: band.sub,
        options: { fontSize: subSize, color: textColor, fontFace: FONT_B },
      },
    ], {
      x: x + 0.20, y: by + 0.04, w: w - 0.40, h: bandH - 0.08,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  });

  if (showVoice && wants("voice")) {
    const vy = y + 3 * bandH + 2 * gap + gap;
    slide.addShape("roundRect", {
      x, y: vy, w, h: voiceH, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    });
    slide.addText(String(o.voiceText || "VOICE - all the way through"), {
      x: x + 0.20, y: vy, w: w - 0.40, h: voiceH,
      fontSize: compact ? 12.5 : 14, bold: true, color: C.WHITE, fontFace: FONT_B,
      align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
    return vy + voiceH;
  }

  return y + 3 * bandH + 2 * gap;
}

/**
 * Draw the four Voice tools as a labelled chip stack (vertical) or row.
 *
 * @param {object} slide
 * @param {object} T
 * @param {object} o  { x, y, w, h, tools, horizontal, focus }
 * @returns {number} y below the block
 */
function drawVoiceTools(slide, T, o) {
  const { C, FONT_B, FONT_H } = T;
  const x = Number(o.x);
  const y = Number(o.y);
  const w = Number(o.w);
  const h = Number(o.h);
  const tools = o.tools || VOICE_TOOLS;
  const focus = o.focus || null;
  const fills = [C.PRIMARY, C.SECONDARY, C.ASSESS, C.SUCCESS];
  const gap = 0.09;
  const n = tools.length;

  if (o.horizontal) {
    const cw = (w - gap * (n - 1)) / n;
    tools.forEach((tool, i) => {
      const cx = x + i * (cw + gap);
      const isFocus = !focus || focus === tool.key;
      slide.addShape("roundRect", {
        x: cx, y, w: cw, h, rectRadius: 0.08,
        fill: { color: isFocus ? fills[i % fills.length] : C.WHITE },
        line: isFocus ? null : { color: C.MUTED, width: 1 },
      });
      slide.addText([
        {
          text: tool.label,
          options: { fontSize: 13, bold: true, color: isFocus ? C.WHITE : C.MUTED, fontFace: FONT_H, breakLine: true },
        },
        {
          text: tool.cue,
          options: { fontSize: 10.5, color: isFocus ? C.WHITE : C.MUTED, fontFace: FONT_B, italic: true },
        },
      ], {
        x: cx + 0.08, y: y + 0.04, w: cw - 0.16, h: h - 0.08,
        align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });
    return y + h;
  }

  const rh = (h - gap * (n - 1)) / n;
  tools.forEach((tool, i) => {
    const ry = y + i * (rh + gap);
    const isFocus = !focus || focus === tool.key;
    slide.addShape("roundRect", {
      x, y: ry, w, h: rh, rectRadius: 0.07,
      fill: { color: isFocus ? fills[i % fills.length] : C.WHITE },
      line: isFocus ? null : { color: C.MUTED, width: 1 },
    });
    slide.addText([
      {
        text: tool.label + "  ",
        options: { fontSize: 13, bold: true, color: isFocus ? C.WHITE : C.MUTED, fontFace: FONT_H },
      },
      {
        text: tool.cue,
        options: { fontSize: 10.5, color: isFocus ? C.WHITE : C.MUTED, fontFace: FONT_B, italic: true },
      },
    ], {
      x: x + 0.16, y: ry, w: w - 0.32, h: rh,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  });
  return y + h;
}

/**
 * Draw a word-bank panel: a heading strip and wrapped word chips.
 * Never lays words out with runs of spaces (slide-text hygiene gate 3).
 *
 * @returns {number} y below the panel
 */
function drawWordChips(slide, T, o) {
  const { C, FONT_B } = T;
  const x = Number(o.x);
  const y = Number(o.y);
  const w = Number(o.w);
  const words = (o.words || []).map(String);
  const perRow = Number(o.perRow) || 4;
  const chipH = Number(o.chipH) || 0.34;
  const gap = Number(o.gap) || 0.08;
  const fill = o.fill || C.BG_LIGHT;
  const textColor = o.textColor || C.CHARCOAL;
  const fontSize = Number(o.fontSize) || 12.5;

  const cw = (w - gap * (perRow - 1)) / perRow;
  words.forEach((word, i) => {
    const row = Math.floor(i / perRow);
    const col = i % perRow;
    const cx = x + col * (cw + gap);
    const cy = y + row * (chipH + gap);
    slide.addShape("roundRect", {
      x: cx, y: cy, w: cw, h: chipH, rectRadius: 0.06,
      fill: { color: fill },
      line: o.line === false ? null : { color: C.MUTED, width: 0.75 },
    });
    slide.addText(word, {
      x: cx + 0.04, y: cy, w: cw - 0.08, h: chipH,
      fontSize, fontFace: FONT_B, color: textColor, bold: Boolean(o.bold),
      align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });
  });

  const rows = Math.ceil(words.length / perRow);
  return y + rows * chipH + Math.max(rows - 1, 0) * gap;
}

/**
 * Word-bank block for a PDF page: heading + comma-free chip grid drawn with
 * pdfkit primitives, so the printed bank matches the slide bank.
 *
 * @param {PDFDocument} doc
 * @param {number} y
 * @param {string} heading
 * @param {string[]} words
 * @param {object} o  { color, perRow, chipH, fontSize, pageLeft, pageWidth }
 * @returns {number} y below the block
 */
function addWordChipsPdf(doc, y, heading, words, o) {
  const opts = o || {};
  const left = opts.pageLeft != null ? opts.pageLeft : 40;
  const width = opts.pageWidth != null ? opts.pageWidth : 515;
  const color = opts.color || "#2C1654";
  const perRow = opts.perRow || 5;
  const chipH = opts.chipH || 20;
  const gap = 5;
  const fontSize = opts.fontSize || 9.5;

  // Keep a heading and its chips together: break the page before the block
  // rather than stranding the heading at the foot of a page (section 42).
  const rowsNeeded = Math.ceil(words.length / perRow);
  const blockH = 23 + rowsNeeded * chipH + Math.max(rowsNeeded - 1, 0) * gap + 10;
  const contentBottom = (opts.pageHeight || 842) - (opts.pageMargin || 40) - 20;
  if (y + blockH > contentBottom) {
    doc.addPage();
    y = opts.pageMargin || 40;
  }

  doc.save();
  doc.roundedRect(left, y, width, 18, 3).fill(color);
  doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text(heading, left + 8, y + 5, { width: width - 16 });
  doc.restore();
  y += 23;

  const cw = (width - gap * (perRow - 1)) / perRow;
  words.forEach((word, i) => {
    const row = Math.floor(i / perRow);
    const col = i % perRow;
    const cx = left + col * (cw + gap);
    const cy = y + row * (chipH + gap);
    doc.save();
    doc.roundedRect(cx, cy, cw, chipH, 3)
      .lineWidth(0.7).strokeColor("#9CA3AF").stroke();
    doc.fontSize(fontSize).font("Sans").fillColor("#000000");
    doc.text(String(word), cx + 2, cy + (chipH - fontSize) / 2 - 1,
      { width: cw - 4, align: "center", lineBreak: false });
    doc.restore();
  });

  const rows = Math.ceil(words.length / perRow);
  return y + rows * chipH + Math.max(rows - 1, 0) * gap + 10;
}

// ─── Shared slide chrome ─────────────────────────────────────────────────────

// Every click-revealed answer bar in this unit sits at the same place, so the
// content floor above it is the same on every slide. Anything drawn below
// CONTENT_FLOOR would be covered by the reveal - the layout diagnostics only
// check withReveal duplicates, never clickBuild steps, so this is on the author.
const REVEAL_Y = 4.28;
const REVEAL_H = 0.78;
const CONTENT_FLOOR = 4.13;

/**
 * Slide chrome for a custom layout: top bar, badge, title. Returns the slide.
 */
function customSlide(pres, T, badgeText, badgeColor, title, opts) {
  const o = opts || {};
  const s = pres.addSlide();
  T.addTopBar(s, badgeColor);
  T.addBadge(s, badgeText, { color: badgeColor, w: o.badgeW || 1.9 });
  T.addTitle(s, title, o.titleOpts);
  return s;
}

/**
 * Exit ticket built as a hero panel rather than the default builder card.
 *
 * The shared exitTicketSlide caps a single prompt card at 1.55" for Y3-6 and
 * offsets it high, which leaves a dead bottom third on a short prompt
 * (megaprompt 15h). Rather than fork the shared builder used by 250+ decks,
 * this unit uses one consistent hero panel across all nine sessions: the task
 * is the largest thing on the slide and the space is filled on purpose.
 *
 * The SC target stays in the notes only, never on the face (section 0a item 18).
 */
function exitTicketPanel(pres, T, cfg) {
  const { C, FONT_H, FONT_B } = T;
  const s = pres.addSlide();
  const strip = C.ASSESS || C.ALERT;

  s.background = { color: C.BG_CARD };
  s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: strip } });
  T.addBadge(s, "Exit Ticket", { color: strip, x: 0.5, y: 0.2, w: 1.85, h: 0.38 });
  s.addText(cfg.title || "Show what you know", {
    x: 0.5, y: 0.68, w: 9, h: 0.60,
    fontSize: 30, fontFace: FONT_H, color: strip, bold: true, margin: 0,
  });

  let y = 1.40;
  if (cfg.topic) {
    T.addTextOnShape(s, String(cfg.topic), {
      x: 0.5, y, w: 9, h: 0.72, rectRadius: 0.08,
      fill: { color: strip },
    }, {
      fontSize: 20, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0.1,
    });
    y += 0.88;
  }

  const cueH = cfg.cue ? 0.60 : 0;
  const taskH = (cfg.cue ? 4.28 : 4.88) - y;
  T.addCard(s, 0.5, y, 9, taskH, { strip, fill: C.WHITE });
  s.addText(String(cfg.task || ""), {
    x: 0.85, y: y + 0.10, w: 8.3, h: taskH - 0.20,
    fontSize: cfg.taskSize || 30, fontFace: FONT_B, color: C.CHARCOAL,
    align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
  });

  if (cfg.cue) {
    T.addTextOnShape(s, String(cfg.cue), {
      x: 0.5, y: 4.28, w: 9, h: cueH, rectRadius: 0.08,
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

module.exports = {
  REVEAL_Y,
  REVEAL_H,
  CONTENT_FLOOR,
  customSlide,
  exitTicketPanel,
  IMAGES,
  ANCHOR_PHRASE,
  ANCHOR_PHRASE_W7,
  STRUCTURE,
  VOICE_TOOLS,
  MODEL_TOPIC,
  EMOTIVE_WORDS,
  RHETORICAL_STARTERS,
  BOSSY_VERBS,
  EXAGGERATION_PROMISES,
  PRECISE_WORDS,
  MODALITY,
  FUNNY_TOPICS,
  ARC,
  ARC_SOURCE_LINE,
  SCHOOL_SOURCE_LINE,
  MM_FRAME,
  drawStructureStrip,
  drawVoiceTools,
  drawWordChips,
  addWordChipsPdf,
};
