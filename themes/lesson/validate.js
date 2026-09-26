"use strict";

/**
 * Lesson spec validation (docs/lesson-spec.md).
 *
 * A spec is content plus intent; the builder makes every layout decision.
 * Validation is deliberately strict and specific: the model that authors a
 * spec is weaker than the one that wrote this pipeline, so every mistake it
 * can make must be named with the path to the field and the fix.
 *
 * Returns { errors: string[], warnings: string[] }. Any error blocks the build.
 */

const { VALID_SUBJECTS, VALID_YEAR_LEVELS } = require("../factory");
const { SUPPORTED_TYPES } = require("../core/visualSpec");
const { PICTOGRAMS } = require("../core/pictograms");
const { ROUTINES, expandSpec } = require("./practice");
const { TEACHER_SOURCE, BEFORE_LOG_SOURCE, positionOf } = require("./taughtLog");

const BANNED_CHARS = /[–—‘’“”…]/;

const BADGE_COLORS = ["primary", "secondary", "accent", "alert", "success", "assess"];

/** kind -> { required, optional, teaching } */
const KINDS = {
  title:         { required: [], optional: ["notes"], teaching: false },
  overview:      { required: ["lines"], optional: ["title", "notes"], teaching: false },
  resources:     { required: [], optional: ["notes"], teaching: false },
  dailyReview:   { required: ["title", "from"], optional: ["prompts", "visual", "reveal", "notes"], teaching: true, numeracy: true },
  fluency:       { required: ["title", "prompts"], optional: ["reveal", "notes"], teaching: true, numeracy: true },
  launch:        { required: ["title"], optional: ["lines", "visual", "label", "prompt", "badge", "badgeColor", "reveal", "notes"], teaching: true },
  li:            { required: ["learningIntention", "successCriteria"], optional: ["notes"], teaching: true },
  keyWord:       { required: ["word", "meaning"], optional: ["example", "pictogram", "image", "routine", "notes"], teaching: true },
  heroVisual:    { required: ["badge", "title", "visual"], optional: ["label", "prompt", "badgeColor", "reveal", "notes"], teaching: true },
  content:       { required: ["badge", "title", "lines"], optional: ["visual", "badgeColor", "reveal", "notes"], teaching: true },
  workedExample: { required: ["stage", "title", "steps"], optional: ["stageLabel", "visual", "reveal", "notes"], teaching: true, numeracy: true },
  choice:        { required: ["badge", "title", "options"], optional: ["prompt", "answer", "badgeColor", "letters", "notes"], teaching: true },
  cfu:           { required: ["title", "technique", "question"], optional: ["badge", "reveal", "notes"], teaching: true },
  youDo:         { required: ["title", "task"], optional: ["steps", "where", "visual", "visualLabel", "frame", "badge", "badgeColor", "extendedTask", "notes"], teaching: true },
  practice:      { required: ["title", "items", "pivot"], optional: ["badge", "badgeColor", "ask", "routine", "thinkTime", "followUp", "notes"], teaching: true },
  textExtract:   { required: ["badge", "title", "extract"], optional: ["highlights", "source", "prompt", "badgeColor", "reveal", "notes"], teaching: true },
  cycle:         { required: ["title", "steps", "centerLabel"], optional: ["badge", "promptTitle", "promptLines", "reveal", "notes"], teaching: true, science: true },
  process:       { required: ["title", "steps"], optional: ["badge", "promptTitle", "promptLines", "notes"], teaching: true, science: true },
  boardBuild:    { required: ["title", "directive"], optional: ["badge", "promptText", "prefilledHints", "badgeColor", "notes"], teaching: true },
  scenario:      { required: ["title", "scenario", "questions"], optional: ["badge", "notes"], teaching: true, wellbeing: true },
  pairShare:     { required: ["title", "questions"], optional: ["notes"], teaching: true },
  exitTicket:    { required: ["questions"], optional: ["title", "visual", "label", "notes"], teaching: true },
  closing:       { required: ["reflectionPrompt"], optional: ["selfAssessment", "takeaways", "notes"], teaching: false },
};

const RESOURCE_KINDS = ["worksheet", "page", "cards"];

// Visual types the PDF twin layer can draw on paper.
const PDF_VISUAL_TYPES = [
  "tensFrame", "fiveFrame", "doubleTensFrame", "dotCard", "dotCards", "numberTrack", "numberLine",
  "fractionStrips", "array", "groupedCounters", "ppwMat", "hundredGrid", "pictogram", "pictograms",
  "text", "table", "chips",
];

function walkStrings(value, pathLabel, visit) {
  if (typeof value === "string") { visit(value, pathLabel); return; }
  if (Array.isArray(value)) { value.forEach((v, i) => walkStrings(v, `${pathLabel}[${i}]`, visit)); return; }
  if (value && typeof value === "object") {
    Object.keys(value).forEach((k) => walkStrings(value[k], `${pathLabel}.${k}`, visit));
  }
}

function isNonEmptyString(v) { return typeof v === "string" && v.trim().length > 0; }
function toArray(v) { return v == null ? [] : (Array.isArray(v) ? v : [v]); }

function validateVisual(visual, where, errors, opts) {
  const o = opts || {};
  if (visual == null) return;
  if (typeof visual !== "object" || Array.isArray(visual)) {
    errors.push(`${where}: visual must be an object like { "type": "tensFrame", "filled": 7 }`);
    return;
  }
  if (!SUPPORTED_TYPES.includes(visual.type)) {
    errors.push(`${where}.type: "${visual.type}" is not a visual type. Use one of: ${SUPPORTED_TYPES.join(", ")}`);
    return;
  }
  if (o.pdf && !PDF_VISUAL_TYPES.includes(visual.type)) {
    errors.push(`${where}.type: "${visual.type}" cannot be drawn on paper. PDF twins exist for: ${PDF_VISUAL_TYPES.join(", ")}`);
  }
  if (visual.type === "custom") {
    errors.push(`${where}: "custom" visuals need JavaScript and are not allowed in a spec. Choose a built-in type or extend themes/core/visualSpec.js.`);
  }
  if (visual.type === "pictogram" && !PICTOGRAMS[visual.name]) {
    errors.push(`${where}.name: "${visual.name}" is not a pictogram. Run listPictograms() or see the Visual Catalogue sheet.`);
  }
  if (visual.type === "pictograms") {
    (visual.items || []).forEach((it, i) => {
      const name = typeof it === "string" ? it : it && it.name;
      if (!PICTOGRAMS[name]) errors.push(`${where}.items[${i}]: "${name}" is not a pictogram.`);
    });
  }
  if (visual.type === "image" && !isNonEmptyString(visual.path)) {
    errors.push(`${where}.path: an image visual needs a local file path.`);
  }
}

function validateNotes(notes, where, kindDef, errors, warnings) {
  if (notes == null) {
    errors.push(`${where}.notes: every slide needs notes (a one-line string for title/resources/closing, a Glance object for teaching slides).`);
    return;
  }
  if (typeof notes === "string") {
    if (kindDef.teaching) warnings.push(`${where}.notes: one-line notes on a teaching slide. Use { answer, beats, trap, stretch, help, prep, tag } (megaprompt 45-47).`);
    return;
  }
  if (typeof notes !== "object" || Array.isArray(notes)) {
    errors.push(`${where}.notes: must be a string or an object { answer, beats, trap, stretch, help, care, prep, sources, tag }.`);
    return;
  }
  const allowed = ["answer", "beats", "trap", "stretch", "help", "care", "prep", "sources", "tag"];
  Object.keys(notes).forEach((k) => {
    if (!allowed.includes(k)) errors.push(`${where}.notes.${k}: unknown notes field. Allowed: ${allowed.join(", ")}`);
  });
  if (!Array.isArray(notes.beats) || notes.beats.length < 2) {
    errors.push(`${where}.notes.beats: 2 to 5 beats in teaching order (each a string or an array of short lines).`);
  } else if (notes.beats.length > 5) {
    errors.push(`${where}.notes.beats: ${notes.beats.length} beats; the live zone allows at most 5.`);
  }
  if (!isNonEmptyString(notes.tag)) {
    errors.push(`${where}.notes.tag: prep-zone tag required, e.g. "[I Do | Explicit teaching | SC2 | HITS 3]".`);
  }
  if (!notes.prep) warnings.push(`${where}.notes.prep: add one purpose line for the prep zone.`);
}

function validateReveal(reveal, where, errors, warnings, slide) {
  if (reveal == null) return;
  if (typeof reveal !== "object") { errors.push(`${where}.reveal: must be an object { answers: [...] }`); return; }
  const answers = Array.isArray(reveal.answers) ? reveal.answers : (reveal.answers != null ? [reveal.answers] : []);
  if (!answers.length || !answers.every(isNonEmptyString)) {
    errors.push(`${where}.reveal.answers: one or more non-empty strings (joined with a visible separator on the answer bar).`);
  }
  if (reveal.separate) {
    if (!reveal.notes || !isNonEmptyString(reveal.notes.answer)) {
      errors.push(`${where}.reveal.notes.answer: a separate reveal slide must carry its own post-reveal notes { answer, beats, prep } (megaprompt 47).`);
    }
  } else if (reveal.notes) {
    warnings.push(`${where}.reveal.notes: ignored for a click reveal (same slide, same notes). Put the REVEAL beat in the slide's notes, or set reveal.separate: true.`);
  }
  if (slide && ["heroVisual", "textExtract", "launch"].includes(slide.kind) && slide.prompt) {
    errors.push(`${where}: a ${slide.kind} with a reveal cannot also have a prompt bar (both sit at the bottom). Put the question in the title or the notes.`);
  }
}

// Visual types that act as answer menus or pictures rather than the item
// itself, so an exit ticket may repeat them (the feelings row, word chips).
const MENU_VISUAL_TYPES = ["pictogram", "pictograms", "chips"];

// Words an exit ticket routine may share with the rest of the deck.
const ROUTINE_WORDS = /^(point|show|write|tell|name|draw|circle|say|turn|hold|chin|look|read|find|which|what|how|why|who|where|when|is|are|do|does|the|a)$/i;

function canonical(value) {
  if (Array.isArray(value)) return `[${value.map(canonical).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.keys(value).sort().map((k) => `${k}:${canonical(value[k])}`).join(",")}}`;
  }
  return JSON.stringify(value);
}

function normaliseText(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9 ]+/g, " ").replace(/\s+/g, " ").trim();
}

function shingles(s, n) {
  const words = normaliseText(s).split(" ").filter(Boolean);
  const out = new Set();
  for (let i = 0; i + n <= words.length; i += 1) out.add(words.slice(i, i + n).join(" "));
  return out;
}

/** Capitalised words that are not the first word of a sentence: character and place names. */
function midSentenceNames(s) {
  const names = new Set();
  String(s).split(/(?<=[.!?:])\s+|\n/).forEach((sentence) => {
    const words = sentence.trim().split(/\s+/);
    words.slice(1).forEach((w) => {
      const m = w.match(/^([A-Z][a-z]+)\b/);
      if (m) names.add(m[1]);
    });
  });
  return names;
}

function collectStrings(value) {
  const out = [];
  walkStrings(value, "", (s) => out.push(s));
  return out;
}

/**
 * The exit ticket is the lesson's evidence, so it must be a new item the
 * students have not already seen modelled, checked or revealed, and every
 * student's answer must be visible to the teacher (megaprompt 53).
 */
function validateExitTicket(spec, slides, exitIndex, errors) {
  const exit = slides[exitIndex];
  const w = `slides[${exitIndex}] (exitTicket)`;
  const questions = toArray(exit.questions).map(String);
  const exitText = questions.concat(exit.label ? [String(exit.label)] : [],
    exit.visual && exit.visual.type === "text" ? [String(exit.visual.text)] : []);

  // Topic words (lesson title, LI, SC, key words) may repeat; they name the learning, not the item.
  const L = spec.lesson || {};
  const topic = normaliseText([L.title, L.subtitle].concat(
    ...slides.filter((s) => s && (s.kind === "li" || s.kind === "keyWord"))
      .map((s) => collectStrings([s.learningIntention, s.successCriteria, s.word, s.meaning]))
  ).join(" "));
  const topicWords = new Set(topic.split(" "));

  const exitNames = new Set();
  exitText.forEach((q) => (q.match(/\b[A-Z][a-z]+\b/g) || []).forEach((n) => exitNames.add(n)));
  const exitShingles = new Set();
  exitText.forEach((q) => shingles(q, 6).forEach((sh) => exitShingles.add(sh)));
  const exitQuestionKeys = questions.map(normaliseText);
  const exitVisualKey = exit.visual && !MENU_VISUAL_TYPES.includes(exit.visual.type) ? canonical(exit.visual) : null;

  const earlier = slides.slice(0, exitIndex).map((s, i) => ({ s, i }))
    .filter(({ s }) => s && KINDS[s.kind] && KINDS[s.kind].teaching && s.kind !== "li");
  // A worksheet item counts as already seen: students have just done it.
  (spec.resources || []).forEach((r, ri) => {
    if (r && r.kind === "worksheet") (r.items || []).forEach((it, j) => earlier.push({ s: it, label: `resources[${ri}].items[${j}]` }));
  });

  const reported = new Set();
  const report = (msg) => { if (!reported.has(msg)) { reported.add(msg); errors.push(msg); } };
  const fix = "The exit ticket must be a new item (a new text, new numbers or a new context), not one already modelled, checked or revealed, so it shows the skill rather than memory of the answer (megaprompt 53).";

  earlier.forEach(({ s, i, label }) => {
    const where = label || `slides[${i}] (${s.kind})`;
    const strings = collectStrings(s);

    const earlierNames = new Set();
    strings.forEach((str) => midSentenceNames(str).forEach((n) => earlierNames.add(n)));
    exitNames.forEach((n) => {
      if (earlierNames.has(n) && !topicWords.has(n.toLowerCase()) && !ROUTINE_WORDS.test(n)) {
        report(`${w}: reuses "${n}" from ${where}. ${fix}`);
      }
    });

    const faceStrings = collectStrings(Object.assign({}, s, { notes: undefined }));
    if (faceStrings.some((str) => [...shingles(str, 6)].some((sh) => exitShingles.has(sh)))) {
      report(`${w}: repeats six or more words in a row from ${where}. ${fix}`);
    }

    if (!exitVisualKey) {
      const keys = [s.question, s.title, s.prompt, s.task].concat(toArray(s.questions)).filter(isNonEmptyString).map(normaliseText);
      if (exitQuestionKeys.some((q) => keys.includes(q))) report(`${w}: asks the same question as ${where} with no new item. ${fix}`);
    }

    if (exitVisualKey) {
      const visuals = [s.visual, s.answerVisual].concat(toArray(s.options).map((o) => o && o.visual), toArray(s.items).map((it) => it && it.visual)).filter(Boolean);
      if (visuals.some((v) => canonical(v) === exitVisualKey)) report(`${w}.visual: the same visual as ${where}. ${fix}`);
    }
  });

  // Individual, visible evidence: talk to a partner or a choral answer hides who knows it.
  const notes = exit.notes && typeof exit.notes === "object" ? exit.notes : null;
  if (notes) {
    const askBeats = toArray(notes.beats).map((b) => toArray(b).join(" ")).filter((b) => /\bASK:/.test(b));
    if (askBeats.some((b) => /turn and tell|partner|together, on three/i.test(b))) {
      errors.push(`${w}.notes: the exit ASK uses partner talk or a choral answer, which hides who can do it. Collect individual evidence the teacher can see: boards (Write it... Chin it... Show me.), fingers, pointing or paper (megaprompt 53).`);
    }
    if (isNonEmptyString(notes.tag) && !/\bSC2\b/.test(notes.tag)) {
      errors.push(`${w}.notes.tag: the exit ticket assesses SC2; name it in the tag, e.g. "[Exit Ticket | Assessment | SC2 | HITS 8]" (megaprompt 53).`);
    }
  }
}

/** A practice round: 3 to 8 items of one kind, each with its answer (megaprompt 82). */
function validatePractice(slide, w, errors) {
  const items = Array.isArray(slide.items) ? slide.items : [];
  if (items.length < 3 || items.length > 8) errors.push(`${w}.items: 3 to 8 items in a round (got ${items.length}). More practice is another round.`);
  if (slide.routine != null && !ROUTINES[slide.routine]) errors.push(`${w}.routine: use one of ${Object.keys(ROUTINES).join(", ")} (default boards).`);
  if (slide.thinkTime != null && (!Number.isInteger(slide.thinkTime) || slide.thinkTime < 2 || slide.thinkTime > 120)) errors.push(`${w}.thinkTime: seconds, 2 to 120.`);
  if (!isNonEmptyString(slide.pivot)) errors.push(`${w}.pivot: the re-teach move for a weak scan, in a different representation, e.g. "count the empty boxes together on the board frame" (megaprompt 38).`);
  const allowed = ["extract", "source", "highlights", "visual", "label", "text", "answer", "expect", "reveal", "say", "pivot"];
  items.forEach((it, j) => {
    const iw = `${w}.items[${j}]`;
    if (!it || typeof it !== "object") { errors.push(`${iw}: an object { visual | extract | text, answer }`); return; }
    Object.keys(it).forEach((k) => { if (!allowed.includes(k)) errors.push(`${iw}.${k}: unknown field. Allowed: ${allowed.join(", ")}`); });
    const forms = ["extract", "visual", "text"].filter((k) => it[k] != null);
    if (forms.length !== 1) errors.push(`${iw}: exactly one of extract, visual or text (the thing students answer).`);
    if (!isNonEmptyString(it.answer)) errors.push(`${iw}.answer: required, in student words.`);
    if (it.visual) validateVisual(it.visual, `${iw}.visual`, errors);
  });
  const n = slide.notes;
  if (!n || typeof n !== "object" || Array.isArray(n)) {
    errors.push(`${w}.notes: { trap, prep, tag } and optionally say, stretch, help. The round's beats are written for you from the cue scripts.`);
  } else {
    const allowedNotes = ["say", "trap", "stretch", "help", "prep", "tag"];
    Object.keys(n).forEach((k) => { if (!allowedNotes.includes(k)) errors.push(`${w}.notes.${k}: unknown field for a practice round. Allowed: ${allowedNotes.join(", ")}`); });
    if (!isNonEmptyString(n.tag)) errors.push(`${w}.notes.tag: required, e.g. "[We Do | Supported application | SC2 | HITS 3, 7]".`);
    if (!n.trap) errors.push(`${w}.notes.trap: the error this round is most likely to surface, with its fix.`);
  }
}

/**
 * Daily Review retrieves taught content, spaced (megaprompt 22, 77, 83).
 * Each dailyReview slide names its source: a taught-log key, "teacher" when
 * the request named the focus, or "before log" for learning that predates
 * the log. With the log loaded, keys must exist and come earlier, and when
 * older learning is available at least one item must reach back two weeks.
 */
function validateReviewSources(spec, slides, errors, taughtLog) {
  const reviews = slides.map((s, i) => ({ s, i })).filter(({ s }) => s && s.kind === "dailyReview");
  if (!reviews.length) return;
  const L = spec.lesson || {};
  const keyRe = /^\d{4}-T[1-4]-W\d{1,2}-S\d$/;
  const special = [TEACHER_SOURCE, BEFORE_LOG_SOURCE];
  reviews.forEach(({ s, i }) => {
    if (s.from == null) return; // reported as a required field
    if (!special.includes(s.from) && !keyRe.test(String(s.from))) {
      errors.push(`slides[${i}] (dailyReview).from: a taught-log key like "2026-T3-W8-S2", "${TEACHER_SOURCE}" (the request named the focus) or "${BEFORE_LOG_SOURCE}". Run node scripts/taught_log.js <spec> to see what is due.`);
    }
  });
  if (!Array.isArray(taughtLog) || !Number.isInteger(L.term) || !Number.isInteger(L.week)) return;
  const here = positionOf(L);
  const byKey = new Map(taughtLog.map((e) => [e.key, e]));
  reviews.forEach(({ s, i }) => {
    if (!keyRe.test(String(s.from))) return;
    const e = byKey.get(s.from);
    if (!e) errors.push(`slides[${i}] (dailyReview).from: "${s.from}" is not in the taught log for ${L.yearLevel} ${L.subject}. Run node scripts/taught_log.js <spec> for the lessons that are.`);
    else if (e.order >= here.order) errors.push(`slides[${i}] (dailyReview).from: "${s.from}" is not earlier than this lesson (${here.key}). Daily Review retrieves learning already taught.`);
  });
  if (reviews.some(({ s }) => special.includes(s.from))) return;
  const olderAvailable = taughtLog.some((e) => e.weekIndex <= here.weekIndex - 2);
  const reachesBack = reviews.some(({ s }) => { const e = byKey.get(s.from); return e && e.weekIndex <= here.weekIndex - 2; });
  if (olderAvailable && !reachesBack) {
    errors.push(`slides (dailyReview): every item comes from the last week. At least one must retrieve learning from two or more weeks ago so it is spaced (megaprompt 77). Run node scripts/taught_log.js <spec>.`);
  }
}

// Beats that ask every student to respond: the school cue scripts (megaprompt 75a).
const ALL_STUDENT_CUE = /Write it|Chin it|boards up|Show me\b|fingers (up|on cue|at your chest)|everyone points|point on cue|Point\.\.\. now|together, on three|turn and tell|Partner [AB] first|thumbs/i;
const INDEPENDENT_MIN = { foundation: 4, grade1: 6, grade2: 6, grade34: 8, grade56: 8 };

/**
 * Practice volume (megaprompt 82): planned whole-class responses and
 * independent items, counted from the spec with practice rounds expanded.
 */
function practiceCounts(spec) {
  const expanded = expandSpec(spec);
  let responses = 0;
  expanded.slides.forEach((s) => {
    const n = s && s.notes;
    if (!n || typeof n !== "object") return;
    toArray(n.beats).forEach((b) => { if (ALL_STUDENT_CUE.test(toArray(b).join(" "))) responses += 1; });
  });
  let independent = 0;
  (spec.resources || []).forEach((r) => { if (r && r.kind === "worksheet") independent += toArray(r.items).length; });
  (spec.slides || []).forEach((s) => {
    if (s && s.kind === "practice" && /you do/i.test(String(s.badge || ""))) independent += toArray(s.items).length;
  });
  const extended = (spec.slides || []).find((s) => s && s.kind === "youDo" && isNonEmptyString(s.extendedTask));
  return { responses, independent, extended: extended ? extended.extendedTask : null };
}

function validatePracticeVolume(spec, errors) {
  const L = spec.lesson || {};
  const minutes = Number.isInteger(L.minutes) ? L.minutes : 60;
  const { responses, independent, extended } = practiceCounts(spec);
  const minResponses = Math.ceil(minutes / 3);
  if (responses < minResponses) {
    errors.push(`practice: ${responses} planned whole-class responses; a ${minutes}-minute lesson needs at least ${minResponses} (one every three minutes, megaprompt 82). Add a practice round ({ "kind": "practice" }) of quick board items, or a response in the I Do.`);
  }
  const minIndependent = INDEPENDENT_MIN[L.yearLevel];
  if (minIndependent && !extended && independent < minIndependent) {
    errors.push(`practice: ${independent} independent items; ${L.yearLevel} needs at least ${minIndependent} (megaprompt 82). Give the You Do a graded worksheet, or a practice round badged "You Do". For one extended task (a paragraph, a labelled diagram), set youDo.extendedTask to say what it is.`);
  }
}

function validateLessonSpec(spec, opts) {
  const vopts = opts || {};
  const errors = [];
  const warnings = [];
  if (!spec || typeof spec !== "object") return { errors: ["spec must be a JSON object"], warnings };

  const L = spec.lesson || {};
  if (!spec.lesson) errors.push("lesson: required block { subject, yearLevel, week, title, ... }");
  if (!VALID_SUBJECTS.includes(L.subject)) errors.push(`lesson.subject: "${L.subject}" must be one of ${VALID_SUBJECTS.join(", ")}`);
  if (!VALID_YEAR_LEVELS.includes(L.yearLevel)) errors.push(`lesson.yearLevel: "${L.yearLevel}" must be one of ${VALID_YEAR_LEVELS.join(", ")}`);
  if (!isNonEmptyString(L.title)) errors.push("lesson.title: required");
  if (L.week != null && (!Number.isInteger(L.week) || L.week < 1)) errors.push("lesson.week: 1-based integer");
  if (L.session != null && (!Number.isInteger(L.session) || L.session < 1)) errors.push("lesson.session: 1-based integer");
  if (!Number.isInteger(L.term) || L.term < 1 || L.term > 4) errors.push("lesson.term: 1 to 4. With week and session it places the lesson in teaching order for the taught log (megaprompt 83).");
  if (L.week == null) errors.push("lesson.week: required, 1-based. With term and session it places the lesson in teaching order.");
  if (L.year != null && (!Number.isInteger(L.year) || L.year < 2020)) errors.push("lesson.year: a four-digit year, or omit it for the current year.");
  if (L.minutes != null && (!Number.isInteger(L.minutes) || L.minutes < 20 || L.minutes > 120)) errors.push("lesson.minutes: session length, 20 to 120 (default 60).");
  if (L.titleVisual) validateVisual(L.titleVisual, "lesson.titleVisual", errors);

  // Banned characters anywhere: the theme sanitises slide text, but PDFs are
  // not sanitised and notes must be authored clean.
  walkStrings(spec, "spec", (s, where) => {
    if (BANNED_CHARS.test(s)) errors.push(`${where}: contains an em/en dash, smart quote or ellipsis character. Use -, straight quotes and ... (CLAUDE.md PptxGenJS rules).`);
    if (/(^|[^-])--(?!-)/.test(s)) errors.push(`${where}: contains "--". Use a single hyphen.`);
    if (/ {3,}/.test(s)) warnings.push(`${where}: 3+ consecutive spaces (layout by spaces). Use separate fields or chips.`);
  });

  const slides = Array.isArray(spec.slides) ? spec.slides : [];
  if (!slides.length) errors.push("slides: required array");

  const subject = L.subject;
  let liIndex = -1;
  let resourcesIndex = -1;
  let closingIndex = -1;
  let exitIndex = -1;
  let dailyIndex = -1;
  let fluencyIndex = -1;
  let launchLikeBeforeLi = false;
  const preLiAllowed = ["title", "overview", "resources", "dailyReview", "fluency", "launch", "content", "heroVisual", "textExtract", "choice", "boardBuild", "pairShare", "scenario"];

  slides.forEach((slide, i) => {
    const where = `slides[${i}]`;
    if (!slide || typeof slide !== "object") { errors.push(`${where}: must be an object with a "kind"`); return; }
    const def = KINDS[slide.kind];
    if (!def) { errors.push(`${where}.kind: "${slide.kind}" is not a slide kind. Use one of: ${Object.keys(KINDS).join(", ")}`); return; }
    const w = `${where} (${slide.kind})`;
    if (def.numeracy && subject !== "numeracy") errors.push(`${w}: only numeracy decks use ${slide.kind}.`);
    if (def.science && subject !== "science") errors.push(`${w}: ${slide.kind} needs the science theme (lesson.subject: "science").`);
    if (def.wellbeing && subject !== "wellbeing") errors.push(`${w}: ${slide.kind} needs the wellbeing theme.`);

    def.required.forEach((f) => {
      if (slide[f] == null || (typeof slide[f] === "string" && !slide[f].trim()) || (Array.isArray(slide[f]) && !slide[f].length)) {
        errors.push(`${w}.${f}: required`);
      }
    });
    Object.keys(slide).forEach((k) => {
      if (k === "kind") return;
      if (!def.required.includes(k) && !def.optional.includes(k)) {
        errors.push(`${w}.${k}: unknown field. Allowed: ${def.required.concat(def.optional).join(", ")}`);
      }
    });

    if (slide.kind === "practice") validatePractice(slide, w, errors);
    else validateNotes(slide.notes, w, def, errors, warnings);
    validateVisual(slide.visual, `${w}.visual`, errors);
    validateReveal(slide.reveal, w, errors, warnings, slide);
    if (slide.badgeColor && !BADGE_COLORS.includes(slide.badgeColor)) {
      errors.push(`${w}.badgeColor: use one of ${BADGE_COLORS.join(", ")}`);
    }

    switch (slide.kind) {
      case "title": if (i !== 0) errors.push(`${w}: the title slide must be first.`); break;
      case "resources": resourcesIndex = i; break;
      case "dailyReview":
        dailyIndex = i;
        if (!slide.visual && !toArray(slide.prompts).length) errors.push(`${w}: needs prompts and/or a visual.`);
        break;
      case "fluency": fluencyIndex = i; break;
      case "li": {
        liIndex = i;
        if (Array.isArray(slide.learningIntention)) errors.push(`${w}.learningIntention: one plain sentence, not a list.`);
        const sc = Array.isArray(slide.successCriteria) ? slide.successCriteria : [];
        if (sc.length !== 3) errors.push(`${w}.successCriteria: exactly 3 "I can..." statements (got ${sc.length}).`);
        sc.forEach((s, j) => { if (!/^I can\b/i.test(String(s))) warnings.push(`${w}.successCriteria[${j}]: should start with "I can".`); });
        break;
      }
      case "keyWord":
        if (!slide.pictogram && !slide.image) errors.push(`${w}: a word card needs a picture: pictogram (see listPictograms()) or image path (megaprompt 29).`);
        if (slide.pictogram && !PICTOGRAMS[slide.pictogram]) errors.push(`${w}.pictogram: "${slide.pictogram}" is not a pictogram.`);
        break;
      case "content":
      case "launch": {
        const lines = slide.lines == null ? [] : (Array.isArray(slide.lines) ? slide.lines : [slide.lines]);
        if (slide.kind === "content" && lines.length > 6) errors.push(`${w}.lines: ${lines.length} lines; keep to 6 or fewer (split the slide).`);
        if (slide.kind === "launch" && !lines.length && !slide.visual) errors.push(`${w}: a launch needs lines and/or a visual.`);
        break;
      }
      case "choice": {
        const opts = Array.isArray(slide.options) ? slide.options : [];
        if (opts.length < 2 || opts.length > 4) errors.push(`${w}.options: 2 to 4 options.`);
        opts.forEach((o, j) => {
          if (typeof o === "string") return;
          if (!o || (!o.visual && !o.text)) errors.push(`${w}.options[${j}]: needs visual and/or text.`);
          if (o && o.visual) validateVisual(o.visual, `${w}.options[${j}].visual`, errors);
          if (o && o.misconception != null && !isNonEmptyString(o.misconception)) errors.push(`${w}.options[${j}].misconception: a short phrase, or omit it on the correct option.`);
          const unknown = o ? Object.keys(o).filter((k) => !["visual", "text", "misconception"].includes(k)) : [];
          unknown.forEach((k) => errors.push(`${w}.options[${j}].${k}: unknown field. Allowed: visual, text, misconception`));
        });
        if (slide.answer != null && (!Number.isInteger(slide.answer) || slide.answer < 0 || slide.answer >= opts.length)) {
          errors.push(`${w}.answer: 0-based index of the correct option.`);
        } else if (slide.answer != null) {
          // A check is only decision-grade when every wrong answer tells the teacher something (megaprompt 37).
          opts.forEach((o, j) => {
            if (j === slide.answer) return;
            if (!o || typeof o === "string" || !isNonEmptyString(o.misconception)) {
              errors.push(`${w}.options[${j}].misconception: every wrong option on a check must name the misconception it catches, e.g. "counts the counters, not the empty boxes" (megaprompt 37). An option that catches nothing is a wasted choice; replace it with one that does. Use { "text": ..., "misconception": ... } for a text option.`);
            }
          });
        }
        break;
      }
      case "youDo": {
        const steps = slide.steps == null ? [] : (Array.isArray(slide.steps) ? slide.steps : [slide.steps]);
        if (steps.length > 3) errors.push(`${w}.steps: at most 3 (First, Next, Then).`);
        if (slide.extendedTask != null && !isNonEmptyString(slide.extendedTask)) errors.push(`${w}.extendedTask: say what the one extended task is, e.g. "write the introduction paragraph".`);
        break;
      }
      case "workedExample":
        if (![1, 2, 3, 4, 5].includes(slide.stage)) errors.push(`${w}.stage: 1-5 (2 = I Do, 3 = We Do).`);
        break;
      case "cycle":
      case "process": {
        const steps = Array.isArray(slide.steps) ? slide.steps : [];
        if (slide.kind === "cycle" && (steps.length < 3 || steps.length > 4)) errors.push(`${w}.steps: a cycle takes 3 or 4 stages.`);
        if (slide.kind === "process" && (steps.length < 2 || steps.length > 6)) errors.push(`${w}.steps: 2 to 6 stages.`);
        steps.forEach((st, j) => {
          if (!st || (typeof st.label !== "string")) errors.push(`${w}.steps[${j}].label: required (use "" to fade the name and keep the detail as the clue)`);
          else if (!st.label.trim() && !isNonEmptyString(st.detail)) errors.push(`${w}.steps[${j}]: a faded label needs a detail clue`);
          if (st && st.icon && !PICTOGRAMS[st.icon]) errors.push(`${w}.steps[${j}].icon: "${st.icon}" is not a pictogram.`);
        });
        break;
      }
      case "exitTicket": {
        exitIndex = i;
        const qs = Array.isArray(slide.questions) ? slide.questions : [slide.questions];
        if (qs.length < 1 || qs.length > 3) errors.push(`${w}.questions: 1 to 3 prompts.`);
        break;
      }
      case "closing": closingIndex = i; break;
      default: break;
    }

    if (liIndex === -1 && i > 0 && !preLiAllowed.includes(slide.kind)) {
      errors.push(`${w}: ${slide.kind} cannot come before the LI and SC slide (megaprompt 0a item 23).`);
    }
    if (liIndex === -1 && ["launch", "content", "heroVisual", "textExtract", "choice", "boardBuild", "pairShare", "scenario"].includes(slide.kind)) {
      launchLikeBeforeLi = true;
    }
  });

  if (slides.length) {
    if (resourcesIndex === -1) errors.push("slides: no resources slide. Add { \"kind\": \"resources\" } straight after the title (megaprompt 44).");
    else if (resourcesIndex > 2) errors.push(`slides[${resourcesIndex}]: the resources slide belongs immediately after the title (an overview may sit between).`);
    if (liIndex === -1) errors.push("slides: no li slide (Learning Intention and Success Criteria).");
    if (liIndex !== -1 && !launchLikeBeforeLi) errors.push("slides: no launch before the LI slide. Every lesson needs a launch (megaprompt 0a item 17).");
    if (subject === "numeracy") {
      if (dailyIndex === -1) errors.push("slides: a numeracy lesson needs a dailyReview slide (megaprompt 22).");
      if (fluencyIndex === -1) errors.push("slides: a numeracy lesson needs a fluency slide (megaprompt 23).");
      if (dailyIndex !== -1 && fluencyIndex !== -1 && fluencyIndex < dailyIndex) errors.push("slides: fluency must follow dailyReview.");
      if (liIndex !== -1 && fluencyIndex > liIndex) errors.push("slides: fluency must come before the LI slide.");
    }
    if (closingIndex === -1) errors.push("slides: no closing slide.");
    else if (closingIndex !== slides.length - 1) errors.push("slides: the closing slide must be last.");
    if (exitIndex === -1) warnings.push("slides: no exitTicket slide. Most lessons collect evidence before the closing (megaprompt 53).");
    else validateExitTicket(spec, slides, exitIndex, errors);
    validateReviewSources(spec, slides, errors, vopts.taughtLog);
    validatePracticeVolume(spec, errors);
    const kinds = slides.map((s) => s && s.kind);
    if (!kinds.includes("cfu") && !kinds.includes("choice")) warnings.push("slides: no cfu or choice slide. Where is the decision-grade check (megaprompt 36, 76)?");
    const youDoRound = slides.some((sl) => sl && sl.kind === "practice" && /you do/i.test(String(sl.badge || "")));
    if (!kinds.includes("youDo") && !youDoRound) warnings.push("slides: no youDo slide or practice round badged \"You Do\".");
  }

  // Resources
  const resources = Array.isArray(spec.resources) ? spec.resources : [];
  resources.forEach((r, i) => {
    const w = `resources[${i}]`;
    if (!r || typeof r !== "object") { errors.push(`${w}: must be an object`); return; }
    if (!RESOURCE_KINDS.includes(r.kind)) errors.push(`${w}.kind: use one of ${RESOURCE_KINDS.join(", ")}`);
    if (!isNonEmptyString(r.label)) errors.push(`${w}.label: teacher-friendly name, e.g. "Make 10 Worksheet" (the Session N prefix is added for you)`);
    if (!isNonEmptyString(r.description)) warnings.push(`${w}.description: one line saying when it is used.`);
    if (r.kind === "worksheet") {
      const items = Array.isArray(r.items) ? r.items : [];
      if (!items.length) errors.push(`${w}.items: at least one item { prompt, visual, answer, answerLines }`);
      items.forEach((it, j) => {
        if (!it || !isNonEmptyString(it.prompt)) errors.push(`${w}.items[${j}].prompt: required`);
        if (it && it.visual) validateVisual(it.visual, `${w}.items[${j}].visual`, errors, { pdf: true });
        if (it && it.answerVisual) validateVisual(it.answerVisual, `${w}.items[${j}].answerVisual`, errors, { pdf: true });
        if (it && r.answerKey !== false && it.answer == null && !it.answerVisual) {
          warnings.push(`${w}.items[${j}]: no answer given; the answer key will show this item unanswered.`);
        }
      });
    }
    if (r.kind === "page") {
      const blocks = Array.isArray(r.blocks) ? r.blocks : [];
      if (!blocks.length) errors.push(`${w}.blocks: at least one block`);
      blocks.forEach((b, j) => {
        if (b && b.visual) validateVisual(b.visual, `${w}.blocks[${j}].visual`, errors, { pdf: true });
        const allowed = ["heading", "text", "tip", "steps", "visual", "lines", "organiser", "box"];
        if (b && !allowed.some((k) => b[k] != null)) errors.push(`${w}.blocks[${j}]: needs one of ${allowed.join(", ")}`);
      });
    }
    if (r.kind === "cards") {
      const cards = Array.isArray(r.cards) ? r.cards : [];
      if (cards.length < 2) errors.push(`${w}.cards: at least two cards { text, visual }`);
      cards.forEach((c, j) => { if (c && c.visual) validateVisual(c.visual, `${w}.cards[${j}].visual`, errors, { pdf: true }); });
    }
  });
  if (resources.length > 2) warnings.push("resources: more than two printed resources. Default is zero or one (megaprompt 0a item 7).");

  return { errors, warnings };
}

module.exports = { validateLessonSpec, practiceCounts, KINDS, RESOURCE_KINDS, PDF_VISUAL_TYPES, BADGE_COLORS };
