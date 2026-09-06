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

const BANNED_CHARS = /[–—‘’“”…]/;

const BADGE_COLORS = ["primary", "secondary", "accent", "alert", "success", "assess"];

/** kind -> { required, optional, teaching } */
const KINDS = {
  title:         { required: [], optional: ["notes"], teaching: false },
  overview:      { required: ["lines"], optional: ["title", "notes"], teaching: false },
  resources:     { required: [], optional: ["notes"], teaching: false },
  dailyReview:   { required: ["title"], optional: ["prompts", "visual", "reveal", "notes"], teaching: true, numeracy: true },
  fluency:       { required: ["title", "prompts"], optional: ["reveal", "notes"], teaching: true, numeracy: true },
  launch:        { required: ["title"], optional: ["lines", "visual", "label", "prompt", "badge", "badgeColor", "reveal", "notes"], teaching: true },
  li:            { required: ["learningIntention", "successCriteria"], optional: ["notes"], teaching: true },
  keyWord:       { required: ["word", "meaning"], optional: ["example", "pictogram", "image", "routine", "notes"], teaching: true },
  heroVisual:    { required: ["badge", "title", "visual"], optional: ["label", "prompt", "badgeColor", "reveal", "notes"], teaching: true },
  content:       { required: ["badge", "title", "lines"], optional: ["visual", "badgeColor", "reveal", "notes"], teaching: true },
  workedExample: { required: ["stage", "title", "steps"], optional: ["stageLabel", "visual", "reveal", "notes"], teaching: true, numeracy: true },
  choice:        { required: ["badge", "title", "options"], optional: ["prompt", "answer", "badgeColor", "letters", "notes"], teaching: true },
  cfu:           { required: ["title", "technique", "question"], optional: ["badge", "reveal", "notes"], teaching: true },
  youDo:         { required: ["title", "task"], optional: ["steps", "where", "visual", "visualLabel", "frame", "badge", "badgeColor", "notes"], teaching: true },
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

function validateLessonSpec(spec) {
  const errors = [];
  const warnings = [];
  if (!spec || typeof spec !== "object") return { errors: ["spec must be a JSON object"], warnings };

  const L = spec.lesson || {};
  if (!spec.lesson) errors.push("lesson: required block { subject, yearLevel, week, title, ... }");
  if (!VALID_SUBJECTS.includes(L.subject)) errors.push(`lesson.subject: "${L.subject}" must be one of ${VALID_SUBJECTS.join(", ")}`);
  if (!VALID_YEAR_LEVELS.includes(L.yearLevel)) errors.push(`lesson.yearLevel: "${L.yearLevel}" must be one of ${VALID_YEAR_LEVELS.join(", ")}`);
  if (!isNonEmptyString(L.title)) errors.push("lesson.title: required");
  if (L.week == null && L.variant == null) warnings.push("lesson.week: not set; variant 0 will be used. Set the week so the unit keeps one palette.");
  if (L.week != null && (!Number.isInteger(L.week) || L.week < 1)) errors.push("lesson.week: 1-based integer");
  if (L.session != null && (!Number.isInteger(L.session) || L.session < 1)) errors.push("lesson.session: 1-based integer");
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

    validateNotes(slide.notes, w, def, errors, warnings);
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
        });
        if (slide.answer != null && (!Number.isInteger(slide.answer) || slide.answer < 0 || slide.answer >= opts.length)) {
          errors.push(`${w}.answer: 0-based index of the correct option.`);
        }
        break;
      }
      case "youDo": {
        const steps = slide.steps == null ? [] : (Array.isArray(slide.steps) ? slide.steps : [slide.steps]);
        if (steps.length > 3) errors.push(`${w}.steps: at most 3 (First, Next, Then).`);
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
    const kinds = slides.map((s) => s && s.kind);
    if (!kinds.includes("cfu") && !kinds.includes("choice")) warnings.push("slides: no cfu or choice slide. Where is the decision-grade check (megaprompt 36, 76)?");
    if (!kinds.includes("youDo")) warnings.push("slides: no youDo slide.");
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

module.exports = { validateLessonSpec, KINDS, RESOURCE_KINDS, PDF_VISUAL_TYPES, BADGE_COLORS };
