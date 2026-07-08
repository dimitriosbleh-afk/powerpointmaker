"use strict";

const { SAFE_BOTTOM } = require("./layout");
const { inferElementType, getElementBounds, getOverlap } = require("./diagnostics");

/**
 * Estimate where a text element's RENDERED text actually ends vertically.
 * Text boxes are routinely taller than their content (padding boxes that run
 * to SAFE_BOTTOM), so comparing box bounds alone gives false positives.
 * Heuristic mirrors the project's text-height rules of thumb.
 */
function estimateRenderedTextBottom(obj) {
  const bounds = getElementBounds(obj);
  if (!bounds) return null;
  const options = obj.options || {};

  // Collect paragraphs: a plain string, or run objects (breakLine = new para)
  const paragraphs = [];
  if (typeof obj.text === "string") {
    paragraphs.push({ text: obj.text, fontSize: options.fontSize });
  } else if (Array.isArray(obj.text)) {
    let current = "";
    let currentSize = null;
    obj.text.forEach((run) => {
      const runText = run && run.text != null ? String(run.text) : "";
      const runOpts = (run && run.options) || {};
      current += runText;
      if (currentSize == null && runOpts.fontSize) currentSize = runOpts.fontSize;
      if (runOpts.breakLine) {
        paragraphs.push({ text: current, fontSize: currentSize || options.fontSize });
        current = "";
        currentSize = null;
      }
    });
    if (current) paragraphs.push({ text: current, fontSize: currentSize || options.fontSize });
  } else {
    return null;
  }

  let textH = 0;
  paragraphs.forEach((para) => {
    const fontSize = Number(para.fontSize) || Number(options.fontSize) || 14;
    // Empirical: ~0.0057" per point per character for the body/heading fonts
    // (an 8.5" box at 44pt renders ~34 chars per line - measured, not guessed)
    const charsPerLine = Math.max(8, Math.floor(bounds.w / (fontSize * 0.0057)));
    const lines = Math.max(1, Math.ceil(String(para.text).trim().length / charsPerLine));
    textH += lines * (fontSize * 0.0185 + 0.02);
  });

  const valign = options.valign || "top";
  if (valign === "middle") {
    return bounds.y + bounds.h / 2 + Math.min(textH, bounds.h) / 2;
  }
  if (valign === "bottom") {
    return bounds.y + bounds.h;
  }
  return bounds.y + Math.min(textH + 0.06, bounds.h);
}

/**
 * Warn when reveal-layer elements (added by revealFn) sit on top of the base
 * slide's rendered text. The build gate treats these WARNs as failures, so a
 * reveal bar can no longer silently hide the last line of a question - the
 * defect previously only caught by visual QA.
 */
function warnIfRevealCoversBaseText(slide, preCount) {
  const objects = Array.isArray(slide && slide._slideObjects) ? slide._slideObjects : [];
  const warnings = [];

  for (let j = preCount; j < objects.length; j += 1) {
    const newObj = objects[j];
    const newType = inferElementType(newObj);
    if (newType !== "text" && newType !== "shape") continue;
    const newBounds = getElementBounds(newObj);
    if (!newBounds) continue;

    for (let i = 0; i < preCount; i += 1) {
      const baseObj = objects[i];
      if (inferElementType(baseObj) !== "text") continue;
      const baseBounds = getElementBounds(baseObj);
      if (!baseBounds) continue;
      // Footer-zone base text is not lesson content
      if (baseBounds.y >= SAFE_BOTTOM + 0.1) continue;

      const overlap = getOverlap(baseBounds, newBounds);
      if (!overlap || overlap.overlapHorizontal < 0.2 || overlap.overlapVertical < 0.05) continue;

      const renderedBottom = estimateRenderedTextBottom(baseObj);
      if (renderedBottom == null) continue;
      if (renderedBottom > newBounds.y + 0.06) {
        const message = `WARN Slide ${slide && slide._slideNum ? slide._slideNum : "?"}: reveal element (y=${newBounds.y.toFixed(2)}") covers base slide text that renders to ~y=${renderedBottom.toFixed(2)}". Shorten the text, raise its box, or move the reveal bar down (CLAUDE.md Reveal Bar Clearance).`;
        warnings.push(message);
        console.warn(message);
        return warnings; // one warning per slide is enough
      }
    }
  }

  return warnings;
}

/**
 * Create a click-to-reveal slide pair.
 * Calls buildFn twice: first call creates the "question" slide (no answer),
 * second call creates an identical slide, then revealFn adds the answer.
 * Teacher clicks "next" in PowerPoint to advance from question -> answer.
 *
 * The reveal layer is automatically checked against the base slide's
 * rendered text (see warnIfRevealCoversBaseText).
 *
 * @param {Function} buildFn  - zero-arg function that calls a slide builder
 *                               and returns the slide (e.g. () => cfuSlide(...))
 * @param {Function} revealFn - callback(slide) that adds answer/reveal content
 * @returns {object} the answer slide (second slide)
 */
function withReveal(buildFn, revealFn) {
  buildFn();              // Slide 1: question only
  const s = buildFn();    // Slide 2: identical base
  const preCount = s && Array.isArray(s._slideObjects) ? s._slideObjects.length : 0;
  revealFn(s);            // Add reveal content to slide 2
  if (s) warnIfRevealCoversBaseText(s, preCount);
  return s;
}

module.exports = { withReveal, warnIfRevealCoversBaseText };
