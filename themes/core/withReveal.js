"use strict";

const { SAFE_BOTTOM } = require("./layout");
const { inferElementType, getElementBounds, getOverlap } = require("./diagnostics");
const { getSlideNotesText } = require("./notes");

const CONTINUATION_INDENT = "   ";
const MAX_LINE_WORDS = 14; // under the 16-word rendered budget in section 46

/**
 * Wrap a long string into short lines, continuation lines indented, so a
 * derived note line can never breach the per-line word budget.
 */
function wrapNoteLine(prefix, text) {
  const words = String(text || "").trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return [prefix.trim()];

  const lines = [];
  let current = prefix;
  let currentWords = prefix.trim().split(/\s+/).length;

  words.forEach((word) => {
    if (currentWords >= MAX_LINE_WORDS) {
      lines.push(current);
      current = CONTINUATION_INDENT + word;
      currentWords = 1;
      return;
    }
    current += (current.endsWith(" ") ? "" : " ") + word;
    currentWords += 1;
  });
  if (current.trim()) lines.push(current);
  return lines;
}

/**
 * Derive post-reveal notes from the base slide's notes.
 *
 * A reveal slide must never carry a byte-copy of its base slide's notes: when
 * the teacher clicks to the answer the notes have to advance with the slide
 * (megaprompt section 47). Authored notes via composeRevealNotes are always
 * better, but 253 of the 264 build scripts predate that rule, so the fallback
 * has to produce something correct rather than leaving the copy in place.
 *
 * The ANSWER line carries the real content, so it becomes the REVEALED line.
 * The prep-zone tag is carried across so the stage / element / SC tag survives.
 *
 * @returns {string|null} derived notes, or null if nothing usable was found
 */
function deriveRevealNotes(baseNotes) {
  const raw = String(baseNotes || "").replace(/\r\n?/g, "\n").trim();
  if (!raw) return null;

  const lines = raw.split("\n");
  const dividerIndex = lines.findIndex((l) => l.trim() === "---");
  const liveLines = (dividerIndex === -1 ? lines : lines.slice(0, dividerIndex));
  const prepLines = dividerIndex === -1 ? [] : lines.slice(dividerIndex + 1);

  // A single-line, zone-less note is a non-teaching slide. Nothing to reveal.
  if (dividerIndex === -1 && liveLines.filter((l) => l.trim()).length <= 1) return null;

  const answerLine = liveLines.find((l) => /^\s*ANSWER:/i.test(l));
  const answer = answerLine ? answerLine.replace(/^\s*ANSWER:\s*/i, "").trim() : "";

  const out = [];
  out.push(...(answer
    ? wrapNoteLine("REVEALED:", answer)
    : ["REVEALED: the answer is now on screen."]));

  out.push("");
  out.push("1. SAY: Tick it or fix it. Check your board against this.");

  out.push("");
  out.push("2. CIRCULATE. Any board that does not match, fix it now.");

  // Carry the prep-zone tag across so the stage and SC focus are not lost.
  const tag = prepLines
    .map((l) => (l.match(/\[[^\]]+\]\s*$/) || [])[0])
    .filter(Boolean)
    .pop();

  out.push("---");
  out.push(tag ? `Reveal slide. ${tag}` : "Reveal slide.");

  return out.join("\n");
}

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
 * Does this reveal object paint an opaque background?
 *
 * An answer bar is a filled shape, so anything wholly behind it is hidden
 * cleanly and on purpose. A bare text run paints nothing and hides nothing.
 */
function isOpaqueRevealElement(obj) {
  const options = (obj && obj.options) || {};
  const fill = options.fill;
  if (!fill) return false;
  if (typeof fill === "string") return true;
  if (typeof fill === "object" && fill.color) {
    // transparency is a percentage: 100 means fully see-through
    const t = Number(fill.transparency);
    return !(isFinite(t) && t >= 90);
  }
  return false;
}

/**
 * Warn when reveal-layer elements (added by revealFn) sit on top of the base
 * slide's rendered text.
 *
 * The defect this exists to catch is a reveal bar clipping text the teacher
 * still needs - the last line of a question disappearing under the answer.
 * That shows up as a PARTIAL overlap: some of the text is covered and the rest
 * bleeds out past the edge of the bar, which looks broken.
 *
 * A reveal element that sits opaquely over the WHOLE of a base text box is the
 * opposite: a deliberate overlay, replacing a spent prompt ("Write the name of
 * each shape on your whiteboard") with the answer. It renders cleanly, because
 * the covered text is completely hidden. Flagging that was wrong, and it was
 * wrong on 133 slides across the library - enough noise to drown the real
 * defect. Full opaque containment is therefore allowed.
 */
function warnIfRevealCoversBaseText(slide, preCount) {
  const objects = Array.isArray(slide && slide._slideObjects) ? slide._slideObjects : [];
  const warnings = [];

  // Coverage is a property of the whole reveal layer, not of one element.
  // addTextOnShape emits a filled shape AND a bare text run: the shape does the
  // hiding, the text run rides on top. Checking them individually would clear
  // the shape and then flag its own label.
  const opaqueCovers = [];
  for (let k = preCount; k < objects.length; k += 1) {
    if (!isOpaqueRevealElement(objects[k])) continue;
    const b = getElementBounds(objects[k]);
    if (b) opaqueCovers.push(b);
  }

  // Text is safely hidden when an opaque reveal element contains it
  // horizontally AND its bottom edge sits at or below where the text stops
  // rendering. The text may start above the element's top edge - a text box
  // usually begins slightly higher than its first glyph - so the top edge is
  // deliberately not part of the test.
  const isHiddenCleanly = (baseBounds, renderedBottom) =>
    opaqueCovers.some((c) =>
      c.x <= baseBounds.x + 0.05 &&
      c.x + c.w >= baseBounds.x + baseBounds.w - 0.05 &&
      c.y + c.h >= renderedBottom - 0.02
    );

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

      if (isHiddenCleanly(baseBounds, renderedBottom)) continue;

      // The defect: text escaping the reveal element and rendering past its
      // bottom edge, so half a sentence pokes out under the answer bar.
      const escapesBelow = renderedBottom > newBounds.y + newBounds.h + 0.02;
      if (escapesBelow) {
        const message = `WARN Slide ${slide && slide._slideNum ? slide._slideNum : "?"}: reveal element (y=${newBounds.y.toFixed(2)}"-${(newBounds.y + newBounds.h).toFixed(2)}") partly covers base slide text that renders to ~y=${renderedBottom.toFixed(2)}", so the text pokes out beneath it. Shorten the text, raise its box, or make the reveal element taller (CLAUDE.md Reveal Bar Clearance).`;
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
 * The reveal slide must NOT carry a byte-copy of the base slide's notes -
 * mid-lesson, the teacher clicks to the answer and the notes must advance
 * with the slide. Pass opts.revealNotes (usually composeRevealNotes(...))
 * and it replaces the reveal slide's notes with the post-reveal script.
 * build_and_check.js Gate 5 fails the build when consecutive slides carry
 * identical notes, so omitting revealNotes is a build error for new decks.
 *
 * @param {Function} buildFn  - zero-arg function that calls a slide builder
 *                               and returns the slide (e.g. () => cfuSlide(...))
 * @param {Function} revealFn - callback(slide) that adds answer/reveal content
 * @param {object}   [opts]
 * @param {string}   [opts.revealNotes] - replacement notes for the reveal slide
 * @returns {object} the answer slide (second slide)
 */
function withReveal(buildFn, revealFn, opts) {
  const o = opts || {};
  buildFn();              // Slide 1: question only
  const s = buildFn();    // Slide 2: identical base
  const preCount = s && Array.isArray(s._slideObjects) ? s._slideObjects.length : 0;
  revealFn(s);            // Add reveal content to slide 2
  if (s) warnIfRevealCoversBaseText(s, preCount);
  if (s && Array.isArray(s._slideObjects)) {
    // Authored notes win. Otherwise derive them, so the reveal slide never
    // ships a byte-copy of the base slide's notes (megaprompt section 47).
    const replacement = o.revealNotes
      ? String(o.revealNotes)
      : deriveRevealNotes(getSlideNotesText(s));

    if (replacement) {
      if (!o.revealNotes) {
        // Advisory on stdout, not stderr: it must not fail the build (legacy
        // decks rely on the fallback) but a NEW deck should never quietly
        // ship a derived reveal script. Section 47 wants it authored.
        console.log(
          `ADVISORY reveal slide used DERIVED notes - author them with ` +
          `composeRevealNotes({ answer, beats }) and pass as withReveal(..., { revealNotes })`
        );
      }
      // Remove the notes copied in by the second buildFn call, then attach the
      // post-reveal notes. Mutate in place: _slideObjects may be a getter.
      for (let i = s._slideObjects.length - 1; i >= 0; i -= 1) {
        const obj = s._slideObjects[i];
        if (obj && obj._type === "notes") s._slideObjects.splice(i, 1);
      }
      s.addNotes(replacement);
    }
  }
  return s;
}

module.exports = { withReveal, warnIfRevealCoversBaseText, deriveRevealNotes };
