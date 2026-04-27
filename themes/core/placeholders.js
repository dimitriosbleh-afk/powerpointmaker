"use strict";

const { validateBounds } = require("./layout");

const PLACEHOLDER_KINDS = ["quote", "extract", "image", "video", "url", "ochre", "blm", "vocab"];

const PLACEHOLDER_COPY = {
  quote:   "Teacher-provided quote goes here.\nUse exact wording from the supplied extract.",
  extract: "Read the selected paragraph from the class text.\nUse the supplied extract only.",
  image:   "Insert the user-provided image here.\nDo not substitute generic clip art.",
  video:   "Teacher-selected school-approved video goes here.\nDo not invent a URL.",
  url:     "Teacher opens the user-provided URL here.",
  ochre:   "Use the supplied OCHRE slide here.",
  blm:     "Use the supplied BLM here.",
  vocab:   "Teacher inserts selected vocabulary from the supplied extract.",
};

/**
 * Anti-hallucination placeholder helper (megaprompt §5 / §5a). When a
 * build script lacks a verified source quote, video, URL, image, or OCHRE
 * slide, it must render a visible "teacher-provided X goes here" stamp
 * instead of inventing content.
 */
function createPlaceholderHelpers(C, FONT_H, FONT_B, el) {

  function addTeacherProvidedPlaceholder(slide, kind, bounds, opts) {
    const o = opts || {};
    const b = bounds || {};
    const x = b.x != null ? b.x : 0.5;
    const y = b.y != null ? b.y : 1.4;
    const w = b.w != null ? b.w : 9;
    const h = b.h != null ? b.h : 2;
    validateBounds("addTeacherProvidedPlaceholder", x, y, w, h);

    const copy = PLACEHOLDER_COPY[kind] || PLACEHOLDER_COPY.extract;
    const label = o.label || `TEACHER-PROVIDED ${String(kind || "SOURCE").toUpperCase()}`;
    const fillColor = o.fill || C.BG_LIGHT;
    const borderColor = o.borderColor || C.ALERT;
    const textColor = o.color || C.CHARCOAL;

    slide.addShape("roundRect", {
      x, y, w, h, rectRadius: 0.08,
      fill: { color: fillColor },
      line: { color: borderColor, width: 2.0, dashType: "dash" },
    });

    const tagW = Math.min(3.4, w * 0.6);
    el.addTextOnShape(slide, label, {
      x: x + 0.18, y: y + 0.14, w: tagW, h: 0.32, rectRadius: 0.06,
      fill: { color: borderColor },
    }, {
      fontSize: 11, fontFace: FONT_B, color: C.WHITE,
      bold: true, align: "center", valign: "middle", margin: 0,
    });

    slide.addText(String(copy), {
      x: x + 0.3, y: y + 0.56,
      w: w - 0.6, h: h - 0.74,
      fontSize: o.fontSize || 16,
      fontFace: FONT_H,
      color: textColor,
      italic: true,
      align: "left", valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  }

  return { addTeacherProvidedPlaceholder, PLACEHOLDER_COPY, PLACEHOLDER_KINDS };
}

module.exports = { createPlaceholderHelpers, PLACEHOLDER_COPY, PLACEHOLDER_KINDS };
