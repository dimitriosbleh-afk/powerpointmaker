"use strict";

const { SLIDE_W, SAFE_BOTTOM, validateBounds } = require("./layout");
const { validateContrast } = require("./contrast");
const { DEFAULT_SIZES, byBand } = require("./gradeBand");
const { mixHex } = require("./color");

/**
 * Create element helpers bound to a specific palette and grade band.
 * All returned functions close over C, FONT_H, FONT_B, cardShadowFn, and the
 * grade-band sizing table S.
 *
 * @param {object}   C             - palette colours object (semantic keys)
 * @param {string}   FONT_H        - heading font name
 * @param {string}   FONT_B        - body font name
 * @param {Function} cardShadowFn  - zero-arg factory that returns a fresh card shadow object
 * @param {object}   [S]           - grade-band sizing table (from getGradeSizes); falls back to Y36 sizes when omitted for back-compat
 * @returns {object} { addTopBar, addBadge, addTitle, addCard, addInstructionCard, addFooter, addIconCircle, addTextOnShape }
 */
function createElements(C, FONT_H, FONT_B, cardShadowFn, S) {
  // Back-compat: callers built before grade-aware sizing supply only 4 args.
  // Default to the upper-primary table — that matches the historical hard-coded values.
  const sz = S || DEFAULT_SIZES;

  function addTopBar(slide, color) {
    slide.background = { color: C.BG_LIGHT };
    // Thick enough to read as a deliberate colour signal for the stage, not
    // a hairline artefact.
    slide.addShape("rect", {
      x: 0, y: 0, w: SLIDE_W, h: 0.09,
      fill: { color: color || C.PRIMARY },
    });
  }

  /** Soft wash of a colour toward the slide background (for large fills). */
  function softOf(color, amount) {
    return mixHex(color || C.PRIMARY, C.BG_LIGHT, amount != null ? amount : 0.88);
  }

  /** Hairline border tone of a colour. */
  function lineOf(color, amount) {
    return mixHex(color || C.PRIMARY, C.BG_LIGHT, amount != null ? amount : 0.6);
  }

  function addBadge(slide, text, opts) {
    const o = opts || {};
    const bandH = byBand(sz, 0.42, 0.40, 0.36);
    const x     = o.x != null ? o.x : 0.5;
    const y     = o.y != null ? o.y : 0.20;
    const w     = o.w || (byBand(sz, 2.1, 1.95, 1.8));
    const h     = o.h || bandH;
    const color = o.color || C.PRIMARY;
    // Pill, not lozenge: reads friendlier and matches the chips elsewhere.
    slide.addShape("roundRect", {
      x, y, w, h, rectRadius: h / 2,
      fill: { color },
    });
    slide.addText(text, {
      x, y, w, h,
      fontSize: o.fontSize || sz.badge,
      fontFace: FONT_B, color: C.WHITE,
      align: "center", valign: "middle", bold: true, margin: 0,
      fit: "shrink", shrinkText: true,
    });
  }

  function addTitle(slide, title, opts) {
    const o = opts || {};
    // Title box bottom must clear CONTENT_TOP (1.30") by >= 0.07" so the
    // descender of "g"/"y" cannot bleed into the LI/SC or content card.
    const bandH = 0.58;
    const bandY = byBand(sz, 0.65, 0.62, 0.62);
    slide.addText(title, {
      x: o.x || 0.5,
      y: o.y != null ? o.y : bandY,
      w: o.w || 9.0,
      h: o.h || bandH,
      fontSize: o.fontSize || sz.titleH1,
      fontFace: FONT_H,
      color: o.color || C.PRIMARY,
      bold: true,
      margin: 0,
      fit: o.fit || "shrink",
      shrinkText: o.shrinkText != null ? o.shrinkText : true,
    });
  }

  /**
   * Card surface. Three variants:
   *   "white"   (default) white fill, soft shadow, optional left strip
   *   "tint"    soft wash of `tone` (or the strip colour), hairline border,
   *             no shadow - the calm hero surface for the main task, question
   *             or model (megaprompt 18a: soften large fills)
   *   "outline" white fill, hairline border in `tone`, no shadow - for option
   *             cards and reading panels
   * Pass `strip` to add the coloured left edge on any variant.
   */
  function addCard(slide, x, y, w, h, opts) {
    const o = opts || {};
    validateBounds("addCard", x, y, w, h);
    const variant = o.variant || "white";
    const tone = o.tone || o.strip || C.PRIMARY;
    if (variant === "tint") {
      slide.addShape("roundRect", {
        x, y, w, h, rectRadius: 0.12,
        fill: { color: o.fill || softOf(tone) },
        line: { color: o.line || lineOf(tone), width: 0.75 },
      });
    } else if (variant === "outline") {
      slide.addShape("roundRect", {
        x, y, w, h, rectRadius: 0.12,
        fill: { color: o.fill || C.WHITE },
        line: { color: o.line || lineOf(tone), width: 1.0 },
      });
    } else {
      slide.addShape("roundRect", {
        x, y, w, h, rectRadius: 0.1,
        fill: { color: o.fill || C.WHITE },
        shadow: o.shadow || cardShadowFn(),
      });
    }
    if (o.strip) {
      slide.addShape("rect", { x, y, w: 0.07, h, fill: { color: o.strip } });
    }
  }

  function addInstructionCard(slide, items, opts) {
    const o = opts || {};
    const x = o.x != null ? o.x : 0.5;
    const y = o.y != null ? o.y : 1.3;
    const w = o.w || 4.5;
    const h = o.h || 2.4;
    const padX = o.padX != null ? o.padX : 0.2;
    const padY = o.padY != null ? o.padY : 0.14;
    const textW = o.textW || (w - padX * 2);
    const textH = o.textH || (h - padY * 2);
    const contentItems = (items || []).filter((item) => item && item.role !== "spacer");
    const bodyItems = contentItems.filter((item) => !item.role || item.role === "body");
    const longestBody = bodyItems.reduce((best, item) => Math.max(best, String(item.text || "").length), 0);
    const bodyCount = Math.max(bodyItems.length, 1);

    let bodyFontSize = o.bodyFontSize;
    if (!bodyFontSize) {
      const baseBody = sz.body || 16.5;
      const denseBody = sz.bodyDense || 14.5;
      if (bodyCount <= 3 && longestBody <= 40) bodyFontSize = baseBody;
      else if (bodyCount <= 4 && longestBody <= 48) bodyFontSize = (baseBody + denseBody) / 2;
      else if (bodyCount <= 5 && longestBody <= 56) bodyFontSize = denseBody;
      else bodyFontSize = denseBody * 0.95;
    }

    const headerCap   = byBand(sz, 24, 20, 17.5);
    const emphasisCap = byBand(sz, 22, 18, 15.5);
    const headerFontSize = o.headerFontSize || Math.min(bodyFontSize + 3.5, headerCap);
    const emphasisFontSize = o.emphasisFontSize || Math.min(bodyFontSize + 1.5, emphasisCap);

    addCard(slide, x, y, w, h, {
      strip: o.strip,
      fill: o.fill,
      shadow: o.shadow,
      variant: o.variant,
      tone: o.tone,
    });

    const textRuns = [];
    (items || []).forEach((item, index) => {
      const role = item && item.role ? item.role : "body";
      const fontSize = item && item.fontSize ? item.fontSize
        : role === "header" ? headerFontSize
        : role === "emphasis" ? emphasisFontSize
        : role === "spacer" ? 5
        : bodyFontSize;
      const color = item && item.color ? item.color
        : role === "header" ? (o.headerColor || o.strip || C.PRIMARY)
        : role === "emphasis" ? (o.emphasisColor || C.ALERT)
        : (o.bodyColor || C.CHARCOAL);
      const breakLine = item && item.breakLine != null ? item.breakLine : index < (items || []).length - 1;
      textRuns.push({
        text: role === "spacer" ? "" : String((item && item.text) || ""),
        options: {
          bold: role === "header" || role === "emphasis" || Boolean(item && item.bold),
          italic: Boolean(item && item.italic),
          breakLine,
          fontSize,
          color,
        },
      });
    });

    slide.addText(textRuns, {
      x: x + padX,
      y: y + padY,
      w: textW,
      h: textH,
      fontFace: o.fontFace || FONT_B,
      margin: 0,
      valign: o.valign || "top",
      fit: o.fit || "shrink",
      paraSpaceAfter: o.paraSpaceAfter != null ? o.paraSpaceAfter : 1,
    });
  }

  function addFooter(slide, text) {
    slide.addText(text, {
      x: 0.5, y: 5.32, w: 9, h: 0.20,
      fontSize: sz.footer || 9,
      fontFace: FONT_B, color: C.MUTED, margin: 0,
    });
  }

  /** Icon in a coloured circle (roundRect for LibreOffice compatibility). */
  function addIconCircle(slide, iconData, cx, cy, r, circleColor) {
    slide.addShape("roundRect", {
      x: cx - r, y: cy - r, w: r * 2, h: r * 2, rectRadius: r,
      fill: { color: circleColor || C.PRIMARY },
    });
    const iconSize = r * 1.1;
    slide.addImage({
      data: iconData,
      x: cx - iconSize / 2,
      y: cy - iconSize / 2,
      w: iconSize, h: iconSize,
    });
  }

  /**
   * Add a shape with centred text overlay — the safe way to put text on a shape.
   * Guarantees valign:"middle", align:"center", margin:0 unless explicitly overridden.
   * Runs contrast validation automatically.
   */
  function addTextOnShape(slide, text, shapeOpts, textOpts) {
    const so = shapeOpts || {};
    const to = textOpts || {};
    const shapeType = so.rectRadius ? "roundRect" : "rect";
    const fillObj = so.fill
      ? (typeof so.fill === "string" ? { color: so.fill } : so.fill)
      : undefined;

    slide.addShape(shapeType, {
      x: so.x, y: so.y, w: so.w, h: so.h,
      rectRadius: so.rectRadius,
      fill: fillObj,
      line: so.line,
      shadow: so.shadow,
    });

    if (to.color && fillObj && fillObj.color) {
      validateContrast(to.color, fillObj.color, "addTextOnShape");
    }

    slide.addText(text, {
      x: so.x, y: so.y, w: so.w, h: so.h,
      align:    to.align    || "center",
      valign:   to.valign   || "middle",
      margin:   to.margin != null ? to.margin : 0,
      fontSize: to.fontSize,
      fontFace: to.fontFace || FONT_B,
      color:    to.color,
      bold:     to.bold,
      italic:   to.italic,
      fit:      to.fit || "shrink",
      shrinkText: to.shrinkText != null ? to.shrinkText : true,
    });
  }

  return { addTopBar, addBadge, addTitle, addCard, addInstructionCard, addFooter, addIconCircle, addTextOnShape, softOf, lineOf };
}

module.exports = { createElements };
