"use strict";

/**
 * Colour utilities shared by the palette retune script, the factory's derived
 * tint keys, and any builder that needs a lighter or darker version of a
 * palette colour.
 *
 * All colours are 6-char hex strings WITHOUT a leading "#" (PptxGenJS corrupts
 * files when it sees one). Every function accepts either form and returns the
 * bare form in upper case.
 */

const { contrastRatio } = require("./contrast");

function normaliseHex(color, fallback) {
  const raw = String(color == null ? "" : color).replace("#", "").trim();
  if (/^[0-9A-Fa-f]{6}$/.test(raw)) return raw.toUpperCase();
  if (/^[0-9A-Fa-f]{3}$/.test(raw)) {
    return raw.split("").map((c) => c + c).join("").toUpperCase();
  }
  return normaliseHex(fallback || "FFFFFF", "FFFFFF");
}

function hexToRgb(color) {
  const value = normaliseHex(color, "FFFFFF");
  return {
    r: parseInt(value.slice(0, 2), 16),
    g: parseInt(value.slice(2, 4), 16),
    b: parseInt(value.slice(4, 6), 16),
  };
}

function rgbToHex(r, g, b) {
  return [r, g, b]
    .map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

/** Linear blend of two colours. amount 0 = colorA, 1 = colorB. */
function mixHex(colorA, colorB, amount) {
  const ratio = Math.max(0, Math.min(1, Number(amount) || 0));
  const a = hexToRgb(colorA);
  const b = hexToRgb(colorB);
  return rgbToHex(
    a.r + (b.r - a.r) * ratio,
    a.g + (b.g - a.g) * ratio,
    a.b + (b.b - a.b) * ratio,
  );
}

function lightenHex(color, amount) {
  return mixHex(color, "FFFFFF", amount);
}

function darkenHex(color, amount) {
  return mixHex(color, "000000", amount);
}

/** RGB (0-255) -> HSL (h 0-360, s 0-1, l 0-1). */
function rgbToHsl(r, g, b) {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === rn) h = ((gn - bn) / d + (gn < bn ? 6 : 0)) * 60;
  else if (max === gn) h = ((bn - rn) / d + 2) * 60;
  else h = ((rn - gn) / d + 4) * 60;
  return { h, s, l };
}

function hslToRgb(h, s, l) {
  const hue = ((h % 360) + 360) % 360;
  if (s === 0) {
    const v = l * 255;
    return { r: v, g: v, b: v };
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const channel = (t) => {
    let tt = t;
    if (tt < 0) tt += 1;
    if (tt > 1) tt -= 1;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };
  const hn = hue / 360;
  return {
    r: channel(hn + 1 / 3) * 255,
    g: channel(hn) * 255,
    b: channel(hn - 1 / 3) * 255,
  };
}

function hexToHsl(color) {
  const { r, g, b } = hexToRgb(color);
  return rgbToHsl(r, g, b);
}

function hslToHex(h, s, l) {
  const { r, g, b } = hslToRgb(h, Math.max(0, Math.min(1, s)), Math.max(0, Math.min(1, l)));
  return rgbToHex(r, g, b);
}

/**
 * Lift (or lower) a colour's lightness until its contrast against `against`
 * is as close as possible to `target` without dropping below it.
 *
 * The palettes were originally authored far darker than WCAG AA needs (many
 * PRIMARY colours sat at 12-18:1 against white) which made every Foundation
 * deck look like a corporate report. This finds the brightest shade of the
 * same hue that still guarantees the pairing, so colour can do its job
 * without the contrast promise weakening.
 *
 * @param {string} color    hex colour to adjust
 * @param {number} target   minimum contrast ratio to keep (e.g. 5.0)
 * @param {string} against  the colour that will sit on/under it (default white)
 * @param {object} [opts]   { minSaturation, maxSaturation, step }
 */
function liftToContrast(color, target, against, opts) {
  const o = opts || {};
  const bg = against || "FFFFFF";
  const { h, s, l } = hexToHsl(color);
  let sat = s;
  if (o.minSaturation != null) sat = Math.max(sat, o.minSaturation);
  if (o.maxSaturation != null) sat = Math.min(sat, o.maxSaturation);
  const step = o.step || 0.005;

  // Direction: if the colour already passes, walk lighter until it would
  // fail; if it fails, walk darker until it passes.
  const passes = (lig) => contrastRatio(hslToHex(h, sat, lig), bg) >= target;
  let lig = l;
  if (passes(lig)) {
    while (lig + step <= 1 && passes(lig + step)) lig += step;
  } else {
    while (lig - step >= 0 && !passes(lig)) lig -= step;
  }
  return hslToHex(h, sat, lig);
}

/**
 * Pick the colour from candidates with the best contrast against bg, or the
 * first that passes `minRatio`. Used to keep on-dark text readable when the
 * dark background is lifted.
 */
function bestContrast(candidates, bg, minRatio) {
  const list = (candidates || []).filter(Boolean);
  const passing = list.find((c) => contrastRatio(c, bg) >= (minRatio || 4.5));
  if (passing) return passing;
  return list.reduce((best, c) => (contrastRatio(c, bg) > contrastRatio(best, bg) ? c : best), list[0]);
}

/**
 * Derived tints for a palette colour. `soft` is a card fill (very light wash
 * of the hue), `line` a hairline border, `deep` a text-safe darker version.
 */
function derivedTones(color, bgLight) {
  const base = normaliseHex(color, "444444");
  const paper = normaliseHex(bgLight || "FFFFFF", "FFFFFF");
  return {
    soft: mixHex(base, paper, 0.88),
    softer: mixHex(base, paper, 0.94),
    line: mixHex(base, paper, 0.6),
    deep: liftToContrast(base, 7.0, paper),
  };
}

module.exports = {
  normaliseHex,
  hexToRgb,
  rgbToHex,
  mixHex,
  lightenHex,
  darkenHex,
  hexToHsl,
  hslToHex,
  liftToContrast,
  bestContrast,
  derivedTones,
};
