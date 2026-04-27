"use strict";

/**
 * Bullet-list and text-fit helpers used by every builder that puts wrapped
 * text inside a fixed-height card. PptxGenJS `fit: "shrink"` on bullet lists
 * is unreliable in LibreOffice and inconsistent in PowerPoint, so we
 * pre-compute a fontSize that we know will fit, and we pre-process the
 * bullet list so build-script empty-string "spacers" become real
 * paragraph-spacing rather than ugly empty bullet markers.
 */

/**
 * Filter empty-string entries out of a raw bullet array. Each empty entry
 * becomes a `extraSpaceAfter` increment on the preceding non-empty item.
 * Leading empty entries are dropped.
 *
 * Input:  ["A", "", "B", "C", "", "D"]
 * Output: [
 *   { text: "A", extraSpaceAfter: 1 },
 *   { text: "B", extraSpaceAfter: 0 },
 *   { text: "C", extraSpaceAfter: 1 },
 *   { text: "D", extraSpaceAfter: 0 },
 * ]
 */
function prepareBullets(rawItems) {
  const out = [];
  const items = Array.isArray(rawItems) ? rawItems : [];
  for (const raw of items) {
    const text = String(raw == null ? "" : raw);
    if (text.trim() === "") {
      if (out.length > 0) {
        out[out.length - 1].extraSpaceAfter += 1;
      }
      continue;
    }
    out.push({ text, extraSpaceAfter: 0 });
  }
  return out;
}

/**
 * Estimate rendered height of a prepared bullet list at the given fontSize.
 * Returns inches.
 *
 * Per-paragraph height = wrappedLines * lineHeight, plus the paragraph
 * space-after that PowerPoint inserts between paragraphs. The per-bullet
 * empty-space boost adds roughly one line-height per `extraSpaceAfter`.
 */
function estimateBulletHeight(items, fontSize, charsPerLine) {
  if (!items || !items.length) return 0;
  const lineHeight = Math.max(0.20, fontSize * 0.018 + 0.06);
  // paraSpaceAfter is in points; PptxGenJS forwards it to PowerPoint where
  // 1 pt ~= 0.0139 inches. Round up to 0.014 for safety.
  const baseSpacePt = fontSize >= 16 ? 5 : 3;
  let total = 0;
  for (let i = 0; i < items.length; i += 1) {
    const item = items[i];
    const text = String(item.text || "");
    const lines = Math.max(1, Math.ceil(text.length / Math.max(charsPerLine, 1)));
    total += lines * lineHeight;
    const extraGapInches = (item.extraSpaceAfter || 0) * lineHeight * 0.9;
    if (i < items.length - 1) {
      total += baseSpacePt * 0.014;
      total += extraGapInches;
    }
  }
  return total;
}

/**
 * Find the largest fontSize in [floor, ideal] (0.5pt steps) at which the
 * prepared bullet list fits within availableH inches. Returns floor if
 * even the floor size would overflow.
 */
function fitBulletFontSize(items, availableH, charsPerLine, ideal, floor) {
  if (!items || !items.length) return ideal;
  if (estimateBulletHeight(items, ideal, charsPerLine) <= availableH) return ideal;
  let size = ideal - 0.5;
  while (size > floor) {
    if (estimateBulletHeight(items, size, charsPerLine) <= availableH) return size;
    size -= 0.5;
  }
  return floor;
}

/**
 * Estimate rendered height of a single multi-line text block (no bullets)
 * at fontSize. Used by the CFU question card where the body is one
 * `addText` call with embedded newlines.
 */
function estimateTextHeight(text, fontSize, charsPerLine) {
  const raw = String(text == null ? "" : text);
  const segments = raw.split("\n");
  const lineHeight = Math.max(0.24, fontSize * 0.020 + 0.06);
  const totalLines = segments.reduce((count, segment) => {
    const trimmed = segment.trim();
    if (!trimmed) return count + 1;
    return count + Math.max(1, Math.ceil(trimmed.length / Math.max(charsPerLine, 1)));
  }, 0);
  return totalLines * lineHeight;
}

/**
 * Find the largest fontSize in [floor, ideal] (0.5pt steps) at which the
 * text block fits within availableH inches.
 */
function fitTextFontSize(text, availableH, charsPerLine, ideal, floor) {
  if (estimateTextHeight(text, ideal, charsPerLine) <= availableH) return ideal;
  let size = ideal - 0.5;
  while (size > floor) {
    if (estimateTextHeight(text, size, charsPerLine) <= availableH) return size;
    size -= 0.5;
  }
  return floor;
}

module.exports = {
  prepareBullets,
  estimateBulletHeight,
  fitBulletFontSize,
  estimateTextHeight,
  fitTextFontSize,
};
