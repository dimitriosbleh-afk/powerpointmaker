"use strict";

const { SLIDE_W, SLIDE_H, SAFE_BOTTOM } = require("./layout");

const EMU_PER_INCH = 914400;

function getSlideDimensions(slide, pres) {
  const layout = (slide && slide._presLayout) || (pres && pres._presLayout) || (pres && pres.presLayout);
  if (layout && layout.width && layout.height) {
    return {
      width: layout.width / EMU_PER_INCH,
      height: layout.height / EMU_PER_INCH,
    };
  }

  return { width: SLIDE_W, height: SLIDE_H };
}

function inferElementType(obj) {
  if (!obj) return "unknown";
  if (obj._type === "image") return "image";
  if (obj._type === "text" && obj.shape === "line") return "line";
  if (obj._type === "text" && obj.text == null) return "shape";
  if (obj._type === "text") return "text";
  return obj._type || "unknown";
}

function getElementBounds(obj) {
  const options = obj && obj.options;
  if (!options || !isFinite(options.x) || !isFinite(options.y) || !isFinite(options.w) || !isFinite(options.h)) {
    return null;
  }

  if (
    obj &&
    obj._type === "image" &&
    options.sizing &&
    options.sizing.type === "crop" &&
    isFinite(options.sizing.w) &&
    isFinite(options.sizing.h)
  ) {
    return {
      x: Number(options.x),
      y: Number(options.y),
      w: Number(options.sizing.w),
      h: Number(options.sizing.h),
    };
  }

  return {
    x: Number(options.x),
    y: Number(options.y),
    w: Number(options.w),
    h: Number(options.h),
  };
}

function isFooterElement(obj) {
  const bounds = getElementBounds(obj);
  if (!bounds || inferElementType(obj) !== "text") return false;
  return bounds.y >= SAFE_BOTTOM + 0.15 && bounds.h <= 0.35 && bounds.w >= 7.5;
}

function describeElement(obj, index) {
  const bounds = getElementBounds(obj);
  const type = inferElementType(obj);
  if (!bounds) {
    return `element ${index} (${type})`;
  }

  const centerX = bounds.x + bounds.w / 2;
  const centerY = bounds.y + bounds.h / 2;
  return `element ${index} (${type}, center_x=${centerX.toFixed(3)}, center_y=${centerY.toFixed(3)})`;
}

function shouldIgnoreIndex(index, obj, opts) {
  const o = opts || {};
  if ((o.ignoreIndices || []).includes(index)) return true;
  if (!obj) return true;
  if (inferElementType(obj) === "line") return true;
  if (o.ignoreFooter !== false && isFooterElement(obj)) return true;
  if (obj._type === "notes") return true;
  return false;
}

function makePairKey(a, b) {
  return `${Math.min(a, b)}:${Math.max(a, b)}`;
}

function shouldIgnorePair(indexA, objA, indexB, objB, opts) {
  if (shouldIgnoreIndex(indexA, objA, opts) || shouldIgnoreIndex(indexB, objB, opts)) {
    return true;
  }

  const typeA = inferElementType(objA);
  const typeB = inferElementType(objB);
  const allowedPairs = new Set(["text:image", "image:text", "text:text", "image:image"]);
  if (!allowedPairs.has(`${typeA}:${typeB}`)) {
    return true;
  }

  const o = opts || {};
  const ignorePairs = o.ignorePairs || [];
  const pairKey = makePairKey(indexA, indexB);
  return ignorePairs.some((pair) => Array.isArray(pair) && makePairKey(pair[0], pair[1]) === pairKey);
}

function getOverlap(boundsA, boundsB) {
  const overlapHorizontal = Math.min(boundsA.x + boundsA.w, boundsB.x + boundsB.w) - Math.max(boundsA.x, boundsB.x);
  const overlapVertical = Math.min(boundsA.y + boundsA.h, boundsB.y + boundsB.h) - Math.max(boundsA.y, boundsB.y);

  if (overlapHorizontal <= 0 || overlapVertical <= 0) {
    return null;
  }

  return {
    overlapHorizontal,
    overlapVertical,
    overlapArea: overlapHorizontal * overlapVertical,
  };
}

function warnIfSlideHasOverlaps(slide, pres, opts) {
  const objects = Array.isArray(slide && slide._slideObjects) ? slide._slideObjects : [];
  const warnings = [];

  for (let i = 0; i < objects.length; i += 1) {
    const boundsA = getElementBounds(objects[i]);
    if (!boundsA) continue;

    for (let j = i + 1; j < objects.length; j += 1) {
      if (shouldIgnorePair(i, objects[i], j, objects[j], opts)) continue;

      const boundsB = getElementBounds(objects[j]);
      if (!boundsB) continue;

      const overlap = getOverlap(boundsA, boundsB);
      if (!overlap) continue;
      if (overlap.overlapHorizontal < 0.04 || overlap.overlapVertical < 0.04) continue;

      const smallerArea = Math.max(Math.min(boundsA.w * boundsA.h, boundsB.w * boundsB.h), 0.001);
      const overlapRatio = overlap.overlapArea / smallerArea;
      const isSevere = overlap.overlapHorizontal >= 0.12 && overlap.overlapVertical >= 0.12 && overlapRatio >= 0.08;
      const message = `${isSevere ? "ERROR" : "WARN"} Slide ${slide && slide._slideNum ? slide._slideNum : "?"}: ` +
        `${isSevere ? "Severe text/image overlap detected" : "Overlap detected"} between ` +
        `${describeElement(objects[i], i)} and ${describeElement(objects[j], j)} ` +
        `(overlap_horizontal=${overlap.overlapHorizontal.toFixed(3)}, ` +
        `overlap_vertical=${overlap.overlapVertical.toFixed(3)}).` +
        `${isSevere ? " THIS MUST BE FIXED." : ""}`;
      warnings.push(message);
      console.warn(message);
    }
  }

  if (warnings.length > 0) {
    console.warn(`WARN Slide ${slide && slide._slideNum ? slide._slideNum : "?"}: Found ${warnings.length} overlapping pair(s).`);
  }

  return warnings;
}

function warnIfSlideElementsOutOfBounds(slide, pres, opts) {
  const objects = Array.isArray(slide && slide._slideObjects) ? slide._slideObjects : [];
  const warnings = [];
  const dims = getSlideDimensions(slide, pres);
  const safeBottom = (opts && isFinite(opts.safeBottom)) ? Number(opts.safeBottom) : SAFE_BOTTOM;
  const respectSafeBottom = !(opts && opts.respectSafeBottom === false);

  objects.forEach((obj, index) => {
    if (shouldIgnoreIndex(index, obj, opts)) return;
    const bounds = getElementBounds(obj);
    if (!bounds) return;

    const issues = [];
    if (bounds.x < -0.01) issues.push(`left ${bounds.x.toFixed(2)}" is outside slide`);
    if (bounds.y < -0.01) issues.push(`top ${bounds.y.toFixed(2)}" is outside slide`);
    if (bounds.x + bounds.w > dims.width + 0.01) {
      issues.push(`right edge ${(bounds.x + bounds.w).toFixed(2)}" exceeds slide width (${dims.width.toFixed(2)}")`);
    }
    if (bounds.y + bounds.h > dims.height + 0.01) {
      issues.push(`bottom ${(bounds.y + bounds.h).toFixed(2)}" exceeds slide height (${dims.height.toFixed(2)}")`);
    } else if (respectSafeBottom && bounds.y + bounds.h > safeBottom + 0.01) {
      issues.push(`bottom ${(bounds.y + bounds.h).toFixed(2)}" exceeds safe content limit (${safeBottom.toFixed(2)}")`);
    }

    if (issues.length === 0) return;

    const message = `WARN Slide ${slide && slide._slideNum ? slide._slideNum : "?"}: ${describeElement(obj, index)} is out of bounds: ${issues.join("; ")}.`;
    warnings.push(message);
    console.warn(message);
  });

  if (warnings.length > 0) {
    console.warn(`WARN Slide ${slide && slide._slideNum ? slide._slideNum : "?"}: Found ${warnings.length} out-of-bounds element(s).`);
  }

  return warnings;
}

function runSlideDiagnostics(slide, pres, opts) {
  return {
    overlaps: warnIfSlideHasOverlaps(slide, pres, opts),
    outOfBounds: warnIfSlideElementsOutOfBounds(slide, pres, opts),
  };
}

module.exports = {
  inferElementType,
  getSlideDimensions,
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
  runSlideDiagnostics,
};
