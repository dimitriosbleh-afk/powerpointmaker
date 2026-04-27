"use strict";

const PptxGenJS = require("pptxgenjs");

// ── Core modules ──
const { SLIDE_W, SLIDE_H, SAFE_RIGHT, SAFE_BOTTOM, CONTENT_TOP, validateBounds } = require("./core/layout");
const { hexToRgb, luminance, contrastRatio, getContrastColor, validateContrast } = require("./core/contrast");
const { iconToBase64Png }          = require("./core/icons");
const { normalizeLessonTargets, sanitizeTeacherNotes, appendSourcesToNotes, installNotesPatch } = require("./core/notes");
const { makeShadow, makeCardShadow } = require("./core/shadows");
const { createElements }           = require("./core/elements");
const { withReveal }               = require("./core/withReveal");
const { createImageHelpers }       = require("./core/images");
const { warnIfSlideHasOverlaps, warnIfSlideElementsOutOfBounds, runSlideDiagnostics } = require("./core/diagnostics");
const { getGradeBand, getGradeSizes } = require("./core/gradeBand");
const { createRoutineHelpers, ROUTINES, ROUTINE_LABELS } = require("./core/routineIcons");
const { createPlaceholderHelpers } = require("./core/placeholders");
const { composeNotes } = require("./core/composeNotes");

// ── Builder factories ──
const { createBaseBuilders }       = require("./builders/base");
const { createLiteracyBuilders }   = require("./builders/literacy");
const { createNumeracyBuilders }   = require("./builders/numeracy");
const { createWellbeingBuilders }  = require("./builders/wellbeing");
const { createInquiryBuilders }    = require("./builders/inquiry");
const { createScienceBuilders }    = require("./builders/science");

// ── Palette registries ──
const { palettes: litPalettes }    = require("./palettes/literacy");
const { palettes: numPalettes }    = require("./palettes/numeracy");
const { palettes: wbPalettes }     = require("./palettes/wellbeing");
const { palettes: inqPalettes }    = require("./palettes/inquiry");
const { palettes: sciPalettes }    = require("./palettes/science");

const SUBJECT_PALETTES = {
  literacy:  litPalettes,
  numeracy:  numPalettes,
  wellbeing: wbPalettes,
  inquiry:   inqPalettes,
  science:   sciPalettes,
};

const SUBJECT_BUILDER_FACTORIES = {
  literacy:  createLiteracyBuilders,
  numeracy:  createNumeracyBuilders,
  wellbeing: createWellbeingBuilders,
  inquiry:   createInquiryBuilders,
  science:   createScienceBuilders,
};

const VALID_SUBJECTS    = Object.keys(SUBJECT_PALETTES);
const VALID_YEAR_LEVELS = ["foundation", "grade1", "grade2", "grade34", "grade56"];
const VARIANTS_PER_LEVEL = 6;

installNotesPatch(PptxGenJS);

/**
 * Create a fully-bound theme object.
 *
 * @param {string} subject    - "literacy"|"numeracy"|"wellbeing"|"inquiry"|"science"
 * @param {string} yearLevel  - "foundation"|"grade1"|"grade2"|"grade34"|"grade56"
 * @param {number} variant    - 0-5 (week rotation index)
 * @returns {object} theme object with all slide builders, element helpers, palette, and utilities
 */
function createTheme(subject, yearLevel, variant) {
  if (variant == null) variant = 0;
  const subjectLower = String(subject).toLowerCase();
  const levelLower   = String(yearLevel).toLowerCase();
  const variantIdx   = Math.max(0, Math.min(VARIANTS_PER_LEVEL - 1, Math.floor(variant)));

  // Validate subject
  if (!SUBJECT_PALETTES[subjectLower]) {
    throw new Error(
      `[createTheme] Unknown subject "${subject}". Valid: ${VALID_SUBJECTS.join(", ")}`
    );
  }

  // Validate year level
  if (!VALID_YEAR_LEVELS.includes(levelLower)) {
    throw new Error(
      `[createTheme] Unknown yearLevel "${yearLevel}". Valid: ${VALID_YEAR_LEVELS.join(", ")}`
    );
  }

  // Resolve palette
  const subjectPalettes = SUBJECT_PALETTES[subjectLower];
  if (!subjectPalettes[levelLower]) {
    throw new Error(`[createTheme] No palettes for ${subjectLower}/${levelLower}`);
  }
  if (!subjectPalettes[levelLower][variantIdx]) {
    throw new Error(
      `[createTheme] Variant ${variantIdx} not found for ${subjectLower}/${levelLower}. ` +
      `Available: 0-${subjectPalettes[levelLower].length - 1}`
    );
  }

  const palette = subjectPalettes[levelLower][variantIdx];

  // Build colour object with backward-compatible aliases for pdf_helpers.js
  const C = { ...palette };
  C.CREAM = C.CREAM || C.BG_LIGHT;
  C.NAVY  = C.NAVY  || C.PRIMARY;
  C.TEAL  = C.TEAL  || C.SECONDARY;
  // C.WHITE and C.MUTED already exist in the palette schema

  const FONT_H = palette.FONT_H;
  const FONT_B = palette.FONT_B;

  // Resolve grade-band sizing table from yearLevel.
  // Drives every slide builder's font sizes and density caps so that the
  // same builder API renders age-appropriate output without callers
  // passing sizes explicitly.
  const gradeBand = getGradeBand(levelLower);
  const S = getGradeSizes(gradeBand);

  // Build shadow factories
  const shadowFn     = makeShadow(palette);
  const cardShadowFn = makeCardShadow(palette);

  // Build bound element helpers
  const el = createElements(C, FONT_H, FONT_B, cardShadowFn, S);
  const img = createImageHelpers(C, FONT_H, FONT_B, el, cardShadowFn);

  // Build bound getContrastColor (needs palette's WHITE/CHARCOAL)
  const boundGetContrastColor = (bgHex) => getContrastColor(bgHex, C.WHITE, C.CHARCOAL);

  // Build base slide builders (all subjects get these)
  const base = createBaseBuilders(C, FONT_H, FONT_B, el, shadowFn, S);

  // Build subject-specific slide builders
  const subjectFactory = SUBJECT_BUILDER_FACTORIES[subjectLower];
  const subjectBuilders = subjectFactory(C, FONT_H, FONT_B, el, S);

  const routine = createRoutineHelpers(C, FONT_B, el);
  const placeholders = createPlaceholderHelpers(C, FONT_H, FONT_B, el);

  // Compose and return
  return {
    // Palette
    C,
    FONT_H,
    FONT_B,

    // Shadow factories
    makeShadow:     shadowFn,
    makeCardShadow: cardShadowFn,

    // Layout constants
    SLIDE_W, SLIDE_H, SAFE_RIGHT, SAFE_BOTTOM, CONTENT_TOP,

    // Contrast utilities
    hexToRgb, luminance, contrastRatio,
    validateContrast,
    getContrastColor: boundGetContrastColor,

    // Bounds validation
    validateBounds,

    // Content normalization
    normalizeLessonTargets,
    sanitizeTeacherNotes,
    appendSourcesToNotes,

    // Icon rendering
    iconToBase64Png,

    // Element helpers
    ...el,
    ...img,

    // Click-to-reveal
    withReveal,

    // Diagnostics
    warnIfSlideHasOverlaps,
    warnIfSlideElementsOutOfBounds,
    runSlideDiagnostics,

    // Base slide builders (all subjects)
    ...base,

    // Subject-specific slide builders (may override base if name collides)
    ...subjectBuilders,

    ...routine,
    ROUTINES,
    ROUTINE_LABELS,
    ...placeholders,
    composeNotes,

    // Grade-band sizing (frozen) — exposed so manual/custom slides
    // can use the same age-appropriate sizes as the builders.
    S,
    _gradeBand: gradeBand,

    // Metadata
    _subject:     subjectLower,
    _yearLevel:   levelLower,
    _variant:     variantIdx,
    _paletteName: `${subjectLower}/${levelLower}/v${variantIdx}`,
  };
}

/**
 * Convert a 1-based week number to a 0-based variant index (cycles 0-5).
 * @param {number} weekNumber - 1-based week number
 * @returns {number} variant index 0-5
 */
function weekToVariant(weekNumber) {
  return ((weekNumber - 1) % VARIANTS_PER_LEVEL + VARIANTS_PER_LEVEL) % VARIANTS_PER_LEVEL;
}

module.exports = {
  createTheme,
  weekToVariant,
  VALID_SUBJECTS,
  VALID_YEAR_LEVELS,
  VARIANTS_PER_LEVEL,
  normalizeLessonTargets,
  sanitizeTeacherNotes,
  appendSourcesToNotes,
  composeNotes,
  getGradeBand,
  getGradeSizes,
  ROUTINES,
  ROUTINE_LABELS,
};
