"use strict";

const PptxGenJS = require("pptxgenjs");

// ── Core modules ──
const { SLIDE_W, SLIDE_H, SAFE_RIGHT, SAFE_BOTTOM, CONTENT_TOP, validateBounds } = require("./core/layout");
const { hexToRgb, luminance, contrastRatio, getContrastColor, validateContrast } = require("./core/contrast");
const { iconToBase64Png }          = require("./core/icons");
const { normalizeLessonTargets, sanitizeTeacherNotes, appendSourcesToNotes, installNotesPatch, installSlideTextPatch } = require("./core/notes");
const { makeShadow, makeCardShadow } = require("./core/shadows");
const { createElements }           = require("./core/elements");
const { withReveal }               = require("./core/withReveal");
const { clickBuild }               = require("./core/animations");
const { createImageHelpers }       = require("./core/images");
const { warnIfSlideHasOverlaps, warnIfSlideElementsOutOfBounds, runSlideDiagnostics } = require("./core/diagnostics");
const { getGradeBand, getGradeSizes } = require("./core/gradeBand");
const { createRoutineHelpers, ROUTINES, ROUTINE_LABELS } = require("./core/routineIcons");
const { createPlaceholderHelpers } = require("./core/placeholders");
const { createManipulatives }      = require("./core/manipulatives");
const { composeNotes, composeGlanceNotes, composeRevealNotes } = require("./core/composeNotes");
const { derivedTones, mixHex, lightenHex, darkenHex } = require("./core/color");
const { createPictogramHelpers, PICTOGRAMS, listPictograms } = require("./core/pictograms");
const { createVisualSpec } = require("./core/visualSpec");

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
installSlideTextPatch(PptxGenJS);

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

  // Derived tints. Large fills (hero panels, question cards, option cards)
  // use the SOFT wash of a role colour so the slide stays calm; the strong
  // colour is reserved for badges, signals, reveals and small accents
  // (megaprompt sections 18a and 50). *_LINE is the matching hairline border.
  ["PRIMARY", "SECONDARY", "ACCENT", "ALERT", "SUCCESS", "ASSESS"].forEach((role) => {
    if (!C[role]) return;
    const tones = derivedTones(C[role], C.BG_LIGHT);
    C[`${role}_SOFT`] = tones.soft;
    C[`${role}_LINE`] = tones.line;
  });
  if (!C.ASSESS) { C.ASSESS = C.ALERT; C.ASSESS_SOFT = C.ALERT_SOFT; C.ASSESS_LINE = C.ALERT_LINE; }
  // A slightly lighter panel colour for use ON the dark title/closing background.
  C.BG_DARK_PANEL = mixHex(C.BG_DARK, "FFFFFF", 0.12);
  /** Soft wash of any hex (e.g. a stage colour a build script chose). */
  const softOf = (hex, amount) => mixHex(hex, C.BG_LIGHT, amount != null ? amount : 0.88);
  /** Hairline border tone of any hex. */
  const lineOf = (hex, amount) => mixHex(hex, C.BG_LIGHT, amount != null ? amount : 0.6);

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
  // Built-in pictograms (sync) and the declarative visual layer. Both are
  // created before the builders so slide builders can place icons and
  // fit visual specs without callers touching raw coordinates.
  const picto = createPictogramHelpers(C, FONT_B, el, S);
  const visual = createVisualSpec(C, FONT_H, FONT_B, el, S, { manips: createManipulatives(C, FONT_B, S), picto });

  const base = createBaseBuilders(C, FONT_H, FONT_B, el, shadowFn, S, {
    exitTicketTitle: subjectLower === "numeracy" ? "Stage 5  |  Show What You Know" : undefined,
    subject: subjectLower,
    picto,
    visual,
    softOf,
    lineOf,
  });

  // Build subject-specific slide builders
  const subjectFactory = SUBJECT_BUILDER_FACTORIES[subjectLower];
  const subjectBuilders = subjectFactory(C, FONT_H, FONT_B, el, S, { picto, visual, softOf, lineOf, subject: subjectLower });

  const routine = createRoutineHelpers(C, FONT_B, el, picto);
  const placeholders = createPlaceholderHelpers(C, FONT_H, FONT_B, el);
  // Visual-anchor manipulative helpers on EVERY theme (a literacy or science
  // lesson may still need a number line, chips or grouped counters).
  const manips = createManipulatives(C, FONT_B, S);

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

    // Colour tools for custom slides
    softOf,
    lineOf,
    mixHex,
    lightenHex,
    darkenHex,

    // Built-in pictograms (white glyph on a coloured circle or tile, or a
    // flat charcoal glyph) - the "visual built in" for vocabulary, science
    // stages, wellbeing feelings, launch hooks. Synchronous.
    ...picto,
    PICTOGRAMS,
    listPictograms,

    // Declarative visuals: drawVisual(slide, { type: "tensFrame", filled: 7 }, frame)
    ...visual,

    // Click-to-reveal. clickBuild is the preferred mechanism (one slide, one
    // element per click); withReveal duplicates the slide and is the fallback
    // for cases a click build cannot express. Megaprompt section 20b.
    clickBuild,
    withReveal,

    // Diagnostics
    warnIfSlideHasOverlaps,
    warnIfSlideElementsOutOfBounds,
    runSlideDiagnostics,

    // Manipulative / visual-anchor helpers (all subjects)
    ...manips,

    // Base slide builders (all subjects)
    ...base,

    // Subject-specific slide builders (may override base if name collides)
    ...subjectBuilders,

    ...routine,
    ROUTINES,
    ROUTINE_LABELS,
    ...placeholders,
    composeNotes,
    composeGlanceNotes,
    composeRevealNotes,

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
