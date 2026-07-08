"use strict";

/**
 * Grade-band sizing and density rules.
 *
 * Maps the megaprompt's age-band rules onto concrete sizing constants used
 * by every slide builder.
 *
 * Three bands:
 *   - F   : Foundation (yearLevel "foundation")
 *   - Y12 : Years 1-2 (yearLevel "grade1", "grade2")
 *   - Y36 : Years 3-6 (yearLevel "grade34", "grade56")
 *
 * Megaprompt §16 hard floors honoured by this table:
 *   F-Y2  : main task >= 44 pt, support text >= 30 pt
 *   Y3-Y4 : main task >= 38 pt, support text >= 26 pt
 *   Y5-Y6 : main task >= 34 pt, support text >= 24 pt, avoid below 20 pt
 *
 * Foundation sits above the F-Y2 floor (heroQuestion 48, body 32). Y12
 * meets the F-Y2 main floor exactly (44/30). Y36 meets the §16 main floor
 * (heroQuestion 34); body and bodyDense (22/19) sit just under the Y5-Y6
 * 24 pt support floor as a practical compromise that still honours the
 * "avoid below 20 pt" rule for body content. Anything below 19 pt is
 * reserved for chips/captions/footers where smaller type is acceptable.
 *
 * The slide canvas is fixed at 10" x 5.625", so "as large as practical"
 * is the operating principle. Card metrics in base.js auto-grow to fit
 * these sizes; if a card runs out of room the metrics function steps the
 * font down to bodyDense, then to baseTight (bodyDense * _shrink) which is
 * the floor for genuinely dense legacy content.
 */

/**
 * Map a yearLevel string to one of the three grade bands.
 * @param {string} yearLevel  "foundation"|"grade1"|"grade2"|"grade34"|"grade56"
 * @returns {string}          "F"|"Y12"|"Y36"
 */
function getGradeBand(yearLevel) {
  const lvl = String(yearLevel || "").toLowerCase();
  if (lvl === "foundation") return "F";
  if (lvl === "grade1" || lvl === "grade2") return "Y12";
  return "Y36";
}

/**
 * Sizing tables. Every size is a font point unless noted.
 *
 * Naming:
 *   titleH1        slide title text (top-of-slide heading)
 *   titleHero      large dark title-slide hero text (lesson opener)
 *   subtitleHero   subtitle on title slide
 *   metaHero       meta line on title slide
 *   closingHero    closing-slide review heading
 *   closingPrompt  reflection prompt text
 *   takeaway       closing-slide takeaway bullet
 *   liHeader       LI/SC card header label
 *   liBody         LI/SC body text
 *   heroQuestion   the main "thing students attend to" — CFU question, hero word
 *   hero           secondary hero (vocab word, named focus)
 *   body           main bullet/body text
 *   bodyDense      body text when card is dense
 *   sectionLabel   "Definition", "Example", "Hypothesis" style header
 *   chip           small chip / pill text (page refs, micro-labels)
 *   caption        captions, sources, tiny label rows
 *   mockupText     text inside visual-mockup previews (poster previews)
 *   badge          top-left coloured badge text
 *   footer         footer line
 *   stepNumber     numbered step prefix in numeracy / inquiry
 *   quote          quote-slide italic body
 *   featureLabel   feature key-card label inside annotatedModelSlide
 *   featureDetail  feature key-card detail inside annotatedModelSlide
 *
 * Density caps:
 *   maxBullets     soft max items per bullet card before triggering "dense"
 *   maxQuestions   max questions per slide before splitting
 *   maxPrompts     max prompts in an instruction card
 */
const SIZES = {
  F: {
    titleH1:       32,
    titleHero:     52,
    subtitleHero:  30,
    metaHero:      18,
    closingHero:   46,
    closingPrompt: 30,
    takeaway:      24,
    liHeader:      22,
    liBody:        26,
    heroQuestion:  48,
    hero:          44,
    body:          32,
    bodyDense:     26,
    sectionLabel:  22,
    chip:          14,
    caption:       12,
    mockupText:    12,
    badge:         14,
    footer:        11,
    stepNumber:    26,
    quote:         30,
    featureLabel:  15,
    featureDetail: 14,
    maxBullets:    3,
    maxQuestions:  1,
    maxPrompts:    2,
  },
  Y12: {
    titleH1:       30,
    titleHero:     48,
    subtitleHero:  28,
    metaHero:      16,
    closingHero:   42,
    closingPrompt: 26,
    takeaway:      22,
    liHeader:      20,
    liBody:        24,
    heroQuestion:  44,
    hero:          38,
    body:          30,
    bodyDense:     24,
    sectionLabel:  20,
    chip:          13,
    caption:       11,
    mockupText:    11,
    badge:         13,
    footer:        10,
    stepNumber:    24,
    quote:         26,
    featureLabel:  14,
    featureDetail: 13,
    maxBullets:    4,
    maxQuestions:  1,
    maxPrompts:    3,
  },
  Y36: {
    titleH1:       28,
    titleHero:     42,
    subtitleHero:  24,
    metaHero:      14,
    closingHero:   38,
    closingPrompt: 22,
    takeaway:      17,
    liHeader:      17,
    liBody:        19,
    heroQuestion:  34,
    hero:          30,
    body:          22,
    bodyDense:     19,
    sectionLabel:  14,
    chip:          12,
    caption:       10.5,
    mockupText:    10.5,
    badge:         11,
    footer:        9.5,
    stepNumber:    18,
    quote:         20,
    featureLabel:  13,
    featureDetail: 12.5,
    maxBullets:    6,
    maxQuestions:  3,
    maxPrompts:    5,
  },
};

/**
 * Multiplier used by density helpers to compute "shrink one notch" sizes.
 * Each band expects different content density, so the shrink ratio differs.
 */
const SHRINK = {
  F:   0.88,
  Y12: 0.86,
  Y36: 0.85,
};

// Cache per band so all callers share a frozen reference (cheap === checks,
// no per-createTheme allocation).
const _CACHE = {};

/**
 * Get the size table for a given grade band. Returns the same frozen object
 * for repeat calls with the same band.
 *
 * @param {string} gradeBand  "F"|"Y12"|"Y36"
 * @returns {object}          frozen sizes object
 */
function getGradeSizes(gradeBand) {
  const band = SIZES[gradeBand] ? gradeBand : "Y36";
  if (!_CACHE[band]) {
    _CACHE[band] = Object.freeze({ ...SIZES[band], _band: band, _shrink: SHRINK[band] });
  }
  return _CACHE[band];
}

/**
 * Single source of truth for the Y36 fallback. Builders that haven't been
 * passed an S table fall back to this rather than maintaining their own
 * partial copies (which silently drift).
 */
const DEFAULT_SIZES = getGradeSizes("Y36");

/**
 * Pick a value by band. Replaces the repeated triple-ternary
 * `sz._band === "F" ? f : sz._band === "Y12" ? y12 : y36` pattern.
 *
 * @param {object} sz   Sizes table (must have `_band`)
 * @param {*} f         Value for Foundation band
 * @param {*} y12       Value for Year 1-2 band
 * @param {*} y36       Value for Year 3-6 band
 */
function byBand(sz, f, y12, y36) {
  const b = sz && sz._band;
  return b === "F" ? f : b === "Y12" ? y12 : y36;
}

module.exports = {
  getGradeBand,
  getGradeSizes,
  DEFAULT_SIZES,
  byBand,
};
