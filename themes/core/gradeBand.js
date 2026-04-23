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
 * Megaprompt rules condensed:
 *   F     : 1 idea/slide, 1 question/slide, hugest text, picture-heavy
 *   Y12   : 1 question/slide, very large text, strong visuals
 *   Y36   : 1-2 questions preferred (3 max), large-but-denser text
 *
 * The slide canvas is fixed at 10" x 5.625", so "as large as practical"
 * is the operating principle for F/Y12 — these sizes leave room for the
 * title bar, badge, footer, and any accompanying visual.
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
    titleH1:       30,
    titleHero:     46,
    subtitleHero:  26,
    metaHero:      15,
    closingHero:   40,
    closingPrompt: 24,
    takeaway:      19,
    liHeader:      18,
    liBody:        18,
    heroQuestion:  34,
    hero:          32,
    body:          22,
    bodyDense:     19,
    sectionLabel:  16,
    chip:          14,
    caption:       12,
    mockupText:    11,
    badge:         13,
    footer:        11,
    stepNumber:    20,
    quote:         24,
    featureLabel:  12,
    featureDetail: 12,
    maxBullets:    3,
    maxQuestions:  1,
    maxPrompts:    2,
  },
  Y12: {
    titleH1:       28,
    titleHero:     42,
    subtitleHero:  24,
    metaHero:      14,
    closingHero:   38,
    closingPrompt: 21,
    takeaway:      16,
    liHeader:      16,
    liBody:        16,
    heroQuestion:  28,
    hero:          28,
    body:          19,
    bodyDense:     16.5,
    sectionLabel:  14,
    chip:          12.5,
    caption:       11,
    mockupText:    10,
    badge:         12,
    footer:        10,
    stepNumber:    18,
    quote:         20,
    featureLabel:  11,
    featureDetail: 10.8,
    maxBullets:    4,
    maxQuestions:  1,
    maxPrompts:    3,
  },
  Y36: {
    titleH1:       26,
    titleHero:     40,
    subtitleHero:  22,
    metaHero:      13,
    closingHero:   36,
    closingPrompt: 18,
    takeaway:      13,
    liHeader:      13,
    liBody:        13,
    heroQuestion:  20,
    hero:          26,
    body:          16.5,
    bodyDense:     14.5,
    sectionLabel:  11,
    chip:          11,
    caption:       9.6,
    mockupText:    9.5,
    badge:         10,
    footer:        9,
    stepNumber:    14,
    quote:         16,
    featureLabel:  9.6,
    featureDetail: 9.4,
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
