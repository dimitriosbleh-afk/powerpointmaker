"use strict";

// ============================================================================
// Numeracy Palette Collection
// 5 year levels x 6 colour-family variants = 30 palettes
// Pure data — no functions, no imports.
//
// Variant families:
//   [0] Blueprint Grid   — deep navy / steel blue / bright teal / red alerts
//   [1] Forest Calculation — dark green / teal / amber / crimson
//   [2] Slate & Copper    — dark slate / charcoal blue / copper-orange / dark red
//   [3] Ocean Logic        — deep ocean blue / aquamarine / gold / brick
//   [4] Graphite & Lime   — dark graphite / forest / lime-gold / berry
//   [5] Cobalt Precision  — cobalt / purple-grey / bronze / maroon
//
// Contrast guarantees (WCAG AA >= 4.5:1):
//   WHITE on PRIMARY / SECONDARY / ACCENT / ALERT / SUCCESS / ASSESS
//   CHARCOAL on BG_LIGHT / BG_CARD
//   TEXT_ON_DARK on BG_DARK
//
// Year-level progression: foundation (most saturated) → grade56 (most muted)
// ============================================================================

const palettes = {

  // ──────────────────────────────────────────────────────────────────────────
  // FOUNDATION — Boldest, most saturated colours. Fun and energetic.
  // Fonts: Arial Black / Calibri
  // ──────────────────────────────────────────────────────────────────────────
  foundation: [
    // [0] Blueprint Grid — deep navy / steel blue / bright teal / red
    {
      PRIMARY: "0A2463", SECONDARY: "1E5078", ACCENT: "0E7C7B",
      ALERT: "B71C1C", SUCCESS: "1A6B3C", ASSESS: "3D2C8D",
      BG_DARK: "0A2463", BG_LIGHT: "F0F4FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "D0DFEF", SUBTITLE: "A8C0D8",
      DECOR_1: "1E5078", DECOR_2: "0E7C7B",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Forest Calculation — dark green / teal / amber / crimson
    {
      PRIMARY: "1B4332", SECONDARY: "0B6E4F", ACCENT: "8B6914",
      ALERT: "8B1A1A", SUCCESS: "1A7A3C", ASSESS: "4A2070",
      BG_DARK: "1B4332", BG_LIGHT: "EFF5F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C8E8D0", SUBTITLE: "A0D0B0",
      DECOR_1: "0B6E4F", DECOR_2: "8B6914",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Slate & Copper — dark slate / charcoal blue / copper-orange / dark red
    {
      PRIMARY: "2D3748", SECONDARY: "1A4A6B", ACCENT: "A0522D",
      ALERT: "9B1B30", SUCCESS: "1B6B40", ASSESS: "5B2D7A",
      BG_DARK: "2D3748", BG_LIGHT: "F0F2F5", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "CBD5E0", SUBTITLE: "A0B0C8",
      DECOR_1: "1A4A6B", DECOR_2: "A0522D",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ocean Logic — deep ocean blue / aquamarine / gold / brick
    {
      PRIMARY: "063B5E", SECONDARY: "0C6B6B", ACCENT: "7A6020",
      ALERT: "A03020", SUCCESS: "1A6B48", ASSESS: "3B2080",
      BG_DARK: "063B5E", BG_LIGHT: "EEF4F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C0D8E8", SUBTITLE: "90C0D8",
      DECOR_1: "0C6B6B", DECOR_2: "7A6020",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Graphite & Lime — dark graphite / forest / lime-gold / berry
    {
      PRIMARY: "2A2E35", SECONDARY: "2D5A3E", ACCENT: "7A7A14",
      ALERT: "7A2050", SUCCESS: "1A7040", ASSESS: "4B3070",
      BG_DARK: "2A2E35", BG_LIGHT: "F0F2F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C8D0C8", SUBTITLE: "A0B0A0",
      DECOR_1: "2D5A3E", DECOR_2: "7A7A14",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Cobalt Precision — cobalt / purple-grey / bronze / maroon
    {
      PRIMARY: "1A3A8A", SECONDARY: "4A3A68", ACCENT: "8A6B28",
      ALERT: "6B1A30", SUCCESS: "1A6B44", ASSESS: "2C2080",
      BG_DARK: "1A3A8A", BG_LIGHT: "F0F2FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C8D0F0", SUBTITLE: "A0B0E0",
      DECOR_1: "4A3A68", DECOR_2: "8A6B28",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ──────────────────────────────────────────────────────────────────────────
  // GRADE 1 — Bold, slightly softer than Foundation.
  // Fonts: Arial Black / Calibri
  // ──────────────────────────────────────────────────────────────────────────
  grade1: [
    // [0] Blueprint Grid
    {
      PRIMARY: "0D2A6B", SECONDARY: "205880", ACCENT: "107A78",
      ALERT: "B52020", SUCCESS: "1C6E3E", ASSESS: "402E8A",
      BG_DARK: "0D2A6B", BG_LIGHT: "F1F5FB", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "CDDDEC", SUBTITLE: "A5BDD5",
      DECOR_1: "205880", DECOR_2: "107A78",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Forest Calculation
    {
      PRIMARY: "1E4836", SECONDARY: "0E7254", ACCENT: "8A6B16",
      ALERT: "8C1C1C", SUCCESS: "1C7C3E", ASSESS: "4C2272",
      BG_DARK: "1E4836", BG_LIGHT: "F0F6F1", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C5E5CD", SUBTITLE: "9ECEAD",
      DECOR_1: "0E7254", DECOR_2: "8A6B16",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Slate & Copper
    {
      PRIMARY: "303B4C", SECONDARY: "1C4E6E", ACCENT: "9E5430",
      ALERT: "9A1D32", SUCCESS: "1D6E42", ASSESS: "5D2F7C",
      BG_DARK: "303B4C", BG_LIGHT: "F1F3F6", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C8D2DD", SUBTITLE: "9EADC5",
      DECOR_1: "1C4E6E", DECOR_2: "9E5430",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ocean Logic
    {
      PRIMARY: "084062", SECONDARY: "0E6E6E", ACCENT: "7C6222",
      ALERT: "9E3222", SUCCESS: "1C6E4A", ASSESS: "3E2282",
      BG_DARK: "084062", BG_LIGHT: "EFF5F9", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "BDD5E5", SUBTITLE: "8EBDD5",
      DECOR_1: "0E6E6E", DECOR_2: "7C6222",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Graphite & Lime
    {
      PRIMARY: "2C3038", SECONDARY: "2F5E42", ACCENT: "7C7C16",
      ALERT: "7C2252", SUCCESS: "1C7242", ASSESS: "4D3272",
      BG_DARK: "2C3038", BG_LIGHT: "F1F3F1", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C5CDC5", SUBTITLE: "9EAD9E",
      DECOR_1: "2F5E42", DECOR_2: "7C7C16",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Cobalt Precision
    {
      PRIMARY: "1C3E8C", SECONDARY: "4C3E6A", ACCENT: "8C6D2A",
      ALERT: "6D1C32", SUCCESS: "1C6E46", ASSESS: "2E2282",
      BG_DARK: "1C3E8C", BG_LIGHT: "F1F3FB", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C5CDED", SUBTITLE: "9EADDD",
      DECOR_1: "4C3E6A", DECOR_2: "8C6D2A",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ──────────────────────────────────────────────────────────────────────────
  // GRADE 2 — Transitional. Noticeable reduction in saturation.
  // Fonts: Trebuchet MS / Calibri
  // ──────────────────────────────────────────────────────────────────────────
  grade2: [
    // [0] Blueprint Grid
    {
      PRIMARY: "12306E", SECONDARY: "256080", ACCENT: "157875",
      ALERT: "B02525", SUCCESS: "1E7040", ASSESS: "453088",
      BG_DARK: "12306E", BG_LIGHT: "F2F6FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "CADAE9", SUBTITLE: "A2BAD2",
      DECOR_1: "256080", DECOR_2: "157875",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Forest Calculation
    {
      PRIMARY: "22503C", SECONDARY: "127858", ACCENT: "886E1A",
      ALERT: "8E2020", SUCCESS: "1E8040", ASSESS: "502674",
      BG_DARK: "22503C", BG_LIGHT: "F1F7F2", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C2E2CA", SUBTITLE: "9CCAAA",
      DECOR_1: "127858", DECOR_2: "886E1A",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Slate & Copper
    {
      PRIMARY: "344050", SECONDARY: "205270", ACCENT: "9C5835",
      ALERT: "982035", SUCCESS: "207044", ASSESS: "60327E",
      BG_DARK: "344050", BG_LIGHT: "F2F4F7", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C5CFDA", SUBTITLE: "9CAAC2",
      DECOR_1: "205270", DECOR_2: "9C5835",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ocean Logic
    {
      PRIMARY: "0B4568", SECONDARY: "107270", ACCENT: "7E6425",
      ALERT: "9C3425", SUCCESS: "1E7050", ASSESS: "422485",
      BG_DARK: "0B4568", BG_LIGHT: "F0F6FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "BAD2E2", SUBTITLE: "8CBAD2",
      DECOR_1: "107270", DECOR_2: "7E6425",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Graphite & Lime
    {
      PRIMARY: "30343C", SECONDARY: "326245", ACCENT: "7E7E1A",
      ALERT: "7E2455", SUCCESS: "1E7545", ASSESS: "503575",
      BG_DARK: "30343C", BG_LIGHT: "F2F4F2", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C2CAC2", SUBTITLE: "9CAA9C",
      DECOR_1: "326245", DECOR_2: "7E7E1A",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Cobalt Precision
    {
      PRIMARY: "20428E", SECONDARY: "50426C", ACCENT: "8E702C",
      ALERT: "701E35", SUCCESS: "1E7048", ASSESS: "322485",
      BG_DARK: "20428E", BG_LIGHT: "F2F4FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C2CAEA", SUBTITLE: "9CAADA",
      DECOR_1: "50426C", DECOR_2: "8E702C",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ──────────────────────────────────────────────────────────────────────────
  // GRADE 3/4 — Balanced, some sophistication. More muted than Grade 2.
  // Fonts: Trebuchet MS / Calibri
  // ──────────────────────────────────────────────────────────────────────────
  grade34: [
    // [0] Blueprint Grid
    {
      PRIMARY: "163572", SECONDARY: "2A6585", ACCENT: "1A7570",
      ALERT: "AA2828", SUCCESS: "207242", ASSESS: "4A3585",
      BG_DARK: "163572", BG_LIGHT: "F3F7FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C8D8E6", SUBTITLE: "A0B8CF",
      DECOR_1: "2A6585", DECOR_2: "1A7570",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Forest Calculation
    {
      PRIMARY: "265540", SECONDARY: "167C5C", ACCENT: "85701E",
      ALERT: "902222", SUCCESS: "208242", ASSESS: "542878",
      BG_DARK: "265540", BG_LIGHT: "F2F8F3", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C0DFC8", SUBTITLE: "9AC6A8",
      DECOR_1: "167C5C", DECOR_2: "85701E",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Slate & Copper
    {
      PRIMARY: "384555", SECONDARY: "255575", ACCENT: "985C38",
      ALERT: "952238", SUCCESS: "227245", ASSESS: "643580",
      BG_DARK: "384555", BG_LIGHT: "F3F5F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C2CCD5", SUBTITLE: "9AA8BF",
      DECOR_1: "255575", DECOR_2: "985C38",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ocean Logic
    {
      PRIMARY: "104A6E", SECONDARY: "147575", ACCENT: "806828",
      ALERT: "9A3628", SUCCESS: "207250", ASSESS: "482888",
      BG_DARK: "104A6E", BG_LIGHT: "F1F7FB", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "B8D0E0", SUBTITLE: "8AB8CF",
      DECOR_1: "147575", DECOR_2: "806828",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Graphite & Lime
    {
      PRIMARY: "353840", SECONDARY: "366548", ACCENT: "7B801E",
      ALERT: "802658", SUCCESS: "207848", ASSESS: "553878",
      BG_DARK: "353840", BG_LIGHT: "F3F5F3", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C0C8C0", SUBTITLE: "9AA89A",
      DECOR_1: "366548", DECOR_2: "7B801E",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Cobalt Precision
    {
      PRIMARY: "244890", SECONDARY: "554570", ACCENT: "907230",
      ALERT: "722038", SUCCESS: "207250", ASSESS: "352888",
      BG_DARK: "244890", BG_LIGHT: "F3F5FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C0C8E8", SUBTITLE: "9AA8D8",
      DECOR_1: "554570", DECOR_2: "907230",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ──────────────────────────────────────────────────────────────────────────
  // GRADE 5/6 — Most muted and sophisticated. Dark but refined.
  // Fonts: Georgia / Calibri
  // ──────────────────────────────────────────────────────────────────────────
  grade56: [
    // [0] Blueprint Grid
    {
      PRIMARY: "1A3875", SECONDARY: "2F6A88", ACCENT: "1E726D",
      ALERT: "A52A2A", SUCCESS: "227545", ASSESS: "4E3882",
      BG_DARK: "1A3875", BG_LIGHT: "F4F8FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C5D5E4", SUBTITLE: "9EB5CC",
      DECOR_1: "2F6A88", DECOR_2: "1E726D",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Forest Calculation
    {
      PRIMARY: "2A5A45", SECONDARY: "1A8060", ACCENT: "827220",
      ALERT: "922525", SUCCESS: "228545", ASSESS: "582A7A",
      BG_DARK: "2A5A45", BG_LIGHT: "F3F8F4", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "BDDCC5", SUBTITLE: "98C4A5",
      DECOR_1: "1A8060", DECOR_2: "827220",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Slate & Copper
    {
      PRIMARY: "3C4A5A", SECONDARY: "2A5A78", ACCENT: "95603C",
      ALERT: "92253A", SUCCESS: "247548", ASSESS: "683882",
      BG_DARK: "3C4A5A", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "C0CAD2", SUBTITLE: "98A5BC",
      DECOR_1: "2A5A78", DECOR_2: "95603C",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ocean Logic
    {
      PRIMARY: "145072", SECONDARY: "187878", ACCENT: "826C2C",
      ALERT: "98382C", SUCCESS: "227555", ASSESS: "4C2A8A",
      BG_DARK: "145072", BG_LIGHT: "F2F8FB", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "B5CDDD", SUBTITLE: "88B5CC",
      DECOR_1: "187878", DECOR_2: "826C2C",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Graphite & Lime
    {
      PRIMARY: "3A3D45", SECONDARY: "3A6A4C", ACCENT: "636E1A",
      ALERT: "82285A", SUCCESS: "227A4A", ASSESS: "583A7A",
      BG_DARK: "3A3D45", BG_LIGHT: "F4F6F4", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "BDC5BD", SUBTITLE: "98A598",
      DECOR_1: "3A6A4C", DECOR_2: "788520",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Cobalt Precision
    {
      PRIMARY: "284C92", SECONDARY: "5A4872", ACCENT: "846A2C",
      ALERT: "74223A", SUCCESS: "227552", ASSESS: "382A8A",
      BG_DARK: "284C92", BG_LIGHT: "F4F6FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B8B", TEXT_ON_DARK: "BDC5E5", SUBTITLE: "98A5D5",
      DECOR_1: "5A4872", DECOR_2: "927535",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],
};

module.exports = { palettes };
