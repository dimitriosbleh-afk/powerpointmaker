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

//
// RETUNED (scripts/retune_palettes.js): every role colour is the brightest
// shade of its hue that still clears a per-band contrast target against
// white (Foundation ~4.9:1 for PRIMARY, rising to ~6.8:1 by Year 5/6;
// supporting roles sit ~0.6 deeper). BG_DARK is rich rather than near-black.
// Edit hues here freely, then re-run the script to re-establish the floors.
const palettes = {

  // ──────────────────────────────────────────────────────────────────────────
  // FOUNDATION — Boldest, most saturated colours. Fun and energetic.
  // Fonts: Arial Black / Calibri
  // ──────────────────────────────────────────────────────────────────────────
  foundation: [
    // [0] Blueprint Grid — deep navy / steel blue / bright teal / red
    {
      PRIMARY: "3067EA", SECONDARY: "296DA3", ACCENT: "0E7E7D",
      ALERT: "CD1F1F", SUCCESS: "1D7743", ASSESS: "6A55CA",
      BG_DARK: "1241B4", BG_LIGHT: "F0F4FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DFE9F4", SUBTITLE: "C3D5E5",
      DECOR_1: "296DA3", DECOR_2: "0E7E7D",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Forest Calculation — dark green / teal / amber / crimson
    {
      PRIMARY: "2F7E5C", SECONDARY: "0C7756", ACCENT: "926E15",
      ALERT: "C92626", SUCCESS: "1A783B", ASSESS: "8A44C9",
      BG_DARK: "225540", BG_LIGHT: "EFF5F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DAEFDF", SUBTITLE: "B5DBC1",
      DECOR_1: "0C7756", DECOR_2: "926E15",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Slate & Copper — dark slate / charcoal blue / copper-orange / dark red
    {
      PRIMARY: "4570BA", SECONDARY: "266D9E", ACCENT: "B25B32",
      ALERT: "C9233E", SUCCESS: "1E7747", ASSESS: "8D47BB",
      BG_DARK: "374C71", BG_LIGHT: "F0F2F5", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "E4E8EF", SUBTITLE: "CDD5E2",
      DECOR_1: "266D9E", DECOR_2: "B25B32",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ocean Logic — deep ocean blue / aquamarine / gold / brick
    {
      PRIMARY: "0C74B9", SECONDARY: "0D7474", ACCENT: "8C6E25",
      ALERT: "BE3926", SUCCESS: "1D7750", ASSESS: "744FD3",
      BG_DARK: "085080", BG_LIGHT: "EEF4F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "E2EDF4", SUBTITLE: "BCDAE8",
      DECOR_1: "0D7474", DECOR_2: "8C6E25",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Graphite & Lime — dark graphite / forest / lime-gold / berry
    {
      PRIMARY: "4570BA", SECONDARY: "2B7547", ACCENT: "767613",
      ALERT: "B93079", SUCCESS: "1B7644", ASSESS: "7E50BE",
      BG_DARK: "374C72", BG_LIGHT: "F0F2F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "E5E9E5", SUBTITLE: "CDD5CD",
      DECOR_1: "2B7547", DECOR_2: "767613",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Cobalt Precision — cobalt / purple-grey / bronze / maroon
    {
      PRIMARY: "3C6ADA", SECONDARY: "7953BF", ACCENT: "8E6E29",
      ALERT: "C12F57", SUCCESS: "1D774C", ASSESS: "6555D5",
      BG_DARK: "1F45A4", BG_LIGHT: "F0F2FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "E5E9F8", SUBTITLE: "C6CFED",
      DECOR_1: "7953BF", DECOR_2: "8E6E29",
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
      PRIMARY: "2963E5", SECONDARY: "266998", ACCENT: "117F7C",
      ALERT: "C42323", SUCCESS: "1D7240", ASSESS: "6851C5",
      BG_DARK: "1440A4", BG_LIGHT: "F1F5FB", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D8E5F0", SUBTITLE: "B9CCDE",
      DECOR_1: "266998", DECOR_2: "117F7C",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Forest Calculation
    {
      PRIMARY: "32795B", SECONDARY: "0E7254", ACCENT: "8C6D16",
      ALERT: "C12727", SUCCESS: "1A743A", ASSESS: "8741C6",
      BG_DARK: "22513D", BG_LIGHT: "F0F6F1", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D2EBD8", SUBTITLE: "AAD4B7",
      DECOR_1: "0E7254", DECOR_2: "8C6D16",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Slate & Copper
    {
      PRIMARY: "456DAA", SECONDARY: "266995", ACCENT: "AC5B34",
      ALERT: "C1243F", SUCCESS: "1E7244", ASSESS: "8945B7",
      BG_DARK: "33496A", BG_LIGHT: "F1F3F6", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DEE4EB", SUBTITLE: "C5CFDC",
      DECOR_1: "266995", DECOR_2: "AC5B34",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ocean Logic
    {
      PRIMARY: "0E71AD", SECONDARY: "0E7070", ACCENT: "8A6D26",
      ALERT: "B53927", SUCCESS: "1D724D", ASSESS: "714AD0",
      BG_DARK: "0A4C75", BG_LIGHT: "EFF5F9", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D7E5EF", SUBTITLE: "B2D2E2",
      DECOR_1: "0E7070", DECOR_2: "8A6D26",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Graphite & Lime
    {
      PRIMARY: "486BB0", SECONDARY: "2E7149", ACCENT: "767615",
      ALERT: "B23176", SUCCESS: "1C7242", ASSESS: "7A4EB7",
      BG_DARK: "35486F", BG_LIGHT: "F1F3F1", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "E0E5E0", SUBTITLE: "C5CFC5",
      DECOR_1: "2E7149", DECOR_2: "767615",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Cobalt Precision
    {
      PRIMARY: "3566D7", SECONDARY: "7251B8", ACCENT: "8C6D2A",
      ALERT: "B82F54", SUCCESS: "1D7249", ASSESS: "6050D1",
      BG_DARK: "1E4397", BG_LIGHT: "F1F3FB", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "E0E5F5", SUBTITLE: "C5CFEB",
      DECOR_1: "7251B8", DECOR_2: "8C6D2A",
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
      PRIMARY: "245FD9", SECONDARY: "276688", ACCENT: "167C79",
      ALERT: "B82727", SUCCESS: "1D6E3F", ASSESS: "684CC0",
      BG_DARK: "184093", BG_LIGHT: "F2F6FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D6E2EE", SUBTITLE: "B7C9DC",
      DECOR_1: "276688", DECOR_2: "167C79",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Forest Calculation
    {
      PRIMARY: "307256", SECONDARY: "106D50", ACCENT: "866C1A",
      ALERT: "B62929", SUCCESS: "1A6F38", ASSESS: "823EBD",
      BG_DARK: "22503C", BG_LIGHT: "F1F7F2", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "CFE8D6", SUBTITLE: "B2D6BD",
      DECOR_1: "106D50", DECOR_2: "866C1A",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Slate & Copper
    {
      PRIMARY: "456999", SECONDARY: "28668C", ACCENT: "A45C38",
      ALERT: "B5263F", SUCCESS: "1F6E43", ASSESS: "8143AA",
      BG_DARK: "304663", BG_LIGHT: "F2F4F7", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D7DEE5", SUBTITLE: "BBC5D5",
      DECOR_1: "28668C", DECOR_2: "A45C38",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ocean Logic
    {
      PRIMARY: "116DA4", SECONDARY: "0F6B69", ACCENT: "866A27",
      ALERT: "AC3929", SUCCESS: "1D6C4D", ASSESS: "6D43CC",
      BG_DARK: "0B486D", BG_LIGHT: "F0F6FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "CFE0EB", SUBTITLE: "A6C9DC",
      DECOR_1: "0F6B69", DECOR_2: "866A27",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Graphite & Lime
    {
      PRIMARY: "4967A2", SECONDARY: "306B48", ACCENT: "737318",
      ALERT: "AA3072", SUCCESS: "1C6D40", ASSESS: "744CAA",
      BG_DARK: "33456A", BG_LIGHT: "F2F4F2", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DADFDA", SUBTITLE: "C3CCC3",
      DECOR_1: "306B48", DECOR_2: "737318",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Cobalt Precision
    {
      PRIMARY: "2F61D0", SECONDARY: "6E4EAF", ACCENT: "886B2A",
      ALERT: "B02F53", SUCCESS: "1D6E47", ASSESS: "5E4BCE",
      BG_DARK: "20428E", BG_LIGHT: "F2F4FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DADFF3", SUBTITLE: "C3CCE8",
      DECOR_1: "6E4EAF", DECOR_2: "886B2A",
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
      PRIMARY: "265BC3", SECONDARY: "285F7D", ACCENT: "1A7772",
      ALERT: "AA2828", SUCCESS: "1D663B", ASSESS: "6347B3",
      BG_DARK: "1A3E85", BG_LIGHT: "F3F7FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "CFDDE9", SUBTITLE: "B5C8DA",
      DECOR_1: "285F7D", DECOR_2: "1A7772",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Forest Calculation
    {
      PRIMARY: "2F6A50", SECONDARY: "12664C", ACCENT: "7D691C",
      ALERT: "AB2828", SUCCESS: "196734", ASSESS: "7B3AAF",
      BG_DARK: "265540", BG_LIGHT: "F2F8F3", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DEEEE2", SUBTITLE: "BAD8C3",
      DECOR_1: "12664C", DECOR_2: "7D691C",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Slate & Copper
    {
      PRIMARY: "476283", SECONDARY: "295F83", ACCENT: "985C38",
      ALERT: "AA2740", SUCCESS: "1E663E", ASSESS: "7B419D",
      BG_DARK: "384555", BG_LIGHT: "F3F5F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D5DCE2", SUBTITLE: "BAC3D4",
      DECOR_1: "295F83", DECOR_2: "985C38",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ocean Logic
    {
      PRIMARY: "166596", SECONDARY: "116464", ACCENT: "806828",
      ALERT: "9E3729", SUCCESS: "1D6648", ASSESS: "6A3CC6",
      BG_DARK: "104A6E", BG_LIGHT: "F1F7FB", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D5E3ED", SUBTITLE: "AFCFDE",
      DECOR_1: "116464", DECOR_2: "806828",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Graphite & Lime
    {
      PRIMARY: "4D5F8E", SECONDARY: "356347", ACCENT: "6B6F1A",
      ALERT: "9D2F6C", SUCCESS: "1B663D", ASSESS: "6E489B",
      BG_DARK: "354162", BG_LIGHT: "F3F5F3", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D5DAD5", SUBTITLE: "BAC3BA",
      DECOR_1: "356347", DECOR_2: "6B6F1A",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Cobalt Precision
    {
      PRIMARY: "2E5CB9", SECONDARY: "664E90", ACCENT: "83672C",
      ALERT: "A42E50", SUCCESS: "1D6648", ASSESS: "5644C8",
      BG_DARK: "244890", BG_LIGHT: "F3F5FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "E5E8F5", SUBTITLE: "C9D0EB",
      DECOR_1: "664E90", DECOR_2: "83672C",
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
      PRIMARY: "2856B4", SECONDARY: "285A73", ACCENT: "1E726D",
      ALERT: "9F2828", SUCCESS: "1C6139", ASSESS: "5F449F",
      BG_DARK: "1C3C7D", BG_LIGHT: "F4F8FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "C5D5E4", SUBTITLE: "AABED2",
      DECOR_1: "285A73", DECOR_2: "1E726D",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Forest Calculation
    {
      PRIMARY: "2F644D", SECONDARY: "146048", ACCENT: "76671D",
      ALERT: "9E2828", SUCCESS: "196233", ASSESS: "7337A0",
      BG_DARK: "2A5A45", BG_LIGHT: "F3F8F4", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "EAF4ED", SUBTITLE: "C8E0CF",
      DECOR_1: "146048", DECOR_2: "76671D",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Slate & Copper
    {
      PRIMARY: "475B73", SECONDARY: "2A5A78", ACCENT: "8E5B39",
      ALERT: "9E283F", SUCCESS: "1E613C", ASSESS: "733E90",
      BG_DARK: "3C4A5A", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DEE3E7", SUBTITLE: "C8CFDC",
      DECOR_1: "2A5A78", DECOR_2: "8E5B39",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ocean Logic
    {
      PRIMARY: "185F88", SECONDARY: "135E5E", ACCENT: "786429",
      ALERT: "92362A", SUCCESS: "1C5F45", ASSESS: "6638B9",
      BG_DARK: "145072", BG_LIGHT: "F2F8FB", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "E1EBF1", SUBTITLE: "B8D3E0",
      DECOR_1: "135E5E", DECOR_2: "786429",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Graphite & Lime
    {
      PRIMARY: "4C5A7D", SECONDARY: "345E44", ACCENT: "616C1A",
      ALERT: "942D66", SUCCESS: "1B603A", ASSESS: "67448F",
      BG_DARK: "353E57", BG_LIGHT: "F4F6F4", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D2D7D2", SUBTITLE: "B9C1B9",
      DECOR_1: "345E44", DECOR_2: "616C1A",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Cobalt Precision
    {
      PRIMARY: "2E57A8", SECONDARY: "604C7B", ACCENT: "7C6429",
      ALERT: "992D4D", SUCCESS: "1C5F43", ASSESS: "503DC4",
      BG_DARK: "284C92", BG_LIGHT: "F4F6FC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "EAEDF6", SUBTITLE: "CFD5EB",
      DECOR_1: "604C7B", DECOR_2: "7C6429",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],
};

module.exports = { palettes };
