"use strict";

/**
 * Literacy colour palettes — pure data, no functions, no imports.
 *
 * 5 year levels x 6 colour-family variants = 30 palettes.
 * Every colour passes WCAG AA contrast (>= 4.5:1) for its intended pairing:
 *   - WHITE (FFFFFF) on PRIMARY, SECONDARY, ACCENT, ALERT, SUCCESS, ASSESS
 *   - CHARCOAL on BG_LIGHT, BG_CARD, WHITE
 *   - TEXT_ON_DARK on BG_DARK
 *
 * Variant families (consistent across all year levels):
 *   [0] Midnight Scholar — deep blue / slate / dark gold / crimson
 *   [1] Plum & Honey    — plum / teal / dark honey / coral
 *   [2] Olive & Parchment — deep olive / burgundy / dark gold / slate
 *   [3] Ink & Paper     — navy ink / charcoal blue / copper / dark red
 *   [4] Autumn Library   — deep brown / forest green / amber / maroon
 *   [5] Twilight Pages   — deep indigo / dusty rose / dark gold / teal
 */

const palettes = {

  // ─────────────────────────────────────────────────────────────────────────
  // FOUNDATION — Boldest, most saturated. Arial Black / Calibri.
  // ─────────────────────────────────────────────────────────────────────────
  foundation: [
    // [0] Midnight Scholar
    {
      PRIMARY: "0E2258", SECONDARY: "284470", ACCENT: "725828", ALERT: "851020",
      SUCCESS: "155A34", ASSESS: "452898",
      BG_DARK: "0E2258", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B8A99", TEXT_ON_DARK: "DCD0C0", SUBTITLE: "C8BEA0",
      DECOR_1: "284470", DECOR_2: "725828",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Plum & Honey
    {
      PRIMARY: "3A1A60", SECONDARY: "125C68", ACCENT: "685020", ALERT: "7A2215",
      SUCCESS: "126838", ASSESS: "581A60",
      BG_DARK: "3A1A60", BG_LIGHT: "F9F4F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B8099", TEXT_ON_DARK: "D0C8D8", SUBTITLE: "C0B0C8",
      DECOR_1: "125C68", DECOR_2: "685020",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Olive & Parchment
    {
      PRIMARY: "243A1C", SECONDARY: "581A2C", ACCENT: "705018", ALERT: "622015",
      SUCCESS: "12623C", ASSESS: "481E65",
      BG_DARK: "243A1C", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B6B", TEXT_ON_DARK: "D8D0B8", SUBTITLE: "C8C0A0",
      DECOR_1: "581A2C", DECOR_2: "705018",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ink & Paper
    {
      PRIMARY: "0E1438", SECONDARY: "1A2840", ACCENT: "703418", ALERT: "680000",
      SUCCESS: "124A38", ASSESS: "381A60",
      BG_DARK: "0E1438", BG_LIGHT: "F5F2EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D0C8B8", SUBTITLE: "C0B8A0",
      DECOR_1: "1A2840", DECOR_2: "703418",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Autumn Library
    {
      PRIMARY: "281614", SECONDARY: "1A4A1C", ACCENT: "704818", ALERT: "480E1C",
      SUCCESS: "245838", ASSESS: "482038",
      BG_DARK: "281614", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7B6B", TEXT_ON_DARK: "DCD0BC", SUBTITLE: "CCC0A8",
      DECOR_1: "1A4A1C", DECOR_2: "704818",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Twilight Pages
    {
      PRIMARY: "1C0C44", SECONDARY: "582848", ACCENT: "624E20", ALERT: "851020",
      SUCCESS: "145840", ASSESS: "3A1A60",
      BG_DARK: "1C0C44", BG_LIGHT: "F6F2F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B7099", TEXT_ON_DARK: "D4C8E0", SUBTITLE: "C4B8D0",
      DECOR_1: "582848", DECOR_2: "624E20",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ─────────────────────────────────────────────────────────────────────────
  // GRADE 1 — Bold but slightly softer than Foundation. Arial Black / Calibri.
  // ─────────────────────────────────────────────────────────────────────────
  grade1: [
    // [0] Midnight Scholar
    {
      PRIMARY: "102450", SECONDARY: "2E4A74", ACCENT: "7A6028", ALERT: "881525",
      SUCCESS: "1A6038", ASSESS: "4E2BA0",
      BG_DARK: "102450", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B8A99", TEXT_ON_DARK: "E0D8C8", SUBTITLE: "D0C4A8",
      DECOR_1: "2E4A74", DECOR_2: "7A6028",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Plum & Honey
    {
      PRIMARY: "421E65", SECONDARY: "186270", ACCENT: "6E5820", ALERT: "822818",
      SUCCESS: "186C3E", ASSESS: "601E65",
      BG_DARK: "421E65", BG_LIGHT: "F9F4F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B8099", TEXT_ON_DARK: "D8CCE0", SUBTITLE: "C8B8D0",
      DECOR_1: "186270", DECOR_2: "6E5820",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Olive & Parchment
    {
      PRIMARY: "2A4020", SECONDARY: "601E30", ACCENT: "785820", ALERT: "6A2418",
      SUCCESS: "1A6640", ASSESS: "502468",
      BG_DARK: "2A4020", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B6B", TEXT_ON_DARK: "DCD4C0", SUBTITLE: "CCC4A8",
      DECOR_1: "601E30", DECOR_2: "785820",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ink & Paper
    {
      PRIMARY: "10183A", SECONDARY: "1E2E44", ACCENT: "783818", ALERT: "700000",
      SUCCESS: "184E3E", ASSESS: "401E65",
      BG_DARK: "10183A", BG_LIGHT: "F5F2EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D4CCC0", SUBTITLE: "C4BCB0",
      DECOR_1: "1E2E44", DECOR_2: "783818",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Autumn Library
    {
      PRIMARY: "2E1A18", SECONDARY: "1E4E20", ACCENT: "784E18", ALERT: "501020",
      SUCCESS: "285C3E", ASSESS: "502440",
      BG_DARK: "2E1A18", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7B6B", TEXT_ON_DARK: "E0D4C4", SUBTITLE: "D0C4B0",
      DECOR_1: "1E4E20", DECOR_2: "784E18",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Twilight Pages
    {
      PRIMARY: "200E4A", SECONDARY: "603050", ACCENT: "6A5420", ALERT: "881525",
      SUCCESS: "1A5E48", ASSESS: "421E65",
      BG_DARK: "200E4A", BG_LIGHT: "F6F2F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B7099", TEXT_ON_DARK: "D8CCE8", SUBTITLE: "C8BCD8",
      DECOR_1: "603050", DECOR_2: "6A5420",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ─────────────────────────────────────────────────────────────────────────
  // GRADE 2 — Transitional, moderately bold. Trebuchet MS / Calibri.
  // ─────────────────────────────────────────────────────────────────────────
  grade2: [
    // [0] Midnight Scholar
    {
      PRIMARY: "14264A", SECONDARY: "355278", ACCENT: "806828", ALERT: "8A1828",
      SUCCESS: "206540", ASSESS: "5530A0",
      BG_DARK: "14264A", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B8A99", TEXT_ON_DARK: "E4DED0", SUBTITLE: "D4CAB0",
      DECOR_1: "355278", DECOR_2: "806828",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Plum & Honey
    {
      PRIMARY: "4A2268", SECONDARY: "1E6878", ACCENT: "755E1E", ALERT: "8B2D20",
      SUCCESS: "1E7044", ASSESS: "682268",
      BG_DARK: "4A2268", BG_LIGHT: "F9F4F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B8099", TEXT_ON_DARK: "DCD0E4", SUBTITLE: "CCC0D4",
      DECOR_1: "1E6878", DECOR_2: "755E1E",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Olive & Parchment
    {
      PRIMARY: "2E4424", SECONDARY: "682434", ACCENT: "805E18", ALERT: "722820",
      SUCCESS: "206B44", ASSESS: "582870",
      BG_DARK: "2E4424", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B6B", TEXT_ON_DARK: "E0DCC8", SUBTITLE: "D0C8B0",
      DECOR_1: "682434", DECOR_2: "805E18",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ink & Paper
    {
      PRIMARY: "121C3E", SECONDARY: "223448", ACCENT: "803E1E", ALERT: "780000",
      SUCCESS: "1E5444", ASSESS: "48226A",
      BG_DARK: "121C3E", BG_LIGHT: "F5F2EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "D8D0C0", SUBTITLE: "C8C0B0",
      DECOR_1: "223448", DECOR_2: "803E1E",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Autumn Library
    {
      PRIMARY: "341E1C", SECONDARY: "245424", ACCENT: "80551A", ALERT: "581420",
      SUCCESS: "2E6044", ASSESS: "582848",
      BG_DARK: "341E1C", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7B6B", TEXT_ON_DARK: "E4D8C8", SUBTITLE: "D4C8B4",
      DECOR_1: "245424", DECOR_2: "80551A",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Twilight Pages
    {
      PRIMARY: "241050", SECONDARY: "683858", ACCENT: "725A20", ALERT: "8A1828",
      SUCCESS: "1E6450", ASSESS: "4A2268",
      BG_DARK: "241050", BG_LIGHT: "F6F2F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B7099", TEXT_ON_DARK: "DCD0EC", SUBTITLE: "CCC0DC",
      DECOR_1: "683858", DECOR_2: "725A20",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ─────────────────────────────────────────────────────────────────────────
  // GRADE 3/4 — Balanced, bold primaries with sophisticated accents.
  //             Georgia / Calibri. ~10% more saturated than Grade 5/6.
  // ─────────────────────────────────────────────────────────────────────────
  grade34: [
    // [0] Midnight Scholar
    {
      PRIMARY: "162840", SECONDARY: "3D5A7A", ACCENT: "8B7328", ALERT: "8B2030",
      SUCCESS: "266B4A", ASSESS: "5E35A0",
      BG_DARK: "162840", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B8A99", TEXT_ON_DARK: "E8E3DA", SUBTITLE: "D4C9B5",
      DECOR_1: "3D5A7A", DECOR_2: "8B7328",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Plum & Honey
    {
      PRIMARY: "52286A", SECONDARY: "246E7D", ACCENT: "7D6520", ALERT: "9C3520",
      SUCCESS: "267548", ASSESS: "70286A",
      BG_DARK: "52286A", BG_LIGHT: "F9F4F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B8099", TEXT_ON_DARK: "E0D8E8", SUBTITLE: "D0C0D8",
      DECOR_1: "246E7D", DECOR_2: "7D6520",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Olive & Parchment
    {
      PRIMARY: "324828", SECONDARY: "702838", ACCENT: "8A6E1E", ALERT: "7B3028",
      SUCCESS: "267048", ASSESS: "603070",
      BG_DARK: "324828", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B6B", TEXT_ON_DARK: "E4E0D4", SUBTITLE: "D4C9B0",
      DECOR_1: "702838", DECOR_2: "8A6E1E",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ink & Paper
    {
      PRIMARY: "152040", SECONDARY: "283850", ACCENT: "8B4520", ALERT: "800000",
      SUCCESS: "265848", ASSESS: "502870",
      BG_DARK: "152040", BG_LIGHT: "F5F2EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DDD8CC", SUBTITLE: "C8BEA8",
      DECOR_1: "283850", DECOR_2: "8B4520",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Autumn Library
    {
      PRIMARY: "382220", SECONDARY: "285828", ACCENT: "8B5E15", ALERT: "601828",
      SUCCESS: "326548", ASSESS: "603050",
      BG_DARK: "382220", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7B6B", TEXT_ON_DARK: "E8DDD0", SUBTITLE: "D4C4A8",
      DECOR_1: "285828", DECOR_2: "8B5E15",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Twilight Pages
    {
      PRIMARY: "281450", SECONDARY: "704060", ACCENT: "7B6222", ALERT: "8B2030",
      SUCCESS: "266858", ASSESS: "52286A",
      BG_DARK: "281450", BG_LIGHT: "F6F2F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B7099", TEXT_ON_DARK: "E0D8F0", SUBTITLE: "D0C0E0",
      DECOR_1: "704060", DECOR_2: "7B6222",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ─────────────────────────────────────────────────────────────────────────
  // GRADE 5/6 — Sophisticated, muted, literary. Georgia / Calibri.
  // ─────────────────────────────────────────────────────────────────────────
  grade56: [
    // [0] Midnight Scholar
    {
      PRIMARY: "1B2A3B", SECONDARY: "4A6583", ACCENT: "8B7328", ALERT: "8B2635",
      SUCCESS: "2D6A4F", ASSESS: "6B3FA0",
      BG_DARK: "1B2A3B", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B8A99", TEXT_ON_DARK: "E8E3DA", SUBTITLE: "D4C9B5",
      DECOR_1: "4A6583", DECOR_2: "8B7328",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Plum & Honey
    {
      PRIMARY: "5B2C6F", SECONDARY: "2A7B88", ACCENT: "856520", ALERT: "A63A28",
      SUCCESS: "2D7B4F", ASSESS: "7B2D6A",
      BG_DARK: "5B2C6F", BG_LIGHT: "F9F4F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B8099", TEXT_ON_DARK: "E0D8E8", SUBTITLE: "D0C0D8",
      DECOR_1: "2A7B88", DECOR_2: "856520",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Olive & Parchment
    {
      PRIMARY: "3A4F2D", SECONDARY: "7B2D3A", ACCENT: "8A6E1E", ALERT: "8B3A2D",
      SUCCESS: "2D6B4A", ASSESS: "6B3A7B",
      BG_DARK: "3A4F2D", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B6B", TEXT_ON_DARK: "E4E0D4", SUBTITLE: "D4C9B0",
      DECOR_1: "7B2D3A", DECOR_2: "8A6E1E",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ink & Paper
    {
      PRIMARY: "1A2744", SECONDARY: "2E4057", ACCENT: "A0522D", ALERT: "8B0000",
      SUCCESS: "2D5F4A", ASSESS: "5B3A7B",
      BG_DARK: "1A2744", BG_LIGHT: "F5F2EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DDD8CC", SUBTITLE: "C8BEA8",
      DECOR_1: "2E4057", DECOR_2: "A0522D",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Autumn Library
    {
      PRIMARY: "3E2723", SECONDARY: "2D5F2D", ACCENT: "996515", ALERT: "6B1D2A",
      SUCCESS: "3A6B4A", ASSESS: "6B3A5B",
      BG_DARK: "3E2723", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7B6B", TEXT_ON_DARK: "E8DDD0", SUBTITLE: "D4C4A8",
      DECOR_1: "2D5F2D", DECOR_2: "996515",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Twilight Pages
    {
      PRIMARY: "2C1654", SECONDARY: "7B4B6A", ACCENT: "7B6222", ALERT: "8B2635",
      SUCCESS: "2D6B5A", ASSESS: "5B2C6F",
      BG_DARK: "2C1654", BG_LIGHT: "F6F2F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B7099", TEXT_ON_DARK: "E0D8F0", SUBTITLE: "D0C0E0",
      DECOR_1: "7B4B6A", DECOR_2: "7B6222",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],
};

module.exports = { palettes };
