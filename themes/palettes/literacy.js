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

//
// RETUNED (scripts/retune_palettes.js): every role colour is the brightest
// shade of its hue that still clears a per-band contrast target against
// white (Foundation ~4.9:1 for PRIMARY, rising to ~6.8:1 by Year 5/6;
// supporting roles sit ~0.6 deeper). BG_DARK is rich rather than near-black.
// Edit hues here freely, then re-run the script to re-establish the floors.
const palettes = {

  // ─────────────────────────────────────────────────────────────────────────
  // FOUNDATION — Boldest, most saturated. Arial Black / Calibri.
  // ─────────────────────────────────────────────────────────────────────────
  foundation: [
    // [0] Midnight Scholar
    {
      PRIMARY: "3B68E0", SECONDARY: "3D69AC", ACCENT: "8E6E32", ALERT: "CE1932",
      SUCCESS: "1C7745", ASSESS: "7150D1",
      BG_DARK: "1C43AE", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "717F8D", TEXT_ON_DARK: "EFE9E2", SUBTITLE: "DAD3BE",
      DECOR_1: "3D69AC", DECOR_2: "8E6E32",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Plum & Honey
    {
      PRIMARY: "8C52D0", SECONDARY: "177382", ACCENT: "8F6E2C", ALERT: "C03521",
      SUCCESS: "157740", ASSESS: "A731B6",
      BG_DARK: "642DA6", BG_LIGHT: "F9F4F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "71768D", TEXT_ON_DARK: "ECE8EF", SUBTITLE: "DACFDE",
      DECOR_1: "177382", DECOR_2: "8F6E2C",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Olive & Parchment
    {
      PRIMARY: "447E2F", SECONDARY: "BA375D", ACCENT: "966B20", ALERT: "BA3D28",
      SUCCESS: "167849", ASSESS: "9041C7",
      BG_DARK: "345428", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B6B", TEXT_ON_DARK: "EDE9DE", SUBTITLE: "DAD5BE",
      DECOR_1: "BA375D", DECOR_2: "966B20",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ink & Paper
    {
      PRIMARY: "5567D4", SECONDARY: "4067AC", ACCENT: "BA5628", ALERT: "D30000",
      SUCCESS: "1C7559", ASSESS: "8148CD",
      BG_DARK: "2C3FB0", BG_LIGHT: "F5F2EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "ECE8E2", SUBTITLE: "D5CFBE",
      DECOR_1: "4067AC", DECOR_2: "BA5628",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Autumn Library
    {
      PRIMARY: "B74F44", SECONDARY: "2A772D", ACCENT: "A06722", ALERT: "C6264D",
      SUCCESS: "2C7649", ASSESS: "AB3F80",
      BG_DARK: "733D37", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7B6B", TEXT_ON_DARK: "EFE9E0", SUBTITLE: "D7CFBB",
      DECOR_1: "2A772D", DECOR_2: "A06722",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Twilight Pages
    {
      PRIMARY: "7D55E1", SECONDARY: "AA3F86", ACCENT: "8A6E2D", ALERT: "CE1932",
      SUCCESS: "1B7555", ASSESS: "8548CD",
      BG_DARK: "5223C8", BG_LIGHT: "F6F2F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B7099", TEXT_ON_DARK: "EDE8F3", SUBTITLE: "D7CFDF",
      DECOR_1: "AA3F86", DECOR_2: "8A6E2D",
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
      PRIMARY: "3265D6", SECONDARY: "3F66A0", ACCENT: "8B6E2E", ALERT: "C41E35",
      SUCCESS: "1F7243", ASSESS: "7149CE",
      BG_DARK: "1E4496", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "717F8D", TEXT_ON_DARK: "EAE5DA", SUBTITLE: "D6CBB2",
      DECOR_1: "3F66A0", DECOR_2: "8B6E2E",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Plum & Honey
    {
      PRIMARY: "8C4DCA", SECONDARY: "1B6D7D", ACCENT: "8A6E28", ALERT: "B63822",
      SUCCESS: "197242", ASSESS: "A132AA",
      BG_DARK: "612C94", BG_LIGHT: "F9F4F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "71768D", TEXT_ON_DARK: "E8E0ED", SUBTITLE: "D5C8DB",
      DECOR_1: "1B6D7D", DECOR_2: "8A6E28",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Olive & Parchment
    {
      PRIMARY: "477731", SECONDARY: "B23859", ACCENT: "926B27", ALERT: "B33D28",
      SUCCESS: "1D7248", ASSESS: "8F40B9",
      BG_DARK: "345027", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B6B", TEXT_ON_DARK: "EAE5DA", SUBTITLE: "D2CBB2",
      DECOR_1: "B23859", DECOR_2: "926B27",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ink & Paper
    {
      PRIMARY: "4B64CD", SECONDARY: "3F669B", ACCENT: "B65524", ALERT: "CC0000",
      SUCCESS: "23715A", ASSESS: "8243C7",
      BG_DARK: "2A409A", BG_LIGHT: "F5F2EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "E8E4DE", SUBTITLE: "D1CBC1",
      DECOR_1: "3F669B", DECOR_2: "B65524",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Autumn Library
    {
      PRIMARY: "AD5047", SECONDARY: "2C732F", ACCENT: "9C651F", ALERT: "BF264C",
      SUCCESS: "2E704A", ASSESS: "A0417E",
      BG_DARK: "6D3A35", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7B6B", TEXT_ON_DARK: "EAE2D7", SUBTITLE: "D6CBB9",
      DECOR_1: "2C732F", DECOR_2: "9C651F",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Twilight Pages
    {
      PRIMARY: "794EDE", SECONDARY: "A04181", ACCENT: "8B6E2A", ALERT: "C41E35",
      SUCCESS: "207257", ASSESS: "8541C7",
      BG_DARK: "5124BC", BG_LIGHT: "F6F2F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B7099", TEXT_ON_DARK: "E8E0F1", SUBTITLE: "D5CBE1",
      DECOR_1: "A04181", DECOR_2: "8B6E2A",
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
      PRIMARY: "3463C0", SECONDARY: "3F628F", ACCENT: "846B29", ALERT: "BA2036",
      SUCCESS: "226D45", ASSESS: "6F44C7",
      BG_DARK: "234382", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "717F8D", TEXT_ON_DARK: "E4DED0", SUBTITLE: "D4CAB0",
      DECOR_1: "3F628F", DECOR_2: "846B29",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Plum & Honey
    {
      PRIMARY: "8C45C2", SECONDARY: "1E6878", ACCENT: "856B22", ALERT: "AC3828",
      SUCCESS: "1D6E43", ASSESS: "9C339C",
      BG_DARK: "602C87", BG_LIGHT: "F9F4F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "71768D", TEXT_ON_DARK: "E4DBEA", SUBTITLE: "D2C8D9",
      DECOR_1: "1E6878", DECOR_2: "856B22",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Olive & Parchment
    {
      PRIMARY: "477233", SECONDARY: "A73A53", ACCENT: "8D671A", ALERT: "A83B2F",
      SUCCESS: "216D45", ASSESS: "893EAE",
      BG_DARK: "304B24", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B6B", TEXT_ON_DARK: "E0DCC8", SUBTITLE: "D0C8B0",
      DECOR_1: "A73A53", DECOR_2: "8D671A",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ink & Paper
    {
      PRIMARY: "4260C8", SECONDARY: "3F638C", ACCENT: "B05529", ALERT: "C50000",
      SUCCESS: "276C58", ASSESS: "823DBF",
      BG_DARK: "283F8B", BG_LIGHT: "F5F2EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "E1DBCF", SUBTITLE: "CFC8B9",
      DECOR_1: "3F638C", DECOR_2: "B05529",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Autumn Library
    {
      PRIMARY: "9F4F47", SECONDARY: "2F6D2F", ACCENT: "97641F", ALERT: "B62942",
      SUCCESS: "306B4A", ASSESS: "96437A",
      BG_DARK: "6A3733", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7B6B", TEXT_ON_DARK: "E7DDCF", SUBTITLE: "D4C8B4",
      DECOR_1: "2F6D2F", DECOR_2: "97641F",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Twilight Pages
    {
      PRIMARY: "7548DA", SECONDARY: "95437A", ACCENT: "886B26", ALERT: "BA2036",
      SUCCESS: "206C56", ASSESS: "853DBB",
      BG_DARK: "4F23B0", BG_LIGHT: "F6F2F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B7099", TEXT_ON_DARK: "E4DBF0", SUBTITLE: "CCC0DC",
      DECOR_1: "95437A", DECOR_2: "886B26",
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
      PRIMARY: "35619B", SECONDARY: "3F5D7D", ACCENT: "7F6925", ALERT: "AA273B",
      SUCCESS: "246546", ASSESS: "6E3EBB",
      BG_DARK: "25436C", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "717F8D", TEXT_ON_DARK: "E8E3DA", SUBTITLE: "D4C9B5",
      DECOR_1: "3F5D7D", DECOR_2: "7F6925",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Plum & Honey
    {
      PRIMARY: "8641AD", SECONDARY: "20626F", ACCENT: "816821", ALERT: "A03621",
      SUCCESS: "21663F", ASSESS: "92348A",
      BG_DARK: "5C2D77", BG_LIGHT: "F9F4F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "71768D", TEXT_ON_DARK: "E0D8E8", SUBTITLE: "D0C0D8",
      DECOR_1: "20626F", DECOR_2: "816821",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Olive & Parchment
    {
      PRIMARY: "476838", SECONDARY: "9B374E", ACCENT: "82671C", ALERT: "9A3C32",
      SUCCESS: "236642", ASSESS: "804095",
      BG_DARK: "314726", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B6B", TEXT_ON_DARK: "E4E0D4", SUBTITLE: "D4C9B0",
      DECOR_1: "9B374E", DECOR_2: "82671C",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ink & Paper
    {
      PRIMARY: "3B5BB5", SECONDARY: "415B81", ACCENT: "A85327", ALERT: "B80000",
      SUCCESS: "2B6452", ASSESS: "7A3DAA",
      BG_DARK: "293E7C", BG_LIGHT: "F5F2EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DDD8CC", SUBTITLE: "CFC6B2",
      DECOR_1: "415B81", DECOR_2: "A85327",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Autumn Library
    {
      PRIMARY: "8A504B", SECONDARY: "2E662E", ACCENT: "8F6116", ALERT: "A72A46",
      SUCCESS: "316347", ASSESS: "85436F",
      BG_DARK: "5F3733", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7B6B", TEXT_ON_DARK: "E8DDD0", SUBTITLE: "D4C4A8",
      DECOR_1: "2E662E", DECOR_2: "8F6116",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Twilight Pages
    {
      PRIMARY: "7040CF", SECONDARY: "81466D", ACCENT: "816724", ALERT: "AA273B",
      SUCCESS: "256455", ASSESS: "7D3DA2",
      BG_DARK: "4E279B", BG_LIGHT: "F6F2F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B7099", TEXT_ON_DARK: "E0D8F0", SUBTITLE: "D0C0E0",
      DECOR_1: "81466D", DECOR_2: "816724",
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
      PRIMARY: "3C5D83", SECONDARY: "405771", ACCENT: "796423", ALERT: "9B2A3B",
      SUCCESS: "285F47", ASSESS: "6A3E9E",
      BG_DARK: "29405A", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "717F8D", TEXT_ON_DARK: "E8E3DA", SUBTITLE: "D4C9B5",
      DECOR_1: "405771", DECOR_2: "796423",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Plum & Honey
    {
      PRIMARY: "7F3D9B", SECONDARY: "205D67", ACCENT: "81621F", ALERT: "963424",
      SUCCESS: "235F3D", ASSESS: "8A3277",
      BG_DARK: "5B2C6F", BG_LIGHT: "F9F4F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "71768D", TEXT_ON_DARK: "E0D8E8", SUBTITLE: "D0C0D8",
      DECOR_1: "205D67", DECOR_2: "81621F",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Olive & Parchment
    {
      PRIMARY: "486238", SECONDARY: "913545", ACCENT: "7D641B", ALERT: "8D3B2E",
      SUCCESS: "285E41", ASSESS: "743F85",
      BG_DARK: "3A4F2D", BG_LIGHT: "F7F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7B6B", TEXT_ON_DARK: "E7E4D9", SUBTITLE: "D9CFB9",
      DECOR_1: "913545", DECOR_2: "7D641B",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Ink & Paper
    {
      PRIMARY: "3B599B", SECONDARY: "3E5675", ACCENT: "A0522D", ALERT: "AF0000",
      SUCCESS: "2C5D49", ASSESS: "69438E",
      BG_DARK: "293D6B", BG_LIGHT: "F5F2EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "5F6E7A", TEXT_ON_DARK: "DDD8CC", SUBTITLE: "C8BEA8",
      DECOR_1: "3E5675", DECOR_2: "A0522D",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Autumn Library
    {
      PRIMARY: "7C4E46", SECONDARY: "2D5F2D", ACCENT: "8C5C13", ALERT: "9B2A3D",
      SUCCESS: "335E41", ASSESS: "7A4268",
      BG_DARK: "563631", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7B6B", TEXT_ON_DARK: "E8DDD0", SUBTITLE: "D4C4A8",
      DECOR_1: "2D5F2D", DECOR_2: "8C5C13",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Twilight Pages
    {
      PRIMARY: "6C38CB", SECONDARY: "734663", ACCENT: "7D6423", ALERT: "9B2A3B",
      SUCCESS: "285E4F", ASSESS: "773A92",
      BG_DARK: "4C2691", BG_LIGHT: "F6F2F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7B7099", TEXT_ON_DARK: "E0D8F0", SUBTITLE: "D0C0E0",
      DECOR_1: "734663", DECOR_2: "7D6423",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],
};

module.exports = { palettes };
