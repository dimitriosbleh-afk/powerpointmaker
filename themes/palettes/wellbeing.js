"use strict";

/**
 * Wellbeing colour palettes — pure data, no functions, no imports.
 *
 * 5 year levels x 6 colour-family variants = 30 palettes.
 * Every colour passes WCAG AA contrast (>= 4.5:1) for its intended pairing:
 *   - WHITE (FFFFFF) on PRIMARY, SECONDARY, ACCENT, ALERT, SUCCESS, ASSESS
 *   - CHARCOAL on BG_LIGHT, BG_CARD, WHITE
 *   - TEXT_ON_DARK on BG_DARK
 *
 * Wellbeing identity: warm, nurturing, safe, calming.
 * Think nature, growth, community. Greens, earthy tones, warm oranges, soft blues.
 *
 * Variant families (consistent across all year levels):
 *   [0] Forest Haven   — deep forest green / sage / warm gold / terracotta
 *   [1] Ocean Calm     — deep ocean blue / seafoam / warm amber / soft coral
 *   [2] Sunset Garden  — deep burgundy / olive / warm orange / earth brown
 *   [3] Mountain Air   — deep teal / lavender-grey / warm copper / russet
 *   [4] Harvest Gold   — deep brown / forest / golden amber / brick
 *   [5] Meadow Mist    — deep sage / plum / soft gold / clay
 *
 * Year-level progression: foundation (most saturated) → grade56 (most muted).
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
    // [0] Forest Haven
    {
      PRIMARY: "208142", SECONDARY: "3D752B", ACCENT: "916D13", ALERT: "B93D20",
      SUCCESS: "1B7737", ASSESS: "6A55CA",
      BG_DARK: "15582D", BG_LIGHT: "F9F6EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "628071", TEXT_ON_DARK: "DEEFD7", SUBTITLE: "B9DDAB",
      DECOR_1: "3D752B", DECOR_2: "916D13",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Ocean Calm
    {
      PRIMARY: "1674BD", SECONDARY: "1B7463", ACCENT: "926C16", ALERT: "AF4632",
      SUCCESS: "1B7752", ASSESS: "6E54C6",
      BG_DARK: "0F4F81", BG_LIGHT: "F5F8FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "627D8C", TEXT_ON_DARK: "DEEBF3", SUBTITLE: "B3D5E5",
      DECOR_1: "1B7463", DECOR_2: "926C16",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Sunset Garden
    {
      PRIMARY: "D02F5E", SECONDARY: "487328", ACCENT: "AA6014", ALERT: "925B27",
      SUCCESS: "2E762C", ASSESS: "A530C2",
      BG_DARK: "8F2041", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8B7878", TEXT_ON_DARK: "F4E5E9", SUBTITLE: "EBCDD5",
      DECOR_1: "487328", DECOR_2: "AA6014",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Mountain Air
    {
      PRIMARY: "177D7D", SECONDARY: "615BC3", ACCENT: "A8621D", ALERT: "B34328",
      SUCCESS: "1D7659", ASSESS: "7E4FC4",
      BG_DARK: "105858", BG_LIGHT: "F5F8F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "677D7D", TEXT_ON_DARK: "E2F1ED", SUBTITLE: "B9DCD5",
      DECOR_1: "615BC3", DECOR_2: "A8621D",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Harvest Gold
    {
      PRIMARY: "9D6133", SECONDARY: "2C7738", ACCENT: "986B12", ALERT: "B93D30",
      SUCCESS: "2C763E", ASSESS: "9F3FAA",
      BG_DARK: "6B4223", BG_LIGHT: "FAF6EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "807667", TEXT_ON_DARK: "EFE6DA", SUBTITLE: "DACFB9",
      DECOR_1: "2C7738", DECOR_2: "986B12",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Meadow Mist
    {
      PRIMARY: "2F7F47", SECONDARY: "A43DA4", ACCENT: "85711A", ALERT: "A84D37",
      SUCCESS: "2C7646", ASSESS: "854BBC",
      BG_DARK: "2A5838", BG_LIGHT: "F6F8F2", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "677D6E", TEXT_ON_DARK: "E5F1E5", SUBTITLE: "C3DFC3",
      DECOR_1: "A43DA4", DECOR_2: "85711A",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ─────────────────────────────────────────────────────────────────────────
  // GRADE 1 — Bold but slightly softer. Arial Black / Calibri.
  // ─────────────────────────────────────────────────────────────────────────
  grade1: [
    // [0] Forest Haven
    {
      PRIMARY: "237C43", SECONDARY: "3D702E", ACCENT: "8D6D1C", ALERT: "B33D2A",
      SUCCESS: "207239", ASSESS: "6A50C4",
      BG_DARK: "1A5C32", BG_LIGHT: "F9F6EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "67806E", TEXT_ON_DARK: "E9F4E5", SUBTITLE: "C3DFBE",
      DECOR_1: "3D702E", DECOR_2: "8D6D1C",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Ocean Calm
    {
      PRIMARY: "1F70AF", SECONDARY: "22705D", ACCENT: "8D6D1C", ALERT: "A8453B",
      SUCCESS: "22724D", ASSESS: "6952BD",
      BG_DARK: "154C76", BG_LIGHT: "F5F8FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "677D8C", TEXT_ON_DARK: "D7E5EF", SUBTITLE: "B1D1DF",
      DECOR_1: "22705D", DECOR_2: "8D6D1C",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Sunset Garden
    {
      PRIMARY: "C1385D", SECONDARY: "4A6F25", ACCENT: "A2621D", ALERT: "895B31",
      SUCCESS: "2F732F", ASSESS: "9C37B4",
      BG_DARK: "84263F", BG_LIGHT: "FBF6F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8A7A78", TEXT_ON_DARK: "F1E0E4", SUBTITLE: "E3C6CF",
      DECOR_1: "4A6F25", DECOR_2: "A2621D",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Mountain Air
    {
      PRIMARY: "1F7878", SECONDARY: "5D58BB", ACCENT: "A0631C", ALERT: "A34A31",
      SUCCESS: "207257", ASSESS: "794CBB",
      BG_DARK: "185E5E", BG_LIGHT: "F5F8F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7D7D", TEXT_ON_DARK: "F4F9F7", SUBTITLE: "CEE5E2",
      DECOR_1: "5D58BB", DECOR_2: "A0631C",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Harvest Gold
    {
      PRIMARY: "91613B", SECONDARY: "2F7238", ACCENT: "916A1A", ALERT: "AA4338",
      SUCCESS: "2E713C", ASSESS: "9742A2",
      BG_DARK: "624128", BG_LIGHT: "FAF6EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "8A8070", TEXT_ON_DARK: "EDE2DA", SUBTITLE: "D5CCB9",
      DECOR_1: "2F7238", DECOR_2: "916A1A",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Meadow Mist
    {
      PRIMARY: "317948", SECONDARY: "9A3F9A", ACCENT: "827122", ALERT: "9B4F3D",
      SUCCESS: "2E7145", ASSESS: "834AB4",
      BG_DARK: "305C3E", BG_LIGHT: "F6F8F2", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7D6E", TEXT_ON_DARK: "F1F6F1", SUBTITLE: "CFE4CF",
      DECOR_1: "9A3F9A", DECOR_2: "827122",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ─────────────────────────────────────────────────────────────────────────
  // GRADE 2 — Transitional, noticeable reduction. Trebuchet MS / Calibri.
  // ─────────────────────────────────────────────────────────────────────────
  grade2: [
    // [0] Forest Haven
    {
      PRIMARY: "29754C", SECONDARY: "3D6B30", ACCENT: "886B24", ALERT: "A34031",
      SUCCESS: "286C3C", ASSESS: "684DBB",
      BG_DARK: "226240", BG_LIGHT: "F8F5EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6C8070", TEXT_ON_DARK: "FAFBF9", SUBTITLE: "D3E5D3",
      DECOR_1: "3D6B30", DECOR_2: "886B24",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Ocean Calm
    {
      PRIMARY: "266CA1", SECONDARY: "276C5A", ACCENT: "886B24", ALERT: "964940",
      SUCCESS: "276C4D", ASSESS: "6450B3",
      BG_DARK: "1C5078", BG_LIGHT: "F4F7F9", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7D8C", TEXT_ON_DARK: "E2ECF3", SUBTITLE: "B8D5E0",
      DECOR_1: "276C5A", DECOR_2: "886B24",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Sunset Garden
    {
      PRIMARY: "B03F65", SECONDARY: "4C692A", ACCENT: "9B6224", ALERT: "7E5936",
      SUCCESS: "306C30", ASSESS: "8E3DA7",
      BG_DARK: "762A43", BG_LIGHT: "FAF5EF", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "887C78", TEXT_ON_DARK: "EDDADE", SUBTITLE: "DEC1CA",
      DECOR_1: "4C692A", DECOR_2: "9B6224",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Mountain Air
    {
      PRIMARY: "247272", SECONDARY: "5953B2", ACCENT: "996324", ALERT: "954C36",
      SUCCESS: "256C55", ASSESS: "744BAD",
      BG_DARK: "1F6262", BG_LIGHT: "F4F7F7", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6E7D7D", TEXT_ON_DARK: "FBFBFB", SUBTITLE: "D7EAE7",
      DECOR_1: "5953B2", DECOR_2: "996324",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Harvest Gold
    {
      PRIMARY: "8A5D3E", SECONDARY: "316C33", ACCENT: "8A6A20", ALERT: "99473D",
      SUCCESS: "316D3B", ASSESS: "8A4396",
      BG_DARK: "5B3F2B", BG_LIGHT: "F9F5EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7D7667", TEXT_ON_DARK: "E8DCD2", SUBTITLE: "CFC8B9",
      DECOR_1: "316C33", DECOR_2: "8A6A20",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Meadow Mist
    {
      PRIMARY: "347347", SECONDARY: "904190", ACCENT: "7E6F29", ALERT: "8F4F40",
      SUCCESS: "316C45", ASSESS: "7C49A4",
      BG_DARK: "376044", BG_LIGHT: "F5F7F0", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6E7D6E", TEXT_ON_DARK: "FBFBFB", SUBTITLE: "D8E8D8",
      DECOR_1: "904190", DECOR_2: "7E6F29",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ─────────────────────────────────────────────────────────────────────────
  // GRADE 3/4 — Balanced, some sophistication. Trebuchet MS / Calibri.
  // ─────────────────────────────────────────────────────────────────────────
  grade34: [
    // [0] Forest Haven
    {
      PRIMARY: "2D6B52", SECONDARY: "366439", ACCENT: "80682B", ALERT: "904439",
      SUCCESS: "2D6540", ASSESS: "624AA9",
      BG_DARK: "28604A", BG_LIGHT: "F7F4EA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6E8072", TEXT_ON_DARK: "F6FBF5", SUBTITLE: "D6E7D6",
      DECOR_1: "366439", DECOR_2: "80682B",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Ocean Calm
    {
      PRIMARY: "2A668C", SECONDARY: "2B6455", ACCENT: "80682B", ALERT: "864A45",
      SUCCESS: "2B6548", ASSESS: "614D9B",
      BG_DARK: "245878", BG_LIGHT: "F3F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6E7D8C", TEXT_ON_DARK: "F3F6F8", SUBTITLE: "CFE0E7",
      DECOR_1: "2B6455", DECOR_2: "80682B",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Sunset Garden
    {
      PRIMARY: "984465", SECONDARY: "47612D", ACCENT: "8E612B", ALERT: "6E563A",
      SUCCESS: "366438", ASSESS: "7D428F",
      BG_DARK: "6C3048", BG_LIGHT: "F9F4EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7D766E", TEXT_ON_DARK: "EADADC", SUBTITLE: "DAC1C8",
      DECOR_1: "47612D", DECOR_2: "8E612B",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Mountain Air
    {
      PRIMARY: "296A6A", SECONDARY: "575298", ACCENT: "8C612A", ALERT: "834C3B",
      SUCCESS: "2A6551", ASSESS: "6F4A94",
      BG_DARK: "256161", BG_LIGHT: "F3F6F6", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "727D7D", TEXT_ON_DARK: "FBFBFB", SUBTITLE: "D5E8E5",
      DECOR_1: "575298", DECOR_2: "8C612A",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Harvest Gold
    {
      PRIMARY: "7B5942", SECONDARY: "36643A", ACCENT: "7E6726", ALERT: "874841",
      SUCCESS: "35643E", ASSESS: "804580",
      BG_DARK: "584030", BG_LIGHT: "F8F4EA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7D7667", TEXT_ON_DARK: "E8DED5", SUBTITLE: "CFCABC",
      DECOR_1: "36643A", DECOR_2: "7E6726",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Meadow Mist
    {
      PRIMARY: "396B4B", SECONDARY: "804580", ACCENT: "746C27", ALERT: "7D4F43",
      SUCCESS: "356346", ASSESS: "754887",
      BG_DARK: "3B6148", BG_LIGHT: "F4F6EF", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "727D6E", TEXT_ON_DARK: "FBFBFB", SUBTITLE: "DBE8DB",
      DECOR_1: "804580", DECOR_2: "746C27",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ─────────────────────────────────────────────────────────────────────────
  // GRADE 5/6 — Most muted and sophisticated. Georgia / Calibri.
  // ─────────────────────────────────────────────────────────────────────────
  grade56: [
    // [0] Forest Haven
    {
      PRIMARY: "336356", SECONDARY: "395D40", ACCENT: "746732", ALERT: "784745",
      SUCCESS: "305D41", ASSESS: "5B4B8C",
      BG_DARK: "305C50", BG_LIGHT: "F6F3E8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "728076", TEXT_ON_DARK: "F4F8F3", SUBTITLE: "CFE2CF",
      DECOR_1: "395D40", DECOR_2: "746732",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [1] Ocean Calm
    {
      PRIMARY: "2D5F7A", SECONDARY: "2E5D55", ACCENT: "746732", ALERT: "724A46",
      SUCCESS: "2E5D48", ASSESS: "5A4D7E",
      BG_DARK: "2C5E78", BG_LIGHT: "F2F5F7", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "727D8C", TEXT_ON_DARK: "FBFBFB", SUBTITLE: "DDE8ED",
      DECOR_1: "2E5D55", DECOR_2: "746732",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [2] Sunset Garden
    {
      PRIMARY: "814868", SECONDARY: "445C37", ACCENT: "7F6232", ALERT: "60533B",
      SUCCESS: "395D3C", ASSESS: "6A4778",
      BG_DARK: "643850", BG_LIGHT: "F8F3EC", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7D766E", TEXT_ON_DARK: "EBDCDE", SUBTITLE: "D7C1C6",
      DECOR_1: "445C37", DECOR_2: "7F6232",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [3] Mountain Air
    {
      PRIMARY: "2C6363", SECONDARY: "505082", ACCENT: "7E6331", ALERT: "724B46",
      SUCCESS: "2B5D50", ASSESS: "674A7A",
      BG_DARK: "2B6161", BG_LIGHT: "F2F5F5", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "767D7D", TEXT_ON_DARK: "FBFBFB", SUBTITLE: "DAEAE7",
      DECOR_1: "505082", DECOR_2: "7E6331",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [4] Harvest Gold
    {
      PRIMARY: "6D5643", SECONDARY: "385C3D", ACCENT: "73682C", ALERT: "744A46",
      SUCCESS: "365C40", ASSESS: "724669",
      BG_DARK: "5C4838", BG_LIGHT: "F7F3E8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "7D7667", TEXT_ON_DARK: "EEE7E2", SUBTITLE: "D4CFC4",
      DECOR_1: "385C3D", DECOR_2: "73682C",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // [5] Meadow Mist
    {
      PRIMARY: "3C624C", SECONDARY: "714671", ACCENT: "676B2C", ALERT: "6E4E44",
      SUCCESS: "375C47", ASSESS: "724672",
      BG_DARK: "3D5F4B", BG_LIGHT: "F3F5EE", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "767D6E", TEXT_ON_DARK: "FBFBFB", SUBTITLE: "DAE7DA",
      DECOR_1: "714671", DECOR_2: "676B2C",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],
};

module.exports = { palettes };
