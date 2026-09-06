"use strict";

// ============================================================================
// Inquiry Theme Palettes
// Exploratory, curious, adventurous. Earthy tones with discovery accents.
// 5 year levels x 6 variants = 30 palettes.
//
// Every colour has been validated:
//   - WHITE on PRIMARY/SECONDARY/ACCENT/ALERT: contrast >= 4.5:1
//   - CHARCOAL on BG_LIGHT: contrast >= 4.5:1
//   - TEXT_ON_DARK on BG_DARK: contrast >= 4.5:1
//   - PRIMARY/SECONDARY/ALERT lifted to the brightest passing shade (see below)
// ============================================================================

//
// RETUNED (scripts/retune_palettes.js): every role colour is the brightest
// shade of its hue that still clears a per-band contrast target against
// white (Foundation ~4.9:1 for PRIMARY, rising to ~6.8:1 by Year 5/6;
// supporting roles sit ~0.6 deeper). BG_DARK is rich rather than near-black.
// Edit hues here freely, then re-run the script to re-establish the floors.
const palettes = {

  // --------------------------------------------------------------------------
  // FOUNDATION  (Arial Black / Calibri — bold, exciting, high-saturation)
  // --------------------------------------------------------------------------
  foundation: [
    // 0 — Explorer: bold olive expedition
    {
      PRIMARY: "4C7D2C",  SECONDARY: "AF4900",  ACCENT: "926D0C",
      ALERT: "4F60BE",  SUCCESS: "1E773B",  ASSESS: "9043C5",
      BG_DARK: "295529",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F5F0E0",  SUBTITLE: "D9CFB0",
      DECOR_1: "AF4900",  DECOR_2: "926D0C",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 1 — Detective: dark teal-grey mystery
    {
      PRIMARY: "39759B",  SECONDARY: "C62A44",  ACCENT: "89711C",
      ALERT: "326BA7",  SUCCESS: "1E773B",  ASSESS: "9043C5",
      BG_DARK: "44448D",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "D5CFC3",
      DECOR_1: "C62A44",  DECOR_2: "89711C",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 2 — Archaeologist: chocolate & ancient gold
    {
      PRIMARY: "AF5641",  SECONDARY: "237823",  ACCENT: "9E660B",
      ALERT: "B73F27",  SUCCESS: "1E773B",  ASSESS: "9043C5",
      BG_DARK: "654530",  BG_LIGHT: "FFFAF5",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F5EDE0",  SUBTITLE: "DECFB5",
      DECOR_1: "237823",  DECOR_2: "9E660B",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 3 — Navigator: midnight navy charting
    {
      PRIMARY: "366DCB",  SECONDARY: "0D7373",  ACCENT: "946D0C",
      ALERT: "CE1313",  SUCCESS: "1E773B",  ASSESS: "9043C5",
      BG_DARK: "254992",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "C8D5E1",
      DECOR_1: "0D7373",  DECOR_2: "946D0C",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 4 — Naturalist: deep forest observation
    {
      PRIMARY: "2D7F2D",  SECONDARY: "4267B2",  ACCENT: "8D6F0C",
      ALERT: "C82447",  SUCCESS: "1E773B",  ASSESS: "9043C5",
      BG_DARK: "295529",  BG_LIGHT: "FFFAF5",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F5F0E0",  SUBTITLE: "C1D6B9",
      DECOR_1: "4267B2",  DECOR_2: "8D6F0C",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 5 — Inventor: dark slate creative
    {
      PRIMARY: "3F74A9",  SECONDARY: "A74F0E",  ACCENT: "0E816A",
      ALERT: "AA3C93",  SUCCESS: "1E773B",  ASSESS: "9043C5",
      BG_DARK: "334E6A",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "C8D5DF",
      DECOR_1: "A74F0E",  DECOR_2: "0E816A",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // --------------------------------------------------------------------------
  // GRADE 1  (Arial Black / Calibri — bold)
  // --------------------------------------------------------------------------
  grade1: [
    // 0 — Explorer: olive & terracotta expedition
    {
      PRIMARY: "4B7831",  SECONDARY: "A44B26",  ACCENT: "8F6D12",
      ALERT: "4062B1",  SUCCESS: "1F743A",  ASSESS: "8844B9",
      BG_DARK: "2C5127",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F5F0E0",  SUBTITLE: "D5CFB5",
      DECOR_1: "A44B26",  DECOR_2: "8F6D12",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 1 — Detective: teal-grey investigation
    {
      PRIMARY: "3D7195",  SECONDARY: "BA2F46",  ACCENT: "8B6D1E",
      ALERT: "32689E",  SUCCESS: "1F743A",  ASSESS: "8844B9",
      BG_DARK: "374771",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "D5CFC6",
      DECOR_1: "BA2F46",  DECOR_2: "8B6D1E",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 2 — Archaeologist: warm brown discovery
    {
      PRIMARY: "A15842",  SECONDARY: "2E722E",  ACCENT: "9F6512",
      ALERT: "AE422C",  SUCCESS: "1F743A",  ASSESS: "8844B9",
      BG_DARK: "5F432E",  BG_LIGHT: "FFFAF5",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F5EDE0",  SUBTITLE: "DBCFB5",
      DECOR_1: "2E722E",  DECOR_2: "9F6512",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 3 — Navigator: deep navy charting
    {
      PRIMARY: "3A6EAF",  SECONDARY: "186F6F",  ACCENT: "8F6B12",
      ALERT: "C32525",  SUCCESS: "1F743A",  ASSESS: "8844B9",
      BG_DARK: "2A4979",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "BBCFE1",
      DECOR_1: "186F6F",  DECOR_2: "8F6B12",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 4 — Naturalist: forest & slate observation
    {
      PRIMARY: "327B32",  SECONDARY: "3E6798",  ACCENT: "8A6E17",
      ALERT: "BD2A4A",  SUCCESS: "1F743A",  ASSESS: "8844B9",
      BG_DARK: "275127",  BG_LIGHT: "FFFAF5",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F5F0E0",  SUBTITLE: "B9CFB2",
      DECOR_1: "3E6798",  DECOR_2: "8A6E17",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 5 — Inventor: slate blue problem-solving
    {
      PRIMARY: "3E7199",  SECONDARY: "9D4F1E",  ACCENT: "1C7F75",
      ALERT: "9D408C",  SUCCESS: "1F743A",  ASSESS: "8844B9",
      BG_DARK: "314B65",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "C3CFDA",
      DECOR_1: "9D4F1E",  DECOR_2: "1C7F75",
      FONT_H:    "Arial Black",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // --------------------------------------------------------------------------
  // GRADE 2  (Trebuchet MS / Calibri — transitional)
  // --------------------------------------------------------------------------
  grade2: [
    // 0 — Explorer: earthy olive expedition
    {
      PRIMARY: "4E7032",  SECONDARY: "9A491B",  ACCENT: "8B6911",
      ALERT: "3D5EA5",  SUCCESS: "206D3C",  ASSESS: "8044B0",
      BG_DARK: "384B24",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F5EDE0",  SUBTITLE: "CFC8B0",
      DECOR_1: "9A491B",  DECOR_2: "8B6911",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 1 — Detective: charcoal-teal investigation
    {
      PRIMARY: "3B6E83",  SECONDARY: "B03141",  ACCENT: "886A1E",
      ALERT: "3D5F9F",  SUCCESS: "206D3C",  ASSESS: "8044B0",
      BG_DARK: "314666",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "CDC8C1",
      DECOR_1: "B03141",  DECOR_2: "886A1E",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 2 — Archaeologist: deep brown & sage
    {
      PRIMARY: "995445",  SECONDARY: "316C31",  ACCENT: "996111",
      ALERT: "A24232",  SUCCESS: "206D3C",  ASSESS: "8044B0",
      BG_DARK: "544229",  BG_LIGHT: "FFFAF5",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F5EDE0",  SUBTITLE: "D6C6AB",
      DECOR_1: "316C31",  DECOR_2: "996111",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 3 — Navigator: navy & teal charting
    {
      PRIMARY: "3B67A9",  SECONDARY: "196B6B",  ACCENT: "8D6811",
      ALERT: "B92424",  SUCCESS: "206D3C",  ASSESS: "8044B0",
      BG_DARK: "25456F",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "BBC8DB",
      DECOR_1: "196B6B",  DECOR_2: "8D6811",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 4 — Naturalist: forest & slate blue
    {
      PRIMARY: "347434",  SECONDARY: "3E638A",  ACCENT: "866B19",
      ALERT: "B22C49",  SUCCESS: "206D3C",  ASSESS: "8044B0",
      BG_DARK: "254E25",  BG_LIGHT: "FFFAF5",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F5F0E0",  SUBTITLE: "B9CFB2",
      DECOR_1: "3E638A",  DECOR_2: "866B19",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 5 — Inventor: slate & burnt orange
    {
      PRIMARY: "3F6B8D",  SECONDARY: "934E1B",  ACCENT: "1D7C6E",
      ALERT: "934281",  SUCCESS: "206D3C",  ASSESS: "8044B0",
      BG_DARK: "2E4760",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "BBC8D5",
      DECOR_1: "934E1B",  DECOR_2: "1D7C6E",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // --------------------------------------------------------------------------
  // GRADE 3/4  (Trebuchet MS / Calibri — curious, structured)
  // --------------------------------------------------------------------------
  grade34: [
    // 0 — Explorer: deep khaki/olive expedition
    {
      PRIMARY: "506838",  SECONDARY: "8F4714",  ACCENT: "876613",
      ALERT: "3B599B",  SUCCESS: "2B654B",  ASSESS: "7044A3",
      BG_DARK: "2E4927",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBD8",  SUBTITLE: "C8C1AB",
      DECOR_1: "8F4714",  DECOR_2: "876613",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 1 — Detective: charcoal & crimson investigation
    {
      PRIMARY: "42647B",  SECONDARY: "A52D3F",  ACCENT: "836816",
      ALERT: "355C89",  SUCCESS: "2B654B",  ASSESS: "7044A3",
      BG_DARK: "32435D",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "C8C7BE",
      DECOR_1: "A52D3F",  DECOR_2: "836816",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 2 — Archaeologist: deep brown & amber
    {
      PRIMARY: "885349",  SECONDARY: "306630",  ACCENT: "926114",
      ALERT: "973F2E",  SUCCESS: "2B654B",  ASSESS: "7044A3",
      BG_DARK: "533D2C",  BG_LIGHT: "FFFAF5",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F5EDE0",  SUBTITLE: "D0C0A0",
      DECOR_1: "306630",  DECOR_2: "926114",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 3 — Navigator: deep navy & teal charting
    {
      PRIMARY: "3A609E",  SECONDARY: "196564",  ACCENT: "876613",
      ALERT: "AE2222",  SUCCESS: "2B654B",  ASSESS: "7044A3",
      BG_DARK: "29426E",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "BBC8DB",
      DECOR_1: "196564",  DECOR_2: "876613",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 4 — Naturalist: forest & slate observation
    {
      PRIMARY: "326D34",  SECONDARY: "3E5D7E",  ACCENT: "806717",
      ALERT: "A52D41",  SUCCESS: "2B654B",  ASSESS: "7044A3",
      BG_DARK: "274827",  BG_LIGHT: "FFFAF5",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBD8",  SUBTITLE: "A8C8A0",
      DECOR_1: "3E5D7E",  DECOR_2: "806717",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 5 — Inventor: dark blue-grey creative
    {
      PRIMARY: "466382",  SECONDARY: "894A1A",  ACCENT: "1D7769",
      ALERT: "884276",  SUCCESS: "2B654B",  ASSESS: "7044A3",
      BG_DARK: "2E445B",  BG_LIGHT: "FFF8F0",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0EBE0",  SUBTITLE: "BBC8D5",
      DECOR_1: "894A1A",  DECOR_2: "1D7769",
      FONT_H:    "Trebuchet MS",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // --------------------------------------------------------------------------
  // GRADE 5/6  (Georgia / Calibri — scholarly, research-oriented, most muted)
  // --------------------------------------------------------------------------
  grade56: [
    // 0 — Explorer: muted olive & warm terracotta
    {
      PRIMARY: "4D603B",  SECONDARY: "824516",  ACCENT: "806316",
      ALERT: "34548F",  SUCCESS: "265F42",  ASSESS: "614696",
      BG_DARK: "324328",  BG_LIGHT: "F7F4EE",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "E8E2D5",  SUBTITLE: "C8C1B5",
      DECOR_1: "824516",  DECOR_2: "806316",
      FONT_H:    "Georgia",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 1 — Detective: sophisticated charcoal & dusty crimson
    {
      PRIMARY: "435E6D",  SECONDARY: "992E3E",  ACCENT: "78661F",
      ALERT: "385682",  SUCCESS: "265F42",  ASSESS: "614696",
      BG_DARK: "30414E",  BG_LIGHT: "F7F4EE",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "E8E2D5",  SUBTITLE: "C3C3BE",
      DECOR_1: "992E3E",  DECOR_2: "78661F",
      FONT_H:    "Georgia",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 2 — Archaeologist: rich brown & scholarly sage
    {
      PRIMARY: "7A503D",  SECONDARY: "2F602F",  ACCENT: "8C5D17",
      ALERT: "903A2A",  SUCCESS: "265F42",  ASSESS: "614696",
      BG_DARK: "4B3C28",  BG_LIGHT: "FFFAF5",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "F0E8D8",  SUBTITLE: "CFC1AB",
      DECOR_1: "2F602F",  DECOR_2: "8C5D17",
      FONT_H:    "Georgia",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 3 — Navigator: deep navy & academic teal
    {
      PRIMARY: "385A96",  SECONDARY: "195E5E",  ACCENT: "816317",
      ALERT: "A42121",  SUCCESS: "265F42",  ASSESS: "614696",
      BG_DARK: "253E64",  BG_LIGHT: "F7F4EE",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "E8E2D5",  SUBTITLE: "B5C1D5",
      DECOR_1: "195E5E",  DECOR_2: "816317",
      FONT_H:    "Georgia",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 4 — Naturalist: refined forest & scholarly slate
    {
      PRIMARY: "376537",  SECONDARY: "3C576F",  ACCENT: "7B6516",
      ALERT: "9A2C42",  SUCCESS: "265F42",  ASSESS: "614696",
      BG_DARK: "2D4529",  BG_LIGHT: "FFFAF5",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "E8E5D8",  SUBTITLE: "ABC8AB",
      DECOR_1: "3C576F",  DECOR_2: "7B6516",
      FONT_H:    "Georgia",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 5 — Inventor: refined blue-grey & copper
    {
      PRIMARY: "425C74",  SECONDARY: "7D471B",  ACCENT: "1F7463",
      ALERT: "7E406B",  SUCCESS: "265F42",  ASSESS: "614696",
      BG_DARK: "2E3F54",  BG_LIGHT: "F7F4EE",  BG_CARD: "FFFFFF",
      WHITE: "FFFFFF",  CHARCOAL: "2C2C2C",  MUTED: "717F7D",
      TEXT_ON_DARK: "E8E2D5",  SUBTITLE: "B5C1CF",
      DECOR_1: "7D471B",  DECOR_2: "1F7463",
      FONT_H:    "Georgia",  FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

};

module.exports = { palettes };
