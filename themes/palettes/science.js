"use strict";

// ============================================================================
// Science Theme Palettes
// Clean, modern, discovery-oriented. Labs, microscopes, periodic tables,
// ecosystems, space. Clean whites/greys with strong STEM colours.
//
// 6 variants per year level, 5 year levels = 30 palettes total.
// All foreground-on-background pairs meet WCAG AA contrast (>= 4.5:1).
// ============================================================================

//
// RETUNED (scripts/retune_palettes.js): every role colour is the brightest
// shade of its hue that still clears a per-band contrast target against
// white (Foundation ~4.9:1 for PRIMARY, rising to ~6.8:1 by Year 5/6;
// supporting roles sit ~0.6 deeper). BG_DARK is rich rather than near-black.
// Edit hues here freely, then re-run the script to re-establish the floors.
const palettes = {

  // ==========================================================================
  // FOUNDATION  (Arial Black / Calibri)
  // Bold, exciting, wonder-filled. Deepest saturation for engagement.
  // ==========================================================================
  foundation: [
    // 1. Lab Coat — deep navy, dark teal, burnt orange, crimson
    {
      PRIMARY: "3872B6", SECONDARY: "0D7373", ACCENT: "AE5D1C",
      ALERT: "CA2323", SUCCESS: "187836", ASSESS: "9139D5",
      BG_DARK: "2B4D7A", BG_LIGHT: "F4F6FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E8EEF4", SUBTITLE: "CDD5DE",
      DECOR_1: "0D7373", DECOR_2: "AE5D1C",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 2. Periodic — deep indigo, element green, dark gold, dark red
    {
      PRIMARY: "5168D1", SECONDARY: "1C7752", ACCENT: "8E6F18",
      ALERT: "C82828", SUCCESS: "1A7742", ASSESS: "8744D1",
      BG_DARK: "284B8B", BG_LIGHT: "F3F5FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E4EBF6", SUBTITLE: "CFD6E2",
      DECOR_1: "1C7752", DECOR_2: "8E6F18",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 3. Ecosystem — dark forest, earth brown, deep sky blue, rust
    {
      PRIMARY: "2F802F", SECONDARY: "935C2A", ACCENT: "2277BC",
      ALERT: "B14636", SUCCESS: "20764F", ASSESS: "943CCF",
      BG_DARK: "295629", BG_LIGHT: "F4F7F2", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7A68", TEXT_ON_DARK: "E8F0E8", SUBTITLE: "C6D5C6",
      DECOR_1: "935C2A", DECOR_2: "2277BC",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 4. Galaxy — deep purple, nebula blue, stellar gold, supernova red
    {
      PRIMARY: "7A5BCD", SECONDARY: "236BAE", ACCENT: "936C13",
      ALERT: "C8223D", SUCCESS: "1B755A", ASSESS: "8641D7",
      BG_DARK: "4D3B9E", BG_LIGHT: "F5F3FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "706880", TEXT_ON_DARK: "EDE7F7", SUBTITLE: "D5CDE2",
      DECOR_1: "236BAE", DECOR_2: "936C13",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 5. Voltage — dark grey-blue, electric teal, amber, crimson
    {
      PRIMARY: "4172AF", SECONDARY: "0B7474", ACCENT: "A16713",
      ALERT: "B63F3F", SUCCESS: "1F774C", ASSESS: "8744D1",
      BG_DARK: "314F6D", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E4EDF2", SUBTITLE: "C8D2DA",
      DECOR_1: "0B7474", DECOR_2: "A16713",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 6. Microscope — dark slate, cell green, warm orange, berry
    {
      PRIMARY: "4073AE", SECONDARY: "1C7752", ACCENT: "A06720",
      ALERT: "AA3F8F", SUCCESS: "22764C", ASSESS: "8248D0",
      BG_DARK: "334E69", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E4EBF2", SUBTITLE: "C8D0DA",
      DECOR_1: "1C7752", DECOR_2: "A06720",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ==========================================================================
  // GRADE 1  (Arial Black / Calibri)
  // Bold, curious. Slightly more refined than Foundation.
  // ==========================================================================
  grade1: [
    // 1. Lab Coat — navy, teal, safety orange, crimson
    {
      PRIMARY: "376DB6", SECONDARY: "0D7070", ACCENT: "A95D1D",
      ALERT: "C42323", SUCCESS: "18732F", ASSESS: "8D37D0",
      BG_DARK: "274979", BG_LIGHT: "F4F6FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E5ECF5", SUBTITLE: "C3CFDC",
      DECOR_1: "0D7070", DECOR_2: "A95D1D",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 2. Periodic — deep indigo, element green, gold, dark red
    {
      PRIMARY: "4A65CC", SECONDARY: "1B734F", ACCENT: "8B6D19",
      ALERT: "C12828", SUCCESS: "1A7246", ASSESS: "8341CC",
      BG_DARK: "2D448D", BG_LIGHT: "F3F5FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E2E8F5", SUBTITLE: "C9CFDE",
      DECOR_1: "1B734F", DECOR_2: "8B6D19",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 3. Ecosystem — forest, earth brown, sky blue, rust
    {
      PRIMARY: "357B32", SECONDARY: "8D5828", ACCENT: "2277B4",
      ALERT: "AE3F34", SUCCESS: "24714D", ASSESS: "9037C9",
      BG_DARK: "275227", BG_LIGHT: "F4F7F3", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7A6A", TEXT_ON_DARK: "E5F0E5", SUBTITLE: "BECFBE",
      DECOR_1: "8D5828", DECOR_2: "2277B4",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 4. Galaxy — deep purple, nebula blue, stellar gold, supernova red
    {
      PRIMARY: "7858C7", SECONDARY: "2467A7", ACCENT: "936A12",
      ALERT: "C1233A", SUCCESS: "1B7151", ASSESS: "893BCE",
      BG_DARK: "4D3993", BG_LIGHT: "F5F4FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "706880", TEXT_ON_DARK: "E5E0F5", SUBTITLE: "CFC8DC",
      DECOR_1: "2467A7", DECOR_2: "936A12",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 5. Voltage — dark grey-blue, electric teal, amber, crimson
    {
      PRIMARY: "436EA4", SECONDARY: "0C7070", ACCENT: "996611",
      ALERT: "AE3E3E", SUCCESS: "1A724C", ASSESS: "863FC9",
      BG_DARK: "2F4B68", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E2EAF2", SUBTITLE: "C2CCD5",
      DECOR_1: "0C7070", DECOR_2: "996611",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 6. Microscope — dark slate, cell green, warm orange, berry
    {
      PRIMARY: "426DA2", SECONDARY: "1B734F", ACCENT: "A1631F",
      ALERT: "A2427D", SUCCESS: "1D7245", ASSESS: "893BCE",
      BG_DARK: "314A65", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E2E8F0", SUBTITLE: "C2CAD5",
      DECOR_1: "1B734F", DECOR_2: "A1631F",
      FONT_H: "Arial Black", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ==========================================================================
  // GRADE 2  (Trebuchet MS / Calibri)
  // Transitional. Slightly more mature palette, still inviting.
  // ==========================================================================
  grade2: [
    // 1. Lab Coat — navy, teal, safety orange, crimson
    {
      PRIMARY: "3768A7", SECONDARY: "0C6A6A", ACCENT: "A65C21",
      ALERT: "BC2121", SUCCESS: "186E30", ASSESS: "882ECE",
      BG_DARK: "284576", BG_LIGHT: "F5F7FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E8EEF5", SUBTITLE: "BECADA",
      DECOR_1: "0C6A6A", DECOR_2: "A65C21",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 2. Periodic — deep indigo, element green, gold, dark red
    {
      PRIMARY: "4160C7", SECONDARY: "1C6E4C", ACCENT: "846C14",
      ALERT: "BA2525", SUCCESS: "1A6E3F", ASSESS: "8039CC",
      BG_DARK: "2D4185", BG_LIGHT: "F3F5FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E2E8F5", SUBTITLE: "C3CADC",
      DECOR_1: "1C6E4C", DECOR_2: "846C14",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 3. Ecosystem — forest, earth brown, sky blue, rust
    {
      PRIMARY: "347436", SECONDARY: "855527", ACCENT: "2272B1",
      ALERT: "A93C32", SUCCESS: "236C4B", ASSESS: "8B33C6",
      BG_DARK: "264E26", BG_LIGHT: "F4F7F3", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7A6A", TEXT_ON_DARK: "E8F0E8", SUBTITLE: "C0CFC0",
      DECOR_1: "855527", DECOR_2: "2272B1",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 4. Galaxy — deep purple, nebula blue, stellar gold, supernova red
    {
      PRIMARY: "7252C3", SECONDARY: "2462A4", ACCENT: "8D6811",
      ALERT: "B92336", SUCCESS: "1B6C4D", ASSESS: "8437CA",
      BG_DARK: "493886", BG_LIGHT: "F5F4FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "706880", TEXT_ON_DARK: "E8E2F5", SUBTITLE: "CAC1DA",
      DECOR_1: "2462A4", DECOR_2: "8D6811",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 5. Voltage — dark grey-blue, electric teal, amber, crimson
    {
      PRIMARY: "446998", SECONDARY: "0A6A6A", ACCENT: "976312",
      ALERT: "AA3A3A", SUCCESS: "1B6E4B", ASSESS: "8638C5",
      BG_DARK: "2F4762", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E2ECF2", SUBTITLE: "BBC7D5",
      DECOR_1: "0A6A6A", DECOR_2: "976312",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 6. Microscope — dark slate, cell green, warm orange, berry
    {
      PRIMARY: "426A92", SECONDARY: "1C6E4C", ACCENT: "99631F",
      ALERT: "954372", SUCCESS: "1D6E46", ASSESS: "8437CA",
      BG_DARK: "2D485E", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E2EAF0", SUBTITLE: "C3CCD5",
      DECOR_1: "1C6E4C", DECOR_2: "99631F",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ==========================================================================
  // GRADE 3/4  (Trebuchet MS / Calibri)
  // Systematic, clear. Balanced saturation — serious but not austere.
  // ==========================================================================
  grade34: [
    // 1. Lab Coat — dark blue, teal, safety orange, crimson
    {
      PRIMARY: "326299", SECONDARY: "0C6464", ACCENT: "A2581D",
      ALERT: "B02121", SUCCESS: "166731", ASSESS: "7838B8",
      BG_DARK: "25436B", BG_LIGHT: "F5F7FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E8EEF4", SUBTITLE: "B7C6D5",
      DECOR_1: "0C6464", DECOR_2: "A2581D",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 2. Periodic — deep indigo, element green, gold, dark red
    {
      PRIMARY: "3959C1", SECONDARY: "1A6747", ACCENT: "806717",
      ALERT: "AB2424", SUCCESS: "19683B", ASSESS: "6F3FB8",
      BG_DARK: "234077", BG_LIGHT: "F3F5FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E0E8F5", SUBTITLE: "BDC5D7",
      DECOR_1: "1A6747", DECOR_2: "806717",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 3. Ecosystem — forest, earth brown, sky blue, rust
    {
      PRIMARY: "326D34", SECONDARY: "7F4F23", ACCENT: "216FA9",
      ALERT: "973F2E", SUCCESS: "2B654B", ASSESS: "704897",
      BG_DARK: "274927", BG_LIGHT: "F4F7F2", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7A68", TEXT_ON_DARK: "E8F0E8", SUBTITLE: "B9CBB9",
      DECOR_1: "7F4F23", DECOR_2: "216FA9",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 4. Galaxy — deep purple, nebula blue, stellar gold, supernova red
    {
      PRIMARY: "6C49C0", SECONDARY: "215D98", ACCENT: "876613",
      ALERT: "AD2034", SUCCESS: "196657", ASSESS: "7C2BCB",
      BG_DARK: "43338A", BG_LIGHT: "F5F3FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "706880", TEXT_ON_DARK: "E8E0F5", SUBTITLE: "CAC1DA",
      DECOR_1: "215D98", DECOR_2: "876613",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 5. Voltage — dark grey-blue, electric teal, amber, crimson
    {
      PRIMARY: "426387", SECONDARY: "096464", ACCENT: "926114",
      ALERT: "9F3636", SUCCESS: "19654B", ASSESS: "7442A7",
      BG_DARK: "2A435D", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E0EAF0", SUBTITLE: "B2C1CF",
      DECOR_1: "096464", DECOR_2: "926114",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 6. Microscope — dark slate, cell green, warm orange, berry
    {
      PRIMARY: "456380", SECONDARY: "1A6747", ACCENT: "935F20",
      ALERT: "8A4070", SUCCESS: "196647", ASSESS: "6F3FB8",
      BG_DARK: "2F4458", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2C2C2C", MUTED: "6B7280", TEXT_ON_DARK: "E0E8F0", SUBTITLE: "BBC5CF",
      DECOR_1: "1A6747", DECOR_2: "935F20",
      FONT_H: "Trebuchet MS", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],

  // ==========================================================================
  // GRADE 5/6  (Georgia / Calibri)
  // Most sophisticated/muted. Professional, research-oriented.
  // ==========================================================================
  grade56: [
    // 1. Lab Coat — steel navy, deep teal, warm amber-orange, crimson
    {
      PRIMARY: "315C91", SECONDARY: "0D5E5E", ACCENT: "9D551C",
      ALERT: "A32222", SUCCESS: "176131", ASSESS: "7A29B8",
      BG_DARK: "253E64", BG_LIGHT: "F5F7FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2D3142", MUTED: "6B7280", TEXT_ON_DARK: "E8F0F8", SUBTITLE: "B0C1D6",
      DECOR_1: "0D5E5E", DECOR_2: "9D551C",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 2. Periodic — deep indigo, element green, dark gold, dark red
    {
      PRIMARY: "3656AC", SECONDARY: "1A6045", ACCENT: "7B6513",
      ALERT: "A12626", SUCCESS: "196139", ASSESS: "742CBA",
      BG_DARK: "283A7D", BG_LIGHT: "F3F5FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2D3142", MUTED: "6B7280", TEXT_ON_DARK: "E2EAF8", SUBTITLE: "B5C1D5",
      DECOR_1: "1A6045", DECOR_2: "7B6513",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 3. Ecosystem — forest, earth brown, sky blue, rust
    {
      PRIMARY: "306530", SECONDARY: "734D26", ACCENT: "226BA1",
      ALERT: "943631", SUCCESS: "225F42", ASSESS: "7A2AB2",
      BG_DARK: "294529", BG_LIGHT: "F4F7F3", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2D3142", MUTED: "6B7A6A", TEXT_ON_DARK: "EAF2EA", SUBTITLE: "B2C8B2",
      DECOR_1: "734D26", DECOR_2: "226BA1",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 4. Galaxy — deep purple, nebula blue, stellar gold, supernova red
    {
      PRIMARY: "6542BB", SECONDARY: "20568A", ACCENT: "7E6212",
      ALERT: "A12236", SUCCESS: "1A6045", ASSESS: "7232B0",
      BG_DARK: "442F83", BG_LIGHT: "F5F4FA", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2D3142", MUTED: "706880", TEXT_ON_DARK: "EAE2F8", SUBTITLE: "C6BBDB",
      DECOR_1: "20568A", DECOR_2: "7E6212",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 5. Voltage — dark grey-blue, electric teal, amber, crimson
    {
      PRIMARY: "3E5D7C", SECONDARY: "0C5E5E", ACCENT: "8C5D15",
      ALERT: "933535", SUCCESS: "1A603E", ASSESS: "7B2EAC",
      BG_DARK: "2E3F54", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2D3142", MUTED: "6B7280", TEXT_ON_DARK: "E5EDF5", SUBTITLE: "B5C3D2",
      DECOR_1: "0C5E5E", DECOR_2: "8C5D15",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
    // 6. Microscope — dark slate, cell green, warm orange, berry
    {
      PRIMARY: "425C74", SECONDARY: "1A6045", ACCENT: "8C5D17",
      ALERT: "7D4166", SUCCESS: "1C6040", ASSESS: "7433AF",
      BG_DARK: "303F54", BG_LIGHT: "F4F6F8", BG_CARD: "FFFFFF", WHITE: "FFFFFF",
      CHARCOAL: "2D3142", MUTED: "6B7280", TEXT_ON_DARK: "E5EDF2", SUBTITLE: "B5C1CF",
      DECOR_1: "1A6045", DECOR_2: "8C5D17",
      FONT_H: "Georgia", FONT_B: "Calibri",
      SHADOW_BLUR: 6, SHADOW_OFFSET: 2, SHADOW_OPACITY: 0.12,
      CARD_SHADOW_BLUR: 4, CARD_SHADOW_OFFSET: 1, CARD_SHADOW_OPACITY: 0.10,
    },
  ],
};

module.exports = { palettes };
