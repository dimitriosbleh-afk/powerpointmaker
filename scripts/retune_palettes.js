"use strict";

/**
 * Retune the 150 theme palettes so colour does its job.
 *
 * The original palettes were authored far darker than WCAG AA requires: most
 * PRIMARY colours sat at 12-18:1 against white when 4.5:1 is the floor. The
 * effect on screen was near-black badges, near-black title slides and
 * counters the colour of a suit, on every deck, including Foundation.
 *
 * This script lifts every role colour to the brightest shade of the SAME hue
 * that still clears a per-band contrast target against white, so the promise
 * "white text on PRIMARY / SECONDARY / ACCENT / ALERT / SUCCESS / ASSESS is
 * always readable" is unchanged while the colours become genuinely colourful.
 * Younger bands get brighter targets; Year 5/6 stays deeper.
 *
 * It also re-derives TEXT_ON_DARK and SUBTITLE so they still read on the
 * lifted BG_DARK, and darkens MUTED where it had drifted too light on the
 * cream background.
 *
 * Files are rewritten in place. Comments and structure are preserved; only
 * the hex values change.
 *
 * Usage:
 *   node scripts/retune_palettes.js            # rewrite themes/palettes/*.js
 *   node scripts/retune_palettes.js --dry-run  # report only
 */

const fs = require("fs");
const path = require("path");
const { contrastRatio } = require("../themes/core/contrast");
const { liftToContrast, mixHex, normaliseHex } = require("../themes/core/color");

const ROOT = path.resolve(__dirname, "..");
const SUBJECTS = ["literacy", "numeracy", "science", "inquiry", "wellbeing"];
const LEVELS = ["foundation", "grade1", "grade2", "grade34", "grade56"];

// Contrast targets against WHITE for the dominant colour, by band. Everything
// is comfortably above the 4.5 floor. Supporting roles sit a little deeper so
// the palette keeps a light/dark rhythm instead of five equal mid-tones.
const BAND = {
  foundation: { primary: 4.9, support: 5.5, accent: 4.7, dark: 8.5, minSat: 0.46 },
  grade1:     { primary: 5.2, support: 5.8, accent: 4.8, dark: 9.0, minSat: 0.42 },
  grade2:     { primary: 5.6, support: 6.2, accent: 5.0, dark: 9.5, minSat: 0.38 },
  grade34:    { primary: 6.2, support: 6.8, accent: 5.3, dark: 10.0, minSat: 0.30 },
  grade56:    { primary: 6.8, support: 7.4, accent: 5.6, dark: 10.5, minSat: 0.24 },
};

function retunePalette(p, level) {
  const t = BAND[level];
  const lift = (hex, target) => liftToContrast(hex, target, "FFFFFF", { minSaturation: t.minSat });
  const out = { ...p };

  out.PRIMARY   = lift(p.PRIMARY, t.primary);
  out.SECONDARY = lift(p.SECONDARY, t.support);
  out.ACCENT    = lift(p.ACCENT, t.accent);
  out.ALERT     = lift(p.ALERT, t.support);
  out.SUCCESS   = lift(p.SUCCESS, t.support);
  out.ASSESS    = lift(p.ASSESS, t.support);
  // Title / closing background: rich and deep, but no longer near-black.
  // One-way: only lift a BG_DARK that is deeper than the target. A palette
  // whose dark was already lighter keeps it (floor 7:1 so white text stays solid).
  out.BG_DARK   = contrastRatio(p.BG_DARK, "FFFFFF") > t.dark
    ? liftToContrast(p.BG_DARK, t.dark, "FFFFFF", { minSaturation: Math.min(t.minSat, 0.35) })
    : liftToContrast(p.BG_DARK, Math.max(7.0, Math.min(t.dark, contrastRatio(p.BG_DARK, "FFFFFF"))), "FFFFFF");

  // Decorative keys are kept for back-compat (title slides no longer draw
  // the blurry circles); point them at the lifted supporting colours.
  out.DECOR_1 = out.SECONDARY;
  out.DECOR_2 = out.ACCENT;

  // On-dark text must still read on the lifted BG_DARK. Lighten toward
  // white until comfortable.
  const ensureOnDark = (hex, minRatio) => {
    let c = normaliseHex(hex, "FFFFFF");
    let steps = 0;
    while (contrastRatio(c, out.BG_DARK) < minRatio && steps < 40) {
      c = mixHex(c, "FFFFFF", 0.12);
      steps += 1;
    }
    return c;
  };
  out.TEXT_ON_DARK = ensureOnDark(p.TEXT_ON_DARK, 7.0);
  out.SUBTITLE     = ensureOnDark(p.SUBTITLE, 5.5);

  // Footer / caption grey: chrome, but it should not vanish on cream.
  let muted = normaliseHex(p.MUTED, "6B7280");
  let guard = 0;
  while (contrastRatio(muted, p.BG_LIGHT) < 3.6 && guard < 40) {
    muted = mixHex(muted, "000000", 0.08);
    guard += 1;
  }
  out.MUTED = muted;

  return out;
}

function rewriteFile(subject, dryRun) {
  const file = path.join(ROOT, "themes", "palettes", `${subject}.js`);
  const src = fs.readFileSync(file, "utf8");
  // Fresh require so a second run sees the current file.
  delete require.cache[require.resolve(file)];
  const { palettes } = require(file);

  const lines = src.split("\n");
  let level = null;
  let variantIdx = -1;
  let inObject = false;
  const retuned = {};
  LEVELS.forEach((lvl) => { retuned[lvl] = palettes[lvl].map((p) => retunePalette(p, lvl)); });

  const out = lines.map((line) => {
    const levelMatch = line.match(/^\s*(foundation|grade1|grade2|grade34|grade56):\s*\[/);
    if (levelMatch) {
      level = levelMatch[1];
      variantIdx = -1;
      return line;
    }
    if (level && /^\s*\{\s*$/.test(line)) {
      variantIdx += 1;
      inObject = true;
      return line;
    }
    if (inObject && /^\s*\},?\s*$/.test(line)) {
      inObject = false;
      return line;
    }
    if (level && /^\s*\],?\s*$/.test(line) && !inObject) {
      level = null;
      return line;
    }
    if (!inObject || level == null || variantIdx < 0) return line;

    const target = retuned[level][variantIdx];
    if (!target) return line;
    return line.replace(/\b([A-Z_0-9]+):\s*"([0-9A-Fa-f]{6})"/g, (m, key, hex) => {
      if (target[key] == null || !/^[0-9A-Fa-f]{6}$/.test(String(target[key]))) return m;
      return `${key}: "${target[key]}"`;
    });
  });

  let text = out.join("\n");

  // Header notes that described the old (over-dark) validation are stale.
  text = text.replace(/^\/\/\s*-\s*PRIMARY\/SECONDARY\/ALERT luminance < 0\.18\s*$/m,
    "//   - PRIMARY/SECONDARY/ALERT lifted to the brightest passing shade (see below)");

  const RETUNE_NOTE = [
    "//",
    "// RETUNED (scripts/retune_palettes.js): every role colour is the brightest",
    "// shade of its hue that still clears a per-band contrast target against",
    "// white (Foundation ~4.9:1 for PRIMARY, rising to ~6.8:1 by Year 5/6;",
    "// supporting roles sit ~0.6 deeper). BG_DARK is rich rather than near-black.",
    "// Edit hues here freely, then re-run the script to re-establish the floors.",
    "",
  ].join("\n");
  if (!text.includes("RETUNED (scripts/retune_palettes.js)")) {
    text = text.replace(/\nconst palettes = \{/, `\n${RETUNE_NOTE}const palettes = {`);
  }

  // Sanity: the rewritten module must still parse and expose 30 palettes.
  const tmp = path.join(ROOT, "themes", "palettes", `.${subject}.retune.tmp.js`);
  fs.writeFileSync(tmp, text);
  delete require.cache[require.resolve(tmp)];
  const check = require(tmp).palettes;
  fs.unlinkSync(tmp);
  LEVELS.forEach((lvl) => {
    if (!Array.isArray(check[lvl]) || check[lvl].length !== palettes[lvl].length) {
      throw new Error(`[retune] ${subject}/${lvl}: variant count changed`);
    }
    check[lvl].forEach((p, i) => {
      ["PRIMARY", "SECONDARY", "ACCENT", "ALERT", "SUCCESS", "ASSESS"].forEach((k) => {
        if (contrastRatio(p[k], "FFFFFF") < 4.5) throw new Error(`[retune] ${subject}/${lvl}[${i}].${k} fails AA`);
      });
      if (contrastRatio(p.TEXT_ON_DARK, p.BG_DARK) < 4.5) throw new Error(`[retune] ${subject}/${lvl}[${i}] TEXT_ON_DARK fails`);
      if (contrastRatio(p.CHARCOAL, p.BG_LIGHT) < 4.5) throw new Error(`[retune] ${subject}/${lvl}[${i}] CHARCOAL fails`);
    });
  });

  if (!dryRun) fs.writeFileSync(file, text);
  return { palettes, retuned: check };
}

function report(subject, before, after) {
  const roles = ["PRIMARY", "SECONDARY", "ACCENT", "ALERT", "SUCCESS", "BG_DARK"];
  console.log(`\n${subject}`);
  LEVELS.forEach((lvl) => {
    const b = before[lvl][0];
    const a = after[lvl][0];
    const row = roles.map((k) =>
      `${k} ${contrastRatio(b[k], "FFFFFF").toFixed(1)}->${contrastRatio(a[k], "FFFFFF").toFixed(1)}`
    ).join("  ");
    console.log(`  ${lvl.padEnd(10)} v0: ${row}`);
  });
}

function main() {
  const dryRun = process.argv.includes("--dry-run");
  SUBJECTS.forEach((subject) => {
    const { palettes, retuned } = rewriteFile(subject, dryRun);
    report(subject, palettes, retuned);
  });
  console.log(dryRun ? "\nDry run - no files written." : "\nPalettes rewritten.");
}

main();
