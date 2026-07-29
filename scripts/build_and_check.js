"use strict";

/**
 * Build gate script — runs a lesson build and enforces seven hard checks:
 *   0. The build itself completes
 *   1. Zero diagnostics errors/warnings in build output
 *   2. markitdown parses the PPTX, and no forbidden output markers survive
 *   3. Slide-face text hygiene (banned characters, layout-by-spaces)
 *   4. Teacher notes format (Glance Format budgets, reveal notes not duplicated)
 *   5. Hyperlink integrity (relative targets resolve)
 *   6. Lesson structure (resources placement, opening order, We Do vs You Do)
 *
 * Usage:
 *   node scripts/build_and_check.js builds/build_foo_lesson1.js
 *
 * Exit codes:
 *   0 = both gates passed
 *   1 = one or more gates failed (fix before visual QA)
 *   2 = bad usage or missing file
 */

const { spawnSync } = require("child_process");
const path = require("path");
const fs = require("fs");
const { scanTextForForbiddenOutput, validateNotesFormat, validateLessonStructure } = require("./qa_lib");

/* ── Patterns ──────────────────────────────────────────────────────────────── */

// Several helpers emit tagged warnings rather than "WARN ..." prefixes.
// Keep safety-related tags in the hard gate instead of letting unsafe decks pass.
const DIAG_PREFIX_RE = /^(ERROR|WARN)\b/;
const SAFETY_WARNING_TAGS = new Set([
  "addresourceslide",
  "bounds",
  "contrast",
  "exitticketslide",
  "lislide",
]);
const SAFETY_WARNING_TEXT_RE = /\bskipping element\b/i;
const PYTHON_ENV = { ...process.env, PYTHONUTF8: "1" };

/* ── Helpers ───────────────────────────────────────────────────────────────── */

function findPptx(dir) {
  if (!fs.existsSync(dir) || !fs.statSync(dir).isDirectory()) return null;
  const hit = fs.readdirSync(dir).find(f => f.endsWith(".pptx"));
  return hit ? path.join(dir, hit) : null;
}

function pluralise(n, word) {
  return n + " " + word + (n === 1 ? "" : "s");
}

function hasSafetyWarningTag(line) {
  return Array.from(line.matchAll(/\[([^\]]+)\]/g)).some((match) =>
    SAFETY_WARNING_TAGS.has(String(match[1]).toLowerCase())
  );
}

function isDiagnosticLine(line) {
  return DIAG_PREFIX_RE.test(line) || hasSafetyWarningTag(line) || SAFETY_WARNING_TEXT_RE.test(line);
}

/* ── Main ──────────────────────────────────────────────────────────────────── */

async function main() {
  const script = process.argv[2];

  if (!script) {
    console.error("Usage: node scripts/build_and_check.js <build-script>");
    console.error("  e.g. node scripts/build_and_check.js builds/build_special_numbers_lesson1.js");
    process.exit(2);
  }

  if (!fs.existsSync(script)) {
    console.error("Not found: " + script);
    process.exit(2);
  }

  let gatesFailed = 0;

  /* ── Gate 0: Build ─────────────────────────────────────────────────────── */

  console.log("\n── Build ─────────────────────────────────────────────");
  const build = spawnSync("node", [script], {
    encoding: "utf8",
    stdio: ["inherit", "pipe", "pipe"],
    env: PYTHON_ENV,
  });

  // Always show stdout (PPTX path, resource confirmations)
  if (build.stdout) process.stdout.write(build.stdout);

  if (build.status !== 0) {
    if (build.stderr) process.stderr.write(build.stderr);
    console.error("\nBUILD FAILED (exit code " + build.status + ")");
    process.exit(1);
  }

  /* ── Gate 1: Diagnostics ───────────────────────────────────────────────── */

  console.log("\n── Diagnostics ───────────────────────────────────────");
  const stderrLines = (build.stderr || "").split(/\r?\n/);
  const diagLines = stderrLines.filter(isDiagnosticLine);
  const errors = diagLines.filter(l => /^ERROR\b/.test(l)).length;
  const warns = diagLines.length - errors;

  if (diagLines.length === 0) {
    console.log("PASS — 0 errors, 0 warnings");
  } else {
    diagLines.forEach(l => console.error("  " + l));
    console.error("FAIL — " + pluralise(errors, "error") + ", " + pluralise(warns, "warning"));
    gatesFailed++;
  }

  /* ── Gate 2: markitdown ────────────────────────────────────────────────── */

  console.log("\n── Content QA (markitdown) ────────────────────────────");

  const pptxPathMatch = (build.stdout || "").match(/PPTX written to (.+)/);
  const pptxRaw = pptxPathMatch ? pptxPathMatch[1].trim() : null;
  let pptxFile = null;
  if (pptxRaw) {
    if (pptxRaw.endsWith(".pptx") && fs.existsSync(pptxRaw)) {
      pptxFile = pptxRaw;
    } else {
      pptxFile = findPptx(pptxRaw);
    }
  }

  if (!pptxFile) {
    console.error("FAIL — could not locate .pptx in build output (expected 'PPTX written to ...' in stdout)");
    gatesFailed++;
  } else {
    const md = spawnSync("python", ["-m", "markitdown", pptxFile], {
      encoding: "utf8",
      stdio: ["inherit", "pipe", "pipe"],
      env: PYTHON_ENV,
    });

    if (md.status !== 0) {
      // Show the tail of the error for context
      const errTail = (md.stderr || "").split(/\r?\n/).filter(Boolean).slice(-6);
      errTail.forEach(l => console.error("  " + l));
      console.error("FAIL — markitdown exit code " + md.status);
      gatesFailed++;
    } else {
      // Forbidden output scan on the extracted text (TODO/TBD/FIXME markers,
      // legacy resource codes) — previously only enforced for merged units.
      const forbidden = scanTextForForbiddenOutput(md.stdout || "", path.basename(pptxFile));
      if (forbidden.length) {
        forbidden.forEach(l => console.error("  ERROR " + l));
        console.error("FAIL — " + forbidden.length + " forbidden output pattern(s)");
        gatesFailed++;
      } else {
        console.log("PASS");
      }
    }
  }

  /* ── Gate 3: slide-face text hygiene ───────────────────────────────────── */
  // Em/en dashes, smart quotes, ellipsis chars and double hyphens are banned
  // on slide faces (the theme's installSlideTextPatch auto-converts them, so
  // any hit here means text bypassed the theme layer). Runs of 3+ spaces are
  // layout-by-spaces hacks: they render as accidental-looking gaps, so they
  // WARN and should be fixed with real layout (chips, columns, breakLine).

  console.log("\n── Slide text hygiene ─────────────────────────────────");

  if (!pptxFile) {
    console.error("SKIP — no PPTX located");
  } else {
    const unzip = spawnSync("unzip", ["-p", pptxFile, "ppt/slides/slide*.xml"], {
      encoding: "utf8",
      maxBuffer: 64 * 1024 * 1024,
    });
    if (unzip.status !== 0 || !unzip.stdout) {
      console.error("FAIL — could not read slide XML (unzip exit " + unzip.status + ")");
      gatesFailed++;
    } else {
      const hygieneErrors = [];
      const hygieneWarns = [];
      const runRe = /<a:t>([^<]*)<\/a:t>/g;
      let m;
      while ((m = runRe.exec(unzip.stdout))) {
        const text = m[1];
        const snippet = JSON.stringify(text.length > 60 ? text.slice(0, 60) + "..." : text);
        if (/[‐-―…‘’“”]/.test(text)) {
          hygieneErrors.push("banned character (dash family, ellipsis or smart quote) in " + snippet);
        }
        if (/(^|[^-])--(?!-)/.test(text)) {
          hygieneErrors.push("double hyphen '--' in " + snippet);
        }
        if (/ {3,}/.test(text)) {
          hygieneWarns.push("3+ consecutive spaces (fix with layout, not spaces) in " + snippet);
        }
      }
      if (hygieneErrors.length === 0 && hygieneWarns.length === 0) {
        console.log("PASS");
      } else {
        hygieneErrors.forEach(l => console.error("  ERROR " + l));
        hygieneWarns.forEach(l => console.error("  WARN  " + l));
        if (hygieneErrors.length > 0) {
          console.error("FAIL — " + hygieneErrors.length + " banned-character issue(s)");
          gatesFailed++;
        } else {
          console.log("PASS with " + hygieneWarns.length + " spacing warning(s) — fix before shipping");
        }
      }
    }
  }

  /* ── Gate 4: teacher notes format ──────────────────────────────────────── */
  // Megaprompt v12.3 sections 45-47: reveal slides must carry their own
  // post-reveal notes (never a byte-copy of the base slide's), and Glance
  // Format live zones must respect the rendered budgets (120 words / 16 words
  // per line / 18 physical lines) so they stay glanceable on an iPad.

  console.log("\n── Teacher notes format ───────────────────────────────");

  if (!pptxFile) {
    console.error("SKIP — no PPTX located");
  } else {
    try {
      const noteIssues = await validateNotesFormat(pptxFile);
      if (noteIssues.length) {
        noteIssues.forEach(l => console.error("  ERROR " + l));
        console.error("FAIL — " + pluralise(noteIssues.length, "notes format issue"));
        gatesFailed++;
      } else {
        console.log("PASS");
      }
    } catch (err) {
      console.error("  ERROR " + err.message);
      console.error("FAIL — could not validate notes format");
      gatesFailed++;
    }
  }

  /* ── Gate 5: hyperlink integrity ───────────────────────────────────────── */
  // Every relative hyperlink in the deck (resource-slide links) must resolve
  // to a file that actually exists next to the PPTX. Catches renamed PDFs,
  // wrong subfolders, and typos that would ship as dead links.

  console.log("\n── Hyperlink integrity ────────────────────────────────");

  if (!pptxFile) {
    console.error("SKIP — no PPTX located");
  } else {
    const rels = spawnSync("bash", ["-c",
      `unzip -p ${JSON.stringify(pptxFile)} 'ppt/slides/_rels/*.rels' 2>/dev/null || true`],
      { encoding: "utf8", maxBuffer: 16 * 1024 * 1024 });
    const relXml = rels.stdout || "";
    const deckDir = path.dirname(pptxFile);
    const linkIssues = [];
    let externalLinks = 0;
    const targetRe = /Target="([^"]+)"[^>]*TargetMode="External"|TargetMode="External"[^>]*Target="([^"]+)"/g;
    let lm;
    while ((lm = targetRe.exec(relXml))) {
      const target = lm[1] || lm[2];
      if (!target) continue;
      externalLinks++;
      if (/^[a-z][a-z0-9+.-]*:/i.test(target)) continue; // http:, https:, mailto: etc.
      // The target lives in XML, so an apostrophe in a resource name arrives
      // as &apos;. Unescape before touching the filesystem, or a perfectly
      // good link reads as broken.
      const unescaped = target
        .replace(/&apos;/g, "'")
        .replace(/&quot;/g, '"')
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&amp;/g, "&");
      const decoded = decodeURIComponent(unescaped.split("#")[0]);
      if (!fs.existsSync(path.join(deckDir, decoded))) {
        linkIssues.push(`link target does not exist: "${decoded}" (relative to ${deckDir})`);
      }
    }
    if (linkIssues.length) {
      linkIssues.forEach(l => console.error("  ERROR " + l));
      console.error("FAIL — " + linkIssues.length + " broken link(s)");
      gatesFailed++;
    } else {
      console.log("PASS (" + externalLinks + " external link(s) checked)");
    }
  }

  /* ── Gate 6: lesson structure ──────────────────────────────────────────── */
  // Rules that were honour-system until now: Teacher Resources near the front
  // (section 0a item 19), the fixed opening order (section 0a item 23), maths
  // decks carrying Fluency (section 23), and You Do not reusing We Do content
  // (section 35). Skipped entirely for decks with no LI slide, so catalogues
  // and smoke tests never trip it.

  console.log("\n── Lesson structure ───────────────────────────────────");

  if (!pptxFile) {
    console.error("SKIP — no PPTX located");
  } else {
    try {
      const structureIssues = await validateLessonStructure(pptxFile);
      if (structureIssues.length) {
        structureIssues.forEach(l => console.error("  ERROR " + l));
        console.error("FAIL — " + pluralise(structureIssues.length, "structure issue"));
        gatesFailed++;
      } else {
        console.log("PASS");
      }
    } catch (err) {
      console.error("  ERROR " + err.message);
      console.error("FAIL — could not validate lesson structure");
      gatesFailed++;
    }
  }

  /* ── Result ────────────────────────────────────────────────────────────── */

  console.log("\n══════════════════════════════════════════════════════");
  if (gatesFailed === 0) {
    console.log("BUILD CHECK PASSED — proceed to visual QA");
    process.exit(0);
  } else {
    console.error("BUILD CHECK FAILED — " + pluralise(gatesFailed, "gate") + " failed. Fix before visual QA.");
    process.exit(1);
  }
}

main().catch((err) => {
  console.error("BUILD CHECK CRASHED — " + (err && err.stack ? err.stack : err));
  process.exit(1);
});
