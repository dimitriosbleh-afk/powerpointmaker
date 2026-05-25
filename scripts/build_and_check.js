"use strict";

/**
 * Build gate script — runs a lesson build and enforces two hard checks:
 *   1. Zero diagnostics errors/warnings in build output
 *   2. markitdown parses the output PPTX without failure
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

/* ── Patterns ──────────────────────────────────────────────────────────────── */

// Matches diagnostics output from themes/core/diagnostics.js and layout.js.
// Layout helpers historically emit "[bounds]" warnings, so include those in
// the same hard gate instead of letting unsafe layouts pass.
const DIAG_LINE_RE = /^(ERROR|WARN) |\[bounds\]/;

/* ── Helpers ───────────────────────────────────────────────────────────────── */

function findPptx(dir) {
  if (!fs.existsSync(dir) || !fs.statSync(dir).isDirectory()) return null;
  const hit = fs.readdirSync(dir).find(f => f.endsWith(".pptx"));
  return hit ? path.join(dir, hit) : null;
}

function pluralise(n, word) {
  return n + " " + word + (n === 1 ? "" : "s");
}

/* ── Main ──────────────────────────────────────────────────────────────────── */

function main() {
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
  const diagLines = stderrLines.filter(l => DIAG_LINE_RE.test(l));
  const errors = diagLines.filter(l => l.startsWith("ERROR")).length;
  const warns = diagLines.filter(l => l.startsWith("WARN") || l.includes("[bounds]")).length;

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
    });

    if (md.status !== 0) {
      // Show the tail of the error for context
      const errTail = (md.stderr || "").split(/\r?\n/).filter(Boolean).slice(-6);
      errTail.forEach(l => console.error("  " + l));
      console.error("FAIL — markitdown exit code " + md.status);
      gatesFailed++;
    } else {
      console.log("PASS");
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

main();
