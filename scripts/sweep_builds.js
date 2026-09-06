"use strict";

/**
 * Corpus-wide build sweep.
 *
 * Runs every build script in builds/ through the full build_and_check gate and
 * reports which fail and why. The per-lesson gate tells you about one deck;
 * this tells you about the whole library, which is what matters when a change
 * to the shared theme layer could affect every deck at once.
 *
 * Usage:
 *   node scripts/sweep_builds.js                 # every builds/build_*.js
 *   node scripts/sweep_builds.js --filter=poetry # only matching names
 *   node scripts/sweep_builds.js --jobs=6        # concurrency (default 4)
 *   node scripts/sweep_builds.js --quick         # build only, skip the gates
 *
 * Exit code 0 when every script passes, 1 otherwise.
 */

const { spawn } = require("child_process");
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");

function parseArgs(argv) {
  const opts = { filter: null, jobs: 4, quick: false };
  argv.forEach((arg) => {
    if (arg.startsWith("--filter=")) opts.filter = arg.slice(9);
    else if (arg.startsWith("--jobs=")) opts.jobs = Math.max(1, parseInt(arg.slice(7), 10) || 4);
    else if (arg === "--quick") opts.quick = true;
  });
  return opts;
}

function listBuildScripts(filter) {
  return fs.readdirSync(path.join(ROOT, "builds"))
    .filter((n) => /^build_.*\.js$/.test(n) || /^exemplar_.*\.json$/.test(n) || /^lesson_.*\.json$/.test(n))
    .filter((n) => !filter || n.includes(filter))
    .sort()
    .map((n) => path.join("builds", n));
}

function runOne(script, quick) {
  return new Promise((resolve) => {
    const args = quick ? [script] : [path.join("scripts", "build_and_check.js"), script];
    const child = spawn("node", args, {
      cwd: ROOT,
      env: { ...process.env, PYTHONUTF8: "1" },
    });

    let out = "";
    child.stdout.on("data", (d) => { out += d; });
    child.stderr.on("data", (d) => { out += d; });
    child.on("close", (code) => {
      const lines = out.split(/\r?\n/);
      // "FAIL - <gate>" is the authoritative signal for WHICH gate failed.
      // Leading WARN/ERROR lines are diagnostic detail and can belong to a
      // gate that still passed, so they must not drive the grouping.
      const gates = lines
        .filter((l) => /^FAIL\b|^BUILD FAILED/.test(l))
        .map((l) => l.trim());
      // Space-run warnings are a Gate 3 advisory that does not fail the build.
      // They are numerous, so leaving them in crowds the actual cause out of
      // the captured detail.
      // Some helpers emit tagged diagnostics ("[contrast] ...") rather than a
      // WARN prefix, and those still fail Gate 1, so they must be captured too.
      const detail = lines
        .filter((l) => /^\s*(ERROR|WARN)\b/.test(l) || /^\s*\[[a-zA-Z]+\]/.test(l) || /skipping element/i.test(l))
        .filter((l) => !/3\+ consecutive spaces/.test(l))
        .map((l) => l.trim())
        .slice(0, 8);
      resolve({ script, code, gates, detail, out });
    });
  });
}

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  const scripts = listBuildScripts(opts.filter);

  if (scripts.length === 0) {
    console.error("No build scripts matched.");
    process.exit(2);
  }

  console.log(`Sweeping ${scripts.length} build script(s) with ${opts.jobs} job(s)` +
    (opts.quick ? " [quick: build only]" : " [full gate]") + "\n");

  const results = [];
  let cursor = 0;
  let done = 0;

  async function worker() {
    while (cursor < scripts.length) {
      const script = scripts[cursor++];
      const result = await runOne(script, opts.quick);
      results.push(result);
      done += 1;
      const mark = result.code === 0 ? "ok  " : "FAIL";
      console.log(`[${String(done).padStart(3)}/${scripts.length}] ${mark} ${path.basename(script)}`);
    }
  }

  await Promise.all(Array.from({ length: opts.jobs }, worker));

  const failures = results.filter((r) => r.code !== 0)
    .sort((a, b) => a.script.localeCompare(b.script));

  console.log("\n" + "=".repeat(70));
  console.log(`${results.length} scripts | ${results.length - failures.length} passed | ${failures.length} failed`);
  console.log("=".repeat(70));

  if (failures.length) {
    // Group by which GATE failed, so a systemic problem reads as one problem.
    const buckets = new Map();
    failures.forEach((f) => {
      const key = f.gates.length
        ? f.gates.map((g) => g.replace(/\s+\d+\s+/, " N ").replace(/^FAIL\s*[-—]\s*/, "")).join(" + ")
        : "no gate line (crashed early)";
      if (!buckets.has(key)) buckets.set(key, []);
      buckets.get(key).push(path.basename(f.script));
    });
    console.log("\nFailures grouped by failing gate:\n");
    [...buckets.entries()]
      .sort((a, b) => b[1].length - a[1].length)
      .forEach(([reason, list]) => {
        console.log(`  [${list.length}] ${reason}`);
        list.slice(0, 6).forEach((s) => console.log(`        ${s}`));
        if (list.length > 6) console.log(`        ...and ${list.length - 6} more`);
        console.log("");
      });

    fs.mkdirSync(path.join(ROOT, "tmp"), { recursive: true });
  fs.writeFileSync(
      path.join(ROOT, "tmp", "sweep_failures.txt"),
      failures.map((f) => `### ${f.script}\nGATES: ${f.gates.join(" | ")}\n${f.detail.join("\n")}\n`).join("\n"),
      "utf8"
    );
    console.log("Full failure detail written to tmp/sweep_failures.txt");
  }

  process.exit(failures.length ? 1 : 0);
}

main().catch((err) => {
  console.error(err && err.stack ? err.stack : err);
  process.exit(1);
});
