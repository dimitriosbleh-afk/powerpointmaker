"use strict";

const { spawnSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const PY_BUILD_UNIT = path.join(ROOT, "scripts", "build_unit.py");
const MANIFESTS_DIR = path.join(ROOT, "builds", "manifests");

function listManifests() {
  if (!fs.existsSync(MANIFESTS_DIR)) return [];
  return fs.readdirSync(MANIFESTS_DIR)
    .filter((name) => name.endsWith(".json"))
    .map((name) => path.join("builds", "manifests", name));
}

function resolveManifest(arg) {
  if (!arg) return null;
  const direct = path.resolve(ROOT, arg);
  if (fs.existsSync(direct)) return direct;
  const shorthand = path.join(MANIFESTS_DIR, `${arg}.json`);
  return fs.existsSync(shorthand) ? shorthand : null;
}

function usage() {
  console.error("Usage: node scripts/build_unit.js <manifest-or-name> [--skip-build]");
  console.error("  e.g. node scripts/build_unit.js builds/manifests/holes.json");
  console.error("  e.g. node scripts/build_unit.js holes --skip-build");
  const manifests = listManifests();
  if (manifests.length) {
    console.error("");
    console.error("Available manifests:");
    manifests.forEach((manifest) => console.error(`  - ${manifest}`));
  }
}

function main() {
  const args = process.argv.slice(2);
  const manifestArg = args.find((arg) => !arg.startsWith("-"));
  const manifest = resolveManifest(manifestArg);

  if (!manifest) {
    usage();
    process.exit(2);
  }

  const pyArgs = [PY_BUILD_UNIT, path.relative(ROOT, manifest)];
  if (args.includes("--skip-build")) pyArgs.push("--skip-build");

  const result = spawnSync("python", pyArgs, {
    cwd: ROOT,
    stdio: "inherit",
  });

  if (result.error) throw result.error;
  process.exit(result.status == null ? 1 : result.status);
}

main();
