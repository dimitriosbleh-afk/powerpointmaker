"use strict";

const fs = require("fs");
const path = require("path");
const {
  ROOT,
  runCommand,
  extractText,
  scanTextForForbiddenOutput,
  validateNotesXml,
  listPdfFiles,
} = require("./qa_lib");

function usage() {
  console.error("Usage: node scripts/qa_unit.js <manifest-or-name> [--skip-build] [--skip-merge]");
  console.error("  e.g. node scripts/qa_unit.js builds/manifests/holes.json");
  console.error("  e.g. node scripts/qa_unit.js holes --skip-build");
}

function resolveManifest(arg) {
  if (!arg) return null;
  const direct = path.resolve(ROOT, arg);
  if (fs.existsSync(direct)) return direct;
  const shorthand = path.join(ROOT, "builds", "manifests", `${arg}.json`);
  return fs.existsSync(shorthand) ? shorthand : null;
}

function loadManifest(filePath) {
  const manifest = JSON.parse(fs.readFileSync(filePath, "utf8"));
  const required = ["unit_folder", "unit_pptx_name", "lessons"];
  const missing = required.filter((key) => manifest[key] == null);
  if (missing.length) {
    throw new Error(`Manifest missing required key(s): ${missing.join(", ")}`);
  }
  if (!Array.isArray(manifest.lessons) || manifest.lessons.length === 0) {
    throw new Error("Manifest has no lessons.");
  }
  manifest.lessons.forEach((lesson, index) => {
    ["build_script", "folder", "session"].forEach((key) => {
      if (lesson[key] == null) {
        throw new Error(`Lesson ${index + 1} missing '${key}'.`);
      }
    });
  });
  return manifest;
}

function findLessonPptx(lessonDir) {
  const pptxFiles = fs.readdirSync(lessonDir)
    .filter((name) => name.toLowerCase().endsWith(".pptx"))
    .sort();
  if (pptxFiles.length !== 1) {
    throw new Error(`${lessonDir} should contain exactly one PPTX; found ${pptxFiles.length}.`);
  }
  return path.join(lessonDir, pptxFiles[0]);
}

async function validatePptx(pptxPath, issues) {
  const pptxText = extractText(pptxPath);
  issues.push(...scanTextForForbiddenOutput(pptxText, path.basename(pptxPath)));
  issues.push(...await validateNotesXml(pptxPath));
}

function validatePdfFolder(resourceDir, issues) {
  if (!fs.existsSync(resourceDir)) return 0;
  const pdfFiles = listPdfFiles(resourceDir);
  for (const pdfPath of pdfFiles) {
    const pdfText = extractText(pdfPath);
    issues.push(...scanTextForForbiddenOutput(pdfText, path.basename(pdfPath)));
  }
  return pdfFiles.length;
}

async function main() {
  const args = process.argv.slice(2);
  const manifestArg = args.find((arg) => !arg.startsWith("-"));
  const skipBuild = args.includes("--skip-build");
  const skipMerge = args.includes("--skip-merge");
  const manifestPath = resolveManifest(manifestArg);

  if (!manifestPath) {
    usage();
    process.exit(2);
  }

  const manifest = loadManifest(manifestPath);
  const issues = [];

  if (!skipBuild) {
    for (const lesson of manifest.lessons) {
      console.log(`\nBuilding ${lesson.build_script}...`);
      const output = runCommand("node", ["scripts/build_and_check.js", lesson.build_script], { timeout: 300000 });
      process.stdout.write(output);
    }
  }

  let expectedResourceCount = 0;
  for (const lesson of manifest.lessons) {
    const lessonDir = path.join(ROOT, "output", lesson.folder);
    if (!fs.existsSync(lessonDir)) {
      issues.push(`Missing lesson output folder: ${lessonDir}`);
      continue;
    }

    const pptxPath = findLessonPptx(lessonDir);
    console.log(`Validating ${path.basename(pptxPath)}...`);
    await validatePptx(pptxPath, issues);

    const resourceDir = path.join(lessonDir, `resources-session${lesson.session}`);
    expectedResourceCount += validatePdfFolder(resourceDir, issues);
  }
  if (manifest.teacher_brief) expectedResourceCount += 1;

  if (skipMerge) {
    console.log("\nSkipping merge; validating existing merged unit output...");
  } else {
    console.log("\nMerging unit...");
    const mergeOutput = runCommand("python", ["scripts/merge_unit.py", path.relative(ROOT, manifestPath)], {
      timeout: 180000,
    });
    process.stdout.write(mergeOutput);
  }

  const unitDir = path.join(ROOT, "output", manifest.unit_folder);
  const unitPptx = path.join(unitDir, manifest.unit_pptx_name);
  const resourcesDir = path.join(unitDir, "Resources");
  if (!fs.existsSync(unitPptx)) {
    issues.push(`Missing merged PPTX: ${unitPptx}`);
  } else {
    console.log(`Validating merged deck ${manifest.unit_pptx_name}...`);
    await validatePptx(unitPptx, issues);
  }

  const deliveredResourceCount = fs.existsSync(resourcesDir)
    ? validatePdfFolder(resourcesDir, issues)
    : 0;
  if (manifest.teacher_brief) {
    const briefName = manifest.teacher_brief.file_name || "Teacher Week Brief.pdf";
    const briefPath = path.join(resourcesDir, briefName);
    if (!fs.existsSync(briefPath)) {
      issues.push(`Missing teacher brief: ${briefPath}`);
    }
  }
  if (deliveredResourceCount !== expectedResourceCount) {
    issues.push(
      `Merged Resources count mismatch: expected ${expectedResourceCount}, found ${deliveredResourceCount}.`
    );
  }

  if (issues.length > 0) {
    console.error("Unit QA failed:");
    issues.forEach((issue) => console.error(`- ${issue}`));
    process.exit(1);
  }

  console.log("Unit QA passed.");
}

main().catch((error) => {
  console.error(error.message || error);
  if (error.output) process.stderr.write(error.output);
  process.exit(1);
});
