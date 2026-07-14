"use strict";

const fs = require("fs");
const path = require("path");

const {
  normaliseTeacherBrief,
  generateTeacherWeekBrief,
} = require("../themes/teacher_brief");

const ROOT = path.resolve(__dirname, "..");

function usage() {
  console.error("Usage: node scripts/build_teacher_brief.js <unit-manifest.json>");
}

async function main() {
  const manifestArg = process.argv[2];
  if (!manifestArg) {
    usage();
    process.exit(2);
  }

  const manifestPath = path.resolve(ROOT, manifestArg);
  if (!fs.existsSync(manifestPath)) {
    throw new Error(`Manifest not found: ${manifestPath}`);
  }

  const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
  if (!manifest.teacher_brief) {
    console.log("No teacher_brief configuration in manifest; nothing to generate.");
    return;
  }

  const opts = {
    unitTitle: path.basename(manifest.unit_pptx_name || "", path.extname(manifest.unit_pptx_name || "")),
    lessonSessions: (manifest.lessons || []).map((lesson) => String(lesson.session)),
  };
  const data = normaliseTeacherBrief(manifest.teacher_brief, opts);
  const outputPath = path.join(ROOT, "output", manifest.unit_folder, "Resources", data.fileName);

  await generateTeacherWeekBrief(manifest.teacher_brief, outputPath, opts);
  console.log(`Teacher brief written to ${outputPath}`);
}

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
