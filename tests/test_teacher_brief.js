"use strict";

const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");

const {
  normaliseTeacherBrief,
  generateTeacherWeekBrief,
} = require("../themes/teacher_brief");

const FIXTURE = {
  title: "Teacher Week Brief",
  unit: "Fractions as numbers",
  grade: "Years 3-4",
  subject: "Mathematics",
  curriculum: "Mathematics 2.0, Number",
  throughline: "Students move from naming equal parts to locating, comparing and explaining fractions on shared visual models.",
  high_leverage_moves: [
    "Pre-cue the question, give think time, then require every student to show an answer.",
    "Keep the fraction strip visible while language and symbols are connected.",
    "Use each check to choose the next move, not simply the next slide.",
  ],
  sessions: [
    {
      session: 1,
      focus: "Name unit fractions",
      students: "Build and label equal parts with strips.",
      check: "All boards show the denominator as the total equal parts.",
      respond: "If mixed, rebuild one whole and count every part before rechecking.",
    },
    {
      session: 2,
      focus: "Locate fractions",
      students: "Place fractions on a number line.",
      check: "Students use equal intervals and start from zero.",
      respond: "If the spacing drifts, fold a strip first and transfer the marks.",
    },
    {
      session: 3,
      focus: "Compare fractions",
      students: "Compare fractions using one shared whole.",
      check: "Students justify the comparison from the model, not the digits alone.",
      respond: "If digits dominate, align two strips over the same-length whole and re-ask.",
    },
  ],
  essential_knowledge: [
    "The denominator names how many equal parts make one whole.",
    "Comparisons are valid only when the wholes are the same size.",
  ],
  misconceptions: [
    "A larger denominator means a larger piece.",
    "The tick marks, rather than the intervals, are counted on a number line.",
  ],
  response_rule: "Secure and complete responses -> move on. Mixed responses -> one more guided pair. A shared misconception -> re-model with a different representation, then recheck everyone.",
  rehearse: [
    "Practise the first strip model and the exact hinge question aloud.",
    "Decide where you will stand to scan every board quickly.",
  ],
  materials: ["Fraction strips", "Mini-whiteboards", "Printed practice sheet"],
};

async function main() {
  const opts = { unitTitle: FIXTURE.unit, lessonSessions: ["1", "2", "3"] };
  const normalised = normaliseTeacherBrief(FIXTURE, opts);
  assert.strictEqual(normalised.sessions.length, 3);
  assert.strictEqual(normalised.fileName, "Teacher Week Brief.pdf");

  assert.throws(
    () => normaliseTeacherBrief(FIXTURE, { ...opts, lessonSessions: ["1", "2", "4"] }),
    /do not match manifest lessons/
  );

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "teacher-brief-test-"));
  const outputPath = path.join(tempDir, normalised.fileName);
  await generateTeacherWeekBrief(FIXTURE, outputPath, opts);
  const pdf = fs.readFileSync(outputPath);
  assert(pdf.length > 10000, "teacher brief PDF is unexpectedly small");
  assert.strictEqual(pdf.subarray(0, 4).toString("ascii"), "%PDF");
  assert.strictEqual((pdf.toString("latin1").match(/\/Type\s*\/Page\b/g) || []).length, 1);
  fs.rmSync(tempDir, { recursive: true, force: true });

  console.log("Teacher brief tests passed.");
}

if (require.main === module) {
  main().catch((error) => {
    console.error(error);
    process.exit(1);
  });
}

module.exports = { FIXTURE };
