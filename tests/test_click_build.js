"use strict";

/**
 * Click build regression guard (megaprompt section 20b).
 *
 * Covers the two things that silently break a reveal:
 *  - the shape id mapping (PptxGenJS id = _slideObjects index + 2, with
 *    non-rendering objects consuming an index)
 *  - the timing tree actually landing in the written file
 *
 * Run:  node tests/test_click_build.js
 */

const assert = require("assert");
const fs = require("fs");
const path = require("path");
const os = require("os");
const JSZip = require("jszip");
const pptxgen = require("pptxgenjs");
const { createTheme } = require("../themes/factory");
const {
  clickBuild,
  buildTimingXml,
  findDanglingTargets,
} = require("../themes/core/animations");

let passed = 0;
function ok(label) {
  console.log("  PASS " + label);
  passed += 1;
}

/* ── 1. id mapping, including the notes-consumes-an-index case ───────────── */

function testIdMapping() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  const s = pres.addSlide();

  s.addText("always visible", { x: 1, y: 0.5, w: 3, h: 0.4 }); // index 0 -> id 2
  s.addNotes("notes consume an index but render no shape");     // index 1 -> no id

  const plan = clickBuild(s, [
    () => { s.addText("step one", { x: 1, y: 1.5, w: 3, h: 0.4 }); },        // index 2 -> id 4
    () => {
      s.addShape("rect", { x: 1, y: 2.5, w: 2, h: 0.4, fill: { color: "FF0000" } }); // index 3 -> id 5
      s.addText("step two b", { x: 1, y: 3.5, w: 3, h: 0.4 });               // index 4 -> id 6
    },
  ]);

  assert.deepStrictEqual(plan, [[4], [5, 6]], "plan should skip the notes index: " + JSON.stringify(plan));
  assert.deepStrictEqual(s._clickBuild, [[4], [5, 6]]);
  ok("shape ids skip non-rendering objects");
}

/* ── 2. timing XML shape ─────────────────────────────────────────────────── */

function testTimingXml() {
  const xml = buildTimingXml([[4], [5, 6]]);

  assert(xml.startsWith("<p:timing>"), "must be a timing element");
  assert(xml.includes('nodeType="tmRoot"'), "needs a tmRoot");
  assert(xml.includes('nodeType="mainSeq"'), "needs a mainSeq");
  assert.strictEqual((xml.match(/nodeType="clickEffect"/g) || []).length, 2,
    "one clickEffect per click group");
  assert.strictEqual((xml.match(/nodeType="withEffect"/g) || []).length, 1,
    "the second shape in click 2 runs alongside the first");
  assert(xml.includes('<p:spTgt spid="4"/>'), "targets shape 4");
  assert(xml.includes('<p:spTgt spid="6"/>'), "targets shape 6");
  assert(xml.includes('<p:strVal val="visible"/>'), "sets visibility");

  // Every cTn id must be unique or PowerPoint rejects the file.
  const ids = [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => m[1]);
  assert.strictEqual(ids.length, new Set(ids).size, "cTn ids must be unique: " + ids.join(","));
  ok("timing tree is well formed with unique ids");
}

/* ── 3. dangling target detection ────────────────────────────────────────── */

function testDanglingDetection() {
  const fakeXml = '<p:cNvPr id="1" name=""/><p:cNvPr id="2" name="Text 0">';
  assert.strictEqual(findDanglingTargets(fakeXml, [[2]]).length, 0, "id 2 exists");
  const bad = findDanglingTargets(fakeXml, [[2], [9]]);
  assert.strictEqual(bad.length, 1, "id 9 is dangling");
  assert.strictEqual(bad[0].step, 2);
  assert.strictEqual(bad[0].spid, 9);
  ok("dangling animation targets are detected");
}

/* ── 4. end to end: timing survives into the written file ────────────────── */

async function testEndToEnd() {
  const T = createTheme("numeracy", "grade56", 0);
  const p = new pptxgen();
  p.layout = "LAYOUT_16x9";

  const s = p.addSlide();
  s.addText("What is 12 divided by 3?", { x: 0.5, y: 1.4, w: 9, h: 0.8, fontSize: 30 });
  s.addNotes("ANSWER: four\n\n1. ASK: What is 12 divided by 3?\n   10 sec. Cue: Write it... Chin it... Show me.\n   EXPECT: four\n---\nClick build reveals the answer after boards are scanned.");
  T.clickBuild(s, [
    () => { s.addText("4", { x: 0.5, y: 2.6, w: 9, h: 1, fontSize: 54, bold: true }); },
  ]);

  // A second slide with NO build must not gain a timing tree.
  const plain = p.addSlide();
  plain.addText("no build here", { x: 1, y: 1, w: 4, h: 0.5 });

  const outDir = fs.mkdtempSync(path.join(os.tmpdir(), "clickbuild-"));
  const outFile = path.join(outDir, "click_build_test.pptx");
  await p.writeFile({ fileName: outFile });

  const zip = await JSZip.loadAsync(fs.readFileSync(outFile));
  const xml1 = await zip.file("ppt/slides/slide1.xml").async("string");
  const xml2 = await zip.file("ppt/slides/slide2.xml").async("string");

  assert(xml1.includes("<p:timing>"), "slide 1 should carry a timing tree");
  assert(xml1.includes('<p:spTgt spid="4"/>'),
    "slide 1 should target the answer shape (index 2 + 2 = 4)");
  assert(xml1.indexOf("<p:timing>") > xml1.indexOf("</p:cSld>"),
    "timing must sit after cSld per the CT_Slide sequence");
  assert(!xml2.includes("<p:timing"), "slide 2 has no build and must stay untouched");

  // The notes rewrite must still have run in the same post-write pass.
  const notesXml = await zip.file("ppt/notesSlides/notesSlide1.xml").async("string");
  assert(notesXml.includes("ANSWER:"), "notes rewrite must survive alongside the timing injection");
  assert(notesXml.includes('b="1"'), "note anchors should still be bolded");

  fs.rmSync(outDir, { recursive: true, force: true });
  ok("timing tree lands in the written file and coexists with the notes rewrite");
}

/* ── run ─────────────────────────────────────────────────────────────────── */

(async () => {
  console.log("\nClick build regressions");
  testIdMapping();
  testTimingXml();
  testDanglingDetection();
  await testEndToEnd();
  console.log(`\n${passed} check(s) passed.\n`);
})().catch((err) => {
  console.error("\nFAIL " + (err && err.stack ? err.stack : err));
  process.exit(1);
});
