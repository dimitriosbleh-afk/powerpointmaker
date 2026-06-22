"use strict";

const assert = require("assert");
const pptxgen = require("pptxgenjs");

const { createTheme } = require("../themes/factory");
const { contrastRatio } = require("../themes/core/contrast");
const { getSlideNotesText, sanitizeTeacherNotes } = require("../themes/core/notes");
const { addResourceSlide } = require("../themes/pdf_helpers");

function textValue(obj) {
  if (!obj) return "";
  if (typeof obj.text === "string") return obj.text;
  if (Array.isArray(obj.text)) {
    return obj.text.map((run) => run && run.text ? run.text : "").join("");
  }
  return "";
}

function assertTextContrast(slide, text, bg, label) {
  const obj = slide._slideObjects.find((entry) => textValue(entry) === text);
  assert(obj, `${label}: expected to find text "${text}"`);
  const color = obj.options && obj.options.color;
  assert(color, `${label}: text object has no colour`);
  assert(
    contrastRatio(color, bg) >= 4.5,
    `${label}: ${color} on ${bg} fails AA contrast`
  );
}

function testNotesAggregation() {
  const slide = {
    _slideObjects: [
      { _type: "notes", text: [{ text: "SAY:\n- Ask the question." }] },
      { _type: "notes", text: [{ text: "TEACHER NOTES:\n- Debrief the answer." }] },
    ],
  };
  assert.strictEqual(
    getSlideNotesText(slide),
    "SAY:\n- Ask the question.\n\nTEACHER NOTES:\n- Debrief the answer."
  );
}

function testDivisionSanitizer() {
  const sanitized = sanitizeTeacherNotes("Expected: 12 ÷ 3 = 4");
  assert(sanitized.includes("12 divided by 3"), sanitized);
  assert(!sanitized.includes("12 / 3"), sanitized);

  const mojibake = sanitizeTeacherNotes("Expected: 12 Ã· 3 = 4");
  assert(mojibake.includes("12 divided by 3"), mojibake);
  assert(!mojibake.includes("12 / 3"), mojibake);

  const legacy = sanitizeTeacherNotes("Expected: 12 \u00C3\u0192\u00C2\u00B7 3 = 4");
  assert(legacy.includes("12 divided by 3"), legacy);
  assert(!legacy.includes("12 / 3"), legacy);
}

function testLiteracyContrast() {
  const levels = ["foundation", "grade1", "grade2", "grade34", "grade56"];
  levels.forEach((level) => {
    for (let variant = 0; variant < 6; variant += 1) {
      const T = createTheme("literacy", level, variant);
      const pres = new pptxgen();
      pres.layout = "LAYOUT_16x9";
      const vocab = T.vocabSlide(
        pres,
        "infer",
        "noun",
        "A careful idea from clues.",
        "I infer the mood from the words.",
        "",
        ""
      );
      assertTextContrast(vocab, "noun", T.C.ACCENT, `vocab ${level}/${variant}`);

      const quote = T.quoteSlide(
        pres,
        "Read Aloud",
        "Chapter 1",
        "The old house waited in silence.",
        "p. 12",
        "What mood is created?",
        "",
        ""
      );
      assertTextContrast(quote, "p. 12", T.C.PRIMARY, `quote ${level}/${variant}`);
    }
  });
}

function testResourceSlideFiveCardsStayInBounds() {
  const T = createTheme("literacy", "grade56", 0);
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  const resources = Array.from({ length: 5 }, (_, index) => ({
    name: `Session ${index + 1} Resource`,
    fileName: `resources-session${index + 1}/Session ${index + 1} Resource.pdf`,
    description: "Teacher printout",
  }));
  const slide = addResourceSlide(pres, resources, T, "Footer", "SAY:\n- Use the links.");

  slide._slideObjects.forEach((obj) => {
    const options = obj && obj.options;
    if (!options || !Number.isFinite(options.y) || !Number.isFinite(options.h)) return;
    const isFooter = options.y >= 5.25 && options.w >= 7.5;
    if (!isFooter) {
      assert(
        options.y + options.h <= 5.11,
        `non-footer element exceeds safe bottom: y=${options.y}, h=${options.h}`
      );
    }
  });
}

function testResourceSlideDenseCardsStayPositive() {
  const T = createTheme("literacy", "grade56", 0);
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  const resources = Array.from({ length: 20 }, (_, index) => ({
    name: `Resource ${index + 1}`,
    fileName: `resources-session1/Resource ${index + 1}.pdf`,
    description: "Teacher printout",
  }));
  const originalWarn = console.warn;
  const warnings = [];
  console.warn = (message) => warnings.push(String(message));
  try {
    const slide = addResourceSlide(pres, resources, T, "", "SAY:\n- Use the links.");
    slide._slideObjects.forEach((obj) => {
      const options = obj && obj.options;
      if (!options) return;
      if (Number.isFinite(options.h)) {
        assert(options.h > 0, `element has non-positive height: ${options.h}`);
      }
      if (Number.isFinite(options.w)) {
        assert(options.w > 0, `element has non-positive width: ${options.w}`);
      }
    });
  } finally {
    console.warn = originalWarn;
  }
  assert(
    warnings.some((message) => message.includes("resource list is too long")),
    "expected dense resource list warning"
  );
}

testNotesAggregation();
testDivisionSanitizer();
testLiteracyContrast();
testResourceSlideFiveCardsStayInBounds();
testResourceSlideDenseCardsStayPositive();

console.log("Review regression tests passed.");
