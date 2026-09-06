"use strict";

/**
 * Build a lesson deck (and its companion PDFs) from a lesson spec.
 *
 * The spec carries content and intent; every layout decision is made here
 * by the theme builders. See docs/lesson-spec.md for the schema and
 * builds/exemplar_*.json for complete lessons.
 *
 *   const { buildLesson } = require("../themes/lesson/buildLesson");
 *   await buildLesson(spec);   // -> { pptxPath, outDir, resources, warnings }
 */

const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../factory");
const { composeGlanceNotes, composeRevealNotes } = require("../core/composeNotes");
const { SAFE_BOTTOM } = require("../core/layout");
const { getElementBounds, inferElementType } = require("../core/diagnostics");
const { byBand } = require("../core/gradeBand");
const { addResourceSlide, getSessionResourceFolder } = require("../pdf_helpers");
const { validateLessonSpec } = require("./validate");
const { buildResources } = require("./resources");

const ROOT = path.resolve(__dirname, "..", "..");

function slugFolder(title) {
  return String(title).replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

function toList(value) {
  if (value == null) return [];
  return Array.isArray(value) ? value.map((v) => String(v)) : [String(value)];
}

function notesFor(slide) {
  const n = slide.notes;
  if (typeof n === "string") return n;
  return composeGlanceNotes(n);
}

function badgeColorFor(T, slide, fallback) {
  const map = { primary: T.C.PRIMARY, secondary: T.C.SECONDARY, accent: T.C.ACCENT, alert: T.C.ALERT, success: T.C.SUCCESS, assess: T.C.ASSESS };
  if (slide.badgeColor && map[slide.badgeColor]) return map[slide.badgeColor];
  const text = String(slide.badge || "").toLowerCase();
  if (/cfu|check|hinge/.test(text)) return T.C.ALERT;
  if (/we do|notice|together/.test(text)) return T.C.SECONDARY;
  if (/you do|your turn/.test(text)) return T.C.SUCCESS;
  if (/exit/.test(text)) return T.C.ASSESS;
  return fallback || T.C.PRIMARY;
}

/** Numeracy decks label the GRR stages "Stage n | Label" (megaprompt 21). */
function badgeTextFor(T, slide, fallback) {
  const text = slide.badge || fallback || "";
  if (T._subject !== "numeracy") return text;
  const stages = { "i do": 2, "we do": 3, "you do": 4 };
  const key = String(text).trim().toLowerCase();
  return stages[key] ? `Stage ${stages[key]}  |  ${text}` : text;
}

/**
 * Place an answer bar under the slide's rendered content on click. The bar
 * takes the bottom of the content area, shrinking if the content runs low,
 * and warns loudly when there is no room at all.
 */
function addClickReveal(T, s, reveal, label) {
  const answers = toList(reveal.answers);
  const objects = Array.isArray(s._slideObjects) ? s._slideObjects : [];
  let contentBottom = 0;
  objects.forEach((obj) => {
    const b = getElementBounds(obj);
    if (!b) return;
    if (inferElementType(obj) === "line") return;
    if (b.y >= SAFE_BOTTOM + 0.1) return; // footer
    contentBottom = Math.max(contentBottom, b.y + b.h);
  });
  const fullH = byBand(T.S, 1.05, 0.95, 0.8);
  let y = SAFE_BOTTOM - fullH;
  if (y < contentBottom + 0.12) y = contentBottom + 0.12;
  const h = SAFE_BOTTOM - y;
  if (h < 0.42) {
    console.warn(`WARN [lesson] ${label}: no room for the answer bar (content ends at y=${contentBottom.toFixed(2)}"). Shorten the content or move the reveal to its own slide (reveal.separate: true).`);
    return;
  }
  const fontSize = h < 0.7 ? byBand(T.S, 24, 21, 18) : undefined;
  T.clickBuild(s, [() => { T.addRevealAnswerBar(s, answers, { y, h, fontSize }); }]);
}

async function buildLesson(spec, opts) {
  const o = opts || {};
  const { errors, warnings } = validateLessonSpec(spec);
  warnings.forEach((w) => console.log(`ADVISORY [spec] ${w}`));
  if (errors.length) {
    const err = new Error(`[lesson spec] ${errors.length} error(s):\n  - ${errors.join("\n  - ")}`);
    err.specErrors = errors;
    throw err;
  }

  const L = spec.lesson;
  const variant = L.variant != null ? L.variant : weekToVariant(L.week || 1);
  const T = createTheme(L.subject, L.yearLevel, variant);
  const { C } = T;
  const session = L.session || 1;
  const footer = L.footer || `${L.title} | ${L.meta || ""}`.replace(/\s\|\s$/, "");
  const outDir = path.resolve(ROOT, o.outRoot || "output", L.outputFolder || slugFolder(L.title));
  const resDir = path.join(outDir, getSessionResourceFolder(session));
  fs.mkdirSync(resDir, { recursive: true });

  // Resources first so the Teacher Resources slide can list them.
  // makeSessionResource returns fileName WITH the resources-session{N}/
  // prefix, so PDFs are written relative to the lesson folder.
  const resourceItems = await buildResources(spec, T, outDir, session);

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  const liSlideSpec = spec.slides.find((s) => s.kind === "li");
  const scItems = liSlideSpec ? toList(liSlideSpec.successCriteria) : [];

  spec.slides.forEach((slide, index) => {
    const label = `slide ${index + 1} (${slide.kind})`;
    const notes = notesFor(slide);
    let s = null;

    switch (slide.kind) {
      case "title":
        s = T.titleSlide(pres, L.title, L.subtitle || "", L.meta || "", notes, L.titleVisual ? { visual: L.titleVisual } : undefined);
        break;

      case "overview":
        s = T.contentSlide(pres, "Teacher overview", C.MUTED, slide.title || "For the teacher: this session", toList(slide.lines), notes, footer);
        break;

      case "resources": {
        const m = spec.materials || {};
        s = addResourceSlide(pres, {
          resources: resourceItems,
          manipulatives: m.manipulatives,
          studentTools: m.studentTools,
          routineIcons: m.routineIcons,
          boardSetup: m.boardSetup,
          videos: m.videos,
          urls: m.urls,
          ochre: m.ochre,
        }, T, footer, notes);
        break;
      }

      case "dailyReview":
        s = T.dailyReviewSlide(pres, slide.title, toList(slide.prompts), notes, footer, slide.visual || undefined);
        break;

      case "fluency":
        s = T.fluencySlide(pres, slide.title, toList(slide.prompts), notes, footer);
        break;

      case "launch":
      case "content": {
        const badge = badgeTextFor(T, slide, slide.kind === "launch" ? "Launch" : "");
        const color = badgeColorFor(T, slide, C.PRIMARY);
        const lines = toList(slide.lines);
        if (slide.kind === "launch" && !lines.length && slide.visual) {
          s = T.heroVisualSlide(pres, badge, slide.title, slide.visual, notes, footer, {
            label: slide.label, prompt: slide.prompt, badgeColor: color, reserveBottom: slide.reveal ? byBand(T.S, 1.15, 1.05, 0.9) : 0,
          });
        } else {
          s = T.contentSlide(pres, badge, color, slide.title, lines, notes, footer, slide.visual || undefined);
        }
        break;
      }

      case "li":
        s = T.liSlide(pres, String(slide.learningIntention), scItems, notes, footer);
        break;

      case "keyWord":
        s = T.keyWordSlide(pres, {
          word: slide.word, meaning: slide.meaning, example: slide.example,
          pictogram: slide.pictogram, image: slide.image, routine: slide.routine,
        }, notes, footer);
        break;

      case "heroVisual":
        s = T.heroVisualSlide(pres, badgeTextFor(T, slide), slide.title, slide.visual, notes, footer, {
          label: slide.label, prompt: slide.prompt, badgeColor: badgeColorFor(T, slide),
          reserveBottom: slide.reveal ? byBand(T.S, 1.15, 1.05, 0.9) : 0,
        });
        break;

      case "workedExample":
        s = T.workedExSlide(pres, slide.stage, slide.stageLabel, slide.title, toList(slide.steps), notes, footer, slide.visual || undefined);
        break;

      case "choice": {
        const options = (slide.options || []).map((opt) => (typeof opt === "string" ? { text: opt } : opt));
        s = T.choiceSlide(pres, badgeTextFor(T, slide), slide.title, slide.prompt, options, notes, footer, {
          badgeColor: badgeColorFor(T, slide), letters: slide.letters,
        });
        if (slide.answer != null) {
          T.clickBuild(s, [() => { T.markChoice(s, slide.answer); }]);
        }
        break;
      }

      case "cfu":
        s = T.cfuSlide(pres, slide.badge || "CFU", slide.title, { technique: slide.technique, question: slide.question }, notes, footer);
        break;

      case "youDo":
        s = T.youDoSlide(pres, slide.title, slide.task, toList(slide.steps), notes, footer, {
          where: slide.where, visual: slide.visual, visualLabel: slide.visualLabel, frame: slide.frame,
          badgeText: slide.badge ? badgeTextFor(T, slide) : undefined,
          badgeColor: slide.badgeColor ? badgeColorFor(T, slide) : undefined,
        });
        break;

      case "textExtract":
        s = T.textExtractSlide(pres, badgeTextFor(T, slide), slide.title, slide.extract, notes, footer, {
          highlights: slide.highlights, source: slide.source, prompt: slide.prompt, badgeColor: badgeColorFor(T, slide),
          reserveBottom: slide.reveal ? byBand(T.S, 1.15, 1.05, 0.9) : 0,
        });
        break;

      case "cycle":
        s = T.cycleDiagramSlide(pres, slide.badge || "I Do", slide.title, slide.promptTitle, toList(slide.promptLines), slide.centerLabel, slide.steps, notes, footer,
          { reserveBottom: slide.reveal ? byBand(T.S, 1.15, 1.05, 0.9) : 0 });
        break;

      case "process":
        s = T.processFlowSlide(pres, slide.badge || "I Do", slide.title, slide.promptTitle, toList(slide.promptLines), slide.steps, notes, footer);
        break;

      case "boardBuild":
        s = T.boardBuildSlide(pres, slide.badge || "Build Together", slide.title, slide.directive, notes, footer, {
          promptText: slide.promptText, prefilledHints: slide.prefilledHints, badgeColor: slide.badgeColor ? badgeColorFor(T, slide) : undefined,
        });
        break;

      case "scenario":
        s = T.scenarioSlide(pres, slide.badge || "Scenario", slide.title, slide.scenario, toList(slide.questions), notes, footer);
        break;

      case "pairShare":
        s = T.pairShareSlide(pres, slide.title, toList(slide.questions), notes, footer);
        break;

      case "exitTicket":
        if (slide.visual) {
          // Junior exit evidence is a representation plus one prompt, never a
          // text-only card (megaprompt 0a item 2).
          s = T.heroVisualSlide(pres, "Exit Ticket", slide.title || "Show what you know", slide.visual, notes, footer, {
            label: slide.label, prompt: toList(slide.questions)[0], badgeColor: C.ASSESS,
          });
        } else {
          s = T.exitTicketSlide(pres, toList(slide.questions), notes, footer, { title: slide.title });
        }
        break;

      case "closing":
        s = T.closingSlide(pres, {
          reflectionPrompt: slide.reflectionPrompt, scItems, selfAssessment: slide.selfAssessment, takeaways: slide.takeaways,
        }, notes);
        break;

      default:
        throw new Error(`[lesson] unhandled slide kind ${slide.kind}`);
    }

    // Reveals: click build on the same slide (default) or a separate slide.
    if (slide.reveal && slide.kind !== "choice") {
      if (slide.reveal.separate) {
        const revealNotes = composeRevealNotes(slide.reveal.notes);
        T.withReveal(() => s, (rs) => {
          const answers = toList(slide.reveal.answers);
          const h = byBand(T.S, 1.05, 0.95, 0.8);
          T.addRevealAnswerBar(rs, answers, { y: SAFE_BOTTOM - h, h });
        }, { revealNotes });
      } else {
        addClickReveal(T, s, slide.reveal, label);
      }
    }
  });

  const fileName = L.fileName || `${L.title}.pptx`;
  const pptxPath = path.join(outDir, fileName);
  await pres.writeFile({ fileName: pptxPath });
  console.log(`PPTX written to ${pptxPath}`);
  resourceItems.forEach((r) => console.log(`Resource: ${path.join(getSessionResourceFolder(session), path.basename(r.fileName))}`));
  return { pptxPath, outDir, resources: resourceItems, warnings };
}

function loadSpec(specPath) {
  const abs = path.resolve(specPath);
  if (abs.endsWith(".json")) return JSON.parse(fs.readFileSync(abs, "utf8"));
  // eslint-disable-next-line global-require, import/no-dynamic-require
  return require(abs);
}

module.exports = { buildLesson, loadSpec };
