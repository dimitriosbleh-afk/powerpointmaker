"use strict";

/**
 * Click-to-reveal builds (megaprompt section 20b).
 *
 * PptxGenJS cannot write animations: it emits no <p:timing> element at all.
 * Until now the only reveal mechanism available to a themed deck was
 * withReveal(), which duplicates the whole slide so that clicking to the NEXT
 * SLIDE shows the answer. That doubles the slide count, splits one teaching
 * moment across two slides, and forces the reveal half to carry its own notes.
 *
 * A click build is the real thing: one slide, and each click makes one more
 * element appear on it. This module records the plan at build time and injects
 * the <p:timing> tree into the finished file.
 *
 * The shape id mapping is the load-bearing fact. PptxGenJS numbers shapes
 * `<p:cNvPr id="N">` where N is the element's index in slide._slideObjects
 * PLUS 2 (id 1 is the spTree root). Objects that render nothing - notably
 * `notes` - still consume an index, so their id is simply skipped. That makes
 * the id predictable from the array position at build time, which is what
 * lets us target shapes without parsing the written XML.
 *
 * Usage in a build script:
 *
 *   const s = T.contentSlide(pres, "Title", [...], notes, footer);
 *   T.clickBuild(s, [
 *     () => { s.addText("first reveal", {...}); },
 *     () => { s.addShape(...); s.addText("second reveal", {...}); },
 *   ]);
 *
 * Each thunk is one click. Everything a thunk adds appears together on that
 * click. Anything added outside a thunk is visible from the start.
 */

const fs = require("fs");
const JSZip = require("jszip");

// PowerPoint "Appear" entrance effect.
const PRESET_ID = 1;
const PRESET_CLASS = "entr";

/**
 * Record a click-to-reveal plan on a slide.
 *
 * @param {object}     slide  a PptxGenJS slide
 * @param {Function[]} steps  one thunk per click; each adds the elements that
 *                            should appear together on that click
 * @returns {number[][]} the recorded plan as arrays of shape ids
 */
function clickBuild(slide, steps) {
  if (!slide || !Array.isArray(slide._slideObjects)) {
    throw new Error("[clickBuild] expected a PptxGenJS slide");
  }
  if (!Array.isArray(steps) || steps.length === 0) {
    throw new Error("[clickBuild] expected a non-empty array of step functions");
  }

  const plan = [];
  steps.forEach((step, stepIndex) => {
    if (typeof step !== "function") {
      throw new Error(`[clickBuild] step ${stepIndex + 1} is not a function`);
    }
    const before = slide._slideObjects.length;
    step();
    const after = slide._slideObjects.length;

    const ids = [];
    for (let i = before; i < after; i += 1) {
      const obj = slide._slideObjects[i];
      // `notes` consumes an array slot but renders no shape, so it has no id
      // to target. Skipping it here keeps the index-to-id maths honest.
      if (obj && obj._type === "notes") continue;
      ids.push(i + 2);
    }

    if (ids.length === 0) {
      console.warn(
        `WARN [clickBuild] Slide ${slide._slideNum || "?"}: step ${stepIndex + 1} added no visible element. ` +
        "A click that reveals nothing leaves the teacher clicking at an unchanged slide."
      );
      return;
    }
    plan.push(ids);
  });

  if (plan.length) {
    // Appending rather than replacing lets a builder pre-seed a plan and a
    // build script extend it without silently dropping the earlier steps.
    slide._clickBuild = (slide._clickBuild || []).concat(plan);
  }
  return plan;
}

/**
 * Build the <p:timing> tree for one slide.
 * @param {number[][]} plan  one inner array per click, holding shape ids
 */
function buildTimingXml(plan) {
  let nextId = 3; // 1 = tmRoot, 2 = mainSeq
  const id = () => nextId++;

  const clickGroups = plan.map((shapeIds) => {
    const effects = shapeIds.map((spid, index) => {
      // The first effect of a click is triggered BY the click; the rest run
      // alongside it, which is how PowerPoint groups simultaneous entrances.
      const nodeType = index === 0 ? "clickEffect" : "withEffect";
      return (
        `<p:par><p:cTn id="${id()}" presetID="${PRESET_ID}" presetClass="${PRESET_CLASS}"` +
        ` presetSubtype="0" fill="hold" nodeType="${nodeType}">` +
        `<p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
        `<p:childTnLst><p:set>` +
        `<p:cBhvr>` +
        `<p:cTn id="${id()}" dur="1" fill="hold">` +
        `<p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
        `</p:cTn>` +
        `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
        `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>` +
        `</p:cBhvr>` +
        `<p:to><p:strVal val="visible"/></p:to>` +
        `</p:set></p:childTnLst>` +
        `</p:cTn></p:par>`
      );
    }).join("");

    return (
      `<p:par><p:cTn id="${id()}" fill="hold">` +
      `<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>` +
      `<p:childTnLst>` +
      `<p:par><p:cTn id="${id()}" fill="hold">` +
      `<p:stCondLst><p:cond delay="0"/></p:stCondLst>` +
      `<p:childTnLst>${effects}</p:childTnLst>` +
      `</p:cTn></p:par>` +
      `</p:childTnLst>` +
      `</p:cTn></p:par>`
    );
  }).join("");

  return (
    `<p:timing><p:tnLst>` +
    `<p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">` +
    `<p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek">` +
    `<p:cTn id="2" dur="indefinite" nodeType="mainSeq">` +
    `<p:childTnLst>${clickGroups}</p:childTnLst>` +
    `</p:cTn>` +
    `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
    `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>` +
    `</p:seq>` +
    `</p:childTnLst>` +
    `</p:cTn></p:par>` +
    `</p:tnLst></p:timing>`
  );
}

/**
 * Verify every targeted shape id exists in the slide XML.
 *
 * A timing tree pointing at a shape that is not there is the defect the
 * supplied-deck work surfaced: PowerPoint shows the slide but the click does
 * nothing, and nothing in the file looks wrong. Catch it at build time.
 */
function findDanglingTargets(slideXml, plan) {
  const present = new Set(
    [...slideXml.matchAll(/<p:cNvPr\s+id="(\d+)"/g)].map((m) => Number(m[1]))
  );
  const dangling = [];
  plan.forEach((ids, stepIndex) => {
    ids.forEach((spid) => {
      if (!present.has(spid)) dangling.push({ step: stepIndex + 1, spid });
    });
  });
  return dangling;
}

/**
 * Inject recorded click builds into a written PPTX, in place.
 *
 * @param {string}   pptxPath
 * @param {object[]} slides   pres.slides, in order
 * @returns {Promise<number>} how many slides received a timing tree
 */
async function injectClickBuildsInFile(pptxPath, slides) {
  if (!pptxPath || !Array.isArray(slides) || slides.length === 0) return 0;
  if (!slides.some((s) => s && Array.isArray(s._clickBuild) && s._clickBuild.length)) return 0;

  const zip = await JSZip.loadAsync(await fs.promises.readFile(pptxPath));
  let injected = 0;

  for (let index = 0; index < slides.length; index += 1) {
    const slide = slides[index];
    const plan = slide && slide._clickBuild;
    if (!Array.isArray(plan) || plan.length === 0) continue;

    const xmlPath = `ppt/slides/slide${index + 1}.xml`;
    const file = zip.file(xmlPath);
    if (!file) continue;

    let xml = await file.async("string");

    if (xml.includes("<p:timing")) {
      console.warn(
        `WARN [clickBuild] ${xmlPath} already carries a timing tree; leaving it alone.`
      );
      continue;
    }

    const dangling = findDanglingTargets(xml, plan);
    if (dangling.length) {
      const detail = dangling.map((d) => `step ${d.step} -> shape ${d.spid}`).join(", ");
      throw new Error(
        `[clickBuild] slide ${index + 1}: timing targets shapes that do not exist (${detail}). ` +
        "The click would do nothing. This usually means elements were removed after clickBuild() ran."
      );
    }

    // CT_Slide order is cSld, clrMapOvr, transition, timing, extLst - so the
    // tree goes immediately before the closing tag.
    xml = xml.replace("</p:sld>", `${buildTimingXml(plan)}</p:sld>`);
    zip.file(xmlPath, xml);
    injected += 1;
  }

  if (injected > 0) {
    const output = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    await fs.promises.writeFile(pptxPath, output);
  }
  return injected;
}

module.exports = {
  clickBuild,
  buildTimingXml,
  findDanglingTargets,
  injectClickBuildsInFile,
};
