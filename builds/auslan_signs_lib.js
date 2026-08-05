"use strict";

// Shared sign-image layer for Auslan decks (AUSLAN_2_SLIDES_PROMPT section 3).
// Used by all ten build_deafsport_lesson*.js scripts.
//
// NEVER draw, generate, mirror or approximate a sign. Images come only from
// assets/auslan_signs/, built by scripts/fetch_auslan_signs.py from Auslan
// Signbank. A gloss with no image gets a lookup card, never a drawing.
//
// SENSE MAP (the important part). Signbank's FIRST entry for a search word is
// often NOT the sense a lesson teaches, and a plausible image of the wrong sign
// is invisible on a rendered slide. Every gloss below was checked against the
// recorded definition in assets/auslan_signs/manifest.json on 2026-08-03.
// Where the default file was the wrong sense, the correct variant is pinned here.

const fs = require("fs");
const path = require("path");

const BANK = path.join(__dirname, "..", "assets", "auslan_signs");

// Glosses where the default <GLOSS>.jpg is the WRONG SENSE. Verified corrections.
const SENSE_CORRECTIONS = {
  WIN: "WIN_3.jpg",           // WIN.jpg is CELEBRATION (party/cheering), not winning
  NEXT: "NEXT_2.jpg",         // NEXT.jpg is DEMOTION, not the next event
  PRACTISE: "PRACTISE_3.jpg", // PRACTISE.jpg is LEARNER/TRAINEE; _3 is train, sport sense
  CORRECT: "CORRECT_2.jpg",   // CORRECT.jpg is TICK (the written mark); _2 is "that's right"
};

// Glosses deliberately given a lookup card even though a file may exist, or
// because Signbank has no entry. Value is the English word for the search URL.
const FORCE_LOOKUP = {
  MEAN: "meaning",  // MEAN.jpg's first sense is SIGN LANGUAGE; do not assert it
  BRONZE: "bronze", // no Signbank entry at all
};

// Signs whose image is usable but whose sense is close enough to be worth a
// teacher rehearsal line. Surfaced into the notes prep zone by the lesson.
const REHEARSE_FIRST = {
  WHAT: "WHAT has one Signbank entry and its first sense is the interactive 'What?!'. Confirm the question form from the school reference.",
  LESS: "LESS shares its sign with SMALL. Confirm the quantity sense before teaching it.",
  HISTORY: "HISTORY.jpg is the long-ago sense. HISTORY_3 is the school-subject form if you want that one.",
  PROUD: "PROUD also carries boast/brag. Confirm the pride sense.",
};

/** Resolve a gloss to a bank file, honouring the sense corrections. */
function resolveSign(gloss) {
  if (FORCE_LOOKUP[gloss]) return null;
  const pinned = SENSE_CORRECTIONS[gloss];
  if (pinned && fs.existsSync(path.join(BANK, pinned))) {
    return path.join(BANK, pinned);
  }
  for (const ext of [".jpg", ".png"]) {
    const p = path.join(BANK, gloss + ext);
    if (fs.existsSync(p)) return p;
  }
  return null;
}

/** Signbank SEARCH url. Never construct an entry url; suffixes are unpredictable. */
function lookupUrl(englishWord) {
  return "https://auslan.org.au/dictionary/search/?query=" +
    encodeURIComponent(String(englishWord).toLowerCase());
}

/**
 * Place one sign on a card. English meaning is the readable label; the gloss is
 * a small caption. Falls back to a lookup card when the bank has no image.
 * Returns { gloss, ok } for the SIGN ASSETS report.
 */
const LABEL_H = 0.42;
const CAP_H = 0.3;
const PAD = 0.12;

/**
 * Height a card needs so the strip fills its width with no dead space.
 * Signbank strips are wide and short (ratio 2.1 to 2.9), so a square-ish image
 * box leaves a band of white above the label. Size from the real pixels.
 */
function cardHeightFor(T, specs, cardW) {
  const imgW = cardW - PAD * 2;
  let tallest = imgW / 2.9; // widest strip in the bank, so the shortest image
  specs.forEach((spec) => {
    const file = resolveSign(spec.gloss);
    if (!file) {
      tallest = Math.max(tallest, 1.05); // lookup card needs room for the link
      return;
    }
    try {
      const d = T.getImageDimensions(file);
      tallest = Math.max(tallest, imgW / (d.width / d.height));
    } catch (err) {
      tallest = Math.max(tallest, imgW / 2.4);
    }
  });
  return tallest + LABEL_H + CAP_H + PAD * 2 + 0.1;
}

function signCard(T, slide, spec, box) {
  const { C, FONT_H, FONT_B, addTextOnShape, addImageWithCaption } = T;
  const { gloss, meaning } = spec;
  const { x, y, w, h } = box;
  const file = resolveSign(gloss);

  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.1,
    fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
  });

  const labelH = LABEL_H;
  const capH = CAP_H;
  const imgY = y + PAD;
  const imgH = h - labelH - capH - PAD * 2 - 0.1;

  if (file) {
    addImageWithCaption(slide, file, {
      x: x + 0.12, y: imgY, w: w - 0.24, h: imgH, fit: "contain",
    });
  } else {
    // Lookup card: no drawing, no empty frame.
    const linkH = 0.28;
    addTextOnShape(slide, "Watch the teacher", {
      x: x + 0.12, y: imgY, w: w - 0.24, h: Math.max(0.4, imgH - linkH - 0.04),
      rectRadius: 0.08,
      fill: { color: C.BG_LIGHT }, line: { color: C.MUTED, width: 1 },
    }, { fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL });
    slide.addText([{
      text: "Look it up: Auslan Signbank",
      options: { hyperlink: { url: lookupUrl(spec.lookupWord || meaning) }, color: C.SECONDARY },
    }], {
      x: x + 0.12, y: imgY + Math.max(0.4, imgH - linkH - 0.04) + 0.04,
      w: w - 0.24, h: linkH,
      fontSize: 9.5, fontFace: FONT_B, align: "center", margin: 0,
    });
  }

  addTextOnShape(slide, String(meaning), {
    x: x + 0.12, y: y + h - labelH - capH - 0.06, w: w - 0.24, h: labelH,
    rectRadius: 0.06, fill: { color: C.PRIMARY },
  }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

  slide.addText(String(gloss), {
    x: x + 0.12, y: y + h - capH - 0.02, w: w - 0.24, h: capH,
    fontSize: 10, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
  });

  return { gloss, ok: Boolean(file) };
}

/**
 * Lay a row of sign cards across the content area, 1 to 4 per row.
 * Card height is derived from the images unless the caller pins `h`, and the
 * row is centred in the band between `y` and SAFE_BOTTOM so it never floats.
 */
function signCardRow(T, slide, specs, opts) {
  const o = opts || {};
  const top = o.y != null ? o.y : 1.5;
  const left = o.x != null ? o.x : 0.5;
  const total = o.w != null ? o.w : 9.0;
  const gap = 0.22;
  const w = (total - gap * (specs.length - 1)) / specs.length;
  const h = o.h != null ? o.h : cardHeightFor(T, specs, w);
  const band = (o.bottom != null ? o.bottom : 5.05) - top;
  const y = o.center === false ? top : top + Math.max(0, (band - h) / 2);
  return specs.map((spec, i) =>
    signCard(T, slide, spec, { x: left + i * (w + gap), y, w, h }));
}

/**
 * addInstructionCard takes {text, role} items, not strings; strings render as
 * an empty card that no diagnostic can see. This wrapper takes a title and a
 * list of plain steps and builds the run objects.
 */
function stepCard(T, slide, opts) {
  const items = [];
  if (opts.title) items.push({ text: opts.title, role: "header" });
  (opts.steps || []).forEach((step) => items.push({ text: step, role: "body" }));
  T.addInstructionCard(slide, items, {
    x: opts.x, y: opts.y, w: opts.w, h: opts.h,
    strip: opts.strip || T.C.SECONDARY,
  });
}

/**
 * A whole "here are today's signs" slide: title, one lead line, a centred row
 * of sign cards, footer, diagnostics. The bank image IS the visual anchor, so
 * this is the standard I Do placement rather than a bullet list.
 */
function signSlide(T, pres, opts) {
  const { C, FONT_B } = T;
  const s = pres.addSlide();
  s.background = { color: C.BG_LIGHT };
  T.addTopBar(s, opts.barColor || C.PRIMARY);
  T.addTitle(s, opts.title);

  let top = 1.42;
  if (opts.lead) {
    s.addText(String(opts.lead), {
      x: 0.5, y: top, w: 9.0, h: 0.42,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
    });
    top += 0.52;
  }

  const results = signCardRow(T, s, opts.signs, { y: top, bottom: 5.02 });
  T.addFooter(s, opts.footer);
  s.addNotes(opts.notes);
  T.runSlideDiagnostics(s, pres);
  return { slide: s, results };
}

/**
 * Retrieval reveal: sign cards with the meaning label hidden until a click.
 * Students sign back or name the meaning before the answer lands. Only build
 * this from glosses that HAVE bank images; a hidden lookup card teaches nothing.
 */
function signRecallSlide(T, pres, opts) {
  const { C, FONT_H, FONT_B, addTextOnShape } = T;
  const usable = opts.signs.filter((sp) => resolveSign(sp.gloss));
  const s = pres.addSlide();
  s.background = { color: C.BG_LIGHT };
  T.addTopBar(s, opts.barColor || C.SECONDARY);
  T.addTitle(s, opts.title);

  let top = 1.42;
  s.addText(String(opts.lead || "What does each one mean?"), {
    x: 0.5, y: top, w: 9.0, h: 0.42,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
  });
  top += 0.52;

  // Cards without their meaning label, then the labels arrive on one click.
  const gap = 0.22;
  const w = (9.0 - gap * (usable.length - 1)) / usable.length;
  const imgW = w - PAD * 2;
  let imgH = imgW / 2.9;
  usable.forEach((sp) => {
    try {
      const d = T.getImageDimensions(resolveSign(sp.gloss));
      imgH = Math.max(imgH, imgW / (d.width / d.height));
    } catch (err) { imgH = Math.max(imgH, imgW / 2.4); }
  });
  const cardH = imgH + LABEL_H + PAD * 2 + 0.1;
  const y = top + Math.max(0, (5.02 - top - cardH) / 2);

  usable.forEach((sp, i) => {
    const x = 0.5 + i * (w + gap);
    s.addShape("roundRect", {
      x, y, w, h: cardH, rectRadius: 0.1,
      fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
    });
    T.addImageWithCaption(s, resolveSign(sp.gloss), {
      x: x + PAD, y: y + PAD, w: imgW, h: imgH, fit: "contain",
    });
  });

  T.clickBuild(s, [() => {
    usable.forEach((sp, i) => {
      const x = 0.5 + i * (w + gap);
      addTextOnShape(s, String(sp.meaning), {
        x: x + PAD, y: y + cardH - LABEL_H - PAD, w: imgW, h: LABEL_H,
        rectRadius: 0.06, fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    });
  }]);

  T.addFooter(s, opts.footer);
  s.addNotes(opts.notes);
  T.runSlideDiagnostics(s, pres);
  return { slide: s, results: usable.map((sp) => ({ gloss: sp.gloss, ok: true })) };
}

/**
 * Teacher Resources slide for a session that prints nothing new. Section 44
 * wants one on every lesson deck, so a boards-only session still names what the
 * teacher needs on the bench and carries the sign-image attribution.
 */
function noPrintResourceSlide(T, pres, opts) {
  const { C, FONT_H, FONT_B, addTextOnShape } = T;
  const s = pres.addSlide();
  s.background = { color: C.BG_LIGHT };
  T.addTopBar(s, C.SECONDARY);
  T.addTitle(s, "Teacher Resources");
  s.addText("Nothing to print this session.", {
    x: 0.5, y: 1.4, w: 9.0, h: 0.4,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
  });
  (opts.items || []).forEach((item, i) => {
    addTextOnShape(s, item, {
      x: 0.5, y: 1.95 + i * 0.68, w: 9.0, h: 0.56, rectRadius: 0.08,
      fill: { color: C.BG_CARD }, line: { color: C.SECONDARY, width: 1 },
    }, { fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL });
  });
  T.addFooter(s, opts.footer);
  s.addNotes(opts.notes);
  T.runSlideDiagnostics(s, pres, { ignoreUnderfill: true });
  return s;
}

/** Collects results across a whole deck and prints the SIGN ASSETS report. */
function createSignReport() {
  const seen = new Map();
  return {
    add(results) {
      (Array.isArray(results) ? results : [results]).forEach((r) => {
        if (r && r.gloss && !seen.has(r.gloss)) seen.set(r.gloss, r.ok);
      });
    },
    rehearseLines(glosses) {
      return glosses.map((g) => REHEARSE_FIRST[g]).filter(Boolean);
    },
    print(label) {
      const found = [...seen.entries()].filter(([, ok]) => ok).map(([g]) => g);
      const missing = [...seen.entries()].filter(([, ok]) => !ok).map(([g]) => g);
      console.log(`SIGN ASSETS (${label}): ${found.length} imaged, ${missing.length} lookup card(s)`);
      console.log("  imaged : " + (found.join(", ") || "none"));
      console.log("  lookup : " + (missing.join(", ") || "none"));
      const corrected = found.filter((g) => SENSE_CORRECTIONS[g]);
      if (corrected.length) {
        console.log("  sense-corrected: " + corrected
          .map((g) => `${g} -> ${SENSE_CORRECTIONS[g]}`).join(", "));
      }
      return { found, missing };
    },
  };
}

const ATTRIBUTION =
  "Sign images: Auslan Signbank (auslan.org.au), CC BY-NC-ND 4.0, used for " +
  "internal school teaching under the schools statutory educational licence.";

module.exports = {
  resolveSign, lookupUrl, signCard, signCardRow, cardHeightFor,
  signSlide, signRecallSlide, stepCard, noPrintResourceSlide, createSignReport,
  SENSE_CORRECTIONS, FORCE_LOOKUP, REHEARSE_FIRST, ATTRIBUTION,
};
