"use strict";

/**
 * Shared sign layer for Auslan decks (AUSLAN_2_SLIDES_PROMPT section 3 and 6).
 *
 * Auslan lessons are built as JavaScript rather than as lesson specs because a
 * spec cannot express what these decks are made of: an animated sign placed on
 * a card, a lookup card when the bank has no usable image, the fixed six-cue
 * strip, and a countdown GIF on a timed slide. The pipeline knows none of those.
 * This module is the shared layer the lesson scripts call, so the next unit's
 * script adds vocabulary, not drawing code.
 *
 * NEVER draw, generate, mirror or approximate a sign. Assets come only from
 * assets/auslan_signs/, built by scripts/fetch_auslan_signs.py from Auslan
 * Signbank. A gloss with no usable asset gets a lookup card, never a drawing.
 *
 * THE SENSE MAP IS THE IMPORTANT PART. Signbank's first entry for a search word
 * is often not the sense a lesson teaches, and a plausible image of the wrong
 * sign is invisible on a rendered slide. Every gloss below was checked against
 * the recorded definition in assets/auslan_signs/manifest.json on 2026-09-17.
 * When the teacher supplies his vetted entry links, refetch with --links and
 * most of this map stops being needed.
 */

const fs = require("fs");
const path = require("path");

const BANK = path.join(__dirname, "..", "assets", "auslan_signs");
const SIGNIT = path.join(BANK, "signit");
const TIMERS = path.join(__dirname, "..", "assets", "timers");

const ATTRIBUTION =
  "Sign images: Auslan Signbank (auslan.org.au), CC BY-NC-ND 4.0, used for " +
  "internal school teaching under the schools statutory educational licence.";

/** Glosses where the default asset is the WRONG SENSE. Verified corrections. */
const SENSE_CORRECTIONS = {
  PRACTISE: "PRACTISE_3",  // PRACTISE is LEARNER/TRAINEE; _3 is regular training
  WIN: "WIN_3",            // WIN is CELEBRATION, not the sporting result
  NEXT: "NEXT_2",          // NEXT is DEMOTION, not the next event
  CORRECT: "CORRECT_2",    // CORRECT is the written tick, not "that is right"
};

/**
 * Glosses given a lookup card even when a file exists, because nothing in
 * Signbank carries what this unit needs. Value is the English word the lookup
 * card searches for. Keep this list short: a lookup card is the last resort,
 * not the first answer to a wrong-looking sign.
 */
const FORCE_LOOKUP = {
  "WHAT MEAN?": "what do you mean",  // a fixed two-sign form the school supplies, not one entry
  MEAN: "meaning",   // MEAN's first sense is SIGN LANGUAGE; do not assert it
  BRONZE: "bronze",  // no Signbank entry at all
};

/**
 * Glosses where the lesson needs a particular SENSE, and the resolver has to go
 * and find the variant that carries it rather than trusting entry order.
 *
 * Signbank records a part-of-speech breakdown per entry, and the manifest keeps
 * it. A question sign is the obvious case: the first entry for "who" is the noun
 * sense (someone) and only the second carries "As Question". Reading the first
 * clause of a definition and giving up produces a lookup card for a sign that is
 * sitting right there, which is not good enough.
 */
const SENSE_NEEDED = {
  WHO: "question", WHAT: "question", WHERE: "question", WHEN: "question",
  WHY: "question", WHICH: "question", HOW: "question",
};

/**
 * Chris's Sign It! page numbers, from the sheet he returned on 3 Oct 2026.
 * A gloss missing here is not in the book. Printed under every sign card.
 */
const SIGNIT_PAGES = {
  SWIM: 57, BASKETBALL: 58, TENNIS: 60, FOOTY: 58, NETBALL: 59, RUGBY: 60,
  SPORT: 80, DEAF: 64, AGAIN: 50, PRACTISE: 83, WHO: 47, WHAT: 47, WHERE: 47,
  WIN: 108, WHEN: 47, "HOW-MANY": 47, YEAR: 65, BEFORE: 68, PAST: 69, THEN: 68,
  START: 84, CHANGE: 105, COMMUNITY: 128,
};

/**
 * Signs Chris vetted from outside Signbank (3 Oct 2026). The fetcher only
 * reads Signbank, so these stay lookup cards, but the card opens his choice
 * rather than a Signbank search.
 */
const TEACHER_LINKS = {
  FUTSAL: "https://find.auslan.fyi/sign/spread-the-sign-auslan/4306",
  AFTER: "https://find.auslan.fyi/sign/spread-the-sign-auslan/7964",
  FINISH: "https://find.auslan.fyi/sign/latrobe-ig/CCHi2nDHEJD",
  GOLD: "https://find.auslan.fyi/sign/spread-the-sign-auslan/6852",
  SILVER: "https://find.auslan.fyi/sign/spread-the-sign-auslan/599",
  BRONZE: "https://find.auslan.fyi/sign/spread-the-sign-auslan/6157",
};

/** The teacher's own overrides, which beat everything else in this file. */
const OVERRIDE_FILE = path.join(
  __dirname, "..", "reference", "auslan", "signbank_links", "overrides.json"
);

let OVERRIDES = null;
function overrides() {
  if (OVERRIDES) return OVERRIDES;
  OVERRIDES = {};
  if (fs.existsSync(OVERRIDE_FILE)) {
    try {
      OVERRIDES = JSON.parse(fs.readFileSync(OVERRIDE_FILE, "utf8"));
    } catch (e) {
      console.warn(`WARN [auslan] could not read ${OVERRIDE_FILE}: ${e.message}`);
    }
  }
  return OVERRIDES;
}

let MANIFEST = null;
function manifest() {
  if (MANIFEST) return MANIFEST;
  MANIFEST = {};
  const mf = path.join(BANK, "manifest.json");
  if (fs.existsSync(mf)) {
    try {
      for (const rec of JSON.parse(fs.readFileSync(mf, "utf8")).signs || []) {
        MANIFEST[rec.gloss] = rec;
      }
    } catch (e) {
      /* a missing or broken manifest just means no sense checking */
    }
  }
  return MANIFEST;
}

const SENSE_PATTERNS = {
  question: /as question/i,
  noun: /as a noun/i,
  verb: /as a verb/i,
};

/**
 * The file in the bank whose recorded definition carries the sense this lesson
 * needs. Returns null when the manifest has nothing to go on.
 */
function fileForSense(gloss, sense) {
  const rec = manifest()[gloss];
  const pattern = SENSE_PATTERNS[sense];
  if (!rec || !pattern) return null;
  for (const img of rec.images || []) {
    if (pattern.test(img.definition || "")) return img.file;
  }
  return null;
}

/** Usable, but close enough to another sense to be worth a rehearsal line. */
const REHEARSE_FIRST = {
  FAVOURITE: "FAVOURITE also carries adoration and obsession. Confirm the preference sense.",
  AGAIN: "AGAIN is the repetition sense. AGAIN_2 is 'one more'. Confirm which you want.",
  WHERE: "WHERE is whereabouts or location. Confirm the question form from Sign It!.",
  PROUD: "PROUD also carries boast. Confirm the pride sense.",
  HISTORY: "HISTORY is the long-ago sense, not the school subject.",
  LIGHT: "LIGHT has a lamp sense and a flashing sense. Confirm which the reference gives.",
};

const EXTS = [".gif", ".jpg", ".png"];

/**
 * Resolve a gloss to a bank asset, in this order of authority:
 *   1. the teacher's override file
 *   2. the entry fetched from the teacher's vetted link
 *   3. the variant whose recorded definition carries the sense the lesson needs
 *   4. a verified sense correction
 *   5. the default entry
 * A lookup card is what happens when all five come up empty, not a shortcut.
 */
function resolveSign(gloss) {
  const exact = (name) => {
    if (!name) return null;
    if (/\.(gif|jpe?g|png)$/i.test(name)) {
      const p = path.join(BANK, name);
      return fs.existsSync(p) ? { path: p, file: name, kind: path.extname(name).slice(1) } : null;
    }
    for (const ext of EXTS) {
      const p = path.join(BANK, name + ext);
      if (fs.existsSync(p)) return { path: p, file: name + ext, kind: ext.slice(1) };
    }
    return null;
  };

  const override = overrides()[gloss];
  if (override) {
    const hit = exact(typeof override === "string" ? override : override.file);
    if (hit) return Object.assign(hit, { why: "teacher override" });
    console.warn(`WARN [auslan] override for ${gloss} is not in the bank: ${JSON.stringify(override)}`);
  }

  if (FORCE_LOOKUP[gloss]) return null;

  // Fetched from the teacher's own vetted link (--links): his choice of entry
  // outranks the sense rules below, which exist to guess when he has not chosen.
  const rec = manifest()[gloss];
  const vetted = rec && (rec.images || []).find((img) => img.vetted);
  const vettedHit = vetted && exact(vetted.file);
  if (vettedHit) return Object.assign(vettedHit, { why: "teacher's vetted link" });

  const sense = SENSE_NEEDED[gloss];
  if (sense) {
    const hit = exact(fileForSense(gloss, sense));
    if (hit) return Object.assign(hit, { why: `${sense} sense` });
  }

  const pinned = exact(SENSE_CORRECTIONS[gloss]);
  if (pinned) return Object.assign(pinned, { why: "sense correction" });

  if (sense && manifest()[gloss]) {
    // The lesson needs a sense no entry in the bank carries. Say so rather than
    // placing the default and hoping.
    console.warn(
      `WARN [auslan] ${gloss}: no variant in the bank carries the ${sense} sense. ` +
      `Fetch more variants, or add an entry to reference/auslan/signbank_links/overrides.json.`
    );
    return null;
  }

  const fallback = exact(gloss);
  return fallback ? Object.assign(fallback, { why: "default entry" }) : null;
}

/** The teacher's own Sign It! crop, if he has scanned that page. */
function resolveSignIt(gloss) {
  const p = path.join(SIGNIT, gloss + ".png");
  return fs.existsSync(p) ? p : null;
}

/** Signbank SEARCH url. Never construct an entry url; suffixes are unpredictable. */
function lookupUrl(englishWord) {
  return "https://auslan.org.au/dictionary/search/?query=" +
    encodeURIComponent(String(englishWord).toLowerCase());
}

/**
 * Real pixel size of a GIF or PNG, so an image is never stretched.
 * Reading a few header bytes beats adding a dependency for two formats.
 */
function imageAspect(file) {
  const buf = fs.readFileSync(file);
  if (buf.slice(0, 3).toString("latin1") === "GIF") {
    return buf.readUInt16LE(6) / buf.readUInt16LE(8);
  }
  if (buf.slice(1, 4).toString("latin1") === "PNG") {
    return buf.readUInt32BE(16) / buf.readUInt32BE(20);
  }
  // JPEG: walk the segments to the frame header.
  let i = 2;
  while (i < buf.length - 9) {
    if (buf[i] !== 0xff) { i += 1; continue; }
    const marker = buf[i + 1];
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return buf.readUInt16BE(i + 7) / buf.readUInt16BE(i + 5);
    }
    i += 2 + buf.readUInt16BE(i + 2);
  }
  return 1;
}

/**
 * The SIGN ASSETS report. Every Auslan build prints one: it is how the gap
 * between what the deck shows and what the teacher has vetted gets closed.
 */
function createSignReport() {
  const placed = [];
  const lookups = [];
  const rehearse = new Set();
  return {
    note(gloss, ok, detail) {
      (ok ? placed : lookups).push(detail ? `${gloss} (${detail})` : gloss);
      if (REHEARSE_FIRST[gloss]) rehearse.add(REHEARSE_FIRST[gloss]);
    },
    rehearsalLines() {
      return Array.from(rehearse);
    },
    print() {
      console.log("\nSIGN ASSETS");
      console.log(`  placed as images (${placed.length}): ${placed.join(", ") || "none"}`);
      console.log(`  lookup cards (${lookups.length}): ${lookups.join(", ") || "none"}`);
      if (lookups.length) {
        console.log("  Lookup cards teach fine, because the teacher models every sign live.");
        console.log("  They close when he supplies vetted Signbank links:");
        console.log("    python scripts/fetch_auslan_signs.py --gif --links <his links file>");
      }
      const unvetted = readManifestUnvetted(placed);
      if (unvetted.length) {
        console.log(`  NOT YET VETTED (${unvetted.length}): ${unvetted.join(", ")}`);
        console.log("  Those came from a Signbank search, not from an entry he chose.");
      }
    },
  };
}

function readManifestUnvetted(placed) {
  const mf = path.join(BANK, "manifest.json");
  if (!fs.existsSync(mf)) return [];
  let signs = [];
  try {
    signs = JSON.parse(fs.readFileSync(mf, "utf8")).signs || [];
  } catch (e) {
    return [];
  }
  const byGloss = new Map(signs.map((s) => [s.gloss, s]));
  return placed
    .map((p) => p.split(" ")[0])
    .filter((g) => {
      const rec = byGloss.get(g);
      return rec && !(rec.images || []).some((i) => i.vetted);
    });
}

/**
 * One sign on a card: the English meaning is what students read, the gloss is a
 * small caption, and the asset is placed to its own aspect so nothing stretches.
 * With no usable asset it becomes a lookup card, which is a real teaching object
 * and not a hole: the meaning, a watch-the-teacher line and a Signbank search.
 */
function addSignCard(T, slide, spec, box, report) {
  const { gloss, meaning } = spec;
  const page = spec.page || SIGNIT_PAGES[gloss];
  const { x, y, w, h } = box;
  const C = T.C;
  const pad = 0.16;
  const capH = 0.28;
  const labelH = 0.46;

  T.addCard(slide, x, y, w, h, { variant: "white" });

  const signIt = resolveSignIt(gloss);
  const found = resolveSign(gloss);
  const asset = signIt || (found || {}).path;
  const imgTop = y + pad;
  const imgH = h - pad * 2 - labelH - capH;

  if (asset) {
    const aspect = imageAspect(asset);
    let iw = imgH * aspect;
    let ih = imgH;
    if (iw > w - pad * 2) {
      iw = w - pad * 2;
      ih = iw / aspect;
    }
    slide.addImage({ path: asset, x: x + (w - iw) / 2, y: imgTop + (imgH - ih) / 2, w: iw, h: ih });
    if (report) {
      report.note(gloss, true,
        signIt ? "Sign It! scan" : `${path.basename(asset).replace(/\.[a-z]+$/, "")}, ${found.why}`);
    }
  } else {
    const word = FORCE_LOOKUP[gloss] || String(meaning || gloss).toLowerCase();
    T.addCard(slide, x + pad, imgTop, w - pad * 2, imgH, { variant: "tint" });
    slide.addText("Watch the teacher", {
      x: x + pad, y: imgTop + imgH / 2 - 0.5, w: w - pad * 2, h: 0.36,
      fontSize: 13, fontFace: T.FONT_B, color: C.MUTED,
      align: "center", valign: "middle", margin: 0,
    });
    slide.addText(
      [TEACHER_LINKS[gloss]
        ? { text: "Watch the sign online", options: { hyperlink: { url: TEACHER_LINKS[gloss] }, color: C.PRIMARY } }
        : { text: "Look it up: Auslan Signbank", options: { hyperlink: { url: lookupUrl(word) }, color: C.PRIMARY } }],
      {
        x: x + pad, y: imgTop + imgH / 2 - 0.06, w: w - pad * 2, h: 0.36,
        fontSize: 12, fontFace: T.FONT_B, bold: true,
        align: "center", valign: "middle", margin: 0,
      }
    );
    if (report) report.note(gloss, false);
  }

  slide.addText(String(meaning), {
    x: x + 0.06, y: y + h - pad - labelH - capH, w: w - 0.12, h: labelH,
    fontSize: 21, fontFace: T.FONT_H, bold: true, color: C.CHARCOAL,
    align: "center", valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
  });
  slide.addText(page ? `${gloss}  |  Sign It! p. ${page}` : gloss, {
    x: x + 0.06, y: y + h - pad - capH, w: w - 0.12, h: capH,
    fontSize: 11, fontFace: T.FONT_B, color: C.MUTED,
    align: "center", valign: "middle", margin: 0,
  });
}

/**
 * The Sign It! front cover, wherever the textbook is referenced or students are
 * told to look something up, so the thing on the screen is the thing on the
 * desk. Chris supplies the cover scan; until it lands this draws a labelled
 * placeholder rather than nothing, so the slot is visible in review.
 */
function addSignItCover(T, slide, box) {
  const { x, y, w, h } = box;
  const C = T.C;
  const cover = path.join(SIGNIT, "COVER.png");
  T.addCard(slide, x, y, w, h, { variant: "outline", tone: C.PRIMARY });
  if (fs.existsSync(cover)) {
    const aspect = imageAspect(cover);
    const pad = 0.16;
    let iw = (h - pad * 2 - 0.34) * aspect;
    let ih = h - pad * 2 - 0.34;
    if (iw > w - pad * 2) {
      iw = w - pad * 2;
      ih = iw / aspect;
    }
    slide.addImage({ path: cover, x: x + (w - iw) / 2, y: y + pad, w: iw, h: ih });
  } else {
    slide.addText("Sign It!", {
      x: x + 0.1, y: y + h / 2 - 0.5, w: w - 0.2, h: 0.6,
      fontSize: 26, fontFace: T.FONT_H, bold: true, color: C.PRIMARY,
      align: "center", valign: "middle", margin: 0,
    });
    slide.addText("cover image to come", {
      x: x + 0.1, y: y + h / 2 + 0.08, w: w - 0.2, h: 0.4,
      fontSize: 12, fontFace: T.FONT_B, color: C.MUTED,
      align: "center", valign: "middle", margin: 0,
    });
  }
  slide.addText("Open it at the back", {
    x: x + 0.1, y: y + h - 0.42, w: w - 0.2, h: 0.32,
    fontSize: 13, fontFace: T.FONT_B, bold: true, color: C.CHARCOAL,
    align: "center", valign: "middle", margin: 0,
  });
  return { x, y, w, h, present: fs.existsSync(cover) };
}

/** A row of sign cards filling the content area, one to four across. */
function addSignCardRow(T, slide, specs, opts) {
  const o = opts || {};
  const n = Math.min(specs.length, 4);
  const top = o.y != null ? o.y : 1.35;
  const bottom = o.bottom != null ? o.bottom : 5.05;
  const gap = n > 3 ? 0.16 : 0.22;
  const w = (9 - gap * (n - 1)) / n;
  const h = bottom - top;
  specs.slice(0, n).forEach((spec, i) => {
    addSignCard(T, slide, spec, { x: 0.5 + i * (w + gap), y: top, w, h }, o.report);
  });
  return { top, bottom, w, h };
}

/**
 * Place the countdown for a timed task. Generates the GIF on first use and
 * caches it, so calling this for every timed slide costs nothing after the
 * first. It plays once, when the slide is advanced to.
 */
function countdownFile(T, seconds) {
  const { execFileSync } = require("child_process");
  const colour = String(T.C.PRIMARY).replace("#", "").toUpperCase();
  const file = path.join(TIMERS, `countdown_${seconds}s_${colour}.gif`);
  if (!fs.existsSync(file)) {
    execFileSync("python", [
      path.join(__dirname, "..", "scripts", "make_countdown_gif.py"),
      String(seconds), "--color", colour,
    ], { stdio: "pipe" });
  }
  return file;
}

function addCountdown(T, slide, seconds, box) {
  const file = countdownFile(T, seconds);
  const d = box.w || 1.05;
  slide.addImage({ path: file, x: box.x, y: box.y, w: d, h: d });
  return { x: box.x, y: box.y, w: d, h: d };
}

const TEAM_ROLE_CARDS = [
  ["Sam", "Swimming", "Perth"], ["Ali", "Swimming", "Perth"],
  ["Jo", "Swimming", "Perth"], ["Kit", "Swimming", "Perth"],
  ["Ren", "Swimming", "Hobart"], ["Bo", "Swimming", "Hobart"],
  ["Tam", "Swimming", "Hobart"], ["Nia", "Swimming", "Hobart"],
  ["Max", "Netball", "Adelaide"], ["Eve", "Netball", "Adelaide"],
  ["Ari", "Netball", "Adelaide"], ["Lou", "Netball", "Adelaide"],
  ["Fin", "Basketball", "Sydney"], ["Zia", "Basketball", "Sydney"],
  ["Rue", "Basketball", "Sydney"], ["Dev", "Basketball", "Sydney"],
  ["Ivy", "Tennis", "Melbourne"], ["Cam", "Tennis", "Melbourne"],
  ["Rio", "Tennis", "Melbourne"], ["Gus", "Tennis", "Melbourne"],
  ["Wren", "Athletics", "Brisbane"], ["Ash", "Athletics", "Brisbane"],
  ["Nell", "Athletics", "Brisbane"], ["Ty", "Athletics", "Brisbane"],
  ["Sol", "Futsal", "Darwin"], ["Pip", "Futsal", "Darwin"],
  ["Quin", "Futsal", "Darwin"], ["Jed", "Futsal", "Darwin"],
];

/**
 * Session 3 Team Role Cards: 28 cards, seven teams of four, cut and laminate.
 * Shared because Challenge and Enrichment play Game 4 with the same set.
 */
function buildTeamRoleCardsPdf(file, footer, color) {
  const P = require("../themes/pdf_helpers");
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const doc = P.createPdf({ title: "Session 3 Team Role Cards" });
  let y = P.addPdfHeader(doc, "Session 3 Team Role Cards", {
    color,
    subtitle: "Seven teams of four. Cut along the lines and laminate. One set does every class.",
    lessonInfo: "Session 3  |  Game 4 Find Your Team  |  one card per student",
    showNameDate: false,
  });

  doc.fontSize(10).font("Sans").fillColor("#374151");
  doc.text(
    "Teams 1 and 2 both play swimming on purpose. A student who stops at the sport question "
    + "joins the wrong team, which is the whole point of the game. Deal so every team of four "
    + "is complete; for a class of 25, hold team 7 back.",
    P.PAGE.MARGIN, y, { width: P.PAGE.CONTENT_W }
  );
  y += 52;  // the intro note runs to three lines; 34 clipped it under the first row

  // Six rows a page, not seven: seven overflows A4 and PDFKit silently starts a
  // new page mid-card, which is how 28 cards became eight pages.
  const cols = 3;
  const rowsPerPage = 6;
  const gap = 10;
  const cw = (P.PAGE.CONTENT_W - gap * (cols - 1)) / cols;
  const ch = 88;
  const perPage = cols * rowsPerPage;
  TEAM_ROLE_CARDS.forEach(([name, sport, city], i) => {
    const col = i % cols;
    const row = Math.floor((i % perPage) / cols);
    if (i > 0 && i % perPage === 0) {
      doc.addPage();
      y = P.PAGE.MARGIN;
    }
    const x = P.PAGE.MARGIN + col * (cw + gap);
    const cy = y + row * (ch + gap);
    doc.save();
    doc.dash(3, { space: 3 }).roundedRect(x, cy, cw, ch, 6)
      .lineWidth(1).strokeColor("#9CA3AF").stroke();
    doc.undash();
    doc.restore();
    doc.fontSize(20).font("Sans-Bold").fillColor("#111827")
      .text(name, x + 12, cy + 10, { width: cw - 24 });
    doc.fontSize(12).font("Sans").fillColor("#374151")
      .text("Sport: " + sport, x + 12, cy + 38, { width: cw - 24 });
    doc.text("City: " + city, x + 12, cy + 56, { width: cw - 24 });
  });

  P.writePdf(doc, file, footer);
}

/**
 * Deaf Sport Timeline Cards (unit document 11.2). Cards 1 to 8 are the Lesson 5
 * set; 9 and 10 join them in the review week. Every date is from Deaf Sports
 * Australia or Deaf Connect, checked 17 Sept 2026. The 1880s club appears under
 * its current name only: the original name is not language we use now.
 */
const TIMELINE_CARDS = [
  ["1880s", "A Deaf cricket club starts in Melbourne. It is still going today as the Melbourne Deaf Cricket Club, and it is one of the oldest Deaf sport clubs in the world."],
  ["1954", "Deaf Sports Australia is set up as the national body for Deaf sport in Australia."],
  ["1964", "The first Australian Deaf Games are held in Sydney, over the summer of 1964 and 1965, with fifteen sports."],
  ["1965", "Australia sends two athletes to the world games for the Deaf. Both of them win a medal."],
  ["1985", "The national Deaf sports body is formally recognised and funded."],
  ["2005", "Melbourne hosts the Deaflympics, with over 3,500 people taking part. It is the only time Australia has hosted them."],
  ["2011", "The Active Deaf Kids school program starts."],
  ["2026", "The Australian Deaf Games are held on the Sunshine Coast from 4 to 11 July, with more than 1,300 competitors and up to twenty sports."],
  ["1924", "The first Deaflympic Games are held in Paris. They have been held every four years since."],
  ["1955", "Deaf Sports Australia joins the international Deaf sports committee."],
];

/** One set of ten cards on one page. Print seven sets, cut and laminate. */
function buildTimelineCardsPdf(file, footer, color) {
  const P = require("../themes/pdf_helpers");
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const doc = P.createPdf({ title: "Session 5 Deaf Sport Timeline Cards" });
  let y = P.addPdfHeader(doc, "Session 5 Deaf Sport Timeline Cards", {
    color,
    subtitle: "Print seven copies: one set per group of four. Cut and laminate. Keep for the review week.",
    lessonInfo: "Session 5  |  Game 6 Order The Years  |  cards 9 and 10 (marked R) join in the review week",
    showNameDate: false,
  });
  y += 8;
  const cols = 2;
  const gap = 10;
  const cw = (P.PAGE.CONTENT_W - gap) / cols;
  const ch = 120;
  // Printed out of order (MEGA_PROMPT 19a): an uncut sheet must not give the answer.
  const PRINT_ORDER = [5, 0, 3, 7, 1, 6, 4, 2, 8, 9];
  PRINT_ORDER.map((k) => TIMELINE_CARDS[k]).forEach(([year, text], i) => {
    const x = P.PAGE.MARGIN + (i % cols) * (cw + gap);
    const cy = y + Math.floor(i / cols) * (ch + gap);
    doc.save();
    doc.dash(3, { space: 3 }).roundedRect(x, cy, cw, ch, 6).lineWidth(1).strokeColor("#9CA3AF").stroke();
    doc.undash();
    doc.restore();
    doc.fontSize(24).font("Sans-Bold").fillColor("#111827").text(year, x + 12, cy + 10, { width: cw - 60 });
    if (i >= 8) {
      doc.fontSize(10).font("Sans-Bold").fillColor("#6B7280").text("R", x + cw - 24, cy + 12, { width: 14 });
    }
    doc.fontSize(11.5).font("Sans").fillColor("#1F2937").text(text, x + 12, cy + 44, { width: cw - 24 });
  });
  P.writePdf(doc, file, footer);
}

/**
 * Medal Tally Sheets A and B (unit document 11.3), for Game 7 How Many Medals.
 * Cricket is swapped for netball, from Chris's twelve sports. The last row is
 * meant to be ridiculous. One per student, half A and half B; consumable.
 */
const MEDAL_TALLY = {
  A: [["Swimming", 7], ["Athletics", 12], ["Netball", 3], ["Basketball", 9], ["Tennis", 5], ["Sock wrestling", 41]],
  B: [["Swimming", 11], ["Athletics", 4], ["Netball", 8], ["Basketball", 6], ["Tennis", 14], ["Sock wrestling", 38]],
};

function buildMedalTallyPdf(file, footer, color) {
  const P = require("../themes/pdf_helpers");
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const doc = P.createPdf({ title: "Session 6 Medal Tally Sheets A and B" });
  ["A", "B"].forEach((sheet, k) => {
    if (k > 0) doc.addPage();
    let y = P.addPdfHeader(doc, "Session 6 Medal Tally Sheet " + sheet, {
      color,
      subtitle: "Your medals are filled in. Write your partner's numbers as they sign them to you.",
      lessonInfo: "Session 6  |  Game 7 How Many Medals  |  print half the class A and half B",
      showNameDate: true,
    });
    y += 14;
    const x0 = P.PAGE.MARGIN;
    const widths = [P.PAGE.CONTENT_W * 0.44, P.PAGE.CONTENT_W * 0.26, P.PAGE.CONTENT_W * 0.30];
    const rowH = 46;
    const rows = [["Sport", "My medals", "My partner's medals"]]
      .concat(MEDAL_TALLY[sheet].map(([sport, n]) => [sport, String(n), ""]))
      .concat([["Total", "", ""]]);
    rows.forEach((row, r) => {
      let x = x0;
      row.forEach((cell, c) => {
        if (r === 0) doc.rect(x, y, widths[c], rowH).fill("#E5E7EB");
        doc.rect(x, y, widths[c], rowH).lineWidth(1).strokeColor("#374151").stroke();
        doc.fontSize(r === 0 ? 13 : 16).font(r === 0 || c === 0 ? "Sans-Bold" : "Sans").fillColor("#111827")
          .text(cell, x + 10, y + (rowH - (r === 0 ? 13 : 16)) / 2 - 2, { width: widths[c] - 20, align: c === 0 ? "left" : "center" });
        x += widths[c];
      });
      y += rowH;
    });
    y += 26;
    doc.fontSize(14).font("Sans-Bold").fillColor("#111827")
      .text("Who has more medals altogether?", x0, y);
    P.addWriteLine(doc, "", y + 26, {});
    y += 70;
    doc.fontSize(14).font("Sans-Bold").fillColor("#111827")
      .text("Who has more sock wrestling medals?", x0, y);
    P.addWriteLine(doc, "", y + 26, {});
    y += 70;
    doc.fontSize(11).font("Sans").fillColor("#4B5563")
      .text("Pencil down while your partner signs. Watch the whole answer, then write.", x0, y, { width: P.PAGE.CONTENT_W });
  });
  P.writePdf(doc, file, footer);
}

module.exports = {
  buildMedalTallyPdf,
  TIMELINE_CARDS,
  buildTimelineCardsPdf,
  SIGNIT_PAGES,
  TEACHER_LINKS,
  buildTeamRoleCardsPdf,
  ATTRIBUTION,
  BANK,
  SENSE_CORRECTIONS,
  SENSE_NEEDED,
  FORCE_LOOKUP,
  fileForSense,
  REHEARSE_FIRST,
  resolveSign,
  resolveSignIt,
  lookupUrl,
  imageAspect,
  createSignReport,
  addSignCard,
  addSignCardRow,
  addSignItCover,
  addCountdown,
  countdownFile,
};
