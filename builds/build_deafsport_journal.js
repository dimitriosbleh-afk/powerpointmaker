"use strict";

/**
 * Deaf Sports in Australia | Student Journal Template, one per cohort
 *
 * AUSLAN_2_SLIDES_PROMPT section 8 and unit document 11.8 and 13.5. Uploaded
 * once to Google Slides, then copied for every student through a Google
 * Classroom assignment. Not printed, never handed in.
 *
 * It carries a cover, the term vocabulary tracker (English words and Sign It!
 * pages only: the journal leaves school control, so no sign images), one page
 * per lesson, and a closing page. Each lesson page reads its intention,
 * criteria, journal line and cultural question from that lesson's build script,
 * so the journal cannot drift from the deck.
 *
 * Tracker and answer areas are real tables and text boxes so students can type
 * in them in Google Slides. Not a lesson, so it builds with `node`, not the
 * lesson gates.
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme, weekToVariant } = require("../themes/factory");
const A = require("./auslan_lib");

const RATINGS = ["Just starting", "Getting there", "Got it"];

// Term vocabulary, unit document section 7 with Chris's twelve sports.
// [English, gloss for the Sign It! page, Enrichment only?]
const TRACKER = [
  ["swimming", "SWIM"], ["athletics", "RUN"], ["futsal", "FUTSAL"], ["basketball", "BASKETBALL"],
  ["tennis", "TENNIS"], ["football", "FOOTY"], ["netball", "NETBALL"], ["chess", "CHESS"],
  ["golf", "GOLF"], ["rugby", "RUGBY"], ["table tennis", "TABLE-TENNIS"], ["lawn bowls", "BOWLS"],
  ["deaf", "DEAF"], ["sport", "SPORT"], ["favourite", "FAVOURITE"], ["team", "TEAM"],
  ["again", "AGAIN"], ["thank you", "THANK-YOU"], ["practise", "PRACTISE"], ["who", "WHO"],
  ["what", "WHAT"], ["where", "WHERE"], ["what do you mean", "WHAT MEAN?"], ["when", "WHEN"],
  ["how many", "HOW-MANY"], ["year", "YEAR"], ["numbers", "NUMBERS"], ["years like 1965", "YEARS"],
  ["medal", "MEDAL"], ["win", "WIN"], ["gold", "GOLD", true], ["silver", "SILVER", true],
  ["bronze", "BRONZE", true], ["slow down", "SLOW"], ["before", "BEFORE"], ["after", "AFTER"],
  ["back then", "PAST"], ["then", "THEN"], ["history", "HISTORY"], ["start", "START"],
  ["light", "LIGHT"], ["flag", "FLAG"], ["change", "CHANGE"], ["community", "COMMUNITY"],
  ["compete", "COMPETE"], ["proud", "PROUD"], ["happen", "HAPPEN"], ["first", "FIRST"],
  ["finish", "FINISH"], ["story", "STORY"], ["tell", "TELL"], ["soon", "SOON"],
  ["later", "LATER"], ["next", "NEXT"],
];

/** Pull a lesson's journal content out of its build script. */
function readLesson(cohort, n) {
  const src = fs.readFileSync(path.join(__dirname, `build_deafsport_${cohort}_s${n}.js`), "utf8");
  const title = src.match(/T\.titleSlide\(\s*pres,\s*"([^"]+)"/)[1];
  const li = src.match(/T\.liSlide\(\s*pres,\s*"([^"]+)",\s*\[\s*"([^"]+)",\s*"([^"]+)",\s*"([^"]+)"/);
  const ex = src.match(/T\.exitTicketSlide\(\s*pres,\s*\[\s*"[^"]+",\s*"([^"]+)",\s*"([^"]+)"/);
  let reflection = ex[1];
  // Lesson 2's exit line points at the journal itself; the page asks for the content.
  if (n === 2) reflection = "Paste the photo of your board here, then write the name and sport you found out.";
  return {
    n, title, li: li[1], sc: [li[2], li[3], li[4]],
    reflection, cultural: ex[2].replace(/^Cultural question:\s*/, ""),
  };
}

function build(cohort) {
  const enrichment = cohort === "enrichment";
  const T = createTheme("literacy", enrichment ? "grade56" : "grade34", weekToVariant(1));
  const C = T.C;
  const cohortName = enrichment ? "Enrichment, Years 5 and 6" : "Challenge, Years 3 and 4";
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  const typeBox = (s, x, y, w, h) => {
    s.addText("Type here", {
      x, y, w, h, fontSize: 13, fontFace: T.FONT_B, color: C.MUTED, italic: true,
      align: "left", valign: "top", margin: 0.08,
      line: { color: C.PRIMARY_LINE || C.MUTED, width: 0.75 }, fill: { color: C.WHITE },
    });
  };

  // 1. Cover
  {
    const s = pres.addSlide();
    s.background = { color: C.BG_DARK };
    s.addText("Deaf Sports in Australia", {
      x: 0.6, y: 1.0, w: 8.8, h: 0.9, fontSize: 38, fontFace: T.FONT_H, bold: true,
      color: C.WHITE, align: "left", valign: "middle", margin: 0,
    });
    s.addText("My Auslan journal  |  " + cohortName + "  |  Term 4 2026", {
      x: 0.6, y: 1.9, w: 8.8, h: 0.5, fontSize: 18, fontFace: T.FONT_B,
      color: C.WHITE, align: "left", valign: "middle", margin: 0,
    });
    s.addText("Name:", {
      x: 0.6, y: 2.8, w: 1.2, h: 0.55, fontSize: 22, fontFace: T.FONT_B, bold: true,
      color: C.WHITE, align: "left", valign: "middle", margin: 0,
    });
    s.addText("Type your name", {
      x: 1.8, y: 2.8, w: 5.2, h: 0.55, fontSize: 20, fontFace: T.FONT_B, italic: true,
      color: C.MUTED, fill: { color: C.WHITE }, align: "left", valign: "middle", margin: 0.1,
    });
    s.addText("This is where your family sees what you can do in Auslan.", {
      x: 0.6, y: 3.7, w: 8.8, h: 0.5, fontSize: 17, fontFace: T.FONT_B,
      color: C.WHITE, align: "left", valign: "middle", margin: 0,
    });
    s.addText("To rate yourself, change the colour of one box: just starting, getting there or got it.", {
      x: 0.6, y: 4.3, w: 8.8, h: 0.5, fontSize: 13, fontFace: T.FONT_B,
      color: C.WHITE, align: "left", valign: "middle", margin: 0,
    });
  }

  // 2. Tracker, split across pages
  const words = TRACKER.filter(([, , eOnly]) => enrichment || !eOnly);
  const perPage = 14;
  const pages = Math.ceil(words.length / perPage);
  for (let p = 0; p < pages; p++) {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "Tracker", { color: C.PRIMARY });
    T.addTitle(s, `Signs this term (${p + 1} of ${pages})`);
    const head = ["Sign", "Sign It! page", "Start of term", "Middle", "End"].map((t) => ({
      text: t, options: { bold: true, color: "FFFFFF", fill: { color: C.PRIMARY }, align: "center" },
    }));
    const rows = [head];
    words.slice(p * perPage, (p + 1) * perPage).forEach(([english, gloss]) => {
      const page = A.SIGNIT_PAGES[gloss];
      rows.push([
        { text: english, options: { bold: true, align: "left" } },
        { text: page ? String(page) : "", options: { align: "center" } },
        { text: "", options: {} }, { text: "", options: {} }, { text: "", options: {} },
      ]);
    });
    s.addTable(rows, {
      x: 0.5, y: 1.3, w: 9, colW: [3.0, 1.5, 1.5, 1.5, 1.5], rowH: 0.24,
      fontSize: 11, fontFace: T.FONT_B, color: C.CHARCOAL,
      border: { type: "solid", color: C.PRIMARY_LINE || "BFBFBF", pt: 0.75 }, margin: 0.04,
    });
  }

  // 3. One page per lesson
  for (let n = 1; n <= 8; n++) {
    const L = readLesson(cohort, n);
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "Lesson " + n, { color: C.PRIMARY });
    T.addTitle(s, L.title);
    s.addText(L.li, {
      x: 0.5, y: 1.25, w: 5.35, h: 0.5, fontSize: 12, fontFace: T.FONT_B, italic: true,
      color: C.CHARCOAL, align: "left", valign: "middle", margin: 0, fit: "shrink",
    });
    L.sc.forEach((sc, i) => {
      const y = 1.8 + i * 0.58;
      s.addText(sc, {
        x: 0.5, y, w: 5.35, h: 0.28, fontSize: 11.5, fontFace: T.FONT_B, bold: true,
        color: C.CHARCOAL, align: "left", valign: "middle", margin: 0, fit: "shrink",
      });
      RATINGS.forEach((r, k) => {
        T.addTextOnShape(
          s,
          r,
          { x: 0.5 + k * 1.3, y: y + 0.3, w: 1.2, h: 0.23, rectRadius: 0.06,
            fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 0.75 } },
          { fontSize: 9.5, fontFace: T.FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0 }
        );
      });
    });
    s.addShape("roundRect", {
      x: 6.1, y: 1.3, w: 3.4, h: 2.1, rectRadius: 0.1,
      fill: { color: C.PRIMARY_SOFT || C.BG_LIGHT }, line: { color: C.PRIMARY, width: 1.25, dashType: "dash" },
    });
    s.addText("Insert your video here", {
      x: 6.1, y: 1.3, w: 3.4, h: 2.1, fontSize: 15, fontFace: T.FONT_B, bold: true,
      color: C.PRIMARY, align: "center", valign: "middle", margin: 0,
    });
    [[L.reflection, 0.5], ["Cultural question: " + L.cultural, 5.1]].forEach(([prompt, x]) => {
      s.addText(prompt, {
        x, y: 3.55, w: 4.4, h: 0.62, fontSize: 11, fontFace: T.FONT_B, bold: true,
        color: C.CHARCOAL, align: "left", valign: "bottom", margin: 0, fit: "shrink",
      });
      typeBox(s, x, 4.22, 4.4, 0.85);
    });
    T.addFooter(s, "Deaf Sports in Australia  |  " + cohortName + "  |  Lesson " + n);
  }

  // 4. Closing page
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "End of term", { color: C.PRIMARY });
    T.addTitle(s, "Three things I can sign now that I could not sign in week 1");
    [0, 1, 2].forEach((i) => {
      const y = 1.5 + i * 1.15;
      s.addText(String(i + 1), {
        x: 0.5, y, w: 0.6, h: 0.9, fontSize: 30, fontFace: T.FONT_H, bold: true,
        color: C.PRIMARY, align: "center", valign: "middle", margin: 0,
      });
      typeBox(s, 1.25, y, 8.25, 0.9);
    });
  }

  return pres;
}

const OUT = path.join("output");
(async () => {
  for (const [cohort, folder, name] of [
    ["enrichment", "Deaf_Sports_Enrichment_Unit", "Deaf Sports Enrichment Student Journal Template.pptx"],
    ["challenge", "Deaf_Sports_Challenge_Unit", "Deaf Sports Challenge Student Journal Template.pptx"],
  ]) {
    const dir = path.join(OUT, folder);
    fs.mkdirSync(dir, { recursive: true });
    const file = path.join(dir, name);
    await build(cohort).writeFile({ fileName: file });
    console.log("PPTX written to " + file);
  }
})();
