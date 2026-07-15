"use strict";

// Big Ideas (Year 5/6) - Session 1 of 2: The Solstice.
// Student-interest enrichment, not a curriculum unit. Brand new content.
//
// Learning target: Melbourne's longest day is in December BECAUSE the southern
// half of Earth leans toward the sun. The axis tilt never changes; only our
// place in the orbit changes. Leaning in -> sun climbs higher (75 deg vs 29 deg)
// and stays up longer (14h 50m vs 9h 35m).
//
// Lesson shape: OBSERVE FIRST. Students watch the side-by-side solstice
// time-lapse and notice the sun's path differs BEFORE any modelling, so the
// I Do answers a question they already have.
//
// Killer misconception: "summer = closer to the sun". For Australians it is
// accidentally confirmed (Earth is nearest the sun in early January). The
// hinge kills it with a fact that needs no numbers: when Melbourne has summer,
// London has winter. Same planet, same day, same distance.
//
// Sessions 1 and 2 are independent topics. Session 2 does not depend on this one.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine, addLinedArea,
} = require("../themes/pdf_helpers");

const T = createTheme("science", "grade56", 2); // variant 2, shared across both sessions
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  keyWordSlide, exitTicketSlide, cycleDiagramSlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  addInstructionCard, getContrastColor,
  withReveal, runSlideDiagnostics,
  composeGlanceNotes,
} = T;

const SESSION = 1;
const TOTAL = 2;
const FOOTER = `Big Ideas | Session ${SESSION} of ${TOTAL} | Year 5/6 Science`;
const OUT_DIR = "output/BigIdeas_Session1_Solstice";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const VIDEO_URL = "https://www.youtube.com/watch?v=UiAUG1HtWIM";
const VIDEO_TITLE = "Summer Solstice vs. Winter Solstice: Side-by-Side Time-lapse";

// Verified figures for Melbourne (latitude about 37.8 S). Noon sun altitude is
// 90 - 37.8 -/+ 23.5, which matches published values of about 75 and 29 degrees.
const DEC_ALT = 75;
const JUN_ALT = 29;
const DEC_DAYLIGHT = "14 h 50 m";
const JUN_DAYLIGHT = "9 h 35 m";

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Label the orbit model, then explain Melbourne's longest and shortest days.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers and look-fors for the Session 1 worksheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

/* ------------------------------------------------------------------ */
/*  Local visual helpers                                               */
/*  No shared helper draws an orbit-with-tilt or a sun path, so these  */
/*  are built here and diagnostics-checked on every slide that uses    */
/*  them. Both use roundRect (ovals do not render in LibreOffice).     */
/* ------------------------------------------------------------------ */

const SUN_COLOR = C.ALERT;    // warm red - the light source
const EARTH_COLOR = C.ACCENT; // ocean blue

function drawCircle(slide, cx, cy, r, fill, lineOpts) {
  slide.addShape("roundRect", {
    x: cx - r, y: cy - r, w: r * 2, h: r * 2, rectRadius: r,
    fill: { color: fill },
    line: lineOpts || { color: fill, width: 0.5 },
  });
}

// The sun: a filled circle with eight short rays so it reads instantly as the
// light source rather than "another planet".
function drawSun(slide, cx, cy, r, label) {
  for (let i = 0; i < 8; i++) {
    const a = (Math.PI * 2 * i) / 8;
    const rm = r + 0.13;
    const px = cx + rm * Math.cos(a);
    const py = cy + rm * Math.sin(a);
    slide.addShape("roundRect", {
      x: px - 0.022, y: py - 0.085, w: 0.044, h: 0.17, rectRadius: 0.02,
      fill: { color: SUN_COLOR },
      line: { color: SUN_COLOR, width: 0.5 },
      rotate: (a * 180) / Math.PI + 90,
    });
  }
  addTextOnShape(slide, label || "Sun", {
    x: cx - r, y: cy - r, w: r * 2, h: r * 2, rectRadius: r,
    fill: { color: SUN_COLOR },
  }, {
    fontSize: 11, fontFace: FONT_H, color: getContrastColor(SUN_COLOR), bold: true,
  });
}

// One Earth on the orbit: globe + axis bar tilted 23.5 degrees with its top
// leaning LEFT. The tilt direction is IDENTICAL for every Earth drawn - that
// is the whole concept, so it is a constant here, never a parameter.
const TILT_DEG = 23.5;

function drawEarth(slide, cx, cy, r, opts) {
  const o = opts || {};
  const axisHalf = r + 0.26;
  const rad = (TILT_DEG * Math.PI) / 180;
  // Top of the axis leans left: offset x is negative at the top.
  const topX = cx - axisHalf * Math.sin(rad);
  const topY = cy - axisHalf * Math.cos(rad);
  const botX = cx + axisHalf * Math.sin(rad);
  const botY = cy + axisHalf * Math.cos(rad);

  // Axis drawn as a thin rotated bar (rotated rects render reliably; a
  // zero-width rotated line does not).
  slide.addShape("roundRect", {
    x: cx - 0.028, y: cy - axisHalf, w: 0.056, h: axisHalf * 2, rectRadius: 0.02,
    fill: { color: C.CHARCOAL },
    line: { color: C.CHARCOAL, width: 0.5 },
    rotate: -TILT_DEG,
  });

  drawCircle(slide, cx, cy, r, EARTH_COLOR, { color: C.WHITE, width: 1.2 });

  // Pole markers. The pole leaning toward the sun is marked in the sun's own
  // colour, so "this end is getting the light" is readable at a glance.
  const northLit = o.litPole === "north";
  const poleR = o.poleR != null ? o.poleR : 0.09;
  drawCircle(slide, topX, topY, poleR, northLit ? SUN_COLOR : C.MUTED);
  drawCircle(slide, botX, botY, poleR, northLit ? C.MUTED : SUN_COLOR);

  const labelSize = o.labelSize != null ? o.labelSize : 12;
  slide.addText("N", {
    x: topX - 0.20, y: topY - 0.38, w: 0.40, h: 0.24,
    fontSize: labelSize, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
    align: "center", margin: 0,
  });
  slide.addText("S", {
    x: botX - 0.20, y: botY + 0.14, w: 0.40, h: 0.24,
    fontSize: labelSize, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
    align: "center", margin: 0,
  });
}

// Sun in the middle, Earth at two points of the orbit. Both axes drawn with
// the SAME tilt so students can lay a ruler across them and see they are
// parallel. Left Earth = December (south leans in), right = June (north).
function drawOrbitTilt(slide, opts) {
  const o = opts || {};
  const cx = o.cx, cy = o.cy, rx = o.rx, ry = o.ry;
  const earthR = o.earthR != null ? o.earthR : 0.38;

  const DOTS = 36;
  for (let i = 0; i < DOTS; i++) {
    const a = (Math.PI * 2 * i) / DOTS;
    slide.addShape("roundRect", {
      x: cx + rx * Math.cos(a) - 0.025,
      y: cy + ry * Math.sin(a) - 0.025,
      w: 0.05, h: 0.05, rectRadius: 0.025,
      fill: { color: C.MUTED },
      line: { color: C.MUTED, width: 0.4 },
    });
  }

  drawSun(slide, cx, cy, o.sunR != null ? o.sunR : 0.40, o.sunLabel);

  const decX = cx - rx;
  const junX = cx + rx;
  const poleOpts = { poleR: o.poleR, labelSize: o.labelSize };
  drawEarth(slide, decX, cy, earthR, { litPole: "south", ...poleOpts });
  drawEarth(slide, junX, cy, earthR, { litPole: "north", ...poleOpts });

  if (o.chips !== false) {
    const chipY = cy + ry + 0.16;
    addTextOnShape(slide, "22 Dec: South leans in", {
      x: decX - 1.05, y: chipY, w: 2.10, h: 0.34, rectRadius: 0.06,
      fill: { color: C.PRIMARY },
    }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true });
    addTextOnShape(slide, "21 Jun: North leans in", {
      x: junX - 1.05, y: chipY, w: 2.10, h: 0.34, rectRadius: 0.06,
      fill: { color: C.SECONDARY },
    }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true });
  }
  return { decX, junX, cy };
}

// The sun's path across one day, drawn as a series of sun positions along an
// arc - the same representation the time-lapse video shows. Peak height is
// scaled from the real noon altitude, so the two Melbourne days are honestly
// proportional to each other.
function drawSunPath(slide, x, y, w, h, opts) {
  const o = opts || {};
  const alt = Number(o.altitudeDeg);
  const color = o.color || C.SECONDARY;
  const labelH = 0.30;
  const horizonY = y + h - labelH - 0.06;
  // Top pad leaves room for the altitude label to sit ABOVE the noon sun.
  // Parked beside it, the label collides with the next sun along the arc.
  const topPad = 0.46;
  const usableH = horizonY - y - topPad;
  const peak = usableH * (alt / 90);
  const innerX = x + 0.22;
  const innerW = w - 0.44;

  // Ground
  slide.addShape("roundRect", {
    x, y: horizonY, w, h: 0.045, rectRadius: 0.02,
    fill: { color: C.CHARCOAL }, line: { color: C.CHARCOAL, width: 0.5 },
  });

  const n = 9;
  for (let i = 0; i < n; i++) {
    const t = i / (n - 1);
    const px = innerX + t * innerW;
    const py = horizonY - peak * 4 * t * (1 - t);
    const isNoon = i === (n - 1) / 2;
    const r = isNoon ? 0.13 : 0.07;
    drawCircle(slide, px, py - (isNoon ? 0.02 : 0), r, color);
  }

  // Noon altitude label, centred ABOVE the peak sun and clear of its neighbours.
  const noonX = innerX + innerW / 2;
  const noonY = horizonY - peak;
  slide.addText(`${alt}° high`, {
    x: noonX - 0.55, y: noonY - 0.42, w: 1.10, h: 0.26,
    fontSize: 12.5, fontFace: FONT_H, color: color, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  if (o.label) {
    slide.addText(String(o.label), {
      x, y: horizonY + 0.10, w, h: labelH - 0.04,
      fontSize: 11.5, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }
  return { horizonY, peak, noonX, noonY };
}

/* ------------------------------------------------------------------ */
/*  Teacher notes                                                      */
/* ------------------------------------------------------------------ */

const NOTES_TITLE = "Session 1 of 2. Enrichment on student interest, not a curriculum unit. Open the deck at the Teacher Resources slide to set up the globe and torch.";

const NOTES_RESOURCES = composeGlanceNotes({
  answer: "open - teacher preparation slide, not taught",
  beats: [
    "SHOW while students settle. SAY: Two resources today, and I need a globe and a torch on the front bench.",
    "CHECK the video link opens and the sound works before the lesson starts.",
  ],
  prep: "Materials: a globe (or a ball with a pencil pushed through it), a torch, mini-whiteboards, the worksheet. Darken the room a little for the torch demo.",
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_OVERVIEW = composeGlanceNotes({
  answer: "open - teacher-facing overview, do not teach from this slide",
  beats: [
    "SKIP in front of students. Read before the lesson.",
  ],
  prep: "Two independent enrichment sessions in one deck; Session 2 does not need Session 1. Decision points today: hinge after I Do, boards after We Do, exit ticket. Assumption flagged: no prior astronomy is assumed.",
  tag: "[Overview | Planning | HITS 1, 2]",
});

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - listen for: one sun climbs high and stays up ages, the other stays low and sets fast",
  beats: [
    "SAY: Two videos of the same sky, side by side. One is the longest day of the year. One is the shortest. Watch the sun.",
    "PLAY the clip. TIME: keep it under 3 minutes, then stop and talk.",
    "ASK: What was different about the sun? 30 sec think. Write it... chin it... show me, boards up. EXPECT: one sun went higher and stayed up longer. ACCEPT: one was up longer.",
    "SCAN boards. 80%+ -> go to the LI, we will explain what they saw. Less -> replay 30 sec, point at the midday sun each time, re-ask.",
  ],
  trap: "saying the sun moved to a different place. Fix: name it as our view of the sun from one fixed spot, student re-describes what changed.",
  prep: "Observe-first launch: students collect the phenomenon before any model, so the I Do answers a question they now own. No prior knowledge needed, so every student can enter. Keep to 6 minutes.",
  tag: "[Launch | Attention, focus and regulation | HITS 3]",
  sources: [`${VIDEO_TITLE} (${VIDEO_URL})`],
});

const NOTES_LI = composeGlanceNotes({
  answer: "open - students read the criteria, no response collected",
  beats: [
    "SAY: You just saw it happen. Today we work out why, and the answer is one word: tilt.",
    "POINT to each criterion. SAY: By the end, everyone can point to the half of Earth leaning towards the sun.",
  ],
  prep: "SC1 is a pointing task so every student reaches it. SC2 is the exit ticket. SC3 is the England comparison.",
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_VOCAB_AXIS = composeGlanceNotes({
  answer: "the imaginary line Earth spins around, from the north pole to the south pole",
  beats: [
    "SHOW the globe. POINT along the metal rod. SAY: This rod is the axis. Earth spins around it, once a day. That is what gives us day and night.",
    "SAY: Say it with me. Everyone, together, on three. One, two, three... axis.",
    "SHOW the tilt. SAY: Notice the rod is not straight up. It leans. Hold that thought, it is the whole lesson.",
  ],
  trap: "thinking the axis is a real rod through the Earth. Fix: say imaginary line, student points to where it would come out at each pole.",
  prep: "First of two words. Concrete before abstract: the globe rod carries the meaning, the slide only names it.",
  tag: "[Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_VOCAB_SOLSTICE = composeGlanceNotes({
  answer: "the day the midday sun is at its highest or its lowest for the year",
  beats: [
    "SAY: Solstice comes from two old Latin words. Sol means sun. Sistere means to stand still.",
    "SAY: For a few days the midday sun stops climbing, seems to pause, then turns back the other way. That is the sun standing still.",
    "ASK: Which video was a solstice? 10 sec, turn and tell your partner. EXPECT: both of them, one is each solstice.",
  ],
  trap: "hearing solstice as the whole season. Fix: it is one day, the turning point, student names the two dates.",
  prep: "Second of two words. The word origin makes the meaning stick and links straight back to the launch clip.",
  tag: "[Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_IDO_TILT = composeGlanceNotes({
  answer: "the axis stays pointing the same way all year; only Earth's place in the orbit changes",
  beats: [
    "POINT to both Earths. SAY: Same planet, same tilt, six months apart. Lay a ruler across both axes. They are parallel.",
    "MODEL with the globe. SAY: Watch. I carry the globe around the sun and I never twist the rod. It keeps pointing the same way, at the same far-off star, all year.",
    "POINT to December. SAY: Here the bottom of the rod leans towards the sun, so the southern half, our half, is leaning in.",
    "POINT to June. SAY: Same rod, same lean, but now we are on the other side. Now the top leans in, and the north gets the sun.",
    "ASK: In December, which half of Earth leans towards the sun? 10 sec, everyone points at the screen. EXPECT: pointing at the southern half.",
  ],
  trap: "thinking Earth tips back and forth like a see-saw. Fix: carry the globe around again with the rod locked, student says what changed and what did not.",
  stretch: "work out which far-off star our axis points at, and check it tonight.",
  help: "hold the globe and walk it around the torch yourself while a partner checks the rod never twists.",
  prep: "Hero model and the anchor for the session: the tilt never changes, only our place changes. Both axes are drawn parallel on purpose.",
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_IDO_DEC = composeGlanceNotes({
  answer: "leaning in means the midday sun climbs to about 75 degrees and stays up about 14 h 50 m",
  beats: [
    "POINT to the arc. SAY: This is our December sun, every hour of one day, exactly like the video.",
    "SAY: Because we lean towards the sun, it climbs high at midday, nearly overhead, and it takes a long slow path to get down.",
    "SAY: So leaning in buys us two things. A higher sun, and a longer day. That is our summer.",
    "ASK: Which of those two would make it hotter? 15 sec, turn and tell. EXPECT: both, more hours and more direct light. ACCEPT: either one with a reason.",
  ],
  trap: "reading the arc as the sun getting physically closer at noon. Fix: point at the fixed horizon, it is the angle we see it at, student re-states.",
  stretch: "explain why a high sun heats the ground more than a low sun, using the torch.",
  help: "trace the arc with a finger, then say higher or lower and longer or shorter.",
  prep: "The effect half of the explanation; the tilt model gave the cause. Melbourne figures are real and verified, not illustrative.",
  tag: "[I Do | Explicit teaching | SC2 | HITS 4]",
  sources: ["Melbourne noon sun altitude and daylight length, latitude about 37.8 S"],
});

const NOTES_CFU_HINGE = composeGlanceNotes({
  answer: "B - distance cannot be the reason, because both places are the same distance on the same day",
  beats: [
    "SAY: Here is the fact. Right now in December we are having summer and London is having winter. Same planet. Same day.",
    "ASK: What does that prove? 30 sec think. Write A, B or C, chin it, boards up. EXPECT: B.",
    "SCAN boards, back row first. 80%+ -> cold call one B board: Convince us. Then reveal and move to We Do. Less -> use the re-teach slide that follows, then re-ask this question.",
    "REVEAL after every board is up.",
  ],
  trap: "choosing A, summer means closer to the sun. Fix: name that Earth is actually nearest the sun in January, our summer, which makes A feel right; then ask what London is doing on that same day, student answers.",
  prep: "The decision point of the lesson. Wrong answers are diagnostic: A is the distance misconception, C is the sun-moves misconception. Both get named, not just marked wrong.",
  tag: "[CFU hinge | Supported application | SC3 | HITS 7, 8]",
});

const NOTES_RETEACH = composeGlanceNotes({
  answer: "the lit half is whichever half leans towards the torch; distance never changes",
  beats: [
    "SKIP this slide if the hinge showed 80%+ on B.",
    "DIM the lights. SHOW the globe and torch. SAY: The torch is the sun. It does not move and it does not get closer.",
    "MODEL: hold the globe with the south leaning towards the torch. POINT: look where the light lands hardest. That is our summer.",
    "MODEL: walk the globe to the far side, rod still locked. POINT: same distance from the torch, but now the north is lit. Australia is in the dim part.",
    "ASK: Did I ever move the torch closer? 10 sec, thumbs only, voices off. EXPECT: thumbs down, no.",
    "SCAN thumbs. 80%+ -> back to the We Do. Less -> hand the globe to a student to walk it around while the class calls the lit half, then re-ask the hinge.",
  ],
  trap: "watching the globe and missing that the torch never moved. Fix: put a chair by the torch as a fixed marker, student re-runs the walk.",
  prep: "Different on-ramp from the I Do: physical and hands-on rather than a diagram, and it isolates distance by holding it constant. Optional, evidence-triggered only.",
  tag: "[Re-teach | Supported application | SC2 | HITS 4, 10]",
});

const NOTES_WEDO_JUN = composeGlanceNotes({
  answer: "north leans in; our midday sun is LOWER, about 29 degrees, and up for only about 9 h 35 m",
  beats: [
    "POINT to the June Earth. SAY: Your turn to use the model. Do not guess, read it off the picture.",
    "ASK: Which pole leans towards the sun in June? 15 sec, boards up on cue. EXPECT: north.",
    "ASK: So is our midday sun higher or lower than December, and is our day longer or shorter? 30 sec, boards up on cue. EXPECT: lower and shorter.",
    "SCAN boards. 80%+ -> cold call one board: How did the model tell you? Then reveal. Less -> re-run the globe walk to June, re-ask the pole question first.",
    "REVEAL after boards scanned. SAY: Lower and shorter. You read that off the model, you did not guess it.",
  ],
  trap: "reading north leans in as our summer. Fix: point at Melbourne on the globe in the dim half, student names our season.",
  stretch: "work out what the arc looks like at the equator, where neither pole leans in.",
  help: "cover the June Earth and ask only: is Melbourne in the bright half or the dim half?",
  prep: "Problem pair: December was modelled, June is theirs, same structure with the surface changed. The reveal completes the side-by-side the video opened with.",
  tag: "[We Do | Supported application | SC2 | HITS 4, 7]",
});

const NOTES_SIDEBYSIDE = composeGlanceNotes({
  answer: "December: 75 degrees and 14 h 50 m. June: 29 degrees and 9 h 35 m. The tilt causes both.",
  beats: [
    "POINT to both arcs. SAY: This is what you watched in the video, drawn as a diagram. Every dot is the sun, one hour apart.",
    "SAY: Nothing about the sun changed. Nothing about the tilt changed. Only which way we were leaning.",
    "ASK: Which arc would give a longer shadow at midday? 20 sec, turn and tell. EXPECT: June, because the sun is low. ACCEPT: the right-hand one.",
  ],
  trap: "reading the flatter June arc as the sun travelling slower. Fix: point at the equal hourly gaps on both arcs, student re-states that June's path is just shorter.",
  stretch: "predict what this diagram looks like at the north pole in June.",
  help: "trace both arcs with a finger and just say which one is up longer.",
  prep: "Consolidation and the payoff the launch promised: they can now explain the clip they watched cold. Keep to 3 minutes.",
  tag: "[We Do | Supported application | SC2 | HITS 4, 6]",
  sources: ["Melbourne noon sun altitude and daylight length, latitude about 37.8 S"],
});

const NOTES_CYCLE = composeGlanceNotes({
  answer: "the two solstices are opposite each other; halfway between, neither pole leans in and days are about equal",
  beats: [
    "POINT around the loop. SAY: The solstices are the two extremes. Everything in between is Earth sliding between them.",
    "POINT to March and September. SAY: Here neither pole leans towards the sun, so day and night are about the same length everywhere.",
    "ASK: When are our days getting longer, after June or after December? 20 sec, turn and tell. EXPECT: after June, we are heading back towards leaning in.",
  ],
  trap: "expecting the hottest day to be the solstice. Fix: the ground keeps soaking up heat for weeks, which is why January is hotter than December.",
  stretch: "find out the real name for the two in-between days.",
  prep: "Answers the question students always ask next: what about the rest of the year. Keep brisk, under 4 minutes.",
  tag: "[We Do | Retention and recall | SC3 | HITS 6]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - look for the word tilt or lean, plus one effect (higher sun or longer day)",
  beats: [
    "SAY: On the worksheet the orbit is already drawn for you. Your job is the thinking, not the drawing.",
    "SAY: Label the two poles, mark which half leans in, then write the explanation in the box.",
    "TIME: 12 minutes. CIRCULATE and read the December explanation first, it is the one the exit ticket needs.",
    "SAY: If you write because it is closer to the sun, cross it out. That is the trap we killed.",
  ],
  trap: "writing it is summer because it is hot. Fix: ask what makes it hot, push back to the lean, student rewrites with the word tilt.",
  stretch: "worksheet challenge box: explain why the north pole gets 24 hours of daylight in June.",
  help: "give the sentence frame: In December the ... half of Earth leans towards the sun, so the sun is ... and the day is ...",
  prep: "Different content from the We Do: they now write the whole explanation rather than read one step off the model. Worksheet Section 1 rebuilds the model for anyone who needs it.",
  tag: "[You Do | Mastery and application | SC2 | HITS 4, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "because the southern half of Earth is leaning towards the sun, so the sun climbs higher and stays up longer",
  beats: [
    "SAY: One question. Your own words. You have four minutes.",
    "COLLECT on a slip or in books. SAY: I am reading for one word: tilt, or lean.",
    "SORT as you collect. Any answer using closer to the sun goes in a pile for tomorrow's first two minutes.",
  ],
  trap: "naming December without naming the lean. Fix: prompt with why does December do that, student adds the cause.",
  prep: "Assesses SC2. Do not show the SC number to students. Sort into: names the lean, names an effect only, still on distance.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - self-assessment against the three criteria",
  beats: [
    "POINT to criterion one. ASK: Can you point to the half leaning towards the sun? 10 sec, thumbs only, voices off. EXPECT: mostly thumbs up.",
    "READ the thumbs across the room. SAY back what you saw, naming the strategy, not the person.",
    "ASK: What is one thing you will notice next time you walk outside at midday? 20 sec, turn and tell. EXPECT: open - how high the sun gets, or how early it is going dark.",
  ],
  prep: "Ties back to the launch clip: they can now explain what they saw in the first five minutes. Note any thumbs-down on criterion two for tomorrow.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* ------------------------------------------------------------------ */
/*  Companion PDFs                                                     */
/*  The orbit model is DRAWN on paper with the same tilt direction and */
/*  layout as the slide, so the worksheet scaffolds the same model     */
/*  students just saw rather than describing it in prose.              */
/* ------------------------------------------------------------------ */

const TILT_SIN = Math.sin((TILT_DEG * Math.PI) / 180);
const TILT_COS = Math.cos((TILT_DEG * Math.PI) / 180);

// One Earth on paper: an OUTLINE globe (so it can be shaded with a pencil),
// its tilted axis, and the equator drawn perpendicular to that axis so the
// two hemispheres are visibly separated for shading.
function drawEarthPdf(doc, cx, cy, r, opts) {
  const o = opts || {};
  const axisHalf = r + 17;
  const topX = cx - axisHalf * TILT_SIN;
  const topY = cy - axisHalf * TILT_COS;
  const botX = cx + axisHalf * TILT_SIN;
  const botY = cy + axisHalf * TILT_COS;

  doc.save();
  doc.lineWidth(1.4).strokeColor(hexPdf(C.CHARCOAL));
  doc.moveTo(topX, topY).lineTo(botX, botY).stroke();

  doc.lineWidth(1.4).strokeColor(hexPdf(EARTH_COLOR));
  doc.circle(cx, cy, r).fillAndStroke("#FFFFFF", hexPdf(EARTH_COLOR));

  // Equator: perpendicular to the axis, so north and south read correctly.
  doc.lineWidth(0.8).strokeColor(hexPdf(C.MUTED)).dash(3, { space: 2 });
  doc.moveTo(cx - r * TILT_COS, cy + r * TILT_SIN)
    .lineTo(cx + r * TILT_COS, cy - r * TILT_SIN).stroke();
  doc.undash();

  // Empty label boxes for N and S, parked just BEYOND each end of the axis so
  // the axis line does not run through the box students have to write in.
  const boxOut = axisHalf + 13;
  [
    [cx - boxOut * TILT_SIN, cy - boxOut * TILT_COS],
    [cx + boxOut * TILT_SIN, cy + boxOut * TILT_COS],
  ].forEach(([px, py]) => {
    doc.lineWidth(1).strokeColor(hexPdf(C.CHARCOAL));
    doc.rect(px - 9, py - 9, 18, 18).fillAndStroke("#FFFFFF", hexPdf(C.CHARCOAL));
  });

  if (o.caption) {
    // Sits clear of the lower pole label box, which reaches to about cy + 59.
    doc.fontSize(9).font("Sans-Bold").fillColor(hexPdf(C.CHARCOAL));
    doc.text(o.caption, cx - 60, cy + r + 46, { width: 120, align: "center" });
  }
  doc.restore();
}

function drawOrbitTiltPdf(doc, cx, cy, rx, ry, opts) {
  const o = opts || {};
  doc.save();

  for (let i = 0; i < 36; i++) {
    const a = (Math.PI * 2 * i) / 36;
    doc.circle(cx + rx * Math.cos(a), cy + ry * Math.sin(a), 1.4)
      .fill(hexPdf(C.MUTED));
  }

  doc.circle(cx, cy, 17).fill(hexPdf(SUN_COLOR));
  doc.fontSize(8).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text("Sun", cx - 17, cy - 4, { width: 34, align: "center" });

  drawEarthPdf(doc, cx - rx, cy, 24, { caption: o.leftCaption });
  drawEarthPdf(doc, cx + rx, cy, 24, { caption: o.rightCaption });

  doc.restore();
}

function hexPdf(color) {
  return "#" + String(color).replace(/^#/, "");
}

async function writeWorksheet() {
  const doc = createPdf({ title: "Session 1 Worksheet" });
  let y = addPdfHeader(doc, "The Solstice", {
    subtitle: "Why is our longest day in December?",
    color: C.PRIMARY,
    showNameDate: true,
  });

  y = addSectionHeading(doc, "1. Label the model", y + 4, { color: C.PRIMARY });
  y = addBodyText(doc,
    "Write N and S in the boxes. Shade the half of Earth that is leaning towards the sun.",
    y, { fontSize: 11 });

  drawOrbitTiltPdf(doc, PAGE_MID, y + 84, 150, 42, {
    leftCaption: "22 December",
    rightCaption: "21 June",
  });
  y += 180;

  y = addSectionHeading(doc, "2. Why is 22 December our longest day?", y, { color: C.PRIMARY });
  y = addBodyText(doc, "Use the word tilt. You can start like this:", y, { fontSize: 11 });
  y = addBodyText(doc,
    "In December the ______ half of Earth leans towards the sun, so the sun is ______ and the day is ______.",
    y, { fontSize: 10.5, italic: true, color: C.MUTED });
  y = addLinedArea(doc, y + 4, 3, { lineSpacing: 26 });

  y = addSectionHeading(doc, "3. Now June", y + 6, { color: C.SECONDARY });
  y = addBodyText(doc,
    "On 21 June, is our midday sun higher or lower than in December? How do you know?",
    y, { fontSize: 11 });
  y = addLinedArea(doc, y + 4, 3, { lineSpacing: 26 });

  addTipBox(doc,
    "Challenge: in June the north pole gets 24 hours of daylight. Use the model to explain why the sun never sets there.",
    y + 6, { color: C.ALERT });

  addPdfFooter(doc, "Big Ideas | Session 1 | The Solstice");
  await writePdf(doc, path.join(RES_DIR, "Session 1 Worksheet.pdf"));
}

async function writeAnswerKey() {
  const doc = createPdf({ title: "Session 1 Answer Key" });
  let y = addPdfHeader(doc, "The Solstice: Answer Key", {
    subtitle: "Teacher copy. Section 2 is the exit-ticket target.",
    color: C.SECONDARY,
    showNameDate: false,
  });

  y = addSectionHeading(doc, "1. Label the model", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc,
    "Both axes lean the same way, top towards the left. On 22 December the SOUTH pole is the one "
    + "tilted towards the sun, so students shade the lower half. On 21 June the NORTH pole is tilted "
    + "towards the sun, so they shade the upper half. If a student has drawn the two axes leaning in "
    + "different directions, that is the see-saw misconception: walk the globe around the torch again "
    + "without twisting the rod.",
    y, { fontSize: 10.5 });

  y = addSectionHeading(doc, "2. Why is 22 December our longest day?", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc,
    "Look for: the southern half of Earth is leaning towards the sun, so the midday sun climbs higher "
    + "(about 75 degrees in Melbourne) and stays up longer (about 14 h 50 m).",
    y, { fontSize: 10.5 });
  y = addBodyText(doc,
    "Accept any answer that names the lean or tilt AND one effect (higher sun OR longer day). "
    + "Do NOT accept 'because Earth is closer to the sun'. That is the misconception this lesson "
    + "targets, and in Australia it feels right because Earth really is nearest the sun in January. "
    + "The reply: on that same day London is freezing, and it is the same distance from the sun.",
    y, { fontSize: 10.5 });

  y = addSectionHeading(doc, "3. Now June", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc,
    "Lower. On 21 June the north pole leans towards the sun, so Melbourne leans away. The midday sun "
    + "only reaches about 29 degrees and we get about 9 h 35 m of daylight. Look for students reading "
    + "this off the model rather than recalling it.",
    y, { fontSize: 10.5 });

  y = addSectionHeading(doc, "Challenge", y + 4, { color: C.ALERT });
  y = addBodyText(doc,
    "In June the north pole is tilted towards the sun, so that whole pole stays in the sunlit half of "
    + "Earth as it spins. It turns around without ever crossing into the dark side, so the sun never "
    + "sets. Strong answers mention the spin AND the lean together.",
    y, { fontSize: 10.5 });

  addTipBox(doc,
    "Sort the exit tickets into three piles: names the lean, names an effect only, still on distance. "
    + "The third pile is your first two minutes tomorrow.",
    y + 6, { color: C.PRIMARY });

  addPdfFooter(doc, "Big Ideas | Session 1 | Answer Key");
  await writePdf(doc, path.join(RES_DIR, "Session 1 Answer Key.pdf"));
}

const PAGE_MID = 595.28 / 2;

/* ------------------------------------------------------------------ */
/*  Build                                                              */
/* ------------------------------------------------------------------ */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  /* 1. Title */
  titleSlide(pres,
    "Why Is Our Longest Day In December?",
    "The summer and winter solstice",
    "Session 1 of 2 | Year 5/6 Science",
    NOTES_TITLE);

  /* 2. Teacher Resources */
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  /* 3. Teacher-facing overview */
  contentSlide(pres, "For the Teacher", C.CHARCOAL,
    "Overview: two sessions in this deck",
    [
      "Session 1 (Science): why Melbourne's longest day is in December. Shape: observe first, then model.",
      "Session 2 (Numeracy): powers of 10, and why Graham's number needed a new way to write.",
      "The two sessions are independent. Session 2 does not build on Session 1.",
      "Today's decision points: the hinge after I Do, boards after We Do, then the exit ticket.",
      "Assumed: no prior astronomy. Everything needed is built in this session.",
    ],
    NOTES_OVERVIEW, FOOTER);

  /* 4. Launch: the video */
  contentSlide(pres, "Launch", C.SECONDARY,
    "Watch the sun",
    [
      "Two videos of the same sky, side by side.",
      "One is the longest day of the year. One is the shortest.",
      "Watch the sun at midday. How high does it get?",
    ],
    NOTES_LAUNCH, FOOTER,
    (s, guide) => {
      const x = guide.rightX;
      const w = guide.rightW;
      const y = guide.panelTopPadded;

      s.addShape("roundRect", {
        x, y, w, h: 1.30, rectRadius: 0.10,
        fill: { color: C.ALERT },
        line: { color: C.ALERT, width: 1 },
      });
      s.addText([{
        text: "Play the video",
        options: {
          hyperlink: { url: VIDEO_URL },
          color: C.WHITE, fontSize: 19, fontFace: FONT_H, bold: true,
        },
      }], {
        x, y, w, h: 1.30, align: "center", valign: "middle", margin: 0,
      });

      addInstructionCard(s, [
        { text: "On your whiteboards", role: "header" },
        { text: "", role: "spacer" },
        { text: "What was different about the sun?", role: "body" },
        { text: "", role: "spacer" },
        { text: "Write it, chin it, show me.", role: "emphasis" },
      ], {
        x, y: y + 1.48, w, h: 1.70,
        strip: C.PRIMARY, fill: C.WHITE,
      });
    });

  /* 5. LI and SC */
  liSlide(pres,
    "I can explain why Melbourne's longest and shortest days happen, using Earth's tilt.",
    [
      "I can point to the half of Earth that is leaning towards the sun.",
      "I can explain why Melbourne's longest day is in December.",
      "I can explain why England has winter when we have summer.",
    ],
    NOTES_LI, FOOTER);

  /* 6-7. Vocabulary: one word per slide */
  keyWordSlide(pres, {
    word: "Axis",
    meaning: "The imaginary line Earth spins around, from the north pole to the south pole.",
    example: "Earth spins around its axis once a day. That gives us day and night.",
    routine: ["Say it", "Point to it", "Use it"],
    color: C.PRIMARY,
  }, NOTES_VOCAB_AXIS, FOOTER);

  keyWordSlide(pres, {
    word: "Solstice",
    meaning: "The day the midday sun is at its highest, or its lowest, for the whole year.",
    example: "Sol means sun. Sistere means to stand still. The sun stops climbing and turns back.",
    routine: ["Say it", "Say why", "Use it"],
    color: C.SECONDARY,
  }, NOTES_VOCAB_SOLSTICE, FOOTER);

  /* 8. I Do: the tilt model (hero) */
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "The tilt never changes");

    drawOrbitTilt(s, {
      cx: 5.0, cy: 2.62, rx: 3.20, ry: 1.02,
      earthR: 0.44, sunR: 0.44, poleR: 0.10, labelSize: 13,
    });

    addTextOnShape(s, "Earth's axis always points the same way. Only our place in the orbit changes.", {
      x: 0.9, y: 4.34, w: 8.2, h: 0.50, rectRadius: 0.08,
      fill: { color: C.BG_LIGHT },
      line: { color: C.PRIMARY, width: 1.2 },
    }, {
      fontSize: 15, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_TILT);
    runSlideDiagnostics(s, pres);
  }

  /* 9. I Do: December */
  contentSlide(pres, "I Do", C.PRIMARY,
    "December: we lean in",
    [
      "The southern half leans towards the sun.",
      "The midday sun climbs high.",
      "The day is long.",
    ],
    NOTES_IDO_DEC, FOOTER,
    (s, guide) => {
      const x = guide.rightX;
      const w = guide.rightW;
      const y = guide.panelTopPadded;
      addCard(s, x, y, w, 2.90, { strip: C.PRIMARY, fill: C.WHITE });
      drawSunPath(s, x + 0.14, y + 0.16, w - 0.28, 2.30, {
        altitudeDeg: DEC_ALT,
        color: C.ALERT,
        label: "22 December",
      });
      addTextOnShape(s, `${DEC_DAYLIGHT} of daylight`, {
        x: x + 0.55, y: y + 2.44, w: w - 1.10, h: 0.34, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 12.5, fontFace: FONT_B, color: C.WHITE, bold: true });
    });

  /* 10. CFU hinge (reveal) */
  withReveal(
    () => cfuSlide(pres, "Check", "Same day, opposite seasons", "Show Me Boards",
      "In December we have summer. On that same day, London has winter.\n\nWhat does that prove?\n\nA: summer means closer to the sun\nB: distance is not the reason\nC: the sun moves to us",
      NOTES_CFU_HINGE, FOOTER),
    (s) => {
      T.addRevealAnswerBar(s, ["B: distance is not the reason"], { y: 4.30, h: 0.56, fontSize: 18 });
    }
  );

  /* 11. Optional re-teach: torch and globe. A physical demo rather than a
     diagram to read - a different on-ramp for anyone the I Do model lost. */
  contentSlide(pres, "Another way to see it", C.SECONDARY,
    "The torch never moves",
    [
      "Optional: skip if the check went well.",
      "Watch the globe, not the torch.",
      "The torch never moves closer.",
      "Only the lean changes.",
    ],
    NOTES_RETEACH, FOOTER,
    (s, guide) => {
      const x = guide.rightX;
      const w = guide.rightW;
      const y = guide.panelTopPadded;
      addCard(s, x, y, w, 2.72, { strip: C.SECONDARY, fill: C.WHITE });

      const midY = y + 1.10;
      drawSun(s, x + 0.78, midY, 0.34, "Torch");
      drawEarth(s, x + 3.30, midY, 0.40, { litPole: "south", poleR: 0.085, labelSize: 11 });

      // Distance bar with end caps: the quantity the demo holds constant.
      const barX = x + 1.28;
      const barW = 1.56;
      s.addShape("roundRect", {
        x: barX, y: midY - 0.02, w: barW, h: 0.045, rectRadius: 0.02,
        fill: { color: C.CHARCOAL }, line: { color: C.CHARCOAL, width: 0.5 },
      });
      [barX, barX + barW].forEach((capX) => {
        s.addShape("roundRect", {
          x: capX - 0.02, y: midY - 0.13, w: 0.04, h: 0.26, rectRadius: 0.015,
          fill: { color: C.CHARCOAL }, line: { color: C.CHARCOAL, width: 0.5 },
        });
      });
      s.addText("Never changes", {
        x: barX - 0.10, y: midY - 0.46, w: barW + 0.20, h: 0.26,
        fontSize: 11.5, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      addTextOnShape(s, "Same distance. Different lean.", {
        x: x + 0.42, y: y + 2.26, w: w - 0.84, h: 0.36, rectRadius: 0.06,
        fill: { color: C.SECONDARY },
      }, { fontSize: 12.5, fontFace: FONT_B, color: C.WHITE, bold: true });
    });

  /* 12. We Do: June (reveal completes the side-by-side) */
  withReveal(
    () => contentSlide(pres, "We Do", C.SECONDARY,
      "Now work out June",
      [
        "Which pole leans in?",
        "Higher sun or lower?",
        "Longer day or shorter?",
      ],
      NOTES_WEDO_JUN, FOOTER,
      (s, guide) => {
        const x = guide.rightX;
        const w = guide.rightW;
        const y = guide.panelTopPadded;
        addCard(s, x, y, w, 2.60, { strip: C.SECONDARY, fill: C.WHITE });
        drawOrbitTilt(s, {
          cx: x + w / 2, cy: y + 1.16, rx: 1.62, ry: 0.62,
          earthR: 0.26, sunR: 0.26, poleR: 0.07, labelSize: 10, chips: false,
        });
        addTextOnShape(s, "Read it off the model", {
          x: x + 0.50, y: y + 2.06, w: w - 1.00, h: 0.34, rectRadius: 0.06,
          fill: { color: C.SECONDARY },
        }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
      }),
    (s) => {
      T.addRevealAnswerBar(s, ["North leans in", "Sun LOWER", "Day SHORTER"],
        { y: 4.30, h: 0.56, fontSize: 17 });
    }
  );

  /* 13. Side by side: the payoff the launch video promised */
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "The same sky, six months apart");

    const cardY = CONTENT_TOP;
    addCard(s, 0.5, cardY, 4.5, 2.86, { strip: C.PRIMARY, fill: C.WHITE });
    drawSunPath(s, 0.64, cardY + 0.14, 4.22, 2.26, {
      altitudeDeg: DEC_ALT, color: C.ALERT, label: "22 December",
    });
    addTextOnShape(s, `${DEC_DAYLIGHT} of daylight`, {
      x: 1.05, y: cardY + 2.44, w: 3.4, h: 0.34, rectRadius: 0.06,
      fill: { color: C.PRIMARY },
    }, { fontSize: 12.5, fontFace: FONT_B, color: C.WHITE, bold: true });

    addCard(s, 5.2, cardY, 4.3, 2.86, { strip: C.SECONDARY, fill: C.WHITE });
    drawSunPath(s, 5.32, cardY + 0.14, 4.06, 2.26, {
      altitudeDeg: JUN_ALT, color: C.SECONDARY, label: "21 June",
    });
    addTextOnShape(s, `${JUN_DAYLIGHT} of daylight`, {
      x: 5.65, y: cardY + 2.44, w: 3.4, h: 0.34, rectRadius: 0.06,
      fill: { color: C.SECONDARY },
    }, { fontSize: 12.5, fontFace: FONT_B, color: C.WHITE, bold: true });

    addTextOnShape(s, "Same city. Same sun. One tilt, two very different days.", {
      x: 1.4, y: 4.36, w: 7.2, h: 0.46, rectRadius: 0.08,
      fill: { color: C.BG_LIGHT },
      line: { color: C.SECONDARY, width: 1.2 },
    }, {
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_SIDEBYSIDE);
    runSlideDiagnostics(s, pres);
  }

  /* 14. The year as a loop */
  cycleDiagramSlide(pres, "We Do",
    "The solstices are two points on a loop",
    "With your partner",
    [
      "The solstices are the two extremes.",
      "What happens halfway between them?",
      "Turn and tell, 20 seconds.",
    ],
    "One year",
    [
      { label: "22 Dec", detail: "South leans in. Longest day.", color: C.PRIMARY },
      { label: "March", detail: "Neither leans. Days equal.", color: C.MUTED },
      { label: "21 Jun", detail: "North leans in. Shortest day.", color: C.SECONDARY },
      { label: "Sept", detail: "Neither leans. Days equal.", color: C.MUTED },
    ],
    NOTES_CYCLE, FOOTER);

  /* 14. You Do */
  contentSlide(pres, "You Do", C.ALERT,
    "Explain it yourself",
    [
      "First: label the poles on the orbit model.",
      "Next: shade the half that leans towards the sun in December.",
      "Then: write why our longest day is in December.",
    ],
    NOTES_YOUDO, FOOTER,
    (s, guide) => {
      const x = guide.rightX;
      const w = guide.rightW;
      const y = guide.panelTopPadded;
      addInstructionCard(s, [
        { text: "Use the word", role: "header" },
        { text: "", role: "spacer" },
        { text: "tilt", role: "body" },
        { text: "", role: "spacer" },
        { text: "Not 'closer to the sun'.", role: "emphasis" },
        { text: "", role: "spacer" },
        { text: "12 minutes.", role: "emphasis" },
      ], {
        x, y, w, h: 2.90,
        strip: C.ALERT, fill: C.WHITE,
      });
    });

  /* 15. Exit ticket */
  exitTicketSlide(pres,
    "It is 22 December in Melbourne. It is our longest day of the year.\n\nExplain why. Use the word tilt.",
    NOTES_EXIT, FOOTER);

  /* 16. Closing */
  closingSlide(pres, {
    reflectionPrompt: "What will you notice next time you are outside at midday?",
    scItems: [
      "I can point to the half of Earth that is leaning towards the sun.",
      "I can explain why Melbourne's longest day is in December.",
      "I can explain why England has winter when we have summer.",
    ],
    selfAssessment: "Thumbs up / sideways / down",
    // One takeaway only: closingSlide silently drops any row that would fall
    // below the safe zone once the SC list and self-assessment are placed.
    takeaways: [
      "The tilt never changes. Our place in the orbit does.",
    ],
  }, NOTES_CLOSING);

  const fileName = path.join(OUT_DIR, "Big Ideas - Session 1 - The Solstice.pptx");
  await pres.writeFile({ fileName });
  console.log("PPTX written to " + fileName);

  await writeWorksheet();
  await writeAnswerKey();
  console.log("PDFs written to " + RES_DIR);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
