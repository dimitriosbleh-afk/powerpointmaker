"use strict";

// =============================================================================
//  Shaping Australia - Influential People Presentation Project
//  Year 6 | Inquiry (The Humanities / History + English)
//
//  PURPOSE (per teacher brief): this is ONE informational deck that explains a
//  multi-session project from start to finish, so the teacher can walk students
//  through it once and students can then work independently across the sessions.
//  It is NOT a single explicit-teaching lesson - it is a project walkthrough.
//
//  The 5 project steps (combined into this one deck):
//    1. Choose & research an influential person
//    2. Write the speech (first person, as the person)
//    3. Practise performing in character
//    4. Deliver the speech (assessed on the Speech Rubric)
//    5. Write & learn Celebration Night questions
//
//  Source-faithful to the teacher's existing resources:
//    - Speech Rubric (Introduction / Body topic ideas / Conclusion / Body
//      Language & Speaking Skills)
//    - Example Questions sheet (5 starter questions, verbatim)
//    - Influential People slideshow (era structure + named example people)
//    - Getting Into Character video (teacher-provided link)
//
//  RESOURCE DECISION: no new printed/PDF resources. The teacher already holds
//  the rubric, question sheet and slideshow; students produce their own prop,
//  cue cards and question ring, and record research in inquiry books / Google
//  Classroom. The deck presents that content on-screen.
//
//  This builds toward "Shaping Australia Day / Celebration Night".
//  Australian spelling throughout. Follows the prior unit on how Australia has
//  changed over time (students bring that background).
// =============================================================================

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");

// Inquiry / grade56 / variant 3 "Navigator: deep navy & academic teal".
// PRIMARY navy 1E3050, SECONDARY teal 1E7070, ACCENT gold 886818,
// ALERT crimson 801A1A, SUCCESS green 2A6848. Dignified, scholarly - fits a
// Year 6 Australian-history presentation project. Georgia headings.
const T = createTheme("inquiry", "grade56", 3);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, closingSlide, investigationSlide,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape,
  runSlideDiagnostics, composeNotes,
} = T;

const OUT_DIR = "output/Shaping Australia - Influential People Project";
fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Shaping Australia - Influential People Project";
pres.author = "James Hooke";

const FOOT_ALL = "Shaping Australia  |  Influential People Project  |  Year 6";
function stepFoot(n, name) {
  return `Shaping Australia  |  Step ${n} of 5: ${name}  |  Year 6`;
}

function notes(obj, opts) {
  return composeNotes(obj, opts);
}

// ---------------------------------------------------------------------------
//  Shared custom-slide helpers
// ---------------------------------------------------------------------------

// Generic custom slide scaffold: top bar, badge, title, draw callback,
// footer, notes, diagnostics. Keeps every custom slide consistent and checked.
function infoSlide({ topColor, badge, badgeColor, badgeW, title, titleColor, draw, footer, notesText }) {
  const s = pres.addSlide();
  addTopBar(s, topColor || C.PRIMARY);
  addBadge(s, badge, { color: badgeColor || C.PRIMARY, w: badgeW || 2.2 });
  addTitle(s, title, titleColor ? { color: titleColor } : undefined);
  if (draw) draw(s);
  if (footer) addFooter(s, footer);
  if (notesText) s.addNotes(notesText);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

// A row of equal feature cards: each item = { tag, color, lines, bullet }.
function featureCards(s, items, { y, h, gap = 0.2, tagSize, lineSize }) {
  const n = items.length;
  const cardW = (9 - gap * (n - 1)) / n;
  items.forEach((it, i) => {
    const x = 0.5 + i * (cardW + gap);
    const color = it.color || C.PRIMARY;
    addCard(s, x, y, cardW, h, { strip: color, fill: C.WHITE });
    addTextOnShape(s, it.tag, {
      x: x + 0.14, y: y + 0.14, w: cardW - 0.28, h: 0.44, rectRadius: 0.06,
      fill: { color },
    }, {
      fontSize: tagSize || 15, fontFace: FONT_H, color: C.WHITE, bold: true,
    });
    const bodyY = y + 0.70;
    const lines = it.lines || [];
    s.addText(lines.map((t, j) => ({
      text: String(t),
      options: {
        bullet: it.bullet === true,
        breakLine: j < lines.length - 1,
        fontSize: lineSize || 12.5,
        color: C.CHARCOAL,
        paraSpaceAfter: 4,
      },
    })), {
      x: x + 0.18, y: bodyY, w: cardW - 0.36, h: y + h - bodyY - 0.12,
      fontFace: FONT_B, valign: "top", margin: 0, fit: "shrink", shrinkText: true,
    });
  });
}

// ===========================================================================
//  1. TITLE
// ===========================================================================

titleSlide(
  pres,
  "Shaping Australia",
  "Research an influential person, then become them and present your speech",
  "Year 6 Inquiry  |  A 5-step project building to our Celebration Night",
  notes({
    say: [
      "We are starting a big project. Each of you will become a person who helped shape Australia.",
      "You will research them, build a display, and present a speech as that person on our Celebration Night.",
      "Today I will walk us through the whole project so you know exactly what you are working towards.",
    ],
    do: [
      "Use this deck to take students through the project from start to finish.",
      "Move to the Project Overview and Teacher Resources slides next.",
    ],
    teacherNotes: {
      text: "This single deck explains the whole multi-session project so students can then work largely independently. It follows your unit on how Australia has changed over time, so students already have that background to draw on.",
    },
    watchFor: [
      "Excitement to pick a person before they understand 'influential' - hold that until the choosing step.",
    ],
    tag: "[Shaping Australia | Title | Project launch]",
  })
);

// ===========================================================================
//  2. TEACHER RESOURCES & MATERIALS  (immediately after title)
// ===========================================================================

contentSlide(
  pres,
  "Before You Teach", C.PRIMARY,
  "Teacher Resources & Materials",
  [
    "Influential People slideshow (teacher-provided) - show it when students are choosing their person.",
    "Speech Rubric (teacher-provided) - share when students start writing; mark with it at the presentation.",
    "Example Questions sheet (teacher-provided) - students copy it and add their own for Celebration Night.",
    "Getting Into Character video (teacher-provided link) - play before students rehearse in character.",
    "Device access + Google Classroom - the research starter link is posted there; students record in inquiry books.",
    "Cue cards and materials for the chosen display (poster, diorama, slideshow or recording). No new printed worksheets.",
  ],
  notes({
    say: [
      "Everything you need for this project is listed here.",
      "Most of it you already have - the slideshow, the rubric and the question sheet.",
    ],
    do: [
      "Have the Influential People slideshow ready to display on the choosing day.",
      "Post the research link on Google Classroom before Step 1.",
      "Book devices for the research and slideshow sessions.",
    ],
    teacherNotes: {
      text: "No new printed resources are needed. Students produce their own display, speech cue cards and Celebration Night question ring. Write down which person each student chooses so no one doubles up and you can support research.",
    },
    watchFor: [
      "Students who cannot access the slideshow or link - have a printed copy of the people list as a backup.",
    ],
    sources: [
      "Influential People slideshow (teacher-provided).",
      "Research starter: theculturetrip.com - The Most Influential Australian People You Should Know (posted on Google Classroom).",
    ],
    tag: "[Shaping Australia | Teacher Resources | Preparation]",
  }),
  FOOT_ALL
);

// ===========================================================================
//  3. PROJECT OVERVIEW  (teacher-facing)
// ===========================================================================

infoSlide({
  badge: "For the Teacher", badgeColor: C.PRIMARY, badgeW: 2.2,
  title: "Project overview",
  footer: FOOT_ALL,
  draw: (s) => {
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    // Left: what students produce
    const lx = 0.5, lw = 4.35;
    addCard(s, lx, cardY, lw, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("What each student produces", {
      x: lx + 0.2, y: cardY + 0.14, w: lw - 0.4, h: 0.3,
      fontSize: 15, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText([
      "A display prop - a poster, diorama, slideshow or recording about their person.",
      "A 3-5 minute speech, written and performed in first person AS their person.",
      "Learned answers to visitor questions, ready for Celebration Night.",
    ].map((t, i, a) => ({
      text: t,
      options: { bullet: true, breakLine: i < a.length - 1, fontSize: 13.5, color: C.CHARCOAL, paraSpaceAfter: 7 },
    })), {
      x: lx + 0.2, y: cardY + 0.52, w: lw - 0.4, h: cardH - 0.66,
      fontFace: FONT_B, valign: "top", margin: 0, fit: "shrink", shrinkText: true,
    });

    // Right: the 5 steps
    const rx = 5.15, rw = 4.35;
    addCard(s, rx, cardY, rw, cardH, { strip: C.PRIMARY, fill: C.BG_CARD });
    s.addText("The five steps", {
      x: rx + 0.2, y: cardY + 0.14, w: rw - 0.4, h: 0.3,
      fontSize: 15, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    const steps = [
      ["1", "Choose & research an influential person"],
      ["2", "Write the speech, in first person"],
      ["3", "Practise performing in character"],
      ["4", "Deliver the speech (marked on the Speech Rubric)"],
      ["5", "Write & learn Celebration Night questions"],
    ];
    const rowH = 0.56, row0 = cardY + 0.58;
    steps.forEach((st, i) => {
      const ry = row0 + i * rowH;
      addTextOnShape(s, st[0], {
        x: rx + 0.2, y: ry, w: 0.42, h: 0.42, rectRadius: 0.21,
        fill: { color: C.PRIMARY },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(st[1], {
        x: rx + 0.74, y: ry, w: rw - 0.94, h: 0.46,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });
  },
  notesText: notes({
    do: [
      "Skim this before you launch the project so the arc and end-point are clear.",
      "Decide how long each step gets and when Celebration Night falls.",
    ],
    teacherNotes: {
      text: "Assessment happens in Step 4 against the Speech Rubric. This project follows your 'how Australia has changed over time' unit, so students bring background on First Nations history, colonisation, immigration and Federation.",
      bullets: [
        "Steps 2-5 success criteria on the goal slides are taken directly from your planner.",
        "The 'create a display' criterion is taught on its own slide in Step 1 (Choose how you'll present).",
      ],
    },
    sensitivity: [
      "What it is: students choose and portray a real person, who may be a First Nations figure or connected to colonisation, conflict or recent events.",
      "Framing language: portray with respect; this is honouring a person's contribution, not impersonating or mocking them.",
      "Watch for: caricature, mock accents, or trivialising sensitive history. Guide choices and portrayals early.",
      "Protocol: if a student picks a culturally significant or sensitive figure, check the portrayal plan with them before they rehearse.",
    ],
    tag: "[Shaping Australia | Project Overview | Teacher planning]",
  }, { requireSay: false, requireDo: false }),
});

// ===========================================================================
//  4. YOUR MISSION + 5-STEP ROADMAP  (student-facing)
// ===========================================================================

infoSlide({
  topColor: C.SECONDARY,
  badge: "The Big Picture", badgeColor: C.SECONDARY, badgeW: 2.2,
  title: "Your mission, in five steps",
  footer: FOOT_ALL,
  draw: (s) => {
    // Hero one-liner
    addTextOnShape(s, "Choose a person who helped shape Australia - then become them and present on our Celebration Night.", {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 0.66, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true, align: "left", valign: "middle", margin: 0.14,
    });

    // 5 roadmap cards
    const y = 2.10, h = 2.95, gap = 0.16;
    const cardW = (9 - gap * 4) / 5;
    const colorCycle = [C.PRIMARY, C.SECONDARY, C.ACCENT, C.SUCCESS, C.PRIMARY];
    const cards = [
      ["1", "Choose & research", "Pick someone who made a positive change. Find out how and when."],
      ["2", "Write your speech", "A 3-5 minute speech as your person, in first person."],
      ["3", "Practise in character", "Use voice, body and face. Rehearse with a partner."],
      ["4", "Deliver your speech", "Present to the class. Marked on the Speech Rubric."],
      ["5", "Celebration Night", "Answer visitors' questions about your person."],
    ];
    cards.forEach((cd, i) => {
      const x = 0.5 + i * (cardW + gap);
      const color = colorCycle[i];
      addCard(s, x, y, cardW, h, { strip: color, fill: C.WHITE });
      const cx = x + cardW / 2;
      addTextOnShape(s, cd[0], {
        x: cx - 0.27, y: y + 0.2, w: 0.54, h: 0.54, rectRadius: 0.27,
        fill: { color },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(cd[1], {
        x: x + 0.1, y: y + 0.86, w: cardW - 0.2, h: 0.62,
        fontSize: 13, fontFace: FONT_H, color, bold: true, align: "center", valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      s.addText(cd[2], {
        x: x + 0.12, y: y + 1.52, w: cardW - 0.24, h: h - 1.66,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });
  },
  notesText: notes({
    say: [
      "Here is the whole journey on one slide. Five steps, finishing on Celebration Night.",
      "You have been learning how Australia changed over time. Now you will spotlight one person who helped shape it.",
      "By the end you will not just talk ABOUT your person - you will present AS them.",
    ],
    do: [
      "Walk left to right through the five steps so students see the destination.",
      "Point out that each step builds on the last.",
    ],
    teacherNotes: {
      text: "This is the student roadmap. Return to it at the start of each step so students always know where they are.",
    },
    watchFor: [
      "Students worried about presenting - reassure them we build up to it with research, writing and practice first.",
    ],
    tag: "[Shaping Australia | Mission & Roadmap | Student overview]",
  }),
});

// ===========================================================================
//  STEP 1 - CHOOSE & RESEARCH
// ===========================================================================

liSlide(
  pres,
  "I am choosing an influential person and researching how they helped shape Australia.",
  [
    "I can choose someone who helped shape Australia (before, during or after British settlement).",
    "I can find information about how my person helped shape Australia.",
    "I can identify when my person had their impact.",
  ],
  notes({
    say: [
      "Step 1 has two jobs: choose your person well, then research them.",
      "Read our three 'I can' statements. Choosing the right person matters most - we will look at that next.",
    ],
    do: [
      "Read the learning intention and success criteria aloud.",
      "Tell students they will see how to choose and how to research over the next few slides.",
    ],
    teacherNotes: {
      text: "The fourth planner criterion - creating a poster, diorama, slideshow or recording - is taught on its own slide later in this step.",
    },
    watchFor: [
      "Students rushing to a favourite celebrity - the next slide sets the 'influential, not just popular' test.",
    ],
    tag: "[Shaping Australia | Step 1 | Goal]",
  }),
  stepFoot(1, "Choose & Research")
);

// Influential vs popular - THE key conversation
infoSlide({
  topColor: C.SECONDARY,
  badge: "Choose Well", badgeColor: C.SECONDARY, badgeW: 2.0,
  title: "Influential, not just popular",
  footer: stepFoot(1, "Choose & Research"),
  draw: (s) => {
    const cardY = CONTENT_TOP, cardH = 2.85;
    const lw = 4.35, lx = 0.5, rx = 5.15, rw = 4.35;

    // Influential (positive)
    addCard(s, lx, cardY, lw, cardH, { strip: C.SUCCESS, fill: C.WHITE });
    addTextOnShape(s, "Influential", {
      x: lx + 0.18, y: cardY + 0.16, w: lw - 0.36, h: 0.5, rectRadius: 0.07,
      fill: { color: C.SUCCESS },
    }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    s.addText([
      "Made a direct, positive change to Australia as a society.",
      "Did something that improved life for others.",
    ].map((t, i, a) => ({
      text: t, options: { bullet: true, breakLine: i < a.length - 1, fontSize: 15, color: C.CHARCOAL, paraSpaceAfter: 8 },
    })), {
      x: lx + 0.2, y: cardY + 0.8, w: lw - 0.4, h: cardH - 0.95,
      fontFace: FONT_B, valign: "top", margin: 0, fit: "shrink", shrinkText: true,
    });

    // Just popular (not enough)
    addCard(s, rx, cardY, rw, cardH, { strip: C.ALERT, fill: C.WHITE });
    addTextOnShape(s, "Just popular", {
      x: rx + 0.18, y: cardY + 0.16, w: rw - 0.36, h: 0.5, rectRadius: 0.07,
      fill: { color: C.ALERT },
    }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    s.addText([
      "Famous, or just someone you like.",
      "Has not necessarily changed Australia for the better.",
      "E.g. picking a player only because they're on the team you support is not enough.",
    ].map((t, i, a) => ({
      text: t, options: { bullet: true, breakLine: i < a.length - 1, fontSize: 13.5, color: C.CHARCOAL, paraSpaceAfter: 6 },
    })), {
      x: rx + 0.2, y: cardY + 0.8, w: rw - 0.4, h: cardH - 0.95,
      fontFace: FONT_B, valign: "top", margin: 0, fit: "shrink", shrinkText: true,
    });

    // Bottom prompt
    addTextOnShape(s, "Ask yourself: What did my person DO to improve our society?", {
      x: 0.5, y: cardY + cardH + 0.16, w: 9, h: 0.55, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
  },
  notesText: notes({
    say: [
      "Before you choose, here is the most important test. Is your person influential, or just popular?",
      "Influential means they made a real, positive change to Australia. Popular just means famous or liked.",
      "Ask: what did this person actually DO to improve our society?",
    ],
    do: [
      "Have a class conversation about the difference, using a couple of student suggestions.",
      "Steer students towards positive contributions, not just fame.",
    ],
    teacherNotes: {
      text: "This is the conversation the planner flags as important. Keep the test simple: a direct, positive change to Australia.",
    },
    sensitivity: [
      "What it is: some famous Australians are admired by some and controversial to others.",
      "Framing language: focus on the positive change a person made and why it mattered.",
      "Watch for: choices made only for shock value or pure fandom - redirect to genuine contribution.",
    ],
    watchFor: [
      "Students defending a choice with 'but they're famous' - return to 'what did they DO to improve society?'.",
    ],
    tag: "[Shaping Australia | Step 1 | Influential vs popular]",
  }),
});

// Who could you choose? - eras + example names
infoSlide({
  badge: "Find Your Person", badgeColor: C.PRIMARY, badgeW: 2.4,
  title: "Who could you choose?",
  footer: stepFoot(1, "Choose & Research"),
  draw: (s) => {
    const cardY = CONTENT_TOP, cardH = 2.95, gap = 0.25;
    const cardW = (9 - gap * 2) / 3;
    const cols = [
      {
        color: C.SECONDARY, pill: "Before 1788", sub: "First Nations peoples",
        names: ["First Nations peoples - here for 60,000+ years", "Choose a group or culture to teach us about"],
      },
      {
        color: C.PRIMARY, pill: "1788 - 1901", sub: "Settlement to nation",
        names: ["Captain Arthur Phillip", "Matthew Flinders", "Bennelong", "Peter Lalor", "Edmund Barton"],
      },
      {
        color: C.ACCENT, pill: "Modern Australia", sub: "1901 - today",
        names: ["Eddie Mabo", "Cathy Freeman", "Adam Goodes", "Steve Irwin", "Julia Gillard"],
      },
    ];
    cols.forEach((col, i) => {
      const x = 0.5 + i * (cardW + gap);
      addCard(s, x, cardY, cardW, cardH, { strip: col.color, fill: C.WHITE });
      addTextOnShape(s, col.pill, {
        x: x + 0.16, y: cardY + 0.14, w: cardW - 0.32, h: 0.42, rectRadius: 0.06,
        fill: { color: col.color },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      s.addText(col.sub, {
        x: x + 0.18, y: cardY + 0.62, w: cardW - 0.36, h: 0.26,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true, bold: true, margin: 0,
      });
      s.addText(col.names.map((t, j, a) => ({
        text: t, options: { bullet: true, breakLine: j < a.length - 1, fontSize: 12, color: C.CHARCOAL, paraSpaceAfter: 5 },
      })), {
        x: x + 0.18, y: cardY + 0.94, w: cardW - 0.36, h: cardH - 1.06,
        fontFace: FONT_B, valign: "top", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "Many more in the Influential People slideshow - pick someone who made a positive change.", {
      x: 0.5, y: cardY + cardH + 0.16, w: 9, h: 0.5, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
  },
  notesText: notes({
    say: [
      "Your person can come from before, during or after British settlement - that is the first success criterion.",
      "Here are a few examples from each time. The Influential People slideshow has many more.",
      "These are just starting points - you might know someone else who shaped Australia.",
    ],
    do: [
      "Display the Influential People slideshow and talk through a few people from each era.",
      "Invite students to suggest others, then apply the influential-not-popular test.",
      "Write down each student's choice so no one doubles up.",
    ],
    teacherNotes: {
      text: "The three bands match success criterion 1 (before, during or after British settlement). Names are examples only - students research the facts themselves.",
    },
    sensitivity: [
      "What it is: choices may include First Nations people and figures tied to colonisation.",
      "Framing language: every choice is about a positive contribution; portray people with respect.",
      "Protocol: discuss respectful portrayal at the point of choosing, not just at performance.",
    ],
    watchFor: [
      "Two students choosing the same person - encourage variety so Celebration Night covers more of our history.",
    ],
    sources: [
      "Influential People slideshow (teacher-provided): era structure and example people (e.g. Captain Arthur Phillip, Matthew Flinders, Bennelong, Eddie Mabo, Adam Goodes, Steve Irwin).",
    ],
    tag: "[Shaping Australia | Step 1 | Choosing a person]",
  }),
});

// Researching your person - investigationSlide (inquiry question + steps)
investigationSlide(
  pres,
  "Research",
  "Researching your influential person",
  "How did my person help shape Australia - and when did it happen?",
  null,
  [
    "Use the slideshow and the research link on Google Classroom to start.",
    "Find out about their life: childhood, family, and the events they are known for.",
    "Focus on what they did to shape Australia, and when it happened.",
    "Record facts in your inquiry book, in your own words.",
    "Note where you found each fact so you can trust it.",
  ],
  notes({
    say: [
      "Now you research. Your big question is on the screen: how did my person help shape Australia, and when?",
      "Start with the slideshow and the link on Google Classroom, then dig deeper.",
      "Write facts in your own words in your inquiry book - we will turn them into a speech next.",
    ],
    do: [
      "Hand out or open inquiry books; make sure the research link is on Google Classroom.",
      "Circulate and help students judge whether a source is reliable.",
    ],
    teacherNotes: {
      text: "Keep research focused on the success criteria: how the person shaped Australia and when. The notes they gather here become the body of their speech.",
    },
    enabling: [
      "ENABLING: give a short fact-finding frame - name, when they lived, one big thing they did, why it mattered.",
      "EXTENDING: ask students to find a quote from or about their person, and check a fact across two sources.",
    ],
    watchFor: [
      "Copy-pasting from websites - prompt students to put facts in their own words.",
      "Collecting trivia that does not show impact - refocus on what shaped Australia.",
    ],
    sources: [
      "Research starter: theculturetrip.com - The Most Influential Australian People You Should Know (posted on Google Classroom).",
    ],
    tag: "[Shaping Australia | Step 1 | Research]",
  }),
  stepFoot(1, "Choose & Research")
);

// Choose how you'll present - the display format (4th planner criterion)
infoSlide({
  topColor: C.ACCENT,
  badge: "Show It Your Way", badgeColor: C.ACCENT, badgeW: 2.6,
  title: "Choose how you'll present",
  footer: stepFoot(1, "Choose & Research"),
  draw: (s) => {
    featureCards(s, [
      { tag: "Poster", color: C.PRIMARY, lines: ["A bold visual display about your person."] },
      { tag: "Diorama", color: C.SECONDARY, lines: ["A 3D scene showing a moment from their life."] },
      { tag: "Slideshow", color: C.ACCENT, lines: ["Digital slides with images and key facts."] },
      { tag: "Recording", color: C.SUCCESS, lines: ["A short video or audio piece you create."] },
    ], { y: CONTENT_TOP, h: 2.8, gap: 0.2, tagSize: 15, lineSize: 12.5 });

    addTextOnShape(s, "Whatever you choose, it is the prop that supports your speech on Celebration Night.", {
      x: 0.5, y: CONTENT_TOP + 2.95, w: 9, h: 0.55, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true });
  },
  notesText: notes({
    say: [
      "You also choose how to display your research: a poster, a diorama, a slideshow or a recording.",
      "Pick the one that suits your person and your strengths.",
      "Remember, this display is the prop for your speech - it supports what you say.",
    ],
    do: [
      "Briefly show what each option could look like.",
      "Help students match their format to their person and the materials available.",
    ],
    teacherNotes: {
      text: "This teaches the planner's 'create a poster, diorama, slideshow or recording' criterion. The display is a prop for the Celebration Night speech, not a separate task.",
    },
    enabling: [
      "ENABLING: offer a poster template or a simple 3-slide structure for students who need a frame.",
      "EXTENDING: challenge students to combine formats, e.g. a recording that plays beside a poster.",
    ],
    watchFor: [
      "Students choosing the most elaborate option without a plan - check it is achievable in the time.",
    ],
    tag: "[Shaping Australia | Step 1 | Choose your display]",
  }),
});

// ===========================================================================
//  STEP 2 - WRITE YOUR SPEECH
// ===========================================================================

liSlide(
  pres,
  "I am preparing a short presentation as my chosen influential person.",
  [
    "I can prepare a short speech as my chosen influential person.",
    "I can write and explain about this person's life, family and important events.",
    "I can use a Bold Beginning and Excellent Ending to grip my audience and leave them thinking.",
  ],
  notes({
    say: [
      "Now that you have researched, you write your speech.",
      "Here is the twist: you write it as your person, speaking in first person - as 'I'.",
      "Read our three 'I can' statements. A Bold Beginning and Excellent Ending are key.",
    ],
    do: [
      "Read the learning intention and success criteria aloud.",
      "Share the Speech Rubric now so students know how the speech is judged.",
    ],
    teacherNotes: {
      text: "Depending on your class, briefly clarify what first person means before students write. The speech becomes the prop-supported presentation for Celebration Night.",
    },
    watchFor: [
      "Students writing 'he' or 'she' - remind them they are the person, so it is 'I'.",
    ],
    tag: "[Shaping Australia | Step 2 | Goal]",
  }),
  stepFoot(2, "Write Your Speech")
);

// Speak AS your person - first person frame
infoSlide({
  topColor: C.SECONDARY,
  badge: "Your Voice", badgeColor: C.SECONDARY, badgeW: 1.9,
  title: "Speak as your person",
  footer: stepFoot(2, "Write Your Speech"),
  draw: (s) => {
    const cardY = CONTENT_TOP, cardH = 2.35;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Write in first person. You ARE your person, telling your own story.", {
      x: 0.78, y: cardY + 0.16, w: 8.4, h: 0.3,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    // First-person frame (placeholders in ACCENT so students see what to fill in)
    const frameRuns = [
      { text: "Hello, I am ", c: C.CHARCOAL }, { text: "[your person]", c: C.ACCENT, b: true }, { text: ".", c: C.CHARCOAL, brk: true },
      { text: "You might know me for ", c: C.CHARCOAL }, { text: "[what you did]", c: C.ACCENT, b: true }, { text: ".", c: C.CHARCOAL, brk: true },
      { text: "I helped shape Australia by ", c: C.CHARCOAL }, { text: "[your impact]", c: C.ACCENT, b: true }, { text: "...", c: C.CHARCOAL, brk: true },
      { text: "...and this happened around ", c: C.CHARCOAL }, { text: "[when]", c: C.ACCENT, b: true }, { text: ".", c: C.CHARCOAL },
    ];
    s.addText(frameRuns.map((r) => ({
      text: r.text, options: { fontSize: 21, fontFace: FONT_H, color: r.c, bold: Boolean(r.b), breakLine: Boolean(r.brk) },
    })), {
      x: 0.78, y: cardY + 0.56, w: 8.4, h: cardH - 0.7,
      valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
    });

    // Bottom reminders
    const by = cardY + cardH + 0.16, bh = 0.55, bw = 4.35;
    addTextOnShape(s, "Keep it to 3-5 minutes", {
      x: 0.5, y: by, w: bw, h: bh, rectRadius: 0.08, fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true });
    addTextOnShape(s, "Write 'I', not 'he' or 'she'", {
      x: 5.15, y: by, w: bw, h: bh, rectRadius: 0.08, fill: { color: C.ACCENT },
    }, { fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true });
  },
  notesText: notes({
    say: [
      "Your speech is told in first person. You step into your person's shoes and say 'I'.",
      "Here is a sentence frame to get you started - just fill in the gaps with your research.",
      "Keep the whole speech to about three to five minutes.",
    ],
    do: [
      "Model the frame aloud with one example person so students hear the first-person voice.",
      "Have students draft an opening line using the frame.",
    ],
    teacherNotes: {
      text: "The frame is a starter, not a script - students replace the brackets with their own researched facts. Do not invent facts for them.",
    },
    enabling: [
      "ENABLING: let students complete the sentence frame as their first paragraph.",
      "EXTENDING: ask students to rewrite the opening without the frame, in a stronger Bold Beginning.",
    ],
    watchFor: [
      "Slipping into third person ('she was born...') - prompt back to 'I was born...'.",
    ],
    tag: "[Shaping Australia | Step 2 | First person voice]",
  }),
});

// Plan your speech - 3 parts (Speech Rubric)
infoSlide({
  badge: "Speech Plan", badgeColor: C.PRIMARY, badgeW: 2.0,
  title: "Plan your speech: beginning, body, ending",
  footer: stepFoot(2, "Write Your Speech"),
  draw: (s) => {
    const x = 0.5, w = 9, gap = 0.15;
    // Introduction band
    let y = CONTENT_TOP, h = 0.85;
    addCard(s, x, y, w, h, { strip: C.SUCCESS, fill: C.WHITE });
    s.addText("Introduction - Bold Beginning", {
      x: x + 0.25, y: y + 0.1, w: w - 0.5, h: 0.3,
      fontSize: 15, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0,
    });
    s.addText("Grab us first, then tell us who you are and what your speech is about.", {
      x: x + 0.25, y: y + 0.44, w: w - 0.5, h: 0.34,
      fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "top", fit: "shrink", shrinkText: true,
    });

    // Body band (taller)
    y = y + h + gap; h = 1.7;
    addCard(s, x, y, w, h, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Body - share the facts", {
      x: x + 0.25, y: y + 0.1, w: w - 0.5, h: 0.3,
      fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Childhood  |  Family life  |  Obstacles and success  |  Historical events  |  Positive contributions to Australia  |  An interesting fact  |  A quote from or about you", {
      x: x + 0.25, y: y + 0.46, w: w - 0.5, h: 0.82,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "top", fit: "shrink", shrinkText: true,
    });
    s.addText("Aim to cover 6-7 of these to really inform your audience.", {
      x: x + 0.25, y: y + 1.30, w: w - 0.5, h: 0.3,
      fontSize: 12.5, fontFace: FONT_B, color: C.ACCENT, bold: true, italic: true, margin: 0,
    });

    // Conclusion band
    y = y + h + gap; h = 0.85;
    addCard(s, x, y, w, h, { strip: C.ACCENT, fill: C.WHITE });
    s.addText("Conclusion - Excellent Ending", {
      x: x + 0.25, y: y + 0.1, w: w - 0.5, h: 0.3,
      fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText("Finish with a memorable quote or strong line that leaves us thinking. Sum up your main points.", {
      x: x + 0.25, y: y + 0.44, w: w - 0.5, h: 0.34,
      fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "top", fit: "shrink", shrinkText: true,
    });
  },
  notesText: notes({
    say: [
      "Every speech has three parts. Look at the plan: a Bold Beginning, a Body full of facts, and an Excellent Ending.",
      "In the Body, choose from these topics - try to cover six or seven so we really get to know your person.",
      "Your ending should leave us thinking, maybe with a quote.",
    ],
    do: [
      "Walk through the Speech Rubric so students see these are the things being marked.",
      "Have students sort their researched facts into the Body topics.",
    ],
    teacherNotes: {
      text: "These three parts and the Body topic list come straight from the Speech Rubric. Covering 6-7 Body topics is the rubric's top 'WOW' band; 3-5 is solid.",
    },
    enabling: [
      "ENABLING: ask for at least three Body topics, well explained.",
      "EXTENDING: ask students to weave a quote from or about their person into the ending.",
    ],
    watchFor: [
      "Speeches that are all facts and no Bold Beginning or Excellent Ending - both are on the rubric.",
    ],
    sources: [
      "Speech Rubric (teacher-provided): Introduction, Body topic ideas, Conclusion, and Body Language & Speaking Skills.",
    ],
    tag: "[Shaping Australia | Step 2 | Speech structure]",
  }),
});

// ===========================================================================
//  STEP 3 - PRACTISE PERFORMING
// ===========================================================================

liSlide(
  pres,
  "I am learning to perform my presentation as my chosen influential person.",
  [
    "I can practise delivering a speech to a partner as my chosen influential person.",
    "I can experiment with voice, body language and gestures to improve my presentation.",
    "I can speak clearly and not read straight off my speech at times.",
  ],
  notes({
    say: [
      "Step 3 is all about performing. A great speech is not just read - it is performed.",
      "You will practise with a partner and try out voice, body and gestures.",
      "Read our three 'I can' statements.",
    ],
    do: [
      "Read the learning intention and success criteria aloud.",
      "Set up partners for rehearsal.",
    ],
    teacherNotes: {
      text: "This step turns the written speech into a performance. The next slides give the success look and the practice routine.",
    },
    watchFor: [
      "Students who want to read every word - the goal is to look up and perform at times.",
    ],
    tag: "[Shaping Australia | Step 3 | Goal]",
  }),
  stepFoot(3, "Practise Performing")
);

// Great presenters: voice, body, face
infoSlide({
  topColor: C.SECONDARY,
  badge: "Perform", badgeColor: C.SECONDARY, badgeW: 1.7,
  title: "Great presenters use voice, body and face",
  footer: stepFoot(3, "Practise Performing"),
  draw: (s) => {
    featureCards(s, [
      { tag: "Voice", color: C.PRIMARY, bullet: true, lines: ["Speak clearly", "Good volume", "Change your pace and tone"] },
      { tag: "Body", color: C.SECONDARY, bullet: true, lines: ["Stand tall", "Use gestures and movement", "Face your audience"] },
      { tag: "Face", color: C.ACCENT, bullet: true, lines: ["Make eye contact", "Show expression", "Look up from your cards"] },
    ], { y: CONTENT_TOP, h: 2.8, gap: 0.25, tagSize: 16, lineSize: 13.5 });

    addTextOnShape(s, "Borrow how your person might have spoken and moved - bring them to life.", {
      x: 0.5, y: CONTENT_TOP + 2.95, w: 9, h: 0.5, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
  },
  notesText: notes({
    say: [
      "Let's brainstorm: what does a great presenter do with their voice, body and face?",
      "Look at the three cards. These are exactly what the rubric looks for.",
      "Think about how your person might have spoken and moved - you can borrow that.",
    ],
    do: [
      "Brainstorm good-presenter features as a class, then match them to the three cards.",
      "Demonstrate one strong and one weak example so the difference is clear.",
    ],
    teacherNotes: {
      text: "These three areas come from the Speech Rubric's Body Language and Speaking Skills row. Keep it active - have students try each one.",
    },
    watchFor: [
      "Quiet, flat delivery - coach volume and tone.",
      "Stillness with hands by sides - coach one or two purposeful gestures.",
    ],
    sources: [
      "Speech Rubric (teacher-provided): Body Language and Speaking Skills - eye contact, volume, pace, tone, gestures and movement.",
    ],
    tag: "[Shaping Australia | Step 3 | Presenter skills]",
  }),
});

// Get into character & practise
infoSlide({
  topColor: C.SECONDARY,
  badge: "Rehearse", badgeColor: C.SECONDARY, badgeW: 1.8,
  title: "Get into character and practise",
  footer: stepFoot(3, "Practise Performing"),
  draw: (s) => {
    featureCards(s, [
      { tag: "Watch", color: C.PRIMARY, lines: ["Watch the Getting Into Character video for ideas."] },
      { tag: "Cue cards", color: C.SECONDARY, lines: ["Put key words on cue cards - not every word."] },
      { tag: "Practise", color: C.SUCCESS, lines: ["Say it out loud with a partner, in character."] },
    ], { y: CONTENT_TOP, h: 2.8, gap: 0.25, tagSize: 16, lineSize: 13 });

    addTextOnShape(s, "Look up from your cards. Don't read every word - perform as your person.", {
      x: 0.5, y: CONTENT_TOP + 2.95, w: 9, h: 0.5, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
  },
  notesText: notes({
    say: [
      "Now let's get into character. We'll watch a short video on how to step into a role.",
      "Move your speech onto cue cards - just key words so you can look up and perform.",
      "Then practise out loud with your partner, in character.",
    ],
    do: [
      "Play the Getting Into Character video (teacher-provided link).",
      "Have students transfer their speech to cue cards, then rehearse in pairs.",
      "Circulate and give quick voice, body and eye-contact feedback.",
    ],
    teacherNotes: {
      text: "Cue cards stop students reading word for word. Partner rehearsal builds confidence before delivering to the class.",
    },
    sensitivity: [
      "What it is: getting into character as a real person, who may be a First Nations figure or from sensitive history.",
      "Framing language: portray with respect and dignity - we are honouring them, not impersonating or mocking.",
      "Watch for: caricature, exaggerated accents, or costumes that stereotype - guide students before they rehearse.",
      "Protocol: check the portrayal plan with any student presenting a culturally significant figure.",
    ],
    watchFor: [
      "Cue cards with full sentences - prompt students to reduce to key words.",
      "Eyes glued to cards - coach looking up at the audience.",
    ],
    sources: [
      "Getting Into Character video (teacher-provided link).",
    ],
    tag: "[Shaping Australia | Step 3 | Get into character]",
  }),
});

// ===========================================================================
//  STEP 4 - DELIVER
// ===========================================================================

liSlide(
  pres,
  "I am delivering my presentation on my chosen influential person to my class or small group.",
  [
    "I can deliver a speech as my chosen influential person.",
    "I can use my voice, body language and gestures to improve my presentation.",
    "I can speak clearly and not read straight off my speech.",
  ],
  notes({
    say: [
      "It's presentation time. You deliver your speech as your person to the class or a small group.",
      "Everything we practised counts now: voice, body, eye contact, and your Bold Beginning and Excellent Ending.",
      "Read our three 'I can' statements.",
    ],
    do: [
      "Read the learning intention and success criteria aloud.",
      "Set up the presentation order and the audience routine.",
    ],
    teacherNotes: {
      text: "Assess each presentation with the Speech Rubric shared earlier. A small group works well for students who find the whole class daunting.",
    },
    watchFor: [
      "Nerves - remind students they are well prepared and in character.",
    ],
    tag: "[Shaping Australia | Step 4 | Goal]",
  }),
  stepFoot(4, "Deliver Your Speech")
);

// Presentation day + PMI feedback
infoSlide({
  topColor: C.PRIMARY,
  badge: "Presentation Day", badgeColor: C.PRIMARY, badgeW: 2.5,
  title: "Presentation day and helpful feedback",
  footer: stepFoot(4, "Deliver Your Speech"),
  draw: (s) => {
    // Top reminder card
    const ty = CONTENT_TOP, th = 1.3;
    addCard(s, 0.5, ty, 9, th, { strip: C.ALERT, fill: C.WHITE });
    s.addText("On presentation day", {
      x: 0.78, y: ty + 0.12, w: 8.4, h: 0.3,
      fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
    });
    s.addText("Present as your person. You'll be marked on your Speech Rubric: a Bold Beginning, the facts in your Body, an Excellent Ending, and your voice, body and eye contact.", {
      x: 0.78, y: ty + 0.46, w: 8.4, h: 0.72,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "top", fit: "shrink", shrinkText: true,
    });

    // PMI feedback label + cards
    const py = ty + th + 0.14;
    s.addText("Audience feedback - PMI", {
      x: 0.5, y: py, w: 9, h: 0.26,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    featureCards(s, [
      { tag: "Plus  (+)", color: C.SUCCESS, lines: ["Something they did well."] },
      { tag: "Minus  (-)", color: C.ALERT, lines: ["One thing to improve."] },
      { tag: "Interesting  (?)", color: C.ACCENT, lines: ["Something that made you think."] },
    ], { y: py + 0.32, h: 1.55, gap: 0.25, tagSize: 15, lineSize: 13 });
  },
  notesText: notes({
    say: [
      "When you present, stay in character and use everything you practised.",
      "I'll be marking with the Speech Rubric, so aim for a Bold Beginning, strong facts and an Excellent Ending.",
      "If we have time, the audience gives PMI feedback: a Plus, a Minus, and something Interesting.",
    ],
    do: [
      "Assess each presentation against the Speech Rubric.",
      "If time permits, ask audience members for one PMI point each.",
    ],
    teacherNotes: {
      text: "PMI keeps peer feedback balanced and kind. Model one PMI response before students give their own.",
    },
    enabling: [
      "ENABLING: let nervous students present to a small group rather than the whole class.",
      "EXTENDING: ask confident presenters to take an in-character question after their speech.",
    ],
    watchFor: [
      "Feedback that is only praise or only criticism - PMI balances both, plus a thinking point.",
    ],
    sources: [
      "Speech Rubric (teacher-provided) - used for assessment.",
    ],
    tag: "[Shaping Australia | Step 4 | Deliver & feedback]",
  }),
});

// ===========================================================================
//  STEP 5 - CELEBRATION NIGHT QUESTIONS
// ===========================================================================

liSlide(
  pres,
  "I am writing questions for visitors to ask me on our Celebration Night.",
  [
    "I can come up with a range of questions about different topics of my person's life.",
    "I can learn and remember the answers to these questions and confidently answer them when asked.",
    "I can portray the influential person I am presenting.",
  ],
  notes({
    say: [
      "On Celebration Night, guests will visit and ask you about your person.",
      "So you will write the questions they might ask - and learn your answers.",
      "Read our three 'I can' statements.",
    ],
    do: [
      "Read the learning intention and success criteria aloud.",
      "Hand out or open the Example Questions sheet.",
    ],
    teacherNotes: {
      text: "Students stay in character to answer. Print question cards and clip them on a ring (Bloom's key rings) to rehearse before the night.",
    },
    watchFor: [
      "Questions that are too narrow - aim for a range across the person's life.",
    ],
    tag: "[Shaping Australia | Step 5 | Goal]",
  }),
  stepFoot(5, "Celebration Night Questions")
);

// Celebration Night questions - Example Questions (verbatim)
infoSlide({
  topColor: C.ACCENT,
  badge: "Celebration Night", badgeColor: C.ACCENT, badgeW: 2.6,
  title: "Get ready for Celebration Night",
  footer: stepFoot(5, "Celebration Night Questions"),
  draw: (s) => {
    const cardY = CONTENT_TOP, cardH = 2.85;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.ACCENT, fill: C.WHITE });
    s.addText("Visitors might ask you...", {
      x: 0.78, y: cardY + 0.14, w: 8.4, h: 0.3,
      fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText([
      "Who is my person and why did I pick them?",
      "What were the thing/s my person did to shape Australia?",
      "What is my person best known for?",
      "Where did my person create change?",
      "What was the most interesting thing I learnt about my person?",
    ].map((t, i, a) => ({
      text: t, options: { bullet: true, breakLine: i < a.length - 1, fontSize: 15, color: C.CHARCOAL, paraSpaceAfter: 7 },
    })), {
      x: 0.85, y: cardY + 0.52, w: 8.3, h: cardH - 0.66,
      fontFace: FONT_B, valign: "top", margin: 0, fit: "shrink", shrinkText: true,
    });

    addTextOnShape(s, "Add your own questions. Learn the answers, clip your cards on a ring, and rehearse.", {
      x: 0.5, y: cardY + cardH + 0.16, w: 9, h: 0.55, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14.5, fontFace: FONT_B, color: C.WHITE, bold: true });
  },
  notesText: notes({
    say: [
      "Here are some questions visitors often ask. Start with these, then add your own.",
      "Write each question and its answer, and learn them so you can reply confidently - in character.",
      "Clip your question cards on a ring so you can rehearse anywhere.",
    ],
    do: [
      "Share the Example Questions sheet; students copy these and add more.",
      "Have students write answers, then test each other in pairs.",
    ],
    teacherNotes: {
      text: "The five starter questions are taken directly from your Example Questions sheet. Students add their own to make a fuller set, then rehearse staying in role.",
    },
    enabling: [
      "ENABLING: students answer the five starter questions first before adding their own.",
      "EXTENDING: students prepare for a tricky or unexpected question and a thoughtful in-character reply.",
    ],
    watchFor: [
      "Answers students cannot yet recall - prompt more rehearsal before the night.",
    ],
    sources: [
      "Example Questions sheet (teacher-provided): the five starter questions are reproduced exactly.",
    ],
    tag: "[Shaping Australia | Step 5 | Celebration Night questions]",
  }),
});

// ===========================================================================
//  PROJECT CHECKLIST  (self-check across all five steps)
// ===========================================================================

infoSlide({
  topColor: C.SUCCESS,
  badge: "Ready Check", badgeColor: C.SUCCESS, badgeW: 2.0,
  title: "Are you ready? Your project checklist",
  footer: FOOT_ALL,
  draw: (s) => {
    const items = [
      "I have chosen an influential person, not just a popular one.",
      "I researched how and when they helped shape Australia.",
      "I have made my poster, diorama, slideshow or recording.",
      "I wrote my speech in first person, with a Bold Beginning and Excellent Ending.",
      "I practised in character, using voice, body and eye contact.",
      "I have my questions and answers ready for Celebration Night.",
    ];
    const rowH = 0.58, y0 = 1.45, boxD = 0.30;
    items.forEach((t, i) => {
      const ry = y0 + i * rowH;
      s.addShape("roundRect", {
        x: 0.7, y: ry, w: boxD, h: boxD, rectRadius: 0.05,
        fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.75 },
      });
      s.addText(t, {
        x: 0.7 + boxD + 0.22, y: ry - 0.06, w: 8.2, h: boxD + 0.12,
        fontSize: 14.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });
  },
  notesText: notes({
    say: [
      "Before Celebration Night, use this checklist to see if you are ready.",
      "Tick each box you can honestly say yes to.",
      "Any boxes you can't tick yet tell you exactly what to work on.",
    ],
    do: [
      "Have students self-check against the six items.",
      "Note who needs support on which step, and plan a quick top-up.",
    ],
    teacherNotes: {
      text: "This pulls the whole project together. It doubles as a readiness check before the event.",
    },
    watchFor: [
      "Students ticking boxes optimistically - ask for the evidence (the cue cards, the display, the answers).",
    ],
    tag: "[Shaping Australia | Checklist | Readiness]",
  }),
});

// ===========================================================================
//  CLOSING REFLECTION
// ===========================================================================

closingSlide(
  pres,
  {
    reflectionPrompt: "Which part of becoming your person are you most proud of, and what will you practise before Celebration Night?",
    scItems: [
      "I can research and choose a person who truly helped shape Australia.",
      "I can present as my person with a clear speech and confident voice and body.",
      "I can answer questions about my person on Celebration Night.",
    ],
    selfAssessment: {
      prompt: "How ready do you feel to present as your person?",
      options: ["Ready to present", "Nearly there", "Need more practice"],
    },
  },
  notes({
    say: [
      "Look back at everything you have done: chosen, researched, written, practised and prepared.",
      "You are bringing a piece of Australia's story to life.",
      "Tell me with a thumb how ready you feel to present as your person.",
    ],
    do: [
      "Run a quick thumbs self-check against the three 'I can' statements.",
      "Note who still needs practice before Celebration Night.",
    ],
    teacherNotes: {
      text: "Use the self-check to decide who needs a short rehearsal top-up. Acknowledge progress - this builds the confidence students need to perform.",
    },
    tag: "[Shaping Australia | Closing | Reflection]",
  })
);

// ===========================================================================
//  WRITE
// ===========================================================================

(async () => {
  const outFile = path.join(OUT_DIR, "Shaping Australia - Influential People Project.pptx");
  await pres.writeFile({ fileName: outFile });
  console.log(`PPTX written to ${outFile}`);
})();
