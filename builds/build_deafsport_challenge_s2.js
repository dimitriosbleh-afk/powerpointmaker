"use strict";

/**
 * Deaf Sports in Australia | Challenge (Years 3 and 4) | Lesson 2 Meeting a signer
 *
 * Built from the unit document section 9.2. Written as a build script rather
 * than a lesson spec because the spec pipeline cannot express an animated sign
 * card, a Signbank lookup card, the fixed six-cue strip or a countdown GIF.
 * The shared parts live in builds/auslan_lib.js so the next unit adds
 * vocabulary, not drawing code.
 *
 * Challenge calibration (unit document 3 and 9.2): one English sentence in the
 * Do Now, ask two people in the Mix and Mingle with the sport card on the desk,
 * and no cold-call follow-up after the boards.
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme, weekToVariant } = require("../themes/factory");
const P = require("../themes/pdf_helpers");
const A = require("./auslan_lib");

const T = createTheme("literacy", "grade34", weekToVariant(1));
const C = T.C;
const FOOTER = "Deaf Sports in Australia  |  Session 2 of 8  |  Years 3-4 Auslan";
const OUT_DIR = "DeafSport_Challenge_S2_Meeting_A_Signer";
const RES_DIR = path.join("output", OUT_DIR, "resources-session2");
const CARD_PDF = "Session 2 Sport Card.pdf";

// Sign It! pages from Chris's sheet, 3 Oct 2026. Blank means not in the book.
const SIGNIT_PAGE = {
  swimming: 57, basketball: 58, tennis: 60, football: 58, netball: 59, rugby: 60,
};

const SPORTS = [
  "swimming", "athletics", "futsal", "basketball", "tennis", "football",
  "netball", "chess", "golf", "rugby", "table tennis", "lawn bowls",
];

const report = A.createSignReport();

/** Session 2 Sport Card: the twelve sports and their pages, two cards a page. */
function buildSportCardPdf() {
  fs.mkdirSync(RES_DIR, { recursive: true });
  const doc = P.createPdf({ title: "Session 2 Sport Card" });
  let y = P.addPdfHeader(doc, "Session 2 Sport Card", {
    color: C.PRIMARY,
    subtitle: "One card per pair, face up on the desk. Two cards a page: cut across the middle.",
    lessonInfo: "Session 2  |  Mix and Mingle  |  point to the sport, then sign it",
    showNameDate: false,
  });
  y += 14;  // clear of the info line, which the dashed cut line otherwise touches
  const cols = 4;
  const gap = 8;
  const cw = (P.PAGE.CONTENT_W - gap * (cols - 1)) / cols;
  const ch = 58;
  const cardH = 26 + 3 * (ch + gap) + 10;
  [0, 1].forEach((k) => {
    const top = y + k * (cardH + 24);
    doc.save();
    doc.dash(4, { space: 4 }).roundedRect(P.PAGE.MARGIN - 6, top - 6, P.PAGE.CONTENT_W + 12, cardH, 8)
      .lineWidth(1).strokeColor("#9CA3AF").stroke();
    doc.undash();
    doc.restore();
    doc.fontSize(13).font("Sans-Bold").fillColor("#111827")
      .text("Twelve sports", P.PAGE.MARGIN, top + 2, { width: P.PAGE.CONTENT_W });
    SPORTS.forEach((word, i) => {
      const x = P.PAGE.MARGIN + (i % cols) * (cw + gap);
      const by = top + 26 + Math.floor(i / cols) * (ch + gap);
      doc.roundedRect(x, by, cw, ch, 5).lineWidth(1).strokeColor("#374151").stroke();
      doc.fontSize(15).font("Sans-Bold").fillColor("#111827")
        .text(word, x + 4, by + 12, { width: cw - 8, align: "center" });
      if (SIGNIT_PAGE[word]) {
        doc.fontSize(10).font("Sans").fillColor("#4B5563")
          .text("Sign It! page " + SIGNIT_PAGE[word], x + 4, by + 36, { width: cw - 8, align: "center" });
      }
    });
  });
  P.writePdf(doc, path.join(RES_DIR, CARD_PDF), "Session 2  |  Sport Card  |  Years 3-4 Auslan");
}

function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  T.titleSlide(
    pres,
    "Meeting a signer",
    "Deaf Sports in Australia",
    "Lesson 2 of 8  |  Years 3 and 4  |  Term 4",
    "Lesson 2 of 8. First lesson with new vocabulary in it. Lesson 1 taught the looking-up routine this one runs on."
  );

  // 2. Teacher Resources
  P.addResourceSlide(pres, [
    {
      name: "Session 2 Sport Card",
      fileName: "resources-session2/" + CARD_PDF,
      note: "One per pair, face up for the Mix and Mingle. Two to a page.",
    },
  ], T, FOOTER, T.composeGlanceNotes({
    beats: [
      ["SHOW while students settle.", "SAY: Journals open on the front page, the tracker."],
      "COLLECT the sport cards at the end; they come back every week.",
    ],
    prep: [
      "Print and cut before the lesson. Clear a floor space; timer visible.",
      A.ATTRIBUTION,
    ],
    tag: "[Setup | Planning | HITS 2]",
  }));

  // 3. Do Now
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "Do Now", { color: C.PRIMARY });
    T.addCueStrip(s, ["voicesOff", "whiteboards"]);
    T.addTitle(s, "Same as last week: write it, then find the signs");
    T.addCard(s, 0.5, 1.45, 9, 1.5, { variant: "tint", tone: C.PRIMARY });
    s.addText("What sport do you like?", {
      x: 0.7, y: 1.45, w: 8.6, h: 1.5, fontSize: 40, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addInstructionCard(s, [
      { role: "header", text: "On your whiteboard" },
      { text: "Write your answer in English. One sentence." },
      { text: "Underline every word you would need a sign for." },
      { text: "Tick the ones you already know on your tracker." },
    ], { x: 0.5, y: 3.15, w: 9, h: 1.55, strip: C.PRIMARY });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        "SILENT from the door. Say nothing; Lesson 1 taught this.",
        ["WATCH for the underlining, not the writing.", "That is the step that was taught."],
        "TIME: 5 minutes. Deal with whatever walked in from the playground.",
        ["SAY: You know how to find a sign now.", "Today you learn twelve of them."],
      ],
      prep: [
        "Seated and silent every week from here. It is the settling time.",
        "If the underlining is shaky, reteach the choosing step for two minutes.",
      ],
      tag: "[Do Now | Attention, focus and regulation | HITS 6]",
    }));
  }

  // 4. LI and SC
  T.liSlide(
    pres,
    "I am learning to open a signed exchange and ask one question about sport.",
    [
      "I can wave or tap to get someone's attention and wait for their eyes.",
      "I can ask one person about their favourite sport and watch their whole answer.",
      "I can tell the class one thing I found out.",
    ],
    T.composeGlanceNotes({
      beats: [
        ["POINT to each criterion.", "SAY: Three things, and the first is just waiting for eyes."],
        "SAY: By the end, everyone has asked somebody what sport they like best.",
      ],
      prep: "Criterion 1 is reachable by everyone. Criterion 2 is what the exit rotation collects.",
      tag: "[LI and SC | Planning | HITS 1]",
    }),
    FOOTER,
    { numberSC: true, strongHeadings: true, separate: true }
  );

  // 5. I Do: the opening bank
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Twelve sports for the whole term");
    const cols = 4;
    const cw = 9 / cols;
    SPORTS.forEach((word, i) => {
      const r = Math.floor(i / cols);
      const cx = 0.5 + (i % cols) * cw;
      T.addTextOnShape(
        s,
        word,
        { x: cx + 0.08, y: 1.5 + r * 0.72, w: cw - 0.16, h: 0.56, rectRadius: 0.1,
          fill: { color: C.BG_LIGHT }, line: { color: C.MUTED, width: 0.75 } },
        { fontSize: 17, fontFace: T.FONT_B, bold: true, color: C.CHARCOAL,
          align: "center", valign: "middle", margin: 0 }
      );
    });
    s.addText("These come back every week. They are the front page of your journal.", {
      x: 0.5, y: 4.2, w: 9, h: 0.5, fontSize: 16, fontFace: T.FONT_B,
      color: C.MUTED, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        "MODEL all twelve, once each, straight through. Nobody produces them today.",
        ["SAY: In July this year the Australian Deaf Games were on at the Sunshine Coast.",
          "More than one thousand three hundred Deaf people competed, across twenty sports.",
          "That is this year, not history."],
        "POINT at the journal tracker. Every sign this term is on that page.",
      ],
      prep: [
        "Receptive only. They count for nothing in the production budget.",
        "Swap any sport for one your classes actually play; keep the count at twelve.",
      ],
      sources: "Deaf Connect, Australian Deaf Games, deafconnect.org.au",
      tag: "[I Do | Explicit teaching | HITS 3]",
    }));
  }

  // 6. I Do: the anchor, four moves
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Eyes up. Watch. Copy.");
    T.addTextOnShape(
      s,
      "Eyes up. Watch. Copy.",
      { x: 0.5, y: 1.4, w: 9, h: 0.72, rectRadius: 0.12, fill: { color: C.PRIMARY } },
      { fontSize: 26, fontFace: T.FONT_H, bold: true, color: C.WHITE,
        align: "center", valign: "middle", margin: 0 }
    );
    const moves = [
      ["Eyes up", "Look at me before I start"],
      ["Watch", "The whole sign, not the first bit"],
      ["Copy", "Once, properly"],
      ["Help", "Tap or wave, then point"],
    ];
    const mw = (9 - 0.22 * 3) / 4;
    moves.forEach(([head, sub], i) => {
      const mx = 0.5 + i * (mw + 0.22);
      T.addCard(s, mx, 2.45, mw, 1.75, { variant: "tint", tone: C.PRIMARY });
      s.addText(String(i + 1), {
        x: mx, y: 2.58, w: mw, h: 0.4, fontSize: 15, fontFace: T.FONT_B, bold: true,
        color: C.PRIMARY, align: "center", valign: "middle", margin: 0,
      });
      s.addText(head, {
        x: mx + 0.06, y: 2.98, w: mw - 0.12, h: 0.5, fontSize: 21, fontFace: T.FONT_H, bold: true,
        color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
      });
      s.addText(sub, {
        x: mx + 0.08, y: 3.5, w: mw - 0.16, h: 0.6, fontSize: 14, fontFace: T.FONT_B,
        color: C.MUTED, align: "center", valign: "top", margin: 0,
      });
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        ["SIGN each of the four and have them copy it back.", "These are classroom signs; they are curriculum."],
        ["SAY: When I sign eyes up, that is your cue.", "Nobody starts until every pair of eyes is on me."],
        ["SAY: If your partner missed it, do not say their name.", "Tap them or wave, then point at me."],
        ["GET IT WRONG: start signing at a room half looking away.", "SAY: Half of you missed that, so I wasted it."],
      ],
      trap: ["telling a neighbour out loud to look up.", "Fix: tap, then point, student redoes it."],
      prep: [
        "The anchor, restated in these exact words in every I Do. Do not reword it.",
        "Tapping and pointing is the Deaf-cultural move, not a workaround for talking.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3]",
    }));
  }

  // 7. I Do: three signs
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Three signs. Watch me, then copy once");
    A.addSignCardRow(T, s, [
      { gloss: "DEAF", meaning: "deaf" },
      { gloss: "SPORT", meaning: "sport" },
      { gloss: "FAVOURITE", meaning: "favourite" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "DEAF, SPORT, FAVOURITE",
      beats: [
        ["MODEL each sign twice, slowly then at pace.", "They copy once. The repetitions come from using it."],
        ["SAY: Deaf with a big D is not a description of ears.",
          "It is the word for a community and the language that holds it together."],
        "CIRCULATE: look for handshapes copied from you, not from a neighbour.",
      ],
      trap: ["a handshape copied from the person beside them.", "Fix: re-model facing them, student redoes it."],
      stretch: "add the sport you play to the exchange from the first turn.",
      help: "three sports circled on their sport card: the answer becomes a point plus a sign.",
      prep: [
        "3 of today's 5 production signs. Lesson 1 is buying routines, not vocabulary.",
        A.REHEARSE_FIRST.FAVOURITE,
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 8. I Do: two more signs
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Two more. One of them gets you out of trouble.");
    A.addSignCardRow(T, s, [
      { gloss: "TEAM", meaning: "team" },
      { gloss: "AGAIN", meaning: "again: show me that again" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "TEAM, AGAIN",
      beats: [
        ["MODEL each twice; they copy once.", "Then use AGAIN live: miss something and ask for it back."],
        ["SAY: I watch the whole answer. Not the first bit,",
          "then off thinking about my next question."],
        ["SAY: If I miss it, I do not guess. I ask again.",
          "That costs four seconds and nothing else."],
      ],
      trap: ["guessing an answer they did not catch.", "Fix: freeze the pair, student signs AGAIN and watches it again."],
      stretch: "use AGAIN once in the You Do without being told to.",
      help: "you sign the answer a second time unprompted while they practise the wait.",
      prep: [
        "AGAIN is the whole repair set this week. WHAT MEAN? arrives in Lesson 2.",
        A.REHEARSE_FIRST.AGAIN,
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 9. We Do: find your three sports and practise them
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.SUCCESS);
    T.addBadge(s, "We Do", { color: C.SUCCESS });
    T.addCueStrip(s, ["voicesOff", "partner", "timer"]);
    T.addTitle(s, "Find your three sports and practise them");
    T.addInstructionCard(s, [
      { role: "header", text: "With your partner" },
      { text: "Pick the three sports you are most likely to be asked about." },
      { text: "Find each one in Sign It! and write the page number." },
      { text: "Practise each one until you can do it without the book." },
      { text: "Check each other. Same handshape? Same place?" },
    ], { x: 0.5, y: 1.4, w: 5.6, h: 2.35, strip: C.SUCCESS });
    A.addSignItCover(T, s, { x: 6.3, y: 1.4, w: 1.8, h: 2.35 });
    A.addCountdown(T, s, 480, { x: 8.3, y: 1.9, w: 1.2 });
    s.addText("8 minutes", {
      x: 8.0, y: 3.15, w: 1.8, h: 0.4, fontSize: 15, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    s.addText(
      "You cannot ask somebody their favourite sport until you can sign a few. This is that time.",
      {
        x: 0.5, y: 3.95, w: 9, h: 0.5, fontSize: 17, fontFace: T.FONT_B, bold: true,
        color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
      }
    );
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "three sport signs each, found and practised without the book",
      beats: [
        ["SET UP: Sign It! one between two, boards out.", "The GIFs stay on screen; go back a slide if they want them."],
        "TIME: advance to start the clock. Eight minutes.",
        ["CIRCULATE: watch the checking, not the looking up.", "A pair who never check will both learn it wrong."],
        "COLLECT: each student can sign three sports with the book shut.",
      ],
      trap: ["copying the picture without checking each other.", "Fix: stop the pair, one signs, one checks, swap."],
      stretch: "find a sport nobody else picked and teach it to another pair.",
      help: "three sports chosen for them, with the page numbers written down.",
      prep: [
        "10 min. This replaced a game, because they cannot play it without the signs.",
        "The GIFs are in Google Classroom, so this can continue at home.",
      ],
      tag: "[We Do | Collaborative learning | HITS 5]",
    }));
  }

  // 10. Primary decision point
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.ASSESS);
    T.addBadge(s, "Check it", { color: C.ASSESS });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "Which sport am I signing?");
    T.addCard(s, 0.5, 1.5, 9, 1.4, { variant: "tint", tone: C.ASSESS });
    s.addText("I sign three sports. You write which sport each one is.", {
      x: 0.7, y: 1.5, w: 8.6, h: 1.4, fontSize: 30, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    const steps = ["Watch all of it", "Write the sport", "Hover, then chin it", "One of you signs it back"];
    const stw = (9 - 0.18 * 3) / 4;
    steps.forEach((step, i) => {
      T.addCard(s, 0.5 + i * (stw + 0.18), 3.08, stw, 1.05, { variant: "outline", tone: C.ASSESS });
      s.addText(String(i + 1), {
        x: 0.5 + i * (stw + 0.18), y: 3.16, w: stw, h: 0.34,
        fontSize: 13, fontFace: T.FONT_B, bold: true, color: C.ASSESS,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(step, {
        x: 0.56 + i * (stw + 0.18), y: 3.5, w: stw - 0.12, h: 0.55,
        fontSize: 15, fontFace: T.FONT_B, bold: true, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0,
      });
    });
    T.addTextOnShape(
      s,
      "Write it. Hover it. Chin it on three. One, two, three.",
      { x: 1.6, y: 4.3, w: 6.8, h: 0.62, rectRadius: 0.12, fill: { color: C.ASSESS } },
      { fontSize: 18, fontFace: T.FONT_B, bold: true, color: C.WHITE,
        align: "center", valign: "middle", margin: 0 }
    );
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "the three sports you signed",
      beats: [
        ["ASK: which sport is this. Sign three, one at a time.",
          "5 sec each. Cue: Write it. Hover it. Chin it on three.",
          "EXPECT: all three named"],
        ["SCAN the room, back row first.",
          "80%+ -> cold-call one student to sign one back, then go to You Do.",
          "Less -> sign each one twice more with the GIF up, then re-ask."],
      ],
      trap: ["guessing from the first handshape.", "Fix: sign it again, they watch it out before writing."],
      prep: [
        "Receiving is the bar: they write what they saw, not produce on command.",
        "Hover then chin is the school routine. Nobody shows early.",
      ],
      tag: "[CFU | Evaluating impact | HITS 7]",
    }));
  }

  // 11. You Do: Mix and Mingle
  {
    const s = T.youDoSlide(
      pres,
      "Ask two people what sport they like best",
      "Stand up, voices off, and ask two different people what sport they like best. Keep your sport card with you. Write each name and sport on your board.",
      [
        "Tap or wave. Wait for their eyes.",
        "Ask, then watch the whole answer.",
        "Photograph your board into your journal.",
      ],
      T.composeGlanceNotes({
        beats: [
          ["SAY: Voices off. Two different people.", "Your sport card comes with you: point, then sign."],
          "TIME: advance to start the clock. Three minutes.",
          ["CIRCULATE: look for the wait, not the sign.", "The wait is what is new today."],
          ["COLLECT: photograph the board into the journal.", "The photo is the evidence, not the board."],
        ],
        stretch: "ask a third person, and find a sport nobody else named.",
        help: "one person, with you standing beside them for the first exchange.",
        prep: [
          "10 min. Game 3 Mix and Mingle, with the sports they practised in the We Do.",
          "Challenge asks two people and keeps the sport card. Enrichment asks three.",
        ],
        tag: "[You Do | Explicit teaching | HITS 10]",
      }),
      FOOTER,
      {
        where: "Voices off  |  3 minutes  |  Anywhere in the room",
        // The countdown is the slide's right-hand visual, so the task keeps the
        // full left column and nothing lands on top of the words.
        visual: { type: "image", path: A.countdownFile(T, 180) },
      }
    );
    T.addCueStrip(s, ["voicesOff", "timer"]);
  }

  // 12. Exit ticket
  T.exitTicketSlide(
    pres,
    [
      "Rate yourself on the three I can statements: just starting, getting there, got it.",
      "Check the photo of your board is in your journal.",
      "Cultural question: the Australian Deaf Games are run by Deaf people, for Deaf people. Why might that matter to a Deaf athlete as much as winning?",
    ],
    T.composeGlanceNotes({
      beats: [
        "COLLECT nothing. Everyone writes in the journal while you call the rotation.",
        ["CALL group 1, one at a time.",
          "They sign DEAF, SPORT and FAVOURITE, then ask you one person's favourite sport."],
        ["RECORD E, C or P on the tracker,", "against the checklist row, not against the lesson."],
        "Nothing is marked afterwards. The tracker is the record.",
      ],
      prep: [
        "5 min. Evidence piece 3, first pass. Students are never told the schedule.",
        "Rotation group 1 is due today; anyone can be called at the end of any lesson.",
      ],
      tag: "[Exit ticket | Evaluating impact | HITS 7]",
    }),
    FOOTER,
    { title: "In your journal" }
  );

  // 13. Closing
  T.closingSlide(
    pres,
    {
      reflectionPrompt: "Turn and tell your partner one sport you can now sign that you could not this morning.",
      scItems: [
        "I can wave or tap to get someone's attention and wait for their eyes.",
        "I can ask one person about their favourite sport and watch their whole answer.",
        "I can tell the class one thing I found out.",
      ],
      selfAssessment: ["Just starting", "Getting there", "Got it"],
    },
    T.composeGlanceNotes({
      beats: [
        ["SAY: Tell your partner one sport you can sign now,", "that you could not sign this morning."],
        "SAY: Next week we add three more questions to the one you used.",
      ],
      care: "talk privately to any Deaf or hard of hearing student before this lesson runs.",
      prep: [
        "Offer them a prepared answer. They are never the live example.",
        "Protocol practised today: waving. Required in every remaining lesson.",
      ],
      tag: "[Closing | Planning | HITS 9]",
    })
  );

  return pres;
}

buildSportCardPdf();
const pres = build();
fs.mkdirSync(path.join("output", OUT_DIR), { recursive: true });
const file = path.join("output", OUT_DIR, "DeafSport Challenge Session 2 Meeting A Signer.pptx");
pres.writeFile({ fileName: file }).then(() => {
  console.log("PPTX written to " + file);
  report.print();
  const lines = report.rehearsalLines();
  if (lines.length) {
    console.log("\nREHEARSE FIRST");
    lines.forEach((l) => console.log("  " + l));
  }
});
