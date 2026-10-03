"use strict";

/**
 * Deaf Sports in Australia | Challenge (Years 3 and 4) | Lesson 5 Back then, and after
 *
 * Built from the unit document section 9.5. A build script rather than a lesson
 * spec for the reasons given in build_deafsport_enrichment_s2.js. The timeline
 * cards are shared with the Challenge deck and the review week, so they live in
 * builds/auslan_lib.js.
 *
 * THEN has no single sign: Chris signs it as LATER or FINISH, or with a pause
 * and a nod (sign sheet row 31). It shows as a "watch the teacher" card.
 *
 * Challenge calibration (unit document 3 and 9.5): two events with one time
 * sign and the cards face up, years read off the card rather than produced,
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
const FOOTER = "Deaf Sports in Australia  |  Session 5 of 8  |  Years 3-4 Auslan";
const OUT_DIR = path.join("output", "DeafSport_Challenge_S5_Back_Then_And_After");
const RES_DIR = path.join(OUT_DIR, "resources-session5");
const CARDS_PDF = "Session 5 Deaf Sport Timeline Cards.pdf";

const report = A.createSignReport();

function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  T.titleSlide(
    pres,
    "Back then, and after",
    "Deaf Sports in Australia",
    "Lesson 5 of 8  |  Years 3 and 4  |  Term 4",
    "Lesson 5 of 8. The time signs, used on real events from Deaf sport history."
  );

  // 2. Teacher Resources
  P.addResourceSlide(pres, [
    {
      name: "Session 5 Deaf Sport Timeline Cards",
      fileName: "resources-session5/" + CARDS_PDF,
      note: "Seven sets of ten, one per group of four. Cut and laminate. Lesson 5 uses cards 1 to 8.",
    },
  ], T, FOOTER, T.composeGlanceNotes({
    beats: [
      ["SHOW while students settle.", "SAY: Groups of four today. Cards stay face down until I say."],
      "COLLECT every set at the end. The review week uses them again.",
    ],
    prep: [
      "Take out the two cards marked R (1924 and 1955) before handing out sets.",
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
    T.addTitle(s, "Same as every week: write it, then find the signs");
    T.addCard(s, 0.5, 1.45, 9, 1.5, { variant: "tint", tone: C.PRIMARY });
    s.addText("What is the oldest club or team you know of?", {
      x: 0.7, y: 1.45, w: 8.6, h: 1.5, fontSize: 30, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addInstructionCard(s, [
      { role: "header", text: "Same as every week" },
      { text: "Write your answer in English. One sentence." },
      { text: "Underline every word you would need a sign for." },
    ], { x: 0.5, y: 3.15, w: 9, h: 1.4, strip: C.PRIMARY });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        "SILENT from the door. Say nothing; the routine runs itself now.",
        ["Retrieval sits inside the prompt: YEAR and HOW-MANY from last week,", "and TEAM from Lesson 2."],
        "TIME: 5 minutes. Deal with whatever walked in from the playground.",
        ["SAY: You can write a year. Now you need to say which one came first,",
          "and you cannot do that with numbers alone."],
      ],
      prep: "5 min. The bridge line is the reason this lesson exists. Say it as written.",
      tag: "[Do Now | Attention, focus and regulation | HITS 6]",
    }));
  }

  // 4. LI and SC
  T.liSlide(
    pres,
    "I am learning to put events in order using before, after, back then and then.",
    [
      "I can put two events in the right order.",
      "I can recount two events from Deaf sport history in order using a time sign.",
      "I can say which event came first and explain how I know.",
    ],
    T.composeGlanceNotes({
      beats: [
        ["POINT to each criterion.", "SAY: Two events in order, then the reason for the order."],
        ["SAY: By the end, you can tell somebody", "a piece of Deaf sport history in order."],
      ],
      prep: "Criterion 2 is what the exit rotation collects.",
      tag: "[LI and SC | Planning | HITS 1]",
    }),
    FOOTER,
    { numberSC: true, strongHeadings: true, separate: true }
  );

  // 5. I Do: back then, history
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Set it back in time first");
    A.addSignCardRow(T, s, [
      { gloss: "PAST", meaning: "back then, in the past" },
      { gloss: "HISTORY", meaning: "history" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "PAST, HISTORY",
      beats: [
        ["SAY: Same three moves, and today the watching does the work.",
          "Miss the time sign at the front and everything after it lands in the wrong year."],
        "MODEL each twice, slowly then at pace. They copy once.",
        ["SAY: I set the whole story back in the 1880s", "before I say anything that happened there."],
      ],
      trap: ["signing the event first and the time sign last.", "Fix: time sign first, student redoes the sentence."],
      stretch: "show a long time ago and recently by changing how big the movement is.",
      help: "copy beside a partner who has it, then on your own.",
      prep: [
        "Chris: the size of the movement and the face add long ago or recently to PAST.",
        A.REHEARSE_FIRST.HISTORY,
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 6. I Do: before, after, then
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Three signs that put things in order");
    A.addSignCardRow(T, s, [
      { gloss: "BEFORE", meaning: "before" },
      { gloss: "AFTER", meaning: "after" },
      { gloss: "THEN", meaning: "then" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "BEFORE, AFTER, THEN",
      beats: [
        ["MODEL each twice; they copy once.", "Use a real pair of events, not an abstract example."],
        ["SAY: THEN does not have one sign of its own.", "Watch me: a pause and a nod, or LATER, or FINISH."],
        ["SAY: The time sign comes and goes fast and it changes everything after it.",
          "Miss it and you have the right event in the wrong century."],
      ],
      trap: ["BEFORE signed with the place form.", "Fix: show the time form again, student redoes it."],
      stretch: "join three events with BEFORE and THEN in one go.",
      help: "the three meanings written up beside where you stand.",
      prep: [
        "BEFORE is Chris's time form. AFTER and THEN are watch-the-teacher cards; AFTER links Chris's pick.",
        "Chris: Signbank's AFTER has no arm movement, which he disagrees with.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 7. I Do: model two real events
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "Two real events, in order");
    const events = [
      ["1880s", "A Deaf cricket club starts in Melbourne. It is still going today."],
      ["1964", "The first Australian Deaf Games are held in Sydney, with fifteen sports."],
    ];
    events.forEach(([yr, txt], i) => {
      const x = i === 0 ? 0.5 : 5.6;
      T.addCard(s, x, 1.5, 3.9, 2.6, { variant: "tint", tone: C.PRIMARY });
      s.addText(yr, {
        x: x + 0.2, y: 1.62, w: 3.5, h: 0.8, fontSize: 38, fontFace: T.FONT_H, bold: true,
        color: C.PRIMARY, align: "left", valign: "middle", margin: 0,
      });
      s.addText(txt, {
        x: x + 0.2, y: 2.5, w: 3.5, h: 1.45, fontSize: 18, fontFace: T.FONT_B,
        color: C.CHARCOAL, align: "left", valign: "top", margin: 0,
      });
    });
    T.addTextOnShape(
      s,
      "before",
      { x: 4.55, y: 2.45, w: 0.9, h: 0.6, rectRadius: 0.1, fill: { color: C.ACCENT } },
      { fontSize: 15, fontFace: T.FONT_B, bold: true, color: C.WHITE,
        align: "center", valign: "middle", margin: 0 }
    );
    s.addText("Time sign first. Then the event.", {
      x: 0.5, y: 4.3, w: 9, h: 0.5, fontSize: 18, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "the cricket club came first",
      beats: [
        ["SAY: Back in the eighteen eighties, a group of Deaf men in Melbourne",
          "started a cricket club. It is still going today."],
        ["SAY: That was before Deaf Sports Australia existed. Then, in the summer",
          "of nineteen sixty four, the first Australian Deaf Games were held in Sydney."],
        ["GET IT WRONG: swap the time signs, then stop.",
          "SAY: I just told you the Games came first. That is a hundred years out."],
        ["ASK: which came first. Thumb toward it.", "5 sec. EXPECT: the cricket club"],
      ],
      trap: ["reading 1964 as earlier because the last two digits are smaller.", "Fix: compare the first two digits, student redoes it."],
      stretch: "add 1954 between them and recount all three.",
      help: "the two cards face up in order while they recount.",
      care: "if the old club name comes up, call it historical and move on.",
      prep: "Facts from Deaf Sports Australia history and Deaf Connect, checked 17 Sept 2026.",
      sources: "Deaf Sports Australia history page; Deaf Connect, Australian Deaf Games",
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 8. We Do: Game 6 Order The Years
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.SUCCESS);
    T.addBadge(s, "We Do", { color: C.SUCCESS });
    T.addCueStrip(s, ["voicesOff", "partner"]);
    T.addTitle(s, "Game 6 Order The Years");
    T.addInstructionCard(s, [
      { role: "header", text: "Groups of four, in a circle" },
      { text: "Take two cards each. Keep them flat on the desk." },
      { text: "Going left, sign one event from a card. Point to its year." },
      { text: "The group decides where it goes. Lay it in the line." },
      { text: "Play until all eight cards are in order." },
    ], { x: 0.5, y: 1.4, w: 5.9, h: 2.5, strip: C.SUCCESS });
    T.addCard(s, 6.65, 1.4, 2.85, 2.5, { variant: "tint", tone: C.ALERT });
    s.addText("One signer at a time.", {
      x: 6.8, y: 1.6, w: 2.55, h: 0.9, fontSize: 19, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    s.addText("Hands off the table while somebody is signing.", {
      x: 6.8, y: 2.55, w: 2.55, h: 1.2, fontSize: 15, fontFace: T.FONT_B,
      color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
    });
    s.addText("Can't place a card yet? Put it to the side and come back to it.", {
      x: 0.5, y: 4.15, w: 9, h: 0.5, fontSize: 16, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "1880s, 1954, 1964, 1965, 1985, 2005, 2011, 2026",
      beats: [
        ["SET UP: groups of four, every face visible.", "One set face down in the middle, cards 1 to 8 only."],
        "SAY: The turn passes left, so nobody has to be called on.",
        ["CIRCULATE: watch for two people signing at once.",
          "That is everybody talking at once."],
        "COLLECT: the whole line read back around the circle, then the set.",
      ],
      trap: ["ordering by the last two digits of the year.", "Fix: compare the first two digits, student re-places the card."],
      stretch: "add one sentence about what changed because of one event.",
      help: "the group places the 1880s and 2026 cards first, as the two ends.",
      prep: [
        "10 min. Game 6 Order The Years. Protocol today: turn taking, one signer at a time.",
        "Challenge reads years: they point to the year on the card, nobody has to sign it.",
      ],
      tag: "[We Do | Collaborative learning | HITS 5]",
    }));
  }

  // 9. Primary decision point
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.ASSESS);
    T.addBadge(s, "Check it", { color: C.ASSESS });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "Which came first?");
    T.addCard(s, 0.5, 1.5, 9, 1.6, { variant: "tint", tone: C.ASSESS });
    s.addText("The Melbourne cricket club, or the first Australian Deaf Games?", {
      x: 0.7, y: 1.5, w: 8.6, h: 1.6, fontSize: 30, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addTextOnShape(
      s,
      "Everyone signs it to me on three. One, two, three.",
      { x: 1.6, y: 3.4, w: 6.8, h: 0.7, rectRadius: 0.12, fill: { color: C.ASSESS } },
      { fontSize: 19, fontFace: T.FONT_B, bold: true, color: C.WHITE,
        align: "center", valign: "middle", margin: 0 }
    );
    s.addText("Five seconds to think first. Hands still until three.", {
      x: 0.5, y: 4.3, w: 9, h: 0.45, fontSize: 16, fontFace: T.FONT_B,
      color: C.MUTED, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "the cricket club, in the 1880s",
      beats: [
        ["ASK: which came first, the cricket club or the first Games.",
          "5 sec think time. Cue: Everyone signs it to me on three. One, two, three.",
          "EXPECT: CRICKET or CLUB, with BEFORE"],
        ["SCAN the room, back row first.",
          "80%+ -> move to You Do with two events.",
          "Less -> eight students hold cards and walk into order while the rest direct.",
          "Re-check with a different pair of events."],
      ],
      trap: ["treating the biggest-looking number as the earliest.", "Fix: the physical line, then re-check."],
      prep: "The decision point that decides whether the You Do runs without extra help.",
      tag: "[CFU | Evaluating impact | HITS 7]",
    }));
  }

  // 10. You Do
  {
    const s = T.youDoSlide(
      pres,
      "Recount two events in order",
      "Pick two cards and keep them face up. Recount the two events to your partner in order, using a time sign. Your partner writes the two years in order.",
      [
        "Time sign first, then the event.",
        "Partner writes the years in order.",
        "Swap, then check against the cards.",
      ],
      T.composeGlanceNotes({
        beats: [
          ["SAY: Voices off, pairs.", "Different from the game: that ordered cards, this tells one person a sequence."],
          "TIME: 10 minutes, both ways.",
          ["CIRCULATE: look for the time sign at the front of each event.", "That is today's whole point."],
          "COLLECT: the partner's two years, checked against the cards.",
        ],
        stretch: "add a third event with a second time sign.",
        help: "the two cards already in order, so they only add the time sign.",
        prep: "10 min. Challenge: two events, one time sign, cards face up. Enrichment does three.",
        tag: "[You Do | Explicit teaching | HITS 10]",
      }),
      FOOTER,
      { where: "Voices off  |  Table partner  |  Timeline cards and boards" }
    );
    T.addCueStrip(s, ["voicesOff", "partner", "whiteboards"]);
  }

  // 11. Exit ticket
  T.exitTicketSlide(
    pres,
    [
      "Rate yourself on the three I can statements: just starting, getting there, got it.",
      "Write one line: the event you would tell somebody about first.",
      "Cultural question: which event on the timeline do you think mattered most to Deaf people, and why?",
    ],
    T.composeGlanceNotes({
      beats: [
        "COLLECT nothing. Everyone writes in the journal while you call the rotation.",
        ["CALL this week's group, one at a time.", "They sign PAST, BEFORE and one sign from Lesson 4."],
        "Then you sign two events and they sign which came first.",
        ["RECORD E, C or P on the tracker,", "against the checklist row, not against the lesson."],
      ],
      prep: "5 min. Evidence piece 3. Students are never told the schedule.",
      tag: "[Exit ticket | Evaluating impact | HITS 7]",
    }),
    FOOTER,
    { title: "In your journal" }
  );

  // 12. Closing
  T.closingSlide(
    pres,
    {
      reflectionPrompt: "Turn and tell your partner the event that surprised you most, and when it happened.",
      scItems: [
        "I can put two events in the right order.",
        "I can recount two events from Deaf sport history in order using a time sign.",
        "I can say which event came first and explain how I know.",
      ],
      selfAssessment: ["Just starting", "Getting there", "Got it"],
    },
    T.composeGlanceNotes({
      beats: [
        "SAY: Tell your partner the event that surprised you most, and when.",
        ["SAY: Next week, how Deaf sport is set up", "so everything you need can be seen."],
      ],
      prep: [
        "Collect every timeline set. The review week uses all ten cards.",
        "Never cut PAST; the whole history strand hangs off it.",
      ],
      tag: "[Closing | Planning | HITS 9]",
    })
  );

  return pres;
}

A.buildTimelineCardsPdf(path.join(RES_DIR, CARDS_PDF), "Session 5  |  Deaf Sport Timeline Cards  |  Years 3-4 Auslan", C.PRIMARY);
const pres = build();
const file = path.join(OUT_DIR, "DeafSport Challenge Session 5 Back Then And After.pptx");
pres.writeFile({ fileName: file }).then(() => {
  console.log("PPTX written to " + file);
  report.print();
  const lines = report.rehearsalLines();
  if (lines.length) {
    console.log("\nREHEARSE FIRST");
    lines.forEach((l) => console.log("  " + l));
  }
});
