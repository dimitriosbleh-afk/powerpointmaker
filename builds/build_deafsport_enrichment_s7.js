"use strict";

/**
 * Deaf Sports in Australia | Enrichment (Years 5 and 6) | Lesson 7 The Deaf profile interview
 *
 * Built from the unit document sections 9.7, 11.4, 11.5 and 13.1. Evidence
 * piece 1 runs in the You Do: a filmed interview, ticked live on the
 * observational checklist. A build script rather than a lesson spec for the
 * reasons given in build_deafsport_enrichment_s2.js. The profile cards, prompt
 * card and checklist are shared with the Challenge deck, so they live in
 * builds/auslan_lib.js.
 *
 * Two profile cards are real people, one living. Every profile stays with what
 * they did in their sport (unit document care note).
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme, weekToVariant } = require("../themes/factory");
const P = require("../themes/pdf_helpers");
const A = require("./auslan_lib");

const T = createTheme("literacy", "grade56", weekToVariant(1));
const C = T.C;
const FOOTER = "Deaf Sports in Australia  |  Session 7 of 8  |  Years 5-6 Auslan";
const OUT_DIR = path.join("output", "DeafSport_Enrichment_S7_The_Deaf_Profile_Interview");
const RES_DIR = path.join(OUT_DIR, "resources-session7");
const PDFS = {
  cards: "Session 7 Deaf Athlete Profile Cards.pdf",
  prompt: "Session 7 Interview Prompt Card.pdf",
  checklist: "Session 7 Observational Checklist.pdf",
};

const report = A.createSignReport();

function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  T.titleSlide(
    pres,
    "The Deaf profile interview",
    "Deaf Sports in Australia",
    "Lesson 7 of 8  |  Years 5 and 6  |  Term 4",
    "Lesson 7 of 8. Four questions as one conversation, on camera. Evidence piece 1."
  );

  // 2. Teacher Resources
  P.addResourceSlide(pres, [
    { name: "Session 7 Deaf Athlete Profile Cards", fileName: "resources-session7/" + PDFS.cards,
      note: "Seven sets of six, one per group of four. Cut and laminate." },
    { name: "Session 7 Interview Prompt Card", fileName: "resources-session7/" + PDFS.prompt,
      note: "Fourteen, one per pair. Laminated so it lies flat." },
    { name: "Session 7 Observational Checklist", fileName: "resources-session7/" + PDFS.checklist,
      note: "One page per class, on a clipboard. Tick live." },
  ], T, FOOTER, T.composeGlanceNotes({
    beats: [
      ["CHECK filming consent against your own list before anyone arrives.",
        "No consent: same interview, live with you, arranged quietly in advance."],
      ["SET UP Google Classroom: one assignment per class, Lesson 7 Interview,",
        "file upload, due at the end of the lesson."],
      "CHARGE five devices. Clipboard and checklist ready.",
    ],
    prep: [
      "A blanket media form does not always cover uploading faces. Read what yours says.",
      A.ATTRIBUTION,
    ],
    tag: "[Setup | Planning | HITS 2]",
  }));

  // 3. Do Now: Auslan90 Australian Deaf Games, Day 7
  A.addVideoDoNow(T, pres, {
    day: 7,
    footer: FOOTER,
    bridge: ["SAY: Everyone in that video was interviewed.", "Today you run an interview of your own, and it goes on camera."],
  });

  // 4. LI and SC
  T.liSlide(
    pres,
    "I am learning to build a four-fact profile and run it as an interview I can record.",
    [
      "I can give two facts about a person on a card.",
      "I can run a four-turn interview with a partner, voice off.",
      "I can add one follow-up question that was not on the card.",
    ],
    T.composeGlanceNotes({
      beats: [
        ["POINT to each criterion.", "SAY: Two facts, then four turns, then a question of your own."],
        "SAY: By the end, your interview is filmed and in your journal.",
      ],
      prep: "Criterion 2 is evidence piece 1, ticked live on the checklist.",
      tag: "[LI and SC | Planning | HITS 1]",
    }),
    FOOTER,
    { numberSC: true, strongHeadings: true, separate: true }
  );

  // 5. I Do: compete, proud, happen
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Three signs for an athlete's story");
    A.addSignCardRow(T, s, [
      { gloss: "COMPETE", meaning: "compete" },
      { gloss: "PROUD", meaning: "proud" },
      { gloss: "HAPPEN", meaning: "happen" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "COMPETE, PROUD, HAPPEN",
      beats: [
        ["SAY: Same three moves, now inside an interview. Get their eyes,",
          "watch the whole answer, copy back what you understood."],
        "MODEL each twice, slowly then at pace. They copy once.",
        ["SAY: What are you proud of is the last question, every time.", "It is the one people remember."],
      ],
      trap: ["COMPETE signed as a race between two people.", "Fix: show the competition form again, student redoes it."],
      stretch: "ask a partner what happened at their last game, using HAPPEN.",
      help: "copy beside a partner who has it, then on your own.",
      prep: [
        "Chris: COMPETE can also mean race. HAPPEN has two variants; show both.",
        A.REHEARSE_FIRST.PROUD,
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 6. I Do: first, finish
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "First to finish, every time");
    A.addSignCardRow(T, s, [
      { gloss: "FIRST", meaning: "first" },
      { gloss: "FINISH", meaning: "finish" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "FIRST, FINISH",
      beats: [
        "MODEL each twice; they copy once.",
        ["SAY: Every profile has four facts in the same order.",
          "Who. What sport. When. What they won or did."],
        ["SAY: FIRST opens it and FINISH closes it.", "Then the person watching knows it is over."],
      ],
      trap: ["stopping without FINISH, so the partner waits.", "Fix: sign FINISH, student redoes the ending."],
      stretch: "use FINISH to mean then, between two facts.",
      help: "the four slot words written up: who, sport, when, won.",
      prep: [
        "FINISH is a watch-the-teacher card linking Chris's pick; Signbank lacks his form.",
        "Chris: FINISH also works as then, a discourse marker. One or two hands.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 7. I Do: a four-fact profile
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "A profile has four facts, in this order");
    const [who, sport, when, won] = A.PROFILE_CARDS[0];
    const slots = [["Who", who], ["What sport", sport], ["When", when], ["What they won", won]];
    const sw = (9 - 0.2 * 3) / 4;
    slots.forEach(([head, body], i) => {
      const x = 0.5 + i * (sw + 0.2);
      T.addCard(s, x, 1.45, sw, 2.85, { variant: "tint", tone: C.PRIMARY });
      s.addText(String(i + 1) + "  " + head, {
        x: x + 0.12, y: 1.55, w: sw - 0.24, h: 0.45, fontSize: 15, fontFace: T.FONT_B, bold: true,
        color: C.PRIMARY, align: "left", valign: "middle", margin: 0,
      });
      s.addText(body, {
        x: x + 0.12, y: 2.05, w: sw - 0.24, h: 2.15, fontSize: i === 3 ? 15 : 21, fontFace: T.FONT_B,
        bold: i !== 3, color: C.CHARCOAL, align: "left", valign: "top", margin: 0,
      });
    });
    s.addText("First to finish. The same order on every card.", {
      x: 0.5, y: 4.5, w: 9, h: 0.4, fontSize: 16, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "who, what sport, when, what they won",
      beats: [
        ["MODEL the whole profile once, at pace, without stopping.",
          "SAY: Barry Knapman. Australia. Diving."],
        ["SAY: He competed at the world games for the Deaf in Washington in 1965",
          "and won gold on the springboard. In 1969 in Belgrade, silver."],
        ["ASK: what is the first fact in every profile.",
          "5 sec. Cue: Everyone signs it to me on three. One, two, three.",
          "EXPECT: WHO"],
      ],
      trap: ["giving the facts in whatever order they come.", "Fix: point at the slots, student re-signs in order."],
      stretch: "sign the Cindy-Lu Bailey card in the same four slots.",
      help: "the card face up, so the load is the signing, not the facts.",
      care: "real people: what they did in their sport only. Never health or family.",
      prep: "Cindy-Lu Bailey is also in Auslan90's Day 6 video, if the class has seen it.",
      sources: "International Committee of Sports for the Deaf athlete records, checked 17 Sept 2026",
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 8. I Do: the four-turn interview, and a repair
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "Four questions, then one of your own");
    const qs = ["Who are you?", "What sport do you play?", "When did you start?", "What are you proud of?"];
    const qw = (9 - 0.2 * 3) / 4;
    qs.forEach((q, i) => {
      const x = 0.5 + i * (qw + 0.2);
      T.addCard(s, x, 1.45, qw, 1.55, { variant: "tint", tone: C.PRIMARY });
      s.addText(String(i + 1), {
        x, y: 1.55, w: qw, h: 0.38, fontSize: 15, fontFace: T.FONT_B, bold: true,
        color: C.PRIMARY, align: "center", valign: "middle", margin: 0,
      });
      s.addText(q, {
        x: x + 0.1, y: 1.95, w: qw - 0.2, h: 0.95, fontSize: 18, fontFace: T.FONT_H, bold: true,
        color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
      });
    });
    T.addTextOnShape(
      s,
      "Then one question of your own",
      { x: 0.5, y: 3.2, w: 9, h: 0.62, rectRadius: 0.12, fill: { color: C.PRIMARY } },
      { fontSize: 19, fontFace: T.FONT_B, bold: true, color: C.WHITE,
        align: "center", valign: "middle", margin: 0 }
    );
    s.addText("Missed an answer? Repair it and keep going. A repair is a pass, not a fail.", {
      x: 0.5, y: 4.05, w: 9, h: 0.6, fontSize: 17, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        ["MODEL the four-turn interview with one student who agreed in advance.",
          "Four questions, four answers, no stopping."],
        ["SAY: My next question is already ready,", "so my eyes never leave the answer."],
        ["GET IT WRONG: miss an answer, repair with AGAIN.", "Miss another, repair with WHAT MEAN?."],
        ["SAY: Two breakdowns and neither ended the interview.", "That is what I want on the recording."],
      ],
      prep: [
        "Light check: four fingers if you can name all four questions in order.",
        "The sequence must be automatic before a camera comes out.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3]",
    }));
  }

  // 9. We Do: Game 8 Which Card Is Mine
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.SUCCESS);
    T.addBadge(s, "We Do", { color: C.SUCCESS });
    T.addCueStrip(s, ["voicesOff", "partner"]);
    T.addTitle(s, "Game 8 Which Card Is Mine");
    T.addInstructionCard(s, [
      { role: "header", text: "Groups of four, in a circle" },
      { text: "All six cards face up in the middle." },
      { text: "Secretly choose one card. Remember it." },
      { text: "Take turns: sign its four facts. Do not say which card." },
      { text: "Everyone points to a card at the end." },
    ], { x: 0.5, y: 1.4, w: 5.9, h: 2.5, strip: C.SUCCESS });
    T.addCard(s, 6.65, 1.4, 2.85, 2.5, { variant: "tint", tone: C.ALERT });
    s.addText("Point at the end, never during.", {
      x: 6.8, y: 1.6, w: 2.55, h: 0.9, fontSize: 19, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    s.addText("Looking at the cards mid-profile means you stopped watching.", {
      x: 6.8, y: 2.55, w: 2.55, h: 1.2, fontSize: 15, fontFace: T.FONT_B,
      color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
    });
    s.addText("The describer says if the group is right. Then the next player goes.", {
      x: 0.5, y: 4.15, w: 9, h: 0.5, fontSize: 16, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "the card whose four facts were signed",
      beats: [
        ["SET UP: groups of four, cards in the middle, never in hands.", "Circle seating so every face is visible."],
        "SAY: Two people can choose the same card. That is fine.",
        ["CIRCULATE: watch for eyes on the cards mid-profile.", "That student has stopped watching the language."],
        "COLLECT: every set back at the end of the game.",
      ],
      trap: ["pointing before the profile finishes.", "Fix: hands down, describer signs it again, then point."],
      stretch: "the describer answers questions instead of signing a monologue.",
      help: "describe one of the two cards modelled on the last slides.",
      prep: "10 min. Game 8 Which Card Is Mine. The harder version bridges into the interview.",
      tag: "[We Do | Collaborative learning | HITS 5]",
    }));
  }

  // 10. Primary decision point
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.ASSESS);
    T.addBadge(s, "Check it", { color: C.ASSESS });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "Before the cameras come out");
    T.addCard(s, 0.5, 1.5, 9, 1.6, { variant: "tint", tone: C.ASSESS });
    s.addText("Sign the first fact in every profile.", {
      x: 0.7, y: 1.5, w: 8.6, h: 1.6, fontSize: 34, fontFace: T.FONT_H, bold: true,
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
      answer: "who it is",
      beats: [
        ["ASK: sign the first fact slot in a profile.",
          "5 sec think time. Cue: Everyone signs it to me on three. One, two, three.",
          "EXPECT: WHO, or a name"],
        ["SCAN the room, back row first.",
          "80%+ -> start filming.",
          "Less -> four slot headings on the board, fill the first and last,",
          "pairs supply the middle two. Re-check before any device comes out."],
        ["FOLLOW UP one student.", "SAY: You started with the name. Say why that has to come first."],
      ],
      trap: ["reciting the four slots as a list, not using them.", "Fix: the partial model, then re-check."],
      prep: "The decision point that decides whether filming starts now or after a rebuild.",
      tag: "[CFU | Evaluating impact | HITS 7]",
    }));
  }

  // 11. You Do: the filmed interview
  {
    const s = T.youDoSlide(
      pres,
      "Two Minute Interview, on camera",
      "Interview your partner: four questions, four answers, then one of your own. Card face down. Then swap.",
      [
        "Sit side on to the camera, faces lit.",
        "One take, no stopping. Repair and keep going.",
        "Put the file in your Lesson 7 journal slide.",
      ],
      T.composeGlanceNotes({
        beats: [
          ["SAY: Voices off. One student per three pairs runs the camera.",
            "Filming students rotate so everybody is recorded."],
          "TIME: advance to start two minutes. Swap, then two minutes back.",
          ["CIRCULATE with the clipboard. Tick the six columns live.", "Do not re-watch recordings to score them."],
          "COLLECT: every file in its journal slide before the bell.",
        ],
        stretch: "answer in role as a person from a profile card.",
        help: "the prompt card face up for the recording, marked S on the checklist.",
        prep: [
          "15 min. Evidence piece 1. One retake only if the recording missed part.",
          "Proficient: four turns, card face down, eyes held, repairs without stopping.",
        ],
        tag: "[You Do | Evaluating impact | HITS 8]",
      }),
      FOOTER,
      {
        where: "Voices off  |  2 minutes each way  |  Prompt card flat",
        visual: { type: "image", path: A.countdownFile(T, 120) },
      }
    );
    T.addCueStrip(s, ["voicesOff", "partner", "timer"]);
  }

  // 12. Exit ticket
  T.exitTicketSlide(
    pres,
    [
      "Rate yourself on the three I can statements: just starting, getting there, got it.",
      "Write one line: the follow-up question you asked that was not on the card.",
      "Cultural question: why do we ask an athlete what they are proud of, and not how they became deaf?",
    ],
    T.composeGlanceNotes({
      beats: [
        "COLLECT nothing. Everyone writes in the journal.",
        ["CHECK the upload: every file is in its Lesson 7 journal slide.", "Fix any missing ones now, not next week."],
        "No rotation this week. The checklist covers the whole class.",
      ],
      prep: "5 min. File the checklist sheet today: about 5 minutes per class.",
      tag: "[Exit ticket | Evaluating impact | HITS 7]",
    }),
    FOOTER,
    { title: "In your journal" }
  );

  // 13. Closing
  T.closingSlide(
    pres,
    {
      reflectionPrompt: "Turn and tell your partner the best answer you got in your interview.",
      scItems: [
        "I can give two facts about a person on a card.",
        "I can run a four-turn interview with a partner, voice off.",
        "I can add one follow-up question that was not on the card.",
      ],
      selfAssessment: ["Just starting", "Getting there", "Got it"],
    },
    T.composeGlanceNotes({
      beats: [
        "SAY: Tell your partner the best answer you got.",
        ["SAY: Next week, how Deaf sport is set up", "so everything you need can be seen."],
      ],
      prep: [
        "Protocol practised today: light on faces, nobody signing in front of a window.",
        "Never cut the recording; evidence piece 1 has no other slot.",
      ],
      tag: "[Closing | Planning | HITS 9]",
    })
  );

  return pres;
}

A.buildProfileCardsPdf(path.join(RES_DIR, PDFS.cards), "Session 7  |  Profile Cards  |  Years 5-6 Auslan", C.PRIMARY);
A.buildInterviewPromptCardPdf(path.join(RES_DIR, PDFS.prompt), "Session 7  |  Interview Prompt Card  |  Years 5-6 Auslan", C.PRIMARY);
A.buildObservationChecklistPdf(path.join(RES_DIR, PDFS.checklist), "Session 7  |  Observational Checklist  |  Years 5-6 Auslan", C.PRIMARY,
  "Enrichment: Proficient is four turns, card face down.");
const pres = build();
const file = path.join(OUT_DIR, "DeafSport Enrichment Session 7 The Deaf Profile Interview.pptx");
pres.writeFile({ fileName: file }).then(() => {
  console.log("PPTX written to " + file);
  report.print();
  const lines = report.rehearsalLines();
  if (lines.length) {
    console.log("\nREHEARSE FIRST");
    lines.forEach((l) => console.log("  " + l));
  }
});
