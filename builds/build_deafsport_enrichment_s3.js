"use strict";

/**
 * Deaf Sports in Australia | Enrichment (Years 5 and 6) | Lesson 3 Who, what, where
 *
 * Built from the unit document section 9.3. A build script rather than a lesson
 * spec for the reasons given in build_deafsport_enrichment_s2.js: animated sign
 * cards, Signbank lookup cards, the cue strip and the countdown are not things
 * a spec can express. Shared parts live in builds/auslan_lib.js.
 *
 * The three question signs resolve by SENSE, not by entry order. Signbank's
 * first entry for "who" is the noun (someone); only the second carries the
 * question use, and the shared layer goes and finds it. Chris overrides any of
 * it from reference/auslan/signbank_links/overrides.json.
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme, weekToVariant } = require("../themes/factory");
const P = require("../themes/pdf_helpers");
const A = require("./auslan_lib");

const T = createTheme("literacy", "grade56", weekToVariant(1));
const C = T.C;
const FOOTER = "Deaf Sports in Australia  |  Session 3 of 8  |  Years 5-6 Auslan";
const OUT_DIR = path.join("output", "DeafSport_Enrichment_S3_Who_What_Where");
const RES_DIR = path.join(OUT_DIR, "resources-session3");
const CARDS_PDF = "Session 3 Team Role Cards.pdf";

const report = A.createSignReport();

function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  T.titleSlide(
    pres,
    "Who, what, where",
    "Deaf Sports in Australia",
    "Lesson 3 of 8  |  Years 5 and 6  |  Term 4",
    "Lesson 3 of 8. The three profile questions, and the repair that gets an explanation rather than a repeat."
  );

  // 2. Teacher Resources
  P.addResourceSlide(pres, [
    {
      name: "Session 3 Team Role Cards",
      fileName: "resources-session3/" + CARDS_PDF,
      note: "One set of 28, cut and laminated. One set does all five classes.",
    },
  ], T, FOOTER, T.composeGlanceNotes({
    beats: [
      ["SHOW while students settle.", "SAY: One card each, flat on the desk. Do not pick it up."],
      "COLLECT every card at the end. The set is laminated and reused all term.",
    ],
    prep: [
      "Print and cut once before term. Boards out, Sign It! one per pair.",
      A.ATTRIBUTION,
    ],
    tag: "[Setup | Planning | HITS 2]",
  }));

  // 3. Do Now: Auslan90 Australian Deaf Games, Day 3
  A.addVideoDoNow(T, pres, {
    day: 3,
    footer: FOOTER,
    bridge: ["SAY: Those interviews were people asking each other questions.", "Today you ask questions of the person opposite you."],
  });

  // 4. LI and SC
  T.liSlide(
    pres,
    "I am learning to ask who, what and where, and to ask for an explanation when I miss something.",
    [
      "I can ask one of the three profile questions.",
      "I can ask all three and record the answers.",
      "I can use WHAT MEAN? when a repeat would not help.",
    ],
    T.composeGlanceNotes({
      beats: [
        ["POINT to each criterion.", "SAY: Three questions, and one way out when you miss an answer."],
        "SAY: By the end, everyone knows three things about the person opposite them.",
      ],
      prep: "Criterion 2 is what the exit rotation collects this week.",
      tag: "[LI and SC | Planning | HITS 1]",
    }),
    FOOTER,
    { numberSC: true, strongHeadings: true, separate: true }
  );

  // 5. I Do: anchor and set it up
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Set the topic before you ask");
    T.addTextOnShape(
      s,
      "Eyes up. Watch. Copy.",
      { x: 0.5, y: 1.4, w: 9, h: 0.72, rectRadius: 0.12, fill: { color: C.PRIMARY } },
      { fontSize: 26, fontFace: T.FONT_H, bold: true, color: C.WHITE,
        align: "center", valign: "middle", margin: 0 }
    );
    const order = ["Who are you?", "What sport do you play?", "Where are you from?"];
    const ow = (9 - 0.22 * 2) / 3;
    order.forEach((q, i) => {
      const ox = 0.5 + i * (ow + 0.22);
      T.addCard(s, ox, 2.5, ow, 1.55, { variant: "tint", tone: C.PRIMARY });
      s.addText(String(i + 1), {
        x: ox, y: 2.62, w: ow, h: 0.4, fontSize: 15, fontFace: T.FONT_B, bold: true,
        color: C.PRIMARY, align: "center", valign: "middle", margin: 0,
      });
      s.addText(q, {
        x: ox + 0.12, y: 3.02, w: ow - 0.24, h: 0.9, fontSize: 21, fontFace: T.FONT_H, bold: true,
        color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
      });
    });
    s.addText("The same three, in the same order, all term.", {
      x: 0.5, y: 4.25, w: 9, h: 0.45, fontSize: 16, fontFace: T.FONT_B,
      color: C.MUTED, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        ["SAY: Same three moves. Eyes up, watch, copy.",
          "What changes today is that you also get to ask."],
        ["SAY: I set the topic first, so they know what I am asking about",
          "before I ask it. If I do not, they spend the whole question",
          "working out what I mean."],
        ["SAY: Three questions, and they come in this order all term,",
          "because an order you do not have to think about",
          "is an order you can use while you are nervous."],
      ],
      prep: [
        "The anchor is restated in these exact words every week. Do not reword it.",
        "Hover then chin is the school routine. Nobody shows early.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3]",
    }));
  }

  // 6. I Do: the three question signs
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Three question signs. Watch me, then copy once");
    A.addSignCardRow(T, s, [
      { gloss: "WHO", meaning: "who" },
      { gloss: "WHAT", meaning: "what" },
      { gloss: "WHERE", meaning: "where" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "WHO, WHAT, WHERE",
      beats: [
        "MODEL each question form twice, slowly then at pace. They copy once.",
        ["SAY: Watch the whole answer even when you already think you know it.",
          "Half the answers this term will surprise you."],
        ["CHECK as they copy: three visibly different questions,",
          "not one form repeated with a different face."],
      ],
      trap: ["WHERE collapsing into WHAT under time pressure.", "Fix: slow the third one down on its own, student redoes it."],
      stretch: "ask a fourth question of your own and say why you chose it.",
      help: "the three question words in English, face up. Then the load is the sign.",
      prep: [
        "CHECK GRAMMAR: confirm where the question sign sits before you teach it.",
        "WHO, WHAT and WHERE are Chris's vetted Signbank entries. The question face carries the question.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 7. I Do: practise and the repair
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "One more sign, and how to fix a broken exchange");
    A.addSignCardRow(T, s, [
      { gloss: "PRACTISE", meaning: "practise, train" },
      { gloss: "WHAT MEAN?", meaning: "what do you mean" },
    ], { y: 1.4, bottom: 4.5, report });
    s.addText("Repair means fixing an exchange without ending it. AGAIN gets the same thing twice. WHAT MEAN? gets an explanation.", {
      x: 0.5, y: 4.55, w: 9, h: 0.45, fontSize: 15, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "PRACTISE, and WHAT MEAN? when a repeat will not help",
      beats: [
        ["GET IT WRONG: ask, look down mid-answer, miss it.",
          "Use AGAIN from last week. Get the same answer, still not understand."],
        ["SAY: AGAIN gave me the same thing twice and I still had nothing.",
          "WHAT MEAN? gets me an explanation."],
        ["MODEL WHAT MEAN? twice; they copy once.", "It is the form this school uses. Do not extend it."],
      ],
      trap: ["using AGAIN twice and giving up.", "Fix: freeze the pair, student signs WHAT MEAN? instead."],
      stretch: "use WHAT MEAN? once in the You Do without being told to.",
      help: "you answer a repair slowly and visibly, so they see what an explanation looks like.",
      prep: [
        "WHAT MEAN? is a fixed two-sign form the school supplies. It has no single Signbank entry.",
        "PRACTISE is Chris's vetted entry, the training sense, not LEARNER.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 8. We Do: Find Your Team
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.SUCCESS);
    T.addBadge(s, "We Do", { color: C.SUCCESS });
    T.addCueStrip(s, ["voicesOff", "partner"]);
    T.addTitle(s, "Game 4 Find Your Team");
    T.addInstructionCard(s, [
      { role: "header", text: "How to play" },
      { text: "Your card has a name, a sport and a city. Leave it flat on the desk." },
      { text: "Get eyes, then ask all three questions." },
      { text: "Find the other three people on your team." },
      { text: "A found team sits down together and keeps checking new arrivals." },
    ], { x: 0.5, y: 1.4, w: 5.7, h: 2.5, strip: C.SUCCESS });
    T.addCard(s, 6.45, 1.4, 3.05, 2.5, { variant: "tint", tone: C.ALERT });
    s.addText("Two teams both play swimming.", {
      x: 6.6, y: 1.6, w: 2.75, h: 0.9, fontSize: 19, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    s.addText("One matching answer is never enough. Ask all three.", {
      x: 6.6, y: 2.55, w: 2.75, h: 1.2, fontSize: 15, fontFace: T.FONT_B,
      color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
    });
    const rules = ["No queueing behind a back", "One person at a time", "Wait for their eyes"];
    const rw = (9 - 0.2 * 2) / 3;
    rules.forEach((r, i) => {
      T.addTextOnShape(
        s,
        r,
        { x: 0.5 + i * (rw + 0.2), y: 4.1, w: rw, h: 0.6, rectRadius: 0.1, fill: { color: C.ALERT } },
        { fontSize: 15, fontFace: T.FONT_B, bold: true, color: C.WHITE,
          align: "center", valign: "middle", margin: 0 }
      );
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "a team is four people with the same sport AND the same city",
      beats: [
        ["SET UP: one card each, flat on the desk station.", "Cards never travel in a hand."],
        "SAY: Half the class moves, half stays. Then swap.",
        ["CIRCULATE: watch for students deciding on the sport answer alone.",
          "That is the error the deal is built to catch."],
        "COLLECT: every team seated together, then every card back.",
      ],
      trap: ["joining the first person with the same sport.", "Fix: send them back to ask the city question."],
      prep: "10 min. Game 4 Find Your Team. Teams 1 and 2 share swimming on purpose.",
      tag: "[We Do | Collaborative learning | HITS 5]",
    }));
  }

  // 9. Primary decision point
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.ASSESS);
    T.addBadge(s, "Check it", { color: C.ASSESS });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "Watch the question. Write what it means.");
    T.addCard(s, 0.5, 1.5, 9, 1.4, { variant: "tint", tone: C.ASSESS });
    s.addText("I sign a whole question. You write what it means in English.", {
      x: 0.7, y: 1.5, w: 8.6, h: 1.4, fontSize: 30, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    const steps = ["Watch all of it", "Write it in English", "Chin it on three", "One of you signs the answer back"];
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
      answer: "where are you from",
      beats: [
        ["ASK: where are you from. Sign the whole question.",
          "5 sec think time. Cue: Write it. Hover it. Chin it on three.",
          "EXPECT: where are you from"],
        ["SCAN the room, back row first.",
          "80%+ -> cold-call one student to sign the answer back, then go to You Do.",
          "Less -> three headings on the board, you sign a question, they point.",
          "Six times fast, then re-ask the whole question."],
        ["FOLLOW UP the student who signed back.", "SAY: Do you agree with that one? Add one thing."],
      ],
      trap: ["reading WHERE as WHAT because they grabbed the nearest question sign.", "Fix: sort, then re-ask."],
      prep: [
        "The decision point that changes Lesson 3. Receiving a whole question is the bar.",
        "Re-check with the same routine after any pivot.",
      ],
      tag: "[CFU | Evaluating impact | HITS 7]",
    }));
  }

  // 10. You Do
  {
    const s = T.youDoSlide(
      pres,
      "Interview the person opposite you",
      "Ask your partner all three questions and write their three answers on your board. Then swap.",
      [
        "Wave or tap. Wait for their eyes.",
        "Ask, then watch the whole answer.",
        "Pencils down while they sign. Write after.",
      ],
      T.composeGlanceNotes({
        beats: [
          ["SAY: Voices off, pairs.",
            "Different from the game: that asked about a card, this asks about the real person."],
          "TIME: 10 minutes, both ways.",
          ["CIRCULATE: look for pencils moving while a partner signs.",
            "That is the thing to stop."],
          "COLLECT: boards hovered, three answers each.",
        ],
        stretch: "ask a fourth question of your own and say why you chose it.",
        help: "the three question words in English, face up on the desk.",
        prep: [
          "10 min. Year 5 keeps the prompt card. Year 6 works with it face down.",
          "Watch the whole answer, then write. Never both at once.",
        ],
        tag: "[You Do | Explicit teaching | HITS 10]",
      }),
      FOOTER,
      { where: "Voices off  |  Table partner  |  Boards out" }
    );
    T.addCueStrip(s, ["voicesOff", "partner", "whiteboards"]);
  }

  // 11. Exit ticket
  T.exitTicketSlide(
    pres,
    [
      "Rate yourself on the three I can statements: just starting, getting there, got it.",
      "Write one line: one thing you found out about your partner.",
      "Cultural question: today you could not ask anything until you had somebody's eyes. What does that tell you about how a Deaf conversation starts?",
    ],
    T.composeGlanceNotes({
      beats: [
        "COLLECT nothing. Everyone writes in the journal while you call the rotation.",
        ["CALL group 2, one at a time.",
          "They sign WHO, WHERE and one sign from Lesson 1."],
        "Then they ask you one profile question and watch the whole answer.",
        ["RECORD E, C or P on the tracker,", "against the checklist row, not against the lesson."],
      ],
      prep: [
        "5 min. Evidence piece 3, second pass. Rotation group 2 is due today.",
        "Students are never told the schedule. Anyone can be called at any lesson.",
      ],
      tag: "[Exit ticket | Evaluating impact | HITS 7]",
    }),
    FOOTER,
    { title: "In your journal" }
  );

  // 12. Closing
  T.closingSlide(
    pres,
    {
      reflectionPrompt: "Turn and tell your partner the one thing you found out about them today.",
      scItems: [
        "I can ask one of the three profile questions.",
        "I can ask all three and record the answers.",
        "I can use WHAT MEAN? when a repeat would not help.",
      ],
      selfAssessment: ["Just starting", "Getting there", "Got it"],
    },
    T.composeGlanceNotes({
      beats: [
        "SAY: Turn and tell your partner the one thing you found out about them.",
        ["SAY: Next week we add when and how many.", "The term fills up with dates from here."],
      ],
      prep: [
        "Protocol practised today: the table tap, for when a wave will not reach.",
        "Required in every remaining lesson alongside waving.",
      ],
      tag: "[Closing | Planning | HITS 9]",
    })
  );

  return pres;
}

A.buildTeamRoleCardsPdf(path.join(RES_DIR, CARDS_PDF), "Session 3  |  Team Role Cards  |  Years 5-6 Auslan", C.PRIMARY);
const pres = build();
const file = path.join(OUT_DIR, "DeafSport Enrichment Session 3 Who What Where.pptx");
pres.writeFile({ fileName: file }).then(() => {
  console.log("PPTX written to " + file);
  report.print();
  const lines = report.rehearsalLines();
  if (lines.length) {
    console.log("\nREHEARSE FIRST");
    lines.forEach((l) => console.log("  " + l));
  }
});
