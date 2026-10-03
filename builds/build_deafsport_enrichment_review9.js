"use strict";

/**
 * Deaf Sports in Australia | Enrichment (Years 5 and 6) | Review lesson, Week 9 game set
 *
 * Built from the unit document section 9.9. The review lesson runs twice, in
 * weeks 8 and 9, with a different game set each time. Nothing new is taught and
 * no evidence is collected. It is also the CRT lesson: every note is written so
 * a teacher with no Auslan can run it without modelling a sign.
 *
 * Nothing new to print. The Week 9 set reuses the Lesson 6 medal tally sheets,
 * named on the resources slide rather than regenerated (AUSLAN_2 section 8).
 *
 * Watch and Write normally has the teacher signing. A relief teacher may not
 * sign, so a student from each team signs a word from the term list instead.
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme, weekToVariant } = require("../themes/factory");

const T = createTheme("literacy", "grade56", weekToVariant(1));
const C = T.C;
const FOOTER = "Deaf Sports in Australia  |  Review, Week 9  |  Years 5-6 Auslan";
const RANGE = "1 to 100";
const OUT_DIR = path.join("output", "DeafSport_Enrichment_Review_Week_9");

const SC = [
  "I can recognise a sign from the term list.",
  "I can produce a sign my team needs on the first try.",
  "I can wait for eyes before I start every exchange.",
];

/** A game slide: rules on the left, the scoring card on the right. */
function gameSlide(pres, badge, tone, title, rules, scoring, footerLine, notes, cues) {
  const s = pres.addSlide();
  T.addTopBar(s, tone);
  T.addBadge(s, badge, { color: tone });
  T.addCueStrip(s, cues);
  T.addTitle(s, title);
  T.addInstructionCard(s, [{ role: "header", text: "How to play" }].concat(rules.map((r) => ({ text: r }))),
    { x: 0.5, y: 1.4, w: 5.9, h: 2.6, strip: tone });
  T.addCard(s, 6.65, 1.4, 2.85, 2.6, { variant: "tint", tone: C.ACCENT });
  s.addText("Points", {
    x: 6.8, y: 1.55, w: 2.55, h: 0.5, fontSize: 19, fontFace: T.FONT_H, bold: true,
    color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
  });
  s.addText(scoring, {
    x: 6.8, y: 2.1, w: 2.55, h: 1.75, fontSize: 15, fontFace: T.FONT_B,
    color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
  });
  s.addText(footerLine, {
    x: 0.5, y: 4.2, w: 9, h: 0.5, fontSize: 16, fontFace: T.FONT_B, bold: true,
    color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
  });
  T.addFooter(s, FOOTER);
  s.addNotes(notes);
  return s;
}

function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  T.titleSlide(
    pres,
    "Review: the Week 9 games",
    "Deaf Sports in Australia",
    "Review lesson  |  Years 5 and 6  |  Term 4",
    "Review lesson, Week 9 set. Game-based review of signs already taught. Nothing new. A teacher with no Auslan can run it."
  );

  // 2. Teacher Resources
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "Teacher Resources", { color: C.PRIMARY, w: 2.6 });
    T.addTitle(s, "Nothing new to print");
    T.addInstructionCard(s, [
      { role: "header", text: "Get these out before the lesson" },
      { text: "The Lesson 6 Medal Tally Sheets A and B, one per student." },
      { text: "A mini whiteboard and marker for every team and every student." },
      { text: "A folder per pair to prop up as a screen." },
      { text: "Sign It!, one per pair, and the student journals." },
    ], { x: 0.5, y: 1.45, w: 9, h: 2.3, strip: C.PRIMARY });
    s.addText("No Auslan needed to run this lesson. If a student asks how to sign something, open Sign It! at the page on the term list. Do not guess a sign.", {
      x: 0.5, y: 3.95, w: 9, h: 0.75, fontSize: 15, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "left", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        ["SHOW while students settle.", "SAY: Journals open on the tracker page."],
        "SET UP four teams before the first game. Mix confident and unsure students.",
      ],
      prep: [
        "The CRT Review Pack has the term list, game rules and a points sheet.",
        "Leave a note: which games ran, which signs the class struggled with.",
      ],
      tag: "[Setup | Planning | HITS 2]",
    }));
  }

  // 3. Do Now
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "Do Now", { color: C.PRIMARY });
    T.addCueStrip(s, ["voicesOff"]);
    T.addTitle(s, "Your tracker page");
    T.addCard(s, 0.5, 1.45, 9, 1.5, { variant: "tint", tone: C.PRIMARY });
    s.addText("Tick every sign you can do now.", {
      x: 0.7, y: 1.45, w: 8.6, h: 1.5, fontSize: 38, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addInstructionCard(s, [
      { role: "header", text: "In your journal" },
      { text: "Open the tracker page at the front." },
      { text: "Tick the End column for every sign you can do without the book." },
    ], { x: 0.5, y: 3.15, w: 9, h: 1.4, strip: C.PRIMARY });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        "SILENT from the door. Journals open, tracker page.",
        "TIME: 5 minutes.",
        ["SAY: Today is a games lesson.", "Every game uses signs you have already learned."],
      ],
      prep: "5 min. Seated and silent, the same as every lesson this term.",
      tag: "[Do Now | Attention, focus and regulation | HITS 6]",
    }));
  }

  // 4. LI and SC
  T.liSlide(
    pres,
    "I am learning to use everything from this term under time pressure.",
    SC,
    T.composeGlanceNotes({
      beats: [
        ["POINT to each criterion.", "SAY: Recognise it, sign it first time, and wait for eyes."],
        "SAY: Points today go to teams that keep their voices off.",
      ],
      prep: "No evidence today; reports are being written. This lesson is retrieval.",
      tag: "[LI and SC | Planning | HITS 1]",
    }),
    FOOTER,
    { numberSC: true, strongHeadings: true, separate: true }
  );

  // 5. I Do: teams, points, the three rules
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "Four teams. Three rules.");
    const rules = ["Voices off during every game", "Eyes on the signer", "Wait for eyes before you start"];
    const rw = (9 - 0.22 * 2) / 3;
    rules.forEach((r, i) => {
      const x = 0.5 + i * (rw + 0.22);
      T.addCard(s, x, 1.45, rw, 1.75, { variant: "tint", tone: C.PRIMARY });
      s.addText(String(i + 1), {
        x, y: 1.55, w: rw, h: 0.42, fontSize: 16, fontFace: T.FONT_B, bold: true,
        color: C.PRIMARY, align: "center", valign: "middle", margin: 0,
      });
      s.addText(r, {
        x: x + 0.12, y: 2.0, w: rw - 0.24, h: 1.05, fontSize: 21, fontFace: T.FONT_H, bold: true,
        color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
      });
    });
    T.addInstructionCard(s, [
      { role: "header", text: "Points" },
      { text: "One point for each correct sign or correct placement." },
      { text: "Two points for a round finished with nobody voicing." },
    ], { x: 0.5, y: 3.4, w: 9, h: 1.3, strip: C.ACCENT });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        ["SAY: Four teams. Points go on the board,", "and I read the total at the end."],
        ["POINT to each rule.", "SAY: Voices off, eyes on the signer, wait for eyes."],
        ["SAY: Two people in each team run the attention moves today:",
          "the wave, the tap, the point. I stay out of it."],
      ],
      prep: "The students run the Deaf-friendly protocols today. Let them.",
      tag: "[I Do | Explicit teaching | HITS 3]",
    }));
  }

  // 6. We Do round 1: Watch and Write as a team round
  gameSlide(pres, "We Do", C.SUCCESS, "Round 1: Watch and Write", [
    "One board per team. Pens down.",
    "A signer from one team signs a word from the term list.",
    "Every other team watches it all, then writes the English.",
    "One, two, three, chin it: every board up at once.",
    "Take turns being the signer. Ten words.",
  ], "A point for every board with the right word.",
  "Nobody writes while the signer is signing.",
  T.composeGlanceNotes({
    answer: "the English word the student signed, from the term list",
    beats: [
      ["SET UP: four teams, one board each.", "Each team picks a signer for the first word."],
      ["POINT at a word on the term list for the signer to sign.",
        "Do not sign it yourself; the signer has learned it this term."],
      "SAY: One, two, three, chin it. Every board comes up at once.",
      "COLLECT: a point per team board with the right word. Then the next signer.",
    ],
    trap: ["writing while the signer is still signing.", "Fix: pens down, the signer signs it again."],
    stretch: "the signer adds a number or a year to the word.",
    help: "the team signs back the word together before writing it.",
    prep: "10 min. Game 2 Watch and Write as a team round, student signers for a relief teacher.",
    tag: "[We Do | Collaborative learning | HITS 6]",
  }), ["voicesOff", "whiteboards"]);

  // 7. We Do round 2: More or Less
  gameSlide(pres, "We Do", C.SUCCESS, "Round 2: More or Less", [
    "Pairs inside your team, facing each other.",
    "Write a hidden number from " + RANGE + " and turn your board over.",
    "Take turns signing a guess. Never write it.",
    "Answer more, less or correct.",
    "First to guess wins the round. Swap who guesses first.",
  ], "A point to your team for every round a pair member wins.",
  "Look down mid-guess and you lose that turn.",
  T.composeGlanceNotes({
    beats: [
      ["SET UP: pairs inside teams, facing.", "Range " + RANGE + " written up where everyone can see it."],
      "SAY: A guess is signed, never written. Boards face down between turns.",
      "TIME: 10 minutes.",
      "COLLECT: each team counts its round wins.",
    ],
    trap: ["writing a guess on the board.", "Fix: board face down, sign the guess again."],
    stretch: "play with years only, range 1880 to 2026.",
    help: "a number line on the desk, so a guess is a point plus a sign.",
    prep: "10 min. Game 5 More or Less. Range " + RANGE + " for this cohort.",
    tag: "[We Do | Collaborative learning | HITS 6]",
  }), ["voicesOff", "partner", "whiteboards"]);

  // 8. You Do: How Many Medals, in pairs
  {
    const s = T.youDoSlide(
      pres,
      "How Many Medals",
      "Pairs: one of you has sheet A, one has sheet B, folder up between you. Take turns asking how many medals your partner has in each sport. Write their numbers, then compare totals.",
      [
        "Ask how many medals, one sport at a time.",
        "Pencil down while they sign. Then write.",
        "Add both totals. Who has more?",
      ],
      T.composeGlanceNotes({
        beats: [
          ["SAY: Pairs. Folder up. The folder hides the sheet, never your face.",
            "Pencils down while your partner signs."],
          "TIME: 10 minutes.",
          ["CIRCULATE: totals are A 77 and B 81.",
            "Last row: who has more sock wrestling medals?"],
          "COLLECT: a point to the team of every pair whose totals are right.",
        ],
        stretch: "work out how many more medals the winner has, and sign that number.",
        help: "a number line on the desk for the bigger numbers.",
        prep: "10 min. Game 7 How Many Medals. Cut this game if you are running late.",
        tag: "[You Do | Collaborative learning | HITS 6]",
      }),
      FOOTER,
      { where: "Voices off  |  Pairs  |  Medal tally sheets" }
    );
    T.addCueStrip(s, ["voicesOff", "partner"]);
  }

  // 9. Points totalled and journal line
  T.exitTicketSlide(
    pres,
    [
      "Points totalled: each team reads its total to the class.",
      "Rate yourself on the three I can statements: just starting, getting there, got it.",
      "Write one line: the sign your team needed most today.",
    ],
    T.composeGlanceNotes({
      beats: [
        "SAY: Team totals, one team at a time. I write them on the board.",
        "SAY: Journals: rate yourself, then one line about the sign your team needed most.",
        "COLLECT nothing. Leave the Auslan teacher a note about which signs were hard.",
      ],
      prep: "5 min. No evidence today. The note is the most useful thing you can leave.",
      tag: "[Exit ticket | Evaluating impact | HITS 7]",
    }),
    FOOTER,
    { title: "Points, then your journal" }
  );

  // 10. Closing
  T.closingSlide(
    pres,
    {
      reflectionPrompt: "Turn and tell your partner the sign you got right first time today.",
      scItems: SC,
      selfAssessment: ["Just starting", "Getting there", "Got it"],
    },
    T.composeGlanceNotes({
      beats: [
        "SAY: Tell your partner the sign you got right first time today.",
        "SAY: That is the end of the games for this term.",
      ],
      prep: "If running late, cut the third game. The team rounds are the lesson.",
      tag: "[Closing | Planning | HITS 9]",
    })
  );

  return pres;
}

const pres = build();
fs.mkdirSync(OUT_DIR, { recursive: true });
const file = path.join(OUT_DIR, "DeafSport Enrichment Review Week 9.pptx");
pres.writeFile({ fileName: file }).then(() => console.log("PPTX written to " + file));
