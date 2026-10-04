"use strict";

/**
 * Deaf Sports in Australia | Enrichment (Years 5 and 6) | Review lesson, Week 9 game set
 *
 * Built from the unit document section 9.9. The review lesson runs twice, in
 * weeks 9 and 10, with a different game set each time. Nothing new is taught and
 * no evidence is collected. It is also the CRT lesson: every note is written so
 * a teacher with no Auslan can run it without modelling a sign.
 *
 * Nothing new to print. The Week 9 set reuses the Lesson 5 timeline cards and
 * the Lesson 7 profile cards, named on the resources slide rather than
 * regenerated (AUSLAN_2 section 8). The points table is a slide.
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme, weekToVariant } = require("../themes/factory");

const T = createTheme("literacy", "grade56", weekToVariant(1));
const C = T.C;
const FOOTER = "Deaf Sports in Australia  |  Review, Week 9  |  Years 5-6 Auslan";
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
      { text: "The Lesson 5 Deaf Sport Timeline Cards, all ten in each set." },
      { text: "The Lesson 7 Deaf Athlete Profile Cards." },
      { text: "Sign It!, one per pair, and the student journals." },
      { text: "A cleared floor space for the first game." },
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

  // 6. We Do round 1: Bob Virus as a team round
  gameSlide(pres, "We Do", C.SUCCESS, "Round 1: Bob Virus", [
    "Everyone stands in the space, eyes closed.",
    "The teacher taps one or two people: they are Bobs.",
    "Move around. Ask each person their name. Answer by fingerspelling.",
    "A Bob fingerspells B-O-B. Meet a Bob and you become one.",
    "Two minutes. Then count the Bobs and the rest.",
  ], "A point for every name your team reads back correctly at the end.",
  "No running. Never avoid somebody who comes up to you.",
  T.composeGlanceNotes({
    beats: [
      ["SET UP: clear the furniture. Students face each other close enough", "to read fingerspelling."],
      ["TAP one or two students on the shoulder while eyes are closed.", "They are the Bobs."],
      "TIME: two minutes on a timer everyone can see.",
      ["COLLECT: each team reads back names it saw fingerspelled.", "A point for each correct one."],
    ],
    trap: ["voicing a name instead of fingerspelling it.", "Fix: that pair restarts the exchange, fingerspelling."],
    stretch: "add Doctors: two taps makes a Doctor, who fingerspells D-R and cures Bobs.",
    help: "a secure partner stands with an unsure one and they answer as one.",
    prep: "10 min. Game 1 Bob Virus as a team round. No signs to model: fingerspelling only.",
    tag: "[We Do | Collaborative learning | HITS 6]",
  }), ["voicesOff", "timer"]);

  // 7. We Do round 2: Order The Years
  gameSlide(pres, "We Do", C.SUCCESS, "Round 2: Order The Years", [
    "Teams of four in a circle. Timeline cards face down in the middle.",
    "Take two cards each. Keep them flat on the desk.",
    "Going left, sign one event from a card, with its year.",
    "Your team decides where it goes in the line.",
    "All ten cards in order, then read the line back around the circle.",
  ], "A point for every card in the right place when your team says done.",
  "One signer at a time. Hands off the table while somebody is signing.",
  T.composeGlanceNotes({
    answer: "1880s, 1924, 1954, 1955, 1964, 1965, 1985, 2005, 2011, 2026",
    beats: [
      ["SET UP: teams of four, every face visible.", "All ten cards this time, including the two marked R."],
      "SAY: The turn passes left, so nobody has to be called on.",
      ["CIRCULATE: check each team's line against the answer above.", "A point per card in the right place."],
      "COLLECT every set at the end.",
    ],
    trap: ["ordering by the last two digits of the year.", "Fix: compare the first two digits, team re-places the card."],
    stretch: "play from memory: each player says the event, not reading the card.",
    help: "the team places the 1880s and 2026 cards first, as the two ends.",
    prep: "10 min. Game 6 Order The Years. Ten cards make the team hold a running order.",
    tag: "[We Do | Collaborative learning | HITS 6]",
  }), ["voicesOff", "partner"]);

  // 8. You Do: Which Card Is Mine, in fours
  {
    const s = T.youDoSlide(
      pres,
      "Which Card Is Mine",
      "In fours, all six profile cards face up in the middle. Secretly pick one. Take turns signing its four facts. Everyone points at the end.",
      [
        "Pick a card. Do not say which.",
        "Sign who, what sport, when, what they won.",
        "Point at the end, never during.",
      ],
      T.composeGlanceNotes({
        beats: [
          ["SAY: Fours. Cards in the middle, never in hands.", "Two people can choose the same card."],
          "TIME: 10 minutes.",
          ["CIRCULATE: watch for eyes on the cards mid-profile.", "That student has stopped watching."],
          "COLLECT: a point per correct guess. Every card set back at the end.",
        ],
        stretch: "a secure team plays with the cards face down, from memory.",
        help: "the describer signs two facts, not four.",
        prep: "10 min. Game 8 Which Card Is Mine. Cut this game if you are running late.",
        tag: "[You Do | Collaborative learning | HITS 6]",
      }),
      FOOTER,
      { where: "Voices off  |  Groups of four  |  Profile cards in the middle" }
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
        "SAY: Next week, a different set of games.",
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
