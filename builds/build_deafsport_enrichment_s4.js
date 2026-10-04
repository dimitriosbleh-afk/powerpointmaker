"use strict";

/**
 * Deaf Sports in Australia | Enrichment (Years 5 and 6) | Lesson 4 When and how many
 *
 * Built from the unit document section 9.4. A build script rather than a lesson
 * spec for the reasons given in build_deafsport_enrichment_s2.js. Shared parts
 * live in builds/auslan_lib.js.
 *
 * How a four-digit year is produced is Chris's answer to row 53 of his sign
 * sheet (3 Oct 2026), not a guess: the first half across the body to the
 * non-dominant side, the second half placed beside it to the dominant side.
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme, weekToVariant } = require("../themes/factory");
const A = require("./auslan_lib");

const T = createTheme("literacy", "grade56", weekToVariant(1));
const C = T.C;
const FOOTER = "Deaf Sports in Australia  |  Session 4 of 8  |  Years 5-6 Auslan";
const OUT_DIR = "DeafSport_Enrichment_S4_When_And_How_Many";
const RANGE = "1 to 100";

const report = A.createSignReport();

/** Four numbered step cards in a row, the routine strip used on game slides. */
function stepRow(s, steps, y, tone) {
  const stw = (9 - 0.18 * (steps.length - 1)) / steps.length;
  steps.forEach((step, i) => {
    const x = 0.5 + i * (stw + 0.18);
    T.addCard(s, x, y, stw, 1.05, { variant: "outline", tone });
    s.addText(String(i + 1), {
      x, y: y + 0.08, w: stw, h: 0.34, fontSize: 13, fontFace: T.FONT_B, bold: true,
      color: tone, align: "center", valign: "middle", margin: 0,
    });
    s.addText(step, {
      x: x + 0.06, y: y + 0.42, w: stw - 0.12, h: 0.55, fontSize: 15, fontFace: T.FONT_B,
      bold: true, color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
    });
  });
}

function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  T.titleSlide(
    pres,
    "When and how many",
    "Deaf Sports in Australia",
    "Lesson 4 of 8  |  Years 5 and 6  |  Term 4",
    "Lesson 4 of 8. Years and numbers, so Lesson 5's timeline can run."
  );

  // 2. Teacher Resources
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "Teacher Resources", { color: C.PRIMARY, w: 2.6 });
    T.addTitle(s, "Nothing to print for this lesson");
    T.addInstructionCard(s, [
      { role: "header", text: "Before the lesson" },
      { text: "A mini whiteboard and marker for every student." },
      { text: "The number range, " + RANGE + ", written up where everyone can see it." },
      { text: "Journals, and one Sign It! per pair." },
    ], { x: 0.5, y: 1.45, w: 9, h: 1.95, strip: C.PRIMARY });
    s.addText(A.ATTRIBUTION, {
      x: 0.5, y: 3.65, w: 9, h: 0.8, fontSize: 11, fontFace: T.FONT_B,
      color: C.MUTED, align: "left", valign: "top", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        ["SHOW while students settle.", "SAY: A board and a marker each. Pens down until I say."],
        "COLLECT nothing today. Boards are the only equipment.",
      ],
      prep: [
        "Rehearse a four-digit year before class: 19 across to your non-dominant side, 65 beside it.",
        A.ATTRIBUTION,
      ],
      tag: "[Setup | Planning | HITS 2]",
    }));
  }

  // 3. Do Now: Auslan90 Australian Deaf Games, Day 4
  A.addVideoDoNow(T, pres, {
    day: 4,
    footer: FOOTER,
    bridge: ["SAY: The youngest and oldest competitors were born years apart.", "Today you learn to sign the years."],
  });

  // 4. LI and SC
  T.liSlide(
    pres,
    "I am learning to ask when something happened and how many, using years and numbers.",
    [
      "I can sign a number in our range.",
      "I can ask when and how many, and read a year in an answer.",
      "I can compare two medal totals and say which is bigger.",
    ],
    T.composeGlanceNotes({
      beats: [
        ["POINT to each criterion.", "SAY: Numbers first, then questions with numbers in the answer."],
        "SAY: By the end, you can catch a year somebody signs and write it down.",
      ],
      prep: "Criterion 2 is what the exit rotation collects. Range today: " + RANGE + ".",
      tag: "[LI and SC | Planning | HITS 1]",
    }),
    FOOTER,
    { numberSC: true, strongHeadings: true, separate: true }
  );

  // 5. I Do: three question and number signs
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Three signs for dates and numbers");
    A.addSignCardRow(T, s, [
      { gloss: "WHEN", meaning: "when" },
      { gloss: "HOW-MANY", meaning: "how many" },
      { gloss: "YEAR", meaning: "year" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "WHEN, HOW-MANY, YEAR",
      beats: [
        ["SAY: Same three moves. Eyes up, watch, copy.",
          "MODEL each sign twice, slowly then at pace. They copy once."],
        ["SAY: With a number question I set up what I am counting first.",
          "Otherwise they get the how many and have no idea what of."],
        ["POINT at WHEN and HOW-MANY side by side.", "Same handshape and movement. Watch where on the face each one sits."],
      ],
      trap: ["WHEN and HOW-MANY blurring together.", "Fix: sign the two in a pair, student names which, then redoes it."],
      stretch: "ask a partner how many years they have played, setting up the sport first.",
      help: "the three meanings written on the board beside where you stand.",
      prep: [
        "3 of today's 5 production signs. Chris: WHEN, HOW-MANY, HOW-OLD share handshape and movement.",
        "YEAR is the letter Y, borrowed from English, like MONTH and the days.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 6. I Do: medal and win
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Two more: medal and win");
    A.addSignCardRow(T, s, [
      { gloss: "MEDAL", meaning: "medal" },
      { gloss: "WIN", meaning: "win" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "MEDAL, WIN",
      beats: [
        "MODEL each twice; they copy once.",
        ["GET IT WRONG: sign a year, then say you caught the number", "but not what it was the year of."],
        ["SAY: AGAIN would get me the same year, and I already had the year.",
          "WHAT MEAN? gets me the explanation I actually needed."],
      ],
      trap: ["signing WIN as a celebration, both arms up.", "Fix: show the sporting sign again, student redoes it."],
      stretch: "sign a sentence with a medal and a year in it.",
      help: "copy MEDAL and WIN beside a partner who has it, then on your own.",
      prep: "WIN is Chris's vetted entry, the sporting result. Signbank's first entry is a celebration.",
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 7. I Do: three years to hold onto
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Three years to hold onto this term");
    ["1965", "2005", "2026"].forEach((yr, i) => {
      T.addTextOnShape(
        s,
        yr,
        { x: 0.5 + i * 3.1, y: 1.5, w: 2.8, h: 1.5, rectRadius: 0.14,
          fill: { color: C.PRIMARY_SOFT }, line: { color: C.PRIMARY, width: 1 } },
        { fontSize: 48, fontFace: T.FONT_H, bold: true, color: C.CHARCOAL,
          align: "center", valign: "middle", margin: 0 }
      );
    });
    T.addInstructionCard(s, [
      { role: "header", text: "A year is signed in two halves" },
      { text: "The first half, then the second half placed beside it." },
      { text: "Watch for the pause between them." },
    ], { x: 0.5, y: 3.25, w: 9, h: 1.4, strip: C.PRIMARY });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "1965, 2005, 2026",
      beats: [
        ["MODEL each year: 19 across to your non-dominant side,", "then 65 placed beside it to your dominant side."],
        ["SAY: Three years to hold onto this term. Nineteen sixty five,",
          "two thousand and five, and this year. You meet all three again next week."],
        ["ASK: which year did I sign. Sign one of the three.",
          "5 sec. Cue: Write it. Hover it. Chin it on three. One, two, three.",
          "EXPECT: the year you signed"],
      ],
      trap: ["catching nineteen and losing sixty five.", "Fix: sign each half with a clear pause, student writes both halves."],
      stretch: "sign the year you were born, in two halves.",
      help: "19 and 65 written in two boxes on the board, signed box by box.",
      prep: [
        "Enrichment produces years. Left-handers reverse the sides.",
        "Chris's row 53 answer. Do not improvise it in front of them.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 8. I Do: gold, silver, bronze, receptive only
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "Gold, silver, bronze: just recognise them");
    A.addSignCardRow(T, s, [
      { gloss: "GOLD", meaning: "gold" },
      { gloss: "SILVER", meaning: "silver" },
      { gloss: "BRONZE", meaning: "bronze" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        "SHOW each once, in that order. Nobody copies these.",
        ["SAY: You only need to know these when you see them.", "Nobody asks you to sign them."],
        ["POINT at the links.", "They open the signs Chris chose, for anyone who wants them at home."],
      ],
      prep: [
        "Receptive only. First to drop if the lesson runs long.",
        "Chris: older signers fingerspell bronze; younger signers use this sign, which also means brown.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3]",
    }));
  }

  // 9. We Do: Game 2 Watch and Write
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.SUCCESS);
    T.addBadge(s, "We Do", { color: C.SUCCESS });
    T.addCueStrip(s, ["eyesUp", "whiteboards"]);
    T.addTitle(s, "Game 2 Watch and Write");
    T.addCard(s, 0.5, 1.45, 9, 1.3, { variant: "tint", tone: C.SUCCESS });
    s.addText("I sign a number, a year or a medal total. You write it.", {
      x: 0.7, y: 1.45, w: 8.6, h: 1.3, fontSize: 28, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    stepRow(s, ["Pens down", "Watch all of it", "Write it", "Chin it on three"], 2.95, C.SUCCESS);
    s.addText("Ten rounds. Nobody writes while I am signing.", {
      x: 0.5, y: 4.2, w: 9, h: 0.5, fontSize: 18, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "whatever you signed, digits only",
      beats: [
        ["SIGN one item. Pens down while you sign.",
          "Cue: One, two, three, chin it. Every board comes up at once."],
        ["Ten items: numbers to 100, the three years,", "then two medal totals. ASK which total is bigger."],
        ["SAY: You watched the whole number before writing.", "That is why you got it."],
        "RE-SIGN an item once if the room splits. Never more than once.",
      ],
      trap: ["writing while you are still signing.", "Fix: pens down, re-sign, student writes after."],
      stretch: "write the English sentence the number belongs in, not just the digits.",
      help: "a number line on the desk, so they find it before writing.",
      prep: "10 min. Game 2 Watch and Write. Receptive: this is the bar for criterion 2.",
      tag: "[We Do | Collaborative learning | HITS 6]",
    }));
  }

  // 10. Primary decision point
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.ASSESS);
    T.addBadge(s, "Check it", { color: C.ASSESS });
    T.addCueStrip(s, ["eyesUp", "whiteboards"]);
    T.addTitle(s, "Watch the year. Write the digits.");
    T.addCard(s, 0.5, 1.5, 9, 1.4, { variant: "tint", tone: C.ASSESS });
    s.addText("I sign one year. You write it in digits.", {
      x: 0.7, y: 1.5, w: 8.6, h: 1.4, fontSize: 32, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    stepRow(s, ["Watch both halves", "Write the digits", "Chin it on three", "One of you signs it back"], 3.08, C.ASSESS);
    T.addTextOnShape(
      s,
      "Write it. Hover it. Chin it on three. One, two, three.",
      { x: 1.6, y: 4.3, w: 6.8, h: 0.62, rectRadius: 0.12, fill: { color: C.ASSESS } },
      { fontSize: 18, fontFace: T.FONT_B, bold: true, color: C.WHITE,
        align: "center", valign: "middle", margin: 0 }
    );
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "1965",
      beats: [
        ["ASK: what year is this. Sign 1965.",
          "5 sec think time. Cue: Write it. Hover it. Chin it on three. One, two, three.",
          "EXPECT: 1965"],
        ["SCAN the room, back row first.",
          "80%+ -> cold-call one student to sign it back, then go to You Do.",
          "Less -> 19 and 65 in two boxes, sign each box, then join.", "Re-check with 2005."],
        ["FOLLOW UP the student who signed back.", "SAY: How did you know it was sixty five and not fifty six?"],
      ],
      trap: ["the two halves fusing, so they catch nineteen and lose sixty five.", "Fix: boxes, then re-check."],
      prep: [
        "The decision point that decides whether Lesson 5's timeline is safe.",
        "Hover then chin is the school routine. Nobody shows early.",
      ],
      tag: "[CFU | Evaluating impact | HITS 7]",
    }));
  }

  // 11. You Do setup: three signs for the game
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.SUCCESS);
    T.addBadge(s, "You Do", { color: C.SUCCESS });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Three signs for More or Less");
    A.addSignCardRow(T, s, [
      { gloss: "MORE", meaning: "more" },
      { gloss: "LESS", meaning: "less" },
      { gloss: "CORRECT", meaning: "correct" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "MORE, LESS, CORRECT",
      beats: [
        "MODEL each once. They copy once. These are game signs, not today's five.",
        ["SAY: Your partner guesses. You answer more, less or correct.", "Nothing else, no voice."],
      ],
      prep: [
        "Chris has not vetted these three. Rehearse them from Signbank before class.",
        "CORRECT is pinned to the right-answer sense; the default entry is the written tick.",
      ],
      tag: "[You Do | Explicit teaching | HITS 3]",
    }));
  }

  // 12. You Do: More or Less, then a real question
  {
    const s = T.youDoSlide(
      pres,
      "More or Less, then a real question",
      "Write a hidden number from " + RANGE + " and turn your board over. Take turns signing guesses. Then ask your partner how many years they have played their sport.",
      [
        "Sign a guess. Never write it.",
        "Answer more, less or correct.",
        "Ask the real question. Digits on the board.",
      ],
      T.composeGlanceNotes({
        beats: [
          ["SAY: Boards face down between turns.", "A guess is signed, never written."],
          "TIME: 10 minutes. Swap who guesses first each round.",
          ["CIRCULATE: look for eyes held through a whole guess.",
            "A player who looks down mid-guess forfeits that turn."],
          "COLLECT: both answers to the real question on the board, digits only.",
        ],
        stretch: "play with years only, range 1880 to 2026.",
        help: "a number line on the desk, so a guess is a point plus a sign.",
        prep: [
          "10 min. Game 5 More or Less: different content from the We Do, which was receptive.",
          "Protocol today: eye gaze held through a whole answer. Name it, then require it.",
        ],
        tag: "[You Do | Explicit teaching | HITS 10]",
      }),
      FOOTER,
      { where: "Voices off  |  Facing your partner  |  Range " + RANGE }
    );
    T.addCueStrip(s, ["voicesOff", "partner", "whiteboards"]);
  }

  // 13. Exit ticket
  T.exitTicketSlide(
    pres,
    [
      "Rate yourself on the three I can statements: just starting, getting there, got it.",
      "Write one line: the biggest number you signed today and what it counted.",
      "Cultural question: the Australian Deaf Games happen every few years and Deaf people run them. Why might that matter more than who wins?",
    ],
    T.composeGlanceNotes({
      beats: [
        "COLLECT nothing. Everyone writes in the journal while you call the rotation.",
        ["CALL this week's group, one at a time.", "They sign WHEN, HOW-MANY and one sign from Lesson 3."],
        "Then you sign a year and they write it.",
        ["RECORD E, C or P on the tracker,", "against the checklist row, not against the lesson."],
      ],
      prep: [
        "5 min. Evidence piece 3. Students are never told the schedule.",
        "Enrichment produces the year; listen for both halves with the pause.",
      ],
      tag: "[Exit ticket | Evaluating impact | HITS 7]",
    }),
    FOOTER,
    { title: "In your journal" }
  );

  // 14. Closing
  T.closingSlide(
    pres,
    {
      reflectionPrompt: "Turn and tell your partner one year you can now catch when someone signs it.",
      scItems: [
        "I can sign a number in our range.",
        "I can ask when and how many, and read a year in an answer.",
        "I can compare two medal totals and say which is bigger.",
      ],
      selfAssessment: ["Just starting", "Getting there", "Got it"],
    },
    T.composeGlanceNotes({
      beats: [
        "SAY: Tell your partner one year you can catch now.",
        ["SAY: Next week those years go in order.", "You will meet 1965 and 2005 again."],
      ],
      prep: [
        "If the decision point fell under 80%, open Lesson 5 with two minutes of years.",
        "Never cut the year practice; Lesson 5 cannot run without it.",
      ],
      tag: "[Closing | Planning | HITS 9]",
    })
  );

  return pres;
}

const pres = build();
fs.mkdirSync(path.join("output", OUT_DIR), { recursive: true });
const file = path.join("output", OUT_DIR, "DeafSport Enrichment Session 4 When And How Many.pptx");
pres.writeFile({ fileName: file }).then(() => {
  console.log("PPTX written to " + file);
  report.print();
  const lines = report.rehearsalLines();
  if (lines.length) {
    console.log("\nREHEARSE FIRST");
    lines.forEach((l) => console.log("  " + l));
  }
});
