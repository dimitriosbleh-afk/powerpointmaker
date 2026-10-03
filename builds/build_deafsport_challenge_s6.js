"use strict";

/**
 * Deaf Sports in Australia | Challenge (Years 3 and 4) | Lesson 6 Lights, flags and what changed
 *
 * Built from the unit document section 9.6. A build script rather than a lesson
 * spec for the reasons given in build_deafsport_enrichment_s2.js. The medal
 * tally sheets are shared with the Challenge deck and the review week, so they
 * live in builds/auslan_lib.js.
 *
 * Every modification is taught as a design problem somebody fixed, never as
 * something an athlete had to overcome (unit document care note).
 *
 * Challenge calibration (unit document 3 and 9.6): name one change and what it
 * replaced, with the three picture prompts on the You Do slide, and no cold-call
 * follow-up after the boards.
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme, weekToVariant } = require("../themes/factory");
const P = require("../themes/pdf_helpers");
const A = require("./auslan_lib");

const T = createTheme("literacy", "grade34", weekToVariant(1));
const C = T.C;
const FOOTER = "Deaf Sports in Australia  |  Session 6 of 8  |  Years 3-4 Auslan";
const OUT_DIR = path.join("output", "DeafSport_Challenge_S6_Lights_Flags_And_What_Changed");
const RES_DIR = path.join(OUT_DIR, "resources-session6");
const TALLY_PDF = "Session 6 Medal Tally Sheets A and B.pdf";

const report = A.createSignReport();

function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  T.titleSlide(
    pres,
    "Lights, flags and what changed",
    "Deaf Sports in Australia",
    "Lesson 6 of 8  |  Years 3 and 4  |  Term 4",
    "Lesson 6 of 8. How Deaf sport turns what you were meant to hear into something you can see."
  );

  // 2. Teacher Resources
  P.addResourceSlide(pres, [
    {
      name: "Session 6 Medal Tally Sheets A and B",
      fileName: "resources-session6/" + TALLY_PDF,
      note: "One per student, half A and half B. Consumable. Plus one propped folder per pair.",
    },
  ], T, FOOTER, T.composeGlanceNotes({
    beats: [
      ["SHOW while students settle.", "SAY: Sheets stay face down until I say. Folders up between partners."],
      "COLLECT the sheets at the end; they go in the journal as a photo.",
    ],
    prep: [
      "The only large print run in the unit. The review week uses the sheets again.",
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
    s.addText("How do you know a race has started?", {
      x: 0.7, y: 1.45, w: 8.6, h: 1.5, fontSize: 36, fontFace: T.FONT_H, bold: true,
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
        ["Retrieval sits inside the prompt: SPORT and TEAM from Lesson 2,", "and the time signs from Lesson 5."],
        "TIME: 5 minutes. Deal with whatever walked in from the playground.",
        ["SAY: Every answer you wrote is a sound or a sight. Today you find out",
          "what happens when the sport decides it will not rely on the sound."],
      ],
      prep: "5 min. Most answers arrive in an order, which is why the time signs come back.",
      tag: "[Do Now | Attention, focus and regulation | HITS 6]",
    }));
  }

  // 4. LI and SC
  T.liSlide(
    pres,
    "I am learning to explain how Deaf sport is set up so everything you need to know can be seen.",
    [
      "I can name one thing in sport that is shown with a light or a flag.",
      "I can explain what problem the light or flag solved.",
      "I can give one more example of a barrier somebody designed away.",
    ],
    T.composeGlanceNotes({
      beats: [
        ["POINT to each criterion.", "SAY: Name it, say what problem it solved, then find another one."],
        "SAY: By the end, you can explain one change and why somebody made it.",
      ],
      prep: "Criterion 2 is only assessable in the You Do explanation. Never cut it.",
      tag: "[LI and SC | Planning | HITS 1]",
    }),
    FOOTER,
    { numberSC: true, strongHeadings: true, separate: true }
  );

  // 5. I Do: start, light, flag
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Three signs for a race you can see");
    A.addSignCardRow(T, s, [
      { gloss: "START", meaning: "start, begin" },
      { gloss: "LIGHT", meaning: "light" },
      { gloss: "FLAG", meaning: "flag" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "START, LIGHT, FLAG",
      beats: [
        ["SAY: Same three moves. Eyes up, watch, copy.", "MODEL each twice, slowly then at pace. They copy once."],
        ["SAY: Deaf athletes play every sport with very little change.",
          "Every change does one job: it turns something you hear into something you see."],
        ["MODEL LIGHT low and facing you, like a starter light in a pool.", "Flick the fingers for a flashing light."],
      ],
      trap: ["signing LIGHT as a ceiling lamp.", "Fix: move it to where the light really is, student redoes it."],
      stretch: "sign where a light would be at the start of a running race.",
      help: "copy beside a partner who has it, then on your own.",
      prep: [
        "Chris: START includes both forms on its Signbank page. LIGHT moves to where the light is.",
        "Source: Deaf Sports Australia, what changes for Deaf athletes.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 6. I Do: three real modifications
  {
    const s = T.choiceSlide(
      pres,
      "I Do",
      "Three things somebody changed",
      "Each one turns a sound into something you can see.",
      [
        { visual: { type: "pictogram", name: "lightbulb" }, text: "Pool light" },
        { visual: { type: "pictogram", name: "flag" }, text: "Flag or wave" },
        { visual: { type: "pictogram", name: "tv" }, text: "Scoreboard" },
      ],
      T.composeGlanceNotes({
        answer: "pool light, flag or wave, scoreboard",
        beats: [
          ["SAY: In swimming, a light beside the pool is wired to the gun.",
            "No light? The starter raises an arm and drops it."],
          ["SAY: A referee gets your attention with a flag, or by waving.",
            "A captain taps the referee on the shoulder. That is normal, not rude."],
          "SAY: Scores go on a board or a screen, so nobody waits to be told.",
          ["GET IT WRONG: SAY: I nearly told you Deaf athletes need special sports.",
            "Same sports, same rules. The equipment got redesigned."],
        ],
        trap: ["saying the athlete had to overcome something.", "Fix: name what got changed, student restates it."],
        stretch: "name what each one replaced: the gun, the whistle, the announcer.",
        help: "point to the picture, then sign LIGHT or FLAG.",
        care: "a design problem somebody fixed, never something an athlete overcame.",
        prep: "Protocol this week: the shoulder tap. It comes straight out of this content.",
        sources: "Deaf Sports Australia, deafsports.org.au",
        tag: "[I Do | Explicit teaching | HITS 3]",
      }),
      FOOTER,
      { letters: false }
    );
    T.addCueStrip(s, ["eyesUp"]);
  }

  // 7. I Do: change, community
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Not the athlete. The equipment changed.");
    A.addSignCardRow(T, s, [
      { gloss: "CHANGE", meaning: "change" },
      { gloss: "COMMUNITY", meaning: "community" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "CHANGE, COMMUNITY",
      beats: [
        "MODEL each twice; they copy once.",
        ["SAY: Notice what changed in every one. Not the athlete.",
          "The gun changed, the referee changed, the scoreboard changed."],
        "SAY: The community decided what to change, and then changed it.",
        ["ASK: what changed at the pool. Thumbs up when you can name it.",
          "5 sec, voices off. EXPECT: most thumbs up fast"],
      ],
      trap: ["CHANGE used to mean the athlete changed.", "Fix: ask what got changed, student re-signs it."],
      stretch: "sign one sentence with CHANGE and COMMUNITY in it.",
      help: "the three pictures from the last slide on the board as prompts.",
      prep: "COMMUNITY is the first sign to drop if the lesson runs heavy.",
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 8. We Do: Game 7 How Many Medals
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.SUCCESS);
    T.addBadge(s, "We Do", { color: C.SUCCESS });
    T.addCueStrip(s, ["voicesOff", "partner"]);
    T.addTitle(s, "Game 7 How Many Medals");
    T.addInstructionCard(s, [
      { role: "header", text: "With your partner" },
      { text: "One of you has sheet A, one has sheet B. Folder up between you." },
      { text: "Take turns: ask how many medals they have in each sport." },
      { text: "Write their number next to that sport." },
      { text: "Add up both totals. The bigger total wins." },
    ], { x: 0.5, y: 1.4, w: 5.9, h: 2.5, strip: C.SUCCESS });
    T.addCard(s, 6.65, 1.4, 2.85, 2.5, { variant: "tint", tone: C.ALERT });
    s.addText("Pencil down while they sign.", {
      x: 6.8, y: 1.6, w: 2.55, h: 0.9, fontSize: 19, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    s.addText("Watch the whole answer, then write.", {
      x: 6.8, y: 2.55, w: 2.55, h: 1.2, fontSize: 15, fontFace: T.FONT_B,
      color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
    });
    s.addText("The folder hides the sheet, never your face.", {
      x: 0.5, y: 4.15, w: 9, h: 0.5, fontSize: 16, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "totals: sheet A 77, sheet B 81",
      beats: [
        ["SET UP: pairs facing, sheet flat behind a propped folder.", "HOW-MANY from Lesson 4 does the asking."],
        ["CIRCULATE: watch for pencils moving while a partner signs.",
          "A pair who writes while watching will be a number out in three turns."],
        ["ASK at the end: who has more sock wrestling medals.", "It is the bit they remember."],
        "COLLECT: photo of the finished sheet into the journal.",
      ],
      trap: ["writing while the partner is still signing.", "Fix: pencils down, partner re-signs, student writes after."],
      stretch: "work out how many more medals the winner has, and sign that number.",
      help: "a number line on the desk for the bigger numbers.",
      prep: "10 min. Game 7 How Many Medals. Netball replaces cricket, from Chris's twelve sports.",
      tag: "[We Do | Collaborative learning | HITS 5]",
    }));
  }

  // 9. Primary decision point
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.ASSESS);
    T.addBadge(s, "Check it", { color: C.ASSESS });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "What starts the race?");
    T.addCard(s, 0.5, 1.5, 9, 1.6, { variant: "tint", tone: C.ASSESS });
    s.addText("What starts a race in Deaf swimming?", {
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
      answer: "a light, wired to the starting gun",
      beats: [
        ["ASK: what starts a race in Deaf swimming.",
          "5 sec think time. Cue: Everyone signs it to me on three. One, two, three.",
          "EXPECT: LIGHT"],
        ["SCAN the room, back row first.",
          "80%+ -> move to You Do.",
          "Less -> three problems left, three fixes right. Pair them on boards,",
          "explain one pairing to a partner, then re-ask."],
      ],
      trap: ["LIGHT as a word with no problem attached.", "Fix: the matching task, then re-ask."],
      prep: "The decision point that decides whether the You Do explanation will land.",
      tag: "[CFU | Evaluating impact | HITS 7]",
    }));
  }

  // 10. You Do: Mix and Mingle, then explain
  {
    const s = T.youDoSlide(
      pres,
      "Find someone who picked a different change",
      "Find someone who picked a different change. Sit together. Sign your change and what it replaced. Then swap.",
      [
        "Mix and mingle: find a different change.",
        "Sign the change and what it replaced.",
        "Partner signs it back to check.",
      ],
      T.composeGlanceNotes({
        beats: [
          ["SAY: Voices off. Three minutes to find your person.",
            "Different from the game: that swapped numbers, this swaps an explanation."],
          "TIME: advance to start the clock. Then sit and explain, both ways.",
          ["CIRCULATE: listen for the problem, not just the change.", "Criterion 2 lives here."],
          "COLLECT: each pair's change and what it replaced, on the board.",
        ],
        stretch: "name one more barrier somebody designed away, in sport or anywhere.",
        help: "point to a picture first, then sign the change.",
        prep: [
          "10 min. Game 3 Mix and Mingle, then the partner explanation.",
          "Challenge explains one change with the pictures up. Time 3 minutes on the class timer.",
        ],
        tag: "[You Do | Explicit teaching | HITS 10]",
      }),
      FOOTER,
      {
        where: "Voices off  |  3 minutes  |  Then sit together",
        visual: { type: "pictograms", items: [{ name: "lightbulb", label: "light" }, { name: "flag", label: "flag" }, { name: "tv", label: "scores" }] },
      }
    );
    T.addCueStrip(s, ["voicesOff", "partner", "timer"]);
  }

  // 11. Exit ticket
  T.exitTicketSlide(
    pres,
    [
      "Rate yourself on the three I can statements: just starting, getting there, got it.",
      "Write one line: the change you explained and what it replaced.",
      "Cultural question: what is one thing at our school that could be designed so it can be seen instead of heard?",
    ],
    T.composeGlanceNotes({
      beats: [
        "COLLECT nothing. Everyone writes in the journal while you call the rotation.",
        ["CALL this week's group, one at a time.", "They sign LIGHT, CHANGE and one sign from Lesson 5."],
        "Then they sign one change and the problem it solved.",
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
      reflectionPrompt: "Turn and tell your partner one thing at school that could be seen instead of heard.",
      scItems: [
        "I can name one thing in sport that is shown with a light or a flag.",
        "I can explain what problem the light or flag solved.",
        "I can give one more example of a barrier somebody designed away.",
      ],
      selfAssessment: ["Just starting", "Getting there", "Got it"],
    },
    T.composeGlanceNotes({
      beats: [
        "SAY: Tell your partner one thing at school that could be seen instead of heard.",
        ["SAY: Next week you put it all together", "and interview somebody properly."],
      ],
      care: "athlete called amazing for competing? Redirect to what they did in the pool.",
      prep: "Protocol practised today: the shoulder tap. Required from here on.",
      tag: "[Closing | Planning | HITS 9]",
    })
  );

  return pres;
}

A.buildMedalTallyPdf(path.join(RES_DIR, TALLY_PDF), "Session 6  |  Medal Tally Sheets  |  Years 3-4 Auslan", C.PRIMARY);
const pres = build();
const file = path.join(OUT_DIR, "DeafSport Challenge Session 6 Lights Flags And What Changed.pptx");
pres.writeFile({ fileName: file }).then(() => {
  console.log("PPTX written to " + file);
  report.print();
  const lines = report.rehearsalLines();
  if (lines.length) {
    console.log("\nREHEARSE FIRST");
    lines.forEach((l) => console.log("  " + l));
  }
});
