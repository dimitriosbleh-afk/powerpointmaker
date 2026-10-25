"use strict";

/**
 * Deaf Sports in Australia | Challenge (Years 3 and 4) | Lesson 6 Showing the story
 *
 * Moved from Lesson 8 to Lesson 6 on 4 Oct 2026: Chris will run the unit into
 * Week 8 only if nothing summative happens that week, and this lesson holds
 * evidence pieces 2 and 3 for the storytelling rows.
 *
 * Built from the unit document sections 9.8, 11.6 and 13.2. Evidence piece 2,
 * the receptive test, runs in the exit slot from a numbered frame on screen.
 * A build script rather than a lesson spec for the reasons given in
 * build_deafsport_enrichment_s2.js. The story strip is shared with the
 * Challenge deck, so it lives in builds/auslan_lib.js.
 *
 * Depicting signs, constructed action and signing space are modelled from
 * Sign It! and the school reference. Nothing here describes how they are
 * produced (CHECK GRAMMAR in the unit document).
 *
 * Challenge calibration (unit document 3, 9.8 and 13.2): set the scene and
 * retell two events with the picture version on screen, test items 1 to 8
 * only, rated out of eight, and no cold-call follow-up after the boards.
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
const OUT_DIR = path.join("output", "DeafSport_Challenge_S6_Showing_The_Story");
const RES_DIR = path.join(OUT_DIR, "resources-session6");
const STRIP_PDF = "Session 6 Story Strip.pdf";
const TEST_ITEMS = 8;

const report = A.createSignReport();

function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  T.titleSlide(
    pres,
    "Showing the story",
    "Deaf Sports in Australia",
    "Lesson 6 of 8  |  Years 3 and 4  |  Term 4",
    "Lesson 6 of 8. Set the scene, then tell what happened in it. Evidence piece 2 runs at the end."
  );

  // 2. Teacher Resources
  P.addResourceSlide(pres, [
    { name: "Session 6 Story Strip", fileName: "resources-session6/" + STRIP_PDF,
      note: "Fourteen, one per pair. Large text, laminated." },
  ], T, FOOTER, T.composeGlanceNotes({
    beats: [
      ["SHOW while students settle.", "SAY: Exercise books out. Plain paper and a pencil each."],
      "COLLECT the story strips at the end. Nothing else is collected.",
    ],
    prep: [
      "The eight test items and answers are in the notes of the test slide. Rehearse them.",
      A.ATTRIBUTION,
    ],
    tag: "[Setup | Planning | HITS 2]",
  }));

  // 3. Do Now: Auslan90 Australian Deaf Games, Day 1 again
  A.addVideoDoNow(T, pres, {
    day: 6,
    footer: FOOTER,
    bridge: ["SAY: Today's story is a swimming race, like the ones you just watched.", "Today you set the place before you tell a story."],
  });

  // 4. LI and SC
  T.liSlide(
    pres,
    "I am learning to set a scene before I tell what happened in it, and to join the events with time signs.",
    [
      "I can show where a story happens before I start it.",
      "I can retell a two-event story with the scene set first.",
      "I can keep two people in the story apart so my partner knows who is who.",
    ],
    T.composeGlanceNotes({
      beats: [
        ["POINT to each criterion.", "SAY: The place first, then the events, then two people kept apart."],
        "SAY: By the end, you can tell a whole story in Auslan.",
      ],
      prep: "Criterion 2 is what the rotation collects today, during the You Do.",
      tag: "[LI and SC | Planning | HITS 1]",
    }),
    FOOTER,
    { numberSC: true, strongHeadings: true, separate: true }
  );

  // 5. I Do: story, tell
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Two signs for telling a story");
    A.addSignCardRow(T, s, [
      { gloss: "STORY", meaning: "story" },
      { gloss: "TELL", meaning: "tell, inform" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "STORY, TELL",
      beats: [
        ["SAY: Same three moves. In a story the watching does the most work,",
          "because the place and the people are set up before anything happens."],
        "MODEL each twice, slowly then at pace. They copy once.",
        "SAY: A story in Auslan starts with where. Then who. Then what happened.",
      ],
      trap: ["starting with the action before the place.", "Fix: place first, student restarts the story."],
      stretch: "tell a partner one thing that happened today, place first.",
      help: "copy beside a partner who has it, then on your own.",
      prep: [
        "TELL is Chris's INFORM entry, the one for telling someone about an event.",
        "STORY is Chris's vetted entry.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 6. I Do: set the scene from the story strip
  T.textExtractSlide(
    pres,
    "I Do",
    "The Race That Started With A Light",
    A.STORY_STRIP.join("\n"),
    T.composeGlanceNotes({
      answer: "four events: gun raised, light on, both gone, her name at the top",
      beats: [
        ["MODEL the scene before any event: a pool, two swimmers,",
          "one in each lane, kept in two different places."],
        ["MODEL the four events in order, keeping the two swimmers", "where you put them."],
        ["GET IT WRONG: retell two events with both swimmers in one place.",
          "SAY: Nothing wrong with my signs. Everything wrong with where I put them."],
        ["ASK: where was the first swimmer. Thumb toward that side.", "5 sec, voices off. EXPECT: thumbs agree"],
      ],
      trap: ["two people in the same place.", "Fix: point to each swimmer's place, student re-signs."],
      stretch: "find which sentences only describe and which ones happen.",
      help: "the strip on the desk, read in English first.",
      care: "constructed action shows what she did. Caricature? Stop once, name it.",
      prep: "CHECK GRAMMAR: model depicting signs and placement from Sign It!. Do not improvise them.",
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }),
    FOOTER,
    { prompt: "Where does it happen? Who is in it? What actually happens?", fontSize: 19 }
  );

  // 7. I Do: soon, later, next
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Three signs that move a story along");
    A.addSignCardRow(T, s, [
      { gloss: "SOON", meaning: "soon" },
      { gloss: "LATER", meaning: "later" },
      { gloss: "NEXT", meaning: "next" },
    ], { y: 1.4, bottom: 5.05, report });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "SOON, LATER, NEXT",
      beats: [
        "MODEL each twice; they copy once.",
        ["SAY: These move a story along rather than place it in history.",
          "Before, after and back then put you in a year."],
        "SAY: Soon, later and next move you through an afternoon.",
        ["MODEL LATER both ways: two short movements is later.",
          "One long movement is much later, some time in the future."],
      ],
      trap: ["NEXT pointed the wrong way for who is next.", "Fix: show I'm next and you're next, student redoes it."],
      stretch: "join all four events with three different time signs.",
      help: "the three meanings written up beside where you stand.",
      prep: [
        "Chris: NEXT moves toward whoever is next. LATER also has a 7 handshape variant, thumb and finger.",
        "NEXT is Chris's NEXT_2; Signbank's default is a demotion sign.",
      ],
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 8. We Do: Game 10 Watch And Draw It
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.SUCCESS);
    T.addBadge(s, "We Do", { color: C.SUCCESS });
    T.addCueStrip(s, ["voicesOff", "partner"]);
    T.addTitle(s, "Game 10 Watch And Draw It");
    T.addInstructionCard(s, [
      { role: "header", text: "With your partner, one round each way" },
      { text: "A signs the scene only: the place and two people." },
      { text: "B watches it all, pencil down, then sketches it." },
      { text: "Compare. Wrong places? A signs it again." },
      { text: "Swap." },
    ], { x: 0.5, y: 1.4, w: 5.9, h: 2.5, strip: C.SUCCESS });
    T.addCard(s, 6.65, 1.4, 2.85, 2.5, { variant: "tint", tone: C.ALERT });
    s.addText("Pencil down. Watch. Then draw.", {
      x: 6.8, y: 1.6, w: 2.55, h: 0.9, fontSize: 19, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    s.addText("Head down drawing means you missed where they put people.", {
      x: 6.8, y: 2.55, w: 2.55, h: 1.2, fontSize: 15, fontFace: T.FONT_B,
      color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
    });
    s.addText("A warm-up for the retell. Fast.", {
      x: 0.5, y: 4.15, w: 9, h: 0.5, fontSize: 16, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "a sketch with the two people where the signer put them",
      beats: [
        ["SET UP: pairs facing at an angle, paper flat, pencils down.", "One round each way, five minutes."],
        "SAY: Scene only. Nothing happens yet.",
        ["CIRCULATE: watch for heads down while the partner signs.", "That student missed the whole content."],
      ],
      trap: ["explaining the scene in writing.", "Fix: sign it again instead, partner redraws."],
      stretch: "three people in the scene instead of two.",
      help: "the pool scene from the strip, so only the placement is new.",
      prep: "5 min. Game 10 Watch And Draw It. Cut it entirely if you are running late.",
      tag: "[We Do | Collaborative learning | HITS 5]",
    }));
  }

  // 9. Primary decision point
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.ASSESS);
    T.addBadge(s, "Check it", { color: C.ASSESS });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "Where were the two swimmers?");
    T.addCard(s, 0.5, 1.5, 9, 1.6, { variant: "tint", tone: C.ASSESS });
    s.addText("Show me where the two swimmers were placed.", {
      x: 0.7, y: 1.5, w: 8.6, h: 1.6, fontSize: 32, fontFace: T.FONT_H, bold: true,
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
      answer: "two different places, one swimmer each",
      beats: [
        ["ASK: show where the two swimmers were placed.",
          "5 sec think time. Cue: Everyone signs it to me on three. One, two, three.",
          "EXPECT: two clearly separate places"],
        ["SCAN the room, back row first.",
          "80%+ -> You Do with the two-event retell.",
          "Less -> two chairs at the front, a swimmer each. Retell pointing at chairs,",
          "then sign it without them. Re-check."],
      ],
      trap: ["copying handshapes and ignoring placement.", "Fix: the two chairs, then re-check."],
      prep: "The decision point that changes what you report for the retell.",
      tag: "[CFU | Evaluating impact | HITS 7]",
    }));
  }

  // 10. You Do: retell, rotation runs here
  {
    const s = T.youDoSlide(
      pres,
      "Retell the story",
      "Scene first. Then two events, joined with soon, later or next. Keep the two swimmers apart. Use the pictures if you need them. Then swap.",
      [
        "Set the place and the two swimmers.",
        "Two events, joined with a time sign.",
        "Partner signs back the first event.",
      ],
      T.composeGlanceNotes({
        beats: [
          ["SAY: Voices off, pairs at an angle.", "Different from the game: that set a scene, this adds four events."],
          "TIME: 10 minutes, both ways.",
          ["CALL the last rotation group while the rest retell.",
            "They sign STORY, NEXT and one sign from Lesson 5, then set the scene for you."],
          "RECORD E, C or P on the tracker against the checklist row.",
        ],
        stretch: "retell all four events, keeping both swimmers apart.",
        help: "you point to each picture as they sign, so the order is held for them.",
        prep: [
          "10 min. The rotation runs here, because the exit slot is the test.",
          "Challenge: scene plus two events, with the picture version on screen.",
        ],
        tag: "[You Do | Explicit teaching | HITS 10]",
      }),
      FOOTER,
      {
        where: "Voices off  |  Pairs at an angle  |  Story strip on the desk",
        visual: { type: "pictograms", items: [
          { name: "waves", label: "on the blocks" }, { name: "eye", label: "watching" },
          { name: "lightbulb", label: "light on, go" }, { name: "trophy", label: "top of the board" },
        ] },
      }
    );
    T.addCueStrip(s, ["voicesOff", "partner"]);
  }

  // 11. The receptive test frame
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.ASSESS);
    T.addBadge(s, "Exit Ticket", { color: C.ASSESS });
    T.addCueStrip(s, ["eyesUp"]);
    T.addTitle(s, "In your exercise book: number 1 to " + TEST_ITEMS);
    const perCol = Math.ceil(TEST_ITEMS / 2);
    const lineH = 0.62;
    for (let i = 0; i < TEST_ITEMS; i++) {
      const col = Math.floor(i / perCol);
      const row = i % perCol;
      const x = 0.7 + col * 4.6;
      const y = 1.5 + row * lineH;
      s.addText(String(i + 1), {
        x, y, w: 0.5, h: 0.45, fontSize: 22, fontFace: T.FONT_H, bold: true,
        color: C.CHARCOAL, align: "right", valign: "bottom", margin: 0,
      });
      s.addShape("line", { x: x + 0.65, y: y + 0.45, w: 3.4, h: 0, line: { color: C.MUTED, width: 1.25 } });
    }
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "2-8: 1965, 2005, medal, win, team, where, before",
      beats: [
        ["SIGN each item once, then once more. Pens down while you sign.",
          "1 a number in range. 2 1965. 3 2005. 4 MEDAL. 5 WIN."],
        "6 TEAM. 7 WHERE. 8 BEFORE.",
        ["At item 8 stop. Go through all eight together.",
          "Students correct their own in a different colour. Scan for outliers."],
      ],
      prep: [
        "Evidence piece 2, Challenge: items 1 to 8, rated out of eight.",
        "Emerging 0 to 2, Consolidating 3 to 5, Proficient 6 to 8. Nothing marked at home.",
      ],
      tag: "[Exit ticket | Evaluating impact | HITS 7]",
    }));
  }

  // 12. Journal, last two minutes
  T.exitTicketSlide(
    pres,
    [
      "Rate yourself on the three I can statements: just starting, getting there, got it.",
      "On your tracker page, tick every sign you can now do.",
      "Cultural question: what does a Deaf storyteller show you that an English sentence does not?",
    ],
    T.composeGlanceNotes({
      beats: [
        "SAY: Last two minutes. Journal and the tracker page.",
        "COLLECT nothing. The test levels go on the tracker after class.",
      ],
      prep: "Allow 10 minutes per class to record the test levels.",
      tag: "[Exit ticket | Evaluating impact | HITS 7]",
    }),
    FOOTER,
    { title: "In your journal" }
  );

  // 13. Closing
  T.closingSlide(
    pres,
    {
      reflectionPrompt: "Turn and tell your partner the sign from this term you are proudest of.",
      scItems: [
        "I can show where a story happens before I start it.",
        "I can retell a two-event story with the scene set first.",
        "I can keep two people in the story apart so my partner knows who is who.",
      ],
      selfAssessment: ["Just starting", "Getting there", "Got it"],
    },
    T.composeGlanceNotes({
      beats: [
        "SAY: Tell your partner the sign from this term you are proudest of.",
        ["SAY: Next week you put it all together", "and interview somebody properly."],
      ],
      prep: [
        "Protocol practised today: signing space and sight lines.",
        "Never cut the test; evidence piece 2 has no other slot.",
      ],
      tag: "[Closing | Planning | HITS 9]",
    })
  );

  return pres;
}

A.buildStoryStripPdf(path.join(RES_DIR, STRIP_PDF), "Session 6  |  Story Strip  |  Years 3-4 Auslan", C.PRIMARY);
const pres = build();
const file = path.join(OUT_DIR, "DeafSport Challenge Session 6 Showing The Story.pptx");
pres.writeFile({ fileName: file }).then(() => {
  console.log("PPTX written to " + file);
  report.print();
  const lines = report.rehearsalLines();
  if (lines.length) {
    console.log("\nREHEARSE FIRST");
    lines.forEach((l) => console.log("  " + l));
  }
});
