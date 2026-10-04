"use strict";

/**
 * Deaf Sports in Australia | Enrichment (Years 5 and 6) | Lesson 1 Finding the signs you need
 *
 * New lesson, added on Chris's feedback of 17 September 2026. The Do Now asks
 * students to read an English sentence, decide which words actually need a sign,
 * and look those up in Sign It!. None of that is obvious, so it gets taught
 * properly here, once. From 4 Oct 2026 the Do Now is Chris's Deaf Games video,
 * and this routine is how students find any new sign they need.
 *
 * Every example is a sport sentence, so the looking-up practice doubles as first
 * exposure to the term's vocabulary.
 *
 * A build script rather than a lesson spec: animated sign cards, lookup cards,
 * the cue strip and the countdown are not things a spec can express. Shared
 * parts live in builds/auslan_lib.js.
 */

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { createTheme, weekToVariant } = require("../themes/factory");
const A = require("./auslan_lib");

const T = createTheme("literacy", "grade56", weekToVariant(1));
const C = T.C;
const FOOTER = "Deaf Sports in Australia  |  Session 1 of 8  |  Years 5-6 Auslan";
const OUT_DIR = "DeafSport_Enrichment_S1_Finding_The_Signs";

const report = A.createSignReport();

function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  T.titleSlide(
    pres,
    "Finding the signs you need",
    "Deaf Sports in Australia",
    "Lesson 1 of 8  |  Years 5 and 6  |  Term 4",
    "Lesson 1 of 8. One job: teach the looking-up routine they use whenever they need a new sign. Every example is a sport sentence, so they meet the term's vocabulary while they practise the process."
  );

  // 2. Teacher Resources
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "Teacher Resources", { color: C.PRIMARY, w: 2.6 });
    T.addTitle(s, "Nothing to print for this lesson");
    T.addInstructionCard(s, [
      { role: "header", text: "On the desks" },
      { text: "One Sign It! per pair, and the class set of mini whiteboards." },
      { text: "Student journals, open on the tracker page." },
    ], { x: 0.5, y: 1.45, w: 5.7, h: 1.7, strip: C.PRIMARY });
    A.addSignItCover(T, s, { x: 6.45, y: 1.45, w: 3.05, h: 2.6 });
    s.addText(A.ATTRIBUTION, {
      x: 0.5, y: 4.3, w: 9, h: 0.6, fontSize: 11, fontFace: T.FONT_B,
      color: C.MUTED, align: "left", valign: "top", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        ["SHOW while students settle.", "SAY: Sign It! open, one between two, and a board each."],
        "COLLECT nothing today. This lesson is a routine, not a hand-in.",
      ],
      prep: [
        "Copy the journal template into the class Google Classroom before the lesson.",
        A.ATTRIBUTION,
      ],
      tag: "[Setup | Planning | HITS 2]",
    }));
  }

  // 3. Do Now: Auslan90 Australian Deaf Games, Day 1
  A.addVideoDoNow(T, pres, {
    day: 1,
    footer: FOOTER,
    bridge: ["SAY: Everyone in that video was signing about sport.", "Today you start finding the signs you need to do the same."],
  });

  // 4. LI and SC
  T.liSlide(
    pres,
    "I am learning to find which words in my sentence need a sign, and look them up.",
    [
      "I can underline the words in my sentence that need a sign.",
      "I can find one of those words in Sign It! and give its page.",
      "I can sign it to my partner without looking back at the book.",
    ],
    T.composeGlanceNotes({
      beats: [
        ["POINT to each criterion.", "SAY: Three steps. Underline, find, sign."],
        ["SAY: By the end you can do this on your own.", "You will use it whenever you need a new sign."],
      ],
      prep: "Criterion 2 is what the exit rotation collects. Criterion 3 is the stretch.",
      tag: "[LI and SC | Planning | HITS 1]",
    }),
    FOOTER,
    { numberSC: true, strongHeadings: true, separate: true }
  );

  // 5. Launch: the sentence the lesson works on
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "Launch", { color: C.PRIMARY });
    T.addCueStrip(s, ["voicesOff", "whiteboards"]);
    T.addTitle(s, "Write one sentence about sport");
    T.addCard(s, 0.5, 1.45, 9, 1.5, { variant: "tint", tone: C.PRIMARY });
    s.addText("What sport do you like, and who do you play it with?", {
      x: 0.7, y: 1.45, w: 8.6, h: 1.5, fontSize: 32, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addInstructionCard(s, [
      { role: "header", text: "On your whiteboard" },
      { text: "Write your answer in English. One or two sentences." },
      { text: "That is all for now. We work on it together next." },
    ], { x: 0.5, y: 3.15, w: 9, h: 1.35, strip: C.PRIMARY });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      beats: [
        "SAY: Boards out. English only, no signs yet.",
        ["SAY: Write it in English. Do not worry about signing it yet.", "That is the lesson."],
        "TIME: 3 minutes.",
        ["SAY: Every word in that sentence is not a sign.", "Today you work out which ones are."],
      ],
      prep: "3 min. Their own sentence is what the I Do and We Do work on.",
      tag: "[Launch | Explicit teaching | HITS 1]",
    }));
  }

  // 6. I Do: which words need a sign
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Not every word is a sign");
    T.addCard(s, 0.5, 1.4, 9, 1.15, { variant: "tint", tone: C.PRIMARY });
    s.addText(
      [
        { text: "I like ", options: { color: C.MUTED } },
        { text: "basketball", options: { color: C.CHARCOAL, bold: true, underline: true } },
        { text: " and I play it with my ", options: { color: C.MUTED } },
        { text: "brother", options: { color: C.CHARCOAL, bold: true, underline: true } },
        { text: ".", options: { color: C.MUTED } },
      ],
      {
        x: 0.7, y: 1.4, w: 8.6, h: 1.15, fontSize: 30, fontFace: T.FONT_H,
        align: "center", valign: "middle", margin: 0,
      }
    );
    const cols = [
      ["Carries the meaning", "basketball, brother, like", C.SUCCESS],
      ["Does not get looked up", "I, and, it, with, my", C.MUTED],
    ];
    cols.forEach(([head, body, tone], i) => {
      const x = 0.5 + i * 4.6;
      T.addCard(s, x, 2.75, 4.4, 1.5, { variant: "outline", tone });
      s.addText(head, {
        x: x + 0.15, y: 2.85, w: 4.1, h: 0.45, fontSize: 17, fontFace: T.FONT_B, bold: true,
        color: tone, align: "left", valign: "middle", margin: 0,
      });
      s.addText(body, {
        x: x + 0.15, y: 3.32, w: 4.1, h: 0.8, fontSize: 19, fontFace: T.FONT_H,
        color: C.CHARCOAL, align: "left", valign: "top", margin: 0,
      });
    });
    s.addText("Underline the ones that carry the meaning. Leave the rest alone.", {
      x: 0.5, y: 4.4, w: 9, h: 0.45, fontSize: 17, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "basketball, brother, like",
      beats: [
        ["MODEL with your own sentence.", "Underline as you talk, so they see the choosing."],
        ["SAY: Every English word is not a sign.", "The small joining words do not get looked up."],
        ["SAY: One question about each word.", "Does it carry the meaning, or hold the sentence together?"],
        ["GET IT WRONG: underline every word, then stop.", "SAY: Now I have nine things to find."],
      ],
      trap: ["underlining the small words too.", "Fix: cross the joining words out together, student redoes it."],
      stretch: "find a word that carries meaning but is not a thing.",
      help: "the sentence already underlined: the task becomes finding, not choosing.",
      prep: "Students were being asked to do this silently from week one. Teach it once, properly.",
      tag: "[I Do | Explicit teaching | HITS 3]",
    }));
  }

  // 7. I Do: how to find it in Sign It!
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.PRIMARY);
    T.addBadge(s, "I Do", { color: C.PRIMARY });
    T.addCueStrip(s, ["eyesUp", "watchCopy"]);
    T.addTitle(s, "Finding a word in Sign It!");
    const steps = [
      ["1", "Go to the back", "The word list is at the back, in alphabetical order."],
      ["2", "Find your word", "Look up basketball. It gives you a page number."],
      ["3", "Page, not chapter", "Page 80 is a page. Chapter 8 is not. Check the bottom of the page."],
    ];
    const sw = (5.7 - 0.16 * 2) / 3;
    steps.forEach(([n, head, body], i) => {
      const x = 0.5 + i * (sw + 0.16);
      T.addCard(s, x, 1.4, sw, 2.6, { variant: "tint", tone: C.PRIMARY });
      s.addText(n, {
        x, y: 1.5, w: sw, h: 0.4, fontSize: 15, fontFace: T.FONT_B, bold: true,
        color: C.PRIMARY, align: "center", valign: "middle", margin: 0,
      });
      s.addText(head, {
        x: x + 0.1, y: 1.9, w: sw - 0.2, h: 0.55, fontSize: 17, fontFace: T.FONT_H, bold: true,
        color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
      });
      s.addText(body, {
        x: x + 0.12, y: 2.5, w: sw - 0.24, h: 1.35, fontSize: 14, fontFace: T.FONT_B,
        color: C.CHARCOAL, align: "center", valign: "top", margin: 0,
      });
    });
    A.addSignItCover(T, s, { x: 6.45, y: 1.4, w: 3.05, h: 2.6 });
    s.addText("Write the page number next to your underlined word.", {
      x: 0.5, y: 4.2, w: 9, h: 0.5, fontSize: 18, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "the word list at the back gives the page, not the chapter",
      beats: [
        ["MODEL with the book held up.", "Back of the book, find basketball, read the number."],
        ["SAY: That number is a page.", "A chapter number is the big one at the start."],
        "SHOW the page number printed at the bottom of a page.",
        ["GET IT WRONG: turn to chapter 8, not page 80.", "SAY: Check the bottom of the page."],
      ],
      trap: ["turning to the chapter number.", "Fix: hold up both, student finds the page."],
      stretch: "find a word that is not in the list, and say what to do next.",
      help: "a partner who already found it, and the page given.",
      prep: "Students confuse chapter and page numbers. This beat exists for that.",
      tag: "[I Do | Explicit teaching | HITS 3, 4]",
    }));
  }

  // 8. We Do: do it together
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.SUCCESS);
    T.addBadge(s, "We Do", { color: C.SUCCESS });
    T.addCueStrip(s, ["voicesOff", "partner", "whiteboards"]);
    T.addTitle(s, "One sentence, together");
    T.addCard(s, 0.5, 1.4, 9, 1.1, { variant: "tint", tone: C.SUCCESS });
    s.addText("On Saturday my sister plays netball at the beach.", {
      x: 0.7, y: 1.4, w: 8.6, h: 1.1, fontSize: 28, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addInstructionCard(s, [
      { role: "header", text: "With your partner" },
      { text: "Copy the sentence onto one board." },
      { text: "Underline the words that carry the meaning." },
      { text: "Find one of them in Sign It! and write the page number." },
    ], { x: 0.5, y: 2.65, w: 6.2, h: 1.85, strip: C.SUCCESS });
    A.addCountdown(T, s, 300, { x: 7.35, y: 2.75, w: 1.6 });
    s.addText("5 minutes", {
      x: 6.95, y: 4.42, w: 2.4, h: 0.4, fontSize: 16, fontFace: T.FONT_B, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "Saturday, sister, netball, beach",
      beats: [
        "SET UP: one board per pair, one Sign It! between two.",
        "TIME: advance to start the clock. Five minutes.",
        ["CIRCULATE: look at what they underlined before you look at the page numbers.", "The choosing is the hard part."],
        "COLLECT: boards on the desk, page numbers visible.",
      ],
      trap: ["underlining on, my and at.", "Fix: ask which of those you could draw a picture of, student redoes it."],
      stretch: "find all four words, not one.",
      help: "the sentence with the four words already underlined.",
      prep: "10 min. Same process as the I Do, with the teacher still in the room for it.",
      tag: "[We Do | Collaborative learning | HITS 5]",
    }));
  }

  // 9. Primary decision point
  {
    const s = pres.addSlide();
    T.addTopBar(s, C.ASSESS);
    T.addBadge(s, "Check it", { color: C.ASSESS });
    T.addCueStrip(s, ["whiteboards"]);
    T.addTitle(s, "Which words would you look up?");
    T.addCard(s, 0.5, 1.5, 9, 1.25, { variant: "tint", tone: C.ASSESS });
    s.addText("My brother and I watch the football on Sunday.", {
      x: 0.7, y: 1.5, w: 8.6, h: 1.25, fontSize: 30, fontFace: T.FONT_H, bold: true,
      color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });
    T.addTextOnShape(
      s,
      "Write the words you would look up. Hover it. Chin it on three.",
      { x: 1.1, y: 3.05, w: 7.8, h: 0.7, rectRadius: 0.12, fill: { color: C.ASSESS } },
      { fontSize: 19, fontFace: T.FONT_B, bold: true, color: C.WHITE,
        align: "center", valign: "middle", margin: 0 }
    );
    s.addText("Hover means the board stays flat until I say chin it, so nobody copies.", {
      x: 0.5, y: 3.95, w: 9, h: 0.5, fontSize: 15, fontFace: T.FONT_B,
      color: C.MUTED, align: "center", valign: "middle", margin: 0,
    });
    T.addFooter(s, FOOTER);
    s.addNotes(T.composeGlanceNotes({
      answer: "brother, watch, football, Sunday",
      beats: [
        ["ASK: which words in that sentence would you look up.",
          "5 sec think time. Cue: Write it. Hover it. Chin it on three. One, two, three.",
          "EXPECT: brother, watch, football, Sunday"],
        ["SCAN the room, back row first.",
          "80%+ -> cold-call one student to say why they left a word out.",
          "Less -> cross the joining words out together, then re-ask with a new sentence."],
        ["FOLLOW UP the student who answered.", "SAY: You left out on. Say how you knew."],
      ],
      trap: ["including and, the and on.", "Fix: cross the joining words out together, student redoes it."],
      prep: [
        "The decision point that decides whether they can look up signs alone from Lesson 2.",
        "Hover then chin is the school routine from prep to Year 6. Nobody shows early.",
      ],
      tag: "[CFU | Evaluating impact | HITS 7]",
    }));
  }

  // 10. You Do
  {
    const s = T.youDoSlide(
      pres,
      "Your own sentence, all the way through",
      "Go back to the sentence you wrote at the start. Underline the words that need a sign, look one up in Sign It!, and sign it to your partner.",
      [
        "Underline the words that carry the meaning.",
        "Find one in Sign It! and write the page.",
        "Sign it to your partner, book closed.",
      ],
      T.composeGlanceNotes({
        beats: [
          "SAY: Your own sentence this time. Same three steps.",
          "TIME: 10 minutes. Book closed for the signing bit.",
          ["CIRCULATE: look for books still open when they sign.", "Closing it is the point."],
          "COLLECT: partners check each other found a real page number.",
        ],
        stretch: "look up a second word and sign both in one go.",
        help: "the word chosen for them, so the task is finding and signing only.",
        prep: [
          "10 min. This is the whole routine, done once with you in the room.",
          "From Lesson 2 they do it in five minutes at the start, silently.",
        ],
        tag: "[You Do | Explicit teaching | HITS 10]",
      }),
      FOOTER,
      { where: "Voices off  |  Table partner  |  Sign It! and your board" }
    );
    T.addCueStrip(s, ["voicesOff", "partner", "whiteboards"]);
  }

  // 11. Exit ticket
  T.exitTicketSlide(
    pres,
    [
      "Rate yourself on the three I can statements: just starting, getting there, got it.",
      "Write one line: the word you looked up and its page number.",
      "Cultural question: Auslan is not English with signs instead of words. Give one thing you noticed today that shows that.",
    ],
    T.composeGlanceNotes({
      beats: [
        "COLLECT nothing. Everyone writes in the journal while you call the rotation.",
        ["CALL group 1, one at a time.", "They name two words they would look up in a sentence you give them."],
        "Then they find one in Sign It! and give you the page number.",
        ["RECORD E, C or P on the tracker,", "against the checklist row, not against the lesson."],
      ],
      prep: [
        "5 min. Evidence piece 3, first pass. Rotation group 1 is due today.",
        "The cultural question wants the joining words they left out, or the order.",
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
      reflectionPrompt: "Turn and tell your partner one word you would never bother looking up, and why.",
      scItems: [
        "I can underline the words in my sentence that need a sign.",
        "I can find one of those words in Sign It! and give its page.",
        "I can sign it to my partner without looking back at the book.",
      ],
      selfAssessment: ["Just starting", "Getting there", "Got it"],
    },
    T.composeGlanceNotes({
      beats: [
        "SAY: Turn and tell your partner one word you would never look up, and why.",
        "SAY: From next week, you look up the signs you need on your own.",
      ],
      prep: [
        "From Lesson 2 the Do Now is the Deaf Games video. This is how they find new signs.",
        "If today did not land, reteach the choosing step at the start of Lesson 2's We Do.",
      ],
      tag: "[Closing | Planning | HITS 9]",
    })
  );

  return pres;
}

const pres = build();
fs.mkdirSync(path.join("output", OUT_DIR), { recursive: true });
const file = path.join("output", OUT_DIR, "DeafSport Enrichment Session 1 Finding The Signs.pptx");
pres.writeFile({ fileName: file }).then(() => {
  console.log("PPTX written to " + file);
  report.print();
  const lines = report.rehearsalLines();
  if (lines.length) {
    console.log("\nREHEARSE FIRST");
    lines.forEach((l) => console.log("  " + l));
  }
});
