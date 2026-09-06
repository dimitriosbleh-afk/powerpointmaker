"use strict";

// Grammar 5-minute routines — Year 5/6
// Weeks 1-10, 5 sessions per week, 50 sessions total.
// Each session = 3 slides: Teach, Practise, Check.
// Short, sharp, sweet: slots into a larger teaching block.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");

// Use variant 2 to differentiate visually from the CSR deck.
const T = createTheme("literacy", "grade56", 2);
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, closingSlide,
  addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal, runSlideDiagnostics,
} = T;

const FOOTER = "Grammar | Year 5/6 | 5-minute routine";
const OUT_DIR = "output/Grammar_Year56";

fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Grammar — Year 5/6 (Weeks 1-10)";

// ────────────────────────────────────────────────────────────────────────────
//  Session divider (precedes every Teach slide — marks session boundaries)
// ────────────────────────────────────────────────────────────────────────────

function sessionDivider(weekLabel, sessionLabel, focusTitle) {
  const s = pres.addSlide();
  s.background = { color: C.BG_LIGHT };

  // Top and bottom accent bars
  s.addShape("rect", { x: 0, y: 0, w: SLIDE_W, h: 0.12, fill: { color: C.PRIMARY } });
  s.addShape("rect", { x: 0, y: SLIDE_H - 0.12, w: SLIDE_W, h: 0.12, fill: { color: C.ACCENT } });

  // Central card
  addCard(s, 0.8, 1.35, 8.4, 3.15, { strip: C.PRIMARY, fill: C.WHITE });

  // Week indicator (small, at top of card)
  s.addText(weekLabel.toUpperCase(), {
    x: 1.1, y: 1.55, w: 7.8, h: 0.32,
    fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });

  // Session heading (very large)
  s.addText(`${sessionLabel} of 5`, {
    x: 1.1, y: 1.92, w: 7.8, h: 0.85,
    fontSize: 40, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0,
  });

  // Session focus / title
  s.addText(focusTitle, {
    x: 1.1, y: 2.85, w: 7.8, h: 1.45,
    fontSize: 20, fontFace: FONT_H, color: C.SECONDARY, margin: 0, valign: "top",
    fit: "shrink",
  });

  // Footer indicator
  s.addText("5-minute routine  |  Teach -> Practise -> Check", {
    x: 0.5, y: 4.75, w: 9, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
  });

  return s;
}

// ────────────────────────────────────────────────────────────────────────────
//  Builders
// ────────────────────────────────────────────────────────────────────────────

function weekDivider(pres, weekNumber, weekTitle, focus, notes) {
  const s = pres.addSlide();
  s.background = { color: C.BG_DARK };
  s.addShape("rect", { x: 0, y: 0, w: 0.12, h: SLIDE_H, fill: { color: C.ACCENT } });

  s.addShape("roundRect", {
    x: 7.5, y: -0.6, w: 3.5, h: 3.5, rectRadius: 1.75,
    fill: { color: C.DECOR_1, transparency: 75 },
  });
  s.addShape("roundRect", {
    x: 8.0, y: 3.5, w: 2.5, h: 2.5, rectRadius: 1.25,
    fill: { color: C.DECOR_2, transparency: 80 },
  });

  s.addText(`Week ${weekNumber}`, {
    x: 0.7, y: 0.9, w: 8, h: 0.5,
    fontSize: 22, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  s.addText(weekTitle, {
    x: 0.7, y: 1.45, w: 8.5, h: 1.2,
    fontSize: 34, fontFace: FONT_H, color: C.WHITE, bold: true, margin: 0,
  });
  s.addText("Focus:", {
    x: 0.7, y: 2.95, w: 1.2, h: 0.35,
    fontSize: 12, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  s.addText(focus, {
    x: 0.7, y: 3.35, w: 8.5, h: 1.2,
    fontSize: 16, fontFace: FONT_B, color: C.WHITE, margin: 0,
  });
  s.addText("5-minute routine  |  5 sessions this week", {
    x: 0.7, y: 4.75, w: 8, h: 0.3,
    fontSize: 12, fontFace: FONT_B, color: C.ACCENT, margin: 0,
  });

  if (notes) s.addNotes(notes);
  return s;
}

/**
 * Teach slide: big rule card + examples.
 * Renders a session divider slide first to mark the session boundary.
 * examples: array of { text, note } — the sentence/example + a short label.
 */
function teachSlide(weekLabel, sessionLabel, title, rule, examples, notes) {
  sessionDivider(weekLabel, sessionLabel, title);

  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  addBadge(s, "Teach", { color: C.PRIMARY, w: 1.25 });
  addBadge(s, weekLabel, { color: C.SECONDARY, w: 2.0, x: 1.85 });
  addBadge(s, sessionLabel, { color: C.ACCENT, w: 1.5, x: 3.95 });
  addTitle(s, title);

  // Rule card
  const ruleH = 1.25;
  addCard(s, 0.5, CONTENT_TOP, 9, ruleH, { strip: C.PRIMARY, fill: C.BG_CARD });
  s.addText("The rule", {
    x: 0.75, y: CONTENT_TOP + 0.10, w: 3, h: 0.28,
    fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  s.addText(rule, {
    x: 0.75, y: CONTENT_TOP + 0.42, w: 8.5, h: ruleH - 0.50,
    fontSize: 15, fontFace: FONT_H, color: C.CHARCOAL, margin: 0, valign: "top",
  });

  // Examples stacked rows (each example full-width for sentences)
  const exY = CONTENT_TOP + ruleH + 0.15;
  const exH = SAFE_BOTTOM - exY;
  const n = Math.min(examples.length, 3);
  const rowH = (exH - 0.05 - (n - 1) * 0.08) / n;

  s.addText("Examples", {
    x: 0.5, y: exY, w: 3, h: 0.28,
    fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });

  const startY = exY + 0.32;
  for (let i = 0; i < n; i++) {
    const ex = examples[i];
    const y = startY + i * (rowH + 0.06);
    const h = rowH - 0.05;
    addCard(s, 0.5, y, 9, h, {
      strip: i === 0 ? C.SECONDARY : (i === 1 ? C.ACCENT : C.ALERT),
      fill: C.WHITE,
    });
    // Note label
    if (ex.note) {
      s.addText(ex.note, {
        x: 0.75, y: y + 0.06, w: 8.3, h: 0.26,
        fontSize: 10.5, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
      });
    }
    s.addText(ex.text, {
      x: 0.75, y: y + (ex.note ? 0.30 : 0.10), w: 8.3, h: h - (ex.note ? 0.36 : 0.16),
      fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, margin: 0, valign: "middle",
      fit: "shrink",
    });
  }

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

/**
 * Practise slide: prompt (left) + sentences to work on (right).
 */
function practiseSlide(weekLabel, sessionLabel, title, prompt, sentences, notes) {
  const s = pres.addSlide();
  addTopBar(s, C.SECONDARY);
  addBadge(s, "Practise", { color: C.SECONDARY, w: 1.4 });
  addBadge(s, weekLabel, { color: C.PRIMARY, w: 2.0, x: 2.0 });
  addBadge(s, sessionLabel, { color: C.ACCENT, w: 1.5, x: 4.1 });
  addTitle(s, title);

  const cardH = SAFE_BOTTOM - CONTENT_TOP;

  // Left prompt
  addCard(s, 0.5, CONTENT_TOP, 4.1, cardH, { strip: C.SECONDARY, fill: C.WHITE });
  s.addText("Your task", {
    x: 0.7, y: CONTENT_TOP + 0.14, w: 3.7, h: 0.32,
    fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });
  s.addText(prompt, {
    x: 0.7, y: CONTENT_TOP + 0.52, w: 3.7, h: cardH - 0.70,
    fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
    paraSpaceAfter: 4,
  });

  // Right sentences
  addCard(s, 4.8, CONTENT_TOP, 4.7, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("Sentences", {
    x: 5.0, y: CONTENT_TOP + 0.14, w: 3, h: 0.32,
    fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });

  const n = sentences.length;
  const listY = CONTENT_TOP + 0.52;
  const listH = cardH - 0.70;
  const itemH = listH / n;

  sentences.forEach((sent, i) => {
    const y = listY + i * itemH;
    s.addText(`${i + 1}.  ${sent}`, {
      x: 5.0, y, w: 4.4, h: itemH - 0.04,
      fontSize: n >= 5 ? 11.5 : n >= 4 ? 12.5 : 13.5,
      fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
    });
  });

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

/**
 * Check slide with reveal.
 */
function checkSlide(weekLabel, sessionLabel, title, question, answer, notes, notesReveal) {
  return withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "Check", { color: C.ALERT, w: 1.25 });
      addBadge(s, weekLabel, { color: C.PRIMARY, w: 2.0, x: 1.85 });
      addBadge(s, sessionLabel, { color: C.ACCENT, w: 1.5, x: 3.95 });
      addTitle(s, title, { color: C.ALERT });

      addCard(s, 0.5, CONTENT_TOP, 9, 2.0, { strip: C.ALERT, fill: C.WHITE });
      s.addText("Quick check", {
        x: 0.75, y: CONTENT_TOP + 0.12, w: 3, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
      });
      s.addText(question, {
        x: 0.75, y: CONTENT_TOP + 0.44, w: 8.5, h: 1.5,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(notes || "");
      return s;
    },
    (s) => {
      const ansY = CONTENT_TOP + 2.15;
      const ansH = 1.0;
      s.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: ansH, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });
      s.addText("Answer", {
        x: 0.75, y: ansY + 0.08, w: 3, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true, margin: 0,
      });
      s.addText(answer, {
        x: 0.75, y: ansY + 0.34, w: 8.5, h: ansH - 0.42,
        fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true,
        valign: "middle", margin: 0, fit: "shrink",
      });
      if (notesReveal) s.addNotes(notesReveal);
    }
  );
}

// ────────────────────────────────────────────────────────────────────────────
//  Title slide
// ────────────────────────────────────────────────────────────────────────────

titleSlide(
  pres,
  "Grammar Routines",
  "5-minute explicit teaching sessions",
  "Year 5/6  |  Weeks 1-10  |  50 sessions",
  `SAY:
- This deck is a 5-minute grammar routine for daily use.
- Each session has three slides: Teach, Practise, then a Check.

DO:
- Navigate to the session you need.
- OG Books ready before you start.

TEACHER NOTES:
Short, sharp, sweet routines. Do not stretch beyond five minutes.

[Grammar | 5-minute routine]`
);

// ===========================================================================
//  WEEK 1 — IDENTIFY RELATIVE CLAUSES
// ===========================================================================

weekDivider(
  pres, 1, "Identify Relative Clauses",
  "Relative clauses add information about a noun. This week we spot them in sentences.",
  `[Week 1 | Identify relative clauses]`
);

// W1 S1: What is a clause
teachSlide(
  "Week 1", "Session 1",
  "What is a clause? A group of words with a subject and a verb",
  "A clause has a subject (who or what) and a verb (the action). Every sentence needs at least one clause.",
  [
    { note: "Subject + verb", text: "The dog barked." },
    { note: "Subject + verb",  text: "My sister is laughing." },
    { note: "Subject + verb",  text: "We walked to school." },
  ],
  `SAY:
- A clause needs two things: a subject and a verb.
- 'The dog barked' -- subject is 'the dog', verb is 'barked'. That's a clause.

DO:
- Underline each subject; circle each verb as you read.

TEACHER NOTES:
This session sets up the rest of the week. A relative clause is a clause, so students must recognise clauses first.

[Grammar | W1 S1 | Teach]`
);

practiseSlide(
  "Week 1", "Session 1",
  "Practise: find the subject and verb",
  "For each sentence:\n1. Underline the subject.\n2. Circle the verb.\n\nCheck: does it have both?",
  [
    "The cat slept.",
    "My friend plays soccer.",
    "We laughed loudly.",
    "The rain fell hard.",
  ],
  `SAY:
- Underline the subject, circle the verb.

DO:
- 90 seconds.
- Pair check.

[Grammar | W1 S1 | Practise]`
);

checkSlide(
  "Week 1", "Session 1",
  "Check: which sentence is a complete clause?",
  "Write A or B:\nA. running through the park\nB. The children ran through the park.",
  "B -- 'The children ran' has a subject and a verb",
  `[Grammar | W1 S1 | Check]`,
  `SAY:
- B. 'Running through the park' has no subject, so it's a phrase, not a clause.`
);

// W1 S2: What is a relative clause
teachSlide(
  "Week 1", "Session 2",
  "A relative clause adds information about a noun",
  "A relative clause is a group of words that tells us more about a noun in the sentence. It usually starts with who, which, that, whose, or whom.",
  [
    { note: "Relative clause in bold", text: "The girl who won the race is my cousin." },
    { note: "Relative clause in bold", text: "The book that I read was great." },
    { note: "Relative clause in bold", text: "The dog, which had muddy paws, ran inside." },
  ],
  `SAY:
- A relative clause adds info about a noun.
- Look at 'The girl who won the race is my cousin.' -- 'who won the race' tells us more about 'the girl'.

DO:
- Read each sentence twice: once with, once without the relative clause.
- Show how the meaning changes.

TEACHER NOTES:
Focus only on spotting. We add them next week.

[Grammar | W1 S2 | Teach]`
);

practiseSlide(
  "Week 1", "Session 2",
  "Practise: find the relative clause",
  "Underline the relative clause in each sentence.\n\nHint: it usually starts with who, which, or that.",
  [
    "The boy who lives next door is funny.",
    "The pie that she baked tasted great.",
    "My teacher, who is very kind, helped me.",
    "The ball which bounced high broke the lamp.",
  ],
  `SAY:
- Hunt for who, which, or that. Then read what follows.

DO:
- 2 minutes.

WATCH FOR:
- Students underlining the whole sentence -- redirect to the clause only.

[Grammar | W1 S2 | Practise]`
);

checkSlide(
  "Week 1", "Session 2",
  "Check: find the relative clause.",
  "In this sentence, which part is the relative clause?\n\nThe car that was stolen was red.",
  "that was stolen",
  `[Grammar | W1 S2 | Check]`,
  `SAY:
- 'That was stolen' tells us which car. That's the relative clause.`
);

// W1 S3: Relative pronouns
teachSlide(
  "Week 1", "Session 3",
  "Relative pronouns: who, which, that, whose",
  "Relative clauses start with a special word called a relative pronoun: who (for people), which (for things/animals), that (people or things), whose (for ownership).",
  [
    { note: "who = people",  text: "The doctor who helped me was kind." },
    { note: "which = things", text: "The car which broke down is old." },
    { note: "whose = ownership", text: "The student whose bag is missing spoke up." },
  ],
  `SAY:
- Different relative pronouns for different jobs.
- Who for people. Which for things or animals. Whose for ownership.

DO:
- Highlight the relative pronoun in each example.

[Grammar | W1 S3 | Teach]`
);

practiseSlide(
  "Week 1", "Session 3",
  "Practise: which pronoun fits?",
  "Fill in the blank with who, which, or whose.",
  [
    "The boy ___ is running is my brother.",
    "The apple ___ fell is bruised.",
    "The girl ___ jacket is red is here.",
    "The book ___ is on the shelf is new.",
  ],
  `SAY:
- Think about what follows: a person, a thing, or ownership?

[Grammar | W1 S3 | Practise]`
);

checkSlide(
  "Week 1", "Session 3",
  "Check: pick the right pronoun.",
  "Fill the blank:\n\nThe teacher ___ marked our tests was happy.",
  "who",
  `[Grammar | W1 S3 | Check]`,
  `SAY:
- Teacher is a person, so we use 'who'.`
);

// W1 S4: Spotting relative clauses
teachSlide(
  "Week 1", "Session 4",
  "Spot the relative clause: two steps",
  "Step 1: Find the relative pronoun (who, which, that, whose). Step 2: Read from there until the clause ends (usually at a comma or before the main verb returns).",
  [
    { note: "Clause: who has red hair", text: "The artist who has red hair won." },
    { note: "Clause: which was quick", text: "The reply, which was quick, helped." },
    { note: "Clause: that we finished", text: "The puzzle that we finished was hard." },
  ],
  `SAY:
- Find the pronoun. Then read what follows until the idea about the noun ends.

[Grammar | W1 S4 | Teach]`
);

practiseSlide(
  "Week 1", "Session 4",
  "Practise: underline the full relative clause",
  "For each sentence:\n1. Find the relative pronoun.\n2. Underline from the pronoun to the end of the clause.",
  [
    "The man who delivers the mail is friendly.",
    "The museum, which opened last year, is busy.",
    "I found the pen that I lost.",
    "The dog whose tail is wagging is mine.",
  ],
  `SAY:
- Pronoun first, then the rest of the clause.

[Grammar | W1 S4 | Practise]`
);

checkSlide(
  "Week 1", "Session 4",
  "Check: underline the relative clause.",
  "The shop that sells comics closed today.",
  "that sells comics",
  `[Grammar | W1 S4 | Check]`,
  `SAY:
- 'That sells comics' tells us which shop. That's the relative clause.`
);

// W1 S5: Mixed identification
teachSlide(
  "Week 1", "Session 5",
  "Relative clauses review: spot and explain",
  "This session mixes everything from this week. Find the relative clause, name the pronoun, and say which noun it describes.",
  [
    { note: "Relative clause describes 'cake'", text: "The cake that mum baked was gone." },
    { note: "Relative clause describes 'kids'", text: "The kids who were running stopped." },
    { note: "Relative clause describes 'hat'", text: "The hat, which was red, blew away." },
  ],
  `SAY:
- Full check: find the clause, name the pronoun, name the noun it describes.

[Grammar | W1 S5 | Teach]`
);

practiseSlide(
  "Week 1", "Session 5",
  "Practise: find clause, pronoun, noun",
  "For each sentence:\n- Underline the relative clause.\n- Circle the relative pronoun.\n- Box the noun it describes.",
  [
    "The girl who was dancing smiled.",
    "The sandwich that I made is gone.",
    "The tree, which is very old, fell.",
    "The player whose shirt is blue scored.",
  ],
  `[Grammar | W1 S5 | Practise]`
);

checkSlide(
  "Week 1", "Session 5",
  "Check: identify the clause, pronoun, and noun.",
  "The chef who cooked dinner smiled.\n\nWhat is the relative clause, the pronoun, and the noun it describes?",
  "Clause: who cooked dinner  |  Pronoun: who  |  Noun: chef",
  `[Grammar | W1 S5 | Check]`,
  `SAY:
- 'Who cooked dinner' describes 'the chef', starting with the pronoun 'who'.`
);

// ===========================================================================
//  WEEK 2 — ADD RELATIVE CLAUSES
// ===========================================================================

weekDivider(
  pres, 2, "Add Relative Clauses",
  "Now we build relative clauses into our own sentences to give the reader more detail.",
  `[Week 2 | Add relative clauses]`
);

// W2 S1: who
teachSlide(
  "Week 2", "Session 1",
  "Add a 'who' clause to describe a person",
  "Pick a person in your sentence. Add 'who' plus extra information about them straight after that person.",
  [
    { note: "Base:", text: "The man was smiling." },
    { note: "Added 'who' clause:", text: "The man who sold us the fruit was smiling." },
    { note: "Pattern:", text: "noun + who + info + rest of sentence" },
  ],
  `SAY:
- Find the person. Add 'who' plus more info straight after them.
- 'The man was smiling' becomes 'The man who sold us the fruit was smiling'.

DO:
- Point at where the 'who' clause slots in.

[Grammar | W2 S1 | Teach]`
);

practiseSlide(
  "Week 2", "Session 1",
  "Practise: add a 'who' clause",
  "For each short sentence, add a 'who' clause about the person.",
  [
    "The teacher smiled.",
    "My aunt came to visit.",
    "The driver waved.",
    "The singer bowed.",
  ],
  `SAY:
- Slide the clause in straight after the person.

[Grammar | W2 S1 | Practise]`
);

checkSlide(
  "Week 2", "Session 1",
  "Check: add a 'who' clause.",
  "Add a 'who' clause:\n\nThe artist painted a tree.",
  "The artist who lives next door painted a tree.  (or similar)",
  `[Grammar | W2 S1 | Check]`,
  `SAY:
- Any sensible 'who' clause works as long as it tells us more about 'the artist'.`
);

// W2 S2: which
teachSlide(
  "Week 2", "Session 2",
  "Add a 'which' clause to describe a thing or animal",
  "Pick a thing or animal in your sentence. Add a pair of commas and a 'which' clause after it.",
  [
    { note: "Base:", text: "The cake was delicious." },
    { note: "Added 'which' clause:", text: "The cake, which mum baked, was delicious." },
    { note: "Pattern:", text: "noun + , which + info + , rest" },
  ],
  `SAY:
- Things and animals get 'which'.
- Often we use commas around a 'which' clause because it's extra info.

DO:
- Show the commas clearly in the example.

[Grammar | W2 S2 | Teach]`
);

practiseSlide(
  "Week 2", "Session 2",
  "Practise: add a 'which' clause",
  "For each sentence, add a 'which' clause with commas.",
  [
    "The book was amazing.",
    "The bike was new.",
    "The cat was purring.",
    "The tree swayed.",
  ],
  `SAY:
- Remember the commas.

[Grammar | W2 S2 | Practise]`
);

checkSlide(
  "Week 2", "Session 2",
  "Check: add a 'which' clause with commas.",
  "Add a 'which' clause:\n\nThe puzzle was hard.",
  "The puzzle, which had 500 pieces, was hard.  (or similar, with commas)",
  `[Grammar | W2 S2 | Check]`,
  `SAY:
- Commas before 'which' and after the clause. That pattern matters here.`
);

// W2 S3: that
teachSlide(
  "Week 2", "Session 3",
  "Add a 'that' clause (no commas)",
  "'That' clauses give essential information. Without them, the sentence loses meaning. No commas.",
  [
    { note: "Base:", text: "The shoes were dirty." },
    { note: "Added 'that' clause:", text: "The shoes that I wore yesterday were dirty." },
    { note: "Pattern:", text: "noun + that + essential info (no commas)" },
  ],
  `SAY:
- 'That' clauses are essential. They tell us exactly which noun we mean.
- No commas with 'that'.

DO:
- Contrast with yesterday's 'which' clause using commas.

[Grammar | W2 S3 | Teach]`
);

practiseSlide(
  "Week 2", "Session 3",
  "Practise: add a 'that' clause",
  "For each sentence, add a 'that' clause. No commas.",
  [
    "The book was lost.",
    "The meal was hot.",
    "The sound was loud.",
    "The puzzle was tricky.",
  ],
  `[Grammar | W2 S3 | Practise]`
);

checkSlide(
  "Week 2", "Session 3",
  "Check: add a 'that' clause.",
  "The movie was scary.",
  "The movie that we watched last night was scary.  (or similar, no commas)",
  `[Grammar | W2 S3 | Check]`,
  `SAY:
- 'That' with no commas. The clause narrows down exactly which movie.`
);

// W2 S4: Placement
teachSlide(
  "Week 2", "Session 4",
  "Where does the relative clause go?",
  "Place the relative clause straight after the noun it describes. If you put it anywhere else, the sentence may not make sense.",
  [
    { note: "Wrong placement:", text: "The girl was happy who won the race." },
    { note: "Right placement:", text: "The girl who won the race was happy." },
    { note: "Pattern reminder:", text: "noun + relative clause + rest of sentence" },
  ],
  `SAY:
- The clause has to hug its noun.

DO:
- Show the drift with a silly misplaced example.

[Grammar | W2 S4 | Teach]`
);

practiseSlide(
  "Week 2", "Session 4",
  "Practise: fix the sentence",
  "Rewrite each sentence so the relative clause sits next to the noun it describes.",
  [
    "The cake tasted great that mum made.",
    "The man waved who was tall.",
    "The book was lost which I loved.",
    "The dog barked whose tail wagged.",
  ],
  `[Grammar | W2 S4 | Practise]`
);

checkSlide(
  "Week 2", "Session 4",
  "Check: rewrite it.",
  "Fix this:\n\nThe student got a prize who studied hard.",
  "The student who studied hard got a prize.",
  `[Grammar | W2 S4 | Check]`,
  `SAY:
- Slide the 'who' clause back beside 'student', where it belongs.`
);

// W2 S5: Mixed
teachSlide(
  "Week 2", "Session 5",
  "Relative clause review: pick the right pronoun and place",
  "Choose the right pronoun (who/which/that), place the clause next to the noun, and check your commas.",
  [
    { note: "People -> who", text: "The teacher who helped me was kind." },
    { note: "Things + extra info -> , which ,", text: "The cake, which was homemade, was gone." },
    { note: "Things + essential info -> that", text: "The song that we played was loud." },
  ],
  `SAY:
- Everything from this week lives here.

[Grammar | W2 S5 | Teach]`
);

practiseSlide(
  "Week 2", "Session 5",
  "Practise: build the sentence",
  "For each base, add a relative clause using the right pronoun.",
  [
    "The baker was smiling.",
    "The hat was blue.",
    "The show was funny.",
    "The artist painted.",
  ],
  `[Grammar | W2 S5 | Practise]`
);

checkSlide(
  "Week 2", "Session 5",
  "Check: add a clause.",
  "Add a relative clause:\n\nThe team won.",
  "The team that practised every day won.  (or: The team, which I support, won.)",
  `[Grammar | W2 S5 | Check]`,
  `SAY:
- Team is a thing/group. 'That' or ', which ,' both work depending on essential vs extra info.`
);

// ===========================================================================
//  WEEK 3 — PHRASES vs CLAUSES
// ===========================================================================

weekDivider(
  pres, 3, "Phrases vs Clauses",
  "A phrase is a group of words. A clause has a subject AND a verb.",
  `[Week 3 | Phrases vs clauses]`
);

// W3 S1: Phrase
teachSlide(
  "Week 3", "Session 1",
  "What is a phrase? A word group with no subject-verb pair",
  "A phrase is a group of words that work together but do NOT have both a subject and a verb.",
  [
    { note: "No subject-verb", text: "under the table" },
    { note: "No subject-verb", text: "a bright yellow umbrella" },
    { note: "No subject-verb", text: "after the rain" },
  ],
  `SAY:
- A phrase is a word group. It can't stand alone as a sentence because it's missing a subject-verb pair.

DO:
- Point at each phrase. Ask: where is the subject? where is the verb? Both missing.

[Grammar | W3 S1 | Teach]`
);

practiseSlide(
  "Week 3", "Session 1",
  "Practise: are these phrases?",
  "For each group of words, decide: is this a phrase?\n\nCheck: is there a subject AND a verb? If no, it's a phrase.",
  [
    "on the roof",
    "a tall tree",
    "after lunch",
    "in the bright sunshine",
  ],
  `[Grammar | W3 S1 | Practise]`
);

checkSlide(
  "Week 3", "Session 1",
  "Check: phrase or not?",
  "Is this a phrase?\n\nbeside the red fence",
  "YES -- it's a phrase. No subject-verb pair.",
  `[Grammar | W3 S1 | Check]`,
  `SAY:
- No subject, no verb. It's a phrase.`
);

// W3 S2: Clause
teachSlide(
  "Week 3", "Session 2",
  "What is a clause? A word group with a subject AND a verb",
  "A clause has both a subject (who or what) and a verb (the action). If it expresses a full thought, it can stand alone as a sentence.",
  [
    { note: "Subject + verb", text: "the dog barked" },
    { note: "Subject + verb", text: "we walked home" },
    { note: "Subject + verb", text: "she laughed loudly" },
  ],
  `SAY:
- Two things in a clause: a subject and a verb.

DO:
- Underline the subject; circle the verb in each.

[Grammar | W3 S2 | Teach]`
);

practiseSlide(
  "Week 3", "Session 2",
  "Practise: clause or phrase?",
  "For each group:\n- If there is a subject AND verb -> clause.\n- If not -> phrase.",
  [
    "the cat purred",
    "under the tree",
    "we ran home",
    "after school",
  ],
  `[Grammar | W3 S2 | Practise]`
);

checkSlide(
  "Week 3", "Session 2",
  "Check: clause or phrase?",
  "the children laughed",
  "Clause -- subject 'the children' + verb 'laughed'.",
  `[Grammar | W3 S2 | Check]`,
  `SAY:
- Both parts are there. It's a clause.`
);

// W3 S3: Compare
teachSlide(
  "Week 3", "Session 3",
  "Compare: phrase vs clause",
  "Same topic, different structure. The phrase has no subject-verb pair. The clause has both.",
  [
    { note: "Phrase", text: "under the bright stars" },
    { note: "Clause",  text: "the stars shone brightly" },
    { note: "Key test", text: "Ask: who or what is doing something?" },
  ],
  `SAY:
- Same idea, two forms. The test is: is someone doing something?

[Grammar | W3 S3 | Teach]`
);

practiseSlide(
  "Week 3", "Session 3",
  "Practise: sort phrases and clauses",
  "Sort each into two columns: Phrase or Clause.\nApply the subject-verb test.",
  [
    "she drew a picture",
    "a pile of books",
    "he smiled",
    "with a big smile",
  ],
  `[Grammar | W3 S3 | Practise]`
);

checkSlide(
  "Week 3", "Session 3",
  "Check: phrase or clause?",
  "before the final bell",
  "Phrase -- no subject-verb pair.",
  `[Grammar | W3 S3 | Check]`,
  `SAY:
- No one is doing anything. Phrase.`
);

// W3 S4: Sort
teachSlide(
  "Week 3", "Session 4",
  "Quick sorting strategy",
  "Two-question test: 1) Who or what is the subject? 2) What are they doing? Both answers -> clause. Missing one -> phrase.",
  [
    { note: "Q1 subject? Q2 verb? Both yes -> clause", text: "the phone rang" },
    { note: "Q1 subject? Q2 verb? No subject -> phrase", text: "ringing loudly" },
    { note: "Q1 subject? Q2 verb? No verb -> phrase", text: "the loud phone" },
  ],
  `SAY:
- Ask both questions. Both yes? Clause. Anything missing? Phrase.

[Grammar | W3 S4 | Teach]`
);

practiseSlide(
  "Week 3", "Session 4",
  "Practise: sort using the two-question test",
  "Label each as P (phrase) or C (clause) and explain why.",
  [
    "the bird sang",
    "singing in the tree",
    "the small bird",
    "it flew away",
  ],
  `[Grammar | W3 S4 | Practise]`
);

checkSlide(
  "Week 3", "Session 4",
  "Check: sort and explain.",
  "Label P or C:\n1. running through the park\n2. the children ran",
  "1. P (no subject)   2. C (subject + verb)",
  `[Grammar | W3 S4 | Check]`,
  `SAY:
- 1: no subject, so phrase. 2: has subject and verb, so clause.`
);

// W3 S5: Mixed
teachSlide(
  "Week 3", "Session 5",
  "Phrases vs clauses review",
  "Phrases fill out a sentence. Clauses carry the main idea. Every complete sentence has at least one clause.",
  [
    { note: "Phrase + clause =", text: "Under the bright stars, we sang." },
    { note: "Clause: 'we sang'", text: "Phrase: 'under the bright stars'" },
    { note: "Full sentence", text: "We sang." },
  ],
  `SAY:
- Phrases and clauses live in the same sentence all the time.

[Grammar | W3 S5 | Teach]`
);

practiseSlide(
  "Week 3", "Session 5",
  "Practise: spot both in one sentence",
  "In each sentence:\n- Underline the clause (subject + verb).\n- Circle a phrase (no subject-verb).",
  [
    "After lunch, we played outside.",
    "In the morning, the birds sang.",
    "Under the table, the cat hid.",
    "Before the test, she revised hard.",
  ],
  `[Grammar | W3 S5 | Practise]`
);

checkSlide(
  "Week 3", "Session 5",
  "Check: find the clause and a phrase.",
  "During the storm, the roof leaked.",
  "Clause: 'the roof leaked'   Phrase: 'During the storm'",
  `[Grammar | W3 S5 | Check]`,
  `SAY:
- 'The roof leaked' has subject + verb. 'During the storm' is a phrase.`
);

// ===========================================================================
//  WEEK 4 — SIMPLE WITH ADVERBIAL PHRASE vs COMPLEX
// ===========================================================================

weekDivider(
  pres, 4, "Simple with Adverbial Phrase vs Complex",
  "This week: recognise that a simple sentence can still have phrases, but a complex sentence adds another clause.",
  `[Week 4 | Simple vs complex]`
);

// W4 S1: Simple sentence
teachSlide(
  "Week 4", "Session 1",
  "What is a simple sentence? One clause",
  "A simple sentence has exactly one clause: one subject-verb group. It can still have phrases added to it.",
  [
    { note: "One clause", text: "The dog barked." },
    { note: "Still one clause (phrase added)", text: "In the morning, the dog barked." },
    { note: "Still one clause (phrase added)", text: "The dog barked loudly." },
  ],
  `SAY:
- A simple sentence is one clause. Phrases can join, but they don't add another clause.

DO:
- Show that adding 'in the morning' or 'loudly' doesn't create a new subject-verb pair.

[Grammar | W4 S1 | Teach]`
);

practiseSlide(
  "Week 4", "Session 1",
  "Practise: count the clauses",
  "For each sentence, count the clauses. If only one, it's a simple sentence.",
  [
    "The baby cried.",
    "In the kitchen, the baby cried.",
    "She laughed loudly.",
    "After lunch, she laughed.",
  ],
  `[Grammar | W4 S1 | Practise]`
);

checkSlide(
  "Week 4", "Session 1",
  "Check: is this a simple sentence?",
  "In the garden, the flowers bloomed.",
  "YES -- one clause ('the flowers bloomed') + phrase",
  `[Grammar | W4 S1 | Check]`,
  `SAY:
- One subject-verb group. Phrase added. Still a simple sentence.`
);

// W4 S2: Adverbial phrase
teachSlide(
  "Week 4", "Session 2",
  "An adverbial phrase tells when, where, or how",
  "An adverbial phrase gives extra information about the verb. It has no subject-verb pair.",
  [
    { note: "When", text: "After the bell, we lined up." },
    { note: "Where", text: "In the hallway, we waited." },
    { note: "How", text: "With a big smile, she waved." },
  ],
  `SAY:
- Adverbial phrases answer when, where, or how.
- They don't have their own subject-verb, so they don't make a new clause.

DO:
- Ask students to identify which question each phrase answers.

[Grammar | W4 S2 | Teach]`
);

practiseSlide(
  "Week 4", "Session 2",
  "Practise: spot the adverbial phrase",
  "Underline the adverbial phrase and write W (when), Wh (where), or H (how).",
  [
    "After school, we played.",
    "In the park, the kids ran.",
    "With excitement, she opened it.",
    "Before dinner, we tidied up.",
  ],
  `[Grammar | W4 S2 | Practise]`
);

checkSlide(
  "Week 4", "Session 2",
  "Check: find the adverbial phrase.",
  "During the storm, the lights flickered.",
  "Adverbial phrase: 'During the storm' -- answers WHEN",
  `[Grammar | W4 S2 | Check]`,
  `SAY:
- It tells us when the lights flickered. No subject-verb in the phrase.`
);

// W4 S3: Complex sentence
teachSlide(
  "Week 4", "Session 3",
  "What is a complex sentence? Two clauses, one main + one subordinate",
  "A complex sentence has TWO clauses. One main clause (makes sense on its own) plus one subordinate clause (usually starts with because, when, if, although, since...).",
  [
    { note: "Main + subordinate", text: "We stayed inside because it was raining." },
    { note: "Main + subordinate", text: "When the bell rang, we lined up." },
    { note: "Main + subordinate", text: "I was happy although I was tired." },
  ],
  `SAY:
- Two clauses. One main, one subordinate.
- Subordinate starters: because, when, if, although, since, while, after, before.

[Grammar | W4 S3 | Teach]`
);

practiseSlide(
  "Week 4", "Session 3",
  "Practise: find the two clauses",
  "Underline the main clause. Put brackets around the subordinate clause.",
  [
    "We left when the bell rang.",
    "I smiled because she waved.",
    "After the game ended, we packed up.",
    "She ran fast although she was tired.",
  ],
  `[Grammar | W4 S3 | Practise]`
);

checkSlide(
  "Week 4", "Session 3",
  "Check: find the two clauses.",
  "I finished my work before the bell rang.",
  "Main: 'I finished my work'   Subordinate: 'before the bell rang'",
  `[Grammar | W4 S3 | Check]`,
  `SAY:
- Each clause has its own subject-verb. 'Before' signals the subordinate clause.`
);

// W4 S4: Tell them apart
teachSlide(
  "Week 4", "Session 4",
  "Simple with adverbial phrase vs complex: the key test",
  "Count the subject-verb pairs. One = simple sentence (with possibly a phrase). Two = complex sentence.",
  [
    { note: "Simple (1 S-V)", text: "After lunch, we played." },
    { note: "Complex (2 S-V)", text: "After lunch finished, we played." },
    { note: "The test", text: "Does 'after lunch' have its own verb? If yes -> clause. If no -> phrase." },
  ],
  `SAY:
- Phrase has no verb. Clause has a verb.
- 'After lunch' (phrase) vs 'after lunch finished' (clause).

DO:
- Point to each subject and verb.

[Grammar | W4 S4 | Teach]`
);

practiseSlide(
  "Week 4", "Session 4",
  "Practise: simple or complex?",
  "Count the subject-verb pairs. Label S (simple) or C (complex).",
  [
    "In the morning, I walked to school.",
    "When the bell rang, I walked to school.",
    "She smiled with excitement.",
    "She smiled because she was happy.",
  ],
  `[Grammar | W4 S4 | Practise]`
);

checkSlide(
  "Week 4", "Session 4",
  "Check: simple or complex?",
  "After the game ended, we cheered.",
  "Complex -- two subject-verb pairs",
  `[Grammar | W4 S4 | Check]`,
  `SAY:
- 'The game ended' and 'we cheered' are two clauses. Complex.`
);

// W4 S5: Mixed
teachSlide(
  "Week 4", "Session 5",
  "Review: simple with phrase vs complex",
  "Steps: 1) Count subject-verb pairs. 2) One pair with a phrase = simple. 3) Two pairs = complex.",
  [
    { note: "Simple", text: "Before the show, we bought tickets." },
    { note: "Complex", text: "Before the show started, we bought tickets." },
    { note: "Simple", text: "With a loud cheer, the crowd stood." },
  ],
  `SAY:
- Count subject-verb pairs. That picks the sentence type.

[Grammar | W4 S5 | Teach]`
);

practiseSlide(
  "Week 4", "Session 5",
  "Practise: mixed",
  "Label each sentence S (simple) or C (complex). Briefly explain.",
  [
    "In the yard, the kids played.",
    "When I arrived, the kids played.",
    "After the storm ended, we walked.",
    "After the storm, we walked.",
  ],
  `[Grammar | W4 S5 | Practise]`
);

checkSlide(
  "Week 4", "Session 5",
  "Check: label and explain.",
  "Label:\n1. Although she was tired, she kept running.\n2. With tired legs, she kept running.",
  "1. C (two clauses)   2. S (one clause + phrase)",
  `[Grammar | W4 S5 | Check]`,
  `SAY:
- 1: 'Although she was tired' has its own subject-verb -> complex. 2: 'With tired legs' is just a phrase -> simple.`
);

// ===========================================================================
//  WEEK 5 — EXPAND SIMPLE TO COMPLEX
// ===========================================================================

weekDivider(
  pres, 5, "Expand Simple to Complex",
  "Take a simple sentence and add a subordinate clause to make it complex. Use because, when, if, and although.",
  `[Week 5 | Expand to complex]`
);

// W5 S1: because
teachSlide(
  "Week 5", "Session 1",
  "Expand with 'because' - add a reason",
  "Take a simple sentence. Add a 'because' clause at the end to give a reason.",
  [
    { note: "Simple:", text: "I wore a coat." },
    { note: "Complex:", text: "I wore a coat because it was cold." },
    { note: "Pattern:", text: "main clause + because + reason" },
  ],
  `SAY:
- 'Because' answers 'why?' -- it's a reason.

DO:
- Ask 'why' out loud before adding the clause.

[Grammar | W5 S1 | Teach]`
);

practiseSlide(
  "Week 5", "Session 1",
  "Practise: add a 'because' clause",
  "For each simple sentence, add a 'because' clause to give a reason.",
  [
    "I was late.",
    "We ate quickly.",
    "She laughed.",
    "The dog barked.",
  ],
  `[Grammar | W5 S1 | Practise]`
);

checkSlide(
  "Week 5", "Session 1",
  "Check: expand with 'because'.",
  "I shut the door.",
  "I shut the door because it was noisy.  (or similar)",
  `[Grammar | W5 S1 | Check]`,
  `SAY:
- Add a reason after 'because'.`
);

// W5 S2: when/while
teachSlide(
  "Week 5", "Session 2",
  "Expand with 'when' or 'while' - add a time",
  "Use 'when' for something that happened at a point in time. Use 'while' for something happening at the same time.",
  [
    { note: "Simple:", text: "I read my book." },
    { note: "Complex with 'when':", text: "I read my book when mum cooked dinner." },
    { note: "Complex with 'while':", text: "I read my book while mum cooked dinner." },
  ],
  `SAY:
- 'When' = moment in time. 'While' = during the same time.

[Grammar | W5 S2 | Teach]`
);

practiseSlide(
  "Week 5", "Session 2",
  "Practise: add a 'when' or 'while' clause",
  "For each simple sentence, add a clause using 'when' or 'while'.",
  [
    "I did my homework.",
    "The dog slept.",
    "We waited.",
    "She sang.",
  ],
  `[Grammar | W5 S2 | Practise]`
);

checkSlide(
  "Week 5", "Session 2",
  "Check: expand with 'when' or 'while'.",
  "We cheered loudly.",
  "We cheered loudly when the team scored.  (or similar)",
  `[Grammar | W5 S2 | Check]`,
  `SAY:
- Add a time clause after 'when' or 'while'.`
);

// W5 S3: although/even though
teachSlide(
  "Week 5", "Session 3",
  "Expand with 'although' or 'even though' - show contrast",
  "Use 'although' or 'even though' when something unexpected happens despite another thing.",
  [
    { note: "Simple:", text: "She won the race." },
    { note: "Complex:", text: "She won the race although she was injured." },
    { note: "Pattern:", text: "main clause + although / even though + unexpected detail" },
  ],
  `SAY:
- 'Although' sets up a twist. Despite one thing, another thing happens.

[Grammar | W5 S3 | Teach]`
);

practiseSlide(
  "Week 5", "Session 3",
  "Practise: add an 'although' clause",
  "For each simple sentence, add a clause using 'although' or 'even though'.",
  [
    "We went outside.",
    "He smiled.",
    "I finished the book.",
    "The team celebrated.",
  ],
  `[Grammar | W5 S3 | Practise]`
);

checkSlide(
  "Week 5", "Session 3",
  "Check: expand with 'although'.",
  "She passed the test.",
  "She passed the test although she was nervous.  (or similar)",
  `[Grammar | W5 S3 | Check]`,
  `SAY:
- The second clause should contrast the first.`
);

// W5 S4: if
teachSlide(
  "Week 5", "Session 4",
  "Expand with 'if' - set up a condition",
  "Use 'if' to describe something that depends on a condition.",
  [
    { note: "Simple:", text: "We will play outside." },
    { note: "Complex with condition:", text: "We will play outside if it stops raining." },
    { note: "Pattern:", text: "main clause + if + condition" },
  ],
  `SAY:
- 'If' means 'on the condition that'.

[Grammar | W5 S4 | Teach]`
);

practiseSlide(
  "Week 5", "Session 4",
  "Practise: add an 'if' clause",
  "For each simple sentence, add an 'if' clause.",
  [
    "We will go to the park.",
    "I will bring a jacket.",
    "She will win.",
    "The plants will grow.",
  ],
  `[Grammar | W5 S4 | Practise]`
);

checkSlide(
  "Week 5", "Session 4",
  "Check: expand with 'if'.",
  "You can come with us.",
  "You can come with us if your parents agree.  (or similar)",
  `[Grammar | W5 S4 | Check]`,
  `SAY:
- The 'if' clause sets the condition.`
);

// W5 S5: Mixed
teachSlide(
  "Week 5", "Session 5",
  "Choose the right connector",
  "Pick the subordinator that matches your meaning: because (reason), when/while (time), although (contrast), if (condition).",
  [
    { note: "Reason", text: "I ate lunch early because I was hungry." },
    { note: "Time", text: "I ate lunch when the bell rang." },
    { note: "Contrast", text: "I ate lunch although I wasn't hungry." },
  ],
  `SAY:
- Decide your meaning first, then pick the connector.

[Grammar | W5 S5 | Teach]`
);

practiseSlide(
  "Week 5", "Session 5",
  "Practise: expand and connect",
  "Expand each simple sentence using a different connector (because, when, although, if).",
  [
    "We stayed home.",
    "He took a deep breath.",
    "She wore boots.",
    "We packed quickly.",
  ],
  `[Grammar | W5 S5 | Practise]`
);

checkSlide(
  "Week 5", "Session 5",
  "Check: choose the best connector.",
  "Combine:\n(a) I was tired.\n(b) I kept going.",
  "I kept going although I was tired.  (contrast)",
  `[Grammar | W5 S5 | Check]`,
  `SAY:
- Tiredness vs continuing -- that's a contrast, so 'although' fits.`
);

// ===========================================================================
//  WEEK 6 — DIRECT SPEECH (TAG AFTER)
// ===========================================================================

weekDivider(
  pres, 6, "Direct Speech - Speaker Tag After",
  "Punctuating direct speech when the 'he said / she said' comes AFTER the spoken words.",
  `[Week 6 | Speech tag after]`
);

// W6 S1: What is direct speech
teachSlide(
  "Week 6", "Session 1",
  "What is direct speech? The exact words spoken",
  "Direct speech shows the exact words a character says. We wrap those words in quotation marks.",
  [
    { note: "Direct speech",  text: '"I am hungry," said Sam.' },
    { note: "Not direct speech",  text: "Sam said that he was hungry." },
    { note: "Key",  text: "Direct = exact words in quote marks." },
  ],
  `SAY:
- Direct speech = exact words in quote marks.
- Indirect speech = reporting without quotes.

[Grammar | W6 S1 | Teach]`
);

practiseSlide(
  "Week 6", "Session 1",
  "Practise: direct or indirect?",
  "Label each D (direct) or I (indirect).",
  [
    '"Run!" shouted Ben.',
    "Ben told us to run.",
    '"I like pizza," said Mia.',
    "Mia said she liked pizza.",
  ],
  `[Grammar | W6 S1 | Practise]`
);

checkSlide(
  "Week 6", "Session 1",
  "Check: direct or indirect?",
  '"Let\'s go home," whispered Dan.',
  "D (direct) -- exact words in quote marks",
  `[Grammar | W6 S1 | Check]`,
  `SAY:
- Quote marks and exact words = direct speech.`
);

// W6 S2: Quotation marks
teachSlide(
  "Week 6", "Session 2",
  "Quotation marks: wrap the spoken words",
  "Open quote marks right before the first spoken word. Close quote marks right after the last piece of punctuation INSIDE the speech.",
  [
    { note: "Open and close quotes", text: '"I am tired," said Eli.' },
    { note: "Everything spoken is inside", text: '"Can we leave?" asked Ada.' },
    { note: "Comma / ? / ! stays inside", text: '"Stop!" yelled Jo.' },
  ],
  `SAY:
- Quote marks hug the exact words AND the punctuation that belongs to those words.

[Grammar | W6 S2 | Teach]`
);

practiseSlide(
  "Week 6", "Session 2",
  "Practise: add the quote marks",
  "Add quotation marks around the spoken words.",
  [
    "I am ready, said Tom.",
    "Watch out! yelled Sue.",
    "Where is it? asked Li.",
    "I will help, whispered Sam.",
  ],
  `[Grammar | W6 S2 | Practise]`
);

checkSlide(
  "Week 6", "Session 2",
  "Check: where do quote marks go?",
  "I can do it, said Kim.",
  '"I can do it," said Kim.',
  `[Grammar | W6 S2 | Check]`,
  `SAY:
- Open before 'I', close after the comma.`
);

// W6 S3: Comma inside quotes
teachSlide(
  "Week 6", "Session 3",
  "Comma inside the quotes (tag after)",
  "When the speaker tag comes AFTER the speech, end the spoken words with a comma INSIDE the closing quote mark.",
  [
    { note: "Comma before closing quote", text: '"I am here," said Eve.' },
    { note: "Still comma (not period)", text: '"I like it," replied Tim.' },
    { note: "Exception - ! or ? replaces comma", text: '"Wait!" shouted Ali.' },
  ],
  `SAY:
- Tag after -> comma inside the quote marks.
- If the speech is a question or exclamation, keep the ? or ! inside -- no extra comma needed.

[Grammar | W6 S3 | Teach]`
);

practiseSlide(
  "Week 6", "Session 3",
  "Practise: put in the comma",
  "Fix each sentence by adding the comma in the right place.",
  [
    '"That was fun" said Jo.',
    '"I am tired" mumbled Ned.',
    '"Come here" called Rae.',
    '"Look at this" whispered Leo.',
  ],
  `[Grammar | W6 S3 | Practise]`
);

checkSlide(
  "Week 6", "Session 3",
  "Check: punctuate correctly.",
  '"I love reading" said Pip.',
  '"I love reading," said Pip.',
  `[Grammar | W6 S3 | Check]`,
  `SAY:
- Comma inside the closing quote, before the speaker tag.`
);

// W6 S4: Capital letter
teachSlide(
  "Week 6", "Session 4",
  "Capital letter for the first spoken word",
  "The first spoken word always starts with a capital, even if the sentence is written in pieces.",
  [
    { note: "Capital inside the speech", text: '"Let\'s go," said Ava.' },
    { note: "Capital even after the opening quote", text: '"Where is it?" asked Jay.' },
    { note: "Speaker tag stays lowercase (unless name)", text: '"I agree," said Mia.' },
  ],
  `SAY:
- First spoken word = capital. Always.
- Speaker tag like 'said' stays lowercase -- unless it's a name.

[Grammar | W6 S4 | Teach]`
);

practiseSlide(
  "Week 6", "Session 4",
  "Practise: fix the capitals",
  "Fix the capitalisation in each sentence.",
  [
    '"let\'s start," said Ben.',
    '"this is cool," whispered Zoe.',
    '"wait for me," called Eli.',
    '"i can help," replied Ada.',
  ],
  `[Grammar | W6 S4 | Practise]`
);

checkSlide(
  "Week 6", "Session 4",
  "Check: fix the capitals.",
  '"come here," said tom.',
  '"Come here," said Tom.',
  `[Grammar | W6 S4 | Check]`,
  `SAY:
- 'Come' starts the speech -> capital. 'Tom' is a name -> capital. 'Said' stays lowercase.`
);

// W6 S5: Mixed
teachSlide(
  "Week 6", "Session 5",
  "Direct speech review (tag after)",
  "Steps: 1) Open quote marks. 2) Capital for the first spoken word. 3) Comma (or ! or ?) inside the closing quote. 4) Lowercase speaker tag. 5) Full stop at the very end.",
  [
    { note: "Full example", text: '"I finished my homework," said Kim.' },
    { note: "Question version",  text: '"Are you coming?" asked Sam.' },
    { note: "Exclamation version",  text: '"Watch out!" yelled Eve.' },
  ],
  `SAY:
- Five quick checks. Run them every time.

[Grammar | W6 S5 | Teach]`
);

practiseSlide(
  "Week 6", "Session 5",
  "Practise: punctuate fully",
  "Punctuate each line of direct speech properly.",
  [
    "i love chocolate said jo",
    "are we late asked ben",
    "stop that yelled sue",
    "i can help whispered mia",
  ],
  `[Grammar | W6 S5 | Practise]`
);

checkSlide(
  "Week 6", "Session 5",
  "Check: punctuate this line.",
  "i am finished said ada",
  '"I am finished," said Ada.',
  `[Grammar | W6 S5 | Check]`,
  `SAY:
- Capitals, quote marks, comma inside, full stop at the end.`
);

// ===========================================================================
//  WEEK 7 — DIRECT SPEECH (TAG BEFORE)
// ===========================================================================

weekDivider(
  pres, 7, "Direct Speech - Speaker Tag Before",
  "Punctuating direct speech when the speaker tag comes BEFORE the spoken words.",
  `[Week 7 | Speech tag before]`
);

// W7 S1: Tag at start
teachSlide(
  "Week 7", "Session 1",
  "Speaker tag before: the pattern",
  "When 'she said' comes first, we use this pattern: Speaker tag + comma + open quote + Capital first word + ...",
  [
    { note: "Pattern", text: 'Ben said, "I am tired."' },
    { note: "Pattern", text: 'Mia whispered, "Look over there."' },
    { note: "Note the comma BEFORE the quote", text: 'Sam called, "Come inside."' },
  ],
  `SAY:
- Tag before -> comma BEFORE the opening quote mark.
- The first spoken word is still a capital.

[Grammar | W7 S1 | Teach]`
);

practiseSlide(
  "Week 7", "Session 1",
  "Practise: rearrange (tag before)",
  "Rewrite each with the speaker tag first.",
  [
    '"I am cold," said Jo. -> Jo said, "I am cold."',
    '"Let\'s go," called Eli.',
    '"I agree," whispered Mia.',
    '"Where is it?" asked Tim.',
  ],
  `[Grammar | W7 S1 | Practise]`
);

checkSlide(
  "Week 7", "Session 1",
  "Check: write tag-first style.",
  'Jo: "I am here".',
  'Jo said, "I am here."',
  `[Grammar | W7 S1 | Check]`,
  `SAY:
- Tag first, then comma, then open quote, then capital.`
);

// W7 S2: Comma after tag
teachSlide(
  "Week 7", "Session 2",
  "Comma after the speaker tag",
  "When the tag comes before the speech, put a comma AFTER the tag and BEFORE the opening quote mark.",
  [
    { note: "Comma after tag", text: 'Lila said, "I love this song."' },
    { note: "Comma after tag", text: 'The coach shouted, "Run faster!"' },
    { note: "Comma after tag", text: 'Mum whispered, "It\'s a secret."' },
  ],
  `SAY:
- The comma sits between the tag and the opening quote mark.

[Grammar | W7 S2 | Teach]`
);

practiseSlide(
  "Week 7", "Session 2",
  "Practise: add the comma",
  "Add the missing comma.",
  [
    'Ned said "I\'m hungry."',
    'Ada whispered "Come look."',
    'The teacher called "Line up."',
    'Jay shouted "That\'s amazing!"',
  ],
  `[Grammar | W7 S2 | Practise]`
);

checkSlide(
  "Week 7", "Session 2",
  "Check: where does the comma go?",
  'Sam said "Let\'s begin."',
  'Sam said, "Let\'s begin."',
  `[Grammar | W7 S2 | Check]`,
  `SAY:
- Between 'said' and the opening quote.`
);

// W7 S3: Capital at start of speech
teachSlide(
  "Week 7", "Session 3",
  "Capital letter on the first spoken word",
  "Even after the speaker tag and comma, the first spoken word still starts with a capital letter.",
  [
    { note: "Capital", text: 'Ava said, "Good morning!"' },
    { note: "Capital", text: 'The dad replied, "That is fine."' },
    { note: "Capital", text: 'Eli laughed, "You\'re so funny!"' },
  ],
  `SAY:
- First spoken word = capital. Every time.

[Grammar | W7 S3 | Teach]`
);

practiseSlide(
  "Week 7", "Session 3",
  "Practise: fix the capital",
  "Fix the capital letter in each one.",
  [
    'Mia said, "good night."',
    'Dan whispered, "look up."',
    'The coach called, "run faster."',
    'Jo replied, "i\'m okay."',
  ],
  `[Grammar | W7 S3 | Practise]`
);

checkSlide(
  "Week 7", "Session 3",
  "Check: fix the capital.",
  'Kim said, "let\'s start."',
  'Kim said, "Let\'s start."',
  `[Grammar | W7 S3 | Check]`,
  `SAY:
- 'Let's' starts the speech -> capital L.`
);

// W7 S4: Ending punctuation inside quotes
teachSlide(
  "Week 7", "Session 4",
  "Ending punctuation sits INSIDE the closing quote mark",
  "The full stop, exclamation, or question mark that ends the spoken sentence goes inside the closing quote mark.",
  [
    { note: "Full stop inside", text: 'Sam said, "I am tired."' },
    { note: "Exclamation inside", text: 'Eve shouted, "Be careful!"' },
    { note: "Question mark inside", text: 'Jay asked, "Where are we?"' },
  ],
  `SAY:
- The ending punctuation lives inside the closing quote mark.

[Grammar | W7 S4 | Teach]`
);

practiseSlide(
  "Week 7", "Session 4",
  "Practise: punctuate inside the quotes",
  "Add the missing end punctuation in the right place.",
  [
    'Ada said, "I am ready"',
    'Tom shouted, "Watch out"',
    'Zoe asked, "Is it time"',
    'Ben whispered, "I hope so"',
  ],
  `[Grammar | W7 S4 | Practise]`
);

checkSlide(
  "Week 7", "Session 4",
  "Check: place the full stop.",
  'Kim replied, "I can come"',
  'Kim replied, "I can come."',
  `[Grammar | W7 S4 | Check]`,
  `SAY:
- Full stop goes inside the closing quote mark.`
);

// W7 S5: Mixed
teachSlide(
  "Week 7", "Session 5",
  "Tag-before review: five-step check",
  "1) Speaker tag. 2) Comma after tag. 3) Open quote. 4) Capital first word. 5) End punctuation inside closing quote.",
  [
    { note: "Five steps all present",  text: 'Mia said, "I am finished."' },
    { note: "Five steps with exclamation", text: 'The coach shouted, "Run faster!"' },
    { note: "Five steps with question", text: 'Dad asked, "Where are we going?"' },
  ],
  `SAY:
- Five-step check. Use it every time.

[Grammar | W7 S5 | Teach]`
);

practiseSlide(
  "Week 7", "Session 5",
  "Practise: write from scratch",
  "Write a direct speech sentence using each speaker tag. Tag before.",
  [
    "Mum said,",
    "The teacher asked,",
    "My friend shouted,",
    "The dog owner whispered,",
  ],
  `[Grammar | W7 S5 | Practise]`
);

checkSlide(
  "Week 7", "Session 5",
  "Check: punctuate fully.",
  "jo said let's head home",
  'Jo said, "Let\'s head home."',
  `[Grammar | W7 S5 | Check]`,
  `SAY:
- All five pieces: capital name, comma after tag, open quote, capital speech, end punctuation inside close quote.`
);

// ===========================================================================
//  WEEK 8 — SIMPLE vs COMPOUND
// ===========================================================================

weekDivider(
  pres, 8, "Simple vs Compound Sentences",
  "A compound sentence joins two simple sentences using a coordinating conjunction: and, but, so, or.",
  `[Week 8 | Compound sentences]`
);

// W8 S1: Simple sentence revisit
teachSlide(
  "Week 8", "Session 1",
  "Simple sentence: one main clause",
  "A simple sentence has one subject-verb pair. It expresses one complete idea.",
  [
    { note: "One main clause", text: "The sun rose." },
    { note: "One main clause", text: "We played soccer." },
    { note: "One main clause", text: "I read my book." },
  ],
  `SAY:
- One clause. One idea. Done.

[Grammar | W8 S1 | Teach]`
);

practiseSlide(
  "Week 8", "Session 1",
  "Practise: simple or not?",
  "Label each S (simple) or NS (not simple).",
  [
    "The rain fell.",
    "I tidied the room, and I washed the car.",
    "She laughed.",
    "We played outside, but it started to rain.",
  ],
  `[Grammar | W8 S1 | Practise]`
);

checkSlide(
  "Week 8", "Session 1",
  "Check: simple or not?",
  "The cat climbed the tree.",
  "Simple -- one subject-verb pair",
  `[Grammar | W8 S1 | Check]`,
  `SAY:
- One clause -> simple.`
);

// W8 S2: Compound
teachSlide(
  "Week 8", "Session 2",
  "Compound sentence: two main clauses joined by 'and, but, so, or'",
  "A compound sentence joins two complete simple sentences with a coordinating conjunction (and, but, so, or), with a comma before the conjunction.",
  [
    { note: "Two main clauses + conjunction", text: "I was hungry, so I made a snack." },
    { note: "Two main clauses + conjunction", text: "We tried hard, but we lost." },
    { note: "Two main clauses + conjunction", text: "You can come, or you can stay." },
  ],
  `SAY:
- Two main clauses. One conjunction. One comma.

[Grammar | W8 S2 | Teach]`
);

practiseSlide(
  "Week 8", "Session 2",
  "Practise: spot compound sentences",
  "Label S (simple) or C (compound). For compounds, circle the conjunction.",
  [
    "I opened the door.",
    "I opened the door, and the dog ran out.",
    "We stayed inside, so we played games.",
    "She sang loudly.",
  ],
  `[Grammar | W8 S2 | Practise]`
);

checkSlide(
  "Week 8", "Session 2",
  "Check: compound or simple?",
  "We were tired, but we kept walking.",
  "Compound -- two main clauses joined by 'but'",
  `[Grammar | W8 S2 | Check]`,
  `SAY:
- Two clauses, each with subject and verb. 'But' joins them. Compound.`
);

// W8 S3: FANBOYS introduction
teachSlide(
  "Week 8", "Session 3",
  "Coordinating conjunctions: for, and, nor, but, or, yet, so",
  "We focus on the four most common in Year 5/6 writing: and, but, so, or. Pick the one that matches the relationship between your two clauses.",
  [
    { note: "and = adds", text: "I finished my homework, and I tidied my room." },
    { note: "but = contrasts", text: "It was cold, but we went outside." },
    { note: "so = result", text: "It was late, so we went home." },
  ],
  `SAY:
- and (adds), but (contrasts), so (result), or (choice).

[Grammar | W8 S3 | Teach]`
);

practiseSlide(
  "Week 8", "Session 3",
  "Practise: choose the right conjunction",
  "Join each pair with and, but, so, or or.",
  [
    "(a) I was hungry. (b) I ate a snack.",
    "(a) The bus was full. (b) We waited for the next one.",
    "(a) You can watch TV. (b) You can read.",
    "(a) I tried my best. (b) I didn't win.",
  ],
  `[Grammar | W8 S3 | Practise]`
);

checkSlide(
  "Week 8", "Session 3",
  "Check: choose the best conjunction.",
  "Join: (a) It started to rain. (b) We went inside.",
  "It started to rain, so we went inside.  ('so' shows the result)",
  `[Grammar | W8 S3 | Check]`,
  `SAY:
- Rain caused us going inside -- 'so' is the best fit.`
);

// W8 S4: Joining practice
teachSlide(
  "Week 8", "Session 4",
  "Join two simple sentences: comma + conjunction",
  "Combine two simple sentences into one compound sentence. Remember: comma before the conjunction.",
  [
    { note: "Two simples:", text: "It was sunny. We went to the beach." },
    { note: "Joined:", text: "It was sunny, so we went to the beach." },
    { note: "Key rule:", text: "comma BEFORE the conjunction" },
  ],
  `SAY:
- Two simples -> one compound. Comma before the conjunction every time.

[Grammar | W8 S4 | Teach]`
);

practiseSlide(
  "Week 8", "Session 4",
  "Practise: join into a compound",
  "Join each pair. Use a comma and the best conjunction.",
  [
    "I love pizza. I don't like olives.",
    "He finished early. He helped his friend.",
    "We can go now. We can wait.",
    "She felt sick. She stayed home.",
  ],
  `[Grammar | W8 S4 | Practise]`
);

checkSlide(
  "Week 8", "Session 4",
  "Check: join them.",
  "(a) The pool was cold. (b) I still jumped in.",
  "The pool was cold, but I still jumped in.",
  `[Grammar | W8 S4 | Check]`,
  `SAY:
- Contrast = 'but'. Comma before 'but'.`
);

// W8 S5: Mixed
teachSlide(
  "Week 8", "Session 5",
  "Review: simple or compound?",
  "Check: one clause = simple. Two main clauses joined by and/but/so/or = compound.",
  [
    { note: "Simple", text: "We ran home." },
    { note: "Compound", text: "We ran home, and we changed quickly." },
    { note: "Compound", text: "I can cook, but I prefer eating out." },
  ],
  `SAY:
- Same test as before: count main clauses.

[Grammar | W8 S5 | Teach]`
);

practiseSlide(
  "Week 8", "Session 5",
  "Practise: mixed",
  "Label each S (simple) or C (compound). Circle the conjunction in any compound.",
  [
    "The storm passed.",
    "The storm passed, and the sun came out.",
    "I can do it, but I need help.",
    "We walked home slowly.",
  ],
  `[Grammar | W8 S5 | Practise]`
);

checkSlide(
  "Week 8", "Session 5",
  "Check: label and explain.",
  "Label:\n1. I laughed.\n2. I laughed, and everyone joined in.",
  "1. S (one clause)   2. C (two main clauses + 'and')",
  `[Grammar | W8 S5 | Check]`,
  `SAY:
- First has one clause. Second has two main clauses joined with 'and'.`
);

// ===========================================================================
//  WEEK 9 — IDENTIFY MAIN CLAUSES + CONJUNCTIONS
// ===========================================================================

weekDivider(
  pres, 9, "Two Main Clauses + Coordinating Conjunction",
  "Zoom in on compound sentences: find each main clause and the coordinating conjunction that joins them.",
  `[Week 9 | Find both main clauses]`
);

// W9 S1: Finding the two main clauses
teachSlide(
  "Week 9", "Session 1",
  "Finding the two main clauses",
  "Cover the conjunction. Each side should be a complete sentence that stands alone. If yes, you have two main clauses.",
  [
    { note: "Main clause 1", text: "[I finished my homework]  and I tidied my room." },
    { note: "Main clause 2", text: "I finished my homework and  [I tidied my room]." },
    { note: "Both sides stand alone -> compound", text: "'I finished my homework.'  'I tidied my room.'" },
  ],
  `SAY:
- Cover the conjunction. Can each side stand alone as a sentence? Yes -> two main clauses.

[Grammar | W9 S1 | Teach]`
);

practiseSlide(
  "Week 9", "Session 1",
  "Practise: split into two clauses",
  "For each compound sentence, write out the two main clauses separately.",
  [
    "I was tired, but I kept reading.",
    "We won the match, so we celebrated.",
    "You can walk there, or you can catch the bus.",
    "She smiled, and he waved.",
  ],
  `[Grammar | W9 S1 | Practise]`
);

checkSlide(
  "Week 9", "Session 1",
  "Check: split into two.",
  "It rained, so we stayed inside.",
  "Clause 1: 'It rained.'   Clause 2: 'We stayed inside.'",
  `[Grammar | W9 S1 | Check]`,
  `SAY:
- Cover 'so'. Each side stands alone. Two main clauses.`
);

// W9 S2: Spotting coordinating conjunctions
teachSlide(
  "Week 9", "Session 2",
  "Spot the coordinating conjunction",
  "Look for the small joining word that connects the two main clauses. Most often: and, but, so, or.",
  [
    { note: "'and' joins", text: "I swept the floor, and I washed the dishes." },
    { note: "'but' joins", text: "I tried hard, but I didn't win." },
    { note: "'so' joins", text: "I was full, so I stopped eating." },
  ],
  `SAY:
- Spot the little joining word.

[Grammar | W9 S2 | Teach]`
);

practiseSlide(
  "Week 9", "Session 2",
  "Practise: circle the coordinating conjunction",
  "Circle the conjunction in each compound sentence.",
  [
    "I felt calm, and I slept well.",
    "We wanted to play, but it was dark.",
    "She was thirsty, so she had a drink.",
    "You can stay, or you can leave.",
  ],
  `[Grammar | W9 S2 | Practise]`
);

checkSlide(
  "Week 9", "Session 2",
  "Check: which word joins the clauses?",
  "The cake was sweet, but the icing was too sugary.",
  "but",
  `[Grammar | W9 S2 | Check]`,
  `SAY:
- 'But' joins the two clauses -- it signals a contrast.`
);

// W9 S3: Label both main clauses
teachSlide(
  "Week 9", "Session 3",
  "Label both main clauses and the conjunction",
  "Mark main clause 1 with [MC1], the conjunction with [CC], and main clause 2 with [MC2].",
  [
    { note: "Full label", text: "[MC1 I was late] [CC, so] [MC2 I ran to school]." },
    { note: "Full label", text: "[MC1 She called] [CC, and] [MC2 we answered]." },
    { note: "Full label", text: "[MC1 He tried] [CC, but] [MC2 he missed]." },
  ],
  `SAY:
- Three labels: MC1, CC, MC2.

[Grammar | W9 S3 | Teach]`
);

practiseSlide(
  "Week 9", "Session 3",
  "Practise: label the parts",
  "Add [MC1], [CC], and [MC2] labels to each sentence.",
  [
    "We set the table, and we served dinner.",
    "I was nervous, but I read my speech.",
    "The wind was strong, so we closed the window.",
    "Take the apple, or take the orange.",
  ],
  `[Grammar | W9 S3 | Practise]`
);

checkSlide(
  "Week 9", "Session 3",
  "Check: label fully.",
  "I packed lunch, and I grabbed my bag.",
  "[MC1 I packed lunch] [CC , and] [MC2 I grabbed my bag]",
  `[Grammar | W9 S3 | Check]`,
  `SAY:
- Three labels: MC1, CC, MC2.`
);

// W9 S4: Mixed identification
teachSlide(
  "Week 9", "Session 4",
  "Is it really compound? Two checks",
  "Check 1: is the joining word a coordinating conjunction (and, but, so, or)? Check 2: can each side stand alone as a sentence?",
  [
    { note: "Compound -- passes both checks", text: "I was ready, and we left early." },
    { note: "NOT compound -- 'because' is subordinating", text: "I was ready because we left early." },
    { note: "NOT compound -- second part can't stand alone", text: "I was ready, and left early." },
  ],
  `SAY:
- Some sentences look compound but aren't. Use both checks.

TEACHER NOTES:
The subordinating vs coordinating distinction is subtle. Keep it simple: 'and/but/so/or with two stand-alone clauses'.

[Grammar | W9 S4 | Teach]`
);

practiseSlide(
  "Week 9", "Session 4",
  "Practise: compound or not?",
  "Label each C (compound) or N (not compound) and say why.",
  [
    "I was hungry, so I ate.",
    "I was hungry because I skipped lunch.",
    "The rain fell, and the streets flooded.",
    "We went home after the show ended.",
  ],
  `[Grammar | W9 S4 | Practise]`
);

checkSlide(
  "Week 9", "Session 4",
  "Check: is this compound?",
  "I stayed home, and I rested.",
  "Compound -- 'and' joins two stand-alone clauses",
  `[Grammar | W9 S4 | Check]`,
  `SAY:
- 'I stayed home' and 'I rested' both stand alone. 'And' joins them. Compound.`
);

// W9 S5: Mixed
teachSlide(
  "Week 9", "Session 5",
  "Compound sentence review",
  "Three things to find every time: MC1, the coordinating conjunction, and MC2. The comma always sits before the conjunction.",
  [
    { note: "All three parts visible", text: "The coach called out, and we came running." },
    { note: "All three parts visible", text: "I was tired, but I finished the race." },
    { note: "All three parts visible", text: "He looked up, so I waved back." },
  ],
  `SAY:
- Three-part pattern: MC1, comma + CC, MC2.

[Grammar | W9 S5 | Teach]`
);

practiseSlide(
  "Week 9", "Session 5",
  "Practise: label and explain",
  "Find MC1, CC, and MC2 in each sentence. Write them out.",
  [
    "The bell rang, and the children lined up.",
    "I felt scared, but I jumped in.",
    "It was late, so we went to bed.",
    "You can sit here, or you can join them.",
  ],
  `[Grammar | W9 S5 | Practise]`
);

checkSlide(
  "Week 9", "Session 5",
  "Check: label the three parts.",
  "I finished the test, and I packed my things.",
  "MC1: 'I finished the test'   CC: ', and'   MC2: 'I packed my things'",
  `[Grammar | W9 S5 | Check]`,
  `SAY:
- Three clean parts: MC1, CC, MC2.`
);

// ===========================================================================
//  WEEK 10 — DIRECT vs INDIRECT SPEECH
// ===========================================================================

weekDivider(
  pres, 10, "Direct vs Indirect Speech",
  "Direct speech uses the speaker's exact words. Indirect speech reports what was said.",
  `[Week 10 | Direct vs indirect]`
);

// W10 S1: Direct vs indirect
teachSlide(
  "Week 10", "Session 1",
  "Direct = exact words; Indirect = reported",
  "Direct speech: the exact words in quotation marks. Indirect speech: a report of what was said (no quote marks, often with 'that').",
  [
    { note: "Direct", text: '"I am going home," said Ben.' },
    { note: "Indirect", text: "Ben said that he was going home." },
    { note: "Key changes", text: "No quotes. Pronoun and verb tense may shift." },
  ],
  `SAY:
- Direct = quotes around exact words.
- Indirect = a reported version, no quotes, often 'that'.

TEACHER NOTES:
Focus on the pattern difference first. Tense shift is nuanced -- keep examples simple.

[Grammar | W10 S1 | Teach]`
);

practiseSlide(
  "Week 10", "Session 1",
  "Practise: direct or indirect?",
  "Label each D or I.",
  [
    '"We are leaving," said Mum.',
    "Mum said we were leaving.",
    "The coach said that we played well.",
    '"You played well," said the coach.',
  ],
  `[Grammar | W10 S1 | Practise]`
);

checkSlide(
  "Week 10", "Session 1",
  "Check: which one is direct?",
  'A. Jo said, "I am tired."\nB. Jo said that he was tired.',
  "A (direct) -- exact words in quote marks",
  `[Grammar | W10 S1 | Check]`,
  `SAY:
- A has the exact words in quote marks. Direct.`
);

// W10 S2: Convert direct to indirect
teachSlide(
  "Week 10", "Session 2",
  "Convert direct to indirect",
  "Three moves: 1) Remove quotation marks. 2) Add 'that' (often). 3) Shift pronouns and tense to match the report.",
  [
    { note: "Direct -> Indirect", text: '"I am happy," said Mia.  ->  Mia said that she was happy.' },
    { note: "Direct -> Indirect", text: '"We will win," said Tom.  ->  Tom said that they would win.' },
    { note: "Direct -> Indirect", text: '"I like pizza," said Jay.  ->  Jay said that he liked pizza.' },
  ],
  `SAY:
- Three moves. Quotes out. 'That' in. Pronouns and tense shift.

TEACHER NOTES:
Verb tense typically shifts backwards: am -> was, will -> would, like -> liked.

[Grammar | W10 S2 | Teach]`
);

practiseSlide(
  "Week 10", "Session 2",
  "Practise: convert to indirect",
  "Rewrite each as indirect speech.",
  [
    '"I am going," said Ada.',
    '"We are ready," said the team.',
    '"I will call you," said Sam.',
    '"I like this book," said Ellie.',
  ],
  `[Grammar | W10 S2 | Practise]`
);

checkSlide(
  "Week 10", "Session 2",
  "Check: convert to indirect.",
  '"I am tired," said Ben.',
  "Ben said that he was tired.",
  `[Grammar | W10 S2 | Check]`,
  `SAY:
- Quotes out. 'That' in. 'I' -> 'he', 'am' -> 'was'.`
);

// W10 S3: Tag before or after review
teachSlide(
  "Week 10", "Session 3",
  "Direct speech punctuation: tag before OR after",
  "Tag BEFORE: Name + comma + \"Capital ... ending punctuation.\"\nTag AFTER: \"Capital ... ,\" said Name.",
  [
    { note: "Tag before", text: 'Kim said, "I can do it."' },
    { note: "Tag after", text: '"I can do it," said Kim.' },
    { note: "Same meaning, different punctuation", text: "Both are direct speech. Punctuation shifts based on tag position." },
  ],
  `SAY:
- Same sentence, two shapes. Review both punctuation rules.

[Grammar | W10 S3 | Teach]`
);

practiseSlide(
  "Week 10", "Session 3",
  "Practise: write it both ways",
  "For each spoken line, write it as tag-before AND tag-after.",
  [
    "Ben: I am tired.",
    "Mia: Let's go now.",
    "Sam: I found it.",
    "Jo: We are late.",
  ],
  `[Grammar | W10 S3 | Practise]`
);

checkSlide(
  "Week 10", "Session 3",
  "Check: write it both ways.",
  "Ada said she was finished.",
  'Tag before: Ada said, "I am finished."\nTag after: "I am finished," said Ada.',
  `[Grammar | W10 S3 | Check]`,
  `SAY:
- Same idea, two punctuation shapes.`
);

// W10 S4: Writing both forms
teachSlide(
  "Week 10", "Session 4",
  "Write both direct and indirect versions",
  "Given one, write the other. This tests both punctuation (direct) and conversion (indirect).",
  [
    { note: "Given direct -> write indirect", text: '"I love maths," said Pip. -> Pip said that he loved maths.' },
    { note: "Given indirect -> write direct", text: 'The teacher said the test was on Friday. -> The teacher said, "The test is on Friday."' },
    { note: "Watch the tense shift", text: "Direct 'is' <-> Indirect 'was'" },
  ],
  `SAY:
- Given one, write the other. Watch the tense shift.

[Grammar | W10 S4 | Teach]`
);

practiseSlide(
  "Week 10", "Session 4",
  "Practise: write the other version",
  "For each one, write the opposite (direct becomes indirect, or indirect becomes direct).",
  [
    '"I am hungry," said Ted.',
    "Mum said that she would call soon.",
    '"We can help," said the twins.',
    "The coach said we had won.",
  ],
  `[Grammar | W10 S4 | Practise]`
);

checkSlide(
  "Week 10", "Session 4",
  "Check: convert.",
  "Ben said that he was cold.",
  'Ben said, "I am cold."  (or: "I am cold," said Ben.)',
  `[Grammar | W10 S4 | Check]`,
  `SAY:
- Indirect -> direct: add quotes, shift 'he' -> 'I', 'was' -> 'am'.`
);

// W10 S5: Mixed
teachSlide(
  "Week 10", "Session 5",
  "Direct and indirect speech review",
  "Full revision: identify, convert, and punctuate. Use the five-step check for direct speech and the three-move conversion for indirect.",
  [
    { note: "Identify", text: "Which is direct? Which is indirect?" },
    { note: "Convert", text: "Rewrite from one form to the other." },
    { note: "Punctuate", text: "Apply the correct punctuation to direct speech." },
  ],
  `SAY:
- All three skills from this term, together.

[Grammar | W10 S5 | Teach]`
);

practiseSlide(
  "Week 10", "Session 5",
  "Practise: do all three",
  "For each sentence: 1) label D or I. 2) Rewrite in the opposite form.",
  [
    '"I am going," said Sam.',
    "Jo said that she had finished.",
    '"We will win," shouted the team.',
    "The teacher said that we were doing well.",
  ],
  `[Grammar | W10 S5 | Practise]`
);

checkSlide(
  "Week 10", "Session 5",
  "Check: convert with proper punctuation.",
  "Sarah said that she liked the song.",
  'Sarah said, "I like the song."   (or: "I like the song," said Sarah.)',
  `[Grammar | W10 S5 | Check]`,
  `SAY:
- Quotes, capital, comma, pronoun shift 'she' -> 'I', tense shift 'liked' -> 'like'.`
);

// ────────────────────────────────────────────────────────────────────────────
//  Closing slide
// ────────────────────────────────────────────────────────────────────────────

closingSlide(
  pres,
  "Which grammar skill have you used most in your writing this term? Share an example with your partner.",
  [
    "Grammar is a 5-minute routine -- drip-feed it, don't cram it.",
    "Each session: Teach, Practise, Check.",
    "Return to these patterns often -- retrieval builds fluency.",
  ],
  `SAY:
- Grammar patterns stack. Clauses and phrases lead into sentence types and speech punctuation.

DO:
- Revisit weeks when the need comes up in writing.

TEACHER NOTES:
No printed student resources are needed -- students write in their OG Books.

[Grammar | Closing]`
);

// ────────────────────────────────────────────────────────────────────────────
//  Write
// ────────────────────────────────────────────────────────────────────────────

(async () => {
  const outFile = path.join(OUT_DIR, "Grammar Year 56 Routines.pptx");
  await pres.writeFile({ fileName: outFile });
  console.log(`PPTX written to ${outFile}`);
})();
