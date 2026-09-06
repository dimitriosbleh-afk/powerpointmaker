"use strict";

// CSR (Orton-Gillingham) 5-minute spelling routines — Year 5/6
// Weeks 1, 3, 5, 7, 9 — 5 sessions per week — 25 sessions total.
// Each session = 3 slides: Teach, Practise, Check.
// Short, sharp, sweet: slots into a larger teaching block.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");

const T = createTheme("literacy", "grade56", 0);
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, closingSlide,
  addCard, addFooter, addTopBar, addBadge, addTitle, addTextOnShape,
  withReveal, runSlideDiagnostics,
} = T;

const FOOTER = "CSR Spelling | Year 5/6 | 5-minute routine";
const OUT_DIR = "output/CSR_Year56_Spelling";

fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "CSR Spelling — Year 5/6 (Weeks 1, 3, 5, 7, 9)";

// ────────────────────────────────────────────────────────────────────────────
//  Stage badge + helper builders
// ────────────────────────────────────────────────────────────────────────────

function addStagePill(slide, label, color, x, y, w) {
  slide.addShape("roundRect", {
    x: x || 0.5, y: y || 1.38, w: w || 1.45, h: 0.32, rectRadius: 0.06,
    fill: { color },
  });
  slide.addText(label, {
    x: x || 0.5, y: y || 1.38, w: w || 1.45, h: 0.32,
    fontSize: 10, fontFace: FONT_B, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// Section divider: week banner
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
    fontSize: 36, fontFace: FONT_H, color: C.WHITE, bold: true, margin: 0,
  });
  s.addText("Focus:", {
    x: 0.7, y: 2.95, w: 1.2, h: 0.35,
    fontSize: 12, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
  });
  s.addText(focus, {
    x: 0.7, y: 3.35, w: 8.5, h: 1.2,
    fontSize: 17, fontFace: FONT_B, color: C.WHITE, margin: 0,
  });
  s.addText("5-minute routine  |  5 sessions this week", {
    x: 0.7, y: 4.75, w: 8, h: 0.3,
    fontSize: 12, fontFace: FONT_B, color: C.ACCENT, margin: 0,
  });

  if (notes) s.addNotes(notes);
  return s;
}

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
//  Mini-session builders (3 slides: Teach → Practise → Check)
// ────────────────────────────────────────────────────────────────────────────

/**
 * Teach slide: rule card (top) + examples row (bottom).
 * Renders a session divider slide first to mark the session boundary.
 * examples: array of { word, highlight } where highlight is the target chunk.
 */
function teachSlide(weekLabel, sessionLabel, title, rule, examples, notes) {
  sessionDivider(weekLabel, sessionLabel, title);

  const s = pres.addSlide();
  addTopBar(s, C.PRIMARY);
  addBadge(s, "Teach", { color: C.PRIMARY, w: 1.25 });
  addBadge(s, weekLabel, { color: C.SECONDARY, w: 2.0, x: 1.85 });
  addBadge(s, sessionLabel, { color: C.ACCENT, w: 1.5, x: 3.95 });
  addTitle(s, title);

  // Rule card (top)
  const ruleH = 1.25;
  addCard(s, 0.5, CONTENT_TOP, 9, ruleH, { strip: C.PRIMARY, fill: C.BG_CARD });
  s.addText("The rule", {
    x: 0.75, y: CONTENT_TOP + 0.10, w: 3, h: 0.28,
    fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });
  s.addText(rule, {
    x: 0.75, y: CONTENT_TOP + 0.42, w: 8.5, h: ruleH - 0.50,
    fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, margin: 0, valign: "top",
  });

  // Examples row (bottom)
  const exY = CONTENT_TOP + ruleH + 0.15;
  const exH = SAFE_BOTTOM - exY;
  const n = Math.min(examples.length, 4);
  const gap = 0.15;
  const totalW = 9;
  const cellW = (totalW - (n - 1) * gap) / n;

  s.addText("Examples", {
    x: 0.5, y: exY, w: 3, h: 0.28,
    fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });

  const boxY = exY + 0.36;
  const boxH = exH - 0.40;
  for (let i = 0; i < n; i++) {
    const ex = examples[i];
    const x = 0.5 + i * (cellW + gap);
    addCard(s, x, boxY, cellW, boxH, { strip: i % 2 === 0 ? C.SECONDARY : C.ACCENT, fill: C.WHITE });

    s.addText(ex.word, {
      x: x + 0.15, y: boxY + 0.18, w: cellW - 0.30, h: 0.55,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0, fit: "shrink",
    });
    if (ex.highlight) {
      s.addText(ex.highlight, {
        x: x + 0.15, y: boxY + boxH - 0.50, w: cellW - 0.30, h: 0.36,
        fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  }

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

/**
 * Practise slide: prompt card (left) + word list or sort (right).
 * items: array of strings to work with.
 */
function practiseSlide(weekLabel, sessionLabel, title, prompt, items, notes) {
  const s = pres.addSlide();
  addTopBar(s, C.SECONDARY);
  addBadge(s, "Practise", { color: C.SECONDARY, w: 1.4 });
  addBadge(s, weekLabel, { color: C.PRIMARY, w: 2.0, x: 2.0 });
  addBadge(s, sessionLabel, { color: C.ACCENT, w: 1.5, x: 4.1 });
  addTitle(s, title);

  // Left prompt card
  const cardH = SAFE_BOTTOM - CONTENT_TOP;
  addCard(s, 0.5, CONTENT_TOP, 4.3, cardH, { strip: C.SECONDARY, fill: C.WHITE });
  s.addText("In your OG Books", {
    x: 0.7, y: CONTENT_TOP + 0.14, w: 3.9, h: 0.32,
    fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
  });
  s.addText(prompt, {
    x: 0.7, y: CONTENT_TOP + 0.52, w: 3.9, h: cardH - 0.70,
    fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
    paraSpaceAfter: 4,
  });

  // Right items card
  addCard(s, 5.0, CONTENT_TOP, 4.5, cardH, { strip: C.ACCENT, fill: C.BG_CARD });
  s.addText("Words", {
    x: 5.2, y: CONTENT_TOP + 0.14, w: 3, h: 0.32,
    fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
  });

  // Render items in a neat grid (2 cols if > 4 items)
  const cols = items.length > 4 ? 2 : 1;
  const rows = Math.ceil(items.length / cols);
  const gx = 5.2;
  const gy = CONTENT_TOP + 0.55;
  const gw = 4.1;
  const gh = cardH - 0.70;
  const cellW = cols === 2 ? (gw - 0.10) / 2 : gw;
  const cellH = gh / rows;

  items.forEach((word, i) => {
    const r = Math.floor(i / cols);
    const col = i % cols;
    const x = gx + col * (cellW + 0.10);
    const y = gy + r * cellH;
    s.addText(`${i + 1}.  ${word}`, {
      x, y, w: cellW, h: cellH - 0.04,
      fontSize: cols === 2 ? 14 : 18, fontFace: FONT_H, color: C.CHARCOAL,
      bold: true, valign: "middle", margin: 0,
    });
  });

  addFooter(s, FOOTER);
  if (notes) s.addNotes(notes);
  runSlideDiagnostics(s, pres, { respectSafeBottom: false });
  return s;
}

/**
 * Check slide (with reveal): a quick CFU question with answer underneath.
 * Uses withReveal pair — second slide shows the answer strip.
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

      // Question card
      addCard(s, 0.5, CONTENT_TOP, 9, 2.0, { strip: C.ALERT, fill: C.WHITE });
      s.addText("Quick check", {
        x: 0.75, y: CONTENT_TOP + 0.12, w: 3, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
      });
      s.addText(question, {
        x: 0.75, y: CONTENT_TOP + 0.44, w: 8.5, h: 1.5,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(notes || "");
      return s;
    },
    (s) => {
      // Answer bar below the question card
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
        fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
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
  "CSR Spelling Routines",
  "Orton-Gillingham inspired 5-minute sessions",
  "Year 5/6  |  Weeks 1, 3, 5, 7, 9",
  `SAY:
- This deck is a 5-minute spelling routine that runs inside a bigger teaching block.
- Each session has three slides: Teach, Practise, then a quick Check.

DO:
- Open the deck to the session you need for the day.
- Have OG Books ready before you start.

TEACHER NOTES:
Short, sharp, sweet routines. Do not stretch a session beyond about five minutes.

WATCH FOR:
- Sessions are cumulative within a week -- do not skip if students missed a day.

[CSR | 5-minute routine]`
);

// ===========================================================================
//  WEEK 1 — PLURAL RULES (-s / -es / -ies)
// ===========================================================================

weekDivider(
  pres,
  1,
  "Plural Rules",
  "How do we make a word plural? Most words take -s, but some need -es, and some change y to i and add -es.",
  `SAY:
- This week's focus is plural rules.
- We will look at -s, -es, and -ies across the five sessions.

DO:
- Display the banner as a transition into the CSR block.

TEACHER NOTES:
Each session targets a single rule, then session 5 mixes them.

[Week 1 | Plurals]`
);

// ── Week 1 Session 1: Basic -s rule ─────────────────────────────────────────
teachSlide(
  "Week 1", "Session 1",
  "Plural -s: add -s to most nouns",
  "For most singular nouns, make them plural by adding -s. No other change.",
  [
    { word: "cat -> cats",  highlight: "+ s" },
    { word: "book -> books", highlight: "+ s" },
    { word: "desk -> desks", highlight: "+ s" },
    { word: "star -> stars", highlight: "+ s" },
  ],
  `SAY:
- Watch me. The word 'cat' becomes 'cats'. I just add one s at the end.
- My rule: for most nouns, I just add s. No spelling change before it.

DO:
- Read each example aloud.
- Underline or point to the final s in each plural.

TEACHER NOTES:
This is the default plural rule. Establish it first so students have a base to compare against in later sessions.

WATCH FOR:
- Students writing 'cat's' with an apostrophe -- correct immediately: apostrophe is for possession, not plurals.

[CSR | Week 1 Session 1 | Teach]`
);

practiseSlide(
  "Week 1", "Session 1",
  "Practise: make these nouns plural",
  "Write each word, then write its plural.\n\nRead your plural out loud.\n\nCheck: did you just add s?",
  ["dog", "book", "hand", "lamp", "tree", "chair"],
  `SAY:
- Open your OG Books. Write each word, then write its plural beside it.

DO:
- Time 90 seconds.
- Circulate, checking for apostrophe errors.
- Ask two students to share one plural each.

TEACHER NOTES:
Keep the list short. The point is the rule, not volume.

WATCH FOR:
- 'boxs' or 'dishs' errors will appear later -- for now, stay on -s only.

[CSR | Week 1 Session 1 | Practise]`
);

checkSlide(
  "Week 1", "Session 1",
  "Check: what is the plural of 'pencil'?",
  "Write the plural of the word 'pencil' in your OG Book. Show it to me when you are ready.",
  "pencils",
  `SAY:
- Think. Write. Show me your OG Book when ready.

DO:
- Scan all OG Books. Look for 'pencils'.
- 80% correct -> move on. Below that -> reteach with one more example.

CFU CHECKPOINT:
Technique: Show Me OG Books
- Scan for: correct 'pencils' with final s, no apostrophe.

WATCH FOR:
- 'pencil's' with apostrophe -- quick correction in the moment.

[CSR | Week 1 Session 1 | Check]`,
  `SAY:
- The plural of pencil is pencils. Just add s.

[CSR | Week 1 Session 1 | Reveal]`
);

// ── Week 1 Session 2: -es after hissing sounds ──────────────────────────────
teachSlide(
  "Week 1", "Session 2",
  "Plural -es: after s, ss, sh, ch, x, z",
  "If the word ends in a hissing sound -- s, ss, sh, ch, x, or z -- add -es. It gives us a new syllable we can hear.",
  [
    { word: "bus -> buses",   highlight: "+ es" },
    { word: "dish -> dishes", highlight: "+ es" },
    { word: "box -> boxes",   highlight: "+ es" },
    { word: "buzz -> buzzes", highlight: "+ es" },
  ],
  `SAY:
- Listen. 'Bus'. If I just add s, I get 'buss'. That does not work.
- I need -es. 'Bus-es'. I can hear the extra syllable.

DO:
- Say each plural slowly so students hear the extra /ez/ syllable.
- Point at the ending letters on the board.

TEACHER NOTES:
The 'hissing sound' cue is more reliable than memorising letters. Say them, do not spell them.

WATCH FOR:
- Students saying 'boxs' -- model the sound trap.

[CSR | Week 1 Session 2 | Teach]`
);

practiseSlide(
  "Week 1", "Session 2",
  "Practise: does it need -s or -es?",
  "Sort each word into two columns in your OG Book.\n\nLabel them -s and -es.\n\nWrite each plural beside the word.",
  ["bus", "cat", "fox", "dish", "lamp", "kiss"],
  `SAY:
- Two columns: -s on the left, -es on the right. Sort each word.

DO:
- 90 seconds.
- Pair-check before you reveal.

TEACHER NOTES:
Sorting forces the rule into use, not just recall.

WATCH FOR:
- Over-applying -es to simple words ('lamps' not 'lampes').

[CSR | Week 1 Session 2 | Practise]`
);

checkSlide(
  "Week 1", "Session 2",
  "Check: plural of 'brush'?",
  "Write the plural of 'brush'. Whisper the ending -- do you hear -s or -es?",
  "brushes",
  `SAY:
- Say it out loud before you write. Hearing helps you choose.

DO:
- Show me.
- Look for the -es ending and the extra syllable.

CFU CHECKPOINT:
Technique: Show Me OG Books
- Scan for: 'brushes' with -es.

[CSR | Week 1 Session 2 | Check]`,
  `SAY:
- Brush ends in -sh, a hissing sound. So we add -es. Brushes.

[CSR | Week 1 Session 2 | Reveal]`
);

// ── Week 1 Session 3: consonant + y -> -ies ─────────────────────────────────
teachSlide(
  "Week 1", "Session 3",
  "Plural -ies: consonant + y -> change y to i, add es",
  "If a word ends in a consonant before y, change the y to i and add -es.",
  [
    { word: "baby -> babies",   highlight: "y -> ies" },
    { word: "city -> cities",   highlight: "y -> ies" },
    { word: "pony -> ponies",   highlight: "y -> ies" },
    { word: "story -> stories", highlight: "y -> ies" },
  ],
  `SAY:
- Watch 'baby'. Before the y is a b, which is a consonant.
- So I change the y to i and add -es. Baby becomes babies.

DO:
- Underline the consonant before each y, then model the change.
- Say the plural clearly as you write.

TEACHER NOTES:
Check the letter immediately before y. That is the trigger for this rule.

WATCH FOR:
- Students writing 'babys' -- redirect to consonant + y check.

[CSR | Week 1 Session 3 | Teach]`
);

practiseSlide(
  "Week 1", "Session 3",
  "Practise: make these plural (consonant + y)",
  "For each word, check the letter before y.\n\nIf it is a consonant, change y to i and add -es.\n\nWrite the plural beside the word.",
  ["lady", "fly", "puppy", "berry", "copy", "cherry"],
  `SAY:
- Look before the y first. Consonant? Change y to i, add es.

DO:
- 90 seconds.
- Cold call two students to spell one plural aloud.

WATCH FOR:
- 'flys' -- reteach the consonant check.

[CSR | Week 1 Session 3 | Practise]`
);

checkSlide(
  "Week 1", "Session 3",
  "Check: plural of 'party'?",
  "Write the plural of 'party'. Show me the rule step in your OG Book.",
  "parties",
  `SAY:
- Check the letter before y. Apply the rule. Write the plural.

DO:
- Show me.
- Scan for 'parties'.

[CSR | Week 1 Session 3 | Check]`,
  `SAY:
- t is a consonant, so y changes to i and we add -es. Parties.

[CSR | Week 1 Session 3 | Reveal]`
);

// ── Week 1 Session 4: vowel + y -> just add s ───────────────────────────────
teachSlide(
  "Week 1", "Session 4",
  "Plural -s after vowel + y: just add s",
  "If a vowel (a, e, i, o, u) comes before the y, leave the y alone and just add -s.",
  [
    { word: "boy -> boys",   highlight: "+ s" },
    { word: "key -> keys",   highlight: "+ s" },
    { word: "day -> days",   highlight: "+ s" },
    { word: "toy -> toys",   highlight: "+ s" },
  ],
  `SAY:
- Look at 'boy'. Before the y is o. That is a vowel.
- So I leave the y and just add s. Boys.

DO:
- Underline the vowel before y in each example.
- Compare back to yesterday's consonant + y examples.

TEACHER NOTES:
Today deliberately contrasts with yesterday. The vowel versus consonant check is the whole rule.

WATCH FOR:
- 'boies' -- over-applying yesterday's rule. Check the letter before y.

[CSR | Week 1 Session 4 | Teach]`
);

practiseSlide(
  "Week 1", "Session 4",
  "Practise: vowel + y or consonant + y?",
  "Write each word. Underline the letter before y.\n\nIf vowel -> add s.\n\nIf consonant -> change y to i, add es.",
  ["boy", "pony", "key", "city", "day", "puppy"],
  `SAY:
- Mix of both rules. Underline first, then decide.

DO:
- 2 minutes.
- Pair-check before reveal.

WATCH FOR:
- Students who apply the same rule to every word.

[CSR | Week 1 Session 4 | Practise]`
);

checkSlide(
  "Week 1", "Session 4",
  "Check: plural of 'tray'?",
  "Write the plural of 'tray'. Which rule did you use?",
  "trays",
  `SAY:
- Check the letter before y, then apply the rule.

DO:
- Show me.
- Watch for 'traies' -- it is a classic slip.

[CSR | Week 1 Session 4 | Check]`,
  `SAY:
- a is a vowel, so we leave the y and just add s. Trays.

[CSR | Week 1 Session 4 | Reveal]`
);

// ── Week 1 Session 5: Mixed plural practice ─────────────────────────────────
teachSlide(
  "Week 1", "Session 5",
  "Plural review: pick the right rule",
  "Check the ending first. Ask: is it a hissing sound? a vowel + y? a consonant + y? plain? Choose the rule that matches.",
  [
    { word: "cat -> cats",      highlight: "+s" },
    { word: "box -> boxes",     highlight: "+es" },
    { word: "city -> cities",   highlight: "y->ies" },
    { word: "boy -> boys",      highlight: "+s" },
  ],
  `SAY:
- This is review week. All four rules live together today.
- The trick: look at the ending before you write.

DO:
- Model the decision aloud for each example.
- Sound out each plural to check.

TEACHER NOTES:
Mixed practice consolidates this week. Do not introduce new rules here.

WATCH FOR:
- Rule confusion between vowel + y and consonant + y.

[CSR | Week 1 Session 5 | Teach]`
);

practiseSlide(
  "Week 1", "Session 5",
  "Practise: mixed plurals",
  "Write the plural for each word.\n\nFor each one, write the rule code beside it:\nS = +s\nES = +es\nIES = y -> ies",
  ["lunch", "baby", "day", "dog", "fox", "story"],
  `SAY:
- Code your rule next to the plural. That is the thinking we are after.

DO:
- 2 minutes.
- Invite two students to explain their rule choice.

WATCH FOR:
- Correct spellings but wrong rule code -- probe for the reasoning.

[CSR | Week 1 Session 5 | Practise]`
);

checkSlide(
  "Week 1", "Session 5",
  "Check: plural of 'watch'?",
  "Write the plural of 'watch' and the rule code you used.",
  "watches  (ES, ends in -ch)",
  `SAY:
- Read the word. Find the ending. Pick the rule.

DO:
- Show me.
- Scan for correct plural AND correct rule code.

[CSR | Week 1 Session 5 | Check]`,
  `SAY:
- 'watch' ends in -ch, a hissing sound. So we add -es. Watches.

[CSR | Week 1 Session 5 | Reveal]`
);

// ===========================================================================
//  WEEK 3 — PAST TENSE -ED RULES
// ===========================================================================

weekDivider(
  pres,
  3,
  "Past Tense -ed Rules",
  "Changing verbs to past tense. Some just add -ed. Some drop an e. Some change y to i. Some double the final consonant.",
  `[Week 3 | -ed rules]`
);

// ── Week 3 Session 1: Basic -ed rule ────────────────────────────────────────
teachSlide(
  "Week 3", "Session 1",
  "Past tense -ed: just add -ed",
  "For most verbs, make them past tense by adding -ed.",
  [
    { word: "jump -> jumped",  highlight: "+ ed" },
    { word: "play -> played",  highlight: "+ ed" },
    { word: "walk -> walked",  highlight: "+ ed" },
    { word: "rain -> rained",  highlight: "+ ed" },
  ],
  `SAY:
- When something already happened, we often show it by adding -ed.
- Most verbs: just add -ed.

DO:
- Read each verb twice, present then past.
- Point at the -ed ending.

TEACHER NOTES:
Note: -ed can sound like /t/ (walked), /d/ (played), or /ed/ (landed). We spell it -ed no matter what.

WATCH FOR:
- Students writing 'jumpt' (by sound). Correct to -ed.

[CSR | Week 3 Session 1 | Teach]`
);

practiseSlide(
  "Week 3", "Session 1",
  "Practise: add -ed to each verb",
  "Write each verb, then its past tense.\n\nRead aloud.\n\nCheck: does your ending spell -ed?",
  ["jump", "play", "kick", "help", "paint", "look"],
  `SAY:
- Write the verb. Add -ed. Read it out loud.

DO:
- 90 seconds.
- Cold call two past-tense words.

WATCH FOR:
- 'helpt' or 'kickt' -- hear the /t/, write the -ed.

[CSR | Week 3 Session 1 | Practise]`
);

checkSlide(
  "Week 3", "Session 1",
  "Check: past tense of 'walk'?",
  "Write the past tense of 'walk'. Remember: we spell it -ed.",
  "walked",
  `SAY:
- Don't be tricked by the /t/ sound. We always spell it -ed.

DO:
- Show me.
- Scan for 'walked'.

[CSR | Week 3 Session 1 | Check]`,
  `SAY:
- Walked. The -ed sounds like /t/, but we write -ed.

[CSR | Week 3 Session 1 | Reveal]`
);

// ── Week 3 Session 2: verbs ending in e -> just add d ───────────────────────
teachSlide(
  "Week 3", "Session 2",
  "Past tense when the verb already ends in e: add d",
  "If the verb already ends in e, don't write another e. Just add d.",
  [
    { word: "bake -> baked",   highlight: "+ d" },
    { word: "smile -> smiled", highlight: "+ d" },
    { word: "live -> lived",   highlight: "+ d" },
    { word: "hope -> hoped",   highlight: "+ d" },
  ],
  `SAY:
- 'Bake' already ends in e. I don't need to write another e.
- Just add d. Baked.

DO:
- Show each verb, cross out the idea of 'ee'.
- Say the pair aloud.

WATCH FOR:
- 'bakeed' -- double e slip. Check for the existing e.

[CSR | Week 3 Session 2 | Teach]`
);

practiseSlide(
  "Week 3", "Session 2",
  "Practise: verbs that already end in e",
  "Write each verb, then its past tense.\n\nNo double e! Just add d.",
  ["bake", "smile", "live", "hope", "move", "dance"],
  `SAY:
- Silent e is already there. Just add d.

DO:
- 90 seconds.
- Pair check.

WATCH FOR:
- 'danceed' -- prompt: does the e already exist?

[CSR | Week 3 Session 2 | Practise]`
);

checkSlide(
  "Week 3", "Session 2",
  "Check: past tense of 'smile'?",
  "Write the past tense of 'smile' in your OG Book.",
  "smiled",
  `SAY:
- One e only. Add d.

[CSR | Week 3 Session 2 | Check]`,
  `SAY:
- 'Smile' already ends in e. Smiled. One e, then d.

[CSR | Week 3 Session 2 | Reveal]`
);

// ── Week 3 Session 3: consonant + y -> change y to i, add ed ────────────────
teachSlide(
  "Week 3", "Session 3",
  "Past tense with consonant + y: change y to i, add -ed",
  "If the verb ends in a consonant before y, change the y to i and add -ed.",
  [
    { word: "cry -> cried",     highlight: "y -> ied" },
    { word: "try -> tried",     highlight: "y -> ied" },
    { word: "carry -> carried", highlight: "y -> ied" },
    { word: "reply -> replied", highlight: "y -> ied" },
  ],
  `SAY:
- Look before the y. 'Cry' -- the r is a consonant.
- So I change y to i and add -ed. Cried.

DO:
- Underline the consonant before each y.
- Compare back to Week 1 session 3 -- same rule, different suffix.

WATCH FOR:
- 'cryed' -- check the letter before y.

[CSR | Week 3 Session 3 | Teach]`
);

practiseSlide(
  "Week 3", "Session 3",
  "Practise: consonant + y verbs",
  "For each verb, check the letter before y.\n\nIf it's a consonant, change y to i and add -ed.\n\nWrite the past tense.",
  ["cry", "try", "fly", "hurry", "copy", "marry"],
  `SAY:
- Careful: some of these are tricky to say out loud. Focus on the spelling.

DO:
- 2 minutes.
- Pair check.

WATCH FOR:
- 'fly' -> past tense 'flew' is irregular, but we'll practise the rule with 'flied' as a demonstration of the pattern. Note to students: flew is the standard past tense.

[CSR | Week 3 Session 3 | Practise]`
);

checkSlide(
  "Week 3", "Session 3",
  "Check: past tense of 'reply'?",
  "Write the past tense of 'reply' and show the rule in your OG Book.",
  "replied",
  `SAY:
- l is the letter before y. l is a consonant. So y becomes i and we add -ed.

[CSR | Week 3 Session 3 | Check]`,
  `SAY:
- Replied. Consonant + y triggers the change.

[CSR | Week 3 Session 3 | Reveal]`
);

// ── Week 3 Session 4: Doubling rule (CVC) ───────────────────────────────────
teachSlide(
  "Week 3", "Session 4",
  "Past tense doubling rule: CVC -> double final consonant + ed",
  "Short one-syllable verbs ending in Consonant-Vowel-Consonant (CVC)? Double the final consonant, then add -ed.",
  [
    { word: "stop -> stopped", highlight: "double p" },
    { word: "hop -> hopped",   highlight: "double p" },
    { word: "plan -> planned", highlight: "double n" },
    { word: "chat -> chatted", highlight: "double t" },
  ],
  `SAY:
- 'Stop'. Last three letters: s-t-o-p. s-t is a blend, but the core pattern is o (vowel) + p (consonant), with a consonant before the vowel.
- CVC pattern. I double the p, then add -ed. Stopped.

DO:
- Mark the C-V-C pattern on each word.
- Compare 'hoped' (silent e) with 'hopped' (CVC + doubled) so students see the difference.

TEACHER NOTES:
The doubling rule protects short vowel sounds. 'Hoped' vs 'hopped' is the classic contrast.

WATCH FOR:
- 'stoped' without doubling -- it changes the sound to a long o.

[CSR | Week 3 Session 4 | Teach]`
);

practiseSlide(
  "Week 3", "Session 4",
  "Practise: doubling or not?",
  "Sort: does the verb fit the CVC doubling rule?\n\nIf YES -> double + ed.\n\nIf NO -> just add -ed.",
  ["stop", "jump", "plan", "paint", "tap", "rest"],
  `SAY:
- Two columns: 'double' and 'just +ed'. Then write each past tense.

DO:
- 2 minutes.
- Invite students to explain one decision.

WATCH FOR:
- 'jumped' being doubled ('jummped'). 'jump' is CCVCC, not CVC.

[CSR | Week 3 Session 4 | Practise]`
);

checkSlide(
  "Week 3", "Session 4",
  "Check: past tense of 'hug'?",
  "Write the past tense of 'hug'. Does it need doubling?",
  "hugged  (CVC -> double)",
  `SAY:
- Check the pattern first, then write.

[CSR | Week 3 Session 4 | Check]`,
  `SAY:
- h-u-g is CVC. Double the g, add -ed. Hugged.

[CSR | Week 3 Session 4 | Reveal]`
);

// ── Week 3 Session 5: Mixed -ed review ──────────────────────────────────────
teachSlide(
  "Week 3", "Session 5",
  "-ed rules review: pick the right move",
  "Steps: 1) Does it end in e? Add d. 2) Consonant + y? y -> i, add ed. 3) CVC? Double, add ed. 4) Otherwise add ed.",
  [
    { word: "smile -> smiled",   highlight: "+d" },
    { word: "cry -> cried",      highlight: "y->ied" },
    { word: "hop -> hopped",     highlight: "double" },
    { word: "walk -> walked",    highlight: "+ed" },
  ],
  `SAY:
- All four rules. Same suffix -ed, four different spellings.

DO:
- Talk through the decision sequence aloud.

TEACHER NOTES:
Keep the four-step decision short. Do not add new information today.

WATCH FOR:
- Rule confusion under time pressure.

[CSR | Week 3 Session 5 | Teach]`
);

practiseSlide(
  "Week 3", "Session 5",
  "Practise: mixed -ed",
  "Write the past tense of each verb.\n\nCode the rule beside it:\nE = already ends in e\nY = y -> ied\nD = CVC double\nP = plain +ed",
  ["bake", "try", "stop", "paint", "hope", "hug"],
  `SAY:
- Code every answer. That is the point of this task.

DO:
- 2 minutes.
- Invite one student to explain their rule choice.

WATCH FOR:
- Spellings that look right but use the wrong rule code.

[CSR | Week 3 Session 5 | Practise]`
);

checkSlide(
  "Week 3", "Session 5",
  "Check: past tense of 'plan'?",
  "Write the past tense of 'plan' and code the rule you used.",
  "planned  (D, CVC -> double)",
  `SAY:
- Apply the four-step decision. Then write.

[CSR | Week 3 Session 5 | Check]`,
  `SAY:
- p-l-a-n is CVC. Double the n, add -ed. Planned.

[CSR | Week 3 Session 5 | Reveal]`
);

// ===========================================================================
//  WEEK 5 — SOFT c / SOFT g RULE
// ===========================================================================

weekDivider(
  pres,
  5,
  "Soft c and Soft g",
  "When c and g come before e, i, or y they often soften. c sounds like /s/, and g often sounds like /j/.",
  `[Week 5 | Soft c and g]`
);

// ── Week 5 Session 1: Soft c ────────────────────────────────────────────────
teachSlide(
  "Week 5", "Session 1",
  "Soft c: c before e, i, or y says /s/",
  "When c is followed by e, i, or y, it softens to the /s/ sound.",
  [
    { word: "cent",   highlight: "c before e = /s/" },
    { word: "city",   highlight: "c before i = /s/" },
    { word: "cycle",  highlight: "c before y = /s/" },
    { word: "dance",  highlight: "c before e = /s/" },
  ],
  `SAY:
- Watch this c. When e, i, or y comes straight after it, the c goes soft.
- It sounds like /s/.

DO:
- Say each word, stretching the /s/ sound.
- Underline the c and the letter after it.

TEACHER NOTES:
Pattern cue: 'E, I, Y' -- the three vowels that soften c.

WATCH FOR:
- Students saying /k/ in 'city' -- remodel the /s/.

[CSR | Week 5 Session 1 | Teach]`
);

practiseSlide(
  "Week 5", "Session 1",
  "Practise: circle the soft c words",
  "Read each word. Underline the letter right after c.\n\nIf it is e, i, or y -> circle the word: soft c.",
  ["city", "cat", "race", "cup", "pencil", "cycle"],
  `SAY:
- Underline first, then decide.

DO:
- 90 seconds.
- Invite students to read one soft c word aloud.

WATCH FOR:
- 'cat' being circled -- the 'a' is not a softener.

[CSR | Week 5 Session 1 | Practise]`
);

checkSlide(
  "Week 5", "Session 1",
  "Check: does 'circle' have a soft c?",
  "Read 'circle' out loud. Write YES or NO in your OG Book, and say why.",
  "YES -- c is before i, so it is soft /s/",
  `SAY:
- Look at the letter right after the first c. Apply the rule.

[CSR | Week 5 Session 1 | Check]`,
  `SAY:
- c is followed by i, so it softens. /s/. Circle.

[CSR | Week 5 Session 1 | Reveal]`
);

// ── Week 5 Session 2: Hard c ────────────────────────────────────────────────
teachSlide(
  "Week 5", "Session 2",
  "Hard c: c before a, o, u, or a consonant says /k/",
  "When c is followed by a, o, u, or a consonant, it stays hard. It sounds like /k/.",
  [
    { word: "cat",     highlight: "c + a = /k/" },
    { word: "cot",     highlight: "c + o = /k/" },
    { word: "cup",     highlight: "c + u = /k/" },
    { word: "clock",   highlight: "c + l = /k/" },
  ],
  `SAY:
- Now c in front of a, o, u, or a consonant. It stays hard.
- /k/.

DO:
- Say each word. Exaggerate the /k/.
- Connect back to yesterday: what softens c? (e, i, y). What keeps it hard? (a, o, u, consonant).

WATCH FOR:
- Students who over-apply /s/ to words like 'cot'.

[CSR | Week 5 Session 2 | Teach]`
);

practiseSlide(
  "Week 5", "Session 2",
  "Practise: soft or hard c?",
  "For each word, underline the letter after c.\n\nWrite S (soft /s/) or H (hard /k/) beside the word.",
  ["cat", "city", "cup", "cycle", "cost", "pencil"],
  `SAY:
- Check the letter after c. Then decide.

DO:
- 2 minutes.
- Pair check.

[CSR | Week 5 Session 2 | Practise]`
);

checkSlide(
  "Week 5", "Session 2",
  "Check: is the c in 'cold' soft or hard?",
  "Write S (soft) or H (hard) in your OG Book. Say why.",
  "H (hard) -- c is before o",
  `SAY:
- Identify the letter after c. Apply the rule.

[CSR | Week 5 Session 2 | Check]`,
  `SAY:
- c is before o. o does not soften. Hard c. /k/.

[CSR | Week 5 Session 2 | Reveal]`
);

// ── Week 5 Session 3: Soft g ────────────────────────────────────────────────
teachSlide(
  "Week 5", "Session 3",
  "Soft g: g before e, i, or y often says /j/",
  "When g is followed by e, i, or y, it often softens to /j/. (Some exceptions exist: 'get', 'girl'.)",
  [
    { word: "gentle", highlight: "g + e = /j/" },
    { word: "giant",  highlight: "g + i = /j/" },
    { word: "gym",    highlight: "g + y = /j/" },
    { word: "page",   highlight: "g + e = /j/" },
  ],
  `SAY:
- g is like c. When e, i, or y comes next, it usually softens.
- Soft g says /j/.

DO:
- Say each word. Exaggerate /j/.
- Flag that 'get' and 'girl' are common exceptions -- they stay hard.

TEACHER NOTES:
Soft g has more exceptions than soft c. Teach 'often', not 'always'.

WATCH FOR:
- 'get' and 'give' confusion -- name them as exceptions, do not over-explain.

[CSR | Week 5 Session 3 | Teach]`
);

practiseSlide(
  "Week 5", "Session 3",
  "Practise: find the soft g",
  "For each word, underline the letter after g.\n\nIf it is e, i, or y, it is probably soft g.\n\nRead aloud to confirm.",
  ["gentle", "goat", "giant", "gold", "page", "gym"],
  `SAY:
- Look after the g. Then say the word. Trust your ear too.

DO:
- 90 seconds.
- Quick share.

WATCH FOR:
- 'goat' circled -- g is before o, so it stays hard.

[CSR | Week 5 Session 3 | Practise]`
);

checkSlide(
  "Week 5", "Session 3",
  "Check: is the g in 'giraffe' soft or hard?",
  "Write S (soft /j/) or H (hard /g/) in your OG Book.",
  "S (soft) -- g is before i",
  `SAY:
- Check the letter after g. Say the word to confirm.

[CSR | Week 5 Session 3 | Check]`,
  `SAY:
- g is before i, and the word says /j/. Soft g.

[CSR | Week 5 Session 3 | Reveal]`
);

// ── Week 5 Session 4: Hard g ────────────────────────────────────────────────
teachSlide(
  "Week 5", "Session 4",
  "Hard g: g before a, o, u, or a consonant says /g/",
  "When g is followed by a, o, u, or a consonant, it stays hard. It sounds like /g/.",
  [
    { word: "gate",   highlight: "g + a = /g/" },
    { word: "goat",   highlight: "g + o = /g/" },
    { word: "gum",    highlight: "g + u = /g/" },
    { word: "glow",   highlight: "g + l = /g/" },
  ],
  `SAY:
- g plus a, o, u, or a consonant stays hard. /g/.

DO:
- Say each word clearly.
- Contrast with yesterday's soft g words.

WATCH FOR:
- Students softening 'goat' or 'gum' -- remodel /g/.

[CSR | Week 5 Session 4 | Teach]`
);

practiseSlide(
  "Week 5", "Session 4",
  "Practise: soft or hard g?",
  "For each word, underline the letter after g.\n\nWrite S (soft /j/) or H (hard /g/) beside the word.",
  ["gum", "gentle", "gold", "giant", "glass", "gym"],
  `SAY:
- Check the letter after g. Sort.

DO:
- 2 minutes.
- Share one pair: one soft, one hard.

[CSR | Week 5 Session 4 | Practise]`
);

checkSlide(
  "Week 5", "Session 4",
  "Check: is the g in 'goal' soft or hard?",
  "Write S or H in your OG Book and explain.",
  "H (hard) -- g is before o",
  `SAY:
- Check the letter after g. Apply the rule.

[CSR | Week 5 Session 4 | Check]`,
  `SAY:
- o does not soften g. Hard g.

[CSR | Week 5 Session 4 | Reveal]`
);

// ── Week 5 Session 5: Mixed c/g review ──────────────────────────────────────
teachSlide(
  "Week 5", "Session 5",
  "c and g review: check the next letter",
  "For c and g, always check the letter straight after. e, i, y -> often soft. a, o, u, consonant -> hard.",
  [
    { word: "cent",   highlight: "soft c" },
    { word: "cold",   highlight: "hard c" },
    { word: "gem",    highlight: "soft g" },
    { word: "goal",   highlight: "hard g" },
  ],
  `SAY:
- Same rule for both letters.

DO:
- Talk through one example of each (soft c, hard c, soft g, hard g).

WATCH FOR:
- Students who guess rather than check the next letter.

[CSR | Week 5 Session 5 | Teach]`
);

practiseSlide(
  "Week 5", "Session 5",
  "Practise: mixed soft and hard",
  "For each word, write one of: SC, HC, SG, HG.\n\nSC = soft c\nHC = hard c\nSG = soft g\nHG = hard g",
  ["cent", "cat", "giant", "goat", "cycle", "page"],
  `SAY:
- Code every answer.

DO:
- 2 minutes.
- Cold call two students to explain one choice.

[CSR | Week 5 Session 5 | Practise]`
);

checkSlide(
  "Week 5", "Session 5",
  "Check: code 'gym' and 'cup'.",
  "Write the code beside each word: SC, HC, SG, or HG.",
  "gym = SG (soft g)     cup = HC (hard c)",
  `SAY:
- Check the letter after c or g. Code it.

[CSR | Week 5 Session 5 | Check]`,
  `SAY:
- gym: g before y -> soft g, SG. cup: c before u -> hard c, HC.

[CSR | Week 5 Session 5 | Reveal]`
);

// ===========================================================================
//  WEEK 7 — VOWEL vs CONSONANT SUFFIXES
// ===========================================================================

weekDivider(
  pres,
  7,
  "Vowel Suffixes vs Consonant Suffixes",
  "A suffix that starts with a vowel can trigger spelling changes. A suffix that starts with a consonant usually does not.",
  `[Week 7 | Suffixes]`
);

// ── Week 7 Session 1: What are vowel and consonant suffixes? ────────────────
teachSlide(
  "Week 7", "Session 1",
  "Vowel suffix or consonant suffix? Check the first letter",
  "A suffix is a word ending we add. If it starts with a vowel, it's a vowel suffix. If it starts with a consonant, it's a consonant suffix.",
  [
    { word: "-ing (vowel)",   highlight: "starts with i" },
    { word: "-ed (vowel)",    highlight: "starts with e" },
    { word: "-ly (consonant)", highlight: "starts with l" },
    { word: "-ful (consonant)", highlight: "starts with f" },
  ],
  `SAY:
- Every suffix starts with a letter. Check that first letter.
- Vowel first = vowel suffix. Consonant first = consonant suffix.

DO:
- Write some suffixes on the board: -ing, -ed, -able, -ly, -ful, -ness.
- Sort them aloud with the class.

TEACHER NOTES:
The distinction matters because vowel suffixes trigger the doubling and drop-the-e rules.

WATCH FOR:
- Students just memorising a list -- keep the check on the first letter.

[CSR | Week 7 Session 1 | Teach]`
);

practiseSlide(
  "Week 7", "Session 1",
  "Practise: sort the suffixes",
  "Sort each suffix into two columns: vowel suffix or consonant suffix.\n\nCheck the first letter.",
  ["-ing", "-ly", "-ed", "-ful", "-able", "-ness"],
  `SAY:
- First letter only.

DO:
- 90 seconds.
- Pair check.

[CSR | Week 7 Session 1 | Practise]`
);

checkSlide(
  "Week 7", "Session 1",
  "Check: is -est a vowel or consonant suffix?",
  "Write V or C in your OG Book and explain.",
  "V -- -est starts with e, which is a vowel",
  `SAY:
- Check the first letter.

[CSR | Week 7 Session 1 | Check]`,
  `SAY:
- -est starts with e. Vowel suffix.

[CSR | Week 7 Session 1 | Reveal]`
);

// ── Week 7 Session 2: Doubling rule with vowel suffixes ─────────────────────
teachSlide(
  "Week 7", "Session 2",
  "CVC + vowel suffix = double the last consonant",
  "If the base word is CVC and we add a vowel suffix, double the last consonant before adding the suffix.",
  [
    { word: "run + ing -> running",  highlight: "double n" },
    { word: "swim + ing -> swimming", highlight: "double m" },
    { word: "hot + est -> hottest",   highlight: "double t" },
    { word: "stop + ed -> stopped",   highlight: "double p" },
  ],
  `SAY:
- 'Run' is CVC. 'Ing' starts with a vowel.
- So I double the n before I add -ing. Running.

DO:
- Underline the CVC.
- Show a non-example: 'walk' (CVCC) does not double.

WATCH FOR:
- 'runing' with one n -- it changes the vowel sound to a long u.

[CSR | Week 7 Session 2 | Teach]`
);

practiseSlide(
  "Week 7", "Session 2",
  "Practise: CVC + vowel suffix",
  "Add the suffix to each word. Apply the doubling rule if it fits.",
  ["run + ing", "jump + ing", "hot + er", "rest + ed", "hop + ed", "sit + ing"],
  `SAY:
- CVC + vowel suffix = double. Check both parts first.

DO:
- 2 minutes.
- Pair check.

WATCH FOR:
- 'jumping' being doubled -- jump is CCVCC, not CVC.

[CSR | Week 7 Session 2 | Practise]`
);

checkSlide(
  "Week 7", "Session 2",
  "Check: shop + ing = ?",
  "Apply the rule. Write your answer.",
  "shopping  (shop is CVC + vowel suffix)",
  `SAY:
- CVC? Vowel suffix? Both yes -> double.

[CSR | Week 7 Session 2 | Check]`,
  `SAY:
- Shop is CVC. -ing is a vowel suffix. Double the p. Shopping.

[CSR | Week 7 Session 2 | Reveal]`
);

// ── Week 7 Session 3: Drop the e rule ───────────────────────────────────────
teachSlide(
  "Week 7", "Session 3",
  "Silent e + vowel suffix = drop the e",
  "If the base word ends in silent e and the suffix starts with a vowel, drop the e first.",
  [
    { word: "bake + ing -> baking",  highlight: "drop e" },
    { word: "hope + ed -> hoped",     highlight: "drop e" },
    { word: "make + er -> maker",     highlight: "drop e" },
    { word: "race + ing -> racing",   highlight: "drop e" },
  ],
  `SAY:
- 'Bake' ends in silent e. '-ing' starts with a vowel.
- So I drop the e before I add -ing. Baking.

DO:
- Cross out the silent e in each example, then write the suffix.
- Contrast: 'bake + s' does NOT drop the e.

WATCH FOR:
- 'bakeing' with the e still there -- check the rule.

[CSR | Week 7 Session 3 | Teach]`
);

practiseSlide(
  "Week 7", "Session 3",
  "Practise: drop the e or not?",
  "Add the suffix. Drop the silent e if the suffix starts with a vowel.",
  ["make + ing", "hope + ful", "ride + ing", "use + ed", "move + ment", "live + ing"],
  `SAY:
- Check the base: silent e?
- Check the suffix: vowel first?
- Both yes -> drop the e.

DO:
- 2 minutes.

WATCH FOR:
- 'hopful' -- -ful is a consonant suffix, so keep the e: hopeful.

[CSR | Week 7 Session 3 | Practise]`
);

checkSlide(
  "Week 7", "Session 3",
  "Check: smile + ing = ?",
  "Apply the rule. Write your answer.",
  "smiling  (drop the e before -ing)",
  `SAY:
- Silent e + vowel suffix -> drop the e.

[CSR | Week 7 Session 3 | Check]`,
  `SAY:
- Smile ends in silent e. -ing starts with a vowel. Drop the e. Smiling.

[CSR | Week 7 Session 3 | Reveal]`
);

// ── Week 7 Session 4: Consonant suffixes don't trigger changes ──────────────
teachSlide(
  "Week 7", "Session 4",
  "Consonant suffix = usually just add it",
  "When a suffix starts with a consonant, usually just add it. Don't drop e. Don't double.",
  [
    { word: "hope + ful -> hopeful",    highlight: "keep e" },
    { word: "kind + ly -> kindly",       highlight: "no change" },
    { word: "care + less -> careless",   highlight: "keep e" },
    { word: "sad + ness -> sadness",     highlight: "no double" },
  ],
  `SAY:
- Consonant suffix? Usually no changes to the base word.
- Just add the suffix.

DO:
- Contrast with Week 7 Sessions 2 and 3 (vowel suffix changes).

WATCH FOR:
- Students dropping the e for -ful or -ly -- reteach: consonant suffix, no change.

[CSR | Week 7 Session 4 | Teach]`
);

practiseSlide(
  "Week 7", "Session 4",
  "Practise: consonant suffixes",
  "Add the suffix. Usually no spelling change in the base word.",
  ["hope + ful", "kind + ly", "care + less", "sad + ness", "help + ful", "love + ly"],
  `SAY:
- Just add it. Keep the base word as it is.

DO:
- 90 seconds.
- Pair check.

WATCH FOR:
- 'sadness' doubled -- consonant suffixes don't trigger doubling.

[CSR | Week 7 Session 4 | Practise]`
);

checkSlide(
  "Week 7", "Session 4",
  "Check: hope + less = ?",
  "Add the consonant suffix. Write your answer.",
  "hopeless  (keep the e, just add -less)",
  `SAY:
- Consonant suffix. No change to the base.

[CSR | Week 7 Session 4 | Check]`,
  `SAY:
- -less starts with a consonant. Keep the e. Hopeless.

[CSR | Week 7 Session 4 | Reveal]`
);

// ── Week 7 Session 5: Mixed suffix review ───────────────────────────────────
teachSlide(
  "Week 7", "Session 5",
  "Suffix review: check suffix type first",
  "Step 1: Vowel or consonant suffix? Step 2: If vowel -> does the base trigger doubling or drop-the-e?",
  [
    { word: "run + ing -> running",  highlight: "CVC + V" },
    { word: "bake + ing -> baking",   highlight: "e + V" },
    { word: "hope + ful -> hopeful",  highlight: "C suffix" },
    { word: "kind + ly -> kindly",    highlight: "C suffix" },
  ],
  `SAY:
- Suffix type first. Then the base-word rule.

DO:
- Walk through all four decisions aloud.

[CSR | Week 7 Session 5 | Teach]`
);

practiseSlide(
  "Week 7", "Session 5",
  "Practise: mixed suffixes",
  "Add each suffix correctly.",
  ["stop + ing", "bake + er", "help + ful", "move + ment", "sit + ing", "care + less"],
  `SAY:
- Apply the right rule for each.

DO:
- 2 minutes.
- Pair check.

[CSR | Week 7 Session 5 | Practise]`
);

checkSlide(
  "Week 7", "Session 5",
  "Check: swim + er = ?",
  "Apply the rule. Write your answer.",
  "swimmer  (CVC + vowel suffix -> double m)",
  `SAY:
- Base CVC? Yes. Suffix starts with vowel? Yes. Double.

[CSR | Week 7 Session 5 | Check]`,
  `SAY:
- swim is CVC. -er is a vowel suffix. Double the m. Swimmer.

[CSR | Week 7 Session 5 | Reveal]`
);

// ===========================================================================
//  WEEK 9 — TIGER RULE (OPEN vs CLOSED SYLLABLES)
// ===========================================================================

weekDivider(
  pres,
  9,
  "The Tiger Rule",
  "The tiger rule helps us read two-syllable words like ti-ger, mu-sic, and ba-by. If a syllable ends in a vowel, the vowel is usually long.",
  `[Week 9 | Tiger rule]`
);

// ── Week 9 Session 1: Open syllable ─────────────────────────────────────────
teachSlide(
  "Week 9", "Session 1",
  "Open syllable: ends in a vowel, vowel is long",
  "An open syllable ends with a vowel. That vowel usually says its long sound.",
  [
    { word: "go",   highlight: "long o" },
    { word: "me",   highlight: "long e" },
    { word: "hi",   highlight: "long i" },
    { word: "we",   highlight: "long e" },
  ],
  `SAY:
- If a syllable ends with a vowel and nothing else closes it, the vowel is open.
- Open syllables have long vowels.

DO:
- Model the long sound aloud for each example.
- Mark the syllable ending with a line.

WATCH FOR:
- Students mixing short and long vowels. Keep modelling /oh/, /ee/, /eye/.

[CSR | Week 9 Session 1 | Teach]`
);

practiseSlide(
  "Week 9", "Session 1",
  "Practise: read and say the long vowel",
  "For each word, read it out loud.\n\nUnderline the vowel at the end.\n\nSay: long or short?",
  ["go", "me", "hi", "so", "be", "no"],
  `SAY:
- All open syllables. Long vowels only.

DO:
- Whole-class choral read.

[CSR | Week 9 Session 1 | Practise]`
);

checkSlide(
  "Week 9", "Session 1",
  "Check: is 'be' an open syllable?",
  "Yes or no? Write in your OG Book and explain.",
  "YES -- 'be' ends in a vowel, so it is open. Long e.",
  `SAY:
- Check the end of the syllable. Vowel alone = open.

[CSR | Week 9 Session 1 | Check]`,
  `SAY:
- 'Be' ends in e, a vowel. Open syllable. Long e.

[CSR | Week 9 Session 1 | Reveal]`
);

// ── Week 9 Session 2: Closed syllable ───────────────────────────────────────
teachSlide(
  "Week 9", "Session 2",
  "Closed syllable: ends in a consonant, vowel is short",
  "A closed syllable ends with a consonant after the vowel. The consonant 'closes' the syllable. The vowel says its short sound.",
  [
    { word: "cat",  highlight: "short a" },
    { word: "pen",  highlight: "short e" },
    { word: "pin",  highlight: "short i" },
    { word: "top",  highlight: "short o" },
  ],
  `SAY:
- A consonant at the end closes the syllable.
- Closed syllables have short vowels.

DO:
- Model the short vowel sound clearly.
- Contrast with yesterday's open syllables.

WATCH FOR:
- Students saying long a in 'cat' -- short a is tight and quick.

[CSR | Week 9 Session 2 | Teach]`
);

practiseSlide(
  "Week 9", "Session 2",
  "Practise: open or closed?",
  "For each word, check the last letter.\n\nVowel -> open -> long\n\nConsonant -> closed -> short",
  ["go", "cat", "me", "pen", "hi", "top"],
  `SAY:
- Look at the last letter. Sort.

DO:
- 90 seconds.
- Invite a student to read each word with the right vowel sound.

[CSR | Week 9 Session 2 | Practise]`
);

checkSlide(
  "Week 9", "Session 2",
  "Check: is 'pig' open or closed?",
  "Write O (open) or C (closed). Say the vowel sound.",
  "C (closed) -- ends in g, short i",
  `SAY:
- Last letter check. Then sound.

[CSR | Week 9 Session 2 | Check]`,
  `SAY:
- 'Pig' ends in g, a consonant. Closed. Short i.

[CSR | Week 9 Session 2 | Reveal]`
);

// ── Week 9 Session 3: Tiger rule (VC/V division) ────────────────────────────
teachSlide(
  "Week 9", "Session 3",
  "Tiger rule: split before the consonant (VC/V) - long first vowel",
  "In words like tiger, we split before the consonant. That gives an open first syllable, so the first vowel is long.",
  [
    { word: "ti-ger",  highlight: "long i" },
    { word: "mu-sic",  highlight: "long u" },
    { word: "ba-by",   highlight: "long a" },
    { word: "o-pen",   highlight: "long o" },
  ],
  `SAY:
- Look at 'tiger'. I could split tig-er or ti-ger.
- If I try ti-ger, the first syllable is open, so the i is long. That matches how we say the word.

DO:
- Mark each split with a line.
- Say each word twice, focusing on the long first vowel.

TEACHER NOTES:
The tiger rule = V/CV division. The first vowel opens, so it's long.

WATCH FOR:
- Students trying 'tig-er' -- remodel: long i sounds right.

[CSR | Week 9 Session 3 | Teach]`
);

practiseSlide(
  "Week 9", "Session 3",
  "Practise: split these tiger-rule words",
  "For each word, try splitting before the consonant.\n\nRead aloud: does the long first vowel sound right?",
  ["tiger", "music", "baby", "open", "paper", "robot"],
  `SAY:
- Split. Say it. Does it sound right?

DO:
- Whole-class read each.
- Mark the split on the board.

[CSR | Week 9 Session 3 | Practise]`
);

checkSlide(
  "Week 9", "Session 3",
  "Check: split 'robot' using the tiger rule.",
  "Write the split and say the first vowel.",
  "ro-bot  (long o)",
  `SAY:
- Split before the consonant. Say the long vowel.

[CSR | Week 9 Session 3 | Check]`,
  `SAY:
- ro-bot. Open first syllable. Long o.

[CSR | Week 9 Session 3 | Reveal]`
);

// ── Week 9 Session 4: Rabbit rule (VC/CV) ───────────────────────────────────
teachSlide(
  "Week 9", "Session 4",
  "Rabbit rule: split between two consonants (VC/CV) - short first vowel",
  "When two consonants sit in the middle (like rabbit), split between them. The first syllable closes, so the vowel is short.",
  [
    { word: "rab-bit", highlight: "short a" },
    { word: "lad-der", highlight: "short a" },
    { word: "pup-py",  highlight: "short u" },
    { word: "nap-kin", highlight: "short a" },
  ],
  `SAY:
- 'Rabbit' has two b's in the middle. I split between them.
- The first syllable is 'rab', which is closed. Short a.

DO:
- Mark the split between the two middle consonants.
- Contrast with yesterday: tiger = one consonant, rabbit = two.

WATCH FOR:
- Students saying long a in 'rabbit' -- remodel closed syllable.

[CSR | Week 9 Session 4 | Teach]`
);

practiseSlide(
  "Week 9", "Session 4",
  "Practise: split these rabbit-rule words",
  "For each word with two middle consonants, split between them.\n\nFirst vowel should sound short.",
  ["rabbit", "letter", "puppy", "napkin", "kitten", "dinner"],
  `SAY:
- Two middle consonants. Split between. Short first vowel.

DO:
- 2 minutes.
- Pair read.

[CSR | Week 9 Session 4 | Practise]`
);

checkSlide(
  "Week 9", "Session 4",
  "Check: split 'kitten' using the rabbit rule.",
  "Write the split and say the first vowel.",
  "kit-ten  (short i)",
  `SAY:
- Two middle consonants? Split between. Short first vowel.

[CSR | Week 9 Session 4 | Check]`,
  `SAY:
- kit-ten. Closed first syllable. Short i.

[CSR | Week 9 Session 4 | Reveal]`
);

// ── Week 9 Session 5: Mixed tiger/rabbit ────────────────────────────────────
teachSlide(
  "Week 9", "Session 5",
  "Tiger vs Rabbit: count the middle consonants",
  "One middle consonant? Try the tiger rule (split before it, long vowel). Two middle consonants? Use the rabbit rule (split between, short vowel).",
  [
    { word: "ti-ger (1 middle)", highlight: "tiger" },
    { word: "rab-bit (2 middle)", highlight: "rabbit" },
    { word: "pa-per (1 middle)",  highlight: "tiger" },
    { word: "pup-py (2 middle)",  highlight: "rabbit" },
  ],
  `SAY:
- Count the middle consonants. That picks the rule.

DO:
- Walk through each example.
- Emphasise the first-vowel sound.

[CSR | Week 9 Session 5 | Teach]`
);

practiseSlide(
  "Week 9", "Session 5",
  "Practise: tiger or rabbit?",
  "For each word, count the middle consonants.\n\n1 = tiger (split before, long)\n2 = rabbit (split between, short)\n\nWrite the split.",
  ["tiger", "rabbit", "music", "letter", "open", "puppy"],
  `SAY:
- Count first. Then split. Then say.

DO:
- 2 minutes.
- Pair check.

[CSR | Week 9 Session 5 | Practise]`
);

checkSlide(
  "Week 9", "Session 5",
  "Check: tiger or rabbit for 'napkin' and 'paper'?",
  "Write each split and the rule code (T or R).",
  "nap-kin = R     pa-per = T",
  `SAY:
- Middle consonants first. Then code.

[CSR | Week 9 Session 5 | Check]`,
  `SAY:
- napkin has two middle consonants (pk), so Rabbit. paper has one (p), so Tiger.

[CSR | Week 9 Session 5 | Reveal]`
);

// ────────────────────────────────────────────────────────────────────────────
//  Closing slide
// ────────────────────────────────────────────────────────────────────────────

closingSlide(
  pres,
  "Which spelling rule helped you most this term? Share one example with your partner.",
  [
    "CSR is a 5-minute routine -- fit it inside your bigger literacy block.",
    "Each session: Teach, Practise, Check.",
    "Return to these rules often -- retrieval builds fluency.",
  ],
  `SAY:
- Every session works the same way: teach the rule, practise it, check it.
- The rules we learned connect across weeks -- plural rules and -ed rules share the same consonant + y check.

DO:
- Revisit any session when the need comes up in reading or writing.

TEACHER NOTES:
No printed student resources are needed. Students write in their OG Books across all sessions.

[CSR | Closing]`
);

// ────────────────────────────────────────────────────────────────────────────
//  Write
// ────────────────────────────────────────────────────────────────────────────

(async () => {
  const outFile = path.join(OUT_DIR, "CSR Year 56 Spelling Routines.pptx");
  await pres.writeFile({ fileName: outFile });
  console.log(`PPTX written to ${outFile}`);
})();
