"use strict";

// Expanded Noun Groups - Year 5/6 Literacy (single session)
// Focus: expand a noun group by adding precise detail before AND after the noun.
// Mixed readiness. Teacher model sentences only (no novel required). Lean deck.
// Variant 3 (Ink & Paper) - standalone lesson, no unit cohesion constraint.

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme } = require("../themes/factory");
const T = createTheme("literacy", "grade56", 3);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  getContrastColor, withReveal, runSlideDiagnostics,
  titleSlide, liSlide, vocabSlide, cfuSlide, exitTicketSlide, closingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  PAGE, hex, lighten,
} = require("../themes/pdf_helpers");

// Colours for the printed builder (6-char hex to match the slide model).
const PDF_BEFORE = "284470"; // SECONDARY-ish (describing words)
const PDF_NOUN = "0E2258";   // PRIMARY (noun)
const PDF_AFTER = "9C4A1E";  // ACCENT-ish copper (detail after)

// A 3-column "before / NOUN / after" builder row. The noun is pre-printed in a
// tinted middle box; the before/after boxes are empty for writing, or carry a
// sample answer on the answer key. Mirrors the colour-coded slide strip so the
// resource matches what students saw. Returns the y below the row.
function builderRow(doc, noun, y, opts) {
  const o = opts || {};
  const x = PAGE.MARGIN;
  const totalW = PAGE.CONTENT_W;
  const gap = 8;
  const nounW = 92;
  const sideW = (totalW - nounW - gap * 2) / 2;
  const labelH = 14;
  const boxH = o.boxH || 48;
  const boxY = y + labelH;

  const cols = [
    { x, w: sideW, label: "describing words (before)", color: PDF_BEFORE, answer: o.before },
    { x: x + sideW + gap, w: nounW, label: "noun", color: PDF_NOUN, noun: true },
    { x: x + sideW + gap + nounW + gap, w: sideW, label: "extra detail (after)", color: PDF_AFTER, answer: o.after },
  ];

  cols.forEach((c) => {
    doc.fontSize(8).font("Sans-Bold").fillColor(hex(c.color));
    doc.text(c.label, c.x, y, { width: c.w, align: "center" });
  });

  cols.forEach((c) => {
    doc.save();
    if (c.noun) {
      doc.roundedRect(c.x, boxY, c.w, boxH, 4).fillAndStroke(lighten(c.color, 0.82), hex(c.color));
      doc.fontSize(15).font("Sans-Bold").fillColor(hex("2D3142"));
      doc.text(noun, c.x, boxY + boxH / 2 - 9, { width: c.w, align: "center" });
    } else {
      doc.roundedRect(c.x, boxY, c.w, boxH, 4).lineWidth(0.9).strokeColor(hex("9CA3AF")).stroke();
      if (c.answer) {
        doc.fontSize(12).font("Sans").fillColor(hex(c.color));
        doc.text(c.answer, c.x + 6, boxY + boxH / 2 - 8, { width: c.w - 12, align: "center" });
      }
    }
    doc.restore();
  });

  return boxY + boxH + 14;
}

const SESSION_NUMBER = 1;
const FOOTER = "Expanded Noun Groups | Year 5/6 Literacy";
const OUT_DIR = "output/Expanded_Noun_Groups_Lesson";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const BUILDER_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Noun Group Builder",
  "Build expanded noun groups: worked example, build-it table, upgrade-a-sentence and create-your-own."
);
const ANSWER_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Sample expanded noun groups and teacher look-fors for the Noun Group Builder."
);
const RESOURCE_ITEMS = [BUILDER_RESOURCE, ANSWER_RESOURCE];
const BUILDER_PDF_PATH = path.join(OUT_DIR, BUILDER_RESOURCE.fileName);
const ANSWER_PDF_PATH = path.join(OUT_DIR, ANSWER_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Custom visual anchor: the colour-coded "noun group" strip.
// A row of labelled chunks read left-to-right as one noun group:
//   [the] [describing words] [NOUN] [detail after]
// Box widths are proportional to text length. White text on coloured fills
// (palette guarantees AA), dark text on the light "the" chunk. Job labels sit
// directly under each box. Returns box geometry so reveals can fill blanks.
// ---------------------------------------------------------------------------

function stripLayout(x, y, w, chunks, gap, boxH) {
  const usable = w - gap * (chunks.length - 1);
  // Size by the visible word OR a `widthText` hint (used so We Do's empty
  // slots are pre-sized for the words the reveal will drop in). Each box gets
  // a minimum width so short nouns ("dog", "garden") never wrap.
  const txts = chunks.map((c) => String(c.widthText != null ? c.widthText : (c.words || "")));
  const minW = chunks.map((c, i) => (c.light ? 0.9 : Math.max(1.4, txts[i].length * 0.14 + 0.35)));
  const sumMin = minW.reduce((a, b) => a + b, 0);
  let widths;
  if (sumMin >= usable) {
    const scale = usable / sumMin;
    widths = minW.map((m) => m * scale);
  } else {
    const extra = usable - sumMin;
    const weights = txts.map((t) => Math.max(t.length, 3));
    const sumWt = weights.reduce((a, b) => a + b, 0);
    widths = minW.map((m, i) => m + (extra * weights[i]) / sumWt);
  }
  const boxes = [];
  let cx = x;
  widths.forEach((bw) => {
    boxes.push({ x: cx, y, w: bw, h: boxH });
    cx += bw + gap;
  });
  return boxes;
}

function drawChunkBox(slide, box, c, labelH) {
  const fill = c.light ? C.BG_LIGHT : (c.color || C.PRIMARY);
  slide.addShape("roundRect", {
    x: box.x, y: box.y, w: box.w, h: box.h, rectRadius: 0.08,
    fill: { color: fill },
    line: { color: c.light ? C.MUTED : fill, width: 1 },
  });
  if (c.words) {
    slide.addText(String(c.words), {
      x: box.x + 0.05, y: box.y + 0.06, w: box.w - 0.1, h: box.h - 0.12,
      fontSize: c.fontSize || 24, fontFace: FONT_H, color: getContrastColor(fill),
      bold: true, align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  }
  if (c.label) {
    slide.addText(String(c.label), {
      x: box.x, y: box.y + box.h + 0.06, w: box.w, h: labelH,
      fontSize: 12, fontFace: FONT_B, color: c.light ? C.MUTED : (c.color || C.CHARCOAL),
      bold: true, align: "center", valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });
  }
}

function nounGroupStrip(slide, x, y, w, chunks, opts) {
  const o = opts || {};
  const boxH = o.boxH || 1.05;
  const labelH = o.labelH || 0.4;
  const gap = o.gap || 0.12;
  const boxes = stripLayout(x, y, w, chunks, gap, boxH);
  chunks.forEach((c, i) => drawChunkBox(slide, boxes[i], c, labelH));
  return { boxes, bottom: y + boxH + labelH + 0.06 };
}

function fillChunkText(slide, box, words, fill) {
  slide.addText(String(words), {
    x: box.x + 0.05, y: box.y + 0.06, w: box.w - 0.1, h: box.h - 0.12,
    fontSize: 22, fontFace: FONT_H, color: getContrastColor(fill),
    bold: true, align: "center", valign: "middle", margin: 0,
    fit: "shrink", shrinkText: true,
  });
}

// ---------------------------------------------------------------------------
// Teacher notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we make our writing paint a clearer picture for the reader.
- We do that by growing a noun group - adding precise detail before and after the noun.
- Some of you may have used describing words before. If this feels new, that is okay.

DO:
- Display the title slide.
- Have whiteboards and markers ready.

TEACHER NOTES:
One tight focus today: expand a noun group with detail before and after the noun. Modelling and practice carry the load, not the slide text.

WATCH FOR:
- Students who already name adjectives - celebrate and connect to today.

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- One builder sheet today and an answer key for me.
- The ${BUILDER_RESOURCE.name} is where you build and upgrade noun groups.

DO:
- Print the Noun Group Builder, one per student.
- Keep the Answer Key for yourself.
- Whiteboards and markers out for every student.

TEACHER NOTES:
The builder uses the same before / noun / after model as the slides, so the worked example matches what students see.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_LAUNCH = `SAY:
- Read both sentences with me.
- Which one puts a clearer picture in your head? Show me A or B on your fingers.
- Tell your partner ONE word that made the difference.

DO:
- Choral read A, then B.
- Show Me Fingers: 1 for A, 2 for B (almost all will pick B).
- Cold call 2-3 students: which word helped you see it?
- Bridge: "Those extra words grew the noun group. Today we learn to grow our own."

TEACHER NOTES:
The launch activates prior knowledge of describing words and connects it to today's target: adding precise detail around a noun.

WATCH FOR:
- Students who point to the describing words or "with muddy paws" - celebrate; these are the parts we name next.
- Students who say B is "just longer" - redirect: "Longer only helps if it adds a clearer picture."

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to expand a noun group with precise detail before and after the noun.
- Three "I can" statements.

DO:
- Choral read the LI and the three success criteria.
- Point out: a noun group is the noun plus the words that describe it.

TEACHER NOTES:
SC1 (all, with support) - add describing words before a noun. SC2 (core, exit ticket) - build detail before AND after. SC3 (depth) - choose precise, exact words. Exit ticket targets SC2.

WATCH FOR:
- Students who think this means "make it longer" - reframe as "make it clearer".

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- The key term: expanded noun group.
- A noun group is a noun plus the words that describe it.
- We expand it by adding describing words BEFORE the noun and extra detail AFTER it.

DO:
- Choral say "expanded noun group".
- Read the example. Ask: which word is the noun? (tower)
- Ask: which words come before? which detail comes after?

TEACHER NOTES:
Keep the definition concrete. The slide example is different from the I Do example on purpose, so the modelling still does the teaching.

WATCH FOR:
- Students who confuse the noun with a describing word - point back to "the thing itself".

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_IDO = `SAY:
- Watch me grow a plain noun group into a clear picture.
- I start with "the dog". That tells you almost nothing.
- Before the noun I add describing words: enormous, shaggy. Now you can picture it.
- After the noun I add detail: with muddy paws. Now you can really see it.
- Notice each part has a job. The noun stays the anchor in the middle.

DO:
- Point to "the dog" first.
- Reveal each coloured chunk as you say it: describing words, then noun, then detail after.
- Read the whole sentence aloud at the bottom.
- Think aloud: "I chose 'enormous', not 'big' - it is more precise."

TEACHER NOTES:
This is the core model. The colour-coded strip is the visual anchor: before / noun / after. Keep the strip up while students try the CFU.

WATCH FOR:
- Students who add vague words ("nice", "good") - flag for the precise-words point coming up.

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. Three noun groups about a car.
- Which one gives the clearest, most precise picture?
- Show me A, B or C on your fingers. Then we check.

DO:
- Read A, B and C aloud.
- Show Me Fingers: 1 (A), 2 (B), 3 (C).
- Scan: most should choose C.
- Cold call: "Why not B?" (because "nice" is vague).

CFU CHECKPOINT:
Technique: Show Me Fingers (1=A, 2=B, 3=C)
Script:
- "Which noun group paints the clearest, most precise picture?"
- Scan for: most students choose C (>=80%).
PROCEED:
- If >=80% choose C, move to We Do.
PIVOT:
- Most likely misconception: more words always means better, so students pick the longest or pick "nice".
- Reteach with the optional slide that follows: a two-question frame (What is it like? Which one exactly?) to build detail with PRECISE words.
- Fresh re-check: on whiteboards, swap "nice" in "the nice car" for a precise word.
- Use the optional re-teach slide that follows.

TEACHER NOTES:
The hinge tests precision, not length. B is longer than A but still vague.

WATCH FOR:
- Students who pick C and can name the before and after parts - ready for We Do.

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The clearest is C.
- "rusty, red" describes it precisely before the noun, and "with a cracked windscreen" adds exact detail after.
- B only adds "nice" - that is vague. It does not help us picture the car.

DO:
- Show the answer banner.
- Point to the precise words in C.

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_RETEACH = `SAY:
- Another way to build the detail. Ask two questions.
- Question 1, BEFORE the noun: what is it like? (wild, overgrown)
- Question 2, AFTER the noun: which one exactly, or where? (behind the old fence)
- Two questions, two pieces of precise detail.

DO:
- Display the re-teach slide.
- Build the garden example LIVE using the two questions.
- Re-check: on whiteboards, students add ONE precise "before" word to "the garden".

TEACHER NOTES:
Optional. Use only if the CFU was below 80%. Different approach from the I Do: here a two-question frame drives the detail, instead of just naming the parts. Skip this slide if students were ready.

WATCH FOR:
- Students who answer the questions with precise words - ready to proceed.
- Students who answer with vague words - model one precise swap, then move on.

[Literacy: Optional Re-teach | VTLM 2.0: Reteach with Different Approach]`;

const NOTES_WEDO = `SAY:
- Together now. The noun is "castle". It is locked in the middle.
- On your whiteboards: what describing words go BEFORE? What detail goes AFTER?
- Aim for precise words that help us picture this exact castle.

DO:
- Read the noun and the two empty slots.
- Take suggestions on whiteboards, then cold call.
- Click to reveal one strong version and compare it with student ideas.

TEACHER NOTES:
We Do uses different content from the I Do (castle, not dog) and from the You Do. Students attempt before the reveal. Any sensible precise version is valid - the reveal is one example, not the only answer.

WATCH FOR:
- Students who add detail in the wrong slot - point back to before vs after.
- Students with precise, vivid words - celebrate and capture for the class.

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Here is one strong version: the ancient, crumbling castle on the misty hill.
- "ancient, crumbling" before the noun. "on the misty hill" after it.
- Did your version paint a clear picture too? Many good answers are possible.

DO:
- Reveal the filled strip.
- Compare with two student versions.

[Literacy: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Your turn on the Noun Group Builder.
- First: look at the worked example - the box.
- Next: build expanded noun groups for the nouns in the table.
- Then: upgrade the plain sentence by growing its noun group.

DO:
- Distribute the Noun Group Builder.
- Set 12-15 minutes.
- Circulate. Ask: "Read me your noun group. What is your most precise word?"

CFU CHECKPOINT:
Technique: Circulate and check
Script:
- "Point to your noun. Now read the detail before and after it."
- Scan for: detail in both positions with at least one precise word.
PROCEED:
- Most students building before AND after - keep going to the exit ticket.
PIVOT:
- If several are stuck, pull a small group and use the two-question frame from the re-teach.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the sentence frame on the builder ("the ___, ___ NOUN ___"). Add one before word and one after detail.
- Extra Notes: The frame and worked example keep SC1 and SC2 reachable.
EXTENDING PROMPT:
- Task: Swap your "after" phrase for a who/which/that clause (e.g. "the cat that prowled the alley"). Then choose your most precise version.
- Extra Notes: Connects to relative clauses without overloading the core task.

TEACHER NOTES:
You Do uses different nouns from the We Do and I Do. The builder's worked example matches the slide model exactly.

WATCH FOR:
- Students who only add words before - prompt: "Now add a detail AFTER the noun."
- Students padding with vague words - prompt: "Pick one more precise word."

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Take the plain noun group "the house".
- Expand it: add precise detail before and after the noun.
- Write it inside a short sentence. Two minutes.

DO:
- Two minutes silent.
- Collect on whiteboards or on the builder.

TEACHER NOTES:
Exit ticket targets SC2 - detail before and after the noun. Look for at least one precise word.

WATCH FOR:
- Detail in only one position - note for next lesson.
- Vague words only - note who needs the precise-words focus again.

[Literacy: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Self-check. Show on your fingers, 1 to 5, for each "I can".
- Partner share: read your best expanded noun group and name your most precise word.

DO:
- Run the fingers check for each success criterion.
- 60 seconds partner share.
- Acknowledge progress: "Your writing now paints a clearer picture than at the start."

TEACHER NOTES:
Use the fingers data to decide who needs another short build session. Connect back to the three success criteria.

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Expanded Noun Groups - Year 5/6 Literacy";

  // SLIDE 1 - Title
  titleSlide(
    pres,
    "Expanded Noun Groups",
    "Adding precise detail to paint a clearer picture",
    "Year 5/6 Literacy  |  Writing & Grammar",
    NOTES_TITLE
  );

  // SLIDE 2 - Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  // SLIDE 3 - Launch (plain vs vivid)
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Launch", { color: C.SECONDARY, w: 1.4 });
    addTitle(s, "Which Sentence Paints a Clearer Picture?");

    const cardY = CONTENT_TOP;
    const cardH = 1.35;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    s.addText("A", {
      x: 0.7, y: cardY + 0.12, w: 0.5, h: 0.5,
      fontSize: 26, fontFace: FONT_H, color: C.MUTED, bold: true, margin: 0,
    });
    s.addText("The dog ran across the yard.", {
      x: 1.25, y: cardY + 0.12, w: 8.0, h: cardH - 0.24,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.2;
    const cardBH = 1.55;
    addCard(s, 0.5, cardBY, 9, cardBH, { strip: C.PRIMARY, fill: C.BG_CARD });
    s.addText("B", {
      x: 0.7, y: cardBY + 0.12, w: 0.5, h: 0.5,
      fontSize: 26, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("The enormous, shaggy dog with muddy paws ran across the yard.", {
      x: 1.25, y: cardBY + 0.12, w: 8.0, h: cardBH - 0.24,
      fontSize: 24, fontFace: FONT_H, color: C.CHARCOAL, bold: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const promptY = cardBY + cardBH + 0.18;
    s.addShape("roundRect", {
      x: 0.5, y: promptY, w: 9, h: SAFE_BOTTOM - promptY, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    });
    s.addText("Show me A or B. Tell your partner ONE word that made the difference.", {
      x: 0.7, y: promptY, w: 8.6, h: SAFE_BOTTOM - promptY,
      fontSize: 15, fontFace: FONT_B, color: getContrastColor(C.SECONDARY), bold: true,
      valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  }

  // SLIDE 4 - LI / SC
  liSlide(
    pres,
    ["We are learning to expand a noun group by adding precise detail before and after the noun"],
    [
      "I can add describing words before a noun",
      "I can build a noun group with detail before and after the noun",
      "I can choose precise words that give the reader a clear, exact picture",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 - Vocabulary
  vocabSlide(
    pres,
    "expanded noun group",
    "grammar",
    "A noun group is a noun plus the words that describe it. We expand it by adding describing words before the noun and extra detail after it, to paint a clearer picture.",
    "a tall, glass tower above the busy city",
    NOTES_VOCAB,
    FOOTER
  );

  // SLIDE 6 - I Do (colour-coded build strip)
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do — Watch Me", { color: C.PRIMARY, w: 2.2 });
    addTitle(s, "Growing a Noun Group");

    // Plain start
    const plainY = CONTENT_TOP;
    addCard(s, 0.5, plainY, 9, 0.7, { strip: C.MUTED, fill: C.WHITE });
    s.addText("Plain:", {
      x: 0.75, y: plainY + 0.1, w: 1.1, h: 0.5,
      fontSize: 14, fontFace: FONT_B, color: C.MUTED, bold: true, valign: "middle", margin: 0,
    });
    s.addText("the dog", {
      x: 2.0, y: plainY + 0.1, w: 7.2, h: 0.5,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, valign: "middle", margin: 0,
    });

    // Hero strip
    nounGroupStrip(s, 0.6, 2.35, 8.8, [
      { words: "the", light: true },
      { words: "enormous, shaggy", color: C.SECONDARY, label: "describing words (before)" },
      { words: "dog", color: C.PRIMARY, label: "noun" },
      { words: "with muddy paws", color: C.ACCENT, label: "extra detail (after)" },
    ], { boxH: 1.05, labelH: 0.38 });

    // Read it in a sentence
    const readY = 4.0;
    addCard(s, 0.5, readY, 9, SAFE_BOTTOM - readY, { strip: C.PRIMARY, fill: C.BG_CARD });
    s.addText("In a sentence:", {
      x: 0.75, y: readY + 0.08, w: 8.5, h: 0.26,
      fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("The enormous, shaggy dog with muddy paws bounded across the yard.", {
      x: 0.75, y: readY + 0.34, w: 8.5, h: SAFE_BOTTOM - readY - 0.44,
      fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
    runSlideDiagnostics(s, pres);
  }

  // SLIDE 7 + 8 - CFU (hinge, with reveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Which Paints the Clearest Picture?", { color: C.ALERT });

      // CHECK wordmark (top-right) - accessible signal beyond colour
      const stampW = 1.3;
      s.addShape("roundRect", {
        x: 9.5 - stampW, y: 0.2, w: stampW, h: 0.32, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      });
      s.addText("CHECK", {
        x: 9.5 - stampW, y: 0.2, w: stampW, h: 0.32,
        fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Technique pill
      s.addShape("roundRect", {
        x: 0.5, y: CONTENT_TOP, w: 3.6, h: 0.42, rectRadius: 0.08, fill: { color: C.ALERT },
      });
      s.addText("Show Me Fingers: 1 (A), 2 (B), 3 (C)", {
        x: 0.5, y: CONTENT_TOP, w: 3.6, h: 0.42,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Options card
      const optY = CONTENT_TOP + 0.58;
      const optH = 2.0;
      addCard(s, 0.5, optY, 9, optH, { strip: C.ALERT, fill: C.WHITE });
      s.addText([
        { text: "A.  the car", options: { breakLine: true } },
        { text: "B.  the nice car", options: { breakLine: true } },
        { text: "C.  the rusty, red car with a cracked windscreen", options: { breakLine: false } },
      ], {
        x: 0.8, y: optY + 0.14, w: 8.4, h: optH - 0.28,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
        paraSpaceAfter: 6,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      return s;
    },
    (slide) => {
      T.addRevealAnswerBar(slide, "C - precise words before (rusty, red) and exact detail after (a cracked windscreen). 'Nice' is vague.", {
        label: "Answer", color: C.SUCCESS, showTickAndFix: false, fontSize: 14,
        y: 4.0, h: 1.0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 9 - Optional Re-teach (two-question build)
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Optional Re-teach", { color: C.SECONDARY, w: 2.6 });
    addTitle(s, "Two Questions Build the Detail");

    // Two guiding questions
    const qY = CONTENT_TOP;
    const qW = 4.4;
    addCard(s, 0.5, qY, qW, 0.95, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Before the noun", {
      x: 0.7, y: qY + 0.1, w: qW - 0.4, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("What is it like?", {
      x: 0.7, y: qY + 0.4, w: qW - 0.4, h: 0.45,
      fontSize: 19, fontFace: FONT_H, color: C.CHARCOAL, bold: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
    addCard(s, 0.5 + qW + 0.2, qY, qW, 0.95, { strip: C.ACCENT, fill: C.WHITE });
    s.addText("After the noun", {
      x: 0.7 + qW + 0.2, y: qY + 0.1, w: qW - 0.4, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText("Which one exactly? Where?", {
      x: 0.7 + qW + 0.2, y: qY + 0.4, w: qW - 0.4, h: 0.45,
      fontSize: 19, fontFace: FONT_H, color: C.CHARCOAL, bold: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Built example strip
    nounGroupStrip(s, 0.6, 2.85, 8.8, [
      { words: "the", light: true },
      { words: "wild, overgrown", color: C.SECONDARY, label: "what is it like?" },
      { words: "garden", color: C.PRIMARY, label: "noun" },
      { words: "behind the old fence", color: C.ACCENT, label: "which one / where?" },
    ], { boxH: 1.0, labelH: 0.36 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_RETEACH);
    runSlideDiagnostics(s, pres);
  }

  // SLIDE 10 + 11 - We Do (guided build with reveal)
  const wedoChunks = [
    { words: "the", light: true },
    { words: "", widthText: "ancient, crumbling", color: C.SECONDARY, label: "describing words (before)" },
    { words: "castle", color: C.PRIMARY, label: "noun" },
    { words: "", widthText: "on the misty hill", color: C.ACCENT, label: "extra detail (after)" },
  ];
  const WEDO_X = 0.6, WEDO_Y = 2.35, WEDO_W = 8.8, WEDO_GAP = 0.12, WEDO_BOXH = 1.05, WEDO_LABELH = 0.38;

  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SUCCESS);
      addBadge(s, "We Do", { color: C.SUCCESS, w: 1.4 });
      addTitle(s, "Build It Together: castle");

      const promptY = CONTENT_TOP;
      s.addShape("roundRect", {
        x: 0.5, y: promptY, w: 9, h: 0.7, rectRadius: 0.08, fill: { color: C.SUCCESS },
      });
      s.addText("On your whiteboards: what goes BEFORE? what detail goes AFTER? Keep it precise.", {
        x: 0.7, y: promptY, w: 8.6, h: 0.7,
        fontSize: 15, fontFace: FONT_B, color: getContrastColor(C.SUCCESS), bold: true,
        valign: "middle", margin: 0,
      });

      nounGroupStrip(s, WEDO_X, WEDO_Y, WEDO_W, wedoChunks,
        { boxH: WEDO_BOXH, labelH: WEDO_LABELH, gap: WEDO_GAP });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      runSlideDiagnostics(s, pres);
      return s;
    },
    (slide) => {
      const boxes = stripLayout(WEDO_X, WEDO_Y, WEDO_W, wedoChunks, WEDO_GAP, WEDO_BOXH);
      fillChunkText(slide, boxes[1], "ancient, crumbling", C.SECONDARY);
      fillChunkText(slide, boxes[3], "on the misty hill", C.ACCENT);
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // SLIDE 12 - You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Build Your Own Expanded Noun Groups");

    const cardY = CONTENT_TOP;
    const cardH = 1.9;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use the Noun Group Builder", {
      x: 0.75, y: cardY + 0.14, w: 8.4, h: 0.34,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   read the worked example (the box).\nNext:   build an expanded noun group for each noun.\nThen:   upgrade the plain sentence by growing its noun group.", {
      x: 0.75, y: cardY + 0.58, w: 8.4, h: cardH - 0.72,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      paraSpaceAfter: 8,
    });

    const tipY = cardY + cardH + 0.2;
    const tipH = 1.4;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Remember", {
      x: 0.75, y: tipY + 0.12, w: 5, h: 0.3,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Detail before AND after the noun.\n- Pick precise words - help the reader see the exact thing.", {
      x: 0.75, y: tipY + 0.5, w: 8.4, h: tipH - 0.62,
      fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
      paraSpaceAfter: 6,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  }

  // SLIDE 13 - Exit Ticket
  exitTicketSlide(
    pres,
    ["Expand \"the house\" into an expanded noun group, then write it inside a short sentence. Add precise detail before AND after the noun."],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, badgeColor: C.ACCENT }
  );

  // SLIDE 14 - Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Partner share: read your best expanded noun group and name your most precise word.",
      scItems: [
        "I can add describing words before a noun",
        "I can build a noun group with detail before and after the noun",
        "I can choose precise words that give the reader a clear, exact picture",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (could teach it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Noun Group Builder (student worksheet) -------------------------
  const wb = createPdf({ title: BUILDER_RESOURCE.name });
  let wy = addPdfHeader(wb, "Noun Group Builder", {
    color: PDF_NOUN,
    subtitle: "Grow each noun group: precise detail before AND after the noun.",
    lessonInfo: "Year 5/6 Literacy | Expanded Noun Groups",
    showNameDate: true,
  });

  wy = addTipBox(wb, "A noun group is a noun plus the words that describe it. Add describing words BEFORE the noun and extra detail AFTER it. Pick precise words that paint a clear picture.", wy, { color: PDF_NOUN });

  wy = addSectionHeading(wb, "Worked example", wy, { color: PDF_NOUN });
  wy = builderRow(wb, "box", wy, { before: "small, wooden", after: "covered in dust" });
  wy = addBodyText(wb, "In a sentence: The small, wooden box covered in dust sat in the corner.", wy, { fontSize: 11, italic: true });
  wy += 4;

  wy = addSectionHeading(wb, "Your turn - build each noun group", wy, { color: PDF_NOUN });
  wy = addBodyText(wb, "Write describing words before the noun and a detail after it.", wy, { fontSize: 10, italic: true });
  wy = builderRow(wb, "cat", wy, {});
  wy = builderRow(wb, "forest", wy, {});
  wy = builderRow(wb, "storm", wy, {});

  addPdfFooter(wb, "Noun Group Builder | Year 5/6 Literacy - Page 1");

  wb.addPage();
  let wy2 = addPdfHeader(wb, "Upgrade & Create", {
    color: PDF_NOUN,
    subtitle: "Grow the noun group to paint a clearer picture.",
    lessonInfo: "Year 5/6 Literacy | Expanded Noun Groups",
    showNameDate: false,
  });

  wy2 = addSectionHeading(wb, "Upgrade the sentence", wy2, { color: PDF_BEFORE });
  wy2 = addBodyText(wb, "Plain sentence:  The bird landed on the branch.", wy2, { fontSize: 12 });
  wy2 = addBodyText(wb, "Rewrite it with an expanded noun group - add precise detail before AND after \"bird\".", wy2, { fontSize: 10, italic: true });
  wy2 = addLinedArea(wb, wy2, 3, { lineSpacing: 26 });
  wy2 += 6;

  wy2 = addSectionHeading(wb, "Create your own", wy2, { color: PDF_AFTER });
  wy2 = addBodyText(wb, "Choose your own noun. Build an expanded noun group, then write it inside a sentence.", wy2, { fontSize: 10, italic: true });
  wy2 = builderRow(wb, "(your noun)", wy2, {});
  wy2 = addLinedArea(wb, wy2, 3, { lineSpacing: 26 });

  addPdfFooter(wb, "Noun Group Builder | Year 5/6 Literacy - Page 2");

  // ---- PDF: Answer Key -----------------------------------------------------
  const ak = createPdf({ title: ANSWER_RESOURCE.name });
  let ay = addPdfHeader(ak, "Noun Group Builder - Answer Key", {
    color: PDF_NOUN,
    subtitle: "Sample expanded noun groups. Many good answers are possible.",
    lessonInfo: "Year 5/6 Literacy | Expanded Noun Groups",
    showNameDate: false,
  });

  ay = addTipBox(ak, "These are samples, not the only answers. Look for detail BEFORE and AFTER the noun, and at least one precise word that helps the reader picture the exact thing.", ay, { color: PDF_NOUN });

  ay = addSectionHeading(ak, "Worked example", ay, { color: PDF_NOUN });
  ay = builderRow(ak, "box", ay, { before: "small, wooden", after: "covered in dust", boxH: 40 });
  ay += 2;

  ay = addSectionHeading(ak, "Sample answers - build each noun group", ay, { color: PDF_NOUN });
  ay = builderRow(ak, "cat", ay, { before: "sleek, black", after: "with green eyes", boxH: 40 });
  ay = builderRow(ak, "forest", ay, { before: "dark, silent", after: "after the storm", boxH: 40 });
  ay = builderRow(ak, "storm", ay, { before: "wild, howling", after: "over the bay", boxH: 40 });
  ay += 4;

  ay = addSectionHeading(ak, "Sample upgrade", ay, { color: PDF_BEFORE });
  ay = addBodyText(ak, "Plain:  The bird landed on the branch.", ay, { fontSize: 11 });
  ay = addBodyText(ak, "Upgraded:  The tiny, startled bird with bright feathers landed on the mossy branch.", ay, { fontSize: 11, italic: true });
  ay += 4;

  ay = addSectionHeading(ak, "Teacher look-fors", ay, { color: PDF_AFTER });
  ay = addBodyText(ak, "- Detail in both positions: describing words before AND a phrase after the noun.", ay, { fontSize: 10 });
  ay = addBodyText(ak, "- At least one precise word (not vague words like nice, good, big).", ay, { fontSize: 10 });
  ay = addBodyText(ak, "- The noun group still reads smoothly inside a full sentence.", ay, { fontSize: 10 });

  addPdfFooter(ak, "Noun Group Builder Answer Key | Year 5/6 Literacy");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Expanded Noun Groups.pptx` }),
    writePdf(wb, BUILDER_PDF_PATH),
    writePdf(ak, ANSWER_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Expanded Noun Groups.pptx`);
  console.log("Done: " + BUILDER_RESOURCE.name);
  console.log("Done: " + ANSWER_RESOURCE.name);
}

build().catch((err) => { console.error(err); process.exit(1); });
