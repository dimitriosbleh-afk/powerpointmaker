(c) 2026 James Hooke. Confidential. Internal use only. Not for redistribution.

# Auslan Session Deck Builder Prompt v1.0
## Step 2 of the Auslan pipeline | Turns the unit document from AUSLAN_1_UNIT_PROMPT.md into decks and printed resources | Runs inside the PPTX lesson generator repo

# 0. Position in the pipeline

The Auslan pipeline has two steps, used in order:

- **Step 1: AUSLAN_1_UNIT_PROMPT.md.** The Auslan teacher (or James on his behalf) runs it in a chat. It produces the unit document: band map, vocabulary bank, weekly plans, printables, assessment, prep calendar. That document is the plan.
- **Step 2: this prompt.** James pastes it into a session in this repo together with the unit document content and a list of which lessons to build. The output is the resources the plan only describes: one PowerPoint deck per teaching session (merged per the multi-session rules) plus the printable PDFs, built through the theme system.

This prompt states only the Auslan deltas. Everything it does not override follows, in this authority order:

1. The pasted unit document (the plan is already made; do not re-plan it)
2. This prompt
3. IMPORTANT/MEGA_PROMPT.md (current version) for slide design, notes, GRR, CFU and QA rules
4. CLAUDE.md and the build gates

**The deck standard is MEGA_PROMPT's, in full.** These are house-standard explicit teaching decks in the school's pedagogy: retrieval opening, LI and SC, modelled-guided-independent release, whole-class checks at decision points, exit ticket, Glance-format teacher notes, HITS-tagged, visual-first, low slide text, click builds, the lot. The Auslan teacher's existing slides are NOT the template; do not imitate their layout, density or design. The one thing his decks demonstrate is the sign-image mechanic - real sign images placed on slides as the thing students learn from - and section 3 turns that mechanic into a system. The division of labour: the unit document supplies WHAT is taught (content, signs, games, scripts, timings); MEGA_PROMPT supplies HOW a deck teaches it (every slide passes its visual anchor, hero task, low text, slide face and layout tests); this prompt supplies only what is different because the room is signing.

**The unit document is a contract, not a suggestion.** Its minutes, sign budgets, games, scripted Say: lines, decision points, band calibrations, care notes and resource specs are already decided and already reviewed. The deck's job is to put that plan on screen and on paper, faithfully. If the plan and good slide design genuinely collide (for example a We Do restatement too long for one slide), solve it with layout, never by rewriting the teaching. If something in the plan cannot be built at all, say so and ask; do not silently substitute.

# 1. Inputs

Paste with this prompt:

- The unit document, whole, or at minimum: unit header (anchor, care blocks), band map, the vocabulary bank rows for every sign in scope, and the full weekly plan of each lesson being built
- Which lessons to build (for example: Lessons 1 to 3, or the whole unit)
- Term and week numbers the lessons run in
- Class size if it differs from the unit document
- Anything the Auslan teacher changed since the document was generated

If the unit document was produced by a different or older prompt, build from what it gives and flag gaps rather than inventing the missing parts.

# 2. Theme and file layout

- Subject: **literacy** (Auslan is a language subject; no dedicated auslan theme exists yet). Year level from the band: Years 5-6 is grade56, Years 3-4 is grade34, and so on. If the unit spans bands taught as one mixed group, use the band the unit document says carries the core target.
- **One variant for the whole unit**, from the first teaching week: `weekToVariant(week)`. Never switch palettes between lessons of one unit.
- Build scripts: `builds/build_<unitprefix>_lesson<n>.js`, one per lesson, written directly in the main context (never by agents). Manifest at `builds/manifests/<unitprefix>.json` when more than one lesson is asked for in a single request; deliver the merged deck per the CLAUDE.md multi-session rules.
- Branch B consolidation lessons are conditional by definition. Do not build them unless the request names them. When built, they are their own sessions named Consolidation Week 1 and 2, appended after the fork's parent lessons.

# 3. Sign images: the hard rules

This is the rule set most likely to prevent real harm. A wrong sign on a slide gets taught to a whole class for a term.

**Never draw, generate, or approximate a sign.** Not with shapes, not with icons, not with AI-generated or stock images, not with a stick figure, not with arrows describing a movement. There is no such thing as a close-enough sign illustration. This extends the unit builder's Auslan language safety rule to visuals.

**Never describe sign production** (handshape, orientation, location, movement) on a slide, in notes, or on a PDF, unless the unit document itself supplies the description. Point to the lookup instead.

## The sign image bank

Sign images live in one shared, growing library: `assets/auslan_signs/`, built from Auslan Signbank by `scripts/fetch_auslan_signs.py`. Each image is a left-to-right sequence strip of frames taken from inside the sign, so the movement reads on a slide. See the bank's README.md for build and licensing detail. The conventions that matter at build time:

- One image per sign, named by gloss in caps: `TEAM.jpg`, `FAVOURITE.jpg`. Multi-word glosses use hyphens: `SLOW-DOWN.jpg`, `THANK-YOU.jpg`.
- Signbank strips are `.jpg`; hand-added line-art scans may be `.png`. Resolve either extension when looking a gloss up.
- A regional or alternate form is `<GLOSS>_2.jpg` (then `_3`). The unnumbered file is Signbank's first entry, which is not automatically the form the school teaches.
- The bank is shared across all Auslan units on purpose. Question signs, greetings, time markers and politeness signs recur every term, so each new unit adds only its own new glosses.
- **Run the fetcher before building a unit whose glosses are not in the bank yet:** `python scripts/fetch_auslan_signs.py --glosses <the unit's vocabulary bank>`. Glosses already present are skipped without a request, so this is cheap; glosses Signbank does not carry are reported MISSING and get lookup cards.
- **The images are gitignored and may be absent.** They are licensed for internal school use and this repo has a public remote, so only the recipe is tracked. If `assets/auslan_signs/` holds no images, rebuild the bank first: `python scripts/fetch_auslan_signs.py --from-file assets/auslan_signs/core_glosses.txt`. Never commit sign images, and never publish a built Auslan deck outside the school.
- Images are placed with the theme image helpers (`addImageWithCaption`, `addInstructionalImageCard`) using `fit: "contain"`. Never stretch, mirror, flip, crop into, or recolour a sign image. Mirroring reverses handedness and teaches the sign wrong.
- **`manifest.json` is the verification record**, holding the Signbank entry URL, keywords and dictionary definition behind every image. Where a lesson needs a sense the entry does not obviously carry, flag it in the notes prep zone rather than assuming the image is right.

## When an image exists vs when it does not

For every sign a slide needs, check the bank:

- **Image exists:** place it on a sign card with the English meaning as the visible label. The gloss may appear as a small caption; the meaning is what students read.
- **Image missing:** build the **lookup card** instead: the English meaning large, a "watch the teacher" line, and a hyperlink labelled `Look it up: Auslan Signbank` pointing at the search URL for the English word:
  `https://auslan.org.au/dictionary/search/?query=<word>`
  That search pattern is the only Signbank URL you may construct. Never guess an entry URL such as `/dictionary/words/team-1.html`; entry suffixes are unpredictable and a wrong link mid-lesson is worse than a search page.
- Never leave an empty image frame, and never fill the gap with a drawing or a stock photo.

## The sign asset report

Every Auslan build script must print, at the end of its run, a `SIGN ASSETS` report: which glosses resolved to bank images and which fell back to lookup cards. A deck with lookup cards in it still builds and still teaches (the teacher models every sign live anyway), but the report is how the gap gets closed, so never omit it and never claim the deck is image-complete when the report says otherwise. Missing glosses are usually fixed by running the fetcher; the ones that survive that are genuinely not in Signbank, which is itself worth telling the teacher, because it usually means the concept is a depicting sign or a phrase rather than a single lexical sign.

## Licensing and attribution

- Auslan Signbank is CC BY-NC-ND 4.0. The bank's strips are frame extractions from its videos, so they rest on the Australian schools statutory educational licence rather than on the CC terms: internal school teaching only, never redistributed outside the school, never commercial.
- Scanned pages from the school's purchased reference (for example Sign It!, Auslan Hub materials) sit under the same statutory educational licence (s113P), and the same internal-use limit.
- Every deck that places bank images carries one attribution line on the Teacher Resources slide, naming the source and the licence basis, for example: `Sign images: Auslan Signbank (auslan.org.au), CC BY-NC-ND 4.0, used for internal school teaching under the schools statutory educational licence.` Match the source name to what the bank README records for the images actually used, and name both sources if a deck mixes them.

# 4. Gloss and sentence safety on slides and PDFs

The unit builder's language rules bind this step too:

- A gloss is a label for looking a sign up, not Auslan and not a sentence. **Never compose a multi-sign gloss string** on any slide or PDF: no `WHEN YOU START?`, no gloss dialogues, no gloss word orders you assembled yourself.
- Where students read a question or answer, print it in plain English. The signs involved appear as separate sign cards or a listed sign set. The teacher models the Auslan form from the school reference.
- The one exception: a modelled structure the unit document or the school reference itself supplies, reproduced exactly, never extended. If the unit document tagged it CHECK GRAMMAR, the deck keeps that flag in the teacher notes prep zone.
- `[CHECK SIGNBANK]` and `[CHECK GRAMMAR]` tags never appear on a student-facing slide face or printable. They live in teacher notes prep zones, worded as the pre-lesson job they are.
- Sequence displays (for example yesterday / today / tomorrow across the slide) are sign cards side by side, each with its own English label. That is a vocabulary set laid out in time order, not a composed sentence, and it is allowed.

# 5. The deck skeleton for one Auslan lesson

The unit document's lesson skeleton (Do Now, LI and SC, I Do, We Do, You Do, Exit ticket) is the school's explicit teaching model wearing Auslan stage names, and it maps directly onto the house deck shape: roughly 12 to 16 unique slides for a 60 minute lesson, every one of them built with the tested theme builders to MEGA_PROMPT's design bar. Stage minutes from the run sheet go in the notes prep zones, not on slide faces.

1. **Title slide.** Unit title, lesson number and focus.
2. **Teacher Resources slide** (when the lesson ships PDFs): links to the session's printables, plus the sign-image attribution line, plus a CATCH-UP style note if the unit document gives one.
3. **Do Now: retrieval slides (1-2).** Previously taught signs shown as bank images with the labels hidden, revealed on click (`clickBuild`), so the class signs back or names the meaning before the answer lands. Signs without bank images are retrieved live by the teacher instead; do not build a retrieval slide around lookup cards. Finish with the bridge line from the unit document in the notes.
4. **LI and SC slide** (`liSlide`): the lesson's single learning intention and exactly the three I can statements from the unit document, verbatim.
5. **I Do: new sign slides (1-3).** The lesson's 3-4 production signs, hero-sized, one to four sign cards per slide. The first I Do slide carries the unit anchor restatement as a visible strip in the anchor's exact wording. The six-repetitions routine, the deliberate error and its Say: lines run live from the notes; do not put stage directions on the slide face. Variation slides (section 6) sit here when the lesson teaches one.
6. **We Do: game slide(s) (1-2).** The game named by the unit document, run with the teacher. An instruction card (`addInstructionCard`) carries the student-facing steps; the restated content the plan says the teacher must not leave the page for (the eight questions, the six fact cards) goes on the slide or its notes exactly as restated in the plan. Answer checks that reveal a gloss or meaning use `clickBuild`.
7. **Primary decision point slide.** The check as its own slide: the cue and what students do, student-facing. What the teacher watches for, the 80 percent proceed move and the full pivot live in the notes, taken from the unit document without compression.
8. **You Do: voice-off practice slide.** The task, the partner structure, and a visible `Voice off` marker. Auslan-only practice is the point of the stage, so the slide carries what students need to run it without the teacher talking: roles, steps, and where the cards sit. No scripted teacher talk here; the plan deliberately does not script the You Do.
9. **Exit ticket slide** (`exitTicketSlide`): the door routine and the self-assessment tick, matching the printed slip.
10. **Closing slide.**

Rules across the skeleton:

- Reveal budget and click mechanics follow MEGA_PROMPT: `clickBuild` is the mechanism, `withReveal` is the fallback for genuinely different layouts.
- Every lesson names its Deaf-friendly protocol for the week; give it one visible line on the slide where it is practised (usually the We Do or You Do), worded from the unit document.
- The care notes in the unit document are teacher-facing: notes zones, never slide faces.
- Support and Extend moves from the plan go in the notes as HELP and STRETCH lines on the core GRR slides, reworded only as much as the Glance format requires.
- Band calibration on a mixed-group deck: build to the carrying band, and put the other band's calibration in the notes line of the affected slides. Do not build parallel slide sets per band unless the request asks for separate class decks.

# 6. Sign images as the visual anchor

In an Auslan deck, the bank image IS the visual anchor MEGA_PROMPT demands: the sign is the content, so the image of it is the thing students look at, exactly as a fraction wall anchors a fractions lesson. Three placements cover nearly every need. Build them with theme primitives and bank images inside otherwise house-standard slides; they are placements, not new drawing helpers, and they must never involve drawing a sign.

- **Sign card.** One bank image on a card, English meaning as the visible label, gloss as a small caption. One to four per slide; at four, all four belong to one taught set.
- **Variation slide ("Which one?").** Two bank images side by side (`<GLOSS>.jpg` and `<GLOSS>_2.jpg`) under a title like `LIBRARY - which one do you use?`. Teaching point in the notes: regional variation is real, two signs can both be right, a student's home sign goes up next to the reference. Only build this when both variant images exist or the unit document explicitly teaches the variation; a variation slide with one lookup card is just a sign card.
- **Retrieval reveal.** Bank image(s) with the label hidden behind a `clickBuild` step. Used in Do Now and for We Do answer checks. The pre-reveal notes ask the class to sign back or name it with the signing response routine; the post-reveal beat confirms and fixes.

# 7. Teacher notes: voice-off deltas to the Glance format

Notes follow the Glance format and all of MEGA_PROMPT sections 45-47, with these substitutions for a signing room:

- **Say: lines come from the unit document verbatim** wherever it scripts them. They are what the teacher communicates (signed, spoken, or both, per school practice); do not convert them into hearing-classroom presenter talk and do not paraphrase them.
- **Response routines are the unit's signing routines,** cued identically in every Auslan deck: `Everyone signs it to me on three... one, two, three.` for expressive checks, and `Write the gloss. Boards up on three... one, two, three.` for receptive checks at bands with a writing load. These replace the school-standard verbal cue bank; the routine-tightness rules (full cue on first use, scripted one-line reset, never hands up) still apply.
- Hands up is never how evidence is collected, and in a signing room it also collides with the language itself: keep hands free for signing. Waving and table taps are attention protocols, not sampling methods.
- **EXPECT: names the gloss or meaning expected** (`EXPECT: FAVOURITE`, `EXPECT: most sign TEAM`), never a description of how the sign is produced.
- **ANSWER: lines state the meaning or gloss**, same rule.
- CHECK SIGNBANK and CHECK GRAMMAR items for the slide's signs go in the prep zone as a rehearse-first line, for example: `Rehearse FAVOURITE (varies) and negated UNDERSTAND in Signbank before class.`
- The prep zone tag keeps its `[Stage | VTLM element | SC | HITS n]` shape, using the HITS numbers the unit document's lesson tag names.
- The "teaching a language you are still learning" stance carries into notes: where a lesson meets a risky sign, the notes may script the lookup move (`Say: I am not sure of that one. Let us look it up together.`) exactly as the unit document words it.

# 8. Printable resources

The unit document's section 11 specifies each lesson's printables: content, band, print instruction, reuse. Build exactly those as companion PDFs with `pdf_helpers.js`, and nothing extra.

- Naming maps `Lesson N <Name>` to `Session N <Name>` to satisfy the pipeline's session-first naming and merge-uniqueness rules: `Session 2 Profile Question Cards`, `Session 7 Barrier and Solution Cards`, `Consolidation Week 1 Interview Grid`. Human-readable, spaces not underscores, no day names, no codes.
- Card sets print the English meaning as the readable text, with target sign glosses listed separately on the card, per section 4. Never a gloss sentence on a card.
- Sign images from the bank may be embedded on PDFs (picture prompt cards, support cards) under the same rules as slides: contain fit, no mirroring, attribution line in the PDF footer when scanned images are used.
- Exit ticket slips and self-assessment slips are the most printed items in the unit: spacious, one item per line, tick boxes drawn as real boxes, sized for the band. They are sorted at the door, never marked, so no marking scaffolding on them.
- Where the plan says a resource is reused from an earlier lesson, do not regenerate it under a new name; the earlier session's PDF is the resource, and the later lesson's Teacher Resources slide links or names it.
- Follow the unit document's print instructions (copies, laminate, cut) on the Teacher Resources slide notes so the deck carries its own prep line.

# 9. QA additions for Auslan decks

Everything in CLAUDE.md's QA section applies unchanged (build gates, markitdown, visual QA, Google Slides pass). On top of it:

- **Sign integrity sweep.** In visual QA, check every sign image slide: nothing stretched, nothing mirrored, nothing cropped into the hands or face, labels under the correct image. An image under the wrong label is the single worst defect this pipeline can produce; check labels against the vocabulary bank rows, not from memory.
- **Sign asset report reviewed.** The missing-sign list is surfaced to James in the final summary. Never report a deck as finished without it.
- **Sign identity checked against the manifest.** For every sign image placed, confirm `manifest.json`'s recorded definition matches the sense the lesson teaches. A right-looking image of the wrong sense is the failure this pipeline is built to prevent, and it is invisible in a rendered slide.
- **Gloss safety sweep.** Search the built deck text (markitdown output) for composed gloss strings and for CHECK tags that leaked onto slide faces. Both are blockers.
- **Link check.** Every Signbank hyperlink uses the search URL pattern and the query term matches the English word on the card.
- **Plan fidelity check.** Re-read the lesson's run sheet against the finished deck: every stage present, in order, minutes in the notes, the anchor restated in its exact words, both decision points present, You Do content different from We Do. The unit document is the contract; the deck is checked against it, not against taste.

# 10. Request format

The user request that accompanies this prompt looks like:

```
Build: Lessons 1-3
Unit prefix: <short prefix, for example deafsport>
Term/weeks: Term 4, weeks 1-3
Class size: 26 (if it differs from the unit document)
Changes since the document was generated: <none / list>

<paste the unit document, or its header + band map + vocabulary bank + the weekly plans in scope>
```

If the request names lessons whose weekly plans were not pasted, ask for those plans; do not reconstruct them from the lesson sequence summary.
