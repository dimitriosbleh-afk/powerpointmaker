(c) 2026 James Hooke. Confidential. Internal use only. Not for redistribution.

# Auslan Session Deck Builder Prompt v2.0
## Step 2 of the Auslan pipeline | Turns the unit document from AUSLAN_1_UNIT_PROMPT.md into decks, the student journal template and printed resources | Runs inside the PPTX lesson generator repo

# 0. Position in the pipeline

The Auslan pipeline has three steps, used in order:

- **Step 1: AUSLAN_1_UNIT_PROMPT.md.** Run in a chat. Produces the unit document: band map, vocabulary bank, weekly plans, printables, assessment, reporting map, prep calendar. That document is the plan.
- **Step 2: this prompt.** James pastes it into a session in this repo together with the unit document content and a list of which lessons to build. The output is one PowerPoint deck per teaching session (merged per the multi-session rules), the student journal template for the unit, and the printable PDFs, all built through the theme system.
- **Step 3: AUSLAN_3_TEACHER_LAYER_PROMPT.md.** Also run in this repo. Produces what the teacher plans and records from: the unit-at-a-glance, the landscape lesson pages, the cross-cohort weekly planner, the evidence tracker spreadsheet and the CRT review pack.

This prompt states only the Auslan deltas. Everything it does not override follows, in this authority order:

1. The pasted unit document (the plan is already made; do not re-plan it)
2. This prompt
3. IMPORTANT/MEGA_PROMPT.md (current version) for slide design, notes, GRR, CFU and QA rules
4. CLAUDE.md and the build gates

**The deck standard is MEGA_PROMPT's, in full.** These are house-standard explicit teaching decks in the school's pedagogy: retrieval opening, LI and SC, modelled-guided-independent release, whole-class checks at decision points, exit ticket, Glance-format teacher notes, HITS-tagged, visual-first, low slide text, click builds. The Auslan teacher liked the simplicity and colour of the first build; keep both. What he asked to change is listed in sections 3, 5, 6 and 7 and is now the rule.

**The unit document is a contract, not a suggestion.** Its minutes, sign budgets, games, scripted Say: lines, decision points, band calibrations, care notes and resource specs are already decided. The deck's job is to put that plan on screen and on paper, faithfully. If the plan and good slide design collide, solve it with layout, never by rewriting the teaching. If something in the plan cannot be built at all, say so and ask; do not silently substitute.

# 1. Inputs

Paste with this prompt:

- The unit document, whole, or at minimum: unit header (anchor, care blocks), band map, the term vocabulary list with Sign It! pages, the vocabulary bank rows for every sign in scope, the section 10 entries for every game in scope, and the full weekly plan of each lesson being built
- Which lessons to build (for example: Lessons 1 to 3, or the whole unit)
- Term and week numbers the lessons run in
- Cohort (Foundation, Discovery, Challenge or Enrichment) and, for Challenge and Enrichment running in parallel, which cohort's deck this is
- Anything the Auslan teacher changed since the document was generated
- Any Sign It! page scans supplied for this unit (see section 3)

If the unit document was produced by an older prompt version, build from what it gives and flag gaps rather than inventing the missing parts.

**Strip the unit document's visual layer on the way in.** The unit document's phase icons, `\newpage` markers and heading numbers exist to make a hundred-page Word document navigable. No icon reaches a slide face, a notes zone or a PDF; the slides have their own cue strip (section 6). Slide text and teacher notes stay ASCII per CLAUDE.md.

# 2. Theme and file layout

- Subject: **literacy** (Auslan is a language subject; no dedicated auslan theme exists yet). Year level from the cohort: Enrichment is grade56, Challenge is grade34, Discovery is grade2 (Year 1 calibration in the notes), Foundation is foundation.
- **One variant for the whole unit**, from the first teaching week: `weekToVariant(week)`. Never switch palettes between lessons of one unit. Challenge and Enrichment decks of the same topic share the variant.
- Lessons are specs: `builds/<unitprefix>_s<n>.json`, one per lesson, written directly in the main context (never by agents), validated and built with `build_and_check.js`. Manifest at `builds/manifests/<unitprefix>.json` when more than one lesson is asked for in a single request; deliver the merged deck per the CLAUDE.md multi-session rules. Write a JavaScript build script only when a spec cannot express a slide, and say why.
- A 50-minute lesson is 10 to 13 unique slides. Do not pad to the 60-minute count.

# 3. Sign visuals: the hard rules

This is the rule set most likely to prevent real harm. A wrong sign on a slide gets taught to a whole class for a term.

**Never draw, generate, or approximate a sign.** Not with shapes, not with icons, not with AI-generated or stock images, not with a stick figure, not with arrows describing a movement. There is no such thing as a close-enough sign illustration.

**Never describe sign production** (handshape, orientation, location, movement) on a slide, in notes, or on a PDF, unless the unit document itself supplies the description. Point to the lookup instead.

**Frame strips are retired.** The first build placed left-to-right frame strips cut from Signbank videos. The teacher's verdict: not clear enough for students to reproduce a sign from. No strip goes on any slide or PDF from v2.0 on. Two things replace them, in this order of preference:

1. **The moving sign.** An animated GIF made from the Signbank video for that gloss, `assets/auslan_signs/<GLOSS>.gif`:

   ```
   python scripts/fetch_auslan_signs.py --gif --links reference/auslan/signbank_links/<unit>.json
   python scripts/fetch_auslan_signs.py --gif --glosses TEAM PROUD COMPETE
   ```

   It starts at rest, runs the whole sign and returns to rest, 14 frames at about 11 a second, 280 pixels tall, looping. It plays in PowerPoint slideshow and in Google Slides. A sign lands around 250KB, so a deck of a dozen is a few megabytes. Same licence basis as the strips.

   **The links file is how the teacher's vetted entry wins over a search.** It is his own map of gloss to Signbank entry URL, JSON or one `GLOSS url` per line, kept in `reference/auslan/signbank_links/`. A gloss listed there is fetched from that exact entry and recorded as vetted in the manifest. Anything not listed falls back to the search and the run prints which glosses those were. Chase them before the deck ships: an unvetted sign is the one most likely to be the wrong sense.
2. **The Sign It! illustration.** Where the teacher has supplied scanned pages of the Auslan Hub textbook Sign It!, the illustration is cropped into `assets/auslan_signs/signit/<GLOSS>.png` and the slide captions it with the page number, so students can open the same page on their desks:

   ```
   python scripts/crop_signit_scans.py --page scans/p34.png --grid
   python scripts/crop_signit_scans.py --page scans/p34.png --crops p34.json
   ```

   The boxes are named by a person, once. The script will not guess which illustration is which sign, because getting one wrong is a class taught the wrong sign for a term. It writes a contact sheet to check every crop and records the page number in `signit_manifest.json`, which is where the slide caption comes from. Never mirrored, never recoloured, never cropped into the hands or face. Scans are teacher-supplied; nothing is fetched from Auslan Hub.

A new-sign slide uses the GIF. Where the Sign It! illustration also exists, it sits beside the GIF as the still students copy from at their own pace, page number under it. Retrieval slides may use either. PDFs, which cannot animate, use the Sign It! illustration where one exists, otherwise the English word and a "Sign It! p. N" reference, otherwise nothing: never a strip and never a still frame.

## The sign image bank

Sign assets live in one shared, growing library: `assets/auslan_signs/`, built from Auslan Signbank by `scripts/fetch_auslan_signs.py`. See the bank's README.md for build and licensing detail. The conventions that matter at build time:

- One asset per sign, named by gloss in caps: `TEAM.gif`, `FAVOURITE.gif`. Multi-word glosses use hyphens: `SLOW-DOWN.gif`, `THANK-YOU.gif`.
- A regional or alternate form is `<GLOSS>_2.gif` (then `_3`). The unnumbered file is Signbank's first entry, which is not automatically the form the school teaches.
- The teacher hands over the exact Signbank entry link for every sign on the term list. Fetch from the supplied entry, not from a search, whenever a link was supplied; record the link in `manifest.json`. Where no link was supplied, the fetcher's search fallback runs and the sign is flagged for him to confirm.
- **The images are gitignored and may be absent.** They are licensed for internal school use and this repo has a public remote, so only the recipe is tracked. Never commit sign images or GIFs, and never publish a built Auslan deck outside the school.
- Assets are placed with the theme image helpers using `fit: "contain"`. Never stretch, mirror, flip, crop into, or recolour a sign asset. Mirroring reverses handedness and teaches the sign wrong.
- **`manifest.json` is the verification record**, holding the Signbank entry URL, keywords and dictionary definition behind every asset. Where a lesson needs a sense the entry does not obviously carry, flag it in the notes prep zone rather than assuming the asset is right.

## When an asset exists vs when it does not

For every sign a slide needs, check the bank:

- **Asset exists:** place it on a sign card with the English meaning as the visible label and the Sign It! page as the caption when known. The gloss may appear small; the meaning is what students read.
- **Asset missing:** build the **lookup card** instead: the English meaning large, the Sign It! page if known, a "watch the teacher" line, and a hyperlink labelled `Look it up: Auslan Signbank` pointing at the search URL for the English word: `https://auslan.org.au/dictionary/search/?query=<word>`. That search pattern is the only Signbank URL you may construct. Never guess an entry URL. A supplied entry link may be used verbatim.
- Never leave an empty image frame, and never fill the gap with a drawing or a stock photo.

## Auslan Hub placeholders

Auslan Hub worksheets, card sets and videos are the teacher's licensed materials and are never reproduced. Where the unit document names one, the deck carries a placeholder card: the resource's Auslan Hub title, what students do with it, and one line "Teacher: attach from Auslan Hub" in the notes prep zone. The Teacher Resources slide may carry an "Open in Auslan Hub" link to the topic page the teacher supplies; it works only on his login and says so in its caption.

## The sign asset report

Every Auslan build must print, at the end of its run, a `SIGN ASSETS` report: which glosses resolved to GIFs, which to Sign It! illustrations, which fell back to lookup cards, and which Auslan Hub placeholders are unfilled. A deck with lookup cards still builds and still teaches, but the report is how the gap gets closed, so never omit it and never claim the deck is image-complete when the report says otherwise.

## Licensing and attribution

- Auslan Signbank is CC BY-NC-ND 4.0. The bank's GIFs are reductions of its videos, so they rest on the Australian schools statutory educational licence rather than on the CC terms: internal school teaching only, never redistributed outside the school, never commercial.
- Sign It! scans sit under the same statutory educational licence (s113P), and the same internal-use limit.
- Every deck that places bank assets carries one attribution line on the Teacher Resources slide, naming each source used and the licence basis, for example: `Sign videos: Auslan Signbank (auslan.org.au), CC BY-NC-ND 4.0. Sign illustrations: Sign It! (Auslan Hub). Both used for internal school teaching under the schools statutory educational licence.`

# 4. Gloss and sentence safety on slides and PDFs

The unit builder's language rules bind this step too:

- A gloss is a label for looking a sign up, not Auslan and not a sentence. **Never compose a multi-sign gloss string** on any slide or PDF.
- Where students read a question or answer, print it in plain English. The signs involved appear as separate sign cards or a listed sign set. The teacher models the Auslan form.
- The Do Now slide for Challenge and Enrichment carries the English prompt only. Students write their own gloss on their boards from Sign It!. A model gloss appears on a click reveal only when the unit document carries one supplied by the teacher, attributed in the notes.
- The one exception: a fixed form the unit document, the teacher or the school reference supplies, for example `WHAT MEAN?`. Reproduce it exactly, never extend it. If the unit document tagged it CHECK GRAMMAR, the deck keeps that flag in the teacher notes prep zone.
- `[CHECK SIGNBANK]` and `[CHECK GRAMMAR]` tags never appear on a student-facing slide face or printable. They live in teacher notes prep zones.
- Sequence displays (yesterday / today / tomorrow across the slide) are sign cards side by side, each with its own English label. That is a vocabulary set laid out in time order, not a composed sentence, and it is allowed.

# 5. The deck skeleton for one Auslan lesson

The unit document's lesson skeleton (Do Now, LI and SC, I Do, We Do, You Do, Exit ticket) maps directly onto the house deck shape: 10 to 13 unique slides for a 50 minute lesson, every one built with the tested theme builders. Stage minutes from the run sheet go in the notes prep zones, not on slide faces.

1. **Title slide.** Unit title, lesson number and focus.
2. **Teacher Resources slide** (when the lesson ships PDFs or names Auslan Hub materials): links to the session's printables, the Auslan Hub placeholders, the attribution line, and a CATCH-UP style note if the unit document gives one.
3. **Do Now slide (1).** Seated, silent, independent. The stage badge reads Do Now and the cue strip shows voices off and whiteboards. Challenge and Enrichment: the English prompt, hero-sized, and nothing else on the face; the notes carry the retrieval items the prompt is built to surface and the bridge line. Foundation and Discovery: the Auslan Hub worksheet placeholder card with the one-line instruction. Never a stand-up task here. A retrieval reveal slide (section 6) may follow only when the unit document's Do Now asks for teacher-signed items.
4. **LI and SC slide** (`liSlide`): the lesson's single learning intention and exactly the three I can statements from the unit document, verbatim.
5. **I Do: new sign slides (1-3).** The lesson's five production signs as moving signs, one to four sign cards per slide, hero-sized. The first I Do slide carries the unit anchor restatement as a visible strip in the anchor's exact wording. The cue strip shows eyes up and watch then copy. The six-repetitions routine, the deliberate error and its Say: lines run live from the notes. Lesson 1 of a unit adds one opening-bank slide: the whole term list as small sign cards with English labels, met receptively, and the same list is the journal's front page.
6. **We Do: game slide(s) (1-2).** The game named by the unit document, run with the teacher. An instruction card carries the student-facing steps; the restated content the plan says the teacher must not leave the page for goes on the slide or its notes exactly as restated. Name the game on the slide by its appendix number and title. Safety rules the students need (no running, no dodging) go on the face. Timed games carry the countdown (section 6).
7. **Primary decision point slide.** The check as its own slide, student-facing: for Challenge and Enrichment the default from the unit document is that the teacher signs a whole question or sentence, students write its English meaning, hover it, chin it on three, then one cold-called student signs the answer back. The face carries the cue and what students do; what the teacher signs, watches for, the 80 percent move and the pivot live in the notes, taken from the unit document without compression. Never build the check as "sign this one word".
8. **You Do: voice-off practice slide.** The task, the partner structure (table partner by default, mix and mingle when the plan says so), a visible `Voice off` cue and the countdown. The slide carries what students need to run it without the teacher talking: roles, steps as chips, where the cards sit. No scripted teacher talk. The extend task is a visible chip.
9. **Exit ticket slide** (`exitTicketSlide`): the journal reflection everyone does (SC self-rating on the three words Just starting / Getting there / Got it, one line, the week's cultural question in a Deaf-culture unit), plus a line saying the teacher will call some students to sign. The rotation itself never appears on the face; the notes name the six or seven due this week from the tracker and the three signs plus the SC2 exchange to ask for. Foundation and Discovery: the picture-card routine, table by table.
10. **Closing slide.**

Rules across the skeleton:

- Reveal budget and click mechanics follow MEGA_PROMPT: `clickBuild` is the mechanism, `withReveal` the fallback for genuinely different layouts.
- Every lesson names its Deaf-friendly protocol for the week; give it one visible line on the slide where it is practised, worded from the unit document.
- Care notes are teacher-facing: notes zones, never slide faces.
- Support and Extend moves go in the notes as HELP and EXTENSION lines on the core GRR slides; the You Do extend chip is the one student-facing exception.
- Challenge and Enrichment decks of the same topic are built as separate decks from the same unit document, each to its own cohort calibration. Do not ship one deck with the other cohort's calibration in the notes.

# 6. The visual system: cue strip, sign cards, countdown

**The cue strip.** `T.addCueStrip(slide, ["voicesOff", "eyesUp", "watchCopy"])` draws it: outlined chips, right-aligned, the same image and the same words for the same cue on every slide, in the same place. `T.listCues()` returns the six. An unknown name prints a gate-failing WARN rather than drawing nothing. The set, and the only set:

| Cue | When it appears |
| --- | --- |
| Voices off | Do Now, You Do, every game, any slide where students sign without voice |
| Eyes up | Every I Do sign slide, every check |
| Watch, then copy | Every I Do sign slide |
| Partner | Every paired task |
| Whiteboards | Every task that writes on boards |
| Timer | Every timed task, beside the countdown |

Each cue keeps its own colour on every slide of every deck, set in the theme, so a student finds it by colour before they read it. They are sized for the back of the room, not for the teacher's screen.

Nothing else gets a cue, and a seventh is a change to this prompt, not to a build script. The stage badge (Do Now, I Do, We Do, You Do, Exit ticket) is larger and higher-contrast than the first build's; it is the first thing a student with working-memory load reads.

The cue strip has its own page in `output/Visual_Catalogue/`. Rebuild and re-inspect it after any change to the cues.

**Sign cards.** In an Auslan deck the sign asset IS the visual anchor MEGA_PROMPT demands. Three placements cover nearly every need; they are placements, not new drawing helpers, and they must never involve drawing a sign.

- **The Sign It! cover, wherever the book is referenced.** Any slide that tells students to look something up shows `addSignItCover`, so the book on the screen is the book on the desk. Chris supplies the cover scan; until it lands the helper draws a labelled placeholder rather than nothing.
- **Sign card.** One asset on a card, English meaning as the visible label, Sign It! page as caption, gloss small. One to four per slide; at four, all four belong to one taught set. One sign is `heroVisualSlide` with `{ type: "image", path }`; two to four are `choiceSlide` options with `caption` set to the meaning and `{ letters: false }`. Never a pictogram in place of a sign.
- **Variation slide ("Which one?").** Two assets side by side (`<GLOSS>.gif` and `<GLOSS>_2.gif`) under a title like `LIBRARY - which one do you use?`, labels `A` / `B`. Build it when the unit document's vocabulary bank fills the Variant column for a sign the lesson teaches, and only when both assets exist. Label which form the class produces and which they only need to recognise.
- **Retrieval reveal.** Sign asset(s) with the label hidden behind a `clickBuild` step. Used for teacher-signed retrieval and We Do answer checks.

**The countdown.** Every timed task (mix and mingle, timed games, You Do) carries a countdown GIF beside the Timer cue:

```
python scripts/make_countdown_gif.py 120 --color <the theme's C.PRIMARY>
```

It writes `assets/timers/countdown_<seconds>s_<COLOR>.gif` and prints the path: one frame per second, a ring that empties clockwise from twelve, the number turning red in the last ten seconds. It plays once, because a countdown that loops back to full is worse than no timer. Pass the theme's `C.PRIMARY` so it matches the deck. The file is cached, so calling it for every timed slide costs nothing after the first.

It starts when the slide appears in slideshow, so the slide is built to be advanced to at the moment the task starts, and the notes say `advance to start the clock`. Google Slides plays it; the LibreOffice preview shows only the first frame, which is expected.

# 7. Teacher notes: what the teacher asked for

He reads the notes on an iPad while teaching, at a glance, and he goes off script by design. The notes are non-negotiables, not a transcript. They still follow the Glance format so the build gates pass, with these Auslan substitutions:

- **Live zone shape:** `ANSWER:` line when the slide asks anything, then three to five numbered beats, each ONE verb-first line with a CAPS anchor (`MODEL`, `POINT`, `TIME`, `CIRCULATE`, `COLLECT`, `REVEAL`), then `TRAP:`, then `HELP:` and `EXTENSION:` on the core GRR slides. Aim for 60 to 80 words, never the 120 ceiling.
- **Say: lines are kept only where the wording matters:** the anchor restatement, the response cue scripts, the deliberate-error line, the feedback line at the decision point, and the bridge. Everything else is an action, not speech. Where the unit document scripts a Say: line the notes keep, reproduce it verbatim.
- **Response routines are the unit's signing routines,** cued identically in every Auslan deck: `Everyone signs it to me on three... one, two, three.` for expressive checks, and `Write it. Hover it. Chin it on three... one, two, three. Chin it.` for receptive checks. Full cue on first use in a deck, shortened after.
- Hands up is never how evidence is collected. Waving and table taps are attention protocols, not sampling methods.
- **EXPECT: names the gloss or meaning expected**, never a description of how the sign is produced. **ANSWER: lines state the meaning or gloss.**
- CHECK SIGNBANK and CHECK GRAMMAR items go in the prep zone as a rehearse-first line.
- The exit ticket slide's notes carry the week's rotation names from the tracker and the three signs to ask for, so he never opens a second document at the door.
- The prep zone keeps its `[Stage | VTLM element | SC | HITS n]` tag and the stage minutes.
- The "teaching a language you are still learning" stance carries into notes only when the unit document carries that block.

# 8. Printable resources and the journal template

**Printables.** The unit document's section 11 specifies each lesson's printables. Build exactly those with `pdf_helpers.js`, and nothing extra. Expect few: Challenge and Enrichment write in exercise books and journals, so a lesson's printables are usually a card set or nothing.

- Naming maps `Lesson N <Name>` to `Session N <Name>`: `Session 2 Team Role Cards`. Human-readable, spaces not underscores, no day names, no codes.
- Card sets print the English meaning as the readable text, with target sign glosses listed separately on the card. Never a gloss sentence on a card.
- Sign visuals on PDFs follow section 3: Sign It! illustration with page number, or the page reference alone. Never a strip.
- Auslan Hub items are not built. They get a one-page cover sheet in the resources folder naming the Auslan Hub title and what to print, so the teacher's prep list is complete.
- Where the plan says a resource is reused from an earlier lesson, do not regenerate it under a new name.

**The receptive test frame.** Piece 2 in the unit document is written in the exercise book from a frame on the screen. Build that frame as a slide in the lesson it lands in (numbered lines, one per item, nothing else on the face) and, for Foundation and Discovery, as the printed circle-or-match sheet the plan specifies.

**The journal template.** One PPTX per unit per cohort, `output/<UnitFolder>/<Unit> Student Journal Template.pptx`, built with the same theme and variant, which the teacher uploads to Google Slides once and students copy into their Google Classroom assignment. It carries:

1. A cover page: unit title, cohort, a name line, and one line saying this is where the family sees what the student can do.
2. The term vocabulary tracker: the opening bank as English words with Sign It! pages in a table, three tick columns (Start of term, Middle, End), no sign images (the journal is copied by students and lives outside the school's asset control).
3. One page per lesson: the learning intention and three I can statements, each with a three-way self-rating in student words that map to the ACARA checklist scale (Just starting = Emerging, Getting there = Consolidating, Got it = Proficient), the same three words on every page and on the closing slide of every deck; an empty video slot outlined and labelled "Insert your video here"; the reflection prompt; and the cultural question for that week.
4. A final page: "Three things I can sign now that I could not sign in week 1".

Build it through the theme with the tested builders; treat it as a deck whose audience is the student and the family, so slide text stays at the band's reading level.

# 9. QA additions for Auslan decks

Everything in CLAUDE.md's QA section applies unchanged (build gates, markitdown, visual QA, Google Slides pass). On top of it:

- **Sign integrity sweep.** In visual QA, check every sign slide: nothing stretched, nothing mirrored, nothing cropped into the hands or face, labels under the correct asset, page numbers matching the vocabulary bank. An asset under the wrong label is the single worst defect this pipeline can produce; check labels against the vocabulary bank rows, not from memory.
- **Sign asset report reviewed.** The lookup-card and unfilled-placeholder lists are surfaced to James in the final summary. Never report a deck as finished without them.
- **Sign identity checked against the manifest.** For every asset placed, confirm `manifest.json`'s recorded definition matches the sense the lesson teaches. The shared layer does this automatically for senses it knows about: a gloss listed in `SENSE_NEEDED` resolves to the variant whose recorded definition carries that sense, so a question sign never ships as the noun. Read a whole definition before concluding a sign is wrong. Signbank lists several uses per entry and the question use is often the last one; the first clause alone told us WHO and WHAT had no question form when both do.
- **A lookup card is the last resort, not the first answer.** If the default entry looks wrong, fetch the variants and check them before shipping a card. Chris overrides any of it from `reference/auslan/signbank_links/overrides.json`, which beats the sense search, the corrections and the default.
- **Animation check.** Open the merged deck in PowerPoint or Google Slides and confirm the GIFs and countdowns play; markitdown and LibreOffice cannot see this.
- **Gloss safety sweep.** Search the built deck text for composed gloss strings and for CHECK tags that leaked onto slide faces. Both are blockers.
- **Link check.** Every Signbank hyperlink is either a supplied entry link or the search URL pattern with the query matching the English word on the card.
- **Cue strip check.** Every slide that should carry a cue carries it, in the same position, with the same image. A cue that drifts between slides has failed.
- **Plan fidelity check.** Re-read the lesson's run sheet against the finished deck: every stage present, in order, minutes in the notes, the anchor restated in its exact words, both decision points present, You Do content different from We Do, the Do Now seated and silent.

# 10. Request format

The user request that accompanies this prompt looks like:

```
Build: Lessons 1-3
Unit prefix: <short prefix, for example deafsport>
Cohort: Enrichment
Term/weeks: Term 4, weeks 1-3
Changes since the document was generated: <none / list>
Sign It! scans supplied: <none / folder>

<paste the unit document, or its header + band map + term list + vocabulary bank + the weekly plans in scope>
```

If the request names lessons whose weekly plans were not pasted, ask for those plans; do not reconstruct them from the lesson sequence summary.
