# Orton-Gillingham (OG) Deck Builder Mega-Prompt v1.0
## Grade 5/6 Enrichment | Diamond Creek East PS | Yoshimoto OG | Template-Locked | Term-Scale Input

---

# 0. PURPOSE

This prompt governs the generation of weekly Orton-Gillingham morphology session decks.
It is the OG equivalent of `MEGA_PROMPT.md` + the teacher notes rules, but the design
philosophy is the OPPOSITE of the general lesson builder:

- The general builder designs slides. The OG builder NEVER designs slides.
- OG slides are a fixed, teacher-trusted instrument. Fonts, positions, colours,
  backgrounds, icons and click animations are locked in a master template. The only
  creative work is CONTENT: which morphemes, which words, which sentences, which notes.
- Students are 10-12 years old (Grade 5/6). The deck colours match the physical card
  deck they drill with: YELLOW = root, GREEN = prefix, RED = suffix. Never repurpose
  these colours.

The pipeline: user pastes a term's content -> you plan each week -> you author a week
spec JSON -> `og_planner/build_og_week.py` clones the master template and fills it ->
you QA the rendered decks. One PPTX per session (day), grouped in a week folder.

---

# 0a. NON-NEGOTIABLE OUTPUT GATE

1. NEVER build OG decks with PptxGenJS, the theme system, or any `builds/` script.
   The ONLY generation path is `og_planner/build_og_week.py` cloning
   `og_planner/OG_MASTER_TEMPLATE.pptx`.
2. NEVER edit `og_planner/OG_MASTER_TEMPLATE.pptx`. If the school issues a new master,
   replace the file wholesale and re-verify the shape-id map at the top of
   `build_og_week.py` (template slide indices and shape ids are hardcoded there).
3. NEVER change layout, fonts, font colours, background colours, icons, or animations
   through the spec. The builder auto-shrinks font sizes only when a word or sentence
   would overflow its box - that is the single permitted deviation, and it is automatic.
4. NEVER invent a morpheme keyword or meaning. They come from `og_planner/morpheme_bank.json`
   (the canon). Missing entry -> look it up in the `OG/` reference library (section 4),
   append it to the bank, then use it. The same morpheme must carry the same keyword and
   meaning every single time it appears, all year.
5. Every deck must pass the builder's gate (no leftover `XYZ`, reopenable file) AND a
   visual render inspection before you may call it done. Google Slides import is the
   final compatibility bar - if you cannot run it, say so explicitly.
6. ASCII-safe text everywhere: straight quotes, `-` not em dash, `...` not ellipsis.
   The builder sanitises, but write clean strings at the source.

---

# 1. THE FIVE-DAY SHAPE

Default week (adjust to the user's "sessions per week"):

| Day | type in spec | Morphology | Learned words | Grammar |
|---|---|---|---|---|
| Monday | `new` | NEW morpheme 1 taught | review 2 + new 1 | day 1 of weekly focus |
| Tuesday | `review` | morpheme 1 revisited (Yoshimoto day-2) | review 2 + new 1 | day 2 |
| Wednesday | `new` | NEW morpheme 2 taught | review 2 + new 1 | day 3 |
| Thursday | `review` | morpheme 2 revisited | review 2 + new 1 | day 4 |
| Friday | `week_review` | no new/review section - all-review spine | review 2, NO new | week review |

Every session, regardless of type, contains: title, weekly overview, morphology review
(10 cards), words to read review (15), sound bank (9), words to spell review (4-6 reveal
slides), learned words, dictation (x2), grammar (I do / We do / You do).

`new` and `review` days additionally contain the New/Review Morphology section (morpheme
card, words to read grid, words to spell grid, You Do activity - section 3d). On `review`
days the builder re-titles those section headers "Review Morphology" / "Words to Read-
Review Morphology" etc.

FILENAMES are auto-generated to the team's convention and must not be changed:
`1a. Monday (ord-ordin).pptx`, `1b. Tuesday (ord-ordin review).pptx`, `2a. Wednesday
(-ible).pptx`, `2b. Thursday (-ible review).pptx`, `3. Friday (week review).pptx` -
number = which of the week's morphemes, `a` = taught, `b` = reviewed, `/` in a morpheme
becomes `-`. Override only if the user asks, via `file_name` per session.

Rationale for review days (Yoshimoto "One Lesson Over 2 Days", in `OG/OG Post Training
Resources/OneLessonOver2Days.pdf`): day 2 is a short review of the same morpheme with the
weight shifted to auditory/spelling work - same morpheme, mostly fresh words.

---

# 2. THE ROLLING REVIEW ENGINE (the heart of every deck)

## 2a. Morphology review cards (10 per session)

- The 10 cards are the 10 MOST RECENTLY taught morphemes, counting backward from today,
  excluding today's new morpheme. A morpheme enters the review deck the session AFTER it
  is taught; when one enters, the oldest of the 10 drops out.
- Friday's 10 include both of this week's morphemes.
- JUMBLE the order every session. Hard rules:
  - Never the same order as any other session that week.
  - No morpheme may sit in the same position two sessions running.
  - No predictable type rotation (never root-prefix-suffix-root-prefix-suffix).
  - Never grouped by type, never alphabetical.
  - The goal: students must retrieve each morpheme cold. If the deck order can be sung,
    it has failed (the Sound Deck Drills note in the template warns against exactly this).
- Each card's notes carry `**Type:** / **Keyword:** / **Meaning:**` from the bank, so
  the teacher can prompt for keyword and meaning without breaking eye contact.
- Yoshimoto's Visual Card Drill (Morphology LessonPlan, see
  `og_planner/OG_LIBRARY_INDEX.md`) has students pronounce the morph, give the
  MEANING, and give a DERIVATIVE - occasionally used in a sentence. Every session,
  2-3 cards' notes append a `**Derivative ask:**` line with 1-2 example answers
  (`**Derivative ask:** who has a flect/flex word? EXPECT: flexible, deflect,
  reflection.`) - rotate which cards get it.

## 2b. Words to Read Review (15 words)

Retrieval weighting - most recent gets most attention. Allocate the 15 slots:

| Source | Slots |
|---|---|
| Last week's morphemes | 6 |
| Two weeks ago | 3 |
| Three weeks ago | 2 |
| Four weeks ago | 1 |
| Five+ weeks ago (rotate deeper history) | 1 |
| Earlier THIS week (from Tuesday onward) | 2 |

- On Monday (nothing taught yet this week), give the 2 this-week slots to last week (8 total).
- Early in a term with a short history, roll unfillable quotas up to the most recent
  teaching that exists. Never pad with untaught morphemes.
- FRESH WORDS DAILY: each session uses different derivative words for the same morphemes.
  A word may return after a gap (spaced retrieval), never two sessions in a row.
- Jumble placement - never group words by morpheme, never place a morpheme's words in
  the same table region across days.
- Every word must contain its morpheme DOING ITS JOB: the meaning must be visible in the
  word ("submerge" = under; "minister" does NOT teach mini = small - reject such words).
- Words the students met in a previous Words to Read/Spell - New Morphology section are
  ideal review candidates - that is the recall loop working.
- THIS SEGMENT IS ENTIRELY VERBAL. The words are on the screen; students have nothing
  in front of them and do not write in their books. Never ask students to underline,
  circle, highlight or write anything here. Students READ aloud, then ANSWER questions
  by looking at the board - a named student finds words, the whole class answers
  quick-fire questions chorally.
- NOTES ARE A TEACHER SCRIPT with CAPS anchors (READ / ASK / EXPECT / CHORAL), no
  bullets, no beat numerals. Required beats, in this exact order:
  1. `READ all rows together first. Today's split:` then EACH SPLIT ON ITS OWN LINE
     (this is how it scans on the iPad mid-drill):
     ```
     READ all rows together first. Today's split:
     Grade 5s - row 1
     Grade 6s - row 2
     Winter birthdays - row 3
     Wearing runners - row 4
     Everyone - row 5
     Push speed on the second read - repetition builds orthographic mapping.
     ```
     Reading the whole board comes FIRST, before any questions. The split is a FUN
     allocation that changes every day, covers every row, and gets nearly every
     student reading at least two rows. Build splits from this menu (invent more in
     the same spirit): grade 5s / grade 6s; boys / girls; birthdays Jan-Jun / Jul-Dec,
     birthday in the last two months, winter birthdays; wearing runners / something
     blue; left-handers / right-handers; walked or rode to school; name contains a
     double letter / the letter a; has a pet / plays a winter sport; everyone (the
     catch-all so no row is orphaned).
  2-4. Three `ASK:` beats tied to the morphemes on the board. Phrase them as the
     teacher would say them - `ASK: Which words have the root flect/flex (bend)? Name
     a student to find them.` - and ALWAYS follow with `EXPECT:` listing the answer
     words, each with a short 10-12yo gloss in parentheses:
     `EXPECT: flexible (bends easily without breaking), deflect (to bend something
     away).` Where the board allows, make one ASK a suffix-rule spot (`Which word had
     a spelling change before its suffix? EXPECT: pianist - piano drops its o`) or a
     transform (`How do I make testimony plural? EXPECT: change the y to i, add -es`).
  5. `5. CHORAL to finish, whole class answers together:` 3-4 quick-fire questions
     answered aloud in unison, EACH with its `EXPECT:` answer - the teacher does not
     memorise answers, so a question without its answer is a defect. Draw from: parts
     of speech, synonyms/antonyms, homophones, plurals/transforms, analogies,
     categories, words associated with..., which word might you hear at/in...
     Once or twice a week, run this beat as Yoshimoto's EXTENSION WHEEL game instead
     (`OG Mid Training Resources/Words to Read-Extension Wheel.pdf`): a named student
     picks a category for points - Rhyming (2), Meaning or use in a sentence (2),
     Synonym (3), Antonym (3), name all the Nouns/Verbs/Adjectives in a row (4),
     Analogy (5) - answers still scripted with EXPECT, next student picks, keep score.
- Every EXPECT answer must be checked against the actual word list on the slide - a
  scripted answer that is not on the board is a build-stopping content bug.
- TAUGHT-ONLY QUESTION RULE (applies to every question you script anywhere in a deck:
  ASK beats, CHORAL, after-checking prompts, extensions, dictation focuses). A
  question may only rely on knowledge the students have actually been taught:
  - Always safe: morpheme hunts and meanings (they are the taught content), transforms
    using taught suffixes (make it plural with -s, add -ly so it tells how), synonyms,
    antonyms, rhymes, categories, "which word might you hear at/in...", "which word
    means...".
  - GATED behind the grammar history: any metalanguage - part of speech (noun, verb,
    adjective, adverb), clause and phrase names, tense labels. Use these ONLY if a
    current or earlier grammar focus in the user's term plan covered them. "Part of
    speech of regulate?" is a defect unless parts of speech have been taught. When
    you want that idea without the label, ask it through the morpheme instead: "-ly
    means in that way - so flexibly means?" rather than "what part of speech is
    flexibly?".
  - Same gate applies to the Extension Wheel's Nouns/Verbs/Adjectives category - skip
    or substitute that sector until parts of speech have been taught.

## 2c. Sound bank (9 boxes)

- The 9 morphemes students will need for TODAY'S spelling work (review + new spelling
  words). Students copy them into the top of their page before spelling.
- Box colour is set by type automatically: yellow root, green prefix, red suffix -
  matching the physical card deck. Give the spec exactly 9 entries.
- GROUP BY TYPE: order the 9 entries so each type sits together as a row (or column) -
  roots together, prefixes together, suffixes together. A 3/3/3 split filling row 1 /
  row 2 / row 3 is the ideal; when counts are uneven, keep each type contiguous.

## 2d. Words to Spell Review (10 words, one reveal slide each)

- TEN words, not five. Weighting: 4 last week, 2 two weeks ago, 1 three weeks ago,
  1 four+ weeks ago, 2 earlier this week (Monday: give those to last week too).
- Must be DIFFERENT words from today's Words to Read Review (reading and spelling are
  different skills; overlap wastes the slot). Words previously READ on earlier days are
  excellent spelling candidates now.
- The morpheme must contribute its meaning in every chosen word.
- Spelling scope: the rest of each word must be spellable - check phonograms/patterns
  against what the group has been taught (OG sequence). When in doubt, choose the more
  decodable word.
- Every word's notes carry THREE things, all required (the builder warns on gaps):
  1. `**Sentence:**` - context sentence (teacher says word -> sentence -> word;
     confusables like intranet/internet MUST have a disambiguating sentence).
  2. `**After checking:**` - one morpheme-aimed question the teacher asks once the
     word is revealed (`Which part of the word tells you the travel is between states?`).
  3. `**Answer:**` - the answer to that question, in the same 10-12yo language the
     teacher can read verbatim (`inter- means between: travel between states.`).
     Teachers do not memorise thirty morpheme breakdowns - a prompt without its answer
     is a defect, not a nice-to-have.
- The builder centres every word at the same position with reveal-on-click and a blank
  screen between words (one slide per word) - this fixes the alignment/reveal feedback
  from the team email; you only supply words, sentences and prompts.

---

# 3. NEW / REVIEW MORPHOLOGY SECTION

## 3a. The morpheme card

- One new morpheme per `new` day. The card slide background matches the type
  (yellow/green/red) - the builder picks the right template slide from `type`.
- Notes: `Type - morph / Keyword: X / Meaning: Y` - keyword and meaning from the bank,
  verbatim, every time. This is the anchor slide the email asked to reinstate; never
  omit it.

## 3b. Words to Read (grid, reveal one word per click)

- AIM FOR 12. Floor is about 8 for a genuinely niche morpheme - reduce deliberately,
  never pad with fake derivatives.
- Layout is automatic and ALWAYS the template's 3-column grid (never reshape it to two
  columns - the school wants the 3-column look preserved): up to 9 words uses the
  template geometry untouched; 10-12 words keeps three columns but the builder widens
  them and stretches them down the page so the fourth row still reads large. Reveal
  stays one word per click either way.
- NOTES ARE A VERBATIM TEACHER SCRIPT (`wtr_new_notes`), not a glossary.
  Non-negotiable line formula, the SUBMARINE MODEL: sub means beneath + marine means
  water, so a submarine is a vessel beneath the water. Every line does exactly that:
  name EVERY morpheme in the word with its meaning joined by `+`, then a colon, then
  the fused whole-word meaning in 10-12 year old language:
  `**disorder** - dis- means apart + ord means order: the order pulled apart. A mess!`
  `**ordinary** - ordin means order + -ary means relating to: relating to the usual
  order of things. Normal.`
  `**insubordinate** - in- means not + sub- means under + ordin means order: refusing
  to sit under your place in the order.`
  For the bare root itself: `**order** - ord is the root on its own, order: things
  arranged in their places, one after another.`
  Two defects to avoid: a line that defines the word without naming its parts
  (`disorder - a mess`), and a circular line that only restates the root
  (`order - ORD means order: the way things are arranged` explains nothing - the
  parts must be seen DOING something to each other). Include the suffixes too (-ary,
  -al, -ate, -ance, -ation, -ly each get named with a kid-friendly meaning). Both
  morpheme spellings (ord AND ordin) must each appear across the script when the
  morpheme has variants.
- End the script with a `**Memory hook:**` line - one sticky retention device students
  can rehearse and reuse in their writing: a hearable slogan (`hear ORD, think "Off to
  my place in the line"`), an action, a family-collection challenge, or a
  build-the-longest-word game. Every new/review morphology session has one; reuse the
  same hook all week so it sticks.
- Grade 5/6 register: choose words that stretch 10-12 year olds (insubordinate,
  ordinance), not infant-level fillers. Same for the example sentences.
- On `review` days: same morpheme as the previous session, mostly NEW derivatives
  (morphological family extensions: coordination, subordination, preordain), with at
  most 2-3 repeats from the previous day.

## 3c. Words to Spell (4, revealed per click)

- Exactly 4, and all 4 MUST come from today's Words to Read grid (this is the one place
  read/spell overlap is required - students spell what they just decoded).
- Rotate the early-finisher extension in the notes; never the same one two days running.
  Menu: use two target words in one sentence with a subordinate clause; write the meaning
  in your own words using the root; find and record an antonym; turn the word into a
  question; add a prefix/suffix and explain the new meaning.
- Extensions must be SELF-CONTAINED: if the task names a concept (subordinating clause,
  adverbial phrase), include a one-line reminder of what it is right in the note, so no
  student stalls waiting for the teacher. Then add a `Stretch:` step roughly two year
  levels harder for the students who eat the first task (e.g. open with the subordinate
  clause and punctuate it; unpack a 3-morpheme word).

## 3d. New Morphology You Do activity (required on every new/review day)

After Words to Spell and before Review Learned Words, every new/review session gets a
dedicated Yoshimoto-style You Do: a magenta section header (auto-retitled `New
Morphology - You Do` / `Review Morphology - You Do`) followed by a DESIGNED task slide
(same structured layout as grammar, but in the New Morphology magenta: rule banner,
hero task card, up to two support lines, routine chip, footer). The generic "write a
fun sentence" is the extension note, not this - this is a real activity slide teachers
actually run. Supply it as `new_morph_activity` in the spec; the builder warns when a
new/review session lacks one.

- The task uses today's morpheme and is done on MINI WHITEBOARDS or in OG WORKBOOKS -
  the `routine` chip says which and how long (`Whiteboards - 3 minutes - boards up`).
- SPREAD THE ACTIVITIES ACROSS THE WEEK: every new/review day in the same week uses a
  DIFFERENT activity type from the menu - with the standard Mon/Tue/Wed/Thu pattern
  that means four different activities in a week, and no type repeats within the
  week. Vary the surface too (whiteboards one day, OG workbooks another). The menu:
  - Word building: add known prefixes/suffixes to the root, build N real words
    (LessonPlan.pdf: students brainstorm derivatives).
  - Morpheme sums, tachistoscope-style: teacher gives the meaning sum, students build
    the word (`together + order + to do = ? coordinate`).
  - Literal meanings: write the morpheme sum beside a given derivative and its fused
    meaning in your own words.
  - Sort derivatives by prefix / by which morpheme they carry (sorting by part of
    speech only once PoS has been taught - taught-only rule, 2b).
  - Sentence that SHOWS the meaning (not just uses the word).
- Every activity involves writing; every activity is startable without teacher help.
- Notes are Glance Format You Do style: numbered beats with SAY/TIME/CIRCULATE/SCAN,
  ANSWER first when the task has answers, TRAP with Fix, STRETCH and HELP, then `---`
  and one prep line.

---

# 4. THE MORPHEME BANK (canon workflow)

`og_planner/morpheme_bank.json` is append-only. For every morpheme in a deck:

1. Look it up in the bank. Found and `verified: true` -> use keyword + meaning verbatim.
2. `keyword: null` -> the anchor keyword is unconfirmed. Ask the user, or propose one
   drawn from the words-to-read family, flag it in your summary, and LOCK it in the bank.
3. Not in the bank -> consult the `OG/` reference library, navigated via
   `og_planner/OG_LIBRARY_INDEX.md` (read that index BEFORE hunting - it maps every
   subfolder and says which document governs which deck section). The PDFs are scans
   with no text layer: render the relevant pages to images (PyMuPDF via miniconda
   `python`) and read them visually. Primary canon sources:
   - `OG/Morphology USB files/Morphology.pdf` + `MorphologyPart2.pdf` - one worksheet per
     suffix/prefix with a definition box (e.g. "-ible is a suffix that means 'able to'").
   - `OG/Morphology USB files/Lower/Upper Level Latin Scrolls.pdf` - one section per
     root: derivative lists and meanings (best source for words-to-read families).
   - `OG/Morphology USB files/Greek Combining Forms.pdf` - Greek forms.
   - `OG/Morphology USB files/Morphology - Scope_Sequence.pdf` - teaching order.
   Append the entry with `source` and `verified: true`, then use it. Keyword
   provenance: Yoshimoto has students select a keyword to remember each morpheme; in
   these decks the keyword is fixed so every exposure matches - once locked, never
   changed. The Latin Scrolls derivative lists are also the first place to look when
   building a Words to Read family (3b).
4. NEVER change an existing verified entry. If the user corrects one, that correction is
   the new canon: update it once, note it in your summary, never drift again.

---

# 5. LEARNED WORDS (red words)

- Every session: REVIEW the two most recently taught learned words, then teach ONE new
  learned word (Friday: review only, no new).
- The word renders white on the dark red slide with the UNFAIR PART highlighted yellow
  (`unfair` field = the exact letters). The unfair part is the piece that breaks the
  spelling patterns students have been taught in Australian English - be precise
  (schwa vowels, silent letters, irregular graphemes), not decorative.
- A learned word must actually be unfair. If the word is fully decodable with taught
  patterns, tell the user it may not belong on the learned-word list - then follow
  their call.
- Notes = the WHY (students always ask). Structure: which part is unfair and what it
  does; an Australian-English pronunciation breakdown (syllables, the tricky sound);
  one usable teaching line. Where natural, link the word to the current inquiry unit
  (e.g. constitution, heritage -> First Nations inquiry) - a short spiel making the
  crossover explicit.
- The fixed procedure notes (far point / near point routine, "red words are STOP and
  think words") live on the section header slides in the template - untouched.

---

# 6. DICTATION (two sentences, both with reveal)

- Sentence 1 EASIER: about 10-12 words. Sentence 2 HARDER: about 14-16 words with
  richer punctuation. Never exceed ~16 words; the slide auto-shrinks but readability
  and working memory both suffer past that.
- Content sourcing, in priority order:
  1. Today's new learned word and/or this week's learned words.
  2. Today's (or yesterday's) morphology words.
  3. The 2b weighting recipe for older morphemes - dictation IS revision.
  Every content word must be spellable from taught material; scope-check phonograms.
- Grade 5/6 punctuation agenda (Victorian Curriculum level 5/6): beyond commas and full
  stops - colons, semicolons, quotation marks with correct comma placement, exclamation
  marks, hyphens, apostrophes, capitals for proper nouns. Weave the week's grammar focus
  in when it fits (grammar focus = colons -> a dictation sentence with a legal colon).
- `targets` = the words being assessed. The builder bolds + underlines them on the slide
  (so teacher and students know what to prioritise) and lists them in the notes with the
  `focus` line. 2-5 targets per sentence.
- CUPS MARKING (required for both sentences). The teacher marks dictation with CUPS -
  Capitals, Understanding (words written in the dictated order), Punctuation, Spelling.
  Give every dictation a `cups` block: `capitals` (the words that must start with a
  capital) and `punctuation` (each assessable mark, named). The builder then:
  - colours all punctuation RED in the revealed sentence (P),
  - keeps spelling targets bold + underlined (S),
  - prints `Score: ___ /N` on the slide, where N = capitals + punctuation marks +
    targets + 1 for Understanding, so students self-score against a visible total,
  - writes a tick-off CUPS checklist in the notes with each category's items and count.
  Sanity-check N against the sentence before shipping - every red mark, underline and
  capital on the slide must be counted once.
- Model pair (from the team email):
  - Easier: We followed the ordinary order, and then we began our work. (targets:
    ordinary, order)
  - Harder: The captain gave a clear order: "Keep every record in an orderly way, and
    check it twice." (targets: captain, order, orderly, record; colon, quotation marks,
    comma, compound structure)

---

# 7. GRAMMAR FINISHER (5 minutes, I do / We do / You do)

The one section where the general MEGA_PROMPT teaching craft applies - and the section
students find boring, so it must earn its five minutes.

- ONE grammar focus per week, taught as a 4-5 day sequence (Mon intro -> Fri week
  review). Link back to previous weeks' grammar when natural.
- ABSENT-STUDENT RULE (non-negotiable): every day's I do restates the core rule in one
  or two lines before extending it. A student who missed yesterday must be able to
  succeed today with only what is on screen. Five minutes is too short to re-teach, so
  the restatement IS the re-teach.
- DESIGNED SLIDES, NOT TEXT DUMPS. Grammar blocks use the structured spec and the
  builder lays them out (this is the one OG section with visual design): a purple rule
  banner across the top, a big soft-tinted HERO CARD carrying the example or task, up
  to two short support lines, a purple routine chip bottom-left, and an italic footer
  bottom-right. Author to that shape:
  - `rule` - the restated rule, one breath (the absent-student line).
  - `example` - ONE hero: the worked sentence, the judgement pair, or the task itself.
    This is what students look at; make it concrete, never meta.
  - `items` - max 2 short support lines (the test, the challenge, the if-stuck starter).
  - `routine` - the all-student response move as a chip: `Whiteboards ready - boards up
    together`, `Thumbs first, then whiteboards`, `90 seconds - boards up when done`.
  - `footer` - tease or stem: `Tomorrow: colon or no colon?`, `Stem: The colon belongs
    there because ...`.
  Never fall back to the legacy `lines` array for new decks - it produces the bland
  text-only slides the school rejected.
- Engagement is mandatory, not optional. Rotate all-student-response moves: mini
  whiteboards, thumbs verdicts, choral answers, beat-the-teacher (teacher makes the
  error, students catch it), partner swap-and-check, exit-ticket boards. Never
  volunteer-hands-only. Never a worksheet.
- You do produces visible evidence (board up, exit ticket) the teacher can scan to
  regroup tomorrow.
- Teacher notes use the Glance Format from `MEGA_PROMPT.md` sections 45-47: LIVE ZONE
  (ANSWER first if the slide asks anything; 2-5 numbered beats fusing action + talk with
  CAPS anchors - SAY/ASK/MODEL/SCAN/TIME/CIRCULATE; TRAP with Fix; STRETCH/HELP), then
  `---`, then a one-line prep zone with the `[Stage | element | focus]` tag. SAY lines
  are warm, natural classroom talk, one breath (~20 words), never clipped fragments.
- Content agenda: Victorian Curriculum English level 5/6 language strand - clause types,
  cohesion, punctuation for effect, complex sentences, quoted speech, apostrophes,
  modality - sequenced across the term by the user's grammar focus list.

---

# 8. TEACHER NOTES MAP (what is fixed, what you write)

FIXED (baked into the template on section headers - never modify):
- Sound Deck Drills variations (morphology review header)
- Words to Read procedure + older-student question menu
- New Morphology procedure (trace 3x, letter names, keyword, write 3x, underline)
- Learned Words far-point/near-point procedure
- Dictation procedure + correction protocol (fingerspell, learned-word wall, rule reminder)

GENERATED (you author per slide). FORMAT RULES - these notes are read live on an iPad
and must follow the MEGA_PROMPT.md Glance Format conventions, not ad-hoc styling:
- Flow notes (review table script, grammar) use NUMBERED BEATS with CAPS anchors
  (READ, ASK, EXPECT, CHORAL, SAY, MODEL, SCAN, TIME, CIRCULATE) - exactly the
  MEGA_PROMPT sections 45-47 voice. No blank lines inside a beat sequence.
- Data notes (cards, spell words, learned words, dictation) use `**bold label:**`
  then a plain value, one per line.
- `**bold**` markup renders as real bold in the notes pane. The builder strips the
  notes master's default bullets - never add your own dot points, dashes-as-bullets,
  or markdown bullets. Never a label followed by a bare hyphen (write `**Suffix:**
  -cy`, not `Suffix - -cy` - the double hyphen reads as noise beside hyphenated
  morphemes).
- Every question you script for the teacher carries its answer (`EXPECT:` or
  `**Answer:**`). Teachers read these cold on the iPad; no answer = defect.

| Slide | Notes content (verbatim template) |
|---|---|
| Morph review card | `**Type:** morph` / `**Keyword:** X` / `**Meaning:** Y` (bank verbatim, colon after the type) |
| Words to Read Review table | numbered verbal script per 2b: `1. READ all rows together first. Today's split: ...` then 3 `ASK: ... EXPECT: ...` beats then `5. CHORAL to finish ...` |
| Sound bank | fixed line (builder default) |
| Spell the word (each) | `**Word to spell:**` / `**Sentence:**` / reveal line / `**After checking:**` prompt / `**Answer:**` (all required) |
| New morph card | `**Type:**` / `**Keyword:**` / `**Meaning:**` |
| Words to Read grid | verbatim script per 3b: every line names the FOCUS MORPHEME in caps + `**Memory hook:**` last |
| Words to Spell grid | rotated extension, self-contained + `Stretch:` |
| Learned word (each) | `**Why learned:**` unfair part / `**Say it:**` AU pronunciation breakdown / `**Link:**` inquiry tie-in |
| Dictation (each) | sentence + CUPS tick-off checklist + `**Score:** /N` + `**Focus:**` |
| Grammar I/We/You do | Glance Format live zone, `---`, one prep line (section 7) |

---

# 9. WEEK SPEC JSON (the artefact you author)

One JSON per week in `og_planner/weeks/`. Exemplar: `og_planner/weeks/sample_term3_week1.json`
(a full Monday `new` + Tuesday `review` + Friday `week_review`). Schema:

```
{
  "term": 3, "week": 1, "cohort": "Enrichment",
  "unit_folder": "OG_Term3_Week1",            // output/<unit_folder>/
  "days": ["Monday", ...],
  "overview": { "<Day>": { "morphology": [label, value] | "text",
                            "grammar": "text",
                            "learned_words": [label, value] | "text" } },
  "sessions": [ {
    "day": "Monday",
    "type": "new" | "review" | "week_review",
    "morphology_review": [ { "morph", "type": root|prefix|suffix, "keyword", "meaning",
                             "derivative_ask"? (on 2-3 cards per session) } x10 ],
    "words_to_read_review": { "words": [15 strings], "notes": "activity script, 2b format" },
    "sound_bank": [ { "morph", "type" } x9 ],
    "words_to_spell_review": [ { "word", "sentence", "prompt", "answer" } x10 ],
    "new_morphology": { "morph", "type", "keyword", "meaning" },   // omit on week_review
    "words_to_read_new": [ { "word", "meaning" } x8-12 ],          // omit on week_review
    "wtr_new_notes": "verbatim teacher script + **Memory hook:** (3b)",
    "words_to_spell_new": [4 words, all from words_to_read_new],   // omit on week_review
    "extension": "self-contained early-finisher + Stretch:",       // omit on week_review
    "new_morph_activity": { "title", "rule", "example", "items": [<=2],
                            "routine", "footer", "time"?, "notes" }, // 3d; omit on week_review
    "file_name"?: "override only if the user asks",
    "learned_words": { "review": [ { "word", "unfair", "notes" } x2 ],
                        "new": { "word", "unfair", "notes" } },     // "new" omitted Friday
    "dictation": [ { "sentence", "targets": [..],
                     "cups": { "capitals": [..], "punctuation": [..] },
                     "focus" } x2 ],
    "grammar": { "header_notes",
                 "i_do":  { "title", "rule", "example", "items": [<=2], "routine", "footer", "notes" },
                 "we_do": { ... }, "you_do": { ... } }
  } ]
}
```

Notes:
- `unfair` must be an exact substring of the word (builder warns if not found).
- `**bold**` markup works in every notes field and in nothing on the slide faces.
- Keep single display words <= ~14 characters where you have a choice; the builder
  shrinks longer ones automatically.
- Straight ASCII only. The builder sanitises smart punctuation as a backstop.

---

# 10. BUILD AND QA (required, in order)

```bash
# deps live in miniconda python (NOT homebrew python3): lxml, python-pptx, pymupdf
python og_planner/build_og_week.py og_planner/weeks/<spec>.json          # all sessions
python og_planner/build_og_week.py og_planner/weeks/<spec>.json --only Tuesday
```

1. The builder's gate must pass (exit 0): file reopens cleanly, no `XYZ` left anywhere.
   Read every WARN - word-count and overflow warnings are content bugs to fix in the
   spec, not noise.
2. Visual QA every session deck: `python scripts/pptx_to_images.py "output/<folder>/<deck>.pptx"`,
   then INSPECT the images: title subtitle on one line; overview table cells complete;
   10 card slides in the jumbled order you specified; 15-word table filled; sound bank
   colours match types; each spell word centred; grid words not wrapping; learned-word
   highlight on the right letters; dictation targets flagged; grammar slides not
   overflowing. Local render substitutes fonts (Luckiest Guy/Lexend may look plain) -
   that is a render artefact, not a bug.
3. Animation spot check (structural): the spell-word, grid, learned-word and dictation
   slides must contain `<p:timing>` (the builder preserves/regenerates them - verify
   after any builder change).
4. Google Slides import pass for final sign-off. If not performed this session, report
   status as "automated gates + local visual QA passed; Google Slides pass pending".
5. Clean up: `python scripts/pptx_to_images.py --clean` (or `python3`; both work for this).

Multi-week asks: build every week, but do NOT merge OG decks into one unit PPTX -
the per-session PPTX in the week folder IS the deliverable (unlike themed lesson units).

---

# 10a. STEP-BY-STEP CHECKLIST (follow mechanically, in order, for every week)

1. Read `og_planner/weeks/sample_term3_week1.json` END TO END before writing anything.
   It is the gold exemplar: every notes format, every field, every recipe is
   demonstrated there. Your specs must match its formats exactly - same bold labels,
   same beat numbering, same field names. Do not invent alternative formats.
2. Build the week's morpheme timeline (which morpheme taught which day) and resolve
   every morpheme against `og_planner/morpheme_bank.json` (section 4). Stop and look
   up / ask about anything unverified BEFORE writing the spec.
3. For each session, in slide order, generate: review-10 (jumbled per 2a), 15 review
   words + verbal script (2b), sound bank 9 grouped by type (2c), 10 spelling words
   with sentence + prompt + answer (2d), morpheme card + words to read + script +
   memory hook + 4 spelling words + extension (3a-3c, skip on Friday), learned words
   2 review + 1 new (5), 2 dictations with cups (6), grammar structured blocks (7).
4. Self-check the spec against section 10b's failure list, line by line.
5. `python og_planner/build_og_week.py og_planner/weeks/<spec>.json` - fix EVERY warn.
6. Render each deck to images and INSPECT them (section 10). Fix, rebuild, re-render.
7. Only then report, stating exactly which QA levels ran.

# 10b. KNOWN FAILURE MODES (each of these has been rejected by the school - check
your spec against every line before building)

- Review cards in the same order two days running, or any singable pattern. (2a)
- A review word whose morpheme is not doing its job (mini in minister). (2b)
- Words to Read Review notes as a definition list, or asking students to underline/
  circle/write - this segment is fully verbal, students have nothing in front of
  them. (2b)
- READ-together missing, listed after the questions, or the row splits crammed onto
  one line instead of one split per line. (2b)
- A grid script line that is circular (`order - ORD means order: the way things are
  arranged`) instead of the submarine model naming every part with `+`. (3b)
- A new/review session without a `new_morph_activity`, or the same activity type
  twice in one week. (3d)
- A question relying on untaught metalanguage - e.g. "part of speech of regulate?"
  when parts of speech are not in the term's grammar history. (2b taught-only rule)
- Renaming output files away from the team convention (`1a. Monday (morph).pptx`). (1)
- A scripted question without its answer (EXPECT / **Answer:**) anywhere in any
  notes. (2b, 2d, 7)
- Fewer than 10 Words to Spell Review, or spelling words that duplicate the same
  day's Words to Read Review. (2d)
- A grid script line that does not name the focus morpheme in caps (writing
  `disorder - a mess` instead of `disorder - ORD means order, dis- means apart:
  the order has come apart`). (3b)
- Missing memory hook, or a new hook every day instead of the same hook all week. (3b)
- Words to Spell New not drawn from that day's Words to Read grid. (3c)
- An extension that names a concept without a one-line reminder of it, or with no
  Stretch step. (3c)
- A morpheme keyword/meaning that differs from the bank, or an invented one. (4)
- A learned word highlight on the wrong letters, or a why-note without the
  Australian pronunciation breakdown. (5)
- A dictation without a cups block, with an uncounted red mark/capital/target, or
  with sentence lengths outside 10-12 / 14-16 words. (6)
- Grammar authored as `lines` (the bland rejected layout) instead of rule / example /
  items / routine / footer. (7)
- An I do that does not restate the rule (absent-student rule). (7)
- Smart quotes, em dashes, or ellipsis characters anywhere in the spec. (0a)

# 11. PLANNING A TERM FROM THE USER'S INPUT

When the user pastes a term block (section 12):

1. Build the term timeline: which morpheme is taught which day, week by week. From it,
   derive every session's rolling review-10, the words-to-read/spell weighting pools,
   and the learned-word review chain (two most recent).
2. Resolve every morpheme against the bank (section 4) BEFORE writing specs. Ask about
   unconfirmed keywords in one batch, not drip-fed.
3. Author week specs in teaching order - later weeks' review pools depend on earlier
   weeks' words, so keep a running record of which derivative words each session used
   (words used in Words to Read/Spell - New Morphology become future review words).
4. Build and QA week by week. Deliver per-week folders under `output/`.
5. In your summary, list: any keyword you had to propose, any word you rejected for
   morpheme-integrity or decodability reasons, and any learned word that looked fully
   decodable.

If the user gives fewer than 5 sessions/week, compress: keep the Mon/Wed `new` pattern
first (e.g. 4 sessions = new, review, new, review; 3 = new, review, new), and fold week
review into the last session's review sections.

---

# 12. USER REQUEST FORMAT (user replaces the XYZs, nothing else)

The user fills in XYZ values only - same convention as MEGA_PROMPT.md. Repeat the
WEEK block for every week of the term. Everything not in this block (derivative
words, review lists, sound banks, spelling words, dictation sentences, grammar
scripts, teacher notes) is YOUR job to generate under the recipes above.

```
User: Generate OG session decks for the following:
Term: " XYZ "
Cohort: " XYZ "
Sessions per week: " XYZ "
Recently taught morphemes (most recent first, aim 10+): " XYZ "
Recently taught learned words (most recent first): " XYZ "

WEEK XYZ
Morphology focus: " XYZ "
Learned words: " XYZ "
Grammar focus: " XYZ "

WEEK XYZ
Morphology focus: " XYZ "
Learned words: " XYZ "
Grammar focus: " XYZ "

(repeat for every week of the term)

Additional notes: " XYZ "

Do not enter plan mode, proceed with the deck creation in bypass permissions. Ensure
you remain active while the decks are being created and continue to be until they are
fully complete, please.
```

Interpreting the block:
- `Morphology focus` lists the week's new morphemes with their type, e.g.
  `ord/ordin (root), -ible (suffix)`. Two per week is the default (taught Mon + Wed).
  One is fine (taught Mon; every other day reviews it).
- `Learned words` is the week's list in teaching order; deal one new word per session
  Mon-Thu, review-only Friday (section 5).
- `Sessions per week` under 5 compresses per section 11.
- `Additional notes` may carry inquiry links, keyword overrides (`ord/ordin =
  coordinate` locks the bank entry), assessment weeks, short weeks. If it is empty,
  proceed with defaults - do not ask.
- If `Recently taught morphemes` is empty and the bank/history gives you nothing,
  ask for it - week 1's review deck cannot be invented.
