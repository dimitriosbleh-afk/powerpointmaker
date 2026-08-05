# Orton-Gillingham (OG) Deck Builder Mega-Prompt v1.1
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
- Students are 10-12 years old (Grade 5/6). The universal morpheme colour key is:
  GREEN = root, YELLOW = prefix, RED = suffix. This applies to review-card
  backgrounds, new/review morpheme-card backgrounds and Sound Bank boxes. Never
  reverse, repurpose or infer a different mapping from an older template slide.

The v1.1 refinement tightens response quality and responsive pacing inside the locked
Yoshimoto sequence. It does not change the five-day shape, required review counts,
morphology sequence, dictation, grammar finisher, master template or output naming.

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
3. NEVER change layout, fonts, font colours, icons, animations or arbitrary background
   colours through the spec. The builder automatically normalises every type-coded
   card and Sound Bank box to GREEN root / YELLOW prefix / RED suffix, even when the
   locked master contains an older colour. It also auto-shrinks font sizes only when a
   word or sentence would overflow its box. These are builder rules, not spec options.
4. NEVER invent, modernise, paraphrase, or web-substitute a morpheme keyword or
   meaning. Before planning any OG lesson, load the three authoritative photographed-
   card files named in section 4. For any of their 277 captured cards, the catalogue's
   `morpheme`, `meaning`, `keyword`, and printed `part_of_speech` are locked. Only a card
   absent from those catalogues may use the fallback workflow in section 4. The same
   physical card must carry the same metadata every time it appears.
5. Every deck must pass the builder's gate (no leftover `XYZ`, reopenable file) AND a
   visual render inspection before you may call it done. Google Slides import is the
   final compatibility bar - if you cannot run it, say so explicitly.
6. ASCII-safe text everywhere: straight quotes, `-` not em dash, `...` not ellipsis.
   The builder sanitises, but write clean strings at the source. Sole exception,
   applied by the builder itself: paired straight quotes render as directional
   speech marks on the dictation slide face (section 6) - never type curly quotes
   into a spec.

---

# 1. THE FIVE-DAY SHAPE

Default week (adjust to the user's "sessions per week"):

| Day | type in spec | Morphology | Learned words | Grammar |
|---|---|---|---|---|
| Monday | `new` | NEW morpheme 1 taught | review 2 + NEW word 1 | day 1 of weekly focus |
| Tuesday | `review` | morpheme 1 revisited (Yoshimoto day-2) | review 2 only | day 2 |
| Wednesday | `new` | NEW morpheme 2 taught | review 2 + NEW word 2 | day 3 |
| Thursday | `review` | morpheme 2 revisited | review 2 only | day 4 |
| Friday | `week_review` | no new/review section - all-review spine | review 2 only | week review |

New learned words land in sessions 1 and 3 only - never back to back. Full schedule
for shorter weeks in section 5.

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
- Each card's notes carry `Type: / Keyword: / Meaning:` from the canonical card
  catalogue (plus `Part of speech:` when the physical card supplies one), so
  the teacher can prompt for keyword and meaning without breaking eye contact.
- Yoshimoto's Visual Card Drill (Morphology LessonPlan, see
  `og_planner/OG_LIBRARY_INDEX.md`) has students pronounce the morph, give the
  MEANING, and give a derivative - occasionally used in a sentence. Every session,
  2-3 cards add an optional `extra_task` object with a direct teacher prompt and
  possible answers. The rendered label is `Extra task:`, never the unclear
  `Derivative ask:`. Example: `SAY: "Give me one flect or flex word, meaning bend."`
  followed on a new line by `Possible answers: flexible, deflect, reflection.` Rotate
  which cards receive it.
- PATTERN MORPHEMES: some bank entries are pronunciation/spelling patterns, not
  meaning morphemes (their `meaning` field says so - e.g. `-ine` "says /in/ or
  /uhn/", `-ciate / -tiate` "says /sh-ee-ate/"). Drill these for SOUND only: the
  card prompt is "what does it say?", never "what does it mean?". Never use a
  pattern morpheme in a meaning-hunt ASK, a derivative-meaning question, or a
  submarine-model script line - it has no meaning to route through.

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
- NOTES ARE A VERBATIM TEACHER SCRIPT with numbered CAPS anchors, one thought per
  line, no bullets and no dense multi-question paragraph. This slide is an explicit
  exception to the general eight-line Glance Format limit because the playful reading
  routine and retrieval questions must remain separate and easy to scan:
  1-4. ONE NUMBERED `READ:` LINE PER GROUP - never a single line listing all four
     allocations. The whole-class warm-up read is NOT scripted: it wastes a line and
     teachers skip it anyway. Go straight to the groups. Use categories such as
     birthday season, first-name initial, handedness/month, or favourite activity;
     every student must belong to a group, and include a self-choice fallback where a
     category could be ambiguous. Write them exactly like this, one per line:
     ```
     1. READ: "Summer birthdays, read row 1."
     2. READ: "Autumn birthdays, read row 2."
     3. READ: "Winter birthdays, read row 3."
     4. READ: "Spring birthdays, read row 4."
     ```
  5. `5. READ:` makes everyone read row 5, then states: `Every student reads at least
     two rows in the group round.` The allocation must genuinely make that true.
  6-8. Three `ASK:` beats tied to morphemes on the board. Phrase each as words the
     teacher can read aloud immediately, then include think time, one response routine,
     and `EXPECT:` on the same line. Example: `4. ASK: "Which words use flect or flex,
     meaning bend?" 5 sec, choral response. EXPECT: flexible, deflect.`
  9-12. Up to four quick checks, each on its OWN numbered line with its own answer. Never
     write one `CHORAL` line containing four unrelated questions. Example:
     `5. ASK: "What is the opposite of include?" 3 sec, choral response. EXPECT:
     exclude.` Draw from: parts
     of speech, synonyms/antonyms, homophones, plurals/transforms, analogies,
     categories, words associated with..., which word might you hear at/in...
     Once or twice a week, run this beat as Yoshimoto's EXTENSION WHEEL game instead
     (`OG Mid Training Resources/Words to Read-Extension Wheel.pdf`): a named student
     picks a category for points - Rhyming (2), Meaning or use in a sentence (2),
     Synonym (3), Antonym (3), name all the Nouns/Verbs/Adjectives in a row (4),
     Analogy (5) - answers still scripted with EXPECT, next student picks, keep score.
- Every EXPECT answer must be checked against the actual word list on the slide - a
  scripted answer that is not on the board is a build-stopping content bug.
- THIS SEGMENT IS REVIEW, SO EVERY QUESTION MUST BE ANSWERABLE FROM WHAT HAS ALREADY
  BEEN TAUGHT. It is not the place to introduce a new meaning, a new morpheme, a new
  word-part split or a new label. If answering needs anything the group has not met,
  the question is a defect - rewrite it around the taught morphemes, the taught
  meanings, and everyday vocabulary. The taught-only rule below is the hard version of
  this for metalanguage; this line covers content as well.
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
  - EXCEPTION - inside the grammar lesson itself: when the user's chosen grammar
    focus requires a building-block term (prepositional phrases need "noun"), the
    grammar slides may use it, but the I do glosses it in passing on first use
    ("the noun after it - the naming word"). The gate still applies everywhere
    OUTSIDE the grammar section.
  - Same gate applies to the Extension Wheel's Nouns/Verbs/Adjectives category - skip
    or substitute that sector until parts of speech have been taught.

## 2c. Sound bank (9 boxes)

- The 9 morphemes students will need for TODAY'S spelling work (review + new spelling
  words). Students copy them into the top of their page before spelling.
- REVIEW DAYS EXCLUDE THE FOCUS MORPHEME. On a `review` day the sound bank must NOT
  contain the day's focus morpheme (the one the Review Morphology section revisits
  from the previous session). Day 2 shifts the weight to auditory/spelling work -
  students spell that morpheme from memory, and handing it to them in the bank defeats
  the retrieval. Give its slot to another review prefix/suffix/root instead: prefer a
  morpheme one of the day's review spelling words actually uses that is missing from
  the bank; otherwise the most recently taught review morpheme of the same type, so
  the type grouping stays intact. On `new` days the new morpheme DOES belong in the
  bank - it is being taught that day, not retrieved.
- Box colour is set by type automatically: green root, yellow prefix, red suffix.
  Give the spec exactly 9 entries. Any other mapping is a build-stopping defect.
- The slide's teacher note states the key as three plain lines and nothing else -
  `green = root`, `yellow = prefix`, `red = suffix`. Do not restore wording like
  "box colour = card type"; the teacher needs the key, not a description of it.
- LABELS MUST RENDER WHOLE. A 2.33" x 1.04" box cannot hold a variant label such as
  `clud(e)/clus(e)` on one line, and left alone it breaks mid-bracket
  (`clud(e)/clus(` then `e)`). The builder now splits any variant label after the
  slash onto a second line at about 28 pt, so it reads `clud(e)/` above `clus(e)`.
  Never work around this by inventing a shortened label - the label must stay
  verbatim from the catalogue.
- GROUP BY TYPE: order the 9 entries so each type sits together as a row (or column) -
  roots together, prefixes together, suffixes together. A 3/3/3 split filling row 1 /
  row 2 / row 3 is the ideal; when counts are uneven, keep each type contiguous.
- CATALOGUED LABELS ONLY: every sound bank entry (and review card) must use a morpheme
  label exactly as it appears in the photographed catalogue (or legacy bank when the
  card has not been photographed) - that label matches
  the physical card students drill with. If a spelling word needs a variant ending,
  roll it up to the taught card (conclusion needs `-tion / -sion`, never a made-up
  `-ion` card). The builder warns on any unbanked label - treat that warning as a
  content bug.

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
  1. `Sentence:` - context sentence, and it must do TWO jobs.
     - CONNECT TO THE STUDENTS. Put the word in a Grade 5/6 world they actually live
       in: their school, their yard, their sport, their science unit, their city.
       `Our school celebrates literacy week every August.` is the standard. A neutral
       dictionary frame (`The report questioned the adequacy of the bridge.`) teaches
       nothing about the word and is the thing the school keeps flagging.
     - DISAMBIGUATE BY EAR. The sentence is HEARD, not seen. Whenever the word has a
       homophone or a near-twin (intranet/internet, past/passed, their/there,
       ordinance/ordnance), the sentence must make the intended word unmistakable
       from the surrounding meaning - WITHOUT defining it. Never rely on the teacher
       spelling it out.
  2. `After checking:` - one morpheme-aimed question the teacher asks once the
     word is revealed (`Which part of the word tells you the travel is between states?`).
  3. `Answer:` - the answer to that question, in the same 10-12yo language the
     teacher can read verbatim (`inter- means between: travel between states.`).
     Teachers do not memorise thirty morpheme breakdowns - a prompt without its answer
     is a defect, not a nice-to-have.
- READ-ALOUD TEST (apply to every `After checking:` / `Answer:` pair). These are read
  VERBATIM by teachers who may not know the morphology. Say the pair out loud. If a
  competent adult would sound stilted or foolish reading it, rewrite it.
  - Do NOT feed the answer into the question and then repeat it back. Defect:
    `After checking: -ule means small, so what is a module?` /
    `Answer: -ule means small: a module is a small separate part of a bigger whole.`
    The teacher has just said the answer, then says it again.
  - Ask the question the way a teacher actually asks it, then answer it in one natural
    sentence: `After checking: Which part of module tells you it is only a small
    piece?` / `Answer: The -ule ending. A module is one small part of a much bigger
    course.`
  - No classroom-management filler either (`Alright boys and girls, who can tell
    me...`). Scripted, natural, and immediately sayable - that is the bar.
  - Keep the whole set consistent: every word on the day uses the same voice.
- The builder centres every word at the same position with reveal-on-click and a blank
  screen between words (one slide per word) - this fixes the alignment/reveal feedback
  from the team email; you only supply words, sentences and prompts.

## 2e. Responsive pacing inside the locked sequence

- The sequence and required item counts do not change. Responsive pacing means changing
  the time, correction and amount of rehearsal within a block, not deleting the block or
  inventing a different OG routine.
- Pre-cue the response, give think time, name one routine and collect the complete response.
  Choral response means every voice. Boards up means every board. A cold call comes after
  the whole group has thought.
- Routines run on the school-standard cue scripts shared with the general lesson builder:
  boards "Write it... Chin it... Show me.", non-verbal signals "voices off", choral
  "Everyone, together, on three". The first use in a deck scripts the full cue; later beats
  may shorten it. When call-outs replace a signal, the reset is one calm scripted line,
  then the response is re-collected.
- Hands up is never the sampling method in any OG segment. After a complete choral or board
  response, an occasional targeted cold-call follow-up ("How do you know?" / "Convince us")
  deepens retrieval without breaking pace - a few times per session, never on every card.
- Do not treat a few confident voices as evidence that the group is secure. If responses are
  incomplete, reset the routine and collect them again.
- Move briskly through secure review only after the complete response shows accuracy and
  fluency. If a common error appears, slow down for the smallest Yoshimoto-aligned correction,
  then require the student or group to redo the corrected response immediately.
- Do not replace the canonical keyword, meaning, visual card drill, sound bank or spelling
  routine with a generic reteach model. Responsiveness stays inside the prescriptive method.
- Every scripted decision point states what the teacher scans or listens for, the secure move,
  the correction move and the fresh re-check. This is evidence-led pacing, not slide-led pacing.

---

# 3. NEW / REVIEW MORPHOLOGY SECTION

## 3a. The morpheme card

- One new morpheme per `new` day. The card slide background matches the locked colour
  key: green root, yellow prefix, red suffix. The builder keeps the correct type icon
  while normalising the background, so an older master colour cannot reverse the key.
- Notes: `Type: morph / Keyword: X / Meaning: Y`, one value per line - keyword and
  meaning from the canonical card source, verbatim, every time. Sound-pattern cards use
  separate `Sound:` and component-meaning lines instead of cramming both jobs into a
  misleading `Meaning:` sentence. This is the anchor slide; never omit it.
- The card is followed by a copy-into-your-book reference slide showing
  `Type / Keyword / Meaning`. It renders at about 35 pt so it carries the room; do
  not shrink it back.
- THE MEMORY HOOK BELONGS ON THAT COPY SLIDE TOO. The builder lifts the `Memory hook:`
  line out of `wtr_new_notes` and appends it to the copy-card notes as a bold
  `MEMORY HOOK:` line, so the teacher delivers it while the anchor card is on screen
  and again at the end of the grid. Author the hook once, in `wtr_new_notes`.
- WHY LINE ON DAY 1: when a morpheme is first taught, the card/copy-slide notes carry
  one SAY line naming what the morpheme BUYS the student - `Learn micro once and you
  can read and spell a whole family of science words without sounding them out.` One
  line, student-felt payoff (reading big words, spelling word families, guessing an
  unseen word's meaning), same standard as the grammar WHY-AND-HOW rule in section 7.
  Never a curriculum reason, never repeated on the review day - the review day's
  payoff is the retrieval itself.

## 3b. Words to Read (grid, reveal one word per click)

- GRID SIZE MUST BE 9 OR 12 - divisible by 3. The template's three columns centre
  independently, so 8, 10 or 11 words leave a short third column floating out of
  alignment at a different size. "Aim for 12, floor about 8" is about morpheme
  integrity, not layout; the layout constraint is absolute.
- MORPHEME STACKING IS THE FIRST SELECTION RULE. Where the family gives you a choice,
  ALWAYS prefer the derivative that also carries a previously taught prefix, root or
  suffix, weighted to the most recently taught. `microscopic` beats `microchip` once
  `-ic` is on the cards; `malnutrition` beats `malaise` once `-tion` is taught. Two
  taught morphemes in one word is worth more than a slightly neater single-morpheme
  word: it is retrieval practice and it teaches students that words are built, not
  memorised. This rule applies to the new grid, the new spelling four, and every
  review reading and spelling list.
- Layout is automatic and ALWAYS the template's 3-column grid (never reshape it to two
  columns - the school wants the 3-column look preserved). The builder widens all three
  columns, removes hidden text insets and centres the block; 10-12 words also stretches
  the columns down the page for a fourth row. Every term must remain unbroken at 27 pt
  or larger. Reveal stays one word per click.
- NOTES ARE A VERBATIM TEACHER SCRIPT (`wtr_new_notes`), not a glossary.
  Non-negotiable line formula, the SUBMARINE MODEL: sub means beneath + marine means
  water, so a submarine is a vessel beneath the water. Every line does exactly that:
  name EVERY morpheme in the word with its meaning joined by `+`, then a colon, then
  the fused whole-word meaning in 10-12 year old language:
  `disorder - dis- means apart + ord means order: the order pulled apart. A mess!`
  `ordinary - ordin means order + -ary means relating to: relating to the usual
  order of things. Normal.`
  `insubordinate - in- means not + sub- means under + ordin means order: refusing
  to sit under your place in the order.`
  For the bare root itself: `order - ord is the root on its own, order: things
  arranged in their places, one after another.`
  THE FUSED MEANING MUST MIRROR THE PARTS YOU JUST NAMED. If the line says
  `micro means small + bio means life + -ology means the study of`, the meaning after
  the colon is `the study of small life` - assembled from those exact words. Do not
  name the parts and then paraphrase past them (`the study of life too small to see`
  is a dictionary gloss; the student cannot see how the parts made it). Many of these
  words are new to the class, so the whole point is that they hear the parts click
  together. Only depart from the literal fusion when the modern meaning genuinely is
  not the sum of its parts, and then say so in the same line.
  Two further defects to avoid: a line that defines the word without naming its parts
  (`disorder - a mess`), and a circular line that only restates the root
  (`order - ORD means order: the way things are arranged` explains nothing - the
  parts must be seen DOING something to each other). Include the suffixes too (-ary,
  -al, -ate, -ance, -ation, -ly each get named with a kid-friendly meaning). Both
  morpheme spellings (ord AND ordin) must each appear across the script when the
  morpheme has variants.
- End the script with a `Memory hook:` line - one sticky retention device students
  can rehearse and reuse in their writing: a hearable slogan (`hear ORD, think "Off to
  my place in the line"`), an action, a family-collection challenge, or a
  build-the-longest-word game. Every new/review morphology session has one; reuse the
  same hook all week so it sticks. The builder also surfaces it on the copy-card slide
  (3a).
- HOOK QUALITY BAR: a hook that only restates the meaning in a slogan voice
  (`hear MICRO, think 'shrink it down' - if micro is at the front, the thing is TINY`)
  is not a memory device, it is the definition wearing a hat. A good hook gives the
  student something to DO or something concrete to picture that survives the week:
  an image that is hard to unsee (`micro is the microscope lens - everything through
  it is too small for your eyes`), a physical action, a countable challenge (`collect
  five micro words from the science room by Friday`), or a rule of thumb they can
  apply to an unseen word. Test it: could a student use this hook next term on a word
  they have never met? If not, rewrite it.
- Grade 5/6 register: choose words that stretch 10-12 year olds (insubordinate,
  ordinance), not infant-level fillers. Same for the example sentences.
- On `review` days: same morpheme as the previous session, mostly NEW derivatives
  (morphological family extensions: coordination, subordination, preordain), with at
  most 2-3 repeats from the previous day.

## 3c. Words to Spell (4, revealed per click)

- Exactly 4, and all 4 MUST come from today's Words to Read grid (this is the one place
  read/spell overlap is required - students spell what they just decoded).
- Apply the morpheme-stacking rule from 3b: of the eligible grid words, choose the four
  that also carry recently taught prefixes/suffixes wherever the family allows it.
- Rotate the early-finisher extension in the notes; never the same one two days running.
  Menu: use two target words in one sentence with a subordinate clause; write the meaning
  in your own words using the root; find and record an antonym; turn the word into a
  question; add a prefix/suffix and explain the new meaning.
- Extensions must be SELF-CONTAINED: if the task names a concept (subordinating clause,
  adverbial phrase), include a one-line reminder of what it is right in the note, so no
  student stalls waiting for the teacher. Then add an `EXTENSION:` step meeting the
  standard in 3e.

## 3d. New Morphology You Do activity (required on every new/review day)

After Words to Spell and before Review Learned Words, every new/review session gets a
dedicated Yoshimoto-style You Do: a magenta section header (auto-retitled `New
  Morphology - You Do` / `Review Morphology - You Do`) followed by a DESIGNED task slide
  (same structured layout as grammar, but in the New Morphology magenta: rule banner,
  hero task card or up to four short prompt lines, routine chip, footer). The generic "write a
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
- ACHIEVABILITY BAR (non-negotiable - the school has rejected You Dos that failed it):
  - The task description on the slide is EXPLICIT: it says exactly what to write and
    where, in words a mixed-readiness 10-12yo can follow with no teacher explanation.
    A terse label (`Write R or P`) is not a task description.
  - The slide carries ONE COMPLETED WORKED EXAMPLE of the very task type, marked
    `One done for you: ...`, before students start. Use the hero card for it. A You Do
    whose first exposure to the task shape is the task itself is a defect.
  - PUT THE WORKED EQUATION ON ITS OWN LINE. Write the hero as
    `One done for you:\nsmall + sound = microphone` - the label on line 1, the
    equation alone on line 2. The builder emits each authored line as its own centred
    paragraph, so an equation reads as an equation instead of trailing off the end of
    a sentence. The same applies to any `word = label` model on a grammar slide.
  - The numbered prompt lines under the hero are centre-aligned by the builder so the
    block reads as a set. Keep each line short enough not to wrap (see 10c).
  - ONE CLEAR ANSWER PER ITEM. Every item must have a single correct answer a student
    can reach from that day's teaching and verify on the check slide. An item with two
    defensible answers (`pottery = P or R/P`) or a judgement call about etymological
    transparency is teacher-led We Do discussion material, never independent work.
  - Prefer CONCRETE BUILD/USE tasks (build the word from a meaning sum, write the sum
    beside a derivative, show the meaning in a sentence) over abstract classification
    or meaning-audit tasks. If a sort is used, the two categories must be plainly
    distinguishable by a 10-12yo using only the printed card meanings.
- Notes are Glance Format You Do style: numbered beats with SAY/TIME/CIRCULATE/SCAN,
  ANSWER first when the task has answers, TRAP with Fix, EXTENSION and HELP, then `---`
  and the prep zone.
- FULL `ANSWERS:` BLOCK AT THE VERY BOTTOM (required on every fixed-answer You Do).
  Below the `---` divider, after the purpose line, write `ANSWERS:` and then one line
  per item restating the QUESTION and its ANSWER, plus the extension answer:
  ```
  ---
  Purpose: ... [You Do | Morpheme sums | Morphology]
  ANSWERS:
  1. small + look at = microscope
  2. small + wave = microwave
  3. small + chip = microchip
  4. small + life + study of = microbiology
  EXTENSION: e.g. microbiologist = micro + bio + -ology + -ist, a person who studies
  tiny life.
  ```
  The one-line `ANSWER:` at the top of the live zone stays - that is the glance fact.
  This block is the fallback for a teacher who gets stuck mid-circulation, so it
  repeats the question rather than assuming they can see the board. The prep zone's
  usual three-line limit does not apply to this block.
- PROMPT LADDER (required): the CIRCULATE beat scripts what the teacher says to a
  stuck student. It must ADD something the board does not already say. Pointing back
  at the worked example and re-reading the rule is what the teacher has just finished
  doing out loud - a student still stuck after that needs a different way in, not a
  louder repeat. Give the ONE piece of knowledge that unlocks the most students:
  the meaning of the unfamiliar second morpheme, a concrete object to picture, a
  spoken version of the sum, or a question that narrows the field
  (`CIRCULATE: to a stuck student say, "scope is the part that means look at - so what
  do you look through in science?"`). Then a second step that starts item one
  together. HELP still changes the task form (word bank, match instead of build,
  first item done). A bare `CIRCULATE: check work`, or one that only restates the
  slide, is a defect.
- The check slide's `ASK:` must be a genuine thinking question, not a spot-the-obvious
  (`Which part of every answer means small?` when every answer visibly starts with
  micro). Ask something that makes them USE the pattern: which of these two words
  changes meaning if you drop the prefix, why does this word need three parts, which
  answer would still make sense with a different root.
- A fixed-answer activity MUST be followed immediately by a duplicate answer slide
  titled `Tick it or fix it - ...`. Supply `check_items` (the completed prompt/answer
  lines) plus `check_notes`. Never use the old `Check and fix` title. The first slide
  protects independent thinking; the next slide shows answers in green and tells
  students to tick correct work and fix errors. The builder blocks a
  fixed-answer activity whose notes start with `ANSWER:` but which lacks `check_items`.
- THE ANSWER SLIDE MUST TEACH, not just display. Bare labels (`winery = P`) reteach
  nothing. Each answer carries its one-line reasoning - the morpheme sum or the rule
  applied - either inside the item or as its own short green line directly beneath
  (`winery = P` / `win + -ery = the place where wine is made`). `check_notes` open by
  MODELLING the reasoning for the first item aloud before students tick, so a student
  who got it wrong hears HOW the answer is built, not just WHAT it was.

## 3e. EXTENSION (the word is EXTENSION, never STRETCH)

Applies to every early-finisher note, every morphology You Do and every grammar You do.

- The label is `EXTENSION:`. `STRETCH:` is retired everywhere in OG decks; the builder
  bolds and line-breaks both, but new specs must use `EXTENSION:`.
- AIM ABOUT 18 MONTHS AHEAD: end of Year 7 / start of Year 8. This cohort is an
  enrichment group and the old "one notch harder" extensions were finished in seconds.
- IT MUST BE THE SAME CONCEPT, PUSHED - not a different task. The test: a student
  doing the extension is still doing today's learning, just at a level that would
  normally sit a year and a half later. If you have to explain a new idea to start
  it, it is a new task, not an extension.
  - Model to hold in mind: if the core task is a Year 1 class counting 1 to 20, the
    extension is counting backwards from 50 - same concept, bigger number, reversed.
  - Morphology examples: core builds `microscope` from a two-part sum, extension
    builds a four-morpheme word and writes the full sum
    (`microbiologist = micro + bio + -ology + -ist`); core writes the sum beside a
    derivative, extension takes an UNSEEN word from the same family and predicts its
    meaning from the parts alone, then checks it.
  - Grammar examples: core labels phrase or clause, extension combines a subordinate
    clause and a main clause into one correctly punctuated sentence and then moves the
    subordinate clause to the end and repunctuates it.
- It must be STARTABLE WITHOUT THE TEACHER, and its answer goes in the `ANSWERS:`
  block so a teacher can confirm it.

---

# 4. THE PHOTOGRAPHED CARD CATALOGUE (canon workflow)

## 4a. Mandatory files and authority order

Every new task starts without relying on conversation memory. Before choosing a
morpheme, keyword, meaning, part of speech, or derivative, read the relevant category
file:

- `og_planner/yoshimoto_cards_suffixes.json` - 89 suffix cards.
- `og_planner/yoshimoto_cards_prefixes.json` - 104 prefix cards.
- `og_planner/yoshimoto_cards_latin_roots.json` - 84 Latin-root cards.

These are direct transcriptions of the teacher's photographed 2007 Yoshimoto set. The
three category JSON files are the source of truth. `yoshimoto_cards_master.json` and
`YOSHIMOTO_CARD_CATALOGUE.md` are generated search/readability views; regenerate them
from the category files rather than editing them as an independent authority.

Resolve every card in this strict order:

1. Match the category/type and the exact physical-card heading in the photographed
   catalogue. For that card, use its `meaning` and `keyword` exactly as stored. Preserve
   `part_of_speech` exactly where present; do not invent one where the card prints none.
2. If - and only if - the photographed catalogues do not contain the card, look in
   `og_planner/morpheme_bank.json` for an older teacher-confirmed entry.
3. If only `og_planner/morpheme_meanings.json` has the entry, treat its wording as an
   unconfirmed reference, not an exact Yoshimoto-card match. Surface that status and
   verify it from the OG source scans or with the teacher before building.
4. Do not use general web definitions to overwrite, blend with, or 'correct' a captured
   card. Online dictionaries may clarify a derivative word, but the physical card is
   authoritative for its own heading and meaning.
5. If a morpheme is not in the meaning catalogue, consult the `OG/` reference library,
   navigated via
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
   Append the meaning with its source, then use it. The Latin Scrolls derivative lists
   are also the first place to look when building a Words to Read family (3b).
6. NEVER change an existing verified meaning casually. A user-supplied correction is
   not a one-off spec override: check it against the photographed card, update the
   relevant category JSON first, regenerate the combined views, record the evidence,
   and then use the corrected canon consistently.

## 4b. Locked fields, collisions, and editor-supplied keywords

- The catalogue already contains the selected teaching keyword. Do not choose the
  first associated word again, substitute a preferred keyword, or 'improve' it from an
  online list. A different keyword/meaning in a week spec is a content error.
- Some physical cards share a heading but not a meaning. Keep them as separate cards.
  In particular, `di-` means either `away, apart, not` or `two`; every `di-` week spec
  must include the intended photographed meaning so resolution is unambiguous. Never
  resolve a collision by heading alone.
- A small number of cards print no associated examples. Their catalogue keyword is an
  explicitly documented, editor-supplied school-friendly anchor (currently `-s` ->
  `cats`, `octo-` -> `octopus`, and `hepta-` -> `heptagon`). Use that catalogue keyword
  consistently, retain the provenance note, and do not misrepresent it as printed text.
- `part_of_speech` records only what is printed on the physical suffix card. Add it to
  teacher-facing card metadata where available. It does not authorise an untaught
  parts-of-speech question; the taught-only gate in section 2b still applies.

## 4c. Associated words and exclusions

- For a captured card, `associated_words` is the default authorised family for new
  Words to Read, new Words to Spell, review pools, extra tasks, and morphology
  examples. Choose for age, decodability, transparency, and lesson purpose from within
  that family; do not silently replace it with a web-generated list.
- `excluded_words` preserves the physical-card provenance but is a hard block on
  automatic or incidental student-facing selection. Do not restore an excluded word
  merely because it is common, appears online, or fits a word count. Use one only when
  the user explicitly requests that exact word and the context is school-appropriate;
  record the reason in the handoff.
- A new morphology spelling word must be a subset of that session's new morphology
  reading grid. Same-day review reading and review spelling lists must remain disjoint.
- An outside derivative is exceptional. Use one only when the captured card has too few
  suitable associated words or the user explicitly requires it. Verify that the
  morpheme is visibly present and contributes the taught meaning, verify the whole-word
  meaning and Australian spelling from an appropriate source, and record it in the
  session's `associated_word_exceptions` with `word`, `reason`, and `source`; repeat it
  in the handoff. Never use an outside derivative merely to avoid reducing a genuinely
  niche family below 12 words.
- Membership in a printed family does not guarantee a transparent modern decomposition.
  Never force a false morpheme sum. If a listed word is etymologically opaque or the
  taught meaning does not operate clearly in it, omit it or teach it only as a
  whole-word/spelling-family example with an explicit teacher note.
- Australian English is compulsory throughout. Preserve catalogue spellings such as
  `haemophiliac` and normalise any verified outside derivative to Australian usage.
  Do not change a catalogue spelling back to US English.

## 4d. Required pre-build card audit

Before writing or building a week spec, verify mechanically:

1. Every captured card matches the correct category, exact heading, meaning, keyword,
   and printed part of speech in the relevant category JSON.
2. Every new-family word is in that card's `associated_words`, or is explicitly logged
   as a verified exception under section 4c.
3. No selected word occurs in any relevant `excluded_words` list.
4. Every selected derivative visibly contains the intended morpheme and uses its taught
   meaning; pattern cards are handled as sound patterns, not meaning morphemes.
5. Every new spelling word occurs in that day's new reading grid, and no review spelling
   word duplicates that day's review reading list.
6. Ambiguous headings such as `di-` include their intended meaning in the spec.

Any mismatch is a blocker, not an advisory note. Fix the source/spec before generating
slides.

---

# 5. LEARNED WORDS (red words)

- Every session REVIEWS the two most recently taught learned words. New learned words
  are dealt on a FIXED SCHEDULE, not one per day until they run out.
- SCHEDULE (the second new word is always taught in SESSION 3, never back to back with
  the first):

  | Sessions in the week | New learned word 1 | New learned word 2 |
  |---|---|---|
  | 5 | session 1 | session 3 |
  | 4 | session 1 | session 3 |
  | 3 | session 1 | session 3 |
  | 2 | session 1 | session 2 |

  Every other session is review-only. Two new red words on consecutive days is the
  defect this replaces - it gives the first word no spacing before the second lands.
  With three or more new words in a week (rare), keep the same principle: never two
  new words on consecutive sessions.
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
- On the day a NEW learned word lands, add one line of payoff naming where the student
  will actually write it (`You will write this word in every science report this
  term.`) - the section 7 WHY standard: a reason the writer feels, because these are
  exactly the words that stall a draft when they have to be guessed.
- If a learned word happens to CONTAIN a morpheme the group has been taught, say so in
  the notes and use it: `seismograph ends with graph, meaning write or record - the
  same graph as pictograph. Only the seis part is unfair.` It shrinks the amount that
  has to be memorised and reinforces the morphology. Do not force it where the part is
  not genuinely the taught morpheme.
- The builder replaces the template's old dense section-note blocks with concise,
  line-separated procedure notes. Never restore the inherited paragraph walls or
  control characters from the master notes.

---

# 6. DICTATION (two sentences, both with reveal)

- MEASURE LENGTH IN CHARACTERS, NOT WORDS. A word count lies: fourteen words made of
  camouflage, seismograph and microprocessor is far harder to hold and write back than
  fourteen short ones, and the group has struggled on exactly that. Count characters
  including spaces and punctuation.
  - Sentence 1 EASIER: **55-75 characters**.
  - Sentence 2 HARDER: **80-100 characters**, richer punctuation. Hard ceiling 105.
  - The reference is the model harder sentence at 92 characters:
    `Before the class began, Mia whispered, "That microchip is smaller than my
    camouflage badge."`
  - Word count is now only a sanity check (roughly 10-12 then 14-16); if the character
    count and the word count disagree, the CHARACTER count wins and you cut words.
- Content sourcing, in priority order:
  1. Today's new learned word and/or this week's learned words.
  2. Today's (or yesterday's) morphology words.
  3. The 2b weighting recipe for older morphemes - dictation IS revision.
  Every content word must be spellable from taught material; scope-check phonograms.
- EVERY DICTATION SENTENCE MUST MAKE SENSE. Two checks, both mandatory, run on every
  sentence before shipping:
  1. REAL WORDS ONLY: every word must be a real English word, spelled exactly. Never
     invent, blend, or misremember a word to force a morpheme in - the school rejected
     `"I will depict the deisomograph."` (not a word; the instrument is a seismograph).
     If a target will not fit, choose a different derivative or a different target.
  2. THE SCENE MUST BE PLAUSIBLE: read the sentence aloud - a 10-12yo must be able to
     picture who is speaking, where, and why in one go. Stringing targets together
     grammatically is not enough (`Although the camouflage worked, our artist said,
     "I will depict the seismograph."` parses but describes no scene anyone can
     picture). If two targets cannot share one sensible moment, split them across the
     two sentences - the targets serve the sentence, never the reverse.
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
  - colours every required CAPITAL LETTER GREEN in the revealed sentence (C),
  - colours all punctuation RED in the revealed sentence (P),
  - keeps spelling targets bold + underlined (S),
  - prints `Score: ___ /N` on the slide, where N = capitals + punctuation marks +
    targets + 1 for Understanding, so students self-score against a visible total,
  - writes a tick-off CUPS checklist in the notes with each category's items and count.
  Sanity-check N against the sentence before shipping - every red mark, underline and
  capital on the slide must be counted once.
- CUPS IS THE BRIDGE TO THEIR OWN WRITING. Once or twice a week (not every session),
  one scripted line at the check step tells students the transfer: `CUPS is not just
  for dictation - run the same four checks on your own draft before anyone else reads
  it.` The section 7 HOW standard: dictation is where the class rehearses the editing
  routine they will use on everything they write.
- NAME EACH PUNCTUATION MARK IN LANGUAGE THE TEACHER AND THE STUDENTS BOTH HAVE.
  `comma after the reporting clause` is unusable: a ten-year veteran cannot explain it
  if a student asks, so they skip it and the mark is never taught. Every entry in
  `punctuation` carries a plain-English reminder in brackets, and the reason the mark
  is there:
  - `comma after the opening subordinate clause (the because/when/although part that
    cannot stand alone - the comma shows where it ends)`
  - `comma before the speech (it separates who said it from what they said)`
  - `opening quotation mark (speech starts here, so the first word takes a capital)`
  - `colon (it introduces the list or the words that follow)`
  - `semicolon (it joins two complete sentences that belong together)`
  - `apostrophe for the missing letter in "didn't"`
  Assume NEITHER the teacher nor the student knows the grammatical label. If a mark
  cannot be explained in one bracketed clause, choose a different sentence.
- ALWAYS state the capital-after-speech-marks rule when the sentence quotes anyone -
  it is the most-missed capital in this cohort's dictation.
- SPEECH MARKS: write straight ASCII quotes in the spec as always. The builder
  converts each straight pair to directional marks on the dictation slide face only
  (`"..."` renders as one opening and one closing curly quote), because the locked
  Lexend sentence box draws the ASCII quote glyph slanted like a CLOSING quote, so
  opening quotes looked backwards on screen - school feedback. Quotes must come in
  pairs; an unpaired quote is a spec defect (the builder warns and leaves it straight).
- COLON vs COMMA vs SEMICOLON, decided once and applied consistently:
  - Speech introduced by a reporting clause takes a **comma**:
    `Mia whispered, "That microchip is tiny."` Never a colon.
  - A colon introduces a list, an explanation, or an instruction that is not reported
    speech: `The captain gave one order: keep every record.`
  - A semicolon joins two complete, closely related sentences with no conjunction.
  Earlier weeks put a colon before reported speech and students copied it; that is the
  error being corrected. When a deck uses a colon or semicolon, the `punctuation`
  entry must say which of these jobs it is doing.
- METER: sentence 1 uses the green gauge (needle near the start), sentence 2 the yellow
  one (needle about two thirds around, everything before it filled). The builder now
  ships the master's real yellow artwork, so the needle genuinely moves - if a harder
  dictation renders with the needle still at the start, that is a build defect, not a
  design choice.
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
- WHY-AND-HOW RULE (non-negotiable, school feedback): labelling is never the point -
  writing is. Every grammar day must answer the two questions a student could fairly
  ask: WHY are we learning this, and HOW do I use it in my own writing?
  - WHY, once per day, in the I do: one line naming the writer's payoff in 10-12yo
    terms, delivered with the rule restatement (fold it into beat 1's SAY, or give it
    its own beat when the live zone has room). It must be a reason a young writer
    FEELS, not a curriculum reason: `Every sentence you write must stand alone - this
    test is how you catch a half-finished sentence before your reader does.` is the
    standard; `Today we are learning about clauses.` and `So you can label clauses.`
    are defects - the label is the tool, never the payoff.
  - HOW, once per day, at the close (the You do or its Tick it or fix it slide): one
    scripted transfer line that sends the concept into the students' own writing -
    when to reach for it and what to do: `When you edit your next draft, read each
    sentence alone; if it leaves you waiting, it needs its verb.` Name writing the
    students actually do (their narrative, their science report, their dictation
    sentences), never writing in the abstract. The footer/check_footer slot is ideal
    on the slide face (`Editing move: read it alone before you hand it in.`).
  - The WEEK states its payoff once: the grammar `header_notes` (or the day-1 purpose
    line) names the writing problem this week's focus solves - fragments, run-ons,
    every sentence opening the same way, speech the reader cannot follow - and each
    day's WHY line is a variation of that payoff, not a new claim each day.
  - BUDGET-NEUTRAL: WHY and HOW are one line each inside existing beats and slots -
    never a new activity, never an extra slide, never a note-budget blowout. If the
    live zone is full, the WHY rides inside beat 1's SAY and the HOW inside the final
    SAY or the footer.
- DESIGNED SLIDES, NOT TEXT DUMPS. Grammar blocks use the structured spec and the
  builder lays them out (this is the one OG section with visual design): a purple rule
  banner across the top, a big soft-tinted HERO CARD carrying the example or task, up
  to two short support lines, a purple routine chip bottom-left, and an italic footer
  bottom-right. Author to that shape:
  - `rule` - the restated rule, one breath (the absent-student line).
  - `example` - ONE hero: the worked sentence, the judgement pair, or the task itself.
    This is what students look at; make it concrete, never meta. EACH `word = label`
    EQUATION GETS ITS OWN LINE - author `under the old bridge = phrase (no verb)` and
    `The boat sank = clause (verb: sank)` as two authored lines, not one run-on
    sentence. The builder renders each authored line as its own centred paragraph and
    sizes the block to the longest line, so every equation stays whole on one line.
    An `A) ... B) ...` judgement pair must fit ONE line (about 36 characters at 30 pt)
    or it breaks mid-item; if it will not fit, shorten both halves.
  - `items` - max 2 short support lines (the test, the challenge, the if-stuck starter).
    Centre-aligned by the builder.
  - `routine` - the all-student response move as a chip. GRAMMAR IS DONE IN THE OG
    BOOK: `OG books - 90 sec - pens down`, `Thumbs verdicts - voices off`,
    `OG books - 2 min - books up`. See the whiteboard rule below.
  - `footer` - tease or stem, and it must MEAN something to a student reading only the
    screen. `Name the verb, or say there is none` is instruction-shaped noise - the
    student does not know whether that is the task, a hint or a check. Write
    `Tomorrow: colon or no colon?`, `Stem: The colon belongs there because ...`, or
    `Check: did you find a doing word?` - a tease, a sentence stem, or a single
    self-check phrased as a question.
  Never fall back to the legacy `lines` array for new decks - it produces the bland
  text-only slides the school rejected.
- WHITEBOARDS ARE FOR THE MORPHOLOGY YOU DO ONLY. Every grammar block - I do, We do and
  You do - is worked in the student's OG book, so the evidence stays in the book the
  teacher already marks. Do not write `whiteboards`, `boards up` or `Write it... Chin
  it... Show me.` into a grammar routine chip or a grammar note. The non-written
  all-student moves (thumbs verdicts, choral, beat-the-teacher) are still available and
  still preferred for the We do.
- Engagement is mandatory, not optional. Rotate all-student-response moves: thumbs
  verdicts, choral answers, beat-the-teacher (teacher makes the error, students catch
  it), partner swap-and-check, books-up exit tickets. Never volunteer-hands-only.
  Never a worksheet.
- You do produces visible evidence (books up, exit ticket) the teacher can scan to
  regroup tomorrow.
- The You do obeys the same achievability bar as the morphology You Do (3d):
  - The hero card shows a COMPLETED WORKED EXAMPLE of the task (`One done for you:
    The keys are under the blue mat. Whole phrase: under the blue mat`), never the
    bare task instruction. Put the instruction in the rule banner or the first item
    (`Your turn: ...`), then give the material items.
  - The notes CIRCULATE beat scripts the prompt ladder for stuck students, and it must
    ADD something the board does not already say (3d); HELP changes the task form.
  - Prefer a fixed-answer core (find/complete/fix given sentences) with production as
    the EXTENSION when the group is still consolidating - open production with no model
    is the hardest form of the task, not the default exit ticket.
  - The grammar You do carries an `EXTENSION:` meeting the 3e standard (about 18 months
    ahead, same concept pushed) and an `ANSWERS:` block at the very bottom of the notes
    listing each item's question and answer plus the extension answer.
- A fixed-answer You do MUST be followed by the same green `Tick it or fix it - ...`
  answer slide as the morphology You Do: supply `check_title`, `check_rule`,
  `check_items`, `check_routine`, `check_footer`, `check_notes` on the `you_do` block.
  The answer slide must teach (3d): each answer carries its one-line reasoning, and
  `check_notes` MODEL the reasoning for the first item before students tick.
- Teacher notes use the Glance Format from `MEGA_PROMPT.md` sections 45-47: LIVE ZONE
  (ANSWER first if the slide asks anything; 2-5 numbered beats fusing action + talk with
  CAPS anchors - SAY/ASK/MODEL/SCAN/TIME/CIRCULATE; TRAP with Fix; EXTENSION/HELP), then
  `---`, then a one-line prep zone with the `[Stage | element | focus]` tag, then the
  `ANSWERS:` block on fixed-answer You dos. SAY lines are warm, natural classroom talk,
  one breath (~20 words), never clipped fragments.
- Every new thought starts a new source line. Never place `ANSWER`, multiple numbered
  beats, `TIME`, `CIRCULATE`, `SCAN`, `TRAP`, `EXTENSION` and `HELP` in one paragraph.
  The live zone is eight lines maximum; the prep zone is three lines plus the
  `ANSWERS:` block.
- Content agenda: Victorian Curriculum English level 5/6 language strand - clause types,
  cohesion, punctuation for effect, complex sentences, quoted speech, apostrophes,
  modality - sequenced across the term by the user's grammar focus list.

---

# 8. TEACHER NOTES MAP (iPad-safe source and rendering)

The builder writes clean notes for every slide, including concise replacements for the
master template's old dense procedure blocks. All authored notes must follow
`MEGA_PROMPT.md` sections 45-47.

- SOURCE TEXT IS PLAIN TEXT. Never type markdown (`**`, `_`, backticks, markdown
  headings or markdown bullets) in a notes field. The builder recognises labels such as
  `ANSWER:`, `1. SAY:`, `SCAN:`, `Keyword:` and `Meaning:` and applies real PowerPoint
  bold formatting without leaving markup characters in the notes pane.
- Every new thought is a separate newline-delimited paragraph. Flow notes use numbered
  beats with CAPS anchors. Do not concatenate `ANSWER: ... 1. SAY: ... 2. CIRCULATE:`
  on one source line; that becomes an unreadable iPad wall.
- Every notes paragraph is left-aligned. Do not inherit centred or indented paragraph
  settings from the master; wrapped lines must return to the left edge on iPad.
- LIVE ZONE: eight nonblank lines maximum in this order when relevant - ANSWER; 2-5
  numbered beats; TRAP with Fix; STRETCH/HELP. No blank lines inside it. The Words to
  Read Review script is the sole exception and follows section 2b's longer numbered
  reading-plus-retrieval structure.
- PREP ZONE: `---` on its own line, followed by at most three short lines. Anything
  needed while students are working must stay above the divider.
- SAY text is natural, direct teacher language that can be read verbatim. Do not use
  filler such as "Okay kids, next we have..." and do not write clipped production
  notes. Each question includes think time, one response routine, and `EXPECT:` or
  `Answer:` in student-friendly language.
- Response routines are completed. Notes never imply that partial choral participation,
  missing boards or volunteer call-outs are sufficient evidence. The teacher resets the
  routine, checks the complete response and then decides whether to keep the brisk pace or
  give the smallest prescribed correction.
- Data notes use one labelled value per line. For a sound-pattern card, do not cram a
  pseudo-meaning into one line: render `Sound: /sh-ee-ate/` and `Meaning of -ate: to
  make or do` separately.
- The builder rejects raw markdown and overlong note lines. Its legacy line-break
  normaliser is a backstop, not permission to author dense notes.

| Slide | Notes content (verbatim template) |
|---|---|
| Morph review card | `Type:` / `Keyword:` / `Meaning:` (catalogue verbatim) / `Part of speech:` where printed / optional `Extra task:` + `SAY:` + `Possible answers:` |
| Words to Read Review table | one numbered `READ:` line per group (no whole-class warm-up beat), then the everyone-row-5 line stating each student reads at least two rows, then one direct question and its `EXPECT:` answer per numbered line |
| Sound bank | `green = root` / `yellow = prefix` / `red = suffix`, one per line (builder default) |
| Spell the word (each) | `Word to spell:` / `Sentence:` / reveal line / `After checking:` / `Answer:` (all required); builder renders labels bold and underlines the target word where it appears inside the `Sentence:` line |
| New morph card | `Type:` / `Keyword:` / `Meaning:` or separate `Sound:` lines; the copy slide adds `MEMORY HOOK:` automatically |
| Words to Read grid | one scripted line per word + `Memory hook:` last |
| Words to Spell grid | line-separated early-finisher cue + `EXTENSION:` (3e) |
| Learned word (each) | `Why learned:` / `Say it:` AU pronunciation / `Link:` |
| Dictation (each) | sentence + line-separated CUPS checklist + `Score:` + `Focus:` |
| Grammar I/We/You do | Glance Format live zone, `---`, prep line, then `ANSWERS:` on a fixed-answer You do (section 7) |

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
                             "extra_task"?: { "prompt", "answers": [..] } } x10 ],
    "words_to_read_review": { "words": [15 strings], "notes": "activity script, 2b format" },
    "sound_bank": [ { "morph", "type" } x9 ],
    "words_to_spell_review": [ { "word", "sentence", "prompt", "answer" } x10 ],
    "new_morphology": { "morph", "type", "keyword", "meaning" },   // omit on week_review
    "words_to_read_new": [ { "word", "meaning" } x8-12 ],          // omit on week_review
    "associated_word_exceptions"?: [ { "word", "reason", "source" } ], // rare; section 4c
    "wtr_new_notes": "verbatim line-separated teacher script + Memory hook: (3b)",
    "words_to_spell_new": [4 words, all from words_to_read_new],   // omit on week_review
    "extension": "self-contained early-finisher + EXTENSION: (3e)", // omit on week_review
    "new_morph_activity": { "title", "rule", "example"?, "items": [<=4],
                            "routine", "footer", "time"?, "notes",
                            "check_title"?: "Tick it or fix it - ...",
                            "check_rule"?, "check_items": [completed answers],
                            "check_routine"?, "check_footer"?, "check_notes" }, // 3d
    "file_name"?: "override only if the user asks",
    "learned_words": { "review": [ { "word", "unfair", "notes" } x2 ],
                        "new": { "word", "unfair", "notes" } },     // "new" omitted Friday
    "dictation": [ { "meter": "green" | "yellow", "sentence", "targets": [..],
                     "cups": { "capitals": [..], "punctuation": [..] },
                     "focus" } x2 ],
    "grammar": { "header_notes",
                 "i_do":  { "title", "rule", "example", "items": [<=2], "routine", "footer", "notes" },
                 "we_do": { ... },
                 "you_do": { ...,                       // fixed-answer You do: same
                             "check_title"?, "check_rule"?, "check_items"?,  // check
                             "check_routine"?, "check_footer"?, "check_notes"? } } // slide as 3d
  } ]
}
```

Notes:
- Card `keyword`/`meaning` fields must mirror the photographed catalogue when the card
  is present there. The builder enforces this and gives the photographed transcription
  priority over the legacy bank and reference meanings. Any differing spec value
  is a blocker even if it appears only as a WARN. For a genuinely ambiguous heading
  such as `di-`, include the intended photographed meaning in the spec.
- New-family words must pass the section 4c/4d associated-word and exclusion audit.
- The builder automatically appends the physical card's printed part of speech to
  teacher notes when the catalogue supplies it; do not invent a spec value.
- `unfair` must be an exact substring of the word (builder warns if not found).
- Notes fields are plain text. Raw markdown such as `**bold**` is a build-stopping error;
  the builder applies real bold formatting to recognised labels.
- Keep single display words <= ~14 characters where you have a choice; the builder
  shrinks longer ones automatically.
- Straight ASCII only. The builder sanitises smart punctuation as a backstop.

---

# 10. BUILD AND QA (required, in order)

```bash
# EVERY python command in this pipeline uses miniconda `python` (NOT homebrew
# python3): lxml, python-pptx and pymupdf live there. That includes the builder,
# scripts/pptx_to_images.py, and any PDF page rendering.
python og_planner/build_og_week.py og_planner/weeks/<spec>.json          # all sessions
python og_planner/build_og_week.py og_planner/weeks/<spec>.json --only Tuesday
python tests/test_og_builder_regressions.py                               # builder + exemplar regression gate
```

1. The builder's gate must pass (exit 0): file reopens cleanly, no `XYZ` left anywhere,
   notes are separate left-aligned PowerPoint paragraphs with real bold labels, each
   spelling target is underlined inside its `Sentence:` line, the standard live zone
   is no more than eight lines, Words to Read Review includes all three required READ
   beats and guarantees every student at least two rows, review questions are not piled
   into one paragraph, every review/new card and Sound Bank box uses green root / yellow
   prefix / red suffix, fixed answers have an immediate green `Tick it or fix it` slide
   (morphology You Do AND any fixed-answer grammar You do),
   new-word grids are at least 27 pt, grammar examples are at least 20 pt, and the first
   dictation uses the green meter while the second/trickier dictation uses yellow.
   Read every WARN - word-count, catalogue mismatch, excluded-word, unbanked-label and
   overflow warnings are content bugs to fix in the spec, not noise. A NOTE about a card
   absent from the photographed catalogues is advisory only when it clearly records an
   unconfirmed fallback; list every such NOTE in the summary. Captured-card metadata is
   never advisory.
2. Visual QA every session deck: `python scripts/pptx_to_images.py "output/<folder>/<deck>.pptx"`,
   then INSPECT the images: title subtitle on one line; overview table cells complete;
   10 card slides in the jumbled order you specified; 15-word table filled; sound bank
   colours match types; each spell word centred; every new-word grid term is unbroken
   and at least 27 pt; task and check slides form a protected pair; learned-word
   highlight is on the right letters; dictation capitals are green, punctuation red and
   targets bold/underlined; grammar slides are not overflowing. Also inspect the notes
   pane/XML: one thought per paragraph, no raw markdown, no dense paragraph walls and no
   control-character artefacts. Local render substitutes fonts (Luckiest Guy/Lexend may
   look plain) - that is a render artefact, not a bug.
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
   demonstrated there. Your specs must match its formats exactly - same plain source
   labels (the builder applies real bold),
   same beat numbering, same field names. Do not invent alternative formats.
2. Load the three category JSONs from section 4, build the week's morpheme timeline
   (which morpheme is taught which day), resolve each exact card, and run the section
   4d audit. Only cards absent from the photographed catalogues may fall back to
   `og_planner/morpheme_bank.json`. Stop and verify anything unconfirmed BEFORE writing
   the spec.
3. For each session, in slide order, generate: review-10 (jumbled per 2a), 15 review
   words + verbal script (2b), sound bank 9 grouped by type, no focus morpheme on
   review days (2c), 10 spelling words
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
- Any type-coded card background or Sound Bank box that is not green for a root,
  yellow for a prefix, or red for a suffix. Never trust an older master colour. (0, 2c, 3a)
- A review word whose morpheme is not doing its job (mini in minister). (2b)
- Words to Read Review notes as a definition list, or asking students to underline/
  circle/write - this segment is fully verbal, students have nothing in front of
  them. (2b)
- READ-together missing, listed after the questions, or the row splits crammed onto
  one line instead of one split per line. (2b)
- Fun row groups that do not cover every student, omit the everyone-reads row, or fail
  to guarantee that every student reads at least two rows. (2b)
- A grid script line that is circular (`order - ORD means order: the way things are
  arranged`) instead of the submarine model naming every part with `+`. (3b)
- A new/review session without a `new_morph_activity`, or the same activity type
  twice in one week. (3d)
- A question relying on untaught metalanguage - e.g. "part of speech of regulate?"
  when parts of speech are not in the term's grammar history. (2b taught-only rule)
- Renaming output files away from the team convention (`1a. Monday (morph).pptx`). (1)
- A scripted question without its answer (`EXPECT:` / `Answer:`) anywhere in any
  notes. (2b, 2d, 7)
- A review block labelled as a check when the notes do not say what the teacher scans or
  hears, or when the pace changes without complete response evidence. (2e, 8)
- Skipping a required review exposure because a few confident students answered correctly.
  Responsive pacing changes time and correction inside the OG sequence, not the sequence. (2e)
- Fewer than 10 Words to Spell Review, or spelling words that duplicate the same
  day's Words to Read Review. (2d)
- A grid script line that does not name the focus morpheme in caps (writing
  `disorder - a mess` instead of `disorder - ORD means order, dis- means apart:
  the order has come apart`). (3b)
- Missing memory hook, or a new hook every day instead of the same hook all week. (3b)
- Words to Spell New not drawn from that day's Words to Read grid. (3c)
- An extension that names a concept without a one-line reminder of it, or with no
  Stretch step. (3c)
- A fixed-answer morphology task without the immediately following green answer-check
  slide titled `Tick it or fix it - ...`, or using the old `Check and fix` wording. (3d)
- A You Do task slide with no completed worked example (`One done for you: ...`), or
  whose description a student cannot follow without the teacher re-explaining it. (3d, 7)
- A You Do item with more than one defensible answer (`pottery = P or R/P`), or an
  abstract classification/meaning-audit task set as independent work - ambiguity
  belongs in the teacher-led We Do, never the You Do. (3d)
- An answer/check slide showing bare labels without the one-line reasoning, or
  `check_notes` that do not MODEL the reasoning for the first item before students
  tick. (3d, 7)
- A fixed-answer grammar You do without its `Tick it or fix it - ...` answer slide. (7)
- A You Do CIRCULATE beat with no scripted prompt for stuck students (a prompt ladder
  is required; `CIRCULATE: check work` is a defect). (3d, 7)
- Notes containing raw markdown, multiple thoughts/numbered beats on one source line,
  an overlong iPad paragraph, non-left-aligned paragraphs, or more than eight live-zone
  lines outside the explicit Words to Read Review exception. (8)
- A new morphology word grid below 27 pt, with a mid-word wrap, or with three narrow
  columns left unadjusted for long words. (3b, 10)
- A morpheme keyword/meaning that differs from the canonical photographed card (or
  the fallback bank when no photo transcription exists), or an invented one. (4)
- A new-family word outside the captured card's `associated_words` without a verified,
  logged exception, or any automatic use of an `excluded_words` entry. (4c)
- Replacing an Australian catalogue spelling with a US spelling, or forcing an opaque
  associated word into a false morpheme-sum explanation. (4c)
- Resolving a duplicated heading such as `di-` without matching its intended meaning.
  (4b)
- A sound-bank or review-card label that is not verbatim in the photographed catalogue
  (or the fallback bank for an uncaptured card), e.g. a made-up `-ion` card when the
  taught card is `-tion / -sion`. (2c)
- A meaning-hunt or submarine script line on a PATTERN morpheme (-ine, -ciate/-tiate) -
  those are drilled for sound, they have no meaning to route through. (2a)
- A learned word highlight on the wrong letters, or a why-note without the
  Australian pronunciation breakdown. (5)
- A dictation without a cups block, with a capital not shown green, with an uncounted
  red mark/capital/target, or with the wrong meter (first = green; second/trickier =
  yellow), or
  with sentence lengths outside 10-12 / 14-16 words. (6)
- A dictation sentence containing an invented or misremembered word (`deisomograph`),
  or one that strings targets into a scene no student can picture when it is read
  aloud. (6)
- An unpaired quotation mark in a dictation sentence, or curly quotes typed into the
  spec - the spec stays straight ASCII and the builder renders the directional marks. (6)
- Grammar authored as `lines` (the bland rejected layout) instead of rule / example /
  items / routine / footer. (7)
- An I do that does not restate the rule (absent-student rule). (7)
- A grammar day that never tells students WHY the focus matters to their own writing,
  or never scripts the HOW transfer line at the close - labelling practised for its
  own sake. (7)
- A new morpheme taught without its one-line day-1 WHY (what the morpheme buys the
  student as a reader and speller). (3a)
- Smart quotes, em dashes, or ellipsis characters anywhere in the spec. (0a)
- A scripted whole-class "read all five rows" warm-up beat, or all four group
  allocations crammed onto one `READ:` line instead of one numbered line each. (2b)
- A review question that cannot be answered from what has already been taught. (2b)
- A sound bank note that describes the colour key instead of stating it, or a variant
  label allowed to break mid-bracket. (2c)
- A review-day sound bank containing the day's focus morpheme - that slot belongs to
  another review prefix/suffix/root. (2c)
- A spelling-review sentence that could be from a dictionary rather than the students'
  own school/science/sport world, or one that leaves a homophone ambiguous by ear. (2d)
- An `After checking:` / `Answer:` pair that feeds the answer into the question and
  repeats it back, or that a non-confident teacher cannot read aloud without sounding
  foolish. (2d read-aloud test)
- A word grid whose size is not 9 or 12. (3b)
- Choosing a derivative with no taught morphemes when the family offered one that
  stacks a recently taught prefix, root or suffix. (3b morpheme stacking)
- A submarine line whose fused meaning does not mirror the parts just named
  (`micro small + bio life + -ology the study of` then `the study of life too small to
  see` instead of `the study of small life`). (3b)
- A memory hook that only restates the meaning in a slogan voice, or a hook missing
  from the copy-card notes. (3a, 3b)
- A worked example written as one run-on line instead of putting the equation on its
  own authored line. (3d, 7)
- The word `STRETCH` anywhere in a new deck, or an "extension" that is one notch harder
  rather than about 18 months ahead. (3e)
- A fixed-answer You Do with no `ANSWERS:` block at the very bottom of the notes. (3d, 7)
- A `CIRCULATE:` prompt that only repeats the rule or points back at the worked example
  the teacher has just read aloud. (3d)
- A check-slide `ASK:` whose answer is visible without thinking. (3d)
- Two new learned words on consecutive sessions instead of sessions 1 and 3. (5)
- A dictation sized by word count alone, outside 55-75 / 80-100 characters. (6)
- A `punctuation` entry naming a grammatical construct with no plain-English reminder
  in brackets, or a colon used before reported speech where a comma belongs. (6)
- A harder dictation rendering with the needle still near the start. (6)
- Whiteboards, `boards up`, or the boards cue script anywhere in a grammar block. (7)
- A grammar footer that is instruction-shaped noise rather than a tease, stem or
  self-check question. (7)

# 10c. STRUCTURED-SLIDE SPACE BUDGET (grammar blocks and the morphology You Do)

`draw_structured_slide` stacks rule banner -> hero example -> item lines from y=1.15
and the build FAILS once content passes y=4.60, where the routine chip sits. Nothing
auto-shrinks except the hero, and nothing clips - an oversized item line silently
overlaps the one below it. Work the budget out BEFORE writing the spec:

| Element | Vertical cost |
|---|---|
| rule banner | 0.76 |
| hero example (with items below) | 1.77 |
| hero example (no items) | 2.57 |
| each item line, `item_size` < 20 | 0.46 |
| each item line, `item_size` >= 20 | 0.58 |

- rule + hero + items therefore allows **2 item lines**. Pack four numbered prompts
  onto two lines (`"1. ... = ?      2. ... = ?"`) rather than dropping the rule banner
  or the worked example - both are required.
- Item lines must stay under `8.2 * 72 / (item_size * 0.55)` characters or they wrap
  into the next line. At 17 pt that is about 63 characters.
- Check slides (rule + up to 4 green answers, no hero) fit comfortably, but the same
  character limit applies at `check_item_size` (default 22 -> about 49 characters).
- Routine chips stay on one line under about 30 characters; footers under about 45.
- Hero text drops to 20 pt minimum, and the gate rejects a grammar example below that -
  keep heroes under roughly 180 characters.

---

# 11. PLANNING A TERM FROM THE USER'S INPUT

When the user pastes a term block (section 12):

1. Build the term timeline: which morpheme is taught which day, week by week. From it,
   derive every session's rolling review-10, the words-to-read/spell weighting pools,
   and the learned-word review chain (two most recent).
2. Resolve every morpheme against the photographed category catalogue and run the
   associated-word/exclusion audit (section 4) BEFORE writing specs. For cards absent
   from the photographed set, batch any genuinely unconfirmed fallback questions.
3. Author week specs in teaching order - later weeks' review pools depend on earlier
   weeks' words, so keep a running record of which derivative words each session used
   (words used in Words to Read/Spell - New Morphology become future review words).
4. Build and QA week by week. Deliver per-week folders under `output/`.
5. In your summary, list: every card absent from the photographed set and how it was
   verified, every authorised outside derivative, every word rejected for exclusion,
   morpheme-integrity or decodability reasons, and any learned word that looked fully
   decodable. Never propose a replacement keyword for a captured card.

Mapping morphemes to sessions (pick the row that matches; ask only if none fits):

| Morphemes | Sessions | Pattern |
|---|---|---|
| 2 | 5 | new, review, new, review, week_review (the default) |
| 3 | 5 | new, review, new, review, new - the team's historic pattern: `1a, 1b, 2a, 2b, 3.` with the third morpheme taught Friday, consolidated next week |
| 3 | 4 | new, new, new, week_review (no per-morpheme review day; the week review carries consolidation - flag this trade-off in your summary) |
| 2 | 4 | new, review, new, review (fold week-review weighting into Thursday's review sections) |
| 1 | any | new, then review days; last session week_review |
| 2-3 | 3 | new, new(, new) - flag that review days are lost and weight the following week's history hard toward these morphemes |

Days with no session: keep them in `days`/`overview` with a "No session this week"
cell so the weekly overview table stays truthful.

---

# 12. USER REQUEST FORMAT (user replaces the XYZs, nothing else)

The user fills in XYZ values only - same convention as MEGA_PROMPT.md. Repeat the
WEEK block for every week of the term. Everything not in this block (derivative
words, review lists, sound banks, spelling words, dictation sentences, grammar
scripts, teacher notes) is YOUR job to generate under the recipes above.

Interpreting the block:
- `Morphology focus` lists the week's new morphemes with their type, e.g.
  `ord/ordin (root), -ible (suffix)`. Two per week is the default (taught Mon + Wed).
  One is fine (taught Mon; every other day reviews it).
- `Learned words` is the week's list in teaching order; deal one new word per session
  Mon-Thu, review-only Friday (section 5).
- `Sessions per week` under 5 compresses per section 11.
- `Additional notes` may carry inquiry links, assessment weeks, short weeks, or a report
  that a catalogue transcription is wrong. It may not silently override a captured
  card's keyword or meaning. Verify any reported correction against the physical card,
  update the relevant category JSON and generated views first, then use the corrected
  canon. If `Additional notes` is empty, proceed with defaults - do not ask.
- If `Recently taught morphemes` is empty and the bank/history gives you nothing,
  ask for it - week 1's review deck cannot be invented.

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
