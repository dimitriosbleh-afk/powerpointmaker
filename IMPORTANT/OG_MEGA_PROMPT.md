# Orton-Gillingham (OG) Deck Builder Mega-Prompt v1.9
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

The v1.1 refinement tightened response quality and responsive pacing inside the locked
Yoshimoto sequence. The v1.2 refinement fixes four defects the school found in shipped
decks, all of them about WHAT STUDENTS HAVE ACTUALLY BEEN TAUGHT at the moment you ask
them something:

1. The Sound Bank was handed students the day's brand-new morpheme, thirteen slides
   before it was taught (2c - it is now excluded on every day type, with a mechanical
   derivation procedure).
2. Word lists were chosen for catalogue tidiness ahead of teaching quality, so grids
   filled with ornate low-value derivatives when a simpler, better-stacked word was
   available (3b selection order, 4c ground 3).
3. Questions and tasks required morphemes the class had never been taught (new section
   2f, the taught-morpheme gate).
4. Grammar and I do notes told the teacher WHAT the answer was but not HOW to
   demonstrate it, so a teacher delivering cold had to invent the demonstration
   (7 teach-it-cold notes).

The v1.3 refinement (school feedback, 19 Aug 2026 - Steph, who teaches the decks)
changes two things and confirms one:

1. DICTATION IS REVISION FROM 2-3 WEEKS AGO, never this week's words. Students were
   copying that week's morphology and learned words straight out of the front of
   their books instead of spelling them (section 6 - sourcing order reversed, builder
   gate added).
2. The morphology bank must never show a morpheme before it is taught - not today's,
   and not one taught later in the same week either (2a/2c - the builder now checks
   the review-10 cards and the Sound Bank against the whole week timeline).
3. The Week 6 shape - `new, review, new, review` with a real review day after each
   morpheme - is the structure the team asked to keep (section 11).

The v1.4 refinement responds to teacher feedback that the notes could meet every
format rule and still sound clipped when read word for word. It permits brief spoken
connectors such as "Okay", "Now", "So" and "Let's" where they make a line flow, while
still excluding repeated greetings, management chatter and extra explanation. The
teacher should be able to read the scripted lines in order without inventing the words
that join them.

The v1.5 refinement (team meeting, 9 Sept 2026) reshapes the New Morphology You Do
after the team reported it was too hard and too repetitive:

1. MEANING SUMS ARE RETIRED. Handing students meanings and asking for the word
   (`house + study of + relating to = ?`) made them hold three or four parts at once,
   and it kept reaching for parts nobody had taught (`-al` in ecological). The four
   remaining types - word building, literal meanings, sorting, showing sentence - are
   the ones the team said work in the room (3d, builder gate).
2. ONE RECALL. The only part a student supplies from memory is today's morpheme; every
   other part is printed with its meaning on the slide face. Taught earlier no longer
   counts as printed on the You Do (2f, 3d, builder gate).
3. Same task for everyone, no tiers on the face; the green Tick it or fix it slide
   stays as it is (3d).

The v1.6 refinement (James, 9 Oct 2026, ahead of Kate - the school's OG lead - reviewing
the decks) fixes two things the room found:

1. PLAIN JOINING IS NOT A TASK. `hyper + link = ?` asks nothing: the answer is the two
   printed parts said together, so no student used hyper's meaning. Every You Do now makes
   the student USE today's morpheme's meaning to decide something - which word fits a
   sentence, whether a sentence uses the word truly, which word matches a meaning, whether
   the letters are the morpheme at all (3d menu).
2. ONE QUESTION PER SLIDE, ANSWER ON THE CLICK. The morphology You Do and the grammar We
   do / You do are ROUNDS (3d, 7): one item per slide at hero size, every student answers,
   the teacher's click reveals a green answer bar with its one-line reason. The reveal is
   the tick-it-or-fix-it moment, so a round needs no separate check slide. Grammar is held
   to five minutes by its shape: I do, one We do item, three You do items (7).

The v1.7 refinement (Kate Watt, the school's Yoshimoto master trainer, observed the
decks on 12-13 Oct 2026; James approved every change on 13 Oct) reshapes the week.
Her message: keep the structure and the content consistent, but 5/6 does not need the
whole lesson four days running, and the decks should leave the teacher room to teach.
Sources: her 1:1 and leadership-meeting advice, Yoshimoto's "One Lesson Over 2 Days"
sheet and the "Advanced Morphology - 2026" training deck (lesson-planning slides). The
training deck and its videos stay inside the school - never copy them into this repo.

1. ONE MORPHEME A WEEK, EVERY DAY, 20-30 MINUTES. The lesson is split across the week
   the way Yoshimoto splits one lesson over two days: Monday teaches, Tuesday spells and
   dictates, Wednesday builds and breaks words, Thursday does parts of speech and
   grammar, Friday reviews. Each day carries only its own blocks (section 1).
2. DICTATION CARRIES TODAY'S TEACHING. Yoshimoto's planning slide: dictation "must
   prioritise what was introduced in this lesson and can also include anything
   previously taught". Copying from the book is fine. One harder word from a week or two
   back checks retention. This reverses the v1.3 rule (section 6).
3. THE SOUND BANK IS DICTATED, NOT COPIED: the teacher says each morpheme or sound,
   students write it, the click shows the box (2c).
4. EVERYONE READS EVERY ROW FIRST in Words to Read Review; the group rows are the
   second, fluency read (2b). This reverses the v1.x ban.
5. PARTS OF SPEECH ARE TAUGHT on suffixes: the card drill says them and Thursday drills
   them (2a, 3g). Kate named this the one gap across the whole school.
6. WRITE THE WORD THAT NAMES... (a meaning given, the word written) and the word matrix
   replace the per-day You Do (3f). Keywords are a memory prompt, never quizzed;
   students choose their own (3a).
7. USEFUL WORDS ONLY. Pick words an adult uses; six good words beat ten with curios
   (3b). Words with an untaught part go last, with its meaning given in one line.
8. COPS replaces CUPS school-wide: Capitals, Organisation, Punctuation, Spelling (6).
9. RESPOND TO THE CLASS. An optional `From our reviews` line in the request plants the
   class's wobbles into next week (12), fed by a quick Friday vote (1).

The v1.8 refinement (James, 22 Oct 2026, after teaching the first v1.7 week; every
point settled with him) keeps the v1.7 week and fixes what the room found:

1. EVERY SESSION OPENS WITH THE CARD DRILL AND WORDS TO READ REVIEW, Monday to Friday,
   fresh random card order each day. Tuesday's separate recap card is gone (1, 2a).
2. 5-4-3-2-1 READING REVIEW (Kate Watt): 5 words from the most recently taught morpheme,
   then 4, 3, 2, 1 from the next four back; the 2 and 1 slots may be words from a while
   back (2b).
3. THE READING REVIEW WAS "SOOOO REPETITIVE" AND KIDS SWITCHED OFF. At most 3 questions,
   one-word answers on the board, at least two as 10 sec turn and talk then choral,
   "why" only as turn and tell, and the question format and second read rotate daily (2b).
4. REAL QUESTIONS: "What is one word built from terr, meaning land?", never "Okay, give
   me..." (2a).
5. MEMORY HOOK: 8 words or fewer, on the new-morpheme card with a small picture, and the
   first line of the notes. No PROCEDURE block on the new-morpheme header (3a).
6. VET EVERY WORD EQUATION against its answer; adult-use words only (3b).
7. AFTER-CHECK QUESTIONS only where the part's meaning plainly shows in the word, at most
   3 a session; no "Click to reveal answer." line (2d). No "SAY each word..." line on
   Words to Spell New (3c).
8. THREE EXTENSION TYPES ONLY: longest word hunt, beat the number, two in one (3e).
9. BREAK IT APART: one empty box per part, root filled in, students write only the parts
   (3f).
10. PARTS OF SPEECH FROM THE SENTENCE: Thursday was "brutal" because the suffixes were
    old and nobody could say why a suffix makes a noun. Each suffix card now shows its
    meaning and a keyword sentence, asks "What job does gravity do in this sentence?"
    with noun / verb / adjective as choices, and reveals the answer and its reason on
    the click (3g). The first v1.8 try, the frames "the ___ / to ___ / very ___", is
    retired (James, 23 Oct 2026: the word never sat in the frames, the answer was on
    the slide, and even a teacher could not tell whether "the gravity" was right).
11. LEARNED WORD ROUTINES rewritten (5), and SHORT WEEKS of 3 or 4 sessions never cut new
    content (11).
12. TOP-TO-BOTTOM HIERARCHY on every slide we build ourselves, and plain-English notes
    with no Purpose line or stage tag (3h, 3c).

A v1.8 week says `"og_version": "1.8"` at the top of its spec. Legacy `new` / `review`
weeks (v1.0-v1.6) and v1.7 weeks (Term 4 Week 3) still build unchanged.

The v1.9 refinement (James, 30 Oct 2026) takes the best of the 5/6 Challenge team's
package - built from a 3/4 teacher's planner decks - and re-pitches each piece for this
Grade 5/6 class (Sabrina spoke highly of their decks). It keeps the v1.8 week and adds:

1. A PICTURE OF ITS MEANING ON EVERY WORDS TO SPELL REVIEW WORD, appearing on the same
   click as the word (2d).
2. TWO NEW WEDNESDAY THIRD ROUNDS, rotating weekly with Kate's "Write the word that
   names...": `Complete the sentence` (a word bank whose size varies - extra, short,
   exact - behind one fixed instruction line) and `Choose the best meaning` (three
   choices, the best wrong one built from a real meaning of one of the word's parts)
   (3f).
3. ENABLER AND EXTENSION SHEETS: for each writing session (Monday, Tuesday, Thursday),
   a one-page A4 enabler on dotted thirds in separated Victorian Modern letters, with
   the week's teaching left as the gaps, and an extension sheet with the same sections
   pushed a step further, on the school worksheet layout (1b). Every printable for the week lands in one PRINT PACK
   with a cover saying who each page is for (1c), so staff print once and organise
   nothing.
4. THE FRIDAY REVIEW WORKSHEET: Friday runs the card drill and Words to Read Review,
   then a two-page books-closed worksheet on this week only, with a teacher check box
   naming the exact skill on every question, an answers page, and a one-page Teacher
   Follow-up sheet saying when a result goes into a student's program and how to
   revisit it next week with no slide edits. The follow-up sheet replaces the
   trickiest-this-week vote (1a).
5. THE 3/4 TEAM'S CARD WORDING IS ADVISORY, NOT ADOPTED. Where it differs from our
   photographed cards the builder raises a NOTE; our slides keep the wording this class
   learned (4e).

A v1.9 week says `"og_version": "1.9"`. v1.8 and older specs still build unchanged;
every new week uses v1.9. Worked example: `og_planner/weeks/sample_v19_term4_week4.json`
(the delivered Week 4 ject teaching plus every v1.9 piece and a Friday).

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

# 1. THE FIVE-DAY SHAPE (v1.7, daily review added in v1.8)

ONE morpheme a week, OG every day, 20-30 minutes a day. The week is one Yoshimoto
lesson split across the days (`OG/OG Post Training Resources/OneLessonOver2Days.pdf`:
"30 minute lesson for each day ... One lesson over 2 days = minimum"), with Kate Watt's
upper-years additions: a build day, a parts-of-speech day, and no new teaching on
Friday. Each day carries ONLY its own blocks - the builder emits nothing else.

| Day | type in spec | What runs (in slide order) | About |
|---|---|---|---|
| Monday | `teach` | card drill (10 cards) - Words to Read Review (15) - NEW morpheme card with memory hook and picture, copy slide, write it 3 times, class brainstorm - Words to Read grid (6, 9 or 12) - Your keyword and one sentence | 25-30 min |
| Tuesday | `spell` | card drill - Words to Read Review - Sound Bank, dictated (2c) - Words to Spell Review (6-8) - Words to Spell New (4, off Monday's grid) - learned words (2 review + NEW word 1) - dictation x2 | 25-30 min |
| Wednesday | `build` | card drill - Words to Read Review - word matrix - Break it apart round - Write the word that names... round (3f) | 20-25 min |
| Thursday | `apply` | card drill - Words to Read Review - suffix parts-of-speech drill, job read from a sentence (3g) - learned words (2 review + NEW word 2) - grammar I do / We do / You do (7) | 25-30 min |
| Friday | `week_review` | v1.9: card drill (10) - Words to Read Review (15) - `Complete this Friday Review` slide, and the deck STOPS; students do the worksheet (1a). v1.8: card drill - Words to Read Review - Sound Bank, dictated - Words to Spell Review (6-10) - learned words review - dictation x2 - grammar (only if no Thursday ran) - trickiest-this-week vote | 25 min |

v1.9 adds the meaning pictures (2d), the Wednesday third-round rotation (3f), the
enabler sheets (1b) and the Friday Review worksheet (1a). In a v1.9 week the Friday
Review replaces the vote; a short week with no Friday keeps the vote on its last
session (11).

v1.8: every day opens with the card drill and the Words to Read Review, so retrieval
runs every session. The table above is the v1.8 shape; a v1.7 spec (no `og_version`)
still builds the v1.7 shape, where Tuesday had a recap card and Wednesday/Thursday had
no opening review.

Why this shape (Kate, 13 Oct 2026): "Big focus on a Monday morning ... tomorrow card
deck, spelling ... Wednesday, let's build some words." "It's still doing it every day,
but the amount that you do every day can be tweaked." The structure is fixed so every
class runs the same week; the teacher's delivery inside it is theirs.

- Monday has no spelling, no You Do, no learned words, no dictation and no grammar.
- Tuesday opens with the card drill and Words to Read Review like every day (v1.8).
  There is no separate recap card: the new morpheme is already on the drill.
- Wednesday is the morphology You Do day: building (matrix) and its opposite (breaking
  apart), then meaning-to-word retrieval on whiteboards.
- Thursday is the only grammar day in a five-session week. The grammar block keeps its
  five-slide shape (section 7), and its You do or EXTENSION puts one of the week's words
  into a sentence in the grammar focus - morphology into writing.
- Friday teaches nothing new. Its closing vote (fingers 1-4: which was trickiest this
  week?) gives the teacher one class-level signal for next week's `From our reviews`
  line (12) without tracking any individual student.

FILENAMES are auto-generated and must not be changed: `1. Monday (port).pptx`,
`2. Tuesday (port spelling).pptx`, `3. Wednesday (port building).pptx`,
`4. Thursday (parts of speech).pptx`, `5. Friday (week review).pptx` - number = session
order, `/` in a morpheme becomes `-`. Override only if the user asks, via `file_name`.

SHORT WEEKS (Kate: "trim something you put back next week, never something you never
revisit"): see section 11.

LEGACY (v1.0-v1.6): `new` / `review` days ran every component every day, two morphemes a
week, files named `1a. Monday (ord-ordin).pptx`. Those specs still build unchanged and
stay the historical record; do not author new weeks in that shape.

## 1a. FRIDAY REVIEW WORKSHEET (v1.9)

Ported from the 5/6 Challenge team, whose sheets copy the school's `Friday OG Review
Term 3 - Challenge` worksheets. Their content was pitched for a 3/4 class (underline the
appositive in "Ruby, my pet dog, sleeps on my bed"; write what NSW stands for). Ours
keeps their layout and their teacher-facing ideas, and sets the content at Grade 5/6.

THE SLIDE. The Friday deck runs the card drill and Words to Read Review, then one
`Complete this Friday Review` slide, and STOPS. The builder writes the slide and its
notes (run it books closed, about 20 minutes, read Part E whole, in chunks, then whole
again). The notes name the files as plain text, because a relative link does not
survive the Google Slides import.

THE FILES, written by the builder into the week folder's `Friday Review/`:
- `Friday Review - Term T Week N.pdf`: two student pages at most, then an answers page
  headed `Teacher copy - do not print for students`. Teachers print only pages 1-2.
- `Friday Review - Term T Week N Teacher Follow-up.pdf`: one landscape page, one row per
  question: `Question` / `What it checks` (the same words as its check boxes) / `Add to a
  student's program if...` (a countable trigger: `2 or more of the 6 jobs wrong`,
  `mortgage misspelled`) / `Revisit it next week (no slide edits)` (one move that slots
  into a section that will already be there - review cards, Words to Read Review, Words
  to Spell Review, dictation - and needs only the teacher's voice or the class
  whiteboard). Its rows are the feed for next week's `From our reviews` line (12).

ONLY THIS WEEK. Every item tests something taught this week: Thursday's grammar focus,
Thursday's suffix parts of speech, this week's learned words, this week's morpheme and
words the class met this week. The taught-morpheme gate (2f) binds every item: a part
the class has not been taught is printed with its meaning on the page (a `Word Bank:`
of parts), so the only part a student supplies from memory is this week's.

THE FIVE PARTS, in this order, each at Grade 5/6:
- PART A: GRAMMAR - this week's grammar focus, 2 question blocks. Students PRODUCE, they
  do not just spot: rewrite with the new punctuation, join two sentences, then write one
  of their own. An underline-only task is a Year 3/4 task.
- PART B: PARTS OF SPEECH FROM A SENTENCE - 5 words built on this week's suffix cards
  (3g), each in a short sentence with the word in CAPITALS; a one-column `codes` task
  takes the job. `Tip:` gives the three jobs in words (noun - names a thing or idea,
  verb - something you do, adjective - describes a noun) on one line and one worked
  sentence on a second line - a two-line Tip, never a side-by-side Example box, so the
  block fits under Part A. Each sentence must prove the job, as on the slides. Never a word whose job depends on
  context (`associate`). (The frames were retired on 23 Oct 2026, 3g.)
- PART C: LEARNED WORDS - this week's two new learned words in sentence gaps, spelled
  from memory (no word bank: a bank turns spelling into copying).
- PART D: MORPHOLOGY - two blocks on this week's morpheme, headed with its label and
  meaning exactly (`ject = throw`; the builder checks the label). Block 1: a `Word / Word
  sum / What the parts mean` table of three words with 3 or more parts, with the other
  parts' meanings in a `Word Bank:`. Block 2, the 5/6 step up: two words NOT met this
  week whose parts are all taught or printed (`ejector = e (out) + ject + or (a thing
  that) =`) - students build the meaning from the parts. An optional Challenge (longest
  ject word with its full sum) goes only if two pages still hold it; it is the first cut.
- PART E: DICTATION - five or more full-width lines. Two sentences built from this
  week's words (each carries a word on this week's morpheme; the builder checks) and
  this week's grammar focus. Section 6 rules apply (sentence types, COPS, real words).

TEACHER CHECK BOXES. Every question block has a `Teacher check:` panel with one box per
skill, at most two, each naming the exact thing assessed: `ject = throw`, `-ity / -tude
= noun`, `mortgage`, `semicolon replaces , so / , but` - never `Meaning`, `Grammar`,
`Spelling` (the builder fails those). The follow-up sheet uses the same labels, and
every follow-up label must match a box (Part E is marked with COPS, no box).

SPACE. A4, 13 mm margins, body 11.5 pt, writing lines 10 mm apart, a one-word answer
line at least 42 mm. TWO STUDENT PAGES, never three: a question never splits across a
page, so a block that does not fit moves whole to the next page. If it runs over, cut
an item or the Challenge - never the writing space. Render every page and check it.

SPEC: `friday_review` on the Friday session - `part_a` ... `part_d` (each `{title,
subblocks}`; a sub-block is `{heading, instruction, tip?, example?, word_bank?, task,
checks, answers}`), optional `challenges`, `dictation_lines`, `dictation` (2 x
`{sentence, chunks, targets, cops}`) and `follow_up` (rows of `{question, checks,
flag_when, next_move}`). Task types: `sentences` (with `answer_lines`), `cloze` (`___`
gaps), `meaning_lines`, `clue_answer`, `codes` (`codes`, `columns`), `match`
(`meanings`), `table` (`headers`, `widths`, `rows` - an empty cell becomes a line),
`syllables`, `build`, `lines`. See the sample week.

## 1b. ENABLER SHEETS (v1.9)

Students who need support write on an enabler sheet instead of a blank page. The
Challenge team's teachers made these by hand in Word; the builder now writes them from
the spec, so they always match the slides. You author nothing extra, but you render and
check every sheet.

WHICH SESSIONS. Monday (`teach`), Tuesday (`spell`), Thursday (`apply`, when it has
learned words or dictation), and a Friday only when it has no Friday Review. A
three-session week's `combined` day gets Thursday's sheet, because it carries Thursday's
learned words and suffix cards (Term 4 Week 5). Wednesday's build is whiteboards and
OG books, and the Friday Review is its own written task. A week with no Friday Review
leaves it off the print pack cover and the Weekly Overview notes.
Files: `Enablers/1. Monday enabler.pdf` etc. (number = session order), named in the
Weekly Overview notes of every deck that week.

THE 5/6 RE-PITCH (from the 3/4 sheets):
- VICTORIAN MODERN LETTERS, SEPARATED, NOT JOINED (James, 30 Oct 2026: "Vic modern
  but not cursive, just regular separated"): `VicDots` (`VICDOTS_.TTF`, dotted) to
  trace and `VicRegular` (`VICREGUL.TTF`, solid) for the morpheme. Never joined
  cursive, never the Foundation-2 pre-cursive, and never `VICCursiveDotted` /
  `VICCursiveSolid`, whose letter widths are only right inside Word (in a PDF the
  letters pile on top of each other). Fonts are read from the installed fonts
  (`C:\Windows\Fonts`, or on a Mac `~/Library/Fonts` after Font Book's Install) and
  never copied into this public repo. VicDots and VicRegular are paid School Fonts fonts; where
  they are missing the builder falls back to the Department's free Victorian Modern
  Cursive (`Vicmorg_.ttf`, the same separated letters) with the tracing letters in
  pale grey instead of dots (31 Oct 2026).
- DOTTED THIRDS at about 4.9 mm (a 14.8 mm row), down to 4.2 mm to fit one page, with a
  NOTE below 4.4 mm. The 3/4 sheets used 6 mm.
- THE TEACHING IS THE GAP. The 3/4 sheets printed whole words to trace. Ours leave the
  week's morpheme blank inside each Words to Spell New word (`pro____or`), every
  dictation target blank in both sentences, the unfair part of each learned word blank
  in its second copy, and the Sound Bank as first letters only (it is dictated, 2c).
- Words to Spell Review: word 1 in full, words 2-4 half shown, the rest first letter
  only. (The 3/4 sheets printed the last word in full as the hardest; our list runs
  most-recent-first, not easy-to-hard, so that step does not carry over.)

SECTIONS. Monday: the morpheme card (`ject = throw`), the morpheme three times to trace,
`My keyword and what it means`, `My sentence with a ject word`. Tuesday: New
Morphology, Sound Bank, Words to Spell, Learned Words (Review), Learned Words (New),
Dictation. Thursday: Learned Words, and Dictation when the short-week rule puts one
there. Every heading sits in its own strip, so no descender runs through it. One page,
always: the builder fails a sheet that cannot fit even at the smallest size. The page
keeps a 24 pt foot margin (Term 4 Week 5: a Tuesday-type sheet with two long review
learned words ran 6 pt past the old 30 pt margin at 4.2 mm) - trim the session, never
the size.

EXTENSION SHEETS (James, 30 Oct 2026: "make an extension page equivalent"). Every
session that gets an enabler also gets `Extension/1. Monday extension.pdf`, for
students ready for more: the enabler's sections in the same order, each pushed one step
further (3e: the same concept, pushed). Where the enabler prints the word to trace with
a gap, the extension prints only what a student needs to start, and the whole word is
theirs:
- New Morphology: Monday - three words on the morpheme and what each means, the
  longest word they can build with its word sum, a sentence with TWO of the words.
  Tuesday - the grid meaning only (`to throw in`), the student writes the word.
- Sound Bank: a box for each part as it is dictated, then a word for two parts.
- Words to Spell: each word spelled, then written again with its parts marked with `/`.
- Learned Words: each written from memory, then the new word in a sentence (on
  Thursday, in the week's grammar focus).
- Dictation: lines for both sentences, then one more word on the morpheme added to
  sentence 1 with a `^`.
- Thursday adds the day's real work the enabler leaves to the slides: a new word for
  each suffix card with its job (noun, verb or adjective), and a sentence using two of
  them that shows each word doing that job.
THE LOOK IS THE SCHOOL WORKSHEET, NOT THE ENABLER (James, 31 Oct 2026: extension
students do not need the enabler's look). No dotted thirds and no Victorian Modern:
`og_planner/render_extension.js` draws the sheet with the shared worksheet layout
(`MEGA_PROMPT.md` 86) - lettered sections, numbered task cards, ruled lines and tables,
the literacy Years 5-6 palette - then the answer key in the same layout with the
answers in colour. One page, or two printed as one sheet double-sided (86); the answer
key follows at the same length, and the print pack cover says which pages are which.
The builder writes it from the spec (`enabler_pdf.extension_sheet`), so nobody
authors it.

THE DECKS NAME THE SHEETS. The title slide's notes say which enabler and extension
sheet the day uses, so the teacher knows what to hand out without opening the pack.

## 1c. THE WEEK'S PRINT PACK (v1.9)

James (30 Oct 2026): better student outcomes, but "we don't want to create more work
for our staff". So the builder puts every printable for the week into ONE file,
`0. Print pack - Term T Week N.pdf` in the week folder, after the last session:
- A cover page: a table of every sheet (day, sheet, who it is for, which pages), then
  "Each day at a glance" - the day's deck and what goes on the desks (OG book; mini
  whiteboard and marker only on days whose routine chips say whiteboards; a pencil and
  the Friday Review on Friday).
- Then the student pages in day order: each day's enabler and extension sheet, then
  the Friday Review's two student pages.
- Teacher copies at the back, shaded on the cover: extension answers, Friday Review
  answers, the Teacher Follow-up sheet.
The Weekly Overview notes of every deck name the pack. You author nothing for it; open
the cover and check it lists every sheet the week made.

---

# 2. THE ROLLING REVIEW ENGINE (the heart of every deck)

## 2a. Morphology review cards (10 per session)

- The 10 cards are the 10 MOST RECENTLY taught morphemes, counting backward from today,
  excluding today's new morpheme. A morpheme enters the review deck the session AFTER it
  is taught; when one enters, the oldest of the 10 drops out.
- NOTHING UNTAUGHT ON A CARD. A morpheme that this week teaches on a LATER day (Monday's
  cards carrying Wednesday's morpheme) is as wrong as today's new one: students drill it
  cold a dozen slides before anyone shows them the card. Same rule for the Sound Bank
  (2c). The builder fails the build on either (today's new morpheme on a `new` day's
  cards; any later-in-week focus morpheme on any day's cards or bank).
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
- THE DRILL IS CHORAL AND FAST (v1.7, Yoshimoto's "modified card drill": "teacher shows
  card, students say meaning"). The whole class says `port means carry` together; a
  suffix card adds the part of speech it makes (`-tude means condition of, noun`).
  Then one student gives a word with it and what that word means. The Yoshimoto
  training videos run a whole deck this way in under two minutes. From v1.8 the cards
  run EVERY session, in a fresh random order each day (section 1); from Tuesday the
  week's new morpheme is one of the 10.
- KEYWORDS ARE A PROMPT, NEVER A QUIZ (Kate, Oct 2026: "the keyword is if they forget").
  Never script "What's the keyword?" as a question. The meaning is what students must
  know; the keyword is what the teacher offers when a student is stuck.
- Yoshimoto's Visual Card Drill (Morphology LessonPlan, see
  `og_planner/OG_LIBRARY_INDEX.md`) has students pronounce the morph, give the
  MEANING, and give a derivative - occasionally used in a sentence. Every session,
  2-3 cards add an optional `extra_task` object with a direct teacher prompt and
  possible answers. The rendered label is `Extra task:`, never the unclear
  `Derivative ask:`. The prompt is a real question (v1.8, James: "Okay, give me..." is not
  a question): `ASK: "What is one word built from flect, meaning bend?"` followed on a
  new line by `Possible answers: flexible, deflect, reflection.` A prompt ending in `?`
  renders as `ASK:`; the builder fails a v1.8 deck containing "give me". Rotate which
  cards receive it.
- PATTERN MORPHEMES: some bank entries are pronunciation/spelling patterns, not
  meaning morphemes (their `meaning` field says so - e.g. `-ine` "says /in/ or
  /uhn/", `-ciate / -tiate` "says /sh-ee-ate/"). Drill these for SOUND only: the
  card prompt is "what does it say?", never "what does it mean?". Never use a
  pattern morpheme in a meaning-hunt ASK, a derivative-meaning question, or a
  submarine-model script line - it has no meaning to route through.

## 2b. Words to Read Review (15 words)

V1.8 SLOTS: 5-4-3-2-1 (Kate Watt: "five, four, three, two, one, meaning most recent").
5 words from the most recently taught morpheme, 4 from the one before, then 3, 2 and 1
from the next three back, counting the week's new morpheme from Tuesday on. The 2 and 1
slots may instead carry words from a while back (a Term 3 morpheme, a sound or spelling
pattern) to keep older learning alive. Jumble the 15 in the grid. Monday's newest
morpheme is last week's, and Monday must not reuse the previous Friday's words.

V1.8 SCRIPT (the v1.7 script below was "SOOOO repetitive", James 22 Oct 2026):
1. `1. READ: "Everyone, read every row with me, row 1 to row 5."` stays first.
2. ONE second read, and it ROTATES day to day: partners on alternate rows, whisper read
   then voices, sides of the room, news-reader read. Birthday groups every day are out.
3. AT MOST 3 numbered `ASK:` beats, each with a ONE-WORD answer that is on the board, so
   a choral answer has no discrepancy. At least two run as `10 sec, turn and talk. Then:
   "Everyone, together, on three."`. A "why" question is turn and tell only, never
   choral. Rotate the formats daily: "Which word means...", "Finish my sentence...",
   "Odd one out: a, b, c", and partner point-and-read (A points, B reads it and says what
   its part means). The builder warns on more than 3 ASKs or fewer than two turn and talk
   beats.

The v1.7 weighting and script below still apply to v1.7 specs and supply the word-choice
rules (fresh words daily, morpheme doing its job, taught-only questions) for every week.

V1.7 retrieval weighting - most recent gets most attention. Allocate the 15 slots:

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
  line, no bullets and no dense multi-question paragraph. Script the small spoken
  connections a teacher needs; do not reduce every line to a bare command. This slide
  is an explicit exception to the general eight-line Glance Format limit because the
  playful reading routine and retrieval questions must remain separate and easy to scan:
  1. EVERYONE READS EVERY ROW FIRST (v1.7, Kate Watt: "get all the kids to read the
     words at once ... who needs to read the words the most? Your front rowers"). The
     first read is decoding, so it is the whole class, every row, never teacher-first
     then echo. Write it as `1. READ: "Everyone, read every row with me, row 1 to
     row 5."` The builder fails a v1.7 deck whose first beat is not this.
  2-5. ONE NUMBERED `READ:` LINE PER GROUP for the second, fluency read - never a single
     line listing all four allocations. Use categories such as birthday season,
     first-name initial, handedness/month, or favourite activity; every student must
     belong to a group, with a self-choice fallback where a category could be
     ambiguous. Write them one per line:
     ```
     2. READ: "Summer and autumn birthdays, read rows 1 and 2."
     3. READ: "Winter and spring birthdays, read rows 3 and 4."
     4. READ: "Everyone, read row 5."
     ```
     (Legacy v1.0-v1.6 specs skipped the whole-class read and stated `Every student
     reads at least two rows in the group round.`; the builder still accepts that for
     those weeks only.)
  Then three numbered `ASK:` beats tied to morphemes on the board. Phrase each as words the
     teacher can read aloud immediately, then include think time, one response routine,
     and `EXPECT:` on the same line. Example: `4. ASK: "Which words use flect or flex,
     meaning bend?" 5 sec, choral response. EXPECT: flexible, deflect.`
  Then up to four quick checks, each on its OWN numbered line with its own answer. Never
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
- Every EXPECT answer to a find-the-word question ("Which words carry...", "Which word
  names...") must be checked against the actual word list on the slide - a scripted
  answer that is not on the board is a build-stopping content bug. Quick checks whose
  answer is said aloud rather than found (opposites, rhymes, as in "opposite of include?
  EXPECT: exclude") may answer with a word that is not on the board. Check rhymes in
  Australian English: wonder does not rhyme with ponder (fonder does).
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
  - PARTS OF SPEECH ARE TAUGHT FROM v1.7: noun, verb, adjective and adverb are taught
    on the suffix cards and drilled every Thursday from Term 4 Week 3 (3g), so a
    part-of-speech question about a suffix the class has met is safe from then on
    ("-ive makes which part of speech?"). Kate Watt: the part of speech is what a
    suffix is FOR.
  - STILL GATED behind the grammar history: clause and phrase names and tense labels.
    Use these ONLY if a current or earlier grammar focus in the user's term plan
    covered them.
  - EXCEPTION - inside the grammar lesson itself: when the user's chosen grammar
    focus requires a building-block term (prepositional phrases need "noun"), the
    grammar slides may use it, but the I do glosses it in passing on first use
    ("the noun after it - the naming word"). The gate still applies everywhere
    OUTSIDE the grammar section.
  - The Extension Wheel's Nouns/Verbs/Adjectives category is open from Term 4 Week 3.
- QUESTIONS MUST STRETCH (Yoshimoto's upper-level planning slide: "WTR extension -
  questions must be more sophisticated"). For 5/6 prefer the questions that make
  students reason with the parts - which word changes part of speech if you drop the
  suffix, which two words share a root and how their meanings differ, an analogy -
  over "which word has port in it?".
  - This rule gates untaught LABELS. Section 2f gates untaught MORPHEMES and is the
    other half of the same idea - a morpheme hunt is only "always safe" when the
    morpheme in question has actually been taught.

## 2c. Sound bank (9 boxes)

- v1.7: THE BANK IS DICTATED (Yoshimoto day 2, "Auditory: Sound"; Kate Watt: the
  students should be retrieving it, not copying it). It runs on Tuesday and Friday. The
  teacher says each item in the order the builder lists in the notes (`port - carry`),
  students write it in their books, and each click shows one box so they tick or fix
  straight away. Everything below about deriving the boxes from the spelling list still
  holds, with one change: a v1.7 bank has 6-9 boxes, because Tuesday and Friday spell
  6-8 review words and a short list rarely needs nine morphemes. Never pad with filler. Two things change because the boxes are hidden until the click:
  - This week's morpheme, taught on Monday, MAY go in Tuesday's and Friday's bank. It
    is retrieval now, and the strongest retrieval of the week. A morpheme not yet taught
    still never goes in (the builder checks).
  - One or two SOUNDS may join the bank (Kate: "even though you're doing morphology,
    throw a sound in") when a spelling word or the class's reviews need one - type
    `sound` with a `say` line (`{"morph": "ou", "type": "sound", "say": "/ow/ as in
    cloud"}`). Sound boxes render white, like the sound deck.
- LEGACY (v1.0-v1.6): students copied all nine boxes. The rules below were written for
  that and still bind legacy weeks.
- WHAT THE BANK IS FOR: the 9 already-taught morphemes students need for the WORDS TO
  SPELL REVIEW list on the very next slides. Students copy the 9 boxes into the top of
  their page and then spell those 10 review words underneath. Supporting that list is
  the bank's only job. It is not a preview, not a summary of the week, and not a
  general morpheme poster.
- NEVER PUT THE DAY'S FOCUS MORPHEME IN THE BANK - AND NEVER A MORPHEME THIS WEEK HAS
  NOT TAUGHT YET. Not on a `new` day, not on a `review` day, not on Friday. Monday's bank
  may not carry Wednesday's morpheme, whatever the spelling list looks like. This is the
  most-reported Sound Bank defect (the school reported it again in Aug 2026: "before
  learning the new morphology it is in the morphology bank") and it has two separate
  causes, both fatal:
  - On a `new` day the Sound Bank slide sits roughly a dozen slides BEFORE the morpheme
    card. The morpheme has not been taught yet. Banking it hands students a spelling
    they have never been shown, pre-empts the reveal the whole New Morphology section is
    built around, and burns one of nine slots on a morpheme that no review spelling word
    uses. "It gets taught later today" does not make it taught now - the students are
    looking at the bank NOW.
  - On a `review` day the morpheme was taught yesterday, so printing it defeats the
    retrieval that is the entire point of day 2 - students spell it from memory.
  The NEW morphology spelling words (3c) never need the bank: they come straight off the
  grid the class has just read, with the morpheme card still on screen.
- HOW TO DERIVE THE 9 - do this mechanically, never by eye:
  1. Write out the day's 10 Words to Spell Review.
  2. Under each word list every morpheme in it that the class has ALREADY been taught.
  3. Strike the day's focus morpheme from that list if it appears at all.
  4. Rank what remains by how many of the 10 words need each morpheme, and take the top 9.
  5. If fewer than 9 distinct morphemes survive, fill the remaining slots with the most
     recently taught morphemes of the types already present, so the type grouping stays
     intact.
  A morpheme that a word on today's page actually uses ALWAYS outranks one that no word
  on the page uses. The rejected junct deck banked `junct/join/joint` (needed by zero of
  its ten review spelling words) while `-ite`, `pict`, `-ule` and `vore` - each needed by
  a spelling word on that same page - were left out entirely.
- AGE IS NOT A CRITERION. The bank is a scaffold students COPY before they spell, not a
  retrieval test, so a Term 1 morpheme earns a slot exactly as readily as last week's -
  what decides it is whether a word on that page needs it. Step 5's "most recently
  taught" preference governs FILLER slots only, the ones no word on the page uses; it
  never outranks a morpheme a spelling word actually needs. Do not swap an old morpheme
  out, and do not re-teach it first, because it feels distant: if a student is shaky on
  it, 2e's in-the-moment correction is the whole response.
- SELF-CHECK BEFORE YOU BUILD: for each of the 9 boxes name the review spelling word it
  serves. Any box you cannot tie to a word is a wasted slot; swap it for a morpheme one
  of the ten words needs.
- THE BUILDER ENFORCES THIS. A bank box matching the day's focus morpheme (any variant)
  is a fatal WARN, and so is a box or review card matching a morpheme that a LATER
  session in the same week spec teaches. A box that appears in none of the day's review spelling words is a
  non-fatal NOTE, because the match is a substring test and an unusual surface spelling
  can read as unused when it is not. Answer every NOTE - usually by swapping the box;
  occasionally the honest answer is that the day needs only 8 morphemes and the 9th is
  step-5 filler.
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

## 2d. Words to Spell Review (one reveal slide each)

- v1.7 COUNTS: Tuesday 6-8 review words (the 4 new words follow), Friday 6-10. Weighting
  stays most-recent-first: about half last week, then two and three weeks ago, then one
  older word. Make them HARDER, not more (Yoshimoto's upper-level slide: "spelling now
  needs to be more complex" - a word like predawn where only aw and pre- are in the
  bank). Legacy weeks used ten every day.
- LEGACY weighting for ten: 4 last week, 2 two weeks ago, 1 three weeks ago,
  1 four+ weeks ago, 2 earlier this week (Monday: give those to last week too).
- Must be DIFFERENT words from today's Words to Read Review (reading and spelling are
  different skills; overlap wastes the slot). Words previously READ on earlier days are
  excellent spelling candidates now.
- The morpheme must contribute its meaning in every chosen word.
- Spelling scope: the rest of each word must be spellable - check phonograms/patterns
  against what the group has been taught (OG sequence). When in doubt, choose the more
  decodable word.
- V1.8 AFTER-CHECK RULE (James, 22 Oct 2026: the cus, pon and -ify questions "made no
  sense"). Only `Sentence:` is required. Add an `After checking:` / `Answer:` pair ONLY
  where the part's meaning plainly shows in the word (astrophysics, hypersonic,
  territory) - never on words whose parts have drifted from their meaning (excusable,
  proponent, justify). At most 3 per session; the builder warns on more. The notes no
  longer carry a "Click to reveal answer." line - a blank line separates the sentence
  from the check.
- v1.7 and earlier: every word's notes carry THREE things, all required (the builder
  warns on gaps):
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
  - Allow brief connective tissue when it makes the words sound human aloud:
    `Okay, now which part of module tells you it is a small piece?` is fine. Do not add
    empty address or management chatter (`Alright boys and girls, who can tell me...`).
    Scripted, natural, concise and immediately sayable - that is the bar.
  - Keep the whole set consistent: every word on the day uses the same voice.
- THE ANSWER MUST ADD SOMETHING THE PROMPT DID NOT. The blunt form of this defect is
  above - the prompt states the answer and the answer repeats it. The subtle form
  survives a quick read and still teaches nothing: the prompt hands over the root
  meaning, and the answer is that same meaning wearing a sentence. `just means right.
  What is justice?` / `The state of things being right.` The student has learned
  nothing they were not just told. Cover the answer and ask what NEW information it
  carries - a different part of the word, a reason, an application, a contrast. If the
  answer is the prompt rearranged, change the QUESTION, not the answer:
  `just means right. Which taught ending turns it into a state?` /
  `The -ice ending, the same one in service. Justice is the state of being right.`
- THE QUESTION MUST AIM AT A TAUGHT PART, OR CARRY THE UNTAUGHT PART'S MEANING WITH IT.
  An `After checking:` prompt is a STUDENT-SUPPLIED answer, so the taught-morpheme gate
  in 2f binds it - and this is the single place that gate is most often missed, because
  you are writing the question here and the gate is written over there. A student may
  never be asked to produce the meaning of a morpheme they have never been taught.
  Three shapes, two safe and one banned:
  - SAFE, aim at a taught part: `capt means hold. Which taught ending turns it into a
    state?` / `The -ity ending, the same one in gravity.` Best of all when that part is
    in the same day's Sound Bank, because it ties the bank to the list it exists to serve.
  - SAFE, hand the untaught meaning over inside the question and ask the student to
    locate or apply it: `Which part of chronograph means write or record?` /
    `The graph part.` The student supplies the location, never the meaning.
  - BANNED, ask a student to supply an untaught part's meaning. Real defects caught in
    review: `What does the a at the front do to it?` (asynchronous, `a-` on no card);
    `What does con add?` (conjoin, `con-` untaught); `Which two parts tell you they had
    to leave?` (evicted, one of them is `e-`). Each was rewritten into a safe shape.
  Read the prompt back and ask: can a student reach this from today's card, the Sound
  Bank, the review-10, the term timeline, or ordinary Grade 5/6 English? If the only
  route runs through a morpheme they have never met, it is a defect - and a common
  Latinate prefix is not "ordinary English" unless the question itself supplies it.
- The builder centres every word at the same position with reveal-on-click and a blank
  screen between words (one slide per word) - this fixes the alignment/reveal feedback
  from the team email; you only supply words, sentences and prompts.
- V1.9 MEANING PICTURE (from the Challenge team's planner decks). Every word carries
  `picture`, and the builder shows it under the word ON THE WORD'S OWN CLICK (Appear,
  With Previous), so the picture never gives the word away before students write it.
  `picture` is a pictogram name (`"plane"`) or a react-icons `pack/Component` name
  (`"fa6/FaShip"`, `"tb/TbTelescope"`, `"md/MdSportsSoccer"`) - thousands of concrete
  objects - optionally `{"icon": ..., "color": "RRGGBB"}`. Pick the thing the word's
  SENTENCE puts in the student's head (exporter -> a cargo ship, supporter -> a footy
  ball), never a picture of the morpheme. An unknown name fails the build.
- THE PICTURE MUST BE CLEAR TO A STUDENT WITH NO EXPLANATION (James, 30 Oct 2026, on a
  bowl standing in for omnivore: "makes NO sense"). Cover the word and ask: would an
  11-year-old name this meaning from the picture alone? If one icon cannot carry it,
  draw a short scene: a LIST of two or three icons side by side - omnivore = bear,
  berries, fish (eats plants and meat); exporter = wheat, arrow, cargo ship (goods
  going out); astronomer = telescope, moon. `{"icon", "color", "highlight": {"box":
  [x0, y0, x1, y1], "color"}}` shades one part of an icon - territory = the Northern
  Territory shaded on a pale map of Australia. Give each icon its natural colour
  (wheat gold, fish blue) so the scene reads at a glance.
- IF NOTHING CAN, SHOW NO PICTURE. `{"none": "reason"}` keeps the slide to the word
  alone (excuse: a reason someone gives cannot be drawn without explaining it). A
  confusing picture is worse than none. At most 2 per list - more means the list is
  too abstract for pictures, so rethink the words. The builder NOTEs each one.

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
- Notes call the session a "session", never a "lesson" (school feedback, Oct 2026: staff
  search decks for "Lesson", and every note that uses the word buries the slide they want).
  The word may still appear as content, in an example sentence or a question about school
  life ("Which word might you hear in a maths lesson?").
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

## 2f. THE TAUGHT-MORPHEME GATE (what a student may be asked to SUPPLY)

The taught-only rule in 2b gates untaught METALANGUAGE. This gates untaught MORPHEMES.
It applies to every segment of every deck and it is the rule most often broken in the
New Morphology You Do.

TEACHER-DELIVERED IS NOT THE SAME AS STUDENT-SUPPLIED. The gate binds only the second.

- TEACHER-DELIVERED is anything the teacher says aloud with the answer in front of them:
  the submarine lines in `wtr_new_notes`, an `Answer:` after a spelling reveal, a MODEL
  beat, a dictation focus line. These may name ANY morpheme, taught or not - the teacher
  supplies the meaning in the same breath. That is how students meet new parts, and it
  is fine.
- STUDENT-SUPPLIED is anything a student must produce, decode, or be marked on: a You Do
  item, an `ASK:` with its `EXPECT:`, a check-slide question, an `After checking:`
  prompt, an extension, a morpheme sum they have to complete.

THE RULE: every student-supplied answer must be reachable from
  (a) TODAY'S FOCUS MORPHEME - its card is on screen, its meaning has just been taught,
      it has a copy-into-books slide and a memory hook; plus
  (b) THE TAUGHT HISTORY - the review-10 cards, the sound bank, and anything earlier on
      the term timeline; plus
  (c) ORDINARY GRADE 5/6 ENGLISH - everyday words, common sense, general knowledge.
Nothing else. If reaching the answer requires a morpheme outside (a) to (c), the item is
a defect - not a stretch, not a challenge, a defect. The student cannot get there and
knows it.

ON THE YOU DO SLIDE, (b) MUST BE PRINTED. Since v1.5 the You Do is a one-recall task
(3d): the taught history still makes a part legitimate to use, but the student does not
retrieve it - its meaning is printed on the slide face. `ASK:`/`EXPECT:`, check-slide
questions, `After checking:` prompts and extensions keep the full (a)-(c) reach.

HEARING A MORPHEME ONCE IS NOT BEING TAUGHT IT. A Words to Read grid holds 9 or 12 words
and the teacher voices a HANDFUL of the script lines as the words appear - never all
twelve, never with the class rehearsing them, and never in a form anyone could revise
from. Therefore:

- No morpheme becomes "taught" because it appeared in a grid script line, a spelling
  `Answer:`, a dictation focus or an extra task earlier in the SAME deck.
- Never build a later task on recall of a specific grid word ("the one we read before",
  "remember injunction"). Assume the student saw the word for two seconds and heard
  nothing about it.
- The only thing the New Morphology section reliably teaches is the focus morpheme
  itself, and that is because it gets a card, a copy slide, a hook and a whole grid.
- A word may still be the ANSWER to a task, as long as the student can BUILD it from
  taught parts rather than retrieve it from memory.

WHEN AN ITEM NEEDS AN UNTAUGHT PART, PRINT THE MEANING ON THE SLIDE. First choice is
always to swap the item for one built from taught parts. If a word genuinely earns its
place, put the untaught part's meaning on the slide face inside the item itself, so the
student READS it instead of remembering it:
`4. in- (onto) + junct + -tion = ?`. A meaning printed on the board is knowledge the
student has; a meaning mentioned nine slides ago is not.

An inline gloss costs characters, and item lines do not shrink or clip - they wrap into
the line below (10c). Count the line before you add one. If the gloss will not fit,
that is the answer: swap the item for one built from taught parts.

Defects, all taken from shipped decks:

- A You Do item `in- + junct + -tion = ?` where `in-` meaning "onto" is untaught and had
  appeared exactly once, inside one grid script line the teacher may never have read.
- An extension asking students to predict what `adjunctive` means when `ad-` has never
  been taught - that is a guess dressed as an extension.
- An `After checking:` question whose `Answer:` turns on a suffix that is on no card the
  class has met.
- A check-slide `ASK:` whose `EXPECT:` names a morpheme meaning the students were never
  given.

WHAT THE BUILDER CATCHES, AND WHAT IT CANNOT. The builder scans the `new_morph_activity`
`example`, `items` and `check_items` - every prompt, including two packed on one line -
and fails the build on two things. A MEANING SUM: an operand or a `= ?` left side that
is a phrase (`small + look at = ?`, `bad function = ?`), or a check-slide sum whose
operand does not survive inside the answer word (`small + wave = microwave`). An
UNPRINTED PART: any affix (`in-`, `-tion`), catalogue root or taught morpheme a student
handles, on either side of the `=`, that is not today's focus morpheme and has no
`part (meaning)` gloss anywhere on the slide face (title, rule, example, items, footer).
Round items (v1.6) are scanned slide by slide: a gloss counts only on the slide that
uses it, so a part glossed on item 1 must be glossed again on item 3.
`-s`, `-es`, `-ed` and `-ing` are universal inflections. A glossed part that is NOT in
the taught set builds with a NOTE, because printed is reachable but a taught part is the
first choice. It CANNOT check: sorting or sentence tasks that name no parts, `EXPECT:`
answers, `After checking:` prompts, or extensions written as prose. Those stay your job
(4d step 8). When the gate fires, there are exactly three honest fixes: rebuild the item
from today's morpheme and everyday words, print the missing meaning on the slide face,
or - for a NOTE only - swap the glossed part for one the class has been taught.

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
- V1.8 MEMORY HOOK: SHORT, SHARP AND SHINY, ON THE CARD WITH A PICTURE (James, 22 Oct
  2026: the old hooks "suck"). Author it as the teach day's `memory_hook`:
  `{"text": "Eject throws the pilot out!", "picture": "plane"}`. The text is 8 words or
  fewer, one line, a concrete picture students can see (an action, an object), never the
  definition reworded. `picture` names a pictogram from `themes/core/pictograms.js`
  (`node -e "console.log(require('./themes/core/pictograms').listPictograms().join(' '))"`);
  when none fits, draw a new one in that file first, then use it. The builder prints the
  hook and picture on the new-morpheme card under the morpheme and puts
  `MEMORY HOOK: ...` as the FIRST line of the card, copy-slide and grid notes. Do not
  also write a `Memory hook:` line in `wtr_new_notes` (the builder warns). The
  new-morpheme header notes carry one line, no PROCEDURE block.
- v1.7 and earlier: THE MEMORY HOOK BELONGS ON THAT COPY SLIDE TOO. The builder lifts the `Memory hook:`
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
- v1.7 MONDAY ROUTINE, Yoshimoto's own steps for a new morpheme (training deck,
  "Introduction of a new suffix" and "Introduction of a Latin root"): the card, the
  copy slide, then students WRITE IT THREE TIMES saying it (the builder adds this line
  to the copy-slide notes), then the class BRAINSTORMS words they already know with it
  before the grid shows any. Author the brainstorm as the spec's `brainstorm` string -
  one ASK line with think time, the routine and EXPECT (`"What words do you already
  know with port in them?" 20 sec, turn and tell. EXPECT: transport, airport,
  passport.`). The grid then confirms and extends what they offered.
- v1.7 YOUR KEYWORD AND ONE SENTENCE close Monday (`keyword_sentence`, a magenta You do
  slide). Kate Watt: show the card's keyword and say why it is a good one - concrete,
  every part visible (homeless, telegram, monobrow) - then each student picks the word
  that will help THEM remember, writes it in their book and stars it, and writes one
  sentence with one of today's words (Yoshimoto's step 11). A slow student may simply
  take the card keyword. Keywords are never quizzed afterwards (2a).
- PICK THE WEEK'S MORPHEME FOR USE, NOT TIDINESS. Kate: "pick only the words that we as
  adults use", and the training deck's "relevance is key - consider what is being
  taught across all subjects" (perimeter in maths, omnivore in science). Concrete,
  high-use roots come first (port, ject, struct, tract, spect, dict, rupt). The user's
  request names the morpheme; when it names a rare one, build it and say in the summary
  which high-use root would serve the class better.

## 3b. Words to Read (grid, reveal one word per click)

- v1.7 THE USEFULNESS TEST COMES FIRST. Kate Watt on astrolabe: "Do we use it as an
  adult? No, we don't. Don't touch it." and "if you don't have ten, do six because it's
  more relatable." Every grid word must pass "would an adult use this word, or would a
  Grade 5/6 student meet it in a book, a science lesson or the news?". A grid of 6 is
  legitimate (6, 9 or 12 - the three columns must balance). One deliberately strange
  word for the brighter students is allowed, flagged as such in its script line ("Here
  is a funny one..."), never more.
- v1.7 WORDS WITH AN UNTAUGHT PART GO LAST (Kate: "astronaut ... naut hasn't been
  taught ... put those last few at the end"). Order the grid so every word whose parts
  are all taught comes first; the last one or two may carry one untaught part, and their
  script line gives that part's meaning in a breath ("naut means sailor - a star
  sailor"). Talk about the meaning of each word as it appears (Kate: the Bentleigh West
  lesson "discusses the meaning of each word first").
- THE CORE CONCRETE ROOTS ARE A RECAP WEEK (Term 4 2026 on: port, ject, struct, tract;
  Sabrina, 13 Oct: "recap and use in our words"). Bentleigh West's Foundation-to-Year-6
  map teaches these in Years 1-2 on top of re-, ex-, in-/im-, trans-, sub-/sup-, de-,
  -able and -er. This cohort started OG in Year 4-5, so those prefixes and suffixes are
  not on its card record, and almost every word in these families carries one. For these
  roots only, those Year 1-3 affixes do not count toward the untaught-part cap: every
  grid line glosses them in a breath, and every You Do slide prints them (`re- (back)`).
  The builder raises a NOTE for each printed one; that NOTE is expected, so say so in the
  summary. Never add them to `taught_morphemes` unless a card actually taught them.
- GRID SIZE MUST BE 9 OR 12 (v1.7: 6, 9 OR 12) - divisible by 3. The template's three columns centre
  independently, so 8, 10 or 11 words leave a short third column floating out of
  alignment at a different size. "Aim for 12, floor about 8" is about morpheme
  integrity, not layout; the layout constraint is absolute.
- HARD CAP: NO GRID WORD LONGER THAN 12 CHARACTERS. The builder sizes the WHOLE grid to
  its single longest word and the gate fails below 27 pt. The arithmetic is fixed and
  identical for a 9-word and a 12-word grid: 12 characters renders at 28 pt, 13
  characters drops to 26 pt and the build stops. `recommendation` (14) and
  `justification` (13) both failed this and had to be swapped after the fact. Check the
  longest word BEFORE you write the script lines, because replacing a grid word means
  also updating `words_to_spell_new`, that word's line in `wtr_new_notes`, its
  `associated_word_exceptions` entry, and any You Do item or extension that names it.
- WORD SELECTION ORDER. Apply these in order for EVERY word list in the deck - the new
  grid, the new spelling four, review reading and review spelling:
  1. TRANSPARENCY - the morpheme is visibly doing its job and the whole-word meaning is
     genuinely the sum of the parts. Reject opaque words outright (4c).
  2. TAUGHT-MORPHEME STACKING - the word also carries previously taught prefixes, roots
     or suffixes, weighted to the most recently taught.
  3. GRADE 5/6 COMPREHENSIBILITY - a 10-12 year old can picture the word, or one plain
     sentence makes it picturable. Stretch words are wanted (insubordinate, ordinance);
     legal-register curios and archaic forms are not.
  4. CATALOGUE MEMBERSHIP (4a-4c).
  The order matters. A simpler, more transparent, better-stacked word BEATS a catalogue
  word that scores worse on 1-3. Section 4c ground 3 exists precisely to authorise that
  substitution - the catalogue is the default family, not a ceiling on teaching quality.
- MORPHEME STACKING IS THE FIRST SELECTION RULE ONCE TRANSPARENCY IS SETTLED. Where the
  family gives you a choice, ALWAYS prefer the derivative that also carries a previously
  taught prefix, root or suffix, weighted to the most recently taught. `microscopic`
  beats `microchip` once `-ic` is on the cards; `malnutrition` beats `malaise` once
  `-tion` is taught. Two taught morphemes in one word is worth more than a slightly
  neater single-morpheme word: it is retrieval practice and it teaches students that
  words are built, not memorised.
- THE SIMPLER SIBLING WINS. When a family offers a short transparent form and a longer
  ornate one built on top of it, take the short one: `rejoin` (re- = back, taught) over
  `rejoinder` (`-der` is taught nowhere); `adjoin` over `conjunctive`; `disjoint` over
  `disjuncture`. Every extra untaught part is a part the teacher must gloss and the
  student cannot reuse. A grid of plainer words that all click teaches the focus
  morpheme better than a grid half-full of words ending in parts nobody has met.
- CAP THE UNTAUGHT-PART COUNT ACROSS THE GRID. At most 2 of the 9 or 12 grid words may
  contain a part the class has not been taught, and each of those must still be a word a Grade 5/6
  student could plausibly meet in a book, a science lesson or the news. If the family
  cannot fill 12 within that cap, build a 9-word grid - never pad to 12 with curios.
- THE SCRIPT MUST BE DELIVERABLE COLD (section 7's teach-it-cold standard). The teacher
  reads a word's line as that word appears, so a line MAY name an untaught part and
  supply its meaning in the same breath - that is the teacher-delivered side of the gate
  in 2f and it is exactly how new parts get met. What is never allowed is treating that
  one mention as teaching: nothing later in the deck may require students to supply a
  meaning they only heard here (2f).
- Layout is automatic and ALWAYS the template's 3-column grid (never reshape it to two
  columns - the school wants the 3-column look preserved). The builder widens all three
  columns, removes hidden text insets and centres the block; 10-12 words also stretches
  the columns down the page for a fourth row. Every term must remain unbroken at 27 pt
  or larger. Reveal stays one word per click.
- V1.8 VET EVERY EQUATION (James, 22 Oct 2026: "word equations sometimes don't match the
  answer"). Before the build, read each `words_to_read_new` meaning against its word:
  every part in the word is named, no part is named that is not in the word, and the
  meaning after the colon is built from exactly those parts. Then ask whether an adult
  uses the word; if not, swap it.
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
- v1.7 and earlier: end the script with a `Memory hook:` line - one sticky retention device students
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
  SMALL FAMILIES: when the card family cannot supply a fresh review grid inside the
  12-letter cap and the two-untaught-part cap, repeat yesterday's words and add plural or
  inflected forms rather than reaching for curios, and flag it in the summary. omni-
  (five card words) did this in Term 4 Week 1: the Thursday grid repeated six of
  Wednesday's nine words. hyper's printed family is mostly 13-15 letters
  (hypersensitive, hypercritical, hyperventilate), so its grids lean on transparent
  hyper + everyday-word compounds (hyperlink, hyperspace, hypermarket).
  GREEK AND LATIN COMBINING FORMS (astro, terr, and the rest of the Greek Combining
  Forms and Upper Level Latin Roots books) almost never join to an everyday word: their
  partners are other untaught parts (naut, nomy, -arium, -ory). The two-untaught-part
  cap cannot be met, so choose words a Grade 5/6 student meets in science, geography or
  the news, have every script line name and gloss the partner part, and print that
  part's meaning wherever a You Do uses it. Term 4 Week 2 did this for astro (astronaut,
  astronomer, asteroid, asterisk, astrolabe) and terr (terrain, territory, terrarium,
  terrier, subterranean). astro's twin spelling aster (asteroid, asterisk) is the same
  Greek word for star; say so in the script line. The scroll prints the Latin root as
  terr (land), not terra - James's request said terra, the card label stays terr.

## 3c. Words to Spell (4, revealed per click)

- Exactly 4, and all 4 MUST come from today's Words to Read grid (this is the one place
  read/spell overlap is required - students spell what they just decoded). v1.7: the 4
  run on TUESDAY and come from MONDAY's grid - students spell what they read yesterday,
  with the recap card still fresh. The builder checks they are on that grid.
- Apply the morpheme-stacking rule from 3b: of the eligible grid words, choose the four
  that also carry recently taught prefixes/suffixes wherever the family allows it.
- PREFER THE FOUR WHOSE PARTS ARE ALL TAUGHT - today's morpheme counts as taught, its
  card is on screen. Students should be able to BUILD each word from parts they hold. A
  spelling word containing a part nobody has met is a memorisation task wearing a
  morphology hat, and it is the one thing this section is not for.
- V1.8 NOTES SHAPE (James, 22 Oct 2026): no "SAY each word, students write it in their
  book, then click to show it." line (the builder warns) - the header carries the
  routine. A blank line goes before the `EXTENSION:` line, and the extension is one of
  the three types in 3e.
- v1.7 and earlier: rotate the early-finisher extension in the notes; never the same one
  two days running. Menu: use two target words in one sentence with a subordinate
  clause; write the meaning in your own words using the root; find and record an
  antonym; turn the word into a question; add a prefix/suffix and explain the new
  meaning.
- Extensions must be SELF-CONTAINED: if the task names a concept (subordinating clause,
  adverbial phrase), include a one-line reminder of what it is right in the note, so no
  student stalls waiting for the teacher. Then add an `EXTENSION:` step meeting the
  standard in 3e.

## 3d. New Morphology You Do activity (legacy new/review days)

v1.7: the per-day You Do is retired (Kate Watt: on a new day it repeated the grid
discussion the class had just had). The morphology You Do is now the whole of
Wednesday (3f). Everything below still governs legacy weeks, and its standards - THE
TASK MUST NEED TODAY'S MEANING, ONE RECALL, rounds with the answer on the click, the
achievability bar - govern every v1.7 build-day round too.

After Words to Spell and before Review Learned Words, every new/review session gets a
dedicated Yoshimoto-style You Do: a magenta section header (auto-retitled `New
Morphology - You Do` / `Review Morphology - You Do`) followed by a ROUND (v1.6): a
"One done for you" slide, then four item slides, one question each. Every slide carries
the same rule banner, the item as the hero, up to three choice chips, and a green answer
bar that the teacher's click reveals; the worked slide shows its answer from the start.
The generic "write a fun sentence" is the extension note, not this. Supply it as
`new_morph_activity` with a `round` (schema in 9); the builder warns when a new/review
session lacks one. The older single task slide plus `Tick it or fix it` check slide still
builds for old specs, but new weeks use rounds.

- THE TASK MUST NEED TODAY'S MEANING (v1.6, the test every item must pass). Cover the
  card meaning and ask: could a student still get this right? If yes, the item teaches
  nothing. `hyper + link = ?` failed it (Term 4 Week 1, flagged by James before Kate's
  review): the answer is the two printed parts said together. The student must use the
  morpheme's MEANING to decide, and must HAVE everything else they need - today's card,
  the printed parts and everyday Grade 5/6 English.
- ROUND SHAPE, why it is the default: every student answers every item (boards or
  workbooks), the teacher sees the whole class before clicking, and a wrong answer is
  fixed seconds later while the reason is still in front of them - the immediate
  corrective feedback OG is built on. It is also the easiest shape for a teacher who did
  not plan the session: read the item, run the cue, scan, click, read the green reason.
  Time: about 20 seconds on the worked slide and 35 seconds an item, inside 3 minutes.
- The task uses today's morpheme and is done on MINI WHITEBOARDS or in OG WORKBOOKS -
  the `routine` chip says which (`Boards - write it, show me`, `OG workbooks - write it`).
- SPREAD THE ACTIVITIES ACROSS THE WEEK: every new/review day in the same week uses a
  DIFFERENT activity type from the menu - with the standard Mon/Tue/Wed/Thu pattern
  that means four different activities in a week, and no type repeats within the
  week. Vary the surface too (whiteboards one day, OG workbooks another). The menu, with
  the day each type suits best (Term 4 Week 2 ran all four, astro then terr):
  - WHICH WORD FITS? (new day). The hero is a sentence with a blank; three chips are
    family words, each with its OTHER part glossed (`asteroid` / `-oid (like)`). The
    student recalls the morpheme's meaning, adds the printed part, and picks the word
    the sentence needs. Build the wrong chips from the item's trap: the edge item's
    sentence mentions sailors and the answer is astrolabe (a star-taker), not
    astronaut (a star sailor).
  - YES OR NO - PROVE IT (review day). The hero is a statement using a family word
    truly or falsely (`An asterisk is a giant star in deep space.`); the student writes
    YES or NO and must be ready to prove it from the parts. Mix YES and NO so guessing
    does not pay, and print any needed parts in the rule banner (`Parts: -isk (little)`).
  - MATCH THE MEANING (new day) - the Yoshimoto worksheet's own task ("Match the words
    with the meanings", on the terr and anti pages). The rule banner is a word bank with
    ONE MORE word than there are items, so the last item is never a forced leftover;
    the hero is a kid-friendly meaning; the student writes the matching word.
  - REAL OR LOOK-ALIKE? (review day). The hero is a sentence with one bold word; the
    student decides whether the letters are really the morpheme doing its job
    (territory, terrarium) or a look-alike (terrible, terrific - from terror, fear). This
    is the OG "morpheme must do its job" test (2b) handed to students, and it suits any
    morpheme with common look-alikes (mini in minister, ped in pedal). The test is
    always the MEANING ("is land part of what the word means?"), never a physical
    stand-in: Term 4 Week 2 first printed "can you stamp on it?", which marks terrarium
    and terrier wrong. A stamp can stay as the memory-hook action, not as the test.
  - Literal meanings and a sentence that SHOWS the meaning stay available, but only with
    a fixed answer a round can reveal; open writing belongs in EXTENSION.
  - WORD BUILDING is narrowed (v1.6): only when the join itself needs a decision the
    class has been taught - a spelling change (`cooperate + -ive = cooperative`,
    `hyper + extend + -ed`) or choosing between two endings. Joining a morpheme to an
    everyday word with no change is retired.
- MEANING SUMS ARE RETIRED (team decision, Sept 2026). The tachistoscope shape - the
  teacher gives the meanings, students produce the word (`small + look at = ?`,
  `cannot be conquered = ?`, `house + study of + relating to = ?`) - is no longer a
  You Do type. It asked students to hold three or four parts they could not all
  remember, and it kept reaching for parts nobody had taught (`-al` in ecological).
  The builder fails any item or check item of that shape. Meaning sums may still be
  worked as teacher-led talk on the grid slides; they are never independent work.
- ONE RECALL. The only part a student supplies from memory is TODAY'S focus morpheme.
  Every other part an item touches is one of: a universal inflection (`-s`, `-es`,
  `-ed`, `-ing`), an everyday whole word (`climate`, `function`), or PRINTED WITH ITS
  MEANING on the slide face - in the rule banner, the worked example or the item
  itself, in the form `part (meaning)`: `dis- (apart)`, `vore (eat)`. Taught earlier
  in the term does NOT count as printed: the review cards and the Sound Bank are where
  retrieval happens, the You Do is where today's morpheme is used. When several old
  parts are needed, put a parts bank in the rule banner (`Parts: dis- (apart), -ive
  (relating to)`) rather than glossing every item (10c character budget). The builder
  fails an unprinted non-focus part.
- SAME TASK FOR EVERYONE. No tiers on the slide face. Differentiation lives in the
  notes: `EXTENSION:` for early finishers, `HELP:` for the stuck.
- Every activity involves writing; every activity is startable without teacher help.
- ACHIEVABILITY BAR (non-negotiable - the school has rejected You Dos that failed it):
  - The task description on the slide is EXPLICIT: it says exactly what to write and
    where, in words a mixed-readiness 10-12yo can follow with no teacher explanation.
    A terse label (`Write R or P`) is not a task description.
  - The slide carries ONE COMPLETED WORKED EXAMPLE of the very task type, marked
    `One done for you: ...`, before students start. Use the hero card for it. A You Do
    whose first exposure to the task shape is the task itself is a defect.
  - PUT THE WORKED EQUATION ON ITS OWN LINE. Write the hero as
    `One done for you:\nmal- + treat = maltreat` - the label on line 1, the
    equation alone on line 2. The builder emits each authored line as its own centred
    paragraph, so an equation reads as an equation instead of trailing off the end of
    a sentence. The same applies to any `word = label` model on a grammar slide.
  - The numbered prompt lines under the hero are centre-aligned by the builder so the
    block reads as a set. Keep each line short enough not to wrap (see 10c).
  - ONE CLEAR ANSWER PER ITEM. Every item must have a single correct answer a student
    can reach from that day's teaching and verify on the check slide. An item with two
    defensible answers (`pottery = P or R/P`) or a judgement call about etymological
    transparency is teacher-led We Do discussion material, never independent work.
  - SOLVABLE FROM THE BOARD (2f + ONE RECALL above - this is the You Do's hardest bar
    and the one most often failed). Every item must be reachable from today's focus
    morpheme plus the parts printed on the slide plus everyday English. Do NOT lean on
    a part taught weeks ago, on a part the class has never been taught, or on recall of
    a particular word from the grid the class just read - the teacher voices a handful
    of those twelve lines, not all of them, and no student revised from it. If a part
    is needed, print its meaning on the slide face
    (`4. in- (onto) + junct + -tion (the act of) = ?`).
  - WALK EACH ANSWER BACKWARDS BEFORE YOU SHIP IT. Write out the answer word and ask:
    could a student who met today's morpheme five minutes ago produce this from the
    board in front of them right now? If the honest answer is "only if they remember it
    from the grid", "only if they already knew the word" or "only if they remember what
    -al means", replace the item or print the part. This check is especially unforgiving
    on a `new` day, where the students have held the focus morpheme for about five
    minutes.
  - Prefer CONCRETE BUILD/USE tasks (join printed parts into the word, write the sum
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
  Purpose: ... [You Do | Word building | Morphology]
  ANSWERS:
  1. mal- + function = malfunction
  2. mal- + nutrition = malnutrition
  3. mal- + formation = malformation
  4. mal- + treatment = maltreatment
  EXTENSION: e.g. maltreated = mal- + treat + -ed, was treated badly.
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
- IN A ROUND the prompt ladder lives in each item's SCAN line: `80%+ -> click to
  reveal. Less -> [one move that adds something the slide does not say], then re-ask.`
  (`Less -> ask whether the blank is a person, a science or a tool, then point at labe
  (taker), then re-ask.`) The TRAP names the near-miss chip and its fix.
- The check slide's `ASK:` must be a genuine thinking question, not a spot-the-obvious
  (`Which part of every answer means small?` when every answer visibly starts with
  micro). Ask something that makes them USE the pattern: which of these two words
  changes meaning if you drop the prefix, why does this word need three parts, which
  answer would still make sense with a different root.
- IN A ROUND, the green answer bar IS the tick-it-or-fix-it slide: each item's
  REVEAL beat says "Tick it or fix it", then the reason. Each item's notes open with
  its own `ANSWER:` line; the worked slide's notes carry the `ANSWERS:` block, the
  `EXTENSION:` and the `HELP:` for the whole round. The rules below about a separate
  check slide apply to the older single-slide shape only.
- A fixed-answer single-slide activity MUST be followed immediately by a duplicate answer slide
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

V1.8: THREE TYPES ONLY (James, 22 Oct 2026). Extensions must be small, extra and
challenging, and built on what the class already holds. Use one of these and nothing
else:
1. LONGEST WORD HUNT - "Build the longest real ject word you can and write its parts."
   (Wednesday's matrix extension was "a ripper".)
2. BEAT THE NUMBER - "Write 5 more real ject words in 2 minutes."
3. TWO IN ONE - "Write one sentence that uses two of today's words and still makes
   sense."
Retired as confusing: "find two words on the board that share a part" and "change one
part of export to make three new real words". The rules below (startable alone, answer
supplied, taught parts only) still apply.

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
  - Morphology examples: core joins two printed parts (`micro + scope (look at) = ?`),
    extension builds a four-morpheme word from a printed parts bank and writes the
    full sum (`microbiologist = micro + bio (life) + -ology (study of) + -ist
    (person)`); core writes the sum beside a
    derivative, extension takes an UNSEEN word from the same family and predicts its
    meaning from the parts alone, then checks it.
  - Grammar examples: core labels phrase or clause, extension combines a subordinate
    clause and a main clause into one correctly punctuated sentence and then moves the
    subordinate clause to the end and repunctuates it.
- It must be STARTABLE WITHOUT THE TEACHER, and its answer goes in the `ANSWERS:`
  block so a teacher can confirm it.
- THE SPELL-SLIDE EXTENSION CARRIES ITS OWN ANSWER. The 3c `extension` field is a bare
  notes string with no `ANSWERS:` block to fall back on, so the confirmable-answer rule
  has to be met inside the note: finish it with one `Answer:` line giving a worked
  example. `EXTENSION: Build the longest real mand or mend word you can, then write its
  full morpheme sum.` / `Answer: e.g. commandments = com (together) + mand (order) +
  ment (the thing) + s (more than one).` Without it the teacher is holding an open task
  they cannot mark.
- THE TAUGHT-MORPHEME GATE APPLIES TO EXTENSIONS TOO (2f). "Predict what `adjunctive`
  means" is not an extension when `ad-` has never been taught - it is a guess, and the
  fastest students are the ones who notice. Push the SAME taught parts further instead:
  more of them in one word, the full sum written out, the word used in a sentence that
  demonstrates the meaning, or an unseen word whose parts are ALL taught.

## 3f. WEDNESDAY: BUILD AND BREAK (v1.7 `build` day, 15-20 minutes)

Kate Watt: "Wednesday, let's build some words, guys ... join an old one." Yoshimoto's
own training deck shows the two halves: a word matrix (the rupt matrix) and a "break it
apart" table (Word | Prefix | Base/Root | Suffix). A Grade 4 model lesson in the same
deck adds the third move: "Please write the word that names the art of beautiful
handwriting." Three blocks, in this order, under one magenta `Build and Break` header:

1. WORD MATRIX (`build.matrix`, one slide). The week's morpheme sits in the centre box
   in its card colour with its meaning; prefixes the class has been taught sit left,
   suffixes right (in a core-concrete-root week, the Year 1-3 recap affixes from 3b
   count), EACH WITH ITS MEANING PRINTED (`trans- (across)`, `-able (able to)`)
   - the one-recall rule (2f) applies, and the audit fails an unglossed part. Students
   build as many REAL words as they can in their OG books and say what each means
   (`transportable - able to be carried across`). Differentiation is built in: how many
   words, and how many parts in one word (Kate: "they need that more of a challenge").
   - 3-6 parts a side. Choose parts that make real, useful words with this morpheme;
     include one or two from weeks ago so old morphemes get used (Kate's "join an old
     one").
   - Notes: the ROUTINE (OG books, 4 minutes), a TRAP (a made-up word that "sounds
     right" - the test is whether it is a real word you could use), and an `ANSWERS:`
     block below the divider listing the real words the matrix makes, with meanings.
2. BREAK IT APART (`build.rounds[0]`, a round of 4 + 1 worked). V1.8 (James: "tough to
   write in the format wanted"): every item carries `parts` (`["re-", "ject", "-ion"]`,
   at most 4) and the round carries `given: ["ject"]`. The slide draws one box per part
   with `+` between them, the week's root already filled in; students write ONLY the
   missing parts in their books, and the meanings arrive on the click's answer bar and
   in the teacher's voice. The worked item shows every box filled. The audit fails a
   v1.8 break round with an item missing `parts`. v1.7 wording follows: the hero is a whole
   word; students write its parts in their books (`trans + port + able`); the click
   shows the answer bar with each part's meaning (`trans- (across) + port (carry) +
   -able (able to)`). Kate: "do the opposite - instead of always building words, break
   them up, because they would go confid- and it's not." Mix a word from this week with
   one from an earlier week.
3. WRITE THE WORD THAT NAMES... (`build.rounds[1]`, a round of 4-5). The hero is a
   kid-friendly meaning (`Write the word that names a person who carries bags at a
   hotel.`); students write the word on whiteboards; the click shows the word and its
   parts. This is retrieval of meaning AND spelling in one move. Answers come from this
   week's grid and Tuesday's spelling - words the class has read and spelled - so it is
   recall, never a guess. It is not a retired meaning sum: the hero is a plain-English
   definition, never `carry + across = ?`.

All three use the 3d round standards: one question a slide, every student answers,
scan, click, tick or fix; `ANSWER:` first in each item's notes; the worked item carries
the round's `ANSWERS:` block, EXTENSION and HELP. Whiteboards for round 3, OG books for
the matrix and round 2.

V1.9 THIRD-ROUND ROTATION (from the Challenge team's activity menu). Round 3 is ONE of
three formats, rotated each week so no format runs two weeks in a row, and Kate's
`Write the word that names...` comes back at least every second week (the builder
NOTEs both against last week's spec):
- WRITE THE WORD THAT NAMES... - as above (meaning in, word out).
- COMPLETE THE SENTENCE (`bank` + `bank_mode` on the round). Each item is a sentence
  with one `___` gap; the bank sits on every slide under the fixed instruction `A word
  may be used twice. Some may not be needed.` - fixed so it never gives the bank size
  away. `bank_mode` varies from one use to the next: `extra` (one or two words fit
  nowhere), `short` (one word is used twice), `exact` (each word once). The worked
  item's answer is a bank word and counts. Every gap has ONE clear answer: a spare word
  must fit no sentence, and in `short` mode the repeated word must be the only word that
  fits both of its sentences. At 5/6, the sentence carries the meaning (`The nurse will
  ___ the vaccine into your arm.` - in, so inject), never a definition in disguise.
- CHOOSE THE BEST MEANING (`choices` on every item, exactly 3, fingers 1-3). The hero is
  a word (`projectile`), the choices are meanings, the answer is one choice word for
  word. THE BEST WRONG ANSWER IS BUILT FROM A REAL MEANING OF ONE OF THE WORD'S PARTS
  (`thrown back` for projectile - ject is right, the direction is wrong), so students
  must combine every part, not spot one. The third choice is plausible and wrong on the
  other part. Use it on words with 3 parts at 5/6, where combining is the skill.
The builder checks the mechanics (round order, bank maths, the `___` gap, 3 choices,
answer among them); you own the one-clear-answer and distractor bars.

## 3g. THURSDAY: SUFFIX PARTS OF SPEECH (v1.7 `apply` day, sentence test in v1.8)

V1.8 (James, 22 Oct 2026: Thursday was "BRUTAL" - the suffixes were old and forgotten,
and teachers could not explain why a suffix makes a noun). Rebuilt 23 Oct 2026 after
the frame version failed in Google Slides: the word was not on the slide, the I do
printed its answer, and the notes said "to supportive - no, very supportive - yes"
without ever saying why. The drill now reads the job from a sentence, the way a
teacher would explain it:
- CHOOSE 4-5 SUFFIXES, NO MORE: suffixes on this week's review cards first; when the
  cards hold none, taught suffixes that build this week's or recent words (Term 4 Week 4
  used -ity, -ive, -ciate, -tude). Noun, verb or adjective only; the builder warns on
  any other `pos_answer`.
- EACH CARD READS TOP TO BOTTOM: the suffix; "-ity means: state, quality"; the keyword
  sentence with the keyword in bold (`"sentence"`, required); the question "What job
  does gravity do in this sentence?"; three choice chips, each a job and its test in
  words (noun - names a thing or idea, verb - something you do, adjective - describes a
  noun); then, behind one click, a green bar with the answer and its reason taken from
  the sentence (`"why": "gravity is the thing that pulls the ball"`, required). The
  builder draws all of this from `morph`, `pos_answer`, `sentence` and `why`. The
  sentence must use the CARD's keyword (-ive: active, never constructive): the slide
  bolds the keyword and asks "What job does active do?", so the builder warns when the
  keyword is missing (Term 4 Week 5).
- CHOOSE THE SENTENCE SO THE JOB IS PROVABLE FROM IT: the keyword is the thing doing or
  receiving the action (noun), the action itself (verb), or the word describing a noun
  that is in the sentence (adjective). The `why` names that evidence.
- CARD NOTES (the builder writes `ANSWER: noun - <why>` first): `1. READ` the card
  chorally, `2. ASK` "What job does gravity do in this sentence: noun, verb or
  adjective?" with `10 sec, turn and talk. Then: "Everyone, together, on three."` and
  EXPECT, a `SCAN:` line for a split or wrong chorus (do not click; ask "What does the
  pulling?", then re-ask on three), `3. REVEAL` after the choral answer with a SAY that
  gives the sentence reason
  AND the suffix rule ("Gravity is the thing that pulls the ball, so it is a noun. -ity
  words name a state or quality, so they are nouns."), then a `TRAP:` naming the likely
  wrong job and the fix ("verb, because pulls is an action. Fix: pulls is the verb; ask
  what does the pulling"). Say "usually" for a suffix whose words are not always one job
  (-ive: detective, objective; -ciate: an associate); -ity and -tude always make nouns.
- `pos_drill.intro` is the I do, every Thursday: title "Noun, verb or adjective?", rule
  "Read the sentence. What job does the bold word do?", a `prompt` sentence with one
  bold word from the week (Week 4: "The **rejection** made her try even harder."), the
  same three `choices`, and `answer` + `why` revealed on the click. Its notes define
  the three jobs, think aloud from the sentence, and show the suffix changing the job
  (reject is something you do; add -ion and it becomes a thing). The builder rejects an
  intro that prints a worked answer on the face (`example` or `items`).

## 3h. TOP-TO-BOTTOM HIERARCHY ON SLIDES WE BUILD (v1.8)

James (23 Oct 2026): "Hierarchy is a very important thing in teaching. Teachers work
from top to bottom." The template slides already do this; the slides this pipeline
draws itself (keyword and sentence, word matrix, break the word, which word is it,
parts-of-speech I do and cards, grammar I do / We do / You do, the vote) must too:
- THE TITLE SAYS THE TASK in words a student could act on: "Noun, verb or adjective?",
  "Break the word into its parts", "Which ject word is it?", "Pick your own ject word".
  Never a label for the method ("Find the job with a frame", "Name it").
- READ ORDER IS WORK ORDER: the banner says what to do, the thing students work on (the
  word, sentence or clue) is the hero straight under it, then the choices or boxes, then
  the answer. Nothing students need sits below something they do later.
- THE THING THEY WORK ON IS ON THE SLIDE, big and bold. A test word that lives only in
  the notes is a failed slide.
- NO ANSWER ON THE FACE BEFORE STUDENTS TRY: answers go in the green bar behind one
  click (`answer` + `why`), except a round item marked `worked` ("One done for you").
- TWO STEPS ARE NUMBERED in the hero ("1. Write your own ject word for throw. Star it."
  / "2. Write one sentence that shows what it means."), never one in the banner and one
  in small print.
- THE NOTES SAY WHY in teacher words. "To supportive - no. Very supportive - yes." is
  not an explanation; "it is the thing doing the pulling, so it is a noun" is.

v1.7 rules follow and still apply to v1.7 specs:

Kate Watt (leadership meeting, 13 Oct 2026): "One thing that's lacking across the
school is ... the suffixes, your part of speech. Is it a verb, is it a noun, is that an
adjective? ... go back, review ones you taught, put the part of speech on." Yoshimoto's
card drill asks for meaning AND part of speech on every suffix card.

- `pos_drill.cards`: 6-8 suffix cards the class has been taught, mixed, older ones
  first in the term so the whole set is swept. Each card shows the suffix; the click
  reveals the part of speech it makes, taken from the PRINTED card
  (`yoshimoto_cards_suffixes.json` prints a part of speech on every suffix card). Where
  a card prints two (`-ant`: adjective, noun), give `pos_answer` = the one this class
  should say first and teach the second in the card's `notes`. The builder fails a
  `pos_answer` the card does not print.
- `pos_drill.intro` (first Thursday of the term, optional later): one magenta I do slide
  teaching the four jobs in student words - noun names a thing, person, place or idea;
  verb is the action; adjective describes a noun; adverb tells how, when or where -
  with one family as the model (`act` verb, `action` noun, `active` adjective,
  `actively` adverb).
- Card notes: `ANSWER:` line, then the card's meaning, then one beat asking a student
  for a word with that suffix used in a short sentence.
- Thursday's grammar block follows (section 7). Use a word from this week's family in
  its You do or EXTENSION so the morpheme lands in the class's own writing.

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
   `og_planner/morpheme_bank.json` for an older teacher-confirmed entry. Bank cards
   already taught (micro, pict/picto, vore, eco, carn/carni, chron, -acy) carry
   `associated_words` (the default family - same role as the photographed field),
   `taught` (the grid, spelling words and memory hook actually used), `avoid_words`
   and `teaching_notes` (the rules and traps taught). Plan from those fields instead of
   re-reading the source scan or old week specs. When a deck teaches an uncaptured card,
   append the same fields to its entry afterwards (the file is append-only).
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

- For a captured card, `associated_words` is the DEFAULT authorised family for new
  Words to Read, new Words to Spell, review pools, extra tasks, and morphology
  examples. Choose for age, decodability, transparency, and lesson purpose from within
  that family; do not silently replace it with a web-generated list.
- It is a default, NOT a ceiling on quality. The catalogue lists what is on the physical
  card; it was not curated for Grade 5/6 teaching value or for stacking with this
  cohort's taught history. Where a word outside the list carries the same morpheme doing
  the same job and teaches it better - simpler, more transparent, or stacking a
  previously taught prefix/root/suffix that no remaining associated word offers - use
  it, under ground 3 below. Selection quality (3b) outranks list membership; the CARD's
  own metadata is what stays locked.
- `excluded_words` preserves the physical-card provenance but is a hard block on
  automatic or incidental student-facing selection. Do not restore an excluded word
  merely because it is common, appears online, or fits a word count. Use one only when
  the user explicitly requests that exact word and the context is school-appropriate;
  record the reason in the handoff.
- A new morphology spelling word must be a subset of that session's new morphology
  reading grid. Same-day review reading and review spelling lists must remain disjoint.
- An outside derivative is allowed on three grounds, and only these three:
  1. the captured card has too few suitable associated words;
  2. the user explicitly requires that exact word; or
  3. IT TEACHES THE MORPHEME BETTER than the catalogue options left to you - it is
     simpler, its morpheme sum is more transparent, or it stacks a previously taught
     prefix/root/suffix that no remaining associated word offers.
  In every case: verify the morpheme is visibly present and contributes the taught
  meaning, verify the whole-word meaning and Australian spelling from an appropriate
  source, and record it in the session's `associated_word_exceptions` with `word`,
  `reason`, and `source`; repeat it in the handoff. A ground-3 entry names the catalogue
  word it displaced and why (`rejoin instead of rejoinder: re- is taught, -der is not`).
  Never use an outside derivative merely to pad a genuinely niche family up to 12 words -
  build a 9-word grid instead (3b).
- WIDENING THE WORD CHOICE NEVER LICENSES EDITING THE CARD. The card's own `morpheme`,
  `meaning`, `keyword` and printed `part_of_speech` stay locked exactly as captured
  (0a, 4a). Ground 3 is about which derivatives you teach, nothing else.
- `excluded_words` remains a hard block. A simpler-and-clearer word that sits on the
  exclusion list is still excluded.
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
7. The sound bank does not contain the day's focus morpheme, and every one of its 9
   boxes can be tied to a word on that day's Words to Spell Review list (2c).
8. Every student-supplied answer in the session - You Do items, `EXPECT:` answers,
   check-slide questions, `After checking:` prompts, extensions - passes the
   taught-morpheme gate (2f). Build the session's taught-morpheme set first (today's
   focus morpheme + the review-10 + the sound bank + the term timeline), then test each
   answer against it. Anything that fails is rewritten or has the missing meaning
   printed on the slide.

Any mismatch is a blocker, not an advisory note. Fix the source/spec before generating
slides.

---

## 4e. THE 3/4 TEAM'S CARD WORDING (v1.9 - advisory only)

`og_planner/school_cards.json` holds the card wording the 3/4 Challenge team was given
(20 Oct 2026, 113 cards). It differs from our photographed cards in small ways: `port`
is `transport / to carry` there and `export / bring, carry` here; `anti-` is
`antifreeze` there and `antidote` here. Our 5/6 class learned the photographed wording,
so the slides keep it (0a rule 4). The builder raises a NOTE, once per card per build,
wherever a card on the deck differs, so the teacher can raise it with Kate before the
two cohorts' card language has to line up (the 2027 plans). Never switch a card's
wording on the strength of this file alone; switch only when Kate or the teacher says
so, and then record it in the catalogue, not here.

---

# 5. LEARNED WORDS (red words)

- V1.8 ROUTINES (James, 22 Oct 2026). The builder writes them into the section header
  notes; each word's notes follow the same steps in its own words:
  - REVIEW word: 1. SAY the word in a sentence; students write it once from hearing it.
    2. SHOW the word; students copy it correctly underneath. 3. COVER both; students
    write it from memory, then tick or fix. Then the `Say it:` line.
  - NEW word: the `Why learned:` and `Say it:` lines, then 1. COPY it twice, saying the
    letter names. 2. COVER both and write it a third time from memory, then tick or fix.
    Then the `Link:` line.

- v1.7 SCHEDULE: learned words run on Tuesday (2 review + NEW word 1), Thursday
  (2 review + NEW word 2) and Friday (review only, both of this week's new words plus
  one older one). Never on Monday or Wednesday. Tuesday's new word is Yoshimoto's
  "new and review learned word" on day 2, and it goes straight into Tuesday's
  dictation (Kate: "New red word ideally should be in there"). Pick topic words the
  class will write this term (Kate called inquiry-linked learned words "bang on
  perfect").
- LEGACY: every session REVIEWED the two most recently taught learned words and new
  words were dealt on a FIXED SCHEDULE, not one per day until they run out.
- LEGACY SCHEDULE (the second new word is always taught in SESSION 3, never back to back
  with the first):

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
- SETTLED VALUES LIVE IN `og_planner/teaching_record.json`. Every learned word already
  taught has its week, its `unfair` part and its Australian `say_it` recorded there.
  Reuse them verbatim on review slides - do not re-decide them. Two drifted before they
  were recorded: tongue (settled on `ongue`) and asthma (AS-muh - the t is silent, not
  AST-muh). When a week is delivered, append its new learned words there.
- REVIEW-ONLY WEEKS CAN REVIEW MORE THAN TWO. The two-review rule is for teaching weeks.
  An end-of-term review week (11a) deals every word of the term across its sessions -
  four or five review words a session - and grouping them by the kind of unfairness
  (silent letters, vowel letters that lie, borrowed spellings, letters doing odd jobs)
  gives each Link line a real job.

---

# 6. DICTATION (two sentences, both with reveal)

- V1.8 SHORT WEEK: in a 4-session week Thursday is the last day and ends with ONE
  dictation sentence (green meter) written in the week's grammar focus and carrying a
  week's word plus a throwback from last week (Term 4 Week 4: `Our teacher rejected the
  report; we had to write it again.`). It runs LAST, after the grammar block and the
  vote, so students write the grammar they have just been taught. The builder places it
  there, fails a last-session apply day with any other count, and an apply day that is
  not the last session carries no dictation.

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
- v1.7 CONTENT SOURCING - DICTATION CARRIES WHAT WAS JUST TAUGHT (Kate Watt, Oct 2026,
  and Yoshimoto's upper-level planning slide: "Dictation - must prioritise what was
  introduced in this lesson and can also include anything previously taught"). Dictation
  runs on Tuesday and Friday. Every dictation day:
  1. At least one target is THIS WEEK'S TEACHING: a word built on this week's morpheme
     and, on Tuesday, the new learned word. Yoshimoto's worked example dictates the
     day's new root (subtraction, distracted) and the day's new learned word (gone).
  2. At least one target is a HARDER WORD FROM A WEEK OR TWO BACK (Kate: "throw in one
     from a week or two ago ... the trickier ones, like entrepreneur"). This answers the
     August copying concern: the throwback shows whether the learning stuck.
  3. COPYING IS FINE. Kate: "if your front rowers are going to copy it ... who cares?
     Because they need to copy it and they get it right." Do not tell students to close
     their books.
  The builder fails a v1.7 dictation day with no target from this week's teaching and
  raises a NOTE when every target is this week's. The rest of this section (length,
  sense, punctuation, COPS) applies unchanged.
- LEGACY (v1.3-v1.6) CONTENT SOURCING - DICTATION WAS REVISION FROM 2-3 WEEKS AGO. School
  feedback (Aug 2026): when dictation targets were that week's new words, students
  copied them from the front of their books instead of spelling them, so the dictation
  measured nothing. Targets therefore came from material OUT of the current pages of the
  book:
  1. Derivatives of morphemes taught 2-3 weeks before this week (the older half of the
     review-10 and the term timeline in `og_planner/taught_morphemes.json`; the current
     week's Words to Spell Review pool is a ready-made source).
  2. Learned words taught 2-3 weeks ago (from the user's `Recently taught learned words`
     list, which is why that list must reach back at least three weeks).
  3. Older still, weighted by the 2b recipe.
  NEVER: today's or this week's new morphology words (anything on this week's Words to
  Read / Words to Spell grids), any word built on this week's focus morphemes, or this
  week's new learned words. Last week's words and the two learned words still on the
  review pair are the fallback of last resort, not the plan.
  Every content word must be spellable from taught material; scope-check phonograms.
  THE BUILDER ENFORCES THE HARD PART: a target that is one of this week's grid words or
  new learned words, or that contains a 4+ letter form of one of this week's focus
  morphemes, is a fatal WARN. A short-form match (`-ice`, `-ery`) and any target that
  was new last week (read from `term<T>_week<N-1>.json` in the same folder) or is on the
  learned-word review pair is a NOTE - answer it, usually by reaching further back.
  The `focus` line names where the targets come from (`Term 2 review: capt, ord`), so
  the teacher can tell students this is revision.
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
- COPS MARKING (required for both sentences). The school marks dictation with COPS -
  Capitals, Organisation (every word there, in the dictated order), Punctuation,
  Spelling. COPS replaced CUPS in Oct 2026: Kate Watt asked for one acronym across the
  whole school. Give every dictation a `cops` block (`cups` is still read so delivered
  specs rebuild): `capitals` (the words that must start with a capital) and
  `punctuation` (each assessable mark, named). The builder then:
  - colours every required CAPITAL LETTER GREEN in the revealed sentence (C),
  - colours all punctuation RED in the revealed sentence (P),
  - keeps spelling targets bold + underlined (S),
  - prints `Score: ___ /N` on the slide, where N = capitals + punctuation marks +
    targets + 1 for Organisation, so students self-score against a visible total,
  - writes a tick-off COPS checklist in the notes with each category's items and count.
  Sanity-check N against the sentence before shipping - every red mark, underline and
  capital on the slide must be counted once.
- COPS IS THE BRIDGE TO THEIR OWN WRITING. Once or twice a week (not every session),
  one scripted line at the check step tells students the transfer: `COPS is not just
  for dictation - run the same four checks on your own draft before anyone else reads
  it.` Reward the brave word: leaders reported students writing "big" instead of
  "enormous" for fear of misspelling it, so the line can add `a big word spelt nearly
  right beats a small word spelt right`. The section 7 HOW standard: dictation is where the class rehearses the editing
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
- NO BRACKETS IN A DICTATION SENTENCE. The builder colours only `, . ; : ! ? " ' -` red,
  so round brackets render in black and never reach the red P count on the slide. Keep
  acronym-and-brackets practice on the grammar slides; weave that grammar week into
  dictation through another mark (a fronted clause, an apostrophe).
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
- SPREAD THE WEEK'S TARGETS: across a week's dictation sentences (four in v1.7, ten in
  legacy weeks), use an older morpheme or red word as a target once, and do not reuse
  last week's throwback targets. This week's morpheme and new learned words are the
  exception: they are meant to recur on Tuesday and Friday. A Term 4 Week 2 draft
  targeted caus in four of ten sentences and repeated salmon, vacuum, captain and
  postpone from Week 1; the reviewed set uses ten different Term 3 sources.
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

- v1.7: GRAMMAR RUNS ON THURSDAY ONLY (Friday instead, in a week with no Thursday). The
  week's one focus gets one well-made block, not five days of it. Its You do (or its
  EXTENSION) uses a word from this week's morpheme family in a sentence in the grammar
  focus - the Grade 4 model lesson in Yoshimoto's training deck ends exactly this way
  (`loom` in a sentence that opens with a subordinating conjunction). The ABSENT-STUDENT
  rule matters more than ever: the I do must stand alone.

- FIVE MINUTES IS A HARD CAP, AND THE SHAPE ENFORCES IT (v1.6). Grammar is five slides
  after its header: one I do slide, ONE We do item, THREE You do items, each item on its
  own slide with the answer on the click - the maths decks' one-question-per-slide
  practice round (MEGA_PROMPT 17, 20, 82). Budget: I do 1 minute, We do 1 minute, You do
  about 45 seconds an item (30 seconds writing, then scan, click, tick or fix). The old
  four-item You do slide plus check slide ran long and asked students to hold four
  sentences at once; do not author it for new weeks.
- WHY ONE ITEM A SLIDE: every student commits to one answer, the teacher scans the whole
  room before the click, and a student who got it wrong fixes it while the reason is on
  screen. The You do items share the I do's exact move, so a student who followed the I
  do and the We do can do all three alone (example set, MEGA_PROMPT 81): easy, typical,
  then an edge case (a past-tense verb, a plural subject, a near-miss chip).
- The last You do item's notes carry the `ANSWERS:` block for all three items plus an
  example answer for every EXTENSION in the round - the teacher holds an open task they
  cannot mark otherwise (3e).
- In an A/B We do, change ONLY the verb (`Possums wrecked our bins` / `Possums knocked
  over our bins`). If the nouns differ too (`wrecked our whole school`), students pick on
  the noun and never judge the verb. Check every chip for a defensible second answer in
  everyday Australian speech: `took the trophy` is ordinary sports talk for won, so it
  cannot be the wrong chip.
- FIXED ANSWERS, WRITTEN SMALL: each You do item gives three chips - the answer, the
  vague or wrong-strength near-miss from the TRAP, and one clearly wrong option - and
  students write only the answer word in the OG book. Writing the whole sentence is the
  item's EXTENSION, not the core, because it doubles the time.
- The We do is a round of one: an A/B judgement or three chips, answered with fingers,
  voices off, then one cold-call "How do you know?" after the reveal.
- The You do needs no separate "One done for you" slide: the We do item, worked together
  seconds earlier, is the model. EXTENSION and HELP go on the first You do item's notes,
  and the HOW transfer line is beat 4 of the last item.

- ONE grammar focus per week: one Thursday block in v1.7 (legacy weeks taught it as a
  4-5 day sequence, Mon intro -> Fri week review). Link back to previous weeks' grammar
  when natural.
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
- TEACH-IT-COLD NOTES (non-negotiable, school feedback). The notes must let a teacher
  who did not plan this lesson, and who may not be confident with the grammar, deliver
  it well straight off the iPad. The slide face never does that on its own: it shows the
  FINISHED example, not how it got there, so a note that only states the answer leaves
  the teacher inventing the demonstration in front of the class. Every I do and We do
  carries all three of these:
  - A BOARD-EXPLICIT DEMONSTRATION BEAT. The `MODEL:` / `WRITE:` / `POINT:` beat names
    the exact thing the teacher does to the visible material, in order: which word to
    point at, what to write, where it goes, and what the board reads when the beat ends.
    `3. MODEL: write both commas.` is not deliverable - the example is already printed
    complete on the slide, so the teacher cannot tell whether to annotate it, rewrite it
    or build it from scratch. Author it as
    `2. WRITE: copy "Our canteen sells pies." on the class board. POINT at canteen.` then
    `3. WRITE: add ", the busiest room," straight after canteen so the class board now
    reads the full sentence.` If the demonstration is annotation rather than writing, say
    so and say what gets marked (`circle canteen on the slide, then draw both commas in`).
    "Class board" means the teacher's own display board - it is not the student
    whiteboard ban below, which is about where STUDENTS write.
  - A `CLARIFY:` LINE, one per I do and We do, IN THE LIVE ZONE. It carries the one
    question this slide predictably provokes with its answer scripted in plain classroom
    English, or the boundary the teacher needs in order to judge a student answer nobody
    planned for:
    `CLARIFY: if a student asks whether the extra words can go at the end, say the
    reader has to hold the name too long, so they sit straight after it.`
    `CLARIFY: any wording counts if you can cover it and the sentence still stands.`
    This is the on-the-fly support - it goes ABOVE the divider, never in the prep zone,
    because it is needed mid-sentence with thirty students watching.
  - A `PREP:` LINE in the prep zone giving the teacher the concept in one plain sentence
    FOR THEMSELVES, not for the students: `PREP: an appositive is just a renaming phrase
    - if you can cover it and the sentence still works, it is one.` A teacher who reads
    nothing but the rule banner and that line should still be able to run the slide.
  LIVE-ZONE BUDGET - the eight-line cap is enforced by the builder, so plan it before
  writing: an I do or We do runs ANSWER + 3-4 beats + TRAP + CLARIFY = 6-7 lines. The
  You do already spends its eight on ANSWER + 4 beats + TRAP + EXTENSION + HELP, so it
  does NOT carry a CLARIFY line - its `CIRCULATE:` prompt ladder does that job. If a
  You do genuinely needs one, drop to three beats.
- THE SAME STANDARD APPLIES WHEREVER THE TEACHER MUST DEMONSTRATE: the morpheme card and
  copy slide (3a) and the Words to Read grid script (3b) are the morphology equivalent of
  an I do. A beat that says what the teacher KNOWS but not what the teacher DOES is
  half-written.
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
  - On a round, the We do item worked together is the completed example. On the older
    single slide, the hero card shows a COMPLETED WORKED EXAMPLE of the task (`One done for you:
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
- In a round, each item's green answer bar replaces the check slide (3d). On the older
  single-slide shape, a fixed-answer You do MUST be followed by the same green `Tick it or fix it - ...`
  answer slide as the morphology You Do: supply `check_title`, `check_rule`,
  `check_items`, `check_routine`, `check_footer`, `check_notes` on the `you_do` block.
  The answer slide must teach (3d): each answer carries its one-line reasoning, and
  `check_notes` MODEL the reasoning for the first item before students tick.
- Teacher notes use the Glance Format from `MEGA_PROMPT.md` sections 45-47: LIVE ZONE
  (ANSWER first if the slide asks anything; 2-5 numbered beats fusing action + talk with
  CAPS anchors - SAY/ASK/MODEL/SCAN/TIME/CIRCULATE; TRAP with Fix; EXTENSION/HELP), then
  `---`, then a one-line prep zone with the `[Stage | element | focus]` tag, then the
  `ANSWERS:` block on fixed-answer You dos. SAY lines are warm, natural classroom talk:
  one short speaking turn, usually 12-24 words, with a brief connector where useful.
  Never compress them into clipped fragments merely to save words.
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
  numbered beats; TRAP with Fix; CLARIFY (required on grammar I do and We do, section 7);
  EXTENSION/HELP. No blank lines inside it. `STRETCH` is retired everywhere in OG decks -
  the label is `EXTENSION` (3e). The Words to Read Review script is the sole exception to
  the eight-line cap and follows section 2b's longer numbered reading-plus-retrieval
  structure.
- V1.8 PREP ZONE (James, 23 Oct 2026: "Purpose and prep - what's all that down the
  bottom? That doesn't make any sense"): no `Purpose:` line and no `[Stage | element |
  focus]` tag in a v1.8 week; the builder fails the build on either. Below `---` sits
  only what a teacher needs before the lesson, in plain words: an `ANSWERS:` block on a
  fixed-answer You Do, and at most one `WHY:` line explaining the idea behind the slide
  ("WHY: the suffix sets the job - reject is a verb, rejection is a noun"). If a slide
  needs neither, it has no `---` at all. The rule below is the v1.7 and legacy one.
- PREP ZONE: `---` on its own line, followed by at most three short lines - the
  `Purpose:` line with its `[Stage | element | focus]` tag, an optional `PREP:` line
  giving the teacher the concept in their own terms (section 7), and at most one more.
  Anything needed while students are working must stay above the divider. The `ANSWERS:`
  block on a fixed-answer You Do is exempt from the three-line limit (3d).
- SAY text is natural, direct teacher language that can be read verbatim. Use
  contractions and brief spoken connectors where a real teacher would: "Okay, let's
  check that word." / "Now, look at the ending." / "So, what changed?" These words are
  useful when they join ideas or soften an abrupt start; they are not a licence for
  repeated greetings or management chatter. Do not write clipped production notes.
  Each question includes think time, one response routine, and `EXPECT:` or `Answer:`
  in student-friendly language.
- Run the WHOLE-SCRIPT READ-ALOUD TEST. Read the scripted SAY, ASK, `After checking:`
  and `Answer:` lines in teaching order. If the sequence sounds like separate labels or
  commands, or the teacher must invent "Okay, so..." between them, include the smallest
  natural connection in the script. Do not add another explanation just to make it
  longer.
- Natural does not mean padded. Prefer `Okay, let's check the ending. What does -ule
  tell us here?` over `Check ending. Meaning?`, but reject `All right, boys and girls,
  everyone looking this way, now let's move on to our next exciting word.`
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
| Sound bank | `green = root` / `yellow = prefix` / `red = suffix`, one per line (builder default). The 9 boxes never include the day's focus morpheme (2c) |
| Spell the word (each) | `Word to spell:` / `Sentence:` / reveal line / `After checking:` / `Answer:` (all required); builder renders labels bold and underlines the target word where it appears inside the `Sentence:` line |
| New morph card | `Type:` / `Keyword:` / `Meaning:` or separate `Sound:` lines; the copy slide adds `MEMORY HOOK:` automatically |
| Words to Read grid | one scripted line per word + `Memory hook:` last |
| Words to Spell grid | line-separated early-finisher cue + `EXTENSION:` (3e) |
| Learned word (each) | `Why learned:` / `Say it:` AU pronunciation / `Link:` |
| Dictation (each) | sentence + line-separated COPS checklist + `Score:` + `Focus:` |
| Round item (You Do, grammar We do / You do) | `ANSWER:` / `1. ASK` or `SAY` with think time and the cue / `2. SCAN ... 80%+ -> click to reveal. Less -> [ladder], then re-ask.` / `3. REVEAL after ... SAY: "Tick it or fix it. [reason]"` / `TRAP:` / `---` / `Purpose:`. Worked item: the round's beats, EXTENSION, HELP and the `ANSWERS:` block |
| Grammar I/We/You do | Glance Format live zone with a board-explicit `MODEL:`/`WRITE:` beat and a `CLARIFY:` line on I do and We do, `---`, `Purpose:` line + optional `PREP:` teacher gloss, then `ANSWERS:` on a fixed-answer You do (section 7) |
| Sound bank (v1.7, dictated) | builder writes the say-order list (`1. port - carry`) and the colour key including `white = sound` |
| Recap card (Tue) | builder writes `Type:`/`Keyword:`/`Meaning:`; spec `notes` adds the CHORAL and ASK beats with EXPECT |
| Keyword and sentence (Mon) | Glance You do with no ANSWER line (open task): `1. SAY` why the card keyword works, `2. SAY` pick yours and star it, `3. TIME` the sentence, `EXTENSION:` and `HELP:` |
| Word matrix (Wed) | Glance You do: beats, `TRAP:` (a made-up word), `EXTENSION:`, `HELP:`, `---`, `Purpose:`, `ANSWERS:` listing the real words with meanings |
| Parts-of-speech card (Thu) | builder writes `ANSWER:` (the part of speech), the card meaning and printed parts of speech; spec `notes` adds one ASK beat with EXPECT. v1.8: builder writes `ANSWER: <job> - <why>`; spec notes are READ / ASK with EXPECT / REVEAL with the sentence reason and the suffix rule / TRAP (3g) |
| Friday vote | `1. SAY` the four options, `2. ASK` fingers 1-4 voices off, `3.` what to write in next week's `From our reviews` line |

---

# 9. WEEK SPEC JSON (the artefact you author)

One JSON per week in `og_planner/weeks/`. v1.8 exemplar: `og_planner/weeks/term4_week4.json`
(ject, four sessions: `teach`, `spell`, `build`, `apply`, with Friday "No session this
week"). v1.8 adds, on top of the v1.7 fields:
`"og_version": "1.8"` at the top; `morphology_review` and `words_to_read_review` on
every session; `memory_hook: {"text", "picture"}` on the teach day (3a); `sentence` on
every `pos_drill.cards` item (3g); `parts` on every Break it apart item and `given` on
the round (3f); `vote` on whichever session is last (1, 11).
v1.7 exemplar: `og_planner/weeks/term4_week3.json`
(port: `teach`, `spell`, `build`, `apply`, `week_review` - the gold standard for every
new week). The legacy schema and exemplar (`og_planner/weeks/sample_term3_week1.json`,
Monday `new` + Tuesday `review` + Friday `week_review`) follow; v1.7 day types use the
same field names for the same blocks, plus the v1.7 fields after it. NAME IT
`term<T>_week<N>.json` exactly: the builder finds last week's spec by that name for its
section 6 dictation NOTE, and `og_history.py` / `audit_week_spec.py` read the same
pattern. A differently named spec silently skips those checks. Schema:

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
    "sound_bank": [ { "morph", "type" } x9 ],  // derived from words_to_spell_review;
                                               // never the day's focus morpheme (2c)
    "words_to_spell_review": [ { "word", "sentence", "prompt", "answer" } x10 ],
    "new_morphology": { "morph", "type", "keyword", "meaning" },   // omit on week_review
    "words_to_read_new": [ { "word", "meaning" } x8-12 ],          // omit on week_review
    "associated_word_exceptions"?: [ { "word", "reason", "source" } ], // rare; section 4c
    "wtr_new_notes": "verbatim line-separated teacher script + Memory hook: (3b)",
    "words_to_spell_new": [4 words, all from words_to_read_new],   // omit on week_review
    "extension": "self-contained early-finisher + EXTENSION: (3e)", // omit on week_review
    "new_morph_activity": { "title", "rule", "routine", "footer", "time"?,
                            "round": [ <round item> x5 ] },   // 3d, v1.6: item 1 is
                                                              // the worked one
       // <round item> = { "prompt" (hero, <=3 lines of <=47 chars, `**bold**` ok),
       //                  "choices"?: [<=3, "word\npart (meaning)", <=24 chars a line],
       //                  "answer", "why" (bar = answer + why, <=58 chars with chips),
       //                  "notes" (ANSWER first), "worked"?: true (answer shown, no click),
       //                  "rule"?/"routine"?/"footer"? per-item overrides }
       // Legacy single slide (old specs only): "example", "items", "notes",
       //   "check_title", "check_rule", "check_items", "check_routine", "check_footer", "check_notes"
    "file_name"?: "override only if the user asks",
    "learned_words": { "review": [ { "word", "unfair", "notes" } x2 ],
                        "new": { "word", "unfair", "notes" } },     // "new" omitted Friday
    "dictation": [ { "meter": "green" | "yellow", "sentence", "targets": [..],
                     "cops": { "capitals": [..], "punctuation": [..] },   // legacy specs: "cups"
                     "focus" } x2 ],
    "grammar": { "header_notes",
                 "i_do":  { "title", "rule", "example", "items": [<=2], "routine", "footer", "notes" },
                 "we_do": { "title", "rule", "round": [ <round item> x1 ] },  // v1.6
                 "you_do": { "title", "rule", "routine", "footer",
                             "round": [ <round item> x3 ] } }                 // section 7
  } ]
}
```

v1.7 DAY TYPES - which fields each day carries (section 1). A day carries ONLY these;
the builder warns when one is missing and emits nothing else:

```
teach (Mon):  morphology_review x10, words_to_read_review, new_morphology,
              words_to_read_new x6|9|12, wtr_new_notes, brainstorm, keyword_sentence,
              associated_word_exceptions?
spell (Tue):  recap, sound_bank x6-9, words_to_spell_review x6-8,
              words_to_spell_new x4 (all on Monday's grid), extension,
              learned_words {review x2, new}, dictation x2
build (Wed):  build
apply (Thu):  pos_drill, learned_words {review x2, new}, grammar
week_review:  morphology_review x10, words_to_read_review, sound_bank x6-9,
  (Fri)       words_to_spell_review x6-10, learned_words {review}, dictation x2,
              grammar? (only when the week has no apply day), vote

"brainstorm": "\"What words do you already know with port in them?\" 20 sec, turn and tell. EXPECT: transport, airport.",
"keyword_sentence": { "title", "rule", "example", "items"?, "routine", "footer", "notes" },
                      // structured slide (10c budget); the hero shows a model keyword + sentence
"recap": { "morph", "type", "notes" },          // Monday's card again; notes = beats + EXPECT
"sound_bank": [ { "morph", "type": root|prefix|suffix|sound, "say"? } x6-9 ],
                      // `say` only on a sound: "/ow/ as in cloud" (2c)
"build": { "time"?: "15 mins",
           "matrix": { "title", "rule", "base": { "morph", "type", "meaning" },
                       "prefixes": ["trans- (across)", ...3-6],
                       "suffixes": ["-able (able to)", ...3-6],
                       "routine", "footer", "notes" (beats, TRAP, ---, ANSWERS:) },
           "rounds": [ { "title", "rule", "routine", "footer",
                         "round": [ <round item> x5 ] } x2 ] },   // break apart, then name it
"pos_drill": { "intro"?: { "title", "rule", "example", "items"?, "routine", "footer", "notes" },
               "cards": [ { "morph", "pos_answer"?, "notes"? } x6-8 ] },  // suffix cards (3g)
"vote": { "title"?, "rule"?, "options": [2-4 short strings], "routine"?, "footer"?, "notes" },
"dictation": [ { ..., "cops": { "capitals": [..], "punctuation": [..] } } x2 ]
```

Notes:
- Card `keyword`/`meaning` fields must mirror the photographed catalogue when the card
  is present there. The builder enforces this and gives the photographed transcription
  priority over the legacy bank and reference meanings. Any differing spec value
  is a blocker even if it appears only as a WARN. For a genuinely ambiguous heading
  such as `di-`, include the intended photographed meaning in the spec.
- New-family words must pass the section 4c/4d associated-word and exclusion audit. A
  word chosen outside `associated_words` on quality grounds (4c ground 3) must be logged
  in `associated_word_exceptions`.
- `taught_morphemes` (optional, WEEK level, alongside `term`/`week`/`sessions`): labels
  taught earlier in the term that no longer appear on the review cards. Strings or
  `{"morph": ...}` objects both work. The builder's taught-morpheme gate already counts
  today's card, the review-10, the sound bank and every earlier session in the same
  spec; this field carries the rest of the term timeline. Since v1.5 the You Do gate
  uses that set only to decide whether a PRINTED part builds silently (taught) or with
  a NOTE (untaught) - an unprinted non-focus part fails either way (3d). Add a label
  here ONLY if it was genuinely taught - the field is the escape hatch, not a mute
  button. The
  authoritative term-by-term record is `og_planner/taught_morphemes.json` (teacher-
  supplied timeline): copy labels from it, never from memory, and never add a label
  that file does not hold.
- The builder enforces PART of sections 2c, 2f and 3d (see section 10). On the You Do
  face it stops meaning sums and unprinted parts; it cannot see prose tasks, `EXPECT:`
  answers or `After checking:` prompts, so the section 4d audits remain author-side: a
  clean build is not proof they passed.
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
python og_planner/og_history.py --from 6 --to 9      # what earlier weeks taught (no spec mining)
python og_planner/og_history.py --index              # word -> every week/list that used it
python og_planner/audit_week_spec.py og_planner/weeks/<spec>.json         # pre-build audit (ISSUE = fix)
python og_planner/build_og_week.py og_planner/weeks/<spec>.json          # all sessions
python og_planner/build_og_week.py og_planner/weeks/<spec>.json --only Tuesday
python tests/test_og_builder_regressions.py                               # builder + exemplar regression gate
# WARN = fatal, NOTE = advisory. Run the regression suite before touching
# validate_sound_bank, validate_activity_morphemes or morph_surface_forms.
```

`audit_week_spec.py` runs the checks the builder does not: dictation character bands
and brackets (6), same-day read/spell overlap (the builder skips it on `week_review`
days), repeats against the previous session, card order and position repeats (2a),
scripted `EXPECT:` answers missing from the review board (2b), sound bank labels that
are not on the taught timeline (2c), text that will wrap on structured slides (10c)
and learned-word values that differ from `teaching_record.json` (5). Fix every ISSUE.
A CHECK is a render-and-look item.

1. The builder's gate must pass (exit 0): file reopens cleanly, no `XYZ` left anywhere,
   notes are separate left-aligned PowerPoint paragraphs with real bold labels, each
   spelling target is underlined inside its `Sentence:` line, the standard live zone
   is no more than eight lines, Words to Read Review includes all three required READ
   beats and guarantees every student at least two rows, review questions are not piled
   into one paragraph, every review/new card and Sound Bank box uses green root / yellow
   prefix / red suffix, fixed answers have an immediate green `Tick it or fix it` slide
   (morphology You Do AND any fixed-answer grammar You do),
   new-word grids are at least 27 pt, grammar examples are at least 20 pt, the first
   dictation uses the green meter while the second/trickier dictation uses yellow, the
   sound bank and review cards do not contain the day's focus morpheme or any morpheme
   this week teaches later (2a/2c), no dictation target is one of this week's new words
   or built on this week's focus morphemes (6), no You Do item or check item is a
   meaning sum, and every part a You Do asks for beyond today's morpheme is printed
   with its meaning on the slide face (2f/3d).
   Read every WARN - word-count, catalogue mismatch, excluded-word, unbanked-label and
   overflow warnings are content bugs to fix in the spec, not noise. Every WARN fails the
   build. A NOTE is advisory and does NOT fail the build - currently the unused-sound-
   bank-box check (2c), the last-week / review-pair / short-form dictation checks (6)
   and cards absent from the photographed catalogues. Advisory does
   not mean ignorable: answer every NOTE and list them in the summary. Captured-card
   metadata is never advisory.
2. Visual QA every session deck: `python scripts/pptx_to_images.py "output/<folder>/<deck>.pptx"`,
   then INSPECT the images: title subtitle on one line; overview table cells complete;
   10 card slides in the jumbled order you specified; 15-word table filled; sound bank
   colours match types AND contains no focus morpheme, with every box traceable to one
   of that day's review spelling words (2c); each spell word centred; every new-word grid term is unbroken
   and at least 27 pt; task and check slides form a protected pair; learned-word
   highlight is on the right letters; dictation capitals are green, punctuation red and
   targets bold/underlined; grammar slides are not overflowing. Also inspect the notes
   pane/XML: one thought per paragraph, no raw markdown, no dense paragraph walls and no
   control-character artefacts. Local render substitutes fonts (Luckiest Guy/Lexend may
   look plain) - that is a render artefact, not a bug.
   Then run the AUTHOR-SIDE half of the 2c/2f audits. The builder covers the mechanical
   part (focus morpheme in the bank, meaning sums and unprinted parts on the You Do
   face); it cannot read prose tasks, `EXPECT:` answers or `After checking:` prompts:
   - SOUND BANK: name the review spelling word each of the 9 boxes serves.
   - TAUGHT-MORPHEME GATE: list every student-supplied answer in the deck - including
     the ones written as prose - and confirm each is reachable from today's morpheme +
     the taught history + everyday English.
   - DICTATION REACH: for each target, name the week (2-3 weeks back or older) it comes
     from. The builder can only see this week and last week's specs.
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

1. Read `og_planner/weeks/term4_week4.json` (v1.8) END TO END before writing anything.
   It is the gold exemplar: every notes format, every field, every recipe is
   demonstrated there. For the v1.9 fields - `picture`, the Wednesday third round,
   `friday_review` - read `og_planner/weeks/sample_v19_term4_week4.json` too. Your specs must match its formats exactly - same plain source
   labels (the builder applies real bold),
   same beat numbering, same field names. Do not invent alternative formats.
1b. Get the history from the records, not by opening old specs one by one:
   `og_planner/taught_morphemes.json` (morphemes by term), `og_planner/teaching_record.json`
   (learned words with settled unfair parts and pronunciations; grammar focuses and the
   class's own wording), `og_planner/morpheme_bank.json` (taught families for uncaptured
   cards) and `python og_planner/og_history.py` (every word each session used, plus
   `--index` to check a word's past uses before choosing it).
   The term's planned learned words and grammar focuses live in `og_planner/term4_plan.json`
   (one file per term); use it when a request leaves them out.
2. Load the three category JSONs from section 4, build the week's morpheme timeline
   (which morpheme is taught which day), resolve each exact card, and run the section
   4d audit. Only cards absent from the photographed catalogues may fall back to
   `og_planner/morpheme_bank.json`. Stop and verify anything unconfirmed BEFORE writing
   the spec.
3. Write down the session's TAUGHT-MORPHEME SET before writing any question: today's
   focus morpheme + the review-10 + the sound bank + everything earlier on the term
   timeline. Every student-supplied answer you script must come out of that set (2f).
4. For each session, in slide order, generate ONLY that day type's blocks (section 1):
   EVERY day (v1.8) - review-10 (jumbled per 2a, fresh order) and 15 review words in
   the 5-4-3-2-1 pattern with the v1.8 script (2b). Then:
   Monday - morpheme card with `memory_hook` and picture + brainstorm + grid (6, 9 or
   12, useful words, untaught-part words last, every equation vetted) + script (3a-3b)
   + keyword and sentence slide. Tuesday - 6-8 review spelling words with sentences and
   at most 3 after-check pairs (2d), THEN the dictated sound bank derived mechanically
   from them (2c), 4 new spelling words off Monday's grid + extension (3c, 3e), learned
   words 2 review + new 1 (5), 2 dictations with cops carrying this week's teaching
   plus one throwback (6). Wednesday - matrix, break-apart round with part boxes,
   third round, rotated from last week (3f, v1.9). Thursday - "Noun, verb or adjective?" I do + 4-5 suffix
   cards with sentences (3g), learned words 2 review + new 2, grammar (7). Friday
   (v1.9) - the `friday_review` worksheet and its follow-up rows, written LAST, from
   what the week actually taught (1a). Every Tuesday spelling word gets a `picture`
   (2d). Short weeks per section 11.
   Note the order: the spelling list comes FIRST and the sound bank is derived from it.
   Choosing bank morphemes before you know the words is how the wrong nine get banked.
   Check the bank and the review-10 against the WHOLE week timeline, not just today:
   nothing taught later this week may appear on an earlier day (2a/2c).
   AUTHORING TIP: write the spec through a short Python generator in your scratchpad
   (Python dicts, notes as lists of lines joined with `"\n"`) that writes the JSON.
   Notes are full of quotes and line breaks, and hand-typed JSON escaping is where
   avoidable build failures come from. The JSON in `og_planner/weeks/` stays the artefact.
5. Self-check the spec against section 10b's failure list, line by line, then run
   `python og_planner/audit_week_spec.py og_planner/weeks/<spec>.json` and fix every ISSUE.
6. `python og_planner/build_og_week.py og_planner/weeks/<spec>.json` - fix EVERY warn.
7. Render each deck to images and INSPECT them (section 10). Fix, rebuild, re-render.
   v1.9: also render every enabler sheet and every Friday Review page (student pages,
   answers page, follow-up sheet) at print size and read them: letters joined and on
   the lines, nothing past a margin, two student pages, one follow-up page.
   A contact sheet per deck (three slides across) makes 40 slides a deck quick to scan;
   re-render only the slides you changed after a fix.
7b. Run the review panel on every session deck (`IMPORTANT/REVIEW_PANEL.md`): Steve,
   James and the Team Leader review independently, meet on one change list, and review
   again after the changes until all three sign off (four rounds at most). Every change
   goes back through the spec, the audit, the builder and a re-render; the locked
   catalogues, template and section 10b rules beat any panel finding.
8. Only then report, stating exactly which QA levels ran and how many panel rounds it took.
9. RECORD THE WEEK once it is delivered, so the next session does not re-mine it:
   new morphemes -> `taught_morphemes.json`; new learned words (unfair part + Say it)
   and the grammar focus with its class wording -> `teaching_record.json`; the grid,
   spelling words, hook and rules for any uncaptured card -> its `morpheme_bank.json`
   entry (4a). Anything you had to work out that is not written down here -> this file.

# 10b. KNOWN FAILURE MODES (each of these has been rejected by the school - check
your spec against every line before building)

- v1.9: a Words to Spell Review word with no `picture`, a picture of the morpheme
  instead of the word's meaning, or a picture that needs explaining (a bowl for
  omnivore). (2d)
- v1.9: the same Wednesday third round two weeks running, a word bank whose size gives
  the answers away, a spare bank word that fits a sentence, or a Choose the best
  meaning item whose wrong answers share no part with the word. (3f)
- v1.9: a Friday Review over two student pages, a check box naming a category
  (`Meaning`), a question on last week's morpheme, an untaught part with no printed
  meaning, or an underline-only grammar task at 5/6. (1a)
- v1.9: an enabler sheet in joined cursive, pre-cursive or VICCursiveDotted, or one
  that prints the week's morpheme in full where the gap should be; an extension sheet
  that is not the enabler's mirror, or that prints anything dotted to trace. (1b)
- v1.9: switching a card's keyword or meaning to the 3/4 team's wording because the
  builder NOTEd it. (4e)
- Review cards in the same order two days running, or any singable pattern. (2a)
- Any type-coded card background or Sound Bank box that is not green for a root,
  yellow for a prefix, or red for a suffix. Never trust an older master colour. (0, 2c, 3a)
- A review word whose morpheme is not doing its job (mini in minister). (2b)
- Words to Read Review notes as a definition list, or asking students to underline/
  circle/write - this segment is fully verbal, students have nothing in front of
  them. (2b)
- v1.7: Words to Read Review that does not OPEN with the whole class reading every row,
  or row splits crammed onto one line instead of one split per line. (2b)
- Fun row groups that do not cover every student. (2b)
- A grid script line that is circular (`order - ORD means order: the way things are
  arranged`) instead of the submarine model naming every part with `+`. (3b)
- LEGACY: a new/review session without a `new_morph_activity`, or the same activity
  type twice in one week. Any week: a meaning-sum item of any kind - retired Sept 2026.
  (3d)
- A You Do part beyond today's morpheme with no printed meaning on the slide face, or
  tiers printed on the face instead of EXTENSION/HELP in the notes. (3d one recall)
- A question relying on untaught metalanguage - clause, phrase or tense labels the
  term's grammar history has not taught. Parts of speech on suffixes are taught from
  Term 4 Week 3 (3g). (2b taught-only rule)
- Renaming output files away from the team convention (`1a. Monday (morph).pptx`). (1)
- A scripted question without its answer (`EXPECT:` / `Answer:`) anywhere in any
  notes. (2b, 2d, 7)
- A review block labelled as a check when the notes do not say what the teacher scans or
  hears, or when the pace changes without complete response evidence. (2e, 8)
- Skipping a required review exposure because a few confident students answered correctly.
  Responsive pacing changes time and correction inside the OG sequence, not the sequence. (2e)
- Words to Spell Review outside 6-8 on Tuesday or 6-10 on Friday (legacy: fewer than
  10), or spelling words that duplicate the same day's Words to Read Review. (2d)
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
- A dictation without a cops block, with a capital not shown green, with an uncounted
  red mark/capital/target, or with the wrong meter (first = green; second/trickier =
  yellow), or
  with sentence lengths outside 10-12 / 14-16 words. (6)
- v1.7: a dictation day with no target from this week's teaching (the week's morpheme,
  Tuesday's new learned word), or with no harder word from a week or two back. Legacy
  weeks ran the reverse rule. (6)
- A Sound Bank box or review card carrying a morpheme this week teaches on a later
  day (Wednesday's morpheme in Monday's bank). (2a/2c)
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
- All the group allocations crammed onto one `READ:` line instead of one numbered line
  each. (2b)
- A review question that cannot be answered from what has already been taught. (2b)
- A sound bank note that describes the colour key instead of stating it, or a variant
  label allowed to break mid-bracket. (2c)
- A sound bank containing a morpheme not yet taught at that point in the week (2c -
  BUILD-STOPPING). Legacy weeks also banned the day's focus morpheme on every day; a
  v1.7 bank is dictated, so this week's Monday morpheme belongs in Tuesday's and
  Friday's banks.
- A sound bank box that no word on that day's Words to Spell Review list needs, while a
  morpheme one of those words DOES need is left out. (2c - the builder raises a NOTE)
- Choosing the nine sound bank morphemes before writing the ten review spelling words -
  the bank is derived from the list, never the other way round. (2c, 10a)
- ANY student-supplied answer that needs a morpheme the class has not been taught: a You
  Do item, an `EXPECT:`, a check-slide `ASK:`, an `After checking:` prompt, an extension.
  (2f - the builder stops meaning sums and unprinted parts on the You Do face; the
  rest is your audit)
- Silencing that gate by dumping the morpheme into `taught_morphemes` when it was never
  actually taught. The field records the term timeline, not your preferences. (2f, 9)
- Treating a morpheme as taught because it appeared once in a Words to Read grid script
  line, a spelling `Answer:` or a dictation earlier in the same deck. The teacher voices
  a handful of grid lines, not twelve. (2f)
- A You Do or extension that requires students to recall a specific word from the grid
  they just read, rather than build it from taught parts. (2f, 3d)
- Any You Do part beyond today's morpheme - untaught OR taught weeks ago - without its
  meaning printed on the slide face (`in- (onto) + junct + -tion (the act of) = ?`, or
  a parts bank in the rule banner). (2f, 3d)
- Choosing an ornate catalogue derivative over a simpler, more transparent sibling that
  stacks taught morphemes (`rejoinder` over `rejoin`, `conjunctive` over `adjoin`), or
  padding a grid to 12 with curios instead of building a clean 9. (3b, 4c)
- More than two grid words containing a part the class has not been taught. (3b)
- A grammar I do or We do whose demonstration beat states the answer but not what the
  teacher does to the board - `MODEL: write both commas` when the finished example is
  already printed on the slide. (7 teach-it-cold)
- A grammar I do or We do with no `CLARIFY:` line in the live zone, or one that hides the
  clarification below the divider where the teacher cannot use it mid-lesson. (7)
- A grammar block with no prep-zone `PREP:` line explaining the concept to the teacher in
  their own terms. (7)
- A spelling-review sentence that could be from a dictionary rather than the students'
  own school/science/sport world, or one that leaves a homophone ambiguous by ear. (2d)
- An `After checking:` / `Answer:` pair that feeds the answer into the question and
  repeats it back, or that a non-confident teacher cannot read aloud without sounding
  foolish. (2d read-aloud test)
- An `After checking:` question whose answer only restates the meaning the prompt just
  handed over (`just means right. What is justice?` / `the state of being right`). The
  answer must carry information the prompt did not. (2d)
- An `After checking:` question that asks a student to SUPPLY the meaning of an untaught
  morpheme (`What does the a at the front do?`, `What does con add?`). Aim the question
  at a taught part, or put the untaught meaning inside the question. (2d, 2f)
- A sound bank slot swapped out, or a morpheme pre-taught, because it was taught long
  ago. The bank is a scaffold students copy; age is not a criterion, need is. (2c)
- A word grid whose size is not 6, 9 or 12 (legacy: 9 or 12), or containing a word
  longer than 12 characters (the grid then renders below the 27 pt floor and the build
  stops). (3b)
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
- A Words to Spell New `extension` note with no `Answer:` line - it has no `ANSWERS:`
  block to fall back on, so the teacher cannot confirm the work. (3e)
- A `CIRCULATE:` prompt that only repeats the rule or points back at the worked example
  the teacher has just read aloud. (3d)
- A check-slide `ASK:` whose answer is visible without thinking. (3d)
- v1.7: new learned words anywhere but Tuesday and Thursday. Legacy: two new learned
  words on consecutive sessions instead of sessions 1 and 3. (5)
- A review word whose card meaning does not operate in it: megaphone on the mega- card
  (the card means million - megaphone means big, not a million). Weeks 6-9 used it; do
  not. The same test catches -ite (dynamite is on the card but names none of its three
  meanings - fine to read, never the answer to an -ite meaning question) and carn
  (carnival and carnation are opaque - spelling targets only). (2b, 4c)
- A sound bank or You Do part drawn from the `_unconfirmed` labels in
  `taught_morphemes.json` (-tion, -sion, -ation, -ment, -ic, -al, -ous, -ly, dis-, de-).
  Weeks 1-5 banked them without the timeline listing them; they stay out until the
  teacher confirms. (2c, 2f)
- A learned-word highlight or Say it line that differs from `teaching_record.json`. (5)
- Brackets inside a dictation sentence - they cannot show red. (6)
- A spec not named `term<T>_week<N>.json`, which silently skips the last-week checks. (9)
- A dictation sized by word count alone, outside 55-75 / 80-100 characters. (6)
- A `punctuation` entry naming a grammatical construct with no plain-English reminder
  in brackets, or a colon used before reported speech where a comma belongs. (6)
- A harder dictation rendering with the needle still near the start. (6)
- Whiteboards, `boards up`, or the boards cue script anywhere in a grammar block. (7)
- A grammar footer that is instruction-shaped noise rather than a tease, stem or
  self-check question. (7)

- A You Do item a student can get right without today's morpheme meaning - plain
  joining like `hyper + link = ?`. (3d, v1.6)
- A new week's You Do or grammar We do / You do authored as one crowded slide plus a
  check slide instead of a round, or a grammar block longer than I do + 1 + 3. (3d, 7)
- A round item whose answer bar shows before the click, or a worked item that hides its
  answer. (3d - the builder checks the first)
- A match-the-meaning bank with no spare word, so the last item is a forced leftover. (3d)

v1.7 additions (Kate Watt, Oct 2026):
- A v1.7 day carrying another day's blocks - spelling or dictation on Monday, a card
  drill on Tuesday or Wednesday, grammar on more than one day. (1)
- A script line or ASK that quizzes the keyword ("What's the keyword?"). Keywords are a
  prompt; the meaning is what is checked. (2a, 3a)
- A grid word an adult would not use (astrolabe), padded in to reach 9 or 12 when 6
  good words were available. (3b)
- A grid word with an untaught part placed before the all-taught words, or with no
  one-breath meaning for that part in its script line. (3b)
- A matrix prefix or suffix printed without its meaning, or a matrix that builds no
  real, useful word. (3f, 2f)
- A name-it item whose answer was not on this week's grid or spelling list, or whose
  hero is a sum of meanings (`carry + across = ?`) instead of a plain definition. (3f)
- A parts-of-speech card whose answer is not printed on the physical card. (3g)
- A Friday with no vote, or a vote with more than four options. (1)

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
- Check slides (rule + up to 4 green answers, no hero) fit comfortably, but their
  answers are BOLD and wrap earlier than that formula says. In Term 3 Week 10 a
  52-character answer wrapped at 20 pt and overlapped the answer below it, while a
  55-character one fitted at 19 pt. Letter shapes decide it, so keep answers near
  `8.2 * 72 / (check_item_size * 0.6)` characters (about 49 at 20 pt, 57 at 17 pt), drop
  `check_item_size` to 17-19 for long ones, and render any line near the limit.
- Rule banners auto-size 18 -> 16 -> 15 pt. In renders, banners of 69 characters or fewer
  always fitted on one line; 70-79 sometimes left one word alone on line two (70, 74, 76,
  77 and 79 did; 71 and 73 did not). Aim for 69 or fewer - shorten, do not pad.
- A hero `example` takes `\n`: each authored line becomes its own centred paragraph.
  Keep it to three lines, and in a multi-line hero keep each line to about 47 characters.
- Routine chips stay on one line under about 30 characters; footers under about 45.
- Hero text drops to 20 pt minimum, and the gate rejects a grammar example below that -
  keep heroes under roughly 180 characters.
- ROUND SLIDES (`draw_round_slide`): rule 0.76, hero 1.10 with chips or 1.62 without,
  chips 0.82, answer bar what is left above 4.56. Keep a prompt to two lines of 47
  characters or fewer, a chip line to 24 (`astronomer` / `-er (person who)` fitted;
  `nom (law) + -er (person)` wrapped to three lines and shrank), and the bar (answer +
  " - " + why) to about 48 characters so it stays near 18 pt - the audit fails over 58
  on a slide with chips. Term 4 Week 2 renders confirmed these numbers.

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

V1.8 SHORT WEEKS NEVER CUT NEW CONTENT (James, 22 Oct 2026: the sessions are "short,
sharp and shiny", about 30 minutes, and nothing important gets dropped; three is the
fewest sessions a week will ever have):

| Sessions | Pattern | Notes |
|---|---|---|
| 5 | teach, spell, build, apply, week_review | the default |
| 4 | teach, spell, build, apply | Friday's review day goes; Thursday carries the vote and ONE dictation sentence in the grammar focus (6) |
| 3 | teach, spell, combined | `combined` = card drill, reading review, word matrix, the week's rotating third round (3f; Term 4 Week 5 ran Complete the sentence, so its build header reads Build words), parts of speech, learned words, grammar, vote - about 30 min. The Break it apart round sits out (leave it out of `build.rounds`); put its words into next week's reading review |

Every day still opens with the card drill and the reading review. The day with no
session stays in `days`/`overview` as "No session this week". Filenames:
`3. Wednesday (ject building and grammar).pptx` for a combined day.

v1.7 mapping (kept for the record; v1.8 weeks use the table above).
Mapping sessions to day types (v1.7, one morpheme a week). Kate Watt: in a short week,
"what can I trim that I lose one week but I put it back in the week after? ... so you're
not trimming off something that you never revisit again." Drop from the END of the
lesson, and put the dropped block back first next week:

| Sessions | Pattern | What waits until next week |
|---|---|---|
| 5 | teach, spell, build, apply, week_review (the default) | nothing |
| 4 | teach, spell, build, week_review (grammar moves onto Friday) | the parts-of-speech drill - open next Monday's card drill with those suffix cards |
| 3 | teach, spell, week_review (grammar onto Friday) | build day - run the matrix next Wednesday with both weeks' morphemes |
| 2 | teach, spell | review, grammar - next week's Monday drill and review lists weight hard toward this morpheme |
| 1 | teach | flag it: next week teaches no new morpheme and runs spell, build, apply, week_review on this one |
| 0 - end-of-term review ("a review of the whole term") | every session `week_review` - see 11a | - |

When the request names two morphemes for one week (the legacy default), teach the first
and carry the second to the next week, unless the user says otherwise - and say so in
the summary. Legacy patterns (two or three morphemes, `new`/`review` days) still build
for old specs; never author them for a new week.

Days with no session: keep them in `days`/`overview` with a "No session this week"
cell so the weekly overview table stays truthful.

## 11a. END-OF-TERM REVIEW WEEK (no new morphemes)

When the Morphology focus is "a review of the whole term", do not stop to ask - this is
the recipe the Term 3 Week 10 decks used (`og_planner/weeks/term3_week10.json` is the
worked example):

- SESSION TYPE: every session is `week_review`. There is no `new_morphology`, grid,
  New Morphology You Do or new learned word, and none are needed. Filenames come out
  `<n>. <Day> (week review).pptx` from the session days (Week 10 ran Tuesday to Friday:
  `1. Tuesday (week review).pptx` ... `4. Friday (week review).pptx`); leave them.
- REVIEW-10 FROM THE WHOLE TERM, not the ten most recent. Spread the term's morphemes
  so each appears on cards once or twice across the week. Give the older weeks the
  second appearance (they dropped off the rolling cards weeks ago), and space repeats
  so a morpheme does not return the next day where you can avoid it. The 2a jumble
  rules still apply.
- TOUCH EVERY TERM MORPHEME EVERY DAY with a complement design. The 10 card morphemes
  get their words on the Words to Read Review board. The morphemes NOT on today's cards
  supply the 10 review spelling words, and the sound bank is derived from those
  (usually 9 of the remaining 11). The one or two left over go on the reading board.
  Mechanical to plan, and nothing is forgotten.
- WORDS: draw on the whole term's grids (`og_history.py --index` shows every past use).
  A word READ on Monday is an ideal SPELLING word on Wednesday. Keep every 2b/2d rule:
  same-day read/spell disjoint, no reading repeat on consecutive sessions.
- LEARNED WORDS: review every word the term taught, split across the sessions (four or
  five a session), grouped by the kind of unfairness (5). Use the settled values in
  `teaching_record.json`.
- GRAMMAR: pair the term's focuses into one day each, in teaching order, choosing pairs
  that genuinely connect: prepositional phrases + adverbials; phrase/clause +
  appositive; simple/compound/complex + trigger words; nominalisation + acronyms. Each
  I do restates both rules (absent-student rule) and reuses the class wording recorded
  in `teaching_record.json`; the You do has a fixed-answer core and a Tick it or fix it
  slide as usual.
- DICTATION: targets from any week at least two back, never last week, never a
  learned word on that day's review list; weave that day's grammar into the sentence.
- OVERVIEW CELLS: `["Term Review", "all <N> term morphemes"]`, the day's grammar pair,
  and the day's review learned words.

---

# 12. USER REQUEST FORMAT (user replaces the XYZs, nothing else)

The user fills in XYZ values only - same convention as MEGA_PROMPT.md. Repeat the
WEEK block for every week of the term. Everything not in this block (derivative
words, review lists, sound banks, spelling words, dictation sentences, grammar
scripts, teacher notes) is YOUR job to generate under the recipes above.

Interpreting the block:
- `Term`: check it against the history before building. If the focus morphemes or the
  "previous weeks" folders say a different term (the Term 3 Week 10 request said Term 4
  while every listed morpheme and folder was Term 3), build the term the history shows
  and flag it in the summary. The week spec and folder name follow the real term.
- `Cohort`: the title slide always says `Enrichment`. The class is a Grade 5/6 composite,
  but the teacher confirmed (Sept 2026) it is known as Enrichment, so a request that
  types "5/6 composite" still gets `"cohort": "Enrichment"`.
- `Morphology focus` names the week's ONE morpheme with its type, e.g. `port (root)`
  (v1.7, section 1). Two listed: teach the first, carry the second (section 11).
- `Learned words` is the week's list in teaching order: the first is new on Tuesday,
  the second on Thursday, both reviewed Friday (section 5).
- `From our reviews` (optional) is the class's wobbles from last week - the Friday
  vote's winner, or a word most boards missed (`aw in drawing, y to i, business`). It is
  ONE line, for the whole class, never per student. Plant each item into next week
  without a new slide: a sound or morpheme goes into Tuesday's dictated bank and the
  review spelling list, a word goes into Friday's or Tuesday's dictation as the
  throwback target, a rule gets one review spelling word that needs it and one
  `After checking:` question about it. Say in the summary where each one landed. Kate
  Watt: "You're wasting your time if you're doing those reviews and you're not acting
  on them." Empty means proceed with defaults.
- `Sessions per week` under 5 compresses per section 11 (v1.8 table); 3 is the minimum.
  Name the days that run if they are not Monday onward.
- `Additional notes` may carry inquiry links, assessment weeks, short weeks, or a report
  that a catalogue transcription is wrong. It may not silently override a captured
  card's keyword or meaning. Verify any reported correction against the physical card,
  update the relevant category JSON and generated views first, then use the corrected
  canon. If `Additional notes` is empty, proceed with defaults - do not ask.
- If `Recently taught morphemes` is empty and the bank/history gives you nothing,
  ask for it - week 1's review deck cannot be invented.
- `Recently taught learned words` should reach back at least three weeks: each
  dictation day needs one harder throwback word from a week or two back (section 6).

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
From our reviews (optional): " XYZ "

WEEK XYZ
Morphology focus: " XYZ "
Learned words: " XYZ "
Grammar focus: " XYZ "
From our reviews (optional): " XYZ "

(repeat for every week of the term)

Additional notes: " XYZ "

Do not enter plan mode, proceed with the deck creation in bypass permissions. Ensure you remain active while the decks are being created and continue to be until they are fully complete, please.```
