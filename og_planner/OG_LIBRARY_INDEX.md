# OG Reference Library Index

The `OG/` folder is the program bible: Ron Yoshimoto's Orton-Gillingham training
materials, which this school's OG program is built on. Deck content must agree with
these documents. Almost every PDF is a SCAN with no text layer - to read one, render
pages to images with PyMuPDF (miniconda `python`) and view the images:

For the evidence trail separating teacher-verified physical-card wording from
exact attributed matches and general reference meanings, see
`og_planner/YOSHIMOTO_CARD_SOURCE_RESEARCH.md`.

```python
import fitz
doc = fitz.open("OG/<path>.pdf")
doc[PAGE].get_pixmap(dpi=110).save("scratch/page.png")   # then Read the image
```

## What governs which deck section

| Deck section | Governing documents |
|---|---|
| Morphology review card drill | `Morphology USB files/Morphology - LessonPlan.pdf` (Visual Card Drill: students pronounce, give MEANING and a DERIVATIVE, may use it in a sentence) |
| New morpheme introduction | same LessonPlan.pdf (show card, pronounce, give meaning, write, students brainstorm derivatives, students select a KEYWORD to remember the meaning, write a sentence) |
| Words to Read + extension | template notes (decode -> discuss unknown meanings -> repeat rows for fluency) + `OG Mid Training Resources/Words to Read-Extension Wheel.pdf` (points game: Rhyming 2, Meaning/use-in-sentence 2, Synonym 3, Antonym 3, Nouns-Verbs-Adjectives 4, Analogy 5) |
| Review-day structure | `OG Post Training Resources/OneLessonOver2Days.pdf` (day 2 = short review of the same morpheme, auditory + spelling emphasis) |
| Weekly planning grid | `OG Post Training Resources/Weekly Lesson Plan.pdf` (R:/S: per day, card deck up to, review + new L.W. rows) |
| Morpheme meanings (canon source) | `Morphology USB files/Morphology.pdf` + `MorphologyPart2.pdf` (one worksheet per prefix/suffix, definition box at top), `Lower/Upper Level Latin Scrolls.pdf` (per-root derivative lists + meanings), `Greek Combining Forms.pdf` |
| Teaching order | `Morphology USB files/Morphology - Scope_Sequence.pdf` (2pp: suffixes -> prefixes -> lower roots -> chameleon prefixes -> number prefixes -> continue suffixes -> upper roots -> Greek forms) |
| Spaced/interleaved review rationale | `OG Post Training Resources/Morkunas - Spaced interleaved and retrieval practice.pdf` (evidence base for the weighting engine) |
| Learned (red) words | `OG USB Resources/DOGI RED Words Cards mini.pdf`; template header notes carry the far-point/near-point procedure |

## Folder map

- `Morphology USB files/` - THE morphology core. Worksheet books (Morphology.pdf 227pp,
  MorphologyPart2.pdf 212pp), Latin Scrolls (330 + 284pp), Greek Combining Forms
  (243pp), Scope & Sequence, LessonPlan, plus activity masters: Tachistoscope (build
  words from prefix/suffix strips - red suffix, yellow prefix card stock), Word
  Building #1-3, Prefix Memory #1-2, Suffix Matching/Sorting, Integration (cross-KLA:
  MarineBiology, SocialStudies), Morph Intro (K-1 note), Base Cards.
- `Morphology POST Training Materials/` - print-and-play versions: 6 x Mini Morph Cards
  (morpheme/meaning matching deck), Morph Word Building Game, Prefix Memory Game,
  Morphology Worksheets, Morph Root Words (scientific roots list), 2022 Spelling BWPS,
  Morphology OG Lesson Plan (a filled handwritten exemplar lesson - useful voice
  reference for how much content fits one session).
- `OG Post Training Resources/` - lesson plan formats (OG Lesson Plan, Revised Lesson
  Plan, Primary Lesson Plan + Student Sheet, Weekly Lesson Plan, OneLessonOver2Days),
  Spelling Generalisations by Yoshimoto Scope and Sequence, Bentleigh West PS Scope
  and Sequence, research fact sheets (IDA structured literacy, spelling, dysgraphia,
  RAN, reading comprehension guides).
- `OG Mid Training Resources/` - activity masters: Words to Read Extension Wheel,
  Silent E jobs, Camel Sliders, Floss Game, kck Sorting, Hints and Tricks ee/ea,
  Australian vowels word list, Alphabet Key Word Desk Strip, Rules & generalisations
  sorting game.
- `OG USB Resources/` - card decks (DOGI Green Words, RED Words mini, Green Phonetic
  HF Words), Kindergarten materials (sound bank, phonetic readings, assessments),
  Weekly Lesson Plan white paper.
- `OG Pre Reading/` - background: Characteristics of the OG Approach, Dyslexia
  (Shaywitz), How Our Brains Learn to Read, Dyslexia fact sheets.
- `OG PowerPoint Presentations/` - the 5-day BASIC training attendee slides (Days 2-5).
- `Morphology Presentation for Attendees.pdf` + `Orton Gillingham Lesson Plan
  Template.docx` at the root.

## Standing rules

- The direct photo transcriptions are the first authority for the captured physical
  set: `yoshimoto_cards_suffixes.json` (89), `yoshimoto_cards_prefixes.json` (104),
  and `yoshimoto_cards_latin_roots.json` (84). The readable combined listing is
  `YOSHIMOTO_CARD_CATALOGUE.md`; the three category JSONs, not the generated combined
  view, are the editable source of truth.
- Use each photographed card's exact heading, printed meaning, selected keyword,
  printed part of speech where present, and `associated_words`. Never silently override
  captured metadata, invent a part of speech, or automatically place an
  `excluded_words` item in a lesson. Associated words have been normalised to
  Australian English. See `IMPORTANT/OG_MEGA_PROMPT.md` section 4 for the mandatory
  exception and pre-build audit rules.
- If a card is not captured, consult `morpheme_bank.json`, then
  `morpheme_meanings.json`. Unconfirmed reference entries must not be described as
  exact card wording. Verify genuinely new entries against the relevant worksheet and
  record the source.
- The activity masters (Tachistoscope, Word Building, Prefix Memory, Extension Wheel)
  are legitimate sources when a teacher asks for OG games or homework - point to them
  rather than inventing new formats.
