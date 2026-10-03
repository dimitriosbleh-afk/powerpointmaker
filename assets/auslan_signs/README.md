# Auslan Sign Image Bank

Shared library of sign images used by Auslan session decks (see
`IMPORTANT/AUSLAN_2_SLIDES_PROMPT.md`). One bank for all Auslan units, because
question signs, greetings, time markers and politeness signs recur every term.

Assets are built automatically from Auslan Signbank:

```bash
# what decks use: an animated GIF of the whole sign
python scripts/fetch_auslan_signs.py --gif --links reference/auslan/signbank_links/<unit>.json
python scripts/fetch_auslan_signs.py --gif --glosses TEAM SCHOOL AGAIN

# still strips, for printables, which cannot animate
python scripts/fetch_auslan_signs.py --from-file assets/auslan_signs/core_glosses.txt
```

**Decks use the GIF.** The teacher's verdict on the still strips was that they
are not clear enough for a student to reproduce a sign from, so a sign card
shows the movement: 14 frames at about 11 a second, 280 pixels tall, starting
and ending at rest, around 250KB a sign. Strips are kept only for paper.

**A links file beats the search.** `--links` takes the teacher's own map of
gloss to Signbank entry URL (JSON, or one `GLOSS url` per line). Those glosses
are fetched from that exact entry and recorded as vetted in the manifest;
everything else falls back to searching, and the run prints which. Chase those
before a deck ships, because a plausible image of the wrong sense is invisible
in a rendered slide.

Sign It! illustrations scanned by the teacher live in `signit/`, cut by
`scripts/crop_signit_scans.py`, with their page numbers in
`signit_manifest.json` so a slide can caption them.

## The images are NOT in git. Rebuild them.

This repo has a public remote, and these images are licensed for internal
school teaching only, so `*.jpg`, `*.png` and `manifest.json` are gitignored.
What IS tracked is the recipe: this README, `core_glosses.txt`, and
`scripts/fetch_auslan_signs.py`. On a fresh clone, or if the folder is ever
emptied, run the command above and the bank rebuilds in a couple of minutes.

Glosses already present are skipped without making a request, so re-running is
free. Adding a new unit means fetching only that unit's new vocabulary; after a
few units most of it is already here.

The script finds each gloss on Signbank, downloads the sign video, picks the
frames that sit inside the sign itself (skipping the signer arriving at and
leaving the rest pose, and preferring sharp frames over motion-blurred ones),
crops to the signer and writes a left-to-right sequence strip that shows the
movement. `--refetch` rebuilds images that already exist.

## Nothing here is a verified sign

The bank is a teaching aid, not an authority. `manifest.json` records, for every
gloss, the exact Signbank entry used, its keywords and its dictionary
definition. That file is the verification worklist: check the entry matches the
meaning the lesson needs, and rehearse from the linked entry (which has the real
video) before teaching. Regional variation is real - where Signbank lists more
than one sign for a word, the extra ones are saved as variants and belong on a
"which one do you use?" slide, not quietly dropped.

## Naming

- One asset per sign, gloss in caps: `TEAM.gif` for decks, `TEAM.jpg` for paper
- Multi-word glosses use hyphens: `THANK-YOU.jpg`, `SLOW-DOWN.jpg`
- Regional or alternate forms are `TEAM_2.jpg`, `TEAM_3.jpg`. The unnumbered
  file is Signbank's first entry, which is not automatically the form your
  school teaches - check it.
- `.gif` holds the animated sign, `.jpg` the photographic strip. Hand-added
  line-art scans may be `.png`. Build scripts resolve any of the three.
- The gloss in the filename must match the gloss in the unit document's
  vocabulary bank exactly. That match is how build scripts find the image.

## Adding images by hand

Scans from the school's own reference (for example Sign It!) are welcome
alongside the Signbank strips - drop them in with the same naming and they take
precedence in review. Scan square-on, crop to one sign with a small margin, at
least ~600 px on the short side. Never mirror or flip an image: that reverses
handedness and teaches the sign wrong. Never crop into the hands or face.

## Licensing

- Auslan Signbank content is CC BY-NC-ND 4.0. The strips here are frame
  extractions, so they are used for internal school teaching only, under the
  Australian schools statutory educational licence, and attributed on every
  deck that places them. Do not redistribute them outside the school and do not
  use them commercially.
- Scans from a purchased reference are copied under the same statutory
  educational licence (see smartcopying.edu.au) and are likewise
  school-internal.
- Every deck placing bank images carries one attribution line on its Teacher
  Resources slide. See `IMPORTANT/AUSLAN_2_SLIDES_PROMPT.md` section 3.

Do not mirror the whole dictionary. Signbank holds 5,172 signs; fetching all of
them would pull roughly 3.7 GB of video off a charity's servers to produce about
680 MB of images, almost none of which any unit would use. Fetch the core list
plus each unit's vocabulary bank.

## Sources on record

| Files | Source | Added |
| --- | --- | --- |
| All `.jpg` strips (222 glosses, 433 files: the core list plus the Deaf Sport unit) | Auslan Signbank, auslan.org.au, CC BY-NC-ND 4.0. Per-file entry URLs in `manifest.json` | 2026-07-29 |
