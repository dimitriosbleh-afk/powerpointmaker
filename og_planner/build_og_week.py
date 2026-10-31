#!/usr/bin/env python3
"""Build Orton-Gillingham session decks from the master template.

Usage:
    python3 og_planner/build_og_week.py og_planner/weeks/<week>.json [--only Monday] [--outdir output]

Reads the week spec JSON (schema documented in IMPORTANT/OG_MEGA_PROMPT.md),
clones slides from og_planner/OG_MASTER_TEMPLATE.pptx at raw-XML level so that
fonts, positions and click animations are preserved exactly. Card and sound-bank
fills are normalised to the school key (root green, prefix yellow, suffix red).
One PPTX per session is written into output/<unit_folder>/.

The master template must not be edited. All content decisions live in the spec.
"""

import base64
import copy
import json
import math
import re
import shutil
import subprocess
import sys
import zipfile
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))


class _LazyModule:
    """Import a v1.9 PDF module (they need PyMuPDF) only when a v1.9 week
    uses it, so v1.8 and older builds keep their old dependencies."""

    def __init__(self, name):
        self._name, self._mod = name, None

    def __getattr__(self, attr):
        if self._mod is None:
            import importlib
            self._mod = importlib.import_module(self._name)
        return getattr(self._mod, attr)


enabler_pdf = _LazyModule("enabler_pdf")              # v1.9 enabler sheets (1b)
friday_review_pdf = _LazyModule("friday_review_pdf")  # v1.9 Friday Review (1a)
print_pack = _LazyModule("print_pack")                # v1.9 week print pack (1c)

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def q(tag):
    pfx, local = tag.split(":")
    return "{%s}%s" % (NS[pfx], local)


def make_compatible_yellow_meter(green_png, yellow_png):
    """Normalise the master's real yellow gauge to the green gauge's geometry.

    The difficulty cue is the NEEDLE POSITION, not just the colour: the green
    gauge sits near the start, the yellow one sweeps about two thirds around
    with everything before the needle filled. An earlier fix recoloured the
    green gauge instead, which silently threw that away and left every harder
    dictation showing an easy meter.

    The real defect was a one-pixel height mismatch (170x89 vs 170x88) against
    the shape extent, which made LibreOffice shift/clamp other slide elements
    during PDF export. Resizing the genuine yellow artwork to the green gauge's
    exact size and re-encoding it keeps the correct needle and drops the defect.
    """
    yellow = Image.open(BytesIO(yellow_png)).convert("RGBA")
    green = Image.open(BytesIO(green_png)).convert("RGBA")
    if yellow.size != green.size:
        yellow = yellow.resize(green.size, Image.LANCZOS)
    output = BytesIO()
    yellow.save(output, format="PNG", dpi=(150, 150))
    return output.getvalue()


ROOT_DIR = Path(__file__).resolve().parent
MASTER = ROOT_DIR / "OG_MASTER_TEMPLATE.pptx"

# ---------------------------------------------------------------- template map
# 1-based slide indices in the master template.
T = {
    "title": 1,
    "overview": 2,
    "hdr_morph_review": 3,
    "card_review_root": 4,
    "card_review_prefix": 5,
    "card_review_suffix": 6,
    "hdr_wtr_review": 7,
    "wtr_review_table": 8,
    "hdr_sound_bank": 9,
    "sound_bank": 10,
    "hdr_wts_review": 11,
    "spell_word": 12,
    "hdr_new_morph": 13,
    "card_new_root": 14,
    "card_new_prefix": 15,
    "card_new_suffix": 16,
    "hdr_wtr_new": 17,
    "wtr_new_grid": 18,
    "hdr_wts_new": 19,
    "wts_new_grid": 20,
    "hdr_review_lw": 21,
    "learned_word": 22,
    "hdr_new_lw": 24,
    "learned_word_new": 25,
    "hdr_dictation": 26,
    "dictation": {"green": 27, "yellow": 28},
    "hdr_grammar": 29,
    "gradual_release": {"i_do": 30, "we_do": 31, "you_do": 32},
}

# Shape ids inside each template slide (stable - taken from the master).
SHAPE = {
    "title_heading": 6391,       # 'OG Planner'
    "title_sub": 6392,           # 'Enrichment / Term 3 Week XYZ'
    "overview_table": 6406,
    "card_review_word": {4: 6431, 5: 6439, 6: 6447},
    "wtr_review_tbl": 6465,
    "sound_boxes": [6482, 6483, 6484, 6486, 6485, 6487, 6488, 6489, 6490],
    # row-major order: (r1: x0.74,3.84,6.93) (r2: 0.74=6486, 3.84=6485, 6.93=6487) (r3)
    "spell_word_box": 6505,
    "hdr_new_morph_txt": 6512,
    "card_new_word": {14: 6523, 15: 6531, 16: 6539},
    "hdr_wtr_new_txt": 6544,
    "wtr_cols": [6556, 6555, 6554],  # left, middle, right
    "hdr_wts_new_txt": 6561,
    "wts_boxes": [6571, 6572],       # left (2 paras), right (2 paras)
    "lw_word": {22: 6598, 25: 6634},
    "dictation_title": {27: 6660, 28: 6671},
    "dictation_sentence": {27: 6663, 28: 6675},
    "dictation_meter": {27: 6664, 28: 6674},
    "gr_title": {30: 6687, 31: 6693, 32: 6699},
    "gr_body": {30: 6688, 31: 6694, 32: 6700},
}

MORPH_FILL = {"root": "ABEA90", "prefix": "FFF2CC", "suffix": "E06666"}
CARD_REVIEW = {"root": 4, "prefix": 5, "suffix": 6}
CARD_NEW = {"root": 14, "prefix": 15, "suffix": 16}

GRAMMAR_PURPLE = "5E58A6"   # sampled from the template's Grammar section header
GRAMMAR_TINT = "EFEEF8"
MORPH_MAGENTA = "A64D79"    # sampled from the New Morphology section header
MORPH_TINT = "F6EBF1"
CHARCOAL = "26262B"
PUNCT_RED = "C00000"
CAPITAL_GREEN = "008A45"
ANSWER_GREEN = "008A45"

HEADER_NOTES = {
    "morph_review": (
        "PROCEDURE:\n"
        "1. SHOW one card at a time in a mixed order.\n"
        "2. CHORAL: \"Everyone, together, on three.\" The class says it, like \"-tude means condition of\". Suffix cards add the part of speech.\n"
        "3. ASK one student for a word with it, and what that word means.\n"
        "4. REPEAT any uncertain card at the end. Keywords are a prompt, never a quiz."
    ),
    "wtr_review": (
        "PROCEDURE:\n"
        "1. READ all rows together before asking the review questions.\n"
        "2. CLARIFY any unfamiliar word briefly.\n"
        "3. READ selected rows again for fluency."
    ),
    "sound_bank": (
        "PROCEDURE:\n"
        "1. SAY each sound and have students repeat it.\n"
        "2. WRITE the matching morpheme before revealing the next one.\n"
        "3. CHECK and fix each response immediately."
    ),
    "wts_review": "Students spell each review word, then use the reveal to check and fix.",
    "new_morph": (
        "PROCEDURE:\n"
        "1. SHOW the morpheme card and say the sound.\n"
        "2. TRACE three times while saying the sound.\n"
        "3. GIVE the letter names, keyword and card meaning.\n"
        "4. WRITE the morpheme three times and underline it in each new word."
    ),
    "wtr_new": "Read each word as it appears, then connect the morpheme to the whole-word meaning.",
    "wts_new": "Students spell the four selected new-family words, then complete the extension.",
    "morph_activity": "Students apply the new morpheme independently before the answer-check slide.",
    "learned_review": (
        "PROCEDURE:\n"
        "1. SAY the word and have students repeat it.\n"
        "2. NAME the unfair part and say the letter names.\n"
        "3. WRITE three times using far-point practice.\n"
        "4. COVER, write from memory, then check and fix."
    ),
    "learned_new": (
        "PROCEDURE:\n"
        "1. SAY the new word and have students repeat it.\n"
        "2. NAME the unfair part and explain why it is unexpected.\n"
        "3. WRITE three times using far-point practice.\n"
        "4. COVER, write from memory, then check and fix."
    ),
    "dictation": (
        "PROCEDURE:\n"
        "1. REVIEW the target morphemes, learned word and punctuation.\n"
        "2. SAY the whole sentence; students repeat it before writing.\n"
        "3. REPEAT by phrase only when the sentence is long.\n"
        "4. REVEAL after all writing is complete.\n"
        "5. TICK capitals, order, punctuation and spelling targets.\n"
        "6. FIX every error before moving on."
    ),
    # v1.7 five-day shape (Kate Watt, Oct 2026)
    "new_morph_v17": (
        "PROCEDURE:\n"
        "1. SHOW the card, say the morpheme and its meaning; the class repeats.\n"
        "2. COPY: students copy the card, then write the morpheme three times, saying it.\n"
        "3. ASK the class for words they already know with it, then show the grid."
    ),
    "sound_bank_v17": (
        "PROCEDURE:\n"
        "1. SAY each morpheme or sound in the order in the slide notes.\n"
        "2. WRITE: students write it in their books before you click.\n"
        "3. REVEAL one box per click; students tick or fix straight away."
    ),
    "recap": (
        "PROCEDURE:\n"
        "1. SHOW yesterday's card.\n"
        "2. CHORAL: \"Everyone, together, on three.\" The class says the morpheme and its meaning.\n"
        "3. ASK two students for a word with it, and what each word means."
    ),
    "build": (
        "PROCEDURE:\n"
        "1. BUILD real words from the matrix in OG books, then say what each means.\n"
        "2. BREAK whole words back into their parts.\n"
        "3. NAME the word from its meaning on whiteboards, answer on the click."
    ),
    "pos_drill": (
        "PROCEDURE:\n"
        "1. SHOW one suffix card at a time.\n"
        "2. CHORAL: \"Everyone, together, on three.\" The class says its meaning and the part of speech it makes.\n"
        "3. REVEAL the part of speech on the click; tick or fix.\n"
        "4. ASK one student for a word with that suffix, used in a short sentence."
    ),
    # v1.8 (James Hooke's classroom feedback, 23 Oct 2026)
    "new_morph_v18": "New morpheme today. The memory hook and its picture are on the next slide.",
    "pos_drill_v18": (
        "PROCEDURE:\n"
        "1. READ the suffix, its meaning and the sentence together.\n"
        "2. ASK what job the bold word does in the sentence: noun, verb or adjective. Turn and talk, then choral.\n"
        "3. REVEAL on the click, then say why, using the sentence and the suffix's meaning."
    ),
    "learned_review_v18": (
        "PROCEDURE:\n"
        "1. SAY the word in a sentence; students write it once from hearing it.\n"
        "2. SHOW the word; students copy it correctly underneath.\n"
        "3. COVER both; students write it from memory, then check and fix."
    ),
    "learned_new_v18": (
        "PROCEDURE:\n"
        "1. SAY the new word and name its unfair part.\n"
        "2. COPY it twice, saying the letter names.\n"
        "3. COVER both; write it a third time from memory, then check and fix."
    ),
}

# v1.7 (Kate Watt, Oct 2026): one morpheme a week, every day for 20-30 minutes,
# one lesson split across the week the way Yoshimoto's "One Lesson Over 2 Days"
# sheet splits it.  Legacy `new` / `review` weeks still build unchanged.
V17_TYPES = {"teach", "spell", "build", "apply", "combined"}
SOUND_FILL = "FFFFFF"   # phonogram (sound) boxes stay white, like the sound deck

WARNINGS = []
NOTES = []
# Folder holding the week specs; main() points it at the built spec's folder so
# the dictation gate can read last week's spec (section 6).
SPEC_DIR = Path(__file__).resolve().parent / "weeks"


def warn(msg):
    WARNINGS.append(msg)
    print(f"  WARN: {msg}")


def note(msg):
    """Advisory finding: printed and summarised, but does not fail the gate.

    For checks that rely on a heuristic (substring morpheme matching) where a
    false positive would block a legitimate build.  Every WARN is fatal; a NOTE
    is a content question for the author to answer.
    """
    NOTES.append(msg)
    print(f"  NOTE: {msg}")


_BANK = None
_MEANINGS = None
_MEANINGS_ANY = None
_PHOTO_CARDS = None


def bank():
    """morph -> legacy verified-card entry."""
    global _BANK
    if _BANK is None:
        data = json.loads((ROOT_DIR / "morpheme_bank.json").read_text())
        _BANK = {m["morph"]: m for m in data["morphemes"]}
    return _BANK


def canon_morph_key(morph):
    """Normalize display punctuation without collapsing genuine variant groups."""
    return str(morph).strip().strip("-").replace(" ", "")


def photo_card_catalogue():
    """(type, normalized morph) -> photographed Yoshimoto card entries.

    A list is retained for each key because the physical set contains genuine
    collisions, notably two different di- cards.  The photographed wording is
    authoritative; the legacy bank is only a compatibility fallback.
    """
    global _PHOTO_CARDS
    if _PHOTO_CARDS is None:
        files = {
            "suffix": "yoshimoto_cards_suffixes.json",
            "prefix": "yoshimoto_cards_prefixes.json",
            "root": "yoshimoto_cards_latin_roots.json",
        }
        catalogue = {}
        for morph_type, filename in files.items():
            data = json.loads((ROOT_DIR / filename).read_text(encoding="utf-8"))
            for card in data["cards"]:
                entry = {
                    "morph": card["morpheme"],
                    "type": morph_type,
                    "meaning": card["meaning"],
                    "keyword": card.get("keyword", ""),
                    "part_of_speech": card.get("part_of_speech", []),
                    "associated_words": card.get("associated_words", []),
                    "excluded_words": card.get("excluded_words", []),
                    "source": f"Photographed Yoshimoto card: {card['source_image']}",
                    "verified": True,
                }
                key = (morph_type, canon_morph_key(card["morpheme"]))
                catalogue.setdefault(key, []).append(entry)
        _PHOTO_CARDS = catalogue
    return _PHOTO_CARDS


def photo_card_entries(morph, morph_type=None):
    key = canon_morph_key(morph)
    catalogue = photo_card_catalogue()
    if morph_type:
        return catalogue.get((morph_type, key), [])
    matches = []
    for (entry_type, entry_key), entries in catalogue.items():
        if entry_key == key:
            matches.extend(entries)
    return matches


def meaning_catalogue():
    """(type, normalized morph) -> canonical meaning from the complete scope."""
    global _MEANINGS, _MEANINGS_ANY
    if _MEANINGS is None:
        data = json.loads((ROOT_DIR / "morpheme_meanings.json").read_text())
        section_types = {
            "suffixes_lower": "suffix",
            "continue_suffixes": "suffix",
            "suffix_card_additions": "suffix",
            "prefixes_first": "prefix",
            "continue_prefixes_chameleon": "prefix",
            "continue_prefixes_other": "prefix",
            "number_prefixes": "prefix",
            "lower_level_roots_start": "root",
            "lower_level_roots_continue": "root",
            "upper_level_roots": "root",
            "greek_combining_forms": "root",
        }
        typed = {}
        any_type = {}
        for section, morph_type in section_types.items():
            for morph, meaning in data[section].items():
                key = canon_morph_key(morph)
                typed[(morph_type, key)] = meaning
                any_type.setdefault(key, set()).add(meaning)
        _MEANINGS = typed
        _MEANINGS_ANY = {
            key: next(iter(values)) for key, values in any_type.items()
            if len(values) == 1
        }
    return _MEANINGS


def catalogue_meaning(morph, morph_type=None):
    typed = meaning_catalogue()
    key = canon_morph_key(morph)
    return typed.get((morph_type, key)) or _MEANINGS_ANY.get(key)


def verified_card_entry(morph, morph_type=None, requested_meaning=None):
    """Return a teacher/card-confirmed entry, tolerant of display hyphens."""
    photographed = photo_card_entries(morph, morph_type)
    if len(photographed) == 1:
        return photographed[0]
    if len(photographed) > 1:
        matching = [entry for entry in photographed
                    if requested_meaning and entry["meaning"] == requested_meaning]
        return matching[0] if len(matching) == 1 else None

    direct = bank().get(morph)
    if direct and (not morph_type or direct.get("type") == morph_type):
        return direct
    key = canon_morph_key(morph)
    for entry in bank().values():
        if (not morph_type or entry.get("type") == morph_type) and canon_morph_key(entry["morph"]) == key:
            return entry
    return None


def resolved_photo_card(morph, morph_type=None, requested_meaning=None):
    """Return the one intended photographed card, including collision handling."""
    entries = photo_card_entries(morph, morph_type)
    if len(entries) == 1:
        return entries[0]
    if len(entries) > 1 and requested_meaning:
        matches = [entry for entry in entries
                   if entry["meaning"] == requested_meaning]
        if len(matches) == 1:
            return matches[0]
    return None


def check_banked(day, where, morph, morph_type=None, requested_meaning=None):
    photographed = photo_card_entries(morph, morph_type)
    if len(photographed) > 1:
        matching = [entry for entry in photographed
                    if requested_meaning and entry["meaning"] == requested_meaning]
        if len(matching) == 1:
            return
        meanings = "; ".join(sorted({entry["meaning"] for entry in photographed}))
        warn(f"{day}: {where} morph {morph!r} matches multiple photographed cards "
             f"({meanings}); include the intended card meaning in the spec")
        return
    if verified_card_entry(morph, morph_type):
        return
    if catalogue_meaning(morph, morph_type):
        warn(f"{day}: {where} morph {morph!r} has a reference meaning, but its "
             "exact physical-card wording has not been confirmed")
    else:
        warn(f"{day}: {where} morph {morph!r} has no verified meaning in the "
             "photographed catalogue, morpheme_bank.json, or morpheme_meanings.json")


def resolve_card(day, where, card):
    """Resolve optional legacy keyword plus required canonical meaning."""
    kw = card.get("keyword", "")
    mean = card.get("meaning", "")
    photographed = photo_card_entries(card["morph"], card.get("type"))
    entry = verified_card_entry(card["morph"], card.get("type"), mean)
    if len(photographed) > 1 and not entry:
        meanings = "; ".join(sorted({item["meaning"] for item in photographed}))
        warn(f"{day}: {where} {card['morph']!r} is ambiguous in the photographed "
             f"card set ({meanings}); supplied keyword and meaning retained")
        return kw, mean
    if entry:
        if entry.get("keyword"):
            if kw and kw != entry["keyword"]:
                warn(f"{day}: {where} {card['morph']!r} spec keyword {kw!r} differs "
                     f"from canonical card value {entry['keyword']!r} - canonical "
                     "value used")
            kw = entry["keyword"]
    if entry:
        if mean and mean != entry["meaning"]:
            warn(f"{day}: {where} {card['morph']!r} spec meaning differs from "
                 "the confirmed physical-card meaning - confirmed value used")
        mean = entry["meaning"]
    else:
        reference_meaning = catalogue_meaning(card["morph"], card.get("type"))
        if reference_meaning:
            if mean and mean != reference_meaning:
                warn(f"{day}: {where} {card['morph']!r} spec meaning differs from "
                     "the unconfirmed reference catalogue - reference value used")
            mean = reference_meaning
    return kw, mean


def card_pos_note(card):
    """Teacher-note line for part of speech printed on the physical card."""
    parts = card_parts_of_speech(card)
    if not parts:
        return ""
    return f"\nPart of speech: {', '.join(parts)}"


def card_parts_of_speech(card):
    """Part-of-speech labels printed on the photographed physical card."""
    entry = resolved_photo_card(card["morph"], card.get("type"),
                                card.get("meaning"))
    return entry.get("part_of_speech", []) if entry else []


def card_meaning_note(meaning):
    """Present a meaning card as clear note lines, including sound-pattern cards."""
    text = str(meaning or "").strip()
    pattern = re.fullmatch(
        r"ending that says\s+([^()]+?)\s*\(the\s+(.+?)\s+part means\s+(.+?)\)",
        text,
        flags=re.IGNORECASE,
    )
    if pattern:
        sound, part, part_meaning = (value.strip() for value in pattern.groups())
        return f"Sound: {sound}\nMeaning of {part}: {part_meaning}"
    pattern = re.fullmatch(r"(?:ending that )?says\s+(.+)", text,
                           flags=re.IGNORECASE)
    if pattern:
        return f"Sound: {pattern.group(1).strip()}"
    return f"Meaning: {text}"


def extra_task_note(card):
    """Optional card retrieval prompt in clear teacher-facing language."""
    task = card.get("extra_task")
    if isinstance(task, dict):
        prompt = str(task.get("prompt", "")).strip()
        answers = task.get("answers", [])
        answer_text = ", ".join(str(answer) for answer in answers)
    else:
        legacy = str(card.get("derivative_ask", "")).strip()
        if not legacy:
            return ""
        warn("legacy derivative_ask is not allowed; use an extra_task object "
             "with prompt and answers")
        prompt, _, answer_text = legacy.partition("EXPECT:")
        prompt = prompt.strip()
        answer_text = answer_text.strip()
    if not prompt:
        return ""
    prompt = prompt[0].upper() + prompt[1:]
    if prompt[-1] not in ".?!":
        prompt += "."
    lines = ["Extra task:", f"{'ASK' if prompt.endswith('?') else 'SAY'}: \"{prompt}\""]
    if answer_text:
        lines.append(f"Possible answers: {answer_text}")
    return "\n" + "\n".join(lines)


def validate_session_card_words(day, session):
    """Enforce photographed-card word families for the new morphology block."""
    nm = session.get("new_morphology")
    if not nm:
        return
    entry = resolved_photo_card(nm["morph"], nm.get("type"), nm.get("meaning"))
    if not entry:
        return  # Uncaptured/ambiguous cards are reported by check_banked().

    exceptions = {}
    for item in session.get("associated_word_exceptions", []):
        if not isinstance(item, dict) or not item.get("word"):
            warn(f"{day}: malformed associated_word_exceptions entry {item!r}")
            continue
        word_key = str(item["word"]).strip().casefold()
        if not item.get("reason") or not item.get("source"):
            warn(f"{day}: outside derivative {item['word']!r} needs both reason and source")
        exceptions[word_key] = item

    allowed = {str(word).strip().casefold()
               for word in entry.get("associated_words", [])}
    excluded = {}
    for item in entry.get("excluded_words", []):
        if isinstance(item, dict):
            excluded[str(item.get("word", "")).strip().casefold()] = item.get("reason", "")
        else:
            excluded[str(item).strip().casefold()] = ""

    new_read = [item.get("word", "") if isinstance(item, dict) else str(item)
                for item in session.get("words_to_read_new", [])]
    selected = {str(word).strip().casefold() for word in new_read}
    for word in new_read:
        word_key = str(word).strip().casefold()
        if word_key in excluded:
            reason = excluded[word_key]
            suffix = f" ({reason})" if reason else ""
            warn(f"{day}: new-family word {word!r} is excluded from automatic lesson use{suffix}")
        elif word_key not in allowed and word_key not in exceptions:
            warn(f"{day}: new-family word {word!r} is not in the photographed card's "
                 "associated_words and has no verified associated_word_exceptions entry")

    for word_key, item in exceptions.items():
        if word_key not in selected:
            warn(f"{day}: associated-word exception {item['word']!r} is not used in "
                 "words_to_read_new")

    new_spell = [str(word).strip().casefold()
                 for word in session.get("words_to_spell_new", [])]
    for word in new_spell:
        if word not in selected:
            warn(f"{day}: new spelling word {word!r} is not in words_to_read_new")

    review_read = {str(word).strip().casefold()
                   for word in session.get("words_to_read_review", {}).get("words", [])}
    review_spell = {str(item.get("word", "")).strip().casefold()
                    for item in session.get("words_to_spell_review", [])}
    overlap = sorted(review_read & review_spell)
    if overlap:
        warn(f"{day}: same-day review reading/spelling overlap: {', '.join(overlap)}")


# ------------------------------------------------- taught-morpheme bookkeeping
# Everyday inflections a Grade 5/6 student uses without a Yoshimoto card.
UNIVERSAL_INFLECTIONS = {"s", "es", "ed", "ing"}

# Morpheme notation inside a You Do item: an affix carrying its hyphen, or a
# slash variant label.  Anchored on non-letters so "T-intersection" and
# "self-check" cannot masquerade as affixes.
AFFIX_TOKEN_RE = re.compile(r"(?<![A-Za-z-])(-[A-Za-z]{1,12}|[A-Za-z]{1,12}-)(?![A-Za-z])")
# A You Do prompt: optional worked-example label, optional "1." numbering.
PROMPT_LABEL_RE = re.compile(r"^\s*(?:one done for you:\s*)?(?:\d+\.\s*)?", re.I)
# A printed meaning beside a part: `in- (onto)`, `vore (eat)`.
GLOSS_RE = re.compile(r"\([^)]*\)")


def morph_surface_forms(label):
    """Lowercase spellings a morpheme label can wear inside a real word.

    Expands slash variants, optional-letter brackets (clud(e) -> clud, clude)
    and the silent e that drops before a vowel suffix (vore -> vor, so
    herbivorous still counts as a vore word).

    The silent-e stem is kept only at 3+ letters.  Without that floor, -ine
    yields "in" and -ice yields "ic", which appear inside half the English
    language and would make any substring test meaningless.
    """
    forms = set()
    for part in str(label or "").split("/"):
        part = part.strip().strip("-").strip().lower()
        if not part:
            continue
        candidates = {part}
        if "(" in part:
            candidates.add(re.sub(r"[()]", "", part))
            candidates.add(re.sub(r"\([^)]*\)", "", part))
        for candidate in list(candidates):
            if candidate.endswith("e") and len(candidate) >= 4:
                candidates.add(candidate[:-1])
        forms |= {c for c in candidates if c}
    return forms


def session_morph_labels(session):
    """Every morpheme label a session puts in front of the class."""
    labels = []
    focus = (session.get("new_morphology") or {}).get("morph")
    if focus:
        labels.append(focus)
    labels += [card["morph"] for card in session.get("morphology_review", [])
               if card.get("morph")]
    labels += [item["morph"] for item in session.get("sound_bank", [])
               if item.get("morph")]
    return labels


def taught_morpheme_forms(week, session):
    """Surface forms of every morpheme taught on or before this session.

    Sources, in the order section 2f of OG_MEGA_PROMPT.md lists them: today's
    focus card, this session's review cards and sound bank, every earlier
    session in the same week spec, and the week-level `taught_morphemes` list
    the author fills from the term timeline.
    """
    forms = set()
    for item in week.get("taught_morphemes", []):
        label = item.get("morph") if isinstance(item, dict) else item
        forms |= morph_surface_forms(label)
    sessions = week.get("sessions", [])
    try:
        stop = sessions.index(session)
    except ValueError:
        stop = len(sessions) - 1
    for earlier in sessions[:stop + 1]:
        for label in session_morph_labels(earlier):
            forms |= morph_surface_forms(label)
    return forms


def untaught_week_focus_labels(week, session):
    """Focus morphemes this week spec teaches AFTER `session`.

    A label counts as untaught only if no session up to and including this one
    has it as its focus - a `review` day repeats yesterday's focus and that
    morpheme is genuinely taught by then.
    """
    sessions = (week or {}).get("sessions", [])
    try:
        stop = sessions.index(session)
    except ValueError:
        return []
    seen = set()
    for earlier in sessions[:stop + 1]:
        seen |= morph_surface_forms((earlier.get("new_morphology") or {}).get("morph"))
    later = []
    for upcoming in sessions[stop + 1:]:
        label = (upcoming.get("new_morphology") or {}).get("morph")
        if label and not (morph_surface_forms(label) & seen):
            later.append(label)
    return later


def validate_sound_bank(day, session, week=None):
    """Sound bank rules from OG_MEGA_PROMPT.md section 2c (and the 2a card rule).

    The bank exists to support the Words to Spell Review list on the slides
    immediately after it, and it is copied into books BEFORE the New Morphology
    section runs.  So it never carries the day's focus morpheme, never carries a
    morpheme this week has not taught yet, and every box should serve a word on
    that page.  The review-10 cards get the same not-yet-taught check.
    """
    bank_items = session.get("sound_bank", [])
    review_cards = session.get("morphology_review", [])
    focus = (session.get("new_morphology") or {}).get("morph")
    if focus:
        focus_forms = morph_surface_forms(focus)
        for item in bank_items:
            if morph_surface_forms(item.get("morph")) & focus_forms:
                warn(f"{day}: sound bank contains the day's focus morpheme "
                     f"{item['morph']!r}. The bank is copied into books before the "
                     "morpheme is taught (new day) or defeats the retrieval it "
                     "exists for (review day) - give the slot to a morpheme one of "
                     "the review spelling words needs (2c)")
        if session.get("type", "new") == "new":
            for card in review_cards:
                if morph_surface_forms(card.get("morph")) & focus_forms:
                    warn(f"{day}: morphology review card {card['morph']!r} is today's "
                         "NEW morpheme - it is drilled a dozen slides before it is "
                         "taught. The review-10 excludes today's new morpheme; it "
                         "enters the deck tomorrow (2a)")
    for label in untaught_week_focus_labels(week, session):
        later_forms = morph_surface_forms(label)
        for where, items in (("sound bank", bank_items),
                             ("morphology review card", review_cards)):
            for item in items:
                if morph_surface_forms(item.get("morph")) & later_forms:
                    warn(f"{day}: {where} contains {item['morph']!r}, which this week "
                         "does not teach until a later session. Students meet it in "
                         "the bank/drill before the morpheme card - swap it for a "
                         "morpheme already taught (2a/2c)")
    if not bank_items:
        return

    spell_words = [str(entry.get("word", "")).strip().lower()
                   for entry in session.get("words_to_spell_review", [])]
    if week_is_v17(week):
        # v1.7 Tuesday: the dictated bank also serves the four new words
        spell_words += [str(w).strip().lower()
                        for w in session.get("words_to_spell_new", []) or []]
    spell_words = [word for word in spell_words if word]
    if not spell_words:
        return
    for item in bank_items:
        forms = morph_surface_forms(item.get("morph"))
        if not any(form in word for form in forms for word in spell_words):
            # Advisory: this is a substring test, so a morpheme that surfaces in
            # an unusual spelling can read as unused when it is not.
            note(f"{day}: sound bank box {item['morph']!r} appears in none of the "
                 "review spelling words on the same page; the 9 boxes should be "
                 "derived from that list (2c)")


def week_book_words(week):
    """Every word this week's spec puts in students' books as NEW material:
    the Words to Read / Words to Spell grids and the new learned words."""
    words = set()
    for session in (week or {}).get("sessions", []):
        for item in session.get("words_to_read_new", []):
            word = item.get("word") if isinstance(item, dict) else item
            words.add(str(word or "").strip().lower())
        for word in session.get("words_to_spell_new", []) or []:
            words.add(str(word or "").strip().lower())
        new_lw = (session.get("learned_words") or {}).get("new") or {}
        words.add(str(new_lw.get("word") or "").strip().lower())
    words.discard("")
    return words


def previous_week_spec(week):
    """The spec for the week before this one, if it sits in SPEC_DIR."""
    try:
        term, num = int(week.get("term")), int(week.get("week"))
    except (TypeError, ValueError):
        return None
    path = SPEC_DIR / f"term{term}_week{num - 1}.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def validate_dictation_sources(day, week, session):
    """Dictation sourcing rule from OG_MEGA_PROMPT.md section 6.

    Dictation targets are revision from 2-3 weeks ago.  Words taught THIS week
    are in the front of the students' books and get copied, not spelled (school
    feedback, Aug 2026).  Exact matches with this week's new material and
    targets built on this week's focus morphemes are fatal; last week's words
    are an advisory NOTE because the ideal reach is 2-3 weeks back.
    """
    dictations = session.get("dictation") or []
    if not dictations:
        return
    if week_is_v17(week):
        validate_dictation_sources_v17(day, week, session)
        return
    this_week = week_book_words(week)
    prev = previous_week_spec(week)
    last_week = week_book_words(prev) if prev else set()
    review_lw = {str(item.get("word") or "").strip().lower()
                 for item in (session.get("learned_words") or {}).get("review", [])}
    focus_forms = {}
    for other in (week or {}).get("sessions", []):
        label = (other.get("new_morphology") or {}).get("morph")
        if label:
            focus_forms[label] = morph_surface_forms(label)
    for entry in dictations:
        for target in entry.get("targets", []) or []:
            word = str(target or "").strip().lower()
            stems = {word}
            for infl in ("es", "s", "ed", "ing", "d", "ly"):
                if word.endswith(infl) and len(word) - len(infl) >= 3:
                    stems.add(word[:-len(infl)])
            if stems & this_week:
                warn(f"{day}: dictation target {target!r} is one of THIS week's new "
                     "words (grid, spelling grid or new learned word). Students copy "
                     "it from their books; dictation targets come from 2-3 weeks "
                     "ago (6)")
                continue
            hit = None
            for label, forms in focus_forms.items():
                for form in forms:
                    if form in word:
                        hit = (label, form)
                        break
                if hit:
                    break
            if hit:
                label, form = hit
                if len(form) >= 4:
                    warn(f"{day}: dictation target {target!r} is built on this "
                         f"week's focus morpheme {label!r}. Dictation revises "
                         "morphemes taught 2-3 weeks ago, never this week's (6)")
                else:
                    note(f"{day}: dictation target {target!r} may contain this "
                         f"week's focus morpheme {label!r} (short-form substring "
                         "match). If it is a this-week word, replace it (6)")
                continue
            if stems & last_week:
                note(f"{day}: dictation target {target!r} was new LAST week; the "
                     "ideal dictation reach is 2-3 weeks back (6)")
            elif word in review_lw:
                note(f"{day}: dictation target {target!r} is a learned word still "
                     "on the review pair; prefer a learned word from 2-3 weeks "
                     "ago (6)")


def week_is_v17(week):
    """True for a v1.7 five-day-shape week (any teach/spell/build/apply day)."""
    return any(s.get("type") in V17_TYPES for s in (week or {}).get("sessions", []))


def week_is_v18(week):
    """True for a v1.8 week (`"og_version": "1.8"`): the daily card drill and
    Words to Read Review open every session (James Hooke, 23 Oct 2026)."""
    return og_version(week) >= (1, 8) and week_is_v17(week)


def og_version(week):
    """`"og_version": "1.9"` as (1, 9), so "1.10" sorts after "1.9"."""
    parts = re.findall(r"\d+", str((week or {}).get("og_version", "")))
    return tuple(int(p) for p in parts[:2]) if parts else (0, 0)


def week_is_v19(week):
    """True for a v1.9 week: meaning pictures, enabler sheets and the Friday
    Review worksheet (ported from the 5/6 Challenge team's package, Oct 2026)."""
    return og_version(week) >= (1, 9) and week_is_v17(week)


def week_focus_labels(week):
    """The week's taught focus morphemes (v1.7 teaches them on `teach` days)."""
    return [s["new_morphology"]["morph"] for s in (week or {}).get("sessions", [])
            if (s.get("new_morphology") or {}).get("morph")]


def taught_so_far_words(week, session):
    """This week's new material up to and including `session`: grid words,
    spelling-grid words and new learned words."""
    sessions = (week or {}).get("sessions", [])
    stop = sessions.index(session) if session in sessions else len(sessions) - 1
    words = set()
    for earlier in sessions[:stop + 1]:
        words |= week_book_words({"sessions": [earlier]})
    return words


def validate_dictation_sources_v17(day, week, session):
    """Dictation sourcing rule from OG_MEGA_PROMPT.md section 6 (v1.7).

    Kate Watt (Oct 2026) and Yoshimoto's own planning slide: dictation "must
    prioritise what was introduced in this lesson and can also include anything
    previously taught".  Copying from the book is fine.  So at least one target
    must be this week's teaching (a word built on the focus morpheme or a new
    learned word taught so far), and one harder word from 1-2+ weeks back keeps
    it honest.  Missing this week's teaching is fatal; a dictation made only of
    this week's words is a NOTE.
    """
    this_week = taught_so_far_words(week, session)
    focus = set()
    for label in week_focus_labels(week):
        focus |= {f for f in morph_surface_forms(label) if len(f) >= 3}
    current, older = [], []
    for entry in session.get("dictation") or []:
        for target in entry.get("targets", []) or []:
            word = str(target or "").strip().lower()
            stems = {word}
            for infl in ("es", "s", "ed", "ing", "d", "ly"):
                if word.endswith(infl) and len(word) - len(infl) >= 3:
                    stems.add(word[:-len(infl)])
            if stems & this_week or any(form in word for form in focus):
                current.append(target)
            else:
                older.append(target)
    if not current:
        warn(f"{day}: no dictation target is this week's teaching. Dictation must "
             "prioritise what was just taught - a word built on this week's "
             "morpheme or the new learned word (6, v1.7)")
    if not older:
        note(f"{day}: every dictation target is this week's - add one harder word "
             "from a week or two back so the dictation also checks retention (6)")


def catalogue_morph_forms():
    """Surface forms of every photographed Yoshimoto card label.

    Lets the You Do gate tell a morpheme (`bio`, `vore`) from an everyday whole
    word (`climate`, `function`) when an operand carries no hyphen.
    """
    forms = set()
    for (_type, _key), entries in photo_card_catalogue().items():
        for entry in entries:
            forms |= morph_surface_forms(entry["morph"])
    return forms


def activity_face_lines(activity):
    """(field, physical line) for everything printed on the You Do slide face."""
    for field in ("title", "rule", "example", "footer"):
        for text in str(activity.get(field, "") or "").splitlines():
            yield field, text
    for field in ("items", "check_items"):
        for line in activity.get(field, []):
            for text in str(line or "").splitlines():
                yield field, text


def parse_sum_prompt(prompt):
    """Split `1. a + b (gloss) = answer` into its parts.

    Returns None when the prompt is not `something = something`.  Operands keep
    their hyphens so the caller can tell an affix from a root or whole word;
    printed glosses are stripped.  `answer` is the answer word when the right
    side is a single word (`microscope`, `microscope (tiny things)`,
    `invincible - vinc is inside it`), else "" - a right side that is a phrase
    is a meaning or an explanation, not an answer to test.  `right_operands`
    holds the parts of a sum written on the RIGHT (`carnivorous = carn + vore
    + -ous: eating flesh`), which is what a student supplies in a
    literal-meanings task.
    """
    body = PROMPT_LABEL_RE.sub("", prompt, count=1)
    if "=" not in body:
        return None
    left, right = body.split("=", 1)
    left = GLOSS_RE.sub("", left)
    if "+" in left:
        operands = [op.strip() for op in left.split("+")]
    else:
        operands = [left.strip()]
    right = GLOSS_RE.sub("", right)
    right = re.split(r":| - |,|=", right, maxsplit=1)[0].strip()
    right_operands = []
    if "+" in right:
        # A part is one token; anything after it on the operand is an
        # unbracketed gloss (`vore eating + -ous full of`).
        right_operands = [op.split()[0] for op in right.split("+") if op.strip()]
    words = right.split()
    answer = re.sub(r"[^A-Za-z]", "", words[0]) if len(words) == 1 else ""
    return {"operands": [op for op in operands if op],
            "is_sum": "+" in left, "answer": answer,
            "right_operands": right_operands}


def part_in_word(part, word):
    """Does `part` visibly survive inside `word`?  cooperate -> cooperative,
    private -> privacy count; small -> microscope does not."""
    stem = part.strip("-").lower()
    word = word.lower()
    forms = set(morph_surface_forms(stem))
    for cut in (1, 2):
        if len(stem) - cut >= 3:
            forms.add(stem[:len(stem) - cut])
    return any(form in word for form in forms)


def validate_activity_morphemes(day, week, session):
    """You Do shape gates from OG_MEGA_PROMPT.md sections 2f and 3d.

    Team decision (Sept 2026): meaning sums are retired and the You Do is a
    one-recall task.  Every part a student handles is either today's focus
    morpheme, a universal inflection, an everyday whole word, or PRINTED with
    its meaning on the slide face.  Taught earlier no longer counts as printed.

    The prompts students work (`example`, `items`, `check_items`) are scanned;
    the rule banner, title and footer are where glosses may live but are not
    themselves scanned for parts.  Three checks:
      1. MEANING SUM - a `+` operand or a `= ?` left side that is a meaning
         phrase (`small + look at = ?`, `bad function = ?`), or a check-slide
         sum whose operand does not survive inside the answer word
         (`small + wave = microwave`).  Fatal.
      2. ONE RECALL - an affix, catalogue root or taught morpheme the student
         supplies (a sum operand on either side of `=`, or morpheme notation
         in a prose prompt) that is not today's focus and has no
         `part (meaning)` gloss anywhere on the face.  Fatal.
      3. UNTAUGHT BUT GLOSSED - reachable, so allowed, but a NOTE asks whether
         a taught part would serve (2f first choice).
    """
    activity = session.get("new_morph_activity")
    if not isinstance(activity, dict):
        return
    focus = morph_surface_forms((session.get("new_morphology") or {}).get("morph"))
    taught = taught_morpheme_forms(week, session)
    catalogue = catalogue_morph_forms()
    # Each slide is its own face: a gloss counts only on the slide that uses it.
    faces = [list(activity_face_lines(activity))]
    for item in activity.get("round", []):
        faces.append(list(round_face_lines(activity, item)))
    reported = set()
    for face_lines in faces:
        _scan_activity_face(day, face_lines, focus, taught, catalogue, reported)


def round_face_lines(activity, item):
    """(field, physical line) printed on one round slide (v1.6)."""
    for field, text in (("title", activity.get("title", "")),
                        ("rule", item.get("rule", activity.get("rule", ""))),
                        ("footer", item.get("footer", activity.get("footer", "")))):
        for line in str(text or "").splitlines():
            yield field, line
    for line in str(item.get("prompt", "")).splitlines():
        yield "example", line
    for choice in item.get("choices", []):
        for line in str(choice).splitlines():
            yield "items", line
    if activity.get("bank"):
        yield "items", bank_line(activity)
    yield "check_items", round_answer_text(item).replace("**", "")


# v1.9 Wednesday third round (ported from the Challenge team's activity menu):
# the instruction never gives the bank size away, so it is fixed.
BANK_RULE = "A word may be used twice. Some may not be needed."
BANK_MODES = ("extra", "short", "exact")


def bank_line(activity):
    return "Word bank:  " + "   ".join(str(w) for w in activity.get("bank", []))


BUILD_ROUND_STEPS = {
    "break_apart": "BREAK whole words back into their parts.",
    "name_it": "NAME the word from its meaning on whiteboards, answer on the click.",
    "complete_sentence": "COMPLETE each sentence with a bank word on whiteboards, answer on the click.",
    "choose_meaning": "CHOOSE the best meaning on whiteboards, answer on the click.",
}


def round_format(rnd):
    """Which Wednesday round this is: break_apart / complete_sentence /
    choose_meaning / name_it (3f, v1.9)."""
    items = [i for i in rnd.get("round", []) if not i.get("worked")]
    if rnd.get("bank"):
        return "complete_sentence"
    if items and all(i.get("parts") for i in items):
        return "break_apart"
    if items and all(i.get("choices") for i in items):
        return "choose_meaning"
    return "name_it"


_SCHOOL_CARDS = None
_SCHOOL_CARDS_NOTED = set()


def school_cards():
    """label -> school card, from og_planner/school_cards.json (4e, v1.9)."""
    global _SCHOOL_CARDS
    if _SCHOOL_CARDS is None:
        _SCHOOL_CARDS = {}
        path = ROOT_DIR / "school_cards.json"
        if path.exists():
            for card in json.loads(path.read_text(encoding="utf-8")).get("cards", []):
                for label in [card["morph"], *card.get("catalogue_cards", [])]:
                    _SCHOOL_CARDS.setdefault(label.split("|")[0], card)
    return _SCHOOL_CARDS


def audit_school_cards(day, session):
    """NOTE (never WARN) where the 3/4 team's taught wording differs from the
    photographed card we teach from (4e, v1.9).  Our wording stays on the slides."""
    shown = [(c["morph"], c.get("type")) for c in session.get("morphology_review", [])]
    if session.get("new_morphology"):
        nm = session["new_morphology"]
        shown.append((nm["morph"], nm.get("type")))
    shown += [(c["morph"], "suffix") for c in (session.get("pos_drill") or {}).get("cards", [])]
    for morph, mtype in shown:
        school = school_cards().get(morph)
        if not school or morph in _SCHOOL_CARDS_NOTED:
            continue
        ours = resolved_photo_card(morph, mtype) or {}
        if (ours.get("keyword"), ours.get("meaning")) == (school["keyword"], school["meaning"]):
            continue
        _SCHOOL_CARDS_NOTED.add(morph)
        note(f"{day}: {morph} - the 3/4 team teaches it as {school['morph']} "
             f"({school['keyword']}) = {school['meaning']}; our card says "
             f"{ours.get('keyword')} = {ours.get('meaning')}. Slides keep our wording; "
             "check with Kate before the classes merge (4e, v1.9)")


BROAD_CHECK_LABELS = {"meaning", "grammar", "rule", "punctuation", "learned words",
                      "morphology", "spelling", "parts of speech", "suffixes"}


def check_friday(day, week, review, report):
    """The Friday Review gate (1a, v1.9): two student pages, an answers page, a
    teacher check box naming the exact skill on every question, a one-page
    follow-up sheet whose labels match the boxes, and this week's content only."""
    for p in report.get("problems", []):
        warn(f"{day}: Friday Review - {p} (1a)")
    if report["student_pages"] > 2:
        warn(f"{day}: Friday Review has {report['student_pages']} student pages - two at "
             "most; cut an item or the Challenge, never the writing space (1a)")
    if report["answer_pages"] < 1:
        warn(f"{day}: Friday Review has no answers page (1a)")
    all_checks = set()
    for block in report["blocks"]:
        all_checks |= {str(c) for c in block["checks"]}
        where = f"{block['part']} {block['heading']!r}"
        if not block["checks"]:
            warn(f"{day}: Friday Review {where} has no teacher check box (1a)")
        if len(block["checks"]) > 2:
            warn(f"{day}: Friday Review {where} has {len(block['checks'])} check boxes - "
                 "at most 2; split it or cut a skill (1a)")
        for label in block["checks"]:
            if str(label).strip().casefold() in BROAD_CHECK_LABELS:
                warn(f"{day}: Friday Review check box {label!r} names a broad category - "
                     "name the exact morpheme, rule, grammar point or word (1a)")
    for key in ("part_a", "part_b", "part_c", "part_d"):
        part = review.get(key) or {}
        if not part.get("subblocks"):
            warn(f"{day}: Friday Review {key} has no questions (1a)")
        for block in part.get("subblocks", []):
            if not block.get("answers"):
                warn(f"{day}: Friday Review {key} question {block.get('heading')!r} has "
                     "no answers for the answers page (1a)")
    focus = week_focus_labels(week)
    for block in (review.get("part_d") or {}).get("subblocks", []):
        head = str(block.get("heading", "")).split("=")[0].strip()
        if head and head not in focus:
            warn(f"{day}: Friday Review Part D heading {head!r} is not this week's "
                 f"morpheme ({', '.join(focus)}) - Part D tests this week only (1a)")
    forms = set()
    for label in focus:
        forms |= morph_surface_forms(label)
    dictation = review.get("dictation") or []
    if len(dictation) != 2:
        warn(f"{day}: Friday Review Part E needs exactly 2 dictation sentences (1a)")
    for n, dt in enumerate(dictation, 1):
        if not str(dt.get("sentence", "")).strip():
            warn(f"{day}: Friday Review dictation {n} has no sentence (1a)")
        if not any(f in str(t).lower() for t in dt.get("targets", []) for f in forms):
            warn(f"{day}: Friday Review dictation {n} has no target built on this week's "
                 "morpheme (1a)")
        missing = [t for t in dt.get("targets", []) if t not in dt.get("sentence", "")]
        if missing:
            warn(f"{day}: Friday Review dictation {n} targets {missing} are not in the "
                 "sentence (1a)")
    follow = review.get("follow_up") or []
    if not follow:
        warn(f"{day}: Friday Review needs follow_up rows for the Teacher Follow-up "
             "sheet (1a)")
    for row in follow:
        for key in ("question", "checks", "flag_when", "next_move"):
            if not row.get(key):
                warn(f"{day}: follow_up row {row.get('question')!r} is missing {key} (1a)")
        labels = row.get("checks", [])
        for label in (labels if isinstance(labels, list) else [labels]):
            if (str(label) not in all_checks
                    and not str(row.get("question", "")).startswith("Part E")):
                warn(f"{day}: follow_up label {label!r} matches no worksheet check "
                     "box (1a)")
    covered = {str(r.get("question", "")) for r in follow}
    rows = [q_ for q_ in covered if not q_.startswith("Part E")]
    if len(rows) < len(report["blocks"]):
        warn(f"{day}: the follow-up sheet has {len(rows)} question rows for "
             f"{len(report['blocks'])} worksheet questions - one row per question, "
             "Challenge included (1a)")
    if not any(q_.startswith("Part E") for q_ in covered):
        warn(f"{day}: follow_up needs a Part E (dictation, COPS) row (1a)")
    if report.get("follow_up_pages", 1) > 1:
        warn(f"{day}: Teacher Follow-up sheet runs to {report['follow_up_pages']} pages "
             "- one page (1a)")


def week_third_round(build):
    """The rotating round's format ('name_it' / 'complete_sentence' /
    'choose_meaning'), or None."""
    fmts = [round_format(r) for r in (build or {}).get("rounds", [])]
    return next((f for f in fmts if f != "break_apart"), None)


def earlier_week_specs(week, back=4):
    """Up to `back` earlier weeks' specs from SPEC_DIR, most recent first."""
    try:
        term, num = int(week.get("term")), int(week.get("week"))
    except (TypeError, ValueError):
        return []
    found = []
    for k in range(1, back + 1):
        path = SPEC_DIR / f"term{term}_week{num - k}.json"
        if path.exists():
            try:
                found.append(json.loads(path.read_text(encoding="utf-8")))
            except (OSError, ValueError):
                pass
    return found


def validate_v19_rounds(day, week, session):
    """v1.9 Wednesday: the third round rotates name_it / complete_sentence /
    choose_meaning (3f).  Mechanical checks only; the distractor quality bar
    lives in the prompt.  A 3-session week's `combined` day carries only the
    rotating round (section 11), so the two-round shape is not required there."""
    bd = session.get("build") or {}
    rounds = bd.get("rounds", [])
    formats = [round_format(r) for r in rounds]
    if session.get("type") != "combined" and (
            formats[:1] != ["break_apart"] or len(rounds) != 2):
        warn(f"{day}: a v1.9 build day has two rounds - Break it apart, then ONE of "
             "Write the word that names... / Complete the sentence / Choose the best "
             "meaning (3f, v1.9)")
    for rnd, fmt in zip(rounds, formats):
        items = [i for i in rnd.get("round", []) if not i.get("worked")]
        if fmt == "choose_meaning":
            for n, it in enumerate(items, 1):
                ch = [str(c) for c in it.get("choices", [])]
                if len(ch) != 3:
                    warn(f"{day}: Choose the best meaning item {n} needs exactly 3 "
                         "choices (fingers 1, 2 or 3) (3f, v1.9)")
                if str(it.get("answer", "")) not in ch:
                    warn(f"{day}: Choose the best meaning item {n}: the answer must be "
                         "one of its choices, word for word (3f, v1.9)")
        if fmt == "complete_sentence":
            bank = [str(w).lower() for w in rnd.get("bank", [])]
            mode = rnd.get("bank_mode")
            if len(bank) != len(set(bank)):
                warn(f"{day}: the word bank lists a word twice - each bank word appears "
                     "once (3f, v1.9)")
            # the "One done for you" answer is a bank word too, so it counts as used
            answers = [str(i.get("answer", "")).lower() for i in rnd.get("round", [])]
            if mode not in BANK_MODES:
                warn(f"{day}: Complete the sentence needs bank_mode extra, short or "
                     "exact (3f, v1.9)")
            if rnd.get("rule") and rnd["rule"] != BANK_RULE:
                warn(f"{day}: Complete the sentence keeps the fixed instruction line "
                     f"\"{BANK_RULE}\" so it never gives the bank size away (3f, v1.9)")
            for a in answers:
                if a not in bank:
                    warn(f"{day}: Complete the sentence answer '{a}' is not in the word "
                         "bank (3f, v1.9)")
            for n, it in enumerate(items, 1):
                if "___" not in str(it.get("prompt", "")):
                    warn(f"{day}: Complete the sentence item {n} needs a ___ gap (3f, v1.9)")
            repeats = len(answers) - len(set(answers))
            spare = len(set(bank) - set(answers))
            if mode == "exact" and (repeats or spare):
                warn(f"{day}: bank_mode exact means every bank word is used once - "
                     f"{repeats} repeated, {spare} unused (3f, v1.9)")
            if mode == "short" and (repeats != 1 or spare):
                warn(f"{day}: bank_mode short means exactly ONE word is used twice and "
                     f"none left over - {repeats} repeated, {spare} unused (3f, v1.9)")
            if mode == "extra" and (repeats or spare not in (1, 2)):
                warn(f"{day}: bank_mode extra means one or two spare words and no "
                     f"repeats - {repeats} repeated, {spare} unused (3f, v1.9)")
    third = week_third_round(bd)
    earlier = [w for w in earlier_week_specs(week) if week_is_v19(w)]
    builds = [next((s_.get("build") for s_ in w.get("sessions", []) if s_.get("build")),
                   None) or {} for w in earlier]
    prev_third = week_third_round(builds[0]) if builds else None
    if third and prev_third == third:
        note(f"{day}: the third round is {third} again - rotate it each week (3f, v1.9)")
    if third and third != "name_it" and prev_third and prev_third != "name_it":
        note(f"{day}: Write the word that names... has missed two weeks running - "
             "bring Kate's meaning-to-word recall back this week (3f, v1.9)")
    # bank_mode varies from one USE to the next, however many weeks apart
    mode = next((r.get("bank_mode") for r in rounds if r.get("bank")), None)
    last_mode = next((r.get("bank_mode") for b_ in builds for r in b_.get("rounds", [])
                      if r.get("bank")), None)
    if mode and mode == last_mode:
        note(f"{day}: bank_mode {mode} again - vary it from one use to the next "
             "(3f, v1.9)")


def _scan_activity_face(day, face_lines, focus, taught, catalogue, reported):
    face_text = "\n".join(text for _field, text in face_lines)

    def glossed(token):
        pattern = r"(?<![A-Za-z])" + re.escape(token) + r"\s*\("
        return re.search(pattern, face_text) is not None

    def check_part(token, hyphenated):
        key = token.strip("-").lower()
        if not key:
            return
        forms = morph_surface_forms(key)
        if key in UNIVERSAL_INFLECTIONS or forms & focus:
            return
        if glossed(token):
            if not forms & taught and ("note", key) not in reported:
                reported.add(("note", key))
                note(f"{day}: You Do prints the meaning of {token!r}, which is not "
                     "in this session's taught set - reachable, but a part the "
                     "class has been taught is the first choice (2f)")
            return
        if (hyphenated or forms & taught or forms & catalogue) and \
                ("warn", key) not in reported:
            reported.add(("warn", key))
            warn(f"{day}: You Do uses morpheme {token!r}, which is not today's "
                 "focus morpheme and has no printed meaning on the slide face. "
                 "Students supply ONE part from memory - today's - so print it "
                 f"as \"{token} (meaning)\" in the rule banner, the worked "
                 "example or the item, or rebuild the item from today's "
                 "morpheme and everyday words (2f/3d one-recall rule)")

    def check_operands(operands):
        for op in operands:
            check_part(op, op.startswith("-") or op.endswith("-"))

    for field, text in face_lines:
        if field not in ("example", "items", "check_items"):
            continue
        # 10c packs two prompts per line, separated by a run of spaces.
        for prompt in (p for p in re.split(r"\s{3,}", text) if p.strip()):
            parsed = parse_sum_prompt(prompt)
            if parsed is None:
                for match in AFFIX_TOKEN_RE.finditer(prompt):
                    check_part(match.group(1), True)
                continue
            phrase = [op for op in parsed["operands"] if re.search(r"\s", op)]
            if phrase:
                warn(f"{day}: You Do prompt {prompt.strip()!r} is a meaning "
                     f"sum - {phrase[0]!r} is a meaning the student must turn "
                     "into a word. Meaning sums are retired (3d). Print the "
                     "parts themselves (`micro + scope = ?`) or ask for the "
                     "sum beside the word (`microscope = ?`)")
                continue
            if parsed["is_sum"]:
                answer = parsed["answer"]
                stray = [op for op in parsed["operands"]
                         if len(answer) >= 3 and not part_in_word(op, answer)]
                if stray:
                    warn(f"{day}: You Do prompt {prompt.strip()!r} is a meaning "
                         f"sum - {stray[0]!r} does not appear inside "
                         f"{answer!r}, so it is a meaning, not a part. Meaning "
                         "sums are retired (3d)")
                    continue
                check_operands(parsed["operands"])
            else:
                # `word = ?` / `word = answer`: the student's work is any sum
                # written on the right of the check item.
                check_operands(parsed["right_operands"])


# ---------------------------------------------------------------- text helpers
def fit_font(text, width_in, base_pt, min_pt=20, char_em=0.62):
    """Largest size <= base_pt at which `text` fits on one line in width_in."""
    n = max(1, len(text))
    fit = int(width_in * 72 / (char_em * n))
    return max(min_pt, min(base_pt, fit))


def fit_font_block(text, width_in, height_in, base_pt, min_pt=20, char_em=0.55,
                   line_factor=1.32):
    """Largest size <= base_pt at which wrapped `text` fits in the box."""
    for sz in range(base_pt, min_pt - 1, -2):
        cpl = max(1, int(width_in * 72 / (char_em * sz)))
        lines = math.ceil(len(text) / cpl)
        if lines * sz * line_factor / 72 <= height_in:
            return sz
    return min_pt


def find_sp(root, spid):
    for sp in root.iter(q("p:sp")):
        nv = sp.find(q("p:nvSpPr") + "/" + q("p:cNvPr"))
        if nv is not None and nv.get("id") == str(spid):
            return sp
    raise KeyError(f"shape id {spid} not found")


def txbody(sp):
    tb = sp.find(q("p:txBody"))
    if tb is None:
        raise KeyError("shape has no txBody")
    return tb


def set_runs(sp, paragraphs, size_pt=None):
    """Replace a shape's paragraphs.

    paragraphs: list of paragraphs; each paragraph is a list of runs;
    each run is (text, color_hex_or_None, bold_or_None, underline_or_None).
    The first existing run's rPr is cloned as base formatting for every run;
    the first existing paragraph's pPr is kept for all paragraphs.
    """
    tb = txbody(sp)
    paras = tb.findall(q("a:p"))
    if not paras:
        raise KeyError("no paragraphs in shape")
    proto_p = paras[0]
    proto_ppr = proto_p.find(q("a:pPr"))
    proto_rpr = None
    for r_el in proto_p.findall(q("a:r")):
        proto_rpr = r_el.find(q("a:rPr"))
        if proto_rpr is not None:
            break
    if proto_rpr is None:  # look in later paragraphs
        for p_el in paras[1:]:
            for r_el in p_el.findall(q("a:r")):
                proto_rpr = r_el.find(q("a:rPr"))
                break
            if proto_rpr is not None:
                break
    for p_el in paras:
        tb.remove(p_el)
    for runs in paragraphs:
        p_el = etree.SubElement(tb, q("a:p"))
        if proto_ppr is not None:
            p_el.append(copy.deepcopy(proto_ppr))
        for text, color, bold, underline in runs:
            r_el = etree.SubElement(p_el, q("a:r"))
            if proto_rpr is not None:
                rpr = copy.deepcopy(proto_rpr)
            else:
                rpr = etree.Element(q("a:rPr"), lang="en-US")
            if size_pt is not None:
                rpr.set("sz", str(int(size_pt * 100)))
            if bold is not None:
                rpr.set("b", "1" if bold else "0")
            if underline:
                rpr.set("u", "sng")
            if color is not None:
                for fill in rpr.findall(q("a:solidFill")):
                    rpr.remove(fill)
                fill = etree.Element(q("a:solidFill"))
                clr = etree.SubElement(fill, q("a:srgbClr"))
                clr.set("val", color)
                fill.insert(0, clr)
                rpr.insert(0, fill)
            r_el.append(rpr)
            t_el = etree.SubElement(r_el, q("a:t"))
            t_el.text = text
        if not runs:
            if proto_rpr is not None:
                end = copy.deepcopy(proto_rpr)
                end.tag = q("a:endParaRPr")
                p_el.append(end)


def set_text(sp, text, size_pt=None, color=None):
    set_runs(sp, [[(text, color, None, None)]], size_pt=size_pt)


def replace_run_text(sp, old, new):
    """Replace `old` with `new` inside existing runs, keeping formatting."""
    hit = False
    for t_el in sp.iter(q("a:t")):
        if t_el.text and old in t_el.text:
            t_el.text = t_el.text.replace(old, new)
            hit = True
    return hit


def set_solid_fill(sp, hexval):
    sppr = sp.find(q("p:spPr"))
    for fill in sppr.findall(q("a:solidFill")):
        sppr.remove(fill)
    fill = etree.Element(q("a:solidFill"))
    clr = etree.SubElement(fill, q("a:srgbClr"))
    clr.set("val", hexval)
    ln = sppr.find(q("a:ln"))
    if ln is not None:
        ln.addprevious(fill)
    else:
        sppr.append(fill)


def set_slide_background(root, hexval):
    """Set a cloned slide's solid background without changing its type icon."""
    csld = root.find(q("p:cSld"))
    if csld is None:
        raise KeyError("slide common data missing")
    bg = csld.find(q("p:bg"))
    if bg is None:
        bg = etree.Element(q("p:bg"))
        csld.insert(0, bg)
    for child in list(bg):
        bg.remove(child)
    bgpr = etree.SubElement(bg, q("p:bgPr"))
    solid = etree.SubElement(bgpr, q("a:solidFill"))
    etree.SubElement(solid, q("a:srgbClr"), val=hexval)


def add_textbox(root, sid, x, y, w, h, lines, size_pt, color=CHARCOAL, bold=False,
                fill=None, rounded=False, align="ctr", anchor="ctr", italic=False,
                outline=None):
    """Add a generated text box/card to a slide. `**bold**` markup supported."""
    a, p = NS["a"], NS["p"]
    geom = "roundRect" if rounded else "rect"
    fill_xml = (f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
                if fill else "<a:noFill/>")
    line_xml = (f'<a:ln w="25400"><a:solidFill><a:srgbClr val="{outline}"/>'
                f'</a:solidFill></a:ln>' if outline else '<a:ln><a:noFill/></a:ln>')
    paras = []
    for ln in lines:
        runs = ""
        for i, seg in enumerate(re.split(r"\*\*(.*?)\*\*", ln)):
            if seg == "":
                continue
            b = "1" if (bold or i % 2 == 1) else "0"
            ital = ' i="1"' if italic else ""
            runs += (f'<a:r><a:rPr lang="en-AU" sz="{int(size_pt * 100)}" b="{b}"{ital}>'
                     f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                     f'<a:latin typeface="Lexend"/></a:rPr><a:t>{escape(seg)}</a:t></a:r>')
        if not runs:
            runs = f'<a:endParaRPr lang="en-AU" sz="{int(size_pt * 100)}"/>'
        paras.append(f'<a:p><a:pPr algn="{align}"/>{runs}</a:p>')
    xml = (
        f'<p:sp xmlns:p="{p}" xmlns:a="{a}">'
        f'<p:nvSpPr><p:cNvPr id="{sid}" name="OGGen{sid}"/><p:cNvSpPr txBox="1"/>'
        f'<p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(x * 914400)}" y="{int(y * 914400)}"/>'
        f'<a:ext cx="{int(w * 914400)}" cy="{int(h * 914400)}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>{fill_xml}'
        f'{line_xml}</p:spPr>'
        f'<p:txBody><a:bodyPr anchor="{anchor}" wrap="square" lIns="91440" '
        f'tIns="45720" rIns="91440" bIns="45720"/><a:lstStyle/>{"".join(paras)}'
        f'</p:txBody></p:sp>')
    sptree = root.find(q("p:cSld") + "/" + q("p:spTree"))
    sptree.append(etree.fromstring(xml))


def table_cells(root, spid):
    for gf in root.iter(q("p:graphicFrame")):
        nv = gf.find(q("p:nvGraphicFramePr") + "/" + q("p:cNvPr"))
        if nv is not None and nv.get("id") == str(spid):
            return gf.findall(".//" + q("a:tr"))
    raise KeyError(f"table {spid} not found")


def set_cell_paras(tc, lines, size_pt=None):
    """Replace a table cell's paragraphs with `lines` (list of strings)."""
    tx = tc.find(q("a:txBody"))
    paras = tx.findall(q("a:p"))
    proto_p = paras[0]
    proto_ppr = proto_p.find(q("a:pPr"))
    proto_rpr = None
    for p_el in paras:
        for r_el in p_el.findall(q("a:r")):
            proto_rpr = r_el.find(q("a:rPr"))
            break
        if proto_rpr is not None:
            break
    for p_el in paras:
        tx.remove(p_el)
    for line in lines:
        p_el = etree.SubElement(tx, q("a:p"))
        if proto_ppr is not None:
            p_el.append(copy.deepcopy(proto_ppr))
        r_el = etree.SubElement(p_el, q("a:r"))
        rpr = copy.deepcopy(proto_rpr) if proto_rpr is not None else etree.Element(q("a:rPr"), lang="en-US")
        if size_pt is not None:
            rpr.set("sz", str(int(size_pt * 100)))
        r_el.append(rpr)
        t_el = etree.SubElement(r_el, q("a:t"))
        t_el.text = line


# ---------------------------------------------------------------- timing regen
def regen_timing(root, targets):
    """Rebuild the click-to-reveal sequence.

    targets: list of (spid, para_index_or_None). One click per entry, in order.
    Uses the slide's existing timing block as the structural pattern.
    """
    timing = root.find(q("p:timing"))
    if timing is None:
        raise KeyError("slide has no timing block to use as pattern")
    seq_ctn = timing.find(".//" + q("p:seq") + "/" + q("p:cTn"))
    child_lst = seq_ctn.find(q("p:childTnLst"))
    pars = child_lst.findall(q("p:par"))
    pattern = copy.deepcopy(pars[0])
    for p_el in pars:
        child_lst.remove(p_el)
    ids = [2000]

    def renumber(el):
        for ctn in el.iter(q("p:cTn")):
            if ctn.get("id"):
                ids[0] += 1
                ctn.set("id", str(ids[0]))

    for spid, para in targets:
        node = copy.deepcopy(pattern)
        renumber(node)
        tgt = node.find(".//" + q("p:spTgt"))
        tgt.set("spid", str(spid))
        txel = tgt.find(q("p:txEl"))
        if para is None:
            if txel is not None:
                tgt.remove(txel)
        else:
            if txel is None:
                txel = etree.SubElement(tgt, q("p:txEl"))
                etree.SubElement(txel, q("p:pRg"))
            prg = txel.find(q("p:pRg"))
            if prg is None:
                for ch in list(txel):
                    txel.remove(ch)
                prg = etree.SubElement(txel, q("p:pRg"))
            prg.set("st", str(para))
            prg.set("end", str(para))
        child_lst.append(node)
    # bldLst: ensure one bldP per shape used, paragraph builds
    bld = timing.find(q("p:bldLst"))
    if bld is not None:
        for b in list(bld):
            bld.remove(b)
        for spid in dict.fromkeys(s for s, _ in targets):
            bp = etree.SubElement(bld, q("p:bldP"))
            bp.set("spid", str(spid))
            bp.set("grpId", "0")
            if any(p is not None for s, p in targets if s == spid):
                bp.set("build", "p")


def remove_timing(root):
    """Make an instructional follow-up slide visible immediately."""
    timing = root.find(q("p:timing"))
    if timing is not None:
        root.remove(timing)


# ---------------------------------------------------------------- notes
NOTE_ANCHOR_RE = re.compile(
    r"(?:\d+\.\s+)?(?:ANSWER|ANSWERS|SAY|ASK|EXPECT|ACCEPT|SCAN|TRAP|FIX|"
    r"EXTENSION|MEMORY HOOK|STRETCH|HELP|"
    r"CARE|TIME|CIRCULATE|CHECK|PREP|POINT|SHOW|MODEL|DRAW|BUILD|COVER|REVEAL|COLLECT|"
    r"READ|ASSIGN|CHORAL|WAIT|REPEAT|CLARIFY|WRITE|TICK|PROCEDURE|CORRECTION|"
    r"Extra task|Possible answers|Root|Prefix|Suffix|Keyword|Meaning|Sound|"
    r"Meaning of [^:]+|Part of speech|Word to spell|Sentence|After checking|"
    r"Answer|Why learned|Say it|Link|C - Capitals(?:, shown green)?|U - Understanding|O - Organisation|"
    r"P - Punctuation, shown red|S - Spelling targets, underlined|Score|Focus|"
    r"Targets|Memory hook|Purpose|WHY|SOURCES):"
)
NUMBERED_ACTION_RE = re.compile(
    r"^\d+\.\s+(?:POINT|SHOW|MODEL|DRAW|BUILD|COVER|REVEAL|TIME|COLLECT|"
    r"CIRCULATE|CHECK|SCAN|READ|WRITE|TICK|WAIT|REPEAT|CLARIFY)\b"
)


def prepare_notes_text(text):
    """Backstop old dense scripts while keeping well-authored line breaks intact."""
    text = sanitize(str(text or "")).replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\x0b", "\n")
    # Older specs sometimes placed every numbered beat and live action on one line.
    # (?<!\d) keeps the break out of the middle of a two-digit beat number ("10. ASK").
    text = re.sub(
        r"(?<!\n)(?<!^)(?<!\d)(?=\d+\.\s+(?:SAY|ASK|SCAN|MODEL|POINT|SHOW|READ|"
        r"REVEAL|CIRCULATE|TIME|WRITE|BUILD|TICK):)",
        "\n",
        text,
    )
    text = re.sub(
        r"(?<!\n)(?<!\d\. )(?=(?:TIME|CIRCULATE|CHECK|SCAN|TRAP|EXTENSION|"
        r"STRETCH|HELP|CARE|ANSWERS|MEMORY HOOK):)",
        "\n",
        text,
    )
    text = re.sub(r"\s*---\s*", "\n---\n", text)
    return "\n".join(line.rstrip() for line in text.splitlines()).strip()


def validate_notes_source(template_idx, original_text, prepared_text):
    """Fail the build on note sources that will become an iPad text wall."""
    if "**" in original_text:
        warn(f"template slide {template_idx}: notes contain markdown ** markers; "
             "use plain labels and let the builder apply bold formatting")
    if re.search(r"[^\x09\x0A\x0D\x20-\x7E]", prepared_text):
        warn(f"template slide {template_idx}: notes contain non-ASCII characters; "
             "use straight quotes, hyphens and plain arrows")
    live_lines = []
    for line_no, line in enumerate(prepared_text.splitlines(), 1):
        if line == "---":
            break
        if line.strip():
            live_lines.append(line)
        if len(line) > 220:
            warn(f"template slide {template_idx}: notes line {line_no} is "
                 f"{len(line)} characters; split each new thought onto its own line")
        if re.fullmatch(r"\d+", line.strip()):
            warn(f"template slide {template_idx}: notes line {line_no} is a bare "
                 "number - a numbered beat has been split mid-line")
        numbered_beats = re.findall(
            r"\b\d+\.\s+(?:SAY|ASK|SCAN|MODEL|POINT|SHOW|READ|REVEAL|"
            r"CIRCULATE|TIME|WRITE|BUILD|TICK|CHECK|WAIT|REPEAT|CLARIFY)\b",
            line,
        )
        if len(numbered_beats) > 1:
            warn(f"template slide {template_idx}: notes line {line_no} contains "
                 "multiple numbered beats")
        if re.match(r"^\d+\.\s+", line) and (
                line.count("EXPECT:") > 1 or line.count("?") > 1):
            warn(f"template slide {template_idx}: notes line {line_no} contains "
                 "multiple questions or answers; use one retrieval prompt per line")
    if "---" in prepared_text.splitlines() and len(live_lines) > 8:
        warn(f"template slide {template_idx}: notes live zone has "
             f"{len(live_lines)} lines; Glance Format allows at most 8")


def note_runs(line, underline_terms=()):
    """Return (text, bold, underline) triples using real PowerPoint formatting."""
    plain = line.replace("**", "")
    bold_ranges = []
    underline_ranges = []
    # Preserve legacy bold intent for old specs while emitting no markdown.
    plain_pos = 0
    cursor = 0
    bold = False
    while cursor < len(line):
        if line.startswith("**", cursor):
            bold = not bold
            cursor += 2
            continue
        start = cursor
        while cursor < len(line) and not line.startswith("**", cursor):
            cursor += 1
        length = cursor - start
        if bold and length:
            bold_ranges.append((plain_pos, plain_pos + length))
        plain_pos += length
    for match in NOTE_ANCHOR_RE.finditer(plain):
        bold_ranges.append(match.span())
    action_match = NUMBERED_ACTION_RE.search(plain)
    if action_match:
        bold_ranges.append(action_match.span())
    for term in underline_terms:
        term = str(term).strip()
        if not term:
            continue
        pattern = rf"(?<![A-Za-z]){re.escape(term)}(?![A-Za-z])"
        underline_ranges.extend(
            match.span() for match in re.finditer(pattern, plain, re.IGNORECASE)
        )
    if not plain:
        return []
    boundaries = {0, len(plain)}
    for start, end in bold_ranges + underline_ranges:
        boundaries.update((start, end))
    points = sorted(boundaries)
    runs = []
    for start, end in zip(points, points[1:]):
        if start == end:
            continue
        is_bold = any(bs <= start and end <= be for bs, be in bold_ranges)
        is_underlined = any(
            us <= start and end <= ue for us, ue in underline_ranges
        )
        runs.append((plain[start:end], is_bold, is_underlined))
    return runs


def set_notes(notes_root, text):
    """Replace the body placeholder text of a notes slide.

    Source notes are plain text. The builder bolds recognised labels as real
    PowerPoint formatting and preserves one paragraph per source line.
    """
    body = None
    for sp in notes_root.iter(q("p:sp")):
        ph = sp.find(".//" + q("p:ph"))
        if ph is not None and ph.get("type") == "body":
            body = sp
            break
    if body is None:
        raise KeyError("notes body placeholder not found")
    original_text = str(text or "")
    text = prepare_notes_text(original_text)
    spelling_match = re.search(
        r"^Word to spell:\s*(\S.*)$",
        text,
        flags=re.MULTILINE,
    )
    spelling_word = spelling_match.group(1).strip() if spelling_match else ""
    tb = txbody(body)
    body_pr = tb.find(q("a:bodyPr"))
    if body_pr is not None:
        body_pr.set("anchor", "t")
    for p_el in tb.findall(q("a:p")):
        tb.remove(p_el)
    for line in text.split("\n"):
        line = re.sub(r" {2,}", " ", line.rstrip())
        p_el = etree.SubElement(tb, q("a:p"))
        # kill the notes master's default bullets - they read as dot-point
        # walls on the iPad speaker-notes view
        ppr = etree.SubElement(p_el, q("a:pPr"))
        ppr.set("algn", "l")
        etree.SubElement(ppr, q("a:buNone"))
        underline_terms = (
            (spelling_word,)
            if spelling_word and line.startswith("Sentence:")
            else ()
        )
        runs = note_runs(line, underline_terms=underline_terms)
        for seg, bold, underlined in runs:
            r_el = etree.SubElement(p_el, q("a:r"))
            rpr = etree.SubElement(r_el, q("a:rPr"), lang="en-AU", sz="1100")
            if bold:
                rpr.set("b", "1")
            if underlined:
                rpr.set("u", "sng")
            t_el = etree.SubElement(r_el, q("a:t"))
            t_el.text = seg
        if not runs:
            etree.SubElement(p_el, q("a:endParaRPr"), lang="en-AU", sz="1100")


def tidy_notes_whitespace(notes_root):
    """Normalise fixed template notes (double spaces read badly on the iPad)."""
    for t_el in notes_root.iter(q("a:t")):
        if t_el.text:
            t_el.text = re.sub(r" {2,}", " ", t_el.text)


# ---------------------------------------------------------------- package
class Package:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}
        if ("ppt/media/image12.png" in self.parts
                and "ppt/media/image13.png" in self.parts):
            self.parts["ppt/media/image12.png"] = make_compatible_yellow_meter(
                self.parts["ppt/media/image13.png"],
                self.parts["ppt/media/image12.png"],
            )
        self.pres = etree.fromstring(self.parts["ppt/presentation.xml"])
        self.pres_rels = etree.fromstring(self.parts["ppt/_rels/presentation.xml.rels"])
        # ordered template slide part names
        rid_to_target = {
            rel.get("Id"): rel.get("Target")
            for rel in self.pres_rels.findall(q("rel:Relationship"))
        }
        self.slide_parts = []
        for sld in self.pres.findall(".//" + q("p:sldId")):
            rid = sld.get(q("r:id"))
            self.slide_parts.append("ppt/" + rid_to_target[rid].lstrip("/").replace("../", ""))

    def slide_xml(self, idx1):
        return self.parts[self.slide_parts[idx1 - 1]]

    def slide_rels(self, idx1):
        name = self.slide_parts[idx1 - 1]
        rn = name.rsplit("/", 1)
        return self.parts.get(f"{rn[0]}/_rels/{rn[1]}.rels")

    def notes_name(self, idx1):
        rels = self.slide_rels(idx1)
        if not rels:
            return None
        root = etree.fromstring(rels)
        for rel in root.findall(q("rel:Relationship")):
            if rel.get("Type").endswith("/notesSlide"):
                tgt = rel.get("Target").replace("../", "ppt/")
                return tgt
        return None


class DeckBuilder:
    """Assembles one output PPTX from template slide clones."""

    def __init__(self, pkg):
        self.pkg = pkg
        self.slides = []  # (slide_xml_root, rels_root, notes_root_or_None)
        self.media = {}   # extra parts, e.g. "ppt/media/oghook_port.png" -> bytes

    def add(self, template_idx, fill=None, notes=None, fill_rels=None):
        """Clone template slide `template_idx`; apply fill(root) mutations."""
        root = etree.fromstring(self.pkg.slide_xml(template_idx))
        rels_bytes = self.pkg.slide_rels(template_idx)
        rels = etree.fromstring(rels_bytes) if rels_bytes else None
        notes_name = self.pkg.notes_name(template_idx)
        notes_root = None
        if notes is not None and notes_name is None:
            # borrow a notes slide from a template slide that has one
            notes_name = self.pkg.notes_name(T["spell_word"])
        if notes_name:
            notes_root = etree.fromstring(self.pkg.parts[notes_name])
            notes_rels = etree.fromstring(
                self.pkg.parts[notes_name.replace("notesSlides/", "notesSlides/_rels/") + ".rels"]
            )
        else:
            notes_rels = None
        if fill:
            fill(root)
        if fill_rels:
            fill_rels(rels)
        if notes is not None and notes_root is not None:
            prepared_notes = prepare_notes_text(str(notes))
            validate_notes_source(template_idx, str(notes), prepared_notes)
            set_notes(notes_root, notes)
        elif notes_root is not None:
            tidy_notes_whitespace(notes_root)
        self.slides.append((root, rels, notes_root, notes_rels))

    def write(self, out_path):
        parts = {}
        # copy every part except slides/notesSlides and their rels and ctypes/pres
        skip_prefixes = ("ppt/slides/", "ppt/notesSlides/")
        for name, data in self.pkg.parts.items():
            if name.startswith(skip_prefixes):
                continue
            if name in ("[Content_Types].xml", "ppt/presentation.xml",
                        "ppt/_rels/presentation.xml.rels"):
                continue
            parts[name] = data
        parts.update(self.media)

        # slides + notes
        for i, (root, rels, notes_root, notes_rels) in enumerate(self.slides, 1):
            sname = f"ppt/slides/slide{i}.xml"
            parts[sname] = etree.tostring(root, xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
            if rels is not None:
                rels2 = copy.deepcopy(rels)
                for rel in rels2.findall(q("rel:Relationship")):
                    if rel.get("Type").endswith("/notesSlide"):
                        if notes_root is not None:
                            rel.set("Target", f"../notesSlides/notesSlide{i}.xml")
                        else:
                            rels2.remove(rel)
                parts[f"ppt/slides/_rels/slide{i}.xml.rels"] = etree.tostring(
                    rels2, xml_declaration=True, encoding="UTF-8", standalone=True)
            if notes_root is not None:
                nname = f"ppt/notesSlides/notesSlide{i}.xml"
                parts[nname] = etree.tostring(notes_root, xml_declaration=True,
                                              encoding="UTF-8", standalone=True)
                nrels = copy.deepcopy(notes_rels)
                for rel in nrels.findall(q("rel:Relationship")):
                    if rel.get("Type").endswith("/slide"):
                        rel.set("Target", f"../slides/slide{i}.xml")
                parts[f"ppt/notesSlides/_rels/notesSlide{i}.xml.rels"] = etree.tostring(
                    nrels, xml_declaration=True, encoding="UTF-8", standalone=True)

        # presentation.xml
        pres = copy.deepcopy(self.pkg.pres)
        sldlst = pres.find(q("p:sldIdLst"))
        for sld in list(sldlst):
            sldlst.remove(sld)
        for i in range(1, len(self.slides) + 1):
            sld = etree.SubElement(sldlst, q("p:sldId"))
            sld.set("id", str(255 + i))
            sld.set(q("r:id"), f"rIdSl{i}")
        parts["ppt/presentation.xml"] = etree.tostring(
            pres, xml_declaration=True, encoding="UTF-8", standalone=True)

        # presentation rels
        prels = copy.deepcopy(self.pkg.pres_rels)
        for rel in prels.findall(q("rel:Relationship")):
            if rel.get("Type").endswith("/slide"):
                prels.remove(rel)
        for i in range(1, len(self.slides) + 1):
            rel = etree.SubElement(prels, q("rel:Relationship"))
            rel.set("Id", f"rIdSl{i}")
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
            rel.set("Target", f"slides/slide{i}.xml")
        parts["ppt/_rels/presentation.xml.rels"] = etree.tostring(
            prels, xml_declaration=True, encoding="UTF-8", standalone=True)

        # content types
        ct = etree.fromstring(self.pkg.parts["[Content_Types].xml"])
        for ov in ct.findall(q("ct:Override")):
            pn = ov.get("PartName")
            if pn.startswith("/ppt/slides/") or pn.startswith("/ppt/notesSlides/"):
                ct.remove(ov)
        for i in range(1, len(self.slides) + 1):
            ov = etree.SubElement(ct, q("ct:Override"))
            ov.set("PartName", f"/ppt/slides/slide{i}.xml")
            ov.set("ContentType",
                   "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
            if self.slides[i - 1][2] is not None:
                ov = etree.SubElement(ct, q("ct:Override"))
                ov.set("PartName", f"/ppt/notesSlides/notesSlide{i}.xml")
                ov.set("ContentType",
                       "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml")
        parts["[Content_Types].xml"] = etree.tostring(
            ct, xml_declaration=True, encoding="UTF-8", standalone=True)

        out_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in sorted(parts.items()):
                z.writestr(name, data)


_PICTOGRAM_CACHE = {}


def pictogram_png(name, color=MORPH_MAGENTA, px=384):
    """PNG bytes for a pictogram from the shared theme library (217 icons,
    themes/core/pictograms.js).  None when the name is unknown."""
    key = (name, color, px)
    if key not in _PICTOGRAM_CACHE:
        repo = ROOT_DIR.parent
        js = ("const p=require('./themes/core/pictograms');"
              f"const r=p.renderPictogramPng({json.dumps(name)},{json.dumps(color)},{px});"
              "process.stdout.write(r?r.split(',')[1]:'')")
        out = subprocess.run(["node", "-e", js], cwd=repo, capture_output=True,
                             text=True)
        _PICTOGRAM_CACHE[key] = base64.b64decode(out.stdout) if out.stdout else None
    return _PICTOGRAM_CACHE[key]


ICON_PACK_RE = re.compile(r"^(ai|bi|bs|ci|fa|fa6|fi|gi|go|gr|hi|hi2|im|io|io5|lu|md|pi|ri|"
                          r"si|sl|tb|tfi|vsc|wi)/[A-Z][A-Za-z0-9]+$")
MEANING_PICTURE_INK = "1F3A5F"


def meaning_picture_png(picture, px=512):
    """PNG bytes for a Words to Spell Review meaning picture (2d, v1.9).

    `picture` is a pictogram name ("plane"), a react-icons "pack/Component" name
    ("fa6/FaEarthOceania" - thousands of concrete objects the 217-icon pictogram
    set lacks), {"icon": either, "color": "RRGGBB"}, or a LIST of two or three of
    those drawn side by side, so a picture can show an idea one icon cannot (bear,
    berries, fish = an omnivore).  None when any part is unknown.
    """
    if isinstance(picture, list):
        parts = [meaning_picture_png(p_, px) for p_ in picture]
        if not parts or any(p_ is None for p_ in parts):
            return None
        ims = [Image.open(BytesIO(b)).convert("RGBA") for b in parts]
        ims = [im.resize((max(1, round(im.width * px / im.height)), px)) for im in ims]
        gap = px // 5
        sheet = Image.new("RGBA", (sum(im.width for im in ims) + gap * (len(ims) - 1), px),
                          (0, 0, 0, 0))
        x = 0
        for im in ims:
            sheet.paste(im, (x, 0), im)
            x += im.width + gap
        out = BytesIO()
        sheet.save(out, format="PNG")
        return out.getvalue()
    if isinstance(picture, dict) and picture.get("highlight"):
        # one region of the icon in a strong colour, the rest pale: a territory
        # on the map of Australia.  box = [x0, y0, x1, y1] as fractions.
        hl = picture["highlight"]
        base = meaning_picture_png({"icon": picture.get("icon", ""),
                                    "color": picture.get("color", "D5D0DA")}, px)
        strong = meaning_picture_png({"icon": picture.get("icon", ""),
                                      "color": hl.get("color", "B5541C")}, px)
        if base is None or strong is None:
            return None
        im = Image.open(BytesIO(base)).convert("RGBA")
        top = Image.open(BytesIO(strong)).convert("RGBA")
        x0, y0, x1, y1 = (round(f * d) for f, d in zip(hl["box"], (im.width, im.height) * 2))
        region = top.crop((x0, y0, x1, y1))
        im.paste(region, (x0, y0), region)
        out = BytesIO()
        im.save(out, format="PNG")
        return out.getvalue()
    if isinstance(picture, dict):
        name, color = picture.get("icon", ""), picture.get("color", MEANING_PICTURE_INK)
    else:
        name, color = str(picture or ""), MEANING_PICTURE_INK
    color = color.lstrip("#").upper()
    if "/" not in name:
        return pictogram_png(name, color, px) if name else None
    if not ICON_PACK_RE.match(name):
        return None
    key = (name, color, px)
    if key not in _PICTOGRAM_CACHE:
        pack, comp = name.split("/")
        js = ("const React=require('react');const S=require('react-dom/server');"
              "const {Resvg}=require('@resvg/resvg-js');let C=null;"
              f"try{{C=require('react-icons/{pack}')[{json.dumps(comp)}]}}catch(e){{}}"
              "if(C){const svg=S.renderToStaticMarkup(React.createElement(C,"
              f"{{color:'#{color}',size:'{px}'}}));"
              f"const png=new Resvg(svg,{{fitTo:{{mode:'width',value:{px}}},"
              "font:{loadSystemFonts:false}}).render().asPng();"
              "process.stdout.write(Buffer.from(png).toString('base64'))}")
        out = subprocess.run(["node", "-e", js], cwd=ROOT_DIR.parent,
                             capture_output=True, text=True)
        _PICTOGRAM_CACHE[key] = base64.b64decode(out.stdout) if out.stdout else None
    return _PICTOGRAM_CACHE[key]


def add_with_effect(root, spid):
    """Make shape `spid` appear WITH the slide's last click (Appear, With Previous).

    The Words to Spell Review picture rides the word's click, so the word and
    its meaning picture land together (2d, v1.9).
    """
    timing = root.find(q("p:timing"))
    clicks = [c for c in timing.iter(q("p:cTn")) if c.get("nodeType") == "clickEffect"]
    click_par = clicks[-1].getparent()
    with_par = copy.deepcopy(click_par)
    with_par.find(q("p:cTn")).set("nodeType", "withEffect")
    with_par.find(".//" + q("p:spTgt")).set("spid", str(spid))
    for ctn in with_par.iter(q("p:cTn")):
        if ctn.get("id"):
            ctn.set("id", str(int(ctn.get("id")) + 500))
    click_par.addnext(with_par)


def add_picture(root, sid, rid, x, y, w, h):
    """Place an image (already related as `rid`) on the slide."""
    a, p, r = NS["a"], NS["p"], NS["r"]
    xml = (
        f'<p:pic xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{r}">'
        f'<p:nvPicPr><p:cNvPr id="{sid}" name="OGPic{sid}"/><p:cNvPicPr/><p:nvPr/>'
        f'</p:nvPicPr><p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/>'
        f'</a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{int(x * 914400)}" '
        f'y="{int(y * 914400)}"/><a:ext cx="{int(w * 914400)}" cy="{int(h * 914400)}"/>'
        f'</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )
    root.find(q("p:cSld") + "/" + q("p:spTree")).append(etree.fromstring(xml))


def relate_image(rels, rid, target):
    rel = etree.SubElement(rels, q("rel:Relationship"))
    rel.set("Id", rid)
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/"
                    "relationships/image")
    rel.set("Target", target)


# ---------------------------------------------------------------- fills
def draw_structured_slide(root, blk, accent, tint, ctx):
    """Rule banner / hero card / support lines / routine chip / footer layout.

    Shared by grammar I-We-You do slides and the new-morphology You Do activity.
    """
    sid, y = 90100, 1.15
    if blk.get("rule"):
        rsz = fit_font_block(blk["rule"], 8.6, 0.58, 18, min_pt=15)
        add_textbox(root, sid, 0.5, y, 9.0, 0.64, [blk["rule"]], rsz,
                    color="FFFFFF", bold=True, fill=accent, rounded=True)
        sid += 1
        y += 0.76
    if blk.get("example"):
        ex_h = 1.65 if blk.get("items") else 2.45
        # A newline in the hero is an authored line break (an equation belongs
        # on its own line). Emit real paragraphs - a raw \n inside one run
        # loses the first line's centring and renders visibly off-axis.
        ex_lines = [ln.strip() for ln in blk["example"].split("\n") if ln.strip()]
        longest = max(ex_lines, key=len)
        sz = fit_font_block(longest, 7.7, (ex_h - 0.25) / max(1, len(ex_lines)),
                            30, min_pt=20, char_em=0.58)
        add_textbox(root, sid, 0.9, y, 8.2, ex_h, ex_lines, sz,
                    fill=tint, rounded=True)
        sid += 1
        y += ex_h + 0.12
    item_size = int(blk.get("item_size", 15))
    item_h = 0.54 if item_size >= 20 else 0.42
    item_step = item_h + 0.04
    for it in blk.get("items", [])[:4]:
        add_textbox(root, sid, 0.9, y, 8.2, item_h, [it], item_size,
                    color=blk.get("item_color", CHARCOAL),
                    bold=bool(blk.get("items_bold")), align="ctr")
        sid += 1
        y += item_step
    if y > 4.60:
        warn(f"{ctx}: content reaches y={y:.2f} - trim items")
    if blk.get("routine"):
        csz = fit_font(blk["routine"], 3.55, 15, min_pt=13, char_em=0.58)
        add_textbox(root, sid, 0.5, 4.66, 3.9, 0.44, [blk["routine"]], csz,
                    color="FFFFFF", bold=True, fill=accent, rounded=True)
        sid += 1
    if blk.get("footer"):
        add_textbox(root, sid, 4.55, 4.66, 4.95, 0.44, [blk["footer"]], 14,
                    align="r", italic=True, color="595959")


def has_fixed_answer_notes(blk):
    """True when a task's notes open with a definite ANSWER (not 'ANSWER: open')."""
    lines = prepare_notes_text(blk.get("notes", "")).splitlines()
    return bool(
        lines
        and lines[0].upper().startswith("ANSWER:")
        and not lines[0].lower().startswith("answer: open")
    )


def add_answer_check_slide(d, blk, tidx, accent, tint, ctx,
                           default_footer="Explain which card meaning fits each word"):
    """Emit the green 'Tick it or fix it' answer slide for a fixed-answer task.

    Shared by the new-morphology You Do and the grammar You Do. Returns True
    when the block supplied check_items and a slide was added.
    """
    check_items = blk.get("check_items", [])
    if not check_items:
        return False
    check_blk = {
        "rule": blk.get(
            "check_rule",
            "Tick each correct answer. Fix any answer that does not match.",
        ),
        "items": check_items,
        "item_color": ANSWER_GREEN,
        "items_bold": True,
        "item_size": blk.get("check_item_size", 22),
        "routine": blk.get("check_routine", "Tick - fix - explain"),
        "footer": blk.get("check_footer", default_footer),
    }

    def fill_check(root, check_blk=check_blk, blk=blk, tidx=tidx):
        tsp = find_sp(root, SHAPE["gr_title"][tidx])
        check_title = blk.get(
            "check_title",
            "Tick it or fix it - " + blk.get("title", ""),
        )
        set_text(
            tsp,
            check_title,
            size_pt=fit_font(check_title, 7.22, 28, min_pt=20, char_em=0.5),
        )
        bsp = find_sp(root, SHAPE["gr_body"][tidx])
        set_runs(bsp, [[]])
        draw_structured_slide(root, check_blk, accent, tint, ctx)

    d.add(tidx, fill_check, notes=blk.get("check_notes", ""))
    return True


# v1.8 parts-of-speech choices (3g): each job in words a Year 5 can test
# against the sentence.
POS_JOBS = {
    "noun": "names a thing or idea",
    "verb": "something you do",
    "adjective": "describes a noun",
}

ROUND_BAR_BOTTOM = 4.56   # the routine chip and footer start at y=4.66


def round_answer_text(item):
    """The green reveal bar: the answer in bold, then its one-line reason."""
    text = f"**{item.get('answer', '')}**"
    if item.get("why"):
        text += " - " + item["why"]
    return text


def draw_round_slide(root, blk, item, accent, tint, ctx):
    """One question per slide (OG_MEGA_PROMPT 3d/7, v1.6).

    Rule banner / hero item / up to three choice chips / green answer bar /
    routine chip / footer.  Returns the answer bar's shape id so the caller
    can hide it behind one click - the reveal IS the tick-it-or-fix-it moment.
    """
    sid, y = 90100, 1.15
    rule = item.get("rule", blk.get("rule"))
    if blk.get("bank"):
        rule = BANK_RULE
    if rule:
        rsz = fit_font_block(rule, 8.6, 0.58, 18, min_pt=15)
        add_textbox(root, sid, 0.5, y, 9.0, 0.64, [rule], rsz,
                    color="FFFFFF", bold=True, fill=accent, rounded=True)
        sid += 1
        y += 0.76
    if blk.get("bank"):
        # v1.9 Complete the sentence: the same bank sits on every item's slide
        line = bank_line(blk)
        bsz = fit_font(line, 8.0, 20, min_pt=14, char_em=0.58)
        add_textbox(root, sid, 0.9, y, 8.2, 0.5, [f"**{line}**"], bsz,
                    fill="FFFFFF", outline=accent, rounded=True)
        sid += 1
        y += 0.58
    choices = [str(c) for c in item.get("choices", [])][:3]
    parts = [str(pt) for pt in item.get("parts", [])][:4]
    hero_h = 1.1 if (choices or parts or blk.get("bank")) else 1.62
    lines = [ln.strip() for ln in str(item.get("prompt", "")).split("\n")
             if ln.strip()] or [""]
    longest = max(lines, key=len)
    sz = fit_font_block(longest, 7.7, (hero_h - 0.2) / len(lines), 32,
                        min_pt=20, char_em=0.58)
    add_textbox(root, sid, 0.9, y, 8.2, hero_h, lines, sz,
                fill=tint, rounded=True)
    sid += 1
    y += hero_h + 0.1
    if choices:
        gap = 0.18
        cw = (8.2 - gap * (len(choices) - 1)) / len(choices)
        chip_lines = [[ln.strip() for ln in c.split("\n") if ln.strip()]
                      for c in choices]
        longest = max((ln for cl in chip_lines for ln in cl), key=len)
        rows = max(len(cl) for cl in chip_lines)
        # Authored **bold** in a chip (a label over its plain gloss) means
        # the author controls the weight; plain chips stay all bold.
        marked = any("**" in c for c in choices)
        csz = fit_font_block(longest.replace("**", ""), cw - 0.2,
                             (0.82 - 0.1) / rows, 20, min_pt=13,
                             char_em=0.5 if marked else 0.58)
        for i, cl in enumerate(chip_lines):
            add_textbox(root, sid, 0.9 + i * (cw + gap), y, cw, 0.82, cl, csz,
                        bold=not marked, fill="FFFFFF", outline=accent, rounded=True)
            sid += 1
        y += 0.82 + 0.08
    elif parts:
        # v1.8 Break it apart (James Hooke, 23 Oct 2026): one box per part with
        # the week's root already in place, so students write only the parts.
        given = {str(g).strip("-").lower()
                 for g in item.get("given", blk.get("given", []))}
        plus_w, gap = 0.45, 0.0
        bw = (8.2 - plus_w * (len(parts) - 1)) / len(parts)
        for i, pt in enumerate(parts):
            show = item.get("worked") or pt.strip("-").lower() in given
            bx = 0.9 + i * (bw + plus_w)
            add_textbox(root, sid, bx, y, bw, 0.82, [pt if show else ""], 26,
                        bold=True, fill="FFFFFF", outline=accent, rounded=True)
            sid += 1
            if i < len(parts) - 1:
                add_textbox(root, sid, bx + bw, y, plus_w, 0.82, ["+"], 26, bold=True)
                sid += 1
        y += 0.82 + 0.08
    bar_h = min(0.8, ROUND_BAR_BOTTOM - y)
    if bar_h < 0.45:
        warn(f"{ctx}: no room for the answer bar - shorten the prompt or drop a choice")
        bar_h = 0.45
    answer = round_answer_text(item)
    asz = fit_font_block(answer.replace("**", ""), 8.0, bar_h - 0.1, 20,
                         min_pt=14, char_em=0.58)
    answer_sid = sid
    add_textbox(root, sid, 0.9, y, 8.2, bar_h, [answer], asz,
                color="FFFFFF", fill=ANSWER_GREEN, rounded=True)
    sid += 1
    routine = item.get("routine", blk.get("round_routine", blk.get("routine")))
    if routine:
        csz = fit_font(routine, 3.55, 15, min_pt=13, char_em=0.58)
        add_textbox(root, sid, 0.5, 4.66, 3.9, 0.44, [routine], csz,
                    color="FFFFFF", bold=True, fill=accent, rounded=True)
        sid += 1
    footer = item.get("footer", blk.get("round_footer", blk.get("footer")))
    if footer:
        add_textbox(root, sid, 4.55, 4.66, 4.95, 0.44, [footer], 14,
                    align="r", italic=True, color="595959")
    return answer_sid


def insert_timing(root, timing):
    """Place a <p:timing> block where PresentationML expects it (before extLst)."""
    old = root.find(q("p:timing"))
    if old is not None:
        root.remove(old)
    ext = root.find(q("p:extLst"))
    if ext is not None:
        ext.addprevious(timing)
    else:
        root.append(timing)


def emit_round(d, blk, tidx, accent, tint, ctx, label):
    """Emit one slide per `round` item, each answer hidden behind one click.

    An item marked `worked` is the "One done for you" model: same layout,
    answer visible from the start.
    """
    timing_src = etree.fromstring(d.pkg.slide_xml(T["spell_word"])).find(q("p:timing"))
    for n, item in enumerate(blk.get("round", []), 1):
        for field in ("prompt", "answer", "notes"):
            if not str(item.get(field, "")).strip():
                warn(f"{ctx}: round item {n} has no {field}")

        def fill(root, item=item, n=n):
            tsp = find_sp(root, SHAPE["gr_title"][tidx])
            set_text(tsp, label + blk.get("title", ""))
            bsp = find_sp(root, SHAPE["gr_body"][tidx])
            set_runs(bsp, [[]])
            answer_sid = draw_round_slide(root, blk, item, accent, tint,
                                          f"{ctx} round item {n}")
            if item.get("worked"):
                return  # "One done for you": same layout, answer already showing
            insert_timing(root, copy.deepcopy(timing_src))
            regen_timing(root, [(answer_sid, None)])
        d.add(tidx, fill, notes=item.get("notes", ""))


def sanitize(s):
    if s is None:
        return s
    for a, b in [("–", "-"), ("—", "-"), ("‘", "'"), ("’", "'"),
                 ("“", '"'), ("”", '"'), ("…", "...")]:
        s = s.replace(a, b)
    return s


def deep_sanitize(obj):
    if isinstance(obj, str):
        return sanitize(obj)
    if isinstance(obj, list):
        return [deep_sanitize(x) for x in obj]
    if isinstance(obj, dict):
        return {k: deep_sanitize(v) for k, v in obj.items()}
    return obj


DICTATION_CUE_RE = re.compile(r"([A-Z]|[,.;:!?\"'“”-]+)")


def smart_dictation_quotes(sentence):
    """Directional speech marks for the dictation slide face only.

    The locked sentence box renders in Lexend, whose ASCII double-quote glyph
    slants like a CLOSING quote, so a straight opening quote looks backwards
    on screen (school feedback). Specs and notes stay ASCII; this converts
    paired straight quotes at render time: first of each pair opens, second
    closes.
    """
    if sentence.count('"') % 2:
        warn(f"dictation sentence has an unpaired quotation mark: {sentence[:40]}")
        return sentence
    out = []
    opening = True
    for ch in sentence:
        if ch == '"':
            out.append("“" if opening else "”")
            opening = not opening
        else:
            out.append(ch)
    return "".join(out)


def cue_runs(text, bold=None, underline=None):
    """Colour capital letters green and punctuation red for COPS checking."""
    runs = []
    for seg in DICTATION_CUE_RE.split(text):
        if not seg:
            continue
        if re.fullmatch(r"[A-Z]", seg):
            color = CAPITAL_GREEN
        elif re.fullmatch(r"[,.;:!?\"'“”-]+", seg):
            color = PUNCT_RED
        else:
            color = None
        runs.append((seg, color, bold, underline))
    return runs


def dictation_runs(sentence, targets):
    """Reveal runs: capitals green, punctuation red, targets bold+underlined."""
    if not targets:
        return [cue_runs(sentence)]
    pattern = re.compile(
        r"\b(" + "|".join(re.escape(t) for t in sorted(targets, key=len, reverse=True)) + r")\b",
        re.IGNORECASE)
    runs = []
    pos = 0
    for m in pattern.finditer(sentence):
        if m.start() > pos:
            runs.extend(cue_runs(sentence[pos:m.start()]))
        runs.extend(cue_runs(m.group(0), bold=True, underline=True))
        pos = m.end()
    if pos < len(sentence):
        runs.extend(cue_runs(sentence[pos:]))
    return [runs]


def lw_runs(word, unfair):
    """White word with the unfair part in yellow."""
    i = word.lower().find(unfair.lower()) if unfair else -1
    if unfair and i < 0:
        warn(f"learned word '{word}': unfair part '{unfair}' not found; highlighting nothing")
    if i < 0:
        return [[(word, "FFFFFF", None, None)]]
    runs = []
    if i > 0:
        runs.append((word[:i], "FFFFFF", None, None))
    runs.append((word[i:i + len(unfair)], "FFFF00", None, None))
    if i + len(unfair) < len(word):
        runs.append((word[i + len(unfair):], "FFFFFF", None, None))
    return [runs]


def validate_build_rounds(day, week, session):
    """The 2f/3d one-recall gate for the v1.7 build-day rounds.

    Break-apart and name-it rounds are You Do work, so every part a slide shows
    beyond the week's morpheme carries its printed meaning, exactly as the
    legacy You Do round does.  The focus is the week's taught morpheme.
    """
    bd = session.get("build")
    if not isinstance(bd, dict):
        return
    focus = set()
    for label in week_focus_labels(week):
        focus |= morph_surface_forms(label)
    taught = taught_morpheme_forms(week, session)
    catalogue = catalogue_morph_forms()
    reported = set()
    for rnd in bd.get("rounds", []):
        for item in rnd.get("round", []):
            _scan_activity_face(day, list(round_face_lines(rnd, item)), focus,
                                taught, catalogue, reported)


def retitle(root, old, new):
    """Replace `old` with `new` in every text run on a slide (header retitles)."""
    hit = False
    for t_el in root.iter(q("a:t")):
        if t_el.text and old in t_el.text:
            t_el.text = t_el.text.replace(old, new)
            hit = True
    return hit


def draw_matrix_slide(root, blk, base, ctx):
    """Word matrix (v1.7 build day): prefixes | the week's morpheme | suffixes.

    Kate Watt's Wednesday task and Yoshimoto's own rupt matrix: students build
    real words from the parts in their OG books and say what each one means.
    Every non-focus part carries its printed meaning (2f one-recall rule).
    """
    sid, y = 90100, 1.15
    if blk.get("rule"):
        rsz = fit_font_block(blk["rule"], 8.6, 0.58, 18, min_pt=15)
        add_textbox(root, sid, 0.5, y, 9.0, 0.64, [blk["rule"]], rsz,
                    color="FFFFFF", bold=True, fill=MORPH_MAGENTA, rounded=True)
        sid += 1
        y += 0.76
    box_h = 4.50 - y
    side = [blk.get("prefixes", []), blk.get("suffixes", [])]
    longest = max((str(p) for col in side for p in col), key=len, default="x")
    rows = max(1, max(len(col) for col in side))
    ssz = fit_font_block(longest, 2.7, (box_h - 0.3) / rows, 24,
                         min_pt=14, char_em=0.55)
    for x, col in ((0.5, side[0]), (6.6, side[1])):
        add_textbox(root, sid, x, y, 2.9, box_h, [str(p) for p in col] or [""],
                    ssz, fill=MORPH_TINT, rounded=True, outline=MORPH_MAGENTA)
        sid += 1
    base_lines = [f"**{base['morph']}**"]
    if base.get("meaning"):
        base_lines.append(f"({base['meaning']})")
    bsz = fit_font_block(base["morph"], 2.6, box_h / 3, 54, min_pt=28, char_em=0.6)
    add_textbox(root, sid, 3.55, y, 2.9, box_h, base_lines[:1], bsz,
                fill=MORPH_FILL.get(base.get("type"), MORPH_TINT), rounded=True)
    sid += 1
    if len(base_lines) > 1:
        msz = fit_font_block(base_lines[1], 2.6, 0.5, 20, min_pt=13, char_em=0.55)
        add_textbox(root, sid, 3.55, y + box_h - 0.62, 2.9, 0.55, base_lines[1:],
                    msz, color=CHARCOAL)
        sid += 1
    if blk.get("routine"):
        csz = fit_font(blk["routine"], 3.55, 15, min_pt=13, char_em=0.58)
        add_textbox(root, sid, 0.5, 4.66, 3.9, 0.44, [blk["routine"]], csz,
                    color="FFFFFF", bold=True, fill=MORPH_MAGENTA, rounded=True)
        sid += 1
    if blk.get("footer"):
        add_textbox(root, sid, 4.55, 4.66, 4.95, 0.44, [blk["footer"]], 14,
                    align="r", italic=True, color="595959")


SESSION_PLANS = {
    # legacy v1.0-v1.6 weeks: every component every day
    "new": {"cards", "wtr", "sound_bank", "wts_review", "morph_hdr", "nm_card",
            "nm_copy", "grid", "wts_new", "you_do", "lw", "dictation", "grammar"},
    "review": {"cards", "wtr", "sound_bank", "wts_review", "morph_hdr", "nm_card",
               "nm_copy", "grid", "wts_new", "you_do", "lw", "dictation", "grammar"},
    "week_review": {"cards", "wtr", "sound_bank", "wts_review", "lw", "dictation",
                    "grammar"},
    # v1.7 five-day shape (section 1)
    "teach": {"cards", "wtr", "morph_hdr", "nm_card", "nm_copy", "grid",
              "keyword_sentence"},
    "spell": {"recap", "sound_bank", "wts_review", "wts_new", "lw", "dictation"},
    "build": {"build"},
    "apply": {"pos_drill", "lw", "grammar"},
    "week_review_v17": {"cards", "wtr", "sound_bank", "wts_review", "lw",
                        "dictation", "grammar", "vote"},
    # v1.8: the daily review (cards + Words to Read Review) opens EVERY session,
    # Tuesday's recap card goes (port is in Tuesday's cards now), and a short
    # week cuts the Friday review day first (section 11).
    "teach_v18": {"cards", "wtr", "morph_hdr", "nm_card", "nm_copy", "grid",
                  "keyword_sentence"},
    "spell_v18": {"cards", "wtr", "sound_bank", "wts_review", "wts_new", "lw",
                  "dictation"},
    "build_v18": {"cards", "wtr", "build"},
    "apply_v18": {"cards", "wtr", "pos_drill", "lw", "grammar", "dictation",
                  "vote"},
    "combined_v18": {"cards", "wtr", "build", "pos_drill", "lw", "grammar",
                     "vote"},
    "week_review_v18": {"cards", "wtr", "sound_bank", "wts_review", "lw",
                        "dictation", "grammar", "vote"},
    # v1.9: Friday opens with the daily review, then the Friday Review
    # worksheet (1a) - the deck stops there; the follow-up sheet replaces the vote
    "week_review_v19": {"cards", "wtr", "friday_review"},
}

V17_REQUIRED = {
    "teach": ["morphology_review", "words_to_read_review", "new_morphology",
              "words_to_read_new", "wtr_new_notes", "keyword_sentence"],
    "spell": ["recap", "sound_bank", "words_to_spell_review",
              "words_to_spell_new", "learned_words", "dictation"],
    "build": ["build"],
    "apply": ["pos_drill", "learned_words", "grammar"],
    "week_review_v17": ["morphology_review", "words_to_read_review", "sound_bank",
                        "words_to_spell_review", "learned_words", "dictation",
                        "vote"],
    "teach_v18": ["morphology_review", "words_to_read_review", "new_morphology",
                  "words_to_read_new", "wtr_new_notes", "keyword_sentence"],
    "spell_v18": ["morphology_review", "words_to_read_review", "sound_bank",
                  "words_to_spell_review", "words_to_spell_new", "learned_words",
                  "dictation"],
    "build_v18": ["morphology_review", "words_to_read_review", "build"],
    "apply_v18": ["morphology_review", "words_to_read_review", "pos_drill",
                  "learned_words", "grammar"],
    "combined_v18": ["morphology_review", "words_to_read_review", "build",
                     "pos_drill", "grammar", "vote"],
    "week_review_v18": ["morphology_review", "words_to_read_review", "sound_bank",
                        "words_to_spell_review", "learned_words", "dictation",
                        "vote"],
    "week_review_v19": ["morphology_review", "words_to_read_review", "friday_review"],
}


def session_plan_key(week, session):
    stype = session.get("type", "new")
    if week_is_v19(week) and stype == "week_review":
        return "week_review_v19"
    if week_is_v18(week) and stype in V17_TYPES | {"week_review"}:
        return stype + "_v18"
    if stype == "week_review" and week_is_v17(week):
        return "week_review_v17"
    return stype


def week_teach_grid(week):
    """Every word on this week's teach-day grids (v1.7 spelling comes from them)."""
    words = set()
    for s in (week or {}).get("sessions", []):
        if s.get("type") == "teach":
            words |= {str(w.get("word") if isinstance(w, dict) else w).lower()
                      for w in s.get("words_to_read_new", [])}
    return words


def build_session(pkg, week, session, out_path):
    d = DeckBuilder(pkg)
    day = session["day"]
    stype = session.get("type", "new")  # see SESSION_PLANS
    plan_key = session_plan_key(week, session)
    if plan_key not in SESSION_PLANS:
        warn(f"{day}: unknown session type {stype!r}")
        plan_key = "new"
    plan = SESSION_PLANS[plan_key]
    v17 = week_is_v17(week)
    v18 = week_is_v18(week)
    v19 = week_is_v19(week)
    term, wknum = week["term"], week["week"]
    if v18:
        # James Hooke (23 Oct 2026): ask real questions ("What is one word built
        # from terr, meaning land?"), never "give me"; the "SAY each word"
        # line above the spelling extension was noise.
        blob = json.dumps(session, ensure_ascii=False)
        if re.search(r"\bgive me\b", blob, flags=re.I):
            warn(f"{day}: a script says \"give me\" - ask the real question instead "
                 "(\"What is one word built from terr, meaning land?\") (2a, v1.8)")
        # James Hooke (23 Oct 2026): "Purpose and prep - what's all that down
        # the bottom?"  v1.8 notes carry no Purpose line and no stage tag; a
        # teacher who needs the reason gets one plain WHY: line (3c).
        if re.search(r'"Purpose:|\[(?:I|We|You) Do \||\[Review \|', blob):
            warn(f"{day}: notes carry a Purpose line or a [Stage | ...] tag - "
                 "v1.8 notes say why in plain words on one WHY: line (3c)")
        if "SAY each word, students write it" in blob:
            warn(f"{day}: drop the \"SAY each word, students write it...\" line from "
                 "the spelling notes (3c, v1.8)")
    cohort = week.get("cohort", "Enrichment")
    marks = {"review_cards": [], "pos_cards": []}
    mh = None

    def want(block):
        return block in plan

    for key in V17_REQUIRED.get(plan_key, []):
        if not session.get(key):
            warn(f"{day}: a {stype} day needs `{key}` (section 1, v1.7)")
    if v18:
        # Short weeks (section 11, v1.8): the review day goes first, so the
        # last session carries the vote, and a 4-session week's apply day
        # keeps a second dictation day with one sentence in the grammar focus.
        last = session is week["sessions"][-1]
        if last and not session.get("vote") and plan_key != "week_review_v19":
            warn(f"{day}: the week's last session needs the trickiest-this-week "
                 "`vote` (section 1, v1.8)")
        if plan_key == "apply_v18":
            n_dict = len(session.get("dictation") or [])
            if last and n_dict != 1:
                warn(f"{day}: a short week's apply day ends with ONE dictation "
                     "sentence in the week's grammar focus (section 11, v1.8)")
            elif not last and n_dict:
                warn(f"{day}: the apply day only carries dictation when it is the "
                     "week's last session (section 11, v1.8)")
    timing_src = etree.fromstring(pkg.slide_xml(T["spell_word"])).find(q("p:timing"))

    # --- 1 title
    def fill_title(root):
        sp = find_sp(root, SHAPE["title_sub"])
        set_runs(sp, [[(cohort, None, None, None)],
                      [(f"Term {term} Week {wknum} - {day}", None, None, None)]])
        # widen the subtitle box (kept centred) so the day fits on one line
        xfrm = sp.find(q("p:spPr") + "/" + q("a:xfrm"))
        off, ext = xfrm.find(q("a:off")), xfrm.find(q("a:ext"))
        new_w = int(4.6 * 914400)
        cx = int(off.get("x")) + int(ext.get("cx")) // 2
        off.set("x", str(cx - new_w // 2))
        ext.set("cx", str(new_w))
    title_notes = f"{cohort} OG session deck. Term {term} Week {wknum}, {day}."
    if v19 and enabler_pdf.wants_sheet(session):
        # v1.9 (1b): the day's sheets named where the teacher starts the deck
        num = week["sessions"].index(session) + 1
        title_notes += (f"\nSHEETS TODAY: Enablers/{enabler_pdf.file_name(num, day)} for "
                        "students who need support; Extension/"
                        f"{enabler_pdf.extension_file_name(num, day)} for students ready "
                        "for more. Both are in the week's print pack.")
    d.add(T["title"], fill_title, notes=title_notes)

    # --- 2 weekly overview
    ov = week["overview"]  # dict day -> {morphology, grammar, learned_words}
    days = week.get("days", ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"])

    def fill_overview(root):
        rows = table_cells(root, SHAPE["overview_table"])
        # rows[0] headers; rows[1] morphology; rows[2] grammar; rows[3] learned words
        if v17:
            # the master says "1 hour"; v1.7 sessions run 20-30 minutes
            set_cell_paras(rows[0].findall(q("a:tc"))[0], ["20-30 min"])
        for ci, dy in enumerate(days, start=1):
            cells_m = rows[1].findall(q("a:tc"))
            cells_g = rows[2].findall(q("a:tc"))
            cells_l = rows[3].findall(q("a:tc"))
            if ci >= len(cells_m):
                break
            o = ov.get(dy, {})
            set_cell_paras(cells_m[ci], o.get("morphology", [""]) if isinstance(o.get("morphology"), list) else [o.get("morphology", "")])
            set_cell_paras(cells_g[ci], [o.get("grammar", "")])
            set_cell_paras(cells_l[ci], o.get("learned_words", [""]) if isinstance(o.get("learned_words"), list) else [o.get("learned_words", "")])
    def cell_text(value):
        # overview cells may be [label, value] lists; notes read them as plain text
        return ": ".join(value) if isinstance(value, list) else str(value or "")
    overview_notes = "\n".join(
        f"{dy} - {cell_text(ov.get(dy, {}).get('morphology'))} | Grammar: "
        f"{cell_text(ov.get(dy, {}).get('grammar'))} | LW: "
        f"{cell_text(ov.get(dy, {}).get('learned_words'))}" for dy in days)
    if v19:
        # v1.9 (1c): the week's print pack named in every deck, as plain text -
        # a relative link would not survive the Google Slides import
        # a short week with no week_review day has no Friday Review to name
        fr_pages = (", extension and Friday Review" if any(
            s.get("friday_review") for s in week["sessions"]) else " and extension")
        overview_notes += (f"\n---\nPRINT PACK: {print_pack.file_name(week)} in the week "
                           f"folder - every enabler{fr_pages} page for "
                           "the week, with a cover saying who each page is for.")
    d.add(T["overview"], fill_overview, notes=overview_notes)

    # --- catalogue and teaching-history gates
    for card in session.get("morphology_review", []):
        check_banked(day, "review card", card["morph"], card.get("type"),
                     card.get("meaning"))
    for item in session.get("sound_bank", []):
        if item.get("type") == "sound":
            continue  # phonograms come from the sound deck, not the morph cards
        check_banked(day, "sound bank", item["morph"], item.get("type"),
                     item.get("meaning"))
    for card in (session.get("pos_drill") or {}).get("cards", []):
        check_banked(day, "parts-of-speech card", card["morph"], "suffix",
                     card.get("meaning"))
    if session.get("new_morphology"):
        new_morph = session["new_morphology"]
        check_banked(day, "new morphology", new_morph["morph"],
                     new_morph.get("type"), new_morph.get("meaning"))
    validate_session_card_words(day, session)
    validate_sound_bank(day, session, week)
    validate_activity_morphemes(day, week, session)
    validate_build_rounds(day, week, session)
    if v19 and session.get("build"):
        validate_v19_rounds(day, week, session)
    validate_dictation_sources(day, week, session)
    if v19:
        audit_school_cards(day, session)

    # --- 3 morphology review cards x10
    if want("cards"):
        d.add(T["hdr_morph_review"], notes=HEADER_NOTES["morph_review"])
        for card in session["morphology_review"]:
            tidx = CARD_REVIEW[card["type"]]

            def fill_card(root, card=card, tidx=tidx):
                set_slide_background(root, MORPH_FILL[card["type"]])
                sp = find_sp(root, SHAPE["card_review_word"][tidx])
                set_text(sp, card["morph"],
                         size_pt=fit_font(card["morph"], 5.75, 56))
            kw, mean = resolve_card(day, "review card", card)
            card_note = (f"{card['type'].capitalize()}: {card['morph']}\n"
                         f"Keyword: {kw}\n"
                         f"{card_meaning_note(mean)}"
                         f"{card_pos_note(card)}"
                         f"{extra_task_note(card)}")
            marks["review_cards"].append((len(d.slides), card))
            d.add(tidx, fill_card, notes=card_note)

    # --- words to read review
    if want("wtr"):
        d.add(T["hdr_wtr_review"], notes=HEADER_NOTES["wtr_review"])
        wtr = session["words_to_read_review"]
        if len(wtr["words"]) != 15:
            warn(f"{day}: words_to_read_review has {len(wtr['words'])} words (template holds 15)")
        review_notes = prepare_notes_text(wtr.get("notes", ""))
        min_reads = 2 if v18 else 3
        if len(re.findall(r"^\d+\. READ:", review_notes, flags=re.MULTILINE)) < min_reads:
            warn(f"{day}: Words to Read Review needs one numbered READ line per "
                 "group plus the whole-class read")
        if v18:
            asks = re.findall(r"^\d+\. ASK:", review_notes, flags=re.MULTILINE)
            if not 1 <= len(asks) <= 3:
                warn(f"{day}: Words to Read Review has {len(asks)} questions; v1.8 "
                     "asks three at most, one-word answers from the board (2b)")
            if len(re.findall(r"turn and (?:talk|tell)", review_notes, flags=re.I)) < 2:
                warn(f"{day}: Words to Read Review needs at least two questions run "
                     "as turn and talk, then a choral answer (2b, v1.8)")
        if v17:
            # Kate Watt (Oct 2026): everyone decodes every row first; the
            # group rows are the second, fluency read.
            if not re.search(r"^1\. READ:.*\b(every|all)\b.*\brows?\b", review_notes,
                             flags=re.MULTILINE | re.IGNORECASE):
                warn(f"{day}: Words to Read Review must open with `1. READ:` the "
                     "whole class reading every row together (2b, v1.7)")
        elif "Every student reads at least two rows" not in review_notes:
            warn(f"{day}: Words to Read Review must state that every student "
                 "reads at least two rows in the group round")

        def fill_wtr_table(root):
            rows = table_cells(root, SHAPE["wtr_review_tbl"])
            n = 0
            for row in rows:
                for tc in row.findall(q("a:tc")):
                    if n < len(wtr["words"]):
                        set_cell_paras(tc, [f"{n + 1}) {wtr['words'][n]}"])
                    else:
                        set_cell_paras(tc, [""])
                    n += 1
        d.add(T["wtr_review_table"], fill_wtr_table, notes=wtr["notes"])

    # --- v1.7 Tuesday: quick recap of Monday's card
    if want("recap") and session.get("recap"):
        rc = session["recap"]
        rtidx = CARD_NEW[rc["type"]]

        def fill_recap_hdr(root):
            retitle(root, "New Morphology", "Quick Recap")
        d.add(T["hdr_new_morph"], fill_recap_hdr, notes=HEADER_NOTES["recap"])

        def fill_recap(root, rc=rc, rtidx=rtidx):
            set_slide_background(root, MORPH_FILL[rc["type"]])
            sp = find_sp(root, SHAPE["card_new_word"][rtidx])
            set_text(sp, rc["morph"], size_pt=fit_font(rc["morph"], 5.75, 56))
        rkw, rmean = resolve_card(day, "recap card", rc)
        recap_note = (f"{rc['type'].capitalize()}: {rc['morph']}\n"
                      f"Keyword: {rkw}\n"
                      f"{card_meaning_note(rmean)}"
                      f"{card_pos_note(rc)}")
        if rc.get("notes"):
            recap_note += "\n" + rc["notes"]
        d.add(rtidx, fill_recap, notes=recap_note)

    # --- sound bank
    if want("sound_bank"):
        sb = session["sound_bank"]
        if len(sb) > 9:
            warn(f"{day}: sound bank has {len(sb)} entries; only 9 boxes - extras dropped")
        d.add(T["hdr_sound_bank"],
              notes=HEADER_NOTES["sound_bank_v17" if v17 else "sound_bank"])

        def fill_sound_bank(root):
            filled = []
            for i, spid in enumerate(SHAPE["sound_boxes"]):
                sp = find_sp(root, spid)
                if i < len(sb):
                    item = sb[i]
                    label = str(item["morph"]).strip()
                    # A 2.33" x 1.04" box cannot hold a variant label such as
                    # clud(e)/clus(e) on one line without shrinking it to an
                    # unreadable size, and left alone it breaks mid-bracket. Split
                    # it after the slash so each half renders whole and large.
                    one_line = fit_font(label, 2.33, 45, min_pt=24)
                    # PowerPoint wraps "pon/pound" mid-word ("pon/poun" + "d")
                    # even when the estimate says it fits (James Hooke, Oct
                    # 2026), so any longer slashed label breaks after a slash.
                    if "/" in label and (one_line <= 28 or len(label) > 6):
                        cuts = [i + 1 for i, ch in enumerate(label) if ch == "/"]
                        cut = min(cuts, key=lambda i: abs(len(label) - 2 * i))
                        head, tail = label[:cut], label[cut:]
                        parts = [head.strip(), tail.strip()]
                        # two lines must still clear the 1.04" box, so drop the
                        # body insets and cap at the height-derived maximum
                        sz = min(28,
                                 fit_font(max(parts, key=len), 2.33, 29, min_pt=22))
                        set_runs(sp, [[(part, None, None, None)] for part in parts],
                                 size_pt=sz)
                        bodypr = sp.find(q("p:txBody") + "/" + q("a:bodyPr"))
                        if bodypr is not None:
                            for inset in ("tIns", "bIns"):
                                bodypr.set(inset, "0")
                    else:
                        set_text(sp, label, size_pt=one_line)
                    set_solid_fill(sp, SOUND_FILL if item.get("type") == "sound"
                                   else MORPH_FILL[item["type"]])
                    filled.append(spid)
                elif v17:
                    # an empty dotted box looks like a missing answer (v1.7
                    # banks hold 6-9), so unused boxes come off the slide
                    sp.getparent().remove(sp)
                else:
                    set_text(sp, "")
                    set_solid_fill(sp, "FFFFFF")
            if v17 and filled:
                # Yoshimoto "Auditory: Sound": the teacher says it, students
                # write it, the click shows the box. Retrieval, not copying.
                insert_timing(root, copy.deepcopy(timing_src))
                regen_timing(root, [(spid, None) for spid in filled])
        if v17:
            say_lines = ["Say each one, students write it, then click to show it:"]
            for n, item in enumerate(sb[:9], 1):
                if item.get("type") == "sound":
                    say_lines.append(f"{n}. {item['morph']} - {item.get('say', 'say the sound')}")
                else:
                    _kw, meaning = resolve_card(day, "sound bank", item)
                    say_lines.append(f"{n}. {item['morph']} - {meaning}")
            say_lines += ["", "green = root", "yellow = prefix", "red = suffix",
                          "white = sound"]
            bank_notes = "\n".join(say_lines)
        else:
            bank_notes = ("Students copy these into their books before the spelling words.\n"
                          "green = root\n"
                          "yellow = prefix\n"
                          "red = suffix")
        marks["sound_bank"] = len(d.slides)
        d.add(T["sound_bank"], fill_sound_bank, notes=bank_notes)

    # --- words to spell review (one reveal slide per word)
    if want("wts_review"):
        d.add(T["hdr_wts_review"], notes=HEADER_NOTES["wts_review"])
        wsr = session["words_to_spell_review"]
        lo, hi = (6, 10) if v17 else (8, 12)
        if not lo <= len(wsr) <= hi:
            warn(f"{day}: words_to_spell_review has {len(wsr)} words "
                 f"(expected {lo}-{hi})")
        if v18 and sum(1 for item in wsr if item.get("prompt")) > 3:
            warn(f"{day}: more than 3 after-check questions; keep them for words "
                 "where the part's meaning plainly shows (2d, v1.8)")
        for n_word, item in enumerate(wsr, 1):
            pic_media, pic_w = None, 0.0
            if v19:
                # v1.9 (from the Challenge team's planner decks): every spelling
                # word carries a picture of its meaning, on the word's own click.
                # A picture must be clear with no explanation (James, 30 Oct 2026);
                # where none can be, {"none": reason} shows the word alone.
                pic = item.get("picture")
                if isinstance(pic, dict) and "none" in pic:
                    n_none = sum(1 for e in wsr if isinstance(e.get("picture"), dict)
                                 and "none" in e["picture"])
                    if not str(pic["none"]).strip():
                        warn(f"{day}: '{item['word']}' has picture none with no reason "
                             "(2d, v1.9)")
                    elif n_none > 2:
                        warn(f"{day}: {n_none} spelling words have no picture - at most "
                             "2; pick words whose meaning can be pictured (2d, v1.9)")
                    else:
                        note(f"{day}: '{item['word']}' shows no picture - "
                             f"{pic['none']} (2d, v1.9)")
                elif not pic:
                    warn(f"{day}: Words to Spell Review word '{item['word']}' needs a "
                         "meaning `picture` (2d, v1.9)")
                else:
                    pic_png = meaning_picture_png(item["picture"])
                    if pic_png is None:
                        warn(f"{day}: meaning picture {item['picture']!r} for "
                             f"'{item['word']}' is not a pictogram or a react-icons "
                             "pack/Component name (2d, v1.9)")
                    else:
                        pic_media = f"ogspell_{day.lower()}_{n_word}.png"
                        d.media[f"ppt/media/{pic_media}"] = pic_png
                        with Image.open(BytesIO(pic_png)) as im_:
                            pic_w = im_.width / im_.height

            def fill_spell(root, item=item, pic_media=pic_media, pic_w=pic_w):
                sp = find_sp(root, SHAPE["spell_word_box"])
                set_text(sp, item["word"], size_pt=fit_font(item["word"], 4.96, 60, min_pt=28))
                if pic_media:
                    # 1.45" tall, centred under the word; a wide composite shrinks
                    # to fit 6" so it never reaches the slide edges
                    h = min(1.45, 6.0 / pic_w)
                    w = h * pic_w
                    add_picture(root, 90070, "rIdOGSpellPic", (10 - w) / 2,
                                3.55 + (1.45 - h) / 2, w, h)
                    add_with_effect(root, 90070)

            def fill_spell_rels(rels, pic_media=pic_media):
                if pic_media:
                    relate_image(rels, "rIdOGSpellPic", f"../media/{pic_media}")
            if not item.get("prompt") and not v18:
                warn(f"{day}: spell review word '{item['word']}' has no after-check prompt")
            elif item.get("prompt") and not item.get("answer"):
                warn(f"{day}: spell review word '{item['word']}' prompt has no answer")
            note_text = (f"Word to spell: {item['word']}\n"
                         f"Sentence: {item['sentence']}\n"
                         + ("" if v18 else "Click to reveal answer."))
            if item.get("prompt"):
                note_text += f"\nAfter checking: {item['prompt']}"
            if item.get("answer"):
                note_text += f"\nAnswer: {item['answer']}"
            d.add(T["spell_word"], fill_spell, notes=note_text,
                  fill_rels=fill_spell_rels if pic_media else None)

    # --- new/review morphology section
    section = "New Morphology" if stype in ("new", "teach", "spell") else "Review Morphology"
    nm = session.get("new_morphology")
    if want("morph_hdr") and nm:
        def fill_nm_hdr(root):
            sp = find_sp(root, SHAPE["hdr_new_morph_txt"])
            replace_run_text(sp, "New Morphology", section)
        d.add(T["hdr_new_morph"], fill_nm_hdr,
              notes=HEADER_NOTES["new_morph_v18" if v18 and stype == "teach"
                                 else "new_morph_v17" if stype == "teach"
                                 else "new_morph"])

    if want("nm_card") and nm:
        tidx = CARD_NEW[nm["type"]]

        # v1.8 (James Hooke, 23 Oct 2026): the memory hook is short, sits at
        # the top of the notes and on the card face with a small picture, so
        # it can be retrieved in an instant.
        mh = session.get("memory_hook") if v18 else None
        hook_png = None
        if v18 and stype == "teach":
            if not mh or not mh.get("text") or not mh.get("picture"):
                warn(f"{day}: a v1.8 teach day needs `memory_hook` with `text` and "
                     "`picture` (3a)")
                mh = None
            else:
                if len(mh["text"].split()) > 8:
                    warn(f"{day}: memory hook {mh['text']!r} is over 8 words (3a, v1.8)")
                hook_png = pictogram_png(mh["picture"])
                if hook_png is None:
                    warn(f"{day}: memory hook picture {mh['picture']!r} is not in the "
                         "pictogram library (themes/core/pictograms.js)")
        hook_media = None
        if hook_png:
            hook_media = f"oghook_{re.sub(r'[^a-z0-9]+', '_', mh['picture'].lower())}.png"
            d.media[f"ppt/media/{hook_media}"] = hook_png

        def fill_nm_card(root, tidx=tidx, nm=nm):
            set_slide_background(root, MORPH_FILL[nm["type"]])
            sp = find_sp(root, SHAPE["card_new_word"][tidx])
            set_text(sp, nm["morph"], size_pt=fit_font(nm["morph"], 5.75, 56))
            if mh:
                add_textbox(root, 90060, 1.45, 3.7, 7.1, 1.2, [""], 12,
                            fill="FFFFFF", rounded=True)
                hsz = fit_font_block(mh["text"], 5.4, 1.0, 24, min_pt=18)
                add_textbox(root, 90061, 2.7 if hook_media else 1.6, 3.75,
                            5.8 if hook_media else 6.8, 1.1, [mh["text"]], hsz,
                            bold=True, align="l" if hook_media else "ctr")
                if hook_media:
                    add_picture(root, 90062, "rIdOGHook", 1.62, 3.8, 1.0, 1.0)

        def fill_nm_rels(rels):
            if hook_media:
                relate_image(rels, "rIdOGHook", f"../media/{hook_media}")
        nm_kw, nm_mean = resolve_card(day, "new morphology", nm)
        marks["nm_card"] = len(d.slides)
        d.add(tidx, fill_nm_card,
              notes=((f"MEMORY HOOK: {mh['text']}\n" if mh else "")
                     + f"{nm['type'].capitalize()}: {nm['morph']}\n"
                     f"Keyword: {nm_kw}\n"
                     f"{card_meaning_note(nm_mean)}"
                     f"{card_pos_note(nm)}"),
              fill_rels=fill_nm_rels if hook_media else None)

        # Immediately follow the delivered morpheme card with a stable,
        # non-animated reference students can copy into their books.
        def fill_nm_copy(root, tidx=tidx, nm=nm, nm_kw=nm_kw,
                         nm_mean=nm_mean):
            set_slide_background(root, MORPH_FILL[nm["type"]])
            sp = find_sp(root, SHAPE["card_new_word"][tidx])
            meaning_lines = card_meaning_note(nm_mean).splitlines()
            fields = [(nm["type"].capitalize(), nm["morph"]),
                      ("Keyword", nm_kw)]
            for line in meaning_lines:
                label, _, value = line.partition(":")
                fields.append((label, value.strip()))
            parts = card_parts_of_speech(nm)
            if parts:
                fields.append(("Part of speech", ", ".join(parts)))
            display_text = "\n".join(
                f"{label}: {value}" for label, value in fields
            )
            sz = fit_font_block(display_text, 6.2, 3.0, 35,
                                min_pt=24, char_em=0.55)
            set_runs(
                sp,
                [[(f"{label}: ", None, True, None),
                  (value, None, False, None)]
                 for label, value in fields],
                size_pt=sz,
            )
            xfrm = sp.find(q("p:spPr") + "/" + q("a:xfrm"))
            xfrm.find(q("a:off")).set("x", str(int(1.7 * 914400)))
            xfrm.find(q("a:off")).set("y", str(int(1.35 * 914400)))
            xfrm.find(q("a:ext")).set("cx", str(int(6.6 * 914400)))
            xfrm.find(q("a:ext")).set("cy", str(int(3.1 * 914400)))
            bodypr = sp.find(q("p:txBody") + "/" + q("a:bodyPr"))
            if bodypr is not None:
                bodypr.set("anchor", "ctr")
            for p_el in sp.findall(q("p:txBody") + "/" + q("a:p")):
                ppr = p_el.find(q("a:pPr"))
                if ppr is not None:
                    ppr.set("algn", "l")
            remove_timing(root)

        copy_note_fields = "morpheme, keyword and meaning"
        if card_parts_of_speech(nm):
            copy_note_fields += ", including the part of speech"
        copy_notes = (
            "1. SAY: \"Copy this morpheme card into your book exactly as shown.\"\n"
            f"2. CHECK: Scan the {copy_note_fields} before moving on."
        )
        if stype == "teach":
            # Yoshimoto's new-morpheme routine (Advanced Morphology training):
            # write it three times saying it, then the class brainstorms words
            # before the grid shows any.
            copy_notes += (f"\n3. WRITE: \"Write {nm['morph']} three times, saying it "
                           "each time.\"")
            if session.get("brainstorm"):
                copy_notes += f"\n4. ASK: {session['brainstorm']}"
        # The memory hook is the retention device for the whole week, so the
        # teacher needs it here, on the anchor card, not only at the end of the
        # word-grid script.
        hook = next(
            (line.partition(":")[2].strip()
             for line in session.get("wtr_new_notes", "").splitlines()
             if line.strip().lower().startswith("memory hook:")),
            "",
        )
        if mh:
            copy_notes = f"MEMORY HOOK: {mh['text']}\n" + copy_notes
        elif hook:
            copy_notes += f"\nMEMORY HOOK: {hook}"
        marks["nm_copy"] = len(d.slides)
        d.add(tidx, fill_nm_copy, notes=copy_notes)

    if want("grid") and nm:
        def fill_wtr_new_hdr(root):
            sp = find_sp(root, SHAPE["hdr_wtr_new_txt"])
            replace_run_text(sp, "New Morphology", section)
        d.add(T["hdr_wtr_new"], fill_wtr_new_hdr, notes=HEADER_NOTES["wtr_new"])

        words = session["words_to_read_new"]  # list of {word, meaning}
        if v17:
            if len(words) not in (6, 9, 12):
                warn(f"{day}: words_to_read_new has {len(words)} words (6, 9 or 12 - "
                     "the three columns must balance)")
        elif not 8 <= len(words) <= 12:
            warn(f"{day}: words_to_read_new has {len(words)} words (aim 9-12, min 8)")

        def fill_wtr_grid(root):
            # always the template's 3-column grid; with 10-12 words the columns
            # widen and stretch down the page so a 4th row still reads big
            n = len(words)
            col_spids = SHAPE["wtr_cols"]
            per = math.ceil(n / 3)
            cols = [words[0:per], words[per:2 * per], words[2 * per:]]
            box_y, box_h = (1.15, 3.95) if n > 9 else (1.48, 3.30)
            col_w = 3.15
            for spid, x in zip(col_spids, (0.25, 3.425, 6.60)):
                sp = find_sp(root, spid)
                xfrm = sp.find(q("p:spPr") + "/" + q("a:xfrm"))
                xfrm.find(q("a:off")).set("x", str(int(x * 914400)))
                xfrm.find(q("a:off")).set("y", str(int(box_y * 914400)))
                xfrm.find(q("a:ext")).set("cx", str(int(col_w * 914400)))
                xfrm.find(q("a:ext")).set("cy", str(int(box_h * 914400)))
                bodypr = sp.find(q("p:txBody") + "/" + q("a:bodyPr"))
                if bodypr is not None:
                    bodypr.set("anchor", "ctr")
                    bodypr.set("lIns", "0")
                    bodypr.set("rIns", "0")
            longest = max(len(w["word"]) for w in words)
            rows_max = max(len(c) for c in cols)
            # Classroom projection needs a readable floor. Use a conservative
            # Lexend estimate, but do not shrink the whole grid to tiny text.
            sz = fit_font("x" * longest, col_w, 50 if rows_max <= 3 else 44,
                          min_pt=26, char_em=0.67)
            sz = min(sz, int(box_h / rows_max * 72 / 1.5))
            if sz < 27:
                warn(f"{day}: new morphology word grid would render at {sz} pt; "
                     "replace/rebalance long words so the grid is at least 27 pt")
            targets = []
            for spid, col in zip(col_spids, cols):
                sp = find_sp(root, spid)
                set_runs(sp, [[(w["word"], None, None, None)] for w in col] or [[]],
                         size_pt=sz)
                for pi in range(len(col)):
                    targets.append((spid, pi))
            regen_timing(root, targets)
        grid_notes = session.get(
            "wtr_new_notes",
            "\n".join(f"{w['word']} - {w['meaning']}" for w in words))
        if v18:
            if re.search(r"^memory hook:", grid_notes, flags=re.I | re.M):
                warn(f"{day}: the grid notes carry an old `Memory hook:` line; v1.8 "
                     "puts the hook in `memory_hook` (3a)")
            if mh:
                grid_notes = f"MEMORY HOOK: {mh['text']}\n" + grid_notes
        d.add(T["wtr_new_grid"], fill_wtr_grid, notes=grid_notes)

    # --- v1.7 Monday close: the student's own keyword and one sentence
    if want("keyword_sentence") and session.get("keyword_sentence"):
        ks = session["keyword_sentence"]
        ks_tidx = T["gradual_release"]["you_do"]

        def fill_ks(root, ks=ks, ks_tidx=ks_tidx):
            tsp = find_sp(root, SHAPE["gr_title"][ks_tidx])
            set_text(tsp, "You do - " + ks.get("title", ""))
            bsp = find_sp(root, SHAPE["gr_body"][ks_tidx])
            set_runs(bsp, [[]])
            draw_structured_slide(root, ks, MORPH_MAGENTA, MORPH_TINT,
                                  f"{day} keyword_sentence")
        d.add(ks_tidx, fill_ks, notes=ks.get("notes", ""))

    if want("wts_new") and session.get("words_to_spell_new"):
        def fill_wts_new_hdr(root):
            sp = find_sp(root, SHAPE["hdr_wts_new_txt"])
            replace_run_text(sp, "New Morphology", section)
        d.add(T["hdr_wts_new"], fill_wts_new_hdr, notes=HEADER_NOTES["wts_new"])

        spell = session["words_to_spell_new"]
        if len(spell) != 4:
            warn(f"{day}: words_to_spell_new has {len(spell)} words (template holds 4)")
        if stype == "spell":
            grid = week_teach_grid(week)
            for w in spell:
                if w.lower() not in grid:
                    warn(f"{day}: words_to_spell_new {w!r} is not on this week's "
                         "teach-day grid - students spell what they read on Monday (3c)")

        def fill_wts_grid(root):
            pairs = [spell[0:2], spell[2:4]]
            longest = max((len(w) for w in spell), default=1)
            sz = fit_font("x" * longest, 4.8, 52, min_pt=28)
            targets = []
            for spid, pair in zip(SHAPE["wts_boxes"], pairs):
                sp = find_sp(root, spid)
                set_runs(sp, [[(w, None, None, None)] for w in pair] or [[]], size_pt=sz)
                for pi in range(len(pair)):
                    targets.append((spid, pi))
            regen_timing(root, targets)
        d.add(T["wts_new_grid"], fill_wts_grid,
              notes=session.get("extension",
                                "Once finished, students write a fun sentence with one or more of the words."))

    # --- Yoshimoto You Do activity (legacy new/review days)
    if want("you_do"):
        act = session.get("new_morph_activity")
        if not act:
            warn(f"{day}: no new_morph_activity (Yoshimoto You Do) provided")
        else:
            def fill_act_hdr(root, section=section):
                sp = find_sp(root, SHAPE["hdr_wts_new_txt"])
                replace_run_text(sp, "Words to Spell- New Morphology",
                                 f"{section} - You Do")
                replace_run_text(sp, "2 mins", act.get("time", "3 mins"))
            d.add(T["hdr_wts_new"], fill_act_hdr,
                  notes=HEADER_NOTES["morph_activity"])

            act_tidx = T["gradual_release"]["you_do"]

            def fill_act(root, act=act, act_tidx=act_tidx):
                tsp = find_sp(root, SHAPE["gr_title"][act_tidx])
                set_text(tsp, "You do - " + act.get("title", ""))
                bsp = find_sp(root, SHAPE["gr_body"][act_tidx])
                set_runs(bsp, [[]])
                draw_structured_slide(root, act, MORPH_MAGENTA, MORPH_TINT,
                                      f"{day} new_morph_activity")
            if act.get("round"):
                # v1.6: one item per slide, each answer revealed on its own
                # click - the reveal replaces the separate check slide. The
                # optional lead slide carries the "One done for you" model.
                if act.get("example"):
                    d.add(act_tidx, fill_act, notes=act.get("notes", ""))
                emit_round(d, act, act_tidx, MORPH_MAGENTA, MORPH_TINT,
                           f"{day} new_morph_activity", "You do - ")
            else:
                d.add(act_tidx, fill_act, notes=act.get("notes", ""))

                added_check = add_answer_check_slide(
                    d, act, act_tidx, MORPH_MAGENTA, MORPH_TINT,
                    f"{day} new_morph_activity answer check")
                if has_fixed_answer_notes(act) and not added_check:
                    warn(f"{day}: fixed-answer new_morph_activity needs a duplicate "
                         "check_items answer slide")

    # --- v1.7 Wednesday: build and break words
    if want("build") and session.get("build"):
        bd = session["build"]
        b_tidx = T["gradual_release"]["you_do"]
        mx = bd.get("matrix")
        base = (mx or {}).get("base") or {}

        # a combined day (11) has no Break it apart round: its header names the
        # rounds that actually run (older break rounds carry no `parts`, so
        # round_format cannot tell them apart - other days keep the fixed text)
        hdr_title, hdr_notes = "Build and Break", HEADER_NOTES["build"]
        if session.get("type") == "combined":
            hdr_title = "Build words"
            steps = [BUILD_ROUND_STEPS[round_format(r)] for r in bd.get("rounds", [])]
            hdr_notes = "\n".join(
                ["PROCEDURE:",
                 "1. BUILD real words from the matrix in OG books, then say what each means."]
                + [f"{n}. {st}" for n, st in enumerate(steps, start=2)])

        def fill_build_hdr(root, bd=bd, hdr_title=hdr_title):
            sp = find_sp(root, SHAPE["hdr_wts_new_txt"])
            replace_run_text(sp, "Words to Spell- New Morphology", hdr_title)
            replace_run_text(sp, "2 mins", bd.get("time", "15 mins"))
        d.add(T["hdr_wts_new"], fill_build_hdr, notes=hdr_notes)
        if mx:
            if base:
                resolve_card(day, "matrix base", base)

            def fill_matrix(root, mx=mx, base=base):
                tsp = find_sp(root, SHAPE["gr_title"][b_tidx])
                set_text(tsp, "You do - " + mx.get("title", ""))
                bsp = find_sp(root, SHAPE["gr_body"][b_tidx])
                set_runs(bsp, [[]])
                draw_matrix_slide(root, mx, base, f"{day} build matrix")
            d.add(b_tidx, fill_matrix, notes=mx.get("notes", ""))
        for rnd in bd.get("rounds", []):
            emit_round(d, rnd, b_tidx, MORPH_MAGENTA, MORPH_TINT,
                       f"{day} build round", "You do - ")

    # --- v1.7 Thursday: suffix cards with their part of speech on the click
    # (after the daily review and any build round, before learned words)
    if want("pos_drill") and session.get("pos_drill"):
        drill = session["pos_drill"]

        def fill_pos_hdr(root, drill=drill):
            retitle(root, "Morphology Review", "Suffix Parts of Speech")
            if drill.get("time"):
                retitle(root, "2 mins", drill["time"])
        d.add(T["hdr_morph_review"], fill_pos_hdr,
              notes=HEADER_NOTES["pos_drill_v18" if v18 else "pos_drill"])
        intro = drill.get("intro")
        if intro:
            itidx = T["gradual_release"]["i_do"]

            if v18:
                # James Hooke (23 Oct 2026): the frames slide put the answer on
                # the face and never showed the word.  The v1.8 I do is one
                # modelled sentence, the three jobs as choices, and the answer
                # with its reason behind one click (3g).
                for field in ("prompt", "choices", "answer", "why"):
                    if not intro.get(field):
                        warn(f"{day}: parts-of-speech I do needs `{field}` - "
                             "sentence, three job choices, answer on the click (3g)")
                if intro.get("example") or intro.get("items"):
                    warn(f"{day}: parts-of-speech I do must not print a worked "
                         "answer on the face - put it in `answer`/`why` (3g)")
            pos_timing = etree.fromstring(
                d.pkg.slide_xml(T["spell_word"])).find(q("p:timing"))

            def fill_pos_intro(root, intro=intro, itidx=itidx, pos_timing=pos_timing):
                tsp = find_sp(root, SHAPE["gr_title"][itidx])
                set_text(tsp, "I do - " + intro.get("title", ""))
                bsp = find_sp(root, SHAPE["gr_body"][itidx])
                set_runs(bsp, [[]])
                if v18:
                    answer_sid = draw_round_slide(root, intro, intro, MORPH_MAGENTA,
                                                  MORPH_TINT, f"{day} pos_drill intro")
                    insert_timing(root, copy.deepcopy(pos_timing))
                    regen_timing(root, [(answer_sid, None)])
                else:
                    draw_structured_slide(root, intro, MORPH_MAGENTA, MORPH_TINT,
                                          f"{day} pos_drill intro")
            d.add(itidx, fill_pos_intro, notes=intro.get("notes", ""))
        for card in drill.get("cards", []):
            card = {**card, "type": "suffix"}
            tidx = CARD_REVIEW["suffix"]
            printed = card_parts_of_speech(card)
            pos = card.get("pos_answer") or (printed[0] if printed else "")
            if not pos:
                warn(f"{day}: parts-of-speech card {card['morph']!r} has no printed "
                     "part of speech and no pos_answer")
            elif printed and pos not in printed:
                warn(f"{day}: parts-of-speech card {card['morph']!r} pos_answer "
                     f"{pos!r} is not printed on the card ({', '.join(printed)})")

            kw, mean = resolve_card(day, "parts-of-speech card", card)
            if v18 and not card.get("sentence"):
                warn(f"{day}: parts-of-speech card {card['morph']!r} needs a "
                     "`sentence` using its keyword - students forget older "
                     "suffixes (3g, v1.8)")
            elif v18 and kw and not re.search(rf"\b{re.escape(kw)}\b",
                                              card["sentence"], re.I):
                # the slide bolds the keyword and asks "What job does <keyword>
                # do?" - a sentence without it asks about a word that is not there
                warn(f"{day}: parts-of-speech card {card['morph']!r} sentence does "
                     f"not use the card keyword {kw!r} (3g)")
            if v18 and pos not in POS_JOBS:
                warn(f"{day}: parts-of-speech card {card['morph']!r} makes a "
                     f"{pos!r}; the card choices are only noun, verb and "
                     "adjective (3g, v1.8)")
            if v18 and not card.get("why"):
                warn(f"{day}: parts-of-speech card {card['morph']!r} needs a `why` "
                     "for the reveal - what the word does in the sentence (3g)")

            def fill_pos(root, card=card, tidx=tidx, pos=pos, kw=kw, mean=mean):
                set_slide_background(root, MORPH_FILL["suffix"])
                sp = find_sp(root, SHAPE["card_review_word"][tidx])
                set_text(sp, card["morph"], size_pt=fit_font(card["morph"], 5.75, 56))
                if v18:
                    # James Hooke (23 Oct 2026): the card reads top to bottom -
                    # suffix, its meaning, the keyword in a sentence, the
                    # question, the three jobs - and the answer with its reason
                    # waits behind the click (3g).  The frames are retired.
                    xfrm = sp.find(q("p:spPr") + "/" + q("a:xfrm"))
                    xfrm.find(q("a:off")).set("y", str(int(0.3 * 914400)))
                    meaning = re.sub(r"^[^:]*:\s*", "",
                                     card_meaning_note(mean).splitlines()[0])
                    add_textbox(root, 90051, 1.0, 1.42, 8.0, 0.5,
                                [f"{card['morph']} means: {meaning}"], 22, bold=True)
                    sentence = card.get("sentence", "")
                    marked = re.sub(rf"\b({re.escape(kw)})\b", r"**\1**", sentence,
                                    count=1, flags=re.I)
                    ssz = fit_font_block(sentence, 8.0, 0.6, 26, min_pt=18)
                    add_textbox(root, 90052, 1.0, 2.0, 8.0, 0.7, [marked], ssz,
                                fill="FFFFFF", rounded=True)
                    question = f"What job does {kw} do in this sentence?"
                    add_textbox(root, 90056, 1.0, 2.78, 8.0, 0.45, [question],
                                fit_font(question, 7.8, 22, min_pt=16, char_em=0.58),
                                bold=True)
                    for i, (job, what) in enumerate(POS_JOBS.items()):
                        add_textbox(root, 90053 + i, 1.0 + i * 2.75, 3.3, 2.5, 0.8,
                                    [f"**{job}**", what], 17, fill="FFFFFF",
                                    outline=CHARCOAL, rounded=True)
                    reveal = f"**{pos}** - {card.get('why', '')}"
                    add_textbox(root, 90050, 1.0, 4.22, 8.0, 0.68, [reveal],
                                fit_font_block(reveal.replace("**", ""), 7.8, 0.58, 22,
                                               min_pt=15, char_em=0.58),
                                color="FFFFFF", fill=ANSWER_GREEN, rounded=True)
                else:
                    add_textbox(root, 90050, 2.5, 4.05, 5.0, 0.75, [pos], 36,
                                color=CHARCOAL, bold=True, fill="FFFFFF", rounded=True)
                insert_timing(root, copy.deepcopy(timing_src))
                regen_timing(root, [(90050, None)])
            pos_note = (f"ANSWER: {pos}\n"
                        f"Suffix: {card['morph']}\n"
                        f"{card_meaning_note(mean)}"
                        f"{card_pos_note(card)}")
            if v18:
                pos_note = f"ANSWER: {pos} - {card.get('why', '')}"
            if card.get("notes"):
                pos_note += "\n" + card["notes"]
            marks["pos_cards"].append((len(d.slides), card))
            d.add(tidx, fill_pos, notes=pos_note)

    # --- learned words
    if want("lw") and session.get("learned_words"):
        lw = session["learned_words"]
        d.add(T["hdr_review_lw"],
              notes=HEADER_NOTES["learned_review_v18" if v18 else "learned_review"])
        for item in lw.get("review", []):
            def fill_lw(root, item=item):
                sp = find_sp(root, SHAPE["lw_word"][22])
                sz = fit_font(item["word"], 7.78, 67, min_pt=32)
                set_runs(sp, lw_runs(item["word"], item.get("unfair", "")), size_pt=sz)
            d.add(T["learned_word"], fill_lw, notes=item["notes"])
        if lw.get("new"):
            item = lw["new"]
            d.add(T["hdr_new_lw"],
                  notes=HEADER_NOTES["learned_new_v18" if v18 else "learned_new"])

            def fill_lw_new(root, item=item):
                sp = find_sp(root, SHAPE["lw_word"][25])
                sz = fit_font(item["word"], 7.78, 67, min_pt=32)
                set_runs(sp, lw_runs(item["word"], item.get("unfair", "")), size_pt=sz)
            d.add(T["learned_word_new"], fill_lw_new, notes=item["notes"])

    # --- dictation
    # v1.8 short week: Thursday's one dictation is written in the week's
    # grammar focus, so it runs after the grammar block and the vote.
    late_dictation = v18 and plan_key == "apply_v18"

    def emit_dictation():
        if want("dictation") and session.get("dictation"):
            d.add(T["hdr_dictation"], notes=HEADER_NOTES["dictation"])
            dictation_items = session["dictation"]
            if len(dictation_items) == 2:
                meters = [item.get("meter") for item in dictation_items]
                if meters != ["green", "yellow"]:
                    warn(f"{day}: dictation meters must progress from green on the "
                         "first sentence to yellow on the trickier second sentence")
            for dictation_index, dt in enumerate(dictation_items):
                meter = dt.get("meter")
                if meter not in T["dictation"]:
                    warn(f"{day}: dictation {dictation_index + 1} needs meter "
                         "'green' or 'yellow'")
                    meter = "green" if dictation_index == 0 else "yellow"
                # The locked yellow slide has no native animation. Clone the green
                # animated slide for both meters, then restore a compatible yellow
                # meter image. This preserves a true appear effect without the
                # LibreOffice/Google Slides position shift caused by transplanting a
                # timing block into the yellow slide.
                clone_tidx = T["dictation"]["green"]
                sentence_spid = SHAPE["dictation_sentence"][clone_tidx]
                # COPS replaced CUPS school-wide (Kate Watt, Oct 2026); `cups` is
                # still read so delivered specs rebuild.
                cops = dt.get("cops") or dt.get("cups")
                targets = dt.get("targets", [])
                score = dt.get("score")
                if score is None and cops:
                    score = (len(cops.get("capitals", [])) + len(cops.get("punctuation", []))
                             + len(targets) + 1)  # +1 = Organisation (dictated order)
                if not cops:
                    warn(f"{day}: dictation '{dt['sentence'][:30]}...' has no cops marking block")

                def fill_dict(root, dt=dt, score=score,
                              sentence_spid=sentence_spid):
                    sp = find_sp(root, sentence_spid)
                    # LibreOffice and Google Slides use slightly taller Lexend metrics
                    # than PowerPoint. Keep long 14-16 word reveals inside the locked
                    # sentence box so they cannot expand upward over the title.
                    display_sentence = smart_dictation_quotes(dt["sentence"])
                    sz = fit_font_block(
                        display_sentence, 9.0, 2.45, 32, min_pt=28,
                        char_em=0.62, line_factor=1.5,
                    )
                    set_runs(sp, dictation_runs(display_sentence, dt.get("targets", [])), size_pt=sz)
                    bodypr = sp.find(q("p:txBody") + "/" + q("a:bodyPr"))
                    if bodypr is not None:
                        bodypr.set("anchor", "t")
                    regen_timing(root, [(sentence_spid, None)])
                    if score:
                        add_textbox(root, 90001, 6.8, 4.55, 2.7, 0.45,
                                    [f"Score: ___ /{score}"], 20, align="r", anchor="ctr")

                def fill_dict_rels(rels, meter=meter):
                    if meter != "yellow":
                        return
                    yellow_rels = etree.fromstring(
                        pkg.slide_rels(T["dictation"]["yellow"])
                    )
                    yellow_image_rel = next(
                        rel for rel in yellow_rels.findall(q("rel:Relationship"))
                        if rel.get("Id") == "rId3"
                    )
                    current_image_rel = next(
                        rel for rel in rels.findall(q("rel:Relationship"))
                        if rel.get("Id") == "rId3"
                    )
                    current_image_rel.set("Target", yellow_image_rel.get("Target"))
                lines = [dt["sentence"], ""]
                if cops:
                    caps = cops.get("capitals", [])
                    punct = cops.get("punctuation", [])
                    # Each punctuation mark now carries a plain-English reminder, so the
                    # joined list is far too long for one paragraph. One mark per line
                    # also gives the teacher a genuine tick-off checklist.
                    lines += ["Mark with COPS, tick each line:",
                              f"C - Capitals, shown green: {', '.join(caps)} ({len(caps)})",
                              "O - Organisation: every word there, in the dictated order (1)",
                              f"P - Punctuation, shown red ({len(punct)}):"]
                    lines += [f"   {mark}" for mark in punct]
                    lines += [f"S - Spelling targets, underlined: {', '.join(targets)} ({len(targets)})",
                              f"Score: /{score}"]
                elif targets:
                    lines.append(f"Targets: {', '.join(targets)}")
                if dt.get("focus"):
                    lines.append(f"Focus: {dt['focus']}")
                d.add(clone_tidx, fill_dict, notes="\n".join(lines),
                      fill_rels=fill_dict_rels)

    if not late_dictation:
        emit_dictation()

    # --- grammar
    if want("grammar") and session.get("grammar"):
        gr = session["grammar"]

        def fill_gr_hdr(root):
            pass
        d.add(T["hdr_grammar"], fill_gr_hdr,
              notes=gr.get("header_notes", ""))
        for key, label in [("i_do", "I do - "), ("we_do", "We do - "), ("you_do", "You do - ")]:
            tidx = T["gradual_release"][key]
            blk = gr[key]
            if "lines" in blk:
                warn(f"{day}: grammar {key} uses the legacy lines layout; use the "
                     "structured rule/example/items format so classroom text stays large")

            def fill_gr(root, tidx=tidx, blk=blk, label=label):
                tsp = find_sp(root, SHAPE["gr_title"][tidx])
                set_text(tsp, label + blk.get("title", ""))
                bsp = find_sp(root, SHAPE["gr_body"][tidx])
                if "lines" in blk:  # legacy plain layout
                    set_runs(bsp, [[(ln, None, None, None)] for ln in blk["lines"]] or [[]])
                    return
                set_runs(bsp, [[]])  # clear the body placeholder; we draw instead
                draw_structured_slide(root, blk, GRAMMAR_PURPLE, GRAMMAR_TINT,
                                      f"{day} grammar {key}")
            if blk.get("round"):
                # v1.6: one question per slide, answer on click (section 7).
                if blk.get("example"):
                    d.add(tidx, fill_gr, notes=blk.get("notes", ""))
                emit_round(d, blk, tidx, GRAMMAR_PURPLE, GRAMMAR_TINT,
                           f"{day} grammar {key}", label)
                continue
            d.add(tidx, fill_gr, notes=blk.get("notes", ""))
            if key == "you_do":
                added_check = add_answer_check_slide(
                    d, blk, tidx, GRAMMAR_PURPLE, GRAMMAR_TINT,
                    f"{day} grammar you_do answer check",
                    default_footer="Check your work against each green answer")
                if has_fixed_answer_notes(blk) and not added_check:
                    warn(f"{day}: fixed-answer grammar you_do needs a duplicate "
                         "check_items answer slide")

    # --- v1.9 Friday: the Friday Review worksheet, and the deck stops (1a)
    if want("friday_review") and session.get("friday_review"):
        fr = session["friday_review"]
        fr_dir = out_path.parent / "Friday Review"
        pdf_name = f"Friday Review - Term {term} Week {wknum}.pdf"
        fu_name = f"Friday Review - Term {term} Week {wknum} Teacher Follow-up.pdf"
        report = friday_review_pdf.build_worksheet(
            fr, fr_dir / pdf_name, f"Week {wknum}", fr.get("dictation", []))
        report["follow_up_pages"] = friday_review_pdf.build_follow_up(
            fr.get("follow_up", []), fr_dir / fu_name,
            f"Friday Review - Term {term} Week {wknum}: Teacher Follow-up")
        check_friday(day, week, fr, report)
        marks["friday_review"] = {"worksheet": fr_dir / pdf_name,
                                  "student_pages": report["student_pages"],
                                  "follow_up": fr_dir / fu_name}
        fr_tidx = T["gradual_release"]["you_do"]
        fr_blk = {
            "rule": "Try your best to complete this review.",
            "example": "Books closed. Work on your own.\nRead each question twice, then write.",
            "items": ["A Grammar  -  B Parts of speech  -  C Learned words",
                      "D Morphology  -  E Dictation"],
            "routine": "On your own - 20 mins",
            "footer": "Finished? Check every answer with COPS",
        }
        lines = [f"ANSWER: on the last page of {pdf_name} (teacher copy).",
                 "1. SAY: \"Books closed. This is our Friday Review. Work on your own "
                 "and try every question. You have about twenty minutes.\"",
                 "2. CIRCULATE: read any question aloud for a student who asks. Never "
                 "read the answer words.",
                 "3. READ Part E when most students reach it. Each sentence whole, then "
                 "in chunks, then whole again:"]
        for n, dt in enumerate(fr.get("dictation", []), 1):
            lines.append(f"   {n}. {dt.get('sentence', '')}")
            if dt.get("chunks"):
                lines.append("   CHUNKS: " + " / ".join(dt["chunks"]))
        lines += ["---",
                  "Purpose: a books-closed check of this week only [Friday Review]",
                  f"PREP: print only the student pages of Friday Review/{pdf_name}. "
                  f"Mark with Friday Review/{fu_name} beside you."]

        def fill_friday(root, fr_blk=fr_blk):
            tsp = find_sp(root, SHAPE["gr_title"][fr_tidx])
            set_text(tsp, "Complete this Friday Review")
            bsp = find_sp(root, SHAPE["gr_body"][fr_tidx])
            set_runs(bsp, [[]])
            draw_structured_slide(root, fr_blk, MORPH_MAGENTA, MORPH_TINT,
                                  f"{day} Friday Review")
        d.add(fr_tidx, fill_friday, notes="\n".join(lines))

    # --- v1.7 Friday close: which was trickiest? (feeds "From our reviews")
    if want("vote") and session.get("vote"):
        vt = session["vote"]
        v_tidx = T["gradual_release"]["you_do"]
        options = vt.get("options", [])
        if not 2 <= len(options) <= 4:
            warn(f"{day}: the Friday vote needs 2-4 options (fingers 1-4)")
        vote_blk = {
            "rule": vt.get("rule", "Which was trickiest this week? Hold up 1, 2, 3 or 4."),
            "example": "\n".join(f"{n}. {opt}" for n, opt in enumerate(options, 1)),
            "routine": vt.get("routine", "Fingers 1-4 - voices off"),
            "footer": vt.get("footer", "Next week we bring the trickiest one back"),
        }

        def fill_vote(root, vt=vt, vote_blk=vote_blk):
            tsp = find_sp(root, SHAPE["gr_title"][v_tidx])
            set_text(tsp, vt.get("title", "Trickiest this week?"))
            bsp = find_sp(root, SHAPE["gr_body"][v_tidx])
            set_runs(bsp, [[]])
            draw_structured_slide(root, vote_blk, MORPH_MAGENTA, MORPH_TINT,
                                  f"{day} vote")
        d.add(v_tidx, fill_vote, notes=vt.get("notes", ""))

    if late_dictation:
        emit_dictation()

    d.write(out_path)
    return len(d.slides), marks


# ---------------------------------------------------------------- QA gate
def qa_deck(path, session=None, marks=None):
    """Reopen the deck and enforce OG readability/teacher-note invariants."""
    problems = []
    try:
        from pptx import Presentation
        from pptx.enum.dml import MSO_COLOR_TYPE

        pres = Presentation(path)

        with zipfile.ZipFile(path) as zf:
            slide_names = sorted(
                (name for name in zf.namelist()
                 if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)),
                key=lambda name: int(re.search(r"slide(\d+)", name).group(1)),
            )
            raw_slide_roots = [
                etree.fromstring(zf.read(name)) for name in slide_names
            ]
            raw_slide_rels = []
            for name in slide_names:
                slide_file = name.rsplit("/", 1)[1]
                rel_name = f"ppt/slides/_rels/{slide_file}.rels"
                raw_slide_rels.append(
                    etree.fromstring(zf.read(rel_name))
                    if rel_name in zf.namelist() else None
                )

        def slide_text(slide):
            return "\n".join(
                shape.text for shape in slide.shapes
                if shape.has_text_frame and shape.text
            )

        def rgb(run):
            try:
                if run.font.color.type == MSO_COLOR_TYPE.RGB:
                    return str(run.font.color.rgb)
            except (AttributeError, TypeError):
                pass
            return None

        def fill_rgb(fill):
            try:
                if fill.fore_color.type == MSO_COLOR_TYPE.RGB:
                    return str(fill.fore_color.rgb)
            except (AttributeError, TypeError):
                pass
            return None

        for i, slide in enumerate(pres.slides, 1):
            for sh in slide.shapes:
                if sh.has_text_frame and "XYZ" in sh.text_frame.text:
                    problems.append(f"slide {i}: leftover XYZ placeholder")
            if slide.has_notes_slide and "XYZ" in slide.notes_slide.notes_text_frame.text:
                problems.append(f"slide {i}: leftover XYZ in notes")

        # Notes must be actual PowerPoint paragraphs, not a single iPad text wall.
        with zipfile.ZipFile(path) as zf:
            note_names = sorted(
                (name for name in zf.namelist()
                 if re.fullmatch(r"ppt/notesSlides/notesSlide\d+\.xml", name)),
                key=lambda name: int(re.search(r"notesSlide(\d+)", name).group(1)),
            )
            for note_name in note_names:
                slide_no = int(re.search(r"notesSlide(\d+)", note_name).group(1))
                root = etree.fromstring(zf.read(note_name))
                body = None
                for sp in root.iter(q("p:sp")):
                    ph = sp.find(".//" + q("p:ph"))
                    if ph is not None and ph.get("type") == "body":
                        body = sp
                        break
                if body is None:
                    problems.append(f"slide {slide_no}: notes body placeholder missing")
                    continue
                tb = body.find(q("p:txBody"))
                lines = []
                spelling_word = ""
                spelling_sentence_checked = False
                for p_el in tb.findall(q("a:p")) if tb is not None else []:
                    run_texts = [t_el.text or "" for t_el in p_el.iter(q("a:t"))]
                    line = "".join(run_texts).strip()
                    if not line:
                        continue
                    lines.append(line)
                    if any("\n" in text or "\r" in text or "\x0b" in text
                           for text in run_texts):
                        problems.append(
                            f"slide {slide_no}: note run contains an embedded line break"
                        )
                    if "**" in line:
                        problems.append(f"slide {slide_no}: raw Markdown in notes")
                    if re.search(r"[^\x09\x20-\x7E]", line):
                        problems.append(f"slide {slide_no}: non-ASCII character in notes")
                    if len(line) > 220:
                        problems.append(
                            f"slide {slide_no}: {len(line)}-character note paragraph"
                        )
                    beats = re.findall(
                        r"\b\d+\.\s+(?:SAY|ASK|SCAN|MODEL|POINT|SHOW|READ|"
                        r"REVEAL|CIRCULATE|TIME|WRITE|BUILD|TICK|CHECK|WAIT|"
                        r"REPEAT|CLARIFY)\b",
                        line,
                    )
                    if len(beats) > 1:
                        problems.append(
                            f"slide {slide_no}: multiple numbered beats in one paragraph"
                        )
                    if re.match(r"^\d+\.\s+", line) and (
                            line.count("EXPECT:") > 1 or line.count("?") > 1):
                        problems.append(
                            f"slide {slide_no}: multiple retrieval questions in one paragraph"
                        )
                    ppr = p_el.find(q("a:pPr"))
                    if ppr is None or ppr.find(q("a:buNone")) is None:
                        problems.append(
                            f"slide {slide_no}: notes paragraph may inherit a bullet"
                        )
                    if ppr is None or ppr.get("algn") != "l":
                        problems.append(
                            f"slide {slide_no}: notes paragraph is not left-aligned"
                        )
                    needs_bold_label = bool(
                        NOTE_ANCHOR_RE.search(line) or NUMBERED_ACTION_RE.search(line)
                    )
                    has_bold_run = any(
                        (run.find(q("a:rPr")) is not None
                         and run.find(q("a:rPr")).get("b") == "1")
                        for run in p_el.findall(q("a:r"))
                    )
                    if needs_bold_label and not has_bold_run:
                        problems.append(
                            f"slide {slide_no}: teacher-note label is not PowerPoint bold"
                        )
                    if line.startswith("Word to spell:"):
                        spelling_word = line.split(":", 1)[1].strip()
                    if line.startswith("Sentence:") and spelling_word:
                        spelling_sentence_checked = True
                        has_underlined_word = any(
                            (run.find(q("a:rPr")) is not None
                             and run.find(q("a:rPr")).get("u") == "sng"
                             and "".join(
                                 t_el.text or "" for t_el in run.iter(q("a:t"))
                             ).strip().casefold() == spelling_word.casefold())
                            for run in p_el.findall(q("a:r"))
                        )
                        if not has_underlined_word:
                            problems.append(
                                f"slide {slide_no}: spelling target is not underlined "
                                "inside the Sentence line"
                            )
                if spelling_word and not spelling_sentence_checked:
                    problems.append(
                        f"slide {slide_no}: spelling notes have no Sentence line "
                        "to carry the underlined target"
                    )
                if "---" in lines and lines.index("---") > 8:
                    problems.append(
                        f"slide {slide_no}: notes live zone exceeds 8 Glance Format lines"
                    )
                joined = "\n".join(lines)
                if "Derivative ask:" in joined:
                    problems.append(
                        f"slide {slide_no}: unclear 'Derivative ask' label in notes"
                    )
                if "Meaning: ending that says" in joined:
                    problems.append(
                        f"slide {slide_no}: sound and component meaning are crammed together"
                    )

        if session:
            # The school colour key is global: root green, prefix yellow, suffix red.
            marks = marks or {}
            for slide_idx, card in marks.get("review_cards", []) + marks.get("pos_cards", []):
                if slide_idx >= len(pres.slides):
                    problems.append("morphology review card slide is missing")
                    break
                slide = pres.slides[slide_idx]
                expected = MORPH_FILL[card["type"]]
                actual = fill_rgb(slide.background.fill)
                if actual != expected:
                    problems.append(
                        f"morphology review card '{card['morph']}' uses {actual} "
                        f"instead of {expected} for {card['type']}"
                    )

            sound_bank = session.get("sound_bank", [])
            bank_idx = marks.get("sound_bank")
            if sound_bank and bank_idx is not None:
                sound_shapes = {
                    shape.shape_id: shape for shape in pres.slides[bank_idx].shapes
                }
                for idx, item in enumerate(sound_bank[:9]):
                    shape = sound_shapes.get(SHAPE["sound_boxes"][idx])
                    expected = (SOUND_FILL if item.get("type") == "sound"
                                else MORPH_FILL[item["type"]])
                    actual = fill_rgb(shape.fill) if shape is not None else None
                    if actual != expected:
                        problems.append(
                            f"sound-bank box '{item['morph']}' uses {actual} "
                            f"instead of {expected} for {item['type']}"
                        )

            new_morph = session.get("new_morphology")
            card_idx, copy_idx = marks.get("nm_card"), marks.get("nm_copy")
            if new_morph and card_idx is not None:
                if card_idx >= len(pres.slides):
                    problems.append("new morphology card slide is missing")
                else:
                    expected = MORPH_FILL[new_morph["type"]]
                    actual = fill_rgb(pres.slides[card_idx].background.fill)
                    if actual != expected:
                        problems.append(
                            f"new morphology card '{new_morph['morph']}' uses "
                            f"{actual} instead of {expected} for "
                            f"{new_morph['type']}"
                        )

                if copy_idx is None or copy_idx >= len(pres.slides):
                    problems.append("new morphology copy-card slide is missing")
                else:
                    copy_slide = pres.slides[copy_idx]
                    copy_text = slide_text(copy_slide)
                    entry = verified_card_entry(
                        new_morph["morph"], new_morph.get("type"),
                        new_morph.get("meaning"),
                    )
                    expected_keyword = (
                        entry.get("keyword", "") if entry
                        else new_morph.get("keyword", "")
                    )
                    expected_meaning = (
                        entry.get("meaning", "") if entry
                        else catalogue_meaning(
                            new_morph["morph"], new_morph.get("type")
                        ) or new_morph.get("meaning", "")
                    )
                    expected_lines = [
                        f"{new_morph['type'].capitalize()}: {new_morph['morph']}",
                        f"Keyword: {expected_keyword}",
                        *card_meaning_note(expected_meaning).splitlines(),
                    ]
                    parts = card_parts_of_speech(new_morph)
                    if parts:
                        expected_lines.append(
                            f"Part of speech: {', '.join(parts)}"
                        )
                    for line in expected_lines:
                        if line not in copy_text:
                            problems.append(
                                f"new morphology copy-card slide is missing {line!r}"
                            )
                    actual = fill_rgb(copy_slide.background.fill)
                    expected = MORPH_FILL[new_morph["type"]]
                    if actual != expected:
                        problems.append(
                            "new morphology copy-card slide does not retain the "
                            f"{new_morph['type']} card colour"
                        )
                    if raw_slide_roots[copy_idx].find(q("p:timing")) is not None:
                        problems.append(
                            "new morphology copy-card slide should be visible "
                            "immediately, without a reveal animation"
                        )

            # New-morphology grids must stay large enough for classroom reading.
            new_words = [
                item["word"] if isinstance(item, dict) else str(item)
                for item in session.get("words_to_read_new", [])
            ]
            if new_words:
                grid_slide = next(
                    (slide for slide in pres.slides
                     if all(word in slide_text(slide) for word in new_words)),
                    None,
                )
                if grid_slide is None:
                    problems.append("new morphology word grid not found")
                else:
                    sizes = {}
                    for shape in grid_slide.shapes:
                        if not shape.has_text_frame:
                            continue
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                word = run.text.strip()
                                if word in new_words and run.font.size is not None:
                                    sizes[word] = run.font.size.pt
                    missing = [word for word in new_words if word not in sizes]
                    if missing:
                        problems.append(
                            "new morphology words lack explicit font sizes: "
                            + ", ".join(missing)
                        )
                    elif min(sizes.values()) < 27:
                        problems.append(
                            f"new morphology grid falls to {min(sizes.values()):.0f} pt"
                        )

            # Fixed-answer activities require an immediate green check-and-fix slide.
            for activity in (session.get("new_morph_activity") or {},
                             (session.get("grammar") or {}).get("you_do") or {}):
                check_items = activity.get("check_items", [])
                if not check_items:
                    continue
                answer_idx = next(
                    (idx for idx, slide in enumerate(pres.slides)
                     if all(item in slide_text(slide) for item in check_items)),
                    None,
                )
                if answer_idx is None:
                    problems.append("fixed-answer check slide not found")
                else:
                    if "Tick it or fix it -" not in slide_text(
                            pres.slides[answer_idx]):
                        problems.append(
                            "fixed-answer slide does not use the 'Tick it or fix it' title"
                        )
                    if answer_idx == 0 or "You do -" not in slide_text(
                            pres.slides[answer_idx - 1]):
                        problems.append(
                            "fixed-answer check slide does not immediately follow the task"
                        )
                    coloured_items = set()
                    for shape in pres.slides[answer_idx].shapes:
                        if not shape.has_text_frame:
                            continue
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if (run.text.strip() in check_items
                                        and rgb(run) == ANSWER_GREEN):
                                    coloured_items.add(run.text.strip())
                    if coloured_items != set(check_items):
                        problems.append(
                            "check-and-fix answers are not all shown in green"
                        )

            # v1.6 rounds: each item slide hides its green answer behind one click.
            round_blocks = [session.get("new_morph_activity") or {}] + [
                (session.get("grammar") or {}).get(key) or {}
                for key in ("i_do", "we_do", "you_do")
            ] + list((session.get("build") or {}).get("rounds", []))
            for blk in round_blocks:
                for item in blk.get("round", []):
                    if item.get("worked"):
                        continue
                    first = str(item.get("prompt", "")).strip().splitlines()[0].replace("**", "")
                    answer = str(item.get("answer", "")).strip()
                    idx = next(
                        (n for n, slide in enumerate(pres.slides)
                         if first in slide_text(slide) and answer in slide_text(slide)),
                        None,
                    )
                    if idx is None:
                        problems.append(f"round item slide not found: {first[:32]}")
                        continue
                    bar = next(
                        (shape for shape in pres.slides[idx].shapes
                         if shape.has_text_frame
                         and shape.text.strip().startswith(answer)
                         and fill_rgb(shape.fill) == ANSWER_GREEN),
                        None,
                    )
                    if bar is None:
                        problems.append(f"round item has no green answer bar: {first[:32]}")
                        continue
                    timing = raw_slide_roots[idx].find(q("p:timing"))
                    timed = {
                        target.get("spid")
                        for target in timing.findall(".//" + q("p:spTgt"))
                    } if timing is not None else set()
                    if str(bar.shape_id) not in timed:
                        problems.append(
                            f"round answer is visible before the click: {first[:32]}")

            # Dictation slides use colour/format as COPS information, so verify it.
            punctuation = set(",.;:!?\"'“”-")
            for item in session.get("dictation", []):
                # The slide face carries directional speech marks; the spec is ASCII.
                sentence = smart_dictation_quotes(item["sentence"])
                target_shape = None
                target_slide = None
                target_slide_index = None
                for slide_index, slide in enumerate(pres.slides):
                    for shape in slide.shapes:
                        if shape.has_text_frame and shape.text.strip() == sentence:
                            target_shape = shape
                            target_slide = slide
                            target_slide_index = slide_index
                            break
                    if target_shape is not None:
                        break
                if target_shape is None:
                    problems.append(f"dictation sentence not found: {sentence[:32]}")
                    continue
                expected_meter = item.get("meter")
                meter_shape_id = SHAPE["dictation_meter"][T["dictation"]["green"]]
                if (meter_shape_id is None
                        or meter_shape_id not in {
                            shape.shape_id for shape in target_slide.shapes
                        }):
                    problems.append(
                        f"dictation does not use the requested {expected_meter} meter"
                    )
                meter_pic = next(
                    (pic for pic in raw_slide_roots[target_slide_index].iter(q("p:pic"))
                     if pic.find(q("p:nvPicPr") + "/" + q("p:cNvPr")).get("id")
                     == str(meter_shape_id)),
                    None,
                )
                if meter_pic is not None:
                    blip = meter_pic.find(q("p:blipFill") + "/" + q("a:blip"))
                    rel_id = blip.get(q("r:embed")) if blip is not None else None
                    rels = raw_slide_rels[target_slide_index]
                    image_target = next(
                        (rel.get("Target") for rel in rels.findall(q("rel:Relationship"))
                         if rel.get("Id") == rel_id),
                        "",
                    ) if rels is not None else ""
                    expected_image = (
                        "image13.png" if expected_meter == "green"
                        else "image12.png"
                    )
                    if not image_target.endswith(expected_image):
                        problems.append(
                            f"dictation does not use the requested {expected_meter} meter image"
                        )
                timing = raw_slide_roots[target_slide_index].find(q("p:timing"))
                timed_shapes = {
                    target.get("spid")
                    for target in timing.findall(".//" + q("p:spTgt"))
                } if timing is not None else set()
                if str(target_shape.shape_id) not in timed_shapes:
                    problems.append(
                        f"{expected_meter} dictation sentence has no click-to-reveal "
                        "animation"
                    )
                chars = []
                for paragraph in target_shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        colour = rgb(run)
                        for char in run.text:
                            chars.append((char, colour, bool(run.font.bold),
                                          bool(run.font.underline)))
                rendered = "".join(char for char, _, _, _ in chars)
                if rendered != sentence:
                    problems.append(
                        f"dictation run text changed: {sentence[:32]}"
                    )
                    continue
                for pos, (char, colour, _, _) in enumerate(chars):
                    if char.isupper() and colour != CAPITAL_GREEN:
                        problems.append(
                            f"dictation capital at character {pos + 1} is not green"
                        )
                    if char in punctuation and colour != PUNCT_RED:
                        problems.append(
                            f"dictation punctuation at character {pos + 1} is not red"
                        )
                lower_sentence = sentence.casefold()
                for target in item.get("targets", []):
                    start = lower_sentence.find(str(target).casefold())
                    if start < 0:
                        problems.append(
                            f"dictation target '{target}' is absent from its sentence"
                        )
                        continue
                    end = start + len(target)
                    if not all(chars[pos][2] and chars[pos][3]
                               for pos in range(start, end)):
                        problems.append(
                            f"dictation target '{target}' is not bold and underlined"
                        )

            # Main grammar examples must not regress to the old tiny template text.
            for key in ("i_do", "we_do", "you_do"):
                block = (session.get("grammar") or {}).get(key) or {}
                example = str(block.get("example", "")).strip()
                if not example:
                    continue
                example_shape = next(
                    (shape for slide in pres.slides for shape in slide.shapes
                     if shape.has_text_frame and shape.text.strip() == example),
                    None,
                )
                if example_shape is None:
                    problems.append(f"grammar {key} example not found")
                    continue
                sizes = [
                    run.font.size.pt
                    for paragraph in example_shape.text_frame.paragraphs
                    for run in paragraph.runs if run.font.size is not None
                ]
                if not sizes or min(sizes) < 20:
                    problems.append(
                        f"grammar {key} example is below the 20 pt readability floor"
                    )
    except Exception as e:
        problems.append(f"failed to reopen: {e}")
    return problems


def build_enabler(week, session, number, folder):
    """Write the session's one-page enabler sheet (1b, v1.9) and its matching
    extension sheet in ../Extension/ (1b)."""
    day = session["day"]
    if not enabler_pdf.fonts_available():
        warn(f"{day}: no VIC handwriting font is installed - install VICDOTS_.TTF and "
             "VICREGUL.TTF (School Fonts), or the Department's free Victorian Modern "
             "Cursive (Vicmorg_.ttf, traced in grey), to build enabler sheets (1b)")
        return
    path = folder / enabler_pdf.file_name(number, day)
    try:
        third = enabler_pdf.make(week, session, path)
    except ValueError as exc:
        warn(f"{day}: {exc} (1b)")
        return
    if third < enabler_pdf.SMALL_THIRD_NOTE:
        note(f"{day}: enabler sheet thirds are {third / 72 * 25.4:.1f} mm to fit one page "
             "- trim a section rather than shrink further (1b)")
    print(f"  Enabler: {path.name} (thirds {third / 72 * 25.4:.1f} mm)")
    ext = folder.parent / "Extension" / enabler_pdf.extension_file_name(number, day)
    try:
        if enabler_pdf.make_extension(week, session, ext):
            print(f"  Extension: {ext.name}")
    except ValueError as exc:
        warn(f"{day}: {exc} (1b)")


def main():
    WARNINGS.clear()
    NOTES.clear()
    _SCHOOL_CARDS_NOTED.clear()
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    only = None
    outdir = Path("output")
    for i, a in enumerate(sys.argv[1:]):
        if a == "--only":
            only = sys.argv[1:][i + 1]
        if a == "--outdir":
            outdir = Path(sys.argv[1:][i + 1])
    if not args:
        sys.exit("usage: build_og_week.py <week_spec.json> [--only Day]")
    spec_path = Path(args[0])
    global SPEC_DIR
    SPEC_DIR = spec_path.resolve().parent
    week = deep_sanitize(json.loads(spec_path.read_text()))
    pkg = Package(MASTER)
    unit = week.get("unit_folder", f"OG_Term{week['term']}_Week{week['week']}")
    built = []
    pack_decks, pack_friday = [], {}  # v1.9 print pack (1c)
    failed = False
    morph_idx = 0
    v17 = week_is_v17(week)
    week_morph = next((s["new_morphology"]["morph"] for s in week["sessions"]
                       if (s.get("new_morphology") or {}).get("morph")), "")
    week_morph = week_morph.replace(" / ", "-").replace("/", "-")
    for n_session, session in enumerate(week["sessions"], 1):
        # team filename convention: "1a. Monday (-eer).pptx", "1b. Tuesday
        # (-eer review).pptx", ..., "3. Friday (week review).pptx"
        stype = session.get("type", "new")
        if stype == "new":
            morph_idx += 1
            label = f"{morph_idx}a"
        elif stype == "review":
            label = f"{max(morph_idx, 1)}b"
        else:
            morph_idx += 1
            label = str(morph_idx)
        morph = session.get("new_morphology", {}).get("morph", "")
        morph = morph.replace(" / ", "-").replace("/", "-")
        if stype == "week_review" or not morph:
            paren = "week review"
        else:
            paren = morph + (" review" if stype == "review" else "")
        if v17:
            # v1.7: one morpheme a week, so sessions number by day order -
            # "1. Monday (port).pptx" ... "5. Friday (week review).pptx"
            label = str(n_session)
            paren = {
                "teach": week_morph,
                "spell": f"{week_morph} spelling",
                "build": f"{week_morph} building",
                "apply": "parts of speech",
                "combined": f"{week_morph} building and grammar",
            }.get(stype, "week review")
        fname = session.get("file_name", f"{label}. {session['day']} ({paren}).pptx")
        if only and session["day"].lower() != only.lower():
            continue
        out_path = outdir / unit / fname
        print(f"Building {out_path} ...")
        n, marks = build_session(pkg, week, session, out_path)
        if week_is_v19(week) and enabler_pdf.wants_sheet(session):
            build_enabler(week, session, n_session, out_path.parent / "Enablers")
        pack_decks.append((session, out_path.name))
        if marks.get("friday_review"):
            pack_friday = marks["friday_review"]
        problems = qa_deck(out_path, session, marks)
        for p in problems:
            print(f"  GATE FAIL: {p}")
        if problems:
            failed = True
        print(f"  {n} slides written.")
        built.append(out_path)
    if week_is_v19(week) and built and not only:
        pack, pages = print_pack.build_pack(week, built[0].parent, pack_decks, pack_friday)
        print(f"\nPrint pack: {pack.name} ({pages} pages)")
    if NOTES:
        print(f"\n{len(NOTES)} advisory note(s) - review above (non-fatal).")
    if WARNINGS:
        print(f"\n{len(WARNINGS)} warning(s) - review above.")
        failed = True
    if failed:
        sys.exit("QA gate failed.")
    print("\nDone:")
    for b in built:
        print(f"  {b}")


if __name__ == "__main__":
    main()
