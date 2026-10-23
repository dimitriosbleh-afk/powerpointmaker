"""Regression gate for the locked-template Orton-Gillingham deck builder.

Run from the repository root:
    python tests/test_og_builder_regressions.py
"""

import importlib.util
import json
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
BUILDER_PATH = ROOT / "og_planner" / "build_og_week.py"
# Delivered before the Aug 2026 dictation-sourcing rule (section 6): its
# dictations target that week's own new words, so it must now FAIL the gate.
RETIRED_SPEC = ROOT / "og_planner" / "weeks" / "term3_week1.json"
SAMPLE_SPEC = ROOT / "og_planner" / "weeks" / "sample_term3_week1.json"

MODULE_SPEC = importlib.util.spec_from_file_location("og_builder", BUILDER_PATH)
OG = importlib.util.module_from_spec(MODULE_SPEC)
MODULE_SPEC.loader.exec_module(OG)


class OgBuilderRegressionTests(unittest.TestCase):
    maxDiff = None

    def build_spec(self, spec_path, output_root):
        result = subprocess.run(
            [
                sys.executable,
                str(BUILDER_PATH),
                str(spec_path),
                "--outdir",
                str(output_root),
            ],
            cwd=ROOT,
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        self.assertEqual(
            result.returncode,
            0,
            msg=(
                f"OG build failed for {spec_path.name}.\n"
                f"STDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}"
            ),
        )
        self.assertNotIn("WARN:", result.stdout)
        return sorted(Path(output_root).rglob("*.pptx"))

    def test_sample_spec_passes_every_output_gate(self):
        with tempfile.TemporaryDirectory(prefix="og_builder_regression_") as temp:
            temp_root = Path(temp)
            sample = self.build_spec(SAMPLE_SPEC, temp_root / "sample")
            self.assertEqual(len(sample), 4)
            self.assert_deck_structure(SAMPLE_SPEC, sample)

    def test_retired_spec_trips_the_dictation_sourcing_gate(self):
        """Pre-rule decks dictated the week's own new words; the gate must say so."""
        week = json.loads(RETIRED_SPEC.read_text(encoding="utf-8"))
        OG.WARNINGS.clear()
        for session in week["sessions"]:
            OG.validate_dictation_sources(session["day"], week, session)
        self.assertTrue(
            any("THIS week's new words" in w for w in OG.WARNINGS), OG.WARNINGS)

    def assert_deck_structure(self, spec_path, decks):
        week = json.loads(spec_path.read_text(encoding="utf-8"))
        self.assertEqual(len(week["sessions"]), len(decks))
        for session, deck in zip(week["sessions"], decks):
            with zipfile.ZipFile(deck) as archive:
                slide_names = sorted(
                    (name for name in archive.namelist()
                     if name.startswith("ppt/slides/slide")
                     and name.endswith(".xml")),
                    key=lambda name: int(Path(name).stem.removeprefix("slide")),
                )
                roots = [etree.fromstring(archive.read(name))
                         for name in slide_names]

            for dictation in session["dictation"]:
                # The slide face carries directional speech marks; specs stay ASCII.
                display_sentence = OG.smart_dictation_quotes(dictation["sentence"])
                sentence_root = next(
                    root for root in roots
                    if display_sentence in "".join(
                        text.text or "" for text in root.iter(OG.q("a:t"))
                    )
                )
                timing = sentence_root.find(OG.q("p:timing"))
                self.assertIsNotNone(timing)
                self.assertTrue(timing.findall(".//" + OG.q("p:spTgt")))
                sentence_shape = next(
                    shape for shape in sentence_root.iter(OG.q("p:sp"))
                    if display_sentence in "".join(
                        text.text or "" for text in shape.iter(OG.q("a:t"))
                    )
                )
                bodypr = sentence_shape.find(
                    OG.q("p:txBody") + "/" + OG.q("a:bodyPr")
                )
                self.assertEqual(bodypr.get("anchor"), "t")
                sizes = [
                    int(run.get("sz"))
                    for run in sentence_shape.iter(OG.q("a:rPr"))
                    if run.get("sz")
                ]
                self.assertTrue(sizes)
                self.assertLessEqual(max(sizes), 3200)

            new_morph = session.get("new_morphology")
            if not new_morph:
                continue
            self.assertGreater(len(roots), 30)
            detail_root = roots[30]
            detail_text = "\n".join(
                text.text or "" for text in detail_root.iter(OG.q("a:t"))
            )
            self.assertIn(
                f"{new_morph['type'].capitalize()}: ", detail_text
            )
            self.assertIn(new_morph["morph"], detail_text)
            self.assertIn("Keyword: ", detail_text)
            self.assertIsNone(detail_root.find(OG.q("p:timing")))

    def test_sound_and_component_meaning_are_split(self):
        note = OG.card_meaning_note(
            "ending that says /sh-ee-ate/ (the -ate part means to make or do)"
        )
        self.assertEqual(
            note,
            "Sound: /sh-ee-ate/\nMeaning of -ate: to make or do",
        )

    def test_morpheme_colour_key_is_locked(self):
        self.assertEqual(
            OG.MORPH_FILL,
            {"root": "ABEA90", "prefix": "FFF2CC", "suffix": "E06666"},
        )

    def test_yellow_dictation_meter_uses_compatible_asset(self):
        package = OG.Package(OG.MASTER)
        yellow = Image.open(BytesIO(package.parts["ppt/media/image12.png"]))
        green = Image.open(BytesIO(package.parts["ppt/media/image13.png"]))
        self.assertEqual(yellow.size, green.size)
        self.assertTrue(
            any(red > 240 and green_channel > 160 and blue < 40 and alpha
                for red, green_channel, blue, alpha in yellow.convert("RGBA").getdata())
        )

    def test_markdown_and_question_piles_are_rejected(self):
        OG.WARNINGS.clear()
        markdown = "**ANSWER:** shut\n1. SAY: Read it.\n---\nPurpose: test."
        OG.validate_notes_source(31, markdown, OG.prepare_notes_text(markdown))
        self.assertTrue(any("markdown" in warning for warning in OG.WARNINGS))

        OG.WARNINGS.clear()
        piled = (
            "1. ASK: Opposite of include? EXPECT: exclude. "
            "Make pedestrian plural? EXPECT: pedestrians."
        )
        OG.validate_notes_source(8, piled, OG.prepare_notes_text(piled))
        self.assertTrue(
            any("multiple questions or answers" in warning
                for warning in OG.WARNINGS)
        )

    def test_legacy_derivative_ask_is_rejected(self):
        OG.WARNINGS.clear()
        rendered = OG.extra_task_note(
            {"derivative_ask": "Give one word. EXPECT: flexible"}
        )
        self.assertIn("Extra task:", rendered)
        self.assertTrue(
            any("legacy derivative_ask" in warning for warning in OG.WARNINGS)
        )

    def test_spelling_sentence_target_uses_real_underline(self):
        word_line = OG.note_runs("Word to spell: preordained")
        self.assertIn(("Word to spell:", True, False), word_line)
        self.assertFalse(any(underlined for _, _, underlined in word_line))

        sentence_runs = OG.note_runs(
            "Sentence: The ending seemed preordained from the first chapter.",
            underline_terms=("preordained",),
        )
        self.assertIn(("Sentence:", True, False), sentence_runs)
        self.assertIn(("preordained", False, True), sentence_runs)

    # ------------------------------------------------ 2c sound bank derivation
    def test_sound_bank_rejects_the_days_focus_morpheme(self):
        """The junct defect: today's new morpheme banked before it is taught."""
        session = {
            "new_morphology": {"morph": "junct/join/joint", "type": "root"},
            "sound_bank": [{"morph": "junct/join/joint", "type": "root"},
                           {"morph": "micro", "type": "root"}],
            "words_to_spell_review": [{"word": "micron"}],
        }
        OG.WARNINGS.clear()
        OG.validate_sound_bank("Wednesday", session)
        self.assertTrue(
            any("focus morpheme" in warning for warning in OG.WARNINGS),
            OG.WARNINGS,
        )

    def test_sound_bank_focus_check_matches_a_single_variant(self):
        """A bank box need only match one variant of the focus label."""
        session = {
            "new_morphology": {"morph": "pict/picto", "type": "root"},
            "sound_bank": [{"morph": "picto", "type": "root"}],
            "words_to_spell_review": [{"word": "pictogram"}],
        }
        OG.WARNINGS.clear()
        OG.validate_sound_bank("Wednesday", session)
        self.assertTrue(any("focus morpheme" in w for w in OG.WARNINGS), OG.WARNINGS)

    def test_unused_sound_bank_box_is_advisory_not_fatal(self):
        """A substring heuristic must not block an otherwise valid build."""
        session = {
            "sound_bank": [{"morph": "-ly", "type": "suffix"}],
            "words_to_spell_review": [{"word": "micron"}, {"word": "carnage"}],
        }
        OG.WARNINGS.clear()
        OG.NOTES.clear()
        OG.validate_sound_bank("Wednesday", session)
        self.assertEqual(OG.WARNINGS, [])
        self.assertTrue(any("appears in none of the" in n for n in OG.NOTES), OG.NOTES)

    def test_sound_bank_rejects_a_morpheme_taught_later_this_week(self):
        """Steph's report (Aug 2026): the new morphology sat in the bank before
        it was taught. Monday's bank/cards may not carry Wednesday's morpheme."""
        monday = {
            "day": "Monday", "type": "new",
            "new_morphology": {"morph": "eco", "type": "root"},
            "morphology_review": [{"morph": "chron", "type": "root"}],
            "sound_bank": [{"morph": "vict/vinc", "type": "root"}],
            "words_to_spell_review": [{"word": "victory"}],
        }
        tuesday = {"day": "Tuesday", "type": "review",
                   "new_morphology": {"morph": "eco", "type": "root"},
                   "morphology_review": [{"morph": "eco", "type": "root"}],
                   "sound_bank": []}
        wednesday = {"day": "Wednesday", "type": "new",
                     "new_morphology": {"morph": "vict/vinc", "type": "root"}}
        thursday = {"day": "Thursday", "type": "new",
                    "new_morphology": {"morph": "chron", "type": "root"}}
        week = {"sessions": [monday, tuesday, wednesday, thursday]}
        OG.WARNINGS.clear()
        OG.validate_sound_bank("Monday", monday, week)
        self.assertEqual(
            len([w for w in OG.WARNINGS if "later session" in w]), 2, OG.WARNINGS)
        # A review day repeating yesterday's focus on its cards is fine.
        OG.WARNINGS.clear()
        OG.validate_sound_bank("Tuesday", tuesday, week)
        self.assertEqual(OG.WARNINGS, [])

    def test_review_cards_reject_todays_new_morpheme(self):
        session = {
            "day": "Monday", "type": "new",
            "new_morphology": {"morph": "eco", "type": "root"},
            "morphology_review": [{"morph": "eco", "type": "root"}],
            "sound_bank": [],
        }
        OG.WARNINGS.clear()
        OG.validate_sound_bank("Monday", session, {"sessions": [session]})
        self.assertTrue(any("today's NEW morpheme" in w for w in OG.WARNINGS), OG.WARNINGS)

    def _dictation_week(self, targets):
        return {"term": 9, "week": 9, "sessions": [{
            "day": "Monday", "type": "new",
            "new_morphology": {"morph": "chron", "type": "root"},
            "words_to_read_new": [{"word": "chronicle"}, {"word": "chronic"}],
            "words_to_spell_new": ["chronicle"],
            "learned_words": {"new": {"word": "leopard"},
                              "review": [{"word": "iron"}]},
            "dictation": [{"meter": "green", "sentence": "x", "targets": targets}],
        }]}

    def test_dictation_rejects_this_weeks_words_and_focus_morpheme(self):
        """Steph's report (Aug 2026): students copy this week's words from their
        books, so dictation targets come from 2-3 weeks ago."""
        week = self._dictation_week(["chronicles", "leopard", "chronometer", "captain"])
        OG.WARNINGS.clear()
        OG.NOTES.clear()
        OG.validate_dictation_sources("Monday", week, week["sessions"][0])
        self.assertEqual(len([w for w in OG.WARNINGS if "THIS week's new words" in w]), 2,
                         OG.WARNINGS)
        self.assertEqual(len([w for w in OG.WARNINGS if "focus morpheme" in w]), 1,
                         OG.WARNINGS)
        self.assertFalse(any("captain" in w for w in OG.WARNINGS + OG.NOTES))

    def test_dictation_review_learned_word_is_advisory(self):
        week = self._dictation_week(["iron"])
        OG.WARNINGS.clear()
        OG.NOTES.clear()
        OG.validate_dictation_sources("Monday", week, week["sessions"][0])
        self.assertEqual(OG.WARNINGS, [])
        self.assertTrue(any("review pair" in n for n in OG.NOTES), OG.NOTES)

    def test_silent_e_stem_does_not_collapse_short_affixes(self):
        """vore -> vor is wanted; -ine -> in would match half of English."""
        self.assertIn("vor", OG.morph_surface_forms("vore"))
        self.assertNotIn("in", OG.morph_surface_forms("-ine"))
        self.assertNotIn("ic", OG.morph_surface_forms("-ice"))
        self.assertEqual(OG.morph_surface_forms("-tion / -sion"), {"tion", "sion"})

    def test_vore_box_counts_herbivorous_as_served(self):
        session = {
            "sound_bank": [{"morph": "vore", "type": "root"}],
            "words_to_spell_review": [{"word": "herbivorous"}],
        }
        OG.WARNINGS.clear()
        OG.NOTES.clear()
        OG.validate_sound_bank("Friday", session)
        self.assertEqual(OG.NOTES, [])

    # ------------------------------------------------ 2f/3d You Do gates
    GLOSSED_RULE = "Parts: dis- (apart), -tion (the act of), -ly (in that way)"

    def _junct_week(self, items, rule="", example="", check_items=()):
        return {
            "taught_morphemes": [],
            "sessions": [{
                "day": "Wednesday",
                "new_morphology": {"morph": "junct/join/joint", "type": "root"},
                "morphology_review": [{"morph": "micro", "type": "root"},
                                      {"morph": "-ly", "type": "suffix"}],
                "sound_bank": [{"morph": "dis-", "type": "prefix"},
                               {"morph": "-tion / -sion", "type": "suffix"}],
                "new_morph_activity": {"rule": rule, "example": example,
                                       "items": items,
                                       "check_items": list(check_items)},
            }],
        }

    def _run_gate(self, week):
        """Run the You Do gate on the LAST session; returns (warnings, notes)."""
        OG.WARNINGS.clear()
        OG.NOTES.clear()
        session = week["sessions"][-1]
        OG.validate_activity_morphemes(session["day"], week, session)
        return list(OG.WARNINGS), list(OG.NOTES)

    def test_you_do_flags_an_untaught_prefix(self):
        """in- appeared once in a grid script line: not taught, and not printed."""
        week = self._junct_week([
            "1. junct + -tion = ?    2. dis- + joint + -ed = ?",
            "3. joint + -ly = ?    4. in- + junct + -tion = ?",
        ], rule=self.GLOSSED_RULE)
        warnings, _ = self._run_gate(week)
        flagged = [w for w in warnings if "You Do uses morpheme" in w]
        self.assertEqual(len(flagged), 1, warnings)
        self.assertIn("'in-'", flagged[0])

    def test_you_do_accepts_printed_parts_and_common_inflections(self):
        """dis-, -tion and -ly are glossed on the rule banner; -ed is universal."""
        week = self._junct_week([
            "1. junct + -tion = ?    2. dis- + joint + -ed = ?",
            "3. joint + -ly = ?",
        ], rule=self.GLOSSED_RULE)
        self.assertEqual(self._run_gate(week), ([], []))

    def test_taught_but_unprinted_part_trips_the_one_recall_rule(self):
        """Team decision (Sept 2026): taught earlier is not printed. -ly is a
        review card on this day and still needs its gloss on the face."""
        week = self._junct_week(["1. joint + -ly = ?"])
        warnings, _ = self._run_gate(week)
        self.assertEqual(len(warnings), 1, warnings)
        self.assertIn("'-ly'", warnings[0])
        self.assertIn("one-recall", warnings[0])

    def test_printed_meaning_clears_an_untaught_part_with_a_note(self):
        """The 2f escape hatch: `in- (onto)` is reachable. Untaught -> NOTE;
        listed in taught_morphemes -> silent."""
        week = self._junct_week(["1. in- (onto) + junct + -tion (the act of) = ?"])
        warnings, notes = self._run_gate(week)
        self.assertEqual(warnings, [])
        self.assertEqual(len(notes), 1, notes)
        self.assertIn("'in-'", notes[0])
        week["taught_morphemes"] = ["in-"]
        self.assertEqual(self._run_gate(week), ([], []))

    def test_you_do_scan_reaches_the_second_prompt_on_a_packed_line(self):
        """10c packs two prompts per line; the tokenizer must see both."""
        text = "1. junct + -tion = ?    2. in- + junct + -tion = ?"
        found = [m.group(1) for m in OG.AFFIX_TOKEN_RE.finditer(text)]
        self.assertIn("in-", found)

    def test_you_do_scan_ignores_hyphenated_ordinary_words(self):
        noise = "T-intersection self-check well-known 10-12yo"
        self.assertEqual([m.group(1) for m in OG.AFFIX_TOKEN_RE.finditer(noise)], [])

    def test_you_do_prose_items_do_not_false_alarm_on_hyphenated_words(self):
        """Prose prompts are scanned now, but re-do is a word, not notation."""
        week = self._junct_week(["Sort these: re-do, un-tie, pre-heat"])
        self.assertEqual(self._run_gate(week), ([], []))

    def test_earlier_session_in_the_week_still_needs_its_gloss(self):
        """eco was taught on Monday: printed it builds silently, unprinted it fails."""
        week = self._junct_week(["1. eco (house) + -tion (the act of) = ?"])
        week["sessions"].insert(0, {
            "day": "Monday",
            "new_morphology": {"morph": "eco", "type": "root"},
            "morphology_review": [], "sound_bank": [],
        })
        self.assertEqual(self._run_gate(week), ([], []))
        week["sessions"][-1]["new_morph_activity"]["items"] = [
            "1. eco + -tion (the act of) = ?"]
        warnings, _ = self._run_gate(week)
        self.assertEqual(len(warnings), 1, warnings)
        self.assertIn("'eco'", warnings[0])

    def test_meaning_sum_items_are_retired(self):
        """The shape the team rejected (Sept 2026): meanings in, whole word out."""
        for item in ("1. small + look at = ?", "bad function = ?",
                     "3. same time, together = ?",
                     "One done for you: house + study of = ecology"):
            week = self._junct_week([item])
            warnings, _ = self._run_gate(week)
            self.assertTrue(any("meaning sum" in w for w in warnings),
                            (item, warnings))

    def test_meaning_sum_is_caught_on_the_check_slide(self):
        """`small + wave = ?` looks like a part sum until the answer shows
        that small is a meaning: it does not survive inside microwave."""
        week = self._junct_week(
            ["1. small + wave = ?"],
            check_items=["1. small + wave = microwave (very small waves)"])
        warnings, _ = self._run_gate(week)
        self.assertTrue(any("'small'" in w and "meaning sum" in w
                            for w in warnings), warnings)

    def _suffix_week(self, morph, items, check_items=(), rule=""):
        return {
            "taught_morphemes": [],
            "sessions": [{
                "day": "Monday",
                "new_morphology": {"morph": morph, "type": "suffix"},
                "morphology_review": [], "sound_bank": [],
                "new_morph_activity": {"rule": rule, "items": items,
                                       "check_items": list(check_items)},
            }],
        }

    def test_word_building_survives_spelling_changes(self):
        """cooperate -> cooperative and private -> privacy are part sums: the
        base visibly survives inside the answer, so nothing fires."""
        week = self._suffix_week(
            "-acy", ["1. private + -acy = ______    2. literate + -acy = ______"],
            check_items=["1. private + -acy = privacy, the state of being private",
                         "2. literate + -acy = literacy, being able to read"])
        self.assertEqual(self._run_gate(week), ([], []))
        week = self._suffix_week(
            "-ive", ["1. cooperate + -ive = ?"],
            check_items=["1. cooperate + -ive = cooperative (working together)"])
        self.assertEqual(self._run_gate(week), ([], []))

    def test_literal_meaning_check_items_scan_the_right_side_sum(self):
        """The -al defect: `microbial = micro + bio + -al` asked students for a
        suffix nobody had printed. The sum on the RIGHT of a check item is the
        student's work."""
        week = {
            "taught_morphemes": [],
            "sessions": [{
                "day": "Tuesday",
                "new_morphology": {"morph": "micro", "type": "root"},
                "morphology_review": [], "sound_bank": [],
                "new_morph_activity": {
                    "rule": "Write the sum beside each word.",
                    "items": ["1. microbial =    2. microsecond ="],
                    "check_items": [
                        "1. microbial = micro + bio + -al = relating to tiny life",
                        "2. microsecond = micro + second = one millionth of a second",
                    ]},
            }],
        }
        warnings, _ = self._run_gate(week)
        self.assertEqual(len(warnings), 1, warnings)
        self.assertIn("'-al'", warnings[0])
        week["sessions"][0]["new_morph_activity"]["rule"] = (
            "Write the sum beside each word. Parts: -al (relating to)")
        warnings, _ = self._run_gate(week)
        self.assertEqual(warnings, [])

    def test_round_gloss_counts_only_on_its_own_slide(self):
        """v1.6 rounds put one item on each slide, so a gloss printed on item 1
        is not on the face of item 2."""
        week = self._junct_week([])
        week["sessions"][0]["new_morph_activity"] = {
            "rule": "Write the word that fits.",
            "round": [
                {"prompt": "A space rock that looks like a star", "answer": "asteroid",
                 "choices": ["asteroid\n-oid (like)"], "why": "it looks like a star"},
                {"prompt": "A rock in space", "answer": "asteroid",
                 "why": "-oid makes it like a star"},
            ],
        }
        warnings, _ = self._run_gate(week)
        self.assertEqual(len(warnings), 1, warnings)
        self.assertIn("'-oid'", warnings[0])

    def test_round_decks_hide_each_answer_behind_a_click(self):
        """Term 4 Week 2 is the first week built on rounds: the gate checks every
        item's green answer bar is click-revealed, and the worked item is not."""
        spec = ROOT / "og_planner" / "weeks" / "term4_week2.json"
        with tempfile.TemporaryDirectory(prefix="og_round_regression_") as temp:
            decks = self.build_spec(spec, Path(temp))
            self.assertEqual(len(decks), 5)
            week = json.loads(spec.read_text(encoding="utf-8"))
            worked = week["sessions"][0]["new_morph_activity"]["round"][0]
            self.assertTrue(worked.get("worked"))
            with zipfile.ZipFile(decks[0]) as archive:
                texts = {}
                for name in archive.namelist():
                    if re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
                        texts[name] = archive.read(name).decode("utf-8")
            first = worked["prompt"].splitlines()[0]
            worked_slides = [x for x in texts.values() if first in x]
            self.assertEqual(len(worked_slides), 1)
            self.assertNotIn("<p:timing", worked_slides[0])

    # ------------------------------------------------ v1.7 five-day shape
    V17_SPEC = ROOT / "og_planner" / "weeks" / "term4_week3.json"

    def _slide_xml(self, deck):
        with zipfile.ZipFile(deck) as archive:
            names = sorted(
                (n for n in archive.namelist()
                 if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                key=lambda n: int(re.search(r"(\d+)\.xml", n).group(1)))
            return [archive.read(n).decode("utf-8") for n in names]

    def test_v17_week_builds_one_deck_per_day_type(self):
        """Kate Watt (Oct 2026): one morpheme a week, each day only its own blocks."""
        with tempfile.TemporaryDirectory(prefix="og_v17_regression_") as temp:
            decks = self.build_spec(self.V17_SPEC, Path(temp))
            names = [d.name for d in decks]
            self.assertEqual(names, [
                "1. Monday (port).pptx", "2. Tuesday (port spelling).pptx",
                "3. Wednesday (port building).pptx", "4. Thursday (parts of speech).pptx",
                "5. Friday (week review).pptx"])
            monday, tuesday, wednesday, thursday, friday = (
                "\n".join(self._slide_xml(d)) for d in decks)
            # Monday teaches: no dictation, spelling or grammar
            self.assertNotIn("Write this sentence", monday)
            self.assertNotIn("Spell the word", monday)
            self.assertIn("Your keyword and one sentence", monday)
            # Tuesday: no card drill, but a dictated bank and dictation
            self.assertNotIn("Morphology Review", tuesday)
            self.assertIn("Write this sentence", tuesday)
            # Wednesday: build and break only
            self.assertIn("Build and Break", wednesday)
            self.assertNotIn("Write this sentence", wednesday)
            # Thursday: parts of speech and the week's one grammar block
            self.assertIn("Suffix Parts of Speech", thursday)
            self.assertIn("I do - ", thursday)
            self.assertNotIn("I do - ", friday.split("Trickiest")[0])
            self.assertIn("Trickiest this week?", friday)

            # The dictated sound bank reveals every box on a click.
            bank = next(x for x in self._slide_xml(decks[1]) if "Write the sound" in x)
            self.assertIn("<p:timing", bank)
            self.assertEqual(bank.count("<p:spTgt"), 9)
            # Parts-of-speech cards hide the answer behind a click.
            pos = [x for x in self._slide_xml(decks[3]) if ">noun<" in x or ">adjective<" in x]
            self.assertTrue(pos)
            self.assertTrue(all("<p:timing" in x for x in pos))
            # COPS replaced CUPS school-wide.
            with zipfile.ZipFile(decks[1]) as archive:
                notes = "".join(archive.read(n).decode("utf-8") for n in archive.namelist()
                                if n.startswith("ppt/notesSlides/notesSlide"))
            self.assertIn("Mark with COPS", notes)
            self.assertIn("O - Organisation", notes)
            self.assertNotIn("CUPS", notes)

    def _v17_week(self, targets, new_lw="economy"):
        teach = {"day": "Monday", "type": "teach",
                 "new_morphology": {"morph": "port", "type": "root"},
                 "words_to_read_new": [{"word": "transport"}, {"word": "export"}]}
        spell = {"day": "Tuesday", "type": "spell",
                 "learned_words": {"new": {"word": new_lw}, "review": []},
                 "dictation": [{"meter": "green", "sentence": "x", "targets": targets}]}
        return {"term": 9, "week": 9, "sessions": [teach, spell]}

    def test_v17_dictation_must_carry_this_weeks_teaching(self):
        """Yoshimoto's planning slide: dictation must prioritise what was introduced."""
        week = self._v17_week(["captain", "salmon"])
        OG.WARNINGS.clear()
        OG.NOTES.clear()
        OG.validate_dictation_sources("Tuesday", week, week["sessions"][1])
        self.assertTrue(any("this week's teaching" in w for w in OG.WARNINGS), OG.WARNINGS)

    def test_v17_dictation_wants_a_throwback(self):
        """All-this-week targets build, with a NOTE asking for one older word."""
        week = self._v17_week(["transported", "economy"])
        OG.WARNINGS.clear()
        OG.NOTES.clear()
        OG.validate_dictation_sources("Tuesday", week, week["sessions"][1])
        self.assertEqual(OG.WARNINGS, [])
        self.assertTrue(any("week or two back" in n for n in OG.NOTES), OG.NOTES)
        week = self._v17_week(["export", "entrepreneur"])
        OG.WARNINGS.clear()
        OG.NOTES.clear()
        OG.validate_dictation_sources("Tuesday", week, week["sessions"][1])
        self.assertEqual((OG.WARNINGS, OG.NOTES), ([], []))

    def test_v17_bank_may_carry_mondays_morpheme_but_not_a_later_one(self):
        """The bank is dictated in v1.7, so Monday's morpheme is retrieval on
        Tuesday; a morpheme taught later in the week is still fatal."""
        teach = {"day": "Monday", "type": "teach",
                 "new_morphology": {"morph": "port", "type": "root"}}
        spell = {"day": "Tuesday", "type": "spell",
                 "sound_bank": [{"morph": "port", "type": "root"},
                                {"morph": "tract", "type": "root"}],
                 "words_to_spell_review": [{"word": "transport"}, {"word": "tractor"}]}
        later = {"day": "Thursday", "type": "teach",
                 "new_morphology": {"morph": "tract", "type": "root"}}
        week = {"sessions": [teach, spell, later]}
        OG.WARNINGS.clear()
        OG.validate_sound_bank("Tuesday", spell, week)
        self.assertEqual(len(OG.WARNINGS), 1, OG.WARNINGS)
        self.assertIn("'tract'", OG.WARNINGS[0])

    V18_SPEC = ROOT / "og_planner" / "weeks" / "term4_week4.json"

    def _notes_text(self, deck):
        with zipfile.ZipFile(deck) as archive:
            return "".join(archive.read(n).decode("utf-8") for n in archive.namelist()
                           if n.startswith("ppt/notesSlides/notesSlide"))

    def test_v18_short_week_opens_every_day_with_review(self):
        """James (22 Oct 2026): cards and reading review every day, hook with a
        picture, frame-based parts of speech, part boxes, four-session week."""
        with tempfile.TemporaryDirectory(prefix="og_v18_regression_") as temp:
            decks = self.build_spec(self.V18_SPEC, Path(temp))
            self.assertEqual([d.name for d in decks], [
                "1. Monday (ject).pptx", "2. Tuesday (ject spelling).pptx",
                "3. Wednesday (ject building).pptx", "4. Thursday (parts of speech).pptx"])
            slides = ["\n".join(self._slide_xml(d)) for d in decks]
            for deck_xml in slides:
                self.assertIn("Morphology Review", deck_xml)
                self.assertIn("Words to Read Review", deck_xml)
            monday, tuesday, wednesday, thursday = slides
            # Memory hook and its picture sit on the new-morpheme card.
            self.assertIn("Eject throws the pilot out!", monday)
            with zipfile.ZipFile(decks[0]) as archive:
                self.assertIn("ppt/media/oghook_plane.png", archive.namelist())
            self.assertTrue(self._notes_text(decks[0]).count("MEMORY HOOK") >= 2)
            self.assertNotIn("PROCEDURE", self._notes_text(decks[0]).split("MEMORY HOOK")[0]
                             .split("New morpheme today")[-1])
            # No recap card on Tuesday; no retired note lines.
            tue_notes = self._notes_text(decks[1])
            self.assertNotIn("Click to reveal answer.", tue_notes)
            self.assertNotIn("SAY each word, students write it", tue_notes)
            # Break it apart: the root box is filled, the others are empty.
            breaks = [x for x in self._slide_xml(decks[2]) if "Break it apart" in x]
            self.assertTrue(breaks)
            self.assertTrue(all(">+<" in x for x in breaks))
            # Thursday: frames on every suffix card, the vote closes the week,
            # one dictation sentence, reading review before parts of speech.
            self.assertEqual(thursday.count(">very ___<"), 4)
            self.assertIn("Trickiest this week?", thursday)
            self.assertEqual(thursday.count("Write this sentence"), 1)
            self.assertLess(thursday.index("Words to Read Review"),
                            thursday.index("Suffix Parts of Speech"))
            for deck in decks:
                self.assertNotIn("give me", self._notes_text(deck).lower())

    def test_v18_rules_fail_the_build(self):
        """The v1.8 rules are gates, not advice."""
        week = json.loads(self.V18_SPEC.read_text())
        mon, tue, wed, thu = week["sessions"]
        mon["morphology_review"][0]["extra_task"] = {
            "prompt": "Okay, give me one port word, meaning carry.", "answers": ["export"]}
        mon["memory_hook"]["text"] = "ject means throw so think of throwing a ball far away"
        thu["pos_drill"]["cards"][0].pop("sentence")
        thu.pop("vote")
        thu["dictation"].append(dict(thu["dictation"][0], meter="yellow"))
        wed["build"]["rounds"][0]["round"][1].pop("parts")
        with tempfile.TemporaryDirectory(prefix="og_v18_fail_") as temp:
            spec = Path(temp) / "term4_week4.json"
            spec.write_text(json.dumps(week))
            result = subprocess.run(
                [sys.executable, str(BUILDER_PATH), str(spec), "--outdir", temp],
                cwd=ROOT, capture_output=True, text=True, timeout=120, check=False)
            out = result.stdout
            self.assertNotEqual(result.returncode, 0, out)
            self.assertIn("give me", out)
            self.assertIn("8 words", out)
            self.assertIn("sentence", out)
            self.assertIn("`vote`", out)
            self.assertIn("ONE dictation", out)
            audit = subprocess.run(
                [sys.executable, str(ROOT / "og_planner" / "audit_week_spec.py"), str(spec)],
                cwd=ROOT, capture_output=True, text=True, timeout=60, check=False)
            self.assertIn("parts", audit.stdout)


if __name__ == "__main__":
    unittest.main(verbosity=2)
