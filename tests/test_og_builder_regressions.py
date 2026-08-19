"""Regression gate for the locked-template Orton-Gillingham deck builder.

Run from the repository root:
    python tests/test_og_builder_regressions.py
"""

import importlib.util
import json
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

    # ------------------------------------------------ 2f taught-morpheme gate
    def _junct_week(self, items):
        return {
            "taught_morphemes": [],
            "sessions": [{
                "day": "Wednesday",
                "new_morphology": {"morph": "junct/join/joint", "type": "root"},
                "morphology_review": [{"morph": "micro", "type": "root"},
                                      {"morph": "-ly", "type": "suffix"}],
                "sound_bank": [{"morph": "dis-", "type": "prefix"},
                               {"morph": "-tion / -sion", "type": "suffix"}],
                "new_morph_activity": {"items": items},
            }],
        }

    def test_you_do_flags_an_untaught_prefix(self):
        """in- appeared once in a grid script line, so it is not taught."""
        week = self._junct_week([
            "1. junct + -tion = ?    2. dis- + joint + -ed = ?",
            "3. joint + -ly = ?    4. in- + junct + -tion = ?",
        ])
        OG.WARNINGS.clear()
        OG.validate_activity_morphemes("Wednesday", week, week["sessions"][0])
        flagged = [w for w in OG.WARNINGS if "You Do uses morpheme" in w]
        self.assertEqual(len(flagged), 1, OG.WARNINGS)
        self.assertIn("'in-'", flagged[0])

    def test_you_do_accepts_taught_parts_and_common_inflections(self):
        """dis- and -tion are banked, -ly is a review card, -ed is universal."""
        week = self._junct_week([
            "1. junct + -tion = ?    2. dis- + joint + -ed = ?",
            "3. joint + -ly = ?",
        ])
        OG.WARNINGS.clear()
        OG.validate_activity_morphemes("Wednesday", week, week["sessions"][0])
        self.assertEqual(OG.WARNINGS, [])

    def test_you_do_scan_reaches_the_second_prompt_on_a_packed_line(self):
        """10c packs two prompts per line; the tokenizer must see both."""
        text = "1. junct + -tion = ?    2. in- + junct + -tion = ?"
        found = [m.group(1) for m in OG.AFFIX_TOKEN_RE.finditer(text)]
        self.assertIn("in-", found)

    def test_you_do_scan_ignores_hyphenated_ordinary_words(self):
        noise = "T-intersection self-check well-known 10-12yo"
        self.assertEqual([m.group(1) for m in OG.AFFIX_TOKEN_RE.finditer(noise)], [])

    def test_you_do_prose_items_are_not_scanned_for_affixes(self):
        """No sum, no morpheme notation - a sorting task must not false-alarm."""
        week = self._junct_week(["Sort these: re-do, un-tie, pre-heat"])
        OG.WARNINGS.clear()
        OG.validate_activity_morphemes("Wednesday", week, week["sessions"][0])
        self.assertEqual(OG.WARNINGS, [])

    def test_week_taught_morphemes_clears_an_earlier_taught_affix(self):
        week = self._junct_week(["1. in- + junct + -tion = ?"])
        OG.WARNINGS.clear()
        OG.validate_activity_morphemes("Wednesday", week, week["sessions"][0])
        self.assertTrue(OG.WARNINGS)
        week["taught_morphemes"] = ["in-"]
        OG.WARNINGS.clear()
        OG.validate_activity_morphemes("Wednesday", week, week["sessions"][0])
        self.assertEqual(OG.WARNINGS, [])

    def test_earlier_session_in_the_week_counts_as_taught(self):
        week = self._junct_week(["1. eco + -tion = ?"])
        week["sessions"].insert(0, {
            "day": "Monday",
            "new_morphology": {"morph": "eco", "type": "root"},
            "morphology_review": [], "sound_bank": [],
        })
        OG.WARNINGS.clear()
        OG.validate_activity_morphemes("Wednesday", week, week["sessions"][1])
        self.assertEqual(OG.WARNINGS, [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
