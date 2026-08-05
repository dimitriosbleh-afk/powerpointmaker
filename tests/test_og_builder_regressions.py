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
CURRENT_SPEC = ROOT / "og_planner" / "weeks" / "term3_week1.json"
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

    def test_current_and_sample_specs_pass_every_output_gate(self):
        with tempfile.TemporaryDirectory(prefix="og_builder_regression_") as temp:
            temp_root = Path(temp)
            current = self.build_spec(CURRENT_SPEC, temp_root / "current")
            sample = self.build_spec(SAMPLE_SPEC, temp_root / "sample")
            self.assertEqual(len(current), 4)
            self.assertEqual(len(sample), 4)
            self.assert_deck_structure(CURRENT_SPEC, current)
            self.assert_deck_structure(SAMPLE_SPEC, sample)

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


if __name__ == "__main__":
    unittest.main(verbosity=2)
