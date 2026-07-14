#!/usr/bin/env python3
"""Build a multi-session unit end-to-end.

Reads a JSON manifest, runs `node scripts/build_and_check.js` for each lesson
build script (aborting on the first failure), runs `merge_unit.py` to produce
a single combined PPTX and a flat Resources/ folder, generates the optional
one-page Teacher Week Brief from manifest data, then runs merged unit QA.

Manifest format (extends merge_unit.py's manifest with a `build_script`
field on each lesson):

    {
      "unit_folder": "Decimals_and_Fractions_Unit",
      "unit_pptx_name": "Decimals and Fractions Unit.pptx",
      "teacher_brief": {"...": "see docs/resource-system.md"},
      "lessons": [
        {
          "build_script": "builds/build_decfrac_lesson1.js",
          "folder": "DecFrac_Lesson1_Place_Value_With_Decimals",
          "session": 1
        }
      ]
    }

Usage:
    python scripts/build_unit.py builds/manifests/<unit>.json
    python scripts/build_unit.py builds/manifests/<unit>.json --skip-build
    python scripts/build_unit.py builds/manifests/<unit>.json --skip-build --skip-qa

The `--skip-build` flag merges already-built lesson folders without re-running
the per-lesson build gate, but still runs merged unit QA. Use `--skip-qa` only
for local merge debugging, not delivery.
"""

import argparse
import json
import subprocess
import sys
from pathlib import Path

from merge_unit import OUTPUT_ROOT, load_manifest, merge_unit

ROOT = Path(__file__).resolve().parent.parent
BUILD_GATE = ROOT / "scripts" / "build_and_check.js"
UNIT_QA = ROOT / "scripts" / "qa_unit.js"


def run_lesson_build(build_script: Path) -> None:
    if not build_script.is_file():
        raise SystemExit(f"Build script not found: {build_script}")

    print(f"\n=== Building {build_script.relative_to(ROOT)} ===")
    result = subprocess.run(
        ["node", str(BUILD_GATE), str(build_script.relative_to(ROOT))],
        cwd=str(ROOT),
    )
    if result.returncode != 0:
        raise SystemExit(
            f"Build gate failed for {build_script} (exit {result.returncode}). "
            f"Fix the build before re-running."
        )


def build_lessons(manifest: dict) -> None:
    for lesson in manifest["lessons"]:
        build_script = lesson.get("build_script")
        if not build_script:
            raise SystemExit(
                f"Lesson missing 'build_script': {lesson}. "
                f"Add the path to the per-lesson build script in the manifest."
            )
        run_lesson_build(ROOT / build_script)


def run_unit_qa(manifest_path: Path) -> None:
    try:
        manifest_arg = str(manifest_path.relative_to(ROOT))
    except ValueError:
        manifest_arg = str(manifest_path)
    print("\n=== Unit QA ===")
    result = subprocess.run(
        [
            "node",
            str(UNIT_QA),
            manifest_arg,
            "--skip-build",
            "--skip-merge",
        ],
        cwd=str(ROOT),
    )
    if result.returncode != 0:
        raise SystemExit(
            f"Unit QA failed for {manifest_path} (exit {result.returncode}). "
            f"Fix the merged output before delivery."
        )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "manifest",
        type=Path,
        help="Path to a unit manifest JSON file.",
    )
    parser.add_argument(
        "--skip-build",
        action="store_true",
        help="Skip the per-lesson build gate and merge already-built lesson folders.",
    )
    parser.add_argument(
        "--skip-qa",
        action="store_true",
        help="Skip merged unit QA after building/merging.",
    )
    args = parser.parse_args()

    if not args.manifest.is_file():
        print(f"Manifest not found: {args.manifest}", file=sys.stderr)
        return 2

    manifest = load_manifest(args.manifest)

    if not args.skip_build:
        build_lessons(manifest)

    print("\n=== Merging unit ===")
    merge_unit(args.manifest)
    if not args.skip_qa:
        run_unit_qa(args.manifest.resolve())
    print("\nDone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
