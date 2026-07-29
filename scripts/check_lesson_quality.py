#!/usr/bin/env python3
"""Check generated PPTX lesson decks for lean-lesson guardrail drift."""

from __future__ import annotations

import argparse
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NOTES_REL_SUFFIX = "/notesSlide"
MARKDOWN_PATTERNS = [
    re.compile(r"\*\*[^*]+\*\*"),
    re.compile(r"_[^_\n]+_"),
    re.compile(r"`[^`]+`"),
    re.compile(r"^\s{0,3}#{1,6}\s+", re.MULTILINE),
    re.compile(r"^\s*[*+]\s+", re.MULTILINE),
]
ASSUMPTION_PATTERNS = [
    re.compile(r"\byou already know\b", re.IGNORECASE),
    re.compile(r"\bstudents know the routine\b", re.IGNORECASE),
    re.compile(r"\bstudents know the drill\b", re.IGNORECASE),
    re.compile(r"\bnot new to you\b", re.IGNORECASE),
    re.compile(r"\bwe(?:'|\u2019)ve done this\b", re.IGNORECASE),
    re.compile(r"\byou know this\b", re.IGNORECASE),
    re.compile(r"\bby week\s+\d+\b[^.\n]*\bknow", re.IGNORECASE),
]
ADVANCED_SC1_RE = re.compile(
    r"\b(analyse|analyze|evaluate|justify|synthesi[sz]e|compare|explain how|identify and explain)\b",
    re.IGNORECASE,
)
NON_ASCII_RE = re.compile(r"[^\x09\x0A\x0D\x20-\x7E]")
RESOURCE_CODE_RE = re.compile(r"\b[A-Z]{2,}\d+(?:_[A-Za-z0-9]+)*\b")
DAY_NAME_RE = re.compile(r"\b(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b", re.IGNORECASE)
LI_NOTE_MISMATCH_PATTERNS = [
    re.compile(r"\b(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s+learning intentions\b", re.IGNORECASE),
    re.compile(r"\b(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s+lis\b", re.IGNORECASE),
    re.compile(r"\bhere are our learning intentions\b", re.IGNORECASE),
    re.compile(r"\blearning intentions today\b", re.IGNORECASE),
    re.compile(r"\bread each li\b", re.IGNORECASE),
    re.compile(r"\bread each learning intention\b", re.IGNORECASE),
    re.compile(r"\bpoint to each li\b", re.IGNORECASE),
    re.compile(r"\bpoint to each learning intention\b", re.IGNORECASE),
    re.compile(r"\bchoral read lis?\b", re.IGNORECASE),
]

# Note: per-section note budgets (SAY / DO / TEACHER NOTES / WATCH FOR bullet
# caps) used to live here. They matched the pre-v11 sectioned note format and
# became dead the moment decks moved to the Glance Format, which has no such
# headers. Glance live-zone budgets are now enforced as hard errors by
# build_and_check.js Gate 4, so they are not duplicated here.
PROFILES = {
    "literacy-60": {
        "max_unique_warn": 14,
        "max_unique_error": 16,
        "max_reveal_pairs": 2,
        "max_explicit_vocab": 2,
    }
}


@dataclass
class SlideData:
    index: int
    slide_path: str
    text_lines: list[str]
    notes_lines: list[str]

    @property
    def slide_text(self) -> str:
        return "\n".join(self.text_lines)

    @property
    def notes_text(self) -> str:
        return "\n".join(self.notes_lines)


@dataclass
class Issue:
    severity: str
    location: str
    message: str


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def normalize_zip_path(base: str, target: str) -> str:
    if target.startswith("/"):
        target = target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), target))


def read_xml(zf: zipfile.ZipFile, path: str) -> ET.Element:
    return ET.fromstring(zf.read(path))


def extract_paragraphs(xml_bytes: bytes) -> list[str]:
    root = ET.fromstring(xml_bytes)
    paragraphs: list[str] = []
    for elem in root.iter():
        if local_name(elem.tag) != "p":
            continue
        parts = [(node.text or "") for node in elem.iter() if local_name(node.tag) == "t"]
        text = re.sub(r"\s+", " ", "".join(parts)).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def ordered_slide_paths(zf: zipfile.ZipFile) -> list[str]:
    pres_path = "ppt/presentation.xml"
    rels_path = "ppt/_rels/presentation.xml.rels"
    pres_root = read_xml(zf, pres_path)
    rels_root = read_xml(zf, rels_path)

    rel_map: dict[str, str] = {}
    for rel in rels_root:
        if local_name(rel.tag) != "Relationship":
            continue
        rel_map[rel.attrib["Id"]] = normalize_zip_path(pres_path, rel.attrib["Target"])

    paths: list[str] = []
    for elem in pres_root.iter():
        if local_name(elem.tag) != "sldId":
            continue
        rel_id = elem.attrib.get(f"{{{REL_NS}}}id")
        if rel_id and rel_id in rel_map:
            paths.append(rel_map[rel_id])
    return paths


def notes_path_for_slide(zf: zipfile.ZipFile, slide_path: str) -> str | None:
    rels_path = posixpath.join(posixpath.dirname(slide_path), "_rels", posixpath.basename(slide_path) + ".rels")
    if rels_path not in zf.namelist():
        return None

    rels_root = read_xml(zf, rels_path)
    for rel in rels_root:
        if local_name(rel.tag) != "Relationship":
            continue
        rel_type = rel.attrib.get("Type", "")
        if rel_type.endswith(NOTES_REL_SUFFIX):
            return normalize_zip_path(slide_path, rel.attrib["Target"])
    return None


def load_slides(pptx_path: Path) -> list[SlideData]:
    with zipfile.ZipFile(pptx_path) as zf:
        slides: list[SlideData] = []
        for idx, slide_path in enumerate(ordered_slide_paths(zf), start=1):
            text_lines = extract_paragraphs(zf.read(slide_path))
            notes_path = notes_path_for_slide(zf, slide_path)
            notes_lines = extract_paragraphs(zf.read(notes_path)) if notes_path and notes_path in zf.namelist() else []
            slides.append(SlideData(idx, slide_path, text_lines, notes_lines))
        return slides


def normalize_for_similarity(text: str) -> set[str]:
    return {
        token
        for token in re.sub(r"[^a-z0-9\s]", " ", text.lower()).split()
        if len(token) > 1
    }


def is_reveal_pair(left: SlideData, right: SlideData) -> bool:
    left_tokens = normalize_for_similarity(left.slide_text)
    right_tokens = normalize_for_similarity(right.slide_text)
    if not left_tokens or not right_tokens:
        return False

    overlap = len(left_tokens & right_tokens) / max(1, min(len(left_tokens), len(right_tokens)))
    same_first_line = bool(left.text_lines and right.text_lines and left.text_lines[0].strip().lower() == right.text_lines[0].strip().lower())
    reveal_hint = "reveal" in left.notes_text.lower() or "reveal" in right.notes_text.lower()
    return len(right_tokens) >= len(left_tokens) and overlap >= 0.82 and (same_first_line or reveal_hint)


def likely_footer(line: str) -> bool:
    lower = line.lower()
    return (
        "|" in line
        or ("lesson" in lower and "week" in lower)
        or ("year" in lower and "week" in lower)
    )


def is_resource_slide(slide: SlideData) -> bool:
    lower_lines = [line.strip().lower() for line in slide.text_lines if line.strip()]
    return "teacher resources" in lower_lines or "printable resources" in lower_lines


def resource_slide_issues(slide: SlideData) -> list[Issue]:
    if not is_resource_slide(slide):
        return []

    issues: list[Issue] = []
    ignored_prefixes = (
        "teacher resources",
        "printable resources",
        "click any resource below",
    )
    ignored_exact = {"pdf"}

    for line in slide.text_lines:
        stripped = line.strip()
        lower = stripped.lower()
        if not stripped or likely_footer(stripped):
            continue
        if lower in ignored_exact or any(lower.startswith(prefix) for prefix in ignored_prefixes):
            continue
        if "_" in stripped:
            issues.append(Issue("warning", f"slide {slide.index}", f"Resource slide contains an underscore-heavy teacher-facing label: '{stripped}'.")) 
        if RESOURCE_CODE_RE.search(stripped):
            issues.append(Issue("warning", f"slide {slide.index}", f"Resource slide contains a code-style resource label: '{stripped}'.")) 
        if DAY_NAME_RE.search(stripped):
            issues.append(Issue("warning", f"slide {slide.index}", f"Resource slide contains a day name; use session numbering instead: '{stripped}'.")) 

    return issues


def li_sc_issues(slide: SlideData) -> list[Issue]:
    lines = [line.strip() for line in slide.text_lines if line.strip()]
    joined = " ".join(lines).lower()
    if "learning intention" not in joined or "success criteria" not in joined:
        return []

    issues: list[Issue] = []
    sc_indices = [idx for idx, line in enumerate(lines) if "success criteria" in line.lower()]
    if not sc_indices:
        return issues
    sc_index = next(
        (idx for idx in sc_indices if "learning intention" not in lines[idx].lower()),
        sc_indices[-1],
    )

    li_candidates = [
        line
        for line in lines[:sc_index]
        if "learning intention" not in line.lower() and "success criteria" not in line.lower() and not likely_footer(line)
    ]
    sc_candidates = [
        line
        for line in lines[sc_index + 1 :]
        if "success criteria" not in line.lower() and not likely_footer(line)
    ]

    if len(li_candidates) != 1:
        issues.append(Issue("warning", f"slide {slide.index}", f"LI/SC slide appears to contain {len(li_candidates)} Learning Intention lines; expected 1."))
    if len(sc_candidates) != 3:
        issues.append(Issue("warning", f"slide {slide.index}", f"LI/SC slide appears to contain {len(sc_candidates)} Success Criteria lines; expected 3."))
    notes_text = "\n".join(
        line.strip()
        for line in slide.notes_lines
        if line.strip() and not line.strip().startswith("[")
    )
    if len(li_candidates) == 1 and any(pattern.search(notes_text) for pattern in LI_NOTE_MISMATCH_PATTERNS):
        issues.append(Issue("warning", f"slide {slide.index} notes", "LI/SC notes refer to multiple learning intentions, but the slide contains a single Learning Intention."))
    if sc_candidates and ADVANCED_SC1_RE.search(sc_candidates[0]):
        issues.append(Issue("warning", f"slide {slide.index}", "SC1 looks too demanding for an ultra-achievable first criterion."))
    return issues


def classify_vocab_slide(slide: SlideData) -> tuple[bool, bool]:
    lower = slide.slide_text.lower()
    incidental = "incidental vocabulary" in lower or "words to listen for" in lower
    explicit = (
        "vocabulary" in lower
        and ("word study" in lower or "definition" in lower or "example" in lower)
        and not incidental
    )
    return explicit, incidental


def scan_assumption_phrases(text: str) -> list[str]:
    hits: list[str] = []
    for pattern in ASSUMPTION_PATTERNS:
        for match in pattern.finditer(text):
            hits.append(match.group(0))
    return hits


def analyze(slides: list[SlideData], profile_name: str) -> tuple[list[Issue], dict[str, int]]:
    profile = PROFILES[profile_name]
    issues: list[Issue] = []

    reveal_pairs = 0
    idx = 0
    while idx < len(slides) - 1:
        if is_reveal_pair(slides[idx], slides[idx + 1]):
            reveal_pairs += 1
            idx += 2
        else:
            idx += 1

    unique_slide_count = len(slides) - reveal_pairs
    explicit_vocab_count = 0
    incidental_vocab_count = 0

    if unique_slide_count > profile["max_unique_error"]:
        issues.append(Issue("error", "deck", f"Unique slide count is {unique_slide_count}; profile limit is {profile['max_unique_error']} before it becomes excessive for a lean 60-minute lesson."))
    elif unique_slide_count > profile["max_unique_warn"]:
        issues.append(Issue("warning", "deck", f"Unique slide count is {unique_slide_count}; profile target is 10-{profile['max_unique_warn']} for a lean 60-minute lesson."))

    if reveal_pairs > profile["max_reveal_pairs"]:
        issues.append(Issue("warning", "deck", f"Detected {reveal_pairs} reveal pairs; profile default is 0-{profile['max_reveal_pairs']}."))

    for slide in slides:
        explicit_vocab, incidental_vocab = classify_vocab_slide(slide)
        explicit_vocab_count += int(explicit_vocab)
        incidental_vocab_count += int(incidental_vocab)

        notes_text = slide.notes_text

        for hit in scan_assumption_phrases(slide.slide_text):
            issues.append(Issue("warning", f"slide {slide.index}", f"Slide text uses mastery-assuming language: '{hit}'."))
        for hit in scan_assumption_phrases(notes_text):
            issues.append(Issue("warning", f"slide {slide.index} notes", f"Notes use mastery-assuming language: '{hit}'."))

        if notes_text:
            if any(pattern.search(notes_text) for pattern in MARKDOWN_PATTERNS):
                issues.append(Issue("warning", f"slide {slide.index} notes", "Notes appear to contain markdown syntax."))
            if NON_ASCII_RE.search(notes_text):
                issues.append(Issue("warning", f"slide {slide.index} notes", "Notes contain non-ASCII characters; PowerPoint notes should stay ASCII-safe."))
            if "PACING OVERVIEW" in notes_text.upper():
                issues.append(Issue("warning", f"slide {slide.index} notes", "Notes contain a PACING OVERVIEW block; lean notes should avoid it."))

        issues.extend(li_sc_issues(slide))
        issues.extend(resource_slide_issues(slide))

    if explicit_vocab_count > profile["max_explicit_vocab"]:
        issues.append(Issue("warning", "deck", f"Detected {explicit_vocab_count} explicit vocabulary slides; profile default is 0-{profile['max_explicit_vocab']}."))
    if incidental_vocab_count:
        issues.append(Issue("warning", "deck", f"Detected {incidental_vocab_count} incidental vocabulary list slide(s); these are off by default in the lean literacy profile."))

    summary = {
        "actual_slides": len(slides),
        "unique_slides": unique_slide_count,
        "reveal_pairs": reveal_pairs,
        "explicit_vocab_slides": explicit_vocab_count,
        "incidental_vocab_slides": incidental_vocab_count,
    }
    return issues, summary


def print_report(pptx_path: Path, profile_name: str, issues: list[Issue], summary: dict[str, int]) -> None:
    print(f"Lesson quality report: {pptx_path}")
    print(
        "Profile: {profile} | slides={actual} unique={unique} reveal_pairs={reveal} explicit_vocab={vocab} incidental_vocab={incidental}".format(
            profile=profile_name,
            actual=summary["actual_slides"],
            unique=summary["unique_slides"],
            reveal=summary["reveal_pairs"],
            vocab=summary["explicit_vocab_slides"],
            incidental=summary["incidental_vocab_slides"],
        )
    )
    if not issues:
        print("No issues found.")
        return
    for issue in issues:
        print(f"{issue.severity.upper()}: {issue.location}: {issue.message}")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path, help="Path to the generated PPTX file")
    parser.add_argument("--profile", default="literacy-60", choices=sorted(PROFILES), help="Lesson quality profile")
    parser.add_argument("--strict", action="store_true", help="Return non-zero if any issues are found")
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    if not args.pptx.exists():
        print(f"PPTX not found: {args.pptx}", file=sys.stderr)
        return 2
    if args.pptx.suffix.lower() != ".pptx":
        print(f"Expected a .pptx file: {args.pptx}", file=sys.stderr)
        return 2

    try:
        slides = load_slides(args.pptx)
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        print(f"Unable to read PPTX: {exc}", file=sys.stderr)
        return 2

    issues, summary = analyze(slides, args.profile)
    print_report(args.pptx, args.profile, issues, summary)

    has_error = any(issue.severity == "error" for issue in issues)
    if has_error or (args.strict and issues):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
