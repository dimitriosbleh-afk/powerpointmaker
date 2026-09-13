#!/usr/bin/env python3
"""Pre-build audit for an OG week spec: the checks build_og_week.py does not run.

Run BEFORE the builder. It never edits anything.

    python og_planner/audit_week_spec.py og_planner/weeks/term3_week10.json

ISSUE = almost certainly a defect (exit code 1). CHECK = needs a human look.
Covers: dictation character bands and bracket use (section 6), same-day and
week-level word repeats (2b, 2d - the builder skips the overlap check on
week_review days), review-card order and position repeats (2a), scripted
EXPECT answers missing from the Words to Read Review board (2b), sound bank
labels not on the taught timeline (2c), structured-slide text that will wrap
(10c) and learned-word values that differ from teaching_record.json (5).
"""

import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import build_og_week as og  # noqa: E402  (morph_surface_forms only)

ISSUES, CHECKS = [], []


def issue(msg):
    ISSUES.append(msg)


def check(msg):
    CHECKS.append(msg)


def taught_timeline():
    data = json.loads((HERE / "taught_morphemes.json").read_text())
    taught, unconfirmed = set(), set()
    for key, labels in data.items():
        if key.startswith("term") and isinstance(labels, list):
            for label in labels:
                taught |= og.morph_surface_forms(label)
    for label in data.get("_unconfirmed", {}).get("labels", []):
        unconfirmed |= og.morph_surface_forms(label)
    return taught, unconfirmed


def settled_learned_words():
    record = json.loads((HERE / "teaching_record.json").read_text())
    settled = {}
    for term in record["learned_words"].values():
        for item in term.get("words", []):
            settled[item["word"]] = item
    return settled


def audit_dictation(day, session):
    for entry in session.get("dictation", []):
        sentence = entry.get("sentence", "")
        length = len(sentence)
        meter = entry.get("meter")
        low, high = (55, 75) if meter == "green" else (80, 100)
        if length > 105 or not low <= length <= high:
            issue(f"{day}: {meter} dictation is {length} characters (band {low}-{high}, ceiling 105): {sentence}")
        for target in entry.get("targets", []):
            if target.lower() not in sentence.lower():
                issue(f"{day}: dictation target {target!r} is not in its sentence")
        if re.search(r"[()\[\]]", sentence):
            issue(f"{day}: dictation uses brackets - the builder does not colour them red, so CUPS "
                  "punctuation would be wrong on the slide. Keep brackets on grammar slides only")


def audit_words(week, prev_last):
    sessions = week["sessions"]
    spelled = {}
    previous_read = set(prev_last.get("words_to_read_review", {}).get("words", [])) if prev_last else set()
    for session in sessions:
        day = session["day"]
        read = session.get("words_to_read_review", {}).get("words", [])
        spell = [w["word"] for w in session.get("words_to_spell_review", [])]
        for name, items, want in (("words_to_read_review", read, 15), ("words_to_spell_review", spell, 10),
                                  ("morphology_review", session.get("morphology_review", []), 10),
                                  ("sound_bank", session.get("sound_bank", []), 9)):
            if len(items) != want:
                issue(f"{day}: {name} has {len(items)} entries (expected {want})")
        overlap = set(read) & set(spell)
        if overlap:
            issue(f"{day}: same word in review reading and review spelling: {', '.join(sorted(overlap))}")
        repeated = set(read) & previous_read
        if repeated:
            issue(f"{day}: review reading word repeated from the previous session: {', '.join(sorted(repeated))}")
        previous_read = set(read)
        for word in spell:
            if word in spelled:
                check(f"{day}: review spelling word {word!r} was already spelled on {spelled[word]} this week")
            spelled[word] = day

        board = set(read)
        for line in session.get("words_to_read_review", {}).get("notes", "").splitlines():
            if not re.search(r"\bWhich\b.*\bwords?\b", line):
                continue
            match = re.search(r"EXPECT:\s*(.+?)\.?\s*$", line)
            if not match:
                continue
            for answer in (a.strip() for a in match.group(1).split(",")):
                if answer and answer not in board:
                    issue(f"{day}: scripted answer {answer!r} is not on the Words to Read Review board")


def audit_cards(week):
    sessions = week["sessions"]
    orders = {}
    for session in sessions:
        order = tuple(card["morph"] for card in session.get("morphology_review", []))
        if order in orders:
            issue(f"{session['day']}: review cards are in the same order as {orders[order]}")
        orders[order] = session["day"]
    for before, after in zip(sessions, sessions[1:]):
        positions = {card["morph"]: i for i, card in enumerate(before.get("morphology_review", []))}
        for i, card in enumerate(after.get("morphology_review", [])):
            if positions.get(card["morph"]) == i:
                issue(f"{after['day']}: card {card['morph']!r} sits in position {i + 1} again after {before['day']}")


def audit_bank(week):
    taught, unconfirmed = taught_timeline()
    for session in week["sessions"]:
        for item in session.get("sound_bank", []):
            forms = og.morph_surface_forms(item["morph"])
            if forms & unconfirmed and not forms & taught:
                issue(f"{session['day']}: sound bank {item['morph']!r} is on the _unconfirmed list in "
                      "taught_morphemes.json - not treated as taught until the teacher confirms")
            elif not forms & taught:
                issue(f"{session['day']}: sound bank {item['morph']!r} is not on the taught timeline")


def audit_text_budget(ctx, block, is_check=False):
    size = block.get("check_item_size", 22) if is_check else block.get("item_size", 15)
    limit = int(8.2 * 72 / (size * 0.55))
    # Check-slide answers are bold and wrap earlier. Verified in renders (Term 3 Week 10):
    # 52 bold characters wrapped at 20 pt, 55 fitted at 19 pt - letter shapes decide it.
    bold_limit = int(8.2 * 72 / (size * 0.6))
    for item in block.get("check_items" if is_check else "items", []):
        if len(item) > limit:
            issue(f"{ctx}: line of {len(item)} characters wraps at {size} pt (limit ~{limit}): {item}")
        elif is_check and len(item) > bold_limit:
            check(f"{ctx}: bold answer of {len(item)} characters may wrap at {size} pt - render to "
                  f"confirm, or shorten to ~{bold_limit}: {item}")
    if is_check:
        rule = block.get("check_rule", "")
    else:
        rule = block.get("rule", "")
        hero = [line for line in str(block.get("example", "")).split("\n") if line.strip()]
        if len(hero) > 3:
            issue(f"{ctx}: hero example has {len(hero)} lines (max 3)")
        if len(hero) > 1 and max(len(line) for line in hero) > 47:
            check(f"{ctx}: a line in a multi-line hero is over 47 characters and may wrap")
    # Renders (Term 3 Week 10): banners of 70-79 characters sometimes left one word alone
    # on line two (70, 74, 76, 77, 79 did; 71 and 73 did not). 69 or fewer always fitted,
    # and 83+ reads as a deliberate two-line banner.
    if 70 <= len(rule) <= 82:
        check(f"{ctx}: rule banner is {len(rule)} characters - may leave one word alone on "
              "line two; render to confirm, or cut to 69 or fewer")
    routine = block.get("check_routine" if is_check else "routine", "")
    footer = block.get("check_footer" if is_check else "footer", "")
    if len(routine) > 30:
        check(f"{ctx}: routine chip is {len(routine)} characters (keep under 30)")
    if len(footer) > 45:
        issue(f"{ctx}: footer is {len(footer)} characters (keep under 45)")


def audit_slides(week):
    for session in week["sessions"]:
        day = session["day"]
        grammar = session.get("grammar", {})
        for key in ("i_do", "we_do", "you_do"):
            if key in grammar:
                audit_text_budget(f"{day} grammar {key}", grammar[key])
        if grammar.get("you_do", {}).get("check_items"):
            audit_text_budget(f"{day} grammar check slide", grammar["you_do"], is_check=True)
        activity = session.get("new_morph_activity")
        if activity:
            audit_text_budget(f"{day} You Do", activity)
            if activity.get("check_items"):
                audit_text_budget(f"{day} You Do check slide", activity, is_check=True)


def audit_learned_words(week):
    settled = settled_learned_words()
    for session in week["sessions"]:
        learned = session.get("learned_words") or {}
        items = list(learned.get("review", [])) + ([learned["new"]] if learned.get("new") else [])
        for item in items:
            record = settled.get(item.get("word"))
            if not record:
                continue
            if item.get("unfair") != record["unfair"]:
                check(f"{session['day']}: {item['word']} highlights {item.get('unfair')!r}; "
                      f"teaching_record.json settled on {record['unfair']!r}")
            if record["say_it"] not in item.get("notes", ""):
                check(f"{session['day']}: {item['word']} notes do not use the settled Say it {record['say_it']!r}")


def main():
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    path = Path(sys.argv[1])
    week = json.loads(path.read_text())
    prev_path = path.with_name(f"term{week['term']}_week{int(week['week']) - 1}.json")
    prev_last = None
    if prev_path.exists():
        prev_last = json.loads(prev_path.read_text())["sessions"][-1]
    else:
        check(f"no previous week spec at {prev_path.name}; first-session repeat check skipped "
              "(name specs term<T>_week<N>.json or the builder's last-week dictation check is skipped too)")

    for session in week["sessions"]:
        audit_dictation(session["day"], session)
    audit_words(week, prev_last)
    audit_cards(week)
    audit_bank(week)
    audit_slides(week)
    audit_learned_words(week)

    for msg in ISSUES:
        print(f"ISSUE: {msg}")
    for msg in CHECKS:
        print(f"CHECK: {msg}")
    print(f"\n{len(ISSUES)} issue(s), {len(CHECKS)} check(s).")
    sys.exit(1 if ISSUES else 0)


if __name__ == "__main__":
    main()
