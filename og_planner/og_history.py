#!/usr/bin/env python3
"""Print the OG teaching history recorded in the week specs.

Replaces hand-mining old specs when planning review lists, dictation reach and
word repeats (OG_MEGA_PROMPT.md sections 2b, 2d, 6 and 11).

Usage:
    python og_planner/og_history.py                  # every term3_week*.json, full detail
    python og_planner/og_history.py --term 3 --from 7 --to 9
    python og_planner/og_history.py --index          # word -> every place it was used
"""

import argparse
import json
import re
from pathlib import Path

WEEKS = Path(__file__).resolve().parent / "weeks"
DAY = {"Monday": "Mon", "Tuesday": "Tue", "Wednesday": "Wed", "Thursday": "Thu", "Friday": "Fri"}


def week_specs(term, first, last):
    found = []
    for path in WEEKS.glob(f"term{term}_week*.json"):
        match = re.fullmatch(rf"term{term}_week(\d+)\.json", path.name)
        if match and first <= int(match.group(1)) <= last:
            found.append((int(match.group(1)), json.loads(path.read_text())))
    return sorted(found, key=lambda item: item[0])


def session_lists(session):
    """(list name, words) pairs for one session, in slide order."""
    learned = session.get("learned_words") or {}
    new_lw = learned.get("new") or {}
    return [
        ("grid", [w["word"] for w in session.get("words_to_read_new", [])]),
        ("spell_new", session.get("words_to_spell_new", []) or []),
        ("read_review", (session.get("words_to_read_review") or {}).get("words", [])),
        ("spell_review", [w["word"] for w in session.get("words_to_spell_review", [])]),
        ("lw_review", [w["word"] for w in learned.get("review", [])]),
        ("lw_new", [new_lw["word"]] if new_lw.get("word") else []),
        ("dictation", [t for d in session.get("dictation", []) for t in d.get("targets", [])]),
    ]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--term", type=int, default=3)
    parser.add_argument("--from", dest="first", type=int, default=1)
    parser.add_argument("--to", dest="last", type=int, default=99)
    parser.add_argument("--index", action="store_true", help="print word -> uses instead of week detail")
    args = parser.parse_args()

    specs = week_specs(args.term, args.first, args.last)
    if not specs:
        raise SystemExit(f"no term{args.term}_week*.json specs in {WEEKS}")

    if args.index:
        uses = {}
        for num, week in specs:
            for session in week["sessions"]:
                where = f"W{num} {DAY.get(session['day'], session['day'])}"
                for name, words in session_lists(session):
                    for word in words:
                        uses.setdefault(word.lower(), []).append(f"{where} {name}")
        for word in sorted(uses):
            print(f"{word:22} {'; '.join(uses[word])}")
        return

    for num, week in specs:
        print(f"\n===== Term {week['term']} Week {num} =====")
        for session in week["sessions"]:
            focus = (session.get("new_morphology") or {}).get("morph") or "-"
            grammar = session.get("grammar", {}).get("i_do", {}).get("title", "")
            print(f"-- {session['day']} [{session.get('type', 'new')}] focus: {focus} | grammar: {grammar}")
            for name, words in session_lists(session):
                if words:
                    print(f"   {name:13} {', '.join(words)}")
            bank = [item["morph"] for item in session.get("sound_bank", [])]
            cards = [card["morph"] for card in session.get("morphology_review", [])]
            print(f"   {'cards':13} {', '.join(cards)}")
            print(f"   {'sound_bank':13} {', '.join(bank)}")


if __name__ == "__main__":
    main()
