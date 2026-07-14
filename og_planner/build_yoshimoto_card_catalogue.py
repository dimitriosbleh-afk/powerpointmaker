#!/usr/bin/env python3
"""Build readable and machine-readable catalogues from photographed OG cards."""

import json
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SETS = (
    ("Suffixes", "suffix", "yoshimoto_cards_suffixes.json"),
    ("Prefixes", "prefix", "yoshimoto_cards_prefixes.json"),
    ("Latin roots", "root", "yoshimoto_cards_latin_roots.json"),
)


def load_sets():
    loaded = []
    for title, card_type, filename in SETS:
        data = json.loads((ROOT / filename).read_text(encoding="utf-8"))
        loaded.append((title, card_type, filename, data["cards"]))
    return loaded


def build_markdown(loaded):
    total = sum(len(cards) for _, _, _, cards in loaded)
    lines = [
        "# Yoshimoto OG Card Catalogue",
        "",
        ("Direct transcription of the teacher's photographed 2007 Ronald H. "
         "Yoshimoto card set. The printed card meaning is authoritative. Keywords "
         "are school-friendly teaching selections; associated words use Australian "
         "English spelling."),
        "",
        f"Total: **{total} cards** (89 suffixes, 104 prefixes, 84 Latin roots).",
        "",
        ("Words under **Excluded from default lessons** were printed on a card but "
         "must not be selected automatically for student-facing materials."),
        "",
    ]

    for title, card_type, filename, cards in loaded:
        lines.extend([f"## {title} ({len(cards)})", ""])
        for card in cards:
            lines.append(f"### {card['morpheme']}")
            lines.append("")
            lines.append(f"- Meaning: {card['meaning']}")
            if card.get("part_of_speech"):
                lines.append(f"- Part of speech: {', '.join(card['part_of_speech'])}")
            if card.get("origin"):
                lines.append(f"- Origin: {card['origin']}")
            lines.append(f"- Keyword: {card.get('keyword') or 'none recorded'}")
            words = card.get("associated_words", [])
            lines.append("- Associated words: " + (", ".join(words) if words else "none printed"))
            excluded = card.get("excluded_words", [])
            if excluded:
                formatted = "; ".join(
                    f"{item['word']} ({item['reason']})" for item in excluded
                )
                lines.append(f"- Excluded from default lessons: {formatted}")
            if card.get("notes"):
                lines.append(f"- Note: {card['notes']}")
            lines.append(f"- Source: {card['source_image']} ({filename})")
            lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def build_master_json(loaded):
    cards = []
    for _, card_type, filename, entries in loaded:
        for entry in entries:
            cards.append({"type": card_type, "catalogue_file": filename, **entry})
    return {
        "description": "Photographed Yoshimoto OG cards, Australian-English lesson catalogue",
        "authority_order": [
            "photographed card transcription",
            "legacy morpheme bank",
            "unconfirmed reference meaning catalogue",
        ],
        "card_count": len(cards),
        "cards": cards,
    }


def main():
    loaded = load_sets()
    (ROOT / "YOSHIMOTO_CARD_CATALOGUE.md").write_text(
        build_markdown(loaded), encoding="utf-8"
    )
    (ROOT / "yoshimoto_cards_master.json").write_text(
        json.dumps(build_master_json(loaded), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"Wrote {sum(len(cards) for _, _, _, cards in loaded)} cards.")


if __name__ == "__main__":
    main()
