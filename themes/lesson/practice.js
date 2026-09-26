"use strict";

/**
 * Practice rounds (megaprompt 82).
 *
 * A `practice` slide in a spec is a run of 3 to 8 quick items of the same
 * kind. Every student answers every item (boards, fingers, pointing or a
 * choral answer), the teacher scans, and the answer is revealed on click.
 * It is how a lesson reaches a high response rate without crowding a slide:
 * one item per slide, hero-sized, and the notes are written here from the
 * school cue scripts (megaprompt 75a) so every round runs the same way.
 *
 * expandSpec() replaces each practice slide with one concrete slide per item
 * (heroVisual, textExtract or content) before the deck is built.
 */

const ROUTINES = {
  boards: {
    full: "Cue: Write it... Chin it... Show me.",
    short: "Cue: boards up on cue.",
    where: "on your board",
    scan: "SCAN every board, back row first.",
    protection: "boards are scanned",
  },
  fingers: {
    full: "Cue: voices off. Fingers at your chest... show me.",
    short: "Cue: fingers on cue, voices off.",
    where: "on your fingers",
    scan: "SCAN every hand, back row first.",
    protection: "fingers are scanned",
  },
  point: {
    full: "Cue: voices off. Point... now. Hold your point.",
    short: "Cue: point on cue, voices off.",
    where: "by pointing",
    scan: "SCAN every point, back row first.",
    protection: "pointing is scanned",
  },
  choral: {
    full: "Cue: Everyone, together, on three.",
    short: "Cue: together, on three.",
    where: "out loud together",
    scan: "LISTEN for every voice; re-cue if it is not everyone.",
    protection: "the choral answer",
  },
};

function isYouDo(slide) {
  return /you do/i.test(String(slide.badge || ""));
}

function itemNotes(round, item, i) {
  const n = round.notes || {};
  // Tolerate a bad routine here; the validator names it.
  const routine = ROUTINES[round.routine] || ROUTINES.boards;
  const count = round.items.length;
  const last = i === count - 1;
  const think = round.thinkTime || 10;

  const beats = [];
  if (i === 0) {
    // The opening may be one line or several short lines of speech.
    const say = n.say != null ? [].concat(n.say) : (isYouDo(round)
      ? ["This round is just for you.", `Answer every one ${routine.where}, with no talking.`]
      : [`Let's practise. ${count} quick ones,`, `and everyone answers every one ${routine.where}.`]);
    beats.push([`SAY: ${say[0]}`, ...say.slice(1)]);
  }
  beats.push([
    `ASK: ${round.ask || round.title}`,
    `${think} sec. ${i === 0 ? routine.full : routine.short}`,
    `EXPECT: ${item.expect || item.answer}`,
  ]);
  beats.push([
    routine.scan,
    i === 0 && round.followUp
      ? `80%+ -> ${round.followUp}, then reveal.`
      : `80%+ -> reveal, ${last ? "then move on." : "then the next one."}`,
    `Less -> ${item.pivot || round.pivot}, re-ask.`,
  ]);
  beats.push([
    `REVEAL after ${routine.protection}.`,
    `SAY: ${item.say || (last ? "Tick it or fix it." : "Tick it or fix it, then get ready for the next one.")}`,
  ]);

  const out = {
    answer: item.answer,
    beats,
    trap: n.trap,
    prep: i === 0 ? n.prep : `Practice round, item ${i + 1} of ${count}. Keep it brisk.`,
    tag: n.tag,
  };
  // STRETCH and HELP ride on the second item: the first carries the round's set-up,
  // and by the second the teacher can see who needs which.
  if (i === Math.min(1, count - 1)) {
    if (n.stretch) out.stretch = n.stretch;
    if (n.help) out.help = n.help;
  }
  return out;
}

function expandPractice(authoredRound) {
  // Malformed items are skipped here and reported by the validator.
  const round = Object.assign({}, authoredRound, { items: authoredRound.items.filter((it) => it && typeof it === "object") });
  return round.items.map((item, i) => {
    const base = {
      badge: round.badge || "We Do",
      title: round.title,
      reveal: { answers: item.reveal ? [].concat(item.reveal) : [item.answer] },
      notes: itemNotes(round, item, i),
      practiceRound: { title: round.title, index: i, count: round.items.length, youDo: isYouDo(round) },
    };
    if (round.badgeColor) base.badgeColor = round.badgeColor;
    if (item.extract) {
      return Object.assign(base, { kind: "textExtract", extract: item.extract, source: item.source, highlights: item.highlights });
    }
    if (item.visual) {
      return Object.assign(base, { kind: "heroVisual", visual: item.visual, label: item.label });
    }
    return Object.assign(base, { kind: "content", lines: [item.text] });
  });
}

/** A copy of the spec with every practice slide replaced by its item slides. */
function expandSpec(spec) {
  const slides = [];
  (spec.slides || []).forEach((s) => {
    if (s && s.kind === "practice" && Array.isArray(s.items)) slides.push(...expandPractice(s));
    else slides.push(s);
  });
  return Object.assign({}, spec, { slides });
}

module.exports = { ROUTINES, expandSpec, expandPractice };
