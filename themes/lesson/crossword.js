"use strict";

/**
 * Crossword resources for lesson specs (docs/lesson-spec.md, kind "crossword").
 *
 *   { "kind": "crossword", "label": "Persuasion Crossword",
 *     "words": [{ "answer": "REBUTTAL", "clue": "Answering the other side" }, ...],
 *     "wordBank": true }
 *
 * The grid is laid out automatically and deterministically: longest word
 * first, then each word placed where it crosses the most letters, with no
 * two words touching side by side. Every word must cross another; a word
 * that cannot be placed is a build error naming it, so the author swaps it
 * rather than shipping a broken puzzle. The answer key uses the same grid.
 */

const P = require("../pdf_helpers");

function cleanAnswer(a) {
  return String(a || "").toUpperCase().replace(/[^A-Z]/g, "");
}

/** Lay the words out on a grid. Returns { placed: [{ word, clue, row, col, dir }], failed: [words] }. */
function layoutCrossword(words, keepFirst) {
  const SIZE = 40;
  const grid = new Map(); // "r,c" -> letter
  const key = (r, c) => `${r},${c}`;
  const at = (r, c) => grid.get(key(r, c));
  const list = words.map((w, i) => ({ word: cleanAnswer(w.answer), clue: w.clue, index: i }))
    .sort((a, b) => (keepFirst && (a.index === 0 || b.index === 0) ? a.index - b.index : 0) || b.word.length - a.word.length || a.index - b.index);

  function fits(word, row, col, dir) {
    const dr = dir === "down" ? 1 : 0;
    const dc = dir === "across" ? 1 : 0;
    if (at(row - dr, col - dc) || at(row + dr * word.length, col + dc * word.length)) return -1;
    let crossings = 0;
    for (let i = 0; i < word.length; i += 1) {
      const r = row + dr * i;
      const c = col + dc * i;
      const cur = at(r, c);
      if (cur) {
        if (cur !== word[i]) return -1;
        crossings += 1;
      } else {
        // No letters beside an empty cell, or two words would run together.
        if (dir === "across" && (at(r - 1, c) || at(r + 1, c))) return -1;
        if (dir === "down" && (at(r, c - 1) || at(r, c + 1))) return -1;
      }
    }
    return crossings;
  }

  function place(item, row, col, dir) {
    const dr = dir === "down" ? 1 : 0;
    const dc = dir === "across" ? 1 : 0;
    for (let i = 0; i < item.word.length; i += 1) grid.set(key(row + dr * i, col + dc * i), item.word[i]);
    placed.push({ word: item.word, clue: item.clue, row, col, dir });
  }

  const placed = [];
  let pending = list.slice();
  if (!pending.length) return { placed, failed: [] };
  const first = pending.shift();
  place(first, SIZE / 2, SIZE / 2 - Math.floor(first.word.length / 2), "across");

  // Keep passing over the unplaced words while progress is made.
  let progress = true;
  while (pending.length && progress) {
    progress = false;
    const still = [];
    pending.forEach((item) => {
      let best = null;
      placed.forEach((p) => {
        for (let i = 0; i < p.word.length; i += 1) {
          const pr = p.row + (p.dir === "down" ? i : 0);
          const pc = p.col + (p.dir === "across" ? i : 0);
          for (let j = 0; j < item.word.length; j += 1) {
            if (item.word[j] !== p.word[i]) continue;
            const dir = p.dir === "across" ? "down" : "across";
            const row = dir === "down" ? pr - j : pr;
            const col = dir === "across" ? pc - j : pc;
            const score = fits(item.word, row, col, dir);
            if (score <= 0) continue;
            // Prefer more crossings, then a placement nearer the centre (a compact grid).
            const spread = Math.abs(row - SIZE / 2) + Math.abs(col - SIZE / 2);
            const rank = score * 100 - spread;
            if (!best || rank > best.rank) best = { row, col, dir, rank };
          }
        }
      });
      if (best) { place(item, best.row, best.col, best.dir); progress = true; } else still.push(item);
    });
    pending = still;
  }
  return { placed, failed: pending.map((p) => p.word) };
}

/** Number the grid the usual way: a cell that starts an across or down word gets the next number. */
function numberCrossword(placed) {
  const rows = placed.flatMap((p) => [p.row, p.row + (p.dir === "down" ? p.word.length - 1 : 0)]);
  const cols = placed.flatMap((p) => [p.col, p.col + (p.dir === "across" ? p.word.length - 1 : 0)]);
  const minR = Math.min(...rows); const maxR = Math.max(...rows);
  const minC = Math.min(...cols); const maxC = Math.max(...cols);
  const cells = new Map();
  placed.forEach((p) => {
    for (let i = 0; i < p.word.length; i += 1) {
      const r = p.row + (p.dir === "down" ? i : 0) - minR;
      const c = p.col + (p.dir === "across" ? i : 0) - minC;
      cells.set(`${r},${c}`, p.word[i]);
    }
  });
  const starts = placed.map((p) => ({ ...p, r: p.row - minR, c: p.col - minC }))
    .sort((a, b) => a.r - b.r || a.c - b.c);
  const numbers = new Map();
  let n = 0;
  starts.forEach((s) => {
    const k = `${s.r},${s.c}`;
    if (!numbers.has(k)) { n += 1; numbers.set(k, n); }
    s.number = numbers.get(k);
  });
  return { cells, numbers, starts, rows: maxR - minR + 1, cols: maxC - minC + 1 };
}

/** Try each word as the starting word (deterministic); keep the most compact full layout. */
function bestLayout(words) {
  let best = null;
  for (let start = 0; start < words.length; start += 1) {
    const order = [words[start]].concat(words.filter((_, i) => i !== start));
    const result = layoutCrossword(order, true);
    const area = result.placed.length ? numberCrossword(result.placed) : null;
    const score = result.failed.length * 10000 + (area ? area.rows * area.cols : 0);
    if (!best || score < best.score) best = { score, result };
  }
  return best.result;
}

function buildCrossword(resource) {
  const words = (resource.words || []).filter((w) => cleanAnswer(w.answer));
  const { placed, failed } = bestLayout(words);
  if (failed.length) {
    throw new Error(`[crossword] "${resource.label}": could not fit ${failed.join(", ")}. Swap or add a word that shares letters with the others.`);
  }
  return numberCrossword(placed);
}

/** Draw the puzzle (or, with isKey, the answers) from y. Returns y after the clues. */
function drawCrossword(doc, resource, y, ctx) {
  const { C, isKey, S } = ctx;
  const primary = P.hex(C.PRIMARY);
  const cw = buildCrossword(resource);
  const x0 = P.PAGE.MARGIN;
  const maxGridH = resource.gridHeight || 360;
  const cell = Math.min(28, P.PAGE.CONTENT_W / cw.cols, maxGridH / cw.rows);
  const gx = x0 + (P.PAGE.CONTENT_W - cell * cw.cols) / 2;
  cw.cells.forEach((letter, k) => {
    const [r, c] = k.split(",").map(Number);
    const cx = gx + c * cell;
    const cy = y + r * cell;
    doc.save().lineWidth(1).strokeColor(P.hex(C.CHARCOAL)).rect(cx, cy, cell, cell).stroke().restore();
    const num = cw.numbers.get(k);
    if (num) doc.save().font("Sans").fontSize(Math.max(6, cell * 0.3)).fillColor(P.hex(C.CHARCOAL)).text(String(num), cx + 1.5, cy + 1, { lineBreak: false }).restore();
    if (isKey) {
      doc.save().font("Sans-Bold").fontSize(cell * 0.55).fillColor(primary)
        .text(letter, cx, cy + cell * 0.28, { width: cell, align: "center", lineBreak: false }).restore();
    }
  });
  y += cw.rows * cell + 14;

  if (resource.wordBank && !isKey) {
    const bank = cw.starts.map((s) => s.word).sort();
    doc.save().font("Sans-Bold").fontSize(S.body - 2).fillColor(primary).text("Word bank:", x0, y, { continued: true })
      .font("Sans").fillColor(P.hex(C.CHARCOAL)).text(`  ${bank.join("   ").toLowerCase()}`, { width: P.PAGE.CONTENT_W }).restore();
    y = doc.y + 10;
  }

  // Clues in two columns: Across | Down.
  const colW = (P.PAGE.CONTENT_W - 20) / 2;
  const clueSize = S.body - 2;
  ["across", "down"].forEach((dir, i) => {
    const cx = x0 + i * (colW + 20);
    let cy = y;
    doc.save().font("Sans-Bold").fontSize(S.body).fillColor(primary).text(dir === "across" ? "Across" : "Down", cx, cy).restore();
    cy += S.body + 6;
    cw.starts.filter((s) => s.dir === dir).sort((a, b) => a.number - b.number).forEach((s) => {
      const line = `${s.number}. ${s.clue} (${s.word.length})${isKey ? ` - ${s.word.toLowerCase()}` : ""}`;
      doc.save().font("Sans").fontSize(clueSize).fillColor(P.hex(C.CHARCOAL)).text(line, cx, cy, { width: colW }).restore();
      cy = doc.y + 4;
    });
    if (i === 0) ctx.acrossEnd = cy; else ctx.downEnd = cy;
  });
  return Math.max(ctx.acrossEnd || y, ctx.downEnd || y) + 8;
}

module.exports = { layoutCrossword, numberCrossword, buildCrossword, drawCrossword, cleanAnswer };
