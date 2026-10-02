"use strict";

/**
 * Geometry and algorithm diagrams drawn once as SVG and rendered to PNG, so
 * slides and paper show exactly the same figure.
 *
 *   { type: "angle", rays: [0, 65], arcs: [{ from: 0, to: 65, label: "?" }], protractor: true }
 *   { type: "angle", rays: [0, 130, 180], arcs: [{ from: 0, to: 130, label: "130°" }, { from: 130, to: 180, label: "x" }] }
 *   { type: "columnSum", numbers: [34567, 12345], op: "+" }
 *
 * Angles: `rays` are directions in degrees, anticlockwise from pointing
 * right, from one vertex. `arcs` mark the angle between two ray directions
 * (anticlockwise from `from` to `to`) with an optional label; `right: true`
 * draws a square marker. `rotate` turns the whole figure. `protractor`
 * overlays a protractor on the first ray (or along `protractorAt`), with an
 * outer and an inner scale.
 * An angle is the same size at any scale, so a printed diagram measures
 * correctly with a real protractor.
 */

const { resolveSansFamily } = require("./fonts");

let Resvg = null;
function getResvg() {
  if (!Resvg) ({ Resvg } = require("@resvg/resvg-js"));
  return Resvg;
}

const ARM = 100;
const rad = (d) => (d * Math.PI) / 180;
const pt = (deg, r) => ({ x: r * Math.cos(rad(deg)), y: -r * Math.sin(rad(deg)) });
const f2 = (n) => Math.round(n * 100) / 100;
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function hexColour(c, fallback) {
  const v = String(c || fallback || "333333").replace(/^#/, "");
  return `#${v}`;
}

function angleSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const rot = Number(spec.rotate) || 0;
  const rays = (Array.isArray(spec.rays) && spec.rays.length ? spec.rays : [0, 60]).map((d) => Number(d) + rot);
  let arcs = Array.isArray(spec.arcs) ? spec.arcs : (rays.length === 2 ? [{ from: spec.rays ? spec.rays[0] : 0, to: spec.rays ? spec.rays[1] : 60 }] : []);
  arcs = arcs.map((a) => ({ ...a, from: Number(a.from) + rot, to: Number(a.to) + rot }));
  // Arms reach past the protractor so the model shows how to read it.
  const arm = spec.protractor ? 118 : ARM;
  const els = [];
  const pts = [{ x: 0, y: 0 }];
  const grow = (p) => pts.push(p);

  // Protractor under the figure, baseline on the first ray.
  if (spec.protractor) {
    // The protractor's flat edge lies along `protractorAt` (default the first ray),
    // curved side anticlockwise from it. A left-pointing arm is read on the outer scale.
    const base = spec.protractorAt != null ? Number(spec.protractorAt) + rot : rays[0];
    const R = 92;
    const semi = [];
    for (let d = 0; d <= 180; d += 3) semi.push(pt(base + d, R));
    const edge = semi.map((p, i) => `${i ? "L" : "M"}${f2(p.x)},${f2(p.y)}`).join(" ");
    els.push(`<path d="${edge} Z" fill="#DCEBFA" fill-opacity="0.85" stroke="#5A7FA8" stroke-width="1"/>`);
    for (let d = 0; d <= 180; d += 1) {
      const len = d % 10 === 0 ? 9 : (d % 5 === 0 ? 6 : 3.5);
      const a = pt(base + d, R);
      const b = pt(base + d, R - len);
      els.push(`<line x1="${f2(a.x)}" y1="${f2(a.y)}" x2="${f2(b.x)}" y2="${f2(b.y)}" stroke="#34506F" stroke-width="${d % 10 === 0 ? 0.9 : 0.5}"/>`);
      if (d % 10 === 0) {
        // Inner scale starts at 0 on the base ray; outer scale starts at 0 on the far side.
        // The end labels lean in a few degrees so they sit above the baseline.
        const at = base + (d === 0 ? 4 : (d === 180 ? 176 : d));
        const outer = pt(at, R - 16);
        const inner = pt(at, R - 31);
        els.push(`<text x="${f2(outer.x)}" y="${f2(outer.y)}" font-size="6.2" fill="#34506F" text-anchor="middle" dominant-baseline="central">${180 - d}</text>`);
        els.push(`<text x="${f2(inner.x)}" y="${f2(inner.y)}" font-size="6.2" font-weight="bold" fill="#1F3550" text-anchor="middle" dominant-baseline="central">${d}</text>`);
      }
    }
    const l = pt(base + 180, R);
    const r = pt(base, R);
    els.push(`<line x1="${f2(l.x)}" y1="${f2(l.y)}" x2="${f2(r.x)}" y2="${f2(r.y)}" stroke="#5A7FA8" stroke-width="1"/>`);
    semi.forEach(grow);
    grow(l); grow(r);
  }

  // Arcs and their labels.
  arcs.forEach((a, i) => {
    let sweep = a.to - a.from;
    while (sweep <= 0) sweep += 360;
    // Reflex arcs sit a little wider so the long way round is unmistakable.
    // Adjacent arcs never overlap, so one radius reads cleanest; reflex arcs sit wider.
    const r = a.r || (spec.protractor ? 16 : (sweep > 180 ? 26 : 22));
    const isRight = a.right === true;
    // Arcs count towards the figure's bounds, or a reflex arc is cropped.
    for (let t = 0; t <= sweep; t += 10) grow(pt(a.from + t, r + 3));
    grow(pt(a.from + sweep, r + 3));
    if (isRight) {
      const s = 11;
      const p1 = pt(a.from, s);
      const p2 = pt(a.from + 45, s * Math.SQRT2);
      const p3 = pt(a.from + 90, s);
      els.push(`<path d="M${f2(p1.x)},${f2(p1.y)} L${f2(p2.x)},${f2(p2.y)} L${f2(p3.x)},${f2(p3.y)}" fill="none" stroke="${accent}" stroke-width="1.8"/>`);
    } else {
      const s = pt(a.from, r);
      const e = pt(a.from + sweep, r);
      const large = sweep > 180 ? 1 : 0;
      els.push(`<path d="M0,0 L${f2(s.x)},${f2(s.y)} A${r},${r} 0 ${large} 0 ${f2(e.x)},${f2(e.y)} Z" fill="${accent}" fill-opacity="0.14" stroke="none"/>`);
      els.push(`<path d="M${f2(s.x)},${f2(s.y)} A${r},${r} 0 ${large} 0 ${f2(e.x)},${f2(e.y)}" fill="none" stroke="${accent}" stroke-width="1.8"/>`);
    }
    if (a.label != null && a.label !== "") {
      // Labels sit well clear of the vertex (further for narrow angles) and
      // large enough to read when the figure spans a full turn.
      const mid = a.from + sweep / 2;
      const lr = Math.max((isRight ? 16 : r) + 20, sweep < 40 ? 58 : 46);
      const p = pt(mid, lr);
      // Trebuchet sets the degree sign wide of its number; pull it in.
      const label = esc(a.label).replace(/°/g, '<tspan dx="-3">°</tspan>');
      els.push(`<text x="${f2(p.x)}" y="${f2(p.y)}" font-size="15" font-weight="bold" fill="${ink}" text-anchor="middle" dominant-baseline="central">${label}</text>`);
      grow({ x: p.x - 22, y: p.y - 10 }); grow({ x: p.x + 22, y: p.y + 10 });
    }
  });

  // A lone arm is a drawing task: reserve the half-turn above it to draw in.
  if (rays.length === 1) for (let d = 0; d <= 180; d += 15) grow(pt(rays[0] + d, arm));

  // Arms and vertex on top.
  rays.forEach((d) => {
    const e = pt(d, arm);
    els.push(`<line x1="0" y1="0" x2="${f2(e.x)}" y2="${f2(e.y)}" stroke="${ink}" stroke-width="2.4" stroke-linecap="round"/>`);
    grow(e);
  });
  els.push(`<circle cx="0" cy="0" r="2.4" fill="${ink}"/>`);

  const pad = 8;
  const minX = Math.min(...pts.map((p) => p.x)) - pad;
  const maxX = Math.max(...pts.map((p) => p.x)) + pad;
  const minY = Math.min(...pts.map((p) => p.y)) - pad;
  const maxY = Math.max(...pts.map((p) => p.y)) + pad;
  const w = maxX - minX;
  const h = maxY - minY;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${f2(minX)} ${f2(minY)} ${f2(w)} ${f2(h)}" width="${f2(w)}" height="${f2(h)}">${els.join("")}</svg>`;
  return { svg, aspect: w / h };
}

function columnSumSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const nums = (spec.numbers || []).map((n) => String(n).replace(/\s+/g, ""));
  const op = spec.op || "+";
  const cols = Math.max(...nums.map((n) => n.length), 1) + 1; // one spare column for a carried digit
  const cell = 26;
  const rowH = 36;
  const left = 30;
  const els = [];
  nums.forEach((n, r) => {
    const digits = n.split("");
    digits.forEach((d, i) => {
      const col = cols - digits.length + i;
      els.push(`<text x="${left + col * cell + cell / 2}" y="${(r + 1) * rowH - 8}" font-size="28" fill="${ink}" text-anchor="middle">${esc(d)}</text>`);
    });
  });
  const opY = nums.length * rowH - 8;
  els.push(`<text x="${left - 8}" y="${opY}" font-size="28" font-weight="bold" fill="${accent}" text-anchor="middle">${esc(op)}</text>`);
  const lineY = nums.length * rowH + 4;
  els.push(`<line x1="${left - 18}" y1="${lineY}" x2="${left + cols * cell}" y2="${lineY}" stroke="${ink}" stroke-width="2.4"/>`);
  const lineY2 = lineY + rowH + 6;
  els.push(`<line x1="${left - 18}" y1="${lineY2}" x2="${left + cols * cell}" y2="${lineY2}" stroke="${ink}" stroke-width="2.4"/>`);
  const w = left + cols * cell + 10;
  const h = lineY2 + 8;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="-10 0 ${w + 10} ${h}" width="${w + 10}" height="${h}">${els.join("")}</svg>`;
  return { svg, aspect: (w + 10) / h };
}

/* ── Shared helpers for the data and space diagrams ─────────────────────── */

const text = (x, y, s, o) => {
  const opt = o || {};
  const attrs = [
    `x="${f2(x)}"`, `y="${f2(y)}"`, `font-size="${opt.size || 11}"`,
    `fill="${opt.fill || "#2B2B2B"}"`, `text-anchor="${opt.anchor || "middle"}"`,
    `dominant-baseline="${opt.baseline || "central"}"`,
  ];
  if (opt.bold) attrs.push('font-weight="bold"');
  if (opt.rotate) attrs.push(`transform="rotate(${opt.rotate} ${f2(x)} ${f2(y)})"`);
  return `<text ${attrs.join(" ")}>${esc(s)}</text>`;
};

function wrapSvg(els, minX, minY, maxX, maxY, pad) {
  const p = pad == null ? 6 : pad;
  const x0 = minX - p;
  const y0 = minY - p;
  const w = maxX - minX + 2 * p;
  const h = maxY - minY + 2 * p;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${f2(x0)} ${f2(y0)} ${f2(w)} ${f2(h)}" width="${f2(w)}" height="${f2(h)}">${els.join("")}</svg>`;
  return { svg, aspect: w / h };
}

/** Round a number for a label: no float noise, ASCII minus. */
function num(n) {
  const r = Math.round(n * 1000) / 1000;
  return String(r).replace(/^-/, "-");
}

/** A nice axis top and step for a set of values. */
function niceScale(maxValue, spec) {
  const start = Number(spec.yStart) || 0;
  const range = Math.max(1, (Number(maxValue) || 1) - start);
  if (spec.yMax && spec.yStep) return { top: spec.yMax, step: spec.yStep };
  const rough = range / 5;
  const mag = Math.pow(10, Math.floor(Math.log10(rough)));
  const step = spec.yStep || [1, 2, 5, 10].map((m) => m * mag).find((s) => s >= rough) || mag * 10;
  const top = spec.yMax || start + Math.ceil(range / step) * step;
  return { top, step };
}

/* ── Coordinate grids ───────────────────────────────────────────────────
 * { type: "grid", x: [0, 10], y: [0, 10], points: [{ x: 3, y: 4, label: "A" }] }
 * { type: "grid", x: [-5, 5], y: [-5, 5], polygon: [[1,1],[4,1],[4,3]] }      four quadrants
 * { type: "grid", reference: true, cols: 6, rows: 5, cells: [{ col: "C", row: 2, label: "tree" }] }
 * Coordinates sit on the LINES; a grid reference names the SPACES.
 */
function gridSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const els = [];
  const cell = 22;
  if (spec.reference) {
    const cols = spec.cols || 6;
    const rows = spec.rows || 6;
    const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    (spec.cells || []).forEach((c) => {
      const ci = letters.indexOf(String(c.col).toUpperCase());
      const ri = Number(c.row) - 1;
      if (ci < 0 || ri < 0) return;
      const x = ci * cell;
      const y = (rows - 1 - ri) * cell;
      if (c.shade !== false) els.push(`<rect x="${x}" y="${y}" width="${cell}" height="${cell}" fill="${accent}" fill-opacity="0.18"/>`);
      if (c.label) els.push(text(x + cell / 2, y + cell / 2, c.label, { size: c.label.length > 3 ? 7 : 10, bold: true, fill: ink }));
    });
    for (let i = 0; i <= cols; i += 1) els.push(`<line x1="${i * cell}" y1="0" x2="${i * cell}" y2="${rows * cell}" stroke="#8A97A8" stroke-width="${i === 0 || i === cols ? 1.4 : 0.8}"/>`);
    for (let j = 0; j <= rows; j += 1) els.push(`<line x1="0" y1="${j * cell}" x2="${cols * cell}" y2="${j * cell}" stroke="#8A97A8" stroke-width="${j === 0 || j === rows ? 1.4 : 0.8}"/>`);
    for (let i = 0; i < cols; i += 1) els.push(text(i * cell + cell / 2, rows * cell + 10, letters[i], { size: 11, bold: true, fill: ink }));
    for (let j = 0; j < rows; j += 1) els.push(text(-10, (rows - 1 - j) * cell + cell / 2, String(j + 1), { size: 11, bold: true, fill: ink }));
    return wrapSvg(els, -20, -4, cols * cell + 4, rows * cell + 20);
  }

  const xr = Array.isArray(spec.x) ? spec.x.map(Number) : [0, 10];
  const yr = Array.isArray(spec.y) ? spec.y.map(Number) : [0, 10];
  const step = Number(spec.step) > 0 ? Number(spec.step) : 1;
  const every = Number(spec.labelEvery) > 0 ? Number(spec.labelEvery) : 1;
  const nx = Math.round((xr[1] - xr[0]) / step);
  const ny = Math.round((yr[1] - yr[0]) / step);
  const X = (v) => ((v - xr[0]) / step) * cell;
  const Y = (v) => ((yr[1] - v) / step) * cell;
  const axisX = xr[0] < 0 && xr[1] > 0 ? 0 : xr[0];   // the y-axis sits at x = axisX
  const axisY = yr[0] < 0 && yr[1] > 0 ? 0 : yr[0];   // the x-axis sits at y = axisY
  const W = nx * cell;
  const H = ny * cell;
  // Bigger grids are drawn smaller on the page, so their text grows to stay readable.
  const fs = 11.5 * Math.min(1.4, Math.max(1, Math.max(nx, ny) / 7));

  for (let i = 0; i <= nx; i += 1) els.push(`<line x1="${i * cell}" y1="0" x2="${i * cell}" y2="${H}" stroke="#B9C3D0" stroke-width="0.7"/>`);
  for (let j = 0; j <= ny; j += 1) els.push(`<line x1="0" y1="${j * cell}" x2="${W}" y2="${j * cell}" stroke="#B9C3D0" stroke-width="0.7"/>`);
  // Axes, with arrowheads at the positive ends.
  els.push(`<line x1="0" y1="${f2(Y(axisY))}" x2="${W + 10}" y2="${f2(Y(axisY))}" stroke="${ink}" stroke-width="1.8"/>`);
  els.push(`<line x1="${f2(X(axisX))}" y1="${H}" x2="${f2(X(axisX))}" y2="-10" stroke="${ink}" stroke-width="1.8"/>`);
  els.push(`<path d="M${W + 14},${f2(Y(axisY))} l-7,-3.5 l0,7 Z" fill="${ink}"/>`);
  els.push(`<path d="M${f2(X(axisX))},-14 l-3.5,7 l7,0 Z" fill="${ink}"/>`);
  if (spec.axisLabels !== false) {
    els.push(text(W + 20, Y(axisY), "x", { size: 12, bold: true, fill: ink, anchor: "start" }));
    els.push(text(X(axisX), -22, "y", { size: 12, bold: true, fill: ink }));
  }
  // Numbers sit on the lines, beside the axes.
  for (let i = 0; i <= nx; i += 1) {
    const v = xr[0] + i * step;
    if (i % every !== 0 || (v === axisX && axisX !== xr[0])) continue;
    els.push(text(X(v), Y(axisY) + fs * 0.9, num(v), { size: fs, fill: ink }));
  }
  for (let j = 0; j <= ny; j += 1) {
    const v = yr[0] + j * step;
    if (j % every !== 0 || v === axisY) continue;
    els.push(text(X(axisX) - 5, Y(v) + (v < 0 ? 3 : 0), num(v), { size: fs, fill: ink, anchor: "end" }));
  }
  if (axisX !== xr[0] || axisY !== yr[0]) els.push(text(X(axisX) - 4, Y(axisY) + fs * 0.9, "0", { size: fs, fill: ink, anchor: "end" }));

  const toPx = (p) => (Array.isArray(p) ? { x: X(Number(p[0])), y: Y(Number(p[1])) } : { x: X(Number(p.x)), y: Y(Number(p.y)) });
  const polys = spec.polygons || (spec.polygon ? [spec.polygon] : []);
  polys.forEach((poly) => {
    const d = poly.map(toPx).map((p, i) => `${i ? "L" : "M"}${f2(p.x)},${f2(p.y)}`).join(" ");
    els.push(`<path d="${d} Z" fill="${accent}" fill-opacity="0.1" stroke="${accent}" stroke-width="2"/>`);
  });
  if (Array.isArray(spec.path) && spec.path.length > 1) {
    const d = spec.path.map(toPx).map((p, i) => `${i ? "L" : "M"}${f2(p.x)},${f2(p.y)}`).join(" ");
    els.push(`<path d="${d}" fill="none" stroke="${accent}" stroke-width="2"/>`);
  }
  (spec.arrows || []).forEach((a) => {
    const s = toPx(a.from);
    const e = toPx(a.to);
    const ang = Math.atan2(e.y - s.y, e.x - s.x);
    const hx = e.x - 7 * Math.cos(ang);
    const hy = e.y - 7 * Math.sin(ang);
    els.push(`<line x1="${f2(s.x)}" y1="${f2(s.y)}" x2="${f2(hx)}" y2="${f2(hy)}" stroke="${ink}" stroke-width="1.6" stroke-dasharray="4 3"/>`);
    const l = { x: hx + 4 * Math.sin(ang), y: hy - 4 * Math.cos(ang) };
    const r = { x: hx - 4 * Math.sin(ang), y: hy + 4 * Math.cos(ang) };
    els.push(`<path d="M${f2(e.x)},${f2(e.y)} L${f2(l.x)},${f2(l.y)} L${f2(r.x)},${f2(r.y)} Z" fill="${ink}"/>`);
  });
  const offsets = { ne: [6, -10, "start"], nw: [-6, -10, "end"], se: [6, 11, "start"], sw: [-6, 11, "end"] };
  (spec.points || []).forEach((p) => {
    const q = toPx(p);
    els.push(`<circle cx="${f2(q.x)}" cy="${f2(q.y)}" r="${f2(4 * fs / 11.5)}" fill="${accent}"/>`);
    const name = p.label != null ? String(p.label) : "";
    const coord = spec.showCoords || p.showCoord ? `(${num(p.x)}, ${num(p.y)})` : "";
    const lab = [name, coord].filter(Boolean).join(" ");
    if (lab) {
      const [dx, dy, anchor] = offsets[p.pos || "ne"] || offsets.ne;
      els.push(text(q.x + dx, q.y + dy, lab, { size: (lab.length > 3 ? 12 : 15) * fs / 11.5, bold: true, fill: ink, anchor }));
    }
  });
  return wrapSvg(els, -10 - 2 * fs, -30, W + 30, H + 8 + fs * 1.4);
}

/* ── Graphs ─────────────────────────────────────────────────────────────
 * { type: "chart", style: "bar", categories: ["Footy", "Netball"], values: [8, 5], title, xLabel, yLabel }
 * style "line" joins the points; style "dot" stacks one dot per count (a dot plot).
 * yStart > 0 starts the axis above zero (for spotting a misleading graph).
 * series: [{ name, values }] draws grouped bars with a key.
 */
const SERIES_FILLS = ["#4A7FC1", "#E39B3A", "#5BA56B", "#B565A7"];

function chartSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "4A7FC1");
  const cats = (spec.categories || []).map(String);
  const series = Array.isArray(spec.series) && spec.series.length
    ? spec.series.map((s, i) => ({ name: s.name, values: s.values.map(Number), fill: SERIES_FILLS[i % SERIES_FILLS.length] }))
    : [{ name: "", values: (spec.values || []).map(Number), fill: accent }];
  const style = spec.style || "bar";
  const els = [];
  const slot = Math.max(30, Math.min(56, 300 / Math.max(cats.length, 1)));
  const W = slot * cats.length;
  const H = 150;
  const all = series.reduce((a, s) => a.concat(s.values), []);

  if (style === "dot") {
    const r = 5.5;
    const maxCount = Math.max(...all, 1);
    const Hd = maxCount * (2 * r + 2);
    els.push(`<line x1="0" y1="${Hd}" x2="${W}" y2="${Hd}" stroke="${ink}" stroke-width="1.6"/>`);
    cats.forEach((c, i) => {
      const cx = i * slot + slot / 2;
      els.push(`<line x1="${cx}" y1="${Hd}" x2="${cx}" y2="${Hd + 4}" stroke="${ink}" stroke-width="1"/>`);
      els.push(text(cx, Hd + 14, c, { size: 12.5, fill: ink }));
      for (let k = 0; k < (series[0].values[i] || 0); k += 1) {
        els.push(`<circle cx="${f2(cx)}" cy="${f2(Hd - r - 1 - k * (2 * r + 2))}" r="${r}" fill="${accent}"/>`);
      }
    });
    if (spec.title) els.push(text(W / 2, -16, spec.title, { size: 14.5, bold: true, fill: ink }));
    if (spec.xLabel) els.push(text(W / 2, Hd + 33, spec.xLabel, { size: 12.5, bold: true, fill: ink }));
    return wrapSvg(els, -6, spec.title ? -28 : -4, W + 6, Hd + (spec.xLabel ? 42 : 24));
  }

  const start = Number(spec.yStart) || 0;
  const { top, step } = niceScale(Math.max(...all, start + 1), spec);
  const Y = (v) => H - ((v - start) / (top - start)) * H;
  for (let v = start; v <= top + 1e-9; v += step) {
    els.push(`<line x1="0" y1="${f2(Y(v))}" x2="${W}" y2="${f2(Y(v))}" stroke="#D3DAE3" stroke-width="0.8"/>`);
    els.push(text(-6, Y(v), num(v), { size: 12, fill: ink, anchor: "end" }));
  }
  if (style === "line") {
    series.forEach((s) => {
      const pts = s.values.map((v, i) => ({ x: i * slot + slot / 2, y: Y(v) }));
      els.push(`<path d="${pts.map((p, i) => `${i ? "L" : "M"}${f2(p.x)},${f2(p.y)}`).join(" ")}" fill="none" stroke="${s.fill}" stroke-width="2.4"/>`);
      pts.forEach((p) => els.push(`<circle cx="${f2(p.x)}" cy="${f2(p.y)}" r="3.4" fill="${s.fill}"/>`));
    });
  } else {
    const n = series.length;
    const groupW = slot * 0.68;
    const barW = groupW / n;
    series.forEach((s, si) => {
      s.values.forEach((v, i) => {
        const x = i * slot + (slot - groupW) / 2 + si * barW;
        const y = Y(Math.max(v, start));
        els.push(`<rect x="${f2(x)}" y="${f2(y)}" width="${f2(barW - (n > 1 ? 1.5 : 0))}" height="${f2(H - y)}" fill="${s.fill}"/>`);
        if (spec.showValues) els.push(text(x + barW / 2, y - 7, num(v), { size: 9, bold: true, fill: ink }));
      });
    });
  }
  els.push(`<line x1="0" y1="${H}" x2="${W}" y2="${H}" stroke="${ink}" stroke-width="1.6"/>`);
  els.push(`<line x1="0" y1="0" x2="0" y2="${H}" stroke="${ink}" stroke-width="1.6"/>`);
  cats.forEach((c, i) => {
    const cx = i * slot + slot / 2;
    const words = c.split(" ");
    if (c.length > 8 && words.length > 1) {
      const half = Math.ceil(words.length / 2);
      els.push(text(cx, H + 11, words.slice(0, half).join(" "), { size: 11, fill: ink }));
      els.push(text(cx, H + 23, words.slice(half).join(" "), { size: 11, fill: ink }));
    } else {
      els.push(text(cx, H + 12, c, { size: c.length > 7 ? 10.5 : 12.5, fill: ink }));
    }
  });
  let minX = -32;
  let maxY = H + 30;
  let maxX = W + 4;
  if (spec.yLabel) { els.push(text(-38, H / 2, spec.yLabel, { size: 12.5, bold: true, fill: ink, rotate: -90 })); minX = -48; }
  if (spec.xLabel) { els.push(text(W / 2, H + 38, spec.xLabel, { size: 12.5, bold: true, fill: ink })); maxY = H + 46; }
  let minY = -8;
  if (spec.title) { els.push(text(W / 2, -18, spec.title, { size: 14.5, bold: true, fill: ink })); minY = -30; }
  if (series.length > 1) {
    series.forEach((s, i) => {
      const ly = 6 + i * 16;
      els.push(`<rect x="${W + 10}" y="${ly - 5}" width="10" height="10" fill="${s.fill}"/>`);
      els.push(text(W + 24, ly, s.name, { size: 12, fill: ink, anchor: "start" }));
    });
    maxX = W + 24 + Math.max(...series.map((s) => String(s.name).length)) * 7;
  }
  return wrapSvg(els, minX, minY, maxX, maxY);
}

/* ── Spinners ───────────────────────────────────────────────────────────
 * { type: "spinner", sectors: [{ label: "red", weight: 2 }, { label: "blue" }], pointer: 40 }
 * Sectors run clockwise from the top. A colour word in the label sets the fill.
 */
const NAMED_FILLS = {
  red: "#F2A19A", blue: "#9EC3F2", green: "#A6DBA9", yellow: "#FBE18A", orange: "#F8BE86",
  purple: "#CDB3EF", pink: "#F6B5D1", white: "#FFFFFF", black: "#8C8C8C", grey: "#CFCFCF", gray: "#CFCFCF",
};
const CYCLE_FILLS = ["#9EC3F2", "#FBE18A", "#A6DBA9", "#F2A19A", "#CDB3EF", "#F8BE86"];

function spinnerSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const sectors = (spec.sectors || []).map((s) => (typeof s === "string" ? { label: s } : s));
  const total = sectors.reduce((t, s) => t + (Number(s.weight) || 1), 0) || 1;
  const R = 80;
  const els = [];
  // Clockwise from the top: a compass bearing b sits at maths angle 90 - b.
  const at = (bearing, r) => pt(90 - bearing, r);
  let b = 0;
  sectors.forEach((s, i) => {
    const sweep = (360 * (Number(s.weight) || 1)) / total;
    const word = String(s.label || "").toLowerCase().split(/\s+/).find((w) => NAMED_FILLS[w]);
    const fill = s.colour ? hexColour(s.colour) : (word ? NAMED_FILLS[word] : CYCLE_FILLS[i % CYCLE_FILLS.length]);
    const p1 = at(b, R);
    const p2 = at(b + sweep, R);
    if (sectors.length === 1) els.push(`<circle cx="0" cy="0" r="${R}" fill="${fill}" stroke="${ink}" stroke-width="1.6"/>`);
    else els.push(`<path d="M0,0 L${f2(p1.x)},${f2(p1.y)} A${R},${R} 0 ${sweep > 180 ? 1 : 0} 1 ${f2(p2.x)},${f2(p2.y)} Z" fill="${fill}" stroke="${ink}" stroke-width="1.6"/>`);
    if (s.label && spec.labels !== false) {
      const m = at(b + sweep / 2, R * (sweep < 50 ? 0.7 : 0.6));
      const size = sweep < 40 ? 9 : (String(s.label).length > 6 ? 10.5 : 12.5);
      els.push(text(m.x, m.y, s.label, { size, bold: true, fill: ink }));
    }
    b += sweep;
  });
  if (spec.pointer !== false) {
    // By default the arrow points just inside the first sector, clear of its label.
    const firstSweep = (360 * (Number(sectors[0] && sectors[0].weight) || 1)) / total;
    const tip = at(spec.pointer != null ? Number(spec.pointer) : firstSweep * 0.12, R * 0.8);
    const ang = Math.atan2(tip.y, tip.x);
    const l = { x: tip.x - 9 * Math.cos(ang) + 4.5 * Math.sin(ang), y: tip.y - 9 * Math.sin(ang) - 4.5 * Math.cos(ang) };
    const r = { x: tip.x - 9 * Math.cos(ang) - 4.5 * Math.sin(ang), y: tip.y - 9 * Math.sin(ang) + 4.5 * Math.cos(ang) };
    els.push(`<line x1="0" y1="0" x2="${f2(tip.x)}" y2="${f2(tip.y)}" stroke="${ink}" stroke-width="3.2" stroke-linecap="round"/>`);
    els.push(`<path d="M${f2(tip.x + 3 * Math.cos(ang))},${f2(tip.y + 3 * Math.sin(ang))} L${f2(l.x)},${f2(l.y)} L${f2(r.x)},${f2(r.y)} Z" fill="${ink}"/>`);
  }
  els.push(`<circle cx="0" cy="0" r="4.5" fill="${ink}"/>`);
  return wrapSvg(els, -R, -R, R, R, 4);
}

/* ── Shapes for area and perimeter ──────────────────────────────────────
 * { type: "shape", points: [[0,0],[8,0],[8,5],[0,5]], labels: ["8 m", "5 m", null, null] }
 * Points are in units (x right, y up). labels[i] sits outside edge i (point i to i+1).
 * grid: true fills the shape with unit squares; splits draw dashed cut lines;
 * shapes: [{ points, labels, centerLabel }] draws several in one coordinate system.
 */
function insidePolygon(p, poly) {
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i, i += 1) {
    const a = poly[i];
    const b = poly[j];
    if ((a.y > p.y) !== (b.y > p.y) && p.x < ((b.x - a.x) * (p.y - a.y)) / (b.y - a.y) + a.x) inside = !inside;
  }
  return inside;
}

function shapeSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const shapes = Array.isArray(spec.shapes) && spec.shapes.length ? spec.shapes : [spec];
  const allPts = shapes.reduce((a, s) => a.concat(s.points || []), []);
  const xs = allPts.map((p) => Number(p[0]));
  const ys = allPts.map((p) => Number(p[1]));
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  const span = Math.max(maxX - minX, (maxY - minY) * 1.2, 1);
  const k = Number(spec.scale) || 220 / span;
  const P2 = (p) => ({ x: (Number(p[0]) - minX) * k, y: (maxY - Number(p[1])) * k });
  const els = [];
  const bounds = [];
  shapes.forEach((s, si) => {
    const poly = (s.points || []).map(P2);
    if (poly.length < 3) return;
    const d = poly.map((p, i) => `${i ? "L" : "M"}${f2(p.x)},${f2(p.y)}`).join(" ") + " Z";
    const shade = s.shade != null ? s.shade : spec.shade;
    els.push(`<path d="${d}" fill="${accent}" fill-opacity="${shade === false ? 0 : 0.1}" stroke="none"/>`);
    if (s.grid || spec.grid) {
      const id = `clip${si}`;
      els.push(`<clipPath id="${id}"><path d="${d}"/></clipPath>`);
      const lines = [];
      for (let gx = Math.ceil(minX); gx <= maxX; gx += 1) lines.push(`<line x1="${f2((gx - minX) * k)}" y1="0" x2="${f2((gx - minX) * k)}" y2="${f2((maxY - minY) * k)}" stroke="#8FA3BA" stroke-width="0.8"/>`);
      for (let gy = Math.ceil(minY); gy <= maxY; gy += 1) lines.push(`<line x1="0" y1="${f2((maxY - gy) * k)}" x2="${f2((maxX - minX) * k)}" y2="${f2((maxY - gy) * k)}" stroke="#8FA3BA" stroke-width="0.8"/>`);
      els.push(`<g clip-path="url(#${id})">${lines.join("")}</g>`);
    }
    (s.splits || []).forEach((seg) => {
      const a = P2(seg[0]);
      const b = P2(seg[1]);
      els.push(`<line x1="${f2(a.x)}" y1="${f2(a.y)}" x2="${f2(b.x)}" y2="${f2(b.y)}" stroke="${accent}" stroke-width="1.8" stroke-dasharray="6 4"/>`);
    });
    els.push(`<path d="${d}" fill="none" stroke="${ink}" stroke-width="2.4" stroke-linejoin="round"/>`);
    // Square corners: a small marker inside every right angle.
    if ((s.rightAngles != null ? s.rightAngles : spec.rightAngles) !== false) {
      poly.forEach((v, i) => {
        const a = poly[(i + poly.length - 1) % poly.length];
        const c = poly[(i + 1) % poly.length];
        const u1 = { x: a.x - v.x, y: a.y - v.y };
        const u2 = { x: c.x - v.x, y: c.y - v.y };
        const l1 = Math.hypot(u1.x, u1.y);
        const l2 = Math.hypot(u2.x, u2.y);
        if (!l1 || !l2 || Math.abs((u1.x * u2.x + u1.y * u2.y) / (l1 * l2)) > 0.02) return;
        const m = Math.min(9, l1 / 4, l2 / 4);
        const e1 = { x: (u1.x / l1) * m, y: (u1.y / l1) * m };
        const e2 = { x: (u2.x / l2) * m, y: (u2.y / l2) * m };
        if (!insidePolygon({ x: v.x + e1.x + e2.x, y: v.y + e1.y + e2.y }, poly)) return;
        els.push(`<path d="M${f2(v.x + e1.x)},${f2(v.y + e1.y)} L${f2(v.x + e1.x + e2.x)},${f2(v.y + e1.y + e2.y)} L${f2(v.x + e2.x)},${f2(v.y + e2.y)}" fill="none" stroke="${ink}" stroke-width="1"/>`);
      });
    }
    // Edge labels sit outside the shape, beside the middle of the edge.
    const labels = Array.isArray(s.labels) ? s.labels : [];
    poly.forEach((p, i) => {
      const lab = labels[i];
      if (lab == null || lab === "") return;
      const q = poly[(i + 1) % poly.length];
      const mid = { x: (p.x + q.x) / 2, y: (p.y + q.y) / 2 };
      const len = Math.hypot(q.x - p.x, q.y - p.y) || 1;
      let nrm = { x: (q.y - p.y) / len, y: -(q.x - p.x) / len };
      if (insidePolygon({ x: mid.x + nrm.x * 4, y: mid.y + nrm.y * 4 }, poly)) nrm = { x: -nrm.x, y: -nrm.y };
      const off = Math.abs(nrm.x) > 0.7 ? 8 : 13;
      const tp = { x: mid.x + nrm.x * off, y: mid.y + nrm.y * off };
      const anchor = Math.abs(nrm.x) > 0.7 ? (nrm.x > 0 ? "start" : "end") : "middle";
      els.push(text(tp.x, tp.y, lab, { size: 17, bold: true, fill: ink, anchor }));
      const tw = String(lab).length * 9.8;
      bounds.push({ x: anchor === "end" ? tp.x - tw : (anchor === "start" ? tp.x : tp.x - tw / 2), y: tp.y - 11 });
      bounds.push({ x: anchor === "end" ? tp.x : (anchor === "start" ? tp.x + tw : tp.x + tw / 2), y: tp.y + 11 });
    });
    if (s.centerLabel) {
      const cx = poly.reduce((t, p) => t + p.x, 0) / poly.length;
      const cy = poly.reduce((t, p) => t + p.y, 0) / poly.length;
      const c = s.labelAt ? P2(s.labelAt) : { x: cx, y: cy };
      els.push(text(c.x, c.y, s.centerLabel, { size: 18, bold: true, fill: accent }));
    }
  });
  const bx = bounds.map((b) => b.x).concat([0, (maxX - minX) * k]);
  const by = bounds.map((b) => b.y).concat([0, (maxY - minY) * k]);
  return wrapSvg(els, Math.min(...bx), Math.min(...by), Math.max(...bx), Math.max(...by), 6);
}

/* ── Short division with the bracket ────────────────────────────────────
 * { type: "shortDivision", dividend: 4728, divisor: 6 }        quotient: "788" shows the answer
 */
function shortDivisionSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const digits = String(spec.dividend).replace(/\s+/g, "").split("");
  const cell = 24;
  const divisor = String(spec.divisor);
  const x0 = divisor.length * 16 + 14;   // where the bracket stands
  const top = 40;
  const base = top + 34;
  const els = [];
  els.push(text(x0 - 12, base - 10, divisor, { size: 28, fill: ink, anchor: "end", baseline: "auto" }));
  els.push(`<path d="M${x0 - 4},${top} Q${x0 + 8},${(top + base) / 2} ${x0 - 4},${base + 4}" fill="none" stroke="${ink}" stroke-width="2.4"/>`);
  const end = x0 + 10 + digits.length * cell + 6;
  els.push(`<line x1="${x0 - 4}" y1="${top}" x2="${end}" y2="${top}" stroke="${ink}" stroke-width="2.4"/>`);
  digits.forEach((d, i) => els.push(text(x0 + 10 + i * cell + cell / 2, base - 10, d, { size: 28, fill: ink, baseline: "auto" })));
  if (spec.quotient != null) {
    const q = String(spec.quotient).replace(/\s+/g, "").split("");
    q.forEach((d, i) => {
      const col = digits.length - q.length + i;
      els.push(text(x0 + 10 + col * cell + cell / 2, top - 8, d, { size: 26, bold: true, fill: accent, baseline: "auto" }));
    });
    if (spec.remainder) els.push(text(end + 4, top - 8, `r ${spec.remainder}`, { size: 20, bold: true, fill: accent, anchor: "start", baseline: "auto" }));
  }
  return wrapSvg(els, 0, 4, end + (spec.remainder ? 40 : 4), base + 8, 6);
}

/* ── Bar model ──────────────────────────────────────────────────────────
 * { type: "barModel", parts: 4, shaded: 1, total: "$80", labels: ["?", "", "", ""], below: ["25%", "25%", "25%", "25%"] }
 */
function barModelSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const n = Math.max(1, Number(spec.parts) || 1);
  const W = 320;
  const H = 52;
  const pw = W / n;
  const els = [];
  const top = spec.total != null ? 30 : 0;
  for (let i = 0; i < n; i += 1) {
    const shaded = i < (Number(spec.shaded) || 0);
    els.push(`<rect x="${f2(i * pw)}" y="${top}" width="${f2(pw)}" height="${H}" fill="${shaded ? accent : "#FFFFFF"}" fill-opacity="${shaded ? 0.3 : 1}" stroke="${ink}" stroke-width="1.8"/>`);
    const lab = (spec.labels || [])[i];
    if (lab) els.push(text(i * pw + pw / 2, top + H / 2, lab, { size: n > 12 ? 11 : 16, bold: true, fill: ink }));
    const below = (spec.below || [])[i];
    if (below) els.push(text(i * pw + pw / 2, top + H + 13, below, { size: n > 12 ? 9.5 : 12.5, fill: ink }));
  }
  if (spec.total != null) {
    els.push(`<path d="M0,${top - 6} L0,${top - 14} L${W},${top - 14} L${W},${top - 6}" fill="none" stroke="${ink}" stroke-width="1.4"/>`);
    els.push(`<rect x="${W / 2 - String(spec.total).length * 4.6 - 6}" y="${top - 24}" width="${String(spec.total).length * 9.2 + 12}" height="18" fill="#FFFFFF"/>`);
    els.push(text(W / 2, top - 15, spec.total, { size: 16, bold: true, fill: ink }));
  }
  return wrapSvg(els, -2, top ? 0 : -2, W + 2, top + H + ((spec.below || []).length ? 22 : 2), 4);
}

/* ── Hundred grid ───────────────────────────────────────────────────────
 * { type: "hundredGrid", shaded: 35 }    fills row by row; fill: "columns" fills down the columns.
 */
function hundredGridSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const cell = 16;
  const n = Math.max(0, Math.min(100, Math.round(Number(spec.shaded) || 0)));
  const els = [];
  for (let i = 0; i < 100; i += 1) {
    const r = spec.fill === "columns" ? i % 10 : Math.floor(i / 10);
    const c = spec.fill === "columns" ? Math.floor(i / 10) : i % 10;
    els.push(`<rect x="${c * cell}" y="${r * cell}" width="${cell}" height="${cell}" fill="${i < n ? accent : "#FFFFFF"}" fill-opacity="${i < n ? 0.55 : 1}" stroke="#7D8A99" stroke-width="0.7"/>`);
  }
  els.push(`<rect x="0" y="0" width="${10 * cell}" height="${10 * cell}" fill="none" stroke="${ink}" stroke-width="1.8"/>`);
  for (let k = 1; k < 10; k += 1) {
    if (k === 5) {
      els.push(`<line x1="${5 * cell}" y1="0" x2="${5 * cell}" y2="${10 * cell}" stroke="${ink}" stroke-width="1.1"/>`);
      els.push(`<line x1="0" y1="${5 * cell}" x2="${10 * cell}" y2="${5 * cell}" stroke="${ink}" stroke-width="1.1"/>`);
    }
  }
  let maxY = 10 * cell;
  if (spec.label) { els.push(text(5 * cell, 10 * cell + 12, spec.label, { size: 12, bold: true, fill: ink })); maxY += 20; }
  return wrapSvg(els, 0, 0, 10 * cell, maxY, 4);
}

/* ── Fraction wall ──────────────────────────────────────────────────────
 * { type: "fractionWall" }                         halves to twelfths, every piece labelled
 * { type: "fractionWall", denoms: [1, 2, 4, 8], shaded: { "4": 3 } }
 * One whole per row, all rows the same width, so equal fractions line up.
 * `shaded` shades the first n pieces of a row; `labels: false` leaves the
 * pieces blank for students to label.
 */
const WALL_FILLS = ["#FDE2E4", "#E2ECE9", "#FFF1C1", "#DCEBFA", "#EADCF8", "#FFE5CC", "#D8F3DC", "#F9DCEB", "#E4E9F2", "#FFF6D6", "#DDF4F7", "#F2E2D2"];
function fractionWallSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const denoms = (Array.isArray(spec.denoms) && spec.denoms.length ? spec.denoms : [1, 2, 3, 4, 5, 6, 8, 10, 12])
    .map((d) => Math.max(1, Math.round(Number(d) || 1)));
  const shaded = spec.shaded || {};
  const W = 360;
  const H = 24;
  const els = [];
  denoms.forEach((d, r) => {
    const y = r * H;
    const pw = W / d;
    const fill = WALL_FILLS[r % WALL_FILLS.length];
    const n = Number(shaded[String(d)]) || 0;
    for (let i = 0; i < d; i += 1) {
      const on = i < n;
      els.push(`<rect x="${f2(i * pw)}" y="${y}" width="${f2(pw)}" height="${H}" fill="${on ? accent : fill}" fill-opacity="${on ? 0.55 : 1}" stroke="${ink}" stroke-width="1.2"/>`);
      if (spec.labels !== false) els.push(text(i * pw + pw / 2, y + H / 2, d === 1 ? "1 whole" : `1/${d}`, { size: d >= 10 ? 9.5 : 11, bold: true, fill: ink }));
    }
  });
  els.push(`<rect x="0" y="0" width="${W}" height="${denoms.length * H}" fill="none" stroke="${ink}" stroke-width="2"/>`);
  return wrapSvg(els, 0, 0, W, denoms.length * H, 4);
}

/* ── Clock face ─────────────────────────────────────────────────────────
 * { type: "clock" }                    a blank face for students to draw the hands
 * { type: "clock", time: "3:45" }      hands drawn at that time
 * { type: "clock", minutes: true }     5, 10, 15 ... labelled round the outside
 * `digital: true` adds an empty box underneath for writing the digital time.
 */
function clockSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const R = 80;
  const els = [];
  els.push(`<circle cx="0" cy="0" r="${R}" fill="#FFFFFF" stroke="${ink}" stroke-width="3"/>`);
  for (let m = 0; m < 60; m += 1) {
    const big = m % 5 === 0;
    const a = pt(90 - m * 6, R);
    const b = pt(90 - m * 6, R - (big ? 10 : 5));
    els.push(`<line x1="${f2(a.x)}" y1="${f2(a.y)}" x2="${f2(b.x)}" y2="${f2(b.y)}" stroke="${ink}" stroke-width="${big ? 2.4 : 1}"/>`);
  }
  for (let h = 1; h <= 12; h += 1) {
    const p = pt(90 - h * 30, R - 22);
    els.push(text(p.x, p.y, String(h), { size: 15, bold: true, fill: ink }));
  }
  let reach = R;
  if (spec.minutes) {
    for (let k = 0; k < 12; k += 1) {
      const p = pt(90 - k * 30, R + 13);
      els.push(text(p.x, p.y, String(k * 5), { size: 9.5, fill: accent, bold: true }));
    }
    reach = R + 22;
  }
  const t = /^(\d{1,2}):(\d{2})$/.exec(String(spec.time || ""));
  if (t) {
    const hh = Number(t[1]) % 12;
    const mm = Number(t[2]);
    const hour = pt(90 - (hh + mm / 60) * 30, R * 0.4);
    const minute = pt(90 - mm * 6, R * 0.62);
    els.push(`<line x1="0" y1="0" x2="${f2(hour.x)}" y2="${f2(hour.y)}" stroke="${ink}" stroke-width="6" stroke-linecap="round"/>`);
    els.push(`<line x1="0" y1="0" x2="${f2(minute.x)}" y2="${f2(minute.y)}" stroke="${accent}" stroke-width="3.5" stroke-linecap="round"/>`);
  }
  els.push(`<circle cx="0" cy="0" r="4.5" fill="${ink}"/>`);
  let maxY = reach;
  if (spec.digital) {
    els.push(`<rect x="-42" y="${reach + 10}" width="84" height="30" rx="6" fill="#FFFFFF" stroke="${ink}" stroke-width="1.6"/>`);
    els.push(text(0, reach + 25, ":", { size: 18, bold: true, fill: ink }));
    maxY = reach + 40;
  }
  if (spec.label) { els.push(text(0, maxY + 12, spec.label, { size: 12, bold: true, fill: ink })); maxY += 22; }
  return wrapSvg(els, -reach, -reach, reach, maxY, 4);
}

/* ── Unit conversion chart ──────────────────────────────────────────────
 * { type: "conversionChart", measure: "length" }   km, m, cm, mm with x and divide arrows
 * measure: length | mass | capacity | time, or your own:
 * { type: "conversionChart", units: ["m", "cm"], factors: [100] }
 * Bigger units sit on the left. The arrows above go right (to the smaller
 * unit: multiply), the arrows below go left (to the bigger unit: divide).
 */
const CONVERSIONS = {
  length: { units: ["km", "m", "cm", "mm"], factors: [1000, 100, 10] },
  mass: { units: ["t", "kg", "g", "mg"], factors: [1000, 1000, 1000] },
  capacity: { units: ["kL", "L", "mL"], factors: [1000, 1000] },
  time: { units: ["days", "hours", "minutes", "seconds"], factors: [24, 60, 60] },
};
function conversionChartSvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const preset = CONVERSIONS[spec.measure] || CONVERSIONS.length;
  const units = Array.isArray(spec.units) && spec.units.length >= 2 ? spec.units.map(String) : preset.units;
  const factors = Array.isArray(spec.factors) && spec.factors.length === units.length - 1 ? spec.factors : preset.factors;
  const BW = units.some((u) => u.length > 4) ? 78 : 58;
  const BH = 40;
  const GAP = 52;
  const els = [];
  const arrowHead = (x, y, dx, dy, fill) => {
    const len = Math.hypot(dx, dy) || 1;
    const ux = dx / len;
    const uy = dy / len;
    const s = 8;
    return `<polygon points="${f2(x)},${f2(y)} ${f2(x - ux * s - uy * s * 0.6)},${f2(y - uy * s + ux * s * 0.6)} ${f2(x - ux * s + uy * s * 0.6)},${f2(y - uy * s - ux * s * 0.6)}" fill="${fill}"/>`;
  };
  units.forEach((u, i) => {
    const x = i * (BW + GAP);
    els.push(`<rect x="${x}" y="0" width="${BW}" height="${BH}" rx="8" fill="${accent}"/>`);
    els.push(text(x + BW / 2, BH / 2, u, { size: 16, bold: true, fill: "#FFFFFF" }));
    if (i === units.length - 1) return;
    const x1 = x + BW * 0.75;
    const x2 = x + BW + GAP + BW * 0.25;
    const mid = (x1 + x2) / 2;
    const f = num(Number(factors[i]));
    // Above: to the smaller unit, multiply.
    els.push(`<path d="M${f2(x1)},-2 Q${f2(mid)},-34 ${f2(x2)},-2" fill="none" stroke="${ink}" stroke-width="1.8"/>`);
    els.push(arrowHead(x2, -2, x2 - mid, 32, ink));
    els.push(text(mid, -27, `× ${f}`, { size: 12, bold: true, fill: ink }));
    // Below: to the bigger unit, divide.
    els.push(`<path d="M${f2(x2)},${BH + 2} Q${f2(mid)},${BH + 34} ${f2(x1)},${BH + 2}" fill="none" stroke="${ink}" stroke-width="1.8"/>`);
    els.push(arrowHead(x1, BH + 2, x1 - mid, -32, ink));
    els.push(text(mid, BH + 27, `÷ ${f}`, { size: 12, bold: true, fill: ink }));
  });
  const W = units.length * BW + (units.length - 1) * GAP;
  let maxY = BH + 36;
  if (spec.label !== false) {
    els.push(text(W / 2, maxY + 10, spec.label || "Bigger unit to smaller: multiply. Smaller to bigger: divide.", { size: 11, fill: ink }));
    maxY += 20;
  }
  return wrapSvg(els, 0, -36, W, maxY, 4);
}

/* ── Tally chart ────────────────────────────────────────────────────────
 * { type: "tally", headers: ["Sport", "Tally", "Frequency"], rows: [["Footy", 8], ["Netball", 5]], counts: true }
 * Tally marks are drawn in fives (four strokes and a gate). counts: false
 * leaves the frequency column blank for students to fill in.
 */
function tallySvg(spec, colours) {
  const ink = hexColour(colours && colours.ink, "2B2B2B");
  const accent = hexColour(colours && colours.accent, "C0392B");
  const rows = (spec.rows || []).map((r) => [String(r[0]), Math.max(0, Number(r[1]) || 0)]);
  const headers = spec.headers || ["Category", "Tally", "Frequency"];
  const rowH = 28;
  const c1 = Math.max(70, ...rows.map((r) => r[0].length * 7.2 + 16), headers[0].length * 7.4 + 16);
  const maxMarks = Math.max(5, ...rows.map((r) => r[1]));
  const c2 = Math.max(90, Math.ceil(maxMarks / 5) * 36 + 16);
  const c3 = spec.counts === "none" ? 0 : 80;
  const W = c1 + c2 + c3;
  const els = [];
  const H = rowH * (rows.length + 1);
  els.push(`<rect x="0" y="0" width="${W}" height="${rowH}" fill="${accent}" fill-opacity="0.85"/>`);
  [[headers[0], c1 / 2], [headers[1], c1 + c2 / 2], [headers[2], c1 + c2 + c3 / 2]].forEach(([h, x], i) => {
    if (i < 2 || c3) els.push(text(x, rowH / 2, h, { size: 12, bold: true, fill: "#FFFFFF" }));
  });
  rows.forEach((r, i) => {
    const y = rowH * (i + 1);
    if (i % 2 === 1) els.push(`<rect x="0" y="${y}" width="${W}" height="${rowH}" fill="${accent}" fill-opacity="0.07"/>`);
    els.push(text(10, y + rowH / 2, r[0], { size: 12, fill: ink, anchor: "start" }));
    // Tally marks: groups of four strokes crossed by a fifth.
    let x = c1 + 10;
    for (let g = 0; g < Math.ceil(r[1] / 5); g += 1) {
      const n = Math.min(5, r[1] - g * 5);
      for (let k = 0; k < Math.min(n, 4); k += 1) {
        els.push(`<line x1="${x + k * 6}" y1="${y + 6}" x2="${x + k * 6}" y2="${y + rowH - 6}" stroke="${ink}" stroke-width="1.8" stroke-linecap="round"/>`);
      }
      if (n === 5) els.push(`<line x1="${x - 3}" y1="${y + rowH - 8}" x2="${x + 3 * 6 + 3}" y2="${y + 8}" stroke="${ink}" stroke-width="1.8" stroke-linecap="round"/>`);
      x += 36;
    }
    if (c3 && spec.counts !== false) els.push(text(c1 + c2 + c3 / 2, y + rowH / 2, String(r[1]), { size: 13, bold: true, fill: ink }));
  });
  for (let i = 0; i <= rows.length + 1; i += 1) els.push(`<line x1="0" y1="${i * rowH}" x2="${W}" y2="${i * rowH}" stroke="#9AA6B5" stroke-width="0.8"/>`);
  [0, c1, c1 + c2, W].forEach((x) => { if (x <= W) els.push(`<line x1="${x}" y1="0" x2="${x}" y2="${H}" stroke="#9AA6B5" stroke-width="0.8"/>`); });
  return wrapSvg(els, 0, 0, W, H, 3);
}

const BUILDERS = {
  tally: tallySvg,
  angle: angleSvg,
  columnSum: columnSumSvg,
  grid: gridSvg,
  chart: chartSvg,
  spinner: spinnerSvg,
  shape: shapeSvg,
  shortDivision: shortDivisionSvg,
  barModel: barModelSvg,
  hundredGrid: hundredGridSvg,
  fractionWall: fractionWallSvg,
  clock: clockSvg,
  conversionChart: conversionChartSvg,
};
const DIAGRAM_TYPES = Object.keys(BUILDERS);

/**
 * Preferred size of each diagram on paper, in points: `w` the width a full
 * question card gives it, `h` the height the worksheet layout plans for.
 * Angles stay big enough to measure with a real protractor.
 */
const PAPER = {
  tally: { w: 300, h: 150 },
  angle: { w: 190, h: 118 },
  columnSum: { w: 150, h: 90 },
  grid: { w: 230, h: 175 },
  chart: { w: 260, h: 150 },
  spinner: { w: 130, h: 110 },
  shape: { w: 220, h: 120 },
  shortDivision: { w: 170, h: 62 },
  barModel: { w: 300, h: 62 },
  hundredGrid: { w: 150, h: 130 },
  fractionWall: { w: 340, h: 215 },
  clock: { w: 140, h: 140 },
  conversionChart: { w: 330, h: 110 },
};

function diagramSvg(spec, colours) {
  const build = BUILDERS[spec.type];
  if (!build) throw new Error(`[diagrams] unknown diagram type ${spec.type}`);
  return build(spec, colours);
}

/** Render a diagram spec to PNG. Returns { buffer, dataUri, aspect }. */
function diagramPng(spec, colours, widthPx) {
  const { svg, aspect } = diagramSvg(spec, colours);
  const family = resolveSansFamily();
  const font = family
    ? { loadSystemFonts: false, fontFiles: [family.regular, family.bold], defaultFontFamily: family.name, sansSerifFamily: family.name }
    : { loadSystemFonts: true };
  const R = getResvg();
  const buffer = new R(svg, { fitTo: { mode: "width", value: widthPx || 1400 }, font, background: "rgba(255,255,255,0)" }).render().asPng();
  return { buffer, dataUri: `data:image/png;base64,${buffer.toString("base64")}`, aspect };
}

module.exports = { DIAGRAM_TYPES, PAPER, diagramSvg, diagramPng };
