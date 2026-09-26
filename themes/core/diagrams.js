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
 * overlays a protractor on the first ray, with an outer and an inner scale.
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
    const base = rays[0];
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
    const r = a.r || (spec.protractor ? 16 : 20 + (i % 2) * 9);
    const isRight = a.right === true;
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
      const mid = a.from + sweep / 2;
      const lr = (isRight ? 16 : r) + (sweep < 35 ? 17 : 12);
      const p = pt(mid, lr);
      els.push(`<text x="${f2(p.x)}" y="${f2(p.y)}" font-size="11" font-weight="bold" fill="${ink}" text-anchor="middle" dominant-baseline="central">${esc(a.label)}</text>`);
      grow({ x: p.x - 16, y: p.y - 8 }); grow({ x: p.x + 16, y: p.y + 8 });
    }
  });

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

const BUILDERS = { angle: angleSvg, columnSum: columnSumSvg };
const DIAGRAM_TYPES = Object.keys(BUILDERS);

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

module.exports = { DIAGRAM_TYPES, diagramSvg, diagramPng };
