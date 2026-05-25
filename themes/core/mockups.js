"use strict";

function normaliseHex(color, fallback) {
  const raw = String(color || fallback || "").replace("#", "").trim();
  return /^[0-9A-Fa-f]{6}$/.test(raw) ? raw.toUpperCase() : String(fallback || "FFFFFF").replace("#", "").toUpperCase();
}

function hexToRgb(color) {
  const value = normaliseHex(color, "FFFFFF");
  return {
    r: parseInt(value.slice(0, 2), 16),
    g: parseInt(value.slice(2, 4), 16),
    b: parseInt(value.slice(4, 6), 16),
  };
}

function mixHex(colorA, colorB, amount) {
  const ratio = Math.max(0, Math.min(1, Number(amount) || 0));
  const a = hexToRgb(colorA);
  const b = hexToRgb(colorB);
  const r = Math.round(a.r + (b.r - a.r) * ratio);
  const g = Math.round(a.g + (b.g - a.g) * ratio);
  const b2 = Math.round(a.b + (b.b - a.b) * ratio);
  return [r, g, b2].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}

function lightenHex(color, amount) {
  return mixHex(color, "FFFFFF", amount);
}

function darkenHex(color, amount) {
  return mixHex(color, "000000", amount);
}

function isStructuredMockupSpec(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value) && Array.isArray(value.components);
}

function looksLikeNav(text) {
  return /\b(home|about|menu|contact|videos|animals|plants|news|sports)\b/i.test(text) ||
    /[|/]/.test(text) ||
    (text.match(/\b[A-Z][a-z]+\b/g) || []).length >= 4;
}

function looksLikeCta(text) {
  return /(click|learn|watch|read|discover|start|join|subscribe|share|talk|apply|download|more|sign up|find out)/i.test(text) ||
    /->/.test(text);
}

function looksLikeQuote(text) {
  return /["“”]| - /.test(text);
}

function resolveHeroMode(component) {
  const raw = String(component.mode || component.scene || "diagram").toLowerCase();
  if (raw === "chart") return "chart";
  if (raw === "diagram") return "diagram";
  if (raw === "browserframe") return "browserFrame";
  if (raw === "cardgrid") return "cardGrid";
  if (["landscape", "wildlife", "photo", "image"].includes(raw)) return "photo";
  if (["person", "peoplesupport", "reach", "reachinghand", "support"].includes(raw)) return "diagram";
  return "diagram";
}

function inferBandKind(component, index, total) {
  const text = String(component.text || "");
  if (index === 0) return "masthead";
  if (looksLikeNav(text)) return "nav";
  if (looksLikeCta(text)) return index === total - 1 ? "footerBand" : "cta";
  if (looksLikeQuote(text)) return "quote";
  if (index === total - 1) return "caption";
  return index <= 2 ? "heading" : "subheading";
}

function normalizeScale(value, fallback) {
  const scale = Number(value);
  return Number.isFinite(scale) && scale > 0 ? scale : fallback;
}

function normalizeMockupComponent(component, index, total) {
  const source = component || {};
  let kind = String(source.kind || "band");
  if (kind === "band") kind = inferBandKind(source, index, total);
  if (kind === "textLines") kind = "textBlock";

  return {
    ...source,
    kind,
    text: source.text != null ? String(source.text) : "",
    scale: normalizeScale(source.scale, 1),
    fill: source.fill ? normaliseHex(source.fill, source.fill) : undefined,
    border: source.border ? normaliseHex(source.border, source.border) : undefined,
    textColor: source.textColor ? normaliseHex(source.textColor, source.textColor) : undefined,
    accent: source.accent ? normaliseHex(source.accent, source.accent) : undefined,
    align: source.align || ((kind === "masthead" || kind === "cta" || kind === "stat" || kind === "footerBand") ? "center" : "left"),
    mode: kind === "hero" ? resolveHeroMode(source) : source.mode,
    count: Math.max(1, Number(source.count) || 3),
    rows: Math.max(1, Number(source.rows) || 2),
    cols: Math.max(1, Number(source.cols) || 2),
    widths: Array.isArray(source.widths) ? source.widths : undefined,
    lineGap: source.lineGap,
    linePad: source.linePad,
    overlayText: source.overlayText != null ? String(source.overlayText) : undefined,
    overlayFill: source.overlayFill ? normaliseHex(source.overlayFill, source.overlayFill) : undefined,
    overlayFontSize: source.overlayFontSize,
    legacyScene: source.scene,
  };
}

function normalizeStructuredMockup(spec) {
  const source = spec || {};
  const pageFill = normaliseHex(source.pageFill, "FFFFFF");
  const accent = normaliseHex(source.accent, "1B3A6B");
  const textColor = normaliseHex(source.textColor, "243142");
  const pageBorder = normaliseHex(source.pageBorder, mixHex(accent, pageFill, 0.72));
  const neutral = normaliseHex(source.neutral, mixHex(pageBorder, pageFill, 0.3));
  const softFill = normaliseHex(source.softFill, mixHex(accent, pageFill, 0.9));
  const softBorder = normaliseHex(source.softBorder, mixHex(accent, pageFill, 0.68));
  const mutedLine = normaliseHex(source.mutedLine, mixHex(textColor, pageFill, 0.72));
  const components = Array.isArray(source.components) ? source.components : [];

  return {
    ...source,
    pageFill,
    pageBorder,
    accent,
    textColor,
    neutral,
    softFill,
    softBorder,
    mutedLine,
    innerPad: source.innerPad,
    gap: source.gap,
    components: components.map((component, index) => normalizeMockupComponent(component, index, components.length)),
  };
}

function createStructuredMockupPlan(spec, frame, opts) {
  const normalized = normalizeStructuredMockup(spec);
  const o = opts || {};
  const x = Number(frame && frame.x) || 0;
  const y = Number(frame && frame.y) || 0;
  const w = Number(frame && frame.w) || 0;
  const h = Number(frame && frame.h) || 0;
  const innerPad = normalized.innerPad != null
    ? normalized.innerPad
    : (o.innerPad != null ? o.innerPad : 0);
  const gap = normalized.gap != null
    ? normalized.gap
    : (o.gap != null ? o.gap : 0);
  const components = normalized.components || [];
  const innerX = x + innerPad;
  const innerY = y + innerPad;
  const innerW = w - innerPad * 2;
  const availableH = h - innerPad * 2 - gap * Math.max(components.length - 1, 0);
  const minBlockH = o.minBlockH != null ? o.minBlockH : 0;
  const totalScale = components.reduce((sum, component) => sum + component.scale, 0) || 1;
  let cursorY = innerY;

  const blocks = components.map((component, index) => {
    const blockH = Math.max(minBlockH, availableH * (component.scale / totalScale));
    const block = {
      component,
      index,
      normalized,
      x: innerX,
      y: cursorY,
      w: innerW,
      h: blockH,
    };
    cursorY += blockH + gap;
    return block;
  });

  return {
    normalized,
    blocks,
    innerPad,
    gap,
    innerX,
    innerY,
    innerW,
    endY: cursorY - (blocks.length ? gap : 0),
  };
}

function forEachStructuredMockupBlock(spec, frame, opts, visitor) {
  const plan = createStructuredMockupPlan(spec, frame, opts);
  plan.blocks.forEach((block) => visitor(block, plan));
  return plan;
}

module.exports = {
  createStructuredMockupPlan,
  darkenHex,
  forEachStructuredMockupBlock,
  isStructuredMockupSpec,
  lightenHex,
  mixHex,
  normaliseHex,
  normalizeStructuredMockup,
};
