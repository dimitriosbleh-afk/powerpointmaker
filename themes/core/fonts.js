"use strict";

/**
 * The sans-serif family used on paper (worksheets, scaffolds, keys) and for
 * text inside rendered diagrams. Trebuchet MS first: friendly, very readable
 * at primary sizes, and installed on both macOS and Windows, so every
 * teacher's machine renders the same sheet. Arial, then Calibri or the Linux
 * families, are fallbacks.
 */

const fs = require("fs");

function candidates() {
  if (process.platform === "win32") {
    return [
      { name: "Trebuchet MS", regular: "C:/Windows/Fonts/trebuc.ttf", bold: "C:/Windows/Fonts/trebucbd.ttf", italic: "C:/Windows/Fonts/trebucit.ttf" },
      { name: "Arial", regular: "C:/Windows/Fonts/arial.ttf", bold: "C:/Windows/Fonts/arialbd.ttf", italic: "C:/Windows/Fonts/ariali.ttf" },
      { name: "Calibri", regular: "C:/Windows/Fonts/calibri.ttf", bold: "C:/Windows/Fonts/calibrib.ttf", italic: "C:/Windows/Fonts/calibrii.ttf" },
    ];
  }
  if (process.platform === "darwin") {
    const dir = "/System/Library/Fonts/Supplemental";
    return [
      { name: "Trebuchet MS", regular: `${dir}/Trebuchet MS.ttf`, bold: `${dir}/Trebuchet MS Bold.ttf`, italic: `${dir}/Trebuchet MS Italic.ttf` },
      { name: "Arial", regular: `${dir}/Arial.ttf`, bold: `${dir}/Arial Bold.ttf`, italic: `${dir}/Arial Italic.ttf` },
    ];
  }
  return [
    { name: "Liberation Sans", regular: "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf", bold: "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf", italic: "/usr/share/fonts/truetype/liberation2/LiberationSans-Italic.ttf" },
    { name: "Liberation Sans", regular: "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", bold: "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", italic: "/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf" },
    { name: "DejaVu Sans", regular: "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", bold: "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", italic: "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf" },
  ];
}

let cached;
/** { name, regular, bold, italic } for the first family fully installed, or null. */
function resolveSansFamily() {
  if (cached !== undefined) return cached;
  cached = candidates().find((f) => [f.regular, f.bold, f.italic].every((p) => fs.existsSync(p))) || null;
  return cached;
}

module.exports = { resolveSansFamily, sansFamilyCandidates: candidates };
