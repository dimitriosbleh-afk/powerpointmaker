"use strict";

const { validateBounds } = require("./layout");
const { byBand, DEFAULT_SIZES } = require("./gradeBand");

/**
 * Concrete-representation manipulative helpers required by the megaprompt's
 * §15a Visual Anchor Test and §24 conceptual-representation list. Designed
 * for Foundation to Year 2 first; usable in upper primary too.
 *
 * Every helper takes (slide, x, y, ...args, opts) and draws on the slide
 * at the given inches-coordinate. Sizing defaults are deliberately generous
 * so that early-years slides feel large from the back of the room.
 *
 * Returns a plain object of helpers ready to be spread onto a theme.
 *
 * @param {object} C       Palette colours (semantic keys)
 * @param {string} FONT_B  Body font name (used for number-track digits)
 * @param {object} [S]     Grade-band size table (from gradeBand) so labels
 *                         and chips scale with the year level
 * @returns {object} factory bag of manipulative helpers
 */
function createManipulatives(C, FONT_B, S) {
  const sz = S || DEFAULT_SIZES;

  /* ------------------------------------------------------------------ */
  /*  Tens frame (2 rows of 5)                                          */
  /* ------------------------------------------------------------------ */

  /**
   * Two rows of five cells. The first `filled` cells get a coloured counter.
   *
   * @param {object} slide   PptxGenJS slide object
   * @param {number} x       Left edge x (inches)
   * @param {number} y       Top edge y (inches)
   * @param {number} w       Total frame width (inches); cell side = w/5
   * @param {number} filled  How many cells are filled (0-10)
   * @param {object} [opts]  { cellH, fillColor, borderColor, dotRatio, lineWidth, fillFromTop }
   */
  function addTensFrame(slide, x, y, w, filled, opts) {
    const o = opts || {};
    const cellW = w / 5;
    const cellH = o.cellH || cellW;
    const fillColor = o.fillColor || C.PRIMARY;
    const borderColor = o.borderColor || C.CHARCOAL;
    const lineW = o.lineWidth || 1.5;
    const dotRatio = o.dotRatio || 0.62;
    const fillFromTop = o.fillFromTop !== false;
    validateBounds("addTensFrame", x, y, w, cellH * 2);

    const fillCount = Math.max(0, Math.min(10, Math.floor(filled)));

    for (let row = 0; row < 2; row += 1) {
      for (let col = 0; col < 5; col += 1) {
        const idx = fillFromTop ? row * 5 + col : (1 - row) * 5 + col;
        const cx = x + col * cellW;
        const cy = y + row * cellH;
        slide.addShape("rect", {
          x: cx, y: cy, w: cellW, h: cellH,
          fill: { color: C.WHITE },
          line: { color: borderColor, width: lineW },
        });
        if (idx < fillCount) {
          const dotW = cellW * dotRatio;
          const dotH = cellH * dotRatio;
          slide.addShape("roundRect", {
            x: cx + (cellW - dotW) / 2,
            y: cy + (cellH - dotH) / 2,
            w: dotW, h: dotH,
            rectRadius: Math.min(dotW, dotH) / 2,
            fill: { color: fillColor },
            line: { color: fillColor, width: 0.2 },
          });
        }
      }
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Five frame (single row of 5)                                       */
  /* ------------------------------------------------------------------ */

  /**
   * Single row of five cells with optional counter fill. Used heavily in
   * Foundation to teach 0-5 cardinality before moving to a tens frame.
   */
  function addFiveFrame(slide, x, y, w, filled, opts) {
    const o = opts || {};
    const cellW = w / 5;
    const cellH = o.cellH || cellW;
    const fillColor = o.fillColor || C.PRIMARY;
    const borderColor = o.borderColor || C.CHARCOAL;
    const lineW = o.lineWidth || 1.5;
    const dotRatio = o.dotRatio || 0.62;
    validateBounds("addFiveFrame", x, y, w, cellH);

    const fillCount = Math.max(0, Math.min(5, Math.floor(filled)));

    for (let col = 0; col < 5; col += 1) {
      const cx = x + col * cellW;
      slide.addShape("rect", {
        x: cx, y, w: cellW, h: cellH,
        fill: { color: C.WHITE },
        line: { color: borderColor, width: lineW },
      });
      if (col < fillCount) {
        const dotW = cellW * dotRatio;
        const dotH = cellH * dotRatio;
        slide.addShape("roundRect", {
          x: cx + (cellW - dotW) / 2,
          y: y + (cellH - dotH) / 2,
          w: dotW, h: dotH,
          rectRadius: Math.min(dotW, dotH) / 2,
          fill: { color: fillColor },
          line: { color: fillColor, width: 0.2 },
        });
      }
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Dot card (dice-style patterns 1-6, split-dice 7-10)                */
  /* ------------------------------------------------------------------ */

  // Relative dot positions inside a 1x1 unit square. Standard dice for 1-6;
  // 7-10 stack two dice patterns side by side via composition in the helper.
  const DICE_PATTERNS = {
    1: [[0.50, 0.50]],
    2: [[0.28, 0.28], [0.72, 0.72]],
    3: [[0.25, 0.25], [0.50, 0.50], [0.75, 0.75]],
    4: [[0.27, 0.27], [0.73, 0.27], [0.27, 0.73], [0.73, 0.73]],
    5: [[0.27, 0.27], [0.73, 0.27], [0.50, 0.50], [0.27, 0.73], [0.73, 0.73]],
    6: [[0.27, 0.20], [0.73, 0.20], [0.27, 0.50], [0.73, 0.50], [0.27, 0.80], [0.73, 0.80]],
  };

  /**
   * Draw a square dot card showing 1-10 dots in a dice-recognisable pattern.
   * Counts 7-10 are rendered as dice-6 + dice-(n-6) inside the same card with
   * a thin divider so students still see the subitised groups.
   *
   * @param {number} count  Dot count (1-10)
   */
  function addDotCard(slide, x, y, size, count, opts) {
    const o = opts || {};
    const dotColor = o.dotColor || C.PRIMARY;
    const cardFill = o.cardFill || C.WHITE;
    const borderColor = o.borderColor || C.CHARCOAL;
    const lineW = o.lineWidth || 1.4;
    const dotRatio = o.dotRatio || 0.18;
    const safe = Math.max(1, Math.min(10, Math.floor(count)));
    validateBounds("addDotCard", x, y, size, size);

    slide.addShape("rect", {
      x, y, w: size, h: size,
      fill: { color: cardFill },
      line: { color: borderColor, width: lineW },
    });

    function placeDots(panelX, panelW, panelY, panelH, n) {
      const pattern = DICE_PATTERNS[n] || [];
      const dotSize = Math.min(panelW, panelH) * dotRatio;
      pattern.forEach((p) => {
        slide.addShape("roundRect", {
          x: panelX + panelW * p[0] - dotSize / 2,
          y: panelY + panelH * p[1] - dotSize / 2,
          w: dotSize, h: dotSize,
          rectRadius: dotSize / 2,
          fill: { color: dotColor },
          line: { color: dotColor, width: 0.2 },
        });
      });
    }

    if (safe <= 6) {
      placeDots(x, size, y, size, safe);
      return;
    }

    const halfW = size / 2;
    placeDots(x, halfW, y, size, 6);
    placeDots(x + halfW, halfW, y, size, safe - 6);
    slide.addShape("line", {
      x: x + halfW, y: y + size * 0.1,
      w: 0, h: size * 0.8,
      line: { color: borderColor, width: lineW * 0.8 },
    });
  }

  /* ------------------------------------------------------------------ */
  /*  Number track (1..N consecutive cells)                              */
  /* ------------------------------------------------------------------ */

  /**
   * Horizontal track of numbered cells from `start` to `end` inclusive.
   * Cells are equal width. Optional `highlight` array of values to colour-fill.
   *
   * @param {number}     w           Total track width (inches)
   * @param {number}     start       First number on the track (typically 1 or 0)
   * @param {number}     end         Last number on the track
   * @param {number[]}   [highlight] Values to fill with the highlight colour
   */
  function addNumberTrack(slide, x, y, w, start, end, highlight, opts) {
    const o = opts || {};
    const count = Math.max(1, end - start + 1);
    const cellW = w / count;
    const cellH = o.cellH || Math.max(0.5, cellW * 0.95);
    const fillColor = o.fillColor || C.ACCENT;
    const borderColor = o.borderColor || C.CHARCOAL;
    const numberColor = o.numberColor || C.CHARCOAL;
    const highlightTextColor = o.highlightTextColor || C.WHITE;
    const lineW = o.lineWidth || 1.2;
    const fontSize = o.fontSize || Math.max(10, Math.min(28, Math.round(cellW * 28)));
    const highlightSet = new Set(Array.isArray(highlight) ? highlight : []);
    validateBounds("addNumberTrack", x, y, w, cellH);

    for (let i = 0; i < count; i += 1) {
      const value = start + i;
      const cx = x + i * cellW;
      const isHi = highlightSet.has(value);
      slide.addShape("rect", {
        x: cx, y, w: cellW, h: cellH,
        fill: { color: isHi ? fillColor : C.WHITE },
        line: { color: borderColor, width: lineW },
      });
      slide.addText(String(value), {
        x: cx, y, w: cellW, h: cellH,
        fontSize, fontFace: FONT_B,
        color: isHi ? highlightTextColor : numberColor,
        bold: true, align: "center", valign: "middle", margin: 0,
      });
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Fraction strip set (stacked equal-parts bars)                      */
  /* ------------------------------------------------------------------ */

  /**
   * Stack of horizontal fraction bars. Each row shows 1 whole partitioned
   * into n equal parts, with optional shaded segments.
   *
   * @param {number}   w             Total bar width (inches)
   * @param {number}   h             Total stacked height (inches)
   * @param {object[]} strips        [{ denom, shaded, label, color }]
   *   - denom:  number of equal parts in this strip
   *   - shaded: how many parts to fill (0..denom)
   *   - label:  optional label drawn to the right of the strip
   *   - color:  optional override fill colour
   */
  function addFractionStripSet(slide, x, y, w, h, strips, opts) {
    const o = opts || {};
    const items = Array.isArray(strips) ? strips : [];
    if (!items.length) return;
    const labelW = o.labelW != null ? o.labelW : 0.8;
    const showLabels = o.showLabels !== false && labelW > 0;
    const stripW = showLabels ? Math.max(0.5, w - labelW - 0.08) : w;
    // Separate wholes must LOOK separate: without a gap, stacked strips fuse
    // into one grid and "3 wholes in quarters" reads as one rectangle cut
    // into 12. Gap defaults on for multi-strip sets.
    const stripGap = o.gap != null ? o.gap : (items.length > 1 ? 0.12 : 0);
    const stripH = (h - stripGap * (items.length - 1)) / items.length;
    const borderColor = o.borderColor || C.CHARCOAL;
    const lineW = o.lineWidth || 1.0;
    const labelFontSize = o.labelFontSize || Math.max(10, Math.min(16, Math.round(stripH * 24)));
    const palette = [C.PRIMARY, C.SECONDARY, C.ACCENT, C.SUCCESS, C.ALERT];
    validateBounds("addFractionStripSet", x, y, w, h);

    items.forEach((item, idx) => {
      const denom = Math.max(1, Math.floor((item && item.denom) || 1));
      const shaded = Math.max(0, Math.min(denom, Math.floor((item && item.shaded) || 0)));
      const stripFill = (item && item.color) || palette[idx % palette.length];
      const cellW = stripW / denom;
      const stripY = y + idx * (stripH + stripGap);

      for (let i = 0; i < denom; i += 1) {
        const cellX = x + i * cellW;
        slide.addShape("rect", {
          x: cellX, y: stripY, w: cellW, h: stripH,
          fill: { color: i < shaded ? stripFill : C.WHITE },
          line: { color: borderColor, width: lineW },
        });
      }

      if (showLabels) {
        const labelText = (item && item.label) || (
          denom === 1 ? "1 whole" : `${shaded}/${denom}`
        );
        slide.addText(String(labelText), {
          x: x + stripW + 0.06, y: stripY,
          w: labelW, h: stripH,
          fontSize: labelFontSize, fontFace: FONT_B,
          color: C.CHARCOAL, bold: true,
          align: "left", valign: "middle", margin: 0,
        });
      }
    });
  }

  /* ------------------------------------------------------------------ */
  /*  Array (rows x cols of dots)                                        */
  /* ------------------------------------------------------------------ */

  /**
   * Rectangular array of dots — rows x cols. Used for early multiplication,
   * "groups of" introduction, and equal-share / repeated-addition modelling.
   *
   * @param {number} rows
   * @param {number} cols
   * @param {object} [opts] { cellSize, dotRatio, color, gap, withFrame }
   */
  function addArray(slide, x, y, rows, cols, opts) {
    const o = opts || {};
    const cellSize = o.cellSize || 0.45;
    const gap = o.gap != null ? o.gap : 0.05;
    const dotRatio = o.dotRatio || 0.7;
    const color = o.color || C.PRIMARY;
    const withFrame = Boolean(o.withFrame);
    const totalW = cols * cellSize + (cols - 1) * gap;
    const totalH = rows * cellSize + (rows - 1) * gap;
    validateBounds("addArray", x, y, totalW, totalH);

    if (withFrame) {
      const pad = cellSize * 0.25;
      slide.addShape("rect", {
        x: x - pad, y: y - pad, w: totalW + pad * 2, h: totalH + pad * 2,
        fill: { color: C.WHITE },
        line: { color: o.frameColor || C.CHARCOAL, width: 1.2 },
      });
    }

    const dotW = cellSize * dotRatio;
    for (let r = 0; r < rows; r += 1) {
      for (let c = 0; c < cols; c += 1) {
        const cx = x + c * (cellSize + gap) + cellSize / 2;
        const cy = y + r * (cellSize + gap) + cellSize / 2;
        slide.addShape("roundRect", {
          x: cx - dotW / 2, y: cy - dotW / 2,
          w: dotW, h: dotW, rectRadius: dotW / 2,
          fill: { color },
          line: { color, width: 0.2 },
        });
      }
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Base-10 / MAB blocks                                               */
  /* ------------------------------------------------------------------ */

  /**
   * Draw hundreds-flat (10x10), tens-rod (1x10), and ones-cubes laid out
   * left to right with a small gap. Standard MAB representation for place
   * value teaching.
   *
   * @param {number} hundreds  Number of hundreds flats
   * @param {number} tens      Number of tens rods
   * @param {number} ones      Number of ones cubes
   * @param {object} [opts]    { unit, color, gap }
   */
  function addBaseTenBlocks(slide, x, y, hundreds, tens, ones, opts) {
    const o = opts || {};
    const unit = o.unit || 0.12;
    const color = o.color || C.SECONDARY;
    const borderColor = o.borderColor || C.CHARCOAL;
    const groupGap = o.gap != null ? o.gap : 0.18;
    const lineW = 0.6;
    let cursorX = x;
    const flatSize = unit * 10;

    function drawBlock(bx, by, bw, bh) {
      slide.addShape("rect", {
        x: bx, y: by, w: bw, h: bh,
        fill: { color },
        line: { color: borderColor, width: lineW },
      });
    }

    function drawGrid(bx, by, cols, rows) {
      drawBlock(bx, by, cols * unit, rows * unit);
      // Faint grid lines so cubes are visible without dominating
      for (let i = 1; i < cols; i += 1) {
        slide.addShape("line", {
          x: bx + i * unit, y: by, w: 0, h: rows * unit,
          line: { color: C.WHITE, width: 0.4 },
        });
      }
      for (let i = 1; i < rows; i += 1) {
        slide.addShape("line", {
          x: bx, y: by + i * unit, w: cols * unit, h: 0,
          line: { color: C.WHITE, width: 0.4 },
        });
      }
    }

    // Hundreds: stack vertically up to 3 across, then wrap (rare for primary)
    for (let i = 0; i < hundreds; i += 1) {
      drawGrid(cursorX, y, 10, 10);
      cursorX += flatSize + 0.05;
    }
    if (hundreds > 0) cursorX += groupGap - 0.05;

    // Tens rods: 1 wide x 10 tall, side by side
    for (let i = 0; i < tens; i += 1) {
      drawGrid(cursorX, y, 1, 10);
      cursorX += unit + 0.04;
    }
    if (tens > 0) cursorX += groupGap - 0.04;

    // Ones cubes: 1x1, in a row
    for (let i = 0; i < ones; i += 1) {
      drawGrid(cursorX, y, 1, 1);
      cursorX += unit + 0.04;
    }

    validateBounds("addBaseTenBlocks", x, y, cursorX - x, flatSize);
  }

  /* ------------------------------------------------------------------ */
  /*  Chip row (evenly spaced choice chips)                              */
  /* ------------------------------------------------------------------ */

  /**
   * Row of evenly spaced chips (rounded rect + centred label). Use for
   * fraction choices, number choices, word options - NEVER lay these out
   * as one string with runs of spaces.
   *
   * @param {number}            w      Total row width (inches)
   * @param {(string|number)[]} items  Chip labels
   * @param {object}            [opts] { gap, chipH, fontSize, fillColor, borderColor, textColor }
   * @returns {object} { chipW, chipH }
   */
  function addChipRow(slide, x, y, w, items, opts) {
    const o = opts || {};
    const list = (Array.isArray(items) ? items : [items]).map((v) => String(v));
    if (!list.length) return { chipW: 0, chipH: 0 };
    const gap = o.gap != null ? o.gap : 0.14;
    const chipH = o.chipH || byBand(sz, 0.8, 0.68, 0.56);
    const chipW = (w - gap * (list.length - 1)) / list.length;
    const fontSize = o.fontSize || byBand(sz, 30, 24, 20);
    const fillColor = o.fillColor || C.WHITE;
    const borderColor = o.borderColor || C.PRIMARY;
    const textColor = o.textColor || C.PRIMARY;
    validateBounds("addChipRow", x, y, w, chipH);

    list.forEach((label, i) => {
      const cx = x + i * (chipW + gap);
      slide.addShape("roundRect", {
        x: cx, y, w: chipW, h: chipH, rectRadius: 0.08,
        fill: { color: fillColor },
        line: { color: borderColor, width: 1.4 },
      });
      slide.addText(label, {
        x: cx, y, w: chipW, h: chipH,
        fontSize, fontFace: FONT_B, color: textColor, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
    });
    return { chipW, chipH };
  }

  /* ------------------------------------------------------------------ */
  /*  Grouped counters ("groups of" division / multiplication)           */
  /* ------------------------------------------------------------------ */

  /**
   * `groups` framed groups of `per` counters each, laid out left to right.
   * The framing makes the grouping readable at a glance (megaprompt
   * Foundation visual clarity rule: organised, never scattered).
   *
   * @param {number} groups  Number of groups
   * @param {number} per     Counters per group
   * @param {object} [opts]  { dot, gap, groupGap, dotColor, frameColor }
   * @returns {object} { totalW, groupW }
   */
  function addGroupedCounters(slide, x, y, groups, per, opts) {
    const o = opts || {};
    const dot = o.dot || byBand(sz, 0.34, 0.3, 0.26);
    const gap = o.gap != null ? o.gap : 0.12;
    const framePad = 0.10;
    const groupW = per * dot + (per - 1) * gap + framePad * 2;
    const groupGap = o.groupGap != null ? o.groupGap : 0.24;
    const dotColor = o.dotColor || C.ACCENT;
    const frameColor = o.frameColor || C.SECONDARY;
    const totalW = groups * groupW + (groups - 1) * groupGap;
    validateBounds("addGroupedCounters", x, y - framePad, totalW, dot + framePad * 2);

    for (let g = 0; g < groups; g += 1) {
      const gx = x + g * (groupW + groupGap);
      slide.addShape("roundRect", {
        x: gx, y: y - framePad, w: groupW, h: dot + framePad * 2, rectRadius: 0.06,
        fill: { color: C.BG_LIGHT || C.WHITE },
        line: { color: frameColor, width: 1.2 },
      });
      for (let i = 0; i < per; i += 1) {
        const cx = gx + framePad + i * (dot + gap);
        slide.addShape("roundRect", {
          x: cx, y, w: dot, h: dot, rectRadius: dot / 2,
          fill: { color: dotColor },
          line: { color: dotColor, width: 0.2 },
        });
      }
    }
    return { totalW, groupW };
  }

  /* ------------------------------------------------------------------ */
  /*  Part-Part-Whole mat                                                */
  /* ------------------------------------------------------------------ */

  /**
   * Classic PPW mat: whole box on top, two part boxes below, connectors.
   * Pass null/undefined for any value to leave that box blank (board
   * build / student fills it in).
   *
   * @param {number} w       Mat width (inches)
   * @param {number} h       Mat height (inches)
   * @param {object} values  { whole, partA, partB } (any may be null for blank)
   * @param {object} [opts]  { captionFontSize, valueFontSize, borderColor, wholeColor, partColor }
   */
  function addPartPartWholeMat(slide, x, y, w, h, values, opts) {
    const v = values || {};
    const o = opts || {};
    // Captions are tiny identifiers; the VALUE is the hero of each box
    const captionFontSize = o.captionFontSize || 11;
    const valueFontSize = o.valueFontSize || byBand(sz, 30, 26, 22);
    const borderColor = o.borderColor || C.CHARCOAL;
    const wholeColor = o.wholeColor || C.PRIMARY;
    const partColor = o.partColor || C.SECONDARY;
    validateBounds("addPartPartWholeMat", x, y, w, h);

    const boxH = h * 0.42;
    const wholeW = w * 0.5;
    const wholeX = x + (w - wholeW) / 2;
    const partW = w * 0.46;
    const partY = y + h - boxH;

    function drawBox(bx, by, bw, caption, value, stripColor) {
      const capH = 0.2;
      slide.addShape("roundRect", {
        x: bx, y: by, w: bw, h: boxH, rectRadius: 0.06,
        fill: { color: C.WHITE },
        line: { color: borderColor, width: 1.4 },
      });
      slide.addShape("rect", {
        x: bx, y: by, w: 0.06, h: boxH,
        fill: { color: stripColor },
      });
      slide.addText(caption, {
        x: bx + 0.12, y: by + 0.03, w: bw - 0.2, h: capH,
        fontSize: captionFontSize, fontFace: FONT_B, color: stripColor, bold: true, margin: 0,
      });
      if (value != null && value !== "") {
        // Value sits BELOW the caption row so the two text boxes never overlap
        const valueY = by + capH + 0.05;
        slide.addText(String(value), {
          x: bx, y: valueY, w: bw, h: Math.max(0.2, by + boxH - valueY - 0.05),
          fontSize: valueFontSize, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
          align: "center", valign: "middle", margin: 0,
          fit: "shrink", shrinkText: true,
        });
      }
    }

    // Connectors first so boxes sit on top (flipH for the leftward line;
    // negative shape widths corrupt PPTX files)
    slide.addShape("line", {
      x: x + partW / 2, y: y + boxH, w: w / 2 - partW / 2, h: partY - y - boxH,
      flipH: true,
      line: { color: borderColor, width: 1.6 },
    });
    slide.addShape("line", {
      x: x + w / 2, y: y + boxH, w: w / 2 - partW / 2, h: partY - y - boxH,
      line: { color: borderColor, width: 1.6 },
    });

    drawBox(wholeX, y, wholeW, "Whole", v.whole, wholeColor);
    drawBox(x, partY, partW, "Part", v.partA, partColor);
    drawBox(x + w - partW, partY, partW, "Part", v.partB, partColor);
  }

  return {
    addTensFrame,
    addFiveFrame,
    addDotCard,
    addNumberTrack,
    addFractionStripSet,
    addArray,
    addBaseTenBlocks,
    addChipRow,
    addGroupedCounters,
    addPartPartWholeMat,
  };
}

module.exports = { createManipulatives };
