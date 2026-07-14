"use strict";

const path = require("path");

const {
  PAGE,
  createPdf,
  writePdf,
  hex,
  lighten,
} = require("./pdf_helpers");
const { sanitizeTeacherNotes } = require("./core/notes");

const DEFAULT_FILE_NAME = "Teacher Week Brief.pdf";
const COLORS = {
  INK: "1F2933",
  MUTED: "5F6B76",
  NAVY: "203A5F",
  TEAL: "2D7E78",
  GOLD: "D5A43B",
  CORAL: "C65D57",
  PAPER: "F7F4EE",
  WHITE: "FFFFFF",
  RULE: "D9DEE3",
};

function cleanLine(value) {
  return sanitizeTeacherNotes(String(value == null ? "" : value))
    .replace(/\s+/g, " ")
    .trim();
}

function cleanList(value) {
  const items = value == null ? [] : Array.isArray(value) ? value : [value];
  return items.map(cleanLine).filter(Boolean);
}

function requireText(value, label, maxChars) {
  const text = cleanLine(value);
  if (!text) throw new Error(`Teacher brief '${label}' is required.`);
  if (text.length > maxChars) {
    throw new Error(
      `Teacher brief '${label}' exceeds ${maxChars} characters (${text.length}). ` +
      "Shorten it so the brief remains a genuine one-page preparation tool."
    );
  }
  return text;
}

function requireList(value, label, minItems, maxItems, maxCharsPerItem) {
  const items = cleanList(value);
  if (items.length < minItems || items.length > maxItems) {
    throw new Error(
      `Teacher brief '${label}' requires ${minItems}-${maxItems} item(s); received ${items.length}.`
    );
  }
  items.forEach((item, index) => {
    if (item.length > maxCharsPerItem) {
      throw new Error(
        `Teacher brief '${label}' item ${index + 1} exceeds ${maxCharsPerItem} characters ` +
        `(${item.length}).`
      );
    }
  });
  return items;
}

function normaliseFileName(value) {
  const fileName = cleanLine(value || DEFAULT_FILE_NAME);
  if (path.basename(fileName) !== fileName || !fileName.toLowerCase().endsWith(".pdf")) {
    throw new Error("Teacher brief file_name must be a plain PDF filename, not a path.");
  }
  return fileName;
}

function normaliseTeacherBrief(config, opts) {
  const c = config || {};
  const o = opts || {};
  const lessonSessions = cleanList(o.lessonSessions).map((value) => Number.parseInt(value, 10));

  const sessions = Array.isArray(c.sessions) ? c.sessions.map((entry, index) => {
    const session = Number.parseInt(entry && entry.session, 10);
    if (!Number.isInteger(session) || session < 1) {
      throw new Error(`Teacher brief session ${index + 1} needs a positive integer 'session'.`);
    }
    return {
      session,
      focus: requireText(entry.focus, `sessions[${index}].focus`, 90),
      students: requireText(entry.students, `sessions[${index}].students`, 120),
      check: requireText(entry.check, `sessions[${index}].check`, 120),
      respond: requireText(entry.respond, `sessions[${index}].respond`, 150),
    };
  }) : [];

  if (sessions.length < 2 || sessions.length > 5) {
    throw new Error(`Teacher brief requires 2-5 session rows; received ${sessions.length}.`);
  }

  const seen = new Set();
  sessions.forEach((session) => {
    if (seen.has(session.session)) {
      throw new Error(`Teacher brief repeats session ${session.session}.`);
    }
    seen.add(session.session);
  });

  if (lessonSessions.length) {
    const expected = [...lessonSessions].sort((a, b) => a - b);
    const actual = sessions.map((entry) => entry.session).sort((a, b) => a - b);
    if (expected.join(",") !== actual.join(",")) {
      throw new Error(
        `Teacher brief sessions (${actual.join(", ")}) do not match manifest lessons ` +
        `(${expected.join(", ")}).`
      );
    }
  }

  return {
    fileName: normaliseFileName(c.file_name),
    title: requireText(c.title || "Teacher Week Brief", "title", 60),
    unit: requireText(c.unit || o.unitTitle, "unit", 90),
    grade: requireText(c.grade, "grade", 30),
    subject: requireText(c.subject, "subject", 40),
    curriculum: requireText(c.curriculum, "curriculum", 100),
    throughline: requireText(c.throughline, "throughline", 260),
    highLeverageMoves: requireList(c.high_leverage_moves, "high_leverage_moves", 3, 3, 125),
    sessions,
    essentialKnowledge: requireList(c.essential_knowledge, "essential_knowledge", 1, 3, 150),
    misconceptions: requireList(c.misconceptions, "misconceptions", 1, 3, 145),
    responseRule: requireText(c.response_rule, "response_rule", 240),
    rehearse: requireList(c.rehearse, "rehearse", 1, 2, 150),
    materials: requireList(c.materials, "materials", 1, 5, 100),
  };
}

function textHeight(doc, text, width, fontSize, font) {
  return doc.font(font || "Sans").fontSize(fontSize).heightOfString(text, {
    width,
    lineGap: 0.5,
  });
}

function assertFits(doc, label, text, width, height, fontSize, font) {
  const measured = textHeight(doc, text, width, fontSize, font);
  if (measured > height + 0.5) {
    throw new Error(
      `Teacher brief '${label}' needs ${measured.toFixed(1)}pt but only ${height.toFixed(1)}pt is available.`
    );
  }
}

function drawText(doc, label, text, x, y, width, height, opts) {
  const o = opts || {};
  const font = o.font || "Sans";
  const fontSize = o.fontSize || 8.5;
  assertFits(doc, label, text, width, height, fontSize, font);
  doc.font(font).fontSize(fontSize).fillColor(hex(o.color || COLORS.INK)).text(text, x, y, {
    width,
    height,
    lineGap: o.lineGap == null ? 0.5 : o.lineGap,
    ellipsis: false,
  });
}

function drawSectionLabel(doc, text, x, y, width) {
  doc.font("Sans-Bold").fontSize(8.2).fillColor(hex(COLORS.NAVY)).text(text.toUpperCase(), x, y, {
    width,
    characterSpacing: 0.8,
  });
}

function drawHighLeverageMoves(doc, items, x, y, width) {
  const gap = 8;
  const cardW = (width - gap * 2) / 3;
  const cardH = 56;
  const accents = [COLORS.TEAL, COLORS.GOLD, COLORS.CORAL];

  items.forEach((item, index) => {
    const cardX = x + index * (cardW + gap);
    doc.roundedRect(cardX, y, cardW, cardH, 6).fill(hex(COLORS.WHITE));
    doc.roundedRect(cardX, y, 5, cardH, 3).fill(hex(accents[index]));
    doc.font("Sans-Bold").fontSize(8).fillColor(hex(accents[index])).text(String(index + 1), cardX + 12, y + 9);
    drawText(doc, `high_leverage_moves[${index}]`, item, cardX + 25, y + 8, cardW - 33, cardH - 16, {
      fontSize: 8.2,
    });
  });

  return y + cardH;
}

function drawSessionTable(doc, sessions, x, y, width) {
  const headerH = 20;
  const rowH = 54;
  const gap = 8;
  const sessionW = 32;
  const lessonW = 172;
  const checkW = 132;
  const respondW = width - sessionW - lessonW - checkW - gap * 3;
  const lessonX = x + sessionW + gap;
  const checkX = lessonX + lessonW + gap;
  const respondX = checkX + checkW + gap;

  doc.roundedRect(x, y, width, headerH, 4).fill(hex(COLORS.NAVY));
  doc.font("Sans-Bold").fontSize(7.6).fillColor(hex(COLORS.WHITE));
  doc.text("S", x, y + 6, { width: sessionW, align: "center" });
  doc.text("FOCUS AND STUDENT WORK", lessonX, y + 6, { width: lessonW });
  doc.text("CHECK", checkX, y + 6, { width: checkW });
  doc.text("RESPOND", respondX, y + 6, { width: respondW });

  let rowY = y + headerH + 4;
  sessions.forEach((session, index) => {
    const fill = index % 2 === 0 ? COLORS.WHITE : lighten(COLORS.NAVY, 0.965).replace("#", "");
    doc.roundedRect(x, rowY, width, rowH, 4).fill(hex(fill));
    doc.circle(x + sessionW / 2, rowY + rowH / 2, 12).fill(hex(COLORS.TEAL));
    doc.font("Sans-Bold").fontSize(9).fillColor(hex(COLORS.WHITE)).text(String(session.session), x + 5, rowY + 21, {
      width: sessionW - 10,
      align: "center",
    });

    drawText(doc, `session ${session.session} focus`, session.focus, lessonX, rowY + 7, lessonW, 16, {
      font: "Sans-Bold",
      fontSize: 8.4,
      color: COLORS.NAVY,
    });
    drawText(doc, `session ${session.session} students`, session.students, lessonX, rowY + 23, lessonW, rowH - 29, {
      fontSize: 7.8,
      color: COLORS.MUTED,
    });
    drawText(doc, `session ${session.session} check`, session.check, checkX, rowY + 7, checkW, rowH - 14, {
      fontSize: 7.8,
    });
    drawText(doc, `session ${session.session} respond`, session.respond, respondX, rowY + 7, respondW, rowH - 14, {
      fontSize: 7.8,
    });
    rowY += rowH + 4;
  });

  return rowY - 4;
}

function formatBullets(items) {
  return items.map((item) => `- ${item}`).join("\n");
}

function drawPreparationCard(doc, title, groups, x, y, width, height, accent) {
  doc.roundedRect(x, y, width, height, 7).fill(hex(COLORS.WHITE));
  doc.roundedRect(x, y, 5, height, 3).fill(hex(accent));
  drawSectionLabel(doc, title, x + 15, y + 13, width - 28);

  let cursorY = y + 34;
  groups.forEach((group, index) => {
    const body = group.items ? formatBullets(group.items) : group.text;
    const bodyHeight = textHeight(doc, body, width - 30, 7.7, "Sans");
    const groupHeight = 13 + bodyHeight + (index === groups.length - 1 ? 0 : 11);
    if (cursorY + groupHeight > y + height - 10) {
      throw new Error(`Teacher brief '${title}' card overflows. Shorten ${group.label}.`);
    }
    doc.font("Sans-Bold").fontSize(7.8).fillColor(hex(accent)).text(group.label.toUpperCase(), x + 15, cursorY, {
      width: width - 30,
      characterSpacing: 0.35,
    });
    cursorY += 13;
    drawText(doc, `${title} ${group.label}`, body, x + 15, cursorY, width - 30, bodyHeight + 1, {
      fontSize: 7.7,
      color: COLORS.INK,
    });
    cursorY += bodyHeight + 11;
  });
}

function renderTeacherWeekBrief(doc, data) {
  const margin = 36;
  const width = PAGE.W - margin * 2;

  doc.rect(0, 0, PAGE.W, 79).fill(hex(COLORS.NAVY));
  drawText(doc, "title", data.title, margin, 22, 350, 28, {
    font: "Sans-Bold",
    fontSize: 22,
    color: COLORS.WHITE,
  });
  drawText(doc, "unit", data.unit, margin, 52, 360, 14, {
    fontSize: 9,
    color: COLORS.WHITE,
  });
  doc.roundedRect(PAGE.W - margin - 92, 24, 92, 27, 13).fill(hex(COLORS.TEAL));
  doc.font("Sans-Bold").fontSize(8).fillColor(hex(COLORS.WHITE)).text("3 MIN READ", PAGE.W - margin - 92, 33, {
    width: 92,
    align: "center",
  });
  drawText(doc, "metadata", `${data.grade} | ${data.subject} | ${data.curriculum}`, PAGE.W - margin - 220, 57, 220, 12, {
    fontSize: 7.2,
    color: COLORS.WHITE,
  });

  let y = 91;
  doc.roundedRect(margin, y, width, 52, 7).fill(hex(COLORS.PAPER));
  drawSectionLabel(doc, "Week throughline", margin + 14, y + 10, width - 28);
  drawText(doc, "throughline", data.throughline, margin + 14, y + 25, width - 28, 21, {
    fontSize: 9,
  });

  y += 66;
  drawSectionLabel(doc, "The 20% to nail", margin, y, width);
  y = drawHighLeverageMoves(doc, data.highLeverageMoves, margin, y + 15, width) + 12;

  drawSectionLabel(doc, "Week at a glance", margin, y, width);
  y = drawSessionTable(doc, data.sessions, margin, y + 15, width) + 12;

  const footerTop = PAGE.H - 33;
  const bottomH = footerTop - y - 8;
  if (bottomH < 150) {
    throw new Error(`Teacher brief bottom preparation zone is too short (${bottomH.toFixed(1)}pt).`);
  }
  const colGap = 10;
  const colW = (width - colGap) / 2;
  drawPreparationCard(doc, "Know before teaching", [
    { label: "Essential knowledge", items: data.essentialKnowledge },
    { label: "Likely misconceptions", items: data.misconceptions },
  ], margin, y, colW, bottomH, COLORS.NAVY);
  drawPreparationCard(doc, "Respond and rehearse", [
    { label: "Response rule", text: data.responseRule },
    { label: "Rehearse", items: data.rehearse },
    { label: "Materials", items: data.materials },
  ], margin + colW + colGap, y, colW, bottomH, COLORS.TEAL);

  doc.moveTo(margin, PAGE.H - 24).lineTo(PAGE.W - margin, PAGE.H - 24).lineWidth(0.5).strokeColor(hex(COLORS.RULE)).stroke();
  doc.font("Sans").fontSize(6.8).fillColor(hex(COLORS.MUTED)).text(
    "Prepare the content, rehearse the hardest moment, then adapt from evidence without changing the learning goal.",
    margin,
    PAGE.H - 18,
    { width, align: "center" }
  );

  const pageRange = doc.bufferedPageRange();
  if (pageRange.count !== 1) {
    throw new Error(`Teacher brief rendered ${pageRange.count} pages; exactly one is required.`);
  }
}

async function generateTeacherWeekBrief(config, outputPath, opts) {
  const data = normaliseTeacherBrief(config, opts);
  // The brief uses explicit page coordinates, including a footer below the
  // standard worksheet margin. A 1pt document margin prevents PDFKit from
  // auto-adding a second page when that footer is drawn.
  const doc = createPdf({ title: data.title, author: "Lesson Generator", margin: 1 });
  renderTeacherWeekBrief(doc, data);
  await writePdf(doc, outputPath);
  return { outputPath, fileName: data.fileName, data };
}

module.exports = {
  DEFAULT_FILE_NAME,
  normaliseTeacherBrief,
  renderTeacherWeekBrief,
  generateTeacherWeekBrief,
};
