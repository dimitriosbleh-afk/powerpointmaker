"use strict";

/**
 * OG extension sheets on the school worksheet layout (MEGA_PROMPT 86): each
 * task on its own numbered card with ruled lines or a table, then the answer
 * key in the same layout with the answers in colour.
 *
 *   node og_planner/render_extension.js <job.json>
 *
 * job: { resource, out, week, info, footer } - resource is a worksheet spec
 * (sections of items) written by enabler_pdf.extension_sheet().  The student
 * sheet comes first - one page, or two printed as one sheet double-sided (86) -
 * then the answer key, which paginates the same way, so the student pages are
 * always the first half of the file (print_pack.py relies on that).
 */

const fs = require("fs");
const P = require("../themes/pdf_helpers");
const { createTheme, weekToVariant } = require("../themes/factory");
const { renderWorksheet } = require("../themes/lesson/worksheetLayout");

async function main() {
  const job = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
  const T = createTheme("literacy", "grade56", weekToVariant(job.week || 1));
  const ctx = { C: T.C, sz: T.S, lessonInfo: job.info || "", footer: job.footer || "" };
  const resource = Object.assign({}, job.resource, { maxPages: 2 });
  const doc = P.createPdf({ title: resource.title });
  doc.page.margins = { top: 0, bottom: 0, left: 0, right: 0 };
  const student = renderWorksheet(doc, resource, Object.assign({}, ctx, { isKey: false })).pages;
  doc.addPage({ size: "A4", margin: 0 });
  const key = renderWorksheet(doc, resource, Object.assign({}, ctx, { isKey: true })).pages;
  if (key !== student) throw new Error(`answer key runs ${key} pages, the sheet ${student}`);
  await P.writePdf(doc, job.out);
}

main().catch((err) => {
  console.error(err.message);
  process.exit(1);
});
