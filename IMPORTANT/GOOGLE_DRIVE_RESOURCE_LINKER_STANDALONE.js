/**
 * Standalone Google Apps Script - Drive Resource Linker
 *
 * Use this when the files already live in Google Drive and you do not want to
 * paste Apps Script into each new Google Slides deck.
 *
 * Setup:
 * 1. Create one standalone Apps Script project at https://script.google.com.
 * 2. Paste this file into Code.gs.
 * 3. Replace appsscript.json with IMPORTANT/GOOGLE_DRIVE_RESOURCE_LINKER_appsscript.json.
 * 4. Deploy > New deployment > Web app.
 * 5. For personal use, set "Execute as" to "Me" and restrict access to yourself
 *    or your school domain. Do not expose an "execute as me" Drive tool publicly.
 * 6. Authorize this Apps Script project once.
 * 7. Open the web app URL, paste a Google Slides URL/ID and a Drive PDF folder
 *    URL/ID, then run the linker.
 *
 * This still requires Google's OAuth authorization once for this Apps Script
 * project. The benefit is that the same authorized project can be reused across
 * decks, so you are not approving a newly pasted script inside every deck.
 */

var SEARCH_SUBFOLDERS_DEFAULT = true;
var MAX_MATCHES_IN_REPORT = 200;

function doGet() {
  return HtmlService.createHtmlOutput(buildWebAppHtml_())
    .setTitle("Drive Resource Linker");
}

function buildWebAppHtml_() {
  return [
    "<!doctype html>",
    '<html><head><base target="_top">',
    '<meta name="viewport" content="width=device-width,initial-scale=1">',
    "<style>",
    "body{font-family:Arial,sans-serif;margin:0;background:#f7f7f8;color:#1f2933;}",
    ".wrap{max-width:820px;margin:0 auto;padding:28px 18px 40px;}",
    "h1{font-size:24px;margin:0 0 8px;}",
    "p{font-size:14px;line-height:1.5;margin:0 0 18px;color:#52606d;}",
    "form{background:#fff;border:1px solid #d9e2ec;border-radius:8px;padding:18px;}",
    "label{display:block;font-size:13px;font-weight:700;margin:0 0 6px;}",
    "input[type=text]{box-sizing:border-box;width:100%;border:1px solid #bcccdc;border-radius:6px;padding:10px;font-size:14px;margin:0 0 14px;}",
    ".row{display:flex;gap:8px;align-items:center;margin:2px 0 16px;}",
    ".row label{font-weight:400;margin:0;}",
    "button{border:0;border-radius:6px;background:#2563eb;color:white;font-weight:700;font-size:14px;padding:10px 14px;cursor:pointer;}",
    "button[disabled]{background:#9fb3c8;cursor:default;}",
    "pre{white-space:pre-wrap;background:#101828;color:#e4e7ec;border-radius:8px;padding:14px;min-height:160px;font-size:12px;line-height:1.45;margin:16px 0 0;}",
    ".hint{font-size:12px;color:#627d98;margin-top:-8px;margin-bottom:14px;}",
    "</style></head><body>",
    '<div class="wrap">',
    "<h1>Drive Resource Linker</h1>",
    "<p>Link resource names in a Google Slides deck to matching PDF files in a Google Drive folder. This reusable web app is authorized once, then reused for future decks.</p>",
    '<form id="linkForm">',
    '<label for="presentationInput">Google Slides URL or ID</label>',
    '<input id="presentationInput" type="text" autocomplete="off" placeholder="https://docs.google.com/presentation/d/.../edit" required>',
    '<div class="hint">Use the converted Google Slides file, not a raw .pptx file URL.</div>',
    '<label for="folderInput">Drive PDF folder URL or ID</label>',
    '<input id="folderInput" type="text" autocomplete="off" placeholder="https://drive.google.com/drive/folders/..." required>',
    '<div class="row"><input id="searchSubfolders" type="checkbox" checked><label for="searchSubfolders">Search subfolders for PDFs</label></div>',
    '<button id="runButton" type="submit">Link resources</button>',
    "</form>",
    '<pre id="report">Paste the two Drive links, then run the linker.</pre>',
    "</div>",
    "<script>",
    "document.getElementById('linkForm').addEventListener('submit',function(event){event.preventDefault();runLinker();});",
    "function runLinker(){",
    "var button=document.getElementById('runButton');",
    "var report=document.getElementById('report');",
    "button.disabled=true;",
    "report.textContent='Scanning Drive PDFs and updating the presentation...';",
    "google.script.run",
    ".withSuccessHandler(function(result){button.disabled=false;report.textContent=result && result.report ? result.report : String(result || 'Done.');})",
    ".withFailureHandler(function(error){button.disabled=false;report.textContent=(error && error.message) ? error.message : String(error);})",
    ".linkResourcesFromDriveWebApp({",
    "presentationInput:document.getElementById('presentationInput').value,",
    "folderInput:document.getElementById('folderInput').value,",
    "searchSubfolders:document.getElementById('searchSubfolders').checked",
    "});",
    "}",
    "</script>",
    "</body></html>"
  ].join("");
}

function linkResourcesFromDriveWebApp(form) {
  var presentationInput = String(form && form.presentationInput || "").trim();
  var folderInput = String(form && form.folderInput || "").trim();
  var searchSubfolders = !(form && (form.searchSubfolders === false || form.searchSubfolders === "false"));

  var presentationId = extractPresentationId(presentationInput);
  if (!presentationId) {
    throw new Error("Could not read a Google Slides presentation ID from the first field.");
  }

  var folderId = extractFolderId(folderInput);
  if (!folderId) {
    throw new Error("Could not read a Google Drive folder ID from the second field.");
  }

  var pdfFolder = DriveApp.getFolderById(folderId);
  var pdfMap = buildPdfMapFromDriveFolder_(pdfFolder, searchSubfolders);
  var pdfCount = uniquePdfCount_(pdfMap);

  if (pdfCount === 0) {
    throw new Error("No PDFs were found in the selected Drive folder.");
  }

  var presentation = SlidesApp.openById(presentationId);
  return linkResourcesInPresentation_(presentation, pdfMap, pdfCount, pdfFolder.getName(), searchSubfolders);
}

function extractPresentationId(input) {
  if (!input) return null;

  var text = String(input).trim();
  var presentationMatch = text.match(/\/presentation\/d\/([a-zA-Z0-9_-]+)/);
  if (presentationMatch) return presentationMatch[1];

  var idParamMatch = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idParamMatch) return idParamMatch[1];

  if (/^[a-zA-Z0-9_-]{10,}$/.test(text)) return text;

  return null;
}

function extractFolderId(input) {
  if (!input) return null;

  var text = String(input).trim();
  var folderMatch = text.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch) return folderMatch[1];

  var idParamMatch = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idParamMatch) return idParamMatch[1];

  if (/^[a-zA-Z0-9_-]{10,}$/.test(text)) return text;

  return null;
}

function buildPdfMapFromDriveFolder_(folder, searchSubfolders) {
  var map = {};
  var visitedFolders = {};
  collectPdfFilesFromFolder_(folder, map, searchSubfolders, visitedFolders);
  return map;
}

function collectPdfFilesFromFolder_(folder, map, searchSubfolders, visitedFolders) {
  var folderId = folder.getId();
  if (visitedFolders[folderId]) return;
  visitedFolders[folderId] = true;

  var files = folder.getFilesByType(MimeType.PDF);
  while (files.hasNext()) {
    addPdfFileToMap_(files.next(), map);
  }

  if (!searchSubfolders) return;

  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    collectPdfFilesFromFolder_(subfolders.next(), map, searchSubfolders, visitedFolders);
  }
}

function addPdfFileToMap_(file, map) {
  var name = file.getName();
  var stem = stripExtension(name);
  var normalizedStem = normalizeName(stem);
  var mapped = {
    id: file.getId(),
    name: name,
    url: file.getUrl(),
    stem: stem
  };

  if (!map[normalizedStem]) {
    map[normalizedStem] = mapped;
  }

  if (normalizedStem !== stem && !map[stem]) {
    map[stem] = mapped;
  }
}

function linkResourcesInPresentation_(presentation, pdfMap, pdfCount, folderName, searchSubfolders) {
  var slides = presentation.getSlides();
  var updatedCount = 0;
  var matchLog = [];

  for (var s = 0; s < slides.length; s++) {
    var slide = slides[s];
    var elements = slide.getPageElements();

    for (var e = 0; e < elements.length; e++) {
      var element = elements[e];

      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        updatedCount += linkShapeRuns_(element.asShape(), s, pdfMap, matchLog);
      }

      if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
        updatedCount += linkTableRuns_(element.asTable(), s, pdfMap, matchLog);
      }
    }
  }

  var report = "Resource Link Update Complete\n\n";
  report += "Presentation: " + presentation.getName() + "\n";
  report += "PDF folder: " + folderName + "\n";
  report += "Searched subfolders: " + (searchSubfolders ? "yes" : "no") + "\n";
  report += "PDFs found: " + pdfCount + "\n";
  report += "Links updated: " + updatedCount + "\n\n";

  if (matchLog.length > 0) {
    report += "Matches:\n";
    for (var m = 0; m < matchLog.length && m < MAX_MATCHES_IN_REPORT; m++) {
      report += "  " + matchLog[m] + "\n";
    }
    if (matchLog.length > MAX_MATCHES_IN_REPORT) {
      report += "  ... " + (matchLog.length - MAX_MATCHES_IN_REPORT) + " more\n";
    }
  }

  if (updatedCount === 0) {
    report += "No matches found. Check that the resource slide text includes the PDF filenames or filename stems.\n";
  }

  return {
    pdfCount: pdfCount,
    updatedCount: updatedCount,
    matches: matchLog,
    report: report
  };
}

function linkShapeRuns_(shape, slideIndex, pdfMap, matchLog) {
  var textRange;
  try {
    textRange = shape.getText();
  } catch (e) {
    return 0;
  }

  return linkTextRuns_(textRange.getRuns(), slideIndex, pdfMap, matchLog, "shape");
}

function linkTableRuns_(table, slideIndex, pdfMap, matchLog) {
  var updatedCount = 0;

  for (var row = 0; row < table.getNumRows(); row++) {
    for (var col = 0; col < table.getNumColumns(); col++) {
      try {
        updatedCount += linkTextRuns_(table.getCell(row, col).getText().getRuns(), slideIndex, pdfMap, matchLog, "table");
      } catch (e) {
        // Merged or inaccessible cells can be skipped.
      }
    }
  }

  return updatedCount;
}

function linkTextRuns_(runs, slideIndex, pdfMap, matchLog, location) {
  var updatedCount = 0;

  for (var r = 0; r < runs.length; r++) {
    var run = runs[r];
    var runText = run.asString().trim();
    var link = run.getTextStyle().getLink();

    var matched = findBestMatch(runText, pdfMap);
    if (!matched && link) {
      matched = findLocalPathMatch_(link, pdfMap);
    }

    if (!matched) continue;

    run.getTextStyle().setLinkUrl(matched.url);
    updatedCount++;
    matchLog.push("Slide " + (slideIndex + 1) + " (" + location + "): \"" + truncate(runText, 48) + "\" -> " + matched.name);
  }

  return updatedCount;
}

function findLocalPathMatch_(link, pdfMap) {
  var linkUrl = "";
  try {
    linkUrl = link.getUrl() || "";
  } catch (e) {
    return null;
  }

  if (!linkUrl || !isLocalPath(linkUrl)) {
    return null;
  }

  var localFilename = extractFilenameFromPath(linkUrl);
  var localStem = stripExtension(localFilename);
  return pdfMap[localStem] || pdfMap[normalizeName(localStem)] || null;
}

function findBestMatch(text, pdfMap) {
  if (!text || text.length < 3) return null;

  var normalizedText = normalizeName(text);
  var minPartialMatchLength = 8;

  if (pdfMap[normalizedText]) return pdfMap[normalizedText];

  var keys = Object.keys(pdfMap);
  var bestMatch = null;
  var bestLength = 0;

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var normalizedKey = normalizeName(key);

    if (normalizedKey.length >= minPartialMatchLength && normalizedText.indexOf(normalizedKey) !== -1) {
      if (normalizedKey.length > bestLength) {
        bestMatch = pdfMap[key];
        bestLength = normalizedKey.length;
      }
    }
  }

  return bestMatch;
}

function uniquePdfCount_(pdfMap) {
  var seen = {};
  var keys = Object.keys(pdfMap || {});

  for (var i = 0; i < keys.length; i++) {
    var entry = pdfMap[keys[i]];
    if (!entry || !entry.url) continue;
    seen[entry.id || entry.url] = true;
  }

  return Object.keys(seen).length;
}

function normalizeName(name) {
  return String(name || "")
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\.pdf$/i, "")
    .replace(/\s+/g, " ")
    .trim();
}

function stripExtension(filename) {
  return String(filename || "").replace(/\.[^.]+$/, "");
}

function isLocalPath(url) {
  if (!url) return false;
  return !/^https?:\/\//.test(url) || url.indexOf("file://") === 0;
}

function extractFilenameFromPath(pathValue) {
  if (!pathValue) return "";
  var parts = String(pathValue).replace(/\\/g, "/").split("/");
  return parts[parts.length - 1] || "";
}

function truncate(str, maxLen) {
  if (!str) return "";
  if (str.length <= maxLen) return str;
  return str.substring(0, maxLen - 3) + "...";
}
