#!/usr/bin/env node

/**
 * Local Google Slides resource linker.
 *
 * This avoids the repeated Apps Script consent problem by using one local OAuth
 * token saved under .google/. Google still requires consent once for the OAuth
 * client, but repeat runs reuse the refresh token.
 */

const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const process = require("process");

const { authenticate } = require("@google-cloud/local-auth");
const { google } = require("googleapis");
const JSZip = require("jszip");

const SCOPES = [
  "https://www.googleapis.com/auth/presentations",
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/drive.file",
];

const DEFAULT_STATE_DIR = path.join(process.cwd(), ".google");
const DEFAULT_TOKEN_PATH = path.join(DEFAULT_STATE_DIR, "google_slides_resource_token.json");
const DEFAULT_CONFIG_PATH = path.join(DEFAULT_STATE_DIR, "google_slides_resource_config.json");
const DEFAULT_CREDENTIAL_PATHS = [
  path.join(process.cwd(), "google_oauth_client.json"),
  path.join(process.cwd(), "credentials.json"),
  path.join(process.cwd(), ".google", "google_oauth_client.json"),
];
const DEFAULT_SERVICE_ACCOUNT_PATHS = [
  path.join(process.cwd(), "service-account.json"),
  path.join(process.cwd(), "google_service_account.json"),
  path.join(process.cwd(), ".google", "service-account.json"),
];

async function main() {
  const argv = parseArgs(process.argv.slice(2));
  const command = argv._[0] || "help";

  if (command === "help" || argv.help || argv.h) {
    printHelp();
    return;
  }

  if (command === "configure") {
    await configure(argv);
    return;
  }

  if (command === "auth") {
    await authorize(argv);
    console.log("Google OAuth token is ready.");
    return;
  }

  if (command === "doctor") {
    await doctor(argv);
    return;
  }

  if (command === "clear-token") {
    await clearToken(argv);
    return;
  }

  if (command === "link") {
    await linkResources(argv);
    return;
  }

  if (command === "manifest") {
    await generateManifest(argv);
    return;
  }

  if (command === "manifest-template" || command === "template-manifest") {
    await generateManifestTemplate(argv);
    return;
  }

  if (command === "patch-pptx") {
    await patchPptxLinks(argv);
    return;
  }

  if (command === "publish") {
    await publishAndLink(argv);
    return;
  }

  throw new Error(`Unknown command: ${command}`);
}

function parseArgs(args) {
  const parsed = { _: [] };

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];

    if (!arg.startsWith("--")) {
      parsed._.push(arg);
      continue;
    }

    const withoutPrefix = arg.slice(2);
    const eqIndex = withoutPrefix.indexOf("=");

    if (eqIndex !== -1) {
      parsed[withoutPrefix.slice(0, eqIndex)] = withoutPrefix.slice(eqIndex + 1);
      continue;
    }

    const next = args[i + 1];
    if (!next || next.startsWith("--")) {
      parsed[withoutPrefix] = true;
      continue;
    }

    parsed[withoutPrefix] = next;
    i++;
  }

  return parsed;
}

async function configure(argv) {
  const existing = await readJsonIfExists(getConfigPath(argv));
  const nextConfig = { ...existing };

  const presentationInput = getFirst(argv, ["presentation", "presentation-id", "deck", "deck-id"]);
  const folderInput = getFirst(argv, ["pdf-folder", "folder", "folder-id"]);
  const driveFolderInput = getFirst(argv, ["drive-folder", "parent-folder", "parent-folder-id"]);

  if (presentationInput) {
    nextConfig.presentationId = extractPresentationId(presentationInput);
    if (!nextConfig.presentationId) {
      throw new Error("Could not extract a Google Slides presentation ID.");
    }
  }

  if (folderInput) {
    nextConfig.pdfFolderId = extractFolderId(folderInput);
    if (!nextConfig.pdfFolderId) {
      throw new Error("Could not extract a Google Drive folder ID.");
    }
  }

  if (driveFolderInput) {
    nextConfig.driveFolderId = extractFolderId(driveFolderInput);
    if (!nextConfig.driveFolderId) {
      throw new Error("Could not extract a parent Google Drive folder ID.");
    }
  }

  if (argv["no-subfolders"]) {
    nextConfig.searchSubfolders = false;
  } else if (argv.subfolders) {
    nextConfig.searchSubfolders = true;
  } else if (typeof nextConfig.searchSubfolders !== "boolean") {
    nextConfig.searchSubfolders = true;
  }

  if (!nextConfig.presentationId && !nextConfig.pdfFolderId && !nextConfig.driveFolderId) {
    throw new Error("Nothing to save. Provide --presentation, --pdf-folder, and/or --drive-folder.");
  }

  await writeJson(getConfigPath(argv), nextConfig);
  console.log(`Saved config to ${getConfigPath(argv)}`);
  if (nextConfig.presentationId) console.log(`Presentation: ${nextConfig.presentationId}`);
  if (nextConfig.pdfFolderId) console.log(`PDF folder: ${nextConfig.pdfFolderId}`);
  if (nextConfig.driveFolderId) console.log(`Parent Drive folder: ${nextConfig.driveFolderId}`);
  console.log(`Search subfolders: ${nextConfig.searchSubfolders !== false}`);
}

async function clearToken(argv) {
  const tokenPath = getTokenPath(argv);
  if (fs.existsSync(tokenPath)) {
    await fsp.unlink(tokenPath);
    console.log(`Deleted ${tokenPath}`);
    return;
  }
  console.log(`No token file found at ${tokenPath}`);
}

async function doctor(argv) {
  const tokenPath = getTokenPath(argv);
  const configPath = getConfigPath(argv);
  const config = await readJsonIfExists(configPath);
  const token = await readJsonIfExists(tokenPath);
  const oauthPath = findDefaultFile(DEFAULT_CREDENTIAL_PATHS);
  const serviceAccountPath = findDefaultFile(DEFAULT_SERVICE_ACCOUNT_PATHS);

  console.log("Google Slides resource linker doctor");
  console.log("");
  console.log(`Config: ${fs.existsSync(configPath) ? configPath : "not found"}`);
  if (config.presentationId) console.log(`  Presentation: ${config.presentationId}`);
  if (config.pdfFolderId) console.log(`  PDF folder: ${config.pdfFolderId}`);
  if (config.driveFolderId) console.log(`  Parent Drive folder: ${config.driveFolderId}`);
  if (!config.presentationId && !config.pdfFolderId && !config.driveFolderId) {
    console.log("  No saved IDs yet.");
  }

  console.log("");
  console.log(`OAuth client: ${oauthPath || "not found"}`);
  console.log(`OAuth token: ${fs.existsSync(tokenPath) ? tokenPath : "not found"}`);
  if (token.refresh_token) {
    console.log("  Refresh token: present");
    console.log(`  Required scopes: ${hasRequiredScopes(token.scope) ? "present" : "missing or outdated"}`);
  }

  console.log("");
  console.log(`Service account: ${serviceAccountPath || "not found"}`);
  console.log(`Delegated subject: ${argv.subject || process.env.GOOGLE_WORKSPACE_SUBJECT || "not set"}`);

  console.log("");
  console.log("Recommended next step:");
  if (serviceAccountPath && (argv.subject || process.env.GOOGLE_WORKSPACE_SUBJECT)) {
    console.log("  Use publish/link with --service-account and --subject, or set GOOGLE_SERVICE_ACCOUNT and GOOGLE_WORKSPACE_SUBJECT.");
  } else if (token.refresh_token && hasRequiredScopes(token.scope)) {
    console.log("  Run publish or link. The saved OAuth token should avoid another browser consent prompt.");
  } else if (oauthPath) {
    console.log(`  Run: node scripts/google_slides_resource_cli.js auth --credentials "${oauthPath}"`);
  } else {
    console.log("  Create an OAuth Desktop client JSON or configure a domain-delegated service account.");
  }

  console.log("");
  console.log("Note: Apps Script pasted into a new deck will still show Google's OAuth consent screen.");
}

async function linkResources(argv) {
  const config = await readJsonIfExists(getConfigPath(argv));
  const presentationId = extractPresentationId(
    getFirst(argv, ["presentation", "presentation-id", "deck", "deck-id"]) || config.presentationId
  );
  const pdfFolderId = extractFolderId(
    getFirst(argv, ["pdf-folder", "folder", "folder-id"]) || config.pdfFolderId
  );
  const searchSubfolders = argv["no-subfolders"] ? false : config.searchSubfolders !== false;
  const dryRun = Boolean(argv["dry-run"]);

  if (!presentationId) {
    throw new Error("Missing presentation. Run configure or pass --presentation <Slides URL or ID>.");
  }
  if (!pdfFolderId) {
    throw new Error("Missing PDF folder. Run configure or pass --pdf-folder <Drive folder URL or ID>.");
  }

  const auth = await authorize(argv);
  const drive = google.drive({ version: "v3", auth });
  const slides = google.slides({ version: "v1", auth });

  const pdfMap = await buildPdfMap(drive, pdfFolderId, searchSubfolders);
  const pdfCount = new Set(Object.values(pdfMap).map((entry) => entry.id)).size;

  if (pdfCount === 0) {
    throw new Error("No PDFs found in the selected Drive folder.");
  }

  const result = await linkPresentationToPdfMap(slides, presentationId, pdfMap, dryRun);

  console.log("Resource link update complete.");
  console.log(`PDFs found: ${pdfCount}`);
  console.log(`Links ${dryRun ? "that would be updated" : "updated"}: ${result.requests.length}`);

  if (result.matchLog.length > 0) {
    console.log("");
    console.log("Matches:");
    for (const line of result.matchLog) {
      console.log(`  ${line}`);
    }
  }

  if (result.requests.length === 0) {
    console.log("");
    console.log("No matches found. Check that resource slide text matches PDF filenames.");
    console.log("PDFs found:");
    const entries = uniquePdfEntries(pdfMap).slice(0, 10);
    for (const entry of entries) {
      console.log(`  ${entry.name}`);
    }
  }
}

async function generateManifest(argv) {
  const config = await readJsonIfExists(getConfigPath(argv));
  const pdfFolderId = extractFolderId(
    getFirst(argv, ["pdf-folder", "folder", "folder-id"]) || config.pdfFolderId
  );
  const searchSubfolders = argv["no-subfolders"] ? false : config.searchSubfolders !== false;

  if (!pdfFolderId) {
    throw new Error("Missing PDF folder. Run configure or pass --pdf-folder <Drive folder URL or ID>.");
  }

  const auth = await authorize(argv);
  const drive = google.drive({ version: "v3", auth });
  const pdfMap = await buildPdfMap(drive, pdfFolderId, searchSubfolders);
  const resources = uniquePdfEntries(pdfMap).map((entry) => ({
    name: entry.name,
    url: entry.url,
  }));
  const manifest = { resources };

  await writeManifestOutput(argv, manifest);
}

async function generateManifestTemplate(argv) {
  const pptxInput = getFirst(argv, ["pptx", "deck-file", "file"]);
  const resourcesInput = getFirst(argv, ["resources", "resource-dir", "resources-dir"]);
  const names = new Set();

  if (pptxInput) {
    const pptxPath = path.resolve(pptxInput);
    if (!fs.existsSync(pptxPath)) {
      throw new Error(`PPTX file not found: ${pptxPath}`);
    }
    for (const name of await collectPptxPdfLinkNames(pptxPath)) {
      names.add(name);
    }
  }

  if (resourcesInput) {
    const resourcesDir = path.resolve(resourcesInput);
    if (!fs.existsSync(resourcesDir)) {
      throw new Error(`Resources folder not found: ${resourcesDir}`);
    }
    for (const pdfPath of await collectLocalPdfs(resourcesDir)) {
      names.add(path.basename(pdfPath));
    }
  }

  if (!pptxInput && !resourcesInput) {
    throw new Error("Pass --pptx <local .pptx path>, --resources <local folder>, or both.");
  }

  const sortedNames = Array.from(names).sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));
  if (sortedNames.length === 0) {
    throw new Error("No PDF resource names found.");
  }

  const manifest = {};
  for (const name of sortedNames) {
    manifest[name] = "";
  }

  await writeManifestOutput(argv, manifest);
}

async function patchPptxLinks(argv) {
  const pptxPath = path.resolve(getFirst(argv, ["pptx", "deck-file", "file"]) || "");
  const manifestPath = path.resolve(getFirst(argv, ["manifest", "manifest-file", "resources-manifest"]) || "");
  const dryRun = Boolean(argv["dry-run"]);

  if (!pptxPath || pptxPath === process.cwd()) {
    throw new Error("Missing PPTX path. Pass --pptx <local .pptx path>.");
  }
  if (!fs.existsSync(pptxPath)) {
    throw new Error(`PPTX file not found: ${pptxPath}`);
  }
  if (path.extname(pptxPath).toLowerCase() !== ".pptx") {
    throw new Error(`Expected a .pptx file: ${pptxPath}`);
  }
  if (!manifestPath || manifestPath === process.cwd()) {
    throw new Error("Missing manifest path. Pass --manifest <resources-manifest.json>.");
  }
  if (!fs.existsSync(manifestPath)) {
    throw new Error(`Manifest file not found: ${manifestPath}`);
  }

  const manifest = parseJsonText(await fsp.readFile(manifestPath, "utf8"));
  const resourceMap = buildResourceMapFromManifest(manifest);
  if (Object.keys(resourceMap).length === 0) {
    throw new Error("Manifest did not contain any filled PDF URLs.");
  }
  const zip = await JSZip.loadAsync(await fsp.readFile(pptxPath));
  const relNames = Object.keys(zip.files).filter((name) => name.endsWith(".rels"));
  const updates = [];
  const unmatchedPdfTargets = new Map();

  for (const relName of relNames) {
    const relFile = zip.file(relName);
    if (!relFile) continue;

    const xml = await relFile.async("string");
    if (!xml.includes("relationships/hyperlink")) continue;

    const nextXml = xml.replace(/<Relationship\b[^>]*>/g, (relationshipXml) => {
      if (!/relationships\/hyperlink/.test(relationshipXml)) {
        return relationshipXml;
      }

      const target = getXmlAttribute(relationshipXml, "Target");
      if (!target) {
        return relationshipXml;
      }

      const match = findManifestMatchForTarget(target, resourceMap);
      if (!match) {
        const filename = extractFilenameFromPath(target);
        if (path.extname(filename).toLowerCase() === ".pdf") {
          unmatchedPdfTargets.set(target, filename);
        }
        return relationshipXml;
      }

      updates.push({
        relName,
        from: target,
        to: match.url,
        name: match.name,
      });

      var updated = setXmlAttribute(relationshipXml, "Target", match.url);
      updated = setXmlAttribute(updated, "TargetMode", "External");
      return updated;
    });

    if (nextXml !== xml && !dryRun) {
      zip.file(relName, nextXml);
    }
  }

  console.log(`PPTX: ${pptxPath}`);
  console.log(`Manifest: ${manifestPath}`);
  console.log(`Hyperlinks ${dryRun ? "that would be patched" : "patched"}: ${updates.length}`);

  for (const update of updates) {
    console.log(`  ${update.name}: ${update.from} -> ${update.to}`);
  }

  if (updates.length === 0) {
    console.log("No matching PPTX resource hyperlinks were found.");
    printUnmatchedPdfTargets(unmatchedPdfTargets);
    return;
  }

  printUnmatchedPdfTargets(unmatchedPdfTargets);

  if (dryRun) {
    return;
  }

  const outputPath = getPatchOutputPath(argv, pptxPath);
  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  await fsp.writeFile(outputPath, buffer);
  console.log(`Wrote patched PPTX: ${outputPath}`);
}

async function publishAndLink(argv) {
  const config = await readJsonIfExists(getConfigPath(argv));
  const pptxPath = path.resolve(getFirst(argv, ["pptx", "deck-file", "file"]) || "");
  const resourcesDir = resolveResourcesDir(argv, pptxPath);
  const parentFolderId = extractFolderId(
    getFirst(argv, ["drive-folder", "parent-folder", "parent-folder-id"]) || config.driveFolderId
  );
  const existingPdfFolderId = extractFolderId(
    getFirst(argv, ["pdf-folder", "folder", "folder-id"]) || config.pdfFolderId
  );
  const dryRun = Boolean(argv["dry-run"]);

  if (!pptxPath || pptxPath === process.cwd()) {
    throw new Error("Missing PPTX path. Pass --pptx <local .pptx path>.");
  }
  if (!fs.existsSync(pptxPath)) {
    throw new Error(`PPTX file not found: ${pptxPath}`);
  }
  if (path.extname(pptxPath).toLowerCase() !== ".pptx") {
    throw new Error(`Expected a .pptx file: ${pptxPath}`);
  }
  if (!resourcesDir || !fs.existsSync(resourcesDir)) {
    throw new Error("Missing resources folder. Pass --resources <local folder>, or place a Resources folder next to the PPTX.");
  }

  const localPdfs = await collectLocalPdfs(resourcesDir);
  if (localPdfs.length === 0) {
    throw new Error(`No PDFs found in resources folder: ${resourcesDir}`);
  }

  const auth = await authorize(argv);
  const drive = google.drive({ version: "v3", auth });
  const slides = google.slides({ version: "v1", auth });

  const deckName = getFirst(argv, ["name", "deck-name"]) || stripExtension(path.basename(pptxPath));
  const pdfFolderName =
    getFirst(argv, ["resource-folder-name", "pdf-folder-name"]) || `${deckName} Resources`;

  console.log(`Uploading and converting PPTX: ${pptxPath}`);
  const uploadedDeck = await uploadPptxAsSlides(drive, pptxPath, deckName, parentFolderId);
  console.log(`Google Slides deck: ${uploadedDeck.name}`);
  console.log(uploadedDeck.webViewLink || `https://docs.google.com/presentation/d/${uploadedDeck.id}/edit`);

  const pdfFolderId = existingPdfFolderId || (await createDriveFolder(drive, pdfFolderName, parentFolderId)).id;
  console.log(`PDF folder: ${pdfFolderId}`);

  const uploadedPdfs = [];
  for (const pdfPath of localPdfs) {
    const uploadedPdf = await uploadPdf(drive, pdfPath, pdfFolderId);
    uploadedPdfs.push(uploadedPdf);
    console.log(`Uploaded PDF: ${uploadedPdf.name}`);
  }

  const pdfMap = buildPdfMapFromFiles(uploadedPdfs);
  await maybeWriteManifestOutput(argv, {
    resources: uploadedPdfs.map((file) => ({
      name: file.name,
      url: file.webViewLink || `https://drive.google.com/file/d/${file.id}/view`,
    })),
  });
  const result = await linkPresentationToPdfMap(slides, uploadedDeck.id, pdfMap, dryRun);

  console.log("");
  console.log("Publish complete.");
  console.log(`Presentation ID: ${uploadedDeck.id}`);
  console.log(`Presentation URL: ${uploadedDeck.webViewLink || `https://docs.google.com/presentation/d/${uploadedDeck.id}/edit`}`);
  console.log(`PDFs uploaded: ${uploadedPdfs.length}`);
  console.log(`Links ${dryRun ? "that would be updated" : "updated"}: ${result.requests.length}`);

  if (result.matchLog.length > 0) {
    console.log("");
    console.log("Matches:");
    for (const line of result.matchLog) {
      console.log(`  ${line}`);
    }
  }

  if (result.requests.length === 0) {
    console.log("");
    console.log("No resource text matched the uploaded PDF filenames.");
  }
}

async function authorize(argv) {
  const serviceAccountPath = getServiceAccountPath(argv);
  if (serviceAccountPath) {
    return authorizeServiceAccount(argv, serviceAccountPath);
  }

  const tokenPath = getTokenPath(argv);
  const credentialsPath = getCredentialsPath(argv);
  const savedToken = fs.existsSync(tokenPath) ? await readJsonIfExists(tokenPath) : null;

  if (savedToken && savedToken.scope && !hasRequiredScopes(savedToken.scope) && !credentialsPath) {
    throw new Error(
      "The saved Google token does not include all required scopes for upload/linking. " +
        "Run auth again with --credentials <oauth-client.json>."
    );
  }

  const savedClient = await loadSavedCredentials(tokenPath, credentialsPath);

  if (savedClient) {
    return savedClient;
  }

  if (!credentialsPath) {
    throw new Error(
      "Missing OAuth client JSON. Create a Google Cloud OAuth client for a Desktop app, " +
        "then save it as google_oauth_client.json or pass --credentials <path>."
    );
  }

  const client = await authenticate({
    scopes: SCOPES,
    keyfilePath: credentialsPath,
  });

  if (!client.credentials || !client.credentials.refresh_token) {
    throw new Error(
      "Google did not return a refresh token. Remove any existing token, revoke this app in your " +
        "Google Account permissions, then run auth again."
    );
  }

  await saveCredentials(client, credentialsPath, tokenPath);
  return client;
}

async function authorizeServiceAccount(argv, serviceAccountPath) {
  const subject = argv.subject || process.env.GOOGLE_WORKSPACE_SUBJECT;
  const key = await readJsonIfExists(serviceAccountPath);

  if (!key.client_email || !key.private_key) {
    throw new Error(`Service account JSON is missing client_email or private_key: ${serviceAccountPath}`);
  }

  if (subject) {
    const client = new google.auth.JWT({
      email: key.client_email,
      key: key.private_key,
      scopes: SCOPES,
      subject,
    });
    await client.authorize();
    console.log(`Using service account ${key.client_email} with delegated subject ${subject}`);
    return client;
  }

  const auth = new google.auth.GoogleAuth({
    keyFile: serviceAccountPath,
    scopes: SCOPES,
  });
  const client = await auth.getClient();
  console.log(`Using service account ${key.client_email}`);
  return client;
}

async function loadSavedCredentials(tokenPath, credentialsPath) {
  if (!fs.existsSync(tokenPath)) {
    return null;
  }

  const token = await readJsonIfExists(tokenPath);
  const credentials = credentialsPath ? await readJsonIfExists(credentialsPath) : null;
  const clientInfo = token.client_id && token.client_secret ? token : getClientInfo(credentials);

  if (token.scope && !hasRequiredScopes(token.scope)) {
    return null;
  }

  if (!token.refresh_token || !clientInfo) {
    return null;
  }

  const client = new google.auth.OAuth2(
    clientInfo.client_id,
    clientInfo.client_secret,
    clientInfo.redirect_uris && clientInfo.redirect_uris[0]
  );
  client.setCredentials({
    access_token: token.access_token,
    expiry_date: token.expiry_date,
    refresh_token: token.refresh_token,
    scope: token.scope,
    token_type: token.token_type,
  });
  return client;
}

async function saveCredentials(client, credentialsPath, tokenPath) {
  const credentials = await readJsonIfExists(credentialsPath);
  const clientInfo = getClientInfo(credentials);

  if (!clientInfo) {
    throw new Error("OAuth client JSON did not contain installed or web client details.");
  }

  await ensureDir(path.dirname(tokenPath));
  await writeJson(tokenPath, {
    type: "authorized_user",
    client_id: clientInfo.client_id,
    client_secret: clientInfo.client_secret,
    refresh_token: client.credentials.refresh_token,
    scope: client.credentials.scope,
    token_type: client.credentials.token_type,
    expiry_date: client.credentials.expiry_date,
  });
  console.log(`Saved OAuth token to ${tokenPath}`);
  console.log(`Using OAuth client from ${credentialsPath}`);
}

async function linkPresentationToPdfMap(slides, presentationId, pdfMap, dryRun) {
  const presentation = await getPresentationWithRetry(slides, presentationId);
  const requests = [];
  const matchLog = [];

  for (const [slideIndex, slide] of (presentation.slides || []).entries()) {
    for (const element of slide.pageElements || []) {
      collectShapeLinkRequests(element, slideIndex, pdfMap, requests, matchLog);
      collectTableLinkRequests(element, slideIndex, pdfMap, requests, matchLog);
    }
  }

  if (!dryRun && requests.length > 0) {
    await runBatchUpdates(slides, presentationId, requests);
  }

  return { requests, matchLog };
}

async function getPresentationWithRetry(slides, presentationId) {
  const maxAttempts = 6;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = await slides.presentations.get({ presentationId });
      return response.data;
    } catch (error) {
      if (attempt === maxAttempts || !isRetryableGoogleError(error)) {
        throw error;
      }
      await sleep(1500 * attempt);
    }
  }

  throw new Error(`Could not open presentation: ${presentationId}`);
}

function collectShapeLinkRequests(element, slideIndex, pdfMap, requests, matchLog) {
  const textElements = element.shape && element.shape.text && element.shape.text.textElements;
  if (!Array.isArray(textElements)) {
    return;
  }

  collectTextRunRequests({
    textElements,
    objectId: element.objectId,
    slideIndex,
    pdfMap,
    requests,
    matchLog,
  });
}

function collectTableLinkRequests(element, slideIndex, pdfMap, requests, matchLog) {
  const rows = element.table && element.table.tableRows;
  if (!Array.isArray(rows)) {
    return;
  }

  for (const [rowIndex, row] of rows.entries()) {
    for (const [columnIndex, cell] of (row.tableCells || []).entries()) {
      const textElements = cell.text && cell.text.textElements;
      if (!Array.isArray(textElements)) {
        continue;
      }

      collectTextRunRequests({
        textElements,
        objectId: element.objectId,
        cellLocation: { rowIndex, columnIndex },
        slideIndex,
        pdfMap,
        requests,
        matchLog,
      });
    }
  }
}

function collectTextRunRequests({ textElements, objectId, cellLocation, slideIndex, pdfMap, requests, matchLog }) {
  for (const textElement of textElements) {
    const textRun = textElement.textRun;
    if (!textRun || typeof textRun.content !== "string") {
      continue;
    }

    const runText = textRun.content.trim();
    if (!runText) {
      continue;
    }

    let matched = findBestMatch(runText, pdfMap);

    if (!matched) {
      const linkUrl = textRun.style && textRun.style.link && textRun.style.link.url;
      if (linkUrl && isLocalPath(linkUrl)) {
        const localFilename = extractFilenameFromPath(linkUrl);
        const localStem = stripExtension(localFilename);
        matched = pdfMap[localStem] || pdfMap[normalizeName(localStem)];
      }
    }

    if (!matched) {
      continue;
    }

    const updateTextStyle = {
      objectId,
      textRange: {
        type: "FIXED_RANGE",
        startIndex: textElement.startIndex,
        endIndex: textElement.endIndex,
      },
      style: {
        link: {
          url: matched.url,
        },
      },
      fields: "link",
    };

    if (cellLocation) {
      updateTextStyle.cellLocation = cellLocation;
    }

    requests.push({ updateTextStyle });
    const location = cellLocation ? "table" : "shape";
    matchLog.push(`Slide ${slideIndex + 1} (${location}): "${truncate(runText, 40)}" -> ${matched.name}`);
  }
}

async function runBatchUpdates(slides, presentationId, requests) {
  const chunkSize = 500;

  for (let start = 0; start < requests.length; start += chunkSize) {
    const chunk = requests.slice(start, start + chunkSize);
    await slides.presentations.batchUpdate({
      presentationId,
      requestBody: { requests: chunk },
    });
  }
}

async function buildPdfMap(drive, rootFolderId, searchSubfolders) {
  const map = {};
  const folderQueue = [rootFolderId];

  while (folderQueue.length > 0) {
    const folderId = folderQueue.shift();

    const pdfs = await listFilesInFolder(drive, folderId, "application/pdf");
    for (const file of pdfs) {
      const stem = stripExtension(file.name);
      const normalizedStem = normalizeName(stem);
      const entry = {
        id: file.id,
        name: file.name,
        stem,
        url: file.webViewLink || `https://drive.google.com/file/d/${file.id}/view`,
      };

      if (!map[normalizedStem]) {
        map[normalizedStem] = entry;
      }
      if (!map[stem] && normalizedStem !== stem) {
        map[stem] = entry;
      }
    }

    if (searchSubfolders) {
      const subfolders = await listFilesInFolder(drive, folderId, "application/vnd.google-apps.folder");
      for (const folder of subfolders) {
        folderQueue.push(folder.id);
      }
    }
  }

  return map;
}

function buildPdfMapFromFiles(files) {
  const map = {};

  for (const file of files) {
    const stem = stripExtension(file.name);
    const normalizedStem = normalizeName(stem);
    const entry = {
      id: file.id,
      name: file.name,
      stem,
      url: file.webViewLink || `https://drive.google.com/file/d/${file.id}/view`,
    };

    if (!map[normalizedStem]) {
      map[normalizedStem] = entry;
    }
    if (!map[stem] && normalizedStem !== stem) {
      map[stem] = entry;
    }
  }

  return map;
}

function buildResourceMapFromManifest(raw) {
  const entries = extractManifestEntries(raw);
  const map = {};

  for (const entry of entries) {
    if (!entry) continue;

    const name = String(entry.name || entry.filename || entry.fileName || "").trim();
    const url = String(entry.url || entry.webViewLink || entry.link || "").trim();
    if (!name || !url) continue;

    const stem = stripExtension(name);
    const normalizedStem = normalizeName(stem);
    const mapped = { name, url, stem };

    map[normalizedStem] = mapped;
    if (normalizedStem !== stem) {
      map[stem] = mapped;
    }
    map[normalizeName(name)] = mapped;
  }

  return map;
}

async function collectPptxPdfLinkNames(pptxPath) {
  const zip = await JSZip.loadAsync(await fsp.readFile(pptxPath));
  const names = new Set();

  for (const relName of Object.keys(zip.files).filter((name) => name.endsWith(".rels"))) {
    const relFile = zip.file(relName);
    if (!relFile) continue;

    const xml = await relFile.async("string");
    if (!xml.includes("relationships/hyperlink")) continue;

    xml.replace(/<Relationship\b[^>]*>/g, (relationshipXml) => {
      if (!/relationships\/hyperlink/.test(relationshipXml)) {
        return relationshipXml;
      }

      const target = getXmlAttribute(relationshipXml, "Target");
      const filename = extractFilenameFromPath(target);
      if (path.extname(filename).toLowerCase() === ".pdf") {
        names.add(filename);
      }
      return relationshipXml;
    });
  }

  return Array.from(names);
}

function extractManifestEntries(raw) {
  if (Array.isArray(raw)) {
    return raw;
  }
  if (raw && Array.isArray(raw.resources)) {
    return raw.resources;
  }
  if (raw && Array.isArray(raw.files)) {
    return raw.files;
  }
  if (raw && typeof raw === "object") {
    return Object.keys(raw).map((key) => ({ name: key, url: raw[key] }));
  }
  throw new Error("Manifest must be a JSON array or object.");
}

async function uploadPptxAsSlides(drive, pptxPath, name, parentFolderId) {
  const requestBody = {
    name,
    mimeType: "application/vnd.google-apps.presentation",
  };

  if (parentFolderId) {
    requestBody.parents = [parentFolderId];
  }

  const response = await drive.files.create({
    requestBody,
    media: {
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      body: fs.createReadStream(pptxPath),
    },
    fields: "id, name, mimeType, webViewLink",
    supportsAllDrives: true,
  });

  return response.data;
}

async function uploadPdf(drive, pdfPath, parentFolderId) {
  const requestBody = {
    name: path.basename(pdfPath),
    mimeType: "application/pdf",
  };

  if (parentFolderId) {
    requestBody.parents = [parentFolderId];
  }

  const response = await drive.files.create({
    requestBody,
    media: {
      mimeType: "application/pdf",
      body: fs.createReadStream(pdfPath),
    },
    fields: "id, name, mimeType, webViewLink",
    supportsAllDrives: true,
  });

  return response.data;
}

async function createDriveFolder(drive, name, parentFolderId) {
  const requestBody = {
    name,
    mimeType: "application/vnd.google-apps.folder",
  };

  if (parentFolderId) {
    requestBody.parents = [parentFolderId];
  }

  const response = await drive.files.create({
    requestBody,
    fields: "id, name, mimeType, webViewLink",
    supportsAllDrives: true,
  });

  return response.data;
}

async function listFilesInFolder(drive, folderId, mimeType) {
  const files = [];
  let pageToken;

  do {
    const response = await drive.files.list({
      q: `'${folderId}' in parents and mimeType='${mimeType}' and trashed=false`,
      fields: "nextPageToken, files(id, name, mimeType, webViewLink)",
      pageSize: 1000,
      pageToken,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
    });

    files.push(...(response.data.files || []));
    pageToken = response.data.nextPageToken;
  } while (pageToken);

  return files;
}

function uniquePdfEntries(pdfMap) {
  const seen = new Set();
  const entries = [];

  for (const entry of Object.values(pdfMap)) {
    if (seen.has(entry.id)) {
      continue;
    }
    seen.add(entry.id);
    entries.push(entry);
  }

  return entries;
}

function findManifestMatchForTarget(target, resourceMap) {
  const decodedTarget = decodeXmlAttribute(target);
  const filename = extractFilenameFromPath(decodedTarget);
  const stem = stripExtension(filename);
  return (
    resourceMap[normalizeName(stem)] ||
    resourceMap[stem] ||
    resourceMap[normalizeName(filename)] ||
    null
  );
}

function printUnmatchedPdfTargets(unmatchedPdfTargets) {
  if (!unmatchedPdfTargets || unmatchedPdfTargets.size === 0) {
    return;
  }

  console.log("");
  console.log("PDF hyperlinks without manifest matches:");
  for (const [target, filename] of unmatchedPdfTargets.entries()) {
    console.log(`  ${filename}: ${target}`);
  }
}

function getXmlAttribute(xml, name) {
  const pattern = new RegExp(`\\b${name}="([^"]*)"`);
  const match = xml.match(pattern);
  return match ? decodeXmlAttribute(match[1]) : "";
}

function setXmlAttribute(xml, name, value) {
  const escapedValue = encodeXmlAttribute(value);
  const pattern = new RegExp(`\\b${name}="[^"]*"`);

  if (pattern.test(xml)) {
    return xml.replace(pattern, `${name}="${escapedValue}"`);
  }

  return xml.replace(/\/?>$/, (ending) => ` ${name}="${escapedValue}"${ending}`);
}

function encodeXmlAttribute(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function decodeXmlAttribute(value) {
  return String(value)
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function getPatchOutputPath(argv, pptxPath) {
  if (argv["in-place"]) {
    return pptxPath;
  }

  if (argv.out || argv.output) {
    return path.resolve(argv.out || argv.output);
  }

  const parsed = path.parse(pptxPath);
  return path.join(parsed.dir, `${parsed.name}_linked${parsed.ext}`);
}

function resolveResourcesDir(argv, pptxPath) {
  const explicit = getFirst(argv, ["resources", "resource-dir", "resources-dir"]);
  if (explicit) {
    return path.resolve(explicit);
  }

  if (!pptxPath || pptxPath === process.cwd()) {
    return null;
  }

  const siblingResources = path.join(path.dirname(pptxPath), "Resources");
  if (fs.existsSync(siblingResources)) {
    return siblingResources;
  }

  return null;
}

async function collectLocalPdfs(dir) {
  const entries = await fsp.readdir(dir, { withFileTypes: true });
  const pdfs = [];

  for (const entry of entries) {
    const entryPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      pdfs.push(...(await collectLocalPdfs(entryPath)));
      continue;
    }
    if (entry.isFile() && path.extname(entry.name).toLowerCase() === ".pdf") {
      pdfs.push(entryPath);
    }
  }

  return pdfs.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));
}

function isRetryableGoogleError(error) {
  const status = error && (error.code || (error.response && error.response.status));
  return status === 404 || status === 409 || status === 429 || (status >= 500 && status < 600);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function findBestMatch(text, pdfMap) {
  if (!text || text.length < 3) return null;

  const normalizedText = normalizeName(text);
  const minPartialMatchLength = 8;

  if (pdfMap[normalizedText]) return pdfMap[normalizedText];

  let bestMatch = null;
  let bestLength = 0;

  for (const key of Object.keys(pdfMap)) {
    const normalizedKey = normalizeName(key);

    if (normalizedKey.length >= minPartialMatchLength && normalizedText.includes(normalizedKey)) {
      if (normalizedKey.length > bestLength) {
        bestMatch = pdfMap[key];
        bestLength = normalizedKey.length;
      }
    }
  }

  return bestMatch;
}

function extractPresentationId(input) {
  if (!input) return null;
  const value = String(input).trim();
  const presentationMatch = value.match(/\/presentation\/d\/([a-zA-Z0-9_-]+)/);
  if (presentationMatch) return presentationMatch[1];
  const idParamMatch = value.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idParamMatch) return idParamMatch[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(value)) return value;
  return null;
}

function extractFolderId(input) {
  if (!input) return null;
  const value = String(input).trim();
  const folderMatch = value.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch) return folderMatch[1];
  const idParamMatch = value.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idParamMatch) return idParamMatch[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(value)) return value;
  return null;
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
  return !/^https?:\/\//.test(url) || url.startsWith("file://");
}

function extractFilenameFromPath(inputPath) {
  if (!inputPath) return "";
  const parts = String(inputPath).replace(/\\/g, "/").split("/");
  return parts[parts.length - 1] || "";
}

function truncate(str, maxLen) {
  if (!str) return "";
  if (str.length <= maxLen) return str;
  return `${str.slice(0, maxLen - 3)}...`;
}

function getFirst(argv, keys) {
  for (const key of keys) {
    if (argv[key]) return argv[key];
  }
  return null;
}

function getTokenPath(argv) {
  return path.resolve(argv.token || process.env.GOOGLE_SLIDES_RESOURCE_TOKEN || DEFAULT_TOKEN_PATH);
}

function getConfigPath(argv) {
  return path.resolve(argv.config || process.env.GOOGLE_SLIDES_RESOURCE_CONFIG || DEFAULT_CONFIG_PATH);
}

function getCredentialsPath(argv) {
  const explicit = argv.credentials || process.env.GOOGLE_OAUTH_CLIENT;
  if (explicit) {
    const resolved = path.resolve(explicit);
    if (!fs.existsSync(resolved)) {
      throw new Error(`OAuth client JSON not found: ${resolved}`);
    }
    return resolved;
  }

  for (const candidate of DEFAULT_CREDENTIAL_PATHS) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return null;
}

function getServiceAccountPath(argv) {
  const explicit = argv["service-account"] || process.env.GOOGLE_SERVICE_ACCOUNT;
  if (explicit) {
    const resolved = path.resolve(explicit);
    if (!fs.existsSync(resolved)) {
      throw new Error(`Service account JSON not found: ${resolved}`);
    }
    return resolved;
  }

  for (const candidate of DEFAULT_SERVICE_ACCOUNT_PATHS) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return null;
}

function findDefaultFile(candidates) {
  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }
  return null;
}

function getClientInfo(credentials) {
  if (!credentials) return null;
  return credentials.installed || credentials.web || null;
}

function hasRequiredScopes(scopeValue) {
  const granted = new Set(String(scopeValue || "").split(/\s+/).filter(Boolean));
  return SCOPES.every((scope) => granted.has(scope));
}

function parseJsonText(text) {
  return JSON.parse(String(text || "").replace(/^\uFEFF/, ""));
}

async function readJsonIfExists(filePath) {
  try {
    return parseJsonText(await fsp.readFile(filePath, "utf8"));
  } catch (error) {
    if (error.code === "ENOENT") {
      return {};
    }
    throw error;
  }
}

async function writeJson(filePath, value) {
  await ensureDir(path.dirname(filePath));
  await fsp.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

async function maybeWriteManifestOutput(argv, manifest) {
  if (!argv["manifest-out"]) {
    return;
  }
  await writeManifestOutput(argv, manifest);
}

async function writeManifestOutput(argv, manifest) {
  const json = `${JSON.stringify(manifest, null, 2)}\n`;
  const outputPath = argv["manifest-out"];

  if (outputPath) {
    const resolved = path.resolve(outputPath);
    await ensureDir(path.dirname(resolved));
    await fsp.writeFile(resolved, json, "utf8");
    console.log(`Resource manifest written to ${resolved}`);
    return;
  }

  console.log(json.trimEnd());
}

async function ensureDir(dir) {
  await fsp.mkdir(dir, { recursive: true });
}

function printHelp() {
  console.log(`
Google Slides resource linker

Commands:
  node scripts/google_slides_resource_cli.js auth --credentials <oauth-client.json>
      Authorize once and save a local refresh token.

  node scripts/google_slides_resource_cli.js doctor
      Show saved config, token, credential, and next-step status.

  node scripts/google_slides_resource_cli.js publish --service-account <key.json> --subject <user@domain> --pptx <local.pptx>
      Run with an admin-approved service account and domain-wide delegation.

  node scripts/google_slides_resource_cli.js publish --pptx <local.pptx> --resources <local-resources-folder>
      Upload and convert a PPTX to Google Slides, upload local PDFs, then link resources.

  node scripts/google_slides_resource_cli.js manifest --pdf-folder <folder-url-or-id> [--manifest-out resources.json]
      Generate JSON for the low-scope Apps Script manifest linker.

  node scripts/google_slides_resource_cli.js manifest-template --pptx <local.pptx> [--resources <folder>] --manifest-out resources.json
      Generate a local fill-in manifest template with PDF names and blank URL values.

  node scripts/google_slides_resource_cli.js patch-pptx --pptx <local.pptx> --manifest <resources.json>
      Rewrite local resource hyperlinks inside the PPTX before uploading to Google Slides.

  node scripts/google_slides_resource_cli.js configure --presentation <slides-url-or-id> --pdf-folder <folder-url-or-id> --drive-folder <folder-url-or-id>
      Save default presentation, PDF folder, and parent upload folder IDs under .google/.

  node scripts/google_slides_resource_cli.js link [--presentation <id>] [--pdf-folder <id>] [--dry-run]
      Link matching resource text in the presentation to PDFs in Drive.

  node scripts/google_slides_resource_cli.js clear-token
      Delete the saved OAuth token.

Options:
  --credentials <path>   OAuth client JSON path. Defaults to google_oauth_client.json,
                         credentials.json, or .google/google_oauth_client.json.
  --service-account <path>
                         Service account key path. Defaults to service-account.json,
                         google_service_account.json, or .google/service-account.json.
  --subject <email>      Google Workspace user to impersonate with domain-wide delegation.
  --token <path>         Token path. Defaults to .google/google_slides_resource_token.json.
  --config <path>        Config path. Defaults to .google/google_slides_resource_config.json.
  --drive-folder <id>    Parent Drive folder for publish uploads.
  --resource-folder-name Folder name to create for uploaded PDFs.
  --manifest-out <path>  Write resource manifest JSON to a file.
  --manifest <path>      Resource manifest JSON for patch-pptx.
  --out <path>           Output PPTX path for patch-pptx. Defaults to *_linked.pptx.
  --in-place             Overwrite the input PPTX when patching links.
  --no-subfolders        Search only the PDF folder root.
  --dry-run              Report matches without editing Slides.

The first auth run opens Google consent once. Later runs reuse the saved refresh token.
`);
}

main().catch((error) => {
  console.error(error.message || error);
  process.exitCode = 1;
});
