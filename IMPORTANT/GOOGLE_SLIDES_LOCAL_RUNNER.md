# Google Slides Resource Linker Without Repeated Apps Script Permissions

Google requires OAuth consent for any tool that reads Drive files and edits Slides. The part we can avoid is approving a new Apps Script project every time you paste code into a new deck.

If the deck and PDFs are already uploaded to Google Drive, use the standalone Drive web app first. It is one reusable Apps Script project that you authorize once, then run from a web page by pasting the Google Slides deck URL and the Drive PDF folder URL.

Use the local runner when the Apps Script web app is blocked by your school domain or when you want the upload/conversion/linking pipeline to run from this repo. It authorizes once with Google, saves a refresh token under `.google/`, and then runs repeatedly without asking for approval each time.

Check setup status at any time:

```bash
node scripts/google_slides_resource_cli.js doctor
```

## Which path solves which permission prompt?

- Standalone Drive web app: one Google-side Apps Script project, one authorization, then reusable across uploaded decks and Drive resource folders.
- Personal OAuth token: one browser consent for your Google account, then repeat runs reuse `.google/google_slides_resource_token.json`.
- Service account with domain-wide delegation: no user browser consent after the Google Workspace admin approves the service account scopes.
- Container-bound Apps Script pasted into each deck: Google treats each pasted script as a new project, so it can still ask for permission on each new deck.
- Standalone Apps Script or Workspace add-on: one reusable Google-side project, but it still needs that project to be authorized/approved once.

## Recommended for uploaded Drive files: standalone web app

Use this when the PowerPoint has already been uploaded/converted to Google Slides and the PDFs are already in a Drive folder.

1. Go to `https://script.google.com` and create a new Apps Script project.
2. Paste `IMPORTANT/GOOGLE_DRIVE_RESOURCE_LINKER_STANDALONE.js` into `Code.gs`.
3. Open project settings and enable `Show appsscript.json manifest file in editor`.
4. Replace `appsscript.json` with `IMPORTANT/GOOGLE_DRIVE_RESOURCE_LINKER_appsscript.json`.
5. Click `Deploy > New deployment > Web app`.
6. For personal use, keep the manifest/access setting as `MYSELF`, or choose your school domain if you intentionally want domain users to run it.
7. Keep `Execute as` set to the deploying user if the tool should use your Drive access.
8. Authorize the web app once.
9. Open the web app URL and paste:
   - the Google Slides deck URL or ID
   - the Drive folder URL or ID containing the resource PDFs

The web app lists PDFs from that Drive folder, opens the Google Slides presentation by ID, and replaces matching local/resource hyperlinks with Drive PDF links.

Do not deploy an `Execute as me` Drive tool publicly. If it runs as you, anyone with access to the web app can make it act using your Drive permissions.

## One-time setup

### Option A: personal OAuth token

1. In Google Cloud, create an OAuth client for a Desktop app.
2. Enable the Google Drive API and Google Slides API for that Cloud project.
3. If the OAuth consent screen is in testing mode, add your Google account as a test user.
4. Download the OAuth client JSON.
5. Save it in this project as one of:
   - `google_oauth_client.json`
   - `credentials.json`
   - `.google/google_oauth_client.json`
6. Run:

```bash
node scripts/google_slides_resource_cli.js auth --credentials google_oauth_client.json
```

Google will ask for permission once. Complete the browser flow. The script saves a token at:

```text
.google/google_slides_resource_token.json
```

That token path is gitignored.

If your school domain blocks unverified OAuth apps, the local runner still cannot bypass that policy. In that case, the durable fix is for the domain admin to trust the OAuth app or for the tool to be published as an approved internal Google Workspace app.

### Option B: service account with domain-wide delegation

For a managed Google Workspace school domain, an admin can approve this once so the runner does not open a browser consent prompt.

1. Create a Google Cloud service account.
2. Enable domain-wide delegation for that service account.
3. In the Google Admin console, grant the service account client ID these scopes:

```text
https://www.googleapis.com/auth/presentations
https://www.googleapis.com/auth/drive.readonly
https://www.googleapis.com/auth/drive.file
```

4. Download the service-account key JSON and save it as one of:
   - `service-account.json`
   - `google_service_account.json`
   - `.google/service-account.json`
5. Run commands with `--subject` set to the Workspace user whose Drive should own/access the files:

```bash
node scripts/google_slides_resource_cli.js publish --service-account service-account.json --subject james.hooke@education.vic.gov.au --pptx "output/UNIT_FOLDER/UNIT.pptx"
```

Without `--subject`, the service account can only access Drive folders/files explicitly shared with the service account. That is usually less convenient for school Drive content.

## Recommended: publish a local deck without Apps Script

For a normal generated unit folder, use the combined PPTX and its local `Resources/` folder:

```bash
node scripts/google_slides_resource_cli.js publish --pptx "output/UNIT_FOLDER/UNIT.pptx" --resources "output/UNIT_FOLDER/Resources"
```

The command will:

1. Upload and convert the PPTX to Google Slides.
2. Create a Drive folder for the PDFs.
3. Upload every local PDF from the resources folder.
4. Link matching resource text in the Google Slides deck to the uploaded PDFs.

To upload everything into a particular Drive folder:

```bash
node scripts/google_slides_resource_cli.js publish --pptx "output/UNIT_FOLDER/UNIT.pptx" --resources "output/UNIT_FOLDER/Resources" --drive-folder "https://drive.google.com/drive/folders/FOLDER_ID"
```

If the PPTX already has a sibling folder named `Resources`, the `--resources` flag can be omitted:

```bash
node scripts/google_slides_resource_cli.js publish --pptx "output/UNIT_FOLDER/UNIT.pptx"
```

## Zero Google automation: patch the PPTX locally

If you want no script or API to ask Google for permissions, patch the PPTX before uploading it.

First generate a fill-in manifest template from the local PPTX:

```bash
node scripts/google_slides_resource_cli.js manifest-template --pptx "output/UNIT_FOLDER/UNIT.pptx" --manifest-out resources-manifest.json
```

This creates JSON with the PDF filenames already filled in and blank URL values:

```json
{
  "Session 1 Worksheet.pdf": "",
  "Session 1 Answer Key.pdf": ""
}
```

Upload the PDFs manually through the browser, copy their Drive links into the JSON, then run:

```bash
node scripts/google_slides_resource_cli.js patch-pptx --pptx "output/UNIT_FOLDER/UNIT.pptx" --manifest resources-manifest.json
```

The script writes a new file beside the original:

```text
output/UNIT_FOLDER/UNIT_linked.pptx
```

Upload that patched PPTX to Google Slides through the browser. Because the links are already web URLs inside the PPTX, no Apps Script or Google API permission is needed for link fixing.

Preview what would change without writing a file:

```bash
node scripts/google_slides_resource_cli.js patch-pptx --pptx "output/UNIT_FOLDER/UNIT.pptx" --manifest resources-manifest.json --dry-run
```

## Link already-uploaded resources

Run with explicit IDs or URLs:

```bash
node scripts/google_slides_resource_cli.js link --presentation "https://docs.google.com/presentation/d/PRESENTATION_ID/edit" --pdf-folder "https://drive.google.com/drive/folders/FOLDER_ID"
```

Or save defaults once:

```bash
node scripts/google_slides_resource_cli.js configure --presentation "https://docs.google.com/presentation/d/PRESENTATION_ID/edit" --pdf-folder "https://drive.google.com/drive/folders/FOLDER_ID"
node scripts/google_slides_resource_cli.js link
```

Preview matches without editing:

```bash
node scripts/google_slides_resource_cli.js link --dry-run
```

Search only the PDF folder root:

```bash
node scripts/google_slides_resource_cli.js link --no-subfolders
```

## Low-scope Apps Script fallback

If you still want to run inside Google Slides, use the manifest command instead of the Drive-folder commands. This avoids `DriveApp` in Apps Script, so the script does not need permission to see Drive files.

1. Generate a manifest from a Drive PDF folder:

```bash
node scripts/google_slides_resource_cli.js manifest --pdf-folder "https://drive.google.com/drive/folders/FOLDER_ID" --manifest-out resources-manifest.json
```

2. In Apps Script, paste `IMPORTANT/GOOGLE_SLIDES_RESOURCE_MANIFEST_ONLY.js` into `Code.gs`.
3. Open the Apps Script project settings and enable `Show appsscript.json manifest file in editor`.
4. Replace `appsscript.json` with `IMPORTANT/GOOGLE_SLIDES_RESOURCE_LOW_SCOPE_appsscript.json`.
5. Reload the Google Slides deck.
6. Run `Resource Tools > Link Resources from Manifest` and load or paste the contents of `resources-manifest.json`.

This mode should only ask for access to the current presentation and the script UI. Use the manifest-only Apps Script file for this path; the all-in-one `GOOGLE_SLIDES_RESOURCE.js` still includes Drive-folder and merge commands for the older workflow.

If you see a manifest parse error, check that you loaded the JSON file contents. Do not paste the CLI status line that says `Resource manifest written to ...`.

## Why the Apps Script still asks

`IMPORTANT/GOOGLE_SLIDES_RESOURCE.js` is an Apps Script file. If you paste it into a different Google Slides file, Google creates a different Apps Script project. A different project means a different OAuth consent grant.

The saved-folder menu in that script removes repeated folder-ID prompts for the same Apps Script project. It cannot remove Google's OAuth screen for a newly pasted script project.

For no repeated permission screens across decks, use the local runner above, or deploy the Apps Script as a real Google Workspace add-on installed once for your account/domain.
