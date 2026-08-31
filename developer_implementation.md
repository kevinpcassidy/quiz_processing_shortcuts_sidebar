# Public add-on implementation guide

This guide describes how to connect the Git repository to Google Apps Script,
associate the script with a standard Google Cloud project, test it as a Google
Sheets Editor add-on, and submit it to Google Workspace Marketplace.

> **Product identity**
>
> - App: **Google Sheets Sidebar**
> - Publisher: **Kevin Cassidy / Kevin's Teacher Tech**
> - Support: **kpcassidy@gmail.com**
> - Product: <https://www.kevinsteachertech.com/google-sheets-sidebar>
> - Privacy: <https://www.kevinsteachertech.com/google-sheets-sidebar/privacy>
> - Terms: <https://www.kevinsteachertech.com/google-sheets-sidebar/terms>
> - Price: **Free**

Google changes console labels and review requirements periodically. Confirm each
submission field against the current official documentation before submitting.
Do not broaden an OAuth scope merely to work around a review or test error.

## 1. Prerequisites and decisions

You need:

- The Google account `kpcassidy@gmail.com` (or a publisher account you control);
- A Google Cloud billing/profile setup if Google requires it for project creation;
- Administrative access to `kevinsteachertech.com` and its DNS or Search Console;
- Node.js and npm locally;
- The Google Cloud CLI (`gcloud`); and
- A small square logo before Marketplace submission. A logo is normally required
  for consent branding and Marketplace assets even though the code can be tested
  without one. Do not use a Google product logo or imply Google endorsement.

Before submission, publish the product, privacy, and terms pages over HTTPS.
Replace the effective-date and jurisdiction placeholders in the policy drafts.

## 2. Run the repository checks

From the repository root:

```bash
npm test
npm run check
```

The checker rejects Google Drive scopes. The app operates only on the spreadsheet
where it is open, using `@OnlyCurrentDoc` and the explicit scopes in
`appsscript.json`.

## 3. Install and authorize `clasp`

Install Google's Apps Script CLI globally or invoke it through `npx`:

```bash
npm install --global @google/clasp
clasp --version
clasp login
```

If `clasp` reports that the Apps Script API is disabled, open
<https://script.google.com/home/usersettings> while signed into the publisher
account, enable the **Google Apps Script API**, wait briefly, and retry.

`clasp login` stores account credentials outside this repository. Never commit
those credentials.

## 4. Create the standalone Apps Script project

For public distribution, use a standalone Apps Script project rather than a
script bound to one private gradebook.

Recommended, explicit method:

1. Open <https://script.google.com/home> with the publisher account.
2. Select **New project**.
3. Rename it **Google Sheets Sidebar**.
4. Open **Project Settings** and copy the **Script ID**.
5. Copy `.clasp.json.example` to `.clasp.json`.
6. Replace the placeholder with the Script ID.
7. Keep `.clasp.json` uncommitted; it is intentionally listed in `.gitignore`.
8. Push the repository files:

```bash
cp .clasp.json.example .clasp.json
# Edit .clasp.json and insert the Script ID.
clasp status
clasp push
clasp open
```

`clasp push` should upload only:

- `shortcuts_sidebar.gs`
- `Sidebar.html`
- `appsscript.json`

The `.claspignore` file excludes policies, tests, and local tooling from the Apps
Script runtime.

If `clasp push` asks to overwrite remote files, inspect the project first. The
initial placeholder `Code.gs` may be replaced, but do not overwrite an existing
project that contains work you need.

## 5. Create the standard Google Cloud project with `gcloud`

Choose a globally unique, permanent project ID. The example below must be
changed if it is unavailable:

```bash
gcloud auth login
gcloud projects create kevins-google-sheets-sidebar \
  --name="Google Sheets Sidebar"
gcloud config set project kevins-google-sheets-sidebar
gcloud projects describe kevins-google-sheets-sidebar \
  --format='value(projectNumber)'
```

Record both the project ID and numeric project number. If the project belongs in
a Google Cloud organization or requires a billing account, supply the appropriate
organization, folder, and billing configuration for your account.

Enable the Google Workspace Marketplace SDK. The service name used by Google has
historically been `appsmarket-component.googleapis.com`; confirm it in the API
Library if Google has renamed it:

```bash
gcloud services enable appsmarket-component.googleapis.com
```

`SpreadsheetApp` is an Apps Script built-in service, so this code does not call
the Google Sheets REST API or Drive REST API and should not need those API scopes.

## 6. Link Apps Script to the Cloud project

This association is performed in Apps Script, not by editing `.clasp.json`:

1. Open the standalone Apps Script project.
2. Open **Project Settings**.
3. Find **Google Cloud Platform (GCP) Project**.
4. Select **Change project**.
5. Enter the numeric project number from the previous step.
6. Confirm the change.
7. Reopen Project Settings and verify the standard Cloud project is shown.

The Apps Script project, OAuth consent configuration, and Marketplace SDK must
refer to this same Cloud project.

## 7. Configure Google Auth Platform / OAuth consent

In Google Cloud Console, select the standard project and open **Google Auth
Platform** (older interfaces call this the **OAuth consent screen**).

Configure:

### Branding

- App name: **Google Sheets Sidebar**
- User support email: **kpcassidy@gmail.com**
- Publisher/developer: **Kevin Cassidy / Kevin's Teacher Tech**
- Home page: <https://www.kevinsteachertech.com/google-sheets-sidebar>
- Privacy policy: <https://www.kevinsteachertech.com/google-sheets-sidebar/privacy>
- Terms: <https://www.kevinsteachertech.com/google-sheets-sidebar/terms>
- Authorized domain: **kevinsteachertech.com**
- Developer contact: **kpcassidy@gmail.com**
- Logo: upload the final non-Google-branded square product logo

Google may require proof of domain ownership through Google Search Console. Use
the same Google account or grant the publisher account verified-owner access.

### Audience

Choose **External** because the add-on is intended for public use. While testing,
add the accounts that will install test deployments as test users. Moving the app
to production does not itself constitute OAuth verification.

### Data access / scopes

The repository declares:

```text
https://www.googleapis.com/auth/spreadsheets.currentonly
https://www.googleapis.com/auth/script.container.ui
```

Use the Apps Script editor's project overview to verify the detected scopes match
the manifest. Do not add `drive.file`, full Drive, or full spreadsheets access
unless future functionality genuinely opens files outside the active spreadsheet.

Suggested scope justification:

> Google Sheets Sidebar uses current-spreadsheet access to read column headers,
> roster rows, grades, formulas, selected ranges, and sheet names in the Google
> spreadsheet where the user opens the add-on. It writes only user-requested
> formulas, selections, and values back to that spreadsheet. Container UI access
> is used to display the add-on menu and sidebar. The app does not request access
> to unrelated Drive files or send spreadsheet contents to a developer server.

## 8. Create and install a test deployment

After every source update:

```bash
npm test
npm run check
clasp push
```

Then in Apps Script:

1. Select **Deploy → Test deployments**.
2. Choose the deployment type for an **Editor add-on**.
3. Select Google Sheets as the host application if prompted.
4. Create/install the test deployment for an authorized test account.
5. Open a test spreadsheet with at least two tabs.
6. Open the add-on from **Extensions** and authorize it.

Do not test initially with real student data. Use fabricated names and grades.

### Manual acceptance checklist

- [ ] The consent screen shows the correct name, publisher, domain, and links.
- [ ] The requested permissions refer only to the current spreadsheet/UI.
- [ ] The add-on menu appears after installation and reopening Sheets.
- [ ] The sidebar loads and excludes the active tab from source choices.
- [ ] Exactly two selections extend without exceeding the sheet row limit.
- [ ] Formula fill handles an empty roster and a populated roster.
- [ ] Header matching ignores capitalization and surrounding whitespace.
- [ ] All matching headers, including the name header, can update.
- [ ] Repeated names pair first-to-first and second-to-second.
- [ ] Repeated names are visibly reported.
- [ ] Unequal duplicate counts leave unpaired destination rows unchanged.
- [ ] Blank source values do not erase existing target values/formulas.
- [ ] Duplicate headers are reported.
- [ ] A source tab cannot update itself.
- [ ] Privacy, terms, and support links open correctly.

## 9. Prepare OAuth verification evidence

For a public external app, submit the consent configuration for verification when
Google requests it. Prepare:

- Verified domain ownership;
- Public product, privacy, and terms pages;
- An accurate explanation for each scope;
- A test account or reviewer instructions if requested;
- A screen recording showing the complete OAuth grant and each permission-backed
  feature; and
- A statement that spreadsheet content is not transferred to an external server.

Suggested demonstration sequence:

1. Start signed out or with the add-on uninstalled.
2. Show the public product and policy pages.
3. Install the test deployment or reviewer build.
4. Show the complete Google consent screen without skipping scope details.
5. Open a fabricated gradebook.
6. Open the sidebar.
7. Demonstrate selection extension.
8. Demonstrate formula fill.
9. Demonstrate a master-roster update with duplicate names.
10. Show the duplicate warning and resulting row-order pairing.
11. State that no developer server or database receives the spreadsheet content.

Keep the video unlisted rather than public if it contains reviewer-only details.

## 10. Configure Google Workspace Marketplace SDK

In the same standard Cloud project:

1. Open **APIs & Services → Library** and confirm **Google Workspace Marketplace
   SDK** is enabled.
2. Open its **Configuration** page.
3. Choose the integration type for a Google Workspace/Apps Script add-on.
4. Enter the standalone Apps Script **Script ID** and the required deployment or
   version information.
5. Enable Google Sheets as the supported host.
6. Choose **Public** visibility/distribution.
7. Enter the app URLs and support email listed at the top of this guide.
8. Confirm the OAuth scopes exactly match the Apps Script manifest.
9. Save the configuration before building the Store Listing.

Do not list Gmail, Calendar, Docs, Drive, or other integrations the code does not
provide.

## 11. Create the Marketplace store listing

Suggested listing copy:

### Short description

> Free Google Sheets tools for roster updates, quick selections, and two-highest
> grade averages.

### Full description

> Google Sheets Sidebar adds practical roster and gradebook tools directly to
> Google Sheets. Extend two selected columns, fill formulas that average the two
> highest of four scores, and update matching columns from a master roster tab.
> Headers and names are matched without regard to capitalization. Repeated names
> are paired by roster order and clearly reported so you can verify the result.
>
> The add-on operates only in the spreadsheet where you open it. It does not
> request broad Google Drive access or send spreadsheet content to a server
> operated by Kevin's Teacher Tech. Google Sheets Sidebar is free.

Suggested categories include **Education** and **Productivity**, subject to the
categories currently offered by Marketplace.

### Required creative assets

Google generally requires application icons and screenshots in specified sizes.
Those requirements can change, so use the dimensions shown in the Marketplace
form. Capture screenshots using fictional roster data. A logo is a remaining
publication dependency; the listing should not be submitted with a generic
Google Sheets logo.

### Reviewer notes

> Install the add-on in Google Sheets and open a spreadsheet containing at least
> two tabs. The active tab is the destination; choose another tab as the master
> roster. Column A is used to match names case-insensitively. Matching column
> headers are updated, including the name column. Duplicate names are paired by
> occurrence order and reported in the sidebar. No spreadsheet data is sent to a
> developer-operated server.

## 12. Submit and release

1. Complete all required Marketplace listing fields and assets.
2. Resolve every validation warning rather than broadening permissions.
3. Submit OAuth verification and Marketplace review in the order Google directs.
4. Monitor `kpcassidy@gmail.com` for reviewer questions.
5. Answer with precise behavior and an updated video when requested.
6. Do not announce general availability until installation works with a Google
   account that is not an OAuth test user.

Google review time is outside the code release process and is not guaranteed.

## 13. Ongoing release workflow

For each release:

```bash
git switch -c feature/short-description
# Make changes.
npm test
npm run check
clasp push
# Complete the manual test checklist in a fabricated spreadsheet.
git add --all
git commit -m "Describe the release"
```

Create a new immutable Apps Script version/deployment as required by the current
add-on workflow. Update the Marketplace configuration to the approved version;
do not silently change scopes. If data handling changes, update and publish the
Privacy Policy before releasing the code change.

## 14. Troubleshooting

### “Unknown developer” or “Google hasn't verified this app”

Confirm that Apps Script is linked to the standard Cloud project, the external
consent configuration is published, the authorized domain is verified, branding
is complete, and verification has been approved for the exact scopes in the
manifest.

### Add-on menu does not appear

Confirm that the test deployment is installed, open a new spreadsheet tab or
reload Sheets, and verify that `onInstall` calls `onOpen`. Check Apps Script
executions for errors.

### Scope mismatch

Run `clasp push`, inspect `appsscript.json` in Apps Script, and compare Apps
Script's detected scopes, Google Auth Platform scopes, and Marketplace SDK scopes.
All three must agree.

### `clasp` pushes documentation or tests

Run `clasp status` and verify `.claspignore` is present at the repository root.
Only the two runtime source files and manifest should be included.
