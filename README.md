# Google Sheets Sidebar

Google Sheets Sidebar is a free Google Sheets Editor add-on by Kevin's Teacher
Tech. It provides quick selection, grade-formula, and master-roster update tools
for the spreadsheet currently open in Google Sheets.

## Privacy-first authorization

The add-on uses `@OnlyCurrentDoc` and explicitly requests only current-spreadsheet
and container-UI scopes. It does not request Google Drive access. Spreadsheet
content is processed by Google Apps Script and is not sent to a developer-owned
server by this version of the project.

## Local checks

```bash
npm test
npm run check
```

## Google Apps Script development

The source of record can remain in Git while `clasp` synchronizes the runnable
copy to Google Apps Script. See [developer_implementation.md](developer_implementation.md)
for project creation, Cloud configuration, test deployment, OAuth verification,
and Google Workspace Marketplace steps.

## Public information

- Product: <https://www.kevinsteachertech.com/google-sheets-sidebar>
- Privacy: <https://www.kevinsteachertech.com/google-sheets-sidebar/privacy>
- Terms: <https://www.kevinsteachertech.com/google-sheets-sidebar/terms>
- Support: <kpcassidy@gmail.com>

Draft website copy is available in [docs/privacy-policy.md](docs/privacy-policy.md)
and [docs/terms-of-use.md](docs/terms-of-use.md). Review and publish those drafts
before requesting public verification.
