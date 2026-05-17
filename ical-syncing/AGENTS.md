# iCal Syncing Agent Notes

## Purpose

This directory contains the Google Apps Script project that syncs remote iCal feeds into Google Calendar.

## Key Files

- `icalFeedSync.gs`: main script logic.
- `icalFeedSync.config.gs`: local real config file used for deploys. This file is gitignored but is uploaded by `clasp push` when present.
- `icalFeedSync.config.example.gs`: checked-in config template.
- `appsscript.json`: Apps Script manifest and scopes.
- `.clasp.json`: local Apps Script project binding for this directory.
- `.claspignore`: excludes tests, docs, and example files from upload.

## Sync To Google

Use the repo-level npm script:

```bash
npm run clasp:push:ical
```

That script runs:

```bash
cd ical-syncing && clasp push
```

Run it from the repo root.

## Preconditions

- `ical-syncing/.clasp.json` must point at the correct Apps Script project ID.
- `ical-syncing/icalFeedSync.config.gs` must exist locally with real settings.
- `clasp` must already be authenticated. If not, run:

```bash
npm run clasp:login
```

- The Apps Script API must be enabled for the Google account doing the push.

## What Gets Uploaded

- `icalFeedSync.gs`
- `icalFeedSync.config.gs` if present locally
- `appsscript.json`

These do not get uploaded:

- `README.md`
- `AGENTS.md`
- `icalFeedSync.test.js`
- `icalFeedSync.config.example.gs`
- `.clasp.json`

## Recommended Workflow

1. Make code changes locally.
2. Run the targeted tests when changing sync behavior:

```bash
npm test -- ical-syncing/icalFeedSync.test.js
```

3. Push the project to Google:

```bash
npm run clasp:push:ical
```

4. In Apps Script, manually run `syncIcalFeeds()` when needed to validate behavior and authorization.

## Notes

- A successful Git push to GitHub does not deploy the Apps Script project.
- Changes to `README.md`, tests, or `AGENTS.md` are local/repo-only and will not affect the deployed Apps Script.
- If `clasp push` fails with a network or auth error, resolve that first rather than assuming the script content is wrong.
