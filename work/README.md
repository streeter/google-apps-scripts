# Work Calendar Scripts

This directory contains scripts focused on day-to-day work calendar management.

## script.google.com setup

1. Create a project at `https://script.google.com/`.
2. Add these files to the project:
   - `blockFromPersonalCalendar.gs`
   - `clearPastBlocks.js`
   - `colorBasedOnAttendees.gs`
   - `scheduleInterviewFeedback.gs`
3. Create config files from examples:
   - `colorAttendeeConfig.gs` (from `colorAttendeeConfig.example.gs`)
   - `getPersonalCalendar.gs` (from `getPersonalCalendar.example.gs`)
4. Keep config files loaded before scripts that use them (for example, first alphabetically).
5. If prompted by a script, add Advanced Services in **Services** (for example Calendar API).
6. Run each script's main function once to authorize.
7. Add time-based triggers for the functions you want to run automatically.

## Deploy with clasp

This repo includes [`@google/clasp`](https://github.com/google/clasp) as a dev dependency and a placeholder `.clasp.json` for this script directory.

1. Enable the Apps Script API at `https://script.google.com/home/usersettings`.
2. Authenticate once:

   ```bash
   npm run clasp:login
   ```

3. Edit `work/.clasp.json` and replace `REPLACE_WITH_APPS_SCRIPT_PROJECT_SCRIPT_ID` with the Apps Script project script ID.
   - In the Apps Script editor, find it under **Project Settings** → **Script ID**.
   - Keep `rootDir` set to `"."`.
4. Ensure the local config files exist with your real values:
   - `work/colorAttendeeConfig.gs` (from `colorAttendeeConfig.example.gs`)
   - `work/getPersonalCalendar.gs` (from `getPersonalCalendar.example.gs`)
5. Push the work files:

   ```bash
   npm run clasp:push:work
   ```

6. Pull remote changes back down when editing in the Apps Script UI:

   ```bash
   npm run clasp:pull:work
   ```

7. Open the Apps Script project when needed:

   ```bash
   npm run clasp:open:work
   ```

`work/.claspignore` excludes local docs, tests, the placeholder `.clasp.json`, and the example configs from uploads. `appsscript.json` must be uploaded because Apps Script requires a manifest. The real `colorAttendeeConfig.gs` and `getPersonalCalendar.gs` remain gitignored, but they will be uploaded by `clasp push` when present locally.

## Scripts

- `colorBasedOnAttendees.gs`: colors each event based on attendees and matching rules.
- `blockFromPersonalCalendar.gs`: blocks time on a target calendar based on events from one or more source calendars.
- `clearPastBlocks.js`: removes old block events created by the blocking script.
- `scheduleInterviewFeedback.gs`: finds interview events and schedules follow-up feedback blocks.

## Config files

- `colorAttendeeConfig.example.gs`: example config for attendee-based coloring rules.
- `getPersonalCalendar.example.gs`: example config for personal/source calendar IDs.

Create local config files from the examples:

- `colorAttendeeConfig.gs`
- `getPersonalCalendar.gs`

Tip: in Apps Script, make sure config files are loaded before scripts that use them (for example, place configs first alphabetically).
