# Work Calendar Scripts

This directory contains scripts focused on day-to-day work calendar management.

## script.google.com setup

1. Create a project at `https://script.google.com/`.
2. Add these files to the project:
   - `blockFromPersonalCalendar.gs`
   - `clearPastBlocks.js`
   - `colorBasedOnAttendees.gs`
   - `scheduleInterviewFeedback.gs`
   - `autoInterviewRooms.gs`
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
   npm run clasp:work:push
   ```

6. Pull remote changes back down when editing in the Apps Script UI:

   ```bash
   npm run clasp:work:pull
   ```

7. Open the Apps Script project when needed:

   ```bash
   npm run clasp:work:open
   ```

`work/.claspignore` excludes local docs, tests, the placeholder `.clasp.json`, and the example configs from uploads. `appsscript.json` must be uploaded because Apps Script requires a manifest. The real `colorAttendeeConfig.gs` and `getPersonalCalendar.gs` remain gitignored, but they will be uploaded by `clasp push` when present locally.

## Scripts

- `colorBasedOnAttendees.gs`: colors each event based on attendees and matching rules.
- `blockFromPersonalCalendar.gs`: blocks time on a target calendar based on events from one or more source calendars.
- `clearPastBlocks.js`: removes old block events created by the blocking script.
- `scheduleInterviewFeedback.gs`: finds interview events and schedules follow-up feedback blocks.
- `autoInterviewRooms.gs`: books meeting rooms for upcoming interviews. See [Auto interview room booker](#auto-interview-room-booker).

## Auto interview room booker

`autoInterviewRooms.gs` automatically books meeting rooms for upcoming interviews on your calendar.

### What it does

1. Scans your calendar for upcoming interviews (identified by "Thanks for interviewing" in the description).
2. Finds an available room for each interview (smallest room first).
3. Books the room with a 15-minute buffer before and after.
4. Cleans up orphaned room bookings whose matching interview no longer exists.

### Setup

1. Set `OFFICE_ID` to your office (defaults to `SSF`). Find your office ID at [Calendar Resource Settings](https://calendar.google.com/calendar/u/0/r/settings/browseresources).
2. Enable the required Advanced Services. `appsscript.json` already declares both, so `clasp push` does this for you. In the Apps Script UI, press **+** next to **Services**:
   - **Google Calendar API** (`Calendar`).
   - **Admin SDK API** (`AdminDirectory`), used to discover rooms in your office.
3. You don't need to deploy. Run the `main()` function once to authorize it. Its log lists the rooms it discovered.

### Usage

- `main()`: run manually or on a schedule. Discovers rooms in `OFFICE_ID`, books rooms for unmatched interviews, and deletes orphans.
- `__deleteAllBookings()`: removes all auto-booked room events.

To run automatically, add a time-driven trigger for `main` that runs every 15–30 minutes.

`main()` tries rooms smallest first, then alphabetically. It skips virtual rooms and rooms larger than `MAX_ROOM_CAPACITY`. The room list is cached per user for 6 hours, so new or renamed rooms show up within that window.

### Configuration

| Constant              | Default                      | Description                                  |
| --------------------- | ---------------------------- | -------------------------------------------- |
| `OFFICE_ID`           | `SSF`                        | Office to discover and book rooms in         |
| `MAX_ROOM_CAPACITY`   | `5`                          | Largest room size the script books           |
| `ROOMS_CACHE_SECONDS` | `21600` (6 hours)            | How long the discovered room list is cached  |
| `BUFFER_MINUTES`      | `15`                         | Buffer added before and after each interview |
| `LOOK_AHEAD_DAYS`     | `14`                         | How far ahead to scan for interviews         |
| `ROOM_EVENT_TITLE`    | `Auto-Booked Interview Room` | Title used for created room events           |

## Config files

- `colorAttendeeConfig.example.gs`: example config for attendee-based coloring rules.
- `getPersonalCalendar.example.gs`: example config for personal/source calendar IDs.

Create local config files from the examples:

- `colorAttendeeConfig.gs`
- `getPersonalCalendar.gs`

Tip: in Apps Script, make sure config files are loaded before scripts that use them (for example, place configs first alphabetically).
