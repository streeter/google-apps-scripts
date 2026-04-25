# iCal Feed Syncing (Google Apps Script)

This directory contains a Google Apps Script that syncs one or more remote iCal (`.ics`) feeds to specific Google Calendars.

## Files

- `icalFeedSync.gs`: main sync script.
- `icalFeedSync.config.example.gs`: example config file. Copy this to `icalFeedSync.config.gs` and fill in your values.

## What the script does

- Pulls each configured iCal feed URL.
- Syncs only events that are on or after the current date into the configured local Google Calendar.
- Forces synced local events to match upstream feed fields on each run (including start/end date and time).
- Overwrites local edits to synced events on subsequent sync runs.
- Optionally creates managed drive-time placeholder events from a configured origin address to event location.
- For drive placeholders, the script first tries to route from the previous non-placeholder event location on the destination calendar; if already there or within threshold, it skips creating drive time.
- When an event description includes `Arrival: N minutes in advance`, creates a managed advanced-arrival placeholder and anchors drive-time before that arrival block.
- Deletes local synced events when an upstream event is canceled.
- Optionally deletes only on/after-today local synced events that are no longer present in the feed (`deleteMissingFromFeed`).
- Adds configured attendee emails to synced events.
- Adds configured attendee emails to synced events and managed placeholder events (drive/arrival).

## script.google.com setup

1. Create a project at `https://script.google.com/`.
2. Add `icalFeedSync.gs` to the project.
3. Add a new file `icalFeedSync.config.gs`, using `icalFeedSync.config.example.gs` as the template.
4. Run `listMyCalendarIds()` once (or run `syncIcalFeeds()` once) to print accessible calendar names and IDs in logs.
5. Edit `icalFeedSync.config.gs`:
   - Set `calendarId` for each feed mapping.
   - Optionally set per-feed `titlePrefix` (for example `[Sports]`).
   - Set `defaultAttendeeEmails` and/or per-feed `attendeeEmails`.
   - Optionally add `placeNameAddressMap` entries when event titles or locations contain venue names instead of full addresses.
   - If using drive placeholders, set `defaultOriginAddress` and set `addDriveTimePlaceholders: true` where needed.
6. In Apps Script editor:
   - Open **Services**.
   - Add **Google Calendar API** (Advanced Google Service).
   - If using drive placeholders (`addDriveTimePlaceholders: true`), no extra Advanced Service is needed for Maps. `Maps` is a built-in Apps Script service (it does not appear in the Add a service dialog).
7. Run `syncIcalFeeds()` once manually to authorize and validate behavior.
8. Run `setupIcalFeedSyncTrigger()` once to create the periodic trigger.

## Config shape

The main script expects:

```javascript
function getIcalSyncConfig() {
  return {
    triggerEveryMinutes: 15,
    deleteMissingFromFeed: true,
    addDriveTimePlaceholders: false,
    defaultOriginAddress: "123 Main St, Brooklyn, NY 11201",
    placeNameAddressMap: {
      "McMoran Park": "1234 McMoran Park Rd, Your City, ST 12345",
    },
    minDriveMinutesToCreate: 10,
    driveEventTitleTemplate: "Drive ({{minutes}}m) to {{title}}",
    defaultAttendeeEmails: ["person@company.com"],
    feedMappings: [
      {
        name: "Cole Streeter",
        feedUrl: "https://example.com/calendar.ics",
        calendarId: "your_calendar_id@group.calendar.google.com",
        titlePrefix: "[Sports]",
        attendeeEmails: [],
        addDriveTimePlaceholders: true,
        originAddress: "",
        placeNameAddressMap: {
          "McMoran Park": "1234 McMoran Park Rd, Your City, ST 12345",
        },
      },
    ],
  };
}
```

`triggerEveryMinutes` supports `1`, `5`, `10`, `15`, `30`, and multiples of `60` (for example `60` hourly, `120` every 2 hours, `1440` daily).

## Notes

- `feedMappings` can include many feed -> calendar routes.
- Per-feed `titlePrefix` prepends synced event titles for that feed.
- `placeNameAddressMap` lets you translate venue names in titles or locations into routable addresses before drive lookup.
- Per-feed `attendeeEmails` overrides `defaultAttendeeEmails` when non-empty.
- Per-feed `addDriveTimePlaceholders` controls whether drive placeholders are managed for that feed.
- Placeholders are only created when computed drive time is strictly greater than `minDriveMinutesToCreate` (default `10`).
- Drive placeholders are skipped for all-day events, events without a location, unroutable addresses, and pre-today events.
- Drive placeholders are tied to source synced events using metadata (`sourceSyncKey` and `sourceEventId`) and are managed/deleted safely.
- Arrival placeholders are tied to source synced events using metadata and are managed/deleted safely.
- The script uses event metadata (`extendedProperties.private`) to track synced items and detect changes.
- Delete operations are guarded to only remove events that are verifiably managed by this script/feed.
- On first `syncIcalFeeds()` run, the script logs all accessible calendar names/IDs once, to help with initial config.

## Local tests

- Run: `npm test`
- Lint formatting with: `npm run lint`
- Auto-format with: `npm run format`
- Test file: `ical-syncing/icalFeedSync.test.js`
- Tests run with Node's built-in test runner and mock Apps Script globals/services.
