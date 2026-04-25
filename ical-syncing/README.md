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
- Deletes local synced events when an upstream event is canceled.
- Optionally deletes only on/after-today local synced events that are no longer present in the feed (`deleteMissingFromFeed`).
- Adds configured attendee emails to synced events.

## Setup

1. In your Apps Script project, add both files from this directory.
2. Run `listMyCalendarIds()` once (or run `syncIcalFeeds()` once) to print accessible calendar names and IDs in logs.
3. Copy `icalFeedSync.config.example.gs` to a new file named `icalFeedSync.config.gs`.
4. Edit `icalFeedSync.config.gs`:
   - Set `calendarId` for each feed mapping.
   - Set `defaultAttendeeEmails` and/or per-feed `attendeeEmails`.
   - If using drive placeholders, set `defaultOriginAddress` and set `addDriveTimePlaceholders: true` where needed.
5. In Apps Script editor:
   - Open **Services**.
   - Add **Calendar API** (Advanced Google Service).
6. Run `setupIcalFeedSyncTrigger()` once to create the periodic trigger.
7. Run `syncIcalFeeds()` once manually to validate permissions and behavior.

## Config shape

The main script expects:

```javascript
function getIcalSyncConfig() {
  return {
    triggerEveryMinutes: 15,
    deleteMissingFromFeed: true,
    addDriveTimePlaceholders: false,
    defaultOriginAddress: "123 Main St, Brooklyn, NY 11201",
    minDriveMinutesToCreate: 10,
    driveEventTitleTemplate: "Drive ({{minutes}}m) to {{title}}",
    defaultAttendeeEmails: ["person@company.com"],
    feedMappings: [
      {
        name: "Cole Streeter",
        feedUrl: "https://example.com/calendar.ics",
        calendarId: "your_calendar_id@group.calendar.google.com",
        attendeeEmails: [],
        addDriveTimePlaceholders: true,
        originAddress: ""
      }
    ]
  };
}
```

## Notes

- `feedMappings` can include many feed -> calendar routes.
- Per-feed `attendeeEmails` overrides `defaultAttendeeEmails` when non-empty.
- Per-feed `addDriveTimePlaceholders` controls whether drive placeholders are managed for that feed.
- Placeholders are only created when computed drive time is strictly greater than `minDriveMinutesToCreate` (default `10`).
- Drive placeholders are skipped for all-day events, events without a location, unroutable addresses, and pre-today events.
- Drive placeholders are tied to source synced events using metadata (`sourceSyncKey` and `sourceEventId`) and are managed/deleted safely.
- The script uses event metadata (`extendedProperties.private`) to track synced items and detect changes.
- Delete operations are guarded to only remove events that are verifiably managed by this script/feed.
- On first `syncIcalFeeds()` run, the script logs all accessible calendar names/IDs once, to help with initial config.

## Local tests

- Run: `npm test`
- Test file: `tests/icalFeedSync.test.js`
- Tests run with Node's built-in test runner and mock Apps Script globals/services.
