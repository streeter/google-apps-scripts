# iCal Feed Syncing (Google Apps Script)

This directory contains a Google Apps Script that syncs one or more remote iCal (`.ics`) feeds to specific Google Calendars.

## Files

- `icalFeedSync.gs`: main sync script.
- `icalFeedSync.config.example.gs`: example config file. Copy this to `icalFeedSync.config.gs` and fill in your values.

## What the script does

- Pulls each configured iCal feed URL.
- Syncs future events into the configured local Google Calendar.
- Updates local events when upstream iCal event details change.
- Deletes local synced events when an upstream event is canceled.
- Optionally deletes future local synced events that are no longer present in the feed (`deleteMissingFromFeed`).
- Adds configured attendee emails to synced events.

## Setup

1. In your Apps Script project, add both files from this directory.
2. Copy `icalFeedSync.config.example.gs` to a new file named `icalFeedSync.config.gs`.
3. Edit `icalFeedSync.config.gs`:
   - Set `calendarId` for each feed mapping.
   - Set `defaultAttendeeEmails` and/or per-feed `attendeeEmails`.
4. In Apps Script editor:
   - Open **Services**.
   - Add **Calendar API** (Advanced Google Service).
5. Run `setupIcalFeedSyncTrigger()` once to create the periodic trigger.
6. Run `syncIcalFeeds()` once manually to validate permissions and behavior.

## Config shape

The main script expects:

```javascript
function getIcalSyncConfig() {
  return {
    triggerEveryMinutes: 15,
    deleteMissingFromFeed: true,
    defaultAttendeeEmails: ["person@company.com"],
    feedMappings: [
      {
        name: "Cole Streeter",
        feedUrl: "https://example.com/calendar.ics",
        calendarId: "your_calendar_id@group.calendar.google.com",
        attendeeEmails: []
      }
    ]
  };
}
```

## Notes

- `feedMappings` can include many feed -> calendar routes.
- Per-feed `attendeeEmails` overrides `defaultAttendeeEmails` when non-empty.
- The script uses event metadata (`extendedProperties.private`) to track synced items and detect changes.
