/**
 * Copy this file to: icalFeedSync.config.gs
 * Then fill in real calendar IDs and attendee emails.
 */

function getIcalSyncConfig() {
  return {
    // Trigger frequency for setupIcalFeedSyncTrigger()
    // Valid Apps Script values: 1, 5, 10, 15, 30
    triggerEveryMinutes: 15,

    // If true, remove future local events previously synced from a feed
    // when they no longer exist in that feed.
    deleteMissingFromFeed: true,

    // Added as attendees on all synced events, unless overridden per feed.
    defaultAttendeeEmails: [
      "example1@yourcompany.com",
      "example2@yourcompany.com"
    ],

    // One mapping per ICS feed -> target Google Calendar
    feedMappings: [
      {
        name: "Cole Streeter",
        feedUrl: "https://ssprodst.blob.core.windows.net/calendars/316/46131.ics",
        calendarId: "REPLACE_WITH_COLE_CALENDAR_ID@group.calendar.google.com",

        // Optional per-feed attendee override.
        // If empty, defaultAttendeeEmails is used.
        attendeeEmails: []
      }

      // Add more mappings, for example:
      // ,
      // {
      //   name: "Another Calendar",
      //   feedUrl: "https://example.com/another.ics",
      //   calendarId: "another_calendar_id@group.calendar.google.com",
      //   attendeeEmails: ["special-person@yourcompany.com"]
      // }
    ]
  };
}
