/**
 * Copy this file to: icalFeedSync.config.gs
 * Then fill in real calendar IDs and attendee emails.
 */

function getIcalSyncConfig() {
  return {
    // Trigger frequency for setupIcalFeedSyncTrigger()
    // Valid Apps Script values: 1, 5, 10, 15, 30
    triggerEveryMinutes: 60,

    // If true, remove future local events previously synced from a feed
    // when they no longer exist in that feed.
    deleteMissingFromFeed: true,

    // Drive-time placeholders (global defaults, can be overridden per feed)
    addDriveTimePlaceholders: false,
    defaultOriginAddress: "123 Main St, Brooklyn, NY 11201",
    // Placeholder is created only when computed drive time is > this threshold.
    minDriveMinutesToCreate: 10,
    driveEventTitleTemplate: "Drive ({{minutes}}m) to {{title}}",

    // Added as attendees on all synced events, unless overridden per feed.
    defaultAttendeeEmails: [
      "example1@yourcompany.com",
      "example2@yourcompany.com",
    ],

    // One mapping per ICS feed -> target Google Calendar
    feedMappings: [
      {
        name: "Cole Streeter",
        feedUrl:
          "https://ssprodst.blob.core.windows.net/calendars/316/46131.ics",
        calendarId: "REPLACE_WITH_COLE_CALENDAR_ID@group.calendar.google.com",

        // Optional per-feed attendee override.
        // If empty, defaultAttendeeEmails is used.
        attendeeEmails: [],

        // Optional per-feed drive placeholder settings.
        addDriveTimePlaceholders: true,
        // Optional per-feed origin override. If empty, defaultOriginAddress is used.
        originAddress: "",
      },

      // Add more mappings, for example:
      // ,
      // {
      //   name: "Another Calendar",
      //   feedUrl: "https://example.com/another.ics",
      //   calendarId: "another_calendar_id@group.calendar.google.com",
      //   attendeeEmails: ["special-person@yourcompany.com"]
      // }
    ],
  };
}
