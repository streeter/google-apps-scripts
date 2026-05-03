/**
 * Copy this file to: icalFeedSync.config.gs
 * Then fill in real calendar IDs and optional attendee emails.
 */

function getIcalSyncConfig() {
  return {
    // Trigger frequency for setupIcalFeedSyncTrigger()
    // Supports 1, 5, 10, 15, 30 and multiples of 60 (60=hourly, 1440=daily)
    triggerEveryMinutes: 60,

    // If true, remove future local events previously synced from a feed
    // when they no longer exist in that feed.
    deleteMissingFromFeed: true,

    // Drive-time placeholders (global defaults, can be overridden per feed)
    addDriveTimePlaceholders: false,
    defaultOriginAddress: "123 Main St, Brooklyn, NY 11201",
    // Optional place-name -> address mappings used to resolve venue names
    // into routable addresses before looking up drive time.
    placeNameAddressMap: {
      "McMoran Park": "1234 McMoran Park Rd, Your City, ST 12345",
    },
    // Placeholder is created only when computed drive time is > this threshold.
    minDriveMinutesToCreate: 10,
    driveEventTitleTemplate: "Drive ({{minutes}}m) to {{title}}",

    // Optional extra attendees added to all synced events, unless overridden per feed.
    // Leave empty to add only the target calendarId.
    defaultAttendeeEmails: [],

    // One mapping per ICS feed -> target Google Calendar
    feedMappings: [
      {
        name: "Cole Streeter",
        feedUrl:
          "https://ssprodst.blob.core.windows.net/calendars/316/46131.ics",
        calendarId: "REPLACE_WITH_COLE_CALENDAR_ID@group.calendar.google.com",

        // Optional per-feed title prefix for all synced event titles.
        // Example: "[Cole]" -> "[Cole] Practice"
        titlePrefix: "",

        // Optional per-feed extra attendee override.
        // If omitted, defaultAttendeeEmails is used.
        // If provided as [], no extra attendees are added.
        attendeeEmails: [],

        // Optional per-feed drive placeholder settings.
        addDriveTimePlaceholders: true,
        // Optional per-feed origin override. If empty, defaultOriginAddress is used.
        originAddress: "",

        // Optional per-feed place-name overrides. These are merged on top of the
        // global placeNameAddressMap and take precedence when keys overlap.
        placeNameAddressMap: {
          "McMoran Park": "1234 McMoran Park Rd, Your City, ST 12345",
        },
      },

      // Add more mappings, for example:
      // ,
      // {
      //   name: "Another Calendar",
      //   feedUrl: "https://example.com/another.ics",
      //   calendarId: "another_calendar_id@group.calendar.google.com",
      //   titlePrefix: "[Another]",
      //   attendeeEmails: ["special-person@yourcompany.com"]
      // }
    ],
  };
}
