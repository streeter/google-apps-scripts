# Work Calendar Scripts

This directory contains scripts focused on day-to-day work calendar management.

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
