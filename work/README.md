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
