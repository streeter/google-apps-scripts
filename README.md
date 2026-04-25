# google-app-scripts

My collection of [Google's App Script](https://script.google.com/) scripts

## 🚀 script.google.com Setup

1. Go to `https://script.google.com/` and create a new project.
2. Add new script files and paste in the `.gs`/`.js` files you want to run.
3. If a script has an `.example` config file, create a real config file in Apps Script and fill in your values.
4. If needed, add Advanced Services in **Services** (for example Calendar API).
5. Run the main function once manually to authorize permissions.
6. Add a time-based trigger in **Triggers** for periodic scripts.

## 💼 Work Scripts

- [`work/README.md`](work/README.md): setup and details for work-calendar scripts.
- [`work/colorBasedOnAttendees.gs`](work/colorBasedOnAttendees.gs): colors each calendar event based on who is in the meeting, or what it is for.
- [`work/blockFromPersonalCalendar.gs`](work/blockFromPersonalCalendar.gs): takes events from a calendar (presumably personal), and blocks those times in another one (presumably professional).
- [`work/clearPastBlocks.js`](work/clearPastBlocks.js): deletes older block events created by the blocking workflow.
- [`work/scheduleInterviewFeedback.gs`](work/scheduleInterviewFeedback.gs): looks for interviews, and schedules feedback blocks after them.
- [`work/colorAttendeeConfig.example.gs`](work/colorAttendeeConfig.example.gs): example config for coloring rules.
- [`work/getPersonalCalendar.example.gs`](work/getPersonalCalendar.example.gs): example config for personal/source calendars.

## 🔁 iCal Syncing

- [`ical-syncing/icalFeedSync.gs`](ical-syncing/icalFeedSync.gs): syncs events from one or more iCal feeds into specific Google Calendars, updates changed events, and adds configured attendees.
- [`ical-syncing/icalFeedSync.config.example.gs`](ical-syncing/icalFeedSync.config.example.gs): example config file for feed mappings and attendee lists.
- [`ical-syncing/README.md`](ical-syncing/README.md): setup, configuration, trigger setup, and behavior details.

## ✅ Testing

- Run tests locally with `npm test`.
- Run formatting lint with `npm run lint`.
- Auto-format checked files with `npm run format`.
- Current suite: [`ical-syncing/icalFeedSync.test.js`](ical-syncing/icalFeedSync.test.js).
