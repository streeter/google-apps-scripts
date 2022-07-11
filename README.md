# google-app-scripts

My collection of [Google's App Script](https://script.google.com/) scripts

## 🗓 Calendar

- [`colorBasedOnAttendees.gs`](colorBasedOnAttendees.gs): colors each calendar event based on who is in the meeting, or what it is for.
- [`blockFromPersonalCalendar.gs`](blockFromPersonalCalendar.gs): takes events from a calendar (pressumably personal), and blocks the times in which there are events in another one (pressumably professional)
- [`clearPastBlocks.gs`](clearPastBlocks.gs): looks at past events that are named in specific ways (see, for example [`blockFromPersonalCalendar.gs`](blockFromPersonalCalendar.gs)), and deletes them

The configurations for the files should live in `colorAttendeeConfig.gs` and `getPersonalCalendar.gs` (of which there are example files included). Of note, the configs should be defined before the other scripts in the App Script project - you can prepend an `a` in front of the file name and then sort the files alphabetically and they'll work.
