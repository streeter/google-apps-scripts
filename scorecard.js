// This is a google apps script to automatically create time blocks on your calendar
// after interviews to give you time between meetings to fill out scorecard feedback.
// It also gives you a link to the scorecard in the event location.
//
// It's very simple in that it
//   1) is not resilient to reschedules/cancellations (it'll leave the block there)
//   2) only schedules a block if there isn't already an event in that timeslot
//      [this doubles as an idempotency mechanism since this script runs repeatedly]
//
// The size of the block is configurable (right below this line)
const TIME_BLOCK_MINS = 15;

const isInterviewEvent = (event) => {
  if (event.getTitle().includes("Team Screen")) return true;
  if (event.getDescription().includes("Thanks for interviewing")) return true;
  if (
    event
      .getGuestList()
      .find((guest) => guest.getName().includes("GoodTime Sync"))
  )
    return true;

  return false;
};

const hasSpaceForBlockAfter = (cal, eventEnds, timeBlockEnds) => {
  const existingEvents = cal.getEvents(eventEnds, timeBlockEnds);
  const existingNotAllDayEvents = existingEvents.filter(
    (evt) => !evt.isAllDayEvent()
  );
  return existingNotAllDayEvents.length === 0; // Don't schedule if there's something there already
};

const getScorecardLink = (event) => {
  // This is a lazy regex; contributions welcome
  const link = event
    .getDescription()
    .match(/https:\/\/app.greenhouse.io\/guides\/[^"]+/);
  return link ? link[0] : "";
};

function main() {
  const today = new Date();
  const twoWeeksFromNow = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);

  const cal = CalendarApp.getDefaultCalendar();
  const myEvents = cal.getEvents(today, twoWeeksFromNow);

  myEvents.forEach((event) => {
    const eventEnds = event.getEndTime();
    const timeBlockEnds = new Date(
      eventEnds.getTime() + TIME_BLOCK_MINS * 60000
    );

    if (
      isInterviewEvent(event) &&
      hasSpaceForBlockAfter(cal, eventEnds, timeBlockEnds)
    ) {
      const location = getScorecardLink(event);
      cal.createEvent(
        "Fill out interview scorecard",
        eventEnds,
        timeBlockEnds,
        { location: location }
      );
    }
  });
}
