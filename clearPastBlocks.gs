/**
 * This script looks at past events that are tagged by `blockFromPersonalCalendar.gs`
 * and deletes them
 */

const ORIGINAL_EVENT_TAG = "blockFromPersonal.originalEventId";

function clearPastBlocks() {
  const anHourAgo = new Date(Date.now() - 1000 * 60 * 60);
  const startDate = new Date(Date.now() - 1000 * 60 * 60 * 24);

  const calendar = CalendarApp.getDefaultCalendar();

  calendar
    .getEvents(startDate, anHourAgo)
    .filter((event) => event.getTag(ORIGINAL_EVENT_TAG))
    .forEach((event) => {
      Logger.log(
        `Deleting ${event.getStartTime().toLocaleString()}-${event
          .getEndTime()
          .toLocaleString()} - ${event.getTitle()}`
      );
      //event.deleteEvent();
    });
}
