/**
 * This script takes events from a calendar (pressumably personal), and blocks the times
 * in which there are events in another one (pressumably professional)
 *
 * This does not handle timezones at all, and just assumes that both calendars are on the same one
 *
 * Configuration:
 * - Follow the instructions on https://support.google.com/calendar/answer/37082 to share your personal calendar with your work one
 * - In your work account, create a new https://script.google.com/ project, inside it a script, and paste the contents of this file
 * - Set a trigger for an hourly run of `blockFromPersonalCalendar`
 *
 * Documentation on the Calendar API here https://developers.google.com/apps-script/reference/calendar/calendar-app
 */
const CONFIG = {
  calendarIds: GetPersonalCalendars(), // (personal) calendars from which to block time
  daysToBlockInAdvance: 30, // how many days to look ahead for
  blockedEventTitle: "Busy", // the title to use in the created events in the (work) calendar
  skipWeekends: true, // if weekend events should be skipped or not
  skipFreeAvailabilityEvents: true, // don't block events that set visibility as "Free" in the personal calendar
  workingHoursStartAt: 830, // any events ending before this time will be skipped. Use 0 if you don't care about working hours
  workingHoursEndAt: 1800, // any events starting after this time will be skipped. Use 2300
  assumeAllDayEventsInWorkCalendarIsOOO: false, // if the work calendar has an all-day event, assume it's an Out Of Office day, and don't block times
  color: CalendarApp.EventColor.YELLOW, // set the color of any newly created events (see https://developers.google.com/apps-script/reference/calendar/event-color)
};

const blockFromPersonalCalendars = () => {
  /**
   * Wrapper for the filtering functions that logs why something was skipped
   */
  const withLogging = (reason, func) => {
    return (event) => {
      const result = func.call(this, event);
      if (!result) {
        console.info(
          `ℹ️ Skipping "${event.getTitle()}" (${event.getStartTime()}) because it's ${reason}`
        );
      }
      return result;
    };
  };

  /**
   * Utility class to  make sure that, when comparing events in a personal calendar with the work's calenedar
   * configuration, things like days and working hours are respected.
   *
   * The trick is that JS stores dates as UTC. Transforming dates to the work calendar's tz as a string, and then back
   * to a Date object, ensures that the absolute numbers for day/hour/minute maintained, which is what we use in the configuration.
   */
  const CalendarAwareTimeConverter = (calendar) => {
    // Load moment.js to be able to do date operations
    eval(
      UrlFetchApp.fetch(
        "https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js"
      ).getContentText()
    );
    eval(
      UrlFetchApp.fetch(
        "https://cdnjs.cloudflare.com/ajax/libs/moment-timezone/0.5.41/moment-timezone-with-data.min.js"
      ).getContentText()
    );

    const timeZone = calendar.getTimeZone();
    const offsetedDate = (date) => moment(date).tz(timeZone);

    /*
     * Return functions that compare the given event (likely from a different calendar),
     * against the timezone of the passed in calendar.
     */
    return {
      isInAWeekend: (event) => {
        const day = offsetedDate(event.getStartTime()).day();
        return day != 0 && day != 6;
      },
      isOutOfWorkHours: (event) => {
        const startingDate = offsetedDate(event.getStartTime());
        const startingTime = startingDate.hour() * 100 + startingDate.minute();
        const endingDate = offsetedDate(event.getEndTime());
        const endingTime = endingDate.hour() * 100 + endingDate.minute();

        // Is the start time of the event within working hours, or is the ending time within working hours?
        return (
          (startingTime >= CONFIG.workingHoursStartAt &&
            startingTime <= CONFIG.workingHoursEndAt) ||
          (endingTime >= CONFIG.workingHoursStartAt &&
            endingTime <= CONFIG.workingHoursEndAt)
        );
      },
      day: (event) => {
        const startTime = offsetedDate(event.getStartTime());
        return `${startTime.year()}${startTime.month()}${startTime.date()}`;
      },
    };
  };

  /**
   * Helper to merge results from using CalendarApp and the advanced API
   * This is inefficient, but gets the best of both worlds: nice JS objects from
   * CalendarApp, and the `transparency` property from the API. If CalendarApp starts
   * exposing that in the future, there won't be a need to continue doing this.
   */
  const getRichEvents = (calendarId, start, end) => {
    const secondaryCalendar = CalendarApp.getCalendarById(calendarId);
    if (!secondaryCalendar) {
      return [];
    }
    const richEvents = secondaryCalendar.getEvents(start, end);
    const freeAvailabilityEvents = new Set(
      secondaryCalendar
        .getEvents(start, end)
        .filter((event) => event.transparency === "transparent")
        .map((event) => event.iCalUID)
    );
    richEvents.forEach((event) => {
      event.showFreeAvailability = freeAvailabilityEvents.has(event.getId());
    });
    return richEvents;
  };

  const eventTagValue = (event) =>
    `${event.getId()}-${event.getStartTime().toISOString()}`;

  const hasTimeChanges = (event, knownEvent) => {
    const eventStartTime = event.getStartTime();
    const knownEventStartTime = knownEvent.getStartTime();
    const eventEndTime = event.getEndTime();
    const knownEventEndTime = knownEvent.getEndTime();
    return (
      eventStartTime.valueOf() !== knownEventStartTime.valueOf() ||
      eventEndTime.valueOf() !== knownEventEndTime.valueOf()
    );
  };

  const mightAttend = (event, calendarId) => {
    return (
      event
        .getGuestsStatus()
        .find((s) => s.getEmail() === calendarId)
        ?.getStatus() !== "no"
    );
  };

  CONFIG.calendarIds.forEach((calendarId) => {
    console.log(`📆 Processing secondary calendar ${calendarId}`);

    const copiedEventTag = calendarEventTag(calendarId);

    const now = new Date();
    const endDate = new Date(
      Date.now() + 1000 * 60 * 60 * 24 * CONFIG.daysToBlockInAdvance
    );

    const primaryCalendar = CalendarApp.getDefaultCalendar();
    const timeZoneAware = CalendarAwareTimeConverter(primaryCalendar);

    const knownEvents = Object.assign(
      {},
      ...primaryCalendar
        .getEvents(now, endDate)
        .filter((event) => event.getTag(copiedEventTag))
        .map((event) => ({ [event.getTag(copiedEventTag)]: event }))
    );

    const knownOutOfOfficeDays = new Set(
      primaryCalendar
        .getEvents(now, endDate)
        .filter((event) => event.isAllDayEvent())
        .map((event) => timeZoneAware.day(event))
    );

    const eventsInSecondaryCalendar = getRichEvents(calendarId, now, endDate);

    const filteredEventsInSecondaryCalendar = eventsInSecondaryCalendar
      .filter(
        withLogging("already known", (event) => {
          return (
            !knownEvents.hasOwnProperty(eventTagValue(event)) ||
            hasTimeChanges(event, knownEvents[eventTagValue(event)])
          );
        })
      )
      .filter(
        withLogging("outside of work hours", (event) =>
          timeZoneAware.isOutOfWorkHours(event)
        )
      )
      .filter(
        withLogging(
          "during a weekend",
          (event) => !CONFIG.skipWeekends || timeZoneAware.isInAWeekend(event)
        )
      )
      .filter(
        withLogging(
          "during an OOO day",
          (event) =>
            !CONFIG.assumeAllDayEventsInWorkCalendarIsOOO ||
            !knownOutOfOfficeDays.has(timeZoneAware.day(event))
        )
      )
      .filter(
        withLogging(
          'marked as "Free" availabilty or is full day',
          (event) =>
            !CONFIG.skipFreeAvailabilityEvents || !event.showFreeAvailability
        )
      )
      .filter((event) => {
        // Look for similar events
        const similarEvents = primaryCalendar.getEvents(
          event.getStartTime(),
          event.getEndTime(),
          {
            search: event.getTitle(),
          }
        );
        // Find events with the same time and titles on the primary calendar. If they exist, ignore this personal event.
        if (similarEvents.length > 0) {
          console.log(
            `ℹ️ Skipping "${event.getTitle()}" (${event.getStartTime()}) because there is one or more similar events on the primary calendar`
          );
          similarEvents.forEach((sevent) => {
            console.log(`  - similar event "${sevent.getTitle()}"`);
          });
          return false;
        }
        return true;
      })
      .filter(
        withLogging(
          "not going",
          // Return events that are confirmed as "Yes", "Maybe", or "Owner". null means created.
          (event) => {
            return (
              [
                null,
                CalendarApp.GuestStatus.MAYBE,
                CalendarApp.GuestStatus.YES,
                CalendarApp.GuestStatus.OWNER,
              ].indexOf(event.getMyStatus()) >= 0
            );
          }
        )
      );

    filteredEventsInSecondaryCalendar.forEach((event) => {
      const knownEvent = knownEvents[eventTagValue(event)];
      if (knownEvent) {
        console.log(
          `📝 Need to edit "${event.getTitle()}" (${event.getStartTime()}) [${event.getId()}]`
        );
        knownEvent.deleteEvent();
      } else {
        console.log(
          `✅ Need to create "${event.getTitle()}" (${event.getStartTime()}) [${event.getId()}]`
        );
      }
      const newEvent = primaryCalendar.createEvent(
        CONFIG.blockedEventTitle,
        event.getStartTime(),
        event.getEndTime(),
        {
          description:
            "Generated with https://github.com/streeter/google-apps-scripts",
        }
      );

      newEvent
        .setTag(copiedEventTag, eventTagValue(event))
        .setVisibility(CalendarApp.Visibility.CONFIDENTIAL)
        .setColor(CONFIG.color)
        .removeAllReminders(); // Avoid double notifications

      // This is now a known event
      knownEvents[eventTagValue(event)] = newEvent;
    });

    // For each secondary event, get the tag for it
    const tagsOnSecondaryCalendar = new Set(
      eventsInSecondaryCalendar.map(eventTagValue)
    );
    // For each known event, if it has a copied event tag, but is not on the secondary calendar events, we should delete the block from the primary calendar.
    Object.values(knownEvents)
      .filter(
        (event) => !tagsOnSecondaryCalendar.has(event.getTag(copiedEventTag))
      )
      .forEach((event) => {
        console.log(
          `🗑️ Need to delete time block event on ${event.getStartTime()}, as it was removed from personal calendar`
        );
        event.deleteEvent();
      });

    if (CONFIG.skipNotAttending) {
      eventsInSecondaryCalendar
        .filter((event) => !mightAttend(event, calendarId))
        .forEach((event) => {
          const knownEvent = knownEvents[eventTagValue(event)];
          if (knownEvent) {
            console.log(
              `🗑️ Need to delete event on ${event.getStartTime()}, as it was marked as not attending`
            );
            knownEvent.deleteEvent();
          }
        });
    }
  });
};

const calendarEventTag = (calendarId) => {
  const calendarHash = Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, calendarId)
  );
  // This is undocumented, but keys fail if they are longer than 44 chars :)
  // The idea behind the SHA is to avoid collisions of the substring when you have similarly-named calendars
  return `blockFromPersonal.${calendarHash.substring(0, 15)}.originalId`;
};

/**
 * Utility function to remove all synced events. This is specially useful if you change configurations,
 * or are doing some testing
 */
const cleanUpAllCalendars = () => {
  const now = new Date();
  const endDate = new Date(
    Date.now() + 1000 * 60 * 60 * 24 * CONFIG.daysToBlockInAdvance
  );
  const tagsOfEventsToDelete = new Set(
    CONFIG.calendarIds.map(calendarEventTag)
  );

  CalendarApp.getDefaultCalendar()
    .getEvents(now, endDate)
    .filter((event) =>
      event.getAllTagKeys().some((tag) => tagsOfEventsToDelete.has(tag))
    )
    .forEach((event) => {
      console.log(
        `🗑️ Need to delete event on ${event.getStartTime()} as part of cleanup`
      );
      event.deleteEvent();
    });
};
