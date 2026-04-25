// ******************************************************************************************************
// *                This is the section that must be edited for anyone deploying this script            *
// ******************************************************************************************************

// Set main variables that are used to choose the calendar and do the right mapping for classifications
const calendarName = Session.getActiveUser().getEmail();
const myalias = calendarName; // Or set to another alias if you have one
const myabp = "i-do-not-have-an-abp";

const domainName = calendarName.split("@")[-1];
const myUsername = calendarName.split("@")[0];
const firstName = myUsername; // Used to detect 1:1s

// These are defined in a colorAttendeeConfig.js
const myorg = GetColorEventOrgNames();
const vips = GetColorEventVips();

// Choose your color coding here
const ColorEventColors = {
  oneOnOne: CalendarApp.EventColor.PALE_GREEN,
  external: CalendarApp.EventColor.MAUVE,
  focus: CalendarApp.EventColor.PALE_BLUE,
  recruiting: CalendarApp.EventColor.GREEN,
  vip: CalendarApp.EventColor.RED,
  org: CalendarApp.EventColor.YELLOW,
  internal: CalendarApp.EventColor.GRAY,
  other: CalendarApp.EventColor.CYAN,
};

const ColorEventStatus = {
  external: 1,
  vip: 1 << 1,
  myorg: 1 << 2,
};

function ColorEvents() {
  const today = new Date();
  const nextweek = new Date();
  nextweek.setDate(today.getDate() + 7);

  // ******************************************************************************************************
  // *                                         Start of the Script                                        *
  // ******************************************************************************************************
  var calendars = CalendarApp.getCalendarsByName(calendarName);
  for (let i = 0; i < calendars.length; i++) {
    let calendar = calendars[i];
    let events = calendar.getEvents(today, nextweek);

    Logger.log(`Processing calendar ${calendar.getName()}`);

    // Get all events in the next week, and loop through each event
    for (let j = 0; j < events.length; j++) {
      // Select an event
      let event = events[j];
      // Get the event title and guest list
      let title = event.getTitle().toLowerCase();
      let guests = event.getGuestList();

      Logger.log(`Looking at event ${title} with ${guests.length} guests`);

      // Catch any blocked time in the calendar or events with no invitees
      if (
        title.includes("dns") ||
        title.includes("focus time") ||
        title.includes("blocked") ||
        guests.length === 0
      ) {
        updateEventColor(event, ColorEventColors.focus);
        continue;
      }

      // Catch any interviews, tropes or other recruiting meetings
      if (
        title.includes("interview") ||
        title.includes("trope") ||
        title.includes("[VC]") ||
        title.includes("recruiting")
      ) {
        updateEventColor(event, ColorEventColors.recruiting);
        continue;
      }

      // Catch any possible 1:1 meetings and color-code them.
      // Identifies possible 1:1's by looking for meetings with only an organiser and one guest
      // Also checks for common words in title to validate
      if (guests.length === 1) {
        // Is this a one on one?
        if (
          (!event.getCreators().includes(myalias) &&
            guests[0].getEmail().includes(myalias)) ||
          event.getCreators().includes(myalias) ||
          event.getCreators().includes(myabp)
        ) {
          if (
            title.includes("1:1") ||
            title.includes(firstName) ||
            title.includes("catch-up")
          ) {
            updateEventColor(event, ColorEventColors.oneOnOne);
            continue;
          }
        }
      }

      // Loop through all guests
      let status = 0;
      for (let i = 0; i < guests.length; i++) {
        const guestEmail = guests[i].getEmail();

        // Check for any Non-domain organisers or attendees
        if (
          checkNonOrg(domainName, guestEmail, event.getCreators().toString())
        ) {
          status |= ColorEventStatus.external;
          continue;
        }

        // Check for any VIP organiser or attendees
        if (checkVIPs(vips, guestEmail, event.getCreators().toString())) {
          status |= ColorEventStatus.vip;
          continue;
        }

        if (checkOrgMeetings(myorg, guestEmail)) {
          status |= ColorEventStatus.myorg;
          continue;
        }
      }

      updateEventColor(event, setPriorityColor(status));
    }
  }
}

// Return one color based on this order
// External Meeting, Org Meeting, VIP Meeting, internal meeting
// Change the order based on your own preference
function setPriorityColor(status) {
  if (status & ColorEventStatus.external) {
    return ColorEventColors.external;
  }

  if (status & ColorEventStatus.myorg) {
    return ColorEventColors.org;
  }

  if (status & ColorEventStatus.vip) {
    return ColorEventColors.vip;
  }

  return ColorEventColors.internal;
}

// This checks to see if a meeting has any of my Tech Services google groups on the invite list but not eileen (which implies it is a TS meeting, not a GTM meeting)
function checkOrgMeetings(orgList, guestEmail) {
  let tsMeeting = false;

  for (var j = 0; j < orgList.length; j++) {
    if (guestEmail.includes(orgList[j])) {
      tsMeeting = true;
    }
    if (guestEmail.includes("eileen@")) {
      tsMeeting = false;
    }
  }
  return tsMeeting;
}

// This checks a guest/organiser for an event against a list of VIP's
function checkVIPs(vipList, guestName, organiserName) {
  for (var n = 0; n < vipList.length; n++) {
    if (guestName.includes(vipList[n]) || organiserName.includes(vipList[n])) {
      return true;
    }
  }
  return false;
}

// This checks a guest/organiser for an event to see if they are from outside the domain
function checkNonOrg(domainName, guestName, organiserName) {
  if (
    !guestName.includes(domainName) &&
    !guestName.includes("calendar.google.com")
  ) {
    return true;
  }
  if (!organiserName.includes(domainName)) {
    return true;
  }
  return false;
}

function updateEventColor(event, color) {
  if (event.getColor() === color) {
    return;
  }

  try {
    event.setColor(color);
  } catch (err) {
    console.error("Unable to set the color with an error", err);
  }
}
