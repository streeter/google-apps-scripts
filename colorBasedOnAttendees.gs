// ******************************************************************************************************
// *                This is the section that must be edited for anyone deploying this script            *
// ******************************************************************************************************

// Set main variables that are used to choose the calendar and do the right mapping for classifications
const myUsername = "streeter";
const domainName = "stripe.com";
const calendarName = myUsername + domainName;
const myalias = calendarName; // Or set to another alias if you have one
const myabp = "i-do-not-have-an-abp";
const firstName = myUsername; // Used to detect 1:1s
const myorg = ["mips-team@", "tech@"];
const vips = [
  // TODO
];

// Choose your color coding here
const color1on1 = CalendarApp.EventColor.PALE_GREEN;
const colorExternal = CalendarApp.EventColor.MAUVE;
const colorFocus = CalendarApp.EventColor.PALE_BLUE;
const colorRecruiting = CalendarApp.EventColor.GREEN;
const colorVIP = CalendarApp.EventColor.RED;
const colorMyOrg = CalendarApp.EventColor.YELLOW;
const colorInternal = CalendarApp.EventColor.CYAN;

const Status = {
  external: 1,
  vip: 1 << 1,
  myorg: 1 << 2,
};

function ColorEvents() {
  const today = new Date();
  const nextweek = new Date().setDate(today.getDate() + 7);

  // ******************************************************************************************************
  // *                                         Start of the Script                                        *
  // ******************************************************************************************************
  var calendars = CalendarApp.getCalendarsByName(calendarName);
  for (let i = 0; i < calendars.length; i++) {
    let calendar = calendars[i];
    let events = calendar.getEvents(today, nextweek);

    // Get all events in the next week, and loop through each event
    for (let j = 0; j < events.length; j++) {
      // Select an event
      let event = events[j];
      // Get the event title and guest list
      let title = event.getTitle().toLowerCase();
      let guests = event.getGuestList();

      // Logger.log('Meeting : "%s"', title)
      // Logger.log('Guests : "%s"', g.length)

      // Catch any blocked time in the calendar or events with no invitees
      if (
        title.includes("dns") ||
        title.includes("focus time") ||
        title.includes("blocked") ||
        g.length === 0
      ) {
        // Logger.log("focus")
        event.setColor(colorFocus);
        continue;
      }

      // Catch any interviews, tropes or other recruiting meetings
      if (
        title.includes("interview") ||
        title.includes("trope") ||
        title.includes("recruiting")
      ) {
        event.setColor(colorRecruiting);
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
            // Logger.log("1:1")
            event.setColor(color1on1);
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
          // Logger.log("external")
          // event.setColor(colorExternal);
          status |= Status.external;
          continue;
        }

        // Check for any VIP organiser or attendees
        if (checkVIPs(vips, guestEmail, event.getCreators().toString())) {
          // Logger.log("vip")
          // event.setColor(colorVIP)
          status |= Status.vip;
          continue;
        }

        if (checkOrgMeetings(myorg, guestEmail)) {
          // Logger.log("myorg")
          // event.setColor(colorMyOrg);
          status |= Status.myorg;
          continue;
        }
      }
      try {
        event.setColor(setPriorityColor(status));
      } catch (err) {
        Logger.log("Unable to set the color with an error " + err);
      }
    }
  }
}

// Return one color based on this order
// External Meeting, Org Meeting, VIP Meeting, internal meeting
// Change the order based on your own preference
function setPriorityColor(status) {
  if (status & Status.external) {
    return colorExternal;
  }

  if (status & Status.myorg) {
    return colorMyOrg;
  }

  if (status & Status.vip) {
    return colorVIP;
  }

  return colorInternal;
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
