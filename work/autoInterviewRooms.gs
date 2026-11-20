// ============================================================
// CONFIG
// ============================================================

/**
 * The (usually) three letter ID of the office to book rooms in.
 * Ex:
 *  DUB = Dublin
 *  SSF = South San Francisco
 *  SEA = Seattle
 * To find the ID of your office, look here: https://calendar.google.com/calendar/u/0/r/settings/browseresources
 */
const OFFICE_ID = "SSF";

/** Rooms with more seats than this are never booked. */
const MAX_ROOM_CAPACITY = 5;

/** How long the discovered room list is cached (Apps Script allows at most 6 hours). */
const ROOMS_CACHE_SECONDS = 6 * 60 * 60;

const BUFFER_MINUTES = 15;
const LOOK_AHEAD_DAYS = 14;
const ROOM_EVENT_TITLE = "Auto-Booked Interview Room";

// ============================================================
// COLLECT EVENTS
// ============================================================

function getTimeWindow() {
  var now = new Date();
  var lookAhead = new Date(
    now.getTime() + LOOK_AHEAD_DAYS * 24 * 60 * 60 * 1000,
  );
  return { now: now, lookAhead: lookAhead };
}

function getUpcomingInterviews() {
  var window = getTimeWindow();
  var events = CalendarApp.getDefaultCalendar().getEvents(
    window.now,
    window.lookAhead,
  );

  return events.filter(function (event) {
    var desc = event.getDescription() || "";
    return desc.includes("Thanks for interviewing");
  });
}

function getExistingRoomEvents() {
  var window = getTimeWindow();
  var events = CalendarApp.getDefaultCalendar().getEvents(
    window.now,
    window.lookAhead,
  );

  return events.filter(function (event) {
    return (
      event.getTitle() === ROOM_EVENT_TITLE &&
      (event.getDescription() || "").includes("Auto-booked")
    );
  });
}

// ============================================================
// PAIR BY EXACT TIME MATCH
// ============================================================

function pairEvents(interviews, roomEvents) {
  var paired = [];
  var unpaired = [];
  var orphanRooms = [];

  var matchedRoomIndices = {};

  for (var i = 0; i < interviews.length; i++) {
    var interview = interviews[i];
    var expectedStart =
      interview.getStartTime().getTime() - BUFFER_MINUTES * 60 * 1000;
    var expectedEnd =
      interview.getEndTime().getTime() + BUFFER_MINUTES * 60 * 1000;
    var foundMatch = false;

    for (var j = 0; j < roomEvents.length; j++) {
      if (matchedRoomIndices[j]) continue;

      var roomStart = roomEvents[j].getStartTime().getTime();
      var roomEnd = roomEvents[j].getEndTime().getTime();

      if (roomStart === expectedStart && roomEnd === expectedEnd) {
        paired.push({ interview: interview, roomEvent: roomEvents[j] });
        matchedRoomIndices[j] = true;
        foundMatch = true;
        break;
      }
    }

    if (!foundMatch) {
      unpaired.push(interview);
    }
  }

  for (var k = 0; k < roomEvents.length; k++) {
    if (!matchedRoomIndices[k]) {
      orphanRooms.push(roomEvents[k]);
    }
  }

  return { paired: paired, unpaired: unpaired, orphanRooms: orphanRooms };
}

// ============================================================
// DELETE ORPHAN ROOM EVENTS
// ============================================================

function deleteOrphanRoomEvents(orphanRooms) {
  for (var i = 0; i < orphanRooms.length; i++) {
    Logger.log(
      "🗑️ Deleting orphan: " +
        orphanRooms[i].getStartTime().toLocaleString() +
        " – " +
        orphanRooms[i].getEndTime().toLocaleString(),
    );
    orphanRooms[i].deleteEvent();
  }
  return orphanRooms.length;
}

// ============================================================
// ROOM DISCOVERY
// ============================================================

/**
 * Lists the bookable rooms in `officeId`, ordered by (capacity, name) so the
 * smallest room is tried first. Skips virtual rooms (no capacity) and rooms
 * larger than MAX_ROOM_CAPACITY.
 */
function listRooms(officeId) {
  const rooms = [];
  let pageToken = null;

  do {
    const response = AdminDirectory.Resources.Calendars.list("my_customer", {
      maxResults: 100,
      pageToken: pageToken,
    });

    for (const room of response.items || []) {
      if (!room.capacity) {
        // ignore virtual rooms
        continue;
      }
      if (room.capacity > MAX_ROOM_CAPACITY) {
        // avoid booking rooms that are too big
        continue;
      }

      if (!room.buildingId || !room.buildingId.startsWith(officeId)) {
        // ignore rooms from other offices
        continue;
      }

      rooms.push({
        id: room.resourceEmail,
        name: room.resourceName,
        capacity: room.capacity,
      });
    }

    pageToken = response.nextPageToken;
  } while (pageToken);

  rooms.sort((a, b) => {
    if (a.capacity !== b.capacity) {
      return a.capacity - b.capacity;
    }
    return a.name.localeCompare(b.name);
  });

  return rooms;
}

/**
 * Same as `listRooms`, but cached so scheduled runs don't page through the
 * whole company's room directory every time. Empty results aren't cached, so
 * a wrong OFFICE_ID is picked up on the next run once it's fixed.
 */
function getCachedRooms(officeId) {
  const cache = CacheService.getUserCache();
  const cacheKey = "auto-interview-rooms:" + officeId;
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  const rooms = listRooms(officeId);
  if (rooms.length > 0) {
    cache.put(cacheKey, JSON.stringify(rooms), ROOMS_CACHE_SECONDS);
  }
  return rooms;
}

// ============================================================
// FIND AVAILABLE ROOM & CREATE EVENT
// ============================================================

function findAvailableRoom(rooms, startTime, endTime) {
  var bufferedStart = new Date(
    startTime.getTime() - BUFFER_MINUTES * 60 * 1000,
  );
  var bufferedEnd = new Date(endTime.getTime() + BUFFER_MINUTES * 60 * 1000);

  for (var i = 0; i < rooms.length; i++) {
    var room = rooms[i];
    var request = {
      timeMin: bufferedStart.toISOString(),
      timeMax: bufferedEnd.toISOString(),
      items: [{ id: room.id }],
    };
    var response = Calendar.Freebusy.query(request);
    var busy = response.calendars[room.id].busy;

    if (busy.length === 0) {
      Logger.log("✅ Available: " + room.name);
      return room;
    }
    Logger.log("❌ Busy: " + room.name);
  }

  Logger.log("⚠️ No rooms available!");
  return null;
}

function createRoomEvent(interview, room) {
  var bufferedStart = new Date(
    interview.getStartTime().getTime() - BUFFER_MINUTES * 60 * 1000,
  );
  var bufferedEnd = new Date(
    interview.getEndTime().getTime() + BUFFER_MINUTES * 60 * 1000,
  );

  var event = {
    summary: ROOM_EVENT_TITLE,
    colorId: "3",
    description:
      "Auto-booked for: " +
      interview.getTitle() +
      "\n" +
      "Interview: " +
      interview.getStartTime().toLocaleString() +
      " – " +
      interview.getEndTime().toLocaleString(),
    start: { dateTime: bufferedStart.toISOString() },
    end: { dateTime: bufferedEnd.toISOString() },
    attendees: [{ email: room.id, resource: true }],
  };

  var created = Calendar.Events.insert(event, "primary", {
    sendUpdates: "all",
  });
  Logger.log(
    "📅 Created: " +
      room.name +
      " @ " +
      bufferedStart.toLocaleString() +
      " – " +
      bufferedEnd.toLocaleString(),
  );
  return created;
}

function createMissingRoomEvents(unpaired, rooms) {
  var created = 0;
  var failed = 0;

  for (var i = 0; i < unpaired.length; i++) {
    var interview = unpaired[i];
    var room = findAvailableRoom(
      rooms,
      interview.getStartTime(),
      interview.getEndTime(),
    );
    if (room) {
      createRoomEvent(interview, room);
      created++;
    } else {
      Logger.log(
        "🚨 No room for: " +
          interview.getTitle() +
          " @ " +
          interview.getStartTime().toLocaleString(),
      );
      failed++;
    }
  }

  return { created: created, failed: failed };
}

// ============================================================
// MAIN
// ============================================================

function main() {
  Logger.log("═══ Discovering rooms in " + OFFICE_ID + " ═══");
  var rooms = getCachedRooms(OFFICE_ID);
  Logger.log(
    "Found " +
      rooms.length +
      " bookable room(s): " +
      rooms.map((room) => room.name).join(", "),
  );

  Logger.log("═══ Collecting events ═══");
  var interviews = getUpcomingInterviews();
  var roomEvents = getExistingRoomEvents();
  Logger.log(
    "Found " +
      interviews.length +
      " interview(s), " +
      roomEvents.length +
      " room event(s)",
  );

  Logger.log("═══ Pairing ═══");
  var result = pairEvents(interviews, roomEvents);
  Logger.log(
    "Paired: " +
      result.paired.length +
      " | Unbooked: " +
      result.unpaired.length +
      " | Orphans: " +
      result.orphanRooms.length,
  );

  Logger.log("═══ Cleaning orphans ═══");
  var deletedCount = deleteOrphanRoomEvents(result.orphanRooms);

  Logger.log("═══ Booking rooms ═══");
  var created = createMissingRoomEvents(result.unpaired, rooms);

  var summary =
    "✅ Done! Kept: " +
    result.paired.length +
    " | Created: " +
    created.created +
    " | Deleted: " +
    deletedCount +
    " | No room: " +
    created.failed;
  Logger.log(summary);
  return summary;
}

function __deleteAllBookings() {
  deleteOrphanRoomEvents(getExistingRoomEvents());
}
