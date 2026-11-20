const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

function buildBaseContext(overrides = {}) {
  const base = {
    console: {
      log: () => {},
      info: () => {},
      error: () => {},
    },
    Date,
    JSON,
    Math,
    Logger: { log: () => {} },
    Calendar: {
      Events: {
        get: () => ({ status: "confirmed" }),
      },
    },
    CalendarApp: {
      EventColor: {
        PALE_GREEN: "PALE_GREEN",
        MAUVE: "MAUVE",
        PALE_BLUE: "PALE_BLUE",
        GREEN: "GREEN",
        RED: "RED",
        YELLOW: "YELLOW",
        GRAY: "GRAY",
        CYAN: "CYAN",
      },
      Visibility: {
        CONFIDENTIAL: "CONFIDENTIAL",
      },
      getDefaultCalendar: () => ({
        getEvents: () => [],
      }),
      getCalendarsByName: () => [],
    },
    Session: {
      getActiveUser: () => ({
        getEmail: () => "me@example.com",
      }),
    },
    GetColorEventOrgNames: () => [],
    GetColorEventVips: () => [],
    GetPersonalCalendars: () => [],
    GetSelfEmail: () => "me@example.com",
    UrlFetchApp: {
      fetch: () => ({
        getContentText: () => "",
      }),
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: "SHA_256" },
      computeDigest: (_algorithm, input) => {
        const hash = crypto
          .createHash("sha256")
          .update(String(input), "utf8")
          .digest();
        return Array.from(hash.values());
      },
      base64Encode: (bytes) => Buffer.from(bytes).toString("base64"),
    },
  };

  return Object.assign(base, overrides);
}

function loadScript(relativePath, overrides = {}) {
  const context = buildBaseContext(overrides);
  vm.createContext(context);
  const scriptPath = path.resolve(__dirname, "..", relativePath);
  const code = fs.readFileSync(scriptPath, "utf8");
  vm.runInContext(code, context, { filename: scriptPath });
  return context;
}

test("clearPastBlocks deletes only generated events with eligible titles", () => {
  const deleted = [];
  const events = [
    makeCalendarEvent(
      "Busy",
      "Generated with github.com/streeter/google-apps-scripts",
      deleted,
      "new-format",
    ),
    makeCalendarEvent(
      "Fill out interview scorecard",
      "Generated with https://github.com/streeter/google-apps-scripts",
      deleted,
      "old-format",
    ),
    makeCalendarEvent("Busy", "random description", deleted, "not-generated"),
    makeCalendarEvent(
      "Other Title",
      "Generated with github.com/streeter/google-apps-scripts",
      deleted,
      "wrong-title",
    ),
  ];

  const ctx = loadScript("work/clearPastBlocks.js", {
    CalendarApp: {
      getDefaultCalendar: () => ({
        getEvents: () => events,
      }),
    },
  });

  ctx.clearPastBlocks();

  assert.deepEqual(deleted.sort(), ["new-format", "old-format"]);
});

test("scheduleInterviewFeedback helper predicates behave as expected", () => {
  const ctx = loadScript("work/scheduleInterviewFeedback.gs");
  const isInterviewEvent = vm.runInContext("isInterviewEvent", ctx);
  const hasSpaceForBlockAfter = vm.runInContext("hasSpaceForBlockAfter", ctx);
  const getScorecardLink = vm.runInContext("getScorecardLink", ctx);

  assert.equal(isInterviewEvent(null), false);

  assert.equal(
    isInterviewEvent(
      makeInterviewEvent({
        title: "Team Screen with Candidate",
      }),
    ),
    true,
  );
  assert.equal(
    isInterviewEvent(
      makeInterviewEvent({
        description: "Thanks for interviewing with our team.",
      }),
    ),
    true,
  );
  assert.equal(
    isInterviewEvent(
      makeInterviewEvent({
        guests: [{ getName: () => "GoodTime Sync Bot" }],
      }),
    ),
    true,
  );
  assert.equal(
    isInterviewEvent(
      makeInterviewEvent({
        title: "Regular Standup",
        description: "No hiring content",
        guests: [{ getName: () => "Alice" }],
      }),
    ),
    false,
  );

  assert.equal(
    getScorecardLink(
      makeInterviewEvent({
        description:
          "Please review https://app.greenhouse.io/guides/12345 before debrief",
      }),
    ),
    "https://app.greenhouse.io/guides/12345",
  );
  assert.equal(
    getScorecardLink(makeInterviewEvent({ description: "no scorecard link" })),
    "",
  );

  const calWithConflict = {
    getEvents: () => [{ isAllDayEvent: () => false }],
  };
  const calWithOnlyAllDay = {
    getEvents: () => [{ isAllDayEvent: () => true }],
  };

  assert.equal(
    hasSpaceForBlockAfter(calWithConflict, new Date(), new Date()),
    false,
  );
  assert.equal(
    hasSpaceForBlockAfter(calWithOnlyAllDay, new Date(), new Date()),
    true,
  );
});

test("scheduleInterviewFeedback tag-protected deletion and interview checks", () => {
  const ctx = loadScript("work/scheduleInterviewFeedback.gs");
  const deleteInterviewBlock = vm.runInContext("deleteInterviewBlock", ctx);
  const interviewInOriginalSpot = vm.runInContext(
    "interviewInOriginalSpot",
    ctx,
  );
  const SCORECARD_TAG = vm.runInContext("SCORECARD_TAG", ctx);

  let deleted = 0;
  const validTaggedEvent = {
    getAllTagKeys: () => [SCORECARD_TAG],
    getTag: () => "source-event-id",
    getTitle: () => "Fill out interview scorecard",
    deleteEvent: () => {
      deleted += 1;
    },
  };
  const invalidTaggedEvent = {
    getAllTagKeys: () => [],
    getTag: () => "",
    getTitle: () => "Fill out interview scorecard",
    deleteEvent: () => {
      deleted += 1;
    },
  };

  deleteInterviewBlock(invalidTaggedEvent);
  deleteInterviewBlock(validTaggedEvent);
  assert.equal(deleted, 1);

  const cal = {
    getId: () => "cal-1",
    getEventById: () => null,
  };
  assert.equal(
    interviewInOriginalSpot(cal, { getStartTime: () => new Date() }, "missing"),
    true,
  );

  const sourceEnd = new Date("2099-01-01T12:00:00Z");
  const blockStartSame = new Date("2099-01-01T12:00:00Z");
  const blockStartDifferent = new Date("2099-01-01T12:30:00Z");

  const interviewEvent = {
    getId: () => "abc123@google.com",
    getEndTime: () => sourceEnd,
  };

  const calCancelled = {
    getId: () => "cal-1",
    getEventById: () => interviewEvent,
  };
  ctx.Calendar.Events.get = () => ({ status: "cancelled" });
  assert.equal(
    interviewInOriginalSpot(
      calCancelled,
      { getStartTime: () => blockStartSame },
      "id",
    ),
    false,
  );

  const calMoved = {
    getId: () => "cal-1",
    getEventById: () => interviewEvent,
  };
  ctx.Calendar.Events.get = () => ({ status: "confirmed" });
  assert.equal(
    interviewInOriginalSpot(
      calMoved,
      { getStartTime: () => blockStartDifferent },
      "id",
    ),
    false,
  );
  assert.equal(
    interviewInOriginalSpot(
      calMoved,
      { getStartTime: () => blockStartSame },
      "id",
    ),
    true,
  );
});

test("colorBasedOnAttendees helper rules and priority behave correctly", () => {
  const ctx = loadScript("work/colorBasedOnAttendees.gs", {
    GetColorEventOrgNames: () => ["eng-team@company.com"],
    GetColorEventVips: () => ["ceo@company.com"],
  });

  assert.equal(
    ctx.checkNonOrg("company.com", "partner@external.com", "me@company.com"),
    true,
  );
  assert.equal(
    ctx.checkNonOrg("company.com", "teammate@company.com", "me@company.com"),
    false,
  );
  assert.equal(
    ctx.checkNonOrg(
      "company.com",
      "teammate@company.com",
      "organizer@external.com",
    ),
    true,
  );

  assert.equal(
    ctx.checkVIPs(["ceo@company.com"], "ceo@company.com", "me@company.com"),
    true,
  );
  assert.equal(
    ctx.checkVIPs(
      ["ceo@company.com"],
      "teammate@company.com",
      "ceo@company.com",
    ),
    true,
  );
  assert.equal(
    ctx.checkVIPs(
      ["ceo@company.com"],
      "teammate@company.com",
      "me@company.com",
    ),
    false,
  );

  assert.equal(
    ctx.checkOrgMeetings(
      ["eng-team@company.com"],
      "eng-team@company.com,eileen@company.com",
    ),
    false,
  );
  assert.equal(
    ctx.checkOrgMeetings(["eng-team@company.com"], "eng-team@company.com"),
    true,
  );

  const ColorEventStatus = vm.runInContext("ColorEventStatus", ctx);
  const ColorEventColors = vm.runInContext("ColorEventColors", ctx);
  assert.equal(
    ctx.setPriorityColor(
      ColorEventStatus.vip | ColorEventStatus.myorg | ColorEventStatus.external,
    ),
    ColorEventColors.external,
  );
  assert.equal(
    ctx.setPriorityColor(ColorEventStatus.myorg | ColorEventStatus.vip),
    ColorEventColors.org,
  );
  assert.equal(
    ctx.setPriorityColor(ColorEventStatus.vip),
    ColorEventColors.vip,
  );
  assert.equal(ctx.setPriorityColor(0), ColorEventColors.internal);

  let setColorCalls = 0;
  const eventSame = {
    getColor: () => "RED",
    setColor: () => {
      setColorCalls += 1;
    },
  };
  const eventDifferent = {
    getColor: () => "BLUE",
    setColor: () => {
      setColorCalls += 1;
    },
  };

  ctx.updateEventColor(eventSame, "RED");
  ctx.updateEventColor(eventDifferent, "GREEN");
  assert.equal(setColorCalls, 1);
});

test("blockFromPersonal calendarEventTag is deterministic and prefixed", () => {
  const ctx = loadScript("work/blockFromPersonalCalendar.gs");
  const calendarEventTag = vm.runInContext("calendarEventTag", ctx);
  const tagA1 = calendarEventTag("cal-a@example.com");
  const tagA2 = calendarEventTag("cal-a@example.com");
  const tagB = calendarEventTag("cal-b@example.com");
  const generatedBy = vm.runInContext("GENERATED_BY_DESCRIPTION", ctx);

  assert.equal(
    generatedBy,
    "Generated with github.com/streeter/google-apps-scripts",
  );
  assert.equal(tagA1, tagA2);
  assert.notEqual(tagA1, tagB);
  assert.equal(tagA1.startsWith("blockFromPersonal."), true);
  assert.equal(tagA1.endsWith(".originalId"), true);
  assert.ok(tagA1.length <= 44);
});

function makeCalendarEvent(title, description, deletedIds, id) {
  return {
    getTitle: () => title,
    getDescription: () => description,
    getStartTime: () => new Date("2026-01-01T10:00:00Z"),
    getEndTime: () => new Date("2026-01-01T11:00:00Z"),
    deleteEvent: () => deletedIds.push(id),
  };
}

function makeInterviewEvent({
  title = "",
  description = "",
  guests = [],
} = {}) {
  return {
    getTitle: () => title,
    getDescription: () => description,
    getGuestList: () => guests,
  };
}

function makeAdminDirectory(pages) {
  const calls = [];
  return {
    calls,
    AdminDirectory: {
      Resources: {
        Calendars: {
          list: (customer, options) => {
            calls.push({ customer, pageToken: options.pageToken });
            const index = options.pageToken ? Number(options.pageToken) : 0;
            const nextIndex = index + 1;
            return {
              items: pages[index],
              nextPageToken:
                nextIndex < pages.length ? String(nextIndex) : undefined,
            };
          },
        },
      },
    },
  };
}

function makeCacheService() {
  const store = {};
  const puts = [];
  return {
    store,
    puts,
    CacheService: {
      getUserCache: () => ({
        get: (key) => (key in store ? store[key] : null),
        put: (key, value, seconds) => {
          puts.push({ key, seconds });
          store[key] = value;
        },
      }),
    },
  };
}

function makeRoomResource(name, buildingId, capacity) {
  return {
    resourceName: name,
    resourceEmail: `room-${name.toLowerCase()}`,
    buildingId,
    capacity,
  };
}

test("autoInterviewRooms listRooms filters by office and capacity, and sorts", () => {
  const admin = makeAdminDirectory([
    [
      makeRoomResource("Zebra", "SSF-1", 4),
      makeRoomResource("Big", "SSF-1", 12),
      makeRoomResource("Virtual", "SSF-1", 0),
      makeRoomResource("Dublin", "DUB-2", 2),
    ],
    [
      makeRoomResource("Apple", "SSF-2", 4),
      makeRoomResource("Phone", "SSF-1", 1),
      {
        resourceName: "NoBuilding",
        resourceEmail: "room-nobuilding",
        capacity: 2,
      },
    ],
  ]);
  const ctx = loadScript("work/autoInterviewRooms.gs", {
    AdminDirectory: admin.AdminDirectory,
  });

  const rooms = ctx.listRooms("SSF");

  assert.deepEqual(JSON.parse(JSON.stringify(rooms)), [
    {
      id: "room-phone",
      name: "Phone",
      capacity: 1,
    },
    {
      id: "room-apple",
      name: "Apple",
      capacity: 4,
    },
    {
      id: "room-zebra",
      name: "Zebra",
      capacity: 4,
    },
  ]);
  assert.deepEqual(
    admin.calls.map((call) => call.pageToken),
    [null, "1"],
  );
});

test("autoInterviewRooms getCachedRooms caches non-empty room lists", () => {
  const admin = makeAdminDirectory([[makeRoomResource("Phone", "SSF-1", 1)]]);
  const cache = makeCacheService();
  const ctx = loadScript("work/autoInterviewRooms.gs", {
    AdminDirectory: admin.AdminDirectory,
    CacheService: cache.CacheService,
  });

  const first = ctx.getCachedRooms("SSF");
  const second = ctx.getCachedRooms("SSF");

  assert.equal(admin.calls.length, 1);
  assert.deepEqual(
    JSON.parse(JSON.stringify(second)),
    JSON.parse(JSON.stringify(first)),
  );
  assert.deepEqual(cache.puts, [
    { key: "auto-interview-rooms:SSF", seconds: 6 * 60 * 60 },
  ]);
});

test("autoInterviewRooms getCachedRooms doesn't cache an empty room list", () => {
  const admin = makeAdminDirectory([[makeRoomResource("Dublin", "DUB-1", 1)]]);
  const cache = makeCacheService();
  const ctx = loadScript("work/autoInterviewRooms.gs", {
    AdminDirectory: admin.AdminDirectory,
    CacheService: cache.CacheService,
  });

  assert.equal(ctx.getCachedRooms("SSF").length, 0);
  assert.equal(ctx.getCachedRooms("SSF").length, 0);

  assert.equal(admin.calls.length, 2);
  assert.deepEqual(cache.puts, []);
});

test("autoInterviewRooms main books the first free room in OFFICE_ID", () => {
  const admin = makeAdminDirectory([
    [
      makeRoomResource("Busy", "SSF-1", 1),
      makeRoomResource("Free", "SSF-1", 2),
      makeRoomResource("Dublin", "DUB-1", 1),
    ],
  ]);
  const cache = makeCacheService();
  const start = new Date("2026-11-21T18:00:00Z");
  const end = new Date("2026-11-21T19:00:00Z");
  const interview = {
    getTitle: () => "Interview with Candidate",
    getDescription: () => "Thanks for interviewing with us",
    getStartTime: () => start,
    getEndTime: () => end,
  };
  const freebusyQueries = [];
  const inserted = [];
  const ctx = loadScript("work/autoInterviewRooms.gs", {
    AdminDirectory: admin.AdminDirectory,
    CacheService: cache.CacheService,
    CalendarApp: {
      getDefaultCalendar: () => ({
        getEvents: () => [interview],
      }),
    },
    Calendar: {
      Freebusy: {
        query: (request) => {
          const id = request.items[0].id;
          freebusyQueries.push(id);
          const busy = id === "room-busy" ? [{ start: "x" }] : [];
          return { calendars: { [id]: { busy } } };
        },
      },
      Events: {
        insert: (event, calendarId) => {
          inserted.push({ event, calendarId });
          return event;
        },
      },
    },
  });

  const summary = ctx.main();

  assert.deepEqual(freebusyQueries, ["room-busy", "room-free"]);
  assert.equal(inserted.length, 1);
  assert.equal(inserted[0].calendarId, "primary");
  assert.deepEqual(JSON.parse(JSON.stringify(inserted[0].event.attendees)), [
    { email: "room-free", resource: true },
  ]);
  assert.equal(inserted[0].event.start.dateTime, "2026-11-21T17:45:00.000Z");
  assert.equal(inserted[0].event.end.dateTime, "2026-11-21T19:15:00.000Z");
  assert.match(summary, /Created: 1/);
  assert.match(summary, /No room: 0/);
});
