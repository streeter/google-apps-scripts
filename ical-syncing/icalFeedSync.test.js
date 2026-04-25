const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

function loadIcalSyncContext() {
  const scriptProperties = new Map();
  const triggerState = {
    method: null,
    value: null,
    created: false,
  };
  const context = {
    console: {
      log: () => {},
      info: () => {},
      error: () => {},
    },
    JSON,
    Date,
    Math,
    Logger: { log: () => {} },
    Session: {
      getScriptTimeZone: () => "America/New_York",
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) =>
          scriptProperties.has(key) ? scriptProperties.get(key) : null,
        setProperty: (key, value) => scriptProperties.set(key, String(value)),
      }),
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: "SHA_256" },
      Charset: { UTF_8: "UTF_8" },
      computeDigest: (_algorithm, input) => {
        const hash = crypto
          .createHash("sha256")
          .update(String(input), "utf8")
          .digest();
        return Array.from(hash.values());
      },
    },
    ScriptApp: {
      getProjectTriggers: () => [],
      deleteTrigger: () => {},
      newTrigger: () => ({
        timeBased: () => ({
          everyMinutes: () => ({
            create: () => {},
          }),
          everyHours: () => ({
            create: () => {},
          }),
          everyDays: () => ({
            create: () => {},
          }),
        }),
      }),
    },
    CalendarApp: {
      getAllCalendars: () => [],
    },
    Calendar: {
      Events: {
        insert: () => {
          throw new Error("Calendar.Events.insert mock not set");
        },
        patch: () => {
          throw new Error("Calendar.Events.patch mock not set");
        },
        remove: () => {
          throw new Error("Calendar.Events.remove mock not set");
        },
        list: () => ({ items: [] }),
      },
    },
    Maps: {
      DirectionFinder: {
        Mode: {
          DRIVING: "DRIVING",
        },
      },
      newDirectionFinder: () => ({
        setOrigin() {
          return this;
        },
        setDestination() {
          return this;
        },
        setMode() {
          return this;
        },
        getDirections() {
          throw new Error(
            "Maps.newDirectionFinder().getDirections() mock not set",
          );
        },
      }),
    },
    UrlFetchApp: {
      fetch: () => {
        throw new Error("UrlFetchApp.fetch mock not set");
      },
    },
    __triggerState: triggerState,
  };

  context.ScriptApp.newTrigger = () => ({
    timeBased: () => ({
      everyMinutes: (n) => ({
        create: () => {
          triggerState.method = "everyMinutes";
          triggerState.value = n;
          triggerState.created = true;
        },
      }),
      everyHours: (n) => ({
        create: () => {
          triggerState.method = "everyHours";
          triggerState.value = n;
          triggerState.created = true;
        },
      }),
      everyDays: (n) => ({
        create: () => {
          triggerState.method = "everyDays";
          triggerState.value = n;
          triggerState.created = true;
        },
      }),
    }),
  });

  vm.createContext(context);
  const scriptPath = path.resolve(__dirname, "../ical-syncing/icalFeedSync.gs");
  const code = fs.readFileSync(scriptPath, "utf8");
  vm.runInContext(code, context, { filename: scriptPath });
  return context;
}

function baseStats() {
  return {
    driveCreated: 0,
    driveUpdated: 0,
    driveDeleted: 0,
    driveSkipped: 0,
    arrivalCreated: 0,
    arrivalUpdated: 0,
    arrivalDeleted: 0,
    arrivalSkipped: 0,
  };
}

test("getIcalSyncConfig_ defaults minDriveMinutesToCreate to 10", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      { feedUrl: "https://example.com/a.ics", calendarId: "cal1" },
    ],
  });

  const cfg = ctx.getIcalSyncConfig_();

  assert.equal(cfg.minDriveMinutesToCreate, 10);
});

test("getIcalSyncConfig_ defaults per-feed titlePrefix to empty string", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      { feedUrl: "https://example.com/a.ics", calendarId: "cal1" },
    ],
  });

  const cfg = ctx.getIcalSyncConfig_();

  assert.equal(cfg.feedMappings[0].titlePrefix, "");
});

test("applyTriggerInterval_ maps minute values to minutes/hours/days", () => {
  const ctx = loadIcalSyncContext();
  const calls = [];
  const clock = {
    everyMinutes: (n) => {
      calls.push(["everyMinutes", n]);
      return { create: () => {} };
    },
    everyHours: (n) => {
      calls.push(["everyHours", n]);
      return { create: () => {} };
    },
    everyDays: (n) => {
      calls.push(["everyDays", n]);
      return { create: () => {} };
    },
  };

  ctx.applyTriggerInterval_(clock, 15);
  ctx.applyTriggerInterval_(clock, 60);
  ctx.applyTriggerInterval_(clock, 1440);

  assert.deepEqual(calls, [
    ["everyMinutes", 15],
    ["everyHours", 1],
    ["everyDays", 1],
  ]);
});

test("setupIcalFeedSyncTrigger uses hourly trigger when triggerEveryMinutes is 60", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    triggerEveryMinutes: 60,
    feedMappings: [
      { feedUrl: "https://example.com/a.ics", calendarId: "cal1" },
    ],
  });

  ctx.setupIcalFeedSyncTrigger();

  assert.equal(ctx.__triggerState.method, "everyHours");
  assert.equal(ctx.__triggerState.value, 1);
  assert.equal(ctx.__triggerState.created, true);
});

test("applyTriggerInterval_ rejects unsupported minute values", () => {
  const ctx = loadIcalSyncContext();
  const clock = {
    everyMinutes: () => ({ create: () => {} }),
    everyHours: () => ({ create: () => {} }),
    everyDays: () => ({ create: () => {} }),
  };

  assert.throws(
    () => ctx.applyTriggerInterval_(clock, 45),
    /Unsupported triggerEveryMinutes value/,
  );
});

test("unescapeIcsText_ converts escaped and double-escaped newlines", () => {
  const ctx = loadIcalSyncContext();

  assert.equal(
    ctx.unescapeIcsText_("Event Type: Practice\\nHome/Away: Home\\n\\n"),
    "Event Type: Practice\nHome/Away: Home\n\n",
  );
  assert.equal(
    ctx.unescapeIcsText_("Event Type: Practice\\\\nHome/Away: Home\\\\n\\\\n"),
    "Event Type: Practice\nHome/Away: Home\n\n",
  );
});

test("applyEventTitlePrefix_ prefixes summary and preserves original event object", () => {
  const ctx = loadIcalSyncContext();
  const evt = {
    uid: "uid-1",
    summary: "Practice",
    start: { type: "dateTime", dateTime: "2099-05-01T15:00:00Z" },
    end: { type: "dateTime", dateTime: "2099-05-01T16:00:00Z" },
  };

  const prefixed = ctx.applyEventTitlePrefix_(evt, "[Sports]");
  assert.equal(prefixed.summary, "[Sports] Practice");
  assert.equal(evt.summary, "Practice");

  const unchanged = ctx.applyEventTitlePrefix_(evt, "   ");
  assert.equal(unchanged, evt);
});

test("shouldSyncEvent_ respects cutoff date", () => {
  const ctx = loadIcalSyncContext();
  const cutoff = new Date("2026-04-25T00:00:00Z");

  const past = {
    cancelled: false,
    start: { type: "dateTime", dateTime: "2026-04-24T10:00:00Z" },
    end: { type: "dateTime", dateTime: "2026-04-24T11:00:00Z" },
  };
  const onCutoff = {
    cancelled: false,
    start: { type: "dateTime", dateTime: "2026-04-25T00:00:00Z" },
    end: { type: "dateTime", dateTime: "2026-04-25T00:15:00Z" },
  };

  assert.equal(ctx.shouldSyncEvent_(past, cutoff), false);
  assert.equal(ctx.shouldSyncEvent_(onCutoff, cutoff), true);
});

test("extractDriveMinutesFromDirections_ rounds up to nearest 15 minutes", () => {
  const ctx = loadIcalSyncContext();

  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 660 } }] }], // 11 minutes
    }),
    15,
  );
  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 300 } }] }], // 5 minutes
    }),
    15,
  );
  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 301 } }] }], // 6 minutes
    }),
    15,
  );
  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 23 * 60 } }] }],
    }),
    30,
  );
  assert.equal(ctx.roundUpMinutesToNearestFifteen_(1), 15);
  assert.equal(ctx.roundUpMinutesToNearestFifteen_(15), 15);
  assert.equal(ctx.roundUpMinutesToNearestFifteen_(23), 30);
});

test("extractArrivalLeadMinutes_ parses Arrival lead time from description", () => {
  const ctx = loadIcalSyncContext();
  assert.equal(
    ctx.extractArrivalLeadMinutes_(
      "Event Type: Practice\nArrival: 30 minutes in advance\nHome/Away: Home",
    ),
    30,
  );
  assert.equal(
    ctx.extractArrivalLeadMinutes_("Arrival: 1 minute in advance"),
    1,
  );
  assert.equal(ctx.extractArrivalLeadMinutes_("No arrival guidance"), null);
});

test("resolveDrivePlan_ prefers previous event location and skips short/no-op drives", () => {
  const ctx = loadIcalSyncContext();
  const driveOpts = {
    originAddress: "123 Main St, Brooklyn, NY",
    minDriveMinutesToCreate: 10,
  };
  const driveEnd = new Date("2099-05-01T15:00:00Z");

  ctx.findPreviousDriveOriginEvent_ = () => ({
    id: "prev-1",
    location: "Gym A",
  });
  let plan = ctx.resolveDrivePlan_(
    "calendar-1",
    "source-1",
    driveEnd,
    "Gym A",
    driveOpts,
    {},
  );
  assert.match(plan.skipReason, /already at destination/);

  ctx.findPreviousDriveOriginEvent_ = () => ({
    id: "prev-2",
    location: "Nearby Gym",
  });
  ctx.getDriveMinutes_ = (origin) => (origin === "Nearby Gym" ? 10 : 30);
  plan = ctx.resolveDrivePlan_(
    "calendar-1",
    "source-1",
    driveEnd,
    "Main Field",
    driveOpts,
    {},
  );
  assert.match(plan.skipReason, /within threshold/);

  ctx.findPreviousDriveOriginEvent_ = () => ({
    id: "prev-3",
    location: "Far Gym",
  });
  ctx.getDriveMinutes_ = (origin) => (origin === "Far Gym" ? 25 : 30);
  plan = ctx.resolveDrivePlan_(
    "calendar-1",
    "source-1",
    driveEnd,
    "Main Field",
    driveOpts,
    {},
  );
  assert.equal(plan.originAddress, "Far Gym");
  assert.equal(plan.driveMinutes, 25);
  assert.equal(plan.previousEventId, "prev-3");
});

test("resolveDrivePlan_ falls back to default origin when previous route lookup fails", () => {
  const ctx = loadIcalSyncContext();
  const driveOpts = {
    originAddress: "123 Main St, Brooklyn, NY",
    minDriveMinutesToCreate: 10,
  };
  const driveEnd = new Date("2099-05-01T15:00:00Z");

  ctx.findPreviousDriveOriginEvent_ = () => ({
    id: "prev-1",
    location: "Unroutable",
  });
  ctx.getDriveMinutes_ = (origin) => (origin === "Unroutable" ? null : 30);

  const plan = ctx.resolveDrivePlan_(
    "calendar-1",
    "source-1",
    driveEnd,
    "Main Field",
    driveOpts,
    {},
  );

  assert.equal(plan.originAddress, "123 Main St, Brooklyn, NY");
  assert.equal(plan.driveMinutes, 30);
  assert.equal(plan.previousEventId, "");
});

test("reconcileDrivePlaceholder_ creates placeholder only when drive time is > threshold", () => {
  const ctx = loadIcalSyncContext();
  const inserted = [];
  ctx.Calendar.Events.insert = (_resource, calendarId) => {
    assert.equal(calendarId, "calendar-1");
    inserted.push(_resource);
    return {
      id: "drive-1",
      start: _resource.start,
      end: _resource.end,
      extendedProperties: _resource.extendedProperties,
    };
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = () => {};

  const evt = {
    uid: "uid-1",
    summary: "Office Meeting",
    location: "1 Infinite Loop, Cupertino, CA",
    start: { type: "dateTime", dateTime: "2099-05-01T15:00:00Z" },
    end: { type: "dateTime", dateTime: "2099-05-01T16:00:00Z" },
  };
  const syncedEvent = {
    id: "source-1",
    start: { dateTime: "2099-05-01T15:00:00Z" },
  };
  const mapping = {
    feedUrl: "https://example.com/feed.ics",
    calendarId: "calendar-1",
  };
  const feedHash = "feedhash123";
  const sourceSyncKey = "feedhash123:abc";
  const driveSyncKey = ctx.buildDriveSyncKey_(sourceSyncKey);
  const driveOpts = {
    enabled: true,
    originAddress: "Brooklyn, NY",
    minDriveMinutesToCreate: 10,
    titleTemplate: "Drive to {{title}} ({{minutes}}m)",
  };
  const existingDriveByKey = {};
  const seenDrive = {};
  const today = new Date("2026-01-01T00:00:00Z");

  ctx.getDriveMinutes_ = () => 10;
  const statsAtThreshold = baseStats();
  ctx.reconcileDrivePlaceholder_(
    evt,
    syncedEvent,
    mapping,
    feedHash,
    sourceSyncKey,
    driveSyncKey,
    driveOpts,
    existingDriveByKey,
    seenDrive,
    today,
    statsAtThreshold,
    {},
    null,
    [],
  );
  assert.equal(statsAtThreshold.driveCreated, 0);
  assert.equal(statsAtThreshold.driveSkipped, 1);
  assert.equal(inserted.length, 0);

  ctx.getDriveMinutes_ = () => 11;
  const statsAboveThreshold = baseStats();
  ctx.reconcileDrivePlaceholder_(
    evt,
    syncedEvent,
    mapping,
    feedHash,
    sourceSyncKey,
    driveSyncKey,
    driveOpts,
    existingDriveByKey,
    {},
    today,
    statsAboveThreshold,
    {},
    null,
    [],
  );
  assert.equal(statsAboveThreshold.driveCreated, 1);
  assert.equal(inserted.length, 1);
});

test("drive placeholder resource carries source linkage metadata", () => {
  const ctx = loadIcalSyncContext();
  const driveStart = new Date("2099-05-01T14:35:00Z");
  const driveEnd = new Date("2099-05-01T15:00:00Z");
  const resource = ctx.buildDrivePlaceholderResource_(
    { feedUrl: "https://example.com/feed.ics" },
    "feedhash123",
    { uid: "uid-1", location: "Destination" },
    "feedhash123:source-sync",
    "drive:feedhash123:source-sync",
    "source-event-123",
    "Drive to Meeting",
    driveStart,
    driveEnd,
    "drivehash123",
    "Origin Address",
    ["a@example.com", "b@example.com"],
  );

  const p = resource.extendedProperties.private;
  assert.equal(p.managedKind, "drive");
  assert.equal(p.sourceSyncKey, "feedhash123:source-sync");
  assert.equal(p.sourceEventId, "source-event-123");
  assert.equal(p.syncKey, "drive:feedhash123:source-sync");
  assert.equal(
    JSON.stringify(resource.attendees),
    JSON.stringify([{ email: "a@example.com" }, { email: "b@example.com" }]),
  );
});

test("arrival placeholder resource carries source linkage metadata", () => {
  const ctx = loadIcalSyncContext();
  const arrivalStart = new Date("2099-05-01T15:00:00Z");
  const arrivalEnd = new Date("2099-05-01T15:30:00Z");
  const resource = ctx.buildArrivalPlaceholderResource_(
    { feedUrl: "https://example.com/feed.ics" },
    "feedhash123",
    { uid: "uid-1", location: "Destination" },
    "feedhash123:source-sync",
    "arrival:feedhash123:source-sync",
    "source-event-123",
    "Advanced arrival for Practice",
    arrivalStart,
    arrivalEnd,
    "arrivalhash123",
    30,
    ["a@example.com", "b@example.com"],
  );

  const p = resource.extendedProperties.private;
  assert.equal(p.managedKind, "arrival");
  assert.equal(p.sourceSyncKey, "feedhash123:source-sync");
  assert.equal(p.sourceEventId, "source-event-123");
  assert.equal(p.arrivalMinutes, "30");
  assert.equal(p.syncKey, "arrival:feedhash123:source-sync");
  assert.equal(
    JSON.stringify(resource.attendees),
    JSON.stringify([{ email: "a@example.com" }, { email: "b@example.com" }]),
  );
});

test("syncOneFeed_ creates source event and tied drive placeholder", () => {
  const ctx = loadIcalSyncContext();
  const inserts = [];

  ctx.fetchIcs_ = () =>
    [
      "BEGIN:VCALENDAR",
      "BEGIN:VEVENT",
      "UID:uid-1",
      "DTSTART:20990501T150000Z",
      "DTEND:20990501T160000Z",
      "SUMMARY:Client Meeting",
      "LOCATION:Seattle, WA",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\n");
  ctx.loadExistingEventsByKey_ = () => ({});
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.getDriveMinutes_ = () => 25;

  ctx.Calendar.Events.insert = (resource, calendarId) => {
    assert.equal(calendarId, "calendar-1");
    inserts.push(resource);
    if (resource.extendedProperties.private.managedKind === "source") {
      return {
        id: "source-created-1",
        start: resource.start,
        end: resource.end,
        extendedProperties: resource.extendedProperties,
      };
    }
    return {
      id: "drive-created-1",
      start: resource.start,
      end: resource.end,
      extendedProperties: resource.extendedProperties,
    };
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = () => {
    throw new Error("unexpected remove");
  };

  const cfg = {
    deleteMissingFromFeed: false,
    defaultAttendeeEmails: ["coach@example.com", "parent@example.com"],
    addDriveTimePlaceholders: false,
    defaultOriginAddress: "New York, NY",
    minDriveMinutesToCreate: 10,
    driveEventTitleTemplate: "Drive ({{minutes}}m) to {{title}}",
  };
  const mapping = {
    name: "Test Feed",
    feedUrl: "https://example.com/feed.ics",
    calendarId: "calendar-1",
    attendeeEmails: [],
    titlePrefix: "[Sports]",
    addDriveTimePlaceholders: true,
    originAddress: "",
  };

  const stats = ctx.syncOneFeed_(
    cfg,
    mapping,
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(stats.created, 1);
  assert.equal(stats.driveCreated, 1);
  assert.equal(inserts.length, 2);

  const source = inserts.find(
    (r) => r.extendedProperties.private.managedKind === "source",
  );
  const drive = inserts.find(
    (r) => r.extendedProperties.private.managedKind === "drive",
  );

  assert.ok(source);
  assert.ok(drive);
  assert.equal(source.summary, "[Sports] Client Meeting");
  assert.equal(drive.summary, "Drive (25m) to [Sports] Client Meeting");
  assert.equal(
    drive.extendedProperties.private.sourceEventId,
    "source-created-1",
  );
  assert.equal(
    drive.extendedProperties.private.syncKey,
    ctx.buildDriveSyncKey_(drive.extendedProperties.private.sourceSyncKey),
  );
});

test("syncOneFeed_ creates arrival placeholder and moves drive before arrival", () => {
  const ctx = loadIcalSyncContext();
  const inserts = [];

  ctx.fetchIcs_ = () =>
    [
      "BEGIN:VCALENDAR",
      "BEGIN:VEVENT",
      "UID:uid-1",
      "DTSTART:20990501T153000Z",
      "DTEND:20990501T163000Z",
      "SUMMARY:Soccer Game",
      "DESCRIPTION:Event Type: Practice\\nArrival: 30 minutes in advance",
      "LOCATION:Seattle, WA",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\n");
  ctx.loadExistingEventsByKey_ = () => ({});
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.getDriveMinutes_ = () => 25;

  ctx.Calendar.Events.insert = (resource, calendarId) => {
    assert.equal(calendarId, "calendar-1");
    inserts.push(resource);
    const kind = resource.extendedProperties.private.managedKind;
    if (kind === "source") {
      return {
        id: "source-created-1",
        start: resource.start,
        end: resource.end,
        extendedProperties: resource.extendedProperties,
      };
    }
    if (kind === "arrival") {
      return {
        id: "arrival-created-1",
        start: resource.start,
        end: resource.end,
        extendedProperties: resource.extendedProperties,
      };
    }
    return {
      id: "drive-created-1",
      start: resource.start,
      end: resource.end,
      extendedProperties: resource.extendedProperties,
    };
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = () => {
    throw new Error("unexpected remove");
  };

  const cfg = {
    deleteMissingFromFeed: false,
    defaultAttendeeEmails: ["coach@example.com", "parent@example.com"],
    addDriveTimePlaceholders: false,
    defaultOriginAddress: "New York, NY",
    minDriveMinutesToCreate: 10,
    driveEventTitleTemplate: "Drive ({{minutes}}m) to {{title}}",
  };
  const mapping = {
    name: "Sports Feed",
    feedUrl: "https://example.com/sports.ics",
    calendarId: "calendar-1",
    attendeeEmails: [],
    titlePrefix: "[Sports]",
    addDriveTimePlaceholders: true,
    originAddress: "",
  };

  const stats = ctx.syncOneFeed_(
    cfg,
    mapping,
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(stats.created, 1);
  assert.equal(stats.arrivalCreated, 1);
  assert.equal(stats.driveCreated, 1);
  assert.equal(inserts.length, 3);
  assert.deepEqual(
    inserts.map((r) => r.extendedProperties.private.managedKind),
    ["source", "arrival", "drive"],
  );

  const source = inserts.find(
    (r) => r.extendedProperties.private.managedKind === "source",
  );
  const arrival = inserts.find(
    (r) => r.extendedProperties.private.managedKind === "arrival",
  );
  const drive = inserts.find(
    (r) => r.extendedProperties.private.managedKind === "drive",
  );

  assert.ok(source);
  assert.ok(arrival);
  assert.ok(drive);
  assert.equal(source.summary, "[Sports] Soccer Game");
  assert.equal(arrival.summary, "Advanced arrival for [Sports] Soccer Game");
  assert.equal(arrival.start.dateTime, "2099-05-01T15:00:00.000Z");
  assert.equal(arrival.end.dateTime, "2099-05-01T15:30:00.000Z");
  assert.equal(drive.end.dateTime, "2099-05-01T15:00:00.000Z");
  assert.equal(drive.start.dateTime, "2099-05-01T14:35:00.000Z");
  assert.equal(
    JSON.stringify(arrival.attendees),
    JSON.stringify([
      { email: "coach@example.com" },
      { email: "parent@example.com" },
    ]),
  );
  assert.equal(
    JSON.stringify(drive.attendees),
    JSON.stringify([
      { email: "coach@example.com" },
      { email: "parent@example.com" },
    ]),
  );
});
