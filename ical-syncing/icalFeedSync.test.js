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
  );

  const p = resource.extendedProperties.private;
  assert.equal(p.managedKind, "drive");
  assert.equal(p.sourceSyncKey, "feedhash123:source-sync");
  assert.equal(p.sourceEventId, "source-event-123");
  assert.equal(p.syncKey, "drive:feedhash123:source-sync");
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
    defaultAttendeeEmails: [],
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
  assert.equal(
    drive.extendedProperties.private.sourceEventId,
    "source-created-1",
  );
  assert.equal(
    drive.extendedProperties.private.syncKey,
    ctx.buildDriveSyncKey_(drive.extendedProperties.private.sourceSyncKey),
  );
});
