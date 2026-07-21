const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

function loadIcalSyncContext() {
  const scriptProperties = new Map();
  const lockState = {
    available: true,
    held: false,
    released: false,
  };
  const triggerState = {
    method: null,
    value: null,
    created: false,
    createdTriggers: [],
  };
  const context = {
    console: {
      log: () => {},
      info: () => {},
      warn: () => {},
      error: () => {},
    },
    JSON,
    Date,
    Math,
    Logger: { log: () => {} },
    Session: {
      getScriptTimeZone: () => "America/Los_Angeles",
      getEffectiveUser: () => ({ getEmail: () => "owner@example.com" }),
      getActiveUser: () => ({ getEmail: () => "owner@example.com" }),
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) =>
          scriptProperties.has(key) ? scriptProperties.get(key) : null,
        setProperty: (key, value) => scriptProperties.set(key, String(value)),
      }),
    },
    LockService: {
      getScriptLock: () => ({
        tryLock: () => {
          if (!lockState.available || lockState.held) return false;
          lockState.held = true;
          return true;
        },
        releaseLock: () => {
          lockState.held = false;
          lockState.released = true;
        },
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
      formatDate: (date, timeZone, pattern) => {
        if (pattern === "Z") {
          const parts = new Intl.DateTimeFormat("en-US", {
            timeZone,
            timeZoneName: "longOffset",
          }).formatToParts(date);
          const offset =
            parts.find((part) => part.type === "timeZoneName")?.value ||
            "GMT+00:00";
          const match = offset.match(/^GMT([+-])(\d{2}):(\d{2})$/);
          if (!match) return "+0000";
          return match[1] + match[2] + match[3];
        }
        if (pattern === "yyyy-MM-dd'T'HH:mm:ss") {
          const parts = new Intl.DateTimeFormat("en-CA", {
            timeZone,
            year: "numeric",
            month: "2-digit",
            day: "2-digit",
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit",
            hourCycle: "h23",
          }).formatToParts(date);
          const values = {};
          parts.forEach((part) => {
            if (part.type !== "literal") values[part.type] = part.value;
          });
          return (
            values.year +
            "-" +
            values.month +
            "-" +
            values.day +
            "T" +
            values.hour +
            ":" +
            values.minute +
            ":" +
            values.second
          );
        }
        throw new Error("Unexpected Utilities.formatDate pattern: " + pattern);
      },
      sleep: () => {},
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
          atHour: () => ({
            nearMinute: () => ({
              everyDays: () => ({
                create: () => {},
              }),
            }),
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
        get: () => {
          throw new Error("Calendar.Events.get mock not set");
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
    __lockState: lockState,
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
      atHour: (hour) => ({
        nearMinute: (minute) => ({
          everyDays: (n) => ({
            create: () => {
              triggerState.method = "scheduledHours";
              triggerState.value = { hour, minute, everyDays: n };
              triggerState.created = true;
              triggerState.createdTriggers.push({
                hour,
                minute,
                everyDays: n,
              });
            },
          }),
        }),
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
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal1",
      },
    ],
  });

  const cfg = ctx.getIcalSyncConfig_();

  assert.equal(cfg.minDriveMinutesToCreate, 10);
});

test("getIcalSyncConfig_ defaults per-feed titlePrefix to empty string", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal1",
      },
    ],
  });

  const cfg = ctx.getIcalSyncConfig_();

  assert.equal(cfg.feedMappings[0].titlePrefix, "");
});

test("getIcalSyncConfig_ defaults per-feed skipAllDayEvents to false", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal1",
      },
    ],
  });

  const cfg = ctx.getIcalSyncConfig_();

  assert.equal(cfg.feedMappings[0].skipAllDayEvents, false);
});

test("getIcalSyncConfig_ defaults destination calendar attendee to true", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal1",
      },
      {
        name: "Feed B",
        feedUrl: "https://example.com/b.ics",
        calendarId: "cal2",
        addDestinationCalendarAsAttendee: false,
      },
    ],
  });

  const cfg = ctx.getIcalSyncConfig_();

  assert.equal(cfg.feedMappings[0].addDestinationCalendarAsAttendee, true);
  assert.equal(cfg.feedMappings[1].addDestinationCalendarAsAttendee, false);
});

test("getIcalSyncConfig_ defaults per-feed timeZone to empty", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal1",
      },
    ],
  });

  const cfg = ctx.getIcalSyncConfig_();

  assert.equal(cfg.feedMappings[0].timeZone, "");
});

test("getIcalSyncConfig_ leaves triggerHours unset for interval scheduling", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal1",
      },
    ],
  });

  const cfg = ctx.getIcalSyncConfig_();

  assert.equal(cfg.triggerHours, undefined);
});

test("getIcalSyncConfig_ requires each feed mapping to have a name", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      { feedUrl: "https://example.com/a.ics", calendarId: "cal1" },
    ],
  });

  assert.throws(
    () => ctx.getIcalSyncConfig_(),
    /feedMappings\[0\] missing name/,
  );
});

test("getIcalSyncConfig_ rejects duplicate feed mapping names", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal1",
      },
      {
        name: " Feed A ",
        feedUrl: "https://example.com/b.ics",
        calendarId: "cal2",
      },
    ],
  });

  assert.throws(
    () => ctx.getIcalSyncConfig_(),
    /feedMappings\[1\] name "Feed A" duplicates feedMappings\[0\]\.name; feed mapping names must be unique/,
  );
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

test("normalizeTriggerHours_ sorts and de-duplicates hours", () => {
  const ctx = loadIcalSyncContext();

  const hours = ctx.normalizeTriggerHours_([22, 6, 8, 6, 20]);

  assert.equal(JSON.stringify(hours), JSON.stringify([6, 8, 20, 22]));
});

test("normalizeTriggerHours_ allows an empty array", () => {
  const ctx = loadIcalSyncContext();

  assert.equal(JSON.stringify(ctx.normalizeTriggerHours_([])), "[]");
});

test("setupIcalFeedSyncTrigger uses hourly trigger when triggerEveryMinutes is 60", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    triggerEveryMinutes: 60,
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal1",
      },
    ],
  });

  ctx.setupIcalFeedSyncTrigger();

  assert.equal(ctx.__triggerState.method, "everyHours");
  assert.equal(ctx.__triggerState.value, 1);
  assert.equal(ctx.__triggerState.created, true);
});

test("setupIcalFeedSyncTrigger uses explicit scheduled hours when configured", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    triggerHours: [22, 6, 8, 10, 12, 14, 16, 18, 20],
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal1",
      },
    ],
  });

  ctx.setupIcalFeedSyncTrigger();

  assert.equal(ctx.__triggerState.method, "scheduledHours");
  assert.deepEqual(ctx.__triggerState.createdTriggers, [
    { hour: 6, minute: 0, everyDays: 1 },
    { hour: 8, minute: 0, everyDays: 1 },
    { hour: 10, minute: 0, everyDays: 1 },
    { hour: 12, minute: 0, everyDays: 1 },
    { hour: 14, minute: 0, everyDays: 1 },
    { hour: 16, minute: 0, everyDays: 1 },
    { hour: 18, minute: 0, everyDays: 1 },
    { hour: 20, minute: 0, everyDays: 1 },
    { hour: 22, minute: 0, everyDays: 1 },
  ]);
});

test("setupIcalFeedSyncTrigger removes triggers for explicit empty triggerHours", () => {
  const ctx = loadIcalSyncContext();
  const syncTrigger = { getHandlerFunction: () => "syncIcalFeeds" };
  const otherTrigger = { getHandlerFunction: () => "anotherFunction" };
  const deleted = [];
  ctx.getIcalSyncConfig = () => ({
    triggerHours: [],
    triggerEveryMinutes: 60,
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal1",
      },
    ],
  });
  ctx.ScriptApp.getProjectTriggers = () => [syncTrigger, otherTrigger];
  ctx.ScriptApp.deleteTrigger = (trigger) => deleted.push(trigger);
  ctx.ScriptApp.newTrigger = () => {
    throw new Error("should not create a replacement trigger");
  };

  ctx.setupIcalFeedSyncTrigger();

  assert.deepEqual(deleted, [syncTrigger]);
});

test("syncIcalFeeds processes all feeds and throws a summary error at the end", () => {
  const ctx = loadIcalSyncContext();
  const logs = [];
  const errors = [];
  const syncedFeeds = [];
  ctx.console.error = (msg) => errors.push(String(msg));
  ctx.Logger.log = (msg) => logs.push(String(msg));
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal-a",
      },
      {
        name: "Feed B",
        feedUrl: "https://example.com/b.ics",
        calendarId: "cal-b",
      },
      {
        name: "Feed C",
        feedUrl: "https://example.com/c.ics",
        calendarId: "cal-c",
      },
    ],
  });
  ctx.syncOneFeed_ = (_cfg, mapping) => {
    syncedFeeds.push(mapping.name);
    if (mapping.name === "Feed B") {
      throw new Error("B failed");
    }
    return { feed: mapping.name, ok: true };
  };

  assert.throws(
    () => ctx.syncIcalFeeds(),
    /syncIcalFeeds completed with 1 error\(s\): Feed B: Error: B failed/,
  );
  assert.deepEqual(syncedFeeds, ["Feed A", "Feed B", "Feed C"]);
  assert.equal(errors.length, 1);
  assert.match(
    errors[0],
    /\[ERROR\] Failed syncing feed Feed B: Error: B failed/,
  );
  assert.equal(logs.length, 1);
  assert.match(logs[0], /"Feed A"/);
  assert.match(logs[0], /"Feed B"/);
  assert.match(logs[0], /"Feed C"/);
});

test("syncIcalFeeds skips an overlapping execution", () => {
  const ctx = loadIcalSyncContext();
  let configCalls = 0;
  const warnings = [];
  ctx.__lockState.available = false;
  ctx.console.warn = (message) => warnings.push(String(message));
  ctx.getIcalSyncConfig = () => {
    configCalls++;
    return {
      feedMappings: [
        {
          name: "Feed A",
          feedUrl: "https://example.com/a.ics",
          calendarId: "cal-a",
        },
      ],
    };
  };

  const result = ctx.syncIcalFeeds();

  assert.equal(result.length, 0);
  assert.equal(configCalls, 0);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /already in progress/);
});

test("syncIcalFeeds releases its script lock after a failure", () => {
  const ctx = loadIcalSyncContext();
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      {
        name: "Feed A",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal-a",
      },
    ],
  });
  ctx.syncOneFeed_ = () => {
    throw new Error("sync failed");
  };

  assert.throws(() => ctx.syncIcalFeeds(), /sync failed/);
  assert.equal(ctx.__lockState.held, false);
  assert.equal(ctx.__lockState.released, true);
});

test("syncIcalFeeds aborts remaining work after a Calendar usage limit", () => {
  const ctx = loadIcalSyncContext();
  const errors = [];
  const syncedFeeds = [];
  let cleanupCalls = 0;

  ctx.console.error = (msg) => errors.push(String(msg));
  ctx.getIcalSyncConfig = () => ({
    feedMappings: [
      {
        name: "Chestnutwold",
        feedUrl: "https://example.com/a.ics",
        calendarId: "cal-a",
      },
      {
        name: "Haverford School District",
        feedUrl: "https://example.com/b.ics",
        calendarId: "cal-b",
      },
    ],
  });
  ctx.syncOneFeed_ = (_cfg, mapping) => {
    syncedFeeds.push(mapping.name);
    throw new Error(
      "GoogleJsonResponseException: API call to calendar.events.insert failed with error: Calendar usage limits exceeded.",
    );
  };
  ctx.cleanupRemovedFeedEvents_ = () => {
    cleanupCalls++;
  };

  assert.throws(
    () => ctx.syncIcalFeeds(),
    /syncIcalFeeds completed with 1 error\(s\): Chestnutwold: Error: GoogleJsonResponseException: API call to calendar\.events\.insert failed with error: Calendar usage limits exceeded\./,
  );
  assert.deepEqual(syncedFeeds, ["Chestnutwold"]);
  assert.equal(cleanupCalls, 0);
  assert.equal(errors.length, 2);
  assert.match(errors[0], /\[ERROR\] Failed syncing feed Chestnutwold:/);
  assert.match(
    errors[1],
    /\[SYNC_ABORT\] Calendar usage limit reached; skipping 1 remaining feed/,
  );
});

test("syncIcalFeeds deletes future managed events from removed feed mappings", () => {
  const ctx = loadIcalSyncContext();
  const removed = [];
  const syncedFeeds = [];
  const activeFeedUrl = "https://example.com/active.ics";
  const removedFeedUrl = "https://example.com/removed.ics";
  const activeFeedHash = ctx.sha256Hex_(activeFeedUrl).slice(0, 16);
  const removedFeedHash = ctx.sha256Hex_(removedFeedUrl).slice(0, 16);
  const activeSourceSyncKey = ctx.buildSyncKey_(
    activeFeedHash,
    "active-uid",
    "",
  );
  const removedSourceSyncKey = ctx.buildSyncKey_(
    removedFeedHash,
    "removed-uid",
    "",
  );
  const removedArrivalSyncKey = ctx.buildArrivalSyncKey_(removedSourceSyncKey);
  const removedDriveSyncKey = ctx.buildDriveSyncKey_(removedSourceSyncKey);

  function managedEvent(id, sourceUrl, sourceFeed, syncKey, kind, start, end) {
    const privateProps = {
      managedKind: kind,
      sourceFeed: sourceFeed,
      sourceUrl: sourceUrl,
      sourceUid: id + "-uid",
      syncKey: syncKey,
    };
    if (kind === "arrival" || kind === "drive") {
      privateProps.sourceSyncKey = removedSourceSyncKey;
      privateProps.sourceEventId = "removed-source";
    }
    return {
      id: id,
      summary: id,
      start: { dateTime: start },
      end: { dateTime: end },
      extendedProperties: { private: privateProps },
    };
  }

  const calendarEvents = [
    managedEvent(
      "removed-source",
      removedFeedUrl,
      removedFeedHash,
      removedSourceSyncKey,
      "source",
      "2099-05-01T15:30:00Z",
      "2099-05-01T16:30:00Z",
    ),
    managedEvent(
      "removed-arrival",
      removedFeedUrl,
      removedFeedHash,
      removedArrivalSyncKey,
      "arrival",
      "2099-05-01T15:00:00Z",
      "2099-05-01T15:30:00Z",
    ),
    managedEvent(
      "removed-drive",
      removedFeedUrl,
      removedFeedHash,
      removedDriveSyncKey,
      "drive",
      "2099-05-01T14:35:00Z",
      "2099-05-01T15:00:00Z",
    ),
    managedEvent(
      "removed-past-source",
      removedFeedUrl,
      removedFeedHash,
      ctx.buildSyncKey_(removedFeedHash, "removed-past-uid", ""),
      "source",
      "2000-05-01T15:30:00Z",
      "2000-05-01T16:30:00Z",
    ),
    managedEvent(
      "active-source",
      activeFeedUrl,
      activeFeedHash,
      activeSourceSyncKey,
      "source",
      "2099-05-01T17:00:00Z",
      "2099-05-01T18:00:00Z",
    ),
    {
      id: "manual-event",
      summary: "Manual Event",
      start: { dateTime: "2099-05-01T19:00:00Z" },
      end: { dateTime: "2099-05-01T20:00:00Z" },
    },
  ];

  ctx.getIcalSyncConfig = () => ({
    deleteMissingFromFeed: true,
    feedMappings: [
      {
        name: "Active Feed",
        feedUrl: activeFeedUrl,
        calendarId: "calendar-1",
      },
    ],
  });
  ctx.syncOneFeed_ = (_cfg, mapping) => {
    syncedFeeds.push(mapping.feedUrl);
    return { feed: mapping.name, ok: true };
  };
  ctx.Calendar.Events.list = (calendarId, opts) => {
    assert.equal(calendarId, "calendar-1");
    const filters = (opts && opts.privateExtendedProperty) || [];
    return {
      items: calendarEvents.filter((event) =>
        filters.every((filter) => {
          const [key, value] = String(filter).split("=");
          const privateProps = (event.extendedProperties || {}).private || {};
          return String(privateProps[key] || "") === value;
        }),
      ),
    };
  };
  ctx.Calendar.Events.insert = () => {
    throw new Error("unexpected insert");
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = (calendarId, eventId) => {
    assert.equal(calendarId, "calendar-1");
    removed.push(eventId);
  };

  ctx.syncIcalFeeds();

  assert.deepEqual(syncedFeeds, [activeFeedUrl]);
  assert.deepEqual(removed.sort(), [
    "removed-arrival",
    "removed-drive",
    "removed-source",
  ]);
});

test("syncIcalFeeds deletes removed-feed events from remembered calendars", () => {
  const ctx = loadIcalSyncContext();
  const removed = [];
  const activeFeedUrl = "https://example.com/active.ics";
  const removedFeedUrl = "https://example.com/removed.ics";
  const removedFeedHash = ctx.sha256Hex_(removedFeedUrl).slice(0, 16);
  const removedSourceSyncKey = ctx.buildSyncKey_(
    removedFeedHash,
    "remembered-removed-uid",
    "",
  );
  const rememberedEvent = {
    id: "remembered-source",
    summary: "Remembered Removed Event",
    start: { dateTime: "2099-05-01T15:30:00Z" },
    end: { dateTime: "2099-05-01T16:30:00Z" },
    extendedProperties: {
      private: {
        managedKind: "source",
        sourceFeed: removedFeedHash,
        sourceUrl: removedFeedUrl,
        sourceUid: "remembered-removed-uid",
        syncKey: removedSourceSyncKey,
      },
    },
  };

  ctx.PropertiesService.getScriptProperties().setProperty(
    "icalSync.managedCalendarIds",
    JSON.stringify(["removed-calendar"]),
  );
  ctx.getIcalSyncConfig = () => ({
    deleteMissingFromFeed: true,
    feedMappings: [
      {
        name: "Active Feed",
        feedUrl: activeFeedUrl,
        calendarId: "active-calendar",
      },
    ],
  });
  ctx.syncOneFeed_ = (_cfg, mapping) => {
    return { feed: mapping.name, ok: true };
  };
  ctx.Calendar.Events.list = (calendarId, opts) => {
    const filters = (opts && opts.privateExtendedProperty) || [];
    if (
      calendarId !== "removed-calendar" ||
      !filters.includes("managedKind=source")
    ) {
      return { items: [] };
    }
    return { items: [rememberedEvent] };
  };
  ctx.Calendar.Events.insert = () => {
    throw new Error("unexpected insert");
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = (calendarId, eventId) => {
    removed.push(calendarId + ":" + eventId);
  };

  ctx.syncIcalFeeds();

  assert.deepEqual(removed, ["removed-calendar:remembered-source"]);
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

test("normalizeTriggerHours_ rejects invalid hours", () => {
  const ctx = loadIcalSyncContext();

  assert.throws(
    () => ctx.normalizeTriggerHours_([6, 24]),
    /triggerHours values must be integers from 0 through 23/,
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

test("extractDriveMinutesFromDirections_ applies drive placeholder rounding rules", () => {
  const ctx = loadIcalSyncContext();

  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 9 * 60 } }] }],
    }),
    0,
  );
  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 10 * 60 } }] }],
    }),
    15,
  );
  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 15 * 60 } }] }],
    }),
    15,
  );
  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 16 * 60 } }] }],
    }),
    20,
  );
  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 20 * 60 } }] }],
    }),
    20,
  );
  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 21 * 60 } }] }],
    }),
    30,
  );
  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 29 * 60 } }] }],
    }),
    30,
  );
  assert.equal(
    ctx.extractDriveMinutesFromDirections_({
      routes: [{ legs: [{ duration: { value: 31 * 60 } }] }],
    }),
    40,
  );
  assert.equal(ctx.roundDriveMinutesForPlaceholder_(1), 0);
  assert.equal(ctx.roundDriveMinutesForPlaceholder_(10), 15);
  assert.equal(ctx.roundDriveMinutesForPlaceholder_(15), 15);
  assert.equal(ctx.roundDriveMinutesForPlaceholder_(16), 20);
  assert.equal(ctx.roundDriveMinutesForPlaceholder_(20), 20);
  assert.equal(ctx.roundDriveMinutesForPlaceholder_(23), 30);
  assert.equal(ctx.roundDriveMinutesForPlaceholder_(31), 40);
});

test("getDriveMinutes_ logs route diagnostics when directions are incomplete", () => {
  const ctx = loadIcalSyncContext();
  const warnings = [];
  const errors = [];
  ctx.console.warn = (msg) => warnings.push(String(msg));
  ctx.console.error = (msg) => errors.push(String(msg));
  ctx.Maps.newDirectionFinder = () => ({
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
      return {
        status: "OK",
        geocoded_waypoints: [
          { geocoder_status: "OK", partial_match: true, place_id: "abc" },
        ],
        routes: [{ legs: [{}] }],
      };
    },
  });

  const minutes = ctx.getDriveMinutes_("Origin", "Destination", {});

  assert.equal(minutes, null);
  assert.ok(
    warnings.some((msg) => msg.includes("returned no usable duration")),
    warnings.join("\n"),
  );
  assert.equal(errors.length, 0);
});

test("logDirectionsDiagnostics_ logs waypoint and status details", () => {
  const ctx = loadIcalSyncContext();
  const warnings = [];
  const infos = [];
  ctx.console.warn = (msg) => warnings.push(String(msg));
  ctx.console.info = (msg) => infos.push(String(msg));

  ctx.logDirectionsDiagnostics_("Origin", "Destination", {
    status: "NOT_FOUND",
    geocoded_waypoints: [
      { geocoder_status: "ZERO_RESULTS", partial_match: true, place_id: "abc" },
    ],
    routes: [],
  });

  assert.ok(
    warnings.some((msg) => msg.includes("Directions status")),
    warnings.join("\n"),
  );
  assert.ok(
    infos.some((msg) => msg.includes("waypoint 0")),
    infos.join("\n"),
  );
  assert.ok(
    warnings.some((msg) => msg.includes("returned no routes")),
    warnings.join("\n"),
  );
});

test("getDriveMinutes_ logs exception details on lookup failure", () => {
  const ctx = loadIcalSyncContext();
  const errors = [];
  ctx.console.error = (msg) => errors.push(String(msg));
  ctx.Maps.newDirectionFinder = () => ({
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
      const err = new Error("quota exceeded");
      err.name = "QuotaError";
      throw err;
    },
  });

  const minutes = ctx.getDriveMinutes_("Origin", "Destination", {});

  assert.equal(minutes, null);
  assert.ok(errors.length > 0);
  assert.match(errors[0], /QuotaError/);
  assert.match(errors[0], /quota exceeded/);
});

test("calendarEventInsert_ fails fast on Calendar usage limits", () => {
  const ctx = loadIcalSyncContext();
  const sleeps = [];
  const warnings = [];
  const errors = [];
  let attempts = 0;
  ctx.Utilities.sleep = (ms) => sleeps.push(ms);
  ctx.console.warn = (msg) => warnings.push(String(msg));
  ctx.console.error = (msg) => errors.push(String(msg));
  ctx.Calendar.Events.insert = () => {
    attempts++;
    throw new Error(
      "GoogleJsonResponseException: Calendar usage limits exceeded.",
    );
  };

  assert.throws(
    () =>
      ctx.calendarEventInsert_(
        {
          summary: "Practice",
          start: { dateTime: "2099-05-01T15:00:00Z" },
          extendedProperties: {
            private: {
              managedKind: "source",
              sourceFeedName: "Practice Feed",
              syncKey: "feedhash:practice-1",
            },
          },
        },
        "calendar-1",
        { sendUpdates: "none" },
      ),
    /Calendar usage limits exceeded/,
  );

  assert.equal(attempts, 1);
  assert.equal(sleeps.length, 0);
  assert.equal(warnings.length, 0);
  assert.equal(errors.length, 1);
  assert.match(errors[0], /\[CALENDAR_WRITE_FAILED\] op=insert/);
  assert.match(errors[0], /errorType=calendar_usage_limits/);
  assert.match(errors[0], /attempt=1\/5/);
  assert.match(errors[0], /retryable=false/);
  assert.match(errors[0], /calendarId="calendar-1"/);
  assert.match(errors[0], /eventKind="source event"/);
  assert.match(errors[0], /title="Practice"/);
  assert.match(errors[0], /eventDate="2099-05-01"/);
  assert.match(errors[0], /feedName="Practice Feed"/);
  assert.doesNotMatch(errors[0], /eventId=|syncKey=/);
  assert.match(errors[0], /writeNumber=1/);
  assert.match(errors[0], /writesSucceeded=0/);
});

test("calendarEventInsert_ logs terminal rate-limit diagnostics", () => {
  const ctx = loadIcalSyncContext();
  const errors = [];
  const warnings = [];
  let attempts = 0;
  ctx.console.error = (msg) => errors.push(String(msg));
  ctx.console.warn = (msg) => warnings.push(String(msg));
  ctx.Utilities.sleep = () => {};
  ctx.Calendar.Events.insert = () => {
    attempts++;
    throw new Error("Rate Limit Exceeded");
  };

  assert.throws(
    () =>
      ctx.calendarEventInsert_(
        {
          summary: "Practice",
          start: { dateTime: "2099-05-01T15:00:00Z" },
          extendedProperties: {
            private: {
              managedKind: "source",
              sourceFeedName: "Practice Feed",
              syncKey: "feedhash:practice-1",
            },
          },
        },
        "calendar-1",
        { sendUpdates: "none" },
      ),
    /Rate Limit Exceeded/,
  );

  assert.equal(attempts, 5);
  assert.equal(errors.length, 1);
  assert.equal(warnings.length, 4);
  assert.match(warnings[0], /\[CALENDAR_WRITE_RETRY\] op=insert/);
  assert.match(warnings[0], /errorType=rate_limit/);
  assert.match(warnings[0], /attempt=1\/5/);
  assert.match(warnings[0], /nextDelayMs=\d+/);
  assert.match(warnings[0], /eventKind="source event"/);
  assert.match(warnings[0], /title="Practice"/);
  assert.match(warnings[0], /eventDate="2099-05-01"/);
  assert.match(warnings[0], /calendarId="calendar-1"/);
  assert.match(warnings[0], /feedName="Practice Feed"/);
  assert.match(warnings[0], /error="Error: Rate Limit Exceeded"/);
  assert.doesNotMatch(warnings[0], /eventId=|syncKey=/);
  assert.match(errors[0], /\[CALENDAR_WRITE_FAILED\] op=insert/);
  assert.match(errors[0], /errorType=rate_limit/);
  assert.match(errors[0], /attempt=5\/5/);
  assert.match(errors[0], /retryable=true/);
  assert.match(errors[0], /writeNumber=1/);
  assert.match(errors[0], /writesSucceeded=0/);
  assert.match(errors[0], /eventKind="source event"/);
  assert.match(errors[0], /title="Practice"/);
  assert.match(errors[0], /eventDate="2099-05-01"/);
  assert.match(errors[0], /feedName="Practice Feed"/);
  assert.doesNotMatch(errors[0], /eventId=|syncKey=/);
});

test("calendarEventInsert_ enforces deterministic Calendar event IDs", () => {
  const ctx = loadIcalSyncContext();
  const insertedResources = [];
  ctx.Calendar.Events.insert = (resource) => {
    insertedResources.push(resource);
    return { id: resource.id };
  };

  function resourceFor(syncKey, suppliedId) {
    return {
      id: suppliedId,
      summary: "Practice",
      extendedProperties: {
        private: {
          managedKind: "source",
          syncKey,
        },
      },
    };
  }

  const first = ctx.calendarEventInsert_(
    resourceFor("feedhash:uid-1", "caller-supplied-id"),
    "calendar-1",
    { sendUpdates: "none" },
  );
  const sameLogicalEvent = ctx.calendarEventInsert_(
    resourceFor("feedhash:uid-1"),
    "calendar-2",
    { sendUpdates: "none" },
  );
  const differentLogicalEvent = ctx.calendarEventInsert_(
    resourceFor("feedhash:uid-2"),
    "calendar-1",
    { sendUpdates: "none" },
  );

  assert.equal(first.id, sameLogicalEvent.id);
  assert.notEqual(first.id, differentLogicalEvent.id);
  assert.equal(
    first.id,
    ctx.buildDeterministicCalendarEventId_("feedhash:uid-1"),
  );
  assert.match(first.id, /^[0-9a-f]{64}$/);
  assert.equal(insertedResources[0].id, first.id);
  assert.notEqual(insertedResources[0].id, "caller-supplied-id");

  assert.throws(
    () =>
      ctx.calendarEventInsert_({ summary: "Missing sync key" }, "calendar-1", {
        sendUpdates: "none",
      }),
    /requires extendedProperties\.private\.syncKey/,
  );
});

test("calendarEventInsert_ recovers a matching deterministic ID conflict", () => {
  const ctx = loadIcalSyncContext();
  const warnings = [];
  const errors = [];
  let getArgs;
  const syncKey = "feedhash:uid-1";
  const eventId = ctx.buildDeterministicCalendarEventId_(syncKey);
  const existing = {
    id: eventId,
    status: "confirmed",
    summary: "Practice",
    extendedProperties: {
      private: {
        managedKind: "source",
        syncKey,
      },
    },
  };
  ctx.console.warn = (msg) => warnings.push(String(msg));
  ctx.console.error = (msg) => errors.push(String(msg));
  ctx.Calendar.Events.insert = () => {
    throw new Error("The requested identifier already exists.");
  };
  ctx.Calendar.Events.get = (calendarId, requestedEventId) => {
    getArgs = [calendarId, requestedEventId];
    return existing;
  };

  const result = ctx.calendarEventInsert_(
    {
      summary: "Practice",
      start: { dateTime: "2099-05-01T15:00:00Z" },
      extendedProperties: {
        private: {
          managedKind: "source",
          sourceFeedName: "Practice Feed",
          syncKey,
        },
      },
    },
    "calendar-1",
    { sendUpdates: "none" },
  );

  assert.equal(result, existing);
  assert.deepEqual(getArgs, ["calendar-1", eventId]);
  assert.equal(errors.length, 0);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /\[CALENDAR_INSERT_RECOVERED\]/);
  assert.match(warnings[0], /eventKind="source event"/);
  assert.match(warnings[0], /title="Practice"/);
  assert.match(warnings[0], /eventDate="2099-05-01"/);
  assert.match(warnings[0], /feedName="Practice Feed"/);
  assert.doesNotMatch(warnings[0], /eventId=|syncKey=/);
  assert.match(warnings[0], /writesSucceeded=0/);
});

test("calendarEventInsert_ rejects a mismatched deterministic ID conflict", () => {
  const ctx = loadIcalSyncContext();
  const errors = [];
  ctx.console.error = (msg) => errors.push(String(msg));
  ctx.Calendar.Events.insert = () => {
    throw new Error("The requested identifier already exists.");
  };
  ctx.Calendar.Events.get = (_calendarId, eventId) => ({
    id: eventId,
    status: "confirmed",
    extendedProperties: {
      private: {
        syncKey: "another-feed:uid-2",
      },
    },
  });

  assert.throws(
    () =>
      ctx.calendarEventInsert_(
        {
          summary: "Practice",
          extendedProperties: {
            private: {
              managedKind: "source",
              syncKey: "feedhash:uid-1",
            },
          },
        },
        "calendar-1",
        { sendUpdates: "none" },
      ),
    /Deterministic Calendar event ID conflict/,
  );

  assert.equal(errors.length, 1);
  assert.match(errors[0], /\[CALENDAR_WRITE_FAILED\] op=insert/);
  assert.match(errors[0], /errorType=non_quota_error/);
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
    placeNameAddressRules: [],
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
    placeNameAddressRules: [],
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

test("resolveDrivePlan_ marks routeLookupFailed when drive cannot be computed", () => {
  const ctx = loadIcalSyncContext();
  const driveOpts = {
    originAddress: "123 Main St, Brooklyn, NY",
    minDriveMinutesToCreate: 10,
    placeNameAddressRules: [],
  };
  const driveEnd = new Date("2099-05-01T15:00:00Z");

  ctx.findPreviousDriveOriginEvent_ = () => null;
  ctx.getDriveMinutes_ = () => null;

  const plan = ctx.resolveDrivePlan_(
    "calendar-1",
    "source-1",
    driveEnd,
    "Main Field",
    driveOpts,
    {},
  );

  assert.equal(plan.routeLookupFailed, true);
  assert.match(plan.skipReason, /route lookup failed/i);
  assert.equal(Array.isArray(plan.lookupFailures), true);
  assert.equal(plan.lookupFailures.length, 1);
});

test("resolvePlaceNameAddress_ maps venue names to canonical addresses", () => {
  const ctx = loadIcalSyncContext();
  const rules = [
    {
      placeName: "McMoran Park",
      placeNameLower: "mcmoran park",
      address: "1000 Park Ave, City, ST",
    },
    {
      placeName: "McMoran",
      placeNameLower: "mcmoran",
      address: "2000 Park Ave, City, ST",
    },
  ];

  const mapped = ctx.resolvePlaceNameAddress_(
    "McMoran Park Field 1 (60)",
    rules,
  );
  assert.equal(mapped.matched, true);
  assert.equal(mapped.matchedPlaceName, "McMoran Park");
  assert.equal(mapped.text, "1000 Park Ave, City, ST");

  const caseInsensitive = ctx.resolvePlaceNameAddress_(
    "mcmoran park field 2",
    rules,
  );
  assert.equal(caseInsensitive.matched, true);
  assert.equal(caseInsensitive.text, "1000 Park Ave, City, ST");

  const untouched = ctx.resolvePlaceNameAddress_("Some Other Venue", rules);
  assert.equal(untouched.matched, false);
  assert.equal(untouched.text, "Some Other Venue");
});

test("applyPlaceNameAddressToEvent_ rewrites matching event locations", () => {
  const ctx = loadIcalSyncContext();
  const evt = {
    uid: "uid-1",
    summary: "Practice",
    location: "Kaiserman JCC- Field 2",
  };
  const rules = [
    {
      placeName: "Kaiserman JCC",
      placeNameLower: "kaiserman jcc",
      address: "45 Haverford Rd, Penn Wynne, PA 19096",
    },
  ];

  const rewritten = ctx.applyPlaceNameAddressToEvent_(evt, rules);

  assert.notEqual(rewritten, evt);
  assert.equal(rewritten.location, "45 Haverford Rd, Penn Wynne, PA 19096");
  assert.equal(evt.location, "Kaiserman JCC- Field 2");
});

test("applyPlaceNameAddressToEvent_ leaves unmatched event locations alone", () => {
  const ctx = loadIcalSyncContext();
  const evt = {
    uid: "uid-1",
    summary: "Practice",
    location: "Some Other Venue",
  };
  const rules = [
    {
      placeName: "Kaiserman JCC",
      placeNameLower: "kaiserman jcc",
      address: "45 Haverford Rd, Penn Wynne, PA 19096",
    },
  ];

  const unchanged = ctx.applyPlaceNameAddressToEvent_(evt, rules);

  assert.equal(unchanged, evt);
  assert.equal(unchanged.location, "Some Other Venue");
});

test("resolveDrivePlan_ applies place-name mappings before route lookup", () => {
  const ctx = loadIcalSyncContext();
  const captured = [];
  const driveOpts = {
    originAddress: "Brooklyn, NY",
    minDriveMinutesToCreate: 10,
    placeNameAddressRules: [
      {
        placeName: "McMoran Park",
        placeNameLower: "mcmoran park",
        address: "1000 Park Ave, City, ST",
      },
    ],
  };
  const driveEnd = new Date("2099-05-01T15:00:00Z");

  ctx.findPreviousDriveOriginEvent_ = () => ({
    id: "prev-1",
    summary: "Lunch",
    location: "North Lot",
  });
  ctx.getDriveMinutes_ = (origin, destination) => {
    captured.push([origin, destination]);
    return 25;
  };

  const plan = ctx.resolveDrivePlan_(
    "calendar-1",
    "source-1",
    driveEnd,
    "McMoran Park Field 1 (60)",
    driveOpts,
    {},
  );

  assert.equal(plan.originAddress, "North Lot");
  assert.equal(plan.driveMinutes, 25);
  assert.deepEqual(captured, [["North Lot", "1000 Park Ave, City, ST"]]);
});

test("findPreviousDriveOriginEvent_ ignores events older than one hour", () => {
  const ctx = loadIcalSyncContext();
  ctx.Calendar.Events.list = () => ({
    items: [
      {
        id: "old-1",
        summary: "Old Practice",
        location: "Old Gym",
        start: { dateTime: "2099-05-01T12:00:00Z" },
        end: { dateTime: "2099-05-01T13:00:00Z" },
      },
      {
        id: "recent-1",
        summary: "Recent Practice",
        location: "Recent Gym",
        start: { dateTime: "2099-05-01T14:15:00Z" },
        end: { dateTime: "2099-05-01T14:45:00Z" },
      },
    ],
  });

  const found = ctx.findPreviousDriveOriginEvent_(
    "calendar-1",
    new Date("2099-05-01T15:00:00Z"),
    "source-1",
    60,
  );

  assert.ok(found);
  assert.equal(found.id, "recent-1");
});

test("reconcileDrivePlaceholder_ skips events without a location", () => {
  const ctx = loadIcalSyncContext();
  const inserted = [];
  ctx.findPreviousDriveOriginEvent_ = () => {
    throw new Error("unexpected previous event lookup");
  };
  ctx.getDriveMinutes_ = () => {
    throw new Error("unexpected drive lookup");
  };
  ctx.Calendar.Events.insert = (resource, calendarId) => {
    assert.equal(calendarId, "b@example.com");
    inserted.push(resource);
    return {
      id: "drive-1",
      start: resource.start,
      end: resource.end,
      extendedProperties: resource.extendedProperties,
    };
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = () => {};

  const evt = {
    uid: "uid-1",
    summary: "McMoran Park Field 1 (60)",
    location: "",
    start: { type: "dateTime", dateTime: "2099-05-01T15:00:00Z" },
    end: { type: "dateTime", dateTime: "2099-05-01T16:00:00Z" },
  };
  const syncedEvent = {
    id: "source-1",
    start: { dateTime: "2099-05-01T15:00:00Z" },
  };
  const mapping = {
    feedUrl: "https://example.com/feed.ics",
    calendarId: "b@example.com",
  };
  const feedHash = "feedhash123";
  const sourceSyncKey = "feedhash123:abc";
  const driveSyncKey = ctx.buildDriveSyncKey_(sourceSyncKey);
  const driveOpts = {
    enabled: true,
    originAddress: "Brooklyn, NY",
    placeNameAddressRules: [
      {
        placeName: "McMoran Park",
        placeNameLower: "mcmoran park",
        address: "1000 Park Ave, City, ST",
      },
    ],
    minDriveMinutesToCreate: 10,
    titleTemplate: "Drive to {{title}}",
  };
  const existingDriveByKey = {};
  const seenDrive = {};
  const today = new Date("2026-01-01T00:00:00Z");
  const stats = baseStats();

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
    stats,
    {},
    null,
    [],
  );

  assert.equal(stats.driveCreated, 0);
  assert.equal(stats.driveSkipped, 1);
  assert.equal(inserted.length, 0);
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
    { name: "Practice Feed", feedUrl: "https://example.com/feed.ics" },
    "feedhash123",
    { uid: "uid-1", location: "Destination" },
    "feedhash123:source-sync",
    "drive:feedhash123:source-sync",
    "source-event-123",
    "Practice",
    "Drive to Meeting",
    driveStart,
    driveEnd,
    "drivehash123",
    "Origin Address",
    "Destination Address",
    ["a@example.com", "b@example.com"],
  );

  const p = resource.extendedProperties.private;
  assert.equal(p.managedKind, "drive");
  assert.equal(p.sourceFeedName, "Practice Feed");
  assert.equal(p.sourceSyncKey, "feedhash123:source-sync");
  assert.equal(p.sourceEventId, "source-event-123");
  assert.equal(p.syncKey, "drive:feedhash123:source-sync");
  assert.match(
    resource.description,
    /<strong>Source event:<\/strong> Practice/,
  );
  assert.match(resource.description, /<strong>From:<\/strong> Origin Address/);
  assert.match(
    resource.description,
    /<strong>To:<\/strong> Destination Address/,
  );
  assert.match(
    resource.description,
    /<a href="https:\/\/www\.google\.com\/maps\/dir\/\?api=1&amp;travelmode=driving&amp;origin=Origin%20Address&amp;destination=Destination%20Address">Open driving directions in Google Maps<\/a>/,
  );
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
    { name: "Practice Feed", feedUrl: "https://example.com/feed.ics" },
    "feedhash123",
    { uid: "uid-1", location: "Destination" },
    "feedhash123:source-sync",
    "arrival:feedhash123:source-sync",
    "source-event-123",
    "Practice",
    "Advanced arrival for Practice",
    arrivalStart,
    arrivalEnd,
    "arrivalhash123",
    30,
    ["a@example.com", "b@example.com"],
  );

  const p = resource.extendedProperties.private;
  assert.equal(p.managedKind, "arrival");
  assert.equal(p.sourceFeedName, "Practice Feed");
  assert.equal(p.sourceSyncKey, "feedhash123:source-sync");
  assert.equal(p.sourceEventId, "source-event-123");
  assert.equal(p.arrivalMinutes, "30");
  assert.equal(p.syncKey, "arrival:feedhash123:source-sync");
  assert.match(resource.description, /Source event: Practice/);
  assert.equal(
    JSON.stringify(resource.attendees),
    JSON.stringify([{ email: "a@example.com" }, { email: "b@example.com" }]),
  );
});

test("attendee selection uses configured attendees and ignores current user", () => {
  const ctx = loadIcalSyncContext();
  const existing = {
    attendees: [
      { email: "owner@example.com", self: true, responseStatus: "declined" },
      { email: "calendar-1", responseStatus: "accepted" },
    ],
  };

  assert.equal(
    JSON.stringify(
      ctx.buildSourceAttendees_(["coach@example.com"], "calendar-1"),
    ),
    JSON.stringify(["coach@example.com", "calendar-1"]),
  );
  assert.equal(
    JSON.stringify(
      ctx.buildSourceAttendees_(["coach@example.com"], "calendar-1", false),
    ),
    JSON.stringify(["coach@example.com"]),
  );
  assert.equal(
    JSON.stringify(ctx.buildSourceAttendees_([], "calendar-1", false)),
    JSON.stringify([]),
  );
  assert.equal(
    ctx.isTargetCalendarDeclinedEvent_(existing, "calendar-1"),
    false,
  );
});

test("syncOneFeed_ uses defaults only when attendeeEmails is omitted", () => {
  const ctx = loadIcalSyncContext();
  const inserts = [];

  ctx.fetchIcs_ = () =>
    [
      "BEGIN:VCALENDAR",
      "BEGIN:VEVENT",
      "UID:uid-1",
      "DTSTART:20990501T150000Z",
      "DTEND:20990501T160000Z",
      "SUMMARY:Practice",
      "LOCATION:Seattle, WA",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\n");
  ctx.loadExistingEventsByKey_ = () => ({});
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.Calendar.Events.insert = (resource) => {
    inserts.push(resource);
    return {
      id: "source-created-1",
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
    defaultAttendeeEmails: ["coach@example.com"],
    addDriveTimePlaceholders: false,
  };

  ctx.syncOneFeed_(
    cfg,
    {
      name: "Omitted Attendees",
      feedUrl: "https://example.com/feed.ics",
      calendarId: "calendar-1",
      titlePrefix: "",
      addDriveTimePlaceholders: false,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(
    JSON.stringify(inserts[0].attendees),
    JSON.stringify([{ email: "coach@example.com" }, { email: "calendar-1" }]),
  );

  inserts.length = 0;
  ctx.syncOneFeed_(
    cfg,
    {
      name: "Destination Not Attendee",
      feedUrl: "https://example.com/feed.ics",
      calendarId: "calendar-1",
      titlePrefix: "",
      attendeeEmails: ["parent@example.com"],
      addDestinationCalendarAsAttendee: false,
      addDriveTimePlaceholders: false,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(
    JSON.stringify(inserts[0].attendees),
    JSON.stringify([{ email: "parent@example.com" }]),
  );
});

test("syncOneFeed_ optionally filters out all-day events for a feed", () => {
  const ctx = loadIcalSyncContext();
  const inserts = [];

  ctx.fetchIcs_ = () =>
    [
      "BEGIN:VCALENDAR",
      "BEGIN:VEVENT",
      "UID:all-day-1",
      "DTSTART;VALUE=DATE:20990501",
      "DTEND;VALUE=DATE:20990502",
      "SUMMARY:Tournament Day",
      "END:VEVENT",
      "BEGIN:VEVENT",
      "UID:timed-1",
      "DTSTART:20990501T150000Z",
      "DTEND:20990501T160000Z",
      "SUMMARY:Practice",
      "LOCATION:Seattle, WA",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\n");
  ctx.loadExistingEventsByKey_ = () => ({});
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.Calendar.Events.insert = (resource) => {
    inserts.push(resource);
    return {
      id: "source-created-" + inserts.length,
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

  const stats = ctx.syncOneFeed_(
    {
      deleteMissingFromFeed: false,
      defaultAttendeeEmails: [],
      addDriveTimePlaceholders: false,
    },
    {
      name: "No All-Day",
      feedUrl: "https://example.com/feed.ics",
      calendarId: "calendar-1",
      titlePrefix: "",
      skipAllDayEvents: true,
      addDriveTimePlaceholders: false,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(stats.created, 1);
  assert.equal(stats.skipped, 1);
  assert.equal(inserts.length, 1);
  assert.equal(inserts[0].summary, "Practice");
});

test("syncOneFeed_ logs the event title when create insert fails", () => {
  const ctx = loadIcalSyncContext();
  const errors = [];

  ctx.console.error = (msg) => errors.push(String(msg));
  ctx.fetchIcs_ = () =>
    [
      "BEGIN:VCALENDAR",
      "BEGIN:VEVENT",
      "UID:uid-1",
      "DTSTART:20990501T150000Z",
      "DTEND:20990501T160000Z",
      "SUMMARY:Practice",
      "LOCATION:Seattle, WA",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\n");
  ctx.loadExistingEventsByKey_ = () => ({});
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.Calendar.Events.insert = () => {
    throw new Error("insert failed");
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = () => {
    throw new Error("unexpected remove");
  };

  assert.throws(() => {
    ctx.syncOneFeed_(
      {
        deleteMissingFromFeed: false,
        defaultAttendeeEmails: [],
        addDriveTimePlaceholders: false,
      },
      {
        name: "Practice Feed",
        feedUrl: "https://example.com/feed.ics",
        calendarId: "calendar-1",
        titlePrefix: "",
        addDriveTimePlaceholders: false,
        originAddress: "",
      },
      new Date("2026-01-01T00:00:00Z"),
    );
  }, /insert failed/);

  assert.ok(
    errors.some(
      (msg) =>
        msg.includes(
          '[ERROR] source event "Practice" on 2099-05-01 in calendar-1 from Practice Feed — create failed:',
        ) && msg.includes("insert failed"),
    ),
  );
});

test("loadExistingEventsByKey_ removes duplicate managed sync keys", () => {
  const ctx = loadIcalSyncContext();
  const removed = [];
  const feedHash = "feedhash123";
  const syncKey = "feedhash123:same-source";
  ctx.Calendar.Events.list = () => ({
    items: [
      {
        id: "source-original",
        created: "2026-07-15T20:00:00Z",
        extendedProperties: {
          private: {
            managedKind: "source",
            sourceFeed: feedHash,
            syncKey,
          },
        },
      },
      {
        id: "source-duplicate-1",
        created: "2026-07-15T20:01:00Z",
        extendedProperties: {
          private: {
            managedKind: "source",
            sourceFeed: feedHash,
            syncKey,
          },
        },
      },
      {
        id: "source-duplicate-2",
        created: "2026-07-15T20:02:00Z",
        extendedProperties: {
          private: {
            managedKind: "source",
            sourceFeed: feedHash,
            syncKey,
          },
        },
      },
    ],
  });
  ctx.Calendar.Events.remove = (calendarId, eventId) => {
    assert.equal(calendarId, "calendar-1");
    removed.push(eventId);
  };

  const existing = ctx.loadExistingEventsByKey_("calendar-1", feedHash);

  assert.equal(existing[syncKey].id, "source-original");
  assert.deepEqual(removed, ["source-duplicate-1", "source-duplicate-2"]);
});

test("eventBoundaryToDate_ honors the timezone on a floating boundary", () => {
  const ctx = loadIcalSyncContext();

  const date = ctx.eventBoundaryToDate_({
    dateTime: "2026-08-18T14:00:00",
    timeZone: "America/Los_Angeles",
  });

  assert.equal(date.toISOString(), "2026-08-18T21:00:00.000Z");
});

test("syncOneFeed_ preserves attendees on existing events when exclusion is fix-forward", () => {
  const ctx = loadIcalSyncContext();
  const feedUrl = "https://example.com/feed.ics";
  const feedHash = ctx.sha256Hex_(feedUrl).slice(0, 16);
  const sourceSyncKey = ctx.buildSyncKey_(feedHash, "uid-1", "");
  let patchCalls = 0;

  const icsText = [
    "BEGIN:VCALENDAR",
    "BEGIN:VEVENT",
    "UID:uid-1",
    "DTSTART:20990501T150000Z",
    "DTEND:20990501T160000Z",
    "SUMMARY:Practice",
    "LOCATION:Seattle, WA",
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\n");
  const parsedEvent = ctx.parseIcs_(icsText).events[0];

  ctx.fetchIcs_ = () => icsText;
  ctx.loadExistingEventsByKey_ = () => ({
    [sourceSyncKey]: {
      id: "source-1",
      summary: "Practice",
      start: { dateTime: "2099-05-01T15:00:00.000Z" },
      end: { dateTime: "2099-05-01T16:00:00.000Z" },
      attendees: [
        { email: "calendar-1", responseStatus: "accepted" },
        { email: "legacy-guest@example.com", responseStatus: "accepted" },
      ],
      extendedProperties: {
        private: {
          managedKind: "source",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: sourceSyncKey,
          syncHash: ctx.computeEventHash_(parsedEvent, [
            "calendar-1",
            "legacy-guest@example.com",
          ]),
        },
      },
    },
  });
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.Calendar.Events.insert = () => {
    throw new Error("unexpected insert");
  };
  ctx.Calendar.Events.patch = () => {
    patchCalls++;
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = () => {
    throw new Error("unexpected remove");
  };

  const stats = ctx.syncOneFeed_(
    {
      deleteMissingFromFeed: false,
      defaultAttendeeEmails: [],
      addDriveTimePlaceholders: false,
    },
    {
      name: "Stable Feed",
      feedUrl: feedUrl,
      calendarId: "calendar-1",
      titlePrefix: "",
      attendeeEmails: [],
      addDestinationCalendarAsAttendee: false,
      addDriveTimePlaceholders: false,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(patchCalls, 0);
  assert.equal(stats.updated, 0);
  assert.equal(stats.unchanged, 1);
});

test("syncOneFeed_ uses a feed timezone and 30-minute fallback end for floating times", () => {
  const ctx = loadIcalSyncContext();
  const inserts = [];
  const icsText = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "BEGIN:VEVENT",
    "UID:17358853@wildwood.piedmont.k12.ca.us",
    "DTSTART:20260818T140000",
    "SUMMARY:2 PM Release Grades 1-3",
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\n");

  ctx.fetchIcs_ = () => icsText;
  ctx.loadExistingEventsByKey_ = () => ({});
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.Calendar.Events.list = () => ({ items: [] });
  ctx.Calendar.Events.insert = (resource) => {
    inserts.push(resource);
    return {
      id: "wildwood-created-1",
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

  const stats = ctx.syncOneFeed_(
    {
      deleteMissingFromFeed: false,
      defaultAttendeeEmails: [],
      addDriveTimePlaceholders: false,
    },
    {
      name: "Wildwood",
      feedUrl: "https://wildwood.piedmont.k12.ca.us/calendar/calendar_356.ics",
      calendarId: "calendar-1",
      titlePrefix: "Wildwood:",
      timeZone: "America/Los_Angeles",
      addDriveTimePlaceholders: false,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(stats.created, 1);
  assert.equal(inserts.length, 1);
  assert.equal(
    JSON.stringify(inserts[0].start),
    JSON.stringify({
      dateTime: "2026-08-18T14:00:00",
      timeZone: "America/Los_Angeles",
    }),
  );
  assert.equal(
    JSON.stringify(inserts[0].end),
    JSON.stringify({
      dateTime: "2026-08-18T14:30:00",
      timeZone: "America/Los_Angeles",
    }),
  );
});

test("syncOneFeed_ treats Calendar offset timestamps as the same feed time", () => {
  const ctx = loadIcalSyncContext();
  const feedUrl =
    "https://wildwood.piedmont.k12.ca.us/calendar/calendar_356.ics";
  const feedHash = ctx.sha256Hex_(feedUrl).slice(0, 16);
  const sourceSyncKey = ctx.buildSyncKey_(
    feedHash,
    "17358853@wildwood.piedmont.k12.ca.us",
    "",
  );
  const icsText = [
    "BEGIN:VCALENDAR",
    "BEGIN:VEVENT",
    "UID:17358853@wildwood.piedmont.k12.ca.us",
    "DTSTART:20260818T140000",
    "SUMMARY:2 PM Release Grades 1-3",
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\n");
  const parsedEvent = ctx.parseIcs_(icsText, "America/Los_Angeles").events[0];
  let patchCalls = 0;

  ctx.fetchIcs_ = () => icsText;
  ctx.loadExistingEventsByKey_ = () => ({
    [sourceSyncKey]: {
      id: "source-1",
      summary: "Wildwood: 2 PM Release Grades 1-3",
      start: { dateTime: "2026-08-18T14:00:00-07:00" },
      end: { dateTime: "2026-08-18T14:30:00-07:00" },
      attendees: [
        { email: "calendar-1", responseStatus: "accepted" },
        { email: "legacy-guest@example.com", responseStatus: "accepted" },
      ],
      extendedProperties: {
        private: {
          managedKind: "source",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "17358853@wildwood.piedmont.k12.ca.us",
          syncKey: sourceSyncKey,
          syncHash: ctx.computeEventHash_(
            ctx.applyEventTitlePrefix_(parsedEvent, "Wildwood:"),
            ["calendar-1"],
          ),
        },
      },
    },
  });
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.Calendar.Events.insert = () => {
    throw new Error("unexpected insert");
  };
  ctx.Calendar.Events.patch = () => {
    patchCalls++;
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = () => {
    throw new Error("unexpected remove");
  };

  const stats = ctx.syncOneFeed_(
    {
      deleteMissingFromFeed: false,
      defaultAttendeeEmails: [],
      addDriveTimePlaceholders: false,
    },
    {
      name: "Wildwood",
      feedUrl,
      calendarId: "calendar-1",
      titlePrefix: "Wildwood:",
      timeZone: "America/Los_Angeles",
      addDriveTimePlaceholders: false,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(patchCalls, 0);
  assert.equal(stats.updated, 0);
  assert.equal(stats.unchanged, 1);
});

test("syncOneFeed_ patches unchanged feed events when destination time drifted", () => {
  const ctx = loadIcalSyncContext();
  const feedUrl = "https://example.com/feed.ics";
  const feedHash = ctx.sha256Hex_(feedUrl).slice(0, 16);
  const sourceSyncKey = ctx.buildSyncKey_(feedHash, "uid-1", "");
  const icsText = [
    "BEGIN:VCALENDAR",
    "X-WR-TIMEZONE:America/Los_Angeles",
    "BEGIN:VEVENT",
    "UID:uid-1",
    "DTSTART:20990501T140000",
    "DTEND:20990501T150000",
    "SUMMARY:Practice",
    "LOCATION:Seattle, WA",
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\n");
  const parsedEvent = ctx.parseIcs_(icsText).events[0];
  let patched = null;

  ctx.fetchIcs_ = () => icsText;
  ctx.loadExistingEventsByKey_ = () => ({
    [sourceSyncKey]: {
      id: "source-1",
      summary: "Practice",
      start: {
        dateTime: "2099-05-01T14:00:00",
        timeZone: "America/New_York",
      },
      end: {
        dateTime: "2099-05-01T15:00:00",
        timeZone: "America/New_York",
      },
      attendees: [
        { email: "calendar-1", responseStatus: "accepted" },
        { email: "legacy-guest@example.com", responseStatus: "accepted" },
      ],
      extendedProperties: {
        private: {
          managedKind: "source",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: sourceSyncKey,
          syncHash: ctx.computeEventHash_(parsedEvent, [
            "calendar-1",
            "legacy-guest@example.com",
          ]),
        },
      },
    },
  });
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.Calendar.Events.insert = () => {
    throw new Error("unexpected insert");
  };
  ctx.Calendar.Events.patch = (resource, calendarId, eventId) => {
    assert.equal(calendarId, "calendar-1");
    assert.equal(eventId, "source-1");
    patched = resource;
    return {
      id: eventId,
      start: resource.start,
      end: resource.end,
      attendees: resource.attendees,
      extendedProperties: resource.extendedProperties,
    };
  };
  ctx.Calendar.Events.remove = () => {
    throw new Error("unexpected remove");
  };

  const stats = ctx.syncOneFeed_(
    {
      deleteMissingFromFeed: false,
      defaultAttendeeEmails: [],
      addDriveTimePlaceholders: false,
    },
    {
      name: "Stable Feed",
      feedUrl: feedUrl,
      calendarId: "calendar-1",
      titlePrefix: "",
      attendeeEmails: [],
      addDestinationCalendarAsAttendee: false,
      addDriveTimePlaceholders: false,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(stats.updated, 1);
  assert.equal(stats.unchanged, 0);
  assert.ok(patched);
  assert.equal(
    JSON.stringify(patched.start),
    JSON.stringify({
      dateTime: "2099-05-01T14:00:00",
      timeZone: "America/Los_Angeles",
    }),
  );
  assert.equal(
    JSON.stringify(patched.end),
    JSON.stringify({
      dateTime: "2099-05-01T15:00:00",
      timeZone: "America/Los_Angeles",
    }),
  );
  assert.equal(
    JSON.stringify(patched.attendees),
    JSON.stringify([
      { email: "calendar-1", responseStatus: "accepted" },
      { email: "legacy-guest@example.com", responseStatus: "accepted" },
    ]),
  );
});

test("syncOneFeed_ skips duplicate event from another active feed", () => {
  const ctx = loadIcalSyncContext();
  const currentFeedUrl = "https://example.com/current.ics";
  const peerFeedUrl = "https://example.com/peer.ics";
  const peerFeedHash = ctx.sha256Hex_(peerFeedUrl).slice(0, 16);

  ctx.fetchIcs_ = () =>
    [
      "BEGIN:VCALENDAR",
      "BEGIN:VEVENT",
      "UID:uid-1",
      "DTSTART:20990501T150000Z",
      "DTEND:20990501T160000Z",
      "SUMMARY:Practice",
      "DESCRIPTION:Shared event",
      "LOCATION:Seattle, WA",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\n");
  ctx.loadExistingEventsByKey_ = () => ({});
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.Calendar.Events.list = () => ({
    items: [
      {
        id: "peer-1",
        summary: "Practice",
        description:
          "Shared event\n\nGenerated with github.com/streeter/google-apps-scripts",
        start: { dateTime: "2099-05-01T15:00:00.000Z" },
        end: { dateTime: "2099-05-01T16:00:00.000Z" },
        extendedProperties: {
          private: {
            managedKind: "source",
            sourceFeed: peerFeedHash,
            syncKey: peerFeedHash + ":peer-uid",
          },
        },
      },
    ],
  });
  ctx.Calendar.Events.insert = () => {
    throw new Error("unexpected insert");
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = () => {
    throw new Error("unexpected remove");
  };

  const stats = ctx.syncOneFeed_(
    {
      deleteMissingFromFeed: false,
      defaultAttendeeEmails: [],
      addDriveTimePlaceholders: false,
      feedMappings: [
        { feedUrl: currentFeedUrl, calendarId: "calendar-1" },
        { feedUrl: peerFeedUrl, calendarId: "calendar-1" },
      ],
    },
    {
      name: "Current Feed",
      feedUrl: currentFeedUrl,
      calendarId: "calendar-1",
      titlePrefix: "",
      addDriveTimePlaceholders: false,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(stats.created, 0);
  assert.equal(stats.skipped, 1);
});

test("syncOneFeed_ deletes unmanaged duplicate before creating managed event", () => {
  const ctx = loadIcalSyncContext();
  const inserted = [];
  const removed = [];
  const currentFeedUrl = "https://example.com/current.ics";

  ctx.fetchIcs_ = () =>
    [
      "BEGIN:VCALENDAR",
      "BEGIN:VEVENT",
      "UID:uid-1",
      "DTSTART:20990501T150000Z",
      "DTEND:20990501T160000Z",
      "SUMMARY:Practice",
      "DESCRIPTION:Shared event",
      "LOCATION:Seattle, WA",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\n");
  ctx.loadExistingEventsByKey_ = () => ({});
  ctx.loadExistingArrivalEventsByKey_ = () => ({});
  ctx.loadExistingDriveEventsByKey_ = () => ({});
  ctx.Calendar.Events.list = () => ({
    items: [
      {
        id: "manual-1",
        summary: "Practice",
        description: "Shared event",
        start: { dateTime: "2099-05-01T15:00:00.000Z" },
        end: { dateTime: "2099-05-01T16:00:00.000Z" },
      },
    ],
  });
  ctx.Calendar.Events.insert = (resource, calendarId) => {
    assert.equal(calendarId, "calendar-1");
    inserted.push(resource);
    return {
      id: "source-created-1",
      start: resource.start,
      end: resource.end,
      extendedProperties: resource.extendedProperties,
    };
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = (calendarId, eventId) => {
    assert.equal(calendarId, "calendar-1");
    removed.push(eventId);
  };

  const stats = ctx.syncOneFeed_(
    {
      deleteMissingFromFeed: false,
      defaultAttendeeEmails: [],
      addDriveTimePlaceholders: false,
      feedMappings: [{ feedUrl: currentFeedUrl, calendarId: "calendar-1" }],
    },
    {
      name: "Current Feed",
      feedUrl: currentFeedUrl,
      calendarId: "calendar-1",
      titlePrefix: "",
      addDriveTimePlaceholders: false,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(stats.deleted, 1);
  assert.equal(stats.created, 1);
  assert.equal(JSON.stringify(removed), JSON.stringify(["manual-1"]));
  assert.equal(inserted.length, 1);
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
    assert.equal(calendarId, "b@example.com");
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
    calendarId: "b@example.com",
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
  assert.equal(source.extendedProperties.private.sourceFeedName, "Test Feed");
  assert.equal(
    JSON.stringify(source.attendees),
    JSON.stringify([{ email: "b@example.com" }]),
  );
  assert.equal(drive.summary, "Drive (25m) to [Sports] Client Meeting");
  assert.equal(drive.extendedProperties.private.sourceFeedName, "Test Feed");
  assert.equal(
    drive.extendedProperties.private.sourceEventId,
    "source-created-1",
  );
  assert.equal(
    drive.extendedProperties.private.syncKey,
    ctx.buildDriveSyncKey_(drive.extendedProperties.private.sourceSyncKey),
  );
});

test("syncOneFeed_ creates attendee-free source and placeholders when configured", () => {
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
    assert.equal(calendarId, "b@example.com");
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
    calendarId: "b@example.com",
    attendeeEmails: [],
    addDestinationCalendarAsAttendee: false,
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
  assert.equal(JSON.stringify(source.attendees), JSON.stringify([]));
  assert.equal(arrival.summary, "Advanced arrival for [Sports] Soccer Game");
  assert.equal(arrival.start.dateTime, "2099-05-01T15:00:00.000Z");
  assert.equal(arrival.end.dateTime, "2099-05-01T15:30:00.000Z");
  assert.equal(drive.end.dateTime, "2099-05-01T15:00:00.000Z");
  assert.equal(drive.start.dateTime, "2099-05-01T14:35:00.000Z");
  assert.equal(JSON.stringify(arrival.attendees), JSON.stringify([]));
  assert.equal(JSON.stringify(drive.attendees), JSON.stringify([]));
});

test("existing arrival and drive placeholders preserve attendees when updated", () => {
  const ctx = loadIcalSyncContext();
  const patched = [];
  const feedHash = "feedhash123";
  const sourceSyncKey = feedHash + ":source-sync";
  const arrivalSyncKey = ctx.buildArrivalSyncKey_(sourceSyncKey);
  const driveSyncKey = ctx.buildDriveSyncKey_(sourceSyncKey);
  const legacyAttendees = [
    { email: "calendar-1", responseStatus: "accepted" },
    { email: "legacy-guest@example.com", responseStatus: "accepted" },
  ];
  const mapping = {
    feedUrl: "https://example.com/feed.ics",
    calendarId: "calendar-1",
    addDestinationCalendarAsAttendee: false,
  };
  const evt = {
    uid: "uid-1",
    summary: "Practice",
    description: "Arrival: 30 minutes in advance",
    location: "Seattle, WA",
    start: { type: "dateTime", dateTime: "2099-05-01T15:30:00Z" },
    end: { type: "dateTime", dateTime: "2099-05-01T16:30:00Z" },
  };
  const syncedEvent = {
    id: "source-1",
    start: { dateTime: "2099-05-01T15:30:00Z" },
  };
  const existingArrivalByKey = {
    [arrivalSyncKey]: {
      id: "arrival-1",
      attendees: legacyAttendees,
      extendedProperties: {
        private: {
          managedKind: "arrival",
          sourceFeed: feedHash,
          sourceUrl: mapping.feedUrl,
          sourceUid: evt.uid,
          syncKey: arrivalSyncKey,
          sourceSyncKey,
          syncHash: "outdated-arrival-hash",
        },
      },
    },
  };
  const existingDriveByKey = {
    [driveSyncKey]: {
      id: "drive-1",
      attendees: legacyAttendees,
      extendedProperties: {
        private: {
          managedKind: "drive",
          sourceFeed: feedHash,
          sourceUrl: mapping.feedUrl,
          sourceUid: evt.uid,
          syncKey: driveSyncKey,
          sourceSyncKey,
          syncHash: "outdated-drive-hash",
        },
      },
    },
  };
  const stats = baseStats();

  ctx.Calendar.Events.insert = () => {
    throw new Error("unexpected insert");
  };
  ctx.Calendar.Events.patch = (resource, calendarId, eventId) => {
    assert.equal(calendarId, "calendar-1");
    patched.push({ resource, eventId });
    return resource;
  };
  ctx.Calendar.Events.remove = () => {
    throw new Error("unexpected remove");
  };
  ctx.findPreviousDriveOriginEvent_ = () => null;
  ctx.getDriveMinutes_ = () => 25;

  ctx.reconcileArrivalPlaceholder_(
    evt,
    syncedEvent,
    mapping,
    feedHash,
    sourceSyncKey,
    arrivalSyncKey,
    existingArrivalByKey,
    {},
    new Date("2026-01-01T00:00:00Z"),
    stats,
    [],
  );
  ctx.reconcileDrivePlaceholder_(
    evt,
    syncedEvent,
    mapping,
    feedHash,
    sourceSyncKey,
    driveSyncKey,
    {
      enabled: true,
      originAddress: "New York, NY",
      placeNameAddressRules: [],
      minDriveMinutesToCreate: 10,
      titleTemplate: "Drive to {{title}}",
    },
    existingDriveByKey,
    {},
    new Date("2026-01-01T00:00:00Z"),
    stats,
    {},
    null,
    [],
  );

  assert.equal(stats.arrivalUpdated, 1);
  assert.equal(stats.driveUpdated, 1);
  assert.equal(patched.length, 2);
  assert.deepEqual(
    patched.map((entry) => entry.eventId),
    ["arrival-1", "drive-1"],
  );
  patched.forEach((entry) => {
    assert.equal(
      JSON.stringify(entry.resource.attendees),
      JSON.stringify(legacyAttendees),
    );
  });
});

test("syncOneFeed_ preserves a target calendar decline and removes placeholders", () => {
  const ctx = loadIcalSyncContext();
  const removed = [];
  let patched = null;
  const feedUrl = "https://example.com/sports.ics";
  const feedHash = ctx.sha256Hex_(feedUrl).slice(0, 16);
  const sourceSyncKey = ctx.buildSyncKey_(feedHash, "uid-1", "");

  ctx.fetchIcs_ = () =>
    [
      "BEGIN:VCALENDAR",
      "BEGIN:VEVENT",
      "UID:uid-1",
      "DTSTART:20990501T153000Z",
      "DTEND:20990501T163000Z",
      "SUMMARY:Soccer Game",
      "LOCATION:Seattle, WA",
      "END:VEVENT",
      "END:VCALENDAR",
    ].join("\n");
  ctx.loadExistingEventsByKey_ = () => ({
    [sourceSyncKey]: {
      id: "source-1",
      summary: "Soccer Game",
      location: "Seattle, WA",
      start: { dateTime: "2099-05-01T15:30:00Z" },
      end: { dateTime: "2099-05-01T16:30:00Z" },
      attendees: [
        {
          email: "calendar-1",
          self: true,
          responseStatus: "declined",
        },
        { email: "coach@example.com", responseStatus: "accepted" },
      ],
      extendedProperties: {
        private: {
          managedKind: "source",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: sourceSyncKey,
          syncHash: "old-hash",
        },
      },
    },
  });
  ctx.loadExistingArrivalEventsByKey_ = () => ({
    [ctx.buildArrivalSyncKey_(sourceSyncKey)]: {
      id: "arrival-1",
      summary: "Advanced arrival for Soccer Game",
      start: { dateTime: "2099-05-01T15:00:00Z" },
      end: { dateTime: "2099-05-01T15:30:00Z" },
      extendedProperties: {
        private: {
          managedKind: "arrival",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: ctx.buildArrivalSyncKey_(sourceSyncKey),
          sourceSyncKey: sourceSyncKey,
        },
      },
    },
  });
  ctx.loadExistingDriveEventsByKey_ = () => ({
    [ctx.buildDriveSyncKey_(sourceSyncKey)]: {
      id: "drive-1",
      summary: "Drive to Soccer Game",
      start: { dateTime: "2099-05-01T14:35:00Z" },
      end: { dateTime: "2099-05-01T15:00:00Z" },
      extendedProperties: {
        private: {
          managedKind: "drive",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: ctx.buildDriveSyncKey_(sourceSyncKey),
          sourceSyncKey: sourceSyncKey,
        },
      },
    },
  });
  ctx.getDriveMinutes_ = () => {
    throw new Error("unexpected drive lookup");
  };
  ctx.Calendar.Events.patch = (resource, calendarId, eventId) => {
    assert.equal(calendarId, "calendar-1");
    assert.equal(eventId, "source-1");
    patched = resource;
    return {
      id: eventId,
      start: resource.start,
      end: resource.end,
      attendees: resource.attendees,
      extendedProperties: resource.extendedProperties,
    };
  };
  ctx.Calendar.Events.insert = () => {
    throw new Error("unexpected insert");
  };
  ctx.Calendar.Events.remove = (calendarId, eventId) => {
    assert.equal(calendarId, "calendar-1");
    removed.push(eventId);
  };

  const cfg = {
    deleteMissingFromFeed: false,
    defaultAttendeeEmails: ["coach@example.com"],
    addDriveTimePlaceholders: true,
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

  assert.equal(stats.updated, 1);
  assert.equal(stats.arrivalDeleted, 1);
  assert.equal(stats.driveDeleted, 1);
  assert.deepEqual(removed.sort(), ["arrival-1", "drive-1"]);
  assert.ok(patched);
  assert.equal(patched.attendees.length, 1);
  assert.equal(
    JSON.stringify(patched.attendees[0]),
    JSON.stringify({
      email: "calendar-1",
      responseStatus: "declined",
      self: true,
    }),
  );
});

test("syncOneFeed_ removes placeholders when all destination attendees declined", () => {
  const ctx = loadIcalSyncContext();
  const removed = [];
  const feedUrl = "https://example.com/sports.ics";
  const feedHash = ctx.sha256Hex_(feedUrl).slice(0, 16);
  const sourceSyncKey = ctx.buildSyncKey_(feedHash, "uid-1", "");
  const arrivalSyncKey = ctx.buildArrivalSyncKey_(sourceSyncKey);
  const driveSyncKey = ctx.buildDriveSyncKey_(sourceSyncKey);
  const icsText = [
    "BEGIN:VCALENDAR",
    "BEGIN:VEVENT",
    "UID:uid-1",
    "DTSTART:20990501T153000Z",
    "DTEND:20990501T163000Z",
    "SUMMARY:Soccer Game",
    "DESCRIPTION:Arrival: 30 minutes in advance",
    "LOCATION:Seattle, WA",
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\n");
  const parsedEvent = ctx.parseIcs_(icsText).events[0];

  ctx.fetchIcs_ = () => icsText;
  ctx.loadExistingEventsByKey_ = () => ({
    [sourceSyncKey]: {
      id: "source-1",
      summary: "Soccer Game",
      location: "Seattle, WA",
      start: { dateTime: "2099-05-01T15:30:00Z" },
      end: { dateTime: "2099-05-01T16:30:00Z" },
      attendees: [
        { email: "coach@example.com", responseStatus: "declined" },
        { email: "parent@example.com", responseStatus: "declined" },
      ],
      extendedProperties: {
        private: {
          managedKind: "source",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: sourceSyncKey,
          syncHash: ctx.computeEventHash_(parsedEvent, [
            "coach@example.com",
            "parent@example.com",
            "calendar-1",
          ]),
        },
      },
    },
  });
  ctx.loadExistingArrivalEventsByKey_ = () => ({
    [arrivalSyncKey]: {
      id: "arrival-1",
      summary: "Advanced arrival for Soccer Game",
      start: { dateTime: "2099-05-01T15:00:00Z" },
      end: { dateTime: "2099-05-01T15:30:00Z" },
      extendedProperties: {
        private: {
          managedKind: "arrival",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: arrivalSyncKey,
          sourceSyncKey: sourceSyncKey,
        },
      },
    },
  });
  ctx.loadExistingDriveEventsByKey_ = () => ({
    [driveSyncKey]: {
      id: "drive-1",
      summary: "Drive to Soccer Game",
      start: { dateTime: "2099-05-01T14:35:00Z" },
      end: { dateTime: "2099-05-01T15:00:00Z" },
      extendedProperties: {
        private: {
          managedKind: "drive",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: driveSyncKey,
          sourceSyncKey: sourceSyncKey,
        },
      },
    },
  });
  ctx.getDriveMinutes_ = () => {
    throw new Error("unexpected drive lookup");
  };
  ctx.Calendar.Events.insert = () => {
    throw new Error("unexpected insert");
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = (calendarId, eventId) => {
    assert.equal(calendarId, "calendar-1");
    removed.push(eventId);
  };

  const stats = ctx.syncOneFeed_(
    {
      deleteMissingFromFeed: false,
      defaultAttendeeEmails: ["coach@example.com", "parent@example.com"],
      addDriveTimePlaceholders: true,
      defaultOriginAddress: "New York, NY",
      minDriveMinutesToCreate: 10,
      driveEventTitleTemplate: "Drive ({{minutes}}m) to {{title}}",
    },
    {
      name: "Sports Feed",
      feedUrl: feedUrl,
      calendarId: "calendar-1",
      titlePrefix: "",
      addDriveTimePlaceholders: true,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(stats.unchanged, 1);
  assert.equal(stats.arrivalDeleted, 1);
  assert.equal(stats.driveDeleted, 1);
  assert.deepEqual(removed.sort(), ["arrival-1", "drive-1"]);
});

test("syncOneFeed_ deletes managed future events missing from feed", () => {
  const ctx = loadIcalSyncContext();
  const removed = [];
  const logs = [];
  const feedUrl = "https://example.com/sports.ics";
  const feedHash = ctx.sha256Hex_(feedUrl).slice(0, 16);
  const sourceSyncKey = ctx.buildSyncKey_(feedHash, "uid-1", "");
  const arrivalSyncKey = ctx.buildArrivalSyncKey_(sourceSyncKey);
  const driveSyncKey = ctx.buildDriveSyncKey_(sourceSyncKey);

  ctx.fetchIcs_ = () => ["BEGIN:VCALENDAR", "END:VCALENDAR"].join("\n");
  ctx.loadExistingEventsByKey_ = () => ({
    [sourceSyncKey]: {
      id: "source-1",
      summary: "Missing Soccer Game",
      start: { dateTime: "2099-05-01T15:30:00Z" },
      end: { dateTime: "2099-05-01T16:30:00Z" },
      extendedProperties: {
        private: {
          managedKind: "source",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: sourceSyncKey,
        },
      },
    },
  });
  ctx.loadExistingArrivalEventsByKey_ = () => ({
    [arrivalSyncKey]: {
      id: "arrival-1",
      summary: "Advanced arrival for Missing Soccer Game",
      start: { dateTime: "2099-05-01T15:00:00Z" },
      end: { dateTime: "2099-05-01T15:30:00Z" },
      extendedProperties: {
        private: {
          managedKind: "arrival",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: arrivalSyncKey,
          sourceSyncKey: sourceSyncKey,
        },
      },
    },
  });
  ctx.loadExistingDriveEventsByKey_ = () => ({
    [driveSyncKey]: {
      id: "drive-1",
      summary: "Drive to Missing Soccer Game",
      start: { dateTime: "2099-05-01T14:35:00Z" },
      end: { dateTime: "2099-05-01T15:00:00Z" },
      extendedProperties: {
        private: {
          managedKind: "drive",
          sourceFeed: feedHash,
          sourceUrl: feedUrl,
          sourceUid: "uid-1",
          syncKey: driveSyncKey,
          sourceSyncKey: sourceSyncKey,
        },
      },
    },
  });
  ctx.Calendar.Events.insert = () => {
    throw new Error("unexpected insert");
  };
  ctx.Calendar.Events.patch = () => {
    throw new Error("unexpected patch");
  };
  ctx.Calendar.Events.remove = (calendarId, eventId) => {
    assert.equal(calendarId, "calendar-1");
    removed.push(eventId);
  };
  ctx.console.log = (message) => logs.push(String(message));

  const stats = ctx.syncOneFeed_(
    {
      deleteMissingFromFeed: true,
      defaultAttendeeEmails: [],
      addDriveTimePlaceholders: true,
      feedMappings: [{ feedUrl: feedUrl, calendarId: "calendar-1" }],
    },
    {
      name: "Sports Feed",
      feedUrl: feedUrl,
      calendarId: "calendar-1",
      attendeeEmails: [],
      titlePrefix: "",
      addDriveTimePlaceholders: true,
      originAddress: "",
    },
    new Date("2026-01-01T00:00:00Z"),
  );

  assert.equal(stats.deleted, 1);
  assert.equal(stats.arrivalDeleted, 1);
  assert.equal(stats.driveDeleted, 1);
  assert.deepEqual(removed.sort(), ["arrival-1", "drive-1", "source-1"]);
  assert.ok(
    logs.includes(
      '[DELETE] source event "Missing Soccer Game" on 2099-05-01 in calendar-1 from Sports Feed — missing from feed',
    ),
  );
});

test("eventStartDateForLog_ supports all-day and missing starts", () => {
  const ctx = loadIcalSyncContext();

  assert.equal(
    ctx.eventStartDateForLog_({ start: { date: "2099-05-02" } }),
    "2099-05-02",
  );
  assert.equal(ctx.eventStartDateForLog_({}), "(Unknown date)");
});

test("formatEventLogContext_ standardizes event metadata without IDs", () => {
  const ctx = loadIcalSyncContext();
  const context = ctx.formatEventLogContext_(
    {
      id: "event-id-not-for-logs",
      summary: "Soccer Game",
      start: { dateTime: "2099-05-03T18:00:00-07:00" },
    },
    "calendar-1",
    "Sports Feed",
    "source event",
  );

  assert.equal(
    context,
    'source event "Soccer Game" on 2099-05-03 in calendar-1 from Sports Feed',
  );
  assert.doesNotMatch(context, /event-id-not-for-logs/);
});
