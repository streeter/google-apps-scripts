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
      warn: () => {},
      error: () => {},
    },
    JSON,
    Date,
    Math,
    Logger: { log: () => {} },
    Session: {
      getScriptTimeZone: () => "America/New_York",
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
    { feedUrl: "https://example.com/feed.ics" },
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
    { feedUrl: "https://example.com/feed.ics" },
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
  assert.equal(
    JSON.stringify(source.attendees),
    JSON.stringify([{ email: "b@example.com" }]),
  );
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
  assert.equal(
    JSON.stringify(source.attendees),
    JSON.stringify([{ email: "b@example.com" }]),
  );
  assert.equal(arrival.summary, "Advanced arrival for [Sports] Soccer Game");
  assert.equal(arrival.start.dateTime, "2099-05-01T15:00:00.000Z");
  assert.equal(arrival.end.dateTime, "2099-05-01T15:30:00.000Z");
  assert.equal(drive.end.dateTime, "2099-05-01T15:00:00.000Z");
  assert.equal(drive.start.dateTime, "2099-05-01T14:35:00.000Z");
  assert.equal(
    JSON.stringify(arrival.attendees),
    JSON.stringify([{ email: "b@example.com" }]),
  );
  assert.equal(
    JSON.stringify(drive.attendees),
    JSON.stringify([{ email: "b@example.com" }]),
  );
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
