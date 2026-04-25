/**
 * iCal -> Google Calendar sync (future events), with update + attendee injection.
 *
 * This script expects a separate config file that defines:
 *   function getIcalSyncConfig() { return { ... }; }
 *
 * Example file: icalFeedSync.config.example.gs
 *
 * Required setup in Apps Script:
 * 1) Add Advanced Google Service: Calendar API
 * 2) Add icalFeedSync.config.gs with getIcalSyncConfig()
 * 3) Run setupIcalFeedSyncTrigger() once
 */

const GENERATED_BY_DESCRIPTION =
  "Generated with github.com/streeter/google-apps-scripts";

/**
 * Creates (or recreates) the periodic time-based trigger for the main sync function.
 */
function setupIcalFeedSyncTrigger() {
  const cfg = getIcalSyncConfig_();
  const fn = "syncIcalFeeds";
  ScriptApp.getProjectTriggers()
    .filter(function (t) {
      return t.getHandlerFunction() === fn;
    })
    .forEach(function (t) {
      ScriptApp.deleteTrigger(t);
    });

  const clockBuilder = ScriptApp.newTrigger(fn).timeBased();
  const scheduledBuilder = applyTriggerInterval_(
    clockBuilder,
    cfg.triggerEveryMinutes,
  );
  scheduledBuilder.create();
}

/**
 * Applies trigger cadence from minute configuration.
 * Supports:
 * - everyMinutes: 1, 5, 10, 15, 30
 * - everyHours: multiples of 60 (e.g. 60, 120, ..., 1380)
 * - everyDays: multiples of 1440 (e.g. 1440, 2880, ...)
 */
function applyTriggerInterval_(clockBuilder, triggerEveryMinutes) {
  const minutes = Number(triggerEveryMinutes);
  if (!isFinite(minutes) || minutes <= 0 || minutes !== Math.floor(minutes)) {
    throw new Error("triggerEveryMinutes must be a positive integer.");
  }

  const everyMinuteAllowed = [1, 5, 10, 15, 30];
  if (everyMinuteAllowed.indexOf(minutes) >= 0) {
    return clockBuilder.everyMinutes(minutes);
  }

  if (minutes % 1440 === 0) {
    const days = minutes / 1440;
    return clockBuilder.everyDays(days);
  }

  if (minutes % 60 === 0) {
    const hours = minutes / 60;
    return clockBuilder.everyHours(hours);
  }

  throw new Error(
    "Unsupported triggerEveryMinutes value: " +
      minutes +
      ". Use 1, 5, 10, 15, 30, or a multiple of 60.",
  );
}

/**
 * Main entry point: logs setup info, loads config, and syncs each feed mapping.
 */
function syncIcalFeeds() {
  logCalendarIdsOnFirstRun_();
  const cfg = getIcalSyncConfig_();
  const today = startOfToday_();
  const results = [];
  console.log(
    "[SYNC] Starting iCal feed sync for " +
      cfg.feedMappings.length +
      " feed(s)",
  );
  console.log("[SYNC] Date cutoff (inclusive): " + today.toISOString());

  cfg.feedMappings.forEach(function (mapping) {
    try {
      results.push(syncOneFeed_(cfg, mapping, today));
    } catch (e) {
      console.error(
        "[ERROR] Failed syncing feed " +
          (mapping.name || mapping.feedUrl) +
          ": " +
          String(e),
      );
      results.push({
        feed: mapping.name || mapping.feedUrl,
        error: String(e),
      });
    }
  });

  console.log("[SYNC] Finished iCal feed sync");
  Logger.log(JSON.stringify(results, null, 2));
}

/**
 * Utility entry point for setup-time logging of all accessible calendar IDs.
 */
function listMyCalendarIds() {
  logAllCalendarIds_();
}

/**
 * Syncs one ICS feed into one target calendar for events on/after the cutoff date.
 */
function syncOneFeed_(cfg, mapping, today) {
  const feedHash = sha256Hex_(mapping.feedUrl).slice(0, 16);
  const feedName = mapping.name || mapping.feedUrl;
  const attendees = uniqueEmails_(
    (mapping.attendeeEmails && mapping.attendeeEmails.length
      ? mapping.attendeeEmails
      : cfg.defaultAttendeeEmails) || [],
  );
  const driveOpts = buildDriveOptions_(cfg, mapping);
  const driveDurationCache = {};
  console.log('[FEED] Processing "' + feedName + '" -> ' + mapping.calendarId);

  const icsText = fetchIcs_(mapping.feedUrl);
  const parsed = parseIcs_(icsText);
  const existingByKey = loadExistingEventsByKey_(mapping.calendarId, feedHash);
  const existingArrivalByKey = loadExistingArrivalEventsByKey_(
    mapping.calendarId,
    feedHash,
  );
  const existingDriveByKey = loadExistingDriveEventsByKey_(
    mapping.calendarId,
    feedHash,
  );
  console.log(
    '[INFO] Feed "' +
      feedName +
      '" has ' +
      parsed.events.length +
      " VEVENT(s); found " +
      Object.keys(existingByKey).length +
      " existing managed event(s)",
  );
  console.log(
    '[INFO] Feed "' +
      feedName +
      '" has ' +
      Object.keys(existingArrivalByKey).length +
      " existing managed arrival placeholder(s)",
  );
  console.log(
    '[INFO] Feed "' +
      feedName +
      '" has ' +
      Object.keys(existingDriveByKey).length +
      " existing managed drive placeholder(s)",
  );

  const seen = {};
  const seenArrival = {};
  const seenDrive = {};
  const stats = {
    feed: feedName,
    created: 0,
    updated: 0,
    deleted: 0,
    unchanged: 0,
    skipped: 0,
    driveCreated: 0,
    driveUpdated: 0,
    driveDeleted: 0,
    driveSkipped: 0,
    arrivalCreated: 0,
    arrivalUpdated: 0,
    arrivalDeleted: 0,
    arrivalSkipped: 0,
  };

  parsed.events.forEach(function (evt) {
    const effectiveEvt = applyEventTitlePrefix_(evt, mapping.titlePrefix);
    const syncKey = buildSyncKey_(feedHash, evt.uid, evt.recurrenceIdKey);
    const arrivalSyncKey = buildArrivalSyncKey_(syncKey);
    const driveSyncKey = buildDriveSyncKey_(syncKey);
    seen[syncKey] = true;
    const existing = existingByKey[syncKey];

    if (evt.cancelled) {
      if (!isEventOnOrAfterCutoff_(evt, today)) {
        stats.skipped++;
        console.info("[SKIP] Not processing cancel for pre-today event");
        return;
      }
      if (existing) {
        if (isManagedEventForFeed_(existing, mapping.feedUrl, feedHash)) {
          Calendar.Events.remove(mapping.calendarId, existing.id);
          stats.deleted++;
          console.log(
            "[DELETE] Deleted canceled event " +
              existing.id +
              " from " +
              feedName,
          );
        } else {
          stats.skipped++;
          console.info(
            "[SKIP] Not deleting non-managed event " +
              existing.id +
              " (cancelled upstream)",
          );
        }
      } else {
        stats.skipped++;
      }
      maybeDeleteArrivalPlaceholder_(
        mapping,
        feedHash,
        arrivalSyncKey,
        existingArrivalByKey,
        today,
        stats,
        "source event canceled",
      );
      maybeDeleteDrivePlaceholder_(
        mapping,
        feedHash,
        driveSyncKey,
        existingDriveByKey,
        today,
        stats,
        "source event canceled",
      );
      return;
    }

    if (!shouldSyncEvent_(effectiveEvt, today)) {
      stats.skipped++;
      console.info(
        '[SKIP] Pre-today event "' +
          (effectiveEvt.summary || "(No title)") +
          '"',
      );
      return;
    }

    const createResource = buildEventResource_(
      effectiveEvt,
      mapping.feedUrl,
      feedHash,
      syncKey,
      attendees,
      parsed.calendarTimezone,
    );

    const newHash = computeEventHash_(effectiveEvt, attendees);
    createResource.extendedProperties.private.syncHash = newHash;

    if (!existing) {
      const inserted = Calendar.Events.insert(
        createResource,
        mapping.calendarId,
        { sendUpdates: "none" },
      );
      stats.created++;
      console.log('[CREATE] "' + (effectiveEvt.summary || "(No title)") + '"');
      const arrivalAnchorStart = reconcileArrivalPlaceholder_(
        effectiveEvt,
        inserted,
        mapping,
        feedHash,
        syncKey,
        arrivalSyncKey,
        existingArrivalByKey,
        seenArrival,
        today,
        stats,
      );
      reconcileDrivePlaceholder_(
        effectiveEvt,
        inserted,
        mapping,
        feedHash,
        syncKey,
        driveSyncKey,
        driveOpts,
        existingDriveByKey,
        seenDrive,
        today,
        stats,
        driveDurationCache,
        arrivalAnchorStart,
      );
      return;
    }

    if (!isManagedEventForFeed_(existing, mapping.feedUrl, feedHash)) {
      stats.skipped++;
      console.info(
        "[SKIP] Not updating event " +
          existing.id +
          " because it is not managed by this feed",
      );
      stats.driveSkipped++;
      console.info(
        "[SKIP] Drive placeholder skipped because source event is unmanaged",
      );
      stats.arrivalSkipped++;
      console.info(
        "[SKIP] Arrival placeholder skipped because source event is unmanaged",
      );
      return;
    }

    const oldHash =
      ((existing.extendedProperties || {}).private || {}).syncHash || "";
    const changedFromLastFeedState = oldHash !== newHash;
    const patchResource = buildEventPatchResource_(
      effectiveEvt,
      mapping.feedUrl,
      feedHash,
      syncKey,
      attendees,
      parsed.calendarTimezone,
    );
    patchResource.extendedProperties.private.syncHash = newHash;
    const patched = Calendar.Events.patch(
      patchResource,
      mapping.calendarId,
      existing.id,
      { sendUpdates: "none" },
    );
    stats.updated++;
    if (changedFromLastFeedState) {
      console.log("[UPDATE] Event " + existing.id + " (feed change detected)");
    } else {
      console.log("[UPDATE] Event " + existing.id + " (forced resync)");
    }
    const arrivalAnchorStart = reconcileArrivalPlaceholder_(
      effectiveEvt,
      patched,
      mapping,
      feedHash,
      syncKey,
      arrivalSyncKey,
      existingArrivalByKey,
      seenArrival,
      today,
      stats,
    );
    reconcileDrivePlaceholder_(
      effectiveEvt,
      patched,
      mapping,
      feedHash,
      syncKey,
      driveSyncKey,
      driveOpts,
      existingDriveByKey,
      seenDrive,
      today,
      stats,
      driveDurationCache,
      arrivalAnchorStart,
    );
  });

  if (cfg.deleteMissingFromFeed) {
    Object.keys(existingByKey).forEach(function (syncKey) {
      if (seen[syncKey]) return;
      const ev = existingByKey[syncKey];
      if (!isManagedEventForFeed_(ev, mapping.feedUrl, feedHash)) {
        console.info("[SKIP] Not deleting non-managed event " + ev.id);
        return;
      }
      if (isFutureEventResource_(ev, today)) {
        Calendar.Events.remove(mapping.calendarId, ev.id);
        stats.deleted++;
        console.log(
          "[DELETE] Deleted feed-missing event " + ev.id + " from " + feedName,
        );
      }
    });

    Object.keys(existingDriveByKey).forEach(function (driveSyncKey) {
      if (seenDrive[driveSyncKey]) return;
      const driveEv = existingDriveByKey[driveSyncKey];
      if (!isManagedDriveEventForFeed_(driveEv, mapping.feedUrl, feedHash)) {
        console.info(
          "[SKIP] Not deleting non-managed drive placeholder " + driveEv.id,
        );
        return;
      }
      if (isFutureEventResource_(driveEv, today)) {
        Calendar.Events.remove(mapping.calendarId, driveEv.id);
        stats.driveDeleted++;
        console.log(
          "[DELETE] Deleted feed-missing drive placeholder " +
            driveEv.id +
            " from " +
            feedName,
        );
      }
    });

    Object.keys(existingArrivalByKey).forEach(function (arrivalSyncKey) {
      if (seenArrival[arrivalSyncKey]) return;
      const arrivalEv = existingArrivalByKey[arrivalSyncKey];
      if (
        !isManagedArrivalEventForFeed_(arrivalEv, mapping.feedUrl, feedHash)
      ) {
        console.info(
          "[SKIP] Not deleting non-managed arrival placeholder " + arrivalEv.id,
        );
        return;
      }
      if (isFutureEventResource_(arrivalEv, today)) {
        Calendar.Events.remove(mapping.calendarId, arrivalEv.id);
        stats.arrivalDeleted++;
        console.log(
          "[DELETE] Deleted feed-missing arrival placeholder " +
            arrivalEv.id +
            " from " +
            feedName,
        );
      }
    });
  }

  console.log(
    '[SUMMARY] Feed "' +
      feedName +
      '": ' +
      "created=" +
      stats.created +
      ", updated=" +
      stats.updated +
      ", deleted=" +
      stats.deleted +
      ", unchanged=" +
      stats.unchanged +
      ", skipped=" +
      stats.skipped +
      ", driveCreated=" +
      stats.driveCreated +
      ", driveUpdated=" +
      stats.driveUpdated +
      ", driveDeleted=" +
      stats.driveDeleted +
      ", driveSkipped=" +
      stats.driveSkipped +
      ", arrivalCreated=" +
      stats.arrivalCreated +
      ", arrivalUpdated=" +
      stats.arrivalUpdated +
      ", arrivalDeleted=" +
      stats.arrivalDeleted +
      ", arrivalSkipped=" +
      stats.arrivalSkipped,
  );
  return stats;
}

/**
 * Reads and validates user config from getIcalSyncConfig(), filling safe defaults.
 */
function getIcalSyncConfig_() {
  if (typeof getIcalSyncConfig !== "function") {
    throw new Error(
      "Missing getIcalSyncConfig(). Create icalFeedSync.config.gs (see icalFeedSync.config.example.gs).",
    );
  }

  const cfg = getIcalSyncConfig();
  if (!cfg || typeof cfg !== "object") {
    throw new Error("getIcalSyncConfig() must return a config object.");
  }
  if (!cfg.triggerEveryMinutes) cfg.triggerEveryMinutes = 15;
  if (typeof cfg.deleteMissingFromFeed !== "boolean")
    cfg.deleteMissingFromFeed = true;
  if (!Array.isArray(cfg.defaultAttendeeEmails)) cfg.defaultAttendeeEmails = [];
  if (typeof cfg.addDriveTimePlaceholders !== "boolean")
    cfg.addDriveTimePlaceholders = false;
  if (typeof cfg.defaultOriginAddress !== "string")
    cfg.defaultOriginAddress = "";
  if (
    typeof cfg.minDriveMinutesToCreate !== "number" ||
    isNaN(cfg.minDriveMinutesToCreate)
  ) {
    cfg.minDriveMinutesToCreate = 10;
  }
  if (cfg.minDriveMinutesToCreate < 1) cfg.minDriveMinutesToCreate = 1;
  if (
    typeof cfg.driveEventTitleTemplate !== "string" ||
    !cfg.driveEventTitleTemplate.trim()
  ) {
    cfg.driveEventTitleTemplate = "Drive to {{title}}";
  }
  if (!Array.isArray(cfg.feedMappings) || !cfg.feedMappings.length) {
    throw new Error("Config feedMappings must be a non-empty array.");
  }

  cfg.feedMappings.forEach(function (m, i) {
    if (!m.feedUrl) throw new Error("feedMappings[" + i + "] missing feedUrl.");
    if (!m.calendarId)
      throw new Error("feedMappings[" + i + "] missing calendarId.");
    if (!Array.isArray(m.attendeeEmails)) m.attendeeEmails = [];
    if (typeof m.titlePrefix !== "string") m.titlePrefix = "";
    if (typeof m.addDriveTimePlaceholders !== "boolean")
      m.addDriveTimePlaceholders = cfg.addDriveTimePlaceholders;
    if (typeof m.originAddress !== "string") m.originAddress = "";
  });

  return cfg;
}

/**
 * Fetches raw ICS text from a remote URL and throws on non-200 responses.
 */
function fetchIcs_(url) {
  const resp = UrlFetchApp.fetch(url, {
    method: "get",
    followRedirects: true,
    muteHttpExceptions: true,
    headers: { "Cache-Control": "no-cache" },
  });
  const code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error("Failed to fetch ICS (" + code + "): " + url);
  }
  return resp.getContentText();
}

/**
 * Parses an ICS document into normalized event objects plus calendar-level timezone.
 */
function parseIcs_(text) {
  const lines = unfoldIcsLines_(text);
  const events = [];
  let inEvent = false;
  let block = [];
  let calendarTimezone = Session.getScriptTimeZone();

  lines.forEach(function (line) {
    const upper = line.toUpperCase();

    if (!inEvent && upper.indexOf("X-WR-TIMEZONE:") === 0) {
      calendarTimezone =
        line.substring(line.indexOf(":") + 1).trim() || calendarTimezone;
      return;
    }

    if (upper === "BEGIN:VEVENT") {
      inEvent = true;
      block = [];
      return;
    }

    if (upper === "END:VEVENT") {
      inEvent = false;
      const evt = parseVEvent_(block, calendarTimezone);
      if (evt) events.push(evt);
      return;
    }

    if (inEvent) block.push(line);
  });

  return { events: events, calendarTimezone: calendarTimezone };
}

/**
 * Parses one VEVENT block into the normalized internal event model.
 */
function parseVEvent_(lines, fallbackTz) {
  const props = {};
  const recurrence = [];

  lines.forEach(function (line) {
    const p = parseIcsLine_(line);
    if (!p) return;

    if (!props[p.name]) props[p.name] = [];
    props[p.name].push(p);

    if (p.name === "RRULE" || p.name === "EXDATE" || p.name === "RDATE") {
      recurrence.push(line);
    }
  });

  const uidProp = firstProp_(props, "UID");
  if (!uidProp) return null;

  const status = (
    (firstProp_(props, "STATUS") || {}).value || ""
  ).toUpperCase();
  const cancelled = status === "CANCELLED";

  const startProp = firstProp_(props, "DTSTART");
  const endProp = firstProp_(props, "DTEND");
  const recIdProp = firstProp_(props, "RECURRENCE-ID");

  const start = startProp ? parseIcsDate_(startProp, fallbackTz) : null;
  let end = endProp ? parseIcsDate_(endProp, fallbackTz) : null;

  if (!cancelled && !start) return null;
  if (!cancelled && start && !end) end = defaultEndFromStart_(start);

  const recurrenceIdKey = recIdProp
    ? recIdProp.value + "|" + (recIdProp.params.TZID || "")
    : "";

  return {
    uid: (uidProp.value || "").trim(),
    recurrenceIdKey: recurrenceIdKey,
    cancelled: cancelled,
    status: status,
    summary: unescapeIcsText_((firstProp_(props, "SUMMARY") || {}).value || ""),
    description: unescapeIcsText_(
      (firstProp_(props, "DESCRIPTION") || {}).value || "",
    ),
    location: unescapeIcsText_(
      (firstProp_(props, "LOCATION") || {}).value || "",
    ),
    start: start,
    end: end,
    recurrence: recurrence,
  };
}

/**
 * Parses a single ICS content line into name/params/value components.
 */
function parseIcsLine_(line) {
  const idx = line.indexOf(":");
  if (idx < 0) return null;

  const left = line.substring(0, idx);
  const value = line.substring(idx + 1);
  const parts = left.split(";");
  const name = parts[0].toUpperCase();
  const params = {};

  for (let i = 1; i < parts.length; i++) {
    const p = parts[i];
    const eq = p.indexOf("=");
    if (eq < 0) {
      params[p.toUpperCase()] = true;
    } else {
      const k = p.substring(0, eq).toUpperCase();
      const v = p.substring(eq + 1).replace(/^"|"$/g, "");
      params[k] = v;
    }
  }

  return { name: name, params: params, value: value };
}

/**
 * Unfolds folded ICS lines (continuation lines starting with space/tab).
 */
function unfoldIcsLines_(text) {
  const raw = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const out = [];

  raw.forEach(function (line) {
    if ((line.indexOf(" ") === 0 || line.indexOf("\t") === 0) && out.length) {
      out[out.length - 1] += line.substring(1);
    } else {
      out.push(line);
    }
  });

  return out;
}

/**
 * Converts ICS date/date-time property values into the internal parsed date shape.
 */
function parseIcsDate_(prop, fallbackTz) {
  const value = (prop.value || "").trim();
  const valueType = (prop.params.VALUE || "").toUpperCase();

  if (valueType === "DATE" || /^\d{8}$/.test(value)) {
    return {
      type: "date",
      date: value.replace(/^(\d{4})(\d{2})(\d{2})$/, "$1-$2-$3"),
    };
  }

  const m = value.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})?(Z)?$/);
  if (!m) return null;

  const sec = m[6] || "00";
  const hasZ = !!m[7];
  const dateTime =
    m[1] +
    "-" +
    m[2] +
    "-" +
    m[3] +
    "T" +
    m[4] +
    ":" +
    m[5] +
    ":" +
    sec +
    (hasZ ? "Z" : "");
  const tzid = prop.params.TZID || fallbackTz || null;

  return {
    type: "dateTime",
    dateTime: dateTime,
    timeZone: hasZ ? null : tzid,
  };
}

/**
 * Produces a fallback end value when DTSTART exists but DTEND is missing.
 */
function defaultEndFromStart_(start) {
  if (start.type === "date") {
    const d = new Date(start.date + "T00:00:00");
    d.setDate(d.getDate() + 1);
    return { type: "date", date: formatYmd_(d) };
  }

  if (start.dateTime.slice(-1) === "Z") {
    const d = new Date(start.dateTime);
    d.setHours(d.getHours() + 1);
    return {
      type: "dateTime",
      dateTime: d.toISOString().replace(".000Z", "Z"),
      timeZone: null,
    };
  }

  const m = start.dateTime.match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/,
  );
  const d = new Date(
    Number(m[1]),
    Number(m[2]) - 1,
    Number(m[3]),
    Number(m[4]),
    Number(m[5]),
    Number(m[6]),
  );
  d.setHours(d.getHours() + 1);

  return {
    type: "dateTime",
    dateTime: formatLocalDateTime_(d),
    timeZone: start.timeZone || null,
  };
}

/**
 * Returns whether an event should be synced given the on/after-cutoff policy.
 */
function shouldSyncEvent_(evt, cutoffDate) {
  if (evt.cancelled) return true;
  if (!evt.start || !evt.end) return false;

  if (evt.recurrence && evt.recurrence.length) {
    return !recurrenceEnded_(evt.recurrence, cutoffDate);
  }

  const end = parsedDateToDate_(evt.end);
  if (!end || isNaN(end.getTime())) return true;
  return end.getTime() >= cutoffDate.getTime();
}

/**
 * Returns true when a recurrence rule has fully ended before the cutoff date.
 */
function recurrenceEnded_(recurrenceLines, cutoffDate) {
  for (let i = 0; i < recurrenceLines.length; i++) {
    const line = recurrenceLines[i];
    if (line.toUpperCase().indexOf("RRULE:") !== 0) continue;

    const rule = line.substring(line.indexOf(":") + 1);
    const parts = rule.split(";");
    for (let j = 0; j < parts.length; j++) {
      if (parts[j].indexOf("UNTIL=") !== 0) continue;
      const untilRaw = parts[j].substring("UNTIL=".length);
      const untilDate = parseUntilDate_(untilRaw);
      if (untilDate && untilDate.getTime() < cutoffDate.getTime()) return true;
    }
  }
  return false;
}

/**
 * Parses RRULE UNTIL values into JavaScript Date values.
 */
function parseUntilDate_(untilRaw) {
  if (/^\d{8}$/.test(untilRaw)) {
    return new Date(
      untilRaw.replace(/^(\d{4})(\d{2})(\d{2})$/, "$1-$2-$3T23:59:59"),
    );
  }
  const m = untilRaw.match(
    /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})?(Z)?$/,
  );
  if (!m) return null;
  const sec = m[6] || "00";
  const iso =
    m[1] +
    "-" +
    m[2] +
    "-" +
    m[3] +
    "T" +
    m[4] +
    ":" +
    m[5] +
    ":" +
    sec +
    (m[7] ? "Z" : "");
  return new Date(iso);
}

/**
 * Builds the Calendar API event resource used for initial event creation.
 */
function buildEventResource_(
  evt,
  feedUrl,
  feedHash,
  syncKey,
  attendees,
  fallbackTz,
) {
  const resource = {
    summary: evt.summary || "(No title)",
    description: addGeneratedByDescription_(evt.description || ""),
    location: evt.location || "",
    start: toGoogleDate_(evt.start, fallbackTz),
    end: toGoogleDate_(evt.end, fallbackTz),
    visibility: "default",
    guestsCanModify: false,
    guestsCanInviteOthers: false,
    guestsCanSeeOtherGuests: true,
    attendees: attendees.map(function (email) {
      return { email: email };
    }),
    extendedProperties: {
      private: {
        managedKind: "source",
        sourceFeed: feedHash,
        sourceUrl: feedUrl,
        sourceUid: evt.uid,
        syncKey: syncKey,
      },
    },
  };

  if (evt.recurrence && evt.recurrence.length) {
    resource.recurrence = evt.recurrence.slice();
  }

  return resource;
}

/**
 * Builds the Calendar API patch resource used to force existing events back to feed state.
 */
function buildEventPatchResource_(
  evt,
  feedUrl,
  feedHash,
  syncKey,
  attendees,
  fallbackTz,
) {
  const resource = {
    summary: evt.summary || "(No title)",
    description: addGeneratedByDescription_(evt.description || ""),
    location: evt.location || "",
    start: toGoogleDate_(evt.start, fallbackTz),
    end: toGoogleDate_(evt.end, fallbackTz),
    visibility: "default",
    guestsCanModify: false,
    guestsCanInviteOthers: false,
    guestsCanSeeOtherGuests: true,
    attendees: attendees.map(function (email) {
      return { email: email };
    }),
    extendedProperties: {
      private: {
        managedKind: "source",
        sourceFeed: feedHash,
        sourceUrl: feedUrl,
        sourceUid: evt.uid,
        syncKey: syncKey,
      },
    },
  };

  // Explicitly include recurrence in patches so existing recurring state is replaced by upstream truth.
  resource.recurrence =
    evt.recurrence && evt.recurrence.length ? evt.recurrence.slice() : [];

  return resource;
}

/**
 * Converts internal parsed date data into Google Calendar API start/end shape.
 */
function toGoogleDate_(parsed, fallbackTz) {
  if (parsed.type === "date") return { date: parsed.date };

  const out = { dateTime: parsed.dateTime };
  if (parsed.dateTime.slice(-1) !== "Z") {
    out.timeZone = parsed.timeZone || fallbackTz || Session.getScriptTimeZone();
  }
  return out;
}

/**
 * Computes a stable hash of feed-driven event fields to track upstream state.
 */
function computeEventHash_(evt, attendees) {
  const normalized = {
    uid: evt.uid,
    recurrenceIdKey: evt.recurrenceIdKey || "",
    summary: evt.summary || "",
    description: evt.description || "",
    location: evt.location || "",
    status: evt.status || "",
    start: evt.start || null,
    end: evt.end || null,
    recurrence: (evt.recurrence || []).slice().sort(),
    attendees: attendees.slice().sort(),
  };
  return sha256Hex_(JSON.stringify(normalized));
}

/**
 * Applies a per-feed title prefix to a feed event title and returns a copied event object.
 */
function applyEventTitlePrefix_(evt, titlePrefix) {
  if (!evt || typeof evt !== "object") return evt;
  const prefix = String(titlePrefix || "").trim();
  if (!prefix) return evt;
  const baseTitle = (evt.summary || "(No title)").trim() || "(No title)";
  const copied = Object.assign({}, evt);
  copied.summary = prefix + " " + baseTitle;
  return copied;
}

/**
 * Builds a deterministic per-feed sync key using UID + recurrence identity.
 */
function buildSyncKey_(feedHash, uid, recurrenceIdKey) {
  const raw = uid + "||" + (recurrenceIdKey || "");
  return feedHash + ":" + sha256Hex_(raw).slice(0, 40);
}

/**
 * Builds the advanced-arrival-placeholder sync key for a given source-event sync key.
 */
function buildArrivalSyncKey_(sourceSyncKey) {
  return "arrival:" + sourceSyncKey;
}

/**
 * Builds the drive-placeholder sync key for a given source-event sync key.
 */
function buildDriveSyncKey_(sourceSyncKey) {
  return "drive:" + sourceSyncKey;
}

/**
 * Builds effective drive-placeholder options from global and per-feed config.
 */
function buildDriveOptions_(cfg, mapping) {
  return {
    enabled: !!mapping.addDriveTimePlaceholders,
    originAddress: (
      mapping.originAddress ||
      cfg.defaultOriginAddress ||
      ""
    ).trim(),
    minDriveMinutesToCreate: cfg.minDriveMinutesToCreate,
    titleTemplate: cfg.driveEventTitleTemplate,
  };
}

/**
 * Loads existing calendar events previously managed by this feed, keyed by syncKey.
 */
function loadExistingEventsByKey_(calendarId, feedHash) {
  const out = {};
  let pageToken;

  do {
    const resp = Calendar.Events.list(calendarId, {
      privateExtendedProperty: ["sourceFeed=" + feedHash],
      showDeleted: false,
      singleEvents: false,
      maxResults: 2500,
      pageToken: pageToken,
    });

    (resp.items || []).forEach(function (ev) {
      const key = ((ev.extendedProperties || {}).private || {}).syncKey || "";
      if (key && !isDriveSyncKey_(key) && !isArrivalSyncKey_(key))
        out[key] = ev;
    });

    pageToken = resp.nextPageToken;
  } while (pageToken);

  return out;
}

/**
 * Loads existing advanced-arrival-placeholder events managed by this feed, keyed by arrival sync key.
 */
function loadExistingArrivalEventsByKey_(calendarId, feedHash) {
  const out = {};
  let pageToken;

  do {
    const resp = Calendar.Events.list(calendarId, {
      privateExtendedProperty: ["sourceFeed=" + feedHash],
      showDeleted: false,
      singleEvents: false,
      maxResults: 2500,
      pageToken: pageToken,
    });

    (resp.items || []).forEach(function (ev) {
      const key = ((ev.extendedProperties || {}).private || {}).syncKey || "";
      if (key && isArrivalSyncKey_(key)) out[key] = ev;
    });

    pageToken = resp.nextPageToken;
  } while (pageToken);

  return out;
}

/**
 * Creates/updates/deletes advanced-arrival placeholders for one synced source event.
 * Returns the arrival start Date when an arrival placeholder should anchor drive-time, else null.
 */
function reconcileArrivalPlaceholder_(
  evt,
  syncedEvent,
  mapping,
  feedHash,
  sourceSyncKey,
  arrivalSyncKey,
  existingArrivalByKey,
  seenArrival,
  today,
  stats,
) {
  const existingArrival = existingArrivalByKey[arrivalSyncKey] || null;

  if (!isEventStartOnOrAfterCutoff_(evt, today)) {
    stats.arrivalSkipped++;
    console.info(
      "[SKIP] Arrival placeholder ignored for pre-today source event",
    );
    return null;
  }

  if (isAllDayEvent_(evt)) {
    stats.arrivalSkipped++;
    console.info("[SKIP] Arrival placeholder ignored for all-day source event");
    maybeDeleteArrivalPlaceholder_(
      mapping,
      feedHash,
      arrivalSyncKey,
      existingArrivalByKey,
      today,
      stats,
      "source event is all-day",
    );
    return null;
  }

  const arrivalMinutes = extractArrivalLeadMinutes_(evt.description);
  if (!arrivalMinutes || arrivalMinutes <= 0) {
    stats.arrivalSkipped++;
    maybeDeleteArrivalPlaceholder_(
      mapping,
      feedHash,
      arrivalSyncKey,
      existingArrivalByKey,
      today,
      stats,
      "source event has no arrival lead-time instruction",
    );
    return null;
  }

  const sourceStart = getSourceEventStartDate_(syncedEvent);
  if (!sourceStart) {
    stats.arrivalSkipped++;
    console.info(
      "[SKIP] Arrival placeholder ignored because source start time is unavailable",
    );
    maybeDeleteArrivalPlaceholder_(
      mapping,
      feedHash,
      arrivalSyncKey,
      existingArrivalByKey,
      today,
      stats,
      "source start time unavailable",
    );
    return null;
  }

  const arrivalEnd = sourceStart;
  const arrivalStart = new Date(
    arrivalEnd.getTime() - arrivalMinutes * 60 * 1000,
  );
  const arrivalTitle = "Advanced arrival for " + (evt.summary || "(No title)");
  const arrivalHash = computeArrivalPlaceholderHash_(
    sourceSyncKey,
    syncedEvent.id,
    arrivalStart,
    arrivalEnd,
    arrivalTitle,
    arrivalMinutes,
  );
  const arrivalResource = buildArrivalPlaceholderResource_(
    mapping,
    feedHash,
    evt,
    sourceSyncKey,
    arrivalSyncKey,
    syncedEvent.id,
    arrivalTitle,
    arrivalStart,
    arrivalEnd,
    arrivalHash,
    arrivalMinutes,
  );
  seenArrival[arrivalSyncKey] = true;

  if (!existingArrival) {
    Calendar.Events.insert(arrivalResource, mapping.calendarId, {
      sendUpdates: "none",
    });
    stats.arrivalCreated++;
    console.log(
      "[CREATE] Arrival placeholder for source event " + syncedEvent.id,
    );
    return arrivalStart;
  }

  if (
    !isManagedArrivalEventForFeed_(existingArrival, mapping.feedUrl, feedHash)
  ) {
    stats.arrivalSkipped++;
    console.info(
      "[SKIP] Not updating unmanaged arrival placeholder " + existingArrival.id,
    );
    return null;
  }

  Calendar.Events.patch(
    arrivalResource,
    mapping.calendarId,
    existingArrival.id,
    {
      sendUpdates: "none",
    },
  );
  stats.arrivalUpdated++;
  console.log(
    "[UPDATE] Arrival placeholder " + existingArrival.id + " (forced resync)",
  );
  return arrivalStart;
}

/**
 * Loads existing drive-placeholder events managed by this feed, keyed by drive sync key.
 */
function loadExistingDriveEventsByKey_(calendarId, feedHash) {
  const out = {};
  let pageToken;

  do {
    const resp = Calendar.Events.list(calendarId, {
      privateExtendedProperty: ["sourceFeed=" + feedHash],
      showDeleted: false,
      singleEvents: false,
      maxResults: 2500,
      pageToken: pageToken,
    });

    (resp.items || []).forEach(function (ev) {
      const key = ((ev.extendedProperties || {}).private || {}).syncKey || "";
      if (key && isDriveSyncKey_(key)) out[key] = ev;
    });

    pageToken = resp.nextPageToken;
  } while (pageToken);

  return out;
}

/**
 * Creates/updates/deletes drive placeholders for one synced source event using configured guardrails.
 */
function reconcileDrivePlaceholder_(
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
  driveDurationCache,
  arrivalAnchorStart,
) {
  const destination = (evt.location || "").trim();
  const existingDrive = existingDriveByKey[driveSyncKey] || null;

  if (!driveOpts.enabled) {
    stats.driveSkipped++;
    maybeDeleteDrivePlaceholder_(
      mapping,
      feedHash,
      driveSyncKey,
      existingDriveByKey,
      today,
      stats,
      "drive placeholders disabled",
    );
    return;
  }

  if (!isEventStartOnOrAfterCutoff_(evt, today)) {
    stats.driveSkipped++;
    console.info("[SKIP] Drive placeholder ignored for pre-today source event");
    return;
  }

  if (isAllDayEvent_(evt)) {
    stats.driveSkipped++;
    console.info("[SKIP] Drive placeholder ignored for all-day source event");
    maybeDeleteDrivePlaceholder_(
      mapping,
      feedHash,
      driveSyncKey,
      existingDriveByKey,
      today,
      stats,
      "source event is all-day",
    );
    return;
  }

  if (!destination) {
    stats.driveSkipped++;
    console.info(
      "[SKIP] Drive placeholder ignored because source event has no location",
    );
    maybeDeleteDrivePlaceholder_(
      mapping,
      feedHash,
      driveSyncKey,
      existingDriveByKey,
      today,
      stats,
      "source event has no location",
    );
    return;
  }

  if (!driveOpts.originAddress) {
    stats.driveSkipped++;
    console.info(
      "[SKIP] Drive placeholder ignored because no default origin address is configured",
    );
    maybeDeleteDrivePlaceholder_(
      mapping,
      feedHash,
      driveSyncKey,
      existingDriveByKey,
      today,
      stats,
      "missing origin address",
    );
    return;
  }

  const sourceStart = getSourceEventStartDate_(syncedEvent);
  if (!sourceStart) {
    stats.driveSkipped++;
    console.info(
      "[SKIP] Drive placeholder ignored because source start time is unavailable",
    );
    maybeDeleteDrivePlaceholder_(
      mapping,
      feedHash,
      driveSyncKey,
      existingDriveByKey,
      today,
      stats,
      "source start time unavailable",
    );
    return;
  }

  const driveMinutes = getDriveMinutes_(
    driveOpts.originAddress,
    destination,
    driveDurationCache,
  );
  if (driveMinutes === null) {
    stats.driveSkipped++;
    console.info(
      "[SKIP] Drive placeholder ignored because route lookup failed",
    );
    maybeDeleteDrivePlaceholder_(
      mapping,
      feedHash,
      driveSyncKey,
      existingDriveByKey,
      today,
      stats,
      "route lookup failed",
    );
    return;
  }

  if (driveMinutes <= driveOpts.minDriveMinutesToCreate) {
    stats.driveSkipped++;
    console.info(
      "[SKIP] Drive placeholder ignored because drive time (" +
        driveMinutes +
        "m) is <= threshold (" +
        driveOpts.minDriveMinutesToCreate +
        "m)",
    );
    maybeDeleteDrivePlaceholder_(
      mapping,
      feedHash,
      driveSyncKey,
      existingDriveByKey,
      today,
      stats,
      "drive time below threshold",
    );
    return;
  }

  const driveEnd =
    arrivalAnchorStart instanceof Date &&
    !isNaN(arrivalAnchorStart.getTime()) &&
    arrivalAnchorStart.getTime() < sourceStart.getTime()
      ? arrivalAnchorStart
      : sourceStart;
  const driveStart = new Date(driveEnd.getTime() - driveMinutes * 60 * 1000);
  const driveTitle = renderDriveEventTitle_(
    driveOpts.titleTemplate,
    evt,
    driveMinutes,
  );
  const driveHash = computeDrivePlaceholderHash_(
    sourceSyncKey,
    syncedEvent.id,
    driveOpts.originAddress,
    destination,
    driveStart,
    driveEnd,
    driveTitle,
  );
  const driveResource = buildDrivePlaceholderResource_(
    mapping,
    feedHash,
    evt,
    sourceSyncKey,
    driveSyncKey,
    syncedEvent.id,
    driveTitle,
    driveStart,
    driveEnd,
    driveHash,
    driveOpts.originAddress,
  );
  seenDrive[driveSyncKey] = true;

  if (!existingDrive) {
    Calendar.Events.insert(driveResource, mapping.calendarId, {
      sendUpdates: "none",
    });
    stats.driveCreated++;
    console.log(
      "[CREATE] Drive placeholder for source event " + syncedEvent.id,
    );
    return;
  }

  if (!isManagedDriveEventForFeed_(existingDrive, mapping.feedUrl, feedHash)) {
    stats.driveSkipped++;
    console.info(
      "[SKIP] Not updating unmanaged drive placeholder " + existingDrive.id,
    );
    return;
  }

  Calendar.Events.patch(driveResource, mapping.calendarId, existingDrive.id, {
    sendUpdates: "none",
  });
  stats.driveUpdated++;
  console.log(
    "[UPDATE] Drive placeholder " + existingDrive.id + " (forced resync)",
  );
}

/**
 * Deletes a drive placeholder when it exists, is managed by this feed, and is on/after cutoff.
 */
function maybeDeleteDrivePlaceholder_(
  mapping,
  feedHash,
  driveSyncKey,
  existingDriveByKey,
  today,
  stats,
  reason,
) {
  const existingDrive = existingDriveByKey[driveSyncKey];
  if (!existingDrive) return;
  if (!isManagedDriveEventForFeed_(existingDrive, mapping.feedUrl, feedHash))
    return;
  if (!isFutureEventResource_(existingDrive, today)) return;

  Calendar.Events.remove(mapping.calendarId, existingDrive.id);
  delete existingDriveByKey[driveSyncKey];
  stats.driveDeleted++;
  console.log(
    "[DELETE] Drive placeholder " + existingDrive.id + " (" + reason + ")",
  );
}

/**
 * Deletes an arrival placeholder when it exists, is managed by this feed, and is on/after cutoff.
 */
function maybeDeleteArrivalPlaceholder_(
  mapping,
  feedHash,
  arrivalSyncKey,
  existingArrivalByKey,
  today,
  stats,
  reason,
) {
  const existingArrival = existingArrivalByKey[arrivalSyncKey];
  if (!existingArrival) return;
  if (
    !isManagedArrivalEventForFeed_(existingArrival, mapping.feedUrl, feedHash)
  )
    return;
  if (!isFutureEventResource_(existingArrival, today)) return;

  Calendar.Events.remove(mapping.calendarId, existingArrival.id);
  delete existingArrivalByKey[arrivalSyncKey];
  stats.arrivalDeleted++;
  console.log(
    "[DELETE] Arrival placeholder " + existingArrival.id + " (" + reason + ")",
  );
}

/**
 * Verifies an event is owned by this script for this specific feed mapping.
 */
function isManagedEventForFeed_(ev, feedUrl, feedHash) {
  const p = (ev.extendedProperties || {}).private || {};
  if (!p.syncKey || typeof p.syncKey !== "string") return false;
  if (p.syncKey.indexOf(feedHash + ":") !== 0) return false;
  if (p.managedKind && p.managedKind !== "source") return false;
  if (p.sourceFeed !== feedHash) return false;
  if (p.sourceUrl !== feedUrl) return false;
  if (!p.sourceUid) return false;
  return true;
}

/**
 * Verifies an arrival placeholder is managed by this script for this feed.
 */
function isManagedArrivalEventForFeed_(ev, feedUrl, feedHash) {
  const p = (ev.extendedProperties || {}).private || {};
  if (!p.syncKey || typeof p.syncKey !== "string") return false;
  if (!isArrivalSyncKey_(p.syncKey)) return false;
  if (p.managedKind && p.managedKind !== "arrival") return false;
  if (p.sourceFeed !== feedHash) return false;
  if (p.sourceUrl !== feedUrl) return false;
  if (!p.sourceUid) return false;
  if (!p.sourceSyncKey) return false;
  return true;
}

/**
 * Verifies a drive placeholder is managed by this script for this feed.
 */
function isManagedDriveEventForFeed_(ev, feedUrl, feedHash) {
  const p = (ev.extendedProperties || {}).private || {};
  if (!p.syncKey || typeof p.syncKey !== "string") return false;
  if (!isDriveSyncKey_(p.syncKey)) return false;
  if (p.managedKind && p.managedKind !== "drive") return false;
  if (p.sourceFeed !== feedHash) return false;
  if (p.sourceUrl !== feedUrl) return false;
  if (!p.sourceUid) return false;
  if (!p.sourceSyncKey) return false;
  return true;
}

/**
 * Returns true when a sync key corresponds to an arrival placeholder.
 */
function isArrivalSyncKey_(syncKey) {
  return typeof syncKey === "string" && syncKey.indexOf("arrival:") === 0;
}

/**
 * Returns true when a sync key corresponds to a drive placeholder.
 */
function isDriveSyncKey_(syncKey) {
  return typeof syncKey === "string" && syncKey.indexOf("drive:") === 0;
}

/**
 * Returns whether an existing Calendar API event is on/after the cutoff date.
 */
function isFutureEventResource_(ev, cutoffDate) {
  if (ev.recurrence && ev.recurrence.length) {
    return !recurrenceEnded_(ev.recurrence, cutoffDate);
  }

  const end = ev.end || ev.start || null;
  if (!end) return false;

  let when;
  if (end.dateTime) when = new Date(end.dateTime);
  if (end.date) when = new Date(end.date + "T00:00:00");

  if (!when || isNaN(when.getTime())) return true;
  return when.getTime() >= cutoffDate.getTime();
}

/**
 * Returns whether a parsed feed event is on/after the cutoff date.
 */
function isEventOnOrAfterCutoff_(evt, cutoffDate) {
  if (!evt || (!evt.start && !evt.end)) return false;
  const anchor = parsedDateToDate_(evt.end || evt.start);
  if (!anchor || isNaN(anchor.getTime())) return false;
  return anchor.getTime() >= cutoffDate.getTime();
}

/**
 * Returns whether the source event start time is on/after the cutoff date.
 */
function isEventStartOnOrAfterCutoff_(evt, cutoffDate) {
  if (!evt || !evt.start) return false;
  const startDate = parsedDateToDate_(evt.start);
  if (!startDate || isNaN(startDate.getTime())) return false;
  return startDate.getTime() >= cutoffDate.getTime();
}

/**
 * Returns true when the parsed source event represents an all-day event.
 */
function isAllDayEvent_(evt) {
  return !!(evt && evt.start && evt.start.type === "date");
}

/**
 * Extracts a JavaScript Date for the source calendar event start dateTime.
 */
function getSourceEventStartDate_(sourceEvent) {
  if (!sourceEvent || !sourceEvent.start || !sourceEvent.start.dateTime)
    return null;
  const d = new Date(sourceEvent.start.dateTime);
  if (isNaN(d.getTime())) return null;
  return d;
}

/**
 * Computes drive minutes from origin to destination using Maps and per-run in-memory cache.
 */
function getDriveMinutes_(originAddress, destinationAddress, cache) {
  const key = (originAddress + "||" + destinationAddress).toLowerCase();
  if (cache.hasOwnProperty(key)) return cache[key];

  try {
    const directions = Maps.newDirectionFinder()
      .setOrigin(originAddress)
      .setDestination(destinationAddress)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();
    const minutes = extractDriveMinutesFromDirections_(directions);
    cache[key] = minutes;
    return minutes;
  } catch (e) {
    console.error(
      '[ERROR] Drive lookup failed from "' +
        originAddress +
        '" to "' +
        destinationAddress +
        '": ' +
        String(e),
    );
    cache[key] = null;
    return null;
  }
}

/**
 * Extracts route duration (minutes) from a Maps Directions response.
 */
function extractDriveMinutesFromDirections_(directions) {
  const routes = directions && directions.routes ? directions.routes : [];
  if (!routes.length || !routes[0].legs || !routes[0].legs.length) return null;
  const leg = routes[0].legs[0];
  const duration = leg && leg.duration ? leg.duration : null;
  if (!duration || typeof duration.value !== "number") return null;
  const minutes = Math.ceil(duration.value / 60);
  return roundUpMinutesToNearestFifteen_(minutes);
}

/**
 * Rounds a minute count up to the nearest 15-minute bucket.
 */
function roundUpMinutesToNearestFifteen_(minutes) {
  if (typeof minutes !== "number" || !isFinite(minutes)) return null;
  if (minutes <= 0) return 0;
  return Math.ceil(minutes / 15) * 15;
}

/**
 * Extracts lead-time minutes from event descriptions like:
 * "Arrival: 30 minutes in advance"
 */
function extractArrivalLeadMinutes_(description) {
  const text = String(description || "");
  const m = text.match(
    /(?:^|\n)\s*Arrival:\s*(\d+)\s*minutes?\s*in\s*advance\b/i,
  );
  if (!m) return null;
  const minutes = Number(m[1]);
  if (!isFinite(minutes) || minutes <= 0) return null;
  return Math.floor(minutes);
}

/**
 * Renders drive placeholder title from a template and source event details.
 */
function renderDriveEventTitle_(template, evt, driveMinutes) {
  return String(template || "Drive to {{title}}")
    .replace(/{{\s*title\s*}}/g, evt.summary || "(No title)")
    .replace(/{{\s*location\s*}}/g, evt.location || "")
    .replace(/{{\s*minutes\s*}}/g, String(driveMinutes));
}

/**
 * Builds the Calendar API resource for a managed drive placeholder event.
 */
function buildDrivePlaceholderResource_(
  mapping,
  feedHash,
  evt,
  sourceSyncKey,
  driveSyncKey,
  sourceEventId,
  driveTitle,
  driveStart,
  driveEnd,
  driveHash,
  originAddress,
) {
  const driveDescription =
    "Managed drive-time placeholder.\n" +
    "From: " +
    originAddress +
    "\n" +
    "To: " +
    (evt.location || "") +
    "\n" +
    "Source event: " +
    sourceEventId;

  return {
    summary: driveTitle,
    description: addGeneratedByDescription_(driveDescription),
    location: evt.location || "",
    start: { dateTime: driveStart.toISOString() },
    end: { dateTime: driveEnd.toISOString() },
    visibility: "default",
    guestsCanModify: false,
    guestsCanInviteOthers: false,
    guestsCanSeeOtherGuests: true,
    extendedProperties: {
      private: {
        managedKind: "drive",
        sourceFeed: feedHash,
        sourceUrl: mapping.feedUrl,
        sourceUid: evt.uid,
        syncKey: driveSyncKey,
        sourceSyncKey: sourceSyncKey,
        sourceEventId: sourceEventId,
        syncHash: driveHash,
      },
    },
  };
}

/**
 * Builds the Calendar API resource for a managed advanced-arrival placeholder event.
 */
function buildArrivalPlaceholderResource_(
  mapping,
  feedHash,
  evt,
  sourceSyncKey,
  arrivalSyncKey,
  sourceEventId,
  arrivalTitle,
  arrivalStart,
  arrivalEnd,
  arrivalHash,
  arrivalMinutes,
) {
  const arrivalDescription =
    "Managed advanced-arrival placeholder.\n" +
    "Lead time: " +
    arrivalMinutes +
    " minutes\n" +
    "Source event: " +
    sourceEventId;

  return {
    summary: arrivalTitle,
    description: addGeneratedByDescription_(arrivalDescription),
    location: evt.location || "",
    start: { dateTime: arrivalStart.toISOString() },
    end: { dateTime: arrivalEnd.toISOString() },
    visibility: "default",
    guestsCanModify: false,
    guestsCanInviteOthers: false,
    guestsCanSeeOtherGuests: true,
    extendedProperties: {
      private: {
        managedKind: "arrival",
        sourceFeed: feedHash,
        sourceUrl: mapping.feedUrl,
        sourceUid: evt.uid,
        syncKey: arrivalSyncKey,
        sourceSyncKey: sourceSyncKey,
        sourceEventId: sourceEventId,
        arrivalMinutes: String(arrivalMinutes),
        syncHash: arrivalHash,
      },
    },
  };
}

/**
 * Computes a stable hash for drive placeholder fields so source linkage is explicit.
 */
function computeDrivePlaceholderHash_(
  sourceSyncKey,
  sourceEventId,
  originAddress,
  destinationAddress,
  driveStart,
  driveEnd,
  driveTitle,
) {
  return sha256Hex_(
    JSON.stringify({
      sourceSyncKey: sourceSyncKey,
      sourceEventId: sourceEventId,
      originAddress: originAddress,
      destinationAddress: destinationAddress,
      driveStart: driveStart.toISOString(),
      driveEnd: driveEnd.toISOString(),
      driveTitle: driveTitle,
    }),
  );
}

/**
 * Computes a stable hash for arrival placeholder fields so source linkage is explicit.
 */
function computeArrivalPlaceholderHash_(
  sourceSyncKey,
  sourceEventId,
  arrivalStart,
  arrivalEnd,
  arrivalTitle,
  arrivalMinutes,
) {
  return sha256Hex_(
    JSON.stringify({
      sourceSyncKey: sourceSyncKey,
      sourceEventId: sourceEventId,
      arrivalStart: arrivalStart.toISOString(),
      arrivalEnd: arrivalEnd.toISOString(),
      arrivalTitle: arrivalTitle,
      arrivalMinutes: arrivalMinutes,
    }),
  );
}

/**
 * Appends the generated-by attribution line to an event description if missing.
 */
function addGeneratedByDescription_(description) {
  const text = String(description || "");
  if (text.indexOf(GENERATED_BY_DESCRIPTION) >= 0) return text;
  if (!text.trim()) return GENERATED_BY_DESCRIPTION;
  return text + "\n\n" + GENERATED_BY_DESCRIPTION;
}

/**
 * Converts the internal parsed date shape into a JavaScript Date for comparisons.
 */
function parsedDateToDate_(parsed) {
  if (!parsed) return null;
  if (parsed.type === "date") return new Date(parsed.date + "T00:00:00");
  return new Date(parsed.dateTime);
}

/**
 * Normalizes/deduplicates attendee emails and returns a unique lowercase list.
 */
function uniqueEmails_(emails) {
  const s = {};
  emails
    .map(function (e) {
      return (e || "").trim().toLowerCase();
    })
    .filter(function (e) {
      return !!e;
    })
    .forEach(function (e) {
      s[e] = true;
    });
  return Object.keys(s);
}

/**
 * Returns the first property entry for an ICS property name, or null.
 */
function firstProp_(props, key) {
  const arr = props[key];
  return arr && arr.length ? arr[0] : null;
}

/**
 * Unescapes ICS text escape sequences (\\n, \\, \\;, etc.).
 */
function unescapeIcsText_(s) {
  return (s || "")
    .replace(/\\\\n/gi, "\n")
    .replace(/\\\\r/gi, "\r")
    .replace(/\\\\/g, "\u0000")
    .replace(/\\n/gi, "\n")
    .replace(/\\r/gi, "\r")
    .replace(/\\,/g, ",")
    .replace(/\\;/g, ";")
    .replace(/\u0000/g, "\\");
}

/**
 * Returns a hex SHA-256 digest for the provided input string.
 */
function sha256Hex_(input) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    input,
    Utilities.Charset.UTF_8,
  );
  return bytes
    .map(function (b) {
      const v = (b + 256) % 256;
      return (v < 16 ? "0" : "") + v.toString(16);
    })
    .join("");
}

/**
 * Formats a Date into YYYY-MM-DD (local clock).
 */
function formatYmd_(d) {
  const p = function (n) {
    return n < 10 ? "0" + n : "" + n;
  };
  return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate());
}

/**
 * Formats a Date into YYYY-MM-DDTHH:mm:ss (local clock, no timezone suffix).
 */
function formatLocalDateTime_(d) {
  const p = function (n) {
    return n < 10 ? "0" + n : "" + n;
  };
  return (
    d.getFullYear() +
    "-" +
    p(d.getMonth() + 1) +
    "-" +
    p(d.getDate()) +
    "T" +
    p(d.getHours()) +
    ":" +
    p(d.getMinutes()) +
    ":" +
    p(d.getSeconds())
  );
}

/**
 * Returns local midnight for "today" to use as sync cutoff.
 */
function startOfToday_() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  return now;
}

/**
 * One-time setup helper: logs all accessible calendars on first sync run.
 */
function logCalendarIdsOnFirstRun_() {
  const key = "icalSync.calendarIdsLoggedOnce";
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(key) === "1") return;
  logAllCalendarIds_();
  props.setProperty(key, "1");
}

/**
 * Logs all accessible calendars as "name => id" for configuration lookup.
 */
function logAllCalendarIds_() {
  const calendars = CalendarApp.getAllCalendars();
  console.log("[SETUP] Accessible calendars: " + calendars.length);
  calendars.forEach(function (cal) {
    console.log("[SETUP] " + cal.getName() + " => " + cal.getId());
  });
}
