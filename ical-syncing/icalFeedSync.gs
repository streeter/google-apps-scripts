/**
 * iCal -> Google Calendar sync (future events), with updates and optional attendees.
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
const MANAGED_CALENDAR_IDS_PROPERTY = "icalSync.managedCalendarIds";
const DRIVE_LOOKBACK_MINUTES = 60;
const CALENDAR_WRITE_MAX_ATTEMPTS = 5;
const CALENDAR_WRITE_BASE_SLEEP_MS = 500;
const SYNC_LOCK_WAIT_MS = 1000;
const DEFAULT_TIMED_EVENT_DURATION_MINUTES = 30;
const CALENDAR_WRITE_DIAGNOSTICS = {
  started: 0,
  succeeded: 0,
};

/**
 * Replaces the time-based triggers for the main sync function.
 * An explicitly empty triggerHours array removes all sync triggers.
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

  if (Array.isArray(cfg.triggerHours)) {
    if (cfg.triggerHours.length) {
      createDailyHourTriggers_(fn, cfg.triggerHours);
    }
    return;
  }

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
 * Creates one daily trigger per hour in the script timezone.
 */
function createDailyHourTriggers_(handlerFunction, triggerHours) {
  normalizeTriggerHours_(triggerHours).forEach(function (hour) {
    ScriptApp.newTrigger(handlerFunction)
      .timeBased()
      .atHour(hour)
      .nearMinute(0)
      .everyDays(1)
      .create();
  });
}

/**
 * Validates, de-duplicates, and sorts configured trigger hours.
 */
function normalizeTriggerHours_(triggerHours) {
  if (!Array.isArray(triggerHours)) {
    throw new Error("triggerHours must be an array when provided.");
  }

  const seen = {};
  const normalized = [];

  triggerHours.forEach(function (hour) {
    const value = Number(hour);
    if (
      !isFinite(value) ||
      value !== Math.floor(value) ||
      value < 0 ||
      value > 23
    ) {
      throw new Error(
        "triggerHours values must be integers from 0 through 23.",
      );
    }
    if (seen[value]) return;
    seen[value] = true;
    normalized.push(value);
  });

  normalized.sort(function (a, b) {
    return a - b;
  });
  return normalized;
}

/**
 * Main entry point: logs setup info, loads config, and syncs each feed mapping.
 */
function syncIcalFeeds() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(SYNC_LOCK_WAIT_MS)) {
    console.warn("[SKIP] iCal feed sync is already in progress");
    return [];
  }

  try {
    return syncIcalFeedsUnlocked_();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Performs one complete sync while the caller holds the script lock.
 */
function syncIcalFeedsUnlocked_() {
  logCalendarIdsOnFirstRun_();
  const cfg = getIcalSyncConfig_();
  const today = startOfToday_();
  const results = [];
  const errors = [];
  let calendarUsageLimitReached = false;
  console.log(
    "[SYNC] Starting iCal feed sync for " +
      cfg.feedMappings.length +
      " feed(s)",
  );
  console.log("[SYNC] Date cutoff (inclusive): " + today.toISOString());

  for (let i = 0; i < cfg.feedMappings.length; i++) {
    const mapping = cfg.feedMappings[i];
    try {
      results.push(syncOneFeed_(cfg, mapping, today));
    } catch (e) {
      const feedName = mapping.name || mapping.feedUrl;
      const errorText = String(e);
      console.error(
        "[ERROR] Failed syncing feed " + feedName + ": " + errorText,
      );
      results.push({
        feed: feedName,
        error: errorText,
      });
      errors.push(feedName + ": " + errorText);
      if (isCalendarUsageLimitError_(e)) {
        calendarUsageLimitReached = true;
        console.error(
          "[SYNC_ABORT] Calendar usage limit reached; skipping " +
            (cfg.feedMappings.length - i - 1) +
            " remaining feed(s) and removed-feed cleanup",
        );
        break;
      }
    }
  }

  if (cfg.deleteMissingFromFeed && !calendarUsageLimitReached) {
    try {
      cleanupRemovedFeedEvents_(cfg, today);
    } catch (e) {
      const errorText = String(e);
      console.error(
        "[ERROR] Failed cleaning up removed feed events: " + errorText,
      );
      errors.push("Removed feed cleanup: " + errorText);
    }
  }
  rememberManagedCalendarIds_(cfg);

  console.log("[SYNC] Finished iCal feed sync");
  Logger.log(JSON.stringify(results, null, 2));
  if (errors.length) {
    const groupedErrors = groupSyncErrors_(errors);
    throw new Error(
      "syncIcalFeeds completed with " +
        groupedErrors.length +
        " error(s): " +
        groupedErrors.join(" | "),
    );
  }
  return results;
}

function groupSyncErrors_(errors) {
  const grouped = {};
  const orderedMessages = [];

  errors.forEach(function (entry) {
    const separator = ": ";
    const splitIndex = entry.indexOf(separator);
    if (splitIndex === -1) {
      if (!grouped[entry]) {
        grouped[entry] = [];
        orderedMessages.push(entry);
      }
      return;
    }

    const feedName = entry.slice(0, splitIndex);
    const errorText = entry.slice(splitIndex + separator.length);
    if (!grouped[errorText]) {
      grouped[errorText] = [];
      orderedMessages.push(errorText);
    }
    grouped[errorText].push(feedName);
  });

  return orderedMessages.map(function (message) {
    const feedNames = grouped[message];
    if (!feedNames || feedNames.length === 0) {
      return message;
    }
    return feedNames.join(", ") + ": " + message;
  });
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
  const hasMappingAttendeeEmails = Object.prototype.hasOwnProperty.call(
    mapping,
    "attendeeEmails",
  );
  const attendees = uniqueEmails_(
    (hasMappingAttendeeEmails
      ? mapping.attendeeEmails
      : cfg.defaultAttendeeEmails) || [],
  );
  const sourceAttendees = buildSourceAttendees_(
    attendees,
    mapping.calendarId,
    mapping.addDestinationCalendarAsAttendee,
  );
  const activePeerFeedHashes = buildActivePeerFeedHashes_(
    cfg,
    mapping.calendarId,
    feedHash,
  );
  const driveOpts = buildDriveOptions_(cfg, mapping);
  const driveDurationCache = {};
  console.log('[FEED] Processing "' + feedName + '" -> ' + mapping.calendarId);

  const icsText = fetchIcs_(mapping.feedUrl);
  const parsed = parseIcs_(icsText, mapping.timeZone);
  const activeSourceSyncKeys = {};
  parsed.events.forEach(function (evt) {
    if (evt.cancelled) return;
    activeSourceSyncKeys[
      buildSyncKey_(feedHash, evt.uid, evt.recurrenceIdKey)
    ] = true;
  });
  const existingByKey = loadExistingEventsByKey_(
    mapping.calendarId,
    feedHash,
    feedName,
  );
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
    const effectiveEvt = applyPlaceNameAddressToEvent_(
      applyEventTitlePrefix_(evt, mapping.titlePrefix),
      driveOpts.placeNameAddressRules,
      mapping,
    );
    if (mapping.skipAllDayEvents && isAllDayEvent_(effectiveEvt)) {
      stats.skipped++;
      console.info(
        "[SKIP] " +
          formatEventLogContext_(
            effectiveEvt,
            mapping.calendarId,
            feedName,
            "source event",
          ) +
          " — filtered by skipAllDayEvents",
      );
      return;
    }
    const syncKey = buildSyncKey_(feedHash, evt.uid, evt.recurrenceIdKey);
    const arrivalSyncKey = buildArrivalSyncKey_(syncKey);
    const driveSyncKey = buildDriveSyncKey_(syncKey);
    seen[syncKey] = true;
    let existing = existingByKey[syncKey];

    if (evt.cancelled) {
      if (!isEventOnOrAfterCutoff_(evt, today)) {
        stats.skipped++;
        console.info(
          "[SKIP] " +
            formatEventLogContext_(
              evt,
              mapping.calendarId,
              feedName,
              "source event",
            ) +
            " — canceled upstream, but event is before the sync cutoff",
        );
        return;
      }
      if (existing) {
        if (isManagedEventForFeed_(existing, mapping.feedUrl, feedHash)) {
          calendarEventRemove_(mapping.calendarId, existing.id, existing);
          stats.deleted++;
          console.log(
            "[DELETE] " +
              formatEventLogContext_(
                existing,
                mapping.calendarId,
                feedName,
                "source event",
              ) +
              " — canceled upstream",
          );
        } else {
          stats.skipped++;
          console.info(
            "[SKIP] " +
              formatEventLogContext_(
                existing,
                mapping.calendarId,
                feedName,
                "source event",
              ) +
              " — canceled upstream, but event is not managed by this feed",
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
        "[SKIP] " +
          formatEventLogContext_(
            effectiveEvt,
            mapping.calendarId,
            feedName,
            "source event",
          ) +
          " — event is before the sync cutoff",
      );
      return;
    }

    const createResource = buildEventResource_(
      effectiveEvt,
      mapping.feedUrl,
      feedHash,
      syncKey,
      feedName,
      sourceAttendees,
      parsed.calendarTimezone,
    );

    const createHash = computeEventHash_(effectiveEvt, sourceAttendees);
    createResource.extendedProperties.private.syncHash = createHash;

    if (!existing) {
      const duplicateResolution = resolveDuplicateBeforeCreate_(
        mapping.calendarId,
        createResource,
        activePeerFeedHashes,
        activeSourceSyncKeys,
      );
      if (duplicateResolution.deleted)
        stats.deleted += duplicateResolution.deleted;
      if (duplicateResolution.adoptExisting) {
        existing = duplicateResolution.adoptExisting;
        const existingPrivateProps =
          ((existing || {}).extendedProperties || {}).private || {};
        const previousSyncKey = existingPrivateProps.syncKey || "";
        reindexExistingManagedEvent_(
          existingByKey,
          previousSyncKey,
          syncKey,
          existing,
        );
        reindexExistingManagedEvent_(
          existingArrivalByKey,
          buildArrivalSyncKey_(previousSyncKey),
          arrivalSyncKey,
        );
        reindexExistingManagedEvent_(
          existingDriveByKey,
          buildDriveSyncKey_(previousSyncKey),
          driveSyncKey,
        );
        console.info(
          "[ADOPT] " +
            formatEventLogContext_(
              effectiveEvt,
              mapping.calendarId,
              feedName,
              "source event",
            ) +
            " — adopting exact same-feed event after " +
            (existingPrivateProps.sourceUid !== evt.uid
              ? "upstream UID changed"
              : "upstream recurrence identity changed"),
        );
      } else {
        if (duplicateResolution.skipCreate) {
          stats.skipped++;
          console.info(
            "[SKIP] " +
              formatEventLogContext_(
                effectiveEvt,
                mapping.calendarId,
                feedName,
                "source event",
              ) +
              " — actively synced peer event already exists",
          );
          return;
        }

        let inserted;
        try {
          inserted = calendarEventInsert_(createResource, mapping.calendarId, {
            sendUpdates: "none",
          });
        } catch (e) {
          console.error(
            "[ERROR] " +
              formatEventLogContext_(
                effectiveEvt,
                mapping.calendarId,
                feedName,
                "source event",
              ) +
              " — create failed: " +
              String(e),
          );
          throw e;
        }
        stats.created++;
        console.log(
          "[CREATE] " +
            formatEventLogContext_(
              effectiveEvt,
              mapping.calendarId,
              feedName,
              "source event",
            ) +
            " — new feed event",
        );
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
          sourceAttendees,
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
          sourceAttendees,
        );
        return;
      }
    }

    if (!isManagedEventForFeed_(existing, mapping.feedUrl, feedHash)) {
      stats.skipped++;
      console.info(
        "[SKIP] " +
          formatEventLogContext_(
            existing,
            mapping.calendarId,
            feedName,
            "source event",
          ) +
          " — event is not managed by this feed; source and placeholders were not updated",
      );
      stats.driveSkipped++;
      stats.arrivalSkipped++;
      return;
    }

    const targetCalendarDeclined = isTargetCalendarDeclinedEvent_(
      existing,
      mapping.calendarId,
    );
    const destinationDeclined =
      targetCalendarDeclined || isAllAttendeesDeclinedEvent_(existing);
    const patchAttendees = targetCalendarDeclined
      ? buildDeclinedAttendees_(existing, mapping.calendarId)
      : resolveAttendeesForExistingManagedEvent_(
          sourceAttendees,
          existing,
          mapping.addDestinationCalendarAsAttendee,
        );
    const patchHashAttendees = targetCalendarDeclined
      ? patchAttendees
      : getAttendeeEmails_(patchAttendees);
    const oldHash =
      ((existing.extendedProperties || {}).private || {}).syncHash || "";
    const patchHash = computeEventHash_(effectiveEvt, patchHashAttendees);
    const changedFromLastFeedState = oldHash !== patchHash;
    const patchResource = buildEventPatchResource_(
      effectiveEvt,
      mapping.feedUrl,
      feedHash,
      syncKey,
      feedName,
      patchAttendees,
      parsed.calendarTimezone,
    );
    patchResource.extendedProperties.private.syncHash = patchHash;
    const changedFromDestinationState = !eventResourceTimingMatches_(
      existing,
      patchResource,
    );
    if (!changedFromLastFeedState && !changedFromDestinationState) {
      stats.unchanged++;
      console.log(
        "[UNCHANGED] " +
          formatEventLogContext_(
            effectiveEvt,
            mapping.calendarId,
            feedName,
            "source event",
          ) +
          " — no feed or destination changes detected",
      );
      if (destinationDeclined) {
        console.info(
          "[DECLINE] " +
            formatEventLogContext_(
              effectiveEvt,
              mapping.calendarId,
              feedName,
              "source event",
            ) +
            " — preserving local decline and removing managed placeholders",
        );
        maybeDeleteArrivalPlaceholder_(
          mapping,
          feedHash,
          arrivalSyncKey,
          existingArrivalByKey,
          today,
          stats,
          "source event declined",
        );
        maybeDeleteDrivePlaceholder_(
          mapping,
          feedHash,
          driveSyncKey,
          existingDriveByKey,
          today,
          stats,
          "source event declined",
        );
        return;
      }
      const unchangedArrivalAnchorStart = reconcileArrivalPlaceholder_(
        effectiveEvt,
        existing,
        mapping,
        feedHash,
        syncKey,
        arrivalSyncKey,
        existingArrivalByKey,
        seenArrival,
        today,
        stats,
        sourceAttendees,
      );
      reconcileDrivePlaceholder_(
        effectiveEvt,
        existing,
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
        unchangedArrivalAnchorStart,
        sourceAttendees,
      );
      return;
    }
    const patched = calendarEventPatch_(
      patchResource,
      mapping.calendarId,
      existing.id,
      { sendUpdates: "none" },
    );
    stats.updated++;
    console.log(
      "[UPDATE] " +
        formatEventLogContext_(
          effectiveEvt,
          mapping.calendarId,
          feedName,
          "source event",
        ) +
        " — " +
        (changedFromLastFeedState
          ? "feed change detected"
          : "destination event time drift detected"),
    );
    if (destinationDeclined) {
      console.info(
        "[DECLINE] " +
          formatEventLogContext_(
            effectiveEvt,
            mapping.calendarId,
            feedName,
            "source event",
          ) +
          " — preserving local decline and removing managed placeholders",
      );
      maybeDeleteArrivalPlaceholder_(
        mapping,
        feedHash,
        arrivalSyncKey,
        existingArrivalByKey,
        today,
        stats,
        "source event declined",
      );
      maybeDeleteDrivePlaceholder_(
        mapping,
        feedHash,
        driveSyncKey,
        existingDriveByKey,
        today,
        stats,
        "source event declined",
      );
      return;
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
      sourceAttendees,
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
      sourceAttendees,
    );
  });

  if (cfg.deleteMissingFromFeed) {
    Object.keys(existingByKey).forEach(function (syncKey) {
      if (seen[syncKey]) return;
      const ev = existingByKey[syncKey];
      if (!isManagedEventForFeed_(ev, mapping.feedUrl, feedHash)) {
        console.info(
          "[SKIP] " +
            formatEventLogContext_(
              ev,
              mapping.calendarId,
              feedName,
              "source event",
            ) +
            " — missing from feed, but event is not managed by this feed",
        );
        return;
      }
      if (isFutureEventResource_(ev, today)) {
        calendarEventRemove_(mapping.calendarId, ev.id, ev);
        stats.deleted++;
        console.log(
          "[DELETE] " +
            formatEventLogContext_(
              ev,
              mapping.calendarId,
              feedName,
              "source event",
            ) +
            " — missing from feed",
        );
      }
    });

    Object.keys(existingDriveByKey).forEach(function (driveSyncKey) {
      if (seenDrive[driveSyncKey]) return;
      const driveEv = existingDriveByKey[driveSyncKey];
      if (!isManagedDriveEventForFeed_(driveEv, mapping.feedUrl, feedHash)) {
        console.info(
          "[SKIP] " +
            formatEventLogContext_(
              driveEv,
              mapping.calendarId,
              feedName,
              "drive placeholder",
            ) +
            " — missing from feed, but placeholder is not managed by this feed",
        );
        return;
      }
      if (isFutureEventResource_(driveEv, today)) {
        calendarEventRemove_(mapping.calendarId, driveEv.id, driveEv);
        stats.driveDeleted++;
        console.log(
          "[DELETE] " +
            formatEventLogContext_(
              driveEv,
              mapping.calendarId,
              feedName,
              "drive placeholder",
            ) +
            " — missing from feed",
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
          "[SKIP] " +
            formatEventLogContext_(
              arrivalEv,
              mapping.calendarId,
              feedName,
              "arrival placeholder",
            ) +
            " — missing from feed, but placeholder is not managed by this feed",
        );
        return;
      }
      if (isFutureEventResource_(arrivalEv, today)) {
        calendarEventRemove_(mapping.calendarId, arrivalEv.id, arrivalEv);
        stats.arrivalDeleted++;
        console.log(
          "[DELETE] " +
            formatEventLogContext_(
              arrivalEv,
              mapping.calendarId,
              feedName,
              "arrival placeholder",
            ) +
            " — missing from feed",
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
 * Returns the calendar-local start date for concise event logging.
 */
function eventStartDateForLog_(eventResource) {
  const start = eventResource && eventResource.start;
  const value = start && (start.date || start.dateTime);
  return value ? String(value).slice(0, 10) : "(Unknown date)";
}

/**
 * Returns a consistent human-readable context for event-level logging.
 */
function formatEventLogContext_(
  eventResource,
  calendarId,
  feedName,
  eventKind,
) {
  const privateProps =
    ((eventResource || {}).extendedProperties || {}).private || {};
  const kind =
    eventKind || managedEventKindForLog_(privateProps.managedKind || "source");
  const title = (eventResource && eventResource.summary) || "(No title)";
  const sourceName =
    feedName ||
    privateProps.sourceFeedName ||
    privateProps.sourceUrl ||
    "(Unknown feed)";
  return (
    kind +
    ' "' +
    title +
    '" on ' +
    eventStartDateForLog_(eventResource) +
    " in " +
    (calendarId || "(Unknown calendar)") +
    " from " +
    sourceName
  );
}

function managedEventKindForLog_(managedKind) {
  if (managedKind === "drive") return "drive placeholder";
  if (managedKind === "arrival") return "arrival placeholder";
  return "source event";
}

function mappingFeedName_(mapping) {
  return (mapping && (mapping.name || mapping.feedUrl)) || "(Unknown feed)";
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
  if (typeof cfg.triggerHours !== "undefined" && cfg.triggerHours !== null) {
    cfg.triggerHours = normalizeTriggerHours_(cfg.triggerHours);
  }
  if (typeof cfg.deleteMissingFromFeed !== "boolean")
    cfg.deleteMissingFromFeed = true;
  if (!Array.isArray(cfg.defaultAttendeeEmails)) cfg.defaultAttendeeEmails = [];
  if (typeof cfg.addDriveTimePlaceholders !== "boolean")
    cfg.addDriveTimePlaceholders = false;
  if (typeof cfg.defaultOriginAddress !== "string")
    cfg.defaultOriginAddress = "";
  if (
    !cfg.placeNameAddressMap ||
    typeof cfg.placeNameAddressMap !== "object" ||
    Array.isArray(cfg.placeNameAddressMap)
  ) {
    cfg.placeNameAddressMap = {};
  }
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

  const feedMappingNameIndexes = Object.create(null);
  cfg.feedMappings.forEach(function (m, i) {
    if (typeof m.name !== "string" || !m.name.trim()) {
      throw new Error("feedMappings[" + i + "] missing name.");
    }
    m.name = m.name.trim();
    if (Object.prototype.hasOwnProperty.call(feedMappingNameIndexes, m.name)) {
      throw new Error(
        "feedMappings[" +
          i +
          '] name "' +
          m.name +
          '" duplicates feedMappings[' +
          feedMappingNameIndexes[m.name] +
          "].name; feed mapping names must be unique.",
      );
    }
    feedMappingNameIndexes[m.name] = i;
    if (!m.feedUrl) throw new Error("feedMappings[" + i + "] missing feedUrl.");
    if (!m.calendarId)
      throw new Error("feedMappings[" + i + "] missing calendarId.");
    if (
      Object.prototype.hasOwnProperty.call(m, "attendeeEmails") &&
      !Array.isArray(m.attendeeEmails)
    ) {
      m.attendeeEmails = [];
    }
    if (typeof m.titlePrefix !== "string") m.titlePrefix = "";
    if (typeof m.timeZone !== "string") m.timeZone = "";
    if (typeof m.skipAllDayEvents !== "boolean") m.skipAllDayEvents = false;
    if (typeof m.addDestinationCalendarAsAttendee !== "boolean")
      m.addDestinationCalendarAsAttendee = true;
    if (typeof m.addDriveTimePlaceholders !== "boolean")
      m.addDriveTimePlaceholders = cfg.addDriveTimePlaceholders;
    if (typeof m.originAddress !== "string") m.originAddress = "";
    if (
      !m.placeNameAddressMap ||
      typeof m.placeNameAddressMap !== "object" ||
      Array.isArray(m.placeNameAddressMap)
    ) {
      m.placeNameAddressMap = {};
    }
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
 * Retries transient Calendar write quota/rate errors. General Calendar usage
 * limits fail immediately so the caller can stop the rest of the sync.
 */
function withCalendarWriteRetry_(opName, context, fn) {
  let lastError = null;
  CALENDAR_WRITE_DIAGNOSTICS.started++;
  context.writeNumber = CALENDAR_WRITE_DIAGNOSTICS.started;

  for (let attempt = 1; attempt <= CALENDAR_WRITE_MAX_ATTEMPTS; attempt++) {
    try {
      const result = fn();
      CALENDAR_WRITE_DIAGNOSTICS.succeeded++;
      return result;
    } catch (e) {
      lastError = e;
      const errorType = classifyCalendarWriteError_(e);
      const retryable = isRetryableCalendarWriteError_(e);
      if (attempt >= CALENDAR_WRITE_MAX_ATTEMPTS || !retryable) {
        console.error(
          "[CALENDAR_WRITE_FAILED] op=" +
            opName +
            " errorType=" +
            errorType +
            " attempt=" +
            attempt +
            "/" +
            CALENDAR_WRITE_MAX_ATTEMPTS +
            " retryable=" +
            retryable +
            formatCalendarWriteContext_(context) +
            " error=" +
            formatCalendarWriteLogValue_(String(e)),
        );
        throw e;
      }
      const sleepMs = computeCalendarWriteBackoffMs_(attempt);
      console.warn(
        "[CALENDAR_WRITE_RETRY] op=" +
          opName +
          " errorType=" +
          errorType +
          " attempt=" +
          attempt +
          "/" +
          CALENDAR_WRITE_MAX_ATTEMPTS +
          " nextDelayMs=" +
          sleepMs +
          formatCalendarWriteContext_(context) +
          " error=" +
          formatCalendarWriteLogValue_(String(e)),
      );
      Utilities.sleep(sleepMs);
    }
  }

  throw lastError;
}

function calendarEventInsert_(resource, calendarId, options) {
  const insertResource = ensureDeterministicCalendarEventId_(resource);
  const context = buildCalendarWriteContext_(insertResource, calendarId);
  return withCalendarWriteRetry_("insert", context, function () {
    try {
      return Calendar.Events.insert(insertResource, calendarId, options);
    } catch (e) {
      if (!isDuplicateCalendarEventIdError_(e)) throw e;
      return recoverDeterministicCalendarEventInsert_(
        insertResource,
        calendarId,
        context,
        e,
      );
    }
  });
}

/**
 * Treats an existing matching deterministic event as a successful insert.
 * This covers an insert that committed in Calendar before the client saw a
 * transient failure and then returned a duplicate-ID conflict on retry.
 */
function recoverDeterministicCalendarEventInsert_(
  resource,
  calendarId,
  context,
  insertError,
) {
  let existing;
  try {
    existing = Calendar.Events.get(calendarId, resource.id);
  } catch (lookupError) {
    console.warn(
      "[CALENDAR_INSERT_RECOVERY_FAILED]" +
        formatCalendarWriteContext_(context) +
        " reason=event lookup after duplicate insert failed" +
        " insertError=" +
        formatCalendarWriteLogValue_(String(insertError)) +
        " lookupError=" +
        formatCalendarWriteLogValue_(String(lookupError)),
    );
    throw insertError;
  }

  const expectedSyncKey = String(
    (((resource || {}).extendedProperties || {}).private || {}).syncKey || "",
  );
  const existingSyncKey = String(
    (((existing || {}).extendedProperties || {}).private || {}).syncKey || "",
  );
  const existingStatus = String((existing && existing.status) || "");

  if (
    existing &&
    existing.id === resource.id &&
    existingStatus !== "cancelled" &&
    existingSyncKey === expectedSyncKey
  ) {
    console.warn(
      "[CALENDAR_INSERT_RECOVERED]" +
        formatCalendarWriteContext_(context) +
        " reason=matching managed event was already stored by Calendar",
    );
    return existing;
  }

  throw new Error(
    "Deterministic Calendar event ID conflict: stored event does not match " +
      "the expected managed event (status " +
      (existingStatus || "unknown") +
      "). Original insert error: " +
      String(insertError),
  );
}

function calendarEventPatch_(resource, calendarId, eventId, options) {
  const context = buildCalendarWriteContext_(resource, calendarId);
  return withCalendarWriteRetry_("patch", context, function () {
    return Calendar.Events.patch(resource, calendarId, eventId, options);
  });
}

function calendarEventRemove_(calendarId, eventId, resource) {
  const context = buildCalendarWriteContext_(resource, calendarId);
  return withCalendarWriteRetry_("remove", context, function () {
    return Calendar.Events.remove(calendarId, eventId);
  });
}

/**
 * Assigns the stable Google Calendar event ID required for every managed insert.
 */
function ensureDeterministicCalendarEventId_(resource) {
  if (!resource || typeof resource !== "object") {
    throw new Error("Calendar event insert requires an event resource.");
  }

  const privateProps = (resource.extendedProperties || {}).private || {};
  const syncKey = String(privateProps.syncKey || "").trim();
  if (!syncKey) {
    throw new Error(
      "Calendar event insert requires extendedProperties.private.syncKey " +
        "to generate a deterministic event ID.",
    );
  }

  resource.id = buildDeterministicCalendarEventId_(syncKey);
  return resource;
}

/**
 * Produces an API-safe, stable event ID from this script's logical sync key.
 * Hex is a subset of Calendar's base32hex ID alphabet.
 */
function buildDeterministicCalendarEventId_(syncKey) {
  const normalized = String(syncKey || "").trim();
  if (!normalized) {
    throw new Error(
      "Cannot build a deterministic event ID without a sync key.",
    );
  }
  return sha256Hex_("ical-sync-event:" + normalized);
}

/**
 * Captures bounded, non-payload context needed to diagnose Calendar writes.
 */
function buildCalendarWriteContext_(resource, calendarId) {
  const privateProps =
    ((resource || {}).extendedProperties || {}).private || {};
  return {
    calendarId: calendarId || "",
    eventKind: managedEventKindForLog_(privateProps.managedKind || "source"),
    title: (resource && resource.summary) || "(No title)",
    eventDate: eventStartDateForLog_(resource),
    feedName:
      privateProps.sourceFeedName || privateProps.sourceUrl || "(Unknown feed)",
  };
}

function formatCalendarWriteContext_(context) {
  const value = context || {};
  return (
    " eventKind=" +
    formatCalendarWriteLogValue_(value.eventKind) +
    " title=" +
    formatCalendarWriteLogValue_(value.title) +
    " eventDate=" +
    formatCalendarWriteLogValue_(value.eventDate) +
    " calendarId=" +
    formatCalendarWriteLogValue_(value.calendarId) +
    " feedName=" +
    formatCalendarWriteLogValue_(value.feedName) +
    " writeNumber=" +
    String(value.writeNumber || "") +
    " writesSucceeded=" +
    CALENDAR_WRITE_DIAGNOSTICS.succeeded
  );
}

function formatCalendarWriteLogValue_(value) {
  const normalized = String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 240);
  return JSON.stringify(normalized);
}

function classifyCalendarWriteError_(err) {
  const text = calendarWriteErrorText_(err);
  if (text.indexOf("calendar usage limits exceeded") >= 0) {
    return "calendar_usage_limits";
  }
  if (text.indexOf("userratelimitexceeded") >= 0) {
    return "user_rate_limit";
  }
  if (
    text.indexOf("ratelimitexceeded") >= 0 ||
    text.indexOf("rate limit exceeded") >= 0
  ) {
    return "rate_limit";
  }
  if (text.indexOf("service invoked too many times in a short time") >= 0) {
    return "apps_script_short_term_rate";
  }
  if (text.indexOf("service invoked too many times") >= 0) {
    return "apps_script_daily_quota";
  }
  if (text.indexOf("quota exceeded") >= 0) return "quota_exceeded";
  return "non_quota_error";
}

function isCalendarUsageLimitError_(err) {
  return classifyCalendarWriteError_(err) === "calendar_usage_limits";
}

function isDuplicateCalendarEventIdError_(err) {
  const text = calendarWriteErrorText_(err);
  return (
    text.indexOf("requested identifier already exists") >= 0 ||
    (text.indexOf("already exists") >= 0 && text.indexOf("duplicate") >= 0) ||
    text.indexOf('reason": "duplicate') >= 0 ||
    text.indexOf("reason: duplicate") >= 0
  );
}

function calendarWriteErrorText_(err) {
  return String(
    (err && (err.message || err.details || (err.toString && err.toString()))) ||
      "",
  ).toLowerCase();
}

function isRetryableCalendarWriteError_(err) {
  const text = calendarWriteErrorText_(err);
  return (
    text.indexOf("rate limit exceeded") >= 0 ||
    text.indexOf("userratelimitexceeded") >= 0 ||
    text.indexOf("ratelimitexceeded") >= 0 ||
    text.indexOf("quota exceeded") >= 0 ||
    text.indexOf("service invoked too many times") >= 0
  );
}

function computeCalendarWriteBackoffMs_(attempt) {
  const base = CALENDAR_WRITE_BASE_SLEEP_MS * Math.pow(2, attempt - 1);
  const jitter = Math.floor(Math.random() * 250);
  return base + jitter;
}

/**
 * Parses an ICS document into normalized event objects plus calendar-level timezone.
 */
function parseIcs_(text, fallbackTimeZone) {
  const lines = unfoldIcsLines_(text);
  const events = [];
  let inEvent = false;
  let block = [];
  let calendarTimezone =
    String(fallbackTimeZone || "").trim() || Session.getScriptTimeZone();

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
    d.setMinutes(d.getMinutes() + DEFAULT_TIMED_EVENT_DURATION_MINUTES);
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
  d.setMinutes(d.getMinutes() + DEFAULT_TIMED_EVENT_DURATION_MINUTES);

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
  feedName,
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
    attendees: attendees.map(function (attendee) {
      return toCalendarAttendeeResource_(attendee);
    }),
    extendedProperties: {
      private: {
        managedKind: "source",
        sourceFeed: feedHash,
        sourceUrl: feedUrl,
        sourceFeedName: feedName,
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
  feedName,
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
    attendees: attendees.map(function (attendee) {
      return toCalendarAttendeeResource_(attendee);
    }),
    extendedProperties: {
      private: {
        managedKind: "source",
        sourceFeed: feedHash,
        sourceUrl: feedUrl,
        sourceFeedName: feedName,
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
 * Rewrites an event location to its configured canonical address when matched.
 */
function applyPlaceNameAddressToEvent_(evt, rules, mapping) {
  if (!evt || typeof evt !== "object") return evt;
  const locationResolution = resolvePlaceNameAddress_(
    evt.location || "",
    rules,
  );
  if (!locationResolution.matched) return evt;

  const copied = Object.assign({}, evt);
  copied.location = locationResolution.text;
  console.info(
    "[INFO] " +
      formatEventLogContext_(
        copied,
        mapping && mapping.calendarId,
        mappingFeedName_(mapping),
        "source event",
      ) +
      ' — rewrote location from "' +
      locationResolution.sourceText +
      '" to "' +
      locationResolution.text +
      '" using place name "' +
      locationResolution.matchedPlaceName +
      '"',
  );
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
    placeNameAddressRules: buildPlaceNameAddressRules_(cfg, mapping),
    minDriveMinutesToCreate: cfg.minDriveMinutesToCreate,
    titleTemplate: cfg.driveEventTitleTemplate,
  };
}

/**
 * Builds normalized place-name replacement rules from global and per-feed config.
 */
function buildPlaceNameAddressRules_(cfg, mapping) {
  const merged = Object.assign(
    {},
    cfg.placeNameAddressMap || {},
    mapping.placeNameAddressMap || {},
  );
  return Object.keys(merged)
    .map(function (name) {
      return {
        placeName: String(name || "").trim(),
        placeNameLower: String(name || "")
          .trim()
          .toLowerCase(),
        address: String(merged[name] || "").trim(),
      };
    })
    .filter(function (rule) {
      return !!rule.placeName && !!rule.address;
    })
    .sort(function (a, b) {
      return b.placeName.length - a.placeName.length;
    });
}

/**
 * Returns configured attendees, optionally including the destination calendar.
 */
function buildSourceAttendees_(
  attendees,
  calendarId,
  addDestinationCalendarAsAttendee,
) {
  const values = (attendees || []).slice();
  if (addDestinationCalendarAsAttendee !== false) values.push(calendarId);
  return uniqueEmails_(values);
}

/**
 * Makes attendee removal fix-forward when the destination calendar is no
 * longer added automatically. Existing managed events retain their complete
 * attendee list, while new events use only the newly configured attendees.
 */
function resolveAttendeesForExistingManagedEvent_(
  configuredAttendees,
  existingEvent,
  addDestinationCalendarAsAttendee,
) {
  if (addDestinationCalendarAsAttendee !== false || !existingEvent) {
    return (configuredAttendees || []).slice();
  }

  return ((existingEvent && existingEvent.attendees) || [])
    .filter(function (attendee) {
      return !!(attendee && attendee.email);
    })
    .map(function (attendee) {
      return Object.assign({}, attendee);
    });
}

function getAttendeeEmails_(attendees) {
  return uniqueEmails_(
    (attendees || []).map(function (attendee) {
      return typeof attendee === "string"
        ? attendee
        : attendee && attendee.email;
    }),
  );
}

/**
 * Returns feed hashes for other configured feeds actively syncing into the same calendar.
 */
function buildActivePeerFeedHashes_(cfg, calendarId, currentFeedHash) {
  const out = {};
  ((cfg && cfg.feedMappings) || []).forEach(function (mapping) {
    if (!mapping || !mapping.feedUrl || mapping.calendarId !== calendarId)
      return;
    const feedHash = sha256Hex_(mapping.feedUrl).slice(0, 16);
    if (feedHash === currentFeedHash) return;
    out[feedHash] = true;
  });
  return out;
}

/**
 * Deletes future managed events whose source feed is no longer configured.
 */
function cleanupRemovedFeedEvents_(cfg, today) {
  const activeFeedsByCalendar = buildActiveFeedIdentityByCalendar_(cfg);
  const calendarIds = collectManagedCalendarIdsForCleanup_(cfg);
  const stats = {
    sourceDeleted: 0,
    arrivalDeleted: 0,
    driveDeleted: 0,
    skippedPast: 0,
  };

  calendarIds.forEach(function (calendarId) {
    ["source", "arrival", "drive"].forEach(function (managedKind) {
      const activeFeeds = activeFeedsByCalendar[calendarId] || {
        feedKeys: {},
      };
      loadManagedEventsByKind_(calendarId, managedKind).forEach(function (ev) {
        if (!isRemovedFeedManagedEvent_(ev, managedKind, activeFeeds)) return;
        if (!isFutureEventResource_(ev, today)) {
          stats.skippedPast++;
          return;
        }

        calendarEventRemove_(calendarId, ev.id, ev);
        if (managedKind === "source") stats.sourceDeleted++;
        if (managedKind === "arrival") stats.arrivalDeleted++;
        if (managedKind === "drive") stats.driveDeleted++;
        console.log(
          "[DELETE] " +
            formatEventLogContext_(
              ev,
              calendarId,
              null,
              managedEventKindForLog_(managedKind),
            ) +
            " — source feed is no longer configured",
        );
      });
    });
  });

  const deleted =
    stats.sourceDeleted + stats.arrivalDeleted + stats.driveDeleted;
  if (deleted || stats.skippedPast) {
    console.log(
      "[SUMMARY] Removed feed cleanup: sourceDeleted=" +
        stats.sourceDeleted +
        ", arrivalDeleted=" +
        stats.arrivalDeleted +
        ", driveDeleted=" +
        stats.driveDeleted +
        ", skippedPast=" +
        stats.skippedPast,
    );
  }
  return stats;
}

/**
 * Returns active feed hashes and URLs by target calendar.
 */
function buildActiveFeedIdentityByCalendar_(cfg) {
  const out = {};
  ((cfg && cfg.feedMappings) || []).forEach(function (mapping) {
    if (!mapping || !mapping.feedUrl || !mapping.calendarId) return;
    const calendarId = String(mapping.calendarId);
    const feedHash = sha256Hex_(mapping.feedUrl).slice(0, 16);
    if (!out[calendarId]) out[calendarId] = { feedKeys: {} };
    out[calendarId].feedKeys[buildFeedIdentityKey_(feedHash, mapping.feedUrl)] =
      true;
  });
  return out;
}

/**
 * Builds a paired feed identity key from the metadata written to synced events.
 */
function buildFeedIdentityKey_(feedHash, feedUrl) {
  return String(feedHash || "") + "\n" + String(feedUrl || "");
}

/**
 * Returns current and previously configured target calendars for orphan cleanup.
 */
function collectManagedCalendarIdsForCleanup_(cfg) {
  const ids = {};
  ((cfg && cfg.feedMappings) || []).forEach(function (mapping) {
    if (mapping && mapping.calendarId) ids[String(mapping.calendarId)] = true;
  });
  getRememberedManagedCalendarIds_().forEach(function (calendarId) {
    ids[calendarId] = true;
  });
  return Object.keys(ids);
}

/**
 * Records target calendars seen by this script so removed-calendar mappings can be cleaned later.
 */
function rememberManagedCalendarIds_(cfg) {
  const ids = {};
  getRememberedManagedCalendarIds_().forEach(function (calendarId) {
    ids[calendarId] = true;
  });
  ((cfg && cfg.feedMappings) || []).forEach(function (mapping) {
    if (mapping && mapping.calendarId) ids[String(mapping.calendarId)] = true;
  });

  try {
    PropertiesService.getScriptProperties().setProperty(
      MANAGED_CALENDAR_IDS_PROPERTY,
      JSON.stringify(Object.keys(ids)),
    );
  } catch (e) {
    console.warn(
      "[WARN] Failed remembering managed calendar IDs for cleanup: " +
        String(e),
    );
  }
}

/**
 * Loads target calendars remembered from prior sync runs.
 */
function getRememberedManagedCalendarIds_() {
  let raw = null;
  try {
    raw = PropertiesService.getScriptProperties().getProperty(
      MANAGED_CALENDAR_IDS_PROPERTY,
    );
  } catch (e) {
    console.warn(
      "[WARN] Failed reading remembered managed calendar IDs: " + String(e),
    );
    return [];
  }
  if (!raw) return [];

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .map(function (calendarId) {
        return String(calendarId || "").trim();
      })
      .filter(function (calendarId) {
        return !!calendarId;
      });
  } catch (e) {
    console.warn(
      "[WARN] Ignoring invalid remembered managed calendar IDs: " + String(e),
    );
    return [];
  }
}

/**
 * Loads managed events for one kind from a calendar.
 */
function loadManagedEventsByKind_(calendarId, managedKind) {
  const out = [];
  let pageToken;

  do {
    const resp = Calendar.Events.list(calendarId, {
      privateExtendedProperty: ["managedKind=" + managedKind],
      showDeleted: false,
      singleEvents: false,
      maxResults: 2500,
      pageToken: pageToken,
    });

    (resp.items || []).forEach(function (ev) {
      out.push(ev);
    });

    pageToken = resp.nextPageToken;
  } while (pageToken);

  return out;
}

/**
 * Returns true for managed events tied to a feed no longer active on the target calendar.
 */
function isRemovedFeedManagedEvent_(ev, managedKind, activeFeeds) {
  const p = ((ev || {}).extendedProperties || {}).private || {};
  if (p.managedKind !== managedKind) return false;
  if (!p.sourceFeed || !p.sourceUrl || !p.syncKey) return false;

  if (managedKind === "source") {
    if (isArrivalSyncKey_(p.syncKey) || isDriveSyncKey_(p.syncKey))
      return false;
  } else if (managedKind === "arrival") {
    if (!isArrivalSyncKey_(p.syncKey) || !p.sourceSyncKey) return false;
  } else if (managedKind === "drive") {
    if (!isDriveSyncKey_(p.syncKey) || !p.sourceSyncKey) return false;
  } else {
    return false;
  }

  return !(
    activeFeeds &&
    activeFeeds.feedKeys &&
    activeFeeds.feedKeys[buildFeedIdentityKey_(p.sourceFeed, p.sourceUrl)]
  );
}

/**
 * Handles exact duplicate events before creating a new managed source event.
 * If a duplicate is already managed by this feed and its old identity is absent from
 * the current feed, adopt it so an upstream identity change can be patched in place.
 * Preserve active same-feed identities. If another active feed manages the duplicate,
 * skip this create. Otherwise remove duplicate non-active events and allow the create.
 */
function resolveDuplicateBeforeCreate_(
  calendarId,
  createResource,
  activePeerFeedHashes,
  activeSourceSyncKeys,
) {
  const duplicates = findDuplicateEventsForResource_(
    calendarId,
    createResource,
  );
  let hasActivePeerDuplicate = false;
  let hasActiveSameFeedDuplicate = false;
  let sameFeedDuplicate = null;
  let deleted = 0;
  const createPrivateProps =
    ((createResource || {}).extendedProperties || {}).private || {};
  const createFeedName =
    createPrivateProps.sourceFeedName || createPrivateProps.sourceUrl;

  duplicates.forEach(function (duplicate) {
    if (
      isManagedEventForFeed_(
        duplicate,
        createPrivateProps.sourceUrl,
        createPrivateProps.sourceFeed,
      )
    ) {
      const duplicatePrivateProps =
        ((duplicate || {}).extendedProperties || {}).private || {};
      const duplicateSyncKey = duplicatePrivateProps.syncKey || "";
      if (
        duplicateSyncKey === createPrivateProps.syncKey ||
        !(activeSourceSyncKeys && activeSourceSyncKeys[duplicateSyncKey])
      ) {
        if (!sameFeedDuplicate) sameFeedDuplicate = duplicate;
      } else {
        hasActiveSameFeedDuplicate = true;
      }
      return;
    }

    if (isActivePeerSourceEvent_(duplicate, activePeerFeedHashes)) {
      hasActivePeerDuplicate = true;
      return;
    }

    calendarEventRemove_(calendarId, duplicate.id, duplicate);
    deleted++;
    console.log(
      "[DELETE] " +
        formatEventLogContext_(
          duplicate,
          calendarId,
          createFeedName,
          "source event",
        ) +
        " — duplicate non-active event",
    );
  });

  return {
    adoptExisting: sameFeedDuplicate,
    skipCreate:
      !sameFeedDuplicate &&
      !hasActiveSameFeedDuplicate &&
      hasActivePeerDuplicate,
    deleted: deleted,
  };
}

/**
 * Moves an already-loaded managed event to its newly adopted sync key.
 * Reindexing immediately prevents later old-identity tombstones and cleanup from
 * acting on the Calendar event after its metadata has been patched in place.
 */
function reindexExistingManagedEvent_(eventsByKey, oldKey, newKey, event) {
  if (!eventsByKey || !oldKey || !newKey || oldKey === newKey) return null;
  const existing = event || eventsByKey[oldKey] || null;
  if (!existing) return null;

  const alreadyIndexed = eventsByKey[newKey];
  if (alreadyIndexed && alreadyIndexed.id !== existing.id) return null;

  delete eventsByKey[oldKey];
  eventsByKey[newKey] = existing;
  return existing;
}

/**
 * Finds exact event duplicates by title, start, end, and description.
 */
function findDuplicateEventsForResource_(calendarId, resource) {
  const out = [];
  const startDate = eventBoundaryToDate_(resource.start);
  const endDate = eventBoundaryToDate_(resource.end || resource.start);
  const listOpts = {
    showDeleted: false,
    singleEvents: false,
    maxResults: 250,
  };

  if (startDate && endDate) {
    listOpts.timeMin = new Date(startDate.getTime() - 60 * 1000).toISOString();
    listOpts.timeMax = new Date(endDate.getTime() + 60 * 1000).toISOString();
  }

  let pageToken;
  do {
    listOpts.pageToken = pageToken;
    const resp = Calendar.Events.list(calendarId, listOpts);
    (resp.items || []).forEach(function (ev) {
      if (isDuplicateEventResource_(ev, resource)) out.push(ev);
    });
    pageToken = resp.nextPageToken;
  } while (pageToken);

  return out;
}

/**
 * Returns true when an event resource exactly duplicates another source event.
 */
function isDuplicateEventResource_(candidate, resource) {
  if (!candidate || !resource) return false;
  return (
    normalizeDuplicateText_(candidate.summary) ===
      normalizeDuplicateText_(resource.summary) &&
    normalizeDuplicateText_(candidate.description) ===
      normalizeDuplicateText_(resource.description) &&
    eventBoundariesMatch_(candidate.start, resource.start) &&
    eventBoundariesMatch_(candidate.end, resource.end)
  );
}

/**
 * Returns true when an existing event has the same start/end as the desired resource.
 */
function eventResourceTimingMatches_(candidate, resource) {
  if (!candidate || !resource) return false;
  return (
    eventBoundariesMatch_(candidate.start, resource.start) &&
    eventBoundariesMatch_(candidate.end, resource.end)
  );
}

/**
 * Compares Calendar boundaries even when one side is a floating local time and
 * the Calendar API has canonicalized the other side to an explicit UTC offset.
 */
function eventBoundariesMatch_(left, right) {
  if (!left || !right) return false;
  if (left.date || right.date) {
    return (
      !!left.date && !!right.date && String(left.date) === String(right.date)
    );
  }
  if (!left.dateTime || !right.dateTime) return false;

  if (eventBoundaryKey_(left) === eventBoundaryKey_(right)) return true;

  const leftHasZone = dateTimeHasExplicitZone_(left.dateTime);
  const rightHasZone = dateTimeHasExplicitZone_(right.dateTime);
  if (leftHasZone === rightHasZone) return false;

  const explicitBoundary = leftHasZone ? left : right;
  const floatingBoundary = leftHasZone ? right : left;
  const explicitDate = new Date(explicitBoundary.dateTime);
  if (isNaN(explicitDate.getTime())) return false;

  const timeZone =
    String(floatingBoundary.timeZone || "").trim() ||
    Session.getScriptTimeZone();
  let explicitLocal;
  try {
    explicitLocal = Utilities.formatDate(
      explicitDate,
      timeZone,
      "yyyy-MM-dd'T'HH:mm:ss",
    );
  } catch (e) {
    return false;
  }
  return (
    normalizeFloatingDateTime_(explicitLocal) ===
    normalizeFloatingDateTime_(floatingBoundary.dateTime)
  );
}

/**
 * Returns true when a duplicate belongs to another active feed syncing here.
 */
function isActivePeerSourceEvent_(ev, activePeerFeedHashes) {
  const privateProps = ((ev || {}).extendedProperties || {}).private || {};
  const feedHash = privateProps.sourceFeed || "";
  const syncKey = privateProps.syncKey || "";
  return !!(
    privateProps.managedKind === "source" &&
    feedHash &&
    activePeerFeedHashes &&
    activePeerFeedHashes[feedHash] &&
    !isDriveSyncKey_(syncKey) &&
    !isArrivalSyncKey_(syncKey)
  );
}

/**
 * Normalizes title/description text for duplicate comparison.
 */
function normalizeDuplicateText_(value) {
  const generatedByPattern = new RegExp(
    "\\n*" + escapeRegExp_(GENERATED_BY_DESCRIPTION) + "\\s*$",
  );
  return String(value || "")
    .replace(/\r\n/g, "\n")
    .replace(generatedByPattern, "")
    .trim();
}

/**
 * Escapes a string for use inside a RegExp.
 */
function escapeRegExp_(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Converts Calendar event start/end shapes into stable comparison keys.
 */
function eventBoundaryKey_(boundary) {
  if (!boundary) return "";
  if (boundary.date) return "date:" + String(boundary.date);
  if (boundary.dateTime) {
    const dateTime = String(boundary.dateTime);
    if (!dateTimeHasExplicitZone_(dateTime)) {
      return (
        "dateTime:" +
        normalizeFloatingDateTime_(dateTime) +
        "|" +
        String(boundary.timeZone || "")
      );
    }
    const parsed = new Date(dateTime);
    if (!isNaN(parsed.getTime())) return "dateTime:" + parsed.toISOString();
    return "dateTime:" + dateTime + "|" + String(boundary.timeZone || "");
  }
  return "";
}

/**
 * Returns true when an RFC3339-ish dateTime string includes Z or a numeric offset.
 */
function dateTimeHasExplicitZone_(dateTime) {
  return /(?:Z|[+-]\d{2}:?\d{2})$/i.test(String(dateTime || ""));
}

/**
 * Normalizes local-clock dateTime values while preserving their floating timezone.
 */
function normalizeFloatingDateTime_(dateTime) {
  return String(dateTime || "")
    .replace(/(\.\d*?)0+$/, "$1")
    .replace(/\.$/, "");
}

/**
 * Converts Calendar event start/end shapes into dates for Calendar list windows.
 */
function eventBoundaryToDate_(boundary) {
  if (!boundary) return null;
  if (boundary.dateTime) {
    const dateTime = dateTimeHasExplicitZone_(boundary.dateTime)
      ? new Date(boundary.dateTime)
      : floatingDateTimeToDate_(
          boundary.dateTime,
          boundary.timeZone || Session.getScriptTimeZone(),
        );
    if (!isNaN(dateTime.getTime())) return dateTime;
  }
  if (boundary.date) {
    const date = new Date(boundary.date + "T00:00:00");
    if (!isNaN(date.getTime())) return date;
  }
  return null;
}

/**
 * Resolves a floating local-clock dateTime to an instant in an IANA timezone.
 */
function floatingDateTimeToDate_(dateTime, timeZone) {
  const match = String(dateTime || "").match(
    /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?$/,
  );
  if (!match) return new Date(NaN);

  const milliseconds = Number((match[7] || "").slice(0, 3).padEnd(3, "0"));
  const localAsUtc = Date.UTC(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4]),
    Number(match[5]),
    Number(match[6]),
    milliseconds,
  );

  let candidate = new Date(localAsUtc);
  for (let i = 0; i < 2; i++) {
    let offset;
    try {
      offset = parseTimeZoneOffsetMinutes_(
        Utilities.formatDate(candidate, timeZone, "Z"),
      );
    } catch (e) {
      return new Date(NaN);
    }
    if (offset === null) return new Date(NaN);
    candidate = new Date(localAsUtc - offset * 60 * 1000);
  }
  return candidate;
}

/**
 * Parses RFC822-style timezone offsets such as -0700 or +05:30.
 */
function parseTimeZoneOffsetMinutes_(offsetText) {
  const match = String(offsetText || "").match(/^([+-])(\d{2}):?(\d{2})$/);
  if (!match) return null;
  const minutes = Number(match[2]) * 60 + Number(match[3]);
  return match[1] === "-" ? -minutes : minutes;
}

/**
 * Loads existing calendar events previously managed by this feed, keyed by syncKey.
 */
function loadExistingEventsByKey_(calendarId, feedHash, feedName) {
  const out = {};
  const duplicates = [];
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
      if (!key || isDriveSyncKey_(key) || isArrivalSyncKey_(key)) return;
      if (!out[key]) {
        out[key] = ev;
        return;
      }

      const existingCreated = Date.parse(out[key].created || "") || Infinity;
      const candidateCreated = Date.parse(ev.created || "") || Infinity;
      if (candidateCreated < existingCreated) {
        duplicates.push(out[key]);
        out[key] = ev;
      } else {
        duplicates.push(ev);
      }
    });

    pageToken = resp.nextPageToken;
  } while (pageToken);

  duplicates.forEach(function (duplicate) {
    calendarEventRemove_(calendarId, duplicate.id, duplicate);
    console.log(
      "[DELETE] " +
        formatEventLogContext_(
          duplicate,
          calendarId,
          feedName,
          "source event",
        ) +
        " — duplicate managed sync key",
    );
  });

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
  attendees,
) {
  const existingArrival = existingArrivalByKey[arrivalSyncKey] || null;

  if (!isEventStartOnOrAfterCutoff_(evt, today)) {
    stats.arrivalSkipped++;
    console.info(
      "[SKIP] " +
        formatEventLogContext_(
          evt,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "source event",
        ) +
        " — arrival placeholder ignored because source event is before the sync cutoff",
    );
    return null;
  }

  if (isAllDayEvent_(evt)) {
    stats.arrivalSkipped++;
    console.info(
      "[SKIP] " +
        formatEventLogContext_(
          evt,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "source event",
        ) +
        " — arrival placeholder ignored because source event is all-day",
    );
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
      "[SKIP] " +
        formatEventLogContext_(
          evt,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "source event",
        ) +
        " — arrival placeholder ignored because source start time is unavailable",
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
    evt.summary || "(No title)",
    arrivalStart,
    arrivalEnd,
    arrivalTitle,
    arrivalMinutes,
  );
  const arrivalAttendees = resolveAttendeesForExistingManagedEvent_(
    attendees,
    existingArrival,
    mapping.addDestinationCalendarAsAttendee,
  );
  const arrivalResource = buildArrivalPlaceholderResource_(
    mapping,
    feedHash,
    evt,
    sourceSyncKey,
    arrivalSyncKey,
    syncedEvent.id,
    evt.summary || "(No title)",
    arrivalTitle,
    arrivalStart,
    arrivalEnd,
    arrivalHash,
    arrivalMinutes,
    arrivalAttendees,
  );
  seenArrival[arrivalSyncKey] = true;
  const existingArrivalHash = (
    (existingArrival && existingArrival.extendedProperties) ||
    {}
  ).private
    ? ((existingArrival.extendedProperties || {}).private || {}).syncHash || ""
    : "";

  if (!existingArrival) {
    calendarEventInsert_(arrivalResource, mapping.calendarId, {
      sendUpdates: "none",
    });
    stats.arrivalCreated++;
    console.log(
      "[CREATE] " +
        formatEventLogContext_(
          arrivalResource,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "arrival placeholder",
        ) +
        " — source event requests an advanced arrival",
    );
    return arrivalStart;
  }

  if (
    !isManagedArrivalEventForFeed_(existingArrival, mapping.feedUrl, feedHash)
  ) {
    stats.arrivalSkipped++;
    console.info(
      "[SKIP] " +
        formatEventLogContext_(
          existingArrival,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "arrival placeholder",
        ) +
        " — placeholder is not managed by this feed",
    );
    return null;
  }

  if (existingArrivalHash === arrivalHash) {
    console.log(
      "[UNCHANGED] " +
        formatEventLogContext_(
          existingArrival,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "arrival placeholder",
        ) +
        " — no feed changes detected",
    );
    return arrivalStart;
  }

  calendarEventPatch_(arrivalResource, mapping.calendarId, existingArrival.id, {
    sendUpdates: "none",
  });
  stats.arrivalUpdated++;
  console.log(
    "[UPDATE] " +
      formatEventLogContext_(
        arrivalResource,
        mapping.calendarId,
        mappingFeedName_(mapping),
        "arrival placeholder",
      ) +
      " — feed change detected",
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
  attendees,
) {
  const destinationCandidate = resolveDriveDestinationCandidate_(evt);
  const sourceTitle = evt.summary || "(No title)";
  if (!destinationCandidate.text) {
    stats.driveSkipped++;
    console.info(
      "[SKIP] " +
        formatEventLogContext_(
          evt,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "source event",
        ) +
        " — drive placeholder ignored because source event has no location",
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
  console.info(
    "[INFO] " +
      formatEventLogContext_(
        evt,
        mapping.calendarId,
        mappingFeedName_(mapping),
        "source event",
      ) +
      " — using event location as drive destination candidate",
  );
  const destination = destinationCandidate.text;
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
    console.info(
      "[SKIP] " +
        formatEventLogContext_(
          evt,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "source event",
        ) +
        " — drive placeholder ignored because source event is before the sync cutoff",
    );
    return;
  }

  if (isAllDayEvent_(evt)) {
    stats.driveSkipped++;
    console.info(
      "[SKIP] " +
        formatEventLogContext_(
          evt,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "source event",
        ) +
        " — drive placeholder ignored because source event is all-day",
    );
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

  const sourceStart = getSourceEventStartDate_(syncedEvent);
  if (!sourceStart) {
    stats.driveSkipped++;
    console.info(
      "[SKIP] " +
        formatEventLogContext_(
          evt,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "source event",
        ) +
        " — drive placeholder ignored because source start time is unavailable",
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

  const driveEnd =
    arrivalAnchorStart instanceof Date &&
    !isNaN(arrivalAnchorStart.getTime()) &&
    arrivalAnchorStart.getTime() < sourceStart.getTime()
      ? arrivalAnchorStart
      : sourceStart;
  const drivePlan = resolveDrivePlan_(
    mapping.calendarId,
    syncedEvent.id,
    driveEnd,
    destination,
    driveOpts,
    driveDurationCache,
  );
  if (drivePlan.skipReason) {
    stats.driveSkipped++;
    if (drivePlan.routeLookupFailed) {
      console.warn(
        "[WARN] " +
          formatEventLogContext_(
            evt,
            mapping.calendarId,
            mappingFeedName_(mapping),
            "source event",
          ) +
          " — could not compute drive time (destination: " +
          destination +
          ")" +
          (drivePlan.lookupFailures && drivePlan.lookupFailures.length
            ? " after attempts: " + drivePlan.lookupFailures.join(", ")
            : ""),
      );
    }
    console.info(
      "[SKIP] " +
        formatEventLogContext_(
          evt,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "source event",
        ) +
        " — drive placeholder ignored because " +
        drivePlan.skipReason,
    );
    maybeDeleteDrivePlaceholder_(
      mapping,
      feedHash,
      driveSyncKey,
      existingDriveByKey,
      today,
      stats,
      drivePlan.skipReason,
    );
    return;
  }

  const driveMinutes = drivePlan.driveMinutes;
  const driveOrigin = drivePlan.originAddress;
  const driveStart = new Date(driveEnd.getTime() - driveMinutes * 60 * 1000);
  const driveTitle = renderDriveEventTitle_(
    driveOpts.titleTemplate,
    evt,
    driveMinutes,
  );
  const driveHash = computeDrivePlaceholderHash_(
    sourceSyncKey,
    syncedEvent.id,
    evt.summary || "(No title)",
    driveOrigin,
    destination,
    driveStart,
    driveEnd,
    driveTitle,
  );
  const driveAttendees = resolveAttendeesForExistingManagedEvent_(
    attendees,
    existingDrive,
    mapping.addDestinationCalendarAsAttendee,
  );
  const driveResource = buildDrivePlaceholderResource_(
    mapping,
    feedHash,
    evt,
    sourceSyncKey,
    driveSyncKey,
    syncedEvent.id,
    evt.summary || "(No title)",
    driveTitle,
    driveStart,
    driveEnd,
    driveHash,
    driveOrigin,
    destination,
    driveAttendees,
    drivePlan.previousEventId || "",
  );
  seenDrive[driveSyncKey] = true;
  const existingDriveHash = (
    (existingDrive && existingDrive.extendedProperties) ||
    {}
  ).private
    ? ((existingDrive.extendedProperties || {}).private || {}).syncHash || ""
    : "";

  if (!existingDrive) {
    calendarEventInsert_(driveResource, mapping.calendarId, {
      sendUpdates: "none",
    });
    stats.driveCreated++;
    console.log(
      "[CREATE] " +
        formatEventLogContext_(
          driveResource,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "drive placeholder",
        ) +
        " — drive time exceeds the configured threshold",
    );
    return;
  }

  if (!isManagedDriveEventForFeed_(existingDrive, mapping.feedUrl, feedHash)) {
    stats.driveSkipped++;
    console.info(
      "[SKIP] " +
        formatEventLogContext_(
          existingDrive,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "drive placeholder",
        ) +
        " — placeholder is not managed by this feed",
    );
    return;
  }

  if (existingDriveHash === driveHash) {
    console.log(
      "[UNCHANGED] " +
        formatEventLogContext_(
          existingDrive,
          mapping.calendarId,
          mappingFeedName_(mapping),
          "drive placeholder",
        ) +
        " — no feed changes detected",
    );
    return;
  }

  calendarEventPatch_(driveResource, mapping.calendarId, existingDrive.id, {
    sendUpdates: "none",
  });
  stats.driveUpdated++;
  console.log(
    "[UPDATE] " +
      formatEventLogContext_(
        driveResource,
        mapping.calendarId,
        mappingFeedName_(mapping),
        "drive placeholder",
      ) +
      " — feed change detected",
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

  calendarEventRemove_(mapping.calendarId, existingDrive.id, existingDrive);
  delete existingDriveByKey[driveSyncKey];
  stats.driveDeleted++;
  console.log(
    "[DELETE] " +
      formatEventLogContext_(
        existingDrive,
        mapping.calendarId,
        mappingFeedName_(mapping),
        "drive placeholder",
      ) +
      " — " +
      reason,
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

  calendarEventRemove_(mapping.calendarId, existingArrival.id, existingArrival);
  delete existingArrivalByKey[arrivalSyncKey];
  stats.arrivalDeleted++;
  console.log(
    "[DELETE] " +
      formatEventLogContext_(
        existingArrival,
        mapping.calendarId,
        mappingFeedName_(mapping),
        "arrival placeholder",
      ) +
      " — " +
      reason,
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
 * Chooses drive origin + duration using prior calendar context, then default origin fallback.
 */
function resolveDrivePlan_(
  calendarId,
  sourceEventId,
  driveEnd,
  destination,
  driveOpts,
  driveDurationCache,
) {
  const lookupFailures = [];
  const destinationResolution = resolvePlaceNameAddress_(
    String(destination || "").trim(),
    driveOpts.placeNameAddressRules,
  );
  if (destinationResolution.matched) {
    console.info(
      '[INFO] Resolved place name "' +
        destinationResolution.matchedPlaceName +
        '" to "' +
        destinationResolution.text +
        '" for destination "' +
        destinationResolution.sourceText +
        '"',
    );
  }
  const resolvedDestination = destinationResolution.text;
  const previousEvent = findPreviousDriveOriginEvent_(
    calendarId,
    driveEnd,
    sourceEventId,
    DRIVE_LOOKBACK_MINUTES,
  );
  if (previousEvent && previousEvent.location) {
    const previousLocationResolution = resolvePlaceNameAddress_(
      String(previousEvent.location || "").trim(),
      driveOpts.placeNameAddressRules,
    );
    if (previousLocationResolution.matched) {
      console.info(
        '[INFO] Resolved place name "' +
          previousLocationResolution.matchedPlaceName +
          '" to "' +
          previousLocationResolution.text +
          '" for previous event "' +
          (previousEvent.summary || previousEvent.id || "(No title)") +
          '"',
      );
    }
    const previousLocation = previousLocationResolution.text;
    if (sameLocation_(previousLocation, resolvedDestination)) {
      return {
        skipReason:
          "previous event is already at destination (" + previousLocation + ")",
      };
    }
    const previousDriveMinutes = getDriveMinutes_(
      previousLocation,
      resolvedDestination,
      driveDurationCache,
    );
    if (previousDriveMinutes !== null) {
      if (previousDriveMinutes <= driveOpts.minDriveMinutesToCreate) {
        return {
          skipReason:
            "previous event location is within threshold (" +
            previousDriveMinutes +
            "m <= " +
            driveOpts.minDriveMinutesToCreate +
            "m)",
        };
      }
      return {
        originAddress: previousLocation,
        driveMinutes: previousDriveMinutes,
        previousEventId: previousEvent.id || "",
      };
    }
    lookupFailures.push(previousLocation + " -> " + resolvedDestination);
    console.info(
      "[INFO] Drive lookup from previous event location failed; falling back to configured origin",
    );
  }

  const originResolution = resolvePlaceNameAddress_(
    String(driveOpts.originAddress || "").trim(),
    driveOpts.placeNameAddressRules,
  );
  const resolvedOrigin = originResolution.text;
  if (!resolvedOrigin) {
    if (lookupFailures.length) {
      return {
        skipReason:
          "route lookup failed and no default origin address is configured",
        routeLookupFailed: true,
        lookupFailures: lookupFailures,
      };
    }
    return { skipReason: "no default origin address is configured" };
  }
  const driveMinutes = getDriveMinutes_(
    resolvedOrigin,
    resolvedDestination,
    driveDurationCache,
  );
  if (driveMinutes === null) {
    lookupFailures.push(resolvedOrigin + " -> " + resolvedDestination);
    return {
      skipReason: "route lookup failed",
      routeLookupFailed: true,
      lookupFailures: lookupFailures,
    };
  }
  if (driveMinutes <= driveOpts.minDriveMinutesToCreate) {
    return {
      skipReason:
        "drive time is within threshold (" +
        driveMinutes +
        "m <= " +
        driveOpts.minDriveMinutesToCreate +
        "m)",
    };
  }
  return {
    originAddress: resolvedOrigin,
    driveMinutes: driveMinutes,
    previousEventId: "",
  };
}

/**
 * Picks the raw location text for a drive lookup.
 */
function resolveDriveDestinationCandidate_(evt) {
  const location = String((evt && evt.location) || "").trim();
  if (location) return { text: location, source: "location" };
  return { text: "", source: "" };
}

/**
 * Replaces configured place-name substrings with canonical addresses when they match.
 */
function resolvePlaceNameAddress_(text, rules) {
  const sourceText = String(text || "").trim();
  if (!sourceText || !rules || !rules.length) {
    return { text: sourceText, matched: false };
  }

  const sourceLower = sourceText.toLowerCase();
  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    if (!rule.placeNameLower) continue;
    if (sourceLower.indexOf(rule.placeNameLower) < 0) continue;
    return {
      text: rule.address,
      matched: true,
      matchedPlaceName: rule.placeName,
      sourceText: sourceText,
    };
  }

  return { text: sourceText, matched: false };
}

/**
 * Finds the most recent non-placeholder event (with location) ending within the lookback window.
 */
function findPreviousDriveOriginEvent_(
  calendarId,
  driveEnd,
  excludeEventId,
  lookbackMinutes,
) {
  let pageToken;
  let best = null;
  let bestEndMs = -1;
  const windowStartMs =
    driveEnd.getTime() -
    (Number(lookbackMinutes) || DRIVE_LOOKBACK_MINUTES) * 60 * 1000;

  do {
    const resp = Calendar.Events.list(calendarId, {
      showDeleted: false,
      singleEvents: true,
      orderBy: "startTime",
      timeMin: new Date(windowStartMs).toISOString(),
      timeMax: driveEnd.toISOString(),
      maxResults: 250,
      pageToken: pageToken,
    });

    (resp.items || []).forEach(function (ev) {
      if (excludeEventId && ev.id === excludeEventId) return;
      if (!ev.location || !String(ev.location).trim()) return;
      if (isManagedPlaceholderEvent_(ev)) return;
      if (
        (ev.start && ev.start.date && !ev.start.dateTime) ||
        (ev.end && ev.end.date && !ev.end.dateTime)
      )
        return;

      const endDate = eventResourceEndDate_(ev);
      if (!endDate || isNaN(endDate.getTime())) return;
      if (endDate.getTime() < windowStartMs) return;
      if (endDate.getTime() > driveEnd.getTime()) return;
      if (endDate.getTime() <= bestEndMs) return;
      best = ev;
      bestEndMs = endDate.getTime();
    });

    pageToken = resp.nextPageToken;
  } while (pageToken);

  return best;
}

/**
 * Returns a Date for an event resource end (or start) for ordering comparisons.
 */
function eventResourceEndDate_(ev) {
  const end = ev && (ev.end || ev.start);
  if (!end) return null;
  if (end.dateTime) return new Date(end.dateTime);
  if (end.date) return new Date(end.date + "T00:00:00");
  return null;
}

/**
 * Returns true when two location strings are equivalent after light normalization.
 */
function sameLocation_(a, b) {
  return normalizeLocation_(a) === normalizeLocation_(b);
}

/**
 * Normalizes location strings for comparisons.
 */
function normalizeLocation_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

/**
 * Returns true when an event is a managed placeholder from this script.
 */
function isManagedPlaceholderEvent_(ev) {
  const kind = (((ev || {}).extendedProperties || {}).private || {})
    .managedKind;
  return kind === "drive" || kind === "arrival";
}

/**
 * Computes drive minutes from origin to destination using Maps and per-run in-memory cache.
 */
function getDriveMinutes_(originAddress, destinationAddress, cache) {
  const key = (originAddress + "||" + destinationAddress).toLowerCase();
  if (cache.hasOwnProperty(key)) {
    console.info(
      '[INFO] Using cached drive time for "' +
        originAddress +
        '" -> "' +
        destinationAddress +
        '"',
    );
    return cache[key];
  }

  try {
    const directions = Maps.newDirectionFinder()
      .setOrigin(originAddress)
      .setDestination(destinationAddress)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();
    logDirectionsDiagnostics_(originAddress, destinationAddress, directions);
    const minutes = extractDriveMinutesFromDirections_(directions);
    cache[key] = minutes;
    return minutes;
  } catch (e) {
    const errName = e && e.name ? e.name : "Error";
    const errMessage = e && e.message ? e.message : String(e);
    console.error(
      '[ERROR] Drive lookup failed from "' +
        originAddress +
        '" to "' +
        destinationAddress +
        '": ' +
        errName +
        ": " +
        errMessage +
        (e && e.stack ? "\n" + e.stack : ""),
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
  return roundDriveMinutesForPlaceholder_(minutes);
}

/**
 * Logs response-level diagnostics for a Maps directions lookup.
 */
function logDirectionsDiagnostics_(
  originAddress,
  destinationAddress,
  directions,
) {
  if (!directions || typeof directions !== "object") {
    console.warn(
      '[WARN] Directions lookup returned no response for "' +
        originAddress +
        '" -> "' +
        destinationAddress +
        '"',
    );
    return;
  }

  if (directions.status && directions.status !== "OK") {
    console.warn(
      '[WARN] Directions status for "' +
        originAddress +
        '" -> "' +
        destinationAddress +
        '": ' +
        directions.status,
    );
  }

  const waypoints = directions.geocoded_waypoints || [];
  waypoints.forEach(function (wp, idx) {
    if (!wp) return;
    const pieces = [];
    if (wp.geocoder_status)
      pieces.push("geocoder_status=" + wp.geocoder_status);
    if (wp.partial_match) pieces.push("partial_match=true");
    if (wp.place_id) pieces.push("place_id=" + wp.place_id);
    if (pieces.length) {
      console.info(
        "[INFO] Directions waypoint " +
          idx +
          ' for "' +
          originAddress +
          '" -> "' +
          destinationAddress +
          '": ' +
          pieces.join(", "),
      );
    }
  });

  const routes = directions.routes || [];
  if (!routes.length) {
    console.warn(
      '[WARN] Directions lookup for "' +
        originAddress +
        '" -> "' +
        destinationAddress +
        '" returned no routes',
    );
    return;
  }

  const legs = routes[0].legs || [];
  if (!legs.length) {
    console.warn(
      '[WARN] Directions lookup for "' +
        originAddress +
        '" -> "' +
        destinationAddress +
        '" returned no legs on first route',
    );
    return;
  }

  const duration = legs[0] && legs[0].duration ? legs[0].duration : null;
  if (!duration || typeof duration.value !== "number") {
    console.warn(
      '[WARN] Directions lookup for "' +
        originAddress +
        '" -> "' +
        destinationAddress +
        '" returned no usable duration on first leg',
    );
  }
}

/**
 * Rounds a minute count for drive placeholders:
 * - < 10 minutes: 0, so the placeholder is skipped by threshold logic
 * - 10-15 minutes: 15
 * - 16-20 minutes: 20
 * - > 20 minutes: next 10-minute bucket
 */
function roundDriveMinutesForPlaceholder_(minutes) {
  if (typeof minutes !== "number" || !isFinite(minutes)) return null;
  if (minutes <= 0) return 0;
  if (minutes < 10) return 0;
  if (minutes <= 15) return 15;
  if (minutes <= 20) return 20;
  return Math.ceil(minutes / 10) * 10;
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
  sourceEventTitle,
  driveTitle,
  driveStart,
  driveEnd,
  driveHash,
  originAddress,
  destinationAddress,
  attendees,
  driveOriginEventId,
) {
  const directionsUrl = buildGoogleMapsDirectionsUrl_(
    originAddress,
    destinationAddress,
  );
  const driveDescription =
    "<p><strong>Managed drive-time placeholder</strong></p>" +
    "<p><strong>From:</strong> " +
    escapeHtml_(originAddress) +
    "<br><strong>To:</strong> " +
    escapeHtml_(destinationAddress) +
    "</p>" +
    '<p><a href="' +
    escapeHtml_(directionsUrl) +
    '">Open driving directions in Google Maps</a></p>' +
    "<p><strong>Source event:</strong> " +
    escapeHtml_(sourceEventTitle || sourceEventId) +
    (driveOriginEventId
      ? "<br><strong>Drive origin event:</strong> " +
        escapeHtml_(driveOriginEventId)
      : "") +
    "</p>";

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
    attendees: (attendees || []).map(function (attendee) {
      return toCalendarAttendeeResource_(attendee);
    }),
    extendedProperties: {
      private: {
        managedKind: "drive",
        sourceFeed: feedHash,
        sourceUrl: mapping.feedUrl,
        sourceFeedName: mappingFeedName_(mapping),
        sourceUid: evt.uid,
        syncKey: driveSyncKey,
        sourceSyncKey: sourceSyncKey,
        sourceEventId: sourceEventId,
        driveOriginEventId: driveOriginEventId || "",
        syncHash: driveHash,
      },
    },
  };
}

/**
 * Builds a Google Maps driving-directions URL from origin and destination.
 */
function buildGoogleMapsDirectionsUrl_(originAddress, destinationAddress) {
  return (
    "https://www.google.com/maps/dir/?api=1&travelmode=driving&origin=" +
    encodeURIComponent(String(originAddress || "")) +
    "&destination=" +
    encodeURIComponent(String(destinationAddress || ""))
  );
}

/**
 * Escapes text for safe use in Calendar event HTML descriptions.
 */
function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
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
  sourceEventTitle,
  arrivalTitle,
  arrivalStart,
  arrivalEnd,
  arrivalHash,
  arrivalMinutes,
  attendees,
) {
  const arrivalDescription =
    "Managed advanced-arrival placeholder.\n" +
    "Lead time: " +
    arrivalMinutes +
    " minutes\n" +
    "Source event: " +
    (sourceEventTitle || sourceEventId);

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
    attendees: (attendees || []).map(function (attendee) {
      return toCalendarAttendeeResource_(attendee);
    }),
    extendedProperties: {
      private: {
        managedKind: "arrival",
        sourceFeed: feedHash,
        sourceUrl: mapping.feedUrl,
        sourceFeedName: mappingFeedName_(mapping),
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
  sourceEventTitle,
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
      sourceEventTitle: sourceEventTitle,
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
  sourceEventTitle,
  arrivalStart,
  arrivalEnd,
  arrivalTitle,
  arrivalMinutes,
) {
  return sha256Hex_(
    JSON.stringify({
      sourceSyncKey: sourceSyncKey,
      sourceEventId: sourceEventId,
      sourceEventTitle: sourceEventTitle,
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
 * Converts an attendee value into the Calendar API attendee resource shape.
 */
function toCalendarAttendeeResource_(attendee) {
  if (!attendee) return { email: "" };
  if (typeof attendee === "string") {
    return { email: attendee };
  }

  const resource = { email: String(attendee.email || "") };
  if (attendee.responseStatus)
    resource.responseStatus = String(attendee.responseStatus);
  if (attendee.self) resource.self = true;
  return resource;
}

/**
 * Returns true when the target calendar's attendee entry is marked declined.
 */
function isTargetCalendarDeclinedEvent_(ev, calendarId) {
  return !!getTargetCalendarAttendee_(ev, calendarId, "declined");
}

/**
 * Returns true when every attendee on a destination event has declined.
 */
function isAllAttendeesDeclinedEvent_(ev) {
  const attendees = (ev && ev.attendees) || [];
  if (!attendees.length) return false;
  for (let i = 0; i < attendees.length; i++) {
    const attendee = attendees[i];
    if (
      !attendee ||
      String(attendee.responseStatus || "").toLowerCase() !== "declined"
    ) {
      return false;
    }
  }
  return true;
}

/**
 * Returns the target calendar's attendee entry when it matches the requested status.
 */
function getTargetCalendarAttendee_(ev, calendarId, responseStatus) {
  const attendees = (ev && ev.attendees) || [];
  const targetEmail = String(calendarId || "")
    .trim()
    .toLowerCase();
  if (!targetEmail) return null;
  for (let i = 0; i < attendees.length; i++) {
    const attendee = attendees[i];
    if (!attendee) continue;
    const email = String(attendee.email || "")
      .trim()
      .toLowerCase();
    if (email !== targetEmail) continue;
    if (
      responseStatus &&
      String(attendee.responseStatus || "").toLowerCase() !== responseStatus
    ) {
      continue;
    }
    return attendee;
  }
  return null;
}

/**
 * Keeps only the target calendar's attendee entry and marks it declined.
 */
function buildDeclinedAttendees_(ev, calendarId) {
  const targetAttendee = getTargetCalendarAttendee_(ev, calendarId, null);
  const email =
    targetAttendee && targetAttendee.email
      ? String(targetAttendee.email).trim().toLowerCase()
      : String(calendarId || "")
          .trim()
          .toLowerCase();
  if (!email) return [];
  const declined = {
    email: email,
    responseStatus: "declined",
  };
  if (targetAttendee && targetAttendee.self) declined.self = true;
  return [declined];
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
