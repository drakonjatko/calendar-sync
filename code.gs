/**
 * calendar-sync — Microsoft Outlook to Google Calendar
 *
 * Fetches busy times from Office 365 ICS feeds and mirrors them into
 * Google Calendar. No event details (titles, attendees, descriptions)
 * are transferred — only start/end times and busy status.
 *
 * Entry point: syncCalendar()
 * Configuration: Config.gs
 * Timezone mappings: Timezones.gs
 */

/**
 * Runs a sync for every calendar source defined in CONFIG.
 * Set this function as a time-based trigger (e.g. every 15 minutes).
 */
function syncCalendar() {
  CONFIG.forEach((config, index) => {
    console.log(`Starting calendar sync for config ${index + 1}: ${config.EVENT_TITLE_PREFIX}`);
    try {
      syncSingleCalendar(config);
    } catch (error) {
      console.error(`Error during sync for config ${index + 1}:`, error);
    }
  });
}

/**
 * Syncs one calendar source end-to-end:
 * fetch ICS → parse events → diff against Google Calendar → create/update/delete.
 *
 * @param {Object} config - A single entry from the CONFIG array.
 */
function syncSingleCalendar(config) {
  // Fetch and parse ICS data
  const icsData = fetchICSData(config);
  if (!icsData) {
    console.error('Failed to fetch ICS data');
    return;
  }
  
  const icsEvents = parseICSEvents(icsData);
  console.log(`Found ${icsEvents.length} busy slots in ICS feed`);
  
  // Get target calendar
  const calendar = CalendarApp.getCalendarById(config.GOOGLE_CALENDAR_ID);
  
  // Get existing synced events:
  //   - Series masters via Advanced Calendar API (singleEvents: false)
  //   - Individual/exception events via CalendarApp
  const existingSeriesMasters = getExistingSeriesMasters(config.GOOGLE_CALENDAR_ID, config);
  const existingSyncedEvents = getExistingSyncedEvents(config.GOOGLE_CALENDAR_ID, config);
  console.log(`Found ${existingSeriesMasters.length} synced series, ${existingSyncedEvents.length} synced instances`);

  const processedTrackingIds = new Set();
  let addedCount = 0;
  let updatedCount = 0;

  // Create or update events from ICS
  icsEvents.forEach(icsEvent => {
    if (!icsEvent.uid) return;

    // Skip all-day events entirely.
    if (icsEvent.isAllDay) {
      console.log('SKIP all-day trackingId:', icsEvent.trackingId);
      return;
    }

    // Any VEVENT with a RECURRENCE-ID is an exception to a series — either a
    // modified occurrence (new time) or a cancelled one. In both cases, remove
    // the corresponding Google Calendar series occurrence first so the original
    // slot does not remain visible alongside the exception.
    if (icsEvent.recurrenceId && !icsEvent.rrule) {
      const parentSeries = findSeriesByTrackingId(existingSeriesMasters, icsEvent.uid);
      if (parentSeries) deleteCancelledOccurrence(config.GOOGLE_CALENDAR_ID, parentSeries.id, icsEvent);
    }

    // Skip cancelled events. By not adding their trackingId to processedTrackingIds,
    // any previously synced standalone version is cleaned up in the deletion pass below.
    if (icsEvent.status === 'CANCELLED') {
      console.log('SKIP cancelled trackingId:', icsEvent.trackingId);
      return;
    }

    processedTrackingIds.add(icsEvent.trackingId);

    if (icsEvent.rrule && !icsEvent.recurrenceId) {
      // Recurring series master — manage via Advanced Calendar API so Google Calendar
      // stores it as a real recurring event (respects DST, shows as series to the user).
      const existingSeries = findSeriesByTrackingId(existingSeriesMasters, icsEvent.trackingId);
      if (existingSeries) {
        if (updateRecurringSeriesIfChanged(existingSeries, icsEvent, config)) updatedCount++;
      } else {
        if (createRecurringGoogleEvent(config.GOOGLE_CALENDAR_ID, icsEvent, config)) addedCount++;
      }
    } else {
      // Single event or modified exception — create/update a standalone event.
      // The series occurrence for this slot was already removed above (if recurrenceId was set).
      const existingEvent = findEventByUID(existingSyncedEvents, icsEvent.trackingId);
      if (existingEvent) {
        if (updateEventIfChanged(existingEvent, icsEvent, config)) updatedCount++;
      } else {
        if (createGoogleEvent(calendar, icsEvent, config)) addedCount++;
      }
    }
  });

  // Delete recurring series that no longer exist in the ICS feed.
  let deletedCount = 0;
  const deletedSeriesIds = new Set();
  existingSeriesMasters.forEach(series => {
    const trackingId = extractUIDFromDescription(series.description);
    if (trackingId && !processedTrackingIds.has(trackingId)) {
      try {
        Calendar.Events.remove(config.GOOGLE_CALENDAR_ID, series.id);
        deletedSeriesIds.add(trackingId); // prevent attempting to delete now-gone instances
        deletedCount++;
        console.log('Deleted series:', series.summary);
      } catch (error) {
        console.error('Error deleting series:', error);
      }
    }
  });

  // Delete individual/exception events that no longer exist in the ICS feed.
  // Skip instances whose parent series was already deleted above.
  // existingSyncedEvents contains only standalone API resources (no series instances),
  // so there is no risk of accidentally matching series occurrences here.
  existingSyncedEvents.forEach(event => {
    const trackingId = extractUIDFromDescription(event.description);
    if (trackingId && !processedTrackingIds.has(trackingId) && !deletedSeriesIds.has(trackingId)) {
      try {
        Calendar.Events.remove(config.GOOGLE_CALENDAR_ID, event.id);
        deletedCount++;
        console.log('Deleted event:', event.summary);
      } catch (error) {
        console.error('Error deleting event:', error);
      }
    }
  });

  console.log(`Sync completed for "${config.EVENT_TITLE_PREFIX}". Added: ${addedCount}, Updated: ${updatedCount}, Deleted: ${deletedCount}`);
}

/**
 * Downloads the ICS calendar feed from the given URL.
 *
 * @param {Object} config - Config entry containing ICS_URL.
 * @returns {string|null} Raw ICS text, or null on error.
 */
function fetchICSData(config) {
  try {
    const response = UrlFetchApp.fetch(config.ICS_URL, {
      method: 'GET',
      headers: {
        'User-Agent': 'GoogleAppsScript-CalendarSync/1.0'
      }
    });
    
    if (response.getResponseCode() === 200) {
      return response.getContentText();
    } else {
      console.error('HTTP Error:', response.getResponseCode());
      return null;
    }
  } catch (error) {
    console.error('Error fetching ICS data:', error);
    return null;
  }
}

/**
 * Parses raw ICS text into an array of event objects.
 * Events without a UID or start time are skipped. Single events and exception
 * occurrences outside the sync window (30 days ago to 1 year ahead) are skipped;
 * series masters (RRULE present) bypass the date filter because their DTSTART
 * can be years in the past while future occurrences still need to be synced.
 *
 * @param {string} icsData - Raw ICS calendar text.
 * @returns {Object[]} Array of parsed event objects.
 */
function parseICSEvents(icsData) {
  const events = [];

  // RFC 5545 line folding: long lines are split by inserting CRLF + a space or tab.
  // Rejoin them before parsing so property values are never truncated.
  const rawLines = icsData.split('\n');
  const lines = [];
  for (let i = 0; i < rawLines.length; i++) {
    const raw = rawLines[i].replace(/\r$/, ''); // strip CR from CRLF
    if ((raw.startsWith(' ') || raw.startsWith('\t')) && lines.length > 0) {
      lines[lines.length - 1] += raw.slice(1); // append continuation (drop leading whitespace)
    } else {
      lines.push(raw);
    }
  }

  let currentEvent = null;
  const today = new Date();
  today.setDate(today.getDate() - 30)
  today.setHours(0, 0, 0, 0);
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();    
    if (line === 'BEGIN:VEVENT') {
      currentEvent = {};
    } else if (line === 'END:VEVENT' && currentEvent) {
      if (!currentEvent.uid) {
        console.log('SKIP no-uid:', JSON.stringify(currentEvent));
      } else if (!currentEvent.startTime) {
        console.log('SKIP no-startTime uid:', currentEvent.uid);
      } else if (!currentEvent.rrule && currentEvent.startTime < today) {
        // Series masters have their original DTSTART (potentially years ago) but their
        // future occurrences still matter — skip the date filter for them.
        console.log('SKIP too-old uid:', currentEvent.uid, 'start:', currentEvent.startTime);
      } else {
        if (currentEvent.rrule && !currentEvent.recurrenceId) {
          // Series master — track by uid alone (represents the whole series).
          currentEvent.trackingId = currentEvent.uid;
        } else if (currentEvent.recurrenceId) {
          // Modified occurrence — track by uid + original scheduled time.
          currentEvent.trackingId = currentEvent.uid + '|' + currentEvent.recurrenceId;
        } else {
          // Single event — track by uid + start time.
          currentEvent.trackingId = currentEvent.uid + '|' + currentEvent.dtStartRaw;
        }
        console.log('PARSED event trackingId:', currentEvent.trackingId, 'start:', currentEvent.startTime, 'status:', currentEvent.status || 'none', 'isAllDay:', !!currentEvent.isAllDay, 'rrule:', currentEvent.rrule || 'none');
        events.push(currentEvent);
      }
      currentEvent = null;
    } else if (currentEvent && line.includes(':')) {
      parseEventProperty(line, currentEvent);
    }
  }
  
  return events;
}

/**
 * Extracts a single ICS property from a VEVENT line and writes it onto
 * the event object. Handles parameterized properties (e.g. DTSTART;TZID=...).
 *
 * @param {string} line  - A single ICS content line (already unfolded).
 * @param {Object} event - The event object being built.
 */
function parseEventProperty(line, event) {
  const colonIndex = line.indexOf(':');
  property = line.substring(0, colonIndex);
  const value = line.substring(colonIndex + 1);
  const propertyParams = []
  if (property.includes(';')) {
    const parts = property.split(';');
    property = parts[0]; // Main property name (e.g., "DTSTART")
    // Parse parameters like "TZID=China Standard Time"
    for (let i = 1; i < parts.length; i++) {
      const param = parts[i];
      if (param.includes('=')) {
        const [paramName, paramValue] = param.split('=');
        propertyParams[paramName] = paramValue;
      }
    }
  }
  switch (property) {
    case 'DTSTART': {
      const startTzid = extractTimezone(propertyParams);
      event.dtStartRaw = value; // raw string used for trackingId
      // Store IANA timezone so recurring events can be created with the correct
      // local time — prevents the event time from shifting across DST transitions.
      event.timezone = startTzid
        ? convertMicrosoftTimezone(startTzid)
        : (value.endsWith('Z') ? 'UTC' : null);
      event.startTime = parseICSDateTime(value, startTzid);
      event.isAllDay = !value.includes('T');
      break;
    }
    case 'DTEND':
      event.endTime = parseICSDateTime(value, extractTimezone(propertyParams));
      break;
    case 'UID':
      event.uid = value;
      break;
    case 'X-MICROSOFT-CDO-BUSYSTATUS':
      event.busyStatus = value; // BUSY, TENTATIVE, FREE
      break;
    case 'STATUS':
      event.status = value; // CONFIRMED, TENTATIVE, CANCELLED
      break;
    case 'RECURRENCE-ID': {
      event.recurrenceId = value;
      const recTzid = extractTimezone(propertyParams);
      if (recTzid) event.recurrenceTzid = recTzid; // used to parse the original time when deleting cancelled occurrences
      break;
    }
    case 'RRULE':
      // Store as a complete RRULE: string ready for Google Calendar's recurrence field.
      event.rrule = 'RRULE:' + value;
      break;
    case 'EXDATE': {
      // Dates to exclude from the series. Convert TZID from Microsoft to IANA format.
      // Google Calendar requires each excluded date on its own EXDATE line,
      // so split comma-separated multi-date values into separate entries.
      if (!event.exdates) event.exdates = [];
      const exTzid = extractTimezone(propertyParams);
      const exIana = exTzid ? convertMicrosoftTimezone(exTzid) : null;
      const prefix = exIana ? 'EXDATE;TZID=' + exIana + ':' : 'EXDATE:';
      value.split(',').forEach(date => event.exdates.push(prefix + date.trim()));
      break;
    }
  }
}

/**
 * Returns the TZID parameter value from a parsed property parameter map,
 * or null if not present.
 */
function extractTimezone(propertyParams) {
  return propertyParams.TZID || null;
}

function getOffsetString(year, month, day, hour, minute, second, tz) {
  const temp = new Date(year, month, day, hour, minute, second);

  return Utilities.formatDate(temp, tz, "XXX"); 
  // returns "+03:00", "+08:00", etc.
}

function pad(n) {
  return n.toString().padStart(2, '0');
}

/**
 * Constructs a UTC Date from a local wall-clock time in a given Microsoft timezone.
 *
 * Google Apps Script's Date constructor treats bare ISO strings as local time,
 * which is unreliable in a server environment. The workaround is:
 *   1. Construct a rough local Date to let Utilities.formatDate() calculate the
 *      correct UTC offset for that instant (accounting for DST).
 *   2. Build a full ISO 8601 string with that offset appended.
 *   3. Parse it with new Date() — the explicit offset makes the result correct.
 *
 * @param {number} year
 * @param {number} month  - 0-based (0 = January).
 * @param {number} day
 * @param {number} hour
 * @param {number} minute
 * @param {number} second
 * @param {string} timezone - Microsoft timezone name (e.g. 'W. Europe Standard Time').
 * @returns {Date} UTC Date object.
 */
function buildDate(year, month, day, hour, minute, second, timezone) {
  const tz = convertMicrosoftTimezone(timezone);

  const offset = getOffsetString(year, month, day, hour, minute, second, tz);

  const iso =
    `${year}-${pad(month+1)}-${pad(day)}T` +
    `${pad(hour)}:${pad(minute)}:${pad(second)}${offset}`;

  return new Date(iso);
}

/**
 * Converts an ICS datetime string to a JavaScript Date.
 *
 * Handles three formats:
 *   - Date-only (YYYYMMDD): returned as a local midnight Date (all-day event).
 *   - UTC datetime (YYYYMMDDTHHmmssZ): parsed directly as UTC.
 *   - Local datetime with TZID param (YYYYMMDDTHHmmss): converted via buildDate().
 *
 * @param {string} dateTimeString - ICS datetime value.
 * @param {string|null} timezone  - Microsoft timezone name from the TZID parameter, or null.
 * @returns {Date}
 */
function parseICSDateTime(dateTimeString, timezone) {
  if (!dateTimeString.includes('T')) {
    // All-day event - no timezone needed
    const year = parseInt(dateTimeString.substring(0, 4));
    const month = parseInt(dateTimeString.substring(4, 6)) - 1;
    const day = parseInt(dateTimeString.substring(6, 8));
    return new Date(year, month, day);
  } else {
    // Timed event
    const year = parseInt(dateTimeString.substring(0, 4));
    const month = parseInt(dateTimeString.substring(4, 6)) - 1;
    const day = parseInt(dateTimeString.substring(6, 8));
    const hour = parseInt(dateTimeString.substring(9, 11));
    const minute = parseInt(dateTimeString.substring(11, 13));
    const second = parseInt(dateTimeString.substring(13, 15)) || 0;
    
    if (dateTimeString.endsWith('Z')) {
      // UTC time
      return new Date(Date.UTC(year, month, day, hour, minute, second));
    } else if (timezone) {
      // Convert from source timezone to UTC
      return buildDate(year, month, day, hour, minute, second, timezone);
    } else {
      // Local time
      return new Date(year, month, day, hour, minute, second);
    }
  }
}

/**
 * Fetches all previously synced standalone (non-series) events from the target
 * Google Calendar within the sync window using the Advanced Calendar API.
 *
 * Using the Advanced API with recurringEventId filtering guarantees we only
 * return truly individual events — not CalendarApp expansions of recurring
 * series occurrences, which share the series master's description and trackingId
 * and would otherwise cause phantom duplicates in the cleanup pass.
 *
 * @param {string} calendarId - Target Google Calendar ID.
 * @param {Object} config     - Config entry.
 * @returns {Object[]} Array of Google Calendar API event resource objects.
 */
function getExistingSyncedEvents(calendarId, config) {
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 30);
  startDate.setHours(0, 0, 0, 0);
  const futureDate = new Date();
  futureDate.setFullYear(futureDate.getFullYear() + 1);
  try {
    const result = Calendar.Events.list(calendarId, {
      singleEvents: true,
      timeMin: startDate.toISOString(),
      timeMax: futureDate.toISOString(),
      maxResults: 2500
    });
    return (result.items || []).filter(event =>
      !event.recurringEventId &&
      event.summary && event.summary.includes(config.EVENT_TITLE_PREFIX) &&
      event.description && event.description.includes('[UID:')
    );
  } catch (error) {
    console.error('Error fetching synced events:', error);
    return [];
  }
}

/**
 * Finds a previously synced Google Calendar event whose description contains
 * the given Microsoft UID marker.
 *
 * @param {Object[]} existingEvents - API event resources from getExistingSyncedEvents.
 * @param {string} uid
 * @returns {Object|undefined}
 */
function findEventByUID(existingEvents, uid) {
  return existingEvents.find(event =>
    extractUIDFromDescription(event.description) === uid
  );
}

/**
 * Extracts the Microsoft event UID from a Google Calendar event description.
 * Returns null if the description is missing or does not contain a [UID:...] marker.
 *
 * @param {string|null} description
 * @returns {string|null}
 */
function extractUIDFromDescription(description) {
  if (!description) return null;
  const match = description.match(/\[UID:([^\]]+)\]/);
  return match ? match[1] : null;
}

/**
 * Strips common formatting issues that cause Google Calendar to reject an RRULE:
 *   - Trailing semicolons (e.g. "RRULE:FREQ=WEEKLY;BYDAY=FR;")
 *   - Stray whitespace or carriage returns
 *
 * @param {string} rrule - Full RRULE: string.
 * @returns {string} Sanitized RRULE string.
 */
function sanitizeRrule(rrule) {
  return rrule.trim().replace(/;+$/, '');
}

/**
 * Builds a Google Calendar API-compatible recurrence array from a parsed ICS event.
 * Includes the RRULE and any EXDATE exclusion lines.
 *
 * @param {Object} icsEvent - Parsed ICS event with rrule and optional exdates.
 * @returns {string[]}
 */
function buildRecurrence(icsEvent) {
  const parts = [sanitizeRrule(icsEvent.rrule)];
  if (icsEvent.exdates) {
    icsEvent.exdates.forEach(ex => parts.push(ex));
  }
  return parts;
}

/**
 * Builds a Google Calendar API start/end object from a UTC Date and IANA timezone.
 * Passing the local datetime + timezone (instead of a bare UTC string) ensures that
 * recurring events stay at the correct wall-clock time across DST transitions.
 *
 * @param {Date}        date     - UTC Date object.
 * @param {string|null} timezone - IANA timezone name, or null for UTC.
 * @returns {{dateTime: string, timeZone: string}}
 */
function googleDateTime(date, timezone) {
  const tz = timezone || 'UTC';
  const localIso = Utilities.formatDate(date, tz, "yyyy-MM-dd'T'HH:mm:ss");
  return { dateTime: localIso, timeZone: tz };
}

/**
 * Fetches all previously synced recurring event series from the target Google Calendar
 * using the Advanced Calendar Service (Calendar.Events.list with singleEvents: false),
 * which returns series masters rather than individual occurrences.
 *
 * PREREQUISITE: Enable the Advanced Calendar Service in the GAS project:
 *   Apps Script editor → Services → Google Calendar API.
 *
 * @param {string} calendarId - Target Google Calendar ID.
 * @param {Object} config     - Config entry.
 * @returns {Object[]} Array of Google Calendar API event resource objects.
 */
function getExistingSeriesMasters(calendarId, config) {
  // Use a wide lookback to catch series that started long ago but are still active.
  const fiveYearsAgo = new Date();
  fiveYearsAgo.setFullYear(fiveYearsAgo.getFullYear() - 5);
  try {
    const result = Calendar.Events.list(calendarId, {
      singleEvents: false,
      timeMin: fiveYearsAgo.toISOString(),
      q: config.EVENT_TITLE_PREFIX,
      maxResults: 500
    });
    return (result.items || []).filter(event =>
      event.recurrence &&
      event.description &&
      event.description.includes('[UID:')
    );
  } catch (error) {
    console.error('Error fetching series masters:', error);
    return [];
  }
}

/**
 * Finds a synced recurring series whose [UID:...] description marker matches the given trackingId.
 *
 * @param {Object[]} seriesMasters - Array of Google Calendar API event resources.
 * @param {string}   trackingId
 * @returns {Object|undefined}
 */
function findSeriesByTrackingId(seriesMasters, trackingId) {
  return seriesMasters.find(series => extractUIDFromDescription(series.description) === trackingId);
}

/**
 * Creates a new recurring event series in Google Calendar using the Advanced Calendar API.
 * The series start time is expressed in the event's original timezone so that future
 * occurrences stay at the correct local wall-clock time across DST transitions.
 *
 * @param {string} calendarId - Target Google Calendar ID.
 * @param {Object} icsEvent   - Parsed ICS series master event.
 * @param {Object} config     - Config entry.
 * @returns {boolean} True on success, false on error.
 */
function createRecurringGoogleEvent(calendarId, icsEvent, config) {
  try {
    const recurrence = buildRecurrence(icsEvent);
    console.log('Creating series recurrence:', JSON.stringify(recurrence),
                'start:', googleDateTime(icsEvent.startTime, icsEvent.timezone),
                'end:',   googleDateTime(icsEvent.endTime,   icsEvent.timezone));
    Calendar.Events.insert({
      summary: `${config.EVENT_TITLE_PREFIX} meeting`,
      description: `Synced from Microsoft Calendar\n[UID:${icsEvent.trackingId}]`,
      start: googleDateTime(icsEvent.startTime, icsEvent.timezone),
      end:   googleDateTime(icsEvent.endTime,   icsEvent.timezone),
      recurrence: recurrence,
      reminders: {
        useDefault: false,
        overrides: [{ method: 'popup', minutes: config.NOTIFICATION_MINUTES }]
      }
    }, calendarId);
    console.log('Created series:', config.EVENT_TITLE_PREFIX, 'starting', icsEvent.startTime, 'rrule:', icsEvent.rrule);
    return true;
  } catch (error) {
    console.error('Error creating recurring event:', error);
    return false;
  }
}

/**
 * Updates a recurring series in Google Calendar if its start time, end time, or
 * recurrence rule has changed in the ICS feed. No-op if nothing changed.
 *
 * @param {Object} seriesEvent - Google Calendar API event resource for the series master.
 * @param {Object} icsEvent    - Updated ICS series master.
 * @param {Object} config      - Config entry.
 * @returns {boolean} True if updated, false otherwise.
 */
function updateRecurringSeriesIfChanged(seriesEvent, icsEvent, config) {
  try {
    const currentStart = new Date(seriesEvent.start.dateTime || seriesEvent.start.date);
    const currentEnd   = new Date(seriesEvent.end.dateTime   || seriesEvent.end.date);
    const currentRRule = (seriesEvent.recurrence || []).find(r => r.startsWith('RRULE:')) || '';

    const startChanged  = currentStart.getTime() !== icsEvent.startTime.getTime();
    const endChanged    = currentEnd.getTime()   !== icsEvent.endTime.getTime();
    const rruleChanged  = currentRRule           !== icsEvent.rrule;

    if (!startChanged && !endChanged && !rruleChanged) return false;

    Calendar.Events.patch({
      start:      googleDateTime(icsEvent.startTime, icsEvent.timezone),
      end:        googleDateTime(icsEvent.endTime,   icsEvent.timezone),
      recurrence: buildRecurrence(icsEvent)
    }, config.GOOGLE_CALENDAR_ID, seriesEvent.id);
    console.log('Updated series:', config.EVENT_TITLE_PREFIX, 'starting', icsEvent.startTime);
    return true;
  } catch (error) {
    console.error('Error updating recurring event:', error);
    return false;
  }
}

/**
 * Creates a new event in the target Google Calendar from a parsed ICS event.
 * The event description always includes [UID:...] so future syncs can
 * identify and update or delete the event.
 *
 * @param {Calendar} calendar  - Target GoogleAppsScript Calendar.
 * @param {Object} icsEvent    - Parsed ICS event object.
 * @param {Object} config      - Config entry.
 * @returns {boolean} True on success, false on error.
 */
function createGoogleEvent(calendar, icsEvent, config) {
  try {
    const description = `Synced from Microsoft Calendar\n[UID:${icsEvent.trackingId}]`;
    // All-day events are filtered out before this function is called.
    const title = `${config.EVENT_TITLE_PREFIX} meeting`;
    
    const newEvent = calendar.createEvent(
      title,
      icsEvent.startTime,
      icsEvent.endTime,
      {
        description: description
      }
    );

    // Set Google Calendar status based on Microsoft status
    if (icsEvent.busyStatus === 'TENTATIVE' || icsEvent.status === 'TENTATIVE') {
      // Google Calendar doesn't have direct "tentative" but we can use transparency
      newEvent.setTransparency(CalendarApp.EventTransparency.TRANSPARENT);
    }
    
    // Add 5-minute notification
    newEvent.addPopupReminder(config.NOTIFICATION_MINUTES);
    
    console.log('Created event:', config.EVENT_TITLE_PREFIX, 'at', icsEvent.startTime);
    return true;
    
  } catch (error) {
    console.error('Error creating event:', error);
    return false;
  }
}

/**
 * Removes a specific occurrence from a Google Calendar recurring series.
 *
 * Used for both cancelled occurrences (STATUS:CANCELLED + RECURRENCE-ID) and
 * modified occurrences (RECURRENCE-ID without CANCELLED, where the occurrence
 * has been rescheduled to a different time). In both cases the original series
 * slot must be removed so it does not appear alongside the exception event.
 *
 * @param {string} calendarId      - Target Google Calendar ID.
 * @param {string} seriesId        - Google Calendar event ID of the series master.
 * @param {Object} exceptionIcsEvent - Parsed ICS event with recurrenceId set.
 */
function deleteCancelledOccurrence(calendarId, seriesId, exceptionIcsEvent) {
  // Prefer the TZID from the RECURRENCE-ID property itself; fall back to the
  // event's DTSTART timezone so floating RECURRENCE-ID values are resolved
  // correctly instead of being treated as GAS server local time.
  const cancelledIcsEvent = exceptionIcsEvent; // kept for internal readability
  try {
    const tzid = cancelledIcsEvent.recurrenceTzid || cancelledIcsEvent.timezone || null;
    const originalTime = parseICSDateTime(
      cancelledIcsEvent.recurrenceId,
      tzid
    );
    const result = Calendar.Events.instances(calendarId, seriesId, {
      timeMin: new Date(originalTime.getTime() - 86400000).toISOString(),
      timeMax: new Date(originalTime.getTime() + 86400000).toISOString(),
      maxResults: 10
    });
    const instance = (result.items || []).find(inst => {
      const orig = inst.originalStartTime;
      if (!orig) return false;
      const t = new Date(orig.dateTime || orig.date).getTime();
      return Math.abs(t - originalTime.getTime()) < 60000; // within 1 minute
    });
    if (instance) {
      Calendar.Events.remove(calendarId, instance.id);
      console.log('Removed cancelled occurrence:', cancelledIcsEvent.recurrenceId);
    } else {
      console.log('Cancelled occurrence not found in series (already removed?):', cancelledIcsEvent.recurrenceId);
    }
  } catch (error) {
    console.error('Error removing cancelled occurrence:', error);
  }
}

/**
 * Updates the start/end time of an existing standalone Google Calendar event if it
 * has changed in the ICS feed. No-op if times are identical.
 * Uses the Advanced Calendar API to work directly with the resource objects
 * returned by getExistingSyncedEvents, avoiding an extra CalendarApp lookup.
 *
 * @param {Object} googleEvent - Google Calendar API event resource.
 * @param {Object} icsEvent    - Updated ICS event.
 * @param {Object} config      - Config entry.
 * @returns {boolean} True if the event was updated, false otherwise.
 */
function updateEventIfChanged(googleEvent, icsEvent, config) {
  try {
    const currentStart = new Date(googleEvent.start.dateTime || googleEvent.start.date);
    const currentEnd   = new Date(googleEvent.end.dateTime   || googleEvent.end.date);
    if (currentStart.getTime() === icsEvent.startTime.getTime() &&
        currentEnd.getTime()   === icsEvent.endTime.getTime()) {
      return false;
    }
    Calendar.Events.patch({
      start: { dateTime: icsEvent.startTime.toISOString() },
      end:   { dateTime: icsEvent.endTime.toISOString() }
    }, config.GOOGLE_CALENDAR_ID, googleEvent.id);
    console.log('Updated event time:', config.EVENT_TITLE_PREFIX, 'to', icsEvent.startTime);
    return true;
  } catch (error) {
    console.error('Error updating event:', error);
    return false;
  }
}