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
  
  // Get existing synced events from today onwards
  const existingSyncedEvents = getExistingSyncedEvents(calendar, config);
  console.log(`Found ${existingSyncedEvents.length} existing synced events`);
  
  // Process events
  const processedUIDs = new Set();
  let addedCount = 0;
  let updatedCount = 0;
  
  // Create or update events from ICS
  icsEvents.forEach(icsEvent => {
    if (!icsEvent.uid) return;
    
    processedUIDs.add(icsEvent.uid);
    const existingEvent = findEventByUID(existingSyncedEvents, icsEvent.uid);
    
    if (existingEvent) {
      // Update existing event if time changed
      if (updateEventIfChanged(existingEvent, icsEvent, config)) {
        updatedCount++;
      }
    } else {
      // Create new event
      if (createGoogleEvent(calendar, icsEvent, config)) {
        addedCount++;
      }
    }
  });
  
  // Delete events that no longer exist in ICS
  let deletedCount = 0;
  existingSyncedEvents.forEach(event => {
    const uid = extractUIDFromDescription(event.getDescription());
    if (uid && !processedUIDs.has(uid)) {
      event.deleteEvent();
      deletedCount++;
      console.log('Deleted event:', event.getTitle());
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
 * Only events with a start time within the sync window (30 days ago to
 * 1 year ahead) are returned. Events without a UID are skipped.
 *
 * @param {string} icsData - Raw ICS calendar text.
 * @returns {Object[]} Array of parsed event objects.
 */
function parseICSEvents(icsData) {
  const events = [];
  const lines = icsData.split('\n');
  let currentEvent = null;
  const today = new Date();
  today.setDate(today.getDate() - 30)
  today.setHours(0, 0, 0, 0);
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    if (line === 'BEGIN:VEVENT') {
      currentEvent = {};
    } else if (line === 'END:VEVENT' && currentEvent) {
      if (currentEvent.startTime && currentEvent.startTime >= today && currentEvent.uid) {
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
 * RRULE lines are intentionally ignored — the Office 365 feed expands
 * recurring events into individual instances.
 *
 * @param {string} line  - A single ICS content line.
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
    case 'DTSTART':
      event.startTime = parseICSDateTime(value, extractTimezone(propertyParams));
      event.isAllDay = !value.includes('T');
      break;
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
    // Ignore RRULE - let ICS feed handle recurrence
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
 * Fetches all previously synced events from the target Google Calendar
 * within the sync window. Synced events are identified by:
 *   - Title containing EVENT_TITLE_PREFIX
 *   - Description containing a [UID:...] marker written by this script
 *
 * @param {Calendar} calendar - GoogleAppsScript Calendar object.
 * @param {Object} config     - Config entry.
 * @returns {CalendarEvent[]} Array of matching Google Calendar events.
 */
function getExistingSyncedEvents(calendar, config) {
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - 30);
  startDate.setHours(0, 0, 0, 0);

  const futureDate = new Date();
  futureDate.setFullYear(futureDate.getFullYear() + 1);
  
  const allEvents = calendar.getEvents(startDate, futureDate);
  return allEvents.filter(event => 
    event.getTitle().includes(config.EVENT_TITLE_PREFIX) && 
    event.getDescription() && 
    event.getDescription().includes('[UID:')
  );
}

/**
 * Finds a previously synced Google Calendar event whose description contains
 * the given Microsoft UID marker.
 *
 * @param {CalendarEvent[]} existingEvents
 * @param {string} uid
 * @returns {CalendarEvent|undefined}
 */
function findEventByUID(existingEvents, uid) {
  return existingEvents.find(event => {
    const eventUID = extractUIDFromDescription(event.getDescription());
    return eventUID === uid;
  });
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
    const description = `Synced from Microsoft Calendar\n[UID:${icsEvent.uid}]`;
    
    // Use different titles for all-day vs timed events
    const title = icsEvent.isAllDay ? `${config.EVENT_TITLE_PREFIX} - FREE` : `${config.EVENT_TITLE_PREFIX} meeting`;
    
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
 * Updates the start/end time of an existing Google Calendar event if it
 * has changed in the ICS feed. No-op if times are identical.
 *
 * @param {CalendarEvent} googleEvent - Existing synced event.
 * @param {Object} icsEvent           - Updated ICS event.
 * @param {Object} config             - Config entry (used for logging).
 * @returns {boolean} True if the event was updated, false otherwise.
 */
function updateEventIfChanged(googleEvent, icsEvent, config) {
  try {
    let hasChanges = false;
    
    // Check if start time changed
    if (googleEvent.getStartTime().getTime() !== icsEvent.startTime.getTime()) {
      googleEvent.setTime(icsEvent.startTime, icsEvent.endTime);
      hasChanges = true;
    }
    
    // Check if end time changed
    if (googleEvent.getEndTime().getTime() !== icsEvent.endTime.getTime()) {
      googleEvent.setTime(icsEvent.startTime, icsEvent.endTime);
      hasChanges = true;
    }
    
    if (hasChanges) {
      console.log('Updated event time:', config.EVENT_TITLE_PREFIX, 'to', icsEvent.startTime);
      return true;
    }
    
    return false;
    
  } catch (error) {
    console.error('Error updating event:', error);
    return false;
  }
}