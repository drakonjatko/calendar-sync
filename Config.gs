/**
 * Sync configuration.
 *
 * Add one entry per Office 365 calendar you want to mirror into Google Calendar.
 * Each entry is synced independently on every run.
 *
 * How to get your ICS_URL:
 *   Outlook Web → Settings → Calendar → Shared calendars → Publish a calendar
 *   Set visibility to "Can view when I'm busy", click Publish, copy the ICS link.
 *
 * How to get your GOOGLE_CALENDAR_ID:
 *   Google Calendar → calendar settings → Integrate calendar → Calendar ID
 *   Looks like "abc123...@group.calendar.google.com" for secondary calendars.
 */
/**
 * Logging verbosity flags.
 *
 * LOG_SKIPPED_EVENTS — log every event that is filtered out during ICS parsing
 *   (no UID, no start time, too old, all-day, cancelled). Useful when events
 *   are unexpectedly missing from Google Calendar.
 *
 * LOG_PARSED_EVENTS — log every VEVENT accepted by the parser with its
 *   trackingId, start time, status, and recurrence info. Useful for verifying
 *   the ICS feed is being read correctly.
 */
const LOG_SKIPPED_EVENTS = false;
const LOG_PARSED_EVENTS  = false;

const CONFIG = [
  {
    // Office 365 ICS export link (get from Outlook Web → calendar sharing settings)
    ICS_URL: 'https://outlook.office365.com/owa/calendar/.../calendar.ics',

    // ID of the target Google Calendar
    GOOGLE_CALENDAR_ID: '...@group.calendar.google.com',

    // Prefix used for all synced event titles.
    // Timed events → "{PREFIX} meeting", all-day events → "{PREFIX} - FREE"
    EVENT_TITLE_PREFIX: 'Work',

    // Minutes before the event to show a popup reminder (0 to disable)
    NOTIFICATION_MINUTES: 5
  },

  // Add more entries here to sync additional calendars, for example:
  // {
  //   ICS_URL: 'https://outlook.office365.com/owa/calendar/.../calendar.ics',
  //   GOOGLE_CALENDAR_ID: '...@group.calendar.google.com',
  //   EVENT_TITLE_PREFIX: 'Client X',
  //   NOTIFICATION_MINUTES: 10
  // },
];