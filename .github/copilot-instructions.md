# calendar-sync

Google Apps Script project that syncs busy times from Microsoft Outlook (Office 365) ICS feeds into Google Calendar. Syncs **busy status only** — no event details are transferred.

## Architecture

```
Office 365 ICS Feed → fetchICSData() → parseICSEvents() → diff with Google Calendar → create/update/delete
```

- **Entry point**: `syncCalendar()` in [code.gs](../code.gs) — iterates over `CONFIG` and calls `syncSingleCalendar()` per source
- **Config**: [Config.gs](../Config.gs) — `CONFIG` array of `{ ICS_URL, GOOGLE_CALENDAR_ID, EVENT_TITLE_PREFIX, NOTIFICATION_MINUTES }` objects
- **Timezone map**: [Timezones.gs](../Timezones.gs) — 156 Microsoft → IANA timezone name mappings

## Deployment

This is a **Google Apps Script** project — there is no local build step.

1. Copy all `.gs` files into a GAS project at [script.google.com](https://script.google.com)
2. Fill in `CONFIG` entries in `Config.gs` with real ICS URLs and Google Calendar IDs
3. Set a time-based trigger on `syncCalendar()` (e.g., every 15 minutes)
4. Required OAuth scope: `https://www.googleapis.com/auth/calendar`

## Conventions

- **UID tracking**: Microsoft event UIDs are embedded in Google Calendar event descriptions as `[UID:value]` — this is the sole mechanism for detecting existing synced events. Do not alter this format.
- **Event titles**: All-day events get title `{PREFIX} - FREE`; timed events get `{PREFIX} meeting`. Title prefix comes from `EVENT_TITLE_PREFIX` in config.
- **Sync window**: Events from 30 days ago to 1 year ahead. Events outside this window are ignored.
- **Stateless**: No database or script properties. All state is derived from Google Calendar event descriptions.
- **ICS parsing**: Custom line-by-line parser (not a library). Does not handle RFC 5545 line folding. Assumes well-formed ICS from Office 365.

## Key Pitfalls

- **Timezone handling**: `buildDate()` uses `Utilities.formatDate()` + ISO string reconstruction — a workaround for GAS date API limitations. DST edge cases may exist; do not simplify without testing.
- **No recurrence support**: `RRULE` lines are intentionally ignored; the ICS feed is expected to expand recurrences itself.
- **UID null collisions**: If `extractUIDFromDescription()` returns `null` (malformed description), `findEventByUID()` could incorrectly match events. Guard any code that relies on UID matching.
- **Quota**: Multiple CONFIG entries run sequentially. High-volume calendars risk hitting Google Calendar API quotas.
- **Google Apps Script APIs only**: Uses `CalendarApp`, `UrlFetchApp`, `Utilities` — no Node.js, no npm.
