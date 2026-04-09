# calendar-sync

Google Apps Script project that syncs busy times from Microsoft Outlook (Office 365) ICS feeds into Google Calendar. Syncs **busy status only** — no event details are transferred.

## Architecture

```
Office 365 ICS Feed → fetchICSData() → parseICSEvents() → diff with Google Calendar → create/update/delete
```

- **Entry point**: `syncCalendar()` in [Code.gs](../Code.gs) — iterates over `CONFIG` and calls `syncSingleCalendar()` per source
- **Config**: [Config.gs](../Config.gs) — `CONFIG` array of `{ ICS_URL, GOOGLE_CALENDAR_ID, EVENT_TITLE_PREFIX, NOTIFICATION_MINUTES }` objects
- **Timezone map**: [Timezones.gs](../Timezones.gs) — 100 Microsoft → IANA timezone name mappings

## Deployment

This is a **Google Apps Script** project — there is no local build step.

1. Copy all `.gs` files into a GAS project at [script.google.com](https://script.google.com)
2. Fill in `CONFIG` entries in `Config.gs` with real ICS URLs and Google Calendar IDs
3. Enable the Advanced Calendar Service: Apps Script editor → Services → Google Calendar API
4. Set a time-based trigger on `syncCalendar()` (e.g., every 15 minutes)
5. Required OAuth scope: `https://www.googleapis.com/auth/calendar`

## Conventions

- **Tracking key**: Each synced Google Calendar event stores a `trackingId` in its description as `[UID:value]`. For non-recurring events `trackingId = uid|dtStartRaw`; for modified series occurrences it is `uid|recurrenceId`; for series masters it is `uid` alone. Do not alter this format or the regex in `extractUIDFromDescription()`.
- **Recurring series**: ICS events with an `RRULE` are created as real Google Calendar recurring events via the Advanced Calendar Service (`Calendar.Events.insert/patch/remove`). The Advanced Calendar Service must be enabled in the GAS project (Services → Google Calendar API). Individual standalone events are created via `CalendarApp.createEvent()`, but reading, updating, and deleting them uses the Advanced Calendar API (`Calendar.Events.list/patch/remove`).
- **Event titles**: Timed events get `{PREFIX} meeting`. Title prefix comes from `EVENT_TITLE_PREFIX` in config.
- **Sync window**: Single events and exception occurrences: 30 days ago to 1 year ahead. Series masters: always included (their DTSTART can be years old).
- **Skipped events**: All-day events are skipped entirely. Events with `STATUS:CANCELLED` are also skipped, but if a cancelled event also has a `RECURRENCE-ID`, `deleteCancelledOccurrence()` is called first to explicitly remove that specific occurrence from the Google Calendar series (otherwise it stays visible). Events with a `RECURRENCE-ID` but **without** `STATUS:CANCELLED` (modified occurrences rescheduled to a different time) also call `deleteCancelledOccurrence()` to remove the original series slot before a standalone event is created at the new time. Neither all-day nor cancelled trackingIds are added to `processedTrackingIds`, so any previously synced standalone version is cleaned up in the deletion pass.
- **ICS parsing**: Custom line-by-line parser (not a library). Handles RFC 5545 line folding: continuation lines (CRLF + space/tab) are rejoined before parsing so property values (e.g. long RRULEs) are never truncated.

## Key Pitfalls

- **Timezone handling**: `buildDate()` uses `Utilities.formatDate()` + ISO string reconstruction — a workaround for GAS date API limitations. DST edge cases may exist; do not simplify without testing. Recurring events use `googleDateTime()` which reconstructs local time from a UTC Date + IANA timezone to keep occurrences at the correct wall-clock time across DST transitions.
- **Advanced Calendar Service**: Required for recurring series and for standalone event management. `getExistingSeriesMasters()` uses `singleEvents: false` with a 5-year lookback so long-running series are found. `getExistingSyncedEvents()` also uses the Advanced API with `singleEvents: true` + `!event.recurringEventId` filter — this is critical to prevent `CalendarApp.getEvents()` from returning series occurrences (which inherit the master's `[UID:...]` description), which would otherwise cause phantom matches in the cleanup pass. Deleting a series via `Calendar.Events.remove()` also removes all instances — the `deletedSeriesIds` set prevents double-delete errors on individual instances. Both functions omit the `q` / summary prefix filter so all synced events are returned regardless of which config created them.
- **Cross-config deduplication**: Both functions fetch without a prefix filter. `syncSingleCalendar` splits the results into `allSeriesMasters`/`allSyncedEvents` (for existence checks) and `ownedSeriesMasters`/`ownedSyncedEvents` (for cleanup — filtered by `EVENT_TITLE_PREFIX`). Before creating, the code checks `allSeriesMasters`/`allSyncedEvents` — if a matching UID already exists under a different prefix, creation is skipped. Before updating, it checks ownership (prefix match) — if the found event belongs to another config it is left untouched. In the cleanup pass, if an owned event has a cross-config duplicate (same UID, different Google Calendar event ID), the newer copy (by `event.created` timestamp) is deleted; the older copy is kept. This is the deterministic tie-breaker.
- **Cancelled occurrences**: `deleteCancelledOccurrence(calendarId, seriesId, exceptionIcsEvent)` is called for ANY VEVENT with a `RECURRENCE-ID` (both cancelled and modified), to remove the original series slot. It uses `Calendar.Events.instances` with a ±24 h window around the `RECURRENCE-ID` time, matches by `originalStartTime` within 60 seconds, and removes the matching instance. The timezone preference order is: `recurrenceTzid` (from the RECURRENCE-ID property’s own TZID parameter) → `timezone` (from DTSTART) → null, to ensure floating RECURRENCE-ID values are parsed in the correct timezone rather than GAS server local time.
- **UID null collisions**: If `extractUIDFromDescription()` returns `null` (malformed description), `findEventByUID()` could incorrectly match events. Guard any code that relies on UID matching.
- **Quota**: Multiple CONFIG entries run sequentially. High-volume calendars risk hitting Google Calendar API quotas.
- **Google Apps Script APIs only**: Uses `CalendarApp` (create standalone events), `Calendar` Advanced Service (list/patch/remove all events, manage series), `UrlFetchApp`, `Utilities` — no Node.js, no npm.
