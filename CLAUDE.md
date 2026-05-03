# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Google Apps Script project that syncs busy times from Office 365 ICS feeds into Google Calendar. It transfers busy/free status only — no event details.

## Development workflow

There is no local build step. All code runs inside Google Apps Script.

- **Deploy**: Copy all `.gs` files into a GAS project at script.google.com. Enable the Advanced Calendar Service (Services → Google Calendar API). Set a time-based trigger on `syncCalendar()`.
- **Test a change**: Paste updated code into the GAS editor and run `syncCalendar()` manually, then check the execution log (View → Logs).
- **No npm, no Node.js** — only GAS APIs: `Calendar` Advanced Service, `UrlFetchApp`, `Utilities`. `CalendarApp` is not used.

## Architecture

```
Office 365 ICS Feed
  → fetchICSData()           # UrlFetchApp.fetch()
  → parseICSEvents()         # custom RFC 5545 line-by-line parser
  → syncSingleCalendar()     # diff against Google Calendar, create/update/delete
```

`syncCalendar()` iterates over `CONFIG` (in Config.gs) and calls `syncSingleCalendar(config)` for each source. Timezones.gs maps Microsoft timezone names to IANA names.

## Critical conventions

**Tracking ID format** — every synced event stores its key in the Google Calendar event description as `[UID:value]`. Three formats:
- Series masters: `uid`
- Standalone events: `uid|dtStartRaw`
- Modified occurrences: `uid|recurrenceId`

Do not change this format or the regex in `extractUIDFromDescription()`.

**All event operations use the Advanced API** — both standalone and recurring events are created, read, updated, and deleted exclusively via `Calendar.Events.insert/patch/list/remove`. `CalendarApp` is not used.

**Event title** — all synced events get the title `{EVENT_TITLE_PREFIX} meeting`. The prefix comes from config.

**Sync window** — standalone events and exceptions: 30 days past to 1 year ahead. Series masters: always included (DTSTART can be years old).

**Skipped events** — all-day events are skipped entirely. `STATUS:CANCELLED` events are also skipped, but if they carry a `RECURRENCE-ID`, `deleteCancelledOccurrence()` is still called to remove that slot from the series. Modified occurrences (RECURRENCE-ID without CANCELLED) also call `deleteCancelledOccurrence()` before a standalone event is created at the new time.

## Key pitfalls

**Pagination is mandatory** — `getExistingSeriesMasters()` and `getExistingSyncedEvents()` both paginate via `nextPageToken` until all pages are retrieved. Without this, series fall off the first page and a new duplicate is created on every trigger run.

**Timezone handling** — `buildDate()` uses `Utilities.formatDate()` + ISO string reconstruction, a workaround for GAS date API limitations. `googleDateTime()` reconstructs wall-clock time from UTC + IANA timezone to survive DST transitions. Do not simplify either without testing DST edge cases.

**Instance filter** — `getExistingSyncedEvents()` uses `singleEvents: true` with `!event.recurringEventId` to exclude series occurrences. Without this filter, recurring instances inherit the master's `[UID:...]` description and cause phantom matches in cleanup.

**Series mastering** — `getExistingSeriesMasters()` uses `singleEvents: false` with a 5-year lookback and a 1-year-ahead cap (`timeMax`). Deleting a series via `Calendar.Events.remove()` removes all its instances server-side; the cleanup pass for standalone events is safe because `getExistingSyncedEvents` already excludes series instances via `!event.recurringEventId`.

**Lookup Maps** — `syncSingleCalendar` calls `buildTrackingIdMap()` once on both `allSeriesMasters` and `allSyncedEvents` to build `Map<trackingId, event[]>` indexes before any loops. All existence checks and cross-config duplicate lookups use `.get()` — O(1) instead of O(n) linear scans.

**Cross-config deduplication** — both fetch functions omit the prefix filter so events from all configs are returned. `syncSingleCalendar` maintains two sets: `all*` (for existence checks across configs) and `owned*` (for cleanup, filtered by prefix). If a UID already exists under a different prefix, creation is skipped. In cleanup, the newer duplicate (by `event.created` timestamp) is deleted.

**Transparency** — `isTentative(icsEvent)` checks `busyStatus === 'TENTATIVE'` or `status === 'TENTATIVE'`. All four create/update functions (`createGoogleEvent`, `createRecurringGoogleEvent`, `updateEventIfChanged`, `updateRecurringSeriesIfChanged`) include `transparency` in every API call, so a flip between tentative and confirmed is always applied on the next sync.

**UID null guard** — `extractUIDFromDescription()` can return `null` on malformed descriptions. `buildTrackingIdMap` skips null keys automatically; any code calling `extractUIDFromDescription()` directly must guard against null before using the result.

**Quota** — CONFIG entries run sequentially; high-volume calendars can hit Google Calendar API quotas.
