# calendar-sync

Sync your Microsoft Outlook (Office 365) calendar into Google Calendar — **busy times only, no event details**.

Use this when you work across two calendar systems and want colleagues or collaborators to see when you're busy in Google Calendar without exposing event titles, attendees, or other private details.

## How it works

The script runs on a schedule inside Google Apps Script. It fetches your Office 365 calendar as an ICS feed, compares the busy slots against your Google Calendar, and creates, updates, or deletes events to keep them in sync.

Only the busy/free status is transferred. All synced events appear with a generic title you define (e.g. `Work meeting`).

## Prerequisites

- A Google account with access to [Google Apps Script](https://script.google.com)
- An Office 365 / Outlook account with calendar sharing enabled (ICS export link)
- A Google Calendar to sync into (can be a secondary calendar you create for this purpose)

## Setup

### 1. Get your Office 365 ICS URL

1. Go to [Outlook Web App](https://outlook.office365.com)
2. Open **Settings → View all Outlook settings → Calendar → Shared calendars**
3. Under **Publish a calendar**, select your calendar and set permissions to **Can view when I'm busy**
4. Click **Publish** and copy the **ICS** link

### 2. Get your Google Calendar ID

1. Open [Google Calendar](https://calendar.google.com)
2. Next to the target calendar, click the three-dot menu → **Settings and sharing**
3. Scroll down to **Integrate calendar** and copy the **Calendar ID**
   - It looks like `abc123...@group.calendar.google.com` for secondary calendars
   - Or your email address for your primary calendar

### 3. Deploy the script

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Create three script files and paste the contents of each `.gs` file:
   - `Code.gs`
   - `Config.gs`
   - `Timezones.gs`
3. In `Config.gs`, replace the placeholder values with your actual ICS URL, Google Calendar ID, and a title prefix for the synced events
4. Enable the **Advanced Calendar Service**:
   - Click **Services** (+ icon in the left sidebar)
   - Find **Google Calendar API** and click **Add**
5. Click **Run → syncCalendar** once to authorize the required permissions (Google will prompt you to grant calendar access)
6. Set up a time-based trigger:
   - Click **Triggers** (clock icon in the left sidebar)
   - Add trigger → function `syncCalendar` → time-driven → every 15 or 30 minutes

### 4. Configure multiple calendars (optional)

Add more entries to the `CONFIG` array in `Config.gs` to sync additional Office 365 calendars into different Google Calendars. Each entry is synced independently on every run.

## Configuration reference

| Field | Description |
|---|---|
| `ICS_URL` | The ICS export link from Office 365 |
| `GOOGLE_CALENDAR_ID` | ID of the target Google Calendar |
| `EVENT_TITLE_PREFIX` | Prefix used for all synced event titles (e.g. `Work`) |
| `NOTIFICATION_MINUTES` | Minutes before event to show a popup reminder |

## Behavior

- **Sync window**: 30 days in the past to 1 year in the future. Events outside this range are ignored.
- **Event titles**: All synced events are titled `{PREFIX} meeting`. All-day events from Outlook are skipped and never appear in Google Calendar.
- **Updates**: If a synced event's time or tentative/confirmed status changes in Outlook, it will be updated on the next sync.
- **Deletions**: If an event is removed from the Outlook feed, it will be deleted from Google Calendar on the next sync.
- **Tentative events**: Events marked as tentative in Outlook are synced as transparent (free) in Google Calendar. If an event changes between tentative and confirmed, the transparency is updated on the next sync.
- **Recurring events**: Office 365 sends recurring meetings as series with a recurrence rule (RRULE). These are created as real recurring events in Google Calendar via the Advanced Calendar API, so they show as a series and respect DST correctly. If an individual occurrence is deleted in Outlook, that specific instance is explicitly removed from the Google Calendar series. If an occurrence is rescheduled to a different time, the original slot is removed and a standalone event is created at the new time.
- **Duplicate meeting detection**: If the same meeting is invited to multiple source accounts (e.g. you have two accounts in a shared meeting), it will only appear once in Google Calendar. The script detects events with the same UID across all config entries syncing into the same calendar, skips re-creation, and cleans up any duplicates that already exist (keeping the older copy).

## Troubleshooting

**Events are not appearing in Google Calendar**
- Check the Apps Script execution log (View → Logs) for HTTP errors when fetching the ICS URL
- Make sure the ICS URL is for the correct calendar and has not expired (Office 365 publishing links can be revoked)
- Verify the Google Calendar ID is correct and that the script has been authorized

**Old events are not being deleted**
- The script only manages events it created (identified by the `[UID:...]` marker in the event description). Manually created events are never touched.

**Timezone issues**
- The script maps Microsoft timezone names to IANA names. If you see events shifted by an hour, check `Timezones.gs` and verify the mapping for your timezone is correct.

## Known limitations

- No bidirectional sync — changes made in Google Calendar are overwritten on the next sync.
- All synced events get the same generic title; individual event names are intentionally not transferred.
