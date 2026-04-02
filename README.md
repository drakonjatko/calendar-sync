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
   - `code.gs`
   - `Config.gs`
   - `Timezones.gs`
3. In `Config.gs`, replace the placeholder values with your actual ICS URL, Google Calendar ID, and a title prefix for the synced events
4. Click **Run → syncCalendar** once to authorize the required permissions (Google will prompt you to grant calendar access)
5. Set up a time-based trigger:
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
- **Event titles**: Timed events are created as `{PREFIX} meeting`; all-day events as `{PREFIX} - FREE`.
- **Updates**: If a synced event's time changes in Outlook, it will be updated on the next sync.
- **Deletions**: If an event is removed from the Outlook feed, it will be deleted from Google Calendar on the next sync.
- **Tentative events**: Events marked as tentative in Outlook are created as transparent (free) in Google Calendar.
- **Recurring events**: The ICS feed from Office 365 expands recurring events into individual instances; they are synced as separate events.

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

- RFC 5545 line folding in ICS data is not supported. This is not an issue for Office 365 feeds in practice.
- No bidirectional sync — changes made in Google Calendar are overwritten on the next sync.
- All synced events get the same generic title; individual event names are intentionally not transferred.
