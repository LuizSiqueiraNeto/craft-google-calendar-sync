# Craft ↔ Google Calendar with Apps Script

A personal bridge between Craft tasks and a dedicated Google Calendar. The complete script is in [Code.gs](Code.gs). This public beta version of V42 requires your own configuration and contains no author's Craft connection URL or remote token.

> **Beta software:** Start with a test Craft space and calendar. Synchronization can edit tasks and reminders, create or move events, and delete events linked to tasks removed from Craft. Back up important data before using it with your main space and calendar.

## License and responsibility

This project is released under the [MIT License](LICENSE). Anyone may copy, modify, use, distribute, or sell it, including for commercial purposes. The copyright notice and license text must remain in copies or substantial portions of the software. The software is provided **as is**, without warranty; the authors disclaim liability for data loss, damage, or other claims arising from its use, as stated in the license. This disclaimer does not guarantee the same legal effect in every jurisdiction.

The beta label describes its testing status; it does not restrict any permission granted by the MIT License.

## How it works

Craft holds the task; Google Calendar displays its time block. The script reads tasks and reminders through the [Craft API](https://connect.craft.do/api-docs/space), reads events through Apps Script `CalendarApp`, and compares both sides with last sync markers in event descriptions. The link uses the task block ID, not its title.

| In Craft | In Google Calendar |
| --- | --- |
| Scheduled date without a time | All-day event |
| Scheduled date and timed reminder | Event on that date at that time |
| `#min15`, `#min30`, `#min45`, or `#min60` in task text | Duration of 15, 30, 45, or 60 minutes |
| Timed task without a duration tag | Defaults to 15 minutes and adds `#min15` to the task |
| Open, completed, or canceled task | Event title starts with ❇️, ✅, or ❎ |
| Recurring task | Next occurrence keeps the same task ID; a completed or canceled occurrence may remain as calendar history |

The script searches 14 days back and 14 days ahead (`SYNC_DAYS_AHEAD`). Events outside this window are not included in normal searches. It uses the calendar's time zone; keep that consistent with Craft's intended time zone.

### Links, conflicts, and deletion

- Event descriptions contain `[CRAFT_BLOCK_ID:...]` and `LAST_SYNC_*` markers. Keep these markers intact.
- Moving or resizing a timed event can update the date, reminder, and duration tag in Craft. Changing a Craft duration tag adjusts the event.
- Moving a one-off timed task to another day in Craft moves its linked event and timed reminder together, keeping the reminder's local clock time. This also repairs an older task whose scheduled date moved while its reminder stayed on the previous event date.
- Moving an all-day event changes the task date after an initial run records `LAST_SYNC_DATE`.
- Completing or canceling a recurring occurrence can preserve its old event as history and create the next one. Completing it in Craft is the more predictable approach.
- When both sides change before a run, conflict resolution depends on the field. Avoid editing the same item simultaneously.
- Deleting a Craft task can remove its linked pending event. Deleting only the calendar event may cause the script to recreate it while the Craft task exists.

Craft distinguishes a scheduled task date from a **timed reminder**. This script reads both. The reminders API is experimental and subject to the conditions of your Craft connection; see the [Craft API documentation](https://connect.craft.do/api-docs/space).

Craft may reject setting a reminder to a time that has already passed. In that case the sync reports the failed task in Apps Script **Executions**; choose a future time in Craft and run it again.

## Setup

### 1. Create a Craft connection

In Craft, open **Imagine** and create an API connection with access to the documents and tasks you want to sync. Use a connection covering the whole space because this script calls `/tasks?scope=all`, `/connection`, and `/reminders`. Copy the connection's base URL, such as `https://connect.craft.do/links/YOUR_ID/api/v1`. Treat the full URL as a **credential**: anyone who has it can use that connection's permissions.

Check the connection page for reminders endpoint availability. Craft documents additional conditions for this experimental resource.

### 2. Create an Apps Script project

1. Create a standalone project in [Apps Script](https://script.google.com/).
2. Replace the entire contents of `Code.gs` in the project with this repository's `Code.gs`. Do not append it to an older copy: duplicate declarations cause syntax errors.
3. Under **Project Settings → Script Properties**, set `CRAFT_API_BASE` to your Craft connection URL. The calendar name defaults to `Task-Craft` through `DEFAULT_CALENDAR_NAME` near the top of `Code.gs`; you can edit that constant. Alternatively, set the `CALENDAR_NAME` Script Property to your preferred name. The Script Property takes precedence. Use a unique calendar name.
4. Set the project's time zone and check the calendar's time zone. Save, then run `syncCraftWithGoogleCalendar` manually to grant the requested permissions. If the named calendar does not exist, this first run creates it.
5. Inspect the dedicated calendar and a test task before enabling automatic runs.

[Script Properties](https://developers.google.com/apps-script/guides/properties) are project settings. Do not copy their values into the repository, screenshots, or public logs.

### 3. Install the Google Calendar change trigger

After the first successful sync, run `installCalendarChangeTrigger` **once** from the Apps Script editor and approve its permissions. It finds the dedicated calendar by its configured name, reads its actual calendar ID, and installs an `onEventUpdated` trigger for `syncCraftWithGoogleCalendar`. Running the installer again does not add a duplicate for the same calendar and function. If the calendar name matches zero or multiple calendars, it stops with an error; choose a unique name. If you later switch to another calendar, remove the old trigger in Apps Script's **Triggers** panel and run the installer again.

You can instead create the same installable trigger in the Apps Script **Triggers** panel: function `syncCraftWithGoogleCalendar`, event source **Calendar**, event type **Event updated**, dedicated calendar. Do not use both methods to create separate triggers for the same calendar.

The [Google Calendar change trigger](https://developers.google.com/apps-script/guides/triggers/events) reports that the calendar changed but does not identify the event or the kind of change. This script scans its sync window each time; it does not use the Calendar advanced service's incremental `nextSyncToken` mechanism. Changes made solely in Craft do not fire this trigger. For those, run the function manually, add a time trigger, or use the optional URL below.

### 4. Optional sync URL

Publishing this source code on GitHub does **not** create a web app for each user. To call a personal copy from Apple Shortcuts, deploy that copy as an [Apps Script web app](https://developers.google.com/apps-script/guides/web). The script already contains `doGet(e)`, which accepts a GET request, checks the token, and starts the sync.

1. Complete the setup above and run `syncCraftWithGoogleCalendar` manually at least once to grant permissions. Then run `prepareRemoteAccess` once in the editor. It creates `REMOTE_SYNC_TOKEN` in **Project Settings → Script Properties**. Copy that value privately; the deployment ID is **not** the token.
2. Select **Deploy → New deployment**. Choose **Web app** as the deployment type, **Execute as: Me**, and **Who has access: Anyone** for a request that does not rely on a signed-in Google session. Authorize and deploy. This publishes the current saved code as a versioned web app.
3. Copy the **Web app URL** ending in `/exec`. Append `?token=` and your own `REMOTE_SYNC_TOKEN` value, with no spaces. For example, replace both placeholders in:

   ```text
   https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?token=YOUR_REMOTE_SYNC_TOKEN
   ```

4. In Apple Shortcuts on Mac, iPhone, or iPad, add a **URL** action containing that complete URL, followed by **Get Contents of URL** with method **GET**. Run the shortcut once and inspect the returned text. A `Show Result` action is useful while testing. You can then use the shortcut in your preferred automation.
5. Keep the complete URL and token private. **Anyone with both can start a sync under your Google account.** Query tokens can appear in browser history, logs, or screenshots. If exposed, replace `REMOTE_SYNC_TOKEN` in Script Properties with a new random value of at least 32 characters and update every shortcut using it.

The response is plain text: `OK` means the run finished; `PENDING` means some Craft reminders still need another attempt; `RECENT` means a remote run was skipped because another one ran within 30 seconds; `BUSY` means a sync already holds the lock. `NOT_CONFIGURED`, `ACCESS_DENIED`, or `ERROR` indicate a setup or execution problem. Inspect **Executions** in Apps Script for details. Check the response text, since this script also returns error messages as text.

The Calendar trigger uses saved code, while the `/exec` URL uses a published version. To update **the same URL** after a code change, create a new version under **Deploy → Manage deployments → Edit → New version → Deploy**. See the [deployment guide](https://developers.google.com/apps-script/concepts/deployments).

## Quick check

1. Create a Craft test task for tomorrow with a future reminder and `#min15`.
2. Run the sync. Check its date, time, duration, and linked block ID in the event description.
3. Move the event to another future time and run the sync. Check the Craft date and reminder.
4. Move the one-off task to another future day in Craft while keeping its time. Run the sync and check that the same Calendar event and the Craft reminder both moved to that day. Run it again to confirm there is no duplicate.
5. Change the Craft tag to `#min30`, run the sync, and check the calendar event duration.
6. After installing the calendar trigger, move the test event once more and inspect **Executions** in Apps Script. The trigger may take some time to run; confirm the final state in Craft.
7. For all-day events created by earlier versions, run V42 once to record their reference date before testing a calendar move.

## Before making the repository public

- Publish only this credential-free copy. Do **not** include older script versions, Git history containing them, real Craft connection URLs, remote tokens, `/exec?token=...` URLs, personal IDs, or screenshots with private data.
- If the Craft connection URL or remote token is exposed, revoke or recreate the connection and replace the token in Script Properties and automations.
- Include `LICENSE` alongside `Code.gs` and `README.md` when publishing.
- Check [Apps Script quotas](https://developers.google.com/apps-script/guides/services/quotas) and [Craft API limits](https://connect.craft.do/api-docs/space) for your task volume.

This is a personal project and is not affiliated with Craft or Google.
