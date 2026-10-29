/**
 * TWO-WAY SYNC: CRAFT ⇄ GOOGLE CALENDAR (V42 BETA)
 * Public example: configure your Craft connection in Script Properties.
 * Released under the MIT License. See LICENSE for the full disclaimer.
 */

// The Craft connection URL is a credential. Never put it in source code.
const CRAFT_API_BASE = (PropertiesService.getScriptProperties()
  .getProperty('CRAFT_API_BASE') || '').replace(/\/$/, '');
// Set your default calendar name here, or override it with the CALENDAR_NAME
// Script Property. This is a name, not the calendar's technical ID.
const DEFAULT_CALENDAR_NAME = 'Task-Craft';
const CALENDAR_NAME = PropertiesService.getScriptProperties()
  .getProperty('CALENDAR_NAME') || DEFAULT_CALENDAR_NAME;
const DEFAULT_DURATION_MIN = 15;
const ALLOWED_DURATIONS = [15, 30, 45, 60];
const SYNC_DAYS_AHEAD = 14;
const REMOTE_MIN_INTERVAL_MS = 30000;
const MAX_CALENDAR_DELETIONS_PER_RUN = 5;

function syncCraftWithGoogleCalendar() {
  return runSyncWithLock_(false);
}

// Run once after configuring CALENDAR_NAME. Re-running does not add a duplicate.
function installCalendarChangeTrigger() {
  const calendars = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (calendars.length !== 1) {
    throw new Error('Expected exactly one calendar named "' + CALENDAR_NAME +
      '"; found ' + calendars.length + '. Set CALENDAR_NAME to a unique name.');
  }
  const calendarId = calendars[0].getId();
  const exists = ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === 'syncCraftWithGoogleCalendar' &&
    trigger.getEventType() === ScriptApp.EventType.ON_EVENT_UPDATED &&
    trigger.getTriggerSourceId() === calendarId);
  if (exists) {
    Logger.log('Calendar change trigger already installed for ' + CALENDAR_NAME + '.');
    return;
  }
  ScriptApp.newTrigger('syncCraftWithGoogleCalendar')
    .forUserCalendar(calendarId).onEventUpdated().create();
  Logger.log('Calendar change trigger installed for ' + CALENDAR_NAME + '.');
}

// Run once in the editor. The token stays in Script Properties.
function prepareRemoteAccess() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('REMOTE_SYNC_TOKEN')) {
    Logger.log('A remote token already exists in Script Properties.');
    return;
  }
  const token = Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');
  props.setProperty('REMOTE_SYNC_TOKEN', token);
  Logger.log('Token created. Read REMOTE_SYNC_TOKEN in Script Properties.');
}

// Web app URL: .../exec?token=YOUR_REMOTE_SYNC_TOKEN
function doGet(e) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('REMOTE_SYNC_TOKEN');
  const received = e && e.parameter && e.parameter.token;
  if (!token || token.length < 32) return remoteText_('NOT_CONFIGURED');
  if (received !== token) return remoteText_('ACCESS_DENIED');

  try {
    return remoteText_(runSyncWithLock_(true));
  } catch (error) {
    Logger.log('Remote sync failed: ' + error.message);
    return remoteText_('ERROR; check Apps Script executions');
  }
}

function remoteText_(message) {
  return ContentService.createTextOutput(message)
    .setMimeType(ContentService.MimeType.TEXT);
}

function runSyncWithLock_(remote) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    if (remote) return 'BUSY; try again';
    throw new Error('Another sync is still running.');
  }
  try {
    const props = PropertiesService.getScriptProperties();
    if (remote) {
      const last = Number(props.getProperty('LAST_REMOTE_SYNC_MS') || 0);
      if (Date.now() - last < REMOTE_MIN_INTERVAL_MS) return 'RECENT; skipped';
    }
    const pendingReminders = syncCraftWithGoogleCalendarCore_();
    if (remote) props.setProperty('LAST_REMOTE_SYNC_MS', String(Date.now()));
    return pendingReminders
      ? 'PENDING: ' + pendingReminders + ' Craft reminder(s)'
      : 'OK';
  } finally {
    lock.releaseLock();
  }
}

function syncCraftWithGoogleCalendarCore_() {
  if (!CRAFT_API_BASE || !/^https:\/\/connect\.craft\.do\//.test(CRAFT_API_BASE)) {
    throw new Error('Set CRAFT_API_BASE in Script Properties.');
  }
  const connInfo = craftRequest('GET', '/connection');
  const spaceId = (connInfo && connInfo.space && connInfo.space.id) ? connInfo.space.id : '';
  const craftAppLinkTemplate = connInfo && connInfo.urlTemplates &&
    connInfo.urlTemplates.app;

  // "all" reduces reads; logbook preserves the previous task selection.
  const allTasksResponse = craftRequest('GET', '/tasks?scope=all');
  if (!allTasksResponse || !Array.isArray(allTasksResponse.items)) {
    throw new Error('Craft did not return a complete task list; no events will be deleted.');
  }
  const allTasks = allTasksResponse.items;
  const logbookTasks = craftRequest('GET', '/tasks?scope=logbook').items || [];
  const logbookIds = new Set(logbookTasks.filter(task => task && task.id)
    .map(task => task.id));
  const taskMap = {};
  allTasks.concat(logbookTasks).forEach(task => {
    if (task && task.id) taskMap[task.id] = task;
  });

  const reminderMap = {};
  let remindersAvailable = true;
  try {
    listReminders().forEach(rem => {
      if (rem.blockId) reminderMap[rem.blockId] = rem;
    });
  } catch (error) {
    remindersAvailable = false;
    Logger.log('Reminders unavailable: ' + error.message);
  }

  let cal = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0];
  if (!cal) {
    cal = CalendarApp.createCalendar(CALENDAR_NAME);
  }
  const timeZone = cal.getTimeZone();

  const now = new Date();
  // Include the previous occurrence so a recurring event can be found when
  // Craft advances its task to the next date.
  const startWindow = new Date(now.getTime() - SYNC_DAYS_AHEAD * 24 * 60 * 60 * 1000);
  const endWindow = new Date(now.getTime() + SYNC_DAYS_AHEAD * 24 * 60 * 60 * 1000);
  const existingEvents = cal.getEvents(startWindow, endWindow);

  const eventByBlockId = {};
  existingEvents.forEach(ev => {
    const desc = ev.getDescription() || '';
    const blockId = extractCraftMarkerId_(desc, 'CRAFT_BLOCK_ID');
    if (blockId) {
      eventByBlockId[blockId] = ev;
    } else {
      const archivedId = extractCraftMarkerId_(desc, 'CRAFT_COMPLETED_OCCURRENCE_OF') ||
        extractCraftMarkerId_(desc, 'CRAFT_CANCELED_OCCURRENCE_OF');
      if (archivedId && taskMap[archivedId]) {
        const fixed = repairCraftLink_(desc,
          craftTaskLink_(craftAppLinkTemplate, spaceId, archivedId));
        if (fixed !== desc) ev.setDescription(fixed);
      }
    }
  });

  const tasks = Object.values(taskMap).filter(task => {
    const state = task.taskInfo && task.taskInfo.state;
    return (state !== 'done' && state !== 'canceled') ||
      logbookIds.has(task.id) || !!eventByBlockId[task.id];
  });
  const failures = [];
  let pendingReminders = 0;
  tasks.forEach(task => {
    try {
    const blockId = task.id;
    const rem = reminderMap[blockId];
    
    let startTime = extractTaskStartTime(task, rem, timeZone);
    const schedDateStr = task.taskInfo && task.taskInfo.scheduleDate;

    if (!schedDateStr) return;

    const parts = schedDateStr.split('-');
    if (parts.length !== 3) return;
    const taskDate = Utilities.parseDate(schedDateStr + ' 12:00', timeZone, 'yyyy-MM-dd HH:mm');

    if (!eventByBlockId[blockId] && (taskDate < startWindow || taskDate > endWindow)) return;

    const parsed = parseTaskWithTags(task.markdown || '');
    const taskState = (task.taskInfo && task.taskInfo.state) ? task.taskInfo.state : 'todo';
    
    let prefix = '❇️ '; 
    if (taskState === 'done') {
      prefix = '✅ ';    
    } else if (taskState === 'canceled') {
      prefix = '❎ ';    
    }

    const expectedTitle = `${prefix}${parsed.cleanTitle}${
      task.location && task.location.title ? ` [${task.location.title}]` : ''
    }`;

    const existingEv = eventByBlockId[blockId];

    if (!existingEv) {
      const deepLink = craftTaskLink_(craftAppLinkTemplate, spaceId, blockId);

      if (!startTime) {
        const allDayDesc = [
          `🔗 Open in Craft: ${deepLink}`,
          `📄 Context: ${(task.location && task.location.title) || 'Daily Note / Inbox'}`,
          ``,
          `--- Sync metadata (Do not edit) ---`,
          `[CRAFT_BLOCK_ID:${blockId}]`,
          `[LAST_SYNC_STATE:${taskState}]`,
          `[LAST_SYNC_DATE:${schedDateStr}]`,
          `[IS_ALL_DAY:true]`
        ].join('\n');
        const dayStart = Utilities.parseDate(schedDateStr + ' 00:00', timeZone, 'yyyy-MM-dd HH:mm');
        cal.createAllDayEvent(expectedTitle, dayStart, { description: allDayDesc });
      } else {
        const durationMin = parsed.durationMin;
        const endTime = new Date(startTime.getTime() + durationMin * 60 * 1000);

        let timedDesc = [
          `🔗 Open in Craft: ${deepLink}`,
          `⏱️ Duration: #min${durationMin}`,
          `📄 Context: ${(task.location && task.location.title) || 'Daily Note / Inbox'}`,
          ``,
          `--- Sync metadata (Do not edit) ---`,
          `[CRAFT_BLOCK_ID:${blockId}]`,
          `[LAST_SYNC_STATE:${taskState}]`,
          rem ? `[CRAFT_REMINDER_ID:${rem.id}]` : `[CRAFT_OFFSET_MODE:true]`,
          `[LAST_SYNC_START:${startTime.toISOString()}]`,
          `[LAST_SYNC_DUR:${durationMin}]`,
          rem && rem.remindAt
            ? `[LAST_SYNC_REMINDER_AT:${new Date(rem.remindAt).toISOString()}]` : '',
          task.repeat && task.repeat.reminder &&
          typeof task.repeat.reminder.dateOffset === 'number'
            ? `[LAST_SYNC_REPEAT_OFFSET:${task.repeat.reminder.dateOffset}]` : ''
        ].join('\n');

        const createdEv = cal.createEvent(expectedTitle, startTime, endTime,
          { description: timedDesc });
        if (!hasDurationTag_(task.markdown || '')) {
          assertItem(craftRequest('PUT', '/tasks', {
            tasksToUpdate: [{ id: blockId,
              markdown: replaceOrAppendDurationTag(task.markdown || '', durationMin) }]
          }), blockId);
          verifyTaskUpdate(blockId, { durationMin: durationMin, requireDurationTag: true });
        }
        if (task.repeat) {
          const reminderResult = ensureRecurringReminder_(blockId, rem,
            startTime, remindersAvailable);
          if (!reminderResult.synced) pendingReminders++;
          timedDesc = describeRecurringReminder_(timedDesc, reminderResult, startTime);
          if (timedDesc !== createdEv.getDescription()) createdEv.setDescription(timedDesc);
        }
      }
    } else {
      let currentEvTitle = existingEv.getTitle();
      let desc = existingEv.getDescription() || '';
      const repairedDesc = repairCraftLink_(desc,
        craftTaskLink_(craftAppLinkTemplate, spaceId, blockId));
      if (repairedDesc !== desc) {
        desc = repairedDesc;
        existingEv.setDescription(desc);
      }
      const wasAllDayInGCal = existingEv.isAllDayEvent();

      // Compare each side with the last synced state. An old ✅ must not
      // complete a task again after it was unchecked directly in Craft.
      const stateMatch = desc.match(/\[LAST_SYNC_STATE:(todo|done|canceled)\]/);
      const lastState = stateMatch ? stateMatch[1] : null;
      const calendarState = currentEvTitle.startsWith('✅') ? 'done'
        : currentEvTitle.startsWith('❎') ? 'canceled' : 'todo';
      const craftStateChanged = lastState !== null && taskState !== lastState;
      const calendarStateChanged = lastState !== null && calendarState !== lastState;

      if (calendarStateChanged && !craftStateChanged) {
        assertItem(craftRequest('PUT', '/tasks', {
          tasksToUpdate: [{ id: blockId, taskInfo: { state: calendarState } }]
        }), blockId);
        verifyTaskUpdate(blockId, { state: calendarState });
        existingEv.setDescription(upsertMarker(desc, 'LAST_SYNC_STATE', calendarState));
        return;
      }

      if (!lastState || craftStateChanged) {
        // For older events without a marker, use Craft as the source of truth.
        // Craft also wins if both sides changed before this sync.
        desc = upsertMarker(desc, 'LAST_SYNC_STATE', taskState);
        existingEv.setDescription(desc);
        if (currentEvTitle !== expectedTitle) {
          existingEv.setTitle(expectedTitle);
          currentEvTitle = expectedTitle;
        }
      }

      // 2. Sync event moves and duration changes between Calendar and Craft.
      if (!wasAllDayInGCal) {
        const evStart = existingEv.getStartTime();
        const evEnd = existingEv.getEndTime();
        
        const rawDuration = Math.round((evEnd.getTime() - evStart.getTime()) / 60000);
        const calendarDuration = snapToAllowedDuration(
          rawDuration > 0 ? rawDuration : DEFAULT_DURATION_MIN);

        const lastSyncMatch = desc.match(/\[LAST_SYNC_START:([^\]]+)\]/);
        const lastSyncISO = lastSyncMatch ? lastSyncMatch[1] : evStart.toISOString();
        const recurring = !!task.repeat;
        const reminderStart = rem && rem.remindAt
          ? new Date(rem.remindAt) : null;
        // Moving a one-off task's scheduled date in Craft can leave its timed
        // reminder on the previous date. Reuse its local clock time on the new
        // task date; the reminder itself is updated below before moving the event.
        if (!recurring && lastSyncMatch && reminderStart &&
            isFinite(reminderStart.getTime())) {
          const previousDate = Utilities.formatDate(new Date(lastSyncISO),
            timeZone, 'yyyy-MM-dd');
          const reminderDate = Utilities.formatDate(reminderStart,
            timeZone, 'yyyy-MM-dd');
          if (schedDateStr !== previousDate && reminderDate === previousDate) {
            startTime = combineTaskDateAndClock_(schedDateStr,
              reminderStart, timeZone);
          }
        }
        
        const lastDurMatch = desc.match(/\[LAST_SYNC_DUR:(\d+)\]/);
        const lastSyncDur = lastDurMatch ? parseInt(lastDurMatch[1], 10) : parsed.durationMin;

        // Compare both sides with the last synced duration. A changed Craft
        // tag sets the event length and wins if both sides changed.
        const craftDurationChanged = parsed.durationMin !== lastSyncDur;
        const gcalResizedByUser = !craftDurationChanged && calendarDuration !== lastSyncDur;
        const syncedDur = craftDurationChanged ? parsed.durationMin : calendarDuration;
        const durationTagMissing = !hasDurationTag_(task.markdown || '');
        let durationTagWritten = false;
        if (rawDuration !== syncedDur) {
          existingEv.setTime(evStart, new Date(evStart.getTime() + syncedDur * 60000));
        }

        const gcalMovedByUser = Math.abs(evStart.getTime() - new Date(lastSyncISO).getTime()) > 60000;
        // Completing an occurrence may advance scheduleDate without changing
        // the repeat clock time, so compare the full start time.
        const craftStartChanged = startTime &&
          Math.abs(startTime.getTime() - new Date(lastSyncISO).getTime()) > 60000;
        const lastReminderMatch = desc.match(/\[LAST_SYNC_REMINDER_AT:([^\]]+)\]/);
        const lastReminderAt = lastReminderMatch ? new Date(lastReminderMatch[1]) : null;
        const reminderDiffersFromEvent = !recurring &&
          (!reminderStart || !isFinite(reminderStart.getTime()) ||
           Math.abs(reminderStart.getTime() - evStart.getTime()) > 60000);
        const craftReminderChanged = lastReminderAt && reminderStart &&
          isFinite(lastReminderAt.getTime()) && isFinite(reminderStart.getTime()) &&
          Math.abs(reminderStart.getTime() - lastReminderAt.getTime()) > 60000;
        if (!recurring && reminderStart && !reminderDiffersFromEvent) {
          const oldDesc = desc;
          if (!lastReminderMatch) {
            desc = upsertMarker(desc, 'LAST_SYNC_REMINDER_AT', reminderStart.toISOString());
          }
          if (rem && rem.id && desc.includes('[CRAFT_OFFSET_MODE:true]')) {
            desc = desc.replace('[CRAFT_OFFSET_MODE:true]',
              '[CRAFT_REMINDER_ID:' + rem.id + ']');
          }
          if (desc !== oldDesc) existingEv.setDescription(desc);
        }
        const localClock = Utilities.formatDate(evStart, timeZone, 'HH:mm').split(':');
        const newOffsetMinutes = Number(localClock[0]) * 60 + Number(localClock[1]);
        const previousOffsetMatch = desc.match(/\[LAST_SYNC_REPEAT_OFFSET:(-?\d+)\]/);
        const previousOffset = previousOffsetMatch ? Number(previousOffsetMatch[1]) : null;
        const craftOffset = recurring && task.repeat.reminder &&
          task.repeat.reminder.enabled ? task.repeat.reminder.dateOffset : null;
        // Recover tasks whose older event markers say they are synced even
        // though the repeat clock time has not changed.
        const recurringTimeMismatch = recurring &&
          (!task.repeat.reminder || !task.repeat.reminder.enabled ||
           task.repeat.reminder.dateOffset !== newOffsetMinutes);
        const craftChangedSinceSync = previousOffset !== null &&
          craftOffset !== null && craftOffset !== previousOffset;
        const craftWins = !gcalMovedByUser && startTime &&
          (craftStartChanged || (recurringTimeMismatch && craftChangedSinceSync) ||
           (!recurring && reminderDiffersFromEvent && craftReminderChanged));

        // Craft keeps the recurring task active under the same ID and archives
        // completed or canceled occurrences inside it. Keep the old event only
        // when its occurrence is found in Craft history.
        const oldDate = Utilities.formatDate(evStart, timeZone, 'yyyy-MM-dd');
        const archivedState = recurring && craftStartChanged &&
          !gcalMovedByUser && startTime.getTime() > evStart.getTime() &&
          getArchivedOccurrenceState_(blockId, oldDate, parsed.cleanTitle);
        if (archivedState) {
          const historyTitle = (archivedState === 'done' ? '✅ ' : '❎ ') +
            parsed.cleanTitle +
            (task.location && task.location.title ? ' [' + task.location.title + ']' : '');
          existingEv.setTitle(historyTitle);
          const historyLink = archivedState === 'done'
            ? 'CRAFT_COMPLETED_OCCURRENCE_OF' : 'CRAFT_CANCELED_OCCURRENCE_OF';
          const historyDate = archivedState === 'done'
            ? 'CRAFT_COMPLETED_DATE' : 'CRAFT_CANCELED_DATE';
          let historyDesc = desc.replace(
            /\[CRAFT_BLOCK_ID:[a-zA-Z0-9\-_\s]+\]/,
            '[' + historyLink + ':' + blockId + ']');
          historyDesc = upsertMarker(historyDesc, 'LAST_SYNC_STATE', archivedState);
          historyDesc = upsertMarker(historyDesc, historyDate, oldDate);
          existingEv.setDescription(historyDesc);
          createNextRecurringEvent_(cal, task, rem, startTime, parsed,
            expectedTitle, spaceId, craftAppLinkTemplate);
          return;
        }

        // Craft exposes the repeat rule and visible reminder separately. If
        // only the reminder changed in Craft, it sets the repeat and event time.
        if (recurring && craftReminderChanged && !gcalMovedByUser &&
            !craftStartChanged && rem && rem.remindAt) {
          const reminderTime = new Date(rem.remindAt);
          const reminderDate = Utilities.formatDate(reminderTime, timeZone, 'yyyy-MM-dd');
          const reminderClock = Utilities.formatDate(reminderTime, timeZone, 'HH:mm').split(':');
          const reminderOffset = Number(reminderClock[0]) * 60 + Number(reminderClock[1]);
          const reminderPayload = { id: blockId,
            repeat: JSON.parse(JSON.stringify(task.repeat)) };
          reminderPayload.repeat.reminder = Object.assign({},
            reminderPayload.repeat.reminder || {},
            { enabled: true, dateOffset: reminderOffset });
          if (reminderDate !== schedDateStr) {
            reminderPayload.taskInfo = { scheduleDate: reminderDate };
          }
          if (gcalResizedByUser && parsed.durationMin !== syncedDur) {
            reminderPayload.markdown = replaceOrAppendDurationTag(task.markdown || '', syncedDur);
          }
          assertItem(craftRequest('PUT', '/tasks',
            { tasksToUpdate: [reminderPayload] }), blockId);
          verifyTaskUpdate(blockId, {
            scheduleDate: reminderPayload.taskInfo && reminderPayload.taskInfo.scheduleDate,
            durationMin: reminderPayload.markdown ? syncedDur : null,
            repeatOffset: reminderOffset
          });
          existingEv.setTime(reminderTime,
            new Date(reminderTime.getTime() + syncedDur * 60000));
          desc = upsertMarker(desc, 'LAST_SYNC_START', reminderTime.toISOString());
          desc = upsertMarker(desc, 'LAST_SYNC_DUR', syncedDur);
          desc = upsertMarker(desc, 'LAST_SYNC_REPEAT_OFFSET', reminderOffset);
          desc = upsertMarker(desc, 'LAST_SYNC_REMINDER_AT', reminderTime.toISOString());
          desc = desc.replace(/\n?\[REMINDER_SYNC_PENDING:true\]/, '');
          desc = desc.replace('[CRAFT_OFFSET_MODE:true]',
            '[CRAFT_REMINDER_ID:' + rem.id + ']');
          desc = desc.replace(/^⏱️[^\n]*$/m, '⏱️ Duration: #min' + syncedDur);
          existingEv.setDescription(desc);
          if (currentEvTitle !== expectedTitle) existingEv.setTitle(expectedTitle);
          return;
        }

        if (craftWins) {
          if (!recurring && rem && rem.id && reminderStart && startTime &&
              Math.abs(reminderStart.getTime() - startTime.getTime()) > 60000) {
            if (!remindersAvailable) {
              throw new Error('Craft reminder could not be moved: reminders endpoint unavailable.');
            }
            assertItem(craftRequest('PUT', '/reminders', {
              remindersToUpdate: [{ id: rem.id,
                remindAt: startTime.toISOString() }]
            }), rem.id);
            verifyReminderTime(blockId, rem.id, startTime);
          }
          if (durationTagMissing || (gcalResizedByUser && parsed.durationMin !== syncedDur)) {
            assertItem(craftRequest('PUT', '/tasks', {
              tasksToUpdate: [{ id: blockId,
                markdown: replaceOrAppendDurationTag(task.markdown || '', syncedDur) }]
            }), blockId);
            verifyTaskUpdate(blockId, { durationMin: syncedDur,
              requireDurationTag: true });
            durationTagWritten = true;
          }
          existingEv.setTime(startTime, new Date(startTime.getTime() + syncedDur * 60000));
          desc = upsertMarker(desc, 'LAST_SYNC_START', startTime.toISOString());
          desc = upsertMarker(desc, 'LAST_SYNC_DUR', syncedDur);
          if (recurring) desc = upsertMarker(desc, 'LAST_SYNC_REPEAT_OFFSET', craftOffset);
          else desc = upsertMarker(desc, 'LAST_SYNC_REMINDER_AT', startTime.toISOString());
          existingEv.setDescription(desc);
        } else if (gcalMovedByUser || gcalResizedByUser || recurringTimeMismatch ||
                   reminderDiffersFromEvent) {
          const newDateStr = Utilities.formatDate(evStart, timeZone, 'yyyy-MM-dd');
          const payload = { id: blockId };
          if (newDateStr !== schedDateStr) {
            payload.taskInfo = { scheduleDate: newDateStr };
          }

          if (durationTagMissing || (gcalResizedByUser && parsed.durationMin !== syncedDur)) {
            payload.markdown = replaceOrAppendDurationTag(task.markdown || '', syncedDur);
          }

          // Recurring task times live in repeat.reminder.dateOffset. Preserve
          // the other recurrence fields.
          if (recurring && (gcalMovedByUser || recurringTimeMismatch)) {
            payload.repeat = JSON.parse(JSON.stringify(task.repeat));
            payload.repeat.reminder = Object.assign({}, payload.repeat.reminder || {}, {
              enabled: true,
              dateOffset: newOffsetMinutes
            });
          }

          if (payload.taskInfo || payload.markdown || payload.repeat) {
            assertItem(craftRequest('PUT', '/tasks', { tasksToUpdate: [payload] }), blockId);
            verifyTaskUpdate(blockId, {
              scheduleDate: payload.taskInfo && payload.taskInfo.scheduleDate,
              durationMin: payload.markdown ? syncedDur : null,
              requireDurationTag: !!payload.markdown,
              repeatOffset: payload.repeat ? newOffsetMinutes : null
            });
            if (payload.markdown) durationTagWritten = true;
          }

          if (!recurring && (gcalMovedByUser || reminderDiffersFromEvent)) {
            if (!remindersAvailable) {
              throw new Error('Time was not sent: reminders endpoint unavailable.');
            }
            let reminderId = rem && rem.id;
            if (rem && rem.id) {
              assertItem(craftRequest('PUT', '/reminders', {
                remindersToUpdate: [{ id: rem.id, remindAt: evStart.toISOString() }]
              }), rem.id);
            } else {
              const created = craftRequest('POST', '/reminders', {
                reminders: [{ blockId: blockId, remindAt: evStart.toISOString() }]
              });
              if (!created || !created.items || !created.items.some(item => item.blockId === blockId && item.id)) {
                throw new Error('Craft did not confirm reminder creation for ' + blockId);
              }
              const newReminder = created.items.find(item => item.blockId === blockId && item.id);
              reminderId = newReminder.id;
              desc = desc.replace(/\[CRAFT_OFFSET_MODE:true\]/, '[CRAFT_REMINDER_ID:' + newReminder.id + ']');
            }
            // A response ID does not prove the time was saved. Advance the
            // sync marker only after reading the reminder back from Craft.
            verifyReminderTime(blockId, reminderId, evStart);
            desc = upsertMarker(desc, 'LAST_SYNC_REMINDER_AT', evStart.toISOString());
          }

          desc = upsertMarker(desc, 'LAST_SYNC_START', evStart.toISOString());
          desc = upsertMarker(desc, 'LAST_SYNC_DUR', syncedDur);
          if (recurring) desc = upsertMarker(desc, 'LAST_SYNC_REPEAT_OFFSET', newOffsetMinutes);
          existingEv.setDescription(desc);
        }

        // Add the default duration to the task text too; otherwise Calendar
        // shows 15 minutes but Craft does not show that value.
        if (durationTagMissing && !durationTagWritten) {
          assertItem(craftRequest('PUT', '/tasks', {
            tasksToUpdate: [{ id: blockId,
              markdown: replaceOrAppendDurationTag(task.markdown || '', syncedDur) }]
          }), blockId);
          verifyTaskUpdate(blockId, { durationMin: syncedDur,
            requireDurationTag: true });
        }

        // Refresh the marker and visible label when only the Craft tag changed.
        const durationLabel = '⏱️ Duration: #min' + syncedDur;
        const durationDesc = /^⏱️[^\n]*$/m.test(desc)
          ? desc.replace(/^⏱️[^\n]*$/m, durationLabel)
          : durationLabel + '\n' + desc;
        const updatedDesc = upsertMarker(durationDesc, 'LAST_SYNC_DUR', syncedDur);
        if (updatedDesc !== desc) existingEv.setDescription(updatedDesc);

        // Changing repeat.reminder.dateOffset may not update the reminder shown
        // in Craft's task list, so update both resources.
        if (recurring) {
          const displayedStart = existingEv.getStartTime();
          const reminderResult = ensureRecurringReminder_(blockId, rem,
            displayedStart, remindersAvailable);
          if (!reminderResult.synced) pendingReminders++;
          const reminderDesc = describeRecurringReminder_(updatedDesc,
            reminderResult, displayedStart);
          if (reminderDesc !== updatedDesc) existingEv.setDescription(reminderDesc);
        }
      }

      // Preserve the historical occurrence for all-day recurring tasks too.
      if (wasAllDayInGCal && !startTime) {
        const oldDate = Utilities.formatDate(existingEv.getStartTime(),
          timeZone, 'yyyy-MM-dd');
        const lastDateMatch = desc.match(/\[LAST_SYNC_DATE:(\d{4}-\d{2}-\d{2})\]/);
        // Events created before V40 lack this marker. On the first sync, keep
        // the previous rule: Craft determines the date.
        const lastDate = lastDateMatch ? lastDateMatch[1] : oldDate;
        const craftDateChanged = schedDateStr !== lastDate;
        const calendarDateChanged = oldDate !== lastDate;
        if (oldDate !== schedDateStr) {
          const dayStart = Utilities.parseDate(schedDateStr + ' 00:00',
            timeZone, 'yyyy-MM-dd HH:mm');
          const archivedState = task.repeat && craftDateChanged &&
            schedDateStr > oldDate
            ? getArchivedOccurrenceState_(blockId, oldDate, parsed.cleanTitle) : null;
          if (archivedState) {
            const historyTitle = (archivedState === 'done' ? '✅ ' : '❎ ') +
              parsed.cleanTitle +
              (task.location && task.location.title ? ' [' + task.location.title + ']' : '');
            existingEv.setTitle(historyTitle);
            const historyLink = archivedState === 'done'
              ? 'CRAFT_COMPLETED_OCCURRENCE_OF' : 'CRAFT_CANCELED_OCCURRENCE_OF';
            const historyDate = archivedState === 'done'
              ? 'CRAFT_COMPLETED_DATE' : 'CRAFT_CANCELED_DATE';
            let historyDesc = desc.replace(
              /\[CRAFT_BLOCK_ID:[a-zA-Z0-9\-_\s]+\]/,
              '[' + historyLink + ':' + blockId + ']');
            historyDesc = upsertMarker(historyDesc, 'LAST_SYNC_STATE', archivedState);
            historyDesc = upsertMarker(historyDesc, historyDate, oldDate);
            existingEv.setDescription(historyDesc);

            const deepLink = craftTaskLink_(craftAppLinkTemplate, spaceId, blockId);
            const newDesc = [
              `🔗 Open in Craft: ${deepLink}`,
              `📄 Context: ${(task.location && task.location.title) || 'Daily Note / Inbox'}`,
              '',
              '--- Sync metadata (Do not edit) ---',
              `[CRAFT_BLOCK_ID:${blockId}]`,
              `[LAST_SYNC_STATE:${taskState}]`,
              `[LAST_SYNC_DATE:${schedDateStr}]`,
              '[IS_ALL_DAY:true]'
            ].join('\n');
            cal.createAllDayEvent(expectedTitle, dayStart, { description: newDesc });
            return;
          }
          if (calendarDateChanged && !craftDateChanged) {
            assertItem(craftRequest('PUT', '/tasks', {
              tasksToUpdate: [{ id: blockId,
                taskInfo: { scheduleDate: oldDate } }]
            }), blockId);
            verifyTaskUpdate(blockId, { scheduleDate: oldDate });
            desc = upsertMarker(desc, 'LAST_SYNC_DATE', oldDate);
          } else {
            // Craft changed, or both sides changed: Craft wins.
            existingEv.setAllDayDate(dayStart);
            desc = upsertMarker(desc, 'LAST_SYNC_DATE', schedDateStr);
          }
        } else if (!lastDateMatch || oldDate !== lastDate) {
          desc = upsertMarker(desc, 'LAST_SYNC_DATE', oldDate);
        }
        if (desc !== existingEv.getDescription()) existingEv.setDescription(desc);
      }

      // 3. Refresh the title if it differs or contains stale formatting.
      const hasResidue = currentEvTitle.includes('(') || currentEvTitle.includes('~') || currentEvTitle.includes('**');
      if (currentEvTitle !== expectedTitle || hasResidue) {
        existingEv.setTitle(expectedTitle);
      }
    }
    } catch (error) {
      Logger.log('Task failed ' + task.id + ': ' + error.message);
      failures.push(task.id + ': ' + error.message);
    }
  });

  // Delete linked events only when the task is absent from Craft's full list.
  // Also clean up older pending events in the window. Confirm with a second
  // read to avoid deleting events after a transient response.
  const deletionStartWindow = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const missingEvents = existingEvents.map(ev => {
    const desc = ev.getDescription() || '';
    const blockId = extractCraftMarkerId_(desc, 'CRAFT_BLOCK_ID');
    const unfinished = /\[LAST_SYNC_STATE:todo\]/.test(desc) ||
      ev.getTitle().startsWith('❇️');
    return blockId && (ev.getStartTime() >= deletionStartWindow || unfinished)
      ? { event: ev, blockId: blockId } : null;
  }).filter(item => item && !taskMap[item.blockId]);
  if (missingEvents.length) {
    const confirmation = craftRequest('GET', '/tasks?scope=all');
    if (!confirmation || !Array.isArray(confirmation.items)) {
      throw new Error('Craft did not confirm the task list; no events will be deleted.');
    }
    const confirmedIds = new Set(confirmation.items.filter(task => task && task.id)
      .map(task => task.id));
    const confirmedMissing = missingEvents.filter(item => !confirmedIds.has(item.blockId));
    if (confirmedMissing.length > MAX_CALENDAR_DELETIONS_PER_RUN) {
      throw new Error('Too many events lost their Craft task (' +
        confirmedMissing.length + '); automatic deletion stopped for review.');
    }
    confirmedMissing.forEach(item => {
      try {
        item.event.deleteEvent();
      } catch (error) {
        Logger.log('Failed to delete event for task ' + item.blockId + ': ' + error.message);
        failures.push(item.blockId + ': event deletion failed: ' + error.message);
      }
    });
  }
  if (failures.length) {
    throw new Error('Sync incomplete (' + failures.length + ' failure(s)). ' +
      failures.slice(0, 3).join(' | '));
  }
  return pendingReminders;
}

function getArchivedOccurrenceState_(blockId, dateStr, expectedTitle) {
  const response = craftRequest('GET', '/blocks?id=' + encodeURIComponent(blockId));

  // The endpoint may return the block directly or inside data/items/blocks.
  // Check its ID before using content from another task's history.
  function findBlock(value, depth) {
    if (!value || depth > 5) return null;
    if (Array.isArray(value)) {
      for (const item of value) {
        const found = findBlock(item, depth + 1);
        if (found) return found;
      }
      return null;
    }
    if (typeof value !== 'object') return null;
    if (value.id === blockId) return value;
    for (const key of ['data', 'items', 'blocks', 'block', 'content', 'result']) {
      const found = findBlock(value[key], depth + 1);
      if (found) return found;
    }
    return null;
  }
  const root = findBlock(response, 0);
  if (!root) {
    const keys = response && typeof response === 'object'
      ? Object.keys(response).join(', ') : typeof response;
    throw new Error('Craft did not return the recurring block ' + blockId +
      ' (received fields: ' + keys + ')');
  }
  // Recognize archived headings created by Craft in multiple locales.
  const [year, month, day] = dateStr.split('-').map(Number);
  const date = new Date(Date.UTC(year, month - 1, day, 12));
  const headings = ['pt-BR', 'en-US', 'en-GB'].map(locale =>
    new Intl.DateTimeFormat(locale, { day: 'numeric', month: 'short',
      timeZone: 'UTC' }).format(date));
  const dateInTitle = new RegExp('\\*\\*(?:' + headings.map(heading => {
    const escaped = heading.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return day < 10 && heading.startsWith(String(day))
      ? '0?' + escaped : escaped;
  }).join('|') + ')\\*\\*', 'i');
  const normalizedTitle = expectedTitle.trim().toLocaleLowerCase();

  const matchedStates = [];
  function visit(block) {
    if (block && block.id !== blockId && block.taskInfo &&
        (block.taskInfo.state === 'done' || block.taskInfo.state === 'canceled') &&
        parseTaskWithTags(block.markdown || '').cleanTitle.toLocaleLowerCase() ===
          normalizedTitle &&
        (block.taskInfo.scheduleDate === dateStr ||
         dateInTitle.test(block.markdown || ''))) {
      matchedStates.push(block.taskInfo.state);
    }
    if (block && Array.isArray(block.content)) block.content.forEach(visit);
  }
  visit(root);
  const states = Array.from(new Set(matchedStates));
  if (states.length > 1) {
    throw new Error('Ambiguous recurrence history for ' + blockId +
      ' on ' + dateStr);
  }
  return states[0] || null;
}

function createNextRecurringEvent_(cal, task, reminder, startTime, parsed,
    title, spaceId, craftAppLinkTemplate) {
  const blockId = task.id;
  const deepLink = craftTaskLink_(craftAppLinkTemplate, spaceId, blockId);
  const reminderAligned = reminder && reminder.remindAt
    ? syncRecurringReminder_(blockId, reminder, startTime) : false;
  const description = [
    `🔗 Open in Craft: ${deepLink}`,
    `⏱️ Duration: #min${parsed.durationMin}`,
    `📄 Context: ${(task.location && task.location.title) || 'Daily Note / Inbox'}`,
    '',
    '--- Sync metadata (Do not edit) ---',
    `[CRAFT_BLOCK_ID:${blockId}]`,
    `[LAST_SYNC_STATE:${task.taskInfo.state || 'todo'}]`,
    reminder ? `[CRAFT_REMINDER_ID:${reminder.id}]` : '[CRAFT_OFFSET_MODE:true]',
    `[LAST_SYNC_START:${startTime.toISOString()}]`,
    `[LAST_SYNC_DUR:${parsed.durationMin}]`,
    reminder && reminder.remindAt
      ? `[LAST_SYNC_REMINDER_AT:${reminderAligned
        ? startTime.toISOString() : new Date(reminder.remindAt).toISOString()}]` : '',
    reminder && reminder.remindAt && !reminderAligned
      ? '[REMINDER_SYNC_PENDING:true]' : '',
    task.repeat && task.repeat.reminder &&
      typeof task.repeat.reminder.dateOffset === 'number'
        ? `[LAST_SYNC_REPEAT_OFFSET:${task.repeat.reminder.dateOffset}]` : ''
  ].join('\n');
  const endTime = new Date(startTime.getTime() + parsed.durationMin * 60000);
  cal.createEvent(title, startTime, endTime, { description: description });
}

// The Craft connection supplies a link template for the task block.
// documentId points to its note, which may change when the task moves.
function craftTaskLink_(template, spaceId, blockId) {
  const safeBlockId = encodeURIComponent(blockId);
  if (typeof template === 'string' &&
      /^craftdocs:\/\/open\?/.test(template) && template.includes('{blockId}')) {
    return template.replace(/\{blockId\}/g, safeBlockId);
  }
  return 'craftdocs://open?' + (spaceId
    ? 'spaceId=' + encodeURIComponent(spaceId) + '&' : '') +
    'blockId=' + safeBlockId;
}

// Repair old descriptions too, even if the link wrapped across lines.
function repairCraftLink_(description, link) {
  const label = '🔗 Open in Craft:';
  const start = description.indexOf('🔗');
  if (start < 0) return label + ' ' + link + '\n' + description;
  const markers = ['⏱️', '📄', '---', '[CRAFT_BLOCK_ID:',
    '[CRAFT_COMPLETED_OCCURRENCE_OF:', '[CRAFT_CANCELED_OCCURRENCE_OF:'];
  const ends = markers.map(marker => description.indexOf(marker, start + 2))
    .filter(index => index >= 0);
  const end = ends.length ? Math.min.apply(null, ends) : description.length;
  const replacement = label + ' ' + link + '\n';
  return description.slice(0, start) + replacement + description.slice(end);
}

function parseTaskWithTags(md) {
  let text = md;

  text = text.replace(/\([^\)]*\)/g, '');
  text = text.replace(/\*\*[^*]+\*\*/g, '');
  text = text.replace(/~/g, '');
  text = text.replace(/^\s*-\s*/, '');
  text = text.replace(/\[\s*[xX\s]?\s*\]/g, '');
  text = text.replace(/\[([^\]]+)\]\(block:\/\/[^\)]+\)/g, '$1');

  text = text.trim();

  const tagRegex = /(?:^|\s)(?:#|\{)min(15|30|45|60)\b\}?/i;
  const match = text.match(tagRegex);

  let durationMin = DEFAULT_DURATION_MIN;
  if (match) {
    durationMin = parseInt(match[1], 10);
    text = text.replace(tagRegex, '').trim();
  }

  text = text.replace(/\s*—\s*$/, '');
  text = text.replace(/\s{2,}/g, ' ').trim();

  return {
    cleanTitle: text || 'Craft task',
    durationMin: durationMin
  };
}

function snapToAllowedDuration(mins) {
  return ALLOWED_DURATIONS.reduce((prev, curr) =>
    Math.abs(curr - mins) < Math.abs(prev - mins) ? curr : prev
  );
}

function hasDurationTag_(md) {
  return /(?:^|\s)(?:#|\{)min(?:15|30|45|60)\b\}?/i.test(md);
}

function replaceOrAppendDurationTag(md, newDur) {
  const tagRegex = /(?:#|\{)min(15|30|45|60)\b\}?/i;
  if (tagRegex.test(md)) {
    return md.replace(tagRegex, `#min${newDur}`);
  }
  return `${md.trim()} #min${newDur}`;
}

function extractTaskStartTime(task, reminderObj, timeZone) {
  const tInfo = task.taskInfo || {};
  const schedDate = tInfo.scheduleDate;
  if (!schedDate) return null;

  let offsetMinutes = null;
  // GET /tasks returns repeat at the task level. For recurring tasks, that
  // time takes priority over a standalone reminder for the same block.
  if (task.repeat && task.repeat.reminder && task.repeat.reminder.enabled &&
      typeof task.repeat.reminder.dateOffset === 'number') {
    offsetMinutes = task.repeat.reminder.dateOffset;
  } else if (reminderObj && reminderObj.remindAt) {
    return new Date(reminderObj.remindAt);
  } else if (tInfo.reminder && tInfo.reminder.enabled && typeof tInfo.reminder.dateOffset === 'number') {
    offsetMinutes = tInfo.reminder.dateOffset;
  } else if (tInfo.repeat && tInfo.repeat.reminder && tInfo.repeat.reminder.enabled && typeof tInfo.repeat.reminder.dateOffset === 'number') {
    offsetMinutes = tInfo.repeat.reminder.dateOffset;
  }

  if (offsetMinutes === null) return null;

  const parts = schedDate.split('-');
  if (parts.length !== 3) return null;

  const hours = Math.floor(offsetMinutes / 60);
  const minutes = offsetMinutes % 60;

  const clock = String(hours).padStart(2, '0') + ':' + String(minutes).padStart(2, '0');
  return Utilities.parseDate(schedDate + ' ' + clock, timeZone, 'yyyy-MM-dd HH:mm');
}

function combineTaskDateAndClock_(date, sourceTime, timeZone) {
  const clock = Utilities.formatDate(sourceTime, timeZone, 'HH:mm');
  return Utilities.parseDate(date + ' ' + clock, timeZone, 'yyyy-MM-dd HH:mm');
}

function assertItem(response, id) {
  if (!response || !Array.isArray(response.items) ||
      !response.items.some(item => item && item.id === id)) {
    throw new Error('Craft did not confirm the update of ' + id);
  }
}

function extractCraftMarkerId_(description, key) {
  const pattern = new RegExp('\\[' + key + ':([a-zA-Z0-9\\-_\\s]+)\\]');
  const match = description.match(pattern);
  return match ? match[1].replace(/\s/g, '') : null;
}

function upsertMarker(description, key, value) {
  const marker = '[' + key + ':' + value + ']';
  const pattern = new RegExp('\\[' + key + ':[^\\]]+\\]');
  return pattern.test(description) ? description.replace(pattern, marker)
    : description + '\n' + marker;
}

function verifyTaskUpdate(blockId, expected) {
  for (let attempt = 0; attempt < 3; attempt++) {
    const result = craftRequest('GET', '/tasks?scope=all');
    const saved = (result.items || []).find(item => item && item.id === blockId);
    const info = saved && saved.taskInfo || {};
    const stateOK = !expected.state || info.state === expected.state;
    const dateOK = !expected.scheduleDate || info.scheduleDate === expected.scheduleDate;
    const durationOK = !expected.durationMin ||
      (saved && parseTaskWithTags(saved.markdown || '').durationMin === expected.durationMin);
    const durationTagOK = !expected.requireDurationTag ||
      (saved && hasDurationTag_(saved.markdown || ''));
    const repeat = saved && saved.repeat && saved.repeat.reminder;
    const repeatOK = expected.repeatOffset == null ||
      (repeat && repeat.enabled && repeat.dateOffset === expected.repeatOffset);
    if (saved && stateOK && dateOK && durationOK && durationTagOK && repeatOK) return;
    if (attempt < 2) Utilities.sleep(500);
  }
  throw new Error('Craft did not confirm task data for ' + blockId);
}

function listReminders() {
  const items = [];
  let cursor = '';
  do {
    const page = craftRequest('GET', '/reminders?status=all&limit=200' +
      (cursor ? '&cursor=' + encodeURIComponent(cursor) : ''));
    items.push.apply(items, page.items || []);
    cursor = page.pagination && page.pagination.nextCursor || '';
  } while (cursor);
  return items;
}

function verifyReminderTime(blockId, reminderId, expectedTime) {
  for (let attempt = 0; attempt < 3; attempt++) {
    const saved = listReminders().find(item =>
      item && item.id === reminderId && item.blockId === blockId);
    if (saved && saved.remindAt &&
        Math.abs(new Date(saved.remindAt).getTime() - expectedTime.getTime()) <= 60000) return;
    if (attempt < 2) Utilities.sleep(500);
  }
  throw new Error('Craft did not save the new reminder time for ' + blockId);
}

function syncRecurringReminder_(blockId, reminder, expectedTime) {
  if (!reminder || !reminder.id || !reminder.remindAt) return false;
  if (Math.abs(new Date(reminder.remindAt).getTime() - expectedTime.getTime()) <= 60000) {
    return true;
  }
  // Craft rejects reminders in the past. Keep the repeat rule and event
  // correct, then retry when the next occurrence is in the future.
  if (expectedTime.getTime() <= Date.now() + 60000) {
    Logger.log('Reminder for task ' + blockId + ' remains at ' +
      reminder.remindAt + ': Craft rejects the past time ' +
      expectedTime.toISOString() + '. The next occurrence will be retried.');
    return false;
  }
  try {
    assertItem(craftRequest('PUT', '/reminders', {
      remindersToUpdate: [{ id: reminder.id, remindAt: expectedTime.toISOString() }]
    }), reminder.id);
  } catch (error) {
    if (/Reminder time must be in the future/i.test(error.message) &&
        expectedTime.getTime() <= Date.now() + 60000) {
      Logger.log('Craft rejected the reminder time for task ' + blockId +
        ' because it just passed. The next occurrence will be retried.');
      return false;
    }
    throw error;
  }
  verifyReminderTime(blockId, reminder.id, expectedTime);
  return true;
}

function ensureRecurringReminder_(blockId, reminder, expectedTime, available) {
  if (reminder && reminder.id && reminder.remindAt) {
    return { id: reminder.id, remindAt: reminder.remindAt,
      synced: available && syncRecurringReminder_(blockId, reminder, expectedTime) };
  }
  if (!available || expectedTime.getTime() <= Date.now() + 60000) {
    Logger.log('Reminder pending for ' + blockId + ': time missing or already past.');
    return { id: null, remindAt: null, synced: false };
  }
  const response = craftRequest('POST', '/reminders', {
    reminders: [{ blockId: blockId, remindAt: expectedTime.toISOString() }]
  });
  const created = response && response.items && response.items.find(item =>
    item && item.blockId === blockId && item.id);
  if (!created) throw new Error('Craft did not confirm the new reminder for ' + blockId);
  verifyReminderTime(blockId, created.id, expectedTime);
  return { id: created.id, remindAt: expectedTime.toISOString(), synced: true };
}

function describeRecurringReminder_(description, result, expectedTime) {
  let updated = description;
  if (result.id) {
    updated = updated.replace('[CRAFT_OFFSET_MODE:true]',
      '[CRAFT_REMINDER_ID:' + result.id + ']');
  }
  if (result.synced) {
    updated = upsertMarker(updated, 'LAST_SYNC_REMINDER_AT',
      expectedTime.toISOString());
    return updated.replace(/\n?\[REMINDER_SYNC_PENDING:true\]/, '');
  }
  if (result.remindAt) {
    updated = upsertMarker(updated, 'LAST_SYNC_REMINDER_AT',
      new Date(result.remindAt).toISOString());
  }
  return upsertMarker(updated, 'REMINDER_SYNC_PENDING', 'true');
}

function craftRequest(method, endpoint, payload) {
  const url = CRAFT_API_BASE + endpoint;
  const options = {
    method: method,
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  if (code >= 200 && code < 300) {
    const content = response.getContentText();
    return content ? JSON.parse(content) : {};
  }

  throw new Error(`API Craft [${method} ${endpoint}] HTTP ${code}: ${response.getContentText().slice(0, 500)}`);
}
