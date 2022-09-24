/************************************************************************************************************
Krahmer Chore Chart
Copyright 2022 Douglas Krahmer

This file is part of Krahmer Chore Chart.

Krahmer Chore Chart is free software: you can redistribute it and/or modify it under the terms of the 
GNU General Public License as published by the Free Software Foundation, either version 3 of the License, 
or (at your option) any later version.

Krahmer Chore Chart is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. 
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with Krahmer Chore Chart.
If not, see <https://www.gnu.org/licenses/>.
************************************************************************************************************/

function processRecurringTasks() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const recurringTasksSheet = sheet.getSheetByName(RECURRING_TASKS_SHEET_NAME);

  const skipRows = recurringTasksSheet.getFrozenRows();
  const rows = recurringTasksSheet.getMaxRows();
  const rowsRange = recurringTasksSheet.getRange(1 + skipRows, 1, rows - skipRows, RECUR_TASKS_COL_COUNT);
  const rowsValues = rowsRange.getValues();

  // processing in reverse order will retain the relative task order since we insert to the top of the task list
  for (let row = rows; row >= 1 + skipRows; row--) {
    const recurringTaskData = rowsValues[row - (1 + skipRows)];
    const recurringTask = {
      enabled: recurringTaskData[RECUR_TASKS_COL_ENABLED - 1],
      isNeedNow: recurringTaskData[RECUR_TASKS_COL_DAY_NEED_NOW - 1],
      name: recurringTaskData[RECUR_TASKS_COL_NAME - 1],
      assignedToArray: [recurringTaskData[RECUR_TASKS_COL_ASSIGNED_TO - 1]],
      nextDueDate: recurringTaskData[RECUR_TASKS_COL_NEXT_DUE_DATE - 1],
      createDaysBeforeDue: recurringTaskData[RECUR_TASKS_COL_CREATE_DAYS_BEFORE_DUE - 1],
      recurringDays: recurringTaskData[RECUR_TASKS_COL_RECURRING_DAYS - 1],
      recurringMonths: recurringTaskData[RECUR_TASKS_COL_RECURRING_MONTHS - 1],
      daysOfWeek: recurringTaskData[RECUR_TASKS_COL_DAYS_OF_WEEK - 1],
      dayOccuranceInMonth: recurringTaskData[RECUR_TASKS_COL_DAY_OCCURANCE_IN_MONTH - 1],
      endDate: recurringTaskData[RECUR_TASKS_COL_END_DATE - 1],
      lastDueDate: recurringTaskData[RECUR_TASKS_LAST_DUE_DATE - 1]
    }

    for (let colIndex = RECUR_TASKS_COL_ASSIGNED_TO + 1; colIndex <= RECUR_TASKS_COL_ASSIGNED_TO_NEXT; colIndex++) {
      recurringTask.assignedToArray.push(recurringTaskData[colIndex - 1]);
    }

    if (!recurringTask.name)
      continue; // no name

    if (!recurringTask.enabled && !recurringTask.isNeedNow)
      continue; // not enabled

    if (!recurringTask.nextDueDate && !recurringTask.isNeedNow)
      continue; // no next date
    
    const currentNewTasks = processRecurringTask(sheet, recurringTask);
    
    if (currentNewTasks) {
      const range = recurringTasksSheet.getRange(row, 1, 1, RECUR_TASKS_COL_COUNT);

      if (range.getValues()[0][RECUR_TASKS_COL_NAME - 1] != recurringTask.name)
        break; // the data row no longer matches. may have been edited while iterating.

      // update recurring task record in sheet
      range.setValues([[
        recurringTask.enabled,
        recurringTask.isNeedNow,
        recurringTask.name,
        ...recurringTask.assignedToArray,
        recurringTask.nextDueDate,
        recurringTask.createDaysBeforeDue,
        recurringTask.recurringDays,
        recurringTask.recurringMonths,
        recurringTask.daysOfWeek,
        recurringTask.dayOccuranceInMonth,
        recurringTask.endDate,
        recurringTask.lastDueDate
      ]]);
      sortTasksSheet(); // sort after each addition
    }
  }

  cleanTasks(sheet);
}

function processRecurringTask(sheet, recurringTask) {
  console.log(`${getFuncName()} - ${recurringTask?.name}...`);
  let newTasks = 0;

  const now = new Date();
  let resume = !(recurringTask.isNeedNow || recurringTask.recurringDays == -1);
  do {
    if (recurringTask.isNeedNow) {
      recurringTask.nextDueDateSave = recurringTask.nextDueDate;
      recurringTask.nextDueDate = new Date();
      recurringTask.nextDueDate.setHours(0, 0, 0, 0);
    }
    else {
      const nextAddDate = getNextRunDate(recurringTask, "nextDueDate");
      
      if (!nextAddDate || nextAddDate > now)
        break;
    }

    recurringTask.dueDate = new Date(recurringTask.nextDueDate);
    addTask(sheet, recurringTask);
    advanceRecurringTask(recurringTask);
    recurringTask.lastDueDate = recurringTask.dueDate;

    newTasks++;
  } while (resume);

  return newTasks;
}

function advanceRecurringTask(recurringTask) {
  console.log(`${getFuncName()}...`);
  if (recurringTask.assignedToArray[1]) {
    const assignedTo = recurringTask.assignedToArray.shift();
    const insertIndex = recurringTask.assignedToArray.indexOf("");
    recurringTask.assignedToArray.splice(insertIndex === -1 ? recurringTask.assignedToArray.length : insertIndex, 0, assignedTo);
  }

  if (recurringTask.isNeedNow) {
      recurringTask.nextDueDate = recurringTask.nextDueDateSave ?? recurringTask.nextDueDate; // restore if saved
      recurringTask.isNeedNow = false;
      return;
  }

  if (recurringTask.recurringDays == -1) {
    recurringTask.nextDueDate = "";
    return;
  }
  if (recurringTask.dayOccuranceInMonth)
    recurringTask.nextDueDate.setDate(1); // reset the day in case the next month has more days

  if (recurringTask.recurringMonths || recurringTask.dayOccuranceInMonth)
    recurringTask.nextDueDate.setMonth(recurringTask.dueDate.getMonth() + (recurringTask.recurringMonths || 1));

  if (recurringTask.recurringDays || (!recurringTask.recurringMonths && !recurringTask.dayOccuranceInMonth))
    recurringTask.nextDueDate.setDate(recurringTask.dueDate.getDate() + (recurringTask.recurringDays || 1));

  if (recurringTask.dayOccuranceInMonth) {
    recurringTask.nextDueDate = getMonthDayOccurance(recurringTask.nextDueDate, recurringTask.daysOfWeek, recurringTask.dayOccuranceInMonth[0])
  }
  else {
    const savedNextDueDate = new Date(recurringTask.nextDueDate);

    // make sure the date is valid
    let attemptsLeft = 400;
    // Increment by 1 day until a valid date is found
    while (!isValidDate(recurringTask, recurringTask.nextDueDate)) {
      if (attemptsLeft-- <= 0) {
        // circuit breaker - to many date checks, give up
        recurringTask.nextDueDate = savedNextDueDate;
        break;
      }
      recurringTask.nextDueDate.setDate(recurringTask.nextDueDate.getDate() + 1);
    }
  }

  if (recurringTask.endDate && recurringTask.nextDueDate > recurringTask.endDate) {
    recurringTask.nextDueDate = ""; // the task has ended
  }
}

function addDaysToRecurringTask(taskName, daysToAdd) {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const recurringTasksSheet = sheet.getSheetByName(RECURRING_TASKS_SHEET_NAME);

  const skipRows = recurringTasksSheet.getFrozenRows();
  const rows = recurringTasksSheet.getMaxRows();
  const rowsRange = recurringTasksSheet.getRange(1 + skipRows, 1, rows - skipRows, RECUR_TASKS_COL_COUNT);
  const rowsValues = rowsRange.getValues();

  // processing in reverse order will retain the relative task order since we insert to the top of the task list
  for (let row = rows; row >= 1 + skipRows; row--) {
    const recurringTaskData = rowsValues[row - (1 + skipRows)];
    const recurringTask = {
      enabled: recurringTaskData[RECUR_TASKS_COL_ENABLED - 1],
      isNeedNow: recurringTaskData[RECUR_TASKS_COL_DAY_NEED_NOW - 1],
      name: recurringTaskData[RECUR_TASKS_COL_NAME - 1],
      assignedToArray: [recurringTaskData[RECUR_TASKS_COL_ASSIGNED_TO - 1]],
      nextDueDate: recurringTaskData[RECUR_TASKS_COL_NEXT_DUE_DATE - 1],
      createDaysBeforeDue: recurringTaskData[RECUR_TASKS_COL_CREATE_DAYS_BEFORE_DUE - 1],
      recurringDays: recurringTaskData[RECUR_TASKS_COL_RECURRING_DAYS - 1],
      recurringMonths: recurringTaskData[RECUR_TASKS_COL_RECURRING_MONTHS - 1],
      daysOfWeek: recurringTaskData[RECUR_TASKS_COL_DAYS_OF_WEEK - 1],
      dayOccuranceInMonth: recurringTaskData[RECUR_TASKS_COL_DAY_OCCURANCE_IN_MONTH - 1],
      endDate: recurringTaskData[RECUR_TASKS_COL_END_DATE - 1],
      lastDueDate: recurringTaskData[RECUR_TASKS_LAST_DUE_DATE - 1]
    }

    for (let colIndex = RECUR_TASKS_COL_ASSIGNED_TO + 1; colIndex <= RECUR_TASKS_COL_ASSIGNED_TO_NEXT; colIndex++) {
      recurringTask.assignedToArray.push(recurringTaskData[colIndex - 1]);
    }

    if (!recurringTask.name)
      continue; // no name

    if (!recurringTask.enabled && !recurringTask.isNeedNow)
      continue; // not enabled

    if (!recurringTask.nextDueDate && !recurringTask.isNeedNow)
      continue; // no next date
    
    if (recurringTask.name !== taskName)
      continue;

    recurringTask.nextDueDate.setDate(recurringTask.nextDueDate.getDate() + daysToAdd);

    const range = recurringTasksSheet.getRange(row, 1, 1, RECUR_TASKS_COL_COUNT);

    if (range.getValues()[0][RECUR_TASKS_COL_NAME - 1] != recurringTask.name)
      break; // the data row no longer matches. may have been edited while iterating.

    // update recurring task record in sheet
    range.setValues([[
      recurringTask.enabled,
      recurringTask.isNeedNow,
      recurringTask.name,
      ...recurringTask.assignedToArray,
      recurringTask.nextDueDate,
      recurringTask.createDaysBeforeDue,
      recurringTask.recurringDays,
      recurringTask.recurringMonths,
      recurringTask.daysOfWeek,
      recurringTask.dayOccuranceInMonth,
      recurringTask.endDate,
      recurringTask.lastDueDate
    ]]);
  }
}
