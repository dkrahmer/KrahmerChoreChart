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

function processRecurringActions() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const recurringActionsSheet = sheet.getSheetByName(RECURRING_ACTIONS_SHEET_NAME);

  const skipRows = recurringActionsSheet.getFrozenRows();
  const rows = recurringActionsSheet.getMaxRows();
  const rowsRange = recurringActionsSheet.getRange(1 + skipRows, 1, rows - skipRows, RECUR_ACTIONS_COL_COUNT);
  const rowsValues = rowsRange.getValues();

  // process in reverse order to be consistent with recurring tasks processing
  for (let row = rows; row >= 1 + skipRows; row--) {
    const recurringActionData = rowsValues[row - (1 + skipRows)];
    const recurringAction = {
      enabled: recurringActionData[RECUR_ACTIONS_COL_ENABLED - 1],
      name: recurringActionData[RECUR_ACTIONS_COL_NAME - 1],
      action: recurringActionData[RECUR_ACTIONS_COL_ACTION - 1],
      parameters: recurringActionData[RECUR_ACTIONS_COL_PARAMETERS - 1],
      nextRunDate: recurringActionData[RECUR_ACTIONS_COL_NEXT_RUN_DATE - 1],
      runAdjustDays: recurringActionData[RECUR_ACTIONS_COL_RUN_ADJUST_DAYS - 1],
      recurringDays: recurringActionData[RECUR_ACTIONS_COL_RECURRING_DAYS - 1],
      recurringMonths: recurringActionData[RECUR_ACTIONS_COL_RECURRING_MONTHS - 1],
      daysOfWeek: recurringActionData[RECUR_ACTIONS_COL_DAYS_OF_WEEK - 1],
      dayOccuranceInMonth: recurringActionData[RECUR_ACTIONS_COL_DAY_OCCURANCE_IN_MONTH - 1],
      endDate: recurringActionData[RECUR_ACTIONS_COL_END_DATE - 1],
      notes: recurringActionData[RECUR_ACTIONS_COL_NOTES - 1],
      lastRunDate: recurringActionData[RECUR_ACTIONS_LAST_RUN_DATE - 1]
    }

    if (!recurringAction.name)
      continue; // no name

    if (!recurringAction.enabled)
      continue; // not enabled

    if (!recurringAction.nextRunDate)
      continue; // no next date

    if (!recurringAction.action)
      continue; // no action specified
    
    const currentNewActions = processRecurringAction(recurringAction);
    
    if (currentNewActions) {
      const range = recurringActionsSheet.getRange(row, 1, 1, RECUR_ACTIONS_COL_COUNT);

      if (range.getValues()[0][RECUR_ACTIONS_COL_NAME - 1] != recurringAction.name)
        break; // the data row no longer matches. may have been edited while iterating.

      // update recurring action record in sheet
      range.setValues([[
        recurringAction.enabled,
        recurringAction.name,
        recurringAction.action,
        recurringAction.parameters,
        recurringAction.nextRunDate,
        recurringAction.runAdjustDays,
        recurringAction.recurringDays,
        recurringAction.recurringMonths,
        recurringAction.daysOfWeek,
        recurringAction.dayOccuranceInMonth,
        recurringAction.endDate,
        recurringAction.notes,
        recurringAction.lastRunDate
      ]]);
    }
  }
}

function processRecurringAction(recurringAction) {
  console.log(`${getFuncName()} - ${recurringAction?.name}...`);
  let newActions = 0;

  const now = new Date();
  let resume = !(recurringAction.recurringDays == -1);
  do {
    const nextRunDate = getNextRunDate(recurringAction, "nextRunDate");
    
    if (!nextRunDate || nextRunDate > now)
      break;

    recurringAction.runDate = new Date(nextRunDate);

    // Run if it is the actual date, otherwise just advance to the next date
    if (now.getFullYear() === recurringAction.runDate.getFullYear()
        && now.getMonth() === recurringAction.runDate.getMonth()
        && now.getDate() === recurringAction.runDate.getDate()) {
      handleAction(recurringAction);
      recurringAction.lastRunDate = recurringAction.runDate;
    }

    advanceRecurringAction(recurringAction);

    newActions++;
  } while (resume);

  return newActions;
}

function advanceRecurringAction(recurringAction, internalCall) {
  console.log(`${getFuncName()}...`);

  if (recurringAction.recurringDays == -1) {
    recurringAction.nextRunDate = "";
    return;
  }
  const originalNextRunDate = recurringAction.nextRunDate;
  if (recurringAction.dayOccuranceInMonth)
    recurringAction.nextRunDate.setDate(1); // reset the day in case the next month has more days

  if (recurringAction.recurringMonths || recurringAction.dayOccuranceInMonth)
    recurringAction.nextRunDate.setMonth(originalNextRunDate.getMonth() + (recurringAction.recurringMonths || 1));

  if (recurringAction.recurringDays || (!recurringAction.recurringMonths && !recurringAction.dayOccuranceInMonth))
    recurringAction.nextRunDate.setDate(originalNextRunDate.getDate() + (recurringAction.recurringDays || 1));

  if (recurringAction.dayOccuranceInMonth) {
    recurringAction.nextRunDate = getMonthDayOccurance(recurringAction.nextRunDate, recurringAction.daysOfWeek, recurringAction.dayOccuranceInMonth[0])
  }
  else {
    if (internalCall)
      return;

    const savedNextRunDate = new Date(recurringAction.nextRunDate);

    // make sure the date is valid
    let attemptsLeft = 20;
    // Increment by 1 day until a valid date is found
    while (!isValidDate(recurringAction, recurringAction.nextRunDate)) {
      advanceRecurringAction(recurringAction, true);
      if (attemptsLeft-- <= 0) {
      //  // circuit breaker - to many date checks, give up
        recurringAction.nextRunDate = savedNextRunDate;
        break;
      }
    }
  }

  if (recurringAction.endDate && recurringAction.nextRunDate > recurringAction.endDate) {
    recurringAction.nextRunDate = ""; // the action has ended
  }
}

function handleAction(recurringAction) {
  console.log(`${getFuncName()} - ${recurringAction.name}...`);

  if (recurringAction.action === "Add days") {
    addDaysToRecurringTask(recurringAction.name, recurringAction.parameters ?? 1);
  }
}
