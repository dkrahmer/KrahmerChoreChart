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

function onEditCustom(e) { // launched from the script scheduler
  console.log(`${getFuncName()}...`);
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
  }
  catch (ex) {
    console.error(ex);
    return;
  }

  try {
    if (!e?.source) {
      console.error("Error: missing e.source");
      return;
    }

    const sourceSheetName = e.source.getSheetName();

    if (sourceSheetName == TASKS_SHEET_NAME) {
      validateTasksSheet();
      sortTasksSheet();
    }
    else if (sourceSheetName == RECURRING_TASKS_SHEET_NAME) {
      processRecurringTasks();
    }
    else if (sourceSheetName == RECURRING_ACTIONS_SHEET_NAME) {
      processRecurringActions();
    }
  }
  finally {
      SpreadsheetApp.flush();
      lock.releaseLock();
  }
}

function onChangeCustom(e) { // launched from the script scheduler
  console.log(`${getFuncName()}...`);
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
  }
  catch (ex) {
    console.error(ex);
    return;
  }

  try {
    if (!e?.source) {
      console.error(`Error: missing e.source`);
      return;
    }

    const sourceSheetName = e.source.getSheetName();

    if (sourceSheetName == ASSIGNEES_SHEET_NAME) {
      processAssigneeColors();
    }
  }
  finally {
      SpreadsheetApp.flush();
      lock.releaseLock();
  }
}
