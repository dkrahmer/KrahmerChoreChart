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

function sortArchivedTasksSheet() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const archivedTasksSheet = sheet.getSheetByName(ARCHIVED_TASKS_SHEET_NAME);

  archivedTasksSheet
    .sort(TASKS_COL_ASSIGNED_TO, true)
    .sort(TASKS_COL_DUE_DATE, false); // sort reverse of tasks
}

function archiveCompletedTasks() {
  console.log(`${getFuncName()}...`);
  archiveCompletedTasksInternal();
}

function archiveCompletedTasksInternal(all) {
  console.log(`${getFuncName()}...`);
  sortTasksSheet();
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const tasksSheet = sheet.getSheetByName(TASKS_SHEET_NAME);
  const archivedTasksSheet = sheet.getSheetByName(ARCHIVED_TASKS_SHEET_NAME);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let archiveCount = 0;
  const topArchiveRow = 2;
  const rowCount = tasksSheet.getMaxRows();
  for (let row = 2; row <= rowCount; row++) {
    const range = tasksSheet.getRange(row, 1, 1, TASKS_COL_COUNT);
    const taskData = range.getValues()[0];
    const taskName = taskData[TASKS_COL_NAME - 1];
    const completed = taskData[TASKS_COL_COMPLETED - 1];
    const dueDate = taskData[TASKS_COL_DUE_DATE - 1];

    if (!taskName)
      continue; // blank row

    if (!dueDate)
      continue; // do not archive if no due date

    if (!completed)
      continue; // do not archive if not completed

    if (dueDate >= today && !all)
      continue; // do not archive if today or newer

    console.log(`Archiving completed task [${taskName}] - due: [${dueDate}]`);
    // tasksSheet.hideRows(row);
    addTaskRow(sheet, archivedTasksSheet, topArchiveRow, taskData, COMPLETED_TASK_DUE_DATE_FORMAT, COMPLETED_TASK_COMPLETED_DATE_FORMAT);
    range.clearContent();
    archiveCount++;
  }

  if (archiveCount > 0) {
    // delete all except 1, even if empty
    let keepRowCount = 0;
    const startRow = 2;
  
    // iterate in reverse since we are deleting rows
    for (let row = rowCount; row >= startRow; row--) {
      if (row <= startRow && keepRowCount == 0)
        break; // no rows have been kept yet so keep this last row

      const name = tasksSheet.getRange(row, TASKS_COL_NAME).getValue();
      if (name) {
        keepRowCount++;
        continue;
      }
      tasksSheet.deleteRow(row);
    }

    sortArchivedTasksSheet();
  }
}