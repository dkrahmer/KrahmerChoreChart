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

function sortTasksSheet() {
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const tasksSheet = sheet.getSheetByName(TASKS_SHEET_NAME);

  tasksSheet
    .sort(TASKS_COL_ASSIGNED_TO, true)
    //.sort(TASKS_COL_COMPLETED_DATE, false)
    //.sort(TASKS_COL_COMPLETED, true)
    .sort(TASKS_COL_DUE_DATE, true);
}

function addTask(sheet, task) {
  console.log(`addTask - ${task.name}`);
  const tasksSheet = sheet.getSheetByName(TASKS_SHEET_NAME);
  const row = 2;

  const taskValues = [
    task.name,
    task.assignedToArray[0],
    task.dueDate,
    task.completed ?? false,
    task.completedDate ?? ""
  ];

  addTaskRow(sheet, tasksSheet, row, taskValues)
}

function addTaskRow(sheet, targetSheet, row, taskValues, dueDateFormat, completedDateFormat) {
  dueDateFormat = dueDateFormat ?? TASK_DUE_DATE_FORMAT;
  completedDateFormat = completedDateFormat ?? TASK_COMPLETED_DATE_FORMAT;

  targetSheet.insertRowBefore(row);
  const range = targetSheet.getRange(row, 1, 1, TASKS_COL_COUNT);

  range.setValues([taskValues]);

  const checkboxCell = targetSheet.getRange(row, TASKS_COL_COMPLETED, 1, 1);
  checkboxCell.insertCheckboxes();

  const completedDateCell = targetSheet.getRange(row, TASKS_COL_COMPLETED_DATE, 1, 1);
  completedDateCell.setNumberFormat(completedDateFormat);

  const dueDateCell = targetSheet.getRange(row, TASKS_COL_DUE_DATE, 1, 1);
  dueDateCell.setNumberFormat(dueDateFormat);

  const assigneesSheet = sheet.getSheetByName(ASSIGNEES_SHEET_NAME);
  const ruleRange = assigneesSheet.getRange('A2:A');
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(ruleRange).build();
  const assignedToCell = targetSheet.getRange(row, TASKS_COL_ASSIGNED_TO, 1, 1);
  assignedToCell.setDataValidation(rule);
}

function validateTasksSheet() {
  console.log(`validateTasksSheet`);

  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const tasksSheet = sheet.getSheetByName(TASKS_SHEET_NAME);

  const dueCells = tasksSheet.getRange(2, TASKS_COL_DUE_DATE, tasksSheet.getMaxRows() - 1, TASKS_COL_COMPLETED_DATE - TASKS_COL_DUE_DATE + 1);
  const dueValues = dueCells.getValues();
  const now = new Date();
  for (let rowIndex = 0; rowIndex < dueValues.length; rowIndex++) {
    const rowValues = dueValues[rowIndex];
    const dueDate = rowValues[0];
    const completed = rowValues[1];
    const completedDate = rowValues[2];

    if ((completed || !!completedDate) && dueDate > now) {  
      console.log(`Task completed before due date`);
      const completedDateCell = tasksSheet.getRange(rowIndex + 2, TASKS_COL_COMPLETED_DATE);
      const completedCell = tasksSheet.getRange(rowIndex + 2, TASKS_COL_COMPLETED);
      completedCell.setValue(false);
      completedDateCell.setValue("");
    }
    else if (completed && !completedDate) {
      console.log(`Completed checkbox checked - adding completed date`);
      const completedDateCell = tasksSheet.getRange(rowIndex + 2, TASKS_COL_COMPLETED_DATE);
      completedDateCell.setValue(new Date());
      completedDateCell.setNumberFormat(TASK_COMPLETED_DATE_FORMAT); // enforce date format after any changes
    }
    else if (!completed && !!completedDate) {
      console.log(`Completed checkbox unchecked - removing completed date`);
      const completedDateCell = tasksSheet.getRange(rowIndex + 2, TASKS_COL_COMPLETED_DATE);
      completedDateCell.setValue("");
    }
  }
}

function cleanTasks(sheet) {
  sheet = sheet ?? SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const tasksSheet = sheet.getSheetByName(TASKS_SHEET_NAME);
  if (tasksSheet.getMaxRows() <= 2)
    return;

  const values = tasksSheet.getRange(2, TASKS_COL_NAME, tasksSheet.getMaxRows() - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (!values[i][0]) {
      tasksSheet.deleteRow(i + 2);
    }
  }
}

function markAllTasksComplete() {
  markTasksComplete();
}

function markTasksComplete(maxCount) {
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const tasksSheet = sheet.getSheetByName(TASKS_SHEET_NAME);
  value = true;

  const range = tasksSheet.getRange(2, TASKS_COL_COMPLETED, tasksSheet.getMaxRows() - 1, 2);
  const values = range.getValues();
  maxCount = maxCount ?? values.length;
  let count = 0;
  for (let i = 0; i < values.length && count < maxCount; i++) {
    const rowValues = values[i];
    if (rowValues[0] !== value) {
      rowValues[0] = value;
      rowValues[1] = value ? new Date() : "";
      count++;
    }
  }

  range.setValues(values);
  validateTasksSheet();
}
