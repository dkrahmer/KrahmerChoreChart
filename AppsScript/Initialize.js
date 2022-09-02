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

/************************************************************************************************************
Release History
  2022-05-01 - v0.0.0: Inception
  2022-09-03 - v1.0.0: Initial public release
************************************************************************************************************/

// ------------------ Begin Settings ------------------

// The number of "Assign to Next" columns
// Set this number before adding additional "Assign to Next" columns to the "Recurring Tasks" tab.
const ASSIGNED_TO_NEXT_COUNT = 2;

const TASK_DUE_DATE_FORMAT = 'mm"/"dd';
const TASK_COMPLETED_DATE_FORMAT = `${TASK_DUE_DATE_FORMAT} hh":"mm" "am/pm`;
const COMPLETED_TASK_DUE_DATE_FORMAT = 'mm"/"dd"/"yyyy (ddd)';
const COMPLETED_TASK_COMPLETED_DATE_FORMAT = `${COMPLETED_TASK_DUE_DATE_FORMAT} hh":"mm" "am/pm`;

const DAY_ABBREVIATIONS = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];

const TASKS_SHEET_NAME = "Tasks";
const RECURRING_TASKS_SHEET_NAME = "Recurring Tasks";
const ARCHIVED_TASKS_SHEET_NAME = "Completed Tasks";
const ASSIGNEES_SHEET_NAME = "Assignees";

// ------------------ End Settings ------------------


var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

// rows and column indexes are 1 based (there is no 0)
let TASKS_COL_COUNT = 0;
const TASKS_COL_NAME = ++TASKS_COL_COUNT;
const TASKS_COL_ASSIGNED_TO = ++TASKS_COL_COUNT;
const TASKS_COL_DUE_DATE = ++TASKS_COL_COUNT;
const TASKS_COL_COMPLETED = ++TASKS_COL_COUNT;
const TASKS_COL_COMPLETED_DATE = ++TASKS_COL_COUNT;

let RECUR_TASKS_COL_COUNT = 0;
const RECUR_TASKS_COL_ENABLED = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_COL_DAY_NEED_NOW = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_COL_NAME = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_COL_ASSIGNED_TO = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_COL_ASSIGNED_TO_NEXT = (RECUR_TASKS_COL_COUNT += ASSIGNED_TO_NEXT_COUNT);
const RECUR_TASKS_COL_NEXT_DUE_DATE = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_COL_CREATE_DAYS_BEFORE_DUE = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_COL_RECURRING_DAYS = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_COL_RECURRING_MONTHS = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_COL_DAYS_OF_WEEK = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_COL_DAY_OCCURANCE_IN_MONTH = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_COL_END_DATE = ++RECUR_TASKS_COL_COUNT;
const RECUR_TASKS_LAST_DUE_DATE = ++RECUR_TASKS_COL_COUNT;

let ASSIGNEES_COL_COUNT = 0;
const ASSIGNEES_COL_NAME = ++ASSIGNEES_COL_COUNT;
const ASSIGNEES_COL_PROPER_NAME = ++ASSIGNEES_COL_COUNT;
const ASSIGNEES_COL_EMAIL = ++ASSIGNEES_COL_COUNT;
const ASSIGNEES_COL_END_OF_DAY_EMAIL_TYPE = ++ASSIGNEES_COL_COUNT;

function initialize() {
  saveSheetKey();
  deleteTasks();
  deleteArchive();
  createTriggers();
}

function reinitialize() {
  SCRIPT_PROP.deleteAllProperties();

  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  initialize();
}

function saveSheetKey() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  SCRIPT_PROP.setProperty("key", sheet.getId());
}

function createTriggers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let scriptName;

  const triggerFunctionNames = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  
  scriptName = "onEditCustom";
  if (!triggerFunctionNames.includes(scriptName)) {
    ScriptApp.newTrigger(scriptName)
      .forSpreadsheet(sheet).onEdit()
      .create();
    console.log(`Scheduled ${scriptName}`);
  }

  scriptName = "onChangeCustom";
  if (!triggerFunctionNames.includes(scriptName)) {
    ScriptApp.newTrigger(scriptName)
      .forSpreadsheet(sheet).onChange()
      .create();
    console.log(`Scheduled ${scriptName}`);
  }

  scriptName = "endOfDayEmailReport";
  if (!triggerFunctionNames.includes(scriptName)) {
    ScriptApp.newTrigger(scriptName)
      .timeBased().everyDays(1).atHour(20).nearMinute(15)
      .create();
    console.log(`Scheduled ${scriptName}`);
  }

  scriptName = "archiveCompletedTasks";
  if (!triggerFunctionNames.includes(scriptName)) {
    ScriptApp.newTrigger(scriptName)
      .timeBased().everyDays(1).atHour(1).nearMinute(0)
      .create();
    console.log(`Scheduled ${scriptName}`);
  }
  
  scriptName = "processRecurringTasks";
  if (!triggerFunctionNames.includes(scriptName)) {
    ScriptApp.newTrigger(scriptName)
      .timeBased().everyDays(1).atHour(1).nearMinute(15)
      .create();
    console.log(`Scheduled ${scriptName}`);
  }
}

function deleteTasks() {
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));

  const tasksSheet = sheet.getSheetByName(TASKS_SHEET_NAME);
  deleteSheetData(tasksSheet);
}

function deleteArchive() {
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));

  const archivedTasksSheet = sheet.getSheetByName(ARCHIVED_TASKS_SHEET_NAME);
  deleteSheetData(archivedTasksSheet);
}
