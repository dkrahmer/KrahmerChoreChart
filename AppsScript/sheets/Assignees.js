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

function processAssigneeColors() {
  console.log(`${getFuncName()}...`);
  console.log(`processAssigneeColors`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));

  const tasksSheet = sheet.getSheetByName(TASKS_SHEET_NAME);
  const tasksRulesRange = tasksSheet.getRange(2, TASKS_COL_ASSIGNED_TO, tasksSheet.getMaxRows() - 1, 1);
  const tasksRules = tasksSheet.getConditionalFormatRules();

  const archivedSheet = sheet.getSheetByName(ARCHIVED_TASKS_SHEET_NAME);
  const archivedRulesRange = archivedSheet.getRange(2, TASKS_COL_ASSIGNED_TO, archivedSheet.getMaxRows() - 1, 1);

  const recurringSheet = sheet.getSheetByName(RECURRING_TASKS_SHEET_NAME);
  const recurringRulesRange = recurringSheet.getRange(2, RECUR_TASKS_COL_ASSIGNED_TO, recurringSheet.getMaxRows() - 1, RECUR_TASKS_COL_ASSIGNED_TO_NEXT - RECUR_TASKS_COL_ASSIGNED_TO + 1);

  const newTasksRules = [];
  const newArchivedRules = [];
  const newRecurringRules = [];
  const assigneesSheet = sheet.getSheetByName(ASSIGNEES_SHEET_NAME);
  const assigneesRange = assigneesSheet.getRange(2, ASSIGNEES_COL_NAME, tasksSheet.getMaxRows() - 1, 1);
  const names = assigneesRange.getValues();
  const foregroundColorObjects = assigneesRange.getFontColorObjects();
  const backgroundObjects = assigneesRange.getBackgroundObjects();

  for (let i = 0; i < names.length; i++) {
    const name = names[i][0];
    if (!name)
      continue;

    const tasksCellTarget = `${String.fromCharCode("A".charCodeAt(0) + TASKS_COL_ASSIGNED_TO - 1)}2`;

    const newTasksRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([tasksRulesRange])
      .whenFormulaSatisfied(`=${tasksCellTarget} = "${name}"`)
      .setFontColorObject(foregroundColorObjects[i][0])
      .setBackgroundObject(backgroundObjects[i][0])
      .build();
    newTasksRules.push(newTasksRule);

    const newArchivedRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([archivedRulesRange])
      .whenFormulaSatisfied(`=${tasksCellTarget} = "${name}"`)
      .setFontColorObject(foregroundColorObjects[i][0])
      .setBackgroundObject(backgroundObjects[i][0])
      .build();
    newArchivedRules.push(newArchivedRule);

    const recurringCellTarget = `${String.fromCharCode("A".charCodeAt(0) + RECUR_TASKS_COL_ASSIGNED_TO - 1)}2`;
    
    const newRecurringRule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([recurringRulesRange])
      .whenFormulaSatisfied(`=${recurringCellTarget} = "${name}"`)
      .setFontColorObject(foregroundColorObjects[i][0])
      .setBackgroundObject(backgroundObjects[i][0])
      .build();
    newRecurringRules.push(newRecurringRule);
  }

  const keepStartRules = 2;
  const keepEndRules = 2;
  const tasksReplacementRules = [...(tasksRules.slice(0, keepStartRules)), ...newTasksRules, ...(tasksRules.slice(-keepEndRules))];
  tasksSheet.setConditionalFormatRules(tasksReplacementRules);

  archivedSheet.setConditionalFormatRules(newArchivedRules);

  recurringSheet.setConditionalFormatRules(newRecurringRules);
}
