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

function endOfDayEmailReport() {
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const tasksSheet = sheet.getSheetByName(TASKS_SHEET_NAME);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const incompleteChores = [];
  const completedChores = [];
  const rowCount = tasksSheet.getMaxRows();
  for (let row = 2; row <= rowCount; row++) {
    const range = tasksSheet.getRange(row, 1, 1, TASKS_COL_COUNT);
    const taskData = range.getValues()[0];
    const taskName = taskData[TASKS_COL_NAME - 1];
    const completed = taskData[TASKS_COL_COMPLETED - 1];
    const dueDate = taskData[TASKS_COL_DUE_DATE - 1];

    if (!taskName)
      continue;

    if (!dueDate)
      continue;

    const chore = {
        taskName,
        assignedTo: taskData[TASKS_COL_ASSIGNED_TO - 1],
        dueDate
      };

    if (completed) {
      completedChores.push(chore);
      continue;
    }
    
    if (dueDate > today)
      continue;

    incompleteChores.push(chore);
  }

  const assigneesSheet = sheet.getSheetByName(ASSIGNEES_SHEET_NAME);
  const assignees = assigneesSheet.getRange(2, 1, assigneesSheet.getMaxRows() - 1, ASSIGNEES_COL_COUNT).getValues();

  const sendToAssignees = assignees
    .filter(assignee => 
      !!assignee[ASSIGNEES_COL_EMAIL - 1] 
        && (assignee[ASSIGNEES_COL_END_OF_DAY_EMAIL_TYPE - 1] === "Always")
          || (incompleteChores.length && assignee[ASSIGNEES_COL_END_OF_DAY_EMAIL_TYPE - 1] === "If any incomplete")
          || (incompleteChores.length && assignee[ASSIGNEES_COL_END_OF_DAY_EMAIL_TYPE - 1] === "If own incomplete")
              && incompleteChores.some(chore => chore.assignedTo === assignee[ASSIGNEES_COL_NAME - 1]))
    .map(assignee => {
      return {
        name: assignee[ASSIGNEES_COL_NAME - 1],
        properName: assignee[ASSIGNEES_COL_PROPER_NAME - 1],
        email: assignee[ASSIGNEES_COL_EMAIL - 1]
      };
    });

  let subject, message = "";
  if (incompleteChores.length) {
    subject = `${incompleteChores.length} incomplete chores`;
    incompleteChores.sort((a, b) => a.assignedTo > b.assignedTo);
    message += `${incompleteChores.length} incomplete chores:\n`
      + incompleteChores
        .map((values, index) => `  ${index + 1}. ${values.assignedTo}: ${values.taskName}`)
        .join("\n")
      + "\n\n";
  }
  else {
    subject = `All chores complete! (${completedChores.length})`;
  }

  if (completedChores.length) {
    message += `${completedChores.length} completed chores:\n`
      + completedChores
        .map((values, index) => `  ${index + 1}. ${values.assignedTo}: ${values.taskName}`)
        .join("\n")
      + "\n\n";
  }

  message += `Chore chart: ${sheet.getUrl()}`;

  sendToAssignees.forEach((assignee) => {
    MailApp.sendEmail(assignee.email, subject, message);
  });
}
