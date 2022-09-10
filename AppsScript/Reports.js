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

function startOfDayEmailReport() {
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const tasks = getTasks();
  const incompleteTasks = tasks
    .filter(t => !t.completed);

  const sendToAssignees = getSendToAssignees(incompleteTasks, ASSIGNEES_COL_START_OF_DAY_EMAIL_TYPE);
  sendToAssignees.forEach(assignee => {
    let subject, message = `Hi ${assignee.properName},`;

    personalizedIncompleteTasks = incompleteTasks.filter(t => isAssigned(t.assignedTo, assignee.name));
    const personalizedIncompleteLateTasks = personalizedIncompleteTasks.filter(t => t.when === "late");
    const personalizedIncompleteCurrentTasks = personalizedIncompleteTasks.filter(t => t.when === "current");
    const personalizedIncompleteFutureTasks = personalizedIncompleteTasks.filter(t => t.when === "future");

    if (assignee.personalized) {
      if (personalizedIncompleteTasks.length) {
        if (personalizedIncompleteCurrentTasks.length) {
          subject = `You have ${personalizedIncompleteCurrentTasks.length} chore${personalizedIncompleteCurrentTasks.length === 1 ? "" : "s"} due today`;
          if (personalizedIncompleteLateTasks.length) {
            subject += ` and ${personalizedIncompleteLateTasks.length} late chore${personalizedIncompleteLateTasks.length === 1 ? "" : "s"}`;
          }
        }
        else if (personalizedIncompleteLateTasks.length) {
          subject = `You have ${personalizedIncompleteLateTasks.length} late chore${personalizedIncompleteLateTasks.length === 1 ? "" : "s"}`;
        }
        else if (personalizedIncompleteFutureTasks.length) {
          subject = `You have ${personalizedIncompleteFutureTasks.length} upcoming chore${personalizedIncompleteFutureTasks.length === 1 ? "" : "s"}`;
        }

        message += `\n\n${subject}.`;

        if (personalizedIncompleteLateTasks.length)
          message += "\n\nLate:\n" + getAssignedTasksText(personalizedIncompleteLateTasks, true);

        if (personalizedIncompleteCurrentTasks.length)
          message += "\n\nCurrent:\n" + getAssignedTasksText(personalizedIncompleteCurrentTasks);

        if (personalizedIncompleteFutureTasks.length)
          message += "\n\nUpcoming:\n" + getAssignedTasksText(personalizedIncompleteFutureTasks, true);
      }
      else {
        subject = "You have no chores due today!";
        message += `\n\n${subject}.`;
      }
    }
    else {
      if (incompleteTasks.length) {
        const incompleteLateTasks = incompleteTasks.filter(t => t.when === "late");
        const incompleteCurrentTasks = incompleteTasks.filter(t => t.when === "current");
        const incompleteFutureTasks = incompleteTasks.filter(t => t.when === "future");
        const personalizedCount = personalizedIncompleteCurrentTasks.length + personalizedIncompleteLateTasks.length;

        if (incompleteTasks.length) {
          subject = `${incompleteCurrentTasks.length} chore${incompleteCurrentTasks.length === 1 ? "" : "s"} ${incompleteCurrentTasks.length === 1 ? "is" : "are"} due today`;
          if (incompleteLateTasks.length) {
            subject += ` and ${incompleteLateTasks.length} late chore${incompleteLateTasks.length === 1 ? "" : "s"}`;
            subject += ` (${personalizedCount} ${personalizedCount === 1 ? "is" : "are"} yours)`;
          }
          else if (incompleteLateTasks.length) {
            subject += `${incompleteLateTasks.length} late chore${incompleteLateTasks.length === 1 ? "" : "s"}`;
          }
          else if (incompleteFutureTasks.length) {
            subject += `${incompleteFutureTasks.length} upcoming chore${incompleteFutureTasks.length === 1 ? "" : "s"}`;
          }

          message += `\n\n${subject}.`;
          if (incompleteLateTasks.length)
            message += `\n\nLate (${incompleteLateTasks.length}):\n` + getAssignedTasksText(incompleteLateTasks, true);

          if (incompleteCurrentTasks.length)
            message += `\n\nCurrent (${incompleteCurrentTasks.length}):\n` + getAssignedTasksText(incompleteCurrentTasks);

          if (incompleteFutureTasks.length)
            message += `\n\nUpcoming (${incompleteFutureTasks.length}):\n` + getAssignedTasksText(incompleteFutureTasks, true);
        }
        else {
          subject = "No chores today!";
          message += `\n\n${subject}.`
        }
      }
    }
    
    message += `\n\nChore chart: ${sheet.getUrl()}`;
    console.log(`Sending email to ${assignee.email}: ${subject}`)
    MailApp.sendEmail(assignee.email, subject, message, { name: EMAIL_REPORT_FROM_NAME });
  });
}

function endOfDayEmailReport() {
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const currentTasks = getTasks();
  const completeTasks = currentTasks.filter(t => t.completed);
  const incompleteTasks = currentTasks.filter(t => !t.completed).filter(t => t.when !== "future");

  const sendToAssignees = getSendToAssignees(incompleteTasks, ASSIGNEES_COL_END_OF_DAY_EMAIL_TYPE);

  sendToAssignees.forEach(assignee => {
    let subject, message = `Hi ${assignee.properName},`;

    personalizedIncompleteTasks = incompleteTasks.filter(t => isAssigned(t.assignedTo, assignee.name))
    personalizedCompleteTasks = completeTasks.filter(t => isAssigned(t.assignedTo, assignee.name))

    if (assignee.personalized) {
      if (personalizedIncompleteTasks.length) {
        subject = `${personalizedIncompleteTasks.length} of your chores are incomplete`;
        message += `\n\n${subject}:\n` + getAssignedTasksText(personalizedIncompleteTasks, true);
      }
      else {
        subject = `All of your chores are complete! (${personalizedCompleteTasks.length})`;
        message += `\n\n${subject}`
      }
    
      if (personalizedCompleteTasks.length) {
        message += `\n\n${personalizedCompleteTasks.length} of your chores are complete:\n`
          + getAssignedTasksText(personalizedCompleteTasks);
      }
    }
    else {
      if (incompleteTasks.length) {
        subject = `${incompleteTasks.length} incomplete chore${incompleteTasks.length === 1 ? "" : "s"} (${personalizedIncompleteTasks.length} ${personalizedIncompleteTasks.length === 1 ? "is" : "are"} yours)`;

        message += `\n\n${subject}:\n` + getAssignedTasksText(incompleteTasks, true);
      }
      else {
        subject = `All chores are complete! (${completeTasks.length})`;
        message += `\n\n${subject}`
      }
    
      if (completeTasks.length) {
        message += `\n\n${completeTasks.length} completed chore${completeTasks.length === 1 ? "" : "s"} (${personalizedCompleteTasks.length} ${personalizedCompleteTasks.length === 1 ? "is" : "are"} yours):\n`
          + getAssignedTasksText(completeTasks);
      }
    }
    
    message += `\n\nChore chart: ${sheet.getUrl()}`;
    console.log(`Sending email to ${assignee.email}: ${subject}`)
    MailApp.sendEmail(assignee.email, subject, message, { name: EMAIL_REPORT_FROM_NAME });
  });
}

function getTasks() {
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const tasksSheet = sheet.getSheetByName(TASKS_SHEET_NAME);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tasks = [];

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

    const task = {
        taskName,
        assignedTo: taskData[TASKS_COL_ASSIGNED_TO - 1],
        dueDate,
        completed: !!completed,
        when: dueDate < today ? "late" : (dueDate > today ? "future" : "current"),
      };
      
    tasks.push(task);
  }

  return tasks.sort((a, b) => a.assignedTo >= b.assignedTo ? 1 : -1);
}

function getAssignedTasksText(tasks, includeDate) {
  return tasks
    .map((t, index) => `  ${index + 1}. ${t.assignedTo}: ${t.taskName}${(includeDate ? " (due " + Utilities.formatDate(t.dueDate, Session.getScriptTimeZone(), EMAIL_TASK_DUE_DATE_FORMAT) + ")" : "")}`)
    .join("\n");
}

function getSendToAssignees(incompleteTasks, emailTypeColumn) {
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const assigneesSheet = sheet.getSheetByName(ASSIGNEES_SHEET_NAME);
  const assignees = assigneesSheet.getRange(2, 1, assigneesSheet.getMaxRows() - 1, ASSIGNEES_COL_COUNT).getValues();

  return assignees
    .filter(assignee => 
      !!assignee[ASSIGNEES_COL_EMAIL - 1]
      && !assignee[ASSIGNEES_COL_EMAIL - 1].includes("?")
      && (
        (assignee[emailTypeColumn - 1] === "Always")
        || (assignee[emailTypeColumn - 1] === "Personalized - Always")
        || (incompleteTasks.length
          && (assignee[emailTypeColumn - 1] === "If any incomplete")
          || (
            (
              (assignee[emailTypeColumn - 1] === "If own incomplete")
              || (assignee[emailTypeColumn - 1] === "Personalized - If own incomplete")
            )
            && incompleteTasks.some(task => isAssigned(task.assignedTo, assignee[ASSIGNEES_COL_NAME - 1]))
          )
        )
      )
    )
    .map(assignee => {
      return {
        name: assignee[ASSIGNEES_COL_NAME - 1],
        properName: assignee[ASSIGNEES_COL_PROPER_NAME - 1],
        email: assignee[ASSIGNEES_COL_EMAIL - 1],
        personalized: !!assignee[emailTypeColumn - 1].includes("Personalized")
      };
    })
    .sort((a, b) => a.name >= b.name ? 1 : -1);
}

function isAssigned(taskAssignee, name) {
  return taskAssignee === name || taskAssignee === "all" || taskAssignee === "anybody";
}
