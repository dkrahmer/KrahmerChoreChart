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

function enableDemoMode() {
  setDemoMode(true);
}

function disableDemoMode() {
  setDemoMode(false);
}

function setDemoMode(enable) {
  const triggerFunctionNames = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  
  scriptName = "markSingleTaskCompleteDemo";
  if (triggerFunctionNames.includes(scriptName) !== enable) {
    if (enable) {
      ScriptApp.newTrigger(scriptName)
        .timeBased().everyHours(1)
        .create();
      console.log(`Scheduled ${scriptName}`);
    }
    else {
      deleteTriggerByScriptName(scriptName);
    }
  }
  
  scriptName = "markAllTasksComplete";
  if (triggerFunctionNames.includes(scriptName) !== enable) {
    if (enable) {
      ScriptApp.newTrigger(scriptName)
        .timeBased().everyDays(1).atHour(19)
        .create();
      console.log(`Scheduled ${scriptName}`);
    }
    else {
      deleteTriggerByScriptName(scriptName);
    }
  }
}

function markSingleTaskCompleteDemo() {
  const hour = new Date().getHours();
  if (hour > 19 || hour < 5)
    return; // do not mark tasks complete during this hour range

  markTasksComplete(1);
}
