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

function deleteSheetData(targetSheet, headerRows = 1) {
  if (targetSheet.getMaxRows() >= headerRows + 2)
    targetSheet.deleteRows(headerRows + 2, targetSheet.getMaxRows() - headerRows - 1);
  targetSheet.insertRowAfter(headerRows + 1);
  targetSheet.deleteRows(headerRows + 1, 1);
}

function getFuncName() {
   return getFuncName.caller.name
}

function deleteTriggerByScriptName(scriptName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    const trigger = triggers[i];
    if (trigger.getHandlerFunction() === scriptName)
      ScriptApp.deleteTrigger(trigger);
  }
}
