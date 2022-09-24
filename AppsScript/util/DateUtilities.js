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

function getMonthDayOccurance(startDate, daysOfWeek, occurance) {
  if (occurance != "L" && occurance < 1 || occurance > 4)
    return startDate; // invalid occurance

  const workDate = new Date(startDate);
  workDate.setDate(1);
  if (occurance == "L") {
    // Get the last day of the month of the start date
    workDate.setMonth(workDate.getMonth() + 1);
    workDate.setDate(0);
    const savedDate = new Date(workDate);
    // keep going back until we find a valid day
    let maxAttempts = 7;
    while (!isDateInDaysOfWeekList(workDate, daysOfWeek)) {
      if (maxAttempts-- <= 0)
        return savedDate; // short circuit if not found within maxAttempts

      workDate.setDate(workDate.getDate() - 1);
    }
    return workDate;
  }

  const savedDate = new Date(workDate);
  // keep going forward until we find the first valid day
  let maxAttempts = 31;
  while (!isDateInDaysOfWeekList(workDate, daysOfWeek)) {
    if (maxAttempts-- <= 0)
      return savedDate; // short circuit if not found within maxAttempts

    workDate.setDate(workDate.getDate() + 1);
  }

  // Add weeks to find the desired occurance
  workDate.setDate(workDate.getDate() + (7 * (occurance - 1)));

  return workDate;
}

function getNextRunDate(recurringTask, propertyName) {
  if (!recurringTask[propertyName])
    return recurringTask[propertyName];

  let runAdjustDays = recurringTask.runAdjustDays;
  if (runAdjustDays)
    runAdjustDays = -runAdjustDays;

  const nextRunDate = new Date(recurringTask[propertyName]);
  nextRunDate.setDate(nextRunDate.getDate() - (runAdjustDays ?? recurringTask.createDaysBeforeDue ?? 0));
  nextRunDate.setHours(0, 0, 0, 0);
  return nextRunDate;
}

function isValidDate(recurringTask, date) {
  if (!isDateInDaysOfWeekList(date, recurringTask.daysOfWeek))
    return false;

  return true;
}

function isDateInDaysOfWeekList(date, daysOfWeekList) {
  if (!daysOfWeekList)
    return true;

  const dayName = DAY_ABBREVIATIONS[date.getDay()];
  return daysOfWeekList.includes(dayName);
}
