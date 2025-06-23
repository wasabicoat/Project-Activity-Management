// Ensure you have these constants at the top of your Code.gs file
const SHEET_ID = '1Cc-I45ehhpLB86dp-qBeOJh0aSj5V_0OOqyY3DG_1ZQ';
const SHEET_NAME = 'action listlog';
const ORDER_SHEET_NAME = 'workstream_order';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Project Activity Management');
}

function getTaskData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues().slice(1); // skip header

  const result = {};
  data.forEach((row, i) => {
    const task = row[1];       // Column B
    const workstream = row[2]; // Column C
    const owner = row[3];      // Column D
    const duedate = formatDate(row[4]); // Column E
    const status = row[5] ? row[5].toString().trim() : 'Not Start'; // Column F
    const remark = row[8];     // Column I

    if (!workstream) return;
    if (!result[workstream]) result[workstream] = [];
    result[workstream].push({ task, workstream, owner, duedate, status, remark, rowIndex: i + 2 });
  });

  const orderSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ORDER_SHEET_NAME);
  let order = Object.keys(result);
  if (orderSheet) {
    const values = orderSheet.getRange(1, 1, orderSheet.getLastRow(), 1).getValues().flat();
    const existing = new Set(order);
    order = values.filter(v => existing.has(v));
    const missing = [...existing].filter(v => !values.includes(v));
    order = [...order, ...missing];
  }

  return { tasks: result, order };
}

function formatDate(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return d ? d.toString().trim() : '';
}

function updateTask(task) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const row = Number(task.rowIndex);
  if (!row) return;
  sheet.getRange(row, 2).setValue(task.task);
  sheet.getRange(row, 3).setValue(task.workstream);
  sheet.getRange(row, 4).setValue(task.owner);
  sheet.getRange(row, 5).setValue(task.duedate);
  sheet.getRange(row, 6).setValue(task.status);
  sheet.getRange(row, 9).setValue(task.remark);
}

/**
 * --- UPDATED FUNCTION ---
 * Adds a new task to the sheet with an index in column A
 * and returns the new task object to the frontend.
 * @param {Object} task - The task object from the frontend.
 * @returns {Object} The complete task object with its new rowIndex.
 */
function addTask(task) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  
  // Get the last row number to use as the index for the new task.
  // This assumes row 1 is a header.
  const newIndex = sheet.getLastRow();
  
  // Append the new row with the calculated index in column A.
  sheet.appendRow([newIndex, task.task, task.workstream, task.owner, task.duedate, task.status, '', '', task.remark]);
  
  // Get the row number of the task we just added.
  const newRowIndex = sheet.getLastRow();
  task.rowIndex = newRowIndex;
  
  // Return the complete task object so the frontend can add it to the UI without a full refresh.
  return task;
}


function renameWorkstream(oldName, newName) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === oldName) {
      sheet.getRange(i + 1, 3).setValue(newName);
    }
  }
  const orderSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ORDER_SHEET_NAME);
  if (orderSheet) {
    const values = orderSheet.getRange(1, 1, orderSheet.getLastRow(), 1).getValues().flat();
    const index = values.indexOf(oldName);
    if (index !== -1) {
      orderSheet.getRange(index + 1, 1).setValue(newName);
    }
  }
}

function moveWorkstream(workstream, direction) {
  const sheet = SpreadsheetApp.openById(SHEET_ID);
  let orderSheet = sheet.getSheetByName(ORDER_SHEET_NAME);
  if (!orderSheet) {
    orderSheet = sheet.insertSheet(ORDER_SHEET_NAME);
    // If we create it, populate it with the current order
    const taskData = getTaskData();
    const initialOrder = taskData.order;
    orderSheet.getRange(1, 1, initialOrder.length, 1).setValues(initialOrder.map(v => [v]));
  }
  const values = orderSheet.getRange(1, 1, orderSheet.getLastRow(), 1).getValues().flat();
  const index = values.indexOf(workstream);
  if (index === -1) return;
  if (direction === 'up' && index > 0) {
    [values[index - 1], values[index]] = [values[index], values[index - 1]];
  } else if (direction === 'down' && index < values.length - 1) {
    [values[index + 1], values[index]] = [values[index], values[index + 1]];
  }
  orderSheet.getRange(1, 1, values.length, 1).setValues(values.map(v => [v]));
}


function moveTask(rowIndex, direction) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const index = rowIndex - 1; // Convert 1-based rowIndex to 0-based array index
  
  if (index < 1 || index >= data.length) return; // Check bounds

  const taskRow = data[index];
  const workstream = taskRow[2];

  let adjacentIndex = -1;
  if (direction === 'up') {
    // Search upwards for the first row with the same workstream
    for (let i = index - 1; i >= 1; i--) {
      if (data[i][2] === workstream) {
        adjacentIndex = i;
        break;
      }
    }
  } else if (direction === 'down') {
    // Search downwards for the first row with the same workstream
    for (let i = index + 1; i < data.length; i++) {
      if (data[i][2] === workstream) {
        adjacentIndex = i;
        break;
      }
    }
  }

  if (adjacentIndex === -1) return; // No task to swap with

  // Get values of the two rows to swap
  const range1 = sheet.getRange(index + 1, 1, 1, data[0].length);
  const range2 = sheet.getRange(adjacentIndex + 1, 1, 1, data[0].length);
  const values1 = range1.getValues();
  const values2 = range2.getValues();

  // Swap the values
  range1.setValues(values2);
  range2.setValues(values1);
}
