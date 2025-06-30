const SHEET_ID = '1Cc-I45ehhpLB86dp-qBeOJh0aSj5V_0OOqyY3DG_1ZQ';
const SHEET_NAME = 'action listlog';
const ORDER_SHEET_NAME = 'workstream_order';

/**
 * Returns all task data and workstream order for frontend.
 */
function getTaskData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const orderSheet = ss.getSheetByName(ORDER_SHEET_NAME);

  // Read workstream order from order sheet (single column)
  let workstreamOrder = [];
  if (orderSheet) {
    workstreamOrder = orderSheet.getRange(1, 1, orderSheet.getLastRow(), 1).getValues().flat().filter(Boolean);
  }

  // Read all rows (including headers)
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tasksArr = data.slice(1);

  // Find the index of required fields
  const idx = {
    number: headers.indexOf('#'),
    actionItemName: headers.indexOf('Action Item Name'),
    category: headers.indexOf('Category'),
    responder: headers.indexOf('Responder'),
    duedate: headers.indexOf('Due Date'),
    status: headers.indexOf('Status'),
    remark: headers.indexOf('Remark'),
    reference: headers.indexOf('Reference'),
    reportttb: headers.indexOf('Report to ttb'),
  };

  // Build tasks as expected by frontend
  const tasksByWorkstream = {};
  tasksArr.forEach((row, i) => {
    // Sheet row index (for update), +2 due to header row (1) and 0-indexing
    const rowIndex = i + 2;
    const workstream = row[idx.category] || 'Uncategorized';
    const task = {
      number: row[idx.number], // <<-- Read from column A
      task: row[idx.actionItemName] || '',
      workstream: workstream,
      owner: row[idx.responder] || '',
      duedate: row[idx.duedate] ? formatDateISO(row[idx.duedate]) : '',
      status: row[idx.status] || 'Not Start',
      remark: row[idx.remark] || '',
      rowIndex: rowIndex
    };
    if (!tasksByWorkstream[workstream]) tasksByWorkstream[workstream] = [];
    tasksByWorkstream[workstream].push(task);
  });

  // If workstream order missing, derive from tasks
  if (!workstreamOrder.length) {
    workstreamOrder = Object.keys(tasksByWorkstream);
  }

  return { tasks: tasksByWorkstream, order: workstreamOrder };
}

// Helper: format a Date or string to 'YYYY-MM-DD'
function formatDateISO(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return val.toISOString().split('T')[0];
  }
  // For dates from sheet that are already string (e.g. 24-Jan-2025)
  const tryDate = new Date(val);
  if (!isNaN(tryDate.getTime())) return tryDate.toISOString().split('T')[0];
  return '';
}

/**
 * Adds a new task to the Google Sheet and returns the new task with rowIndex.
 */
function addTask(newTask) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  // Find header indexes
  const headers = sheet.getDataRange().getValues()[0];
  const idx = {
    number: headers.indexOf('#'),
    actionItemName: headers.indexOf('Action Item Name'),
    category: headers.indexOf('Category'),
    responder: headers.indexOf('Responder'),
    duedate: headers.indexOf('Due Date'),
    status: headers.indexOf('Status'),
    remark: headers.indexOf('Remark'),
    reference: headers.indexOf('Reference'),
    reportttb: headers.indexOf('Report to ttb'),
  };

  // Find the latest number in column A (ignoring header)
  let lastNumber = 0;
  if (idx.number > -1) {
    const numbers = sheet.getRange(2, idx.number + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    numbers.forEach(n => {
      if (!isNaN(n) && n !== '') lastNumber = Math.max(lastNumber, Number(n));
    });
  }

  // Append the new row at the bottom
  const values = Array(headers.length).fill('');
  if (idx.number > -1) values[idx.number] = lastNumber + 1;
  values[idx.actionItemName] = newTask.task;
  values[idx.category] = newTask.workstream;
  values[idx.responder] = newTask.owner;
  values[idx.duedate] = newTask.duedate ? new Date(newTask.duedate) : '';
  values[idx.status] = newTask.status || 'Not Start';
  values[idx.remark] = newTask.remark || '';
  sheet.appendRow(values);
  const newRowIndex = sheet.getLastRow();

  // Return new task in frontend format (with new rowIndex and number)
  return {
    ...newTask,
    number: lastNumber + 1,
    rowIndex: newRowIndex
  };
}

/**
 * Updates a task (row) in the sheet by rowIndex.
 */
function updateTask(task) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const headers = sheet.getDataRange().getValues()[0];
  const idx = {
    number: headers.indexOf('#'),
    actionItemName: headers.indexOf('Action Item Name'),
    category: headers.indexOf('Category'),
    responder: headers.indexOf('Responder'),
    duedate: headers.indexOf('Due Date'),
    status: headers.indexOf('Status'),
    remark: headers.indexOf('Remark'),
    reference: headers.indexOf('Reference'),
    reportttb: headers.indexOf('Report to ttb'),
  };
  const row = task.rowIndex;
  if (row < 2) return; // never update header

  if (idx.actionItemName > -1) sheet.getRange(row, idx.actionItemName + 1).setValue(task.task);
  if (idx.category > -1)        sheet.getRange(row, idx.category + 1).setValue(task.workstream);
  if (idx.responder > -1)       sheet.getRange(row, idx.responder + 1).setValue(task.owner);
  if (idx.duedate > -1)         sheet.getRange(row, idx.duedate + 1).setValue(task.duedate ? new Date(task.duedate) : '');
  if (idx.status > -1)          sheet.getRange(row, idx.status + 1).setValue(task.status);
  if (idx.remark > -1)          sheet.getRange(row, idx.remark + 1).setValue(task.remark);
  // You can also update reference/reportttb if you wish.
}

/**
 * Renames a workstream across all tasks and updates order sheet.
 */
function renameWorkstream(oldName, newName) {
  if (oldName === newName) return;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const catIdx = headers.indexOf('Category');
  // Update tasks
  for (let r = 1; r < data.length; r++) {
    if (data[r][catIdx] === oldName) {
      sheet.getRange(r + 1, catIdx + 1).setValue(newName);
    }
  }
  // Update order sheet
  const orderSheet = ss.getSheetByName(ORDER_SHEET_NAME);
  if (orderSheet) {
    const values = orderSheet.getRange(1, 1, orderSheet.getLastRow(), 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === oldName) {
        orderSheet.getRange(i + 1, 1).setValue(newName);
      }
    }
  }
}

/**
 * Moves a workstream up or down in the order sheet.
 */
function moveWorkstream(workstream, direction) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const orderSheet = ss.getSheetByName(ORDER_SHEET_NAME);
  if (!orderSheet) return;
  const values = orderSheet.getRange(1, 1, orderSheet.getLastRow(), 1).getValues().flat();
  const idx = values.indexOf(workstream);
  if (idx < 0) return;
  if (direction === 'up' && idx > 0) {
    [values[idx - 1], values[idx]] = [values[idx], values[idx - 1]];
  } else if (direction === 'down' && idx < values.length - 1) {
    [values[idx + 1], values[idx]] = [values[idx], values[idx + 1]];
  }
  // Write back new order
  for (let i = 0; i < values.length; i++) {
    orderSheet.getRange(i + 1, 1).setValue(values[i]);
  }
}

/**
 * Moves a task up or down within its workstream and swaps the index in column A.
 */
function moveTask(rowIndex, direction) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const catIdx = headers.indexOf('Category');
  const idxNumber = headers.indexOf('#');
  const idxStatus = headers.indexOf('Status');
  if (catIdx === -1 || idxNumber === -1 || idxStatus === -1) return;

  const taskRowIdx = rowIndex - 1;
  const taskWorkstream = data[taskRowIdx][catIdx];
  const taskStatus = data[taskRowIdx][idxStatus];

  // Collect all task rows in this workstream with the same status (skip header)
  const taskRows = [];
  for (let r = 1; r < data.length; r++) {
    if (data[r][catIdx] === taskWorkstream && data[r][idxStatus] === taskStatus) {
      taskRows.push(r);
    }
  }
  const pos = taskRows.indexOf(taskRowIdx);
  if (pos < 0) return;

  let swapWith = null;
  if (direction === 'up' && pos > 0) {
    swapWith = taskRows[pos - 1];
  } else if (direction === 'down' && pos < taskRows.length - 1) {
    swapWith = taskRows[pos + 1];
  }
  if (swapWith !== null) {
    // Swap entire rows (including the # index)
    const rowA = sheet.getRange(taskRowIdx + 1, 1, 1, headers.length).getValues()[0];
    const rowB = sheet.getRange(swapWith + 1, 1, 1, headers.length).getValues()[0];

    // Swap the # value too!
    const tempNum = rowA[idxNumber];
    rowA[idxNumber] = rowB[idxNumber];
    rowB[idxNumber] = tempNum;

    // Write swapped rows back to the sheet
    sheet.getRange(taskRowIdx + 1, 1, 1, headers.length).setValues([rowB]);
    sheet.getRange(swapWith + 1, 1, 1, headers.length).setValues([rowA]);
  }
}


// For web app
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Project Activity Management');
}

function addWorkstream(wsName, status) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const orderSheet = ss.getSheetByName(ORDER_SHEET_NAME);

  // Add workstream to order sheet
  if (orderSheet) {
    orderSheet.appendRow([wsName]);
  }

  // Get header indices
  const headers = sheet.getDataRange().getValues()[0];
  const idx = {
    number: headers.indexOf('#'),
    actionItemName: headers.indexOf('Action Item Name'),
    category: headers.indexOf('Category'),
    responder: headers.indexOf('Responder'),
    duedate: headers.indexOf('Due Date'),
    status: headers.indexOf('Status'),
    remark: headers.indexOf('Remark')
  };

  // Find max number
  let lastNumber = 0;
  if (idx.number > -1) {
    const numbers = sheet.getRange(2, idx.number + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    numbers.forEach(n => { if (!isNaN(n) && n !== '') lastNumber = Math.max(lastNumber, Number(n)); });
  }

  // Append first task for new workstream
  const values = Array(headers.length).fill('');
  if (idx.number > -1) values[idx.number] = lastNumber + 1;
  values[idx.actionItemName] = "New Task";
  values[idx.category] = wsName;
  values[idx.responder] = '';
  values[idx.duedate] = '';
  values[idx.status] = status || 'Not Start';
  values[idx.remark] = '';
  sheet.appendRow(values);

  return {
    workstream: wsName,
    task: {
      number: lastNumber + 1,
      task: "New Task",
      workstream: wsName,
      owner: '',
      duedate: '',
      status: status || 'Not Start',
      remark: '',
      rowIndex: sheet.getLastRow()
    }
  };
}

function deleteTask(rowIndex) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (rowIndex > 1) sheet.deleteRow(rowIndex);
}

function reindexAll() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const headers = sheet.getDataRange().getValues()[0];
  const idxNumber = headers.indexOf('#');
  const lastRow = sheet.getLastRow();
  for (let i = 2; i <= lastRow; i++) {
    sheet.getRange(i, idxNumber + 1).setValue(i - 1);
  }
}

// Add this to moveWorkstream, and duplicate for left/right for kanban
function moveWorkstreamLR(workstream, direction) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const orderSheet = ss.getSheetByName(ORDER_SHEET_NAME);
  if (!orderSheet) return;
  const values = orderSheet.getRange(1, 1, orderSheet.getLastRow(), 1).getValues().flat();
  const idx = values.indexOf(workstream);
  if (idx < 0) return;
  if (direction === 'left' && idx > 0) {
    [values[idx - 1], values[idx]] = [values[idx], values[idx - 1]];
  } else if (direction === 'right' && idx < values.length - 1) {
    [values[idx + 1], values[idx]] = [values[idx], values[idx + 1]];
  }
  for (let i = 0; i < values.length; i++) {
    orderSheet.getRange(i + 1, 1).setValue(values[i]);
  }
}
