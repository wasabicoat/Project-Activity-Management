const SHEET_ID = '1Cc-I45ehhpLB86dp-qBeOJh0aSj5V_0OOqyY3DG_1ZQ';
const SHEET_NAME = 'action listlog';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Project Activity Management');
}

function getTaskData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  const idx = {
    number: header.indexOf('#'),
    task: header.indexOf('Action Item Name'),
    workstream: header.indexOf('Category'),
    owner: header.indexOf('Responder'),
    duedate: header.indexOf('Due Date'),
    status: header.indexOf('Status'),
    remark: header.indexOf('Remark')
  };

  // Flat list, with rowIndex for later updating
  const tasks = rows.map((row, i) => ({
    number: row[idx.number],
    task: row[idx.task],
    workstream: row[idx.workstream],
    owner: row[idx.owner],
    duedate: row[idx.duedate] ? formatDateForInput(row[idx.duedate]) : '',
    status: row[idx.status],
    remark: row[idx.remark],
    rowIndex: i + 2 // because header is row 1
  }));

  return { tasks };
}

function formatDateForInput(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  if (date instanceof Date && !isNaN(date)) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return '';
}

function updateTask(task) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idx = {
    number: header.indexOf('#') + 1,
    task: header.indexOf('Action Item Name') + 1,
    workstream: header.indexOf('Category') + 1,
    owner: header.indexOf('Responder') + 1,
    duedate: header.indexOf('Due Date') + 1,
    status: header.indexOf('Status') + 1,
    remark: header.indexOf('Remark') + 1
  };
  const row = task.rowIndex;
  sheet.getRange(row, idx.task).setValue(task.task);
  sheet.getRange(row, idx.owner).setValue(task.owner);
  sheet.getRange(row, idx.duedate).setValue(task.duedate || '');
  sheet.getRange(row, idx.status).setValue(task.status);
  sheet.getRange(row, idx.remark).setValue(task.remark);
  sheet.getRange(row, idx.workstream).setValue(task.workstream);
}

function addTask(task) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idx = {
    number: header.indexOf('#'),
    task: header.indexOf('Action Item Name'),
    workstream: header.indexOf('Category'),
    owner: header.indexOf('Responder'),
    duedate: header.indexOf('Due Date'),
    status: header.indexOf('Status'),
    remark: header.indexOf('Remark')
  };

  // Find max # in the column for new index
  let lastNumber = 0;
  for (let i = 1; i < data.length; i++) {
    const n = data[i][idx.number];
    if (!isNaN(n) && n !== '') lastNumber = Math.max(lastNumber, Number(n));
  }
  const values = Array(header.length).fill('');
  values[idx.number] = lastNumber + 1;
  values[idx.task] = task.task || 'New Task';
  values[idx.workstream] = task.workstream;
  values[idx.owner] = task.owner || '';
  values[idx.duedate] = task.duedate || '';
  values[idx.status] = task.status || 'Not Start';
  values[idx.remark] = task.remark || '';
  sheet.appendRow(values);
  const newRowIdx = sheet.getLastRow();

  return {
    number: lastNumber + 1,
    task: values[idx.task],
    workstream: values[idx.workstream],
    owner: values[idx.owner],
    duedate: values[idx.duedate],
    status: values[idx.status],
    remark: values[idx.remark],
    rowIndex: newRowIdx
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
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idxNumber = header.indexOf('#');
  for (let i = 1; i < data.length; i++) {
    sheet.getRange(i + 1, idxNumber + 1).setValue(i);
  }
}

function moveWorkstreamInSheet(workstream, direction) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idxWS = header.indexOf('Category');
  const rows = data.slice(1);

  // Group rows by workstream
  let wsMap = {};
  rows.forEach((row, i) => {
    const ws = row[idxWS];
    if (!wsMap[ws]) wsMap[ws] = [];
    wsMap[ws].push(i); // index in rows
  });
  let order = Object.keys(wsMap);
  let pos = order.indexOf(workstream);
  if (pos === -1) return;
  let targetPos = null;
  if (direction === 'up' && pos > 0) targetPos = pos - 1;
  if (direction === 'down' && pos < order.length - 1) targetPos = pos + 1;
  if (targetPos === null) return;

  // Swap
  const newOrder = order.slice();
  [newOrder[pos], newOrder[targetPos]] = [newOrder[targetPos], newOrder[pos]];

  // Build new rows array
  const newRows = [];
  newOrder.forEach(ws => {
    wsMap[ws].forEach(idx => newRows.push(rows[idx]));
  });

  // Write back to sheet (overwrite all rows below header)
  if (newRows.length > 0)
    sheet.getRange(2, 1, newRows.length, newRows[0].length).setValues(newRows);

  // Remove extra rows if any
  const lastRow = sheet.getLastRow();
  if (lastRow > newRows.length + 1) sheet.deleteRows(newRows.length + 2, lastRow - (newRows.length + 1));

  reindexAll();
}

function renameWorkstream(oldName, newName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const idx = header.indexOf('Category') + 1;
  for (let i = 2; i <= data.length; i++) {
    if (sheet.getRange(i, idx).getValue() === oldName) {
      sheet.getRange(i, idx).setValue(newName);
    }
  }
}
