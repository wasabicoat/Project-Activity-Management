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
    const task = row[1];         // Column B
    const workstream = row[2];   // Column C
    const owner = row[3];        // Column D
    const duedate = formatDate(row[4]); // Column E
    const status = row[5] ? row[5].toString().trim() : 'Not Start'; // Column F
    const remark = row[8];       // Column I

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

function addTask(task) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  sheet.appendRow(['', task.task, task.workstream, task.owner, task.duedate, task.status, '', '', task.remark]);
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
  }
  const values = orderSheet.getRange(1, 1, orderSheet.getLastRow(), 1).getValues().flat();
  const index = values.indexOf(workstream);
  if (index === -1) return;
  if (direction === 'up' && index > 0) {
    [values[index - 1], values[index]] = [values[index], values[index - 1]];
  } else if (direction === 'down' && index < values.length - 1) {
    [values[index + 1], values[index]] = [values[index], values[index + 1]];
  }
  orderSheet.clear();
  values.forEach((v, i) => orderSheet.getRange(i + 1, 1).setValue(v));
}

function moveTask(rowIndex, direction) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const index = rowIndex - 1;
  const task = data[index];
  const workstream = task[2];

  let adjacentIndex = -1;
  if (direction === 'up') {
    for (let i = index - 1; i >= 1; i--) {
      if (data[i][2] === workstream) {
        adjacentIndex = i;
        break;
      }
    }
  } else if (direction === 'down') {
    for (let i = index + 1; i < data.length; i++) {
      if (data[i][2] === workstream) {
        adjacentIndex = i;
        break;
      }
    }
  }

  if (adjacentIndex === -1) return;

  const range1 = sheet.getRange(index + 1, 1, 1, data[0].length);
  const range2 = sheet.getRange(adjacentIndex + 1, 1, 1, data[0].length);

  const temp = range1.getValues();
  range1.setValues(range2.getValues());
  range2.setValues(temp);
}
