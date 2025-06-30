# Project Activity Management

Project Activity Management is a Google Apps Script web app for tracking tasks in a Google Sheet. Tasks are organized by "workstreams" and can be viewed in either a simple list or a kanban board.

## Features

- **List or Kanban View** – toggle between a list of tasks or a kanban-style board.
- **Workstreams** – group tasks by workstream; rename or reorder workstreams.
- **Task Fields** – each task records the name, owner, due date, status, and a rich-text remark.
- **Filtering** – filter tasks by status such as In-Progress, Not Start, Complete, On hold, TBC from ttb, or Closed.
- **Add, Update, Delete** – edit tasks in place and sync changes with the sheet.

## Setup

1. Create (or open) a Google Sheet and add the headers `#`, `Action Item Name`, `Category`, `Responder`, `Due Date`, `Status`, and `Remark`.
2. In Apps Script, add the files `Code.gs` and `index.html` from this repository.
3. Replace the placeholders in `Code.gs` with your sheet ID and tab name:
   ```javascript
   const SHEET_ID = "<YOUR_SHEET_ID>";
   const SHEET_NAME = "<SHEET_TAB_NAME>";
   ```
4. Deploy the Apps Script as a Web App and note the generated URL.

## How It Works

`index.html` renders the interface with Tailwind CSS and Quill.js. The `doGet` function serves the page:
```javascript
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Project Activity Management');
}
```
Server-side functions such as `addTask` update the Google Sheet:
```javascript
function addTask(task) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  sheet.appendRow([
    task.task,
    task.workstream,
    // ...additional fields
  ]);
}
```

## License

This project is released under the MIT License. See [LICENSE](LICENSE) if provided.
