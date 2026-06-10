// ════════════════════════════════════════════════════════════════
//  EMPLOYEE DAILY WORK REPORT — Google Apps Script Backend
//  Receives form submissions and writes to Google Sheets
//
//  HOW TO DEPLOY (step-by-step):
//  1. Go to https://script.google.com → New Project
//  2. Paste this entire file into the editor
//  3. Change SHEET_ID below to your Google Sheet ID
//  4. Click Deploy → New Deployment
//  5. Type: Web App
//  6. Execute as: Me
//  7. Who has access: Anyone  (so employees can submit)
//  8. Click Deploy → Copy the Web App URL
//  9. Paste that URL into Employee_Report_Form.html where it says:
//     const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbx2sCj2VzD2om5SODQ1VlRbbtbV8-_kZo6BM1ORiGbj4taG6pPLJ58-m_Qfq18exKOT0Q/exec";
// ════════════════════════════════════════════════════════════════

// ── CONFIGURATION ──────────────────────────────────────────────
var SHEET_ID   = "YOUR_GOOGLE_SHEET_ID_HERE";   // ← paste your Sheet ID here
                                                   // (from the URL: /d/SHEET_ID/edit)
var SHEET_NAME_REPORTS  = "Daily Reports";         // main submissions tab
var SHEET_NAME_TASKS    = "Task Details";          // individual task rows
var SHEET_NAME_MEETINGS = "Meeting Details";       // individual meeting rows
// ───────────────────────────────────────────────────────────────


// ════════════════════════════════════════════════════════════════
//  doPost — called when the HTML form submits
// ════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var params = e.parameter;
    
    var ss;
    if (SHEET_ID && SHEET_ID !== "YOUR_GOOGLE_SHEET_ID_HERE") {
      ss = SpreadsheetApp.openById(SHEET_ID);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    if (!ss) {
      throw new Error("Spreadsheet not found. Please set SHEET_ID in the script configuration or ensure the script is container-bound.");
    }

    setupSheets(ss);  // ensure all sheets & headers exist

    var reportId = writeMainReport(ss, params);
    writeTasks(ss, params, reportId);
    writeMeetings(ss, params, reportId);
    sendConfirmationEmail(params);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", reportId: reportId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ════════════════════════════════════════════════════════════════
//  doGet — health check (visit the URL in browser to test)
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  return ContentService
    .createTextOutput("✅ Employee Report API is live. Use POST to submit reports.")
    .setMimeType(ContentService.MimeType.TEXT);
}


// ════════════════════════════════════════════════════════════════
//  SHEET SETUP — creates tabs and headers if they don't exist
// ════════════════════════════════════════════════════════════════
function setupSheets(ss) {

  // ── Daily Reports sheet ──
  var ws = getOrCreateSheet(ss, SHEET_NAME_REPORTS);
  if (ws.getLastRow() === 0) {
    var reportHeaders = [
      "Report ID", "Submitted At", "Report Date",
      "Employee Name", "Employee ID", "Department", "Designation",
      "Manager", "Work Mode", "Shift Start", "Shift End", "Total Hours",
      "Attendance",
      "Accomplishments", "Deliverables", "Challenges", "Risks",
      "Dependencies", "Tomorrow Plan", "Support Required", "Additional Notes",
      "Goals Achieved %", "Mood",
      "Total Tasks", "Completed Tasks", "Total Hours Logged",
      "Total Meetings"
    ];
    ws.appendRow(reportHeaders);
    styleHeaderRow(ws, reportHeaders.length);
  }

  // ── Task Details sheet ──
  var tw = getOrCreateSheet(ss, SHEET_NAME_TASKS);
  if (tw.getLastRow() === 0) {
    var taskHeaders = [
      "Report ID", "Report Date", "Employee Name", "Employee ID",
      "Task Title", "Category", "Priority", "Status",
      "Estimated Hours", "Actual Hours", "Progress %", "Notes"
    ];
    tw.appendRow(taskHeaders);
    styleHeaderRow(tw, taskHeaders.length);
  }

  // ── Meeting Details sheet ──
  var mw = getOrCreateSheet(ss, SHEET_NAME_MEETINGS);
  if (mw.getLastRow() === 0) {
    var mtgHeaders = [
      "Report ID", "Report Date", "Employee Name", "Employee ID",
      "Meeting Title", "Stakeholders", "Type",
      "Start Time", "End Time", "Outcome", "Action Items", "Follow-up?"
    ];
    mw.appendRow(mtgHeaders);
    styleHeaderRow(mw, mtgHeaders.length);
  }
}


// ════════════════════════════════════════════════════════════════
//  WRITE MAIN REPORT ROW
// ════════════════════════════════════════════════════════════════
function writeMainReport(ss, p) {
  var ws       = ss.getSheetByName(SHEET_NAME_REPORTS);
  var reportId = "RPT-" + Utilities.formatDate(new Date(), "GMT", "yyyyMMdd") +
                 "-" + Math.floor(Math.random() * 9000 + 1000);

  var tasks    = safeParseJSON(p.tasks    || "[]");
  var meetings = safeParseJSON(p.meetings || "[]");

  var totalTasks    = tasks.length;
  var completedTasks = tasks.filter(function(t){ return t && t.status === "Completed"; }).length;
  var totalHoursLogged = tasks.reduce(function(sum, t){
    return sum + (t && parseFloat(t.actual) || 0);
  }, 0);
  var totalMeetings = meetings.length;

  var row = [
    reportId,
    new Date(),
    p.reportDate        || "",
    p.fullName          || "",
    p.employeeId        || "",
    p.department        || "",
    p.designation       || "",
    p.managerName       || "",
    p.workMode          || "",
    p.shiftStart        || "",
    p.shiftEnd          || "",
    p.totalHours        || "",
    p.attendance        || "",
    p.accomplishments   || "",
    p.deliverables      || "",
    p.challenges        || "",
    p.risks             || "",
    p.dependencies      || "",
    p.tomorrowPlan      || "",
    p.supportRequired   || "",
    p.additionalNotes   || "",
    p.goalsAchieved     || "",
    p.mood              || "",
    totalTasks,
    completedTasks,
    totalHoursLogged,
    totalMeetings
  ];

  ws.appendRow(row);
  applyRowZebra(ws, ws.getLastRow());

  return reportId;
}


// ════════════════════════════════════════════════════════════════
//  WRITE TASK ROWS
// ════════════════════════════════════════════════════════════════
function writeTasks(ss, p, reportId) {
  var tw    = ss.getSheetByName(SHEET_NAME_TASKS);
  var tasks = safeParseJSON(p.tasks || "[]");

  tasks.forEach(function(t) {
    if (!t || !t.title) return;
    var row = [
      reportId,
      p.reportDate   || "",
      p.fullName     || "",
      p.employeeId   || "",
      t.title        || "",
      t.cat          || "",
      t.priority     || "",
      t.status       || "",
      t.est          || "",
      t.actual       || "",
      t.prog         || "",
      t.notes        || ""
    ];
    tw.appendRow(row);
    applyPriorityColor(tw, tw.getLastRow(), t.priority, t.status);
  });
}


// ════════════════════════════════════════════════════════════════
//  WRITE MEETING ROWS
// ════════════════════════════════════════════════════════════════
function writeMeetings(ss, p, reportId) {
  var mw       = ss.getSheetByName(SHEET_NAME_MEETINGS);
  var meetings = safeParseJSON(p.meetings || "[]");

  meetings.forEach(function(m) {
    if (!m || !m.title) return;
    var row = [
      reportId,
      p.reportDate      || "",
      p.fullName        || "",
      p.employeeId      || "",
      m.title           || "",
      m.stakeholders    || "",
      m.type            || "",
      m.start           || "",
      m.end             || "",
      m.outcome         || "",
      m.actions         || "",
      m.followup        || ""
    ];
    mw.appendRow(row);
    applyFollowUpColor(mw, mw.getLastRow(), m.followup);
  });
}


// ════════════════════════════════════════════════════════════════
//  OPTIONAL: Send confirmation email to employee
//  (Remove or comment out if not needed)
// ════════════════════════════════════════════════════════════════
function sendConfirmationEmail(p) {
  try {
    var employeeEmail = p.employeeEmail || "";  // add email field to form if needed
    if (!employeeEmail) return;

    var subject = "✅ Daily Report Submitted — " + p.reportDate;
    var body = "Hi " + (p.fullName || "Team Member") + ",\n\n" +
      "Your daily work report for " + p.reportDate + " has been submitted successfully.\n\n" +
      "Report ID: RPT-" + p.reportDate + "\n" +
      "Department: " + (p.department || "—") + "\n" +
      "Tasks logged: " + (safeParseJSON(p.tasks || "[]").length) + "\n" +
      "Meetings logged: " + (safeParseJSON(p.meetings || "[]").length) + "\n\n" +
      "Have a great evening!\n— HR System";

    MailApp.sendEmail(employeeEmail, subject, body);
  } catch(e) {
    // email is optional — silently ignore errors
  }
}


// ════════════════════════════════════════════════════════════════
//  STYLE HELPERS
// ════════════════════════════════════════════════════════════════

function getOrCreateSheet(ss, name) {
  var ws = ss.getSheetByName(name);
  if (!ws) ws = ss.insertSheet(name);
  return ws;
}

function styleHeaderRow(ws, numCols) {
  var headerRange = ws.getRange(1, 1, 1, numCols);
  headerRange.setBackground("#0F172A");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment("center");
  headerRange.setWrap(true);
  ws.setFrozenRows(1);
  ws.setColumnWidth(1, 140);
}

function applyRowZebra(ws, rowNum) {
  var numCols = ws.getLastColumn();
  var color = (rowNum % 2 === 0) ? "#F8FAFC" : "#FFFFFF";
  ws.getRange(rowNum, 1, 1, numCols).setBackground(color);
}

function applyPriorityColor(tw, rowNum, priority, status) {
  var numCols = tw.getLastColumn();
  var bg = "#FFFFFF";
  if (status === "Completed") {
    bg = "#D1FAE5";
  } else if (priority === "Critical") {
    bg = "#FEE2E2";
  } else if (priority === "High") {
    bg = "#FED7AA";
  } else if (priority === "Medium") {
    bg = "#DBEAFE";
  }
  if (rowNum % 2 === 0 && bg === "#FFFFFF") bg = "#F8FAFC";
  tw.getRange(rowNum, 1, 1, numCols).setBackground(bg);
}

function applyFollowUpColor(mw, rowNum, followup) {
  var numCols = mw.getLastColumn();
  var bg = (rowNum % 2 === 0) ? "#F8FAFC" : "#FFFFFF";
  if (followup === "Yes") bg = "#FEE2E2";
  else if (followup === "No") bg = "#D1FAE5";
  mw.getRange(rowNum, 1, 1, numCols).setBackground(bg);
}

function safeParseJSON(str) {
  try { return JSON.parse(str) || []; }
  catch(e) { return []; }
}


// ════════════════════════════════════════════════════════════════
//  BONUS: Auto-create a Google Sheet with correct structure
//  Run this function ONCE manually to set up a fresh Sheet
// ════════════════════════════════════════════════════════════════
function createFreshSheet() {
  var ss = SpreadsheetApp.create("Employee Daily Work Reports");
  Logger.log("New Sheet ID: " + ss.getId());
  Logger.log("New Sheet URL: " + ss.getUrl());

  // Update SHEET_ID above and re-deploy
  setupSheets(ss);

  // Auto-resize all columns for readability
  ss.getSheets().forEach(function(ws) {
    ws.autoResizeColumns(1, ws.getLastColumn());
  });

  return ss.getUrl();
}
