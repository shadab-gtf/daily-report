// ════════════════════════════════════════════════════════════════
//  EMPLOYEE DAILY WORK REPORT — Google Apps Script Backend
//  Receives form submissions, manages employees, and queries sheets
// ════════════════════════════════════════════════════════════════

// ── CONFIGURATION ──────────────────────────────────────────────
var SHEET_ID   = "1AmoAsv4YvMQwHs9FESvszUhXSKt6Lr7HLquTDLz7-Wk";   // ← paste your Sheet ID here
                                                   // (from the URL: /d/SHEET_ID/edit)
var SHEET_NAME_REPORTS   = "Daily Reports";        // main submissions tab
var SHEET_NAME_TASKS     = "Task Details";         // individual task rows
var SHEET_NAME_EMPLOYEES = "Employees";            // registered employees tab
// ───────────────────────────────────────────────────────────────


// ════════════════════════════════════════════════════════════════
//  doPost — handles write operations (login, submits, creations)
// ════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var params = e.parameter;
    var action = params.action;
    
    var ss = getSpreadsheet();
    setupSheets(ss);
    

    if (action === "submitReport") {
      return handleSubmitReport(ss, params);
    } else if (action === "createEmployee") {
      return handleCreateEmployee(ss, params);
    } else if (action === "login") {
      return handleLogin(ss, params);
    } else if (action === "getReports") {
      return handleGetReports(ss, params);
    } else if (action === "getEmployees") {
      return handleGetEmployees(ss, params);
    } else if (action === "assignTask") {
      return handleAssignTask(ss, params);
    }// Default fallback to submit report if no action is provided
    return handleSubmitReport(ss, params);
  } catch (err) {
    return createJSONResponse({ status: "error", message: err.toString() });
  }
}


// ════════════════════════════════════════════════════════════════
//  doGet — handles read operations (avoiding preflight CORS)
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var params = e.parameter;
    var action = params.action;
    
    var ss = getSpreadsheet();
    setupSheets(ss);
    
    if (action === "login") {
      return handleLogin(ss, params);
    } else if (action === "getReports") {
      return handleGetReports(ss, params);
    } else if (action === "getEmployees") {
      return handleGetEmployees(ss, params);
    }
    
    return createJSONResponse({ status: "success", message: "✅ API is active." });
  } catch (err) {
    return createJSONResponse({ status: "error", message: err.toString() });
  }
}


// ════════════════════════════════════════════════════════════════
//  SPREADSHEET RESOLVER
// ════════════════════════════════════════════════════════════════
function getSpreadsheet() {
  var ss;
  if (SHEET_ID && SHEET_ID !== "YOUR_GOOGLE_SHEET_ID_HERE") {
    ss = SpreadsheetApp.openById(SHEET_ID);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) {
    throw new Error("Spreadsheet not found. Please set SHEET_ID or run from bound sheet.");
  }
  return ss;
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
      "Total Tasks", "Completed Tasks", "Total Hours Logged"
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

  // ── Employees sheet ──
  var ew = getOrCreateSheet(ss, SHEET_NAME_EMPLOYEES);
  if (ew.getLastRow() === 0) {
    var empHeaders = [
      "Employee ID", "Employee Name", "Designation", "Department", "Created At"
    ];
    ew.appendRow(empHeaders);
    styleHeaderRow(ew, empHeaders.length);
    
    // Seed manager and initial test employee
    ew.appendRow(["saurabh.yadav", "Saurabh Yadav", "Manager", "Web Development", new Date()]);
    ew.appendRow(["GTF001", "Shadab", "Frontend", "Web Development", new Date()]);
  }
}


// ════════════════════════════════════════════════════════════════
//  LOGIN HANDLER
// ════════════════════════════════════════════════════════════════
function handleLogin(ss, p) {
  var employeeId = p.employeeId;
  if (!employeeId) {
    return createJSONResponse({ status: "error", message: "Employee ID is required." });
  }
  
  employeeId = employeeId.trim();
  
  if (employeeId.toLowerCase() === "saurabh.yadav") {
    return createJSONResponse({
      status: "success",
      user: {
        employeeId: "saurabh.yadav",
        fullName: "Saurabh Yadav",
        role: "manager",
        designation: "Manager",
        department: "Web Development"
      }
    });
  }
  
  var ew = ss.getSheetByName(SHEET_NAME_EMPLOYEES);
  var data = ew.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === employeeId.toLowerCase()) {
      return createJSONResponse({
        status: "success",
        user: {
          employeeId: data[i][0].toString().trim(),
          fullName: data[i][1].toString().trim(),
          designation: data[i][2].toString().trim(),
          department: data[i][3].toString().trim(),
          role: "employee"
        }
      });
    }
  }
  
  return createJSONResponse({ status: "error", message: "Employee ID not found." });
}


// ════════════════════════════════════════════════════════════════
//  CREATE EMPLOYEE HANDLER (Manager only)
// ════════════════════════════════════════════════════════════════
function handleCreateEmployee(ss, p) {
  var creatorId = p.creatorId;
  if (!creatorId || creatorId.toLowerCase() !== "saurabh.yadav") {
    return createJSONResponse({ status: "error", message: "Unauthorized. Only Saurabh Yadav can create employees." });
  }
  
  var employeeId = p.employeeId;
  var fullName = p.fullName;
  var designation = p.designation;
  var department = p.department || "Web Development";
  
  if (!employeeId || !fullName || !designation) {
    return createJSONResponse({ status: "error", message: "Missing required employee details." });
  }
  
  employeeId = employeeId.trim();
  
  if (!employeeId.toUpperCase().startsWith("GTF")) {
    return createJSONResponse({ status: "error", message: "Employee ID must start with GTF." });
  }
  
  var ew = ss.getSheetByName(SHEET_NAME_EMPLOYEES);
  var data = ew.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === employeeId.toLowerCase()) {
      return createJSONResponse({ status: "error", message: "Employee ID already exists." });
    }
  }
  
  ew.appendRow([employeeId, fullName, designation, department, new Date()]);
  return createJSONResponse({ status: "success", message: "Employee " + employeeId + " created successfully." });
}


// ════════════════════════════════════════════════════════════════
//  LIST EMPLOYEES HANDLER (Manager only)
// ════════════════════════════════════════════════════════════════
function handleGetEmployees(ss, p) {
  var requesterId = p.requesterId;
  if (!requesterId || requesterId.toLowerCase() !== "saurabh.yadav") {
    return createJSONResponse({ status: "error", message: "Unauthorized." });
  }
  
  var ew = ss.getSheetByName(SHEET_NAME_EMPLOYEES);
  var data = ew.getDataRange().getValues();
  var employees = [];
  
  for (var i = 1; i < data.length; i++) {
    employees.push({
      employeeId: data[i][0],
      fullName: data[i][1],
      designation: data[i][2],
      department: data[i][3],
      createdAt: data[i][4]
    });
  }
  
  return createJSONResponse({ status: "success", employees: employees });
}


// ════════════════════════════════════════════════════════════════
//  GET REPORTS & TASKS HANDLER
// ════════════════════════════════════════════════════════════════
function handleGetReports(ss, p) {
  var requesterId = p.requesterId;
  if (!requesterId) {
    return createJSONResponse({ status: "error", message: "Requester ID is required." });
  }
  
  var isManager = (requesterId.toLowerCase() === "saurabh.yadav");
  var targetEmployeeId = p.targetEmployeeId || "";
  
  var rws = ss.getSheetByName(SHEET_NAME_REPORTS);
  var reportsData = rws.getDataRange().getValues();
  var reportsHeaders = reportsData[0];
  
  var tws = ss.getSheetByName(SHEET_NAME_TASKS);
  var tasksData = tws.getDataRange().getValues();
  var tasksHeaders = tasksData[0];
  
  var reports = [];
  
  function rowToObject(headers, row) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j];
      var cleanKey = key.replace(/[^a-zA-Z0-9]/g, "");
      cleanKey = cleanKey.charAt(0).toLowerCase() + cleanKey.slice(1);
      
      if (cleanKey === "reportID") cleanKey = "reportId";
      if (cleanKey === "employeeID") cleanKey = "employeeId";
      if (cleanKey === "employeeName") cleanKey = "employeeName";
      if (cleanKey === "goalsAchieved") cleanKey = "goalsAchieved";
      
      var val = row[j];
      if (val instanceof Date) {
        if (val.getFullYear() === 1899) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
        }
      }
      obj[cleanKey] = val;
    }
    return obj;
  }
  
  for (var i = 1; i < reportsData.length; i++) {
    var rRow = reportsData[i];
    var empId = rRow[4] ? rRow[4].toString().trim() : "";
    
    if (isManager) {
      if (targetEmployeeId && empId.toLowerCase() !== targetEmployeeId.toLowerCase()) {
        continue;
      }
    } else {
      if (empId.toLowerCase() !== requesterId.toLowerCase()) {
        continue;
      }
    }
    
    var rObj = rowToObject(reportsHeaders, rRow);
    rObj.tasks = [];
    reports.push(rObj);
  }
  
  var tasksByReportId = {};
  var assignedTasks = [];
  for (var k = 1; k < tasksData.length; k++) {
    var tRow = tasksData[k];
    var rId = tRow[0] ? tRow[0].toString().trim() : "";
    var empId = tRow[3] ? tRow[3].toString().trim() : "";
    var tObj = rowToObject(tasksHeaders, tRow);
    
    if (rId === "ASSIGNED") {
      if (isManager) {
        if (!targetEmployeeId || empId.toLowerCase() === targetEmployeeId.toLowerCase()) {
          assignedTasks.push(tObj);
        }
      } else {
        if (empId.toLowerCase() === requesterId.toLowerCase() && tObj.status !== "Completed") {
          assignedTasks.push(tObj);
        }
      }
    } else {
      if (!tasksByReportId[rId]) {
        tasksByReportId[rId] = [];
      }
      tasksByReportId[rId].push(tObj);
    }
  }
  
  reports.forEach(function(r) {
    r.tasks = tasksByReportId[r.reportId] || [];
  });
  
  reports.sort(function(a, b) {
    return new Date(b.submittedAt) - new Date(a.submittedAt);
  });
  
  var spreadsheetUrl = "";
  if (isManager) {
    spreadsheetUrl = ss.getUrl();
  }
  
  return createJSONResponse({ 
    status: "success", 
    reports: reports, 
    assignedTasks: assignedTasks, 
    spreadsheetUrl: spreadsheetUrl 
  });
}


// ════════════════════════════════════════════════════════════════
//  SUBMIT / UPDATE DAILY REPORT HANDLER
// ════════════════════════════════════════════════════════════════
function handleSubmitReport(ss, p) {
  var reportDate = p.reportDate;
  var employeeId = p.employeeId;
  
  if (!reportDate || !employeeId) {
    return createJSONResponse({ status: "error", message: "Missing Report Date or Employee ID." });
  }
  
  reportDate = reportDate.trim();
  employeeId = employeeId.trim();
  
  var rws = ss.getSheetByName(SHEET_NAME_REPORTS);
  var reportsData = rws.getDataRange().getValues();
  
  var existingRowIndex = -1;
  var reportId = "";
  
  // Find report for same date and employee
  for (var i = 1; i < reportsData.length; i++) {
    var rRow = reportsData[i];
    var rowDate = rRow[2] ? rRow[2].toString().trim() : "";
    var rowEmpId = rRow[4] ? rRow[4].toString().trim() : "";
    
    if (formatDateString(rowDate) === formatDateString(reportDate) && 
        rowEmpId.toLowerCase() === employeeId.toLowerCase()) {
      existingRowIndex = i + 1;
      reportId = rRow[0].toString().trim();
      break;
    }
  }
  
  var tasks = safeParseJSON(p.tasks || "[]");
  var totalTasks = tasks.length;
  var completedTasks = tasks.filter(function(t){ return t && t.status === "Completed"; }).length;
  var totalHoursLogged = tasks.reduce(function(sum, t){
    return sum + (t && parseFloat(t.actual) || 0);
  }, 0);
  
  var row = [
    reportId || ("RPT-" + Utilities.formatDate(new Date(), "GMT", "yyyyMMdd") + "-" + Math.floor(Math.random() * 9000 + 1000)),
    new Date(),
    reportDate,
    p.fullName          || "",
    employeeId,
    p.department        || "Web Development",
    p.designation       || "",
    p.managerName       || "Saurabh Yadav",
    p.workMode          || "",
    p.shiftStart        || "",
    p.shiftEnd          || "",
    p.totalHours        || "",
    p.attendance        || "Present",
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
    totalHoursLogged
  ];
  
  if (!reportId) {
    reportId = row[0];
  }
  
  if (existingRowIndex > 0) {
    var range = rws.getRange(existingRowIndex, 1, 1, row.length);
    range.setValues([row]);
    applyRowZebra(rws, existingRowIndex);
    deleteTasksForReport(ss, reportId);
  } else {
    rws.appendRow(row);
    applyRowZebra(rws, rws.getLastRow());
  }
  
  clearAssignedTasks(ss, employeeId, tasks);
  writeTasks(ss, p, reportId);
  sendConfirmationEmail(p);
  
  return createJSONResponse({ 
    status: "success", 
    reportId: reportId, 
    message: existingRowIndex > 0 ? "Report updated successfully!" : "Report submitted successfully!" 
  });
}

function clearAssignedTasks(ss, employeeId, tasks) {
  try {
    var tw = ss.getSheetByName(SHEET_NAME_TASKS);
    var data = tw.getDataRange().getValues();
    
    var submittedTitles = {};
    tasks.forEach(function(t) {
      if (t && t.title) {
        submittedTitles[t.title.trim().toLowerCase()] = true;
      }
    });

    for (var i = data.length - 1; i >= 1; i--) {
      var rId = data[i][0] ? data[i][0].toString().trim() : "";
      var empId = data[i][3] ? data[i][3].toString().trim() : "";
      var title = data[i][4] ? data[i][4].toString().trim() : "";
      
      if (rId === "ASSIGNED" && 
          empId.toLowerCase() === employeeId.toLowerCase() && 
          submittedTitles[title.toLowerCase()]) {
        tw.deleteRow(i + 1);
      }
    }
  } catch(e) {
    // Fail-safe
  }
}

function handleAssignTask(ss, p) {
  var employeeId = p.employeeId;
  var taskTitle = p.taskTitle;
  var category = p.category || "Web Development";
  var priority = p.priority || "Medium";
  var estHours = p.estHours || "";
  var notes = p.notes || "";
  
  if (!employeeId || !taskTitle) {
    return createJSONResponse({ status: "error", message: "Missing Employee ID or Task Title." });
  }

  // Find employee name from directory
  var ew = ss.getSheetByName(SHEET_NAME_EMPLOYEES);
  var data = ew.getDataRange().getValues();
  var employeeName = "";
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === employeeId.trim().toLowerCase()) {
      employeeName = data[i][1].toString().trim();
      break;
    }
  }

  var tw = ss.getSheetByName(SHEET_NAME_TASKS);
  var row = [
    "ASSIGNED",               // Report ID
    new Date(),               // Report Date (assigned date)
    employeeName || "",
    employeeId.toUpperCase(),
    taskTitle,
    category,
    priority,
    "Not Started",
    estHours,
    0,                        // Actual Hours
    0,                        // Progress %
    notes
  ];
  
  tw.appendRow(row);
  applyPriorityColor(tw, tw.getLastRow(), priority, "Not Started");
  
  return createJSONResponse({ status: "success", message: "Task assigned successfully to " + (employeeName || employeeId) + "." });
}


// ════════════════════════════════════════════════════════════════
//  TASK WRITER Helper
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
//  DELETE TASKS Helper (For Updates)
// ════════════════════════════════════════════════════════════════
function deleteTasksForReport(ss, reportId) {
  var tw = ss.getSheetByName(SHEET_NAME_TASKS);
  var data = tw.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] && data[i][0].toString().trim() === reportId) {
      tw.deleteRow(i + 1);
    }
  }
}


// ════════════════════════════════════════════════════════════════
//  CONFIRMATION EMAIL Helper
// ════════════════════════════════════════════════════════════════
function sendConfirmationEmail(p) {
  try {
    var employeeEmail = p.employeeEmail || "";
    if (!employeeEmail) return;

    var subject = "✅ Daily Report Submitted — " + p.reportDate;
    var body = "Hi " + (p.fullName || "Team Member") + ",\n\n" +
      "Your daily work report for " + p.reportDate + " has been submitted successfully.\n\n" +
      "Report ID: RPT-" + p.reportDate + "\n" +
      "Department: " + (p.department || "—") + "\n" +
      "Tasks logged: " + (safeParseJSON(p.tasks || "[]").length) + "\n\n" +
      "Have a great evening!\n— HR System";

    MailApp.sendEmail(employeeEmail, subject, body);
  } catch(e) {
    // optional email - ignore errors
  }
}


// ════════════════════════════════════════════════════════════════
//  SHEET STYLING & CONVENIENCE HELPERS
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

function safeParseJSON(str) {
  try { return JSON.parse(str) || []; }
  catch(e) { return []; }
}

function formatDateString(str) {
  if (!str) return "";
  if (str instanceof Date) {
    return Utilities.formatDate(str, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = str.toString().trim();
  if (s.indexOf("T") > 0) {
    s = s.split("T")[0];
  }
  var match = s.match(/\d{4}-\d{2}-\d{2}/);
  if (match) {
    return match[0];
  }
  return s;
}

function createJSONResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
//  Fresh sheet setup script (Run once manually to bootstrap)
// ════════════════════════════════════════════════════════════════
function createFreshSheet() {
  var ss = SpreadsheetApp.create("Employee Daily Work Reports");
  setupSheets(ss);
  ss.getSheets().forEach(function(ws) {
    ws.autoResizeColumns(1, ws.getLastColumn());
  });
  return ss.getUrl();
}

// ════════════════════════════════════════════════════════════════
//  AUTOMATED DAILY COMPLETED TASKS EMAIL TRIGGER
// ════════════════════════════════════════════════════════════════

/**
 * Scheduled function to send a daily summary of completed tasks to the manager.
 */
function sendDailyCompletedTasksEmail() {
  try {
    var ss = getSpreadsheet();
    var tw = ss.getSheetByName(SHEET_NAME_TASKS);
    var data = tw.getDataRange().getValues();
    
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var completedTasks = [];

    // Columns: "Report ID", "Report Date", "Employee Name", "Employee ID", "Task Title", "Category", "Priority", "Status", "Estimated Hours", "Actual Hours", "Progress %", "Notes"
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][1];
      var status = data[i][7] ? data[i][7].toString().trim() : "";
      
      if (rowDate && status === "Completed") {
        var formattedRowDate = formatDateString(rowDate);
        if (formattedRowDate === todayStr) {
          completedTasks.push({
            employeeId: data[i][3],
            employeeName: data[i][2],
            title: data[i][4],
            category: data[i][5],
            priority: data[i][6],
            actual: data[i][9],
            notes: data[i][11]
          });
        }
      }
    }

    if (completedTasks.length === 0) {
      MailApp.sendEmail(
        "saurabh.yadav@gtftechnologies.com",
        "📋 Daily Completed Tasks Summary — " + todayStr,
        "No completed tasks were logged today."
      );
      return;
    }

    // Build HTML table for email
    var html = "<h3>Daily Summary of Completed Tasks — " + todayStr + "</h3>";
    html += "<p>Below is the summary of tasks completed by team members today:</p>";
    html += "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; font-family: sans-serif; font-size: 13px; border-color: #DADCE0;'>";
    html += "<tr style='background-color: #0F172A; color: white;'>";
    html += "<th>Employee</th>";
    html += "<th>Task Description</th>";
    html += "<th>Category</th>";
    html += "<th>Priority</th>";
    html += "<th>Actual Hours</th>";
    html += "<th>Notes</th>";
    html += "</tr>";

    completedTasks.forEach(function(t) {
      html += "<tr>";
      html += "<td><strong>" + t.employeeName + "</strong> (" + t.employeeId + ")</td>";
      html += "<td>" + t.title + "</td>";
      html += "<td>" + t.category + "</td>";
      html += "<td>" + t.priority + "</td>";
      html += "<td style='text-align: center;'>" + parseFloat(t.actual || 0).toFixed(1) + "</td>";
      html += "<td>" + (t.notes || "—") + "</td>";
      html += "</tr>";
    });

    html += "</table>";
    html += "<br><p>This is an automated system email.</p>";

    MailApp.sendEmail({
      to: "saurabh.yadav@gtftechnologies.com",
      subject: "✅ Daily Completed Tasks Summary — " + todayStr,
      htmlBody: html
    });
  } catch(e) {
    Logger.log("Error sending daily report: " + e.toString());
  }
}

/**
 * Run this function once manually in the script editor to set up the daily 8:00 PM email trigger.
 */
function createDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendDailyCompletedTasksEmail") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger("sendDailyCompletedTasksEmail")
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .nearMinute(0)
    .create();
}
