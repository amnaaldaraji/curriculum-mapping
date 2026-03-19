// ───────────────────────────────────────────────────────────────────────────────
// CURRICULUM MAPPING TOOL — Google Apps Script
// Paste this in your Google Sheet: Extensions → Apps Script → Save → Deploy
// ───────────────────────────────────────────────────────────────────────────────
const SHEET_NAME = "Responses";
const OUTCOMES = [
  "Describe organizational and bureaucratic structures involved in policy development",
  "Use knowledge and abilities to solve a problem in any context",
  "Develop ethically defensible solutions to issues",
  "Formulate strategies to implement new policies",
  "Effectively communicate ideas orally and in writing",
  "Work effectively as a member of a team",
];

// Column layout (1-indexed):
//  1  Timestamp
//  2  First Name
//  3  Last Name
//  4  Email
//  5  Title
//  6  Course Name
//  7  Course Code
//  8  Credits (hp)
//  9  Program
// 10  OS1 Level | 11 OS1 Activities | 12 OS1 Assessment
// 13  OS2 Level | 14 OS2 Activities | 15 OS2 Assessment
// ... (each outcome = 3 cols)
// 28  Indirect Measures

const INFO_COLS   = 9;
const COLS_PER_OS = 3; // Level, Activities, Assessment

function osLevelCol(i)      { return INFO_COLS + 1 + i * COLS_PER_OS; }       // col of OSi Level
function osActivitiesCol(i) { return INFO_COLS + 2 + i * COLS_PER_OS; }
function osAssessmentCol(i) { return INFO_COLS + 3 + i * COLS_PER_OS; }

// — Receives POST from the HTML form
function doPost(e) {
  try {
    const data = JSON.parse(e.parameter.data);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      setupHeader(sheet);
    } else if (sheet.getLastRow() === 0) {
      setupHeader(sheet);
    }
    const row = buildRow(data);
    sheet.appendRow(row);
    formatDataRow(sheet, sheet.getLastRow());
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput("Curriculum Mapping Tool - API active")
    .setMimeType(ContentService.MimeType.TEXT);
}

// — Build flat row: info cols, then for each outcome: Level / Activities / Assessment
function buildRow(data) {
  const row = [
    new Date(data.submitted),
    data.instructor.firstName,
    data.instructor.lastName,
    data.instructor.email,
    data.instructor.title,
    data.course.name,
    data.course.code,
    Number(data.course.hp) || data.course.hp,
    data.program,
  ];
  OUTCOMES.forEach((_, i) => {
    const m = data.mapping ? data.mapping[i] : null;
    row.push(m ? (m.level      || "") : "");
    row.push(m ? (m.activities || "") : "");
    row.push(m ? (m.assessment || "") : "");
  });
  row.push(data.indirect || "");
  return row;
}

// — Write and style the header row
function setupHeader(sheet) {
  const headers = [
    "Timestamp", "First Name", "Last Name", "Email", "Title",
    "Course Name", "Course Code", "Credits (hp)", "Program",
  ];
  OUTCOMES.forEach((_, i) => {
    headers.push("OS" + (i+1) + " Level");
    headers.push("OS" + (i+1) + " Activities");
    headers.push("OS" + (i+1) + " Assessment");
  });
  headers.push("Indirect Measures");
  sheet.appendRow(headers);

  const totalCols = headers.length;
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange
    .setBackground("#0B1F33")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(10)
    .setFontFamily("Arial")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrap(true);
  sheet.setRowHeight(1, 50);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  // Color each outcome group with alternating teal shades so groups are visually distinct
  const groupColors = ["#006D77", "#005962", "#006D77", "#005962", "#006D77", "#005962"];
  OUTCOMES.forEach((outcome, i) => {
    const lvlCol  = osLevelCol(i);
    const actCol  = osActivitiesCol(i);
    const assCol  = osAssessmentCol(i);
    const bg      = groupColors[i];
    sheet.getRange(1, lvlCol).setBackground(bg).setNote("Outcome " + (i+1) + ":\n" + outcome);
    sheet.getRange(1, actCol).setBackground(bg);
    sheet.getRange(1, assCol).setBackground(bg);
    // Add a light vertical border before each group to visually separate them
    sheet.getRange(1, lvlCol, sheet.getMaxRows(), 1)
      .setBorder(false, true, false, false, false, false, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  // — Column widths
  sheet.setColumnWidth(1,  110);  // Timestamp
  sheet.setColumnWidth(2,  85);   // First Name
  sheet.setColumnWidth(3,  95);   // Last Name
  sheet.setColumnWidth(4,  170);  // Email
  sheet.setColumnWidth(5,  120);  // Title
  sheet.setColumnWidth(6,  160);  // Course Name
  sheet.setColumnWidth(7,  80);   // Course Code
  sheet.setColumnWidth(8,  65);   // Credits
  sheet.setColumnWidth(9,  160);  // Program

  OUTCOMES.forEach((_, i) => {
    sheet.setColumnWidth(osLevelCol(i),      55);   // Level — narrow, just I/A/M
    sheet.setColumnWidth(osActivitiesCol(i), 140);  // Activities
    sheet.setColumnWidth(osAssessmentCol(i), 140);  // Assessment
  });

  // Indirect Measures
  sheet.setColumnWidth(INFO_COLS + OUTCOMES.length * COLS_PER_OS + 1, 180);

  // Legend on A1
  sheet.getRange(1, 1).setNote(
    "I = Introductory: recall or explain basic concepts\n" +
    "A = Advanced: apply procedures or analyse\n" +
    "M = Mastery: evaluate evidence or create novel work\n\n" +
    "Each OS group: Level | Activities | Assessment\n" +
    "Hover the Level header to see the full outcome text."
  );
}

// — Format each new data row
function formatDataRow(sheet, rowIndex) {
  const lastCol = sheet.getLastColumn();
  const fullRow = sheet.getRange(rowIndex, 1, 1, lastCol);
  fullRow
    .setFontSize(10)
    .setFontFamily("Arial")
    .setVerticalAlignment("middle")
    .setWrap(true);
  sheet.setRowHeight(rowIndex, 36);

  const bgColor = (rowIndex % 2 === 0) ? "#F7F8FB" : "#FFFFFF";
  fullRow.setBackground(bgColor);

  OUTCOMES.forEach((_, i) => {
    const cell = sheet.getRange(rowIndex, osLevelCol(i));
    const val  = cell.getValue();
    cell.setHorizontalAlignment("center").setFontWeight("bold").setFontSize(12);
    if      (val === "I") { cell.setBackground("#E6F4F5").setFontColor("#006D77"); }
    else if (val === "A") { cell.setBackground("#FBF3E4").setFontColor("#B5985A"); }
    else if (val === "M") { cell.setBackground("#E8F0F8").setFontColor("#2C5F8A"); }
    else                  { cell.setBackground("#F0F2F5").setFontColor("#BBBBBB"); }
  });

  fullRow.setBorder(
    false, false, true, false, false, false,
    "#DDE2EA", SpreadsheetApp.BorderStyle.SOLID
  );
}

// — Run this ONCE manually after pasting to reformat all existing rows
function reformatAllRows() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    formatDataRow(sheet, r);
  }
  SpreadsheetApp.getUi().alert("All rows reformatted!");
}
