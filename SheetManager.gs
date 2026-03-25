/**
 * SheetManager.gs
 * Handles all Google Sheet read/write operations, config management, and tab creation.
 */

// ============================================================
// CONFIG OPERATIONS
// ============================================================

/**
 * Returns the active spreadsheet (cached per execution).
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Gets or creates the Config sheet. Returns the sheet object.
 */
function getConfigSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    initializeConfigSheet_(sheet);
  }
  return sheet;
}

/**
 * Initializes the Config sheet with default structure.
 */
function initializeConfigSheet_(sheet) {
  var headers = [
    ['Setting', 'Value'],
    ['Canvas URL', ''],
    ['Canvas API Token', ''],
    ['Gemini API Key', ''],
    ['Gemini Model', 'gemini-2.0-flash'],
    ['Gemini Thinking Level', ''],
    ['Gemini Prompt', getDefaultGeminiPrompt()],
    ['AI Tools', 'ChekhovBot, GothBot, Gemini, ChatGPT, Claude, Other'],
    [''],
    ['Course ID', 'Course Name', 'Active']
  ];
  sheet.getRange(1, 1, headers.length, 3).setValues(
    headers.map(function(row) {
      while (row.length < 3) row.push('');
      return row;
    })
  );
  // Bold header row
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  sheet.getRange(9, 1, 1, 3).setFontWeight('bold');
  // Set column widths
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 500);
}

/**
 * Reads a config value by setting name (e.g., "Canvas URL").
 */
function getConfigValue(settingName) {
  var sheet = getConfigSheet();
  var data = sheet.getRange(1, 1, 8, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === settingName) {
      return data[i][1];
    }
  }
  return '';
}

/**
 * Writes a config value by setting name.
 */
function setConfigValue(settingName, value) {
  var sheet = getConfigSheet();
  var data = sheet.getRange(1, 1, 8, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === settingName) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
}

/**
 * Returns the list of AI tools from Config as an array.
 */
function getAITools() {
  var toolString = getConfigValue('AI Tools');
  if (!toolString) return ['ChekhovBot', 'GothBot', 'Gemini', 'ChatGPT', 'Claude', 'Other'];
  return toolString.split(',').map(function(t) { return t.trim(); });
}

/**
 * Returns all courses from Config (rows 10+).
 * Each course: {id, name, active}
 */
function getConfigCourses() {
  var sheet = getConfigSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 10) return [];
  var data = sheet.getRange(10, 1, lastRow - 9, 3).getValues();
  return data
    .filter(function(row) { return row[0] !== ''; })
    .map(function(row) {
      return {
        id: String(row[0]),
        name: String(row[1]),
        active: row[2] === true || row[2] === 'TRUE' || row[2] === true
      };
    });
}

/**
 * Returns only active courses.
 */
function getActiveCourses() {
  return getConfigCourses().filter(function(c) { return c.active; });
}

/**
 * Writes courses to Config sheet (starting at row 10).
 * courses: [{id, name, active}]
 */
function setConfigCourses(courses) {
  var sheet = getConfigSheet();
  // Clear existing course rows
  var lastRow = sheet.getLastRow();
  if (lastRow >= 10) {
    sheet.getRange(10, 1, lastRow - 9, 3).clearContent();
  }
  if (courses.length === 0) return;
  var data = courses.map(function(c) {
    return [c.id, c.name, c.active ? 'TRUE' : 'FALSE'];
  });
  sheet.getRange(10, 1, data.length, 3).setValues(data);
}

// ============================================================
// COURSE TAB OPERATIONS
// ============================================================

/**
 * Column headers for course data tabs.
 */
var COURSE_HEADERS = [
  'Timestamp', 'Student Name', 'Student Email', 'Assignment',
  'Due Date', 'AI Tool', 'Other Tool', 'Time (min)',
  'Reflection', 'Raw Transcript', 'Cleaned Transcript',
  'AI Use Summary', 'Status'
];

/**
 * Creates a course data tab if it doesn't exist. Returns the sheet.
 */
function getOrCreateCourseSheet(courseName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(courseName);
  if (!sheet) {
    sheet = ss.insertSheet(courseName);
    initializeCourseSheet_(sheet);
  }
  return sheet;
}

/**
 * Initializes a course data tab with headers and formatting.
 */
function initializeCourseSheet_(sheet) {
  sheet.getRange(1, 1, 1, COURSE_HEADERS.length).setValues([COURSE_HEADERS]);
  sheet.getRange(1, 1, 1, COURSE_HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  // Set reasonable column widths
  var widths = [160, 150, 200, 200, 100, 120, 120, 80, 300, 300, 300, 200, 80];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}

/**
 * Appends a submission row to a course tab. Returns the row number.
 */
function appendSubmission(courseName, rowData) {
  var sheet = getOrCreateCourseSheet(courseName);
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  // Cap row height at 2 lines (42px) so transcripts don't blow up the sheet
  sheet.setRowHeight(newRow, 42);
  return newRow;
}

/**
 * Updates Gemini processing results for a specific row.
 * Bolds speaker tags (Student:, AI:) in the cleaned transcript using RichText.
 */
function updateProcessingResults(courseName, rowNumber, cleanedTranscript, summary, status) {
  var sheet = getOrCreateCourseSheet(courseName);

  // Write cleaned transcript with bold speaker tags
  if (cleanedTranscript) {
    var richText = buildBoldSpeakerTags_(cleanedTranscript);
    sheet.getRange(rowNumber, 11).setRichTextValue(richText);
  } else {
    sheet.getRange(rowNumber, 11).setValue('');
  }

  // Columns L (12), M (13)
  sheet.getRange(rowNumber, 12).setValue(summary);
  sheet.getRange(rowNumber, 13).setValue(status);

  // Ensure row stays capped at 2 lines
  sheet.setRowHeight(rowNumber, 42);
}

/**
 * Builds a RichTextValue with bold "Student:" and "AI:" speaker tags.
 */
function buildBoldSpeakerTags_(text) {
  var builder = SpreadsheetApp.newRichTextValue().setText(text);
  var boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();

  var patterns = ['Student:', 'AI:'];
  for (var p = 0; p < patterns.length; p++) {
    var tag = patterns[p];
    var startIdx = 0;
    while (true) {
      var idx = text.indexOf(tag, startIdx);
      if (idx === -1) break;
      builder.setTextStyle(idx, idx + tag.length, boldStyle);
      startIdx = idx + tag.length;
    }
  }

  return builder.build();
}

/**
 * Returns all submission data for a course tab.
 * Returns array of objects with column names as keys.
 */
function getSubmissions(courseName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(courseName);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, COURSE_HEADERS.length).getValues();
  return data.map(function(row, index) {
    var obj = {};
    for (var i = 0; i < COURSE_HEADERS.length; i++) {
      obj[COURSE_HEADERS[i]] = row[i];
    }
    obj._rowNumber = index + 2;
    return obj;
  });
}

// ============================================================
// ASSIGNMENTS CACHE TAB
// ============================================================

var ASSIGNMENTS_HEADERS = ['Course ID', 'Course Name', 'Assignment ID', 'Assignment Name', 'Due Date'];

/**
 * Gets or creates the hidden _Assignments cache tab.
 */
function getAssignmentsSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('_Assignments');
  if (!sheet) {
    sheet = ss.insertSheet('_Assignments');
    sheet.getRange(1, 1, 1, ASSIGNMENTS_HEADERS.length).setValues([ASSIGNMENTS_HEADERS]);
    sheet.getRange(1, 1, 1, ASSIGNMENTS_HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Replaces all cached assignments with new data.
 * assignments: [{courseId, courseName, assignmentId, name, dueDate}]
 */
function writeAssignmentsCache(assignments) {
  var sheet = getAssignmentsSheet();
  // Clear existing data (keep headers)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, ASSIGNMENTS_HEADERS.length).clearContent();
  }
  if (assignments.length === 0) return;
  var data = assignments.map(function(a) {
    return [a.courseId, a.courseName, a.assignmentId, a.name, a.dueDate || ''];
  });
  sheet.getRange(2, 1, data.length, ASSIGNMENTS_HEADERS.length).setValues(data);
}

/**
 * Returns cached assignments, optionally filtered by courseId.
 */
function getCachedAssignments(courseId) {
  var sheet = getAssignmentsSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, ASSIGNMENTS_HEADERS.length).getValues();
  var assignments = data
    .filter(function(row) { return row[0] !== ''; })
    .map(function(row) {
      return {
        courseId: String(row[0]),
        courseName: String(row[1]),
        assignmentId: String(row[2]),
        name: String(row[3]),
        dueDate: row[4]
      };
    });
  if (courseId) {
    assignments = assignments.filter(function(a) { return a.courseId === String(courseId); });
  }
  return assignments;
}

// ============================================================
// DEFAULT GEMINI PROMPT
// ============================================================

function getDefaultGeminiPrompt() {
  return 'You will receive a student\'s chat transcript with an AI tool. Produce two outputs:\n\n' +
    '1. CLEANED TRANSCRIPT: Normalize the transcript with clear "Student:" and "AI:" speaker labels. ' +
    'Remove formatting artifacts. Preserve the substantive content faithfully.\n\n' +
    '2. AI USE SUMMARY: A comma-separated list of max 6 words total describing what the student did ' +
    'with the AI. Examples: "brainstorming, drafting, revision" or "source analysis, outlining" or ' +
    '"proofreading, idea generation"\n\n' +
    'Format your response exactly as:\n' +
    '===CLEANED===\n[cleaned transcript here]\n===SUMMARY===\n[summary here]';
}
