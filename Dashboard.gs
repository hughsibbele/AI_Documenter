/**
 * Dashboard.gs
 * Interactive dashboard: student view and assignment view.
 * Data is written to the Dashboard sheet tab, driven by dropdown selections.
 */

/**
 * Initializes the Dashboard sheet with structure and dropdowns.
 */
function initializeDashboardSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Dashboard');
  if (!sheet) {
    sheet = ss.insertSheet('Dashboard');
  }

  // Clear everything
  sheet.clear();
  sheet.clearFormats();

  // Row 1: View type
  sheet.getRange('A1').setValue('View').setFontWeight('bold').setFontColor('#5f6368');
  var viewRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Student View', 'Assignment View'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B1').setDataValidation(viewRule).setValue('Student View');

  // Row 2: Course selector
  sheet.getRange('A2').setValue('Course').setFontWeight('bold').setFontColor('#5f6368');
  populateDashboardCourseDropdown();

  // Row 3: Select (student or assignment)
  sheet.getRange('A3').setValue('Select').setFontWeight('bold').setFontColor('#5f6368');
  sheet.getRange('B3').setValue('');

  // Formatting
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidths(3, 4, 150);

  // Add instruction
  sheet.getRange('A5').setValue('Select a view, course, and student/assignment above to see data.')
    .setFontColor('#80868b').setFontStyle('italic');
}

/**
 * Populates the Course dropdown on the Dashboard from active courses.
 */
function populateDashboardCourseDropdown() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Dashboard');
  if (!sheet) return;

  var courses = getActiveCourses();
  if (courses.length === 0) return;

  var courseNames = courses.map(function(c) { return c.name; });
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(courseNames)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B2').setDataValidation(rule);
  if (!sheet.getRange('B2').getValue()) {
    sheet.getRange('B2').setValue(courseNames[0]);
  }
  populateDashboardSelectDropdown();
}

/**
 * Populates the Select dropdown based on view type and course.
 */
function populateDashboardSelectDropdown() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Dashboard');
  if (!sheet) return;

  var viewType = sheet.getRange('B1').getValue();
  var courseName = sheet.getRange('B2').getValue();
  if (!courseName) return;

  var items = [];
  if (viewType === 'Student View') {
    items = getUniqueStudents_(courseName);
  } else {
    items = getUniqueAssignments_(courseName);
  }

  if (items.length === 0) {
    sheet.getRange('B3').clearDataValidations().setValue('(no data yet)');
    return;
  }

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(items)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B3').setDataValidation(rule).setValue(items[0]);
}

/**
 * Main dashboard refresh — reads dropdowns and populates data area.
 */
function refreshDashboard() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Dashboard');
  if (!sheet) {
    initializeDashboardSheet();
    return;
  }

  var viewType = sheet.getRange('B1').getValue();
  var courseName = sheet.getRange('B2').getValue();
  var selection = sheet.getRange('B3').getValue();

  if (!viewType || !courseName || !selection || selection === '(no data yet)') return;

  // Clear data area (row 5+)
  var lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    sheet.getRange(5, 1, lastRow - 4, 8).clear();
  }

  if (viewType === 'Student View') {
    renderStudentView_(sheet, courseName, selection);
  } else {
    renderAssignmentView_(sheet, courseName, selection);
  }
}

// ============================================================
// STUDENT VIEW
// ============================================================

function renderStudentView_(sheet, courseName, studentName) {
  var submissions = getSubmissions(courseName).filter(function(s) {
    return s['Student Name'] === studentName;
  });

  if (submissions.length === 0) {
    sheet.getRange('A5').setValue('No submissions found for this student.').setFontColor('#80868b');
    return;
  }

  // Summary stats
  var totalTime = submissions.reduce(function(sum, s) { return sum + (Number(s['Time (min)']) || 0); }, 0);
  var tools = {};
  submissions.forEach(function(s) {
    var tool = s['AI Tool'] || 'Unknown';
    tools[tool] = (tools[tool] || 0) + 1;
  });
  var toolList = Object.keys(tools).map(function(t) { return t + ' (' + tools[t] + ')'; }).join(', ');

  // Row 5: Summary header
  sheet.getRange('A5').setValue('Student:').setFontWeight('bold');
  sheet.getRange('B5').setValue(studentName);
  sheet.getRange('C5').setValue('Submissions: ' + submissions.length).setFontColor('#5f6368');
  sheet.getRange('D5').setValue('Total AI Time: ' + totalTime + ' min').setFontColor('#5f6368');

  // Row 6: Tools
  sheet.getRange('A6').setValue('Tools Used:').setFontWeight('bold');
  sheet.getRange('B6').setValue(toolList).setFontColor('#5f6368');

  // Row 8: Headers
  var headers = ['Assignment', 'Due Date', 'Tool', 'Time (min)', 'AI Use Summary', 'Reflection (excerpt)'];
  sheet.getRange(8, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');

  // Row 9+: Data rows
  var rows = submissions.map(function(s) {
    var reflection = String(s['Reflection'] || '');
    if (reflection.length > 100) reflection = reflection.substring(0, 100) + '...';
    return [
      s['Assignment'] || '',
      s['Due Date'] || '',
      s['AI Tool'] || '',
      s['Time (min)'] || '',
      s['AI Use Summary'] || '',
      reflection
    ];
  });

  // Sort by due date
  rows.sort(function(a, b) { return String(a[1]).localeCompare(String(b[1])); });

  if (rows.length > 0) {
    sheet.getRange(9, 1, rows.length, headers.length).setValues(rows);

    // Alternating row colors
    for (var i = 0; i < rows.length; i++) {
      if (i % 2 === 1) {
        sheet.getRange(9 + i, 1, 1, headers.length).setBackground('#f8f9fa');
      }
    }
  }

  // Auto-resize
  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }
}

// ============================================================
// ASSIGNMENT VIEW
// ============================================================

function renderAssignmentView_(sheet, courseName, assignmentName) {
  var submissions = getSubmissions(courseName).filter(function(s) {
    return s['Assignment'] === assignmentName;
  });

  if (submissions.length === 0) {
    sheet.getRange('A5').setValue('No submissions found for this assignment.').setFontColor('#80868b');
    return;
  }

  // Summary stats
  var totalTime = submissions.reduce(function(sum, s) { return sum + (Number(s['Time (min)']) || 0); }, 0);
  var avgTime = Math.round(totalTime / submissions.length);
  var tools = {};
  submissions.forEach(function(s) {
    var tool = s['AI Tool'] || 'Unknown';
    tools[tool] = (tools[tool] || 0) + 1;
  });
  var toolDist = Object.keys(tools).map(function(t) { return t + ': ' + tools[t]; }).join(', ');

  // Get due date from first submission
  var dueDate = submissions[0]['Due Date'] || '';

  // Row 5: Summary header
  sheet.getRange('A5').setValue('Assignment:').setFontWeight('bold');
  sheet.getRange('B5').setValue(assignmentName);
  sheet.getRange('C5').setValue('Due: ' + dueDate).setFontColor('#5f6368');
  sheet.getRange('D5').setValue('Submissions: ' + submissions.length).setFontColor('#5f6368');

  // Row 6: Stats
  sheet.getRange('A6').setValue('Avg AI Time:').setFontWeight('bold');
  sheet.getRange('B6').setValue(avgTime + ' min').setFontColor('#5f6368');
  sheet.getRange('C6').setValue('Tools: ' + toolDist).setFontColor('#5f6368');

  // Row 8: Headers
  var headers = ['Student', 'Tool', 'Time (min)', 'AI Use Summary', 'Reflection (excerpt)'];
  sheet.getRange(8, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#f1f3f4');

  // Row 9+: Data rows
  var rows = submissions.map(function(s) {
    var reflection = String(s['Reflection'] || '');
    if (reflection.length > 100) reflection = reflection.substring(0, 100) + '...';
    return [
      s['Student Name'] || '',
      s['AI Tool'] || '',
      s['Time (min)'] || '',
      s['AI Use Summary'] || '',
      reflection
    ];
  });

  // Sort by student name
  rows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });

  if (rows.length > 0) {
    sheet.getRange(9, 1, rows.length, headers.length).setValues(rows);

    // Alternating row colors
    for (var i = 0; i < rows.length; i++) {
      if (i % 2 === 1) {
        sheet.getRange(9 + i, 1, 1, headers.length).setBackground('#f8f9fa');
      }
    }
  }

  // Auto-resize
  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }
}

// ============================================================
// HELPERS
// ============================================================

/**
 * Returns unique student names from a course tab, sorted.
 */
function getUniqueStudents_(courseName) {
  var submissions = getSubmissions(courseName);
  var names = {};
  submissions.forEach(function(s) {
    if (s['Student Name']) names[s['Student Name']] = true;
  });
  return Object.keys(names).sort();
}

/**
 * Returns unique assignment names from a course tab, sorted.
 */
function getUniqueAssignments_(courseName) {
  var submissions = getSubmissions(courseName);
  var assignments = {};
  submissions.forEach(function(s) {
    if (s['Assignment']) assignments[s['Assignment']] = true;
  });
  return Object.keys(assignments).sort();
}
