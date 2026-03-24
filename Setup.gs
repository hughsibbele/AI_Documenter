/**
 * Setup.gs
 * Setup wizard backend functions.
 */

/**
 * Saves Canvas configuration and returns test results.
 */
function setupSaveCanvas(canvasUrl, apiToken) {
  var result = testCanvasConnection(canvasUrl, apiToken);
  if (result.success) {
    setConfigValue('Canvas URL', canvasUrl);
    setConfigValue('Canvas API Token', apiToken);
  }
  return result;
}

/**
 * Saves selected courses to Config and creates their data tabs.
 * courses: [{id, name}]
 */
function setupSaveCourses(courses) {
  var courseData = courses.map(function(c) {
    return { id: c.id, name: c.name, active: true };
  });
  setConfigCourses(courseData);

  // Create a data tab for each course
  for (var i = 0; i < courses.length; i++) {
    getOrCreateCourseSheet(courses[i].name);
  }

  return { success: true, message: 'Created ' + courses.length + ' course tab(s).' };
}

/**
 * Tests the Gemini API connection.
 */
function setupTestGemini(apiKey, model) {
  try {
    var modelName = model || getConfigValue('Gemini Model') || 'gemini-2.0-flash';
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + modelName + ':generateContent?key=' + apiKey;
    var payload = {
      contents: [{ parts: [{ text: 'Say "hello" and nothing else.' }] }]
    };
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code !== 200) {
      return { success: false, message: 'Gemini API returned status ' + code + '. Check your API key.' };
    }

    return { success: true, message: 'Gemini connection successful!' };
  } catch (err) {
    return { success: false, message: 'Connection failed: ' + err.message };
  }
}

/**
 * Saves Gemini configuration.
 */
function setupSaveGemini(apiKey, model) {
  setConfigValue('Gemini API Key', apiKey);
  if (model) setConfigValue('Gemini Model', model);
  return { success: true };
}

/**
 * Saves AI tools list to Config.
 */
function setupSaveTools(toolsString) {
  setConfigValue('AI Tools', toolsString);
  return { success: true };
}

/**
 * Final setup step: sync assignments and set up triggers.
 */
function setupFinalize() {
  try {
    // Sync assignments from Canvas
    var courses = getActiveCourses();
    var allAssignments = [];
    for (var i = 0; i < courses.length; i++) {
      try {
        var assignments = fetchAssignments(courses[i].id);
        for (var j = 0; j < assignments.length; j++) {
          allAssignments.push({
            courseId: courses[i].id,
            courseName: courses[i].name,
            assignmentId: assignments[j].assignmentId,
            name: assignments[j].name,
            dueDate: assignments[j].dueDate
          });
        }
      } catch (err) {
        Logger.log('Sync error for ' + courses[i].name + ': ' + err.message);
      }
    }
    writeAssignmentsCache(allAssignments);

    // Set up daily Canvas sync trigger
    setupCanvasSyncTrigger();

    // Initialize Dashboard tab
    initializeDashboardSheet();

    // Organize and color-code tabs
    organizeTabs_(courses);

    return {
      success: true,
      message: 'Setup complete! Synced ' + allAssignments.length + ' assignments from ' +
               courses.length + ' course(s). Daily sync enabled.\n\n' +
               'Next step: Deploy this script as a web app.\n' +
               '1. Click Deploy → New deployment\n' +
               '2. Select type: Web app\n' +
               '3. Set "Who has access" to "Anyone"\n' +
               '4. Click Deploy and copy the URL\n' +
               '5. Share the URL with your students!'
    };
  } catch (err) {
    return { success: false, message: 'Finalization error: ' + err.message };
  }
}

/**
 * Menu-callable version: reorganizes and color-codes all tabs
 * without re-running the full setup wizard.
 */
function reformatTabs() {
  var courses = getActiveCourses();
  organizeTabs_(courses);
  SpreadsheetApp.getUi().alert('Tabs reorganized and color-coded.');
}

/**
 * Organizes tab order and applies color coding after setup.
 * Order: Dashboard, course tabs (alphabetical), Config, _Assignments
 * Removes the default "Sheet1" tab if it's still blank.
 */
function organizeTabs_(courses) {
  var ss = getSpreadsheet();

  // Remove default "Sheet1" if it exists and is empty
  var sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() <= 1 && sheet1.getLastColumn() <= 1) {
    ss.deleteSheet(sheet1);
  }

  // Color coding
  var dashboard = ss.getSheetByName('Dashboard');
  if (dashboard) dashboard.setTabColor('#1a73e8'); // blue

  var courseNames = courses.map(function(c) { return c.name; }).sort();
  for (var i = 0; i < courseNames.length; i++) {
    var sheet = ss.getSheetByName(courseNames[i]);
    if (sheet) sheet.setTabColor('#34a853'); // green
  }

  var config = ss.getSheetByName('Config');
  if (config) config.setTabColor('#80868b'); // gray

  var assignments = ss.getSheetByName('_Assignments');
  if (assignments) assignments.setTabColor('#80868b'); // gray

  // Reorder tabs: Dashboard first, then courses alphabetically, then Config, then _Assignments
  var position = 0;
  if (dashboard) { ss.setActiveSheet(dashboard); ss.moveActiveSheet(position + 1); position++; }
  for (var i = 0; i < courseNames.length; i++) {
    var sheet = ss.getSheetByName(courseNames[i]);
    if (sheet) { ss.setActiveSheet(sheet); ss.moveActiveSheet(position + 1); position++; }
  }
  if (config) { ss.setActiveSheet(config); ss.moveActiveSheet(position + 1); position++; }
  if (assignments) { ss.setActiveSheet(assignments); ss.moveActiveSheet(position + 1); position++; }

  // Return focus to Dashboard
  if (dashboard) ss.setActiveSheet(dashboard);
}
