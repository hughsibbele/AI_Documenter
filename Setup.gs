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
function setupTestGemini(apiKey) {
  try {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;
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
function setupSaveGemini(apiKey) {
  setConfigValue('Gemini API Key', apiKey);
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
