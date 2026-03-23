/**
 * Canvas.gs
 * Canvas LMS API integration: fetch courses, fetch assignments, sync to cache.
 */

/**
 * Builds the Canvas API base URL from config.
 */
function getCanvasBaseUrl() {
  var url = getConfigValue('Canvas URL');
  if (!url) throw new Error('Canvas URL not configured. Run Setup Wizard first.');
  // Ensure https:// prefix and /api/v1/ suffix
  if (url.indexOf('http') !== 0) url = 'https://' + url;
  if (url.charAt(url.length - 1) !== '/') url += '/';
  if (url.indexOf('/api/v1/') === -1) url += 'api/v1/';
  return url;
}

/**
 * Returns the Canvas API token from config.
 */
function getCanvasToken() {
  var token = getConfigValue('Canvas API Token');
  if (!token) throw new Error('Canvas API Token not configured. Run Setup Wizard first.');
  return token;
}

/**
 * Makes an authenticated GET request to the Canvas API.
 * Handles pagination automatically, returning all results.
 */
function canvasApiGet(endpoint) {
  var baseUrl = getCanvasBaseUrl();
  var token = getCanvasToken();
  var allResults = [];
  var url = baseUrl + endpoint;

  // Add per_page if not already in the URL
  if (url.indexOf('per_page') === -1) {
    url += (url.indexOf('?') === -1 ? '?' : '&') + 'per_page=100';
  }

  while (url) {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code !== 200) {
      throw new Error('Canvas API error (' + code + '): ' + response.getContentText().substring(0, 200));
    }

    var results = JSON.parse(response.getContentText());
    allResults = allResults.concat(results);

    // Check for pagination via Link header
    url = getNextPageUrl_(response);
  }

  return allResults;
}

/**
 * Parses the Link header to find the next page URL.
 */
function getNextPageUrl_(response) {
  var linkHeader = response.getHeaders()['Link'] || response.getHeaders()['link'];
  if (!linkHeader) return null;

  var links = linkHeader.split(',');
  for (var i = 0; i < links.length; i++) {
    var parts = links[i].split(';');
    if (parts.length === 2 && parts[1].trim() === 'rel="next"') {
      return parts[0].trim().replace(/^<|>$/g, '');
    }
  }
  return null;
}

/**
 * Fetches all courses where the user is a teacher.
 * Returns [{id, name}]
 */
function fetchCourses() {
  var courses = canvasApiGet('courses?enrollment_type=teacher&enrollment_state=active&state[]=available');
  return courses.map(function(c) {
    return {
      id: String(c.id),
      name: c.name || c.course_code || ('Course ' + c.id)
    };
  });
}

/**
 * Fetches published assignments for a specific course.
 * Returns [{id, name, dueDate}]
 */
function fetchAssignments(courseId) {
  var assignments = canvasApiGet('courses/' + courseId + '/assignments?order_by=due_at');
  return assignments
    .filter(function(a) { return a.published === true; })
    .map(function(a) {
      return {
        assignmentId: String(a.id),
        name: a.name,
        dueDate: a.due_at ? formatCanvasDate_(a.due_at) : 'No due date'
      };
    });
}

/**
 * Formats a Canvas ISO date string to a readable format (e.g., "Mar 25").
 */
function formatCanvasDate_(isoString) {
  if (!isoString) return 'No due date';
  var date = new Date(isoString);
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return months[date.getMonth()] + ' ' + date.getDate();
}

/**
 * Syncs assignments from Canvas for all active courses into the _Assignments cache.
 */
function syncAllAssignments() {
  var courses = getActiveCourses();
  if (courses.length === 0) {
    SpreadsheetApp.getUi().alert('No active courses configured. Run Setup Wizard first.');
    return;
  }

  var allAssignments = [];
  var errors = [];

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
      errors.push(courses[i].name + ': ' + err.message);
    }
  }

  writeAssignmentsCache(allAssignments);

  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Synced ' + allAssignments.length + ' assignments, but had errors:\n\n' + errors.join('\n')
    );
  } else {
    SpreadsheetApp.getUi().alert('Synced ' + allAssignments.length + ' assignments from ' + courses.length + ' course(s).');
  }
}

/**
 * Tests the Canvas connection with the configured URL and token.
 * Returns {success: boolean, message: string, courses: [{id, name}]}
 */
function testCanvasConnection(url, token) {
  try {
    // Temporarily use provided values (not yet saved to config)
    var testUrl = url;
    if (testUrl.indexOf('http') !== 0) testUrl = 'https://' + testUrl;
    if (testUrl.charAt(testUrl.length - 1) !== '/') testUrl += '/';
    if (testUrl.indexOf('/api/v1/') === -1) testUrl += 'api/v1/';

    var response = UrlFetchApp.fetch(
      testUrl + 'courses?enrollment_type=teacher&enrollment_state=active&state[]=available&per_page=100',
      {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );

    var code = response.getResponseCode();
    if (code !== 200) {
      return { success: false, message: 'API returned status ' + code + '. Check your URL and token.' };
    }

    var courses = JSON.parse(response.getContentText());
    return {
      success: true,
      message: 'Connected! Found ' + courses.length + ' course(s).',
      courses: courses.map(function(c) {
        return { id: String(c.id), name: c.name || c.course_code || ('Course ' + c.id) };
      })
    };
  } catch (err) {
    return { success: false, message: 'Connection failed: ' + err.message };
  }
}

/**
 * Sets up a daily trigger to sync assignments from Canvas.
 */
function setupCanvasSyncTrigger() {
  // Remove any existing sync triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncAllAssignmentsSilent') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create new daily trigger at 5 AM
  ScriptApp.newTrigger('syncAllAssignmentsSilent')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();
}

/**
 * Silent version of syncAllAssignments (no UI alerts) for use with triggers.
 */
function syncAllAssignmentsSilent() {
  var courses = getActiveCourses();
  if (courses.length === 0) return;

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
}
