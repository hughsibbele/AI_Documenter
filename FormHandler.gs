/**
 * FormHandler.gs
 * Backend for the student-facing form: getCourses, getAssignments, submitForm.
 */

/**
 * Returns active courses for the form dropdown.
 * Called by the form frontend via google.script.run.
 */
function getCourses() {
  var courses = getActiveCourses();
  return courses.map(function(c) {
    return { id: c.id, name: c.name };
  });
}

/**
 * Returns assignments for a course, formatted for the form dropdown.
 * Falls back to live Canvas fetch if cache is empty.
 * Called by the form frontend via google.script.run.
 */
function getAssignmentsForCourse(courseId) {
  var assignments = getCachedAssignments(courseId);

  // Fallback: if cache is empty for this course, try fetching live from Canvas
  if (assignments.length === 0) {
    Logger.log('Assignment cache empty for courseId: ' + courseId + '. Fetching live from Canvas.');
    try {
      assignments = fetchAssignments(courseId).map(function(a) {
        return {
          courseId: String(courseId),
          courseName: '',
          assignmentId: a.assignmentId,
          name: a.name,
          dueDate: a.dueDate
        };
      });
    } catch (err) {
      Logger.log('Live Canvas fetch failed: ' + err.message);
    }
  }

  return assignments.map(function(a) {
    var label = a.name;
    if (a.dueDate && a.dueDate !== 'No due date') {
      label += ' (Due: ' + a.dueDate + ')';
    }
    return {
      name: a.name,
      dueDate: a.dueDate,
      label: label
    };
  });
}

/**
 * Returns the list of AI tools for the form dropdown.
 */
function getFormTools() {
  return getAITools();
}

/**
 * Handles form submission from the student form.
 * formData: {name, email, courseId, courseName, assignment, dueDate, tool, otherTool, time, reflection, transcript}
 */
function submitForm(formData) {
  try {
    // Validate required fields
    var required = ['name', 'email', 'courseName', 'assignment', 'tool', 'time', 'reflection', 'transcript'];
    for (var i = 0; i < required.length; i++) {
      if (!formData[required[i]] || String(formData[required[i]]).trim() === '') {
        return { success: false, message: 'Missing required field: ' + required[i] };
      }
    }

    var timestamp = new Date();
    var courseName = formData.courseName;

    // Build the row data (matches COURSE_HEADERS order)
    var rowData = [
      timestamp,                                    // A: Timestamp
      formData.name.trim(),                         // B: Student Name
      formData.email.trim(),                        // C: Student Email
      formData.assignment,                          // D: Assignment
      formData.dueDate || '',                       // E: Due Date
      formData.tool,                                // F: AI Tool
      formData.tool === 'Other' ? (formData.otherTool || '').trim() : '',  // G: Other Tool
      parseInt(formData.time, 10) || 0,             // H: Time (min)
      formData.reflection.trim(),                   // I: Reflection
      formData.transcript.trim(),                   // J: Raw Transcript
      '',                                           // K: Cleaned Transcript (Gemini)
      '',                                           // L: AI Use Summary (Gemini)
      'pending'                                     // M: Processing Status
    ];

    // Append to the correct course tab
    var rowNumber = appendSubmission(courseName, rowData);

    // Process with Gemini immediately
    try {
      processSubmission(courseName, rowNumber);
    } catch (geminiErr) {
      // If Gemini fails, the submission is still saved — just log the error
      Logger.log('Gemini processing failed for row ' + rowNumber + ': ' + geminiErr.message);
      updateProcessingResults(courseName, rowNumber, '', '', 'error');
    }

    return { success: true, message: 'Your AI use documentation has been submitted. Thank you!' };
  } catch (err) {
    Logger.log('submitForm error: ' + err.message);
    return { success: false, message: 'Submission error: ' + err.message };
  }
}
