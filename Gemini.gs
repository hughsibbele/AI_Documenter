/**
 * Gemini.gs
 * Gemini API integration for transcript processing.
 */

/**
 * Calls the Gemini API with the given prompt and content.
 * Returns the raw text response.
 */
function callGemini(prompt, content) {
  var apiKey = getConfigValue('Gemini API Key');
  if (!apiKey) throw new Error('Gemini API Key not configured.');

  var model = getConfigValue('Gemini Model') || 'gemini-2.0-flash';
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model +
            ':generateContent?key=' + apiKey;

  var generationConfig = {
    temperature: 0.3,
    maxOutputTokens: 8192
  };

  // Add thinking config for models that support it (Gemini 2.5+/3+)
  var thinkingLevel = getConfigValue('Gemini Thinking Level');
  if (thinkingLevel) {
    generationConfig.thinkingConfig = { thinkingLevel: thinkingLevel };
  }

  var payload = {
    contents: [{
      parts: [{ text: prompt + '\n\n--- STUDENT TRANSCRIPT ---\n\n' + content }]
    }],
    generationConfig: generationConfig
  };

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    var errorText = response.getContentText().substring(0, 300);
    throw new Error('Gemini API error (' + code + '): ' + errorText);
  }

  var result = JSON.parse(response.getContentText());

  // Extract the text from the response, filtering out thinking parts
  if (result.candidates && result.candidates.length > 0 &&
      result.candidates[0].content && result.candidates[0].content.parts) {
    return result.candidates[0].content.parts
      .filter(function(p) { return !p.thought; })
      .map(function(p) { return p.text || ''; })
      .join('');
  }

  throw new Error('Unexpected Gemini response format.');
}

/**
 * Parses the Gemini response into cleaned transcript and summary.
 * Expects the format:
 *   ===CLEANED===
 *   [transcript]
 *   ===SUMMARY===
 *   [summary]
 */
function parseGeminiResponse(responseText) {
  var cleanedMatch = responseText.match(/===CLEANED===\s*([\s\S]*?)(?====SUMMARY===)/);
  var summaryMatch = responseText.match(/===SUMMARY===\s*([\s\S]*?)$/);

  var cleaned = cleanedMatch ? cleanedMatch[1].trim() : '';
  var summary = summaryMatch ? summaryMatch[1].trim() : '';

  // If parsing failed, use the whole response as a fallback
  if (!cleaned && !summary) {
    cleaned = responseText;
    summary = '(parsing error — see cleaned transcript)';
  }

  return {
    cleanedTranscript: cleaned,
    summary: summary
  };
}

/**
 * Processes a single submission row: calls Gemini and writes results back.
 * courseName: the sheet tab name
 * rowNumber: the row to process (1-indexed)
 */
function processSubmission(courseName, rowNumber) {
  var sheet = getSpreadsheet().getSheetByName(courseName);
  if (!sheet) throw new Error('Course tab "' + courseName + '" not found.');

  // Read the raw transcript (column J = 10)
  var rawTranscript = sheet.getRange(rowNumber, 10).getValue();
  if (!rawTranscript || String(rawTranscript).trim() === '') {
    updateProcessingResults(courseName, rowNumber, '', 'No transcript provided', 'complete');
    return;
  }

  // Get the prompt from Config
  var prompt = getConfigValue('Gemini Prompt');
  if (!prompt) prompt = getDefaultGeminiPrompt();

  // Call Gemini
  var responseText = callGemini(prompt, String(rawTranscript));

  // Parse the response
  var parsed = parseGeminiResponse(responseText);

  // Write results back
  updateProcessingResults(
    courseName,
    rowNumber,
    parsed.cleanedTranscript,
    parsed.summary,
    'complete'
  );
}

/**
 * Retries all failed submissions across all course tabs.
 * Called from the instructor menu.
 */
function retryFailedProcessing() {
  var courses = getActiveCourses();
  var retried = 0;
  var errors = 0;

  for (var i = 0; i < courses.length; i++) {
    var sheet = getSpreadsheet().getSheetByName(courses[i].name);
    if (!sheet) continue;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    // Read status column (M = 13)
    var statuses = sheet.getRange(2, 13, lastRow - 1, 1).getValues();
    for (var j = 0; j < statuses.length; j++) {
      if (statuses[j][0] === 'error' || statuses[j][0] === 'pending') {
        var rowNumber = j + 2;
        try {
          processSubmission(courses[i].name, rowNumber);
          retried++;
        } catch (err) {
          errors++;
          Logger.log('Retry failed for ' + courses[i].name + ' row ' + rowNumber + ': ' + err.message);
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    'Retry complete.\n' +
    'Successfully processed: ' + retried + '\n' +
    'Still failing: ' + errors
  );
}
