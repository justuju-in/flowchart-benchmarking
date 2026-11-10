/**
 * Service for handling form submissions.
 * Handles the processing of new form submissions and their entry into the master sheet.
 */

/**
 * Handler for form submissions.
 * @param {Object} e - The form submission event.
 */
function processFormSubmission(e) {
  try {
    if (!e || !e.range) {
      throw new Error("This function must run from a spreadsheet form-submit trigger.");
    }

    var submissionData = extractFormSubmissionEventData_(e);
    var spreadsheet = e.source || e.range.getSheet().getParent();
    var masterSheet = spreadsheet.getSheetByName(getConfig().masterSheet);
    if (!masterSheet) {
      masterSheet = spreadsheet.insertSheet(getConfig().masterSheet);
      masterSheet.appendRow(getConfig().masterSheetColumns);
      masterSheet.getRange(1, 1, 1, getConfig().masterSheetColumns.length)
        .setFontWeight("bold");
    }

    var newRow = createMasterSheetRow(submissionData);
    while (newRow.length < masterSheet.getLastColumn()) {
      newRow.push("");
    }
    masterSheet.appendRow(newRow);
    SpreadsheetApp.flush();
    console.log("New form submission added to Master.");
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    throw err;
  }
}

/**
 * Installs the spreadsheet form-submit trigger used by this project.
 *
 * Run this function once from the Apps Script editor after deploying the
 * local code. It is intentionally not called automatically: creating an
 * installable trigger requires the spreadsheet owner's authorization.
 * Existing matching triggers are kept, so running setup twice does not
 * create duplicate Master rows.
 *
 * @return {string} Human-readable setup result.
 */
function installFormSubmissionTrigger() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("Open the bound spreadsheet before installing the form submission trigger.");
  }

  var handlerName = "processFormSubmission";
  var spreadsheetId = spreadsheet.getId();
  var triggerExists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === handlerName &&
      trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT &&
      trigger.getTriggerSourceId() === spreadsheetId;
  });

  if (triggerExists) {
    return "Form submission trigger is already installed for this spreadsheet.";
  }

  ScriptApp.newTrigger(handlerName)
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();

  return "Form submission trigger installed for this spreadsheet.";
}

/**
 * Installs the recurring pipeline for the copied test system.
 *
 * This deliberately does not run during deployment. After the copied script's
 * Script Properties have been set and one end-to-end test succeeds, run this
 * function once. It is idempotent and never creates duplicate timers.
 */
function installCopiedFormPipelineTrigger() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("Open the copied test spreadsheet before installing the pipeline trigger.");
  }
  if (!getGeminiApiKey()) {
    throw new Error("Set GEMINI_API_KEY in copied-test Script Properties first.");
  }
  getConfiguredDomjudge_();

  var handlerName = "runCopiedFormPipeline";
  var exists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === handlerName;
  });
  if (exists) return "Copied-form pipeline trigger is already installed.";

  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .everyMinutes(5)
    .create();
  return "Copied-form pipeline trigger installed; it runs every 5 minutes.";
}

/** Runs the copied-form pipeline. Historical Master rows are filtered out. */
function runCopiedFormPipeline() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return "Another copied-form pipeline run is active.";
  try {
    processGemini();
    triggerDomJudgeProcessing();
    triggerVerdictPolling();
    return "Copied-form pipeline completed.";
  } finally {
    lock.releaseLock();
  }
}

/**
 * Extracts submission data from a form submission.
 * @param {Object} e - The form submission event.
 * @return {Object} - The extracted submission data.
 */
function extractFormSubmissionEventData_(e) {
  var responseSheet = e.range.getSheet();
  var headers = responseSheet
    .getRange(1, 1, 1, responseSheet.getLastColumn())
    .getDisplayValues()[0];
  var values = responseSheet
    .getRange(e.range.getRow(), 1, 1, responseSheet.getLastColumn())
    .getDisplayValues()[0];

  function valueFor(columnNames, fallbackPattern) {
    var candidates = Array.isArray(columnNames) ? columnNames : [columnNames];
    for (var i = 0; i < candidates.length; i++) {
      var index = headers.indexOf(candidates[i]);
      if (index >= 0 && values[index]) {
        return values[index];
      }
    }

    if (fallbackPattern) {
      for (var j = 0; j < values.length; j++) {
        if (fallbackPattern.test(String(values[j] || ""))) {
          return values[j];
        }
      }
    }
    return "";
  }

  return {
    timestamp: valueFor("Timestamp"),
    email: valueFor("Email Address"),
    // The copied form's problem field is currently labelled "Column 2".
    // Keep the original label for backwards compatibility and fall back to
    // any submitted value containing an FCP code.
    problemCode: valueFor(
      ["Print User Input", "Problem Number", "Column 2"],
      /\bFCP\d{3}\b/i
    ),
    // Older forms use "Upload your Flowchart"; the copied test form uses
    // "Please upload image". Both map to the same Master column.
    flowchartUrl: valueFor(
      ["Please upload image", "Upload your Flowchart"],
      /drive\.google\.com|docs\.google\.com/i
    )
  };
}

/**
 * Creates a new row for the master sheet from submission data.
 * @param {Object} data - The submission data.
 * @return {Array} - Array representing the new row.
 */
function createMasterSheetRow(data) {
  return [
    data.timestamp,
    data.email,
    data.problemCode,
    data.flowchartUrl,
    getConfig().NewEntryStatus,
    "", // Empty placeholders for remaining columns
    "",
    "",
    ""
  ];
}
