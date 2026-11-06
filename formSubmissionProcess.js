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

  function valueFor(columnName) {
    var index = headers.indexOf(columnName);
    return index >= 0 ? values[index] || "" : "";
  }

  return {
    timestamp: valueFor("Timestamp"),
    email: valueFor("Email Address"),
    problemCode: valueFor("Print User Input"),
    // Older forms use "Upload your Flowchart"; the copied test form uses
    // "Please upload image". Both map to the same Master column.
    flowchartUrl: valueFor("Please upload image") || valueFor("Upload your Flowchart")
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
