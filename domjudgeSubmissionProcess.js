/**
 * Service for submitting code to DomJudge.
 * Handles the submission of generated code to the DomJudge system for evaluation.
 */

/**
 * Submits generated code to DomJudge for evaluation.
 * @param {Sheet} sheet - The master sheet.
 * @param {Array[]} data - Sheet data as 2D array.
 * @param {number} rowIndex - Index of the row in data to process.
 * @return {boolean} - True if submission was successful.
 */
function submitCodeToDomJudge(sheet, data, rowIndex) {
  const realRow = rowIndex + 1; // 1-indexed for actual sheet operations
  
  // Extract necessary data
  var codeFileUrl = getSheetValueByColumnName(data, rowIndex, sheet, getConfig().CodeFileUrlColumnName);
  if (!codeFileUrl) {
    logEvent("No code file URL found for row: " + realRow, getConfig().logLevelError);
    console.log("No code file URL found for row: " + realRow);
    return false;
  }
  
  var code = getCodeFromDriveUrl(codeFileUrl);
  if (!code) {
    logEvent("Failed to retrieve code from URL: " + codeFileUrl, getConfig().logLevelError);
    console.log("Failed to retrieve code from URL: " + codeFileUrl);
    return false;
  }
  
  var problemNumber = getSheetValueByColumnName(data, rowIndex, sheet, getConfig().ProblemNumberColumnName);
  logEvent("Raw problem number from sheet: " + problemNumber, getConfig().logLevelDebug);
  console.log("Raw problem number from sheet: " + problemNumber);
  
  // Extract just the problem code (e.g., "FCP045" from "FCP045 - Title")
  var problemCode = String(problemNumber).split(" ")[0].trim();
  logEvent("Extracted problem code: " + problemCode, getConfig().logLevelDebug);
  console.log("Extracted problem code: " + problemCode);
  
  var problemId = getProblemIdFromMeta(problemCode);
  logEvent("Resolved problem ID: " + problemId, getConfig().logLevelDebug);
  console.log("Resolved problem ID: " + problemId);
  
  if (problemId === null || problemId === undefined) {
    // Log this as a critical error to ensure visibility
    logEvent("CRITICAL: No valid problem ID found for problem code: " + problemCode, getConfig().logLevelError);
    console.log("CRITICAL: No valid problem ID found for problem code: " + problemCode);
    
    // Create meaningful error message to display in sheet
    setCellValueByColumnName(sheet, realRow, getConfig().VerdictColumnName, 
                            "ERROR: Problem code " + problemCode + " not found in Meta sheet");
    
    // Update status to indicate the problem
    setCellValueByColumnName(sheet, realRow, getConfig().StatusColumnName, 
                            getConfig().CannotProcessStatus);
    
    return false;
  }
  
  // Oldjudge's public API identifies PRE problems by their external code
  // (for example FCP023), while some installations use the numeric Meta ID.
  // Resolve this in one place so submission and polling use the same ID.
  var apiProblemId = getDomjudgeProblemIdentifier_(problemCode, problemId, getConfig());
  logEvent("Submitting DOMjudge problem identifier: " + apiProblemId, getConfig().logLevelDebug);

  // Submit to DomJudge
  var submissionTimestamp = new Date();
  var submissionId = submitToDomjudge(code, apiProblemId);
  console.log("submission attempt for: " + problemCode + ", result: " + submissionId);
  
  if (!submissionId) {
    return false;
  }
  
  // Save submission data
  saveSubmissionMetadata(sheet, realRow, submissionId, submissionTimestamp);
  return true;
}


// Remove this function since it's now in domjudge.js
// Just call submitToDomjudge() from domjudge.js directly
/**
 * Saves metadata about a DomJudge submission.
 * @param {Sheet} sheet - The sheet to update.
 * @param {number} row - The row number.
 * @param {string} submissionId - The submission ID from DomJudge.
 * @param {Date} timestamp - Submission timestamp.
 */
function saveSubmissionMetadata(sheet, row, submissionId, timestamp) {
  var metadata = {
    [getConfig().SubmissionIdColumnName]: submissionId,
    [getConfig().SubmissionTimestampColumnName]: timestamp,
    [getConfig().SubmissionStatusColumnName]: getConfig().SubmissionStatus,
    [getConfig().StatusColumnName]: getConfig().JudgeSubmittedStatus
  };
  
  updateSheetMetadata(sheet, row, metadata);
}

/**
 * Triggers DomJudge submission for all eligible entries in the master sheet.
 */
function triggerDomJudgeProcessing() {
  var masterSheet = getOrCreateSheet(
    getConfig().masterSheet,
    getConfig().masterSheetColumns
  );
  // Get all sheet data using helper function
  var data = getSheetData(masterSheet);
  var copiedFormSubmissionKeys = getCopiedFormSubmissionKeys_(masterSheet.getParent());
  
  for (var i = 1; i < data.length; i++) {
    var status = getSheetValueByColumnName(data, i, masterSheet, getConfig().StatusColumnName);
    
    if (status === getConfig().GeminiDoneStatus &&
        isCopiedFormSubmission_(masterSheet, data, i, copiedFormSubmissionKeys)) {
      submitCodeToDomJudge(masterSheet, data, i);
    }
  }
}

/**
 * Builds stable identity keys for all responses received through the copied
 * form. Historical rows imported into Master are deliberately not included.
 *
 * @param {Spreadsheet} spreadsheet The bound test spreadsheet.
 * @return {Object<string, boolean>} Form-response identity keys.
 */
function getCopiedFormSubmissionKeys_(spreadsheet) {
  var responseSheet = spreadsheet.getSheetByName(getConfig().sourceResponseSheet);
  if (!responseSheet || responseSheet.getLastRow() < 2) {
    return {};
  }

  var rows = responseSheet.getDataRange().getDisplayValues();
  var headers = rows[0];
  var emailIndex = findCopiedResponseHeaderIndex_(headers, "email");
  var problemIndex = findCopiedResponseHeaderIndex_(headers, "problem");
  var uploadIndex = findCopiedResponseHeaderIndex_(headers, "upload");
  if (emailIndex < 0 || problemIndex < 0 || uploadIndex < 0) {
    throw new Error(
      "The copied form response sheet is missing an email, problem, or flowchart-upload column. " +
      "Found headers: " + headers.join(" | ")
    );
  }

  return rows.slice(1).reduce(function(keys, row) {
    var key = copiedFormSubmissionKey_(row[emailIndex], row[problemIndex], row[uploadIndex]);
    if (key) keys[key] = true;
    return keys;
  }, {});
}

/**
 * Locates the copied Google Form columns even if its question label has been
 * edited. This keeps the copied test workflow independent of exact wording.
 */
function findCopiedResponseHeaderIndex_(headers, kind) {
  var normalised = headers.map(function(header) {
    return String(header || "").trim().toLowerCase().replace(/\s+/g, " ");
  });
  var exact = {
    email: ["email address", "email"],
    problem: ["print user input", "problem number", "problem"],
    upload: ["please upload image", "upload your flowchart", "upload flowchart"]
  }[kind] || [];

  for (var i = 0; i < exact.length; i++) {
    var exactIndex = normalised.indexOf(exact[i]);
    if (exactIndex >= 0) return exactIndex;
  }

  return normalised.findIndex(function(header) {
    if (kind === "email") return header.indexOf("email") >= 0;
    if (kind === "problem") {
      return header.indexOf("problem") >= 0 ||
        (header.indexOf("print") >= 0 && header.indexOf("user input") >= 0);
    }
    return header.indexOf("upload") >= 0 &&
      (header.indexOf("flowchart") >= 0 || header.indexOf("image") >= 0);
  });
}

/**
 * Checks that a Master row came from the copied form before any test judge
 * submission is made.
 */
function isCopiedFormSubmission_(sheet, data, rowIndex, copiedFormSubmissionKeys) {
  var email = getSheetValueByColumnName(data, rowIndex, sheet, getConfig().EmailAddressColumnName);
  var problem = getSheetValueByColumnName(data, rowIndex, sheet, getConfig().ProblemNumberColumnName);
  var flowchartUrl = getSheetValueByColumnName(data, rowIndex, sheet, getConfig().UploadFlowchartColumnName);
  var key = copiedFormSubmissionKey_(email, problem, flowchartUrl);
  return Boolean(key && copiedFormSubmissionKeys[key]);
}

function copiedFormSubmissionKey_(email, problem, flowchartUrl) {
  var fileId = getFileIdFromUrl(String(flowchartUrl || ""));
  if (!email || !problem || !fileId) return "";
  return [String(email).trim().toLowerCase(), String(problem).trim(), fileId].join("|");
}
