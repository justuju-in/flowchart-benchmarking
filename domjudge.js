/**
 * domjudge.js - DOMjudge API specific utilities
 * 
 * This file contains specialized utilities for interacting with the DOMjudge API.
 * These functions build on the generic API utilities from apiUtils.js.
 */

/**
 * Creates a DOMjudge submission URL from contest ID.
 * 
 * @param {string} baseUrl - The base DOMjudge API URL
 * @param {string} contestId - The contest ID
 * @returns {string} The full submissions URL
 */
function createDomjudgeSubmissionUrl(baseUrl, contestId) {
  return `${baseUrl}/contests/${contestId}/submissions`;
}

/**
 * Creates a DOMjudge judgements URL with optional submission ID filter.
 * 
 * @param {string} baseUrl - The base DOMjudge API URL
 * @param {string} contestId - The contest ID
 * @param {string} [submissionId] - Optional submission ID to filter judgements
 * @returns {string} The full judgements URL
 */
function createDomjudgeJudgementsUrl(baseUrl, contestId, submissionId) {
  let url = `${baseUrl}/contests/${contestId}/judgements`;
  if (submissionId) {
    url += `?submission_id=${encodeURIComponent(submissionId)}&strict=false`;
  }
  return url;
}

/**
 * Creates authorization headers for DOMjudge API using student credentials.
 * 
 * @returns {Object} Headers with student authorization
 */
function createDomjudgeStudentAuthHeaders(config) {
  const username = config.domjudgeUser;
  const password = config.domjudgePass;
  return createBasicAuthHeaders(username, password);
}

/**
 * Creates authorization headers for DOMjudge API using admin credentials.
 * 
 * @returns {Object} Headers with admin authorization
 */
function createDomjudgeAdminAuthHeaders(config) {
  const username = config.domjudgeAdminUser;
  const password = config.domjudgeAdminPass;
  return createBasicAuthHeaders(username, password);
}

/** Ensures this copied script cannot call a judge until configured safely. */
function getConfiguredDomjudge_() {
  const config = getConfig();
  const missing = [
    ["DOMJUDGE_BASE_URL", config.domjudgeApiUrl],
    ["DOMJUDGE_CONTEST_ID", config.domjudgeContestId],
    ["DOMJUDGE_USERNAME", config.domjudgeUser],
    ["DOMJUDGE_PASSWORD", config.domjudgePass]
  ].filter(function(entry) { return !entry[1]; })
    .map(function(entry) { return entry[0]; });

  if (missing.length) {
    throw new Error("Set copied-test Script Properties before using DOMjudge: " + missing.join(", "));
  }
  return config;
}

/**
 * Submits code to DOMjudge using the centralized API utility.
 * 
 * @param {string} code - The code to submit
 * @param {string|number} problemId - The problem ID
 * @returns {string} The submission ID if successful, empty string otherwise
 */
function submitToDomjudge(code, problemId) {
  // Get configuration values
  const config = getConfiguredDomjudge_();
  const contestId = config.domjudgeContestId;
  const languageId = config.domjudgeLanguageId;
  const filename = config.domjudgeSolutionFilename;
  const teamId = config.domjudgeTeamId;
  const domjudgeUrl = config.domjudgeApiUrl;
  
  // Prepare file for submission
  const codeBlob = Utilities.newBlob(code, 'text/x-python', filename);
  const zipBlob = Utilities.zip([codeBlob], config.domjudgeZipFilename);
  const encodedZip = Utilities.base64Encode(zipBlob.getBytes());
  
  // Create submission URL
  const url = createDomjudgeSubmissionUrl(domjudgeUrl, contestId);
  
  // Ensure problemId is properly formatted (numeric if possible)
  const parsedProblemId = !isNaN(problemId) ? Number(problemId) : String(problemId);
  
  // Create submission payload
  const payload = {
    problem_id: parsedProblemId,
    language_id: languageId,
    team_id: teamId,
    files: [
      {
        filename: config.domjudgeZipFilename,
        data: encodedZip,
      },
    ],
  };
  if (teamId) payload.team_id = teamId;
  
  // Create request headers and options
  const headers = createDomjudgeStudentAuthHeaders(config);
  const options = createPostRequestOptions(headers, payload);
  
  // Make the API call using the centralized utility
  const response = makeDomjudgeRequest(url, options);
  
  // Process the response
  if (response.success) {
    logEvent("Successfully submitted. Judge ID: " + response.json.id, getConfig().logLevelDebug);
    console.log("Submission successful: " + JSON.stringify(response.json));
    return response.json.id;
  } else {
    logEvent("DomJudge Error: " + response.code + " - " + response.body, getConfig().logLevelError);
    return "";
  }
}
