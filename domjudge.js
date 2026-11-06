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
    ["DOMJUDGE_USERNAME", config.domjudgeUser],
    ["DOMJUDGE_PASSWORD", config.domjudgePass]
  ].filter(function(entry) { return !entry[1]; })
    .map(function(entry) { return entry[0]; });

  if (missing.length) {
    throw new Error("Set copied-test Script Properties before using DOMjudge: " + missing.join(", "));
  }
  if (!getDomjudgeContestIds_(config).length) {
    throw new Error("Set DOMJUDGE_CONTEST_IDS (or the legacy DOMJUDGE_CONTEST_ID) in copied-test Script Properties.");
  }
  return config;
}

/**
 * Returns the configured contest IDs. DOMJUDGE_CONTEST_ID remains supported
 * for a one-contest setup; DOMJUDGE_CONTEST_IDS is used for If-Else/Loop/Array.
 */
function getDomjudgeContestIds_(config) {
  var raw = config.domjudgeContestIds || config.domjudgeContestId || "";
  return String(raw)
    .split(",")
    .map(function(value) { return value.trim(); })
    .filter(function(value, index, values) {
      return value && values.indexOf(value) === index;
    });
}

/**
 * Finds the configured contest that contains a DOMjudge problem ID.
 * This avoids a fragile manual FCP-to-category mapping in the copied system.
 */
function getDomjudgeContestForProblem_(problemId, config) {
  var contestIds = getDomjudgeContestIds_(config);
  if (contestIds.length === 1) return contestIds[0];

  var cache = CacheService.getScriptCache();
  var cacheKey = "domjudge-contest:" + contestIds.join(",") + ":" + String(problemId);
  var cachedContestId = cache.get(cacheKey);
  if (cachedContestId) return cachedContestId;

  var headers = createDomjudgeAdminAuthHeaders(config);
  var options = createGetRequestOptions(headers);
  var matches = [];

  contestIds.forEach(function(contestId) {
    var url = config.domjudgeApiUrl + "/contests/" + encodeURIComponent(contestId) + "/problems";
    var response = makeDomjudgeRequest(url, options);
    if (!response.success) {
      throw new Error("Could not read problems for contest " + contestId + ": HTTP " + response.code);
    }

    var problems = Array.isArray(response.json) ? response.json : [];
    var found = problems.some(function(problem) {
      return String(problem.id) === String(problemId);
    });
    if (found) matches.push(contestId);
  });

  if (matches.length !== 1) {
    throw new Error(matches.length
      ? "Problem " + problemId + " exists in multiple configured contests: " + matches.join(", ")
      : "Problem " + problemId + " is not present in configured contests: " + contestIds.join(", "));
  }

  cache.put(cacheKey, matches[0], 21600);
  return matches[0];
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
  let contestId;
  try {
    contestId = getDomjudgeContestForProblem_(problemId, config);
  } catch (error) {
    logEvent("DomJudge contest routing error: " + error.message, getConfig().logLevelError);
    return "";
  }
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
