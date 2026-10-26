/**
 * Batch mode utility functions
 */

// Get or create a folder in Google Drive
function getOrCreateFolder_(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

// Upload a JSONL file to Gemini API
function uploadJsonlToGemini_(driveFile, apiKey) {
  const baseUrl = getConfig().geminiEndpoint;
  const bytes = driveFile.getBlob().getBytes();
  const mime = "application/jsonl";
  const displayName = "BatchInput-" + driveFile.getName();

  // Debug: log upload parameters
  console.log("uploadJsonlToGemini_: baseUrl", baseUrl);
  console.log("uploadJsonlToGemini_: displayName", displayName);
  console.log("uploadJsonlToGemini_: bytes.length", bytes.length);

  // start
  const startResp = UrlFetchApp.fetch(`${baseUrl}/upload/v1beta/files?key=${encodeURIComponent(apiKey)}`, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ file: { display_name: displayName } }),
    muteHttpExceptions: true,
    headers: {
  "X-Goog-Upload-Protocol": "resumable",
  "X-Goog-Upload-Command": "start",
  // DO NOT set X-Goog-Upload-Header-Content-Length, Apps Script manages it automatically
  "X-Goog-Upload-Header-Content-Type": mime
    }
  });
  console.log("uploadJsonlToGemini_: startResp code", startResp.getResponseCode());
  console.log("uploadJsonlToGemini_: startResp headers", startResp.getHeaders());
  if (startResp.getResponseCode() < 200 || startResp.getResponseCode() >= 300)
    throw new Error("Resumable start failed: " + startResp.getContentText());

  const uploadUrl = startResp.getHeaders()["X-Goog-Upload-Url"] || startResp.getHeaders()["x-goog-upload-url"];
  console.log("uploadJsonlToGemini_: uploadUrl", uploadUrl);
  if (!uploadUrl) throw new Error("Missing X-Goog-Upload-Url");

  // upload + finalize
  const upResp = UrlFetchApp.fetch(uploadUrl, {
    method: "post",
    contentType: mime,
    payload: bytes,
    muteHttpExceptions: true,
    headers: {
      "X-Goog-Upload-Command": "upload, finalize",
      "X-Goog-Upload-Offset": "0"
      // DO NOT set Content-Length, Apps Script manages it automatically
    }
  });
  console.log("uploadJsonlToGemini_: upResp code", upResp.getResponseCode());
  console.log("uploadJsonlToGemini_: upResp text", upResp.getContentText());
  if (upResp.getResponseCode() < 200 || upResp.getResponseCode() >= 300)
    throw new Error("Upload failed: " + upResp.getContentText());

  const json = JSON.parse(upResp.getContentText());
  console.log("uploadJsonlToGemini_: upload response json", json);
  const fileName = json?.file?.name;
  if (!fileName) throw new Error("Upload response missing file.name");
  return fileName; // e.g., "files/abc123"
}

// Create a batch with a file in Gemini API
function createBatchWithFile_(baseUrl, apiKey, model, fileName, displayName) {
  // The Gemini batch endpoint requires the API key as a query parameter, not as a header
  // Use /models/{model}:batchGenerateContent for Gemini batch API
  const url = `${baseUrl}/v1beta/models/${model}:batchGenerateContent?key=${encodeURIComponent(apiKey)}`;
  const body = { batch: { displayName, inputConfig: { fileName } } };
  console.log("createBatchWithFile_: url", url);
  console.log("createBatchWithFile_: body", body);
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
    // No need for x-goog-api-key header
  });
  console.log("createBatchWithFile_: resp code", resp.getResponseCode());
  console.log("createBatchWithFile_: resp text", resp.getContentText());
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300)
    throw new Error("Batch create failed: " + resp.getContentText());
  return JSON.parse(resp.getContentText());
}

// Get batch information from Gemini API
function getBatch_(baseUrl, apiKey, batchName) {
  const url = `${baseUrl}/v1beta/${batchName}`;
  console.log("getBatch_: url", url);
  const resp = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: { "x-goog-api-key": apiKey, "Content-Type": "application/json" }
  });
  console.log("getBatch_: resp code", resp.getResponseCode());
  console.log("getBatch_: resp text", resp.getContentText());
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300)
    throw new Error("getBatch failed: " + resp.getContentText());
  return JSON.parse(resp.getContentText());
}

// Check if a batch is in a terminal state
function isTerminalState_(state) {
  return new Set([
    "BATCH_STATE_SUCCEEDED","BATCH_STATE_FAILED","BATCH_STATE_CANCELLED","BATCH_STATE_EXPIRED",
    "JOB_STATE_SUCCEEDED","JOB_STATE_FAILED","JOB_STATE_CANCELLED","JOB_STATE_EXPIRED"
  ]).has(state);
}

// Download response file from Gemini API
function downloadResponsesFile_(baseUrl, apiKey, responsesFile) {
  const url = `${baseUrl}/download/v1beta/${responsesFile}:download?alt=media`;
  console.log("downloadResponsesFile_: url", url);
  const resp = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: { "x-goog-api-key": apiKey }
  });
  console.log("downloadResponsesFile_: resp code", resp.getResponseCode());
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300)
    throw new Error("downloadResponsesFile failed: " + resp.getContentText());
  return resp.getBlob().getBytes();
}

// Test function to start Gemini batch upload and creation
function testGeminiBatchUploadAndCreate(apiKey, driveFileId, model) {
  try {
    console.log("testGeminiBatchUploadAndCreate: apiKey", apiKey);
    console.log("testGeminiBatchUploadAndCreate: driveFileId", driveFileId);
    console.log("testGeminiBatchUploadAndCreate: model", model);
    var driveFile = DriveApp.getFileById(driveFileId);
    var uploadedFileName = uploadJsonlToGemini_(driveFile, apiKey);
    console.log("testGeminiBatchUploadAndCreate: uploadedFileName", uploadedFileName);
    var batch = createBatchWithFile_(getConfig().geminiEndpoint, apiKey, model, uploadedFileName, "flowchart-judge-test");
    console.log("testGeminiBatchUploadAndCreate: batch", batch);
    return batch;
  } catch (e) {
    console.log("testGeminiBatchUploadAndCreate: error", e);
    throw e;
  }
}
