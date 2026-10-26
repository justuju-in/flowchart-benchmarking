/**
 * GeminiAPI.js
 * 
 * Low-level implementation of Gemini API calls that can be reused 
 * across different parts of the system.
 */

// ====================== API COMMUNICATION ======================

/**
 * Make a content generation request to the Gemini API.
 * @param {Object} requestData - The structured request data
 * @param {string} [model] - Optional model to override config default
 * @param {Object} [options] - Additional options
 * @returns {Object} The API response or error
 */
function geminiGenerateContent(requestData, model, options = {}) {
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    throw new Error("Missing GEMINI_API_KEY in Script Properties.");
  }
  
  const modelName = model || getConfig().geminiModel;
  const endpoint = `${getConfig().geminiEndpoint}${modelName}:generateContent?key=${apiKey}`;
  
  try {
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestData),
      muteHttpExceptions: true,
      ...options
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      logEvent(`Gemini API error: ${responseCode} - ${responseText}`, getConfig().logLevelError);
      return { success: false, error: `API error: ${responseCode}`, details: responseText };
    }
    
    return { 
      success: true, 
      data: JSON.parse(responseText)
    };
  } catch (error) {
    logEvent(`Error in geminiGenerateContent: ${error.toString()}`, getConfig().logLevelError);
    return { success: false, error: error.toString() };
  }
}

/**
 * Create a batch request with a file.
 * @param {string} fileName - The file name to use for batch processing
 * @param {string} [displayName] - Optional display name for the batch
 * @param {string} [model] - Optional model to use
 * @returns {Object} Batch creation result
 */
function geminiBatchCreate(fileName, displayName = "flowchart-judge-batch", model) {
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    throw new Error("Missing GEMINI_API_KEY in Script Properties.");
  }
  
  const modelName = model || getConfig().geminiModel;
  const baseUrl = getConfig().geminiEndpointBase;
  const url = `${baseUrl}/v1beta/models/${modelName}:batchGenerateContent?key=${encodeURIComponent(apiKey)}`;
  const body = { batch: { displayName, inputConfig: { fileName } } };
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode < 200 || responseCode >= 300) {
      logEvent(`Batch create failed: ${responseText}`, getConfig().logLevelError);
      return { success: false, error: `Batch create failed: ${responseCode}`, details: responseText };
    }
    
    return { 
      success: true, 
      data: JSON.parse(responseText)
    };
  } catch (error) {
    logEvent(`Error in geminiBatchCreate: ${error.toString()}`, getConfig().logLevelError);
    return { success: false, error: error.toString() };
  }
}

/**
 * Get batch status information.
 * @param {string} batchName - The batch name
 * @returns {Object} Batch status result
 */
function geminiBatchStatus(batchName) {
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    throw new Error("Missing GEMINI_API_KEY in Script Properties.");
  }
  
  const baseUrl = getConfig().geminiEndpointBase;
  const url = `${baseUrl}/v1beta/${batchName}`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      muteHttpExceptions: true,
      headers: { "x-goog-api-key": apiKey, "Content-Type": "application/json" }
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode < 200 || responseCode >= 300) {
      logEvent(`getBatch failed: ${responseText}`, getConfig().logLevelError);
      return { success: false, error: `Get batch failed: ${responseCode}`, details: responseText };
    }
    
    return { 
      success: true, 
      data: JSON.parse(responseText)
    };
  } catch (error) {
    logEvent(`Error in geminiBatchStatus: ${error.toString()}`, getConfig().logLevelError);
    return { success: false, error: error.toString() };
  }
}

/**
 * Upload a file to Gemini for batch processing.
 * @param {DriveFile} driveFile - Google Drive file to upload
 * @returns {Object} Upload result with file name
 */
function geminiUploadFile(driveFile) {
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    throw new Error("Missing GEMINI_API_KEY in Script Properties.");
  }
  
  const baseUrl = getConfig().geminiEndpointBase;
  const bytes = driveFile.getBlob().getBytes();
  const mime = "application/jsonl";
  const displayName = "BatchInput-" + driveFile.getName();

  try {
    // Start upload
    const startResp = UrlFetchApp.fetch(`${baseUrl}/upload/v1beta/files?key=${encodeURIComponent(apiKey)}`, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ file: { display_name: displayName } }),
      muteHttpExceptions: true,
      headers: {
        "X-Goog-Upload-Protocol": "resumable",
        "X-Goog-Upload-Command": "start",
        "X-Goog-Upload-Header-Content-Type": mime
      }
    });
    
    if (startResp.getResponseCode() < 200 || startResp.getResponseCode() >= 300) {
      return { success: false, error: "Resumable start failed", details: startResp.getContentText() };
    }

    const uploadUrl = startResp.getHeaders()["X-Goog-Upload-Url"] || startResp.getHeaders()["x-goog-upload-url"];
    if (!uploadUrl) {
      return { success: false, error: "Missing X-Goog-Upload-Url" };
    }

    // Upload + finalize
    const upResp = UrlFetchApp.fetch(uploadUrl, {
      method: "post",
      contentType: mime,
      payload: bytes,
      muteHttpExceptions: true,
      headers: {
        "X-Goog-Upload-Command": "upload, finalize",
        "X-Goog-Upload-Offset": "0"
      }
    });
    
    if (upResp.getResponseCode() < 200 || upResp.getResponseCode() >= 300) {
      return { success: false, error: "Upload failed", details: upResp.getContentText() };
    }

    const json = JSON.parse(upResp.getContentText());
    const fileName = json?.file?.name;
    if (!fileName) {
      return { success: false, error: "Upload response missing file.name" };
    }
    
    return { success: true, fileName: fileName };
  } catch (error) {
    logEvent(`Error in geminiUploadFile: ${error.toString()}`, getConfig().logLevelError);
    return { success: false, error: error.toString() };
  }
}

/**
 * Download a file from Gemini API.
 * @param {string} responsesFile - The file path to download
 * @returns {Object} Download result with file bytes
 */
function geminiDownloadFile(responsesFile) {
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    throw new Error("Missing GEMINI_API_KEY in Script Properties.");
  }
  
  const baseUrl = getConfig().geminiEndpointBase;
  const url = `${baseUrl}/download/v1beta/${responsesFile}:download?alt=media`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      muteHttpExceptions: true,
      headers: { "x-goog-api-key": apiKey }
    });
    
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      return { 
        success: false, 
        error: "Download failed", 
        details: response.getContentText() 
      };
    }
    
    return { 
      success: true,
      bytes: response.getBlob().getBytes(),
      blob: response.getBlob()
    };
  } catch (error) {
    logEvent(`Error in geminiDownloadFile: ${error.toString()}`, getConfig().logLevelError);
    return { success: false, error: error.toString() };
  }
}

// ====================== HELPER FUNCTIONS ======================

/**
 * Check if a batch state is terminal (complete).
 * @param {string} state - The batch state string
 * @returns {boolean} True if the state is terminal
 */
function isGeminiBatchTerminal(state) {
  return new Set([
    "BATCH_STATE_SUCCEEDED",
    "BATCH_STATE_FAILED",
    "BATCH_STATE_CANCELLED",
    "BATCH_STATE_EXPIRED",
    "JOB_STATE_SUCCEEDED",
    "JOB_STATE_FAILED",
    "JOB_STATE_CANCELLED",
    "JOB_STATE_EXPIRED"
  ]).has(state);
}

/**
 * Prepare an image request for Gemini.
 * @param {string} fileId - Google Drive file ID
 * @param {string} promptText - Prompt text
 * @param {number} temperature - Generation temperature (0-1)
 * @returns {Object} Request data object
 */
function prepareGeminiImageRequest(fileId, promptText, temperature = 0) {
  try {
    // Get the image as base64
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const mimeType = blob.getContentType();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    
    // Prepare request data
    return {
      contents: [
        {
          parts: [
            { 
              inline_data: {
                mime_type: mimeType,
                data: base64Data
              }
            },
            { text: promptText }
          ]
        }
      ],
      generationConfig: {
        temperature: temperature
      }
    };
  } catch (error) {
    logEvent(`Error preparing Gemini image request: ${error.toString()}`, getConfig().logLevelError);
    return null;
  }
}

/**
 * Extract and parse content from Gemini response.
 * @param {Object} response - The Gemini API response
 * @returns {Object} Extracted content and metadata
 */
function extractGeminiContent(response) {
  if (!response || !response.success || !response.data) {
    return { success: false };
  }
  
  try {
    const data = response.data;
    const candidates = data.candidates || [];
    
    if (candidates.length === 0 || !candidates[0].content) {
      return { success: false, error: "No valid content in response" };
    }
    
    const parts = candidates[0].content.parts || [];
    const textPart = parts.find(part => typeof part.text === "string");
    
    if (!textPart) {
      return { success: false, error: "No text content in response" };
    }
    
    // Extract text and clean it
    const rawText = textPart.text;
    const code = cleanCodeBlock(rawText);
    
    // Extract metadata
    const metadata = extractGeminiMetadata(data);
    
    return {
      success: true,
      content: code,
      rawText: rawText,
      metadata: metadata
    };
  } catch (error) {
    logEvent(`Error extracting Gemini content: ${error.toString()}`, getConfig().logLevelError);
    return { success: false, error: error.toString() };
  }
}

/**
 * Extract metadata from Gemini API response.
 * @param {Object} response - The raw Gemini API response
 * @returns {Object} - Formatted metadata
 */
function extractGeminiMetadata(response) {
  const metadata = {};
  
  try {
    // Extract usage metadata
    const usage = response.usageMetadata || {};
    metadata.inputTokens = usage.promptTokenCount || 0;
    metadata.outputTokens = usage.candidatesTokenCount || 0;
    metadata.totalTokens = usage.totalTokenCount || 0;
    metadata.responseTime = usage.responseTimeMs || null;
    
    // Extract token details
    metadata.thoughtsTokenCount = usage.thoughtsTokenCount || 0;
    
    const promptTokensDetails = usage.promptTokensDetails || [];
    metadata.textTokenCount = (promptTokensDetails.find(d => d.modality === "TEXT") || {}).tokenCount || 0;
    metadata.imageTokenCount = (promptTokensDetails.find(d => d.modality === "IMAGE") || 
                               promptTokensDetails.find(d => d.modality === "DOCUMENT") || {}).tokenCount || 0;
    
    // Extract candidate metadata
    const candidate = (response.candidates || [])[0] || {};
    metadata.safetyRatings = JSON.stringify(candidate.safetyRatings || []);
    metadata.finishReason = candidate.finishReason || getConfig().defaultFinishReason;
    metadata.citationMetadata = JSON.stringify(candidate.citationMetadata || {});
    
    // Extract model info
    metadata.modelVersion = response.modelVersion || getConfig().geminiModel;
    metadata.responseId = response.responseId || "";
    
  } catch (error) {
    logEvent(`Error extracting Gemini metadata: ${error.toString()}`, getConfig().logLevelError);
  }
  
  return metadata;
}


