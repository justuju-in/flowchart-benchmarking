/**
 * Handles both individual and batch processing of flowchart images with Gemini API.
 */

// ========================= MAIN ENTRY POINTS =========================

/**
 * Main entry point for Gemini processing with automatic mode selection.
 * Decides between batch and individual processing based on configuration and row count.
 */
function processGemini() {
  const cfg = getConfig();
  
  // Early return if there's an issue with the sheet
  const masterSheet = getOrCreateSheet(cfg.masterSheet, cfg.masterSheetColumns);
  if (!masterSheet) {
    logEvent("Failed to access master sheet", cfg.logLevelError);
    return;
  }
  
  // Get all data
  const data = getSheetData(masterSheet);
  
  // Find eligible rows for processing
  const eligibleRows = findEligibleRows(masterSheet, data);
  if (eligibleRows.length === 0) {
    logEvent("No eligible rows found for processing", cfg.logLevelDebug);
    return;
  }
  
  // Decide whether to use batch mode
  if (cfg.IsBatchMode && eligibleRows.length >= cfg.geminiBatchThreshold) {
    // Use batch processing
    logEvent(`Using batch mode for ${eligibleRows.length} rows`, cfg.logLevelDebug);
    return enqueueBatch(cfg.maxRowCount); 
  } else {
    // Use individual processing
    logEvent(`Using individual processing for ${eligibleRows.length} rows`, cfg.logLevelDebug);
    return processItemsIndividually(masterSheet, data, eligibleRows);
  }
}

/**
 * Legacy entry point for individual processing only.
 * Kept for backward compatibility.
 */
function processGeminiIndividually() {
  // Get master sheet and data
  const masterSheet = getOrCreateSheet(
    getConfig().masterSheet,
    getConfig().masterSheetColumns
  );
  const data = getSheetData(masterSheet);
  
  // Find eligible rows for processing
  const eligibleRows = findEligibleRows(masterSheet, data);
  
  // Process each eligible row
  return processItemsIndividually(masterSheet, data, eligibleRows);
}

/**
 * Manually enqueue eligible rows for batch processing.
 * @param {number} maxRows - Maximum number of rows to process
 * @returns {string} Status message
 */
function enqueueBatch(maxRows) {
  const cfg = getConfig();
  const apiKey = getGeminiApiKey();
  
  if (!apiKey) throw new Error("Missing GEMINI_API_KEY in Script Properties.");

  const master = getOrCreateSheet(cfg.masterSheet, cfg.masterSheetColumns);
  const values = master.getDataRange().getValues();
  
  if (values.length <= 1) return "No data";

  const rows = getRowsToEnqueue(master, cfg, maxRows);
  if (!rows.length) return "No NEW rows to enqueue.";

  const promptText = getPrompt();
  const { jsonlLines, manifestRows } = buildBatchData(rows, master, cfg, promptText);
  
  if (!jsonlLines.length) return "No valid files among NEW rows.";

  const manifest = {
    createdAt: new Date().toISOString(),
    model: cfg.geminiModel,
    sheetName: cfg.masterSheet,
    rows: manifestRows,
    batchName: null,
    responsesFile: null
  };

  const folder = getOrCreateFolder(cfg.BatchesFolderName);
  const jsonlName = `batch_input_${new Date().toISOString().replace(/[:.]/g, '-')}.jsonl`;
  const jsonlFile = folder.createFile(jsonlName, jsonlLines.join("\n"), "application/jsonl");

  const uploadedFileName = uploadJsonlToGemini(jsonlFile, apiKey);
  const batch = createBatchWithFile(cfg.geminiEndpointBase, apiKey, cfg.geminiModel, uploadedFileName, "flowchart-judge");
  const batchName = batch.name;
  manifest.batchName = batchName;

  markRowsAsQueued(rows, master, cfg);

  const manifestFile = saveManifest(folder, batchName, manifest);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  trackBatch(ss, batchName, manifestFile.getId(), rows.length);

  return `Enqueued ${rows.length} rows. Batch=${batchName}. ManifestId=${manifestFile.getId()}`;
}

/**
 * Poll for batch results and process them.
 * @param {string} manifestFileId - Optional manifest file ID to poll for
 * @returns {string} Status message
 */
function pollBatchResults(manifestFileId) {
  if (!manifestFileId) {
    return pollLatestBatch();
  } else {
    return pollBatchById(manifestFileId);
  }
}

// ========================= INDIVIDUAL PROCESSING =========================

/**
 * Process multiple items individually using the standard pipeline.
 * @param {Sheet} sheet - The master sheet
 * @param {Array[]} data - Sheet data as 2D array
 * @param {number[]} rowIndices - Array of eligible row indices to process
 * @returns {string} Status message
 */
function processItemsIndividually(sheet, data, rowIndices) {
  let successCount = 0;
  let failureCount = 0;
  
  rowIndices.forEach(rowIndex => {
    try {
      const success = processFlowchartWithGemini(sheet, data, rowIndex);
      if (success) {
        successCount++;
      } else {
        failureCount++;
      }
    } catch (error) {
      logEvent(`Error processing row ${rowIndex + 1}: ${error.message}`, getConfig().logLevelError);
      failureCount++;
    }
  });
  
  return `Processed ${successCount} items successfully, ${failureCount} failed`;
}

/**
 * Processes a single flowchart submission with Gemini.
 * @param {Sheet} sheet - The master sheet
 * @param {Array[]} data - Sheet data as 2D array
 * @param {number} rowIndex - Index of the row in data to process
 * @returns {boolean} True if processing was successful
 */
function processFlowchartWithGemini(sheet, data, rowIndex) {
  const realRow = rowIndex + 1; // 1-indexed for actual sheet operations
  
  // Extract necessary data
  const submissionData = extractSubmissionData(sheet, data, rowIndex);
  if (!submissionData || !submissionData.fileId) {
    return false;
  }
  
  // Save image MIME type
  const mimeType = getImageMimeType(submissionData.fileId);
  setCellValueByColumnName(sheet, realRow, getConfig().ImageMimeTypeColumnName, mimeType);
  
  // Process with Gemini using our centralized image processing function
  const gemResult = processGeminiImageRequest(
    submissionData.fileId, 
    getPrompt(), 
    getConfig().geminiTemperature
  );
  
  if (!gemResult) {
    logEvent("Gemini processing failed for row: " + realRow, getConfig().logLevelError);
    return false;
  }
  
  // Extract and clean code
  const code = gemResult.content;
  
  // Save code and get URL
  const generationTimestamp = new Date();
  const codeFileUrl = saveGeneratedCode(
    code,
    submissionData.email,
    submissionData.problemId,
    generationTimestamp
  );
  
  if (!codeFileUrl) {
    logEvent("Failed to save code to Drive for row: " + realRow, getConfig().logLevelError);
    return false;
  }
  
  // Save all metadata (basic and extended)
  saveGeminiMetadata(sheet, realRow, codeFileUrl, generationTimestamp, gemResult.metadata);
  
  // Update status
  setCellValueByColumnName(sheet, realRow, getConfig().StatusColumnName, getConfig().GeminiDoneStatus);
  
  return true;
}

/**
 * Processes an image with Gemini API.
 * @param {string} fileId - Google Drive file ID of the image
 * @param {string} promptText - The prompt to send to Gemini
 * @param {number} temperature - Temperature for generation (0-1)
 * @return {Object|null} - Response object with content and metadata or null on error
 */
function processGeminiImageRequest(fileId, promptText, temperature = 0) {
  try {
    const apiKey = getGeminiApiKey();
    if (!apiKey) {
      throw new Error("Missing GEMINI_API_KEY in Script Properties.");
    }
    
    // Get the image as base64
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const mimeType = blob.getContentType();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    
    // Prepare request to Gemini API
    const endpoint = `${getConfig().geminiEndpoint}${getConfig().geminiModel}:generateContent?key=${apiKey}`;
    
    const requestBody = {
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
    
    // Make the API call
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });
    
    // Handle response
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      logEvent(`Gemini API error: ${responseCode} - ${responseText}`, getConfig().logLevelError);
      return null;
    }
    
    // Parse response and extract code
    const jsonResponse = JSON.parse(responseText);
    const candidates = jsonResponse.candidates || [];
    
    if (candidates.length === 0 || !candidates[0].content) {
      logEvent("No valid response from Gemini API", getConfig().logLevelError);
      return null;
    }
    
    const parts = candidates[0].content.parts || [];
    const textPart = parts.find(part => typeof part.text === "string");
    
    if (!textPart) {
      logEvent("No text content in Gemini response", getConfig().logLevelError);
      return null;
    }
    
    // Extract code and clean it
    const code = cleanCodeBlock(textPart.text);
    
    // Extract metadata for tracking
    const metadata = extractGeminiMetadata(jsonResponse);
    
    return {
      content: code,
      metadata: metadata
    };
    
  } catch (error) {
    logEvent(`Error in processGeminiImageRequest: ${error.toString()}`, getConfig().logLevelError);
    return null;
  }
}

/**
 * Finds rows that are eligible for Gemini processing.
 * @param {Sheet} sheet - The master sheet
 * @param {Array[]} data - Sheet data as 2D array
 * @returns {number[]} Array of eligible row indices
 */
function findEligibleRows(sheet, data) {
  const eligibleRows = [];
  
  for (let i = 1; i < data.length; i++) {
    const status = getValueFromDataByColumnName(
      data, i, sheet, getConfig().StatusColumnName
    );
    
    const flowchartUrl = getValueFromDataByColumnName(
      data, i, sheet, getConfig().UploadFlowchartColumnName
    );
    
    if (status === getConfig().NewEntryStatus && flowchartUrl) {
      eligibleRows.push(i);
    }
  }
  
  return eligibleRows;
}

/**
 * Extracts submission data from a row.
 * @param {Sheet} sheet - The master sheet
 * @param {Array[]} data - Sheet data as 2D array
 * @param {number} rowIndex - Index of the row in data to process
 * @returns {Object|null} Submission data or null if invalid
 */
function extractSubmissionData(sheet, data, rowIndex) {
  const flowchartUrl = getValueFromDataByColumnName(
    data, rowIndex, sheet, getConfig().UploadFlowchartColumnName
  );
  
  const fileId = getFileIdFromUrl(flowchartUrl);
  if (!fileId) {
    logEvent("Invalid fileId for flowchart URL: " + flowchartUrl, getConfig().logLevelError);
    return null;
  }
  
  const email = getValueFromDataByColumnName(
    data, rowIndex, sheet, getConfig().EmailAddressColumnName
  );
  
  const problemCode = getValueFromDataByColumnName(
    data, rowIndex, sheet, getConfig().ProblemNumberColumnName
  );
  
  // Extract just the problem code part (e.g., "FCP045" from "FCP045 - Title")
  const problemCodePart = problemCode.split(" ")[0].trim();
  const problemId = getProblemIdFromMeta(problemCodePart);
  
  return {
    fileId: fileId,
    email: email,
    problemCode: problemCodePart,
    problemId: problemId
  };
}

/**
 * Extracts and formats metadata from Gemini API response.
 * @param {Object} response - The raw Gemini API response
 * @return {Object} - Formatted metadata
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

/**
 * Cleans a code block from Gemini output.
 * @param {string} text - Raw text from Gemini
 * @return {string} - Cleaned code
 */
function cleanCodeBlock(text) {
  // Remove markdown code block markers if present
  let cleanedText = text.replace(/```[\w]*\n/g, '').replace(/```$/g, '');
  
  // Handle special case where there's no code block markers but explanation text
  const codeBlockIndex = cleanedText.indexOf('```');
  if (codeBlockIndex >= 0) {
    // Extract just the code block
    cleanedText = cleanedText.substring(codeBlockIndex + 3);
    cleanedText = cleanedText.replace(/```$/g, '');
    
    // Remove language identifier if present
    const firstLineBreak = cleanedText.indexOf('\n');
    if (firstLineBreak > 0 && firstLineBreak < 20) { // Language identifier should be short
      cleanedText = cleanedText.substring(firstLineBreak + 1);
    }
  }
  
  return cleanedText.trim();
}

// ========================= BATCH PROCESSING =========================

/**
 * Find rows eligible for batch processing.
 * @param {Sheet} master - The master sheet
 * @param {Object} cfg - Configuration object
 * @param {number} maxRows - Maximum number of rows to return
 * @returns {Array} Array of eligible row objects
 */
function getRowsToEnqueue(master, cfg, maxRows) {
  const col = name => getColumnIndexByName(master, name) - 1;
  const idxStatus = col(cfg.StatusColumnName);
  const idxFlowchartURL = col(cfg.UploadFlowchartColumnName);
  const rows = [];
  const values = master.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row[idxStatus] === cfg.NewEntryStatus && row[idxFlowchartURL]) {
      rows.push({ r, row });
      if (maxRows && rows.length >= maxRows) break;
    }
  }
  return rows;
}

/**
 * Build batch data for submission.
 * @param {Array} rows - Eligible rows
 * @param {Sheet} master - Master sheet
 * @param {Object} cfg - Configuration object
 * @param {string} promptText - Prompt text
 * @returns {Object} Object with jsonlLines and manifestRows
 */
function buildBatchData(rows, master, cfg, promptText) {
  const col = name => getColumnIndexByName(master, name) - 1;
  const idxTimestamp = col(cfg.TimestampColumnName);
  const idxEmail = col(cfg.EmailAddressColumnName);
  const idxProblem = col(cfg.ProblemNumberColumnName);
  const idxFlowchartURL = col(cfg.UploadFlowchartColumnName);
  const jsonlLines = [];
  const manifestRows = [];
  rows.forEach(({ r, row }) => {
    const flowchartUrl = row[idxFlowchartURL];
    const fileId = getFileIdFromUrl(flowchartUrl);
    if (!fileId) {
      logEvent(`Skipping row ${r+1}, invalid fileId for URL: ${flowchartUrl}`, getConfig().logLevelError);
      return;
    }
    let file, blob, mimeType;
    try {
      file = DriveApp.getFileById(fileId);
      blob = file.getBlob();
      mimeType = blob.getContentType();
    } catch (e) {
      logEvent(`Skipping row ${r+1}, error getting file or mimeType: ${e}`, getConfig().logLevelError);
      return;
    }
    if (!mimeType) {
      logEvent(`Skipping row ${r+1}, mimeType is null for fileId: ${fileId}`, getConfig().logLevelError);
      return;
    }
    const b64 = Utilities.base64Encode(blob.getBytes());
    const key = `row-${r+1}`;
    const line = {
      key,
      request: {
        contents: [
          {
            parts: [
              { inline_data: { mime_type: mimeType, data: b64 } },
              { text: promptText }
            ]
          }
        ],
        generationConfig: { temperature: cfg.geminiTemperature || 0 }
      }
    };
    jsonlLines.push(JSON.stringify(line));
    manifestRows.push({
      key,
      r,
      timestamp: row[idxTimestamp],
      email: row[idxEmail],
      problem: row[idxProblem],
      mimeType: mimeType
    });
  });
  return { jsonlLines, manifestRows };
}

/**
 * Mark rows as queued in the master sheet.
 * @param {Array} rows - Array of row objects
 * @param {Sheet} master - Master sheet
 * @param {Object} cfg - Configuration object
 */
function markRowsAsQueued(rows, master, cfg) {
  const col = name => getColumnIndexByName(master, name) - 1;
  const idxStatus = col(cfg.StatusColumnName);
  rows.forEach(({ r }) => master.getRange(r+1, idxStatus+1).setValue(cfg.GeminiQueuedStatus));
}

/**
 * Poll the latest batch for results.
 * @returns {string} Status message
 */
function pollLatestBatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GeminiBatches');
  if (!sheet) throw new Error('GeminiBatches sheet not found');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('No batch records found');
  const lastRow = data[data.length - 1];
  const manifestFileId = lastRow[2]; // Column 3 is ManifestFileId
  return pollBatchById(manifestFileId);
}

/**
 * Poll for batch results by manifest ID.
 * @param {string} manifestFileId - Manifest file ID
 * @returns {string} Status message
 */
function pollBatchById(manifestFileId) {
  const cfg = getConfig();
  const apiKey = getGeminiApiKey();
  
  if (!apiKey) throw new Error("Missing GEMINI_API_KEY");

  const manifestFile = DriveApp.getFileById(manifestFileId);
  const manifest = JSON.parse(manifestFile.getBlob().getDataAsString());
  
  if (!manifest.batchName) throw new Error("Manifest missing batchName");

  const batchInfo = getBatch(cfg.geminiEndpointBase, apiKey, manifest.batchName);
  const state = batchInfo?.metadata?.state;
  
  // Log batch polling status
  const details = {
    manifestFileId: manifestFileId,
    totalCount: manifest.rows?.length || 0,
    completedCount: isTerminalState(state) ? (manifest.rows?.length || 0) : 0,
    responseTime: batchInfo?.metadata?.createTime ? 
      (new Date() - new Date(batchInfo.metadata.createTime)) : null,
    manifestCreatedAt: manifest.createdAt
  };
  
  logBatchPollingStatus(manifest.batchName, state, details);
  
  if (!isTerminalState(state)) return `Batch still running: ${state}`;

  // Terminal: download results
  const responsesFile = batchInfo?.response?.responsesFile;
  
  if (!responsesFile) {
    if (batchInfo?.response?.inlinedResponses) {
      return processInlineResults(manifest, batchInfo.response.inlinedResponses);
    }
    throw new Error("No responsesFile or inlinedResponses in batch result.");
  }

  // Download results JSONL
  const bytes = downloadResponsesFile(cfg.geminiEndpointBase, apiKey, responsesFile);
  const folder = getOrCreateFolder(cfg.BatchesFolderName);
  const out = folder.createFile(
    `results_${manifest.batchName.replace(/\//g,'_')}.jsonl`,
    Utilities.newBlob(bytes, "application/jsonl").getDataAsString(),
    "application/jsonl"
  );

  // Parse and write back to rows
  const text = out.getBlob().getDataAsString();
  return processJsonlResults(manifest, text);
}

/**
 * Process inline batch results.
 * @param {Object} manifest - Batch manifest
 * @param {Array} inlinedResponses - Array of responses
 * @returns {string} Status message
 */
function processInlineResults(manifest, inlinedResponses) {
  const lines = inlinedResponses.map((resp, i) => JSON.stringify({ key: manifest.rows[i]?.key, response: resp }));
  return processJsonlResults(manifest, lines.join("\n"));
}

/**
 * Process JSONL batch results.
 * @param {Object} manifest - Batch manifest
 * @param {string} jsonlText - JSONL results text
 * @returns {string} Status message
 */
function processJsonlResults(manifest, jsonlText) {
  const cfg = getConfig();
  const master = getOrCreateSheet(cfg.masterSheet, cfg.masterSheetColumns);
  const col = name => getColumnIndexByName(master, name) - 1;

  const idxStatus        = col(cfg.StatusColumnName);
  const idxCodeURL       = col(cfg.CodeFileUrlColumnName);
  const idxModelUsed     = col(cfg.ModelUsedColumnName);
  const idxPromptVersion = col(cfg.PromptVersionColumnName);
  const idxGenTimestamp  = col(cfg.GenerationTimestampColumnName);
  const idxInTok         = col(cfg.InputTokensColumnName);
  const idxOutTok        = col(cfg.OutputTokensColumnName);
  const idxTotTok        = col(cfg.TotalTokensColumnName);
  const idxRespTime      = col(cfg.ResponseTimeColumnName);
  const idxSafety        = col(cfg.SafetyRatingsColumnName);
  const idxFinish        = col(cfg.FinishReasonColumnName);
  const idxCite          = col(cfg.CitationMetadataColumnName);
  const idxModelVer      = col(cfg.ModelVersionColumnName);
  const idxRespId        = col(cfg.ResponseIdColumnName);
  const idxThoughtTok    = col(cfg.ThoughtsTokenCountColumnName);
  const idxTextTok       = col(cfg.TextTokenCountColumnName);
  const idxImageTok      = col(cfg.ImageTokenCountColumnName);
  const idxImgMime       = col(cfg.ImageMimeTypeColumnName);
  const idxTimestamp     = col(cfg.TimestampColumnName);
  const idxEmail         = col(cfg.EmailAddressColumnName);
  const idxProblem       = col(cfg.ProblemNumberColumnName);

  const mapKeyToRow = new Map(manifest.rows.map(x => [x.key, x.r]));
  const lines = jsonlText.split(/\r?\n/).filter(Boolean);
  let ok = 0, err = 0;

  lines.forEach((line) => {
    let obj;
    try { obj = JSON.parse(line); } catch { err++; return; }

    const key = obj.key || obj.metadata?.key;
    const r = mapKeyToRow.get(key);
    if (r == null) { err++; return; }

    const rowIndex1 = r + 1;
    if (obj.error) {
      master.getRange(rowIndex1, idxStatus+1).setValue(cfg.CannotProcessStatus);
      err++;
      return;
    }

    const parts = obj.response?.candidates?.[0]?.content?.parts || [];
    const textPart = parts.find(p => typeof p.text === "string");
    const code = textPart ? cleanCodeBlock(textPart.text) : "";

    const ts = master.getRange(rowIndex1, idxTimestamp+1).getValue();
    const email = master.getRange(rowIndex1, idxEmail+1).getValue();
    const problemCode = master.getRange(rowIndex1, idxProblem+1).getValue();
    const problemId = getProblemIdFromMeta(String(problemCode).split(" ")[0]);
    const codeUrl = saveCodeToDrive(ts, email, problemId, code);

    // Metadata
    const usage = obj.response?.usageMetadata || {};
    const finishReason = obj.response?.candidates?.[0]?.finishReason || cfg.defaultFinishReason;
    const safety = JSON.stringify(obj.response?.candidates?.[0]?.safetyRatings || []);
    const cite   = JSON.stringify(obj.response?.candidates?.[0]?.citationMetadata || {});
    const modelVersion = obj.response?.modelVersion || cfg.geminiModel;

    master.getRange(rowIndex1, idxCodeURL+1).setValue(codeUrl);
    master.getRange(rowIndex1, idxModelUsed+1).setValue(cfg.geminiModel);
    master.getRange(rowIndex1, idxPromptVersion+1).setValue(cfg.promptVersion);
    master.getRange(rowIndex1, idxGenTimestamp+1).setValue(new Date());

    master.getRange(rowIndex1, idxInTok+1).setValue(usage.promptTokenCount || 0);
    master.getRange(rowIndex1, idxOutTok+1).setValue(usage.candidatesTokenCount || 0);
    master.getRange(rowIndex1, idxTotTok+1).setValue(usage.totalTokenCount || 0);
    master.getRange(rowIndex1, idxRespTime+1).setValue(usage.responseTimeMs || "");
    master.getRange(rowIndex1, idxSafety+1).setValue(safety);
    master.getRange(rowIndex1, idxFinish+1).setValue(finishReason);
    master.getRange(rowIndex1, idxCite+1).setValue(cite);
    master.getRange(rowIndex1, idxModelVer+1).setValue(modelVersion);
    master.getRange(rowIndex1, idxRespId+1).setValue(obj.response?.responseId || "");
    master.getRange(rowIndex1, idxThoughtTok+1).setValue(usage.thoughtsTokenCount || 0);
    const ptd = usage.promptTokensDetails || [];
    master.getRange(rowIndex1, idxTextTok+1).setValue((ptd.find(d => d.modality === "TEXT")   || {}).tokenCount || 0);
    master.getRange(rowIndex1, idxImageTok+1).setValue((ptd.find(d => d.modality === "IMAGE")  || ptd.find(d => d.modality === "DOCUMENT") || {}).tokenCount || 0);

    master.getRange(rowIndex1, idxStatus+1).setValue(cfg.GeminiDoneStatus);
    ok++;
  });

  // Log the final status after processing results
  const finalStatus = "BATCH_RESULTS_PROCESSED";
  const finalDetails = {
    totalCount: lines.length,
    successCount: ok,
    errorCount: err,
    manifestFileId: manifestFileId || "" // Use the manifest file ID we already have
  };
  logBatchPollingStatus(manifest.batchName, finalStatus, finalDetails);
  
  return `Ingested results for batch ${manifest.batchName}: OK=${ok}, ERR=${err}`;
}

// ========================= UTILITY FUNCTIONS =========================

/**
 * Get or create a folder in Google Drive.
 * @param {string} name - Folder name
 * @returns {Folder} Drive folder
 */
function getOrCreateFolder(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

/**
 * Upload a JSONL file to Gemini API.
 * @param {File} driveFile - Drive file object
 * @param {string} apiKey - Gemini API key
 * @returns {string} Uploaded file name
 */
function uploadJsonlToGemini(driveFile, apiKey) {
  const baseUrl = getConfig().geminiEndpointBase;
  const bytes = driveFile.getBlob().getBytes();
  const mime = "application/jsonl";
  const displayName = "BatchInput-" + driveFile.getName();

  // start upload
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
  
  if (startResp.getResponseCode() < 200 || startResp.getResponseCode() >= 300)
    throw new Error("Resumable start failed: " + startResp.getContentText());

  const uploadUrl = startResp.getHeaders()["X-Goog-Upload-Url"] || startResp.getHeaders()["x-goog-upload-url"];
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
    }
  });
  
  if (upResp.getResponseCode() < 200 || upResp.getResponseCode() >= 300)
    throw new Error("Upload failed: " + upResp.getContentText());

  const json = JSON.parse(upResp.getContentText());
  const fileName = json?.file?.name;
  if (!fileName) throw new Error("Upload response missing file.name");
  return fileName; // e.g., "files/abc123"
}

/**
 * Create a batch with a file in Gemini API.
 * @param {string} baseUrl - Gemini API base URL
 * @param {string} apiKey - Gemini API key
 * @param {string} model - Model name
 * @param {string} fileName - Uploaded file name
 * @param {string} displayName - Display name for batch
 * @returns {Object} Batch object
 */
function createBatchWithFile(baseUrl, apiKey, model, fileName, displayName) {
  const url = `${baseUrl}/v1beta/models/${model}:batchGenerateContent?key=${encodeURIComponent(apiKey)}`;
  const body = { batch: { displayName, inputConfig: { fileName } } };
  
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300)
    throw new Error("Batch create failed: " + resp.getContentText());
  
  return JSON.parse(resp.getContentText());
}

/**
 * Get batch information from Gemini API.
 * @param {string} baseUrl - Gemini API base URL
 * @param {string} apiKey - Gemini API key
 * @param {string} batchName - Batch name
 * @returns {Object} Batch object
 */
function getBatch(baseUrl, apiKey, batchName) {
  const url = `${baseUrl}/v1beta/${batchName}`;
  
  const resp = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: { "x-goog-api-key": apiKey, "Content-Type": "application/json" }
  });
  
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300)
    throw new Error("getBatch failed: " + resp.getContentText());
  
  return JSON.parse(resp.getContentText());
}

/**
 * Check if a batch is in a terminal state.
 * @param {string} state - Batch state
 * @returns {boolean} True if terminal state
 */
function isTerminalState(state) {
  return new Set([
    "BATCH_STATE_SUCCEEDED","BATCH_STATE_FAILED","BATCH_STATE_CANCELLED","BATCH_STATE_EXPIRED",
    "JOB_STATE_SUCCEEDED","JOB_STATE_FAILED","JOB_STATE_CANCELLED","JOB_STATE_EXPIRED"
  ]).has(state);
}

/**
 * Download response file from Gemini API.
 * @param {string} baseUrl - Gemini API base URL
 * @param {string} apiKey - Gemini API key
 * @param {string} responsesFile - Responses file name
 * @returns {Bytes} File bytes
 */
function downloadResponsesFile(baseUrl, apiKey, responsesFile) {
  const url = `${baseUrl}/download/v1beta/${responsesFile}:download?alt=media`;
  
  const resp = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: { "x-goog-api-key": apiKey }
  });
  
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300)
    throw new Error("downloadResponsesFile failed: " + resp.getContentText());
  
  return resp.getBlob().getBytes();
}

/**
 * Save batch manifest to Drive.
 * @param {Folder} folder - Drive folder
 * @param {string} batchName - Batch name
 * @param {Object} manifest - Manifest object
 * @returns {File} Manifest file
 */
function saveManifest(folder, batchName, manifest) {
  const manifestName = `manifest_${batchName.replace(/\//g,'_')}.json`;
  
  // Check for existing file with same name and delete
  const files = folder.getFilesByName(manifestName);
  while (files.hasNext()) {
    const file = files.next();
    file.setTrashed(true);
  }
  
  // Create file
  const manifestJson = JSON.stringify(manifest, null, 2);
  const file = folder.createFile(manifestName, manifestJson, "application/json");
  
  return file;
}

/**
 * Track batch in GeminiBatches sheet.
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} batchName - Batch name
 * @param {string} manifestFileId - Manifest file ID
 * @param {number} rowCount - Number of rows in batch
 */
function trackBatch(ss, batchName, manifestFileId, rowCount) {
  let batchesSheet = ss.getSheetByName('GeminiBatches');
  if (!batchesSheet) {
    batchesSheet = ss.insertSheet('GeminiBatches');
    batchesSheet.appendRow(['Timestamp', 'BatchName', 'ManifestFileId', 'RowCount']);
  }
  batchesSheet.appendRow([
    new Date(),
    batchName,
    manifestFileId,
    rowCount
  ]);
}
