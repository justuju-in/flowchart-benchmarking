/**
 * Core functions for batch mode processing
 */

// Helper to poll and ingest the latest batch automatically
function pollAndIngestLatestBatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GeminiBatches');
  if (!sheet) throw new Error('GeminiBatches sheet not found');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('No batch records found');
  const lastRow = data[data.length - 1];
  const manifestFileId = lastRow[2]; // Column 3 is ManifestFileId
  return pollAndIngestBatchByManifestId(manifestFileId);
}

// Main entry point for enqueueing rows for batch processing
function enqueueGeminiBatchFromSheet(maxRows) {
  const cfg = getConfig();
  const apiKey = PropertiesService.getScriptProperties().getProperty(cfg.geminiApiKeyProperty);
  console.log("enqueueGeminiBatchFromSheet: config", cfg);
  console.log("enqueueGeminiBatchFromSheet: apiKey", apiKey);
  if (!apiKey) throw new Error("Missing GEMINI_API_KEY in Script Properties.");

  const master = getOrCreateSheet(cfg.masterSheet, cfg.masterSheetColumns);
  const values = master.getDataRange().getValues();
  console.log("enqueueGeminiBatchFromSheet: values.length", values.length);
  if (values.length <= 1) return "No data";

  const rows = getRowsToEnqueue(master, cfg, maxRows);
  console.log("enqueueGeminiBatchFromSheet: rows to enqueue", rows.length);
  if (!rows.length) return "No NEW rows to enqueue.";

  const promptText = getPrompt();
  const { jsonlLines, manifestRows } = buildJsonlLinesAndManifest(rows, master, cfg, promptText);
  console.log("enqueueGeminiBatchFromSheet: jsonlLines.length", jsonlLines.length);
  if (!jsonlLines.length) return "No valid files among NEW rows.";

  const manifest = {
    createdAt: new Date().toISOString(),
    model: cfg.geminiModel,
    sheetName: cfg.masterSheet,
    rows: manifestRows,
    batchName: null,
    responsesFile: null
  };

  const folder = getOrCreateFolder_(cfg.BatchesFolderName);
  const jsonlName = `batch_input_${new Date().toISOString().replace(/[:.]/g, '-')}.jsonl`;
  const jsonlFile = folder.createFile(jsonlName, jsonlLines.join("\n"), "application/jsonl");
  console.log("enqueueGeminiBatchFromSheet: jsonlFile id", jsonlFile.getId());

  const uploadedFileName = uploadJsonlToGemini_(jsonlFile, apiKey);
  console.log("enqueueGeminiBatchFromSheet: uploadedFileName", uploadedFileName);
  const batch = createBatchWithFile_(cfg.geminiEndpoint, apiKey, cfg.geminiModel, uploadedFileName, "flowchart-judge");
  const batchName = batch.name;
  manifest.batchName = batchName;
  console.log("enqueueGeminiBatchFromSheet: batchName", batchName);

  markRowsQueued(rows, master, cfg);

  const manifestFile = saveManifest(folder, batchName, manifest);
  console.log("enqueueGeminiBatchFromSheet: manifestFile id", manifestFile.getId());

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  trackBatch(ss, batchName, manifestFile.getId(), rows.length);

  return `Enqueued ${rows.length} rows. Batch=${batchName}. ManifestId=${manifestFile.getId()}`;
}

// Poll for batch results and process them
function pollAndIngestBatchByManifestId(manifestFileId) {
  const cfg = getConfig();
  const apiKey = PropertiesService.getScriptProperties().getProperty(cfg.geminiApiKeyProperty);
  console.log("pollAndIngestBatchByManifestId: manifestFileId", manifestFileId);
  if (!apiKey) throw new Error("Missing GEMINI_API_KEY");

  const manifestFile = DriveApp.getFileById(manifestFileId);
  const manifest = JSON.parse(manifestFile.getBlob().getDataAsString());
  console.log("pollAndIngestBatchByManifestId: manifest", manifest);
  if (!manifest.batchName) throw new Error("Manifest missing batchName");

  const batchInfo = getBatch_(cfg.geminiEndpoint, apiKey, manifest.batchName);
  console.log("pollAndIngestBatchByManifestId: batchInfo", batchInfo);
  const state = batchInfo?.metadata?.state;
  console.log(`Batch ${manifest.batchName} state = ${state}`);

  if (!isTerminalState_(state)) return `Batch still running: ${state}`;

  // Terminal: download results
  const responsesFile = batchInfo?.response?.responsesFile;
  console.log("pollAndIngestBatchByManifestId: responsesFile", responsesFile);
  if (!responsesFile) {
    if (batchInfo?.response?.inlinedResponses) {
      console.log("pollAndIngestBatchByManifestId: inlinedResponses", batchInfo.response.inlinedResponses);
      return ingestInlinedResults_(manifest, batchInfo.response.inlinedResponses);
    }
    throw new Error("No responsesFile or inlinedResponses in batch result.");
  }

  // Download results JSONL
  const bytes = downloadResponsesFile_(cfg.geminiEndpoint, apiKey, responsesFile);
  const folder = getOrCreateFolder_(cfg.BatchesFolderName);
  const out = folder.createFile(
    `results_${manifest.batchName.replace(/\//g,'_')}.jsonl`,
    Utilities.newBlob(bytes, "application/jsonl").getDataAsString(),
    "application/jsonl"
  );
  console.log("pollAndIngestBatchByManifestId: results file id", out.getId());

  // Parse and write back to rows
  const text = out.getBlob().getDataAsString();
  return ingestJsonlResultsToSheet_(manifest, text);
}

// Helper functions for batch processing

// Find rows that are ready to be processed in batch
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

// Prepare data for batch processing
function buildJsonlLinesAndManifest(rows, master, cfg, promptText) {
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
      console.log(`buildJsonlLinesAndManifest: Skipping row ${r+1}, invalid fileId for URL: ${flowchartUrl}`);
      return;
    }
    let file, blob, mimeType;
    try {
      file = DriveApp.getFileById(fileId);
      blob = file.getBlob();
      mimeType = blob.getContentType();
      console.log(`buildJsonlLinesAndManifest: Row ${r+1} mimeType:`, mimeType);
    } catch (e) {
      console.log(`buildJsonlLinesAndManifest: Skipping row ${r+1}, error getting file or mimeType: ${e}`);
      return;
    }
    if (!mimeType) {
      console.log(`buildJsonlLinesAndManifest: Skipping row ${r+1}, mimeType is null for fileId: ${fileId}`);
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

// Update row status to GEMINI_QUEUED
function markRowsQueued(rows, master, cfg) {
  const col = name => getColumnIndexByName(master, name) - 1;
  const idxStatus = col(cfg.StatusColumnName);
  rows.forEach(({ r }) => master.getRange(r+1, idxStatus+1).setValue(cfg.GeminiQueuedStatus));
}

// Save batch manifest to file
function saveManifest(folder, batchName, manifest) {
  const manifestName = `manifest_${batchName.replace(/\//g,'_')}.json`;
  // Debug: Log manifest name and folder
  try {
    console.log("saveManifest: manifestName", manifestName);
    console.log("saveManifest: folder id", folder && folder.getId ? folder.getId() : folder);
    // Check for existing file with same name and delete
    const files = folder.getFilesByName(manifestName);
    while (files.hasNext()) {
      const file = files.next();
      file.setTrashed(true);
      console.log("saveManifest: Deleted existing manifest file", file.getId());
    }
    // Validate manifest object
    let manifestJson;
    try {
      manifestJson = JSON.stringify(manifest, null, 2);
    } catch (e) {
      console.log("saveManifest: Error serializing manifest", e);
      throw new Error("Manifest serialization failed: " + e);
    }
    console.log("saveManifest: manifest JSON", manifestJson);
    // Create file
    const file = folder.createFile(manifestName, manifestJson, "application/json");
    console.log("saveManifest: Created manifest file", file.getId());
    return file;
  } catch (e) {
    console.log("saveManifest: Error creating manifest file", e);
    throw e;
  }
}

// Track batch in GeminiBatches sheet
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

// Process inline responses from Gemini
function ingestInlinedResults_(manifest, inlinedResponses) {
  const lines = inlinedResponses.map((resp, i) => JSON.stringify({ key: manifest.rows[i]?.key, response: resp }));
  return ingestJsonlResultsToSheet_(manifest, lines.join("\n"));
}

// Process JSONL results and update sheet
function ingestJsonlResultsToSheet_(manifest, jsonlText) {
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

  console.log("ingestJsonlResultsToSheet_: lines.length", lines.length);
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

  console.log(`ingestJsonlResultsToSheet_: OK=${ok}, ERR=${err}`);
  return `Ingested results for batch ${manifest.batchName}: OK=${ok}, ERR=${err}`;
}
