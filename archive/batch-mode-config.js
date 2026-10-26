/**
 * Config values needed for batch mode
 */

// Add these to your CONFIG object in config.js
var BATCH_MODE_CONFIG = {
  // New status for batch mode
  GeminiQueuedStatus: "GEMINI_QUEUED",
  
  // Folder for batch files in Google Drive
  BatchesFolderName: "GeminiBatches",
  
  // Threshold for using batch mode
  geminiBatchThreshold: 5, // Use batch mode if requests >= this number

  // Other constants needed for batch API
  geminiEndpoint: "https://generativelanguage.googleapis.com", // Base endpoint without /v1beta/models/
  CannotProcessStatus: "CANNOT_PROCESS", // Status for rows that can't be processed
  defaultFinishReason: "STOP" // Default finish reason if none provided
};

// UI Menu items to add
/*
function legacyArchiveOnOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Evaluate Flowcharts")
    .addItem("Enqueue Gemini Batch (NEW rows)", "enqueueGeminiBatchFromSheet")
    .addItem("Poll & Ingest Gemini Batch Results", "pollAndIngestBatchByManifestId")
    // ... other menu items ...
    .addToUi();
}
*/
