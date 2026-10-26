/**
 * Integration module that bridges between single-item and batch processing modes.
 * This file serves as the interface between the main pipeline and the batch processing system.
 */

/**
 * Dispatches Gemini processing tasks to either batch or individual processing
 * based on configuration settings and the number of eligible rows.
 * This is the main entry point that should be called from the UI or triggers.
 */
function processGeminiWithModeSelection() {
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
    return enqueueGeminiBatchFromSheet(cfg.maxRowCount); 
  } else {
    // Use single item processing
    logEvent(`Using single item processing for ${eligibleRows.length} rows`, cfg.logLevelDebug);
    return processItemsIndividually(masterSheet, data, eligibleRows);
  }
}

/**
 * Processes multiple items using the original single-item pipeline.
 * @param {Sheet} sheet - The master sheet
 * @param {Array[]} data - Sheet data as 2D array
 * @param {number[]} rowIndices - Array of eligible row indices to process
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
 * Polls for Gemini batch results and processes them.
 * This polls for a specific batch result or the latest batch if no manifest ID is provided.
 * @param {string} manifestFileId - Optional manifest file ID to poll for
 */
function pollGeminiBatchResults(manifestFileId) {
  if (!manifestFileId) {
    return pollAndIngestLatestBatch();
  } else {
    return pollAndIngestBatchByManifestId(manifestFileId);
  }
}

/**
 * Main entry point for Gemini processing with batch support.
 * This function replaces the original triggerGeminiProcessing with improved
 * capabilities for both individual and batch processing.
 */
function triggerGeminiProcessingV2() {
  return processGeminiWithModeSelection();
}
