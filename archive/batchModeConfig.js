/**
 * Integrated batch mode configuration that extends the main config.
 * This ensures that the batch mode configuration is properly integrated with the main config.
 */

/**
 * Updates the main CONFIG object with batch mode specific settings.
 * Call this function once at the beginning of your script execution.
 */
function initBatchModeConfig() {
  // Add batch mode configuration to the main CONFIG object
  CONFIG.GeminiQueuedStatus = "GEMINI_QUEUED";
  CONFIG.BatchesFolderName = "GeminiBatches";
  CONFIG.geminiBatchThreshold = 5; // Use batch mode if requests >= this number
  
  // Make sure the geminiEndpoint is correctly set for batch mode API
  // In batch mode, we need the base URL without the /v1beta/models/ path
  CONFIG.geminiEndpointBase = "https://generativelanguage.googleapis.com";
  
  // Log the initialization
  logEvent("Batch mode configuration initialized", CONFIG.logLevelDebug);
}

// Initialize batch mode config when this script is loaded
(function() {
  try {
    initBatchModeConfig();
  } catch (e) {
    // Silent initialization - errors will be caught later when functions are called
    console.log("Error initializing batch mode config:", e);
  }
})();
