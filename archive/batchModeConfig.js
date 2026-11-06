/**
 * Integrated batch mode configuration that extends the main config.
 * This ensures that the batch mode configuration is properly integrated with the main config.
 */

/**
 * Updates the main CONFIG object with batch mode specific settings.
 * Call this function once at the beginning of your script execution.
 */
function initBatchModeConfig() {
  // Files are evaluated in an unspecified order. geminiConfig.js performs
  // the active initialization after CONFIG exists, so an early archive load
  // must be a no-op rather than emitting a misleading startup error.
  if (typeof CONFIG === "undefined" || !CONFIG) return false;
  // Add batch mode configuration to the main CONFIG object
  CONFIG.GeminiQueuedStatus = "GEMINI_QUEUED";
  CONFIG.BatchesFolderName = "GeminiBatches";
  CONFIG.geminiBatchThreshold = 5; // Use batch mode if requests >= this number
  
  // Make sure the geminiEndpoint is correctly set for batch mode API
  // In batch mode, we need the base URL without the /v1beta/models/ path
  CONFIG.geminiEndpointBase = "https://generativelanguage.googleapis.com";
  
  // Log the initialization
  return true;
}

// Initialize batch mode config when this script is loaded
(function() {
  try {
    initBatchModeConfig();
  } catch (e) {
    // The active configuration is initialized by geminiConfig.js.
  }
})();
