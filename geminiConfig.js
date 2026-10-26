/**
 * Unified Gemini configuration file.
 * Contains all configuration settings for both individual and batch processing.
 */

/**
 * Initialize Gemini configuration settings.
 * Adds Gemini-specific configuration to the main CONFIG object.
 */
function initGeminiConfig() {
  // Add Gemini API configuration
  CONFIG.geminiApiKeyProperty = "GEMINI_API_KEY";
  CONFIG.geminiModel = "gemini-2.5-pro";
  CONFIG.geminiTemperature = 0;
  CONFIG.geminiEndpoint = "https://generativelanguage.googleapis.com/v1beta/models/";
  
  // Add batch mode configuration
  CONFIG.IsBatchMode = true;
  CONFIG.maxRowCount = 100;
  CONFIG.GeminiQueuedStatus = "GEMINI_QUEUED";
  CONFIG.BatchesFolderName = "GeminiBatches";
  CONFIG.geminiBatchThreshold = 5; // Use batch mode if requests >= this number
  CONFIG.geminiEndpointBase = "https://generativelanguage.googleapis.com";
  CONFIG.CannotProcessStatus = "CANNOT_PROCESS";
  CONFIG.defaultFinishReason = "STOP";
  
  // Add batch performance optimization settings
  CONFIG.batchRetryLimit = 3;
  CONFIG.batchMaxSizeBytes = 10 * 1024 * 1024; // 10MB max batch size
  CONFIG.batchPollingIntervalSeconds = 60;
  
  // Initialize batch analytics
  if (!CONFIG.batchAnalytics) {
    CONFIG.batchAnalytics = {
      totalBatchesProcessed: 0,
      totalItemsProcessed: 0,
      totalTokensSaved: 0,
      estimatedCostSaved: 0
    };
  }
}

/**
 * Get Gemini API key from script properties.
 * @returns {string|null} The API key or null if not found
 */
function getGeminiApiKey() {
  return PropertiesService.getScriptProperties().getProperty(CONFIG.geminiApiKeyProperty);
}

// Initialize configuration when this file is loaded
(function() {
  try {
    initGeminiConfig();
  } catch (e) {
    console.log("Error initializing Gemini config:", e);
  }
})();
