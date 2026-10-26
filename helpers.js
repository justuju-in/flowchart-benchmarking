function getProblemIdFromMeta(problemCode) {
  var metaSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getConfig().metaSheet);
  if (!metaSheet) {
    logEvent('Meta sheet not found', getConfig().logLevelError);
    return null;
  }
  var data = metaSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === problemCode) {
      // Convert numeric IDs to Number type for DOMjudge API
      var problemId = data[i][1];
      if (!isNaN(problemId)) {
        problemId = Number(problemId);
      }
      return problemId;
    }
  }
  logEvent('Problem code ' + problemCode + ' not found in Meta sheet', getConfig().logLevelError);
  return null;
}

/**
 * Clean code block from Gemini response text.
 * @param {string} text - Raw text from Gemini
 * @returns {string} Cleaned code
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
