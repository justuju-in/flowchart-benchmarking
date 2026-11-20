function createBenchmarkRun() {
  return createConfiguredBenchmarkRun(
    'BenchmarkSetA',
    'gemini-3.1-pro-preview',
    'medium'
  );
}

/** Creates a run from explicitly selected benchmark settings. */
function createConfiguredBenchmarkRun(setName, model, reasoningLevel, sampleSize) {
  const SPREADSHEET_ID = '1U-FahEyZvmMa5Y2uaLmoBUDL-q-Zqp14ss2fYqFwWU4';
  const SET_NAME = String(setName || '').trim();
  const MODEL = String(model || '').trim();
  const REASONING_LEVEL = String(reasoningLevel || '').trim();
  const SAMPLE_SIZE = sampleSize == null || sampleSize === '' ? 5 : Number(sampleSize);

  if (!/^BenchmarkSet[AB]$/.test(SET_NAME)) {
    throw new Error('Choose BenchmarkSetA or BenchmarkSetB.');
  }
  if (!MODEL || !REASONING_LEVEL) {
    throw new Error('Model and reasoning level are required.');
  }
  if (!Number.isInteger(SAMPLE_SIZE) || SAMPLE_SIZE < 1 || SAMPLE_SIZE > 10) {
    throw new Error('For the pilot, choose between 1 and 10 items.');
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const setSheet = ss.getSheetByName(SET_NAME);
  const runsSheet = ss.getSheetByName('BenchmarkRuns');

  if (!setSheet) {
    throw new Error(`${SET_NAME} tab not found.`);
  }

  if (!runsSheet) {
    throw new Error('BenchmarkRuns tab not found.');
  }

  const itemCount = Math.max(0, setSheet.getLastRow() - 1);

  if (itemCount === 0) {
    throw new Error(`${SET_NAME} has no items.`);
  }
  if (SAMPLE_SIZE > itemCount) throw new Error('Sample size exceeds available set items.');
  const sampleHeader = runsSheet.getRange(1, 8).getValue();
  if (sampleHeader && sampleHeader !== 'Sample Size') {
    throw new Error('BenchmarkRuns column H is already in use; cannot save sample size safely.');
  }
  if (!sampleHeader) runsSheet.getRange(1, 8).setValue('Sample Size');

  const existingRuns = runsSheet
    .getDataRange()
    .getValues()
    .slice(1)
    .map(row => String(row[0] || ''));

  const highestRunNumber = existingRuns.reduce((highest, runId) => {
    const match = runId.match(/^RUN-(\d+)$/);
    return match ? Math.max(highest, Number(match[1])) : highest;
  }, 0);

  const runId = `RUN-${String(highestRunNumber + 1).padStart(3, '0')}`;

  runsSheet.appendRow([
    runId,
    SET_NAME,
    MODEL,
    REASONING_LEVEL,
    'NEW',
    new Date(),
    '',
    SAMPLE_SIZE
  ]);

  ensureBenchmarkResultsCostColumn_(ss);
  return `${runId} created for the first ${SAMPLE_SIZE} ${SET_NAME} items.`;
}

/** Opens a small dialog so a user can select a set, model and reasoning level. */
function showBenchmarkRunDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>body{font:14px Arial;padding:16px}label{display:block;margin-top:12px}select,input,button{width:100%;box-sizing:border-box;padding:8px;margin-top:4px}button{margin-top:18px;background:#1a73e8;color:#fff;border:0;border-radius:4px}</style>
    <h3>Start benchmark run</h3>
    <label>Benchmark set<select id="set"><option>BenchmarkSetA</option><option>BenchmarkSetB</option></select></label>
    <label>Gemini model<input id="model" value="gemini-3.1-pro-preview"></label>
    <label>Reasoning level<select id="reasoning"><option>low</option><option selected>medium</option><option>high</option></select></label>
    <label>Sample size<input id="sample" type="number" min="1" max="10" value="5"></label>
    <button onclick="start()">Create run</button><p id="message"></p>
    <script>function start(){const b=document.querySelector('button');b.disabled=true;google.script.run.withSuccessHandler(r=>{document.getElementById('message').textContent=r;b.disabled=false;}).withFailureHandler(e=>{document.getElementById('message').textContent=e.message;b.disabled=false;}).createConfiguredBenchmarkRun(document.getElementById('set').value,document.getElementById('model').value,document.getElementById('reasoning').value,document.getElementById('sample').value);}</script>`)
    .setWidth(360).setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, 'Start benchmark run');
}

function ensureBenchmarkResultsCostColumn_(ss) {
  const sheet = ss.getSheetByName('BenchmarkResults');
  if (!sheet) return;
  const headerCell = sheet.getRange(1, 19);
  if (!headerCell.getValue()) headerCell.setValue('Estimated Cost (USD)');
}

function getBenchmarkEstimatedCost_(ss, model, usage) {
  const pricingSheet = ss.getSheetByName('ModelPricing');
  if (!pricingSheet) return '';
  const priceRows = pricingSheet.getDataRange().getValues().slice(1);
  const price = priceRows.find(row => String(row[0]).trim() === String(model).trim());
  if (!price) return '';

  const inputRate = Number(price[1]);
  const outputRate = Number(price[2]);
  if (!Number.isFinite(inputRate) || !Number.isFinite(outputRate)) return '';
  const inputTokens = Number(usage.promptTokenCount || 0);
  const outputTokens = Number(usage.candidatesTokenCount || 0) + Number(usage.thoughtsTokenCount || 0);
  return Number(((inputTokens * inputRate + outputTokens * outputRate) / 1000000).toFixed(9));
}
function queueBenchmarkRun() {
  const SPREADSHEET_ID = '1U-FahEyZvmMa5Y2uaLmoBUDL-q-Zqp14ss2fYqFwWU4';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const runsSheet = ss.getSheetByName('BenchmarkRuns');
  const resultsSheet = ss.getSheetByName('BenchmarkResults');

  if (!runsSheet || !resultsSheet) {
    throw new Error('BenchmarkRuns or BenchmarkResults tab not found.');
  }

  // Latest NEW run find karo.
  const runValues = runsSheet.getDataRange().getValues();
  let runRow = -1;

  for (let i = runValues.length - 1; i >= 1; i--) {
    if (String(runValues[i][4]).trim() === 'NEW') {
      runRow = i + 1;
      break;
    }
  }

  if (runRow === -1) {
    throw new Error('No NEW benchmark run found.');
  }

  const run = runValues[runRow - 1];
  const runId = run[0];
  const setName = run[1];
  const model = run[2];
  const reasoningLevel = run[3];

  const setSheet = ss.getSheetByName(setName);
  if (!setSheet) {
    throw new Error(`${setName} tab not found.`);
  }

  // Same run ko dobara queue hone se roko.
  const resultLastRow = resultsSheet.getLastRow();
  if (resultLastRow > 1) {
    const existingRunIds = resultsSheet
      .getRange(2, 1, resultLastRow - 1, 1)
      .getValues()
      .flat();

    if (existingRunIds.includes(runId)) {
      throw new Error(`${runId} is already queued.`);
    }
  }

  const setValues = setSheet.getDataRange().getValues();
  if (setValues.length < 2) {
    throw new Error(`${setName} has no benchmark items.`);
  }

  const requestedSize = run[7] === '' || run[7] == null ? setValues.length - 1 : Number(run[7]);
  if (!Number.isInteger(requestedSize) || requestedSize < 1 || requestedSize > setValues.length - 1) {
    throw new Error('Invalid or unavailable sample size for this run.');
  }
  const items = setValues.slice(1, requestedSize + 1);

  const queuedRows = items.map((item, index) => [
    runId,               // Run ID
    setName,             // Set Name
    index + 2,           // Item Row in BenchmarkSetA
    item[3],             // Problem ID
    item[4],             // Flowchart URL
    item[5],             // Baseline Submission ID
    item[6],             // Baseline Verdict
    model,               // Model
    reasoningLevel,      // Reasoning Level
    'QUEUED',            // Status
    '', '', '', '', '', '', '', '' // actual result fields abhi blank
  ]);

  resultsSheet
    .getRange(resultsSheet.getLastRow() + 1, 1, queuedRows.length, queuedRows[0].length)
    .setValues(queuedRows);

  // BenchmarkRuns mein status update.
  runsSheet.getRange(runRow, 5).setValue('QUEUED');

  return `${runId} queued with ${queuedRows.length} items.`;
}
function generateBenchmarkCode(fileId, model, reasoningLevel) {
  const apiKey = getGeminiApiKey();

  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is missing.');
  }

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const startTime = new Date().getTime();

  const payload = {
    contents: [{
      parts: [
        {
          inline_data: {
            mime_type: blob.getContentType(),
            data: Utilities.base64Encode(blob.getBytes())
          }
        },
        { text: getPrompt() }
      ]
    }],
    generationConfig: {
      temperature: 0,
      thinkingConfig: {
        thinkingLevel: reasoningLevel
      }
    }
  };

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    model +
    ':generateContent?key=' +
    apiKey;

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const responseTimeMs = new Date().getTime() - startTime;

  if (response.getResponseCode() !== 200) {
    throw new Error('Gemini error: ' + response.getContentText());
  }

  const json = JSON.parse(response.getContentText());
  const text = json.candidates?.[0]?.content?.parts?.find(
    part => typeof part.text === 'string'
  )?.text;

  if (!text) {
    throw new Error('Gemini returned no generated code.');
  }

  return {
    code: cleanCodeBlock(text),
    responseTimeMs: responseTimeMs,
    usage: json.usageMetadata || {}
  };
}
function getBenchmarkProblemId(ss, problemCode) {
  const metaSheet = ss.getSheetByName('Meta');

  if (!metaSheet) {
    throw new Error('Meta tab not found.');
  }

  const metaRows = metaSheet.getDataRange().getValues();

  for (let i = 1; i < metaRows.length; i++) {
    if (String(metaRows[i][0]).trim() === problemCode) {
      return metaRows[i][1];
    }
  }

  throw new Error(`Problem ID not found in Meta for ${problemCode}.`);
}

/**
 * Keep benchmark submissions on the same DOMjudge path as evaluation.
 * Meta provides the internal ID, while oldjudge expects the public FCP code.
 */
function getBenchmarkSubmissionProblemId_(ss, problemCode) {
  const internalId = getBenchmarkProblemId(ss, problemCode);
  return getDomjudgeProblemIdentifier_(problemCode, internalId, getConfig());
}


function processOneBenchmarkItem() {
  const SPREADSHEET_ID = '1U-FahEyZvmMa5Y2uaLmoBUDL-q-Zqp14ss2fYqFwWU4';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resultsSheet = ss.getSheetByName('BenchmarkResults');
  const runsSheet = ss.getSheetByName('BenchmarkRuns');

  const rows = resultsSheet.getDataRange().getValues();

  // First QUEUED item find karo.
  let rowNumber = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][9]).trim() === 'QUEUED') {
      rowNumber = i + 1;
      break;
    }
  }

  if (rowNumber === -1) {
    throw new Error('No QUEUED benchmark item found.');
  }

  const item = rows[rowNumber - 1];
  const runId = item[0];
  const problemText = item[3];
  const flowchartUrl = item[4];
  const model = item[7];
  const reasoningLevel = item[8];
  const problemCode = String(problemText).split(' ')[0].trim();

  // Status = PROCESSING
  resultsSheet.getRange(rowNumber, 10).setValue('PROCESSING');

  try {
    const fileId = getFileIdFromUrl(flowchartUrl);

    if (!fileId) {
      throw new Error('Invalid Flowchart URL.');
    }

    const geminiResult = generateBenchmarkCode(
      fileId,
      model,
      reasoningLevel
    );

    const codeUrl = saveCodeToDrive(
      new Date(),
      'benchmark',
      problemCode,
      geminiResult.code
    );

    if (!codeUrl) {
      throw new Error('Generated code could not be saved to Drive.');
    }

    const domjudgeProblemId = getBenchmarkSubmissionProblemId_(ss, problemCode);
    const submissionId = submitToDomjudge(
      geminiResult.code,
      domjudgeProblemId
    );

    if (!submissionId) {
      throw new Error('DOMjudge submission failed.');
    }

    const usage = geminiResult.usage;
    const estimatedCost = getBenchmarkEstimatedCost_(ss, model, usage);
    ensureBenchmarkResultsCostColumn_(ss);

    // Result row update
    resultsSheet.getRange(rowNumber, 11, 1, 7).setValues([[
      codeUrl,                                    // K Generated Code URL
      submissionId,                               // L DOMjudge Submission ID
      '',                                         // M New Verdict (polling ke baad)
      usage.promptTokenCount || 0,                // N Input Tokens
      usage.candidatesTokenCount || 0,            // O Output Tokens
      usage.totalTokenCount || 0,                 // P Total Tokens
      geminiResult.responseTimeMs                 // Q Response Time
    ]]);

    resultsSheet.getRange(rowNumber, 10)
      .setValue('JUDGE_SUBMITTED');
    resultsSheet.getRange(rowNumber, 19).setValue(estimatedCost);

    // Run ko RUNNING mark karo.
    const runRows = runsSheet.getDataRange().getValues();

    for (let i = 1; i < runRows.length; i++) {
      if (runRows[i][0] === runId) {
        runsSheet.getRange(i + 1, 5).setValue('RUNNING');
        break;
      }
    }

    return `${runId}: ${problemCode} submitted to DOMjudge.`;

  } catch (error) {
    resultsSheet.getRange(rowNumber, 10).setValue('ERROR');
    resultsSheet.getRange(rowNumber, 18)
      .setValue(String(error.message || error));

    throw error;
  }
}
function generateOneBenchmarkItem(targetRunId) {
  const SPREADSHEET_ID = '1U-FahEyZvmMa5Y2uaLmoBUDL-q-Zqp14ss2fYqFwWU4';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resultsSheet = ss.getSheetByName('BenchmarkResults');
  const runsSheet = ss.getSheetByName('BenchmarkRuns');

  if (!resultsSheet || !runsSheet) {
    throw new Error('BenchmarkResults or BenchmarkRuns tab not found.');
  }

  const rows = resultsSheet.getDataRange().getValues();
  let rowNumber = -1;

  // Sirf pehla QUEUED item generate hoga.
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][9]).trim() === 'QUEUED' && (!targetRunId || String(rows[i][0]).trim() === String(targetRunId).trim())) {
      rowNumber = i + 1;
      break;
    }
  }

  if (rowNumber === -1) {
    throw new Error(targetRunId ? `No QUEUED benchmark item found for ${targetRunId}.` : 'No QUEUED benchmark item found.');
  }

  const item = rows[rowNumber - 1];
  const runId = item[0];
  const problemText = item[3];
  const flowchartUrl = item[4];
  const model = item[7];
  const reasoningLevel = item[8];
  const problemCode = String(problemText).split(' ')[0].trim();

  resultsSheet.getRange(rowNumber, 10).setValue('GENERATING');

  try {
    const fileId = getFileIdFromUrl(flowchartUrl);

    if (!fileId) {
      throw new Error('Invalid Flowchart URL.');
    }

    const geminiResult = generateBenchmarkCode(
      fileId,
      model,
      reasoningLevel
    );

    const codeUrl = saveCodeToDrive(
      new Date(),
      'benchmark',
      problemCode,
      geminiResult.code
    );

    if (!codeUrl) {
      throw new Error('Generated code could not be saved to Drive.');
    }

    const usage = geminiResult.usage;
    const estimatedCost = getBenchmarkEstimatedCost_(ss, model, usage);
    ensureBenchmarkResultsCostColumn_(ss);

    // K–Q: generated code aur Gemini details save hongi.
    resultsSheet.getRange(rowNumber, 11, 1, 7).setValues([[
      codeUrl,
      '',
      '',
      usage.promptTokenCount || 0,
      usage.candidatesTokenCount || 0,
      usage.totalTokenCount || 0,
      geminiResult.responseTimeMs
    ]]);

    resultsSheet.getRange(rowNumber, 10).setValue('GENERATED');
    resultsSheet.getRange(rowNumber, 18).setValue('');
    resultsSheet.getRange(rowNumber, 19).setValue(estimatedCost);

    // Parent run ko RUNNING mark karo.
    const runRows = runsSheet.getDataRange().getValues();
    for (let i = 1; i < runRows.length; i++) {
      if (String(runRows[i][0]) === String(runId)) {
        runsSheet.getRange(i + 1, 5).setValue('RUNNING');
        break;
      }
    }

    return `${runId}: ${problemCode} code generated and saved.`;

  } catch (error) {
    resultsSheet.getRange(rowNumber, 10).setValue('GENERATION_ERROR');
    resultsSheet.getRange(rowNumber, 18)
      .setValue(String(error.message || error));

    throw error;
  }
}
function submitOneBenchmarkItem(targetRunId) {
  const SPREADSHEET_ID = '1U-FahEyZvmMa5Y2uaLmoBUDL-q-Zqp14ss2fYqFwWU4';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resultsSheet = ss.getSheetByName('BenchmarkResults');
  const rows = resultsSheet.getDataRange().getValues();
  let rowNumber = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][9]).trim() === 'GENERATED' && (!targetRunId || String(rows[i][0]).trim() === String(targetRunId).trim())) {
      rowNumber = i + 1;
      break;
    }
  }

  if (rowNumber === -1) {
    throw new Error(targetRunId ? `No GENERATED item found for ${targetRunId}.` : 'No GENERATED item found.');
  }

  const item = rows[rowNumber - 1];
  const problemCode = String(item[3]).split(' ')[0].trim();
  const codeUrl = item[10];

  // Submit se pehle status save karte hain.
  resultsSheet.getRange(rowNumber, 10).setValue('SUBMITTING');

  try {
    const codeFileId = getFileIdFromUrl(codeUrl);
    const code = DriveApp.getFileById(codeFileId)
      .getBlob()
      .getDataAsString();

    const domjudgeProblemId = getBenchmarkSubmissionProblemId_(ss, problemCode);
    const submissionId = submitToDomjudge(code, domjudgeProblemId);

    if (!submissionId) {
      throw new Error('DOMjudge submission failed; see event-log for the response.');
    }

    resultsSheet.getRange(rowNumber, 12).setValue(String(submissionId));
    resultsSheet.getRange(rowNumber, 10).setValue('JUDGE_SUBMITTED');
    resultsSheet.getRange(rowNumber, 18).setValue('');

    return `${problemCode} submitted: ${submissionId}`;

  } catch (error) {
    resultsSheet.getRange(rowNumber, 10).setValue('SUBMISSION_ERROR');
    resultsSheet.getRange(rowNumber, 18)
      .setValue(String(error.message || error));

    throw error;
  }
}

/**
 * Fetches the DOMjudge verdict for the first submitted benchmark item.
 * It deliberately uses the central polling helper, so benchmark and normal
 * flowchart evaluation call DOMjudge in the same way.
 */
function pollOneBenchmarkVerdict(targetRunId) {
  const SPREADSHEET_ID = '1U-FahEyZvmMa5Y2uaLmoBUDL-q-Zqp14ss2fYqFwWU4';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resultsSheet = ss.getSheetByName('BenchmarkResults');
  const rows = resultsSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const item = rows[i];
    if (String(item[9]).trim() !== 'JUDGE_SUBMITTED') continue;
    if (targetRunId && String(item[0]).trim() !== String(targetRunId).trim()) continue;

    const rowNumber = i + 1;
    const problemCode = String(item[3]).split(' ')[0].trim();
    const submissionId = String(item[11] || '').trim();
    if (!submissionId) {
      throw new Error(`${problemCode} is JUDGE_SUBMITTED but has no DOMjudge submission ID.`);
    }

    const domjudgeProblemId = getBenchmarkSubmissionProblemId_(ss, problemCode);
    const verdict = pollJudgement(submissionId, domjudgeProblemId);
    if (!verdict) {
      return `${problemCode} (${submissionId}) is still waiting for a verdict.`;
    }

    resultsSheet.getRange(rowNumber, 13).setValue(verdict); // New Verdict
    resultsSheet.getRange(rowNumber, 10).setValue('VERDICT_READY');
    resultsSheet.getRange(rowNumber, 18).setValue('');
    return `${problemCode} (${submissionId}) verdict: ${verdict}`;
  }

  throw new Error(targetRunId ? `No JUDGE_SUBMITTED benchmark item found for ${targetRunId}.` : 'No JUDGE_SUBMITTED benchmark item found.');
}

function promptBenchmarkRunId_(action) {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt('Benchmark Runner', `Enter the Run ID to ${action}, for example RUN-002:`, ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return '';
  const runId = String(prompt.getResponseText() || '').trim();
  if (!/^RUN-\d+$/.test(runId)) throw new Error('Enter a valid Run ID, for example RUN-002.');
  return runId;
}

function generateNextBenchmarkItemForRun() {
  const runId = promptBenchmarkRunId_('generate the next item');
  return runId ? generateOneBenchmarkItem(runId) : '';
}

function submitNextBenchmarkItemForRun() {
  const runId = promptBenchmarkRunId_('submit the next generated item');
  return runId ? submitOneBenchmarkItem(runId) : '';
}

function pollNextBenchmarkVerdictForRun() {
  const runId = promptBenchmarkRunId_('poll the next verdict');
  return runId ? pollOneBenchmarkVerdict(runId) : '';
}

/** Retries the first failed submission that already has generated code. */
function retryOneBenchmarkSubmission() {
  const SPREADSHEET_ID = '1U-FahEyZvmMa5Y2uaLmoBUDL-q-Zqp14ss2fYqFwWU4';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resultsSheet = ss.getSheetByName('BenchmarkResults');
  const rows = resultsSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const item = rows[i];
    if (String(item[9]).trim() !== 'SUBMISSION_ERROR' || !String(item[10]).trim()) continue;

    const rowNumber = i + 1;
    const problemCode = String(item[3]).split(' ')[0].trim();
    resultsSheet.getRange(rowNumber, 10).setValue('SUBMITTING');
    try {
      const code = DriveApp.getFileById(getFileIdFromUrl(item[10])).getBlob().getDataAsString();
      const submissionId = submitToDomjudge(code, getBenchmarkSubmissionProblemId_(ss, problemCode));
      if (!submissionId) throw new Error('DOMjudge submission failed; see event-log for the response.');
      resultsSheet.getRange(rowNumber, 12).setValue(String(submissionId));
      resultsSheet.getRange(rowNumber, 10).setValue('JUDGE_SUBMITTED');
      resultsSheet.getRange(rowNumber, 18).setValue('');
      return `${problemCode} retried as submission ${submissionId}`;
    } catch (error) {
      resultsSheet.getRange(rowNumber, 10).setValue('SUBMISSION_ERROR');
      resultsSheet.getRange(rowNumber, 18).setValue(String(error.message || error));
      throw error;
    }
  }
  throw new Error('No retryable SUBMISSION_ERROR item with generated code found.');
}

/** Prompts for two run IDs, then creates a row-level comparison report. */
function promptBenchmarkRunComparison() {
  const ui = SpreadsheetApp.getUi();
  const first = ui.prompt('Benchmark comparison', 'Enter first Run ID, for example RUN-002:', ui.ButtonSet.OK_CANCEL);
  if (first.getSelectedButton() !== ui.Button.OK) return;
  const second = ui.prompt('Benchmark comparison', 'Enter second Run ID:', ui.ButtonSet.OK_CANCEL);
  if (second.getSelectedButton() !== ui.Button.OK) return;
  const message = buildBenchmarkRunComparisonReport(first.getResponseText(), second.getResponseText());
  ui.alert(message);
}

/** Builds a report with verdict changes for any two benchmark runs. */
function buildBenchmarkRunComparisonReport(runA, runB) {
  const SPREADSHEET_ID = '1U-FahEyZvmMa5Y2uaLmoBUDL-q-Zqp14ss2fYqFwWU4';
  const firstRun = String(runA || '').trim();
  const secondRun = String(runB || '').trim();
  if (!firstRun || !secondRun || firstRun === secondRun) {
    throw new Error('Provide two different Run IDs.');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const resultRows = ss.getSheetByName('BenchmarkResults').getDataRange().getValues();
  const byRun = {};
  resultRows.slice(1).forEach(row => {
    const id = String(row[0]).trim();
    if (id !== firstRun && id !== secondRun) return;
    const key = `${row[2]}|${String(row[3]).split(' ')[0]}`;
    if (!byRun[id]) byRun[id] = {};
    byRun[id][key] = row;
  });
  if (!Object.keys(byRun[firstRun] || {}).length || !Object.keys(byRun[secondRun] || {}).length) {
    throw new Error('Both Run IDs must have rows in BenchmarkResults.');
  }

  const setsA = [...new Set(Object.values(byRun[firstRun]).map(row => String(row[1])))];
  const setsB = [...new Set(Object.values(byRun[secondRun]).map(row => String(row[1])))];
  if (setsA.length !== 1 || setsB.length !== 1 || setsA[0] !== setsB[0]) {
    throw new Error('Both runs must use the same benchmark set.');
  }
  const keys = [...new Set([...Object.keys(byRun[firstRun]), ...Object.keys(byRun[secondRun])])].sort();
  const rows = [[
    'Run A', 'Run B', 'Item Row', 'Problem ID', 'Verdict A', 'Verdict B', 'Comparison'
  ]];
  keys.forEach(key => {
    const a = byRun[firstRun][key] || [];
    const b = byRun[secondRun][key] || [];
    const verdictA = String(a[12] || 'PENDING');
    const verdictB = String(b[12] || 'PENDING');
    let comparison = 'PENDING';
    if (verdictA !== 'PENDING' && verdictB !== 'PENDING') {
      comparison = verdictA === verdictB ? 'SAME_VERDICT' : `${verdictA}→${verdictB}`;
    }
    rows.push([firstRun, secondRun, a[2] || b[2] || '', a[3] || b[3] || '', verdictA, verdictB, comparison]);
  });

  const reportName = `BenchmarkComparison_${firstRun}_${secondRun}`.replace(/[^a-zA-Z0-9_-]/g, '_');
  if (ss.getSheetByName(reportName)) throw new Error('This comparison report already exists; existing report was not changed.');
  const report = ss.insertSheet(reportName);
  report.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  report.setFrozenRows(1);
  return `${reportName} created for ${firstRun} and ${secondRun}.`;
}
