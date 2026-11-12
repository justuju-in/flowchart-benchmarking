/** Fixed benchmark-set storage and Set B selection. */

var BENCHMARK_SET_HEADERS = [
  "Set ID", "Name", "Description", "Source Spreadsheet ID", "Source Sheet",
  "Status", "Item Count", "Selection SHA256", "Created At"
];

var BENCHMARK_SET_ITEM_HEADERS = [
  "Set ID", "Item ID", "Source Row", "Flowchart URL", "Flowchart File ID",
  "Source Reference", "Problem Code", "Issue Reported", "RCA", "Debugging Status",
  "Debugging Comments", "Included", "Created At"
];

var BENCHMARK_SET_B = {
  id: "SET_B_FLAGGED_VISHAL_V1",
  name: "Set B - Vishal Flagged Submissions",
  description: "Fixed regression set of unique flowcharts previously flagged by Vishal.",
  sourceSpreadsheetId: "1VHFvbcIukZhSB9wmTMSFXuSpRu49ySVq63_SpUz_KUs",
  sourceSheet: "flagged_runs_vishal"
};

function benchmarkNormalizeProblemCode_(value) {
  var text = String(value || "").trim().toUpperCase();
  if (!text) return "";
  if (/^FCP\d+$/.test(text)) return "FCP" + text.slice(3).padStart(3, "0");
  return "";
}

function benchmarkDriveFileId_(url) {
  var match = String(url || "").match(/[-\w]{25,}/);
  return match ? match[0] : "";
}

function benchmarkAppendUniqueText_(current, next) {
  var values = String(current || "").split("\n---\n").filter(Boolean);
  var value = String(next || "").trim();
  if (value && values.indexOf(value) < 0) values.push(value);
  return values.join("\n---\n");
}

function benchmarkBuildSetBItems_(values) {
  if (!values || values.length < 2) return [];
  var headers = values[0].map(function (value) { return String(value).trim(); });
  var required = ["flowchart", "f_id", "issue_reported", "rca", "debugging_status", "debugging_comments"];
  required.forEach(function (name) {
    if (headers.indexOf(name) < 0) throw new Error("Set B source is missing column: " + name);
  });
  var read = function (row, name) { return row[headers.indexOf(name)] || ""; };
  var byFile = {};
  values.slice(1).forEach(function (row, index) {
    var url = String(read(row, "flowchart") || "").trim();
    var fileId = benchmarkDriveFileId_(url);
    if (!fileId) return;
    if (!byFile[fileId]) {
      byFile[fileId] = {
        sourceRow: index + 2,
        flowchartUrl: url,
        flowchartFileId: fileId,
        sourceReference: String(read(row, "f_id") || "").trim(),
        problemCode: benchmarkNormalizeProblemCode_(read(row, "f_id")),
        issueReported: "",
        rca: "",
        debuggingStatus: "",
        debuggingComments: ""
      };
    }
    var item = byFile[fileId];
    item.issueReported = benchmarkAppendUniqueText_(item.issueReported, read(row, "issue_reported"));
    item.rca = benchmarkAppendUniqueText_(item.rca, read(row, "rca"));
    item.debuggingStatus = benchmarkAppendUniqueText_(item.debuggingStatus, read(row, "debugging_status"));
    item.debuggingComments = benchmarkAppendUniqueText_(item.debuggingComments, read(row, "debugging_comments"));
  });
  return Object.keys(byFile).sort().map(function (key) { return byFile[key]; });
}

function benchmarkHash_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text))
    .map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, "0"); })
    .join("");
}

/**
 * Creates the immutable first version of Set B from flagged_runs_vishal.
 * This only stores the selection; it does not call Gemini or DOMjudge.
 */
function createBenchmarkSetB() {
  var target = SpreadsheetApp.getActiveSpreadsheet();
  var sets = getOrCreateSheet("BenchmarkSets", BENCHMARK_SET_HEADERS);
  var itemsSheet = getOrCreateSheet("BenchmarkSetItems", BENCHMARK_SET_ITEM_HEADERS);
  var existing = sets.getDataRange().getDisplayValues().slice(1).some(function (row) {
    return row[0] === BENCHMARK_SET_B.id;
  });
  if (existing) throw new Error(BENCHMARK_SET_B.id + " already exists. Fixed benchmark sets cannot be overwritten.");

  var source = SpreadsheetApp.openById(BENCHMARK_SET_B.sourceSpreadsheetId)
    .getSheetByName(BENCHMARK_SET_B.sourceSheet);
  if (!source) throw new Error("Set B source sheet not found: " + BENCHMARK_SET_B.sourceSheet);
  var items = benchmarkBuildSetBItems_(source.getDataRange().getDisplayValues());
  if (!items.length) throw new Error("No valid flagged flowcharts found for Set B.");

  var createdAt = new Date();
  var rows = items.map(function (item) {
    return [
      BENCHMARK_SET_B.id,
      Utilities.getUuid(),
      item.sourceRow,
      item.flowchartUrl,
      item.flowchartFileId,
      item.sourceReference,
      item.problemCode,
      item.issueReported,
      item.rca,
      item.debuggingStatus,
      item.debuggingComments,
      true,
      createdAt
    ];
  });
  itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, rows.length, BENCHMARK_SET_ITEM_HEADERS.length).setValues(rows);
  var selection = items.map(function (item) {
    return [item.flowchartFileId, item.sourceReference, item.problemCode, item.issueReported, item.rca];
  });
  sets.appendRow([
    BENCHMARK_SET_B.id,
    BENCHMARK_SET_B.name,
    BENCHMARK_SET_B.description,
    BENCHMARK_SET_B.sourceSpreadsheetId,
    BENCHMARK_SET_B.sourceSheet,
    "FIXED",
    rows.length,
    benchmarkHash_(JSON.stringify(selection)),
    createdAt
  ]);
  SpreadsheetApp.flush();
  target.setActiveSheet(itemsSheet);
  return { setId: BENCHMARK_SET_B.id, itemCount: rows.length };
}
