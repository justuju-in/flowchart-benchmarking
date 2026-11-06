// config.js - Centralized configuration for flowchart-judge
// All configurable variables and sheet names

var CONFIG = {
  IsBatchMode: true,
  maxRowCount: 5,
  SubmissionIdColumnName: "Submission ID",
  SubmissionTimestampColumnName: "Submission Timestamp",
  SubmissionStatusColumnName: "Submission Status",
  StatusColumnName: "Status",
  CodeFileUrlColumnName: "Code File URL",
  ProblemNumberColumnName: "Problem Number",
  TimestampColumnName: "Timestamp",
  EmailAddressColumnName: "Email Address",
  UploadFlowchartColumnName: "Upload your Flowchart",
  VerdictColumnName: "Verdict",
  ModelUsedColumnName: "Model Used",
  PromptVersionColumnName: "Prompt Version",
  GenerationTimestampColumnName: "Generation Timestamp",
  InputTokensColumnName: "Input Tokens",
  OutputTokensColumnName: "Output Tokens",
  TotalTokensColumnName: "Total Tokens",
  ResponseTimeColumnName: "Response Time (ms)",
  SafetyRatingsColumnName: "Safety Ratings",
  FinishReasonColumnName: "Finish Reason",
  CitationMetadataColumnName: "Citation Metadata",
  ModelVersionColumnName: "Model Used",
  ResponseIdColumnName: "Response ID",
  ThoughtsTokenCountColumnName: "Thoughts Token Count",
  TextTokenCountColumnName: "Text Token Count",
  ImageTokenCountColumnName: "Image Token Count",
  ImageMimeTypeColumnName: "Image MIME Type",
  promptVersion: "Production - 1st Aug 2025",
  // Status constants
  NewEntryStatus: "NEW",
  GeminiDoneStatus: "GEMINI_DONE",
  JudgeSubmittedStatus: "JUDGE_SUBMITTED",
  VerdictReadyStatus: "VERDICT_READY",
  SubmissionStatus: "SUBMITTED",
  generatedCodesFolder: "GeneratedCodes",
  masterSheet: "Master",
  // Only submissions originating from this response tab are allowed to reach
  // the test DOMjudge workflow. This keeps copied historical Master rows safe.
  sourceResponseSheet: "Form Responses 4",
  logEventSheet: "event-log",
  metaSheet: "Meta",
  geminiApiKeyProperty: "GEMINI_API_KEY", // Fixed to match your script property name
  // DOMjudge connection details are deliberately read from Script Properties.
  // Never put test or production credentials in this repository.
  domjudgeApiUrl: "",
  domjudgeContestId: "",
  domjudgeTeamId: "",
  domjudgeUser: "",
  domjudgePass: "",
  // gemini-2.5-pro is retired for this project. Keep this default aligned
  // with the runtime configuration in geminiConfig.js.
  geminiModel: "gemini-3.1-pro-preview",
  geminiTemperature: 0,
  CannotProcessStatus: "CANNOT_PROCESS",
  // Use base endpoint for Gemini API with models path
  geminiEndpoint: "https://generativelanguage.googleapis.com/v1beta/models/",
  logLevelDebug: 1,
  logLevelError: 2,
  loggingLevel: 2,
  // Logging configuration
  logDateHeader: "DATE",
  logLevelHeader: "LOG_LEVEL", 
  logMessageHeader: "MESSAGE",
  logDebugString: "DEBUG",
  logErrorString: "ERROR",
  // Gemini response defaults
  defaultFinishReason: "STOP",
  domjudgeAdminUser: "",
  domjudgeAdminPass: "",
  // DomJudge submission configuration
  domjudgeLanguageId: "python3",
  domjudgeSolutionFilename: "solution.py",
  domjudgeZipFilename: "solution.zip",
  generatedCodeFileExtension: ".py",
  masterSheetColumns: [
    "Timestamp",
    "Email Address",
    "Problem Number",
    "Upload your Flowchart",
    "Status",
    "Code File URL",
    "Model Used",
    "Prompt Version",
    "Generation Timestamp",
    "Submission ID",
    "Submission Timestamp",
    "Submission Status",
    "Verdict",
    "Processing Duration",
    "Input Tokens",
    "Output Tokens",
    "Total Tokens",
    "Response Time (ms)",
    "Safety Ratings",
    "Finish Reason",
    "Citation Metadata",
    "Model Version",
    "Response ID",
    "Thoughts Token Count",
    "Text Token Count",
    "Image Token Count"
  ],
  metaSheetColumns: ["Problem Code", "Problem ID"],
  domjudgeSubmissionColumns: ["Submission ID", "Problem ID", "Verdict"],
};

function getConfig() {
  var config = Object.assign({}, CONFIG);
  var properties = PropertiesService.getScriptProperties().getProperties();
  var propertyMap = {
    DOMJUDGE_BASE_URL: "domjudgeApiUrl",
    DOMJUDGE_CONTEST_ID: "domjudgeContestId",
    DOMJUDGE_TEAM_ID: "domjudgeTeamId",
    DOMJUDGE_USERNAME: "domjudgeUser",
    DOMJUDGE_PASSWORD: "domjudgePass",
    DOMJUDGE_ADMIN_USERNAME: "domjudgeAdminUser",
    DOMJUDGE_ADMIN_PASSWORD: "domjudgeAdminPass",
    DOMJUDGE_LANGUAGE_ID: "domjudgeLanguageId"
  };

  Object.keys(propertyMap).forEach(function(propertyName) {
    if (properties[propertyName] !== undefined && properties[propertyName] !== "") {
      config[propertyMap[propertyName]] = properties[propertyName];
    }
  });

  config.domjudgeApiUrl = String(config.domjudgeApiUrl || "").replace(/\/+$/, "");
  // A separate admin account is optional for the copied test environment.
  config.domjudgeAdminUser = config.domjudgeAdminUser || config.domjudgeUser;
  config.domjudgeAdminPass = config.domjudgeAdminPass || config.domjudgePass;
  return config;
}
