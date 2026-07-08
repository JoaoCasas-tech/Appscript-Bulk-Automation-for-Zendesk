/*******************************************************
 * Zendesk Bulk Ticket Automation
 *
 * Expected source sheet columns:
 * customer_email | ticket_subject | ticket_comment | tags
 *
 * This script:
 * 1. Reads the Input CSV Google Sheet tab
 * 2. Creates Zendesk tickets in batches of up to 100
 * 3. Sets the ticket comment as public
 * 4. Adds tags from the CSV/sheet
 * 5. Solves the ticket immediately
 * 6. Appends all audit logs into ONE fixed sheet
 *
 * Authentication:
 * Zendesk API Token using Basic Auth:
 * email/token:api_token
 *******************************************************/

const CONFIG = {
  // IMPORTANT:
  // Your audit log showed the source sheet as "Input_CSV".
  // If your tab is named "Input CSV" instead, change this value.
  SOURCE_SHEET_NAME: "Input_CSV",

  // One single audit sheet for every run and every batch.
  LOG_SHEET_NAME: "Zendesk Bulk Audit Log",

  // Zendesk bulk ticket creation supports up to 100 tickets per request.
  BATCH_SIZE: 100,

  // Keep true for safe testing.
  // Set Script Property DRY_RUN=false when ready to send real Zendesk tickets.
  DRY_RUN_DEFAULT: true,

  // Bulk creation returns a job, so polling improves audit accuracy.
  POLL_JOB_STATUS: true,
  POLL_INTERVAL_MS: 5000,
  POLL_MAX_ATTEMPTS: 18,

  // Small pause between batches to be safer with low rate limits.
  SLEEP_MS_BETWEEN_BATCHES: 2000,

  MAX_RETRIES: 5
};

const REQUIRED_HEADERS = [
  "customer_email",
  "ticket_subject",
  "ticket_comment",
  "tags"
];

/**
 * Main function to run.
 */
function processZendeskBulkTickets() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error("Another Zendesk bulk process is already running. Try again later.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const zendesk = getZendeskConfig_();
    const sourceSheet = getSourceSheet_(ss);
    const sourceSheetName = sourceSheet.getName();

    const runId = Utilities.getUuid();
    const startedAt = new Date();

    const logSheet = getOrCreateAuditLogSheet_(ss);
    const records = readSourceRecords_(sourceSheet);

    const validRecords = [];
    const immediateLogs = [];

    records.forEach(record => {
      const validation = validateRecord_(record);
      const externalId = buildExternalId_(runId, record.sourceRowNumber);

      if (!validation.ok) {
        immediateLogs.push(makeLogRow_({
          runId,
          sourceSheetName,
          sourceRowNumber: record.sourceRowNumber,
          batchNumber: "",
          batchSize: "",
          payloadIndex: "",
          customerEmail: record.customer_email,
          ticketSubject: record.ticket_subject,
          zendeskHttpStatus: "",
          jobId: "",
          jobStatus: "not_sent_validation_failed",
          ticketId: "",
          externalId,
          success: false,
          message: validation.errors.join("; "),
          rawResponse: ""
        }));
        return;
      }

      validRecords.push(record);
    });

    appendLogRows_(logSheet, immediateLogs);

    const chunks = chunkArray_(validRecords, CONFIG.BATCH_SIZE);

    chunks.forEach((chunk, chunkIndex) => {
      const batchNumber = chunkIndex + 1;
      const batchSize = chunk.length;

      const tickets = chunk.map(record => {
        return buildZendeskTicketPayload_(record, runId);
      });

      if (zendesk.dryRun) {
        const dryRunLogs = chunk.map((record, index) => {
          const externalId = buildExternalId_(runId, record.sourceRowNumber);

          return makeLogRow_({
            runId,
            sourceSheetName,
            sourceRowNumber: record.sourceRowNumber,
            batchNumber,
            batchSize,
            payloadIndex: index,
            customerEmail: record.customer_email,
            ticketSubject: record.ticket_subject,
            zendeskHttpStatus: "DRY_RUN",
            jobId: "",
            jobStatus: "dry_run_not_sent",
            ticketId: "",
            externalId,
            success: "dry_run",
            message: "Payload built successfully. No Zendesk request was sent.",
            rawResponse: JSON.stringify(tickets[index])
          });
        });

        appendLogRows_(logSheet, dryRunLogs);
        SpreadsheetApp.flush();
        return;
      }

      let postResult;

      try {
        postResult = zendeskRequest_({
          method: "post",
          url: `${zendesk.baseUrl}/tickets/create_many.json`,
          email: zendesk.email,
          apiToken: zendesk.apiToken,
          payload: { tickets }
        });
      } catch (error) {
        const errorLogs = chunk.map((record, index) => {
          const externalId = buildExternalId_(runId, record.sourceRowNumber);

          return makeLogRow_({
            runId,
            sourceSheetName,
            sourceRowNumber: record.sourceRowNumber,
            batchNumber,
            batchSize,
            payloadIndex: index,
            customerEmail: record.customer_email,
            ticketSubject: record.ticket_subject,
            zendeskHttpStatus: "SCRIPT_EXCEPTION",
            jobId: "",
            jobStatus: "request_failed",
            ticketId: "",
            externalId,
            success: false,
            message: error.message,
            rawResponse: ""
          });
        });

        appendLogRows_(logSheet, errorLogs);
        SpreadsheetApp.flush();
        return;
      }

      const httpStatus = postResult.statusCode;
      const responseJson = postResult.bodyJson || {};
      const initialJob = responseJson.job_status || {};
      const jobId = initialJob.id || "";

      let finalJob = initialJob;
      let finalJobRawResponse = postResult.bodyText;

      if (isSuccessfulHttp_(httpStatus) && CONFIG.POLL_JOB_STATUS && jobId) {
        const pollResult = pollJobStatus_(zendesk, jobId);
        finalJob = pollResult.job || initialJob;
        finalJobRawResponse = pollResult.rawResponse || finalJobRawResponse;
      }

      const resultLogs = buildChunkLogRows_({
        runId,
        sourceSheetName,
        chunk,
        batchNumber,
        batchSize,
        httpStatus,
        jobId,
        finalJob,
        rawResponse: finalJobRawResponse
      });

      appendLogRows_(logSheet, resultLogs);
      SpreadsheetApp.flush();

      if (chunkIndex < chunks.length - 1) {
        Utilities.sleep(CONFIG.SLEEP_MS_BETWEEN_BATCHES);
      }
    });

    Logger.log(`Zendesk bulk run completed. Run ID: ${runId}. Started at: ${startedAt}`);

  } finally {
    lock.releaseLock();
  }
}

/***********************
 * Data reading
 ***********************/

function getSourceSheet_(ss) {
  if (CONFIG.SOURCE_SHEET_NAME) {
    const configuredSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME);

    if (!configuredSheet) {
      throw new Error(`Source sheet not found: ${CONFIG.SOURCE_SHEET_NAME}`);
    }

    if (configuredSheet.getName() === CONFIG.LOG_SHEET_NAME) {
      throw new Error("The configured source sheet cannot be the audit log sheet.");
    }

    return configuredSheet;
  }

  const activeSheet = ss.getActiveSheet();

  if (activeSheet && activeSheet.getName() !== CONFIG.LOG_SHEET_NAME) {
    return activeSheet;
  }

  const source = ss.getSheets().find(sheet => {
    return sheet.getName() !== CONFIG.LOG_SHEET_NAME;
  });

  if (!source) {
    throw new Error("No source sheet found. Please create/import a sheet with the required CSV columns.");
  }

  return source;
}

function readSourceRecords_(sheet) {
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error("Source sheet must have a header row and at least one data row.");
  }

  const headers = values[0].map(normalizeHeader_);
  const headerIndex = {};

  headers.forEach((header, index) => {
    headerIndex[header] = index;
  });

  const missingHeaders = REQUIRED_HEADERS.filter(header => !(header in headerIndex));

  if (missingHeaders.length > 0) {
    throw new Error(`Missing required columns: ${missingHeaders.join(", ")}`);
  }

  const records = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    if (row.every(cell => String(cell).trim() === "")) {
      continue;
    }

    records.push({
      sourceRowNumber: i + 1,
      customer_email: cleanCell_(row[headerIndex.customer_email]),
      ticket_subject: cleanCell_(row[headerIndex.ticket_subject]),
      ticket_comment: cleanCell_(row[headerIndex.ticket_comment]),
      tags: cleanCell_(row[headerIndex.tags])
    });
  }

  return records;
}

function validateRecord_(record) {
  const errors = [];

  if (!record.customer_email || !isValidEmail_(record.customer_email)) {
    errors.push("Invalid or missing customer_email");
  }

  if (!record.ticket_subject) {
    errors.push("Missing ticket_subject");
  }

  if (!record.ticket_comment) {
    errors.push("Missing ticket_comment");
  }

  return {
    ok: errors.length === 0,
    errors
  };
}

/***********************
 * Zendesk payload
 ***********************/

function buildZendeskTicketPayload_(record, runId) {
  return {
    subject: record.ticket_subject,
    requester: {
      name: deriveRequesterName_(record.customer_email),
      email: record.customer_email
    },
    comment: {
      body: formatCommentBody_(record.ticket_comment),
      public: true
    },
    tags: parseTags_(record.tags),
    status: "solved",
    external_id: buildExternalId_(runId, record.sourceRowNumber)
  };
}

function buildExternalId_(runId, sourceRowNumber) {
  return `zendesk_bulk_${runId}_row_${sourceRowNumber}`;
}

function deriveRequesterName_(email) {
  const localPart = String(email || "").split("@")[0] || "Customer";

  const cleanedName = localPart
    .replace(/[._+-]+/g, " ")
    .replace(/\d+/g, " ")
    .trim();

  if (!cleanedName) {
    return "Customer";
  }

  return toTitleCase_(cleanedName);
}

function toTitleCase_(value) {
  return String(value)
    .toLowerCase()
    .replace(/\b\w/g, char => char.toUpperCase());
}

function formatCommentBody_(comment) {
  return String(comment || "")
    .replace(/\\r\\n/g, "\n")
    .replace(/\\n/g, "\n")
    .replace(/\\r/g, "\n")
    .trim();
}

function parseTags_(tagsCell) {
  if (!tagsCell) {
    return [];
  }

  return String(tagsCell)
    .split(/[;,|]/)
    .map(tag => tag.trim())
    .filter(Boolean)
    .map(tag => normalizeZendeskTag_(tag));
}

function normalizeZendeskTag_(tag) {
  return String(tag)
    .toLowerCase()
    .trim()
    .replace(/\s+/g, "_")
    .replace(/[^a-z0-9_-]/g, "");
}

/***********************
 * Zendesk API
 ***********************/

function getZendeskConfig_() {
  const props = PropertiesService.getScriptProperties();

  const subdomain = props.getProperty("ZENDESK_SUBDOMAIN");
  const email = props.getProperty("ZENDESK_EMAIL");
  const apiToken = props.getProperty("ZENDESK_API_TOKEN");
  const dryRunProperty = props.getProperty("DRY_RUN");

  const dryRun = dryRunProperty === null
    ? CONFIG.DRY_RUN_DEFAULT
    : String(dryRunProperty).toLowerCase() === "true";

  if (!subdomain) {
    throw new Error("Missing Script Property: ZENDESK_SUBDOMAIN");
  }

  if (!dryRun && !email) {
    throw new Error("Missing Script Property: ZENDESK_EMAIL");
  }

  if (!dryRun && !apiToken) {
    throw new Error("Missing Script Property: ZENDESK_API_TOKEN");
  }

  return {
    baseUrl: `https://${subdomain}.zendesk.com/api/v2`,
    email,
    apiToken,
    dryRun
  };
}

function zendeskRequest_({ method, url, email, apiToken, payload }) {
  let attempt = 0;

  while (attempt <= CONFIG.MAX_RETRIES) {
    const authString = `${email}/token:${apiToken}`;
    const encodedAuth = Utilities.base64Encode(authString);

    const options = {
      method,
      muteHttpExceptions: true,
      contentType: "application/json",
      headers: {
        Authorization: `Basic ${encodedAuth}`,
        Accept: "application/json"
      }
    };

    if (payload !== undefined && payload !== null) {
      options.payload = JSON.stringify(payload);
    }

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const bodyText = response.getContentText();
    const headers = response.getAllHeaders();

    if (shouldRetry_(statusCode) && attempt < CONFIG.MAX_RETRIES) {
      const waitSeconds = getRetryWaitSeconds_(headers, attempt);
      Logger.log(`Zendesk API retry. Status: ${statusCode}. Waiting ${waitSeconds}s.`);
      Utilities.sleep(waitSeconds * 1000);
      attempt++;
      continue;
    }

    return {
      statusCode,
      headers,
      bodyText,
      bodyJson: parseJsonSafely_(bodyText)
    };
  }

  throw new Error("Zendesk request failed after max retries.");
}

function pollJobStatus_(zendesk, jobId) {
  let lastResult = null;

  for (let attempt = 0; attempt < CONFIG.POLL_MAX_ATTEMPTS; attempt++) {
    const result = zendeskRequest_({
      method: "get",
      url: `${zendesk.baseUrl}/job_statuses/${encodeURIComponent(jobId)}.json`,
      email: zendesk.email,
      apiToken: zendesk.apiToken
    });

    const job = result.bodyJson && result.bodyJson.job_status
      ? result.bodyJson.job_status
      : result.bodyJson;

    lastResult = {
      job,
      rawResponse: result.bodyText
    };

    if (job && ["completed", "failed"].indexOf(job.status) !== -1) {
      return lastResult;
    }

    Utilities.sleep(CONFIG.POLL_INTERVAL_MS);
  }

  return {
    job: {
      id: jobId,
      status: "timeout_waiting_for_job_completion",
      message: "Job did not complete within the configured polling window."
    },
    rawResponse: lastResult ? lastResult.rawResponse : ""
  };
}

function shouldRetry_(statusCode) {
  return statusCode === 429 || statusCode === 503 || statusCode >= 500;
}

function getRetryWaitSeconds_(headers, attempt) {
  const retryAfter = getHeaderCaseInsensitive_(headers, "Retry-After");

  if (retryAfter) {
    const retryAfterNumber = Number(retryAfter);

    if (!isNaN(retryAfterNumber)) {
      return retryAfterNumber + 1;
    }
  }

  const rateLimitReset = getHeaderCaseInsensitive_(headers, "ratelimit-reset");

  if (rateLimitReset) {
    const resetNumber = Number(rateLimitReset);

    if (!isNaN(resetNumber)) {
      const nowSeconds = Math.floor(Date.now() / 1000);

      if (resetNumber > nowSeconds) {
        return Math.max(resetNumber - nowSeconds + 1, 1);
      }

      return Math.max(resetNumber + 1, 1);
    }
  }

  return Math.min(Math.pow(2, attempt + 1), 60);
}

/***********************
 * Audit logging
 ***********************/

function getOrCreateAuditLogSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "run_id",
      "source_sheet_name",
      "source_row_number",
      "batch_number",
      "batch_size",
      "payload_index",
      "customer_email",
      "ticket_subject",
      "zendesk_http_status",
      "job_id",
      "job_status",
      "ticket_id",
      "external_id",
      "success",
      "message",
      "processed_at",
      "raw_response"
    ]);

    sheet.setFrozenRows(1);
  }

  return sheet;
}

function buildChunkLogRows_({
  runId,
  sourceSheetName,
  chunk,
  batchNumber,
  batchSize,
  httpStatus,
  jobId,
  finalJob,
  rawResponse
}) {
  const results = Array.isArray(finalJob.results) ? finalJob.results : [];

  return chunk.map((record, index) => {
    const result = findJobResultForIndex_(results, index);
    const externalId = buildExternalId_(runId, record.sourceRowNumber);

    return makeLogRow_({
      runId,
      sourceSheetName,
      sourceRowNumber: record.sourceRowNumber,
      batchNumber,
      batchSize,
      payloadIndex: index,
      customerEmail: record.customer_email,
      ticketSubject: record.ticket_subject,
      zendeskHttpStatus: httpStatus,
      jobId,
      jobStatus: finalJob.status || "",
      ticketId: extractTicketId_(result),
      externalId,
      success: calculateRowSuccess_(httpStatus, finalJob, result),
      message: extractResultMessage_(finalJob, result),
      rawResponse
    });
  });
}

function findJobResultForIndex_(results, index) {
  if (!results || results.length === 0) {
    return null;
  }

  const byIndex = results.find(result => Number(result.index) === index);

  if (byIndex) {
    return byIndex;
  }

  return results[index] || null;
}

function extractTicketId_(result) {
  if (!result) {
    return "";
  }

  return result.id || result.ticket_id || "";
}

function extractResultMessage_(finalJob, result) {
  if (!result) {
    return finalJob.message || "";
  }

  if (result.error) {
    return `${result.error}: ${result.details || result.message || ""}`;
  }

  if (result.errors) {
    return JSON.stringify(result.errors);
  }

  if (result.details && !result.id && !result.ticket_id) {
    return JSON.stringify(result.details);
  }

  return result.message || result.status || finalJob.message || "";
}

function calculateRowSuccess_(httpStatus, finalJob, result) {
  if (!isSuccessfulHttp_(httpStatus)) {
    return false;
  }

  if (finalJob.status === "timeout_waiting_for_job_completion") {
    return "pending_job_status_timeout";
  }

  if (finalJob.status === "failed") {
    return false;
  }

  if (result && (result.error || result.errors || result.success === false)) {
    return false;
  }

  if (result && (result.id || result.ticket_id)) {
    return true;
  }

  if (finalJob.status === "completed" && !result) {
    return "completed_no_row_result";
  }

  return "accepted_pending_job_completion";
}

function makeLogRow_(args) {
  return [
    args.runId || "",
    args.sourceSheetName || "",
    args.sourceRowNumber || "",
    args.batchNumber || "",
    args.batchSize || "",
    args.payloadIndex === 0 ? 0 : args.payloadIndex || "",
    args.customerEmail || "",
    args.ticketSubject || "",
    args.zendeskHttpStatus || "",
    args.jobId || "",
    args.jobStatus || "",
    args.ticketId || "",
    args.externalId || "",
    args.success,
    args.message || "",
    new Date(),
    truncateForSheet_(args.rawResponse || "")
  ];
}

function appendLogRows_(sheet, rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}

/***********************
 * Helpers
 ***********************/

function normalizeHeader_(value) {
  return String(value).trim().toLowerCase();
}

function cleanCell_(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).trim());
}

function isSuccessfulHttp_(statusCode) {
  return Number(statusCode) >= 200 && Number(statusCode) < 300;
}

function parseJsonSafely_(text) {
  try {
    return JSON.parse(text);
  } catch (error) {
    return {};
  }
}

function getHeaderCaseInsensitive_(headers, targetName) {
  const target = targetName.toLowerCase();

  for (const key in headers) {
    if (String(key).toLowerCase() === target) {
      return headers[key];
    }
  }

  return null;
}

function chunkArray_(array, size) {
  const chunks = [];

  for (let i = 0; i < array.length; i += size) {
    chunks.push(array.slice(i, i + size));
  }

  return chunks;
}

function truncateForSheet_(value) {
  const text = typeof value === "string" ? value : JSON.stringify(value);
  const limit = 45000;

  if (text.length <= limit) {
    return text;
  }

  return text.substring(0, limit) + "...[truncated]";
}