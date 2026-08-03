// Data service for Invoice Requests.

var INVOICE_REQUESTS_SPREADSHEET_ID =
  "1-5b98hkm-wsrTlyt-1j2PJcTvu1VNyAh0EjRhT3ZBg0";
var INVOICE_REQUESTS_SHEET_NAME = "Requests";
var INVOICE_REQUESTS_INFORMATION_SHEET = "Information";
var INVOICE_REQUESTS_LISTS_SHEET = "Lists";
var INVOICE_REQUESTS_FIRST_COLUMN = 2; // B
var INVOICE_REQUESTS_COLUMN_COUNT = 18; // B:S
var INVOICE_REQUESTS_LAST_COLUMN = 19; // S
var INVOICE_REQUESTS_STATUS_FIRST_OFFSET = 4; // F
var INVOICE_REQUESTS_STATUS_COUNT = 8; // F:M
var INVOICE_REQUESTS_PROJECT_OFFSET = 0; // B
var INVOICE_REQUESTS_DETAILS_OFFSET = 1; // C
var INVOICE_REQUESTS_COMMENT_OFFSET = 2; // D
var INVOICE_REQUESTS_AUTHOR_OFFSET = 3; // E
var INVOICE_REQUESTS_RATE_FILE_OFFSET = 12; // N
var INVOICE_REQUESTS_CLIENT_FOLDER_OFFSET = 13; // O
var INVOICE_REQUESTS_CREATED_BY_OFFSET = 14; // P
var INVOICE_REQUESTS_CREATED_AT_OFFSET = 15; // Q
var INVOICE_REQUESTS_EDITED_BY_OFFSET = 16; // R
var INVOICE_REQUESTS_EDITED_AT_OFFSET = 17; // S
var INVOICE_REQUESTS_TRAILING_COLUMN_COUNT = 6; // N:S
// Stored in the status cell. No background color is used for N/A.
var INVOICE_REQUESTS_NOT_APPLICABLE = "✕";
var INVOICE_REQUESTS_LIST_CACHE_KEY = "invoiceRequestsList";
var INVOICE_REQUESTS_LEGACY_NOT_APPLICABLE = "⊟";

function isInvoiceRequestStatusColumn_(columnOffset) {
  return (
    columnOffset >= INVOICE_REQUESTS_STATUS_FIRST_OFFSET &&
    columnOffset <
      INVOICE_REQUESTS_STATUS_FIRST_OFFSET +
        INVOICE_REQUESTS_STATUS_COUNT
  );
}

function isInvoiceRequestContentColumn_(columnOffset) {
  return (
    columnOffset === INVOICE_REQUESTS_PROJECT_OFFSET ||
    columnOffset === INVOICE_REQUESTS_DETAILS_OFFSET ||
    columnOffset === INVOICE_REQUESTS_COMMENT_OFFSET
  );
}

function isInvoiceRequestClientEditableColumn_(columnOffset) {
  return (
    isInvoiceRequestContentColumn_(columnOffset) ||
    isInvoiceRequestStatusColumn_(columnOffset)
  );
}

function formatInvoiceRequestTimestamp_(date) {
  return Utilities.formatDate(
    date || new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );
}

// App timestamps are always day/month/year. Do not use Date.parse on
// slash-dates: in JS they are treated as MM/DD and break sorting.
function parseInvoiceRequestDdMmTimestamp_(value) {
  var text = String(value || "").trim();
  if (!text) return 0;
  var match = text.match(
    /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})(?:[,\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/
  );
  if (!match) return 0;
  var day = Number(match[1]);
  var month = Number(match[2]);
  var year = Number(match[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return 0;
  var parsed = new Date(
    year,
    month - 1,
    day,
    Number(match[4] || 0),
    Number(match[5] || 0),
    Number(match[6] || 0)
  );
  return isNaN(parsed.getTime()) ? 0 : parsed.getTime();
}

function parseInvoiceRequestTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getTime();
  }
  return parseInvoiceRequestDdMmTimestamp_(value);
}

function isInvoiceRequestTimestampColumn_(columnOffset) {
  return (
    columnOffset === INVOICE_REQUESTS_CREATED_AT_OFFSET ||
    columnOffset === INVOICE_REQUESTS_EDITED_AT_OFFSET
  );
}

function resolveInvoiceRequestActivityMs_(value, displayValue) {
  // Trust the real cell Date first. Sheets display strings follow the
  // spreadsheet locale (often MM/DD), so parsing them as dd/MM flips
  // days and months (e.g. 3 Aug -> 08/03/2026) and breaks sort order.
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getTime();
  }
  if (typeof value === "number" && Number.isFinite(value) && value > 0) {
    if (value < 1000000) {
      return new Date(Math.round((value - 25569) * 86400000)).getTime();
    }
    return value;
  }
  // Legacy app-written text timestamps used dd/MM/yyyy HH:mm.
  var fromValue = parseInvoiceRequestDdMmTimestamp_(value);
  if (fromValue) return fromValue;
  return parseInvoiceRequestDdMmTimestamp_(displayValue);
}

function serializeInvoiceRequestTimestampCell_(value, displayValue) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return formatInvoiceRequestTimestamp_(value);
  }
  var fromValue = parseInvoiceRequestDdMmTimestamp_(value);
  if (fromValue) return formatInvoiceRequestTimestamp_(new Date(fromValue));
  var fromDisplay = parseInvoiceRequestDdMmTimestamp_(displayValue);
  if (fromDisplay) {
    return formatInvoiceRequestTimestamp_(new Date(fromDisplay));
  }
  return String(displayValue || value || "").trim();
}

function getInvoiceRequestActivityAt_(createdAt, editedAt) {
  var editedMs = parseInvoiceRequestTimestamp_(editedAt);
  if (editedMs) return editedMs;
  return parseInvoiceRequestTimestamp_(createdAt);
}

function getInvoiceRequestActivityAtFromCells_(
  createdValue,
  createdDisplay,
  editedValue,
  editedDisplay
) {
  var editedMs = resolveInvoiceRequestActivityMs_(
    editedValue,
    editedDisplay
  );
  if (editedMs) return editedMs;
  return resolveInvoiceRequestActivityMs_(createdValue, createdDisplay);
}

function sheetColumnForInvoiceRequestOffset_(columnOffset) {
  return INVOICE_REQUESTS_FIRST_COLUMN + columnOffset;
}

function resetInvoiceRequestStatuses_(sheet, sheetRow) {
  var statusRange = sheet.getRange(
    sheetRow,
    sheetColumnForInvoiceRequestOffset_(
      INVOICE_REQUESTS_STATUS_FIRST_OFFSET
    ),
    1,
    INVOICE_REQUESTS_STATUS_COUNT
  );
  var uncheckedRow = [];
  for (var i = 0; i < INVOICE_REQUESTS_STATUS_COUNT; i++) {
    uncheckedRow.push(false);
  }
  statusRange.insertCheckboxes();
  statusRange.setValues([uncheckedRow]);
  statusRange.setBackground(null);
}

// Projects list for dropdowns / validation. Does not read rate or folder links.
function getInvoiceRequestProjects_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    INVOICE_REQUESTS_INFORMATION_SHEET
  );
  if (!sheet) throw new Error('Sheet "Information" was not found.');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
  var seen = {};
  var projects = [];
  for (var i = 0; i < values.length; i++) {
    var project = String(values[i][0] || "").trim();
    if (!project || seen[project]) continue;
    seen[project] = true;
    projects.push(project);
  }
  projects.sort(function (left, right) {
    return left.localeCompare(right, undefined, { sensitivity: "base" });
  });
  return projects;
}

// Rate file + client folder from Information. Used only when creating a request.
function getInvoiceRequestInformationLookup_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    INVOICE_REQUESTS_INFORMATION_SHEET
  );
  if (!sheet) throw new Error('Sheet "Information" was not found.');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { projects: [], rateByProject: {}, folderByProject: {} };
  }

  // B:E in one read: project in B, client-folder link in C, rate-file link in E.
  var range = sheet.getRange(2, 2, lastRow - 1, 4);
  var displayValues = range.getDisplayValues();
  var richTextValues = range.getRichTextValues();
  var seen = {};
  var projects = [];
  var rateByProject = {};
  var folderByProject = {};
  for (var i = 0; i < displayValues.length; i++) {
    var project = String(displayValues[i][0] || "").trim();
    if (!project) continue;
    if (!seen[project]) {
      seen[project] = true;
      projects.push(project);
    }
    if (!folderByProject[project]) {
      var folderRichText = richTextValues[i][1];
      var folderLink = folderRichText
        ? folderRichText.getLinkUrl() || ""
        : "";
      var folderValue = String(displayValues[i][1] || "").trim();
      if (!folderValue && folderLink) folderValue = folderLink;
      folderByProject[project] = {
        value: folderValue,
        link: folderLink,
      };
    }
    if (rateByProject[project]) continue;
    var richText = richTextValues[i][3];
    var link = richText ? richText.getLinkUrl() || "" : "";
    var value = String(displayValues[i][3] || "").trim();
    if (!value && link) value = link;
    rateByProject[project] = { value: value, link: link };
  }
  projects.sort(function (left, right) {
    return left.localeCompare(right, undefined, { sensitivity: "base" });
  });
  return {
    projects: projects,
    rateByProject: rateByProject,
    folderByProject: folderByProject,
  };
}

function getInvoiceRequestListsLookup_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(INVOICE_REQUESTS_LISTS_SHEET);
  if (!sheet) throw new Error('Sheet "Lists" was not found.');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { authorByEmail: {}, notificationEmails: [] };
  }

  // A:D — Full Name, Email, unused, notification recipients.
  var values = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  var authorByEmail = {};
  var notificationEmails = [];
  var seenNotify = {};
  for (var i = 0; i < values.length; i++) {
    var fullName = String(values[i][0] || "").trim();
    var listEmail = String(values[i][1] || "").trim().toLowerCase();
    if (listEmail && !authorByEmail[listEmail]) {
      authorByEmail[listEmail] = fullName || listEmail;
    }
    var notifyEmail = String(values[i][3] || "").trim().toLowerCase();
    if (!notifyEmail || seenNotify[notifyEmail]) continue;
    seenNotify[notifyEmail] = true;
    notificationEmails.push(notifyEmail);
  }
  return {
    authorByEmail: authorByEmail,
    notificationEmails: notificationEmails,
  };
}

function resolveInvoiceRequestAuthor_(spreadsheet, email, listsLookup) {
  var normalizedEmail = String(email || "").trim().toLowerCase();
  if (!normalizedEmail) return "";
  var lookup = listsLookup || getInvoiceRequestListsLookup_(spreadsheet);
  return lookup.authorByEmail[normalizedEmail] || normalizedEmail;
}

function getInvoiceRequestsPageUrl_() {
  return ScriptApp.getService().getUrl() + "?page=InvoiceRequests";
}

function escapeInvoiceRequestHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function sendInvoiceRequestNotification_(
  spreadsheet,
  kind,
  project,
  author,
  listsLookup
) {
  try {
    var lookup = listsLookup || getInvoiceRequestListsLookup_(spreadsheet);
    var recipients = lookup.notificationEmails || [];
    if (!recipients.length) return;

    var isEdited = kind === "edited";
    var subject = isEdited
      ? "Invoice request edited"
      : "Invoice request created";
    var paragraph = isEdited
      ? "Заявку на інвойс було змінено."
      : "Було створено заявку на інвойс.";
    var pageUrl = getInvoiceRequestsPageUrl_();
    var safeProject = escapeInvoiceRequestHtml_(project);
    var safeAuthor = escapeInvoiceRequestHtml_(author);
    var safeUrl = escapeInvoiceRequestHtml_(pageUrl);

    var plainBody =
      "Привіт.\n\n" +
      paragraph +
      "\n\n" +
      "Проект: " +
      String(project || "") +
      "\n" +
      "Автор: " +
      String(author || "") +
      "\n\n" +
      "Деталі за посиланням: " +
      pageUrl;

    var htmlBody =
      "<p>Привіт.</p>" +
      "<p>" +
      escapeInvoiceRequestHtml_(paragraph) +
      "</p>" +
      "<p><b>Проект:</b> " +
      safeProject +
      "<br>" +
      "<b>Автор:</b> " +
      safeAuthor +
      "</p>" +
      "<p>Деталі за <a href=\"" +
      safeUrl +
      "\">посиланням</a></p>";

    MailApp.sendEmail({
      to: recipients[0],
      bcc: recipients.slice(1).join(","),
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
    });
  } catch (error) {
    console.error("Invoice request notification failed:", error);
  }
}

/**
 * Public entry for deferred notifications after create/edit succeeded.
 * Called from the browser after the UI already received the save result.
 */
function sendInvoiceRequestNotifications(notifications) {
  assertInvoiceRequestsAccess_();
  if (!Array.isArray(notifications) || notifications.length === 0) {
    return { success: true, sent: 0 };
  }

  var spreadsheet = SpreadsheetApp.openById(
    INVOICE_REQUESTS_SPREADSHEET_ID
  );
  var listsLookup = getInvoiceRequestListsLookup_(spreadsheet);
  var sent = 0;
  for (var i = 0; i < notifications.length; i++) {
    var item = notifications[i] || {};
    var kind = item.kind === "edited" ? "edited" : "created";
    var project = String(item.project || "").trim();
    var author = String(item.author || "").trim();
    if (!project) continue;
    sendInvoiceRequestNotification_(
      spreadsheet,
      kind,
      project,
      author,
      listsLookup
    );
    sent++;
  }
  return { success: true, sent: sent };
}

function resolveInvoiceRequestRateFile_(
  spreadsheet,
  project,
  informationLookup
) {
  var normalizedProject = String(project || "").trim();
  if (!normalizedProject) return { value: "", link: "" };
  var lookup =
    informationLookup || getInvoiceRequestInformationLookup_(spreadsheet);
  return lookup.rateByProject[normalizedProject] || { value: "", link: "" };
}

function resolveInvoiceRequestClientFolder_(
  spreadsheet,
  project,
  informationLookup
) {
  var normalizedProject = String(project || "").trim();
  if (!normalizedProject) return { value: "", link: "" };
  var lookup =
    informationLookup || getInvoiceRequestInformationLookup_(spreadsheet);
  return lookup.folderByProject[normalizedProject] || { value: "", link: "" };
}

function writeInvoiceRequestLinkedValue_(range, value, link) {
  // Plain URL text in the cell — same approach as Clients Information.
  var url = String(link || value || "").trim();
  range.clearDataValidations();
  range.setBackground(null);
  if (url) {
    range.setValue(url);
    return;
  }
  range.clearContent();
}

function assertInvoiceRequestProject_(project, projects) {
  if (!projects || projects.indexOf(project) === -1) {
    return {
      success: false,
      validation: true,
      message: "Select a project from the list.",
    };
  }
  return null;
}

function assertInvoiceRequestsAccess_() {
  if (!canAccessInvoiceRequests(getCurrentUserEmail())) {
    throw new Error("No permission to access Invoice Requests.");
  }
}

function getInvoiceRequestAccessMode_() {
  return isInvoiceRequestsFullAccess() ? "full" : "limited";
}

function invoiceRequestRowOwnedByEmail_(rowValues, email) {
  var start = INVOICE_REQUESTS_FIRST_COLUMN - 1;
  var createdBy = String(
    rowValues[start + INVOICE_REQUESTS_CREATED_BY_OFFSET] || ""
  )
    .trim()
    .toLowerCase();
  var editedBy = String(
    rowValues[start + INVOICE_REQUESTS_EDITED_BY_OFFSET] || ""
  )
    .trim()
    .toLowerCase();
  var normalized = String(email || "")
    .trim()
    .toLowerCase();
  return createdBy === normalized || editedBy === normalized;
}

function serializeInvoiceRequestValue_(value, displayValue) {
  if (value instanceof Date) return displayValue || "";
  if (value === null || value === undefined) return "";
  return value;
}

function comparableInvoiceRequestValue_(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd'T'HH:mm:ss"
    );
  }
  if (typeof value === "boolean" || typeof value === "number") return value;
  return String(value);
}

function getInvoiceRequestCheckboxStatus_(value) {
  var text = String(value || "").trim();
  if (
    text === INVOICE_REQUESTS_NOT_APPLICABLE ||
    text === INVOICE_REQUESTS_LEGACY_NOT_APPLICABLE
  ) {
    return "notApplicable";
  }
  if (value === true) return "checked";
  return "unchecked";
}

function getInvoiceRequestOriginalToken_(value, columnOffset) {
  if (!isInvoiceRequestStatusColumn_(columnOffset)) {
    return comparableInvoiceRequestValue_(value);
  }
  return getInvoiceRequestCheckboxStatus_(value);
}

function invoiceRequestLinkFromCell_(value, displayValue) {
  var display = String(displayValue || "").trim();
  var raw = value === null || value === undefined ? "" : String(value).trim();
  var candidate = display || raw;
  if (/^https?:\/\//i.test(candidate)) return candidate;
  if (/^https?:\/\//i.test(raw)) return raw;
  return candidate;
}

function invalidateInvoiceRequestsListCache_() {
  removeCachedJson_(INVOICE_REQUESTS_LIST_CACHE_KEY);
}

function invoiceRequestPayloadRowOwnedByEmail_(row, email) {
  var cells = row && row.cells ? row.cells : [];
  var createdBy = String(
    (cells[INVOICE_REQUESTS_CREATED_BY_OFFSET] &&
      cells[INVOICE_REQUESTS_CREATED_BY_OFFSET].value) ||
      ""
  )
    .trim()
    .toLowerCase();
  var editedBy = String(
    (cells[INVOICE_REQUESTS_EDITED_BY_OFFSET] &&
      cells[INVOICE_REQUESTS_EDITED_BY_OFFSET].value) ||
      ""
  )
    .trim()
    .toLowerCase();
  var normalized = String(email || "")
    .trim()
    .toLowerCase();
  return createdBy === normalized || editedBy === normalized;
}

function buildInvoiceRequestsPayload_(spreadsheet, projectsOrLookup) {
  var accessMode = getInvoiceRequestAccessMode_();
  var email = getCurrentUserEmail();
  var cached = getCachedJson_(INVOICE_REQUESTS_LIST_CACHE_KEY);
  var base;
  if (cached && cached.headers && cached.rows) {
    base = cached;
  } else {
    base = readInvoiceRequestsSheetPayload_(spreadsheet, projectsOrLookup);
    putCachedJson_(
      INVOICE_REQUESTS_LIST_CACHE_KEY,
      {
        headers: base.headers,
        rows: base.rows,
        projects: base.projects,
      },
      DATA_LIST_CACHE_TTL_SECONDS
    );
  }

  var rows = base.rows || [];
  if (accessMode === "limited") {
    rows = rows.filter(function (row) {
      return invoiceRequestPayloadRowOwnedByEmail_(row, email);
    });
  }

  return {
    headers: base.headers || [],
    rows: rows,
    projects: base.projects || [],
    accessMode: accessMode,
    showStatusColumns: accessMode === "full",
    showAuthorColumn: accessMode === "full",
    showClientFolderColumn: accessMode === "full",
  };
}

function readInvoiceRequestsSheetPayload_(spreadsheet, projectsOrLookup) {
  var sheet = spreadsheet.getSheetByName(INVOICE_REQUESTS_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "Requests" was not found.');

  var projects = Array.isArray(projectsOrLookup)
    ? projectsOrLookup
    : projectsOrLookup && projectsOrLookup.projects
    ? projectsOrLookup.projects
    : getInvoiceRequestProjects_(spreadsheet);
  var emptyPayload = {
    headers: [],
    rows: [],
    projects: projects,
  };
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return emptyPayload;

  // Same pattern as Clients Information: values + display only.
  // Links are plain URL text; N/A is a cell symbol (no background).
  var range = sheet.getRange(
    1,
    1,
    lastRow,
    INVOICE_REQUESTS_FIRST_COLUMN +
      INVOICE_REQUESTS_COLUMN_COUNT -
      1
  );
  var values = range.getValues();
  var displayValues = range.getDisplayValues();
  var headers = displayValues[0].slice(
    INVOICE_REQUESTS_FIRST_COLUMN - 1,
    INVOICE_REQUESTS_FIRST_COLUMN -
      1 +
      INVOICE_REQUESTS_COLUMN_COUNT
  );
  var rows = [];
  var seenIds = {};
  var start = INVOICE_REQUESTS_FIRST_COLUMN - 1;

  for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
    var rowId = String(values[rowIndex][0] || "").trim();
    if (!rowId) continue;
    if (seenIds[rowId]) {
      throw new Error("Duplicate row ID in column A: " + rowId);
    }
    seenIds[rowId] = true;

    var cells = [];
    for (
      var columnOffset = 0;
      columnOffset < INVOICE_REQUESTS_COLUMN_COUNT;
      columnOffset++
    ) {
      var sourceColumn = start + columnOffset;
      var isStatusColumn = isInvoiceRequestStatusColumn_(columnOffset);
      var checkboxStatus = isStatusColumn
        ? getInvoiceRequestCheckboxStatus_(values[rowIndex][sourceColumn])
        : null;
      var displayValue = isInvoiceRequestTimestampColumn_(columnOffset)
        ? serializeInvoiceRequestTimestampCell_(
            values[rowIndex][sourceColumn],
            displayValues[rowIndex][sourceColumn]
          )
        : displayValues[rowIndex][sourceColumn] || "";
      var cellValue = isStatusColumn
        ? checkboxStatus
        : isInvoiceRequestTimestampColumn_(columnOffset)
        ? displayValue
        : serializeInvoiceRequestValue_(
            values[rowIndex][sourceColumn],
            displayValues[rowIndex][sourceColumn]
          );
      var link = "";
      if (
        columnOffset === INVOICE_REQUESTS_RATE_FILE_OFFSET ||
        columnOffset === INVOICE_REQUESTS_CLIENT_FOLDER_OFFSET
      ) {
        link = invoiceRequestLinkFromCell_(
          values[rowIndex][sourceColumn],
          displayValues[rowIndex][sourceColumn]
        );
        if (!cellValue && link) cellValue = link;
        if (!displayValue && link) displayValue = link;
      }
      cells.push({
        value: cellValue,
        originalToken: getInvoiceRequestOriginalToken_(
          values[rowIndex][sourceColumn],
          columnOffset
        ),
        displayValue: displayValue,
        link: link,
      });
    }

    var createdAtColumn = start + INVOICE_REQUESTS_CREATED_AT_OFFSET;
    var editedAtColumn = start + INVOICE_REQUESTS_EDITED_AT_OFFSET;
    rows.push({
      id: rowId,
      cells: cells,
      activityAt: getInvoiceRequestActivityAtFromCells_(
        values[rowIndex][createdAtColumn],
        displayValues[rowIndex][createdAtColumn],
        values[rowIndex][editedAtColumn],
        displayValues[rowIndex][editedAtColumn]
      ),
    });
  }

  rows.sort(function (left, right) {
    return (right.activityAt || 0) - (left.activityAt || 0);
  });

  return {
    headers: headers,
    rows: rows,
    projects: projects,
  };
}

function getInvoiceRequests(forceRefresh) {
  assertInvoiceRequestsAccess_();
  if (forceRefresh === true) {
    invalidateInvoiceRequestsListCache_();
  }
  var spreadsheet = SpreadsheetApp.openById(
    INVOICE_REQUESTS_SPREADSHEET_ID
  );
  return buildInvoiceRequestsPayload_(spreadsheet);
}

function writeInvoiceRequestStatus_(range, status) {
  if (status === "notApplicable") {
    range.clearDataValidations();
    range.setValue(INVOICE_REQUESTS_NOT_APPLICABLE);
    range.setBackground(null);
    return;
  }
  if (status !== "checked" && status !== "unchecked") {
    throw new Error("Invalid checkbox status.");
  }
  range.insertCheckboxes();
  range.setValue(status === "checked");
  range.setBackground(null);
}

function createInvoiceRequest(data) {
  assertInvoiceRequestsAccess_();
  var cells = data && Array.isArray(data.cells) ? data.cells : [];
  var project = String(cells[INVOICE_REQUESTS_PROJECT_OFFSET] || "").trim();
  var details = String(cells[INVOICE_REQUESTS_DETAILS_OFFSET] || "").trim();
  var comment = String(cells[INVOICE_REQUESTS_COMMENT_OFFSET] || "");
  if (!project || !details) {
    return {
      success: false,
      validation: true,
      message: "Project and Details are required.",
    };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("The sheet is busy. Please try saving again.");
  }

  var sheet = null;
  var newRow = -1;
  try {
    var spreadsheet = SpreadsheetApp.openById(
      INVOICE_REQUESTS_SPREADSHEET_ID
    );
    var informationLookup = getInvoiceRequestInformationLookup_(
      spreadsheet
    );
    var projectError = assertInvoiceRequestProject_(
      project,
      informationLookup.projects
    );
    if (projectError) return projectError;

    sheet = spreadsheet.getSheetByName(
      INVOICE_REQUESTS_SHEET_NAME
    );
    if (!sheet) throw new Error('Sheet "Requests" was not found.');

    var email = getCurrentUserEmail();
    var listsLookup = getInvoiceRequestListsLookup_(spreadsheet);
    var author = resolveInvoiceRequestAuthor_(
      spreadsheet,
      email,
      listsLookup
    );
    var rateFile = resolveInvoiceRequestRateFile_(
      spreadsheet,
      project,
      informationLookup
    );
    var clientFolder = resolveInvoiceRequestClientFolder_(
      spreadsheet,
      project,
      informationLookup
    );
    var createdAt = new Date();

    var lastRow = sheet.getLastRow();
    var existingIds =
      lastRow > 1
        ? sheet.getRange(2, 1, lastRow - 1, 1).getValues()
        : [];
    var maxId = 0;
    for (var idIndex = 0; idIndex < existingIds.length; idIndex++) {
      var numericId = Number(existingIds[idIndex][0]);
      if (Number.isFinite(numericId)) maxId = Math.max(maxId, numericId);
    }
    var newId = maxId + 1;

    sheet.insertRowAfter(Math.max(lastRow, 1));
    newRow = Math.max(lastRow, 1) + 1;
    if (lastRow > 1) {
      sheet
        .getRange(lastRow, 1, 1, INVOICE_REQUESTS_LAST_COLUMN)
        .copyTo(
          sheet.getRange(newRow, 1, 1, INVOICE_REQUESTS_LAST_COLUMN),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
    }

    sheet
      .getRange(newRow, 1, 1, 5)
      .setValues([[newId, project, details, comment, author]]);

    // Processing statuses F:M belong to the person handling the request.
    // A newly submitted request must leave them completely empty.
    var statusRange = sheet.getRange(
      newRow,
      sheetColumnForInvoiceRequestOffset_(
        INVOICE_REQUESTS_STATUS_FIRST_OFFSET
      ),
      1,
      INVOICE_REQUESTS_STATUS_COUNT
    );
    statusRange.clearContent();
    statusRange.clearDataValidations();
    statusRange.setBackground(null);

    var trailingRange = sheet.getRange(
      newRow,
      sheetColumnForInvoiceRequestOffset_(
        INVOICE_REQUESTS_RATE_FILE_OFFSET
      ),
      1,
      INVOICE_REQUESTS_TRAILING_COLUMN_COUNT
    );
    trailingRange.clearContent();
    trailingRange.clearDataValidations();
    trailingRange.setBackground(null);

    writeInvoiceRequestLinkedValue_(
      sheet.getRange(
        newRow,
        sheetColumnForInvoiceRequestOffset_(
          INVOICE_REQUESTS_RATE_FILE_OFFSET
        )
      ),
      rateFile.value,
      rateFile.link
    );
    writeInvoiceRequestLinkedValue_(
      sheet.getRange(
        newRow,
        sheetColumnForInvoiceRequestOffset_(
          INVOICE_REQUESTS_CLIENT_FOLDER_OFFSET
        )
      ),
      clientFolder.value,
      clientFolder.link
    );
    sheet
      .getRange(
        newRow,
        sheetColumnForInvoiceRequestOffset_(
          INVOICE_REQUESTS_CREATED_BY_OFFSET
        )
      )
      .setValue(email);
    sheet
      .getRange(
        newRow,
        sheetColumnForInvoiceRequestOffset_(
          INVOICE_REQUESTS_CREATED_AT_OFFSET
        )
      )
      .setValue(createdAt);

    SpreadsheetApp.flush();
    invalidateInvoiceRequestsListCache_();
    var payload = buildInvoiceRequestsPayload_(
      spreadsheet,
      informationLookup
    );
    return {
      success: true,
      id: String(newId),
      headers: payload.headers,
      rows: payload.rows,
      projects: payload.projects,
      accessMode: payload.accessMode,
      showStatusColumns: payload.showStatusColumns,
      showAuthorColumn: payload.showAuthorColumn,
      showClientFolderColumn: payload.showClientFolderColumn,
      // Sent by a separate client call so the UI is not blocked on MailApp.
      notifications: [
        {
          kind: "created",
          project: project,
          author: author,
        },
      ],
    };
  } catch (error) {
    if (sheet && newRow > 0) {
      try {
        sheet.deleteRow(newRow);
      } catch (rollbackError) {
        console.error(
          "Could not roll back Invoice Request row:",
          rollbackError
        );
      }
    }
    throw error;
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      console.warn("Could not release Invoice Requests lock:", error);
    }
  }
}

function findInvoiceRequestSheetRowsByIds_(sheet, rowIds) {
  var wanted = {};
  var hasWanted = false;
  for (var i = 0; i < rowIds.length; i++) {
    var id = String(rowIds[i] || "").trim();
    if (!id || wanted[id]) continue;
    wanted[id] = true;
    hasWanted = true;
  }
  if (!hasWanted) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var map = {};
  var found = 0;
  var wantedCount = Object.keys(wanted).length;
  for (var rowIndex = 0; rowIndex < idValues.length; rowIndex++) {
    var rowId = String(idValues[rowIndex][0] || "").trim();
    if (!rowId || !wanted[rowId] || map[rowId]) continue;
    map[rowId] = rowIndex + 2;
    found++;
    if (found >= wantedCount) break;
  }
  return map;
}

function buildInvoiceRequestRowPayloadFromSheet_(sheet, sheetRow, rowId) {
  var range = sheet.getRange(sheetRow, 1, 1, INVOICE_REQUESTS_LAST_COLUMN);
  var valuesRow = range.getValues()[0];
  var displayRow = range.getDisplayValues()[0];
  var start = INVOICE_REQUESTS_FIRST_COLUMN - 1;
  var cells = [];
  for (
    var columnOffset = 0;
    columnOffset < INVOICE_REQUESTS_COLUMN_COUNT;
    columnOffset++
  ) {
    var sourceColumn = start + columnOffset;
    var isStatusColumn = isInvoiceRequestStatusColumn_(columnOffset);
    var checkboxStatus = isStatusColumn
      ? getInvoiceRequestCheckboxStatus_(valuesRow[sourceColumn])
      : null;
    var displayValue = isInvoiceRequestTimestampColumn_(columnOffset)
      ? serializeInvoiceRequestTimestampCell_(
          valuesRow[sourceColumn],
          displayRow[sourceColumn]
        )
      : displayRow[sourceColumn] || "";
    var cellValue = isStatusColumn
      ? checkboxStatus
      : isInvoiceRequestTimestampColumn_(columnOffset)
      ? displayValue
      : serializeInvoiceRequestValue_(
          valuesRow[sourceColumn],
          displayRow[sourceColumn]
        );
    var link = "";
    if (
      columnOffset === INVOICE_REQUESTS_RATE_FILE_OFFSET ||
      columnOffset === INVOICE_REQUESTS_CLIENT_FOLDER_OFFSET
    ) {
      link = invoiceRequestLinkFromCell_(
        valuesRow[sourceColumn],
        displayRow[sourceColumn]
      );
      if (!cellValue && link) cellValue = link;
      if (!displayValue && link) displayValue = link;
    }
    cells.push({
      value: cellValue,
      originalToken: getInvoiceRequestOriginalToken_(
        valuesRow[sourceColumn],
        columnOffset
      ),
      displayValue: displayValue,
      link: link,
    });
  }
  var createdAtColumn = start + INVOICE_REQUESTS_CREATED_AT_OFFSET;
  var editedAtColumn = start + INVOICE_REQUESTS_EDITED_AT_OFFSET;
  return {
    id: String(rowId),
    cells: cells,
    activityAt: getInvoiceRequestActivityAtFromCells_(
      valuesRow[createdAtColumn],
      displayRow[createdAtColumn],
      valuesRow[editedAtColumn],
      displayRow[editedAtColumn]
    ),
  };
}

function getInvoiceRequestProjectOptions() {
  assertInvoiceRequestsAccess_();
  return getInvoiceRequestProjects_(
    SpreadsheetApp.openById(INVOICE_REQUESTS_SPREADSHEET_ID)
  );
}

function saveInvoiceRequestChanges(changes) {
  assertInvoiceRequestsAccess_();
  if (!Array.isArray(changes) || changes.length === 0) {
    return { success: true, updated: 0, patch: true, applied: [] };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("The sheet is busy. Please try saving again.");
  }

  try {
    var accessMode = getInvoiceRequestAccessMode_();
    var email = getCurrentUserEmail();
    var spreadsheet = SpreadsheetApp.openById(
      INVOICE_REQUESTS_SPREADSHEET_ID
    );
    var sheet = spreadsheet.getSheetByName(INVOICE_REQUESTS_SHEET_NAME);
    if (!sheet) throw new Error('Sheet "Requests" was not found.');

    var requestedIds = [];
    for (var prep = 0; prep < changes.length; prep++) {
      requestedIds.push(String((changes[prep] && changes[prep].id) || ""));
    }
    var idToSheetRow = findInvoiceRequestSheetRowsByIds_(sheet, requestedIds);
    var rowValuesCache = {};
    function rowValuesFor_(sheetRow) {
      if (!rowValuesCache[sheetRow]) {
        rowValuesCache[sheetRow] = sheet
          .getRange(sheetRow, 1, 1, INVOICE_REQUESTS_LAST_COLUMN)
          .getValues()[0];
      }
      return rowValuesCache[sheetRow];
    }

    var projects = null;
    var conflicts = [];
    var validated = [];
    var contentEditsByRow = {};

    for (var changeIndex = 0; changeIndex < changes.length; changeIndex++) {
      var change = changes[changeIndex] || {};
      var rowId = String(change.id || "").trim();
      var columnOffset = Number(change.columnOffset);
      if (
        !rowId ||
        !Number.isInteger(columnOffset) ||
        !isInvoiceRequestClientEditableColumn_(columnOffset)
      ) {
        throw new Error("Invalid change payload.");
      }
      if (
        accessMode === "limited" &&
        isInvoiceRequestStatusColumn_(columnOffset)
      ) {
        throw new Error("No permission to edit status columns.");
      }

      var sheetRow = idToSheetRow[rowId];
      if (!sheetRow) {
        conflicts.push({ id: rowId, columnOffset: columnOffset });
        continue;
      }

      var rowValues = rowValuesFor_(sheetRow);
      if (
        accessMode === "limited" &&
        !invoiceRequestRowOwnedByEmail_(rowValues, email)
      ) {
        throw new Error("No permission to edit this invoice request.");
      }

      var sourceColumn =
        INVOICE_REQUESTS_FIRST_COLUMN - 1 + columnOffset;
      if (
        getInvoiceRequestOriginalToken_(
          rowValues[sourceColumn],
          columnOffset
        ) !== change.originalToken
      ) {
        conflicts.push({ id: rowId, columnOffset: columnOffset });
        continue;
      }

      var nextValue =
        change.value === null || change.value === undefined
          ? ""
          : String(change.value);
      if (
        isInvoiceRequestStatusColumn_(columnOffset) &&
        nextValue !== "checked" &&
        nextValue !== "unchecked" &&
        nextValue !== "notApplicable"
      ) {
        throw new Error("Invalid checkbox status.");
      }
      if (
        columnOffset === INVOICE_REQUESTS_PROJECT_OFFSET &&
        String(nextValue || "").trim()
      ) {
        if (!projects) {
          projects = getInvoiceRequestProjects_(spreadsheet);
        }
        var projectError = assertInvoiceRequestProject_(
          String(nextValue).trim(),
          projects
        );
        if (projectError) return projectError;
      }
      if (
        columnOffset === INVOICE_REQUESTS_PROJECT_OFFSET ||
        columnOffset === INVOICE_REQUESTS_DETAILS_OFFSET
      ) {
        if (!String(nextValue || "").trim()) {
          return {
            success: false,
            validation: true,
            message: "Project and Details are required.",
          };
        }
      }

      validated.push({
        sheetRow: sheetRow,
        sheetColumn: sheetColumnForInvoiceRequestOffset_(columnOffset),
        columnOffset: columnOffset,
        value: nextValue,
        rowId: rowId,
      });
      if (isInvoiceRequestContentColumn_(columnOffset)) {
        if (!contentEditsByRow[rowId]) {
          contentEditsByRow[rowId] = {
            sheetRow: sheetRow,
            project: String(
              rowValues[
                INVOICE_REQUESTS_FIRST_COLUMN -
                  1 +
                  INVOICE_REQUESTS_PROJECT_OFFSET
              ] || ""
            ).trim(),
          };
        }
        if (columnOffset === INVOICE_REQUESTS_PROJECT_OFFSET) {
          contentEditsByRow[rowId].project = String(nextValue || "").trim();
        }
      }
    }

    if (conflicts.length > 0) {
      return {
        success: false,
        conflict: true,
        conflicts: conflicts,
        message:
          "Some cells changed in Google Sheets after this page was loaded. Reload before saving.",
      };
    }

    for (var writeIndex = 0; writeIndex < validated.length; writeIndex++) {
      var item = validated[writeIndex];
      if (
        isInvoiceRequestStatusColumn_(item.columnOffset) &&
        contentEditsByRow[item.rowId]
      ) {
        continue;
      }
      var target = sheet.getRange(item.sheetRow, item.sheetColumn);
      if (isInvoiceRequestStatusColumn_(item.columnOffset)) {
        writeInvoiceRequestStatus_(target, item.value);
      } else {
        target.setValue(item.value);
      }
    }

    var author = "";
    var listsLookup = null;
    var editedAt = new Date();
    var contentRowIds = Object.keys(contentEditsByRow);
    if (contentRowIds.length > 0) {
      listsLookup = getInvoiceRequestListsLookup_(spreadsheet);
      author = resolveInvoiceRequestAuthor_(
        spreadsheet,
        email,
        listsLookup
      );
    }
    for (var metaIndex = 0; metaIndex < contentRowIds.length; metaIndex++) {
      var meta = contentEditsByRow[contentRowIds[metaIndex]];
      sheet
        .getRange(
          meta.sheetRow,
          sheetColumnForInvoiceRequestOffset_(
            INVOICE_REQUESTS_AUTHOR_OFFSET
          )
        )
        .setValue(author);
      sheet
        .getRange(
          meta.sheetRow,
          sheetColumnForInvoiceRequestOffset_(
            INVOICE_REQUESTS_EDITED_BY_OFFSET
          )
        )
        .setValue(email);
      sheet
        .getRange(
          meta.sheetRow,
          sheetColumnForInvoiceRequestOffset_(
            INVOICE_REQUESTS_EDITED_AT_OFFSET
          )
        )
        .setValue(editedAt);
      resetInvoiceRequestStatuses_(sheet, meta.sheetRow);
    }

    SpreadsheetApp.flush();
    invalidateInvoiceRequestsListCache_();

    var notifications = [];
    for (var notifyIndex = 0; notifyIndex < contentRowIds.length; notifyIndex++) {
      notifications.push({
        kind: "edited",
        project: contentEditsByRow[contentRowIds[notifyIndex]].project,
        author: author,
      });
    }

    // Content edits: return only touched rows. Status-only: return applied cells.
    if (contentRowIds.length > 0) {
      var patchedRows = [];
      for (var p = 0; p < contentRowIds.length; p++) {
        var patchedId = contentRowIds[p];
        patchedRows.push(
          buildInvoiceRequestRowPayloadFromSheet_(
            sheet,
            contentEditsByRow[patchedId].sheetRow,
            patchedId
          )
        );
      }
      return {
        success: true,
        updated: validated.length,
        patch: true,
        rows: patchedRows,
        notifications: notifications,
      };
    }

    var applied = [];
    for (var a = 0; a < validated.length; a++) {
      var appliedItem = validated[a];
      applied.push({
        id: appliedItem.rowId,
        columnOffset: appliedItem.columnOffset,
        value: appliedItem.value,
        originalToken: isInvoiceRequestStatusColumn_(appliedItem.columnOffset)
          ? appliedItem.value
          : comparableInvoiceRequestValue_(appliedItem.value),
      });
    }
    return {
      success: true,
      updated: validated.length,
      patch: true,
      applied: applied,
      notifications: notifications,
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      console.warn("Could not release Invoice Requests lock:", error);
    }
  }
}
