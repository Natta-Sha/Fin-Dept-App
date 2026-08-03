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
var INVOICE_REQUESTS_NOT_APPLICABLE = "⊟";
var INVOICE_REQUESTS_NOT_APPLICABLE_BACKGROUND = "#d9ead3";

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
  var text = String(value || "").trim();
  var url = String(link || "").trim();
  if (url) {
    range.setRichTextValue(
      SpreadsheetApp.newRichTextValue()
        .setText(text || url)
        .setLinkUrl(url)
        .build()
    );
    return;
  }
  range.clear();
  if (text) range.setValue(text);
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

function normalizeInvoiceRequestBackground_(background) {
  return String(background || "").trim().toLowerCase();
}

function isNeutralInvoiceRequestBackground_(background) {
  var normalized = normalizeInvoiceRequestBackground_(background);
  return (
    normalized === "" ||
    normalized === "#ffffff" ||
    normalized === "#fff" ||
    normalized === "white"
  );
}

function getInvoiceRequestCheckboxStatus_(value, background) {
  if (
    String(value || "").trim() ===
    INVOICE_REQUESTS_NOT_APPLICABLE
  ) {
    return "notApplicable";
  }
  if (value === true) return "checked";
  if (value === false && !isNeutralInvoiceRequestBackground_(background)) {
    return "notApplicable";
  }
  return "unchecked";
}

function getInvoiceRequestOriginalToken_(
  value,
  background,
  columnOffset
) {
  if (!isInvoiceRequestStatusColumn_(columnOffset)) {
    return comparableInvoiceRequestValue_(value);
  }
  return JSON.stringify([
    comparableInvoiceRequestValue_(value),
    normalizeInvoiceRequestBackground_(background),
  ]);
}

function migrateLegacyInvoiceRequestStatuses_(
  sheet,
  values,
  backgrounds
) {
  var migrated = 0;
  for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
    for (
      var columnOffset = INVOICE_REQUESTS_STATUS_FIRST_OFFSET;
      columnOffset <
      INVOICE_REQUESTS_STATUS_FIRST_OFFSET +
        INVOICE_REQUESTS_STATUS_COUNT;
      columnOffset++
    ) {
      var sourceColumn =
        INVOICE_REQUESTS_FIRST_COLUMN - 1 + columnOffset;
      if (
        values[rowIndex][sourceColumn] === false &&
        !isNeutralInvoiceRequestBackground_(
          backgrounds[rowIndex][sourceColumn]
        )
      ) {
        var cell = sheet.getRange(rowIndex + 1, sourceColumn + 1);
        cell.clearDataValidations();
        cell.setValue(INVOICE_REQUESTS_NOT_APPLICABLE);
        cell.setBackground(
          INVOICE_REQUESTS_NOT_APPLICABLE_BACKGROUND
        );
        values[rowIndex][sourceColumn] =
          INVOICE_REQUESTS_NOT_APPLICABLE;
        backgrounds[rowIndex][sourceColumn] =
          INVOICE_REQUESTS_NOT_APPLICABLE_BACKGROUND;
        migrated++;
      }
    }
  }
  if (migrated > 0) SpreadsheetApp.flush();
}

function buildInvoiceRequestsPayload_(spreadsheet, projectsOrLookup) {
  var sheet = spreadsheet.getSheetByName(INVOICE_REQUESTS_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "Requests" was not found.');

  var accessMode = getInvoiceRequestAccessMode_();
  var email = getCurrentUserEmail();
  var projects = Array.isArray(projectsOrLookup)
    ? projectsOrLookup
    : projectsOrLookup && projectsOrLookup.projects
    ? projectsOrLookup.projects
    : getInvoiceRequestProjects_(spreadsheet);
  var emptyPayload = {
    headers: [],
    rows: [],
    projects: projects,
    accessMode: accessMode,
    showStatusColumns: accessMode === "full",
    showAuthorColumn: accessMode === "full",
    showClientFolderColumn: accessMode === "full",
  };
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return emptyPayload;

  // Column A is read only as a stable hidden row key.
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
  var richTextValues = range.getRichTextValues();
  var backgrounds = range.getBackgrounds();
  if (accessMode === "full") {
    migrateLegacyInvoiceRequestStatuses_(sheet, values, backgrounds);
  }
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
    if (
      accessMode === "limited" &&
      !invoiceRequestRowOwnedByEmail_(values[rowIndex], email)
    ) {
      continue;
    }

    var cells = [];
    for (
      var columnOffset = 0;
      columnOffset < INVOICE_REQUESTS_COLUMN_COUNT;
      columnOffset++
    ) {
      var sourceColumn = start + columnOffset;
      var richText = richTextValues[rowIndex][sourceColumn];
      var isStatusColumn = isInvoiceRequestStatusColumn_(columnOffset);
      var checkboxStatus = isStatusColumn
        ? getInvoiceRequestCheckboxStatus_(
            values[rowIndex][sourceColumn],
            backgrounds[rowIndex][sourceColumn]
          )
        : null;
      cells.push({
        value: isStatusColumn
          ? checkboxStatus
          : isInvoiceRequestTimestampColumn_(columnOffset)
          ? serializeInvoiceRequestTimestampCell_(
              values[rowIndex][sourceColumn],
              displayValues[rowIndex][sourceColumn]
            )
          : serializeInvoiceRequestValue_(
              values[rowIndex][sourceColumn],
              displayValues[rowIndex][sourceColumn]
            ),
        originalToken: getInvoiceRequestOriginalToken_(
          values[rowIndex][sourceColumn],
          backgrounds[rowIndex][sourceColumn],
          columnOffset
        ),
        displayValue: isInvoiceRequestTimestampColumn_(columnOffset)
          ? serializeInvoiceRequestTimestampCell_(
              values[rowIndex][sourceColumn],
              displayValues[rowIndex][sourceColumn]
            )
          : displayValues[rowIndex][sourceColumn] || "",
        link: richText ? richText.getLinkUrl() || "" : "",
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
    accessMode: accessMode,
    showStatusColumns: accessMode === "full",
    showAuthorColumn: accessMode === "full",
    showClientFolderColumn: accessMode === "full",
  };
}

function getInvoiceRequests() {
  assertInvoiceRequestsAccess_();
  var spreadsheet = SpreadsheetApp.openById(
    INVOICE_REQUESTS_SPREADSHEET_ID
  );
  return buildInvoiceRequestsPayload_(spreadsheet);
}

function writeInvoiceRequestStatus_(range, status) {
  if (status === "notApplicable") {
    range.clearDataValidations();
    range.setValue(INVOICE_REQUESTS_NOT_APPLICABLE);
    range.setBackground(INVOICE_REQUESTS_NOT_APPLICABLE_BACKGROUND);
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

function saveInvoiceRequestChanges(changes) {
  assertInvoiceRequestsAccess_();
  if (!Array.isArray(changes) || changes.length === 0) {
    return { success: true, updated: 0 };
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

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("No editable rows were found.");

    var sourceRange = sheet.getRange(
      2,
      1,
      lastRow - 1,
      INVOICE_REQUESTS_FIRST_COLUMN +
        INVOICE_REQUESTS_COLUMN_COUNT -
        1
    );
    var source = sourceRange.getValues();
    var sourceBackgrounds = sourceRange.getBackgrounds();
    var projects = null;
    var rowsById = {};
    for (var sourceIndex = 0; sourceIndex < source.length; sourceIndex++) {
      var sourceId = String(source[sourceIndex][0] || "").trim();
      if (sourceId && rowsById[sourceId] !== undefined) {
        throw new Error("Duplicate row ID in column A: " + sourceId);
      }
      if (sourceId) {
        rowsById[sourceId] = {
          sheetRow: sourceIndex + 2,
          values: source[sourceIndex],
          backgrounds: sourceBackgrounds[sourceIndex],
        };
      }
    }

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

      var targetRow = rowsById[rowId];
      if (!targetRow) {
        conflicts.push({ id: rowId, columnOffset: columnOffset });
        continue;
      }
      if (
        accessMode === "limited" &&
        !invoiceRequestRowOwnedByEmail_(targetRow.values, email)
      ) {
        throw new Error("No permission to edit this invoice request.");
      }

      var sourceColumn =
        INVOICE_REQUESTS_FIRST_COLUMN - 1 + columnOffset;
      if (
        getInvoiceRequestOriginalToken_(
          targetRow.values[sourceColumn],
          targetRow.backgrounds[sourceColumn],
          columnOffset
        ) !==
        change.originalToken
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
        sheetRow: targetRow.sheetRow,
        sheetColumn: sheetColumnForInvoiceRequestOffset_(columnOffset),
        columnOffset: columnOffset,
        value: nextValue,
        rowId: rowId,
      });
      if (isInvoiceRequestContentColumn_(columnOffset)) {
        if (!contentEditsByRow[rowId]) {
          contentEditsByRow[rowId] = {
            sheetRow: targetRow.sheetRow,
            project: String(
              targetRow.values[
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

    // The browser makes one save request. Checkbox cells are written
    // individually because N/A deliberately uses a symbol instead of Sheets
    // checkbox validation. Status clicks on rows with content edits are
    // ignored because those statuses are reset below.
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
      // Rate file and client folder are snapshotted only on create.
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

    var notifications = [];
    for (var notifyIndex = 0; notifyIndex < contentRowIds.length; notifyIndex++) {
      notifications.push({
        kind: "edited",
        project: contentEditsByRow[contentRowIds[notifyIndex]].project,
        author: author,
      });
    }

    var payload = buildInvoiceRequestsPayload_(spreadsheet, projects);
    return {
      success: true,
      updated: validated.length,
      headers: payload.headers,
      rows: payload.rows,
      projects: payload.projects,
      accessMode: payload.accessMode,
      showStatusColumns: payload.showStatusColumns,
      showAuthorColumn: payload.showAuthorColumn,
      showClientFolderColumn: payload.showClientFolderColumn,
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
