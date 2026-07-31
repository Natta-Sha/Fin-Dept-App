// Isolated data service for the Tabulator invoice-requests experiment.
// Removing this file and its page/navigation hooks removes the feature.

var TEST_INVOICE_REQUESTS_SPREADSHEET_ID =
  "1-5b98hkm-wsrTlyt-1j2PJcTvu1VNyAh0EjRhT3ZBg0";
var TEST_INVOICE_REQUESTS_SHEET_NAME = "Requests";
var TEST_INVOICE_REQUESTS_INFORMATION_SHEET = "Information";
var TEST_INVOICE_REQUESTS_LISTS_SHEET = "Lists";
var TEST_INVOICE_REQUESTS_FIRST_COLUMN = 2; // B
var TEST_INVOICE_REQUESTS_COLUMN_COUNT = 17; // B:R
var TEST_INVOICE_REQUESTS_LAST_COLUMN = 18; // R
var TEST_INVOICE_REQUESTS_STATUS_FIRST_OFFSET = 4; // F
var TEST_INVOICE_REQUESTS_STATUS_COUNT = 8; // F:M
var TEST_INVOICE_REQUESTS_PROJECT_OFFSET = 0; // B
var TEST_INVOICE_REQUESTS_DETAILS_OFFSET = 1; // C
var TEST_INVOICE_REQUESTS_COMMENT_OFFSET = 2; // D
var TEST_INVOICE_REQUESTS_AUTHOR_OFFSET = 3; // E
var TEST_INVOICE_REQUESTS_RATE_FILE_OFFSET = 12; // N
var TEST_INVOICE_REQUESTS_CREATED_BY_OFFSET = 13; // O
var TEST_INVOICE_REQUESTS_CREATED_AT_OFFSET = 14; // P
var TEST_INVOICE_REQUESTS_EDITED_BY_OFFSET = 15; // Q
var TEST_INVOICE_REQUESTS_EDITED_AT_OFFSET = 16; // R
var TEST_INVOICE_REQUESTS_NOT_APPLICABLE = "⊟";
var TEST_INVOICE_REQUESTS_NOT_APPLICABLE_BACKGROUND = "#d9ead3";

function isTestInvoiceRequestStatusColumn_(columnOffset) {
  return (
    columnOffset >= TEST_INVOICE_REQUESTS_STATUS_FIRST_OFFSET &&
    columnOffset <
      TEST_INVOICE_REQUESTS_STATUS_FIRST_OFFSET +
        TEST_INVOICE_REQUESTS_STATUS_COUNT
  );
}

function isTestInvoiceRequestContentColumn_(columnOffset) {
  return (
    columnOffset === TEST_INVOICE_REQUESTS_PROJECT_OFFSET ||
    columnOffset === TEST_INVOICE_REQUESTS_DETAILS_OFFSET ||
    columnOffset === TEST_INVOICE_REQUESTS_COMMENT_OFFSET
  );
}

function isTestInvoiceRequestClientEditableColumn_(columnOffset) {
  return (
    isTestInvoiceRequestContentColumn_(columnOffset) ||
    isTestInvoiceRequestStatusColumn_(columnOffset)
  );
}

function formatTestInvoiceRequestTimestamp_(date) {
  return Utilities.formatDate(
    date || new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );
}

function parseTestInvoiceRequestTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getTime();
  }
  var text = String(value || "").trim();
  if (!text) return 0;
  var match = text.match(
    /^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (!match) return 0;
  return new Date(
    Number(match[3]),
    Number(match[2]) - 1,
    Number(match[1]),
    Number(match[4] || 0),
    Number(match[5] || 0),
    Number(match[6] || 0)
  ).getTime();
}

function getTestInvoiceRequestActivityAt_(createdAt, editedAt) {
  var editedMs = parseTestInvoiceRequestTimestamp_(editedAt);
  if (editedMs) return editedMs;
  return parseTestInvoiceRequestTimestamp_(createdAt);
}

function sheetColumnForTestInvoiceRequestOffset_(columnOffset) {
  return TEST_INVOICE_REQUESTS_FIRST_COLUMN + columnOffset;
}

function resetTestInvoiceRequestStatuses_(sheet, sheetRow) {
  var statusRange = sheet.getRange(
    sheetRow,
    sheetColumnForTestInvoiceRequestOffset_(
      TEST_INVOICE_REQUESTS_STATUS_FIRST_OFFSET
    ),
    1,
    TEST_INVOICE_REQUESTS_STATUS_COUNT
  );
  var uncheckedRow = [];
  for (var i = 0; i < TEST_INVOICE_REQUESTS_STATUS_COUNT; i++) {
    uncheckedRow.push(false);
  }
  statusRange.insertCheckboxes();
  statusRange.setValues([uncheckedRow]);
  statusRange.setBackground(null);
}

function getTestInvoiceRequestInformationLookup_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(
    TEST_INVOICE_REQUESTS_INFORMATION_SHEET
  );
  if (!sheet) throw new Error('Sheet "Information" was not found.');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { projects: [], rateByProject: {} };
  }

  // B:E in one read: project in B, rate-file link in E.
  var range = sheet.getRange(2, 2, lastRow - 1, 4);
  var displayValues = range.getDisplayValues();
  var richTextValues = range.getRichTextValues();
  var seen = {};
  var projects = [];
  var rateByProject = {};
  for (var i = 0; i < displayValues.length; i++) {
    var project = String(displayValues[i][0] || "").trim();
    if (!project) continue;
    if (!seen[project]) {
      seen[project] = true;
      projects.push(project);
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
  return { projects: projects, rateByProject: rateByProject };
}

function getTestInvoiceRequestProjects_(spreadsheet) {
  return getTestInvoiceRequestInformationLookup_(spreadsheet).projects;
}

function resolveTestInvoiceRequestAuthor_(spreadsheet, email) {
  var normalizedEmail = String(email || "").trim().toLowerCase();
  if (!normalizedEmail) return "";

  var sheet = spreadsheet.getSheetByName(TEST_INVOICE_REQUESTS_LISTS_SHEET);
  if (!sheet) throw new Error('Sheet "Lists" was not found.');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return normalizedEmail;

  var values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var listEmail = String(values[i][1] || "").trim().toLowerCase();
    if (listEmail === normalizedEmail) {
      var fullName = String(values[i][0] || "").trim();
      return fullName || normalizedEmail;
    }
  }
  return normalizedEmail;
}

function resolveTestInvoiceRequestRateFile_(
  spreadsheet,
  project,
  informationLookup
) {
  var normalizedProject = String(project || "").trim();
  if (!normalizedProject) return { value: "", link: "" };
  var lookup =
    informationLookup || getTestInvoiceRequestInformationLookup_(spreadsheet);
  return lookup.rateByProject[normalizedProject] || { value: "", link: "" };
}

function writeTestInvoiceRequestLinkedValue_(range, value, link) {
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

function assertTestInvoiceRequestProject_(project, projects) {
  if (!projects || projects.indexOf(project) === -1) {
    return {
      success: false,
      validation: true,
      message: "Select a project from the list.",
    };
  }
  return null;
}

function assertTestInvoiceRequestsAccess_() {
  if (!isFullAccessUser(getCurrentUserEmail())) {
    throw new Error("No permission to access Test Invoice Requests.");
  }
}

function serializeTestInvoiceRequestValue_(value, displayValue) {
  if (value instanceof Date) return displayValue || "";
  if (value === null || value === undefined) return "";
  return value;
}

function comparableTestInvoiceRequestValue_(value) {
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

function normalizeTestInvoiceRequestBackground_(background) {
  return String(background || "").trim().toLowerCase();
}

function isNeutralTestInvoiceRequestBackground_(background) {
  var normalized = normalizeTestInvoiceRequestBackground_(background);
  return (
    normalized === "" ||
    normalized === "#ffffff" ||
    normalized === "#fff" ||
    normalized === "white"
  );
}

function getTestInvoiceRequestCheckboxStatus_(value, background) {
  if (
    String(value || "").trim() ===
    TEST_INVOICE_REQUESTS_NOT_APPLICABLE
  ) {
    return "notApplicable";
  }
  if (value === true) return "checked";
  if (value === false && !isNeutralTestInvoiceRequestBackground_(background)) {
    return "notApplicable";
  }
  return "unchecked";
}

function getTestInvoiceRequestOriginalToken_(
  value,
  background,
  columnOffset
) {
  if (!isTestInvoiceRequestStatusColumn_(columnOffset)) {
    return comparableTestInvoiceRequestValue_(value);
  }
  return JSON.stringify([
    comparableTestInvoiceRequestValue_(value),
    normalizeTestInvoiceRequestBackground_(background),
  ]);
}

function migrateLegacyTestInvoiceRequestStatuses_(
  sheet,
  values,
  backgrounds
) {
  var migrated = 0;
  for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
    for (
      var columnOffset = TEST_INVOICE_REQUESTS_STATUS_FIRST_OFFSET;
      columnOffset <
      TEST_INVOICE_REQUESTS_STATUS_FIRST_OFFSET +
        TEST_INVOICE_REQUESTS_STATUS_COUNT;
      columnOffset++
    ) {
      var sourceColumn =
        TEST_INVOICE_REQUESTS_FIRST_COLUMN - 1 + columnOffset;
      if (
        values[rowIndex][sourceColumn] === false &&
        !isNeutralTestInvoiceRequestBackground_(
          backgrounds[rowIndex][sourceColumn]
        )
      ) {
        var cell = sheet.getRange(rowIndex + 1, sourceColumn + 1);
        cell.clearDataValidations();
        cell.setValue(TEST_INVOICE_REQUESTS_NOT_APPLICABLE);
        cell.setBackground(
          TEST_INVOICE_REQUESTS_NOT_APPLICABLE_BACKGROUND
        );
        values[rowIndex][sourceColumn] =
          TEST_INVOICE_REQUESTS_NOT_APPLICABLE;
        backgrounds[rowIndex][sourceColumn] =
          TEST_INVOICE_REQUESTS_NOT_APPLICABLE_BACKGROUND;
        migrated++;
      }
    }
  }
  if (migrated > 0) SpreadsheetApp.flush();
}

function buildTestInvoiceRequestsPayload_(spreadsheet, informationLookup) {
  var sheet = spreadsheet.getSheetByName(TEST_INVOICE_REQUESTS_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "Requests" was not found.');

  var lookup =
    informationLookup || getTestInvoiceRequestInformationLookup_(spreadsheet);
  var projects = lookup.projects;
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return { headers: [], rows: [], projects: projects };

  // Column A is read only as a stable hidden row key.
  var range = sheet.getRange(
    1,
    1,
    lastRow,
    TEST_INVOICE_REQUESTS_FIRST_COLUMN +
      TEST_INVOICE_REQUESTS_COLUMN_COUNT -
      1
  );
  var values = range.getValues();
  var displayValues = range.getDisplayValues();
  var richTextValues = range.getRichTextValues();
  var backgrounds = range.getBackgrounds();
  migrateLegacyTestInvoiceRequestStatuses_(sheet, values, backgrounds);
  var headers = displayValues[0].slice(
    TEST_INVOICE_REQUESTS_FIRST_COLUMN - 1,
    TEST_INVOICE_REQUESTS_FIRST_COLUMN -
      1 +
      TEST_INVOICE_REQUESTS_COLUMN_COUNT
  );
  var rows = [];
  var seenIds = {};

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
      columnOffset < TEST_INVOICE_REQUESTS_COLUMN_COUNT;
      columnOffset++
    ) {
      var sourceColumn =
        TEST_INVOICE_REQUESTS_FIRST_COLUMN - 1 + columnOffset;
      var richText = richTextValues[rowIndex][sourceColumn];
      var isStatusColumn = isTestInvoiceRequestStatusColumn_(columnOffset);
      var checkboxStatus = isStatusColumn
        ? getTestInvoiceRequestCheckboxStatus_(
            values[rowIndex][sourceColumn],
            backgrounds[rowIndex][sourceColumn]
          )
        : null;
      cells.push({
        value: isStatusColumn
          ? checkboxStatus
          : serializeTestInvoiceRequestValue_(
              values[rowIndex][sourceColumn],
              displayValues[rowIndex][sourceColumn]
            ),
        originalToken: getTestInvoiceRequestOriginalToken_(
          values[rowIndex][sourceColumn],
          backgrounds[rowIndex][sourceColumn],
          columnOffset
        ),
        displayValue: displayValues[rowIndex][sourceColumn] || "",
        link: richText ? richText.getLinkUrl() || "" : "",
      });
    }

    var createdAtColumn =
      TEST_INVOICE_REQUESTS_FIRST_COLUMN -
      1 +
      TEST_INVOICE_REQUESTS_CREATED_AT_OFFSET;
    var editedAtColumn =
      TEST_INVOICE_REQUESTS_FIRST_COLUMN -
      1 +
      TEST_INVOICE_REQUESTS_EDITED_AT_OFFSET;
    rows.push({
      id: rowId,
      cells: cells,
      activityAt: getTestInvoiceRequestActivityAt_(
        displayValues[rowIndex][createdAtColumn] ||
          values[rowIndex][createdAtColumn],
        displayValues[rowIndex][editedAtColumn] ||
          values[rowIndex][editedAtColumn]
      ),
    });
  }

  rows.sort(function (left, right) {
    return (right.activityAt || 0) - (left.activityAt || 0);
  });

  return { headers: headers, rows: rows, projects: projects };
}

function getTestInvoiceRequests() {
  assertTestInvoiceRequestsAccess_();
  var spreadsheet = SpreadsheetApp.openById(
    TEST_INVOICE_REQUESTS_SPREADSHEET_ID
  );
  return buildTestInvoiceRequestsPayload_(spreadsheet);
}

function writeTestInvoiceRequestStatus_(range, status) {
  if (status === "notApplicable") {
    range.clearDataValidations();
    range.setValue(TEST_INVOICE_REQUESTS_NOT_APPLICABLE);
    range.setBackground(TEST_INVOICE_REQUESTS_NOT_APPLICABLE_BACKGROUND);
    return;
  }
  if (status !== "checked" && status !== "unchecked") {
    throw new Error("Invalid checkbox status.");
  }
  range.insertCheckboxes();
  range.setValue(status === "checked");
  range.setBackground(null);
}

function createTestInvoiceRequest(data) {
  assertTestInvoiceRequestsAccess_();
  var cells = data && Array.isArray(data.cells) ? data.cells : [];
  var project = String(cells[TEST_INVOICE_REQUESTS_PROJECT_OFFSET] || "").trim();
  var details = String(cells[TEST_INVOICE_REQUESTS_DETAILS_OFFSET] || "").trim();
  var comment = String(cells[TEST_INVOICE_REQUESTS_COMMENT_OFFSET] || "");
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
      TEST_INVOICE_REQUESTS_SPREADSHEET_ID
    );
    var informationLookup = getTestInvoiceRequestInformationLookup_(
      spreadsheet
    );
    var projectError = assertTestInvoiceRequestProject_(
      project,
      informationLookup.projects
    );
    if (projectError) return projectError;

    sheet = spreadsheet.getSheetByName(
      TEST_INVOICE_REQUESTS_SHEET_NAME
    );
    if (!sheet) throw new Error('Sheet "Requests" was not found.');

    var email = getCurrentUserEmail();
    var author = resolveTestInvoiceRequestAuthor_(spreadsheet, email);
    var rateFile = resolveTestInvoiceRequestRateFile_(
      spreadsheet,
      project,
      informationLookup
    );
    var createdAt = formatTestInvoiceRequestTimestamp_();

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
        .getRange(lastRow, 1, 1, TEST_INVOICE_REQUESTS_LAST_COLUMN)
        .copyTo(
          sheet.getRange(newRow, 1, 1, TEST_INVOICE_REQUESTS_LAST_COLUMN),
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
      sheetColumnForTestInvoiceRequestOffset_(
        TEST_INVOICE_REQUESTS_STATUS_FIRST_OFFSET
      ),
      1,
      TEST_INVOICE_REQUESTS_STATUS_COUNT
    );
    statusRange.clearContent();
    statusRange.clearDataValidations();
    statusRange.setBackground(null);

    var trailingRange = sheet.getRange(
      newRow,
      sheetColumnForTestInvoiceRequestOffset_(
        TEST_INVOICE_REQUESTS_RATE_FILE_OFFSET
      ),
      1,
      5
    );
    trailingRange.clearContent();
    trailingRange.clearDataValidations();
    trailingRange.setBackground(null);

    writeTestInvoiceRequestLinkedValue_(
      sheet.getRange(
        newRow,
        sheetColumnForTestInvoiceRequestOffset_(
          TEST_INVOICE_REQUESTS_RATE_FILE_OFFSET
        )
      ),
      rateFile.value,
      rateFile.link
    );
    sheet
      .getRange(
        newRow,
        sheetColumnForTestInvoiceRequestOffset_(
          TEST_INVOICE_REQUESTS_CREATED_BY_OFFSET
        )
      )
      .setValue(email);
    sheet
      .getRange(
        newRow,
        sheetColumnForTestInvoiceRequestOffset_(
          TEST_INVOICE_REQUESTS_CREATED_AT_OFFSET
        )
      )
      .setValue(createdAt);

    SpreadsheetApp.flush();
    var payload = buildTestInvoiceRequestsPayload_(
      spreadsheet,
      informationLookup
    );
    return {
      success: true,
      id: String(newId),
      headers: payload.headers,
      rows: payload.rows,
      projects: payload.projects,
    };
  } catch (error) {
    if (sheet && newRow > 0) {
      try {
        sheet.deleteRow(newRow);
      } catch (rollbackError) {
        console.error(
          "Could not roll back Test Invoice Request row:",
          rollbackError
        );
      }
    }
    throw error;
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      console.warn("Could not release Test Invoice Requests lock:", error);
    }
  }
}

function saveTestInvoiceRequestChanges(changes) {
  assertTestInvoiceRequestsAccess_();
  if (!Array.isArray(changes) || changes.length === 0) {
    return { success: true, updated: 0 };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("The sheet is busy. Please try saving again.");
  }

  try {
    var spreadsheet = SpreadsheetApp.openById(
      TEST_INVOICE_REQUESTS_SPREADSHEET_ID
    );
    var sheet = spreadsheet.getSheetByName(TEST_INVOICE_REQUESTS_SHEET_NAME);
    if (!sheet) throw new Error('Sheet "Requests" was not found.');

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("No editable rows were found.");

    var sourceRange = sheet.getRange(
      2,
      1,
      lastRow - 1,
      TEST_INVOICE_REQUESTS_FIRST_COLUMN +
        TEST_INVOICE_REQUESTS_COLUMN_COUNT -
        1
    );
    var source = sourceRange.getValues();
    var sourceBackgrounds = sourceRange.getBackgrounds();
    var informationLookup = null;
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
        !isTestInvoiceRequestClientEditableColumn_(columnOffset)
      ) {
        throw new Error("Invalid change payload.");
      }

      var targetRow = rowsById[rowId];
      if (!targetRow) {
        conflicts.push({ id: rowId, columnOffset: columnOffset });
        continue;
      }

      var sourceColumn =
        TEST_INVOICE_REQUESTS_FIRST_COLUMN - 1 + columnOffset;
      if (
        getTestInvoiceRequestOriginalToken_(
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
        isTestInvoiceRequestStatusColumn_(columnOffset) &&
        nextValue !== "checked" &&
        nextValue !== "unchecked" &&
        nextValue !== "notApplicable"
      ) {
        throw new Error("Invalid checkbox status.");
      }
      if (
        columnOffset === TEST_INVOICE_REQUESTS_PROJECT_OFFSET &&
        String(nextValue || "").trim()
      ) {
        if (!informationLookup) {
          informationLookup = getTestInvoiceRequestInformationLookup_(
            spreadsheet
          );
        }
        var projectError = assertTestInvoiceRequestProject_(
          String(nextValue).trim(),
          informationLookup.projects
        );
        if (projectError) return projectError;
      }
      if (
        columnOffset === TEST_INVOICE_REQUESTS_PROJECT_OFFSET ||
        columnOffset === TEST_INVOICE_REQUESTS_DETAILS_OFFSET
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
        sheetColumn: sheetColumnForTestInvoiceRequestOffset_(columnOffset),
        columnOffset: columnOffset,
        value: nextValue,
        rowId: rowId,
      });
      if (isTestInvoiceRequestContentColumn_(columnOffset)) {
        if (!contentEditsByRow[rowId]) {
          contentEditsByRow[rowId] = {
            sheetRow: targetRow.sheetRow,
            project: String(
              targetRow.values[
                TEST_INVOICE_REQUESTS_FIRST_COLUMN -
                  1 +
                  TEST_INVOICE_REQUESTS_PROJECT_OFFSET
              ] || ""
            ).trim(),
          };
        }
        if (columnOffset === TEST_INVOICE_REQUESTS_PROJECT_OFFSET) {
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
        isTestInvoiceRequestStatusColumn_(item.columnOffset) &&
        contentEditsByRow[item.rowId]
      ) {
        continue;
      }
      var target = sheet.getRange(item.sheetRow, item.sheetColumn);
      if (isTestInvoiceRequestStatusColumn_(item.columnOffset)) {
        writeTestInvoiceRequestStatus_(target, item.value);
      } else {
        target.setValue(item.value);
      }
    }

    var email = getCurrentUserEmail();
    var author = "";
    var editedAt = formatTestInvoiceRequestTimestamp_();
    var contentRowIds = Object.keys(contentEditsByRow);
    if (contentRowIds.length > 0) {
      if (!informationLookup) {
        informationLookup = getTestInvoiceRequestInformationLookup_(
          spreadsheet
        );
      }
      author = resolveTestInvoiceRequestAuthor_(spreadsheet, email);
    }
    for (var metaIndex = 0; metaIndex < contentRowIds.length; metaIndex++) {
      var meta = contentEditsByRow[contentRowIds[metaIndex]];
      var rateFile = resolveTestInvoiceRequestRateFile_(
        spreadsheet,
        meta.project,
        informationLookup
      );
      sheet
        .getRange(
          meta.sheetRow,
          sheetColumnForTestInvoiceRequestOffset_(
            TEST_INVOICE_REQUESTS_AUTHOR_OFFSET
          )
        )
        .setValue(author);
      writeTestInvoiceRequestLinkedValue_(
        sheet.getRange(
          meta.sheetRow,
          sheetColumnForTestInvoiceRequestOffset_(
            TEST_INVOICE_REQUESTS_RATE_FILE_OFFSET
          )
        ),
        rateFile.value,
        rateFile.link
      );
      sheet
        .getRange(
          meta.sheetRow,
          sheetColumnForTestInvoiceRequestOffset_(
            TEST_INVOICE_REQUESTS_EDITED_BY_OFFSET
          )
        )
        .setValue(email);
      sheet
        .getRange(
          meta.sheetRow,
          sheetColumnForTestInvoiceRequestOffset_(
            TEST_INVOICE_REQUESTS_EDITED_AT_OFFSET
          )
        )
        .setValue(editedAt);
      resetTestInvoiceRequestStatuses_(sheet, meta.sheetRow);
    }

    SpreadsheetApp.flush();

    var payload = buildTestInvoiceRequestsPayload_(
      spreadsheet,
      informationLookup
    );
    return {
      success: true,
      updated: validated.length,
      headers: payload.headers,
      rows: payload.rows,
      projects: payload.projects,
    };
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      console.warn("Could not release Test Invoice Requests lock:", error);
    }
  }
}
