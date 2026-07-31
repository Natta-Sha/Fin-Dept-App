// Isolated data service for the Tabulator invoice-requests experiment.
// Removing this file and its page/navigation hooks removes the feature.

var TEST_INVOICE_REQUESTS_SPREADSHEET_ID =
  "1-5b98hkm-wsrTlyt-1j2PJcTvu1VNyAh0EjRhT3ZBg0";
var TEST_INVOICE_REQUESTS_SHEET_NAME = "Requests";
var TEST_INVOICE_REQUESTS_FIRST_COLUMN = 2; // B
var TEST_INVOICE_REQUESTS_COLUMN_COUNT = 12; // B:M
var TEST_INVOICE_REQUESTS_NOT_APPLICABLE = "⊟";
var TEST_INVOICE_REQUESTS_NOT_APPLICABLE_BACKGROUND = "#d9ead3";

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
  if (columnOffset < 4) {
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
      var columnOffset = 4;
      columnOffset < TEST_INVOICE_REQUESTS_COLUMN_COUNT;
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

function getTestInvoiceRequests() {
  assertTestInvoiceRequestsAccess_();

  var spreadsheet = SpreadsheetApp.openById(
    TEST_INVOICE_REQUESTS_SPREADSHEET_ID
  );
  var sheet = spreadsheet.getSheetByName(TEST_INVOICE_REQUESTS_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "Requests" was not found.');

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return { headers: [], rows: [] };

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
      var checkboxStatus =
        columnOffset >= 4
          ? getTestInvoiceRequestCheckboxStatus_(
              values[rowIndex][sourceColumn],
              backgrounds[rowIndex][sourceColumn]
            )
          : null;
      cells.push({
        value:
          columnOffset >= 4
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

    rows.push({ id: rowId, cells: cells });
  }

  return { headers: headers, rows: rows };
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
  var project = String(cells[0] || "").trim();
  var details = String(cells[1] || "").trim();
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
    sheet = spreadsheet.getSheetByName(
      TEST_INVOICE_REQUESTS_SHEET_NAME
    );
    if (!sheet) throw new Error('Sheet "Requests" was not found.');

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
        .getRange(lastRow, 1, 1, 13)
        .copyTo(
          sheet.getRange(newRow, 1, 1, 13),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
    }

    var textValues = [
      newId,
      project,
      details,
      String(cells[2] || ""),
      String(cells[3] || ""),
    ];
    sheet.getRange(newRow, 1, 1, textValues.length).setValues([textValues]);

    // Processing statuses F:M belong to the person handling the request.
    // A newly submitted request must leave them completely empty.
    var statusRange = sheet.getRange(
      newRow,
      TEST_INVOICE_REQUESTS_FIRST_COLUMN + 4,
      1,
      TEST_INVOICE_REQUESTS_COLUMN_COUNT - 4
    );
    statusRange.clearContent();
    statusRange.clearDataValidations();
    statusRange.setBackground(null);
    SpreadsheetApp.flush();
    return { success: true, id: String(newId) };
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

    var source = sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        TEST_INVOICE_REQUESTS_FIRST_COLUMN +
          TEST_INVOICE_REQUESTS_COLUMN_COUNT -
          1
      )
      .getValues();
    var sourceBackgrounds = sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        TEST_INVOICE_REQUESTS_FIRST_COLUMN +
          TEST_INVOICE_REQUESTS_COLUMN_COUNT -
          1
      )
      .getBackgrounds();
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
    for (var changeIndex = 0; changeIndex < changes.length; changeIndex++) {
      var change = changes[changeIndex] || {};
      var rowId = String(change.id || "").trim();
      var columnOffset = Number(change.columnOffset);
      if (
        !rowId ||
        !Number.isInteger(columnOffset) ||
        columnOffset < 0 ||
        columnOffset >= TEST_INVOICE_REQUESTS_COLUMN_COUNT
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
        columnOffset >= 4 &&
        nextValue !== "checked" &&
        nextValue !== "unchecked" &&
        nextValue !== "notApplicable"
      ) {
        throw new Error("Invalid checkbox status.");
      }
      validated.push({
        sheetRow: targetRow.sheetRow,
        sheetColumn: TEST_INVOICE_REQUESTS_FIRST_COLUMN + columnOffset,
        columnOffset: columnOffset,
        value: nextValue,
      });
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
    // checkbox validation.
    for (var writeIndex = 0; writeIndex < validated.length; writeIndex++) {
      var item = validated[writeIndex];
      var target = sheet.getRange(item.sheetRow, item.sheetColumn);
      if (item.columnOffset < 4) {
        target.setValue(item.value);
      } else {
        writeTestInvoiceRequestStatus_(target, item.value);
      }
    }
    SpreadsheetApp.flush();

    return { success: true, updated: validated.length };
  } finally {
    try {
      lock.releaseLock();
    } catch (error) {
      console.warn("Could not release Test Invoice Requests lock:", error);
    }
  }
}
