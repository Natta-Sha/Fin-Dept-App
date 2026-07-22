// Isolated data service for the Tabulator invoice-requests experiment.
// Removing this file and its page/navigation hooks removes the feature.

var TEST_INVOICE_REQUESTS_SPREADSHEET_ID =
  "1-5b98hkm-wsrTlyt-1j2PJcTvu1VNyAh0EjRhT3ZBg0";
var TEST_INVOICE_REQUESTS_SHEET_NAME = "анкета";
var TEST_INVOICE_REQUESTS_FIRST_COLUMN = 2; // B
var TEST_INVOICE_REQUESTS_COLUMN_COUNT = 12; // B:M

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

function getTestInvoiceRequests() {
  assertTestInvoiceRequestsAccess_();

  var spreadsheet = SpreadsheetApp.openById(
    TEST_INVOICE_REQUESTS_SPREADSHEET_ID
  );
  var sheet = spreadsheet.getSheetByName(TEST_INVOICE_REQUESTS_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "анкета" was not found.');

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
      cells.push({
        value: serializeTestInvoiceRequestValue_(
          values[rowIndex][sourceColumn],
          displayValues[rowIndex][sourceColumn]
        ),
        originalToken: comparableTestInvoiceRequestValue_(
          values[rowIndex][sourceColumn]
        ),
        displayValue: displayValues[rowIndex][sourceColumn] || "",
        link: richText ? richText.getLinkUrl() || "" : "",
      });
    }

    rows.push({ id: rowId, cells: cells });
  }

  return { headers: headers, rows: rows };
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
    if (!sheet) throw new Error('Sheet "анкета" was not found.');

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
        comparableTestInvoiceRequestValue_(targetRow.values[sourceColumn]) !==
        change.originalToken
      ) {
        conflicts.push({ id: rowId, columnOffset: columnOffset });
        continue;
      }

      var nextValue =
        columnOffset >= 4 ? change.value === true : String(change.value || "");
      validated.push({
        sheetRow: targetRow.sheetRow,
        sheetColumn: TEST_INVOICE_REQUESTS_FIRST_COLUMN + columnOffset,
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

    // Send adjacent cells as one Sheets write while leaving unrelated cells
    // untouched. The browser still makes only one server request per save.
    validated.sort(function (left, right) {
      return (
        left.sheetRow - right.sheetRow ||
        left.sheetColumn - right.sheetColumn
      );
    });
    var writeIndex = 0;
    while (writeIndex < validated.length) {
      var first = validated[writeIndex];
      var segmentValues = [first.value];
      var nextIndex = writeIndex + 1;
      while (
        nextIndex < validated.length &&
        validated[nextIndex].sheetRow === first.sheetRow &&
        validated[nextIndex].sheetColumn ===
          first.sheetColumn + segmentValues.length
      ) {
        segmentValues.push(validated[nextIndex].value);
        nextIndex++;
      }
      sheet
        .getRange(
          first.sheetRow,
          first.sheetColumn,
          1,
          segmentValues.length
        )
        .setValues([segmentValues]);
      writeIndex = nextIndex;
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
