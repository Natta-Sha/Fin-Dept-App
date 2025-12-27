// Data access layer for spreadsheet operations

/**
 * Get project names from the Lists sheet
 * @returns {Array} Array of unique project names
 */
function getProjectNamesFromData() {
  try {
    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.LISTS);
    const values = sheet.getRange("A:A").getValues().flat().filter(String);
    return [...new Set(values)].sort((a, b) => a.localeCompare(b));
  } catch (error) {
    console.error("Error getting project names:", error);
    return [];
  }
}

/**
 * Get project details from the Lists sheet
 * @param {string} projectName - Name of the project
 * @returns {Object} Project details object
 */
function getProjectDetailsFromData(projectName) {
  try {
    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.LISTS);
    const values = sheet.getDataRange().getValues();

    const templateMap = new Map();
    const bankMap = new Map();
    let projectRow = null;

    // Process all rows to build maps and find project
    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      // Find project row
      const name = (row[CONFIG.COLUMNS.PROJECT_NAME] || "").toString().trim();
      if (
        !projectRow &&
        name.toLowerCase() === projectName.toString().trim().toLowerCase()
      ) {
        projectRow = row;
      }

      // Build template map (columns T/U → 19/20)
      const templateName = (row[CONFIG.COLUMNS.TEMPLATE_NAME_COL] || "")
        .toString()
        .trim();
      const templateId = (row[CONFIG.COLUMNS.TEMPLATE_ID_COL] || "")
        .toString()
        .trim();
      if (templateName && templateId) {
        templateMap.set(templateName.toLowerCase(), templateId);
      }

      // Build bank map (columns Q/R → 16/17)
      const short = (row[CONFIG.COLUMNS.BANK_SHORT_COL] || "")
        .toString()
        .trim();
      const full = (row[CONFIG.COLUMNS.BANK_FULL_COL] || "").toString().trim();
      if (short && full) {
        bankMap.set(short, full);
      }
    }

    if (!projectRow) {
      throw new Error(ERROR_MESSAGES.PROJECT_NOT_FOUND(projectName));
    }

    const selectedTemplateName = (
      projectRow[CONFIG.COLUMNS.TEMPLATE_NAME] || ""
    )
      .toString()
      .trim();
    if (!selectedTemplateName) {
      throw new Error(ERROR_MESSAGES.NO_TEMPLATE_NAME(projectName));
    }

    const selectedTemplateId = templateMap.get(
      selectedTemplateName.toLowerCase()
    );
    if (!selectedTemplateId) {
      throw new Error(ERROR_MESSAGES.NO_TEMPLATE_FOUND(selectedTemplateName));
    }

    // Process tax rate
    const tax =
      typeof projectRow[CONFIG.COLUMNS.TAX_RATE] === "number"
        ? projectRow[CONFIG.COLUMNS.TAX_RATE] * 100
        : parseFloat(projectRow[CONFIG.COLUMNS.TAX_RATE]);

    // Get bank details
    const shortBank1 = (projectRow[CONFIG.COLUMNS.BANK_SHORT_1] || "")
      .toString()
      .trim();
    const shortBank2 = (projectRow[CONFIG.COLUMNS.BANK_SHORT_2] || "")
      .toString()
      .trim();

    Logger.log("Selected templateId: " + selectedTemplateId);
    Logger.log(
      "Bank details lookup - shortBank1: '" +
        shortBank1 +
        "', shortBank2: '" +
        shortBank2 +
        "'"
    );
    Logger.log("Bank map size: " + bankMap.size);
    Logger.log(
      "Bank map entries: " +
        Array.from(bankMap.entries())
          .map(([k, v]) => k + "->" + v.substring(0, 50) + "...")
          .join(", ")
    );

    return {
      clientName: projectRow[CONFIG.COLUMNS.CLIENT_NAME] || "",
      clientNumber: `${projectRow[CONFIG.COLUMNS.CLIENT_NUMBER_PART1] || ""} ${
        projectRow[CONFIG.COLUMNS.CLIENT_NUMBER_PART2] || ""
      }`.trim(),
      clientAddress: projectRow[CONFIG.COLUMNS.CLIENT_ADDRESS] || "",
      tax: isNaN(tax) ? 0 : tax.toFixed(0),
      currency:
        CONFIG.CURRENCY_SYMBOLS[projectRow[CONFIG.COLUMNS.CURRENCY]] ||
        projectRow[CONFIG.COLUMNS.CURRENCY],
      paymentDelay: parseInt(projectRow[CONFIG.COLUMNS.PAYMENT_DELAY]) || 0,
      dayType: (projectRow[CONFIG.COLUMNS.DAY_TYPE] || "")
        .toString()
        .trim()
        .toUpperCase(),
      ourCompany: projectRow[CONFIG.COLUMNS.OUR_COMPANY] || "",
      bankDetails1: bankMap.get(shortBank1) || "",
      bankDetails2: bankMap.get(shortBank2) || "",
      templateId: selectedTemplateId,
    };
  } catch (error) {
    console.error("Error getting project details:", error);
    throw error;
  }
}

/**
 * Get invoice list from the Invoices sheet
 * @returns {Array} Array of invoice objects
 */
function getInvoiceListFromData() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("invoiceList");
    if (cached) {
      return JSON.parse(cached);
    }

    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.INVOICES);
    const data = sheet.getDataRange().getValues();

    if (data.length < 2) return [];

    const headers = data[0].map((h) => (h || "").toString().trim());

    const colIndex = {
      id: headers.indexOf("ID"),
      projectName: headers.indexOf("Project Name"),
      invoiceNumber: headers.indexOf("Invoice Number"),
      invoiceDate: headers.indexOf("Invoice Date"),
      dueDate: headers.indexOf("Due Date"),
      total: headers.indexOf("Total"),
      currency: headers.indexOf("Currency"),
    };

    // Validate required columns
    for (let key in colIndex) {
      if (colIndex[key] === -1) {
        throw new Error(ERROR_MESSAGES.MISSING_COLUMN(key));
      }
    }

    const result = data.slice(1).map((row) => ({
      id: row[colIndex.id] || "",
      projectName: row[colIndex.projectName] || "",
      invoiceNumber: row[colIndex.invoiceNumber] || "",
      invoiceDate: formatDate(row[colIndex.invoiceDate]),
      dueDate: formatDate(row[colIndex.dueDate]),
      total:
        row[colIndex.total] !== undefined && row[colIndex.total] !== ""
          ? parseFloat(row[colIndex.total]).toFixed(2)
          : "",
      currency: row[colIndex.currency] || "",
    }));

    cache.put("invoiceList", JSON.stringify(result), 300); // cache for 5 minutes
    return result;
  } catch (error) {
    console.error("Error getting invoice list:", error);
    return [];
  }
}

/**
 * Get credit note list from the Credit Notes sheet
 * @returns {Array} Array of credit note objects
 */
function getCreditNoteListFromData() {
  try {
    Logger.log("=== getCreditNoteListFromData: Starting ===");

    var cache = CacheService.getScriptCache();
    var cached = cache.get("creditNoteList");
    if (cached) {
      Logger.log("getCreditNoteListFromData: Returning cached data");
      return JSON.parse(cached);
    }

    Logger.log(
      "getCreditNoteListFromData: CONFIG.SHEETS.CREDITNOTES = " +
        CONFIG.SHEETS.CREDITNOTES
    );
    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.CREDITNOTES);
    Logger.log("getCreditNoteListFromData: Sheet found: " + sheet.getName());

    const data = sheet.getDataRange().getValues();
    Logger.log("getCreditNoteListFromData: Data rows count: " + data.length);

    if (data.length < 2) {
      Logger.log("getCreditNoteListFromData: No data rows found (less than 2)");
      return [];
    }

    const headers = data[0].map((h) => (h || "").toString().trim());
    Logger.log(
      "getCreditNoteListFromData: Headers found: " + JSON.stringify(headers)
    );

    const colIndex = {
      id: headers.indexOf("ID"),
      projectName: headers.indexOf("Project Name"),
      creditNoteNumber: headers.indexOf("CN Number"),
      creditNoteDate: headers.indexOf("CN Date"),
      total: headers.indexOf("Total"),
      currency: headers.indexOf("Currency"),
    };

    Logger.log(
      "getCreditNoteListFromData: Column indexes: " + JSON.stringify(colIndex)
    );

    // Validate required columns
    for (let key in colIndex) {
      if (colIndex[key] === -1) {
        Logger.log("getCreditNoteListFromData: Missing column: " + key);
        throw new Error(ERROR_MESSAGES.MISSING_COLUMN(key));
      }
    }

    const result = data.slice(1).map((row, index) => {
      const rowData = {
        id: row[colIndex.id] || "",
        projectName: row[colIndex.projectName] || "",
        creditNoteNumber: row[colIndex.creditNoteNumber] || "",
        creditNoteDate: formatDate(row[colIndex.creditNoteDate]),
        total:
          row[colIndex.total] !== undefined && row[colIndex.total] !== ""
            ? parseFloat(row[colIndex.total]).toFixed(2)
            : "",
        currency: row[colIndex.currency] || "",
      };

      if (index < 3) {
        // Log first 3 rows for debugging
        Logger.log(
          "getCreditNoteListFromData: Row " +
            (index + 1) +
            ": " +
            JSON.stringify(rowData)
        );
      }

      return rowData;
    });

    Logger.log(
      "getCreditNoteListFromData: Processed " + result.length + " rows"
    );
    Logger.log(
      "getCreditNoteListFromData: First result: " +
        JSON.stringify(result[0] || {})
    );

    cache.put("creditNoteList", JSON.stringify(result), 300); // cache for 5 minutes
    return result;
  } catch (error) {
    Logger.log("getCreditNoteListFromData: ERROR - " + error.toString());
    Logger.log("getCreditNoteListFromData: Stack trace: " + error.stack);
    console.error("Error getting credit note list:", error);
    return [];
  }
}

/**
 * Get invoice data by ID
 * @param {string} id - Invoice ID
 * @returns {Object} Invoice data object
 */
function getInvoiceDataByIdFromData(id) {
  try {
    // Validate input
    if (!id || id.toString().trim() === "") {
      console.log("Invalid ID provided to getInvoiceDataByIdFromData");
      return {};
    }

    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.INVOICES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const indexMap = headers.reduce((acc, h, i) => {
      acc[h] = i;
      return acc;
    }, {});

    let row = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][indexMap["ID"]] === id) {
        row = data[i];
        break;
      }
    }
    if (!row) {
      console.log(`Invoice with ID ${id} not found.`);
      return {};
    }

    const items = [];
    for (let i = 0; i < CONFIG.INVOICE_TABLE.MAX_ROWS; i++) {
      const base = 21 + i * CONFIG.INVOICE_TABLE.COLUMNS_PER_ROW;
      const item = row.slice(base, base + CONFIG.INVOICE_TABLE.COLUMNS_PER_ROW);
      if (item.some((cell) => cell && cell.toString().trim() !== "")) {
        items.push(item);
      }
    }

    return {
      projectName: row[indexMap["Project Name"]],
      invoiceNumber: row[indexMap["Invoice Number"]],
      clientName: row[indexMap["Client Name"]],
      clientAddress: row[indexMap["Client Address"]],
      clientNumber: row[indexMap["Client Number"]],
      invoiceDate: formatDateForInput(row[indexMap["Invoice Date"]]),
      dueDate: formatDateForInput(row[indexMap["Due Date"]]),
      tax: row[indexMap["Tax Rate (%)"]],
      subtotal: row[indexMap["Subtotal"]],
      total: row[indexMap["Total"]],
      exchangeRate: row[indexMap["Exchange Rate"]],
      currency: row[indexMap["Currency"]],
      amountInEUR: row[indexMap["Amount in EUR"]],
      bankDetails1: row[indexMap["Bank Details 1"]] || "",
      bankDetails2: row[indexMap["Bank Details 2"]] || "",
      ourCompany: row[indexMap["Our Company"]],
      comment: row[indexMap["Comment"]],
      items: items,
    };
  } catch (error) {
    console.error("Error getting invoice data by ID:", error);
    return {}; // ⚠️ тоже вернём пустой объект
  }
}

/**
 * Get credit note data by ID from spreadsheet
 * @param {string} id - Credit note ID
 * @returns {Object} Credit note data
 */
function getCreditNoteDataByIdFromData(id) {
  try {
    // Validate input
    if (!id || id.toString().trim() === "") {
      console.log("Invalid ID provided to getCreditNoteDataByIdFromData");
      return null; // Return null to match what client expects
    }

    console.log("getCreditNoteDataByIdFromData: Looking for ID:", id);
    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.CREDITNOTES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    console.log("getCreditNoteDataByIdFromData: Headers:", headers);
    console.log(
      "getCreditNoteDataByIdFromData: Total rows in sheet:",
      data.length
    );
    console.log(
      "getCreditNoteDataByIdFromData: First few IDs in sheet:",
      data.slice(1, 4).map((row) => row[0])
    );

    const indexMap = headers.reduce((acc, h, i) => {
      acc[h] = i;
      return acc;
    }, {});
    console.log("getCreditNoteDataByIdFromData: Index map:", indexMap);

    let row = null;
    for (let i = 1; i < data.length; i++) {
      const rowId = data[i][indexMap["ID"]];
      console.log(
        `getCreditNoteDataByIdFromData: Checking row ${i}, ID: ${rowId} (type: ${typeof rowId}), looking for: ${id} (type: ${typeof id})`
      );

      // Try both string and number comparison
      if (rowId == id || rowId === id || rowId.toString() === id.toString()) {
        row = data[i];
        console.log("getCreditNoteDataByIdFromData: Found matching row:", row);
        break;
      }
    }
    if (!row) {
      console.log(`Credit note with ID ${id} not found.`);
      console.log(
        "Available IDs in sheet:",
        data.slice(1).map((row, index) => ({ row: index + 1, id: row[0] }))
      );
      return null; // Return null to match what client expects
    }

    const items = [];
    console.log(
      "getCreditNoteDataByIdFromData: CONFIG.CREDIT_NOTE_TABLE.COLUMNS_PER_ROW =",
      CONFIG.CREDIT_NOTE_TABLE.COLUMNS_PER_ROW
    );
    // Base fields: 18 columns (ID through PDF Link)
    // First table row starts at column 19 (Row 1 #, Row 1 Description, Row 1 Period, Row 1 Amount)
    for (let i = 0; i < CONFIG.CREDIT_NOTE_TABLE.MAX_ROWS; i++) {
      const base = 18 + i * 4; // Start from column 19, 4 columns per row
      const item = row.slice(base, base + 4);
      console.log(
        `getCreditNoteDataByIdFromData: Row ${i}, base=${base}, item=`,
        item
      );
      if (item.some((cell) => cell && cell.toString().trim() !== "")) {
        items.push(item);
        console.log(`getCreditNoteDataByIdFromData: Added item ${i}:`, item);
      }
    }

    const result = {
      projectName: row[indexMap["Project Name"]],
      creditNoteNumber: row[indexMap["CN Number"]],
      clientName: row[indexMap["Client Name"]],
      clientAddress: row[indexMap["Client Address"]],
      clientNumber: row[indexMap["Client Number"]],
      creditNoteDate: formatDateForInputFromUtils(row[indexMap["CN Date"]]),
      tax: row[indexMap["Tax Rate (%)"]],
      subtotal: row[indexMap["Subtotal"]],
      total: row[indexMap["Total"]],
      exchangeRate: row[indexMap["Exchange Rate"]],
      currency: row[indexMap["Currency"]],
      amountInEUR: row[indexMap["Amount in EUR"]],
      ourCompany: row[indexMap["Our Company"]],
      comment: row[indexMap["Comment"]],
      items: items,
    };

    console.log("getCreditNoteDataByIdFromData: Returning result:", result);
    return result;
  } catch (error) {
    console.error("Error getting credit note data by ID:", error);
    console.error("Error stack:", error.stack);
    return null; // Return null to match what client expects
  }
}

/**
 * Save invoice data to spreadsheet
 * @param {Object} data - Invoice data to save
 * @returns {Object} Result with doc and PDF URLs
 */
function saveInvoiceData(data) {
  try {
    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.INVOICES);
    const uniqueId = Utilities.getUuid();

    // Parse DD/MM/YYYY date
    const [day, month, year] = data.dueDate.split("/");
    const dueDateObject = new Date(year, month - 1, day);

    const newRow = [
      uniqueId,
      data.projectName,
      data.invoiceNumber,
      data.clientName,
      data.clientAddress,
      data.clientNumber,
      new Date(data.invoiceDate), // YYYY-MM-DD is fine
      dueDateObject,
      data.tax,
      data.subtotal,
      calculateTaxAmountFromUtils(data.subtotal, data.tax),
      calculateTotalAmountFromUtils(
        data.subtotal,
        calculateTaxAmountFromUtils(data.subtotal, data.tax)
      ),
      data.currency === "$" ? data.exchangeRate : "",
      data.currency,
      data.currency === "$" ? data.amountInEUR : "",
      data.ourCompany || "",
      data.comment || "",
      "", // Placeholder for Doc URL
      "", // Placeholder for PDF URL
    ];

    // Process items to force Period field to be saved as text
    const itemCells = [];
    data.items.forEach((row, i) => {
      const newRow = [...row];
      // Force Period field (index 2) to be saved as text to prevent Google Sheets from converting it to date
      if (newRow[2]) {
        newRow[2] = `'${newRow[2].toString()}`; // Add single quote prefix to force text format
      }
      itemCells.push(...newRow);
    });
    const fullRow = newRow.concat(itemCells);

    const newRowIndex = sheet.getLastRow() + 1;
    sheet.getRange(newRowIndex, 1, 1, fullRow.length).setValues([fullRow]);
    CacheService.getScriptCache().remove("invoiceList");

    return { newRowIndex, uniqueId };
  } catch (error) {
    console.error("Error saving invoice data:", error);
    throw error;
  }
}

function processFormFromData(data) {
  try {
    Logger.log("processFormFromData: Starting invoice creation.");
    Logger.log(
      `processFormFromData: Received data for project: ${data.projectName}, invoice: ${data.invoiceNumber}`
    );

    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.INVOICES);
    const uniqueId = Utilities.getUuid();
    Logger.log(`processFormFromData: Generated new unique ID: ${uniqueId}`);

    if (sheet.getLastRow() === 0) {
      const baseHeaders = [
        "ID",
        "Project Name",
        "Invoice Number",
        "Client Name",
        "Client Address",
        "Client Number",
        "Invoice Date",
        "Due Date",
        "Tax Rate (%)",
        "Subtotal",
        "Tax Amount",
        "Total",
        "Exchange Rate",
        "Currency",
        "Amount in EUR",
        "Bank Details 1",
        "Bank Details 2",
        "Our Company",
        "Comment",
        "Google Doc Link",
        "PDF Link",
      ];

      const itemHeaders = [];
      for (let i = 1; i <= CONFIG.INVOICE_TABLE.MAX_ROWS; i++) {
        itemHeaders.push(
          `Row ${i} #`,
          `Row ${i} Service`,
          `Row ${i} Period`,
          `Row ${i} Quantity`,
          `Row ${i} Rate/hour`,
          `Row ${i} Amount`
        );
      }
      sheet.appendRow([...baseHeaders, ...itemHeaders]);
      Logger.log("processFormFromData: Sheet was empty, headers created.");
    }

    const formattedDate = formatDate(data.invoiceDate);

    const [day, month, year] = data.dueDate.split("/");
    const dueDateObject = new Date(year, month - 1, day);
    const formattedDueDate = formatDate(dueDateObject);

    const subtotalNum = parseFloat(data.subtotal) || 0;
    const taxRate = parseFloat(data.tax) || 0;
    const taxAmount = (subtotalNum * taxRate) / 100;
    const totalAmount = subtotalNum + taxAmount;

    const itemCells = [];
    data.items.forEach((row, i) => {
      const newRow = [...row];
      newRow[0] = (i + 1).toString();
      // Force Period field (index 2) to be saved as text to prevent Google Sheets from converting it to date
      if (newRow[2]) {
        newRow[2] = `'${newRow[2].toString()}`; // Add single quote prefix to force text format
      }
      itemCells.push(...newRow);
    });

    const row = [
      uniqueId,
      data.projectName,
      data.invoiceNumber,
      data.clientName,
      data.clientAddress,
      data.clientNumber,
      new Date(data.invoiceDate),
      dueDateObject,
      taxRate.toFixed(0),
      subtotalNum.toFixed(2),
      taxAmount.toFixed(2),
      totalAmount.toFixed(2),
      data.currency === "$" ? parseFloat(data.exchangeRate).toFixed(4) : "",
      data.currency,
      data.currency === "$" ? parseFloat(data.amountInEUR).toFixed(2) : "",
      data.bankDetails1 || "",
      data.bankDetails2 || "",
      data.ourCompany || "",
      data.comment || "",
      "",
      "", // placeholders for doc & pdf
    ].concat(itemCells);

    const newRowIndex = sheet.getLastRow() + 1;
    sheet.getRange(newRowIndex, 1, 1, row.length).setValues([row]);
    Logger.log(
      `processFormFromData: Wrote main data to sheet '${CONFIG.SHEETS.INVOICES}' at row ${newRowIndex}.`
    );

    const folderId = getProjectFolderId(data.projectName);
    Logger.log(">>> Resolved folderId: " + folderId);

    const doc = createInvoiceDoc(
      data,
      formattedDate,
      formattedDueDate,
      subtotalNum,
      taxRate,
      taxAmount,
      totalAmount,
      data.templateId,
      folderId
    );
    if (!doc) {
      Logger.log(
        "processFormFromData: ERROR - createInvoiceDoc returned null or undefined."
      );
      throw new Error(
        "Failed to create the Google Doc. The returned document object was empty."
      );
    }
    Logger.log(
      `processFormFromData: createInvoiceDoc successful. Doc ID: ${doc.getId()}, URL: ${doc.getUrl()}`
    );

    Utilities.sleep(1000);
    Logger.log("processFormFromData: Woke up from 1-second sleep.");

    const pdf = doc.getAs("application/pdf");
    if (!pdf) {
      Logger.log(
        "processFormFromData: ERROR - doc.getAs('application/pdf') returned a null blob."
      );
      throw new Error("Failed to generate PDF content from the document.");
    }
    Logger.log(
      `processFormFromData: Got PDF blob. Name: ${pdf.getName()}, Type: ${pdf.getContentType()}, Size: ${
        pdf.getBytes().length
      } bytes.`
    );

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);

    const cleanCompany = (data.ourCompany || "")
      .replace(/[\\/:*?"<>|]/g, "")
      .trim();
    const cleanClient = (data.clientName || "")
      .replace(/[\\/:*?"<>|]/g, "")
      .trim();
    const filename = `${data.invoiceDate}_Invoice${data.invoiceNumber}_${cleanCompany}-${cleanClient}`;

    const pdfFile = folder.createFile(pdf).setName(`${filename}.pdf`);
    Logger.log(
      `processFormFromData: Created PDF file. ID: ${pdfFile.getId()}, URL: ${pdfFile.getUrl()}`
    );

    sheet.getRange(newRowIndex, 20).setValue(doc.getUrl());
    sheet.getRange(newRowIndex, 21).setValue(pdfFile.getUrl());
    SpreadsheetApp.flush();
    Logger.log(
      `processFormFromData: Wrote Doc and PDF URLs to sheet at row ${newRowIndex}.`
    );

    const result = {
      docUrl: doc.getUrl(),
      pdfUrl: pdfFile.getUrl(),
    };
    Logger.log(
      "processFormFromData: Successfully completed. Returning URLs to client."
    );

    CacheService.getScriptCache().remove("invoiceList");

    return result;
  } catch (e) {
    Logger.log(`processFormFromData: CRITICAL ERROR - ${e.toString()}`);
    Logger.log(`Stack Trace: ${e.stack}`);
    // Re-throw the error so the client-side `.withFailureHandler` can catch it if one is added.
    throw e;
  }
}

// updateSpreadsheetWithUrls is handled in documentService.js

/**
 * Delete invoice by ID from the Invoices sheet
 * @param {string} id - Invoice ID
 * @returns {Object} { success: true } or { success: false, message }
 */
function deleteInvoiceByIdFromData(id) {
  try {
    // Validate input
    if (!id || id.toString().trim() === "") {
      console.log("Invalid ID provided to deleteInvoiceByIdFromData");
      return { success: false, message: "Invalid invoice ID provided" };
    }

    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.INVOICES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const idCol = headers.indexOf("ID");
    const docLinkCol = headers.indexOf("Google Doc Link");
    const pdfLinkCol = headers.indexOf("PDF Link");

    if (idCol === -1) throw new Error("ID column not found.");

    let rowToDelete = -1;
    let docUrl = "";
    let pdfUrl = "";

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === id) {
        rowToDelete = i + 1; // 1-based index
        docUrl = data[i][docLinkCol] || "";
        pdfUrl = data[i][pdfLinkCol] || "";
        break;
      }
    }

    if (rowToDelete === -1) {
      return { success: false, message: "Invoice not found." };
    }

    // 🔹 Удаляем файлы (если есть), логируем ошибки
    let deletedNotes = [];

    if (docUrl && docUrl.trim() !== "") {
      try {
        const docId = extractFileIdFromUrl(docUrl);
        if (docId) {
          DriveApp.getFileById(docId).setTrashed(true);
        }
      } catch (err) {
        const msg = "Google Doc already deleted or not found.";
        Logger.log(msg + " " + err.message);
        deletedNotes.push(msg);
      }
    }

    if (pdfUrl && pdfUrl.trim() !== "") {
      try {
        const pdfId = extractFileIdFromUrl(pdfUrl);
        if (pdfId) {
          DriveApp.getFileById(pdfId).setTrashed(true);
        }
      } catch (err) {
        const msg = "PDF already deleted or not found.";
        Logger.log(msg + " " + err.message);
        deletedNotes.push(msg);
      }
    }

    // 🧹 Удаляем строку
    sheet.deleteRow(rowToDelete);

    // 🧼 Очищаем кэш
    CacheService.getScriptCache().remove("invoiceList");

    // ✅ Возвращаем результат
    return {
      success: true,
      note: deletedNotes.length ? deletedNotes.join(" ") : undefined,
    };
  } catch (error) {
    console.error("Error deleting invoice:", error);
    return { success: false, message: error.message };
  }
}

function extractFileIdFromUrl(url) {
  if (!url || typeof url !== "string") {
    throw new Error("Invalid URL provided: " + url);
  }

  const match = url.match(/[-\w]{25,}/);
  if (!match) {
    throw new Error("Invalid file URL: " + url);
  }
  return match[0];
}

/**
 * Update existing invoice row by ID with new data
 * Replaces the entire row's data while preserving Doc/PDF links
 * @param {string} id - Invoice ID to update
 * @param {Object} data - New invoice data (same shape as creation payload)
 * @returns {Object} { success: true }
 */
function updateInvoiceByIdFromData(id, data) {
  try {
    if (!id || id.toString().trim() === "") {
      return { success: false, message: "Invalid invoice ID provided" };
    }

    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.INVOICES);
    const table = sheet.getDataRange().getValues();
    const headers = table[0];

    const indexMap = headers.reduce((acc, h, i) => {
      acc[h] = i;
      return acc;
    }, {});

    const idCol = indexMap["ID"];
    if (idCol === undefined) throw new Error("ID column not found.");

    let rowIndex = -1; // 0-based within data array (header at 0)
    for (let i = 1; i < table.length; i++) {
      if (table[i][idCol] === id) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1)
      return { success: false, message: "Invoice not found." };

    // Delete old Doc/PDF (best effort)
    const oldDocUrl = table[rowIndex][indexMap["Google Doc Link"]] || "";
    const oldPdfUrl = table[rowIndex][indexMap["PDF Link"]] || "";
    try {
      if (oldDocUrl) {
        const oldDocId = extractFileIdFromUrl(oldDocUrl);
        if (oldDocId) DriveApp.getFileById(oldDocId).setTrashed(true);
      }
    } catch (e) {}
    try {
      if (oldPdfUrl) {
        const oldPdfId = extractFileIdFromUrl(oldPdfUrl);
        if (oldPdfId) DriveApp.getFileById(oldPdfId).setTrashed(true);
      }
    } catch (e) {}

    // Recompute like in creation
    const formattedDate = formatDate(data.invoiceDate);
    const [day, month, year] = (data.dueDate || "01/01/1970").split("/");
    const dueDateObject = new Date(year, month - 1, day);
    const formattedDueDate = formatDate(dueDateObject);

    const subtotalNum = parseFloat(data.subtotal) || 0;
    const taxRate = parseFloat(data.tax) || 0;
    const taxAmount = (subtotalNum * taxRate) / 100;
    const totalAmount = subtotalNum + taxAmount;

    // Resolve template and folder like in creation flow
    const detailsForTemplate = getProjectDetailsFromData(data.projectName);
    const templateId = detailsForTemplate && detailsForTemplate.templateId;
    if (!templateId) {
      throw new Error(ERROR_MESSAGES.NO_TEMPLATE_ID);
    }
    const folderId = getProjectFolderId(data.projectName);

    const doc = createInvoiceDoc(
      data,
      formattedDate,
      formattedDueDate,
      subtotalNum,
      taxRate,
      taxAmount,
      totalAmount,
      templateId,
      folderId
    );
    const pdf = doc.getAs("application/pdf");
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const cleanCompany = (data.ourCompany || "")
      .replace(/[\\/:*?"<>|]/g, "")
      .trim();
    const cleanClient = (data.clientName || "")
      .replace(/[\\/:*?"<>|]/g, "")
      .trim();
    const filename = `${data.invoiceDate}_Invoice${data.invoiceNumber}_${cleanCompany}-${cleanClient}`;
    const pdfFile = folder.createFile(pdf).setName(`${filename}.pdf`);

    // Build row exactly by headers
    const fullRow = new Array(headers.length).fill("");
    fullRow[indexMap["ID"]] = id;
    fullRow[indexMap["Project Name"]] = data.projectName;
    fullRow[indexMap["Invoice Number"]] = data.invoiceNumber;
    fullRow[indexMap["Client Name"]] = data.clientName;
    fullRow[indexMap["Client Address"]] = data.clientAddress;
    fullRow[indexMap["Client Number"]] = data.clientNumber;
    fullRow[indexMap["Invoice Date"]] = new Date(data.invoiceDate);
    fullRow[indexMap["Due Date"]] = dueDateObject;
    fullRow[indexMap["Tax Rate (%)"]] = taxRate.toFixed(0);
    fullRow[indexMap["Subtotal"]] = subtotalNum.toFixed(2);
    fullRow[indexMap["Tax Amount"]] = taxAmount.toFixed(2);
    fullRow[indexMap["Total"]] = totalAmount.toFixed(2);
    fullRow[indexMap["Exchange Rate"]] =
      data.currency === "$"
        ? parseFloat(data.exchangeRate || 0).toFixed(4)
        : "";
    fullRow[indexMap["Currency"]] = data.currency;
    fullRow[indexMap["Amount in EUR"]] =
      data.currency === "$" ? parseFloat(data.amountInEUR || 0).toFixed(2) : "";
    fullRow[indexMap["Bank Details 1"]] = data.bankDetails1 || "";
    fullRow[indexMap["Bank Details 2"]] = data.bankDetails2 || "";
    fullRow[indexMap["Our Company"]] = data.ourCompany || "";
    fullRow[indexMap["Comment"]] = data.comment || "";
    fullRow[indexMap["Google Doc Link"]] = doc.getUrl();
    fullRow[indexMap["PDF Link"]] = pdfFile.getUrl();

    // Items: start from 'Row 1 #' header
    const firstItemIdx = headers.indexOf("Row 1 #");
    const itemsCapacity =
      CONFIG.INVOICE_TABLE.MAX_ROWS * CONFIG.INVOICE_TABLE.COLUMNS_PER_ROW;
    let flatItems = [];
    (data.items || []).forEach((row, i) => {
      const r = [...row];
      r[0] = (i + 1).toString();
      if (r[2]) r[2] = `'${r[2].toString()}`; // force Period as text
      flatItems.push(...r);
    });
    if (flatItems.length > itemsCapacity)
      flatItems = flatItems.slice(0, itemsCapacity);
    while (flatItems.length < itemsCapacity) flatItems.push("");
    if (firstItemIdx !== -1) {
      for (let j = 0; j < itemsCapacity; j++) {
        const target = firstItemIdx + j;
        if (target < headers.length) fullRow[target] = flatItems[j];
      }
    }

    // Write back row (1-based)
    const sheetRow = rowIndex + 1;
    sheet.getRange(sheetRow, 1, 1, fullRow.length).setValues([fullRow]);
    CacheService.getScriptCache().remove("invoiceList");
    return { success: true, docUrl: doc.getUrl(), pdfUrl: pdfFile.getUrl() };
  } catch (error) {
    console.error("Error updating invoice:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Process credit note form data and create documents
 * @param {Object} data - Credit note form data
 * @returns {Object} Result with document and PDF URLs
 */
function processCreditNoteFormFromData(data) {
  try {
    Logger.log("processCreditNoteFormFromData: Starting credit note creation.");
    Logger.log(
      `processCreditNoteFormFromData: Received data for project: ${data.projectName}, credit note: ${data.creditNoteNumber}`
    );

    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.CREDITNOTES);
    const uniqueId = Utilities.getUuid();
    Logger.log(
      `processCreditNoteFormFromData: Generated new unique ID: ${uniqueId}`
    );

    if (sheet.getLastRow() === 0) {
      const baseHeaders = [
        "ID",
        "Project Name",
        "CN Number",
        "Client Name",
        "Client Address",
        "Client Number",
        "CN Date",
        "Tax Rate (%)",
        "Subtotal",
        "Tax Amount",
        "Total",
        "Exchange Rate",
        "Currency",
        "Amount in EUR",
        "Our Company",
        "Comment",
        "Google Doc Link",
        "PDF Link",
      ];

      const itemHeaders = [];
      for (let i = 1; i <= 20; i++) {
        // Max 20 rows as requested
        itemHeaders.push(
          `Row ${i} #`,
          `Row ${i} Description`,
          `Row ${i} Period`,
          `Row ${i} Amount`
        );
      }
      sheet.appendRow([...baseHeaders, ...itemHeaders]);
      Logger.log(
        "processCreditNoteFormFromData: Sheet was empty, headers created."
      );
    }

    const formattedDate = formatDate(data.creditNoteDate);

    const subtotalNum = parseFloat(data.subtotal) || 0;
    const taxRate = parseFloat(data.tax) || 0;
    const taxAmount = (subtotalNum * taxRate) / 100;
    const totalAmount = subtotalNum + taxAmount;

    const itemCells = [];
    data.items.forEach((row, i) => {
      const newRow = [...row];
      newRow[0] = (i + 1).toString();
      // Force Period field (index 2) to be saved as text to prevent Google Sheets from converting it to date
      if (newRow[2]) {
        newRow[2] = `'${newRow[2].toString()}`; // Add single quote prefix to force text format
      }
      itemCells.push(...newRow);
    });

    const row = [
      uniqueId,
      data.projectName,
      data.creditNoteNumber,
      data.clientName,
      data.clientAddress,
      data.clientNumber,
      new Date(data.creditNoteDate),
      taxRate.toFixed(0),
      subtotalNum.toFixed(2),
      taxAmount.toFixed(2),
      totalAmount.toFixed(2),
      data.currency === "$" ? parseFloat(data.exchangeRate).toFixed(4) : "",
      data.currency,
      data.currency === "$" ? parseFloat(data.amountInEUR).toFixed(2) : "",
      data.ourCompany || "",
      data.comment || "",
      "",
      "", // placeholders for doc & pdf
    ].concat(itemCells);

    const newRowIndex = sheet.getLastRow() + 1;
    sheet.getRange(newRowIndex, 1, 1, row.length).setValues([row]);
    Logger.log(
      `processCreditNoteFormFromData: Data saved to row ${newRowIndex}`
    );

    // Create Google Doc and PDF
    const folderId = getProjectFolderId(data.projectName);
    Logger.log(">>> Resolved folderId: " + folderId);

    // Hardcoded template ID as requested
    const templateId = "1yCKAx3nyIz-L_u3FPSK1zMof5Mo0m2-gsNzl1cCQsuQ";

    const doc = createCreditNoteDoc(
      data,
      formattedDate,
      subtotalNum,
      taxRate,
      taxAmount,
      totalAmount,
      templateId,
      folderId
    );
    if (!doc) {
      Logger.log(
        "processCreditNoteFormFromData: ERROR - createCreditNoteDoc returned null or undefined."
      );
      throw new Error(
        "Failed to create the Google Doc. The returned document object was empty."
      );
    }
    Logger.log(
      `processCreditNoteFormFromData: createCreditNoteDoc successful. Doc ID: ${doc.getId()}, URL: ${doc.getUrl()}`
    );

    Utilities.sleep(1000);
    Logger.log("processCreditNoteFormFromData: Woke up from 1-second sleep.");

    const pdf = doc.getAs("application/pdf");
    if (!pdf) {
      Logger.log(
        "processCreditNoteFormFromData: ERROR - doc.getAs('application/pdf') returned a null blob."
      );
      throw new Error("Failed to generate PDF content from the document.");
    }
    Logger.log(
      `processCreditNoteFormFromData: Got PDF blob. Name: ${pdf.getName()}, Type: ${pdf.getContentType()}, Size: ${
        pdf.getBytes().length
      } bytes.`
    );

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);

    const cleanCompany = (data.ourCompany || "")
      .replace(/[\\/:*?"<>|]/g, "")
      .trim();
    const cleanClient = (data.clientName || "")
      .replace(/[\\/:*?"<>|]/g, "")
      .trim();
    const filename = `${data.creditNoteDate}_CreditNote${data.creditNoteNumber}_${cleanCompany}-${cleanClient}`;

    const pdfFile = folder.createFile(pdf).setName(`${filename}.pdf`);
    Logger.log(
      `processCreditNoteFormFromData: Created PDF file. ID: ${pdfFile.getId()}, URL: ${pdfFile.getUrl()}`
    );

    // Update the row with Doc and PDF URLs (columns 17 and 18)
    sheet.getRange(newRowIndex, 17).setValue(doc.getUrl());
    sheet.getRange(newRowIndex, 18).setValue(pdfFile.getUrl());
    SpreadsheetApp.flush();
    Logger.log(
      `processCreditNoteFormFromData: Wrote Doc and PDF URLs to sheet at row ${newRowIndex}.`
    );

    const result = {
      docUrl: doc.getUrl(),
      pdfUrl: pdfFile.getUrl(),
    };
    Logger.log(
      "processCreditNoteFormFromData: Successfully completed. Returning URLs to client."
    );

    CacheService.getScriptCache().remove("creditNoteList");

    return result;
  } catch (error) {
    Logger.log(`processCreditNoteFormFromData: ERROR - ${error.toString()}`);
    Logger.log(`Stack Trace: ${error.stack}`);
    throw error;
  }
}

/**
 * Delete credit note by ID from the Credit Notes sheet
 * @param {string} id - Credit Note ID
 * @returns {Object} { success: true } or { success: false, message }
 */
function deleteCreditNoteByIdFromData(id) {
  try {
    // Validate input
    if (!id || id.toString().trim() === "") {
      console.log("Invalid ID provided to deleteCreditNoteByIdFromData");
      return { success: false, message: "Invalid credit note ID provided" };
    }

    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.CREDITNOTES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const idCol = headers.indexOf("ID");
    const docLinkCol = headers.indexOf("Google Doc Link");
    const pdfLinkCol = headers.indexOf("PDF Link");

    if (idCol === -1) throw new Error("ID column not found.");

    let rowToDelete = -1;
    let docUrl = "";
    let pdfUrl = "";

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === id) {
        rowToDelete = i + 1; // 1-based index
        docUrl = data[i][docLinkCol] || "";
        pdfUrl = data[i][pdfLinkCol] || "";
        break;
      }
    }

    if (rowToDelete === -1) {
      return { success: false, message: "Credit note not found." };
    }

    // 🔹 Удаляем файлы (если есть), логируем ошибки
    let deletedNotes = [];

    if (docUrl && docUrl.trim() !== "") {
      try {
        const docId = extractFileIdFromUrl(docUrl);
        if (docId) {
          DriveApp.getFileById(docId).setTrashed(true);
        }
      } catch (err) {
        const msg = "Google Doc already deleted or not found.";
        Logger.log(msg + " " + err.message);
        deletedNotes.push(msg);
      }
    }

    if (pdfUrl && pdfUrl.trim() !== "") {
      try {
        const pdfId = extractFileIdFromUrl(pdfUrl);
        if (pdfId) {
          DriveApp.getFileById(pdfId).setTrashed(true);
        }
      } catch (err) {
        const msg = "PDF already deleted or not found.";
        Logger.log(msg + " " + err.message);
        deletedNotes.push(msg);
      }
    }

    // 🧹 Удаляем строку
    sheet.deleteRow(rowToDelete);

    // 🧼 Очищаем кэш
    CacheService.getScriptCache().remove("creditNoteList");

    // ✅ Возвращаем результат
    return {
      success: true,
      note: deletedNotes.length ? deletedNotes.join(" ") : undefined,
    };
  } catch (error) {
    console.error("Error deleting credit note:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Update existing credit note by ID with new data
 * @param {Object} data - Credit note data with id
 * @returns {Object} { success: true } or { success: false, message }
 */
function updateCreditNoteByIdFromData(data) {
  try {
    if (!data.id) {
      return { success: false, message: "Credit note ID is required" };
    }

    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.CREDITNOTES);
    const sheetData = sheet.getDataRange().getValues();
    const headers = sheetData[0];

    const idCol = headers.indexOf("ID");
    if (idCol === -1) throw new Error("ID column not found");

    let rowToUpdate = -1;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][idCol] === data.id) {
        rowToUpdate = i + 1; // 1-based index
        break;
      }
    }

    if (rowToUpdate === -1) {
      return { success: false, message: "Credit note not found" };
    }

    // Prepare row data
    const row = new Array(headers.length);

    // Map data to columns
    const colMap = {
      "Project Name": data.projectName || "",
      "CN Number": data.creditNoteNumber || "",
      "Client Name": data.clientName || "",
      "Client Address": data.clientAddress || "",
      "Client Number": data.clientNumber || "",
      "CN Date": data.creditNoteDate ? new Date(data.creditNoteDate) : "",
      "Tax Rate (%)": data.tax || "0",
      Subtotal: data.subtotal || "0",
      Total: data.total || "0",
      "Exchange Rate": data.exchangeRate || "1.0000",
      Currency: data.currency || "$",
      "Amount in EUR": data.amountInEUR || "",
      "Our Company": data.ourCompany || "",
      Comment: data.comment || "",
    };

    // Fill row with mapped data
    headers.forEach((header, index) => {
      if (colMap.hasOwnProperty(header)) {
        row[index] = colMap[header];
      } else if (header === "ID") {
        row[index] = data.id;
      } else {
        row[index] = sheetData[rowToUpdate - 1][index]; // Keep existing value
      }
    });

    // Add items data
    if (data.items && data.items.length > 0) {
      const baseCol = 18; // Start from column 19 (index 18)
      data.items.forEach((item, i) => {
        const startCol = baseCol + i * 4;
        if (startCol + 3 < headers.length) {
          row[startCol] = item[0] || ""; // #
          row[startCol + 1] = item[1] || ""; // Description
          row[startCol + 2] = item[2] || ""; // Period
          row[startCol + 3] = item[3] || ""; // Amount
        }
      });
    }

    // Delete old Doc/PDF (best effort)
    const oldDocUrl =
      sheetData[rowToUpdate - 1][headers.indexOf("Google Doc Link")] || "";
    const oldPdfUrl =
      sheetData[rowToUpdate - 1][headers.indexOf("PDF Link")] || "";
    try {
      if (oldDocUrl) {
        const oldDocId = extractFileIdFromUrl(oldDocUrl);
        if (oldDocId) DriveApp.getFileById(oldDocId).setTrashed(true);
      }
    } catch (e) {}
    try {
      if (oldPdfUrl) {
        const oldPdfId = extractFileIdFromUrl(oldPdfUrl);
        if (oldPdfId) DriveApp.getFileById(oldPdfId).setTrashed(true);
      }
    } catch (e) {}

    // Generate new documents
    const formattedDate = formatDate(data.creditNoteDate);
    const subtotalNum = parseFloat(data.subtotal || 0);
    const taxRate = parseFloat(data.tax || 0);
    const taxAmount = calculateTaxAmountFromUtils(subtotalNum, taxRate);
    const totalAmount = calculateTotalAmountFromUtils(subtotalNum, taxAmount);

    // Resolve template and folder like in creation flow
    const detailsForTemplate = getProjectDetailsFromData(data.projectName);
    // Hardcoded template ID for credit notes as in creation flow
    const templateId = "1yCKAx3nyIz-L_u3FPSK1zMof5Mo0m2-gsNzl1cCQsuQ";
    const folderId = getProjectFolderId(data.projectName);

    // Create updated data object with project details for template filling
    const updatedData = {
      ...data,
      clientName: detailsForTemplate.clientName || data.clientName,
      clientAddress: detailsForTemplate.clientAddress || data.clientAddress,
      clientNumber: detailsForTemplate.clientNumber || data.clientNumber,
      ourCompany: detailsForTemplate.ourCompany || data.ourCompany,
      tax: detailsForTemplate.tax || data.tax,
      currency: detailsForTemplate.currency || data.currency,
    };

    const doc = createCreditNoteDoc(
      updatedData,
      formattedDate,
      subtotalNum,
      taxRate,
      taxAmount,
      totalAmount,
      templateId,
      folderId
    );
    const pdf = doc.getAs("application/pdf");
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const cleanCompany = (updatedData.ourCompany || "")
      .replace(/[\\/:*?"<>|]/g, "")
      .trim();
    const cleanClient = (updatedData.clientName || "")
      .replace(/[\\/:*?"<>|]/g, "")
      .trim();
    const filename = `${data.creditNoteDate}_CreditNote${data.creditNoteNumber}_${cleanCompany}-${cleanClient}`;
    const pdfFile = folder.createFile(pdf).setName(`${filename}.pdf`);

    // Update the row with new document URLs
    row[headers.indexOf("Google Doc Link")] = doc.getUrl();
    row[headers.indexOf("PDF Link")] = pdfFile.getUrl();

    // Update the row
    sheet.getRange(rowToUpdate, 1, 1, headers.length).setValues([row]);

    // Clear cache
    CacheService.getScriptCache().remove("creditNoteList");

    return { success: true, docUrl: doc.getUrl(), pdfUrl: pdfFile.getUrl() };
  } catch (error) {
    console.error("Error updating credit note:", error);
    return { success: false, message: error.message };
  }
}

// ============================================
// CONTRACTS FUNCTIONS
// ============================================

/**
 * Get dropdown options for contract form from Lists sheet
 * @returns {Object} Object with arrays for each dropdown
 */
function getContractDropdownOptionsFromData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.CONTRACTORS_SPREADSHEET_ID);
    const listsSheet = spreadsheet.getSheetByName("Lists");
    
    if (!listsSheet) {
      console.log("Lists sheet not found in contractors spreadsheet");
      return {
        cooperationTypes: [],
        ourCompanies: [],
        serviceTypes: [],
        peOptions: [],
        accountTypes: [],
        currencies: [],
        documentTypes: [],
      };
    }
    
    const data = listsSheet.getDataRange().getValues();
    
    // Extract unique values from each column (skip header row)
    const cooperationTypes = []; // Column A
    const ourCompanies = [];     // Column B
    const serviceTypes = [];     // Column C
    const peOptions = [];        // Column D
    const accountTypes = [];     // Column E
    const currencies = [];       // Column F
    const documentTypes = [];    // Column G
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[0] && row[0].toString().trim() !== "") {
        const val = row[0].toString().trim();
        if (!cooperationTypes.includes(val)) cooperationTypes.push(val);
      }
      if (row[1] && row[1].toString().trim() !== "") {
        const val = row[1].toString().trim();
        if (!ourCompanies.includes(val)) ourCompanies.push(val);
      }
      if (row[2] && row[2].toString().trim() !== "") {
        const val = row[2].toString().trim();
        if (!serviceTypes.includes(val)) serviceTypes.push(val);
      }
      if (row[3] && row[3].toString().trim() !== "") {
        const val = row[3].toString().trim();
        if (!peOptions.includes(val)) peOptions.push(val);
      }
      if (row[4] && row[4].toString().trim() !== "") {
        const val = row[4].toString().trim();
        if (!accountTypes.includes(val)) accountTypes.push(val);
      }
      if (row[5] && row[5].toString().trim() !== "") {
        const val = row[5].toString().trim();
        if (!currencies.includes(val)) currencies.push(val);
      }
      if (row[6] && row[6].toString().trim() !== "") {
        const val = row[6].toString().trim();
        if (!documentTypes.includes(val)) documentTypes.push(val);
      }
    }
    
    return {
      cooperationTypes: cooperationTypes.sort(),
      ourCompanies: ourCompanies.sort(),
      serviceTypes: serviceTypes.sort(),
      peOptions: peOptions.sort(),
      accountTypes: accountTypes.sort(),
      currencies: currencies.sort(),
      documentTypes: documentTypes.sort(),
    };
  } catch (error) {
    console.error("Error getting contract dropdown options:", error);
    return {
      cooperationTypes: [],
      ourCompanies: [],
      serviceTypes: [],
      peOptions: [],
      accountTypes: [],
      currencies: [],
      documentTypes: [],
    };
  }
}

/**
 * Get contract templates from Templates sheet
 * Filters by cooperation type, our company, and service type
 * @param {string} cooperationType
 * @param {string} ourCompany
 * @param {string} serviceType
 * @returns {Array} Array of template objects {name, link}
 */
function getContractTemplatesFromData(cooperationType, ourCompany, serviceType, documentType) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.CONTRACTORS_SPREADSHEET_ID);
    const templatesSheet = spreadsheet.getSheetByName("Templates");
    
    if (!templatesSheet) {
      console.log("Templates sheet not found");
      return [];
    }
    
    const data = templatesSheet.getDataRange().getValues();
    const templates = [];
    
    // Columns: A = Type of cooperation, B = Our company, C = Type of services, D = Document type, E = Template link
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCoopType = (row[0] || "").toString().trim();
      const rowCompany = (row[1] || "").toString().trim();
      const rowServiceType = (row[2] || "").toString().trim();
      const rowDocType = (row[3] || "").toString().trim();
      const templateLink = (row[4] || "").toString().trim();
      
      // Filter by matching criteria (if provided)
      const matchCoop = !cooperationType || rowCoopType === cooperationType;
      const matchCompany = !ourCompany || rowCompany === ourCompany;
      const matchService = !serviceType || rowServiceType === serviceType;
      const matchDocType = !documentType || rowDocType === documentType;
      
      if (matchCoop && matchCompany && matchService && matchDocType && templateLink) {
        templates.push({
          cooperationType: rowCoopType,
          ourCompany: rowCompany,
          serviceType: rowServiceType,
          documentType: rowDocType,
          link: templateLink,
          name: `${rowCoopType} - ${rowCompany} - ${rowServiceType} - ${rowDocType}`,
        });
      }
    }
    
    return templates;
  } catch (error) {
    console.error("Error getting contract templates:", error);
    return [];
  }
}

/**
 * Get contract list from the Contracts sheet
 * @returns {Array} Array of contract objects
 */
function getContractListFromData() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get("contractList");
    if (cached) {
      return JSON.parse(cached);
    }

    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.CONTRACTS);
    const data = sheet.getDataRange().getValues();

    if (data.length < 2) return [];

    const headers = data[0].map((h) => (h || "").toString().trim());

    const colIndex = {
      id: headers.indexOf("ID"),
      folderLink: headers.indexOf("Folder Link"),
      contractorName: headers.indexOf("Contractor Name"),
      ourCompany: headers.indexOf("Our Company"),
      serviceType: headers.indexOf("Service Type"),
      cooperationType: headers.indexOf("Cooperation Type"),
      contractNumber: headers.indexOf("Contract Number"),
      contractDate: headers.indexOf("Contract Date"),
      templateLink: headers.indexOf("Template Link"),
    };

    // Validate required columns - only check essential ones
    const requiredCols = ["id", "contractorName", "contractNumber"];
    for (let key of requiredCols) {
      if (colIndex[key] === -1) {
        console.log("Missing column in Contracts sheet: " + key);
        return [];
      }
    }

    const result = data.slice(1).map((row) => ({
      id: row[colIndex.id] || "",
      folderLink: colIndex.folderLink !== -1 ? row[colIndex.folderLink] || "" : "",
      contractorName: row[colIndex.contractorName] || "",
      ourCompany: colIndex.ourCompany !== -1 ? row[colIndex.ourCompany] || "" : "",
      serviceType: colIndex.serviceType !== -1 ? row[colIndex.serviceType] || "" : "",
      cooperationType: colIndex.cooperationType !== -1 ? row[colIndex.cooperationType] || "" : "",
      contractNumber: row[colIndex.contractNumber] || "",
      contractDate: colIndex.contractDate !== -1 ? formatDate(row[colIndex.contractDate]) : "",
      templateLink: colIndex.templateLink !== -1 ? row[colIndex.templateLink] || "" : "",
    }));

    cache.put("contractList", JSON.stringify(result), 300);
    return result;
  } catch (error) {
    console.error("Error getting contract list:", error);
    return [];
  }
}

/**
 * Get contract data by ID from spreadsheet
 * @param {string} id - Contract ID
 * @returns {Object} Contract data
 */
function getContractDataByIdFromData(id) {
  try {
    if (!id || id.toString().trim() === "") {
      console.log("Invalid ID provided to getContractDataByIdFromData");
      return null;
    }

    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = getSheet(spreadsheet, CONFIG.SHEETS.CONTRACTS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const indexMap = headers.reduce((acc, h, i) => {
      acc[h] = i;
      return acc;
    }, {});

    let row = null;
    for (let i = 1; i < data.length; i++) {
      const rowId = data[i][indexMap["ID"]];
      if (rowId == id || rowId === id || rowId.toString() === id.toString()) {
        row = data[i];
        break;
      }
    }

    if (!row) {
      console.log(`Contract with ID ${id} not found.`);
      return null;
    }

    return {
      id: row[indexMap["ID"]] || "",
      folderLink: row[indexMap["Folder Link"]] || "",
      contractorName: row[indexMap["Contractor Name"]] || "",
      ourCompany: row[indexMap["Our Company"]] || "",
      serviceType: row[indexMap["Service Type"]] || "",
      cooperationType: row[indexMap["Cooperation Type"]] || "",
      contractNumber: row[indexMap["Contract Number"]] || "",
      contractDate: formatDateForInputFromUtils(row[indexMap["Contract Date"]]),
      templateLink: row[indexMap["Template Link"]] || "",
    };
  } catch (error) {
    console.error("Error getting contract data by ID:", error);
    return null;
  }
}

/**
 * Field mapping: form field ID -> sheet column name
 */
const CONTRACT_FIELD_MAPPING = {
  folderLink: "Ссылка на папку с дого",
  contractorName: "Название контрактора",
  pe: "ФОП",
  ourCompany: "Наша компания",
  serviceType: "Вид услуг",
  cooperationType: "Вид сотрудничества",
  documentType: "Вид документа",
  contractNumber: "№ договора",
  contractDate: "Дата договора",
  probationPeriod: "Срок ИС",
  terminationDate: "Дата окончания договора",
  registrationNumber: "№ гос.регистрации",
  registrationDate: "Дата гос.регистрации",
  contractorId: "Номер контрактора",
  contractorVatId: "Номер НДС контрактора",
  contractorJurisdiction: "Юрисдикция контрактора",
  contractorAddress: "Адрес контрактора",
  bankAccountUAH: "Счет (грн)",
  bankAccountUSD: "Счет (долл)",
  bankAccountEUR: "Счет (евро)",
  bankName: "Банк",
  accountType: "Тип счета",
  bankCode: "Код банка",
  contractorEmail: "Эл.почта",
  contractorRole: "Роль контрактора",
  contractorRate: "Рейт контрактора",
  currencyOfRate: "Валюта рейта",
  attachmentNumber: "Номер приложения",
  sowStartDateRequired: "Дата старта термин",
  sowStartDate: "Дата старта",
  templateLink: "Шаблон договора",
};

/**
 * Placeholder mapping: form field ID -> placeholder in template
 */
const CONTRACT_PLACEHOLDER_MAPPING = {
  contractorName: "{Название контрактора}",
  pe: "{ФОП}",
  ourCompany: "{Наша компания}",
  serviceType: "{Вид услуг}",
  cooperationType: "{Вид сотрудничества}",
  contractNumber: "{№ договора}",
  contractDate: "{Дата договора}",
  probationPeriod: "{Срок ИС}",
  terminationDate: "{Дата окончания договора}",
  registrationNumber: "{№ гос.регистрации}",
  registrationDate: "{Дата гос.регистрации}",
  contractorId: "{Номер контрактора}",
  contractorVatId: "{Номер НДС контрактора}",
  contractorJurisdiction: "{Юрисдикция контрактора}",
  contractorAddress: "{Адрес контрактора}",
  bankAccountUAH: "{Счет грн}",
  bankAccountUSD: "{Счет долл}",
  bankAccountEUR: "{Счет евро}",
  bankName: "{Банк}",
  accountType: "{Тип счета}",
  bankCode: "{Код банка}",
  contractorEmail: "{Эл.почта}",
  contractorRole: "{Роль контрактора}",
  contractorRate: "{Рейт контрактора}",
  currencyOfRate: "{Валюта рейта}",
  attachmentNumber: "{Номер приложения}",
  sowStartDateRequired: "{Дата старта термин}",
  sowStartDate: "{Дата старта}",
};

/**
 * Extract folder ID from Google Drive folder URL
 * @param {string} folderUrl - Google Drive folder URL
 * @returns {string|null} Folder ID or null
 */
function extractFolderIdFromUrl(folderUrl) {
  if (!folderUrl) return null;
  
  // Match patterns like:
  // https://drive.google.com/drive/folders/FOLDER_ID
  // https://drive.google.com/drive/u/0/folders/FOLDER_ID
  const match = folderUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Extract document ID from Google Docs URL
 * @param {string} docUrl - Google Docs URL
 * @returns {string|null} Document ID or null
 */
function extractDocIdFromUrl(docUrl) {
  if (!docUrl) return null;
  
  // Match patterns like:
  // https://docs.google.com/document/d/DOC_ID/edit
  // https://docs.google.com/document/d/DOC_ID/preview
  const match = docUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Generate document name based on document type
 * @param {Object} formData - Form data
 * @returns {string} Document name
 */
function generateContractDocumentName(formData) {
  const documentType = (formData.documentType || "").toLowerCase();
  const ourCompany = formData.ourCompany || "Company";
  const contractorName = formData.contractorName || "Contractor";
  
  if (documentType.includes("attachment") || documentType.includes("addendum")) {
    // Addendum_[attachmentNumber]_[sowStartDate]_[ourCompany]-[contractorName]
    const attachmentNumber = formData.attachmentNumber || "1";
    const startDate = formData.sowStartDate || "";
    return "Addendum_" + attachmentNumber + "_" + startDate + "_" + ourCompany + "-" + contractorName;
  } else {
    // Contract_[contractNumber]_[contractDate]_[ourCompany]-[contractorName]
    const contractNumber = formData.contractNumber || "";
    const contractDate = formData.contractDate || "";
    return "Contract_" + contractNumber + "_" + contractDate + "_" + ourCompany + "-" + contractorName;
  }
}

/**
 * Copy template to folder and replace placeholders
 * @param {string} templateUrl - URL of the template document
 * @param {string} folderUrl - URL of the destination folder
 * @param {Object} formData - Form data for replacements
 * @returns {Object} Result with document URL or error
 */
function createContractDocument(templateUrl, folderUrl, formData) {
  try {
    // Extract IDs from URLs
    const templateId = extractDocIdFromUrl(templateUrl);
    const folderId = extractFolderIdFromUrl(folderUrl);
    
    if (!templateId) {
      throw new Error("Invalid template URL");
    }
    
    if (!folderId) {
      throw new Error("Invalid folder URL");
    }
    
    // Get template and destination folder
    const templateFile = DriveApp.getFileById(templateId);
    const destFolder = DriveApp.getFolderById(folderId);
    
    // Generate document name
    const docName = generateContractDocumentName(formData);
    
    // Copy template to destination folder
    const copiedFile = templateFile.makeCopy(docName, destFolder);
    const copiedDocId = copiedFile.getId();
    
    // Open the copied document and replace placeholders
    const doc = DocumentApp.openById(copiedDocId);
    const body = doc.getBody();
    
    // Replace all placeholders
    Object.keys(CONTRACT_PLACEHOLDER_MAPPING).forEach(function(fieldId) {
      const placeholder = CONTRACT_PLACEHOLDER_MAPPING[fieldId];
      const value = formData[fieldId] || "";
      body.replaceText(placeholder.replace(/[{}]/g, "\\$&"), value);
    });
    
    // Save and close
    doc.saveAndClose();
    
    // Get the URL of the created document
    const docUrl = "https://docs.google.com/document/d/" + copiedDocId + "/edit";
    
    console.log("Contract document created:", docUrl);
    
    return {
      success: true,
      documentUrl: docUrl,
      documentId: copiedDocId
    };
    
  } catch (error) {
    console.error("Error creating contract document:", error);
    return {
      success: false,
      documentUrl: null,
      error: error.toString()
    };
  }
}

/**
 * Save a new contract to the Contracts sheet
 * @param {Object} formData - Object with form field values
 * @returns {Object} Result with success status and contract ID
 */
function saveContractToData(formData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.CONTRACTORS_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("Contracts");
    
    if (!sheet) {
      throw new Error("Contracts sheet not found");
    }
    
    // Get headers from first row
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Build column index map: column name -> column index (0-based)
    const columnMap = {};
    headers.forEach((header, index) => {
      if (header) {
        columnMap[header.toString().trim()] = index;
      }
    });
    
    // Generate unique ID
    const contractId = Utilities.getUuid();
    
    // Step 1: Create the contract document from template
    let documentUrl = "";
    if (formData.templateLink && formData.folderLink) {
      const docResult = createContractDocument(
        formData.templateLink,
        formData.folderLink,
        formData
      );
      
      if (docResult.success) {
        documentUrl = docResult.documentUrl;
      } else {
        console.warn("Could not create document:", docResult.error);
        // Continue saving data even if document creation fails
      }
    }
    
    // Step 2: Prepare row data array (fill with empty strings)
    const rowData = new Array(headers.length).fill("");
    
    // Column A (index 0) = ID
    rowData[0] = contractId;
    
    // Column B (index 1) = Document link
    rowData[1] = documentUrl;
    
    // Map form data to columns using the mapping
    Object.keys(formData).forEach(function(fieldId) {
      const columnName = CONTRACT_FIELD_MAPPING[fieldId];
      if (columnName && columnMap.hasOwnProperty(columnName)) {
        const colIndex = columnMap[columnName];
        rowData[colIndex] = formData[fieldId] || "";
      }
    });
    
    // Append the row to the sheet
    sheet.appendRow(rowData);
    
    console.log("Contract saved successfully with ID:", contractId);
    
    return {
      success: true,
      id: contractId,
      documentUrl: documentUrl,
      message: documentUrl 
        ? "Contract saved and document created successfully" 
        : "Contract saved (document creation skipped)"
    };
    
  } catch (error) {
    console.error("Error saving contract:", error);
    return {
      success: false,
      id: null,
      message: "Error saving contract: " + error.toString()
    };
  }
}
