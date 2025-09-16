// Document service for Google Docs and PDF operations

/**
 * Create invoice document from template
 * @param {Object} data - Invoice data
 * @param {string} formattedDate - Formatted invoice date
 * @param {string} formattedDueDate - Formatted due date
 * @param {number} subtotal - Subtotal amount
 * @param {number} taxRate - Tax rate
 * @param {number} taxAmount - Tax amount
 * @param {number} totalAmount - Total amount
 * @param {string} templateId - Template document ID
 * @returns {Document} Google Document object
 */
function createInvoiceDoc(
  data,
  formattedDate,
  formattedDueDate,
  subtotal,
  taxRate,
  taxAmount,
  totalAmount,
  templateId,
  folderId
) {
  Logger.log(`createInvoiceDoc: Starting for template ID: ${templateId}`);
  if (!templateId) {
    Logger.log("createInvoiceDoc: ERROR - No templateId provided.");
    throw new Error(ERROR_MESSAGES.NO_TEMPLATE_ID);
  }

  try {
    const template = DriveApp.getFileById(templateId);
    Logger.log(">>> Using folderId: " + folderId);

    const folder = DriveApp.getFolderById(folderId);

    const filename = generateInvoiceFilenameFromUtils(data);
    Logger.log(`createInvoiceDoc: Generated filename: ${filename}`);

    const copy = template.makeCopy(filename, folder);
    Logger.log(`createInvoiceDoc: Created copy with ID: ${copy.getId()}`);

    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    // Handle exchange rate section
    handleExchangeRateSection(body, data);

    // Update invoice table
    updateInvoiceTable(body, data);

    // Replace placeholders
    replaceDocumentPlaceholders(
      body,
      data,
      formattedDate,
      formattedDueDate,
      taxRate,
      taxAmount,
      totalAmount
    );

    Logger.log(
      `createInvoiceDoc: Placeholders replaced. Saving and closing doc.`
    );
    doc.saveAndClose();
    Logger.log(
      `createInvoiceDoc: Document saved and closed. Returning doc object.`
    );
    return doc;
  } catch (error) {
    Logger.log(`createInvoiceDoc: CRITICAL ERROR - ${error.toString()}`);
    Logger.log(`Stack Trace: ${error.stack}`);
    console.error("Error creating invoice document:", error);
    throw error;
  }
}

/**
 * Handle exchange rate section in document
 * @param {Body} body - Document body
 * @param {Object} data - Invoice data
 */
function handleExchangeRateSection(body, data) {
  if (data.currency !== "$") {
    // Remove exchange rate notice for non-USD currencies
    const paragraphs = body.getParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
      const text = paragraphs[i].getText();
      if (text.includes("Exchange Rate Notice")) {
        // Check if we can safely remove paragraphs
        if (paragraphs.length > 1) {
          paragraphs[i].removeFromParent();
          if (i < paragraphs.length - 1) {
            paragraphs[i + 1].removeFromParent();
          }
        } else {
          // Instead of removing, just clear the text
          paragraphs[i].clear();
        }
        break;
      }
    }
  } else {
    // Update exchange rate placeholders for USD
    body.replaceText(
      "\\{Exchange Rate\\}",
      parseFloat(data.exchangeRate).toFixed(4)
    );
    body.replaceText(
      "\\{Amount in EUR\\}",
      `€${parseFloat(data.amountInEUR).toFixed(2)}`
    );
  }
}

/**
 * Update invoice table in document
 * @param {Body} body - Document body
 * @param {Object} data - Invoice data
 */
function updateInvoiceTable(body, data) {
  const tables = body.getTables();
  let targetTable = null;

  // Find the correct table
  for (const table of tables) {
    const headers = [];
    for (let i = 0; i < table.getRow(0).getNumCells(); i++) {
      headers.push(table.getRow(0).getCell(i).getText().trim());
    }

    if (
      headers.length >= 6 &&
      headers[0] === "#" &&
      headers[1] === "Services" &&
      headers[2] === "Period" &&
      headers[3] === "Quantity" &&
      headers[4] === "Rate/hour" &&
      headers[5] === "Amount"
    ) {
      targetTable = table;
      break;
    }
  }

  if (!targetTable) {
    throw new Error(ERROR_MESSAGES.TABLE_NOT_FOUND);
  }

  // Clear existing rows (keep header)
  const numRows = targetTable.getNumRows();
  for (let i = numRows - 1; i > 0; i--) {
    targetTable.removeRow(i);
  }

  // Add new rows
  data.items.forEach((row) => {
    const newRow = targetTable.appendTableRow();
    row.forEach((cell, index) => {
      const cellElement = newRow.appendTableCell(
        index === 4 || index === 5
          ? cell
            ? formatCurrencyFromUtils(cell, data.currency)
            : ""
          : cell || ""
      );

      // Выравнивание вправо для сумм (Rate/hour и Amount)
      if (index === 4 || index === 5) {
        cellElement
          .getChild(0)
          .asParagraph()
          .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      }
    });
  });
}

/**
 * Replace placeholders in document
 * @param {Body} body - Document body
 * @param {Object} data - Invoice data
 * @param {string} formattedDate - Formatted invoice date
 * @param {string} formattedDueDate - Formatted due date
 * @param {number} taxRate - Tax rate
 * @param {number} taxAmount - Tax amount
 * @param {number} totalAmount - Total amount
 */
function replaceDocumentPlaceholders(
  body,
  data,
  formattedDate,
  formattedDueDate,
  taxRate,
  taxAmount,
  totalAmount
) {
  // Basic invoice information
  const replacements = {
    "\\{Project Name\\}": data.projectName,
    "\\{Название клиента\\}": data.clientName,
    "\\{Адрес клиента\\}": data.clientAddress,
    "\\{Номер клиента\\}": data.clientNumber,
    "\\{Номер счета\\}": data.invoiceNumber,
    "\\{Дата счета\\}": formattedDate,
    "\\{Due date\\}": formattedDueDate,
    "\\{VAT%\\}": taxRate.toFixed(0),
    "\\{Сумма НДС\\}": formatCurrencyFromUtils(taxAmount, data.currency),
    "\\{Сумма общая\\}": formatCurrencyFromUtils(totalAmount, data.currency),
    "\\{Банковские реквизиты1\\}": data.bankDetails1,
    "\\{Банковские реквизиты2\\}": data.bankDetails2,
    "\\{Комментарий\\}": data.comment || "",
  };

  // Apply basic replacements
  Object.entries(replacements).forEach(([placeholder, value]) => {
    body.replaceText(placeholder, value);
  });

  // Replace item-specific placeholders
  for (let i = 0; i < CONFIG.INVOICE_TABLE.MAX_ROWS; i++) {
    const item = data.items[i];
    if (item) {
      const itemReplacements = {
        [`\\{Вид работ-${i + 1}\\}`]: item[1] || "",
        [`\\{Период работы-${i + 1}\\}`]: item[2] || "",
        [`\\{Часы-${i + 1}\\}`]: item[3] || "",
      };

      Object.entries(itemReplacements).forEach(([placeholder, value]) => {
        body.replaceText(placeholder, value);
      });

      // Handle currency fields
      if (item[4]) {
        body.replaceText(
          `\\{Рейт-${i + 1}\\}`,
          formatCurrencyFromUtils(item[4], data.currency)
        );
      }
      if (item[5]) {
        body.replaceText(
          `\\{Сумма-${i + 1}\\}`,
          formatCurrencyFromUtils(item[5], data.currency)
        );
      }
    }
  }
}

/**
 * Generate PDF from document and save to Drive
 * @param {Document} doc - Google Document object
 * @param {string} filename - Filename for PDF
 * @returns {File} PDF file object
 */
function generateAndSavePDF(doc, filename) {
  try {
    // Wait a bit for document to be fully processed
    Utilities.sleep(500);

    const pdf = doc.getAs("application/pdf");
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);

    const pdfFile = folder.createFile(pdf).setName(`${filename}.pdf`);

    return pdfFile;
  } catch (error) {
    console.error("Error generating PDF:", error);
    throw error;
  }
}

/**
 * Update spreadsheet with document URLs
 * @param {number} rowIndex - Row index in spreadsheet
 * @param {string} docUrl - Document URL
 * @param {string} pdfUrl - PDF URL
 */
function updateSpreadsheetWithUrls(rowIndex, docUrl, pdfUrl) {
  try {
    const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheets()[0];

    // Update Google Doc Link (column 19)
    sheet.getRange(rowIndex, 20).setValue(docUrl);
    // Update PDF Link (column 20)
    sheet.getRange(rowIndex, 21).setValue(pdfUrl);
  } catch (error) {
    console.error("Error updating spreadsheet with URLs:", error);
    throw error;
  }
}

/**
 * Get Google Drive folder ID for the given project name from the Lists sheet.
 * Falls back to CONFIG.FOLDER_ID if folder link is missing or invalid.
 * @param {string} projectName - The project name to look up.
 * @returns {string} The extracted folder ID.
 */
function getProjectFolderId(projectName) {
  const spreadsheet = getSpreadsheet(CONFIG.SPREADSHEET_ID);
  const sheet = getSheet(spreadsheet, CONFIG.SHEETS.LISTS);
  const values = sheet.getDataRange().getValues();

  Logger.log(
    ">>> Looking for project folder. Input projectName: " + projectName
  );

  for (let i = 1; i < values.length; i++) {
    const rowName = (values[i][CONFIG.COLUMNS.PROJECT_NAME] || "")
      .toString()
      .trim();
    Logger.log(`Row ${i}: rowName="${rowName}"`);

    if (!rowName || !projectName) continue;

    if (rowName.toLowerCase() === projectName.toString().trim().toLowerCase()) {
      const folderUrl = (values[i][12] || "").toString().trim(); // column M
      const match = folderUrl.match(/[-\w]{25,}/);

      Logger.log(`>>> Found folder URL for ${projectName}: ${folderUrl}`);
      Logger.log(`>>> Extracted folderId: ${match ? match[0] : "NONE"}`);

      if (match) {
        return match[0];
      } else {
        Logger.log(
          `>>> No valid folder ID for ${projectName}, fallback to CONFIG.FOLDER_ID`
        );
        return CONFIG.FOLDER_ID;
      }
    }
  }

  Logger.log(
    `>>> Project ${projectName} not found in Lists, fallback to CONFIG.FOLDER_ID`
  );
  return CONFIG.FOLDER_ID;
}

/**
 * Create credit note document from template
 * @param {Object} data - Credit note data
 * @param {string} formattedDate - Formatted credit note date
 * @param {number} subtotal - Subtotal amount
 * @param {number} taxRate - Tax rate
 * @param {number} taxAmount - Tax amount
 * @param {number} totalAmount - Total amount
 * @param {string} templateId - Template document ID
 * @param {string} folderId - Folder ID for saving
 * @returns {Document} Google Document object
 */
function createCreditNoteDoc(
  data,
  formattedDate,
  subtotal,
  taxRate,
  taxAmount,
  totalAmount,
  templateId,
  folderId
) {
  Logger.log(`createCreditNoteDoc: Starting for template ID: ${templateId}`);
  if (!templateId) {
    Logger.log("createCreditNoteDoc: ERROR - No templateId provided.");
    throw new Error("No credit note template found for the selected project.");
  }

  try {
    const template = DriveApp.getFileById(templateId);
    Logger.log(">>> Using folderId: " + folderId);

    const folder = DriveApp.getFolderById(folderId);

    const filename = generateCreditNoteFilenameFromUtils(data);
    Logger.log(`createCreditNoteDoc: Generated filename: ${filename}`);

    const copy = template.makeCopy(filename, folder);
    Logger.log(`createCreditNoteDoc: Created copy with ID: ${copy.getId()}`);

    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    // Handle exchange rate section
    handleCreditNoteExchangeRateSection(body, data);

    // Update credit note table
    updateCreditNoteTable(body, data);

    // Replace placeholders
    replaceCreditNoteDocumentPlaceholders(
      body,
      data,
      formattedDate,
      taxRate,
      taxAmount,
      totalAmount
    );

    Logger.log(
      `createCreditNoteDoc: Placeholders replaced. Saving and closing doc.`
    );
    doc.saveAndClose();
    Logger.log(
      `createCreditNoteDoc: Document saved and closed. Returning doc object.`
    );
    return doc;
  } catch (error) {
    Logger.log(`createCreditNoteDoc: CRITICAL ERROR - ${error.toString()}`);
    Logger.log(`Stack Trace: ${error.stack}`);
    console.error("Error creating credit note document:", error);
    throw error;
  }
}

/**
 * Handle exchange rate section in credit note document
 * @param {Body} body - Document body
 * @param {Object} data - Credit note data
 */
function handleCreditNoteExchangeRateSection(body, data) {
  if (data.currency !== "$") {
    // Clear exchange rate notice for non-USD currencies (instead of removing paragraphs)
    const paragraphs = body.getParagraphs();
    for (let i = 0; i < paragraphs.length; i++) {
      const text = paragraphs[i].getText();
      if (text.includes("Exchange Rate Notice")) {
        // Clear the text instead of removing the paragraph
        paragraphs[i].clear();
        if (i + 1 < paragraphs.length) {
          paragraphs[i + 1].clear();
        }
        break;
      }
    }
  } else {
    // Update exchange rate placeholders for USD
    body.replaceText(
      "\\{Exchange Rate\\}",
      parseFloat(data.exchangeRate).toFixed(4)
    );
    body.replaceText(
      "\\{Amount in EUR\\}",
      `€${parseFloat(data.amountInEUR).toFixed(2)}`
    );
  }
}

/**
 * Update credit note table in document
 * @param {Body} body - Document body
 * @param {Object} data - Credit note data
 */
function updateCreditNoteTable(body, data) {
  try {
    const tables = body.getTables();
    let targetTable = null;

    Logger.log(
      `updateCreditNoteTable: Found ${tables.length} tables in document`
    );

    // Find the correct table - look for table with specific headers
    for (let tableIndex = 0; tableIndex < tables.length; tableIndex++) {
      const table = tables[tableIndex];
      if (table.getNumRows() === 0) {
        Logger.log(
          `updateCreditNoteTable: Table ${tableIndex} has no rows, skipping`
        );
        continue;
      }

      const headers = [];
      const headerRow = table.getRow(0);
      for (let i = 0; i < headerRow.getNumCells(); i++) {
        headers.push(headerRow.getCell(i).getText().trim());
      }

      Logger.log(
        `updateCreditNoteTable: Table ${tableIndex} headers: [${headers.join(
          ", "
        )}]`
      );

      // Check if this is the credit note items table
      // Look for a table with headers like "#", "Description", "Period", "Amount"
      if (
        headers.length >= 4 &&
        (headers[0] === "#" || headers[0] === "№") &&
        (headers[1].toLowerCase().includes("description") ||
          headers[1].toLowerCase().includes("описание") ||
          headers[1].toLowerCase().includes("services")) &&
        (headers[2].toLowerCase().includes("period") ||
          headers[2].toLowerCase().includes("период")) &&
        (headers[3].toLowerCase().includes("amount") ||
          headers[3].toLowerCase().includes("сумма"))
      ) {
        targetTable = table;
        Logger.log(
          `updateCreditNoteTable: Found matching table at index ${tableIndex}`
        );
        break;
      }
    }

    if (!targetTable) {
      Logger.log(
        "updateCreditNoteTable: No suitable table found, skipping table update"
      );
      return;
    }

    // Clear existing rows (keep header)
    const numRows = targetTable.getNumRows();
    for (let i = numRows - 1; i > 0; i--) {
      targetTable.removeRow(i);
    }

    // Add new rows
    if (data.items && data.items.length > 0) {
      data.items.forEach((row) => {
        const newRow = targetTable.appendTableRow();
        row.forEach((cell, index) => {
          const cellElement = newRow.appendTableCell(
            index === 3 // Amount column
              ? cell
                ? formatCurrencyFromUtils(cell, data.currency)
                : ""
              : cell || ""
          );

          // Right-align amount column
          if (index === 3) {
            cellElement
              .getChild(0)
              .asParagraph()
              .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
          }
        });
      });
    } else {
      Logger.log("updateCreditNoteTable: No items to add to table");
    }
  } catch (error) {
    Logger.log(`updateCreditNoteTable: Error occurred - ${error.toString()}`);
    Logger.log(
      `updateCreditNoteTable: Continuing with placeholder replacement...`
    );
    // Don't throw error - let the main function continue
  }
}

/**
 * Replace placeholders in credit note document
 * @param {Body} body - Document body
 * @param {Object} data - Credit note data
 * @param {string} formattedDate - Formatted credit note date
 * @param {number} taxRate - Tax rate
 * @param {number} taxAmount - Tax amount
 * @param {number} totalAmount - Total amount
 */
function replaceCreditNoteDocumentPlaceholders(
  body,
  data,
  formattedDate,
  taxRate,
  taxAmount,
  totalAmount
) {
  Logger.log(
    `replaceCreditNoteDocumentPlaceholders: STARTING - taxRate=${taxRate}, taxAmount=${taxAmount}, totalAmount=${totalAmount}`
  );

  // Basic credit note information
  const replacements = {
    "\\{Номер CN\\}": data.creditNoteNumber,
    "\\{Название клиента\\}": data.clientName,
    "\\{Адрес клиента\\}": data.clientAddress,
    "\\{Номер клиента\\}": data.clientNumber,
    "\\{Дата CN\\}": formattedDate,
    "\\{VAT%\\}": taxRate.toFixed(0),
    "\\{Сумма НДС\\}": formatCurrencyFromUtils(taxAmount, data.currency),
    "\\{Сумма общая\\}": formatCurrencyFromUtils(totalAmount, data.currency),
    "\\{Комментарий\\}": data.comment || "",
  };

  Logger.log(
    `replaceCreditNoteDocumentPlaceholders: About to replace ${
      Object.keys(replacements).length
    } placeholders`
  );

  // Apply basic replacements
  Object.entries(replacements).forEach(([placeholder, value]) => {
    Logger.log(`Replacing ${placeholder} with "${value}"`);
    const result = body.replaceText(placeholder, value);
    Logger.log(`Replace result: ${result}`);
  });

  Logger.log(`replaceCreditNoteDocumentPlaceholders: COMPLETED`);

  // Item-specific placeholders are handled by updateCreditNoteTable function
  // No need to replace them here as the table is already updated
}
