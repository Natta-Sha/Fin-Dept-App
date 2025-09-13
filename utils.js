// Utility functions for common operations

/**
 * Format date from YYYY-MM-DD to DD/MM/YYYY
 * @param {string|Date} dateStr - Date string or Date object
 * @returns {string} Formatted date string
 */
function formatDateFromUtils(dateStr) {
  if (!dateStr) return "";

  // If it's already a Date object
  if (dateStr instanceof Date) {
    const dd = String(dateStr.getDate()).padStart(2, "0");
    const mm = String(dateStr.getMonth() + 1).padStart(2, "0");
    const yyyy = dateStr.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }

  // If it's an ISO string (2025-06-17T00:00:00Z)
  if (typeof dateStr === "string" && dateStr.includes("T")) {
    const [yyyy, mm, dd] = dateStr.split("T")[0].split("-");
    return `${dd}/${mm}/${yyyy}`;
  }

  // If it's already in DD/MM/YYYY format
  if (typeof dateStr === "string" && dateStr.indexOf("-") === -1) {
    return dateStr;
  }

  // If it's in YYYY-MM-DD format
  if (typeof dateStr === "string" && dateStr.indexOf("-") !== -1) {
    const [yyyy, mm, dd] = dateStr.split("-");
    return `${dd}/${mm}/${yyyy}`;
  }

  return dateStr.toString();
}

/**
 * Format date from any format to YYYY-MM-DD for input fields
 * @param {string|Date} val - Date value
 * @returns {string} YYYY-MM-DD formatted string
 */
function formatDateForInputFromUtils(val) {
  if (val instanceof Date) {
    const yyyy = val.getFullYear();
    const mm = String(val.getMonth() + 1).padStart(2, "0");
    const dd = String(val.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }
  return val;
}

/**
 * Clean filename by removing invalid characters
 * @param {string} filename - Original filename
 * @returns {string} Cleaned filename
 */
function cleanFilename(filename) {
  return (filename || "").replace(/[\\/:*?"<>|]/g, "").trim();
}

/**
 * Generate invoice filename
 * @param {Object} data - Invoice data
 * @returns {string} Generated filename
 */
function generateInvoiceFilenameFromUtils(data) {
  const cleanCompany = cleanFilename(data.ourCompany);
  const cleanClient = cleanFilename(data.clientName);
  return `${data.invoiceDate}_Invoice${data.invoiceNumber}_${cleanCompany}-${cleanClient}`;
}

/**
 * Format currency amount with symbol
 * @param {number|string} amount - Amount to format
 * @param {string} currency - Currency symbol
 * @returns {string} Formatted amount
 */
function formatCurrencyFromUtils(amount, currency) {
  if (!amount) return "";
  const num = parseFloat(amount);
  if (isNaN(num)) return "";

  // Format the number first, then add currency symbol in the correct position
  const formattedNumber = num.toLocaleString(undefined, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

  // For negative numbers, put the minus sign before the currency symbol
  if (num < 0) {
    return `-${currency}${formattedNumber.substring(1)}`; // Remove the minus from formatted number and add it before currency
  }

  return `${currency}${formattedNumber}`;
}

/**
 * Calculate tax amount
 * @param {number} subtotal - Subtotal amount
 * @param {number} taxRate - Tax rate percentage
 * @returns {number} Tax amount
 */
function calculateTaxAmountFromUtils(subtotal, taxRate) {
  const subtotalNum = parseFloat(subtotal) || 0;
  const rate = parseFloat(taxRate) || 0;
  return (subtotalNum * rate) / 100;
}

/**
 * Calculate total amount
 * @param {number} subtotal - Subtotal amount
 * @param {number} taxAmount - Tax amount
 * @returns {number} Total amount
 */
function calculateTotalAmountFromUtils(subtotal, taxAmount) {
  const subtotalNum = parseFloat(subtotal) || 0;
  const taxNum = parseFloat(taxAmount) || 0;
  return subtotalNum + taxNum;
}

/**
 * Validate required fields
 * @param {Object} data - Data object to validate
 * @param {Array} requiredFields - Array of required field names
 * @returns {Object} Validation result with isValid and errors
 */
function validateRequiredFieldsFromUtils(data, requiredFields) {
  const errors = [];

  requiredFields.forEach((field) => {
    if (!data[field] || data[field].toString().trim() === "") {
      errors.push(`${field} is required`);
    }
  });

  return {
    isValid: errors.length === 0,
    errors: errors,
  };
}

/**
 * Get spreadsheet by ID with error handling
 * @param {string} spreadsheetId - Google Sheets ID
 * @returns {Spreadsheet} Spreadsheet object
 */
function getSpreadsheet(spreadsheetId) {
  try {
    if (
      !spreadsheetId ||
      typeof spreadsheetId !== "string" ||
      spreadsheetId.trim() === ""
    ) {
      throw new Error("Invalid spreadsheet ID provided");
    }
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw new Error(`Failed to open spreadsheet: ${error.message}`);
  }
}

/**
 * Get sheet by name with error handling
 * @param {Spreadsheet} spreadsheet - Spreadsheet object
 * @param {string} sheetName - Sheet name
 * @returns {Sheet} Sheet object
 */
function getSheet(spreadsheet, sheetName) {
  try {
    if (!spreadsheet) {
      throw new Error("Spreadsheet object is null or undefined");
    }
    if (
      !sheetName ||
      typeof sheetName !== "string" ||
      sheetName.trim() === ""
    ) {
      throw new Error("Invalid sheet name provided");
    }
    return spreadsheet.getSheetByName(sheetName);
  } catch (error) {
    throw new Error(`Failed to get sheet "${sheetName}": ${error.message}`);
  }
}
