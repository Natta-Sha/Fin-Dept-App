// Main application entry point - Optimized version
// This file contains the web app endpoints and main business logic

// ── Access Control (per-user cache, 30 min TTL) ─────────────────────────────

var ACCESS_CACHE_TTL_SECONDS = 1800; // 30 minutes
var _userAccessMap = null; // in-memory: avoids repeated CacheService calls within one request

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail().toLowerCase();
}

function getPageSection(page) {
  return CONFIG.ACCESS_CONTROL.PAGE_TO_SECTION[page] || null;
}

/**
 * Read emails from a sheet (column A). Skips empty cells and optional header row.
 */
function getEmailsFromAccessSheet(spreadsheetId, sheetName) {
  try {
    var spreadsheet = getSpreadsheet(spreadsheetId);
    var sheet = getSheet(spreadsheet, sheetName);
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return [];
    var values = sheet.getRange(1, 1, lastRow, 1).getValues();
    var emails = [];
    for (var i = 0; i < values.length; i++) {
      var cell = values[i][0];
      if (cell === null || cell === undefined) continue;
      var s = String(cell).trim().toLowerCase();
      if (s === "") continue;
      if (i === 0 && (s === "email" || s === "e-mail")) continue;
      emails.push(s);
    }
    return emails;
  } catch (e) {
    console.error("getEmailsFromAccessSheet(" + sheetName + "):", e);
    return [];
  }
}

/**
 * Build full access map for a user by reading all access sheets.
 * Called only on cache miss (once per 30 min per user).
 */
function buildUserAccessMap(email) {
  var ac = CONFIG.ACCESS_CONTROL;
  var normalizedEmail = email.toLowerCase();

  var fullAccessEmails = getEmailsFromAccessSheet(ac.SPREADSHEET_ID, ac.SHEETS.FULL_ACCESS);
  var isFullAccess = fullAccessEmails.indexOf(normalizedEmail) !== -1;

  var allSections = {};
  Object.keys(ac.PAGE_TO_SECTION).forEach(function (page) {
    allSections[ac.PAGE_TO_SECTION[page]] = true;
  });

  var access = {};
  Object.keys(allSections).forEach(function (section) {
    if (isFullAccess) {
      access[section] = true;
      return;
    }
    var sheetName = ac.SECTION_SHEETS[section];
    if (!sheetName) {
      access[section] = false;
      return;
    }
    var sectionEmails = getEmailsFromAccessSheet(ac.SPREADSHEET_ID, sheetName);
    access[section] = sectionEmails.indexOf(normalizedEmail) !== -1;
  });

  return access;
}

/**
 * Get the access map for the current user. One CacheService call per page load,
 * or zero if already loaded during this request.
 */
function getUserNavAccess(email) {
  if (_userAccessMap) return _userAccessMap;

  var cacheKey = "access_user_" + email.toLowerCase();
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);

  if (cached !== null) {
    _userAccessMap = JSON.parse(cached);
    return _userAccessMap;
  }

  _userAccessMap = buildUserAccessMap(email);
  cache.put(cacheKey, JSON.stringify(_userAccessMap), ACCESS_CACHE_TTL_SECONDS);
  return _userAccessMap;
}

function hasAccessToSection(email, section) {
  var access = getUserNavAccess(email);
  return access[section] === true;
}

function hasAccessToPage(email, page) {
  var section = getPageSection(page);
  if (!section) {
    var access = getUserNavAccess(email);
    return Object.keys(access).some(function (k) { return access[k]; });
  }
  return hasAccessToSection(email, section);
}

/**
 * Clear all access caches. Call after editing the access spreadsheet.
 * Can be run from Apps Script editor or attached to a button in the spreadsheet.
 */
function clearAccessCache() {
  var cache = CacheService.getScriptCache();
  cache.remove("access_user_" + getCurrentUserEmail());

  var ac = CONFIG.ACCESS_CONTROL;
  var allSheets = [ac.SHEETS.FULL_ACCESS];
  Object.keys(ac.SECTION_SHEETS).forEach(function (section) {
    allSheets.push(ac.SECTION_SHEETS[section]);
  });

  var spreadsheet = getSpreadsheet(ac.SPREADSHEET_ID);
  allSheets.forEach(function (sheetName) {
    var sheet = getSheet(spreadsheet, sheetName);
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return;
    var values = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var s = String(values[i][0] || "").trim().toLowerCase();
      if (s === "" || s === "email" || s === "e-mail") continue;
      cache.remove("access_user_" + s);
    }
  });

  _userAccessMap = null;
  return { success: true, message: "Access cache cleared" };
}

// ── Web App Entry Point ─────────────────────────────────────────────────────

/**
 * Main web app entry point
 * @param {Object} e - Event object with parameters
 * @returns {HtmlOutput} Rendered HTML page
 */
function doGet(e) {
  try {
    var email = getCurrentUserEmail();
    var page = e.parameter.page || "Home";

    // Access check
    if (!hasAccessToPage(email, page)) {
      var navAccess = getUserNavAccess(email);
      var hasAnyAccess = Object.keys(navAccess).some(function (k) {
        return navAccess[k];
      });

      if (!hasAnyAccess) {
        return HtmlService.createHtmlOutput(
          '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
            "<style>" +
            "body{display:flex;justify-content:center;align-items:center;height:100vh;margin:0;" +
            'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#f8f9fa;}' +
            ".card{text-align:center;background:#fff;border-radius:15px;padding:3rem;" +
            "box-shadow:0 10px 30px rgba(0,0,0,.1);max-width:460px;}" +
            "h1{color:#6c757d;font-size:2rem;margin-bottom:.75rem;}" +
            "p{color:#adb5bd;font-size:1.05rem;margin:0;}" +
            ".email{font-size:.85rem;color:#ced4da;margin-top:1.5rem;}" +
            "</style></head><body>" +
            '<div class="card"><h1>No access</h1>' +
            "<p>You do not have permission to use this application.</p>" +
            '<p class="email">' +
            email +
            "</p></div></body></html>"
        ).setTitle("No Access");
      }

      var denied = HtmlService.createTemplateFromFile("AccessDenied");
      denied.baseUrl = ScriptApp.getService().getUrl();
      denied.activePage = "";
      return denied.evaluate().setTitle("No Access");
    }

    var template = HtmlService.createTemplateFromFile(page);
    template.baseUrl = ScriptApp.getService().getUrl();
    template.invoiceId = e.parameter.invoiceId || e.parameter.id || "";
    template.creditNoteId = e.parameter.invoiceId || e.parameter.id || "";
    template.contractId = e.parameter.contractId || e.parameter.id || "";
    template.mode = e.parameter.mode || "";

    // Set active page for navigation
    template.activePage = getActivePageForNavigation(page, e.parameter);

    // Pass invoice ID if provided
    if (e.parameter.invoiceId || e.parameter.id) {
      template.invoiceId = e.parameter.invoiceId || e.parameter.id;
    }
    // Pass credit note ID if provided (same as invoice ID for now)
    if (e.parameter.invoiceId || e.parameter.id) {
      template.creditNoteId = e.parameter.invoiceId || e.parameter.id;
    }
    // Pass contract ID if provided
    if (e.parameter.contractId || e.parameter.id) {
      template.contractId = e.parameter.contractId || e.parameter.id;
    }
    if (e.parameter.mode) {
      template.mode = e.parameter.mode;
    }

    return template.evaluate().setTitle(page);
  } catch (error) {
    console.error("Error in doGet:", error);
    return HtmlService.createHtmlOutput(
      `<h1>Error</h1><p>${error.message}</p>`
    );
  }
}

/**
 * Load HTML page content
 * @param {string} name - Page name
 * @returns {string} HTML content
 */
function loadPage(name) {
  try {
    return typeof dataService !== "undefined" && dataService.loadPage
      ? dataService.loadPage(name)
      : HtmlService.createHtmlOutputFromFile(name).getContent();
  } catch (error) {
    console.error(`Error loading page ${name}:`, error);
    return `<h1>Error</h1><p>Failed to load page: ${error.message}</p>`;
  }
}

/**
 * Process invoice form data - Main business logic
 * @param {Object} data - Form data from frontend
 * @returns {Object} Result with document and PDF URLs
 */
function processForm(data) {
  return processInvoiceCreation(data);
}

function processCreditNoteForm(data) {
  return processCreditNoteCreation(data);
}

// Export functions for use in other modules
// Note: In Google Apps Script, all functions are globally available
// These comments help with documentation and IDE support

/**
 * Get project names for dropdown
 * @returns {Array} Array of project names
 */
function getProjectNames() {
  return getProjectNamesFromData();
}

/**
 * Get project details by name
 * @param {string} projectName - Project name
 * @returns {Object} Project details
 */
function getProjectDetails(projectName) {
  return getProjectDetailsFromData(projectName);
}

/**
 * Get list of all invoices
 * @returns {Array} Array of invoice objects
 */
function getInvoiceList() {
  return getInvoiceListFromData();
}

/**
 * Get invoice data by ID
 * @param {string} id - Invoice ID
 * @returns {Object} Invoice data
 */
function getInvoiceDataById(id) {
  return getInvoiceDataByIdFromData(id);
}

/**
 * Get credit note data by ID
 * @param {string} id - Credit note ID
 * @returns {Object} Credit note data
 */
function getCreditNoteDataById(id) {
  console.log("getCreditNoteDataById called with ID:", id);
  const result = getCreditNoteDataByIdFromData(id);
  console.log("getCreditNoteDataById returning:", result);
  return result;
}

/**
 * Get credit note list
 * @returns {Array} Credit note list
 */
function getCreditNoteList() {
  return getCreditNoteListFromData();
}

/**
 * Get list of all contracts
 * @returns {Array} Array of contract objects
 */
function getContractList() {
  return getContractListFromData();
}

/**
 * Get contract data by ID
 * @param {string} id - Contract ID
 * @returns {Object} Contract data
 */
function getContractDataById(id) {
  return getContractDataByIdFromData(id);
}

/**
 * Get dropdown options for contract form
 * @returns {Object} Dropdown options
 */
function getContractDropdownOptions() {
  return getContractDropdownOptionsFromData();
}

/**
 * Get contract templates filtered by criteria
 * @param {string} cooperationType
 * @param {string} ourCompany
 * @param {string} serviceType
 * @returns {Array} Array of templates
 */
function getContractTemplates(
  cooperationType,
  ourCompany,
  serviceType,
  documentType
) {
  return getContractTemplatesFromData(
    cooperationType,
    ourCompany,
    serviceType,
    documentType
  );
}

/**
 * Save a new contract
 * @param {Object} formData - Form data from the contract generator
 * @returns {Object} Result with success status and contract ID
 */
function saveContract(formData) {
  return saveContractToData(formData);
}

/**
 * Delete a contract by ID
 * @param {string} id - Contract ID
 * @returns {Object} Result with success status
 */
function deleteContract(id) {
  return deleteContractFromData(id);
}

/**
 * Update an existing contract
 * @param {Object} formData - Contract form data including id
 * @returns {Object} Result with success status
 */
function updateContract(formData) {
  return updateContractToData(formData);
}

/**
 * Get list of all bills
 * @returns {Array} Array of bill objects
 */
function getBillList() {
  return getBillListFromData();
}

// Error handling and performance monitoring removed for cleaner code

// Performance monitoring removed for cleaner code

function validateRequiredFields(data, requiredFields) {
  return validateRequiredFieldsFromUtils(data, requiredFields);
}

function calculateTaxAmount(subtotal, taxRate) {
  return calculateTaxAmountFromUtils(subtotal, taxRate);
}

function calculateTotalAmount(subtotal, taxAmount) {
  return calculateTotalAmountFromUtils(subtotal, taxAmount);
}

// saveInvoiceData is handled directly in businessService.js

// Document service functions are available globally from documentService.js
// No wrapper functions needed in Google Apps Script

function formatDate(dateStr) {
  return formatDateFromUtils(dateStr);
}

function formatDateForInput(val) {
  return formatDateForInputFromUtils(val);
}

/**
 * Delete invoice by ID (global endpoint for frontend)
 * @param {string} id - Invoice ID
 * @returns {Object} { success: true } or { success: false, message }
 */
function deleteInvoiceById(id) {
  return deleteInvoiceByIdFromData(id);
}

/**
 * Delete credit note by ID (global endpoint for frontend)
 * @param {string} id - Credit Note ID
 * @returns {Object} { success: true } or { success: false, message }
 */
function deleteCreditNoteById(id) {
  return deleteCreditNoteByIdFromData(id);
}

/**
 * Update credit note by ID (global endpoint for frontend)
 * @param {Object} data - Credit note data with id
 * @returns {Object} { success: true } or { success: false, message }
 */
function updateCreditNoteById(data) {
  return updateCreditNoteByIdFromData(data);
}

function testLogger(message) {
  Logger.log(`[CLIENT TEST]: ${message}`);
}

/**
 * Update invoice by ID (frontend endpoint)
 * @param {string} id
 * @param {Object} data
 * @returns {Object}
 */
function updateInvoiceById(id, data) {
  return updateInvoiceByIdFromData(id, data);
}

/**
 * Get navigation HTML with active page highlighting
 * @param {string} activePage - Current active page identifier
 * @returns {string} Navigation HTML
 */
function getNavigation(activePage) {
  activePage = activePage || "";
  var template = HtmlService.createTemplateFromFile("Navigation");
  template.activePage = activePage;
  template.baseUrl = ScriptApp.getService().getUrl();
  var email = getCurrentUserEmail();
  template.navAccess = getUserNavAccess(email);
  return template.evaluate().getContent();
}

/**
 * Determine active page for navigation based on current page and parameters
 * @param {string} page - Current page name
 * @param {Object} params - URL parameters
 * @returns {string} Active page identifier for navigation
 */
function getActivePageForNavigation(page, params = {}) {
  switch (page) {
    case "Home":
      return "home";
    case "InvoicesList":
      return "invoices";
    case "InvoiceGenerator":
      // InvoiceGenerator is part of invoices section
      return "invoices";
    case "CreditNotesList":
      return "creditnotes";
    case "CreditNotesGenerator":
      // CreditNotesGenerator is part of credit notes section
      return "creditnotes";
    case "ContractsList":
      return "contracts";
    case "ContractGenerator":
      // ContractGenerator is part of contracts section
      return "contracts";
    case "BillsList":
      return "bills";
    default:
      return "";
  }
}
