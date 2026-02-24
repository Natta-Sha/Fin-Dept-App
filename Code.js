// Main application entry point - Optimized version
// This file contains the web app endpoints and main business logic

// ── Access Control (emails from spreadsheet, cached 5 min) ───────────────────

var ACCESS_CACHE_TTL_SECONDS = 300; // 5 minutes
var ACCESS_CACHE_KEYS = {
  FULL: "access_emails_full",
  HOME: "access_emails_limited_home",
  CONTRACTS: "access_emails_limited_contracts",
};

// In-memory store — avoids repeated CacheService calls within a single request
var _accessMemory = {};

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
 * Two-layer cache: in-memory (instant, within request) → CacheService (5 min, across requests) → Spreadsheet.
 */
function getCachedEmails(cacheKey, sheetName) {
  if (_accessMemory[cacheKey]) return _accessMemory[cacheKey];

  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached !== null) {
    var parsed = JSON.parse(cached);
    _accessMemory[cacheKey] = parsed;
    return parsed;
  }

  var ac = CONFIG.ACCESS_CONTROL;
  var emails = getEmailsFromAccessSheet(ac.SPREADSHEET_ID, sheetName);
  cache.put(cacheKey, JSON.stringify(emails), ACCESS_CACHE_TTL_SECONDS);
  _accessMemory[cacheKey] = emails;
  return emails;
}

function getDefaultEmails() {
  return getCachedEmails(ACCESS_CACHE_KEYS.FULL, CONFIG.ACCESS_CONTROL.SHEETS.FULL_ACCESS);
}

function getSectionExtraEmails(section) {
  var ac = CONFIG.ACCESS_CONTROL;
  var sheetName = ac.SECTION_SHEETS[section];
  if (!sheetName) return [];
  var cacheKey = "access_emails_limited_" + section;
  return getCachedEmails(cacheKey, sheetName);
}

function hasAccessToSection(email, section) {
  var normalizedEmail = email.toLowerCase();
  var defaultEmails = getDefaultEmails();
  for (var i = 0; i < defaultEmails.length; i++) {
    if (defaultEmails[i] === normalizedEmail) return true;
  }
  var extraEmails = getSectionExtraEmails(section);
  for (var j = 0; j < extraEmails.length; j++) {
    if (extraEmails[j] === normalizedEmail) return true;
  }
  return false;
}

function hasAccessToPage(email, page) {
  var section = getPageSection(page);
  if (!section) {
    var defaultEmails = getDefaultEmails();
    var normalized = email.toLowerCase();
    for (var i = 0; i < defaultEmails.length; i++) {
      if (defaultEmails[i] === normalized) return true;
    }
    return false;
  }
  return hasAccessToSection(email, section);
}

/**
 * Build a map of {sectionName: boolean} for every section found in PAGE_TO_SECTION.
 */
function getUserNavAccess(email) {
  var ac = CONFIG.ACCESS_CONTROL;
  var sections = {};
  Object.keys(ac.PAGE_TO_SECTION).forEach(function (page) {
    sections[ac.PAGE_TO_SECTION[page]] = true;
  });
  var access = {};
  Object.keys(sections).forEach(function (section) {
    access[section] = hasAccessToSection(email, section);
  });
  return access;
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
    default:
      return "";
  }
}
