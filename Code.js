// Main application entry point - Optimized version
// This file contains the web app endpoints and main business logic

/**
 * Main web app entry point
 * @param {Object} e - Event object with parameters
 * @returns {HtmlOutput} Rendered HTML page
 */
function doGet(e) {
  try {
    const page = e.parameter.page || "Home";
    const template = HtmlService.createTemplateFromFile(page);
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
function getNavigation(activePage = "") {
  const template = HtmlService.createTemplateFromFile("Navigation");
  template.activePage = activePage;
  template.baseUrl = ScriptApp.getService().getUrl();
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
