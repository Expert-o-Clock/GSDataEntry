/**
 * Professional Grade Google Apps Script Backend
 * Enhanced with error handling, validation, and performance
 * @version 3.2.0
 */

const CONFIG = {
  SHEETS: {
    DAILY_LOGS: "Daily_Logs",
    WORK_ORDERS: "Work_Orders",
    LISTS: "Lists",
    SALES_PERSONS: "Sales_Persons",
    PREFERENCES: "Preferences"
  },
  COLUMNS: {
    WO_NUMBER: 1,
    STATUS: 3,
    DATE: 7
  },
  CACHE: {
    DROPDOWN_TTL: 300 // 5 minutes cache
  }
};

const { SHEETS, COLUMNS } = CONFIG;

/**
 * Retrieves optimized dropdown options with caching and validation
 */
function getDropdownOptions(sheetName, rangeAddress) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `dropdown_${sheetName}_${rangeAddress.replace(/:/g,'-')}`;
  
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);
    
    const range = sheet.getRange(rangeAddress);
    const values = range.getDisplayValues()
      .flat()
      .filter(value => value && value.toString().trim() !== "")
      .filter((v, i, a) => a.indexOf(v) === i);

    cache.put(cacheKey, JSON.stringify(values), CONFIG.CACHE.DROPDOWN_TTL);
    return values;
  } catch (error) {
    console.error(`getDropdownOptions error: ${error.message}`);
    throw new Error(`Failed to load options: ${error.message}`);
  }
}

/**
 * Check for existing Work Order number
 */
function checkWoExists(woNumber) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.WORK_ORDERS);
    if (!sheet) return false;
    
    const woNumbers = sheet.getRange(
      `${COLUMNS.WO_NUMBER}2:${COLUMNS.WO_NUMBER}`
    ).getDisplayValues().flat();
    
    return woNumbers.includes(woNumber.toString());
  } catch (error) {
    console.error(`checkWoExists error: ${error.message}`);
    return false;
  }
}

/**
 * Submit Work Order Data with duplicate handling
 */
function submitWorkOrderData(data, overwrite = false) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEETS.WORK_ORDERS);
    if (!sheet) throw new Error("Work Orders sheet not found");

    // Validate required fields
    if (!data.workOrderNumber) throw new Error("Work Order Number required");
    if (!data.date) throw new Error("Date is required");

    // Check for duplicates
    const exists = checkWoExists(data.workOrderNumber);
    if (exists && !overwrite) throw new Error("DUPLICATE_WO");

    // Prepare formulas
    const row = exists ? 
      sheet.getRange(COLUMNS.WO_NUMBER, 1).getValues().flat().indexOf(data.workOrderNumber) + 2 : 
      sheet.getLastRow() + 1;

    const formulas = [
      data.workOrderNumber,
      data.workOrderName,
      `=XLOOKUP(A${row}, '${SHEETS.DAILY_LOGS}'!A:A, '${SHEETS.DAILY_LOGS}'!E:E, "Not Found", 0, -1)`,
      data.recurringNumber,
      data.salesPerson,
      data.crewLeader,
      null,
      Utilities.formatDate(new Date(data.date), 
      ss.getSpreadsheetTimeZone(), 
      "MM/dd/yyyy"),
      // ... rest of formulas
    ];

    // Update or append
    if (exists) {
      sheet.getRange(row, 1, 1, formulas.length).setValues([formulas]);
      return { action: "updated", row };
    } else {
      sheet.appendRow(formulas);
      return { action: "created", row: sheet.getLastRow() };
    }
  } catch (error) {
    console.error(`submitWorkOrderData error: ${error.message}`);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Submit Daily Data with enhanced validation
 */
function submitDailyData(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEETS.DAILY_LOGS);
    
    // Validation
    if (!sheet) throw new Error("Daily Logs sheet not found");
    if (!data.workOrderNumber) throw new Error("Work Order Number required");
    if (!data.date) throw new Error("Date is required");

    // Prepare payload
    const payload = [
      Utilities.formatDate(new Date(data.date), 
        ss.getSpreadsheetTimeZone(), 
        "MM/dd/yyyy"),
      data.status,
      data.workOrderNumber,
      parseFloat(data.actHours),
      parseFloat(data.actRevenue),
      data.jsaDaily ? "TRUE" : "FALSE",
      data.goBacks ? "TRUE" : "FALSE",
      data.propertyDamage ? "TRUE" : "FALSE"
    ];

    // Batch write
    sheet.appendRow(payload);
    
    return { success: true, row: sheet.getLastRow() };
  } catch (error) {
    console.error(`submitDailyData error: ${error.message}`);
    throw new Error(`Submission failed: ${error.message}`);
  } finally {
    lock.releaseLock();
  }
}

// UI Functions
function showWorkOrderSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('form_WorkOrderEntry')
    .setTitle('Work Order Management')
  SpreadsheetApp.getUi().showSidebar(html);
}

function showDailySidebar() {
  const html = HtmlService.createHtmlOutputFromFile('form_DailyDataEntry')
    .setTitle('Daily Log Entry')
  SpreadsheetApp.getUi().showSidebar(html);
}