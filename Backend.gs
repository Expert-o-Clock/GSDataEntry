const ui = SpreadsheetApp.getUi();

/**
 * Professional UI Handling with Theme Support
 * Enhanced error handling and responsive design
 */

const UI_CONFIG = {
  THEME: {
    PRIMARY: "#2c3e50",
    SECONDARY: "#3498db",
    SUCCESS: "#27ae60",
    DANGER: "#e74c3c"
  },
  DIMENSIONS: {
    SIDEBAR: 400,
    MODAL: { width: 600, height: 720 }
  }
};

function onOpen() {
  try {
    ui.createMenu('🏝️ Oahu Tree Works')
      .addSubMenu(ui.createMenu('⌨️ Data Entry')
        .addItem('📅 Daily Logs', 'showDailySidebar')
        .addItem('📝 Work Order', 'showWorkOrderSidebar'))
      .addSeparator()
      .addItem('🔄 Refresh Data', 'refreshData')
      // .addSeparator()
      // .addSubMenu(ui.createMenu('Auto-Refresh Crew Leaders Names')
      //   .addItem('Activate Schedule', 'createTrigger')
      //   .addItem('Deactivate Schedule', 'deleteTrigger'))
    .addToUi();
  } catch (error) {
    console.error(`Menu initialization failed: ${error.message}`);
  }
}

function showDailySidebar() {
  try {
    const html = HtmlService.createTemplateFromFile('form_DailyDataEntry_Log')
      .evaluate()
      .setTitle('Daily Logs')
      .setWidth(UI_CONFIG.DIMENSIONS.SIDEBAR)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

    ui.showSidebar(html);
  } catch (error) {
    showErrorToast('Failed to load Daily Logs interface');
  }
}

function showWorkOrderSidebar() {
  try {
    const html = HtmlService.createTemplateFromFile('form_WorkOrderEntry')
      .evaluate()
      .setTitle('Work Orders')
      .setWidth(UI_CONFIG.DIMENSIONS.SIDEBAR);

    ui.showSidebar(html);
  } catch (error) {
    showErrorToast('Failed to load Work Orders interface');
  }
}

function refreshData() {
  try {
    CacheService.getScriptCache().removeAll([]);
    SpreadsheetApp.getActive().toast('Data cache refreshed successfully', '✅ Success');
  } catch (error) {
    showErrorToast('Refresh failed');
  }
}

// Helper function
function showErrorToast(message) {
  SpreadsheetApp.getActive().toast(message, '⚠️ Error', 6);
}



// UI Functions
function showWorkOrderSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('form_WorkOrderEntry')
    .setTitle('Work Order Management')
  ii.showSidebar(html);
}

function showDailySidebar() {
  const html = HtmlService.createHtmlOutputFromFile('form_DailyDataEntry')
    .setTitle('Daily Log Entry')
  ui.showSidebar(html);
}




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
    DATE: 1,
    WO_NUMBER: 2,
    STATUS: 5
  },
  CACHE: {
    DROPDOWN_TTL: 300 // 5 minutes cache
  }
};

const { SHEETS, COLUMNS } = CONFIG;


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
    if (exists && !overwrite) {
      ui.alert("Work Order Already Exists.");
      throw new Error("DUPLICATE_WO");
      };

    // Prepare formulas
    const row = exists ? 
      sheet.getRange(COLUMNS.WO_NUMBER, 1).getValues().flat().indexOf(data.workOrderNumber) + 2 : 
      sheet.getLastRow() + 1;

    const formulas = [
      Utilities.formatDate(new Date(data.date), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy"),
      "'" + data.workOrderNumber.toString(),
      data.workOrderName,
      data.recurringNumber,
      data.estHours,
      data.salesPerson,
      data.crewLeader,
      null,
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
      Utilities.formatDate(new Date(data.date), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy"),
      "'" + data.workOrderNumber.toString(),
      parseFloat(data.actHours),
      parseFloat(data.actRevenue),
      data.status,
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

