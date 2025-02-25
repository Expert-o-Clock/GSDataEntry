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
    SpreadsheetApp.getUi()
      .createMenu('🏝️ Oahu Tree Works')
      .addItem('📅 Daily Logs', 'showDailySidebar')
      .addItem('📝 Work Orders', 'showWorkOrderSidebar')
      .addSeparator()
      .addItem('🔄 Refresh Data', 'refreshData')
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

    SpreadsheetApp.getUi().showSidebar(html);
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

    SpreadsheetApp.getUi().showSidebar(html);
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