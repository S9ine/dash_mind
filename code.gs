/**
 * @OnlyCurrentDoc
 */

// --- CONFIGURATION ---
const CONFIG = {
  sheetName: "ผู้จัดการ",
  settingsSheetName: "Settings", // [NEW] Sheet for authorized users
  pageTitle: "App Dashboard"
};

/**
 * [UPDATED] Main function to handle different pages and user access.
 */
function doGet(e) {
  // Page routing
  if (e.parameter.page === 'settings') {
    // Only the owner can access the settings page
    if (isOwner_()) {
      const template = HtmlService.createTemplateFromFile('Settings');
      template.dashboardUrl = ScriptApp.getService().getUrl();
      return template.evaluate().setTitle("Settings | " + CONFIG.pageTitle);
    } else {
      return HtmlService.createHtmlOutputFromFile('AccessDenied').setTitle("Access Denied");
    }
  }

  // Default page (Dashboard)
  if (checkUserAccess_()) {
    const template = HtmlService.createTemplateFromFile('Dashboard');
    template.menuItems = getMenuItems_();
    template.isOwner = isOwner_(); // [NEW] Pass owner status to the template
    return template.evaluate()
      .setTitle(CONFIG.pageTitle)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } else {
    return HtmlService.createHtmlOutputFromFile('AccessDenied').setTitle("Access Denied");
  }
}


/**
 * Helper function to include other HTML files (like CSS) into the main template.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Fetches and processes data from the spreadsheet.
 */
function getMenuItems_() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    const range = sheet.getRange(`A2:D${sheet.getLastRow()}`);
    const values = range.getValues();

    return values.map(row => {
      const [name, webAppUrl, sheetUrl, devUrl] = row;
      if (name && name.toString().trim() !== '') {
        return { name, webAppUrl, sheetUrl, devUrl };
      }
      return null;
    }).filter(Boolean);

  } catch (e) {
    console.error("getMenuItems_ Error: " + e.message);
    return [];
  }
}

// ===============================================
// === [NEW] ACCESS CONTROL & SETTINGS FUNCTIONS ===
// ===============================================

/**
 * Checks if the current user is the owner of the spreadsheet.
 * @returns {boolean} True if the user is the owner.
 */
function isOwner_() {
  return Session.getActiveUser().getEmail() === Session.getEffectiveUser().getEmail();
}

/**
 * Checks if the current user has access to the web app.
 * Access is granted if the user is the owner or their email is in the Settings sheet.
 * @returns {boolean} True if access is granted.
 */
function checkUserAccess_() {
  if (isOwner_()) return true; // Owner always has access

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.settingsSheetName);
  if (!settingsSheet || settingsSheet.getLastRow() < 1) return false; // No settings sheet or no emails listed

  const allowedEmails = new Set(
    settingsSheet.getRange(`A1:A${settingsSheet.getLastRow()}`)
      .getValues()
      .flat()
      .map(e => e.toString().trim().toLowerCase())
      .filter(Boolean)
  );
  
  const currentUser = Session.getActiveUser().getEmail().toLowerCase();
  return allowedEmails.has(currentUser);
}

/**
 * [WEB APP FUNCTION] Gets the list of authorized emails for the settings page.
 * @returns {string[]} An array of email addresses.
 */
function getAuthorizedEmails() {
  if (!isOwner_()) return []; // Security check

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.settingsSheetName);
  if (!settingsSheet || settingsSheet.getLastRow() < 1) return [];

  return settingsSheet.getRange(`A1:A${settingsSheet.getLastRow()}`).getValues().flat().filter(Boolean);
}

/**
 * [WEB APP FUNCTION] Saves the list of authorized emails from the settings page.
 * @param {string[]} emails - An array of email addresses to save.
 * @returns {object} A success or error message.
 */
function saveAuthorizedEmails(emails) {
  if (!isOwner_()) {
    return { success: false, message: "Only the owner can save settings." };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let settingsSheet = ss.getSheetByName(CONFIG.settingsSheetName);

    // Create the sheet if it doesn't exist
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet(CONFIG.settingsSheetName);
      settingsSheet.getRange("A1").setValue("Authorized Emails").setFontWeight("bold");
    }
    
    // Clear old data and write new data
    settingsSheet.getRange("A2:A").clearContent();
    if (emails && emails.length > 0) {
      settingsSheet.getRange(2, 1, emails.length, 1).setValues(emails.map(e => [e]));
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
