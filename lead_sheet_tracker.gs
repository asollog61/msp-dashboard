/**
 * Lead Sheet Change Tracker
 * 
 * Auto-stamps "Last Modified" and "Modified By" on every edit.
 * Logs all changes to a "Lead Sheet Changelog" sheet.
 * 
 * SETUP:
 * 1. Open the Lead Sheet Google Sheet
 * 2. Extensions → Apps Script
 * 3. Paste this entire script
 * 4. Click Save
 * 5. Run "setup" function once (it will ask for permissions — approve them)
 * 6. Done! Changes are tracked automatically.
 */

// Config — update these if your sheet name differs
var LEAD_SHEET_GID = 1762616490;
var CHANGELOG_SHEET_NAME = "Lead Sheet Changelog";

function setup() {
  // Create changelog sheet if it doesn't exist
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var changelog = ss.getSheetByName(CHANGELOG_SHEET_NAME);
  if (!changelog) {
    changelog = ss.insertSheet(CHANGELOG_SHEET_NAME);
    changelog.appendRow(["Timestamp", "User", "Section", "Row", "Tenant", "Column", "Old Value", "New Value"]);
    changelog.getRange("A1:H1").setFontWeight("bold");
    changelog.setColumnWidth(1, 160);
    changelog.setColumnWidth(2, 180);
    changelog.setColumnWidth(3, 80);
    changelog.setColumnWidth(6, 120);
    changelog.setColumnWidth(7, 150);
    changelog.setColumnWidth(8, 150);
  }
  
  // Create the onEdit trigger (installable trigger — needed for email/user info)
  var triggers = ScriptApp.getProjectTriggers();
  var hasEditTrigger = triggers.some(function(t) { return t.getHandlerFunction() === "onEditTracker"; });
  if (!hasEditTrigger) {
    ScriptApp.newTrigger("onEditTracker")
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }
  
  Logger.log("Setup complete! Changelog sheet created and trigger installed.");
  SpreadsheetApp.getUi().alert("Setup complete! Change tracking is now active.");
}

function onEditTracker(e) {
  if (!e) return;
  
  var sheet = e.range.getSheet();
  
  // Only track edits on the Lead Sheet tab
  if (sheet.getSheetId() !== LEAD_SHEET_GID) return;
  
  var row = e.range.getRow();
  var col = e.range.getColumn();
  
  // Skip header row
  if (row <= 1) return;
  
  var user = Session.getActiveUser().getEmail() || "Unknown";
  var timestamp = new Date();
  var oldValue = e.oldValue || "";
  var newValue = e.range.getValue() || "";
  
  // Skip if nothing actually changed
  if (String(oldValue) === String(newValue)) return;
  
  // Get header name for the edited column
  var header = sheet.getRange(1, col).getValue() || ("Column " + col);
  
  // Get tenant name (first column of the row, or first column of the section)
  var tenant = sheet.getRange(row, 1).getValue() || "";
  
  // Determine section (Retail vs Office) based on column position
  // Find the separator column (empty header)
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var sepCol = -1;
  for (var i = 1; i < headers.length; i++) {
    if (!String(headers[i]).trim() && i + 1 < headers.length && String(headers[i + 1]).trim()) {
      sepCol = i + 1; // 1-indexed
      break;
    }
  }
  
  var section = "Retail";
  if (sepCol > 0 && col > sepCol) {
    section = "Office";
    // Get the tenant from the office section's first column
    tenant = sheet.getRange(row, sepCol + 1).getValue() || "";
  }
  
  // Log to changelog
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var changelog = ss.getSheetByName(CHANGELOG_SHEET_NAME);
  if (!changelog) return;
  
  changelog.appendRow([
    timestamp,
    user,
    section,
    row,
    tenant,
    header,
    oldValue,
    newValue
  ]);
  
  // Format the timestamp cell
  var lastRow = changelog.getLastRow();
  changelog.getRange(lastRow, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");
}
