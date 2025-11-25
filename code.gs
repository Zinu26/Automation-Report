/**
 * @OnlyCurrentDoc
 */

// --- CONFIGURATION ---
const MAIN_SHEET_NAME = 'Sheet1';
const EXCLUDED_COLUMN_INDEX = 10; // Column K (0-based index)

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sheet Organizer')
    .addItem('Sync All Sheets', 'runFullSync')
    .addToUi();
}

function runFullSync() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (_) {
    ui = null; // Non-UI context (trigger/webhook)
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);

  if (!mainSheet) {
    ui ? ui.alert(`Main sheet "${MAIN_SHEET_NAME}" not found.`)
       : console.error(`Main sheet missing.`);
    return;
  }

  try {
    if (ui) ui.alert("Sync started. Please wait.");

    const allData = mainSheet.getDataRange().getValues();
    const header = allData.shift();
    const newHeader = header.filter((_, i) => i !== EXCLUDED_COLUMN_INDEX);

    const groups = {};
    allData.forEach((row, i) => {
      const groupCode = row[0];
      if (!groupCode) return;

      if (!groups[groupCode]) groups[groupCode] = [];
      groups[groupCode].push({
        filteredRow: row.filter((_, idx) => idx !== EXCLUDED_COLUMN_INDEX),
        originalRow: i + 2
      });
    });

    const existingSheets = ss.getSheets().map(s => s.getName());
    const groupNames = Object.keys(groups);

    existingSheets.forEach(name => {
      if (name !== MAIN_SHEET_NAME && !groupNames.includes(name)) {
        ss.getSheetByName(name).clear();
      }
    });

    groupNames.forEach(groupCode => {
      let sheet = ss.getSheetByName(groupCode);
      if (!sheet) sheet = ss.insertSheet(groupCode);

      sheet.clear();
      sheet.getRange(1, 1, 1, newHeader.length).setValues([newHeader]);

      const values = groups[groupCode].map(g => g.filteredRow);
      if (values.length > 0) {
        sheet.getRange(2, 1, values.length, newHeader.length).setValues(values);
      }

      for (let col = 0; col < newHeader.length; col++) {
        const originalCol = col < EXCLUDED_COLUMN_INDEX ? col : col + 1;
        mainSheet.getRange(1, originalCol + 1).copyTo(
          sheet.getRange(1, col + 1),
          { formatOnly: true }
        );
        sheet.setColumnWidth(col + 1, mainSheet.getColumnWidth(originalCol + 1));
      }
    });

    SpreadsheetApp.flush();
    ui ? ui.alert("Success! All sheets synchronized.") 
       : console.log("Sync completed.");

  } catch (e) {
    const msg = "Sync failed: " + e.message;
    ui ? ui.alert(msg) : console.error(msg);
  }
}
