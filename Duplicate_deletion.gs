function removeDuplicatesColumn2Priority() {
  // ======= USER SETTINGS =======
  const sheetName = "All Companies"; // <-- Change to your subsheet name
  const columnCheckIndex = 2;    // Which column to use for duplicate detection (B = 2)
  const colorCheckColumn = 2;    // Which column's color to use for priority comparison
  // =============================

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`The sheet "${sheetName}" was not found!`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No data found to process.");
    return;
  }

  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();

  // --- Define Color Priority ---
  const COLOR_PRIORITY = {
    "#00ff00": 1, // Green - highest
    "#0000ff": 2, // Blue
    "#ffff00": 3  // Yellow
  };
  const DEFAULT_PRIORITY = 4; // For any other color

  let uniqueMap = {};
  let rowsToDelete = [];

  // Loop through all rows
  values.forEach((row, i) => {
    const key = row[columnCheckIndex - 1]; // Only Column 2 value
    const rowColor = backgrounds[i][colorCheckColumn - 1].toLowerCase();
    const priority = COLOR_PRIORITY[rowColor] || DEFAULT_PRIORITY;

    if (!(key in uniqueMap)) {
      // First time seeing this value
      uniqueMap[key] = { index: i, priority: priority };
    } else {
      // Already exists — check priority
      if (priority < uniqueMap[key].priority) {
        // New row has better priority; delete old
        rowsToDelete.push(uniqueMap[key].index + 2); // Stored old row index
        uniqueMap[key] = { index: i, priority: priority };
      } else {
        // Keep old, delete current
        rowsToDelete.push(i + 2);
      }
    }
  });

  // Sort rows descending so deletion does not shift indexes
  rowsToDelete = [...new Set(rowsToDelete)];
  rowsToDelete.sort((a, b) => b - a);
  rowsToDelete.forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });

  SpreadsheetApp.getUi().alert(`Duplicate removal complete. Deleted ${rowsToDelete.length} rows from "${sheetName}" based on Column ${columnCheckIndex}.`);
}
