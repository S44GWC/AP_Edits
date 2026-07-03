/**
 * ONE-OFF HELPER: Map Workbook Structure (Grid Format)
 * Scans all sheets. Places Sheet Names in Row 1 as headers,
 * and lists the actual sheet headers vertically down the columns.
 */
function mapWorkbookStructureGrid() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');

  // 1. Create or clear the Config sheet
  if (!configSheet) {
    configSheet = ss.insertSheet('Config', 0);
  } else {
    configSheet.clear();
  }

  const sheets = ss.getSheets();
  let columnIndex = 1;

  // 2. Loop through all sheets and extract headers
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Skip the Config sheet to avoid an endless recursive loop
    if (sheetName === 'Config') return;

    // Set the Sheet Name as the top header for this column
    configSheet.getRange(1, columnIndex)
      .setValue(sheetName)
      .setFontWeight('bold')
      .setBackground('#202124')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');

    const lastCol = sheet.getLastColumn();
    
    if (lastCol > 0) {
      // Grab the first row up to the last column with data
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

      // Filter out any blank cells
      const cleanedHeaders = headers.filter(String);

      if (cleanedHeaders.length > 0) {
        // Convert the horizontal 1D array into a vertical 2D array
        const verticalHeaders = cleanedHeaders.map(header => [header]);
        
        // Paste the headers vertically underneath the Sheet Name
        configSheet.getRange(2, columnIndex, verticalHeaders.length, 1).setValues(verticalHeaders);
      } else {
        configSheet.getRange(2, columnIndex).setValue('[No Headers]').setFontStyle('italic').setFontColor('#5f6368');
      }
    } else {
      configSheet.getRange(2, columnIndex).setValue('[Empty Sheet]').setFontStyle('italic').setFontColor('#5f6368');
    }

    // Move to the next column for the next sheet
    columnIndex++;
  });

  // 3. Tidy up the column widths for legibility
  if (columnIndex > 1) {
    configSheet.autoResizeColumns(1, columnIndex - 1);
  }

  // Alert the user upon completion
  SpreadsheetApp.getUi().alert('Success', 'Workbook structure mapped! The Config sheet has been formatted as a grid.', SpreadsheetApp.getUi().ButtonSet.OK);
}
