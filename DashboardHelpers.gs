/**
 * File: DashboardHelpers.gs
 * Contains one-off maintenance, migration, and dashboard formatting tools.
 * These do not need to run on a timer; run them manually from the editor when needed.
 */

function repairMathColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit Log');
  if (!sheet) return;
  
  // Re-injects the master math formulas into Row 2
  sheet.getRange('J2').setFormula(`=ARRAYFORMULA(IF(lower(G2:G)="completed", IFERROR(VALUE(F2:F), IFERROR(1/0)), IFERROR(1/0)))`);
  sheet.getRange('M2').setFormula(`=ARRAYFORMULA(IF(lower(G2:G)="completed", IFERROR(J2:J / D2:D, IFERROR(1/0)), IFERROR(1/0)))`);
  
  SpreadsheetApp.getUi().alert("Math Columns Repaired! Check your Dashboard.");
}

function updateDashboardLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');
  if (!dashboardSheet) return;

  // --- 1. REMOVE 'ELAPSED TIME' (COLUMN F) ---
  // Update Headers to only 4 columns
  const activeHeaders = [['School', 'Editor', 'Start Time', 'Status']];
  dashboardSheet.getRange('B4:E4').setValues(activeHeaders)
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Clear the old F and G columns entirely
  dashboardSheet.getRange('F4:G15').clear();

  // Re-apply borders to just B through E
  dashboardSheet.getRange('B5:E15').setBorder(true, true, true, true, true, true, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
  dashboardSheet.getRange('B5:E15').setHorizontalAlignment('center');

  // --- 2. RELAX THE ESTIMATES QUERY ---
  // We use 'contains' to catch accidental trailing spaces (e.g. "Completed ")
  const pivotRow = 18;
  const estRow = pivotRow + 15;
  
  const estFormula = `=IFERROR(QUERY('Edit Log'!A:M, "SELECT C, COUNT(C), AVG(J), AVG(M) WHERE lower(G) contains 'completed' AND C != '' GROUP BY C ORDER BY AVG(J) DESC LABEL C 'Style', COUNT(C) 'Jobs', AVG(J) 'Avg Job', AVG(M) 'Avg Class' FORMAT AVG(J) '[hh]:mm:ss', AVG(M) '[hh]:mm:ss'", 1), "Still no data. Check Column C!")`;
  
  dashboardSheet.getRange(estRow + 2, 2).setFormula(estFormula);
}

function fixEstimatesQuery() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');
  if (!dashboardSheet) return;

  const pivotRow = 18;
  const estRow = pivotRow + 15;
  
  // THE FIX: Added "AND J > 0" to force the query to only look at pure numbers and ignore text/errors.
  const estFormula = `=IFERROR(QUERY('Edit Log'!A:M, "SELECT C, COUNT(C), AVG(J), AVG(M) WHERE lower(G) contains 'completed' AND C != '' AND J > 0 GROUP BY C ORDER BY AVG(J) DESC LABEL C 'Style', COUNT(C) 'Jobs', AVG(J) 'Avg Job', AVG(M) 'Avg Class' FORMAT AVG(J) '[hh]:mm:ss', AVG(M) '[hh]:mm:ss'", 1), "Still no data. Math columns might be empty!")`;
  
  dashboardSheet.getRange(estRow + 2, 2).setFormula(estFormula);
}

function migrateOldData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheet = ss.getSheetByName('Edit Log');
  const newSheet = ss.getSheetByName('System Data');

  if (!oldSheet || !newSheet) {
    SpreadsheetApp.getUi().alert("Couldn't find the sheets!");
    return;
  }

  const oldData = oldSheet.getDataRange().getValues();
  let migratedCount = 0;

  // Loop through old data (skipping the header row)
  for (let i = 1; i < oldData.length; i++) {
    const row = oldData[i];
    const status = String(row[6]).toLowerCase().trim();

    // Only migrate the 44 actually completed jobs
    if (status === 'completed') {
       const school = row[0];
       const editor = row[1];
       const style = row[2];
       const classes = row[3];
       const reedit = row[4];
       
       // Old Column J (Index 9) holds the numeric time fraction. 
       // We multiply by 86,400,000 to turn it into Milliseconds for the new engine.
       const numericDuration = Number(row[9]) || 0; 
       const totalMs = numericDuration * 86400000;

       // Append to the new System Data sheet
       // Format: School, Editor, Style, Classes, Re-Edit?, Status, Edit(ms), Proof(ms), LastChecked, CardID
       newSheet.appendRow([
         school, editor, style, classes, reedit, 
         'Completed', totalMs, "", new Date(), "migrated_historical_data"
       ]);
       
       migratedCount++;
    }
  }
  
  SpreadsheetApp.getUi().alert(`Success! Migrated ${migratedCount} completed jobs to the new System Data sheet.`);
}
