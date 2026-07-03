/**
 * BUILDS THE V2 MULTI-PHASE DASHBOARD (SheetSetup.gs)
 * Creates a clean backend log and a highly professional, legible front-end dashboard.
 */
function buildProDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // --- 1. BUILD BACKEND (The Engine Room) ---
    let backend = ss.getSheetByName('System Data');
    if (!backend) {
      backend = ss.insertSheet('System Data', ss.getNumSheets());
    }
    
    // We don't want to clear the backend if it already has data, 
    // so we only write headers if the sheet is blank (Row 1 is empty).
    if (backend.getRange('A1').isBlank()) {
      const backendHeaders = [
        'School', 'Editor', 'Style', 'Classes', 'Re-Edit?', 
        'Current Status', 'Edit Time (ms)', 'Proof Time (ms)', 
        'Last Checked', 'Card ID', 'Formatted Edit', 'Formatted Proof', 'Total Time'
      ];
      backend.getRange(1, 1, 1, backendHeaders.length).setValues([backendHeaders]).setFontWeight('bold');
      
      // FIX: Output raw numbers (fractions of a day), not text strings, so QUERY can average them later.
      backend.getRange('K2:K').setFormula(`=ARRAYFORMULA(IF(G2:G<>"", G2:G/86400000, ""))`);
      backend.getRange('L2:L').setFormula(`=ARRAYFORMULA(IF(H2:H<>"", H2:H/86400000, ""))`);
      backend.getRange('M2:M').setFormula(`=ARRAYFORMULA(IF((G2:G<>"")+(H2:H<>""), (G2:G+H2:H)/86400000, ""))`);
      
      // Apply native duration formatting to the columns instead of using TEXT()
      backend.getRange('K2:M').setNumberFormat('[hh]:mm:ss');
      
      backend.setColumnWidths(7, 4, 80); // Shrink math columns to keep it tidy
    }

    // --- 2. BUILD FRONTEND DASHBOARD ---
    let dash = ss.getSheetByName('Production Dashboard');
    if (dash) {
      dash.clear();
      dash.getCharts().forEach(c => dash.removeChart(c));
    } else {
      dash = ss.insertSheet('Production Dashboard', 0);
    }

    dash.setHiddenGridlines(true);
    dash.setColumnWidth(1, 20); // Spacer
    
    // TITLE
    dash.getRange('B2').setValue('PRODUCTION OVERVIEW').setFontSize(22).setFontWeight('bold').setFontColor('#1a73e8');
    dash.getRange('B3').setFormula(`="Last updated: " & TEXT(NOW(), "dd/MM/yyyy hh:mm AM/PM")`).setFontColor('#5f6368').setFontStyle('italic');

    // === SECTION A: ACTIVE PIPELINE ===
    dash.getRange('B5').setValue('Live Pipeline').setFontSize(14).setFontWeight('bold').setFontColor('#202124');
    
    const activeHeaders = [['School', 'Editor', 'Phase / Status', 'Re-Edit', 'Edit Time', 'Proof Time']];
    dash.getRange('B7:G7').setValues(activeHeaders)
      .setBackground('#202124').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

    // Query for anything NOT completed
    dash.getRange('B8').setFormula(`=IFERROR(QUERY('System Data'!A2:M, "SELECT A, B, F, E, K, L WHERE lower(F) != 'completed' AND A is not null ORDER BY F ASC", 0), "No active jobs.")`);
    
    // Ensure the imported time columns in the dashboard are formatted properly
    dash.getRange('F8:G20').setNumberFormat('[hh]:mm:ss');
    dash.getRange('B7:G20').setBorder(true, true, true, true, true, true, '#daddad', SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment('center');
    
    dash.setColumnWidth(2, 250); // School
    dash.setColumnWidth(3, 100); // Editor
    dash.setColumnWidth(4, 180); // Status
    dash.setColumnWidth(5, 80);  // Re-edit
    
    // === SECTION B: PERFORMANCE METRICS ===
    const metricsRow = 23;
    dash.getRange(metricsRow, 2).setValue('Style Performance (Averages)').setFontSize(14).setFontWeight('bold').setFontColor('#202124');

    const perfHeaders = [['Style', 'Jobs', 'Avg Edit Time', 'Avg Proof Time', 'Avg Total Time']];
    dash.getRange(metricsRow + 2, 2, 1, 5).setValues(perfHeaders)
      .setBackground('#1a73e8').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

    // FIX: Query the natively formatted K, L, M columns instead of trying to average raw milliseconds (G, H)
    const perfFormula = `=IFERROR(QUERY('System Data'!A2:M, "SELECT C, COUNT(C), AVG(K), AVG(L), AVG(M) WHERE lower(F)='completed' AND C is not null GROUP BY C LABEL C '', COUNT(C) '', AVG(K) '', AVG(L) '', AVG(M) '' FORMAT AVG(K) '[hh]:mm:ss', AVG(L) '[hh]:mm:ss', AVG(M) '[hh]:mm:ss'", 0), "No completed data.")`;
    dash.getRange(metricsRow + 3, 2).setFormula(perfFormula);
    
    // Apply formatting to the query results
    dash.getRange(metricsRow + 3, 4, 7, 3).setNumberFormat('[hh]:mm:ss');
    dash.getRange(metricsRow + 2, 2, 8, 5).setBorder(true, true, true, true, true, true, '#daddad', SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment('center');

    ui.alert('Success', 'V2 Dashboard built! You can hide the System Data sheet if you prefer.', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}/**
 * BUILDS THE V2 MULTI-PHASE DASHBOARD (SheetSetup.gs)
 * Creates a clean backend log and a highly professional, legible front-end dashboard.
 */
function buildProDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // --- 1. BUILD BACKEND (The Engine Room) ---
    let backend = ss.getSheetByName('System Data');
    if (!backend) {
      backend = ss.insertSheet('System Data', ss.getNumSheets());
    }
    
    // We don't want to clear the backend if it already has data, 
    // so we only write headers if the sheet is blank (Row 1 is empty).
    if (backend.getRange('A1').isBlank()) {
      const backendHeaders = [
        'School', 'Editor', 'Style', 'Classes', 'Re-Edit?', 
        'Current Status', 'Edit Time (ms)', 'Proof Time (ms)', 
        'Last Checked', 'Card ID', 'Formatted Edit', 'Formatted Proof', 'Total Time'
      ];
      backend.getRange(1, 1, 1, backendHeaders.length).setValues([backendHeaders]).setFontWeight('bold');
      
      // FIX: Output raw numbers (fractions of a day), not text strings, so QUERY can average them later.
      backend.getRange('K2:K').setFormula(`=ARRAYFORMULA(IF(G2:G<>"", G2:G/86400000, ""))`);
      backend.getRange('L2:L').setFormula(`=ARRAYFORMULA(IF(H2:H<>"", H2:H/86400000, ""))`);
      backend.getRange('M2:M').setFormula(`=ARRAYFORMULA(IF((G2:G<>"")+(H2:H<>""), (G2:G+H2:H)/86400000, ""))`);
      
      // Apply native duration formatting to the columns instead of using TEXT()
      backend.getRange('K2:M').setNumberFormat('[hh]:mm:ss');
      
      backend.setColumnWidths(7, 4, 80); // Shrink math columns to keep it tidy
    }

    // --- 2. BUILD FRONTEND DASHBOARD ---
    let dash = ss.getSheetByName('Production Dashboard');
    if (dash) {
      dash.clear();
      dash.getCharts().forEach(c => dash.removeChart(c));
    } else {
      dash = ss.insertSheet('Production Dashboard', 0);
    }

    dash.setHiddenGridlines(true);
    dash.setColumnWidth(1, 20); // Spacer
    
    // TITLE
    dash.getRange('B2').setValue('PRODUCTION OVERVIEW').setFontSize(22).setFontWeight('bold').setFontColor('#1a73e8');
    dash.getRange('B3').setFormula(`="Last updated: " & TEXT(NOW(), "dd/MM/yyyy hh:mm AM/PM")`).setFontColor('#5f6368').setFontStyle('italic');

    // === SECTION A: ACTIVE PIPELINE ===
    dash.getRange('B5').setValue('Live Pipeline').setFontSize(14).setFontWeight('bold').setFontColor('#202124');
    
    const activeHeaders = [['School', 'Editor', 'Phase / Status', 'Re-Edit', 'Edit Time', 'Proof Time']];
    dash.getRange('B7:G7').setValues(activeHeaders)
      .setBackground('#202124').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

    // Query for anything NOT completed
    dash.getRange('B8').setFormula(`=IFERROR(QUERY('System Data'!A2:M, "SELECT A, B, F, E, K, L WHERE lower(F) != 'completed' AND A is not null ORDER BY F ASC", 0), "No active jobs.")`);
    
    // Ensure the imported time columns in the dashboard are formatted properly
    dash.getRange('F8:G20').setNumberFormat('[hh]:mm:ss');
    dash.getRange('B7:G20').setBorder(true, true, true, true, true, true, '#daddad', SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment('center');
    
    dash.setColumnWidth(2, 250); // School
    dash.setColumnWidth(3, 100); // Editor
    dash.setColumnWidth(4, 180); // Status
    dash.setColumnWidth(5, 80);  // Re-edit
    
    // === SECTION B: PERFORMANCE METRICS ===
    const metricsRow = 23;
    dash.getRange(metricsRow, 2).setValue('Style Performance (Averages)').setFontSize(14).setFontWeight('bold').setFontColor('#202124');

    const perfHeaders = [['Style', 'Jobs', 'Avg Edit Time', 'Avg Proof Time', 'Avg Total Time']];
    dash.getRange(metricsRow + 2, 2, 1, 5).setValues(perfHeaders)
      .setBackground('#1a73e8').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

    // FIX: Query the natively formatted K, L, M columns instead of trying to average raw milliseconds (G, H)
    const perfFormula = `=IFERROR(QUERY('System Data'!A2:M, "SELECT C, COUNT(C), AVG(K), AVG(L), AVG(M) WHERE lower(F)='completed' AND C is not null GROUP BY C LABEL C '', COUNT(C) '', AVG(K) '', AVG(L) '', AVG(M) '' FORMAT AVG(K) '[hh]:mm:ss', AVG(L) '[hh]:mm:ss', AVG(M) '[hh]:mm:ss'", 0), "No completed data.")`;
    dash.getRange(metricsRow + 3, 2).setFormula(perfFormula);
    
    // Apply formatting to the query results
    dash.getRange(metricsRow + 3, 4, 7, 3).setNumberFormat('[hh]:mm:ss');
    dash.getRange(metricsRow + 2, 2, 8, 5).setBorder(true, true, true, true, true, true, '#daddad', SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment('center');

    ui.alert('Success', 'V2 Dashboard built! You can hide the System Data sheet if you prefer.', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}
