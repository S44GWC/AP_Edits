/**
 * Runs when the spreadsheet is opened.
 * This function adds a custom menu to the spreadsheet UI, allowing users
 * to easily open the sidebar and setup/reset sheets.
 */
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
       .createMenu('Edit Logger')
       .addItem('Open Sidebar', 'showSidebar')
       .addSeparator() // Add a separator for better organization
       .addItem('Setup/Reset Sheet Structure', 'setupSheetStructure') // New menu item
       .addToUi();
}

/**
 * Displays the HTML sidebar.
 * This function is called when the 'Open Sidebar' menu item is clicked.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
       .setTitle('Log a School Edit')
       .setWidth(300); // You can adjust the width as needed
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Handles the initial setup when the add-on is installed.
 * This ensures the custom menu appears immediately upon installation.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Sets up or verifies the required sheet structure for the Edit Logger.
 * This includes 'Edit Log', 'Schools List', and 'Dashboard' sheets with headers and formulas.
 */
function setupSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // 1. Setup 'Edit Log' sheet
    let editLogSheet = ss.getSheetByName('Edit Log');
    if (!editLogSheet) {
      editLogSheet = ss.insertSheet('Edit Log', 0); // Insert as the first sheet
    }
    // Updated Headers including new pause columns (K, L)
    const editLogHeaders = [
      'School Code + Name', // Column A (1)
      'Editor',             // Column B (2)
      'Shoot Style',        // Column C (3)
      'Number of Classes',  // Column D (4)
      'Re-edit?',           // Column E (5)
      'Duration',           // Column F (6) - hh:mm:ss string for actual work time
      'Status',             // Column G (7) - Open, Paused, Completed
      'Start Time',         // Column H (8)
      'End Time',           // Column I (9)
      'Numeric Duration',   // Column J (10) - Hidden helper for calculations (total work time in days)
      'Pause Start Time',   // Column K (11) - Timestamp when job was paused
      'Total Paused Duration (ms)' // Column L (12) - Accumulated pause time in milliseconds
    ];
    
    // Clear existing content in the first row before setting new headers
    editLogSheet.getRange(1, 1, 1, editLogSheet.getMaxColumns()).clearContent(); 
    editLogSheet.getRange(1, 1, 1, editLogHeaders.length).setValues([editLogHeaders]).setFontWeight('bold');
    
    // Set column widths
    editLogSheet.setColumnWidth(1, 180); // School Code + Name
    editLogSheet.setColumnWidth(2, 120); // Editor
    editLogSheet.setColumnWidth(3, 100); // Shoot Style
    editLogSheet.setColumnWidth(4, 70); // Number of Classes
    editLogSheet.setColumnWidth(5, 70); // Re-edit?
    editLogSheet.setColumnWidth(6, 100); // Duration
    editLogSheet.setColumnWidth(7, 80); // Status
    editLogSheet.setColumnWidth(8, 150); // Start Time
    editLogSheet.setColumnWidth(9, 150); // End Time
    editLogSheet.setColumnWidth(10, 120); // Numeric Duration
    editLogSheet.setColumnWidth(11, 150); // Pause Start Time
    editLogSheet.setColumnWidth(12, 100); // Total Paused Duration (ms)

    // Set ARRAYFORMULA for 'Numeric Duration' (Column J)
    // Converts 'Duration' (F) text to a numeric value (fraction of a day)
    // IFERROR handles cases where TIMEVALUE might fail (e.g., empty F:F), returning blank
    // VALUE ensures the output is treated as a number by other functions
    // This formula will only calculate for 'Completed' jobs
    editLogSheet.getRange('J2').setFormula(`=ARRAYFORMULA(IF(G2:G="Completed", IFERROR(VALUE(TIMEVALUE(F2:F)), ""), ""))`);
    
    // Hide helper columns
    editLogSheet.hideColumns(10); // Numeric Duration
    editLogSheet.hideColumns(12); // Total Paused Duration (ms)

    // Hide any columns beyond L that might exist from previous setups
    if (editLogSheet.getMaxColumns() > editLogHeaders.length) {
      editLogSheet.hideColumns(editLogHeaders.length + 1, editLogSheet.getMaxColumns() - editLogHeaders.length);
    }

    // 2. Setup 'Schools List' sheet
    let schoolsListSheet = ss.getSheetByName('Schools List');
    if (!schoolsListSheet) {
      schoolsListSheet = ss.insertSheet('Schools List', 1); // Insert as the second sheet
    }
    const schoolsListHeaders = ['School Name'];
    // Only set headers and sample data if the sheet is new or empty
    if (schoolsListSheet.getRange(1, 1).isBlank()) {
      schoolsListSheet.getRange(1, 1, 1, schoolsListHeaders.length).setValues([schoolsListHeaders]).setFontWeight('bold');
      if (schoolsListSheet.getLastRow() < 2) { // Add sample data if completely empty
        schoolsListSheet.getRange(2, 1, 3, 1).setValues([['Sample School A'], ['Sample School B'], ['Sample School C']]);
      }
    }
    schoolsListSheet.setColumnWidth(1, 250);

    // 3. Setup 'Dashboard' sheet
    let dashboardSheet = ss.getSheetByName('Dashboard');
    if (!dashboardSheet) {
      dashboardSheet = ss.insertSheet('Dashboard', 2); // Insert as the third sheet
    }
    dashboardSheet.clearContents(); // Clear existing content to set up fresh

    // --- Active Job Overview Section ---
    dashboardSheet.getRange('A1').setValue('Active Job Overview').setFontSize(20).setFontWeight('bold');
    dashboardSheet.getRange('A2').setFormula(`="Currently " & COUNTIF('Edit Log'!G:G, "Open") + COUNTIF('Edit Log'!G:G, "Paused") & " Active Job(s):"`).setFontWeight('bold');
    
    // Headers for the Query results + Elapsed Time calculation
    dashboardSheet.getRange('A3').setValue('School').setFontWeight('bold');
    dashboardSheet.getRange('B3').setValue('Editor').setFontWeight('bold');
    dashboardSheet.getRange('C3').setValue('Start Time').setFontWeight('bold');
    dashboardSheet.getRange('D3').setValue('Current Status').setFontWeight('bold'); // New Status Column
    dashboardSheet.getRange('E3').setValue('Elapsed Work Time').setFontWeight('bold'); // Adjusted to elapsed work time

    // Main QUERY formula to pull active (Open or Paused) jobs from 'Edit Log' (A, B, H, G from A:L range)
    // Uses the new ARRAYFORMULA with TRIM and LOWER for robustness
    const queryFormula = `=IFERROR(QUERY(ARRAYFORMULA({'Edit Log'!A2:F, TRIM(LWER('Edit Log'!G2:G)), 'Edit Log'!H2:L}), "SELECT Col1, Col2, Col8, Col7 WHERE Col7 = 'open' OR Col7 = 'paused' ORDER BY Col7 DESC, Col8 ASC", 0), "No active jobs found.")`;
    dashboardSheet.getRange('A4').setFormula(queryFormula);

    // ARRAYFORMULA for Elapsed Work Time - calculates live duration from Start Time (C4:C)
    // This will need to be calculated considering pause times dynamically, which is hard with pure sheet formulas.
    // We'll primarily rely on the sidebar's timer for *live* elapsed time,
    // but we can put a placeholder here or a simple difference that doesn't account for pauses.
    // For a dashboard, a static time until last refresh is usually acceptable for complex calculations.
    // A more accurate sheet formula for Elapsed Work Time would involve the Paused Duration (Col12/L).
    // Let's use a formula that pulls Start Time and Total Paused Duration from the query result and calculates.
    // This assumes the query pulls Col1, Col2, Col8, Col7. Let's adjust the query to get Col12 (Total Paused Duration (ms)).
    
    // Updated QUERY formula to also select Col12 (Total Paused Duration in ms) which is from original L column
    // The columns in the virtual array are:
    // Col1: A (School)
    // Col2: B (Editor)
    // Col3: C (Shoot Style)
    // Col4: D (Num Classes)
    // Col5: E (Re-edit?)
    // Col6: F (Duration)
    // Col7: G (Status - Trimmed & Lowercased)
    // Col8: H (Start Time)
    // Col9: I (End Time)
    // Col10: J (Numeric Duration)
    // Col11: K (Pause Start Time)
    // Col12: L (Total Paused Duration (ms))

    // Re-adjusting query to pull necessary data for Elapsed Work Time calculation on dashboard
    const queryFormulaWithPauseData = `=IFERROR(QUERY(ARRAYFORMULA({'Edit Log'!A2:F, TRIM(LOWER('Edit Log'!G2:G)), 'Edit Log'!H2:L}), "SELECT Col1, Col2, Col8, Col7, Col12 WHERE Col7 = 'open' OR Col7 = 'paused' ORDER BY Col7 DESC, Col8 ASC", 0), "No active jobs found.")`;
    dashboardSheet.getRange('A4').setFormula(queryFormulaWithPauseData);

    // ARRAYFORMULA for Elapsed Work Time (Column E on Dashboard)
    // C4:C is Start Time (Col8 from query), D4:D is Status (Col7 from query), E4:E will be Paused Duration (Col12 from query)
    // Note: The Query is outputting: School (A), Editor (B), Start Time (C), Status (D), Total Paused Duration (E)
    // So, in this context, Start Time is in C4:C, Status in D4:D, Paused Duration is in E4:E
    const elapsedTimeCalcFormula = `=ARRAYFORMULA(IF(LEN(C4:C), TEXT((NOW()-C4:C) - IF(ISNUMBER(E4:E), E4:E/86400000, 0), "hh:mm:ss"), ""))`;
    dashboardSheet.getRange('E4').setFormula(elapsedTimeCalcFormula);

    // Set column widths for Active Jobs section
    dashboardSheet.setColumnWidth(1, 200); // School
    dashboardSheet.setColumnWidth(2, 120); // Editor
    dashboardSheet.setColumnWidth(3, 160); // Start Time (Formatted)
    dashboardSheet.setColumnWidth(4, 100); // Current Status
    dashboardSheet.setColumnWidth(5, 120); // Elapsed Work Time
    dashboardSheet.getRange('A:E').setWrap(true); // Set text wrapping

    // --- Metrics Section ---
    const metricsStartRow = 7; // Start metrics below active jobs, leave a gap
    const metricsStartCol = 1; // Column A for metrics

    dashboardSheet.getRange(metricsStartRow, metricsStartCol).setValue('Key Performance Metrics').setFontSize(16).setFontWeight('bold');

    // Average Edit Time (for Completed Jobs)
    dashboardSheet.getRange(metricsStartRow + 1, metricsStartCol).setValue('Average Completed Edit Time:').setFontWeight('bold');
    dashboardSheet.getRange(metricsStartRow + 1, metricsStartCol + 1)
        .setFormula(`=IFERROR(AVERAGE(FILTER('Edit Log'!J:J, 'Edit Log'!G:G="Completed", 'Edit Log'!J:J<>"")), TIME(0,0,0))`)
        .setNumberFormat('[hh]:mm:ss'); // Format as duration

    // Number of Unique Schools Edited
    dashboardSheet.getRange(metricsStartRow + 2, metricsStartCol).setValue('Unique Schools Edited:').setFontWeight('bold');
    dashboardSheet.getRange(metricsStartRow + 2, metricsStartCol + 1)
        .setFormula(`=COUNTA(UNIQUE(FILTER('Edit Log'!A:A, 'Edit Log'!G:G="Completed", 'Edit Log'!A:A<>"")))`);

    // Schools Remaining
    dashboardSheet.getRange(metricsStartRow + 3, metricsStartCol).setValue('Schools Remaining:').setFontWeight('bold');
    dashboardSheet.getRange(metricsStartRow + 3, metricsStartCol + 1)
        .setFormula(`=COUNTA('Schools List'!A:A) - ${dashboardSheet.getRange(metricsStartRow + 2, metricsStartCol + 1).getA1Notation()}`);
    
    // --- Editor Metrics Table ---
    const editorMetricsStartRow = metricsStartRow + 6;
    dashboardSheet.getRange(editorMetricsStartRow, metricsStartCol).setValue('Editor Performance').setFontSize(16).setFontWeight('bold');
    dashboardSheet.getRange(editorMetricsStartRow + 1, metricsStartCol)
        .setFormula(`=IFERROR(QUERY('Edit Log'!A:L, "SELECT B, COUNT(B), AVG(J) WHERE G = 'Completed' GROUP BY B ORDER BY AVG(J) ASC LABEL COUNT(B) 'Completed Jobs', AVG(J) 'Avg Time' FORMAT AVG(J) 'hh:mm:ss'", 1), "No completed jobs to display editor performance.")`);
    
    // --- Shoot Style Breakdown Table ---
    const shootStyleMetricsStartRow = editorMetricsStartRow + 10;
    dashboardSheet.getRange(shootStyleMetricsStartRow, metricsStartCol).setValue('Shoot Style Breakdown').setFontSize(16).setFontWeight('bold');
    dashboardSheet.getRange(shootStyleMetricsStartRow + 1, metricsStartCol)
        .setFormula(`=IFERROR(QUERY('Edit Log'!A:L, "SELECT C, COUNT(C), AVG(J) WHERE G = 'Completed' GROUP BY C ORDER BY AVG(J) ASC LABEL COUNT(C) 'Completed Jobs', AVG(J) 'Avg Time' FORMAT AVG(J) 'hh:mm:ss'", 1), "No completed jobs to display shoot style breakdown.")`);

    ui.alert('Sheet Setup Complete', 'The "Edit Log", "Schools List", and "Dashboard" sheets have been created/updated successfully with comprehensive metrics. You can now use the sidebar!', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error Setting Up Sheets', 'An error occurred during setup: ' + e.message + '\n\nPlease ensure you have permission to create/modify sheets.', ui.ButtonSet.OK);
    console.error('Sheet Setup Error: ' + e.message, e.stack);
  }
}

/**
 * Logs a new edit job to the Google Sheet.
 * @param {Object} formData - The data from the HTML form.
 * @returns {number} The row number of the newly logged job.
 */
function logEdit(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit Log');
  if (!sheet) {
    throw new Error('Sheet "Edit Log" not found. Please create a sheet named "Edit Log" using the "Setup Sheet Structure" menu option.');
  }

  const now = new Date();
  
  // Prepare row data, ensuring order matches your 12 sheet columns (A-L):
  // A: School Code + Name
  // B: Editor
  // C: Shoot Style
  // D: Number of Classes
  // E: Re-edit?
  // F: Duration (Empty initially, calculated on job end)
  // G: Status (Set to 'Open')
  // H: Start Time (Date object, Sheets will format)
  // I: End Time (Empty initially)
  // J: Numeric Duration (calculated by sheet formula)
  // K: Pause Start Time (Empty initially)
  // L: Total Paused Duration (ms) (0 initially)
  const rowData = [
    formData.school,
    formData.editor,
    formData.style,
    parseInt(formData.classes),
    formData.reedit ? 'Yes' : 'No', // Convert boolean to 'Yes'/'No'
    '', // Duration (F)
    'Open', // Status (G)
    now, // Start Time (H)
    '', // End Time (I)
    '', // Numeric Duration (J) - handled by sheet formula
    '', // Pause Start Time (K)
    0 // Total Paused Duration (ms) (L)
  ];

  sheet.appendRow(rowData);
  const newRow = sheet.getLastRow();

  // Set the active cell in the sheet to the new job's start time for easy access/visuals
  const startTimeColumnIndex = 8; // Column H
  sheet.getRange(newRow, startTimeColumnIndex).activate();

  return newRow;
}

/**
 * Ends an active job in the Google Sheet.
 * @param {number} rowNum - The row number of the job to end.
 * @returns {string} Success message.
 */
function endJob(rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit Log');
  if (!sheet) {
    throw new Error('Sheet "Edit Log" not found. Please create a sheet named "Edit Log" using the "Setup Sheet Structure" menu option.');
  }

  // Define column indices (1-based for getRange calls)
  const statusCol = 7;             // Column G
  const endTimeCol = 9;            // Column I
  const durationCol = 6;           // Column F
  const startTimeCol = 8;          // Column H
  const pauseStartTimeCol = 11;    // Column K
  const totalPausedDurationCol = 12; // Column L

  // Get current status and all relevant times
  const rowData = sheet.getRange(rowNum, 1, 1, totalPausedDurationCol).getValues()[0];
  const currentStatus = String(rowData[statusCol - 1]).trim();
  const startTimeJsDate = rowData[startTimeCol - 1];
  let totalPausedMs = Number(rowData[totalPausedDurationCol - 1]);
  const pauseStartTimeJsDate = rowData[pauseStartTimeCol - 1];

  if (currentStatus === 'Completed') {
    throw new Error(`Job on row ${rowNum} is already completed.`);
  }

  const now = new Date();

  // If the job is currently paused, first account for the final pause segment
  if (currentStatus === 'Paused') {
    if (!(pauseStartTimeJsDate instanceof Date) || isNaN(pauseStartTimeJsDate.getTime())) {
      console.warn('Invalid pause start time for paused job on endJob. Row:', rowNum, 'Value:', pauseStartTimeJsDate);
      // Try to proceed, assuming pause duration might be missing for this last segment
    } else {
      const currentPauseSegmentMs = now.getTime() - pauseStartTimeJsDate.getTime();
      if (currentPauseSegmentMs > 0) {
        totalPausedMs += currentPauseSegmentMs;
      }
    }
    sheet.getRange(rowNum, pauseStartTimeCol).clearContent(); // Clear pause start time
  }
  
  // Validate startTimeJsDate is a valid Date object before using getTime()
  if (!(startTimeJsDate instanceof Date) || isNaN(startTimeJsDate.getTime())) {
    console.error('Invalid start time Date object for row:', rowNum, 'Value:', startTimeJsDate, 'Type:', typeof startTimeJsDate);
    throw new Error('Could not calculate duration: Invalid start time data in sheet (Column H should be a valid date/time).');
  }

  // Calculate total elapsed time, then subtract total paused time
  const totalRawElapsedMs = now.getTime() - startTimeJsDate.getTime();
  const actualWorkMs = totalRawElapsedMs - totalPausedMs;

  if (actualWorkMs < 0) {
    console.warn('Calculated actual work time is negative for row:', rowNum);
    // This could happen if pause time exceeds total elapsed, suggesting data issue.
    // Set to 0 to avoid negative duration display.
    actualWorkMs = 0; 
  }

  const totalSeconds = Math.floor(actualWorkMs / 1000);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  const formattedDuration = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;

  // Update cells (using 1-based column indices)
  sheet.getRange(rowNum, statusCol).setValue('Completed');
  sheet.getRange(rowNum, endTimeCol).setValue(now); // Set end time
  sheet.getRange(rowNum, durationCol).setValue(formattedDuration); // Set actual work duration
  sheet.getRange(rowNum, totalPausedDurationCol).setValue(totalPausedMs); // Store final total paused duration

  return 'Job ended successfully.';
}

/**
 * Pauses an active job in the Google Sheet.
 * @param {number} rowNum - The row number of the job to pause.
 * @returns {string} Success message.
 */
function pauseJob(rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit Log');
  if (!sheet) {
    throw new Error('Sheet "Edit Log" not found. Please create a sheet named "Edit Log" using the "Setup Sheet Structure" menu option.');
  }

  const statusCol = 7;          // Column G
  const pauseStartTimeCol = 11; // Column K

  const currentStatus = String(sheet.getRange(rowNum, statusCol).getValue()).trim();

  if (currentStatus === 'Paused') {
    throw new Error(`Job on row ${rowNum} is already paused.`);
  }
  if (currentStatus === 'Completed') {
    throw new Error(`Job on row ${rowNum} is already completed and cannot be paused.`);
  }

  sheet.getRange(rowNum, statusCol).setValue('Paused');
  sheet.getRange(rowNum, pauseStartTimeCol).setValue(new Date());

  return 'Job paused successfully.';
}

/**
 * Resumes a paused job in the Google Sheet.
 * @param {number} rowNum - The row number of the job to resume.
 * @returns {string} Success message.
 */
function resumeJob(rowNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit Log');
  if (!sheet) {
    throw new Error('Sheet "Edit Log" not found. Please create a sheet named "Edit Log" using the "Setup Sheet Structure" menu option.');
  }

  const statusCol = 7;             // Column G
  const pauseStartTimeCol = 11;    // Column K
  const totalPausedDurationCol = 12; // Column L

  // Read current data
  const rowData = sheet.getRange(rowNum, 1, 1, totalPausedDurationCol).getValues()[0];
  const currentStatus = String(rowData[statusCol - 1]).trim();
  const pauseStartTimeJsDate = rowData[pauseStartTimeCol - 1];
  let totalPausedMs = Number(rowData[totalPausedDurationCol - 1]);

  if (currentStatus === 'Open') {
    throw new Error(`Job on row ${rowNum} is already open.`);
  }
  if (currentStatus === 'Completed') {
    throw new Error(`Job on row ${rowNum} is already completed and cannot be resumed.`);
  }
  if (currentStatus !== 'Paused') {
     throw new Error(`Job on row ${rowNum} has an unexpected status: ${currentStatus}.`);
  }

  const now = new Date();

  // Calculate pause duration and add to total paused time
  if (!(pauseStartTimeJsDate instanceof Date) || isNaN(pauseStartTimeJsDate.getTime())) {
    console.warn('Invalid pause start time on resumeJob. Row:', rowNum, 'Value:', pauseStartTimeJsDate);
    // If pause start time is invalid, we can't calculate this segment, proceed without adding to total.
  } else {
    const currentPauseSegmentMs = now.getTime() - pauseStartTimeJsDate.getTime();
    if (currentPauseSegmentMs > 0) {
      totalPausedMs += currentPauseSegmentMs;
    }
  }

  // Update sheet
  sheet.getRange(rowNum, statusCol).setValue('Open');
  sheet.getRange(rowNum, pauseStartTimeCol).clearContent(); // Clear pause start time
  sheet.getRange(rowNum, totalPausedDurationCol).setValue(totalPausedMs); // Update total paused duration

  return 'Job resumed successfully.';
}


/**
 * Retrieves the list of schools from a designated sheet (e.g., 'Schools List').
 * Assumes schools are in the first column of that sheet.
 * @returns {string[]} An array of school names.
 */
function getSchoolsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schoolsSheet = ss.getSheetByName('Schools List'); 
  if (!schoolsSheet) {
    console.warn('Sheet "Schools List" not found. Please create a sheet named "Schools List" with schools in column A, or use the "Setup Sheet Structure" menu option.');
    return ['Sample School A1', 'Sample School B2', 'Another Test School C3']; // Fallback
  }
  const lastRow = schoolsSheet.getLastRow();
  if (lastRow < 2) return []; // Return empty array if only headers or no data
  
  // Get all values from Column A, starting from row 2 (to skip header)
  const schools = schoolsSheet.getRange('A2:A' + lastRow).getValues();
  // Flatten array, trim whitespace, filter out empty strings, and sort alphabetically
  return schools.map(row => String(row[0]).trim()).filter(String).sort();
}

/**
 * Retrieves the hardcoded list of editors.
 * @returns {string[]} An array of editor names.
 */
function getEditorsList() {
  return ['Josh', 'George'];
}

/**
 * Retrieves the hardcoded list of shoot styles.
 * @returns {string[]} An array of shoot style names.
 */
function getShootStylesList() {
  return ['Traditionals', 'Funstitch', 'Website', 'Staff Product'];
}

/**
 * Fetches current open and paused jobs from the 'Edit Log' sheet.
 * @returns {Object[]} An array of objects, each representing an active job.
 */
function getOpenJobs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit Log');
  if (!sheet || sheet.getLastRow() < 2) {
    console.log('getOpenJobs: No "Edit Log" sheet or no data found (only headers).');
    return []; // No jobs or just headers
  }

  // Define column indices (1-based for getRange calls, but used for array access as 0-based)
  // Fetch up to Column L (Total Paused Duration (ms)) which is the 12th column
  const schoolCol = 1;             // Col A (0-indexed)
  const editorCol = 2;             // Col B (1-indexed)
  const statusCol = 7;             // Col G (6-indexed)
  const startTimeCol = 8;          // Col H (7-indexed)
  const pauseStartTimeCol = 11;    // Col K (10-indexed)
  const totalPausedDurationCol = 12; // Col L (11-indexed)

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, totalPausedDurationCol).getValues();

  const activeJobs = [];
  values.forEach((row, index) => {
    // Map fetched values to their original column meanings (0-indexed array)
    const currentStatus = String(row[statusCol - 1]).trim(); // Column G
    
    // Only include jobs that are 'Open' or 'Paused'
    if (currentStatus === 'Open' || currentStatus === 'Paused') {
      const currentStartTimeJsDate = row[startTimeCol - 1]; // Column H
      const currentPauseStartTimeJsDate = row[pauseStartTimeCol - 1]; // Column K
      const currentTotalPausedDuration = Number(row[totalPausedDurationCol - 1]); // Column L

      // Ensure startTime is a valid Date object
      const startTimeMs = (currentStartTimeJsDate instanceof Date && !isNaN(currentStartTimeJsDate.getTime())) 
                         ? currentStartTimeJsDate.getTime() 
                         : null;

      // Ensure pauseStartTime is a valid Date object if status is Paused
      const pauseStartTimeMs = (currentStatus === 'Paused' && currentPauseStartTimeJsDate instanceof Date && !isNaN(currentPauseStartTimeJsDate.getTime()))
                              ? currentPauseStartTimeJsDate.getTime()
                              : null;

      activeJobs.push({
        row: index + 2, // +2 because data starts from row 2 and index is 0-based
        school: String(row[schoolCol - 1]).trim(),
        editor: String(row[editorCol - 1]).trim(),
        status: currentStatus,
        startTime: startTimeMs,
        pauseStartTime: pauseStartTimeMs,
        totalPausedDuration: currentTotalPausedDuration // Already in milliseconds
      });
    }
  });
  console.log('getOpenJobs: Found', activeJobs.length, 'active jobs:', JSON.stringify(activeJobs));
  return activeJobs;
}

/**
 * Checks the active cell in the spreadsheet to see if it's part of an active job.
 * (This function is no longer strictly used by the sidebar's primary flow,
 * as loadOpenJobsIntoSelect handles fetching all active jobs)
 * @returns {Object|null} Job details if an active job is found, otherwise null.
 */
function getJobFromActiveCell() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Edit Log');
  if (!sheet) {
    console.log('getJobFromActiveCell: No "Edit Log" sheet found.');
    return null;
  }

  const activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange || activeRange.getHeight() !== 1) {
    console.log('getJobFromActiveCell: No single row active range or multiple rows selected.');
    return null;
  }

  const row = activeRange.getRow();
  if (row < 2) {
    console.log('getJobFromActiveCell: Active cell is in header row or above.');
    return null;
  }

  // Define column indices (0-based for array access of fetched rowData)
  const schoolColIndex = 0; // Column A
  const editorColIndex = 1; // Column B
  const statusColIndex = 6; // Column G
  const startTimeColIndex = 7; // Column H
  const pauseStartTimeColIndex = 10; // Column K
  const totalPausedDurationColIndex = 11; // Column L

  // Fetch the relevant row data up to Column L (12 columns total for A-L)
  const rowData = sheet.getRange(row, 1, 1, 12).getValues()[0];

  const currentStatus = String(rowData[statusColIndex]).trim();
  if (currentStatus === 'Open' || currentStatus === 'Paused') {
    const startTimeJsDate = rowData[startTimeColIndex];
    const pauseStartTimeJsDate = rowData[pauseStartTimeColIndex];
    const totalPausedDuration = Number(rowData[totalPausedDurationColIndex]);

    if (startTimeJsDate instanceof Date && !isNaN(startTimeJsDate.getTime())) {
      const jobDetails = {
        row: row,
        school: String(rowData[schoolColIndex]).trim(),
        editor: String(rowData[editorColIndex]).trim(),
        status: currentStatus,
        startTime: startTimeJsDate.getTime(),
        pauseStartTime: (currentStatus === 'Paused' && pauseStartTimeJsDate instanceof Date && !isNaN(pauseStartTimeJsDate.getTime())) 
                        ? pauseStartTimeJsDate.getTime() 
                        : null,
        totalPausedDuration: totalPausedDuration
      };
      console.log('getJobFromActiveCell: Found active job from active cell:', JSON.stringify(jobDetails));
      return jobDetails;
    }
  }
  console.log('getJobFromActiveCell: No active job found from active cell or start time is invalid.');
  return null;
}