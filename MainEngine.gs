/**
 * File: MainEngine.gs
 * The core Delta-Time synchronisation engine. Run this via a Time-Driven Trigger.
 */

function runSyncEngine() {
  const now = new Date();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    console.error(`Sheet '${CONFIG.SHEET_NAME}' not found.`);
    return;
  }

  // --- 1. PULL ACTIVE TRELLO DATA ---
  const activeListIds = [
    CONFIG.LISTS.EDITING, CONFIG.LISTS.EDIT_PAUSED, 
    CONFIG.LISTS.PAUSED_PROOF, CONFIG.LISTS.PROOFING, CONFIG.LISTS.RE_EDITS
  ];
  
  const activeCardsMap = {}; 
  activeListIds.forEach(listId => {
    const cards = fetchCardsInList(listId);
    cards.forEach(c => { activeCardsMap[c.id] = { listId: listId, data: c }; });
  });

  // --- 2. UPDATE EXISTING ROWS (IN-MEMORY BATCH PROCESSING) ---
  const lastRow = sheet.getLastRow();
  let existingCardIds = new Set();
  
  if (lastRow > 1) {
    // Grab all 15 columns (Indices 0 through 14)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 15);
    const sheetData = dataRange.getValues();
    
    sheetData.forEach(row => {
      const cardId = row[11]; // Column L (Index 11)
      if (!cardId || String(cardId).trim() === '') return; 
      
      existingCardIds.add(cardId);
      const status = row[7]; // Current Status (Index 7)
      if (status === 'Completed') return; 

      const lastChecked = row[10] instanceof Date ? row[10] : now;
      const deltaMs = now.getTime() - lastChecked.getTime();
      
      let editMs = Number(row[8]) || 0; // Edit Time (Index 8)
      let proofMs = Number(row[9]) || 0; // Proof Time (Index 9)
      let newStatus = status;
      let reEditFlag = row[4]; // Re-Edit? (Index 4)

      if (activeCardsMap[cardId]) {
        const currentList = activeCardsMap[cardId].listId;
        
        // Self-Healing Start Time (Checks Index 5)
        if (!row[5] || String(row[5]).trim() === '') {
          row[5] = now; 
        }

        if (currentList === CONFIG.LISTS.EDITING) {
          newStatus = 'Editing';
          editMs += deltaMs;
        } else if (currentList === CONFIG.LISTS.EDIT_PAUSED) {
          newStatus = 'Edit Paused';
        } else if (currentList === CONFIG.LISTS.PAUSED_PROOF) {
          newStatus = 'Paused for Proofing';
        } else if (currentList === CONFIG.LISTS.PROOFING) {
          newStatus = 'Proofing';
          proofMs += deltaMs; 
        } else if (currentList === CONFIG.LISTS.RE_EDITS) {
          newStatus = 'Re-Edits Required';
          reEditFlag = 'Yes'; 
        }

        // Apply changes to the array
        row[4] = reEditFlag;
        row[7] = newStatus;
        row[8] = editMs;
        row[9] = proofMs;
        row[10] = now; 

      } else {
        // Card missing from active. Did it hit the Admin Board?
        if (isCardCompleted(cardId)) {
          row[7] = 'Completed'; // Status
          row[6] = now;         // End Time (Index 6)
          row[10] = now;        // Last Checked
        }
      }

      // Auto-format readable strings
      row[12] = formatMsToReadable(editMs);              // Formatted Edit
      row[13] = formatMsToReadable(proofMs);             // Formatted Proof
      row[14] = formatMsToReadable(editMs + proofMs);    // Formatted Total
    });

    dataRange.setValues(sheetData);
  }

  // --- 3. INGEST NEW JOBS ---
  const editingCards = fetchCardsInList(CONFIG.LISTS.EDITING);
  const newJobsToAppend = [];

  editingCards.forEach(card => {
    if (!existingCardIds.has(card.id)) {
      const meta = extractCardMetadata(card);
      
      if (meta.editor !== "Unassigned") {
        console.log(`Ingesting New Job: ${card.name}`);
        
        // Array matches the 15-column layout perfectly
        newJobsToAppend.push([
          meta.school,        // 0: School
          meta.editor,        // 1: Editor
          meta.style,         // 2: Style
          meta.classes,       // 3: Classes
          'No',               // 4: Re-Edit?
          now,                // 5: Start Time
          "",                 // 6: End Time (Blank on ingest)
          'Editing',          // 7: Status
          0,                  // 8: Edit Time (ms)
          0,                  // 9: Proof Time (ms)
          now,                // 10: Last Checked
          card.id,            // 11: Card ID
          "00:00:00",         // 12: Formatted Edit
          "00:00:00",         // 13: Formatted Proof
          "00:00:00"          // 14: Formatted Total
        ]);
      }
    }
  });

  if (newJobsToAppend.length > 0) {
    const startRow = lastRow === 0 ? 2 : lastRow + 1;
    sheet.getRange(startRow, 1, newJobsToAppend.length, 15).setValues(newJobsToAppend);
  }
}

function formatMsToReadable(ms) {
  if (!ms || ms <= 0) return "00:00:00";
  const totalSeconds = Math.floor(ms / 1000);
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);
  const s = totalSeconds % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
}
