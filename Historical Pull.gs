/**
 * ONE-OFF HELPER: Admin Board Data Recovery
 * Scans all Admin Board lists, captures history across board movements,
 * handles rate limits, and maps perfectly to your 15-column System Data sheet.
 */
function runAdminBoardRecovery() {
  const UI = SpreadsheetApp.getUi();
  
  // --- CONFIGURATION ---
  const KEY = '74a83598cab63d8a5c9113b86df26955'; 
  const TOKEN = '';
  const KNOWN_EDITORS = ['Josh', 'George', 'Liv', 'Sarah'];
  const EDITING_LIST_ID = '685d4e3ad7103a8995464035'; 
  const EDITING_BOARD_ID = '64c24b08a69b6f4904a5ba1b';
  const REEDIT_LIST_ID = '662131782d8e2d0b64737e0c';

  const TARGET_LIST_IDS = [
    '6928242a2d58e16ad170298a', '6928242a2d58e16ad170298b', '6928242a2d58e16ad170298c', 
    '6928242a2d58e16ad170298d', '6928242a2d58e16ad170299d', '6928242a2d58e16ad170298e', 
    '6928242a2d58e16ad170298f', '6928242a2d58e16ad1702993', '6928242a2d58e16ad1702992', 
    '6928242a2d58e16ad1702991', '6928242a2d58e16ad1702999', '6928242a2d58e16ad1702994'
  ];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('System Data');
  if (!sheet) return UI.alert('Error: Could not find System Data sheet.');

  // --- 1. BUILD DE-DUPE BLOCKLIST ---
  const lastRow = sheet.getLastRow();
  let existingCardIds = new Set();
  if (lastRow > 1) {
    sheet.getRange(2, 12, lastRow - 1, 1).getValues().forEach(row => {
      if (row[0]) existingCardIds.add(String(row[0]).trim());
    });
  }

  // --- 2. FETCH MISSING CARDS ---
  let allTargetCards = [];
  TARGET_LIST_IDS.forEach(listId => {
    const data = fetchTrelloWithBackoff(`https://api.trello.com/1/lists/${listId}/cards?key=${KEY}&token=${TOKEN}&fields=name,labels,desc`);
    if (data) allTargetCards = allTargetCards.concat(data);
  });

  const missingCards = allTargetCards.filter(card => !existingCardIds.has(card.id));
  if (missingCards.length === 0) return UI.alert('Complete', 'No missing cards found.', UI.ButtonSet.OK);

  // --- 3. RECOVERY LOOP ---
  const recoveredRows = [];
  missingCards.forEach(card => {
    const parts = card.name.split(' - ');
    const school = parts[0] ? parts[0].trim() : card.name;
    const style = parts.length > 1 ? parts[1].trim() : "Standard";
    
    let editor = "Unassigned";
    if (card.labels) {
      const found = card.labels.filter(l => KNOWN_EDITORS.includes(l.name)).map(l => l.name);
      if (found.length > 0) editor = found.join(", ");
    }

    let classes = 1;
    if (card.desc) {
      const match = card.desc.match(/Number of Groups\D*(\d+(\.\d+)?)/i);
      if (match && match[1]) classes = Number(match[1]);
    }

    let startTime = "", endTime = "", isReedit = 'No';
    const actions = fetchTrelloWithBackoff(`https://api.trello.com/1/cards/${card.id}/actions?key=${KEY}&token=${TOKEN}&filter=all&limit=100`);
    
    if (actions) {
      // Find Start Time: Captures standard list moves OR board moves from Editing Board
      const startAction = actions.find(a => 
        (a.data.listAfter?.id === EDITING_LIST_ID) || 
        (a.type === 'moveCardToBoard' && a.data.boardSource?.id === EDITING_BOARD_ID)
      );
      if (startAction) startTime = new Date(startAction.date);

      // Find End Time: Arrival at Admin
      const endAction = actions.find(a => {
        const lid = (a.data.listAfter?.id) || (a.data.list?.id);
        return lid === '6928242a2d58e16ad1702989' || lid === '6928242a2d58e16ad170298a';
      });
      if (endAction) endTime = new Date(endAction.date);

      if (actions.find(a => (a.data.listAfter?.id === REEDIT_LIST_ID) || (a.data.list?.id === REEDIT_LIST_ID))) isReedit = 'Yes';
    }

    let totalMs = 0;
    if (startTime instanceof Date && endTime instanceof Date) totalMs = Math.max(0, endTime.getTime() - startTime.getTime());

    const h = Math.floor(totalMs / 3600000);
    const m = Math.floor((totalMs % 3600000) / 60000);
    const s = Math.floor((totalMs % 60000) / 1000);
    const formattedStr = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;

    recoveredRows.push([school, editor, style, classes, isReedit, startTime, endTime, 'Completed', totalMs, 0, new Date(), card.id, formattedStr, "00:00:00", formattedStr]);
  });

  if (recoveredRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, recoveredRows.length, 15).setValues(recoveredRows);
    UI.alert('Recovery Complete', `Recovered ${recoveredRows.length} jobs.`, UI.ButtonSet.OK);
  }
}

function fetchTrelloWithBackoff(url) {
  for (let i = 0; i < 3; i++) {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) return JSON.parse(response.getContentText());
    if (response.getResponseCode() === 429) {
      Utilities.sleep(Math.pow(2, i) * 1000);
    } else break;
  }
  return null;
}
