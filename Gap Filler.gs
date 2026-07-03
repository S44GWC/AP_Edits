function runGapFiller() {
  const UI = SpreadsheetApp.getUi();
  const KEY = '74a83598cab63d8a5c9113b86df26955'; 
  const TOKEN = '';
  const EDITING_LIST_ID = '685d4e3ad7103a8995464035'; 

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('System Data');
  const lastRow = sheet.getLastRow();
  
  // Get all Card IDs (Col L) and Start Times (Col F) in one go
  const idRange = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
  const timeRange = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
  
  let updatesCount = 0;

  for (let i = 0; i < idRange.length; i++) {
    const cardId = idRange[i][0];
    const startTime = timeRange[i][0];

    // Only process if we have an ID and Start Time is empty
    if (cardId && !startTime) {
      const historyUrl = `https://api.trello.com/1/cards/${cardId}/actions?key=${KEY}&token=${TOKEN}&filter=all&limit=100`;
      const actions = fetchTrelloWithBackoff(historyUrl);

      if (actions) {
        const startAction = actions.find(a => (a.data.listAfter?.id === EDITING_LIST_ID) || (a.data.list?.id === EDITING_LIST_ID));
        
        if (startAction) {
          const newStartTime = new Date(startAction.date);
          // Write to Column F (Col 6), row i+2 because we started at row 2
          sheet.getRange(i + 2, 6).setValue(newStartTime);
          updatesCount++;
        }
      }
      Utilities.sleep(150);
    }
  }

  UI.alert('Gap Filler', `Finished! Filled ${updatesCount} missing Start Times in Column F.`, UI.ButtonSet.OK);
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

function testHistoryFetch() {
  const CARD_ID = '69e0b899e9d6e94a5ed1d5ad';
  const KEY = '74a83598cab63d8a5c9113b86df26955';
  const TOKEN = '';
  const url = `https://api.trello.com/1/cards/${CARD_ID}/actions?key=${KEY}&token=${TOKEN}&filter=all&limit=100`;
  
  const response = UrlFetchApp.fetch(url);
  Logger.log(response.getContentText());
}
