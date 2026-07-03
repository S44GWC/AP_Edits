/**
 * File: TrelloAPI.gs
 * Handles all outgoing requests to the Trello API and metadata extraction.
 */

function fetchCardsInList(listId) {
  if (!listId || listId === 'PASTE_ID_HERE') return [];
  
  const url = `https://api.trello.com/1/lists/${listId}/cards?key=${CONFIG.TRELLO_KEY}&token=${CONFIG.TRELLO_TOKEN}&fields=name,labels,desc`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }
  } catch (e) {
    console.warn(`Failed to fetch list ${listId}: ${e.message}`);
  }
  return [];
}

function isCardCompleted(cardId) {
  const url = `https://api.trello.com/1/cards/${cardId}?key=${CONFIG.TRELLO_KEY}&token=${CONFIG.TRELLO_TOKEN}&fields=idList,closed`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const cardData = JSON.parse(response.getContentText());
      return (cardData.idList === CONFIG.LISTS.DONE || cardData.closed === true);
    }
  } catch (e) {
    console.warn(`Could not verify completion for card ${cardId}: ${e.message}`);
  }
  return false;
}

function extractCardMetadata(card) {
  const parts = card.name.split(' - ');
  const school = parts[0] ? parts[0].trim() : card.name;
  const style = parts.length > 1 ? parts[1].trim() : "Standard";
  
  let editor = "Unassigned";
  let foundEditors = [];
  
  if (card.labels) {
    card.labels.forEach(l => { 
      if (CONFIG.KNOWN_EDITORS.includes(l.name)) foundEditors.push(l.name); 
    });
  }
  if (foundEditors.length > 0) editor = foundEditors.join(", ");

  let classes = 1;
  if (card.desc) {
    const match = card.desc.match(/Number of Groups\D*(\d+(\.\d+)?)/i);
    if (match && match[1]) classes = Number(match[1]);
  }

  return { school, editor, style, classes };
}
