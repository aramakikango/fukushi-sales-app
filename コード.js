function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function addRecord(data) {
  var ss = SpreadsheetApp.openById('1bTRSe5l7RTMk1taHNtYaAUMcFBEIwGUf6Yz0icPtp2M');
  var sheet = ss.getSheetByName('施設リスト');
  sheet.appendRow([data.name, data.address, data.phone, data.contact, data.notes]);
}
