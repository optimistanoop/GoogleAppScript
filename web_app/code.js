function doGet() {
  return HtmlService
      .createTemplateFromFile('Index')
      .evaluate();
}

function getData(){
        
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheets = ss.getSheetByName("Construction Expense");
 var sheet = ss.getSheets()[0];

 // This logs the value in the very last cell of this sheet
 var lastRow = sheet.getLastRow();
 var lastColumn = sheet.getLastColumn();
 var lastCell = sheet.getRange(lastRow, lastColumn);
  Logger.log(lastCell.getValue());
 return lastCell.getValue();
  //return lastCell.getValue();
  
}