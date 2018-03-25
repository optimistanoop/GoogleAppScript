function setDataValid_(range, sourceRange) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange, true).build();
   range.setDataValidation(rule);
}

function onEdit() {
  var aSheet = SpreadsheetApp.getActiveSheet();
  var aCell = aSheet.getActiveCell();
  var aColumn = aCell.getColumn();
  
  if (aColumn == 6 && aSheet.getName() == 'Refund Overview') {
    var range = aSheet.getRange(aCell.getRow(), aColumn + 1);
    var sourceRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('Data!' + aCell.getValue())
    setDataValid_(range, sourceRange)
  }
 
}