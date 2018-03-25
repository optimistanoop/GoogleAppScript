function updateUrls(colName) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  var lastRow = data.length;
  var result = [];
  var cell = sheet.getRange("D1:D"+lastRow);
  Logger.log(cell.getValues());
  if (col != -1) {
    for(var d in data){
      result.push(data[d][col]);
      //var cell = sheet.getRange("D");
			//cell.setFormula('=HYPERLINK("http://www.google.com/","Google")');
      // update it
    }
    
  }
  return result;
}

function myFunction() {
 Logger.log(updateUrls('Edit URL'))
}
