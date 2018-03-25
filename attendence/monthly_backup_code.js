
//This function creates a new sheet with the name of current month then copy data from attendance sheet and pastes into newly created month sheet. 
function monthlyBackup(){

   var monthSheetName = createMonthName();
  
  SpreadsheetApp.getActive().insertSheet(monthSheetName, sheetCount());  
  
  CopyInfo(monthSheetName);
  
  CopyData(monthSheetName);
  
  deleteResponsesForNewMonth();
  
}
//Create MonthName Function
function createMonthName() {

 //Get month name from current date
var date = new Date();
var mt = date.getMonth();
var yr = date.getYear();
var currentD 
var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if  (mt === 0) {
   currentD = "JANUARY";
  }if (mt === 1) {
   currentD = "FEBRUARY";
  }if (mt === 2) {
   currentD = "MARCH";
  }if (mt === 3) {
   currentD = "APRIL";
  }if (mt === 4) {
   currentD = "MAY";
  }if (mt === 5) {
   currentD = "JUNE";   
  }if (mt === 6) { 
   currentD = "JULY"; 
  }if (mt === 7) {
   currentD = "AUGUST";  
  }if (mt === 8) {
   currentD = "SEPTEMBER";  
  }if (mt === 9) {
   currentD = "OCTOBER";  
  }if (mt === 10){
  currentD = "NOVOMBER";
  }else if (mt === 11){
  currentD = "DECEMBER";
  }
  var monthSheetName = currentD +", "+ yr;
  Logger.log(monthSheetName); 
  //Logger.log();
  
  return monthSheetName;
}

//Get Sheet count to use in new sheet creating function
function sheetCount(){
  return SpreadsheetApp.getActive().getSheets().length;
}




//This Function copies Name Column to newly created monthly sheet
function CopyInfo(monthlySheetName) {
 var sss = SpreadsheetApp.getActiveSpreadsheet(); //replace with source ID
  var activesheet = sss.getActiveSheet();
 var ss = sss.getSheetByName('Attendance Sheet'); //replace with source Sheet tab name
 var range = ss.getRange(1,2,200,1); //assign the range you want to copy
 var data = range.getValues();

 var tss = SpreadsheetApp.getActiveSpreadsheet(); //replace with destination ID
 var ts = tss.getSheetByName(monthlySheetName); //replace with destination Sheet tab name
 ts.getRange(1, 1,data.length, data[0].length).setValues(data); //you will need to define the size of the copied data see getRange()
}

//This function copies attendance data columns to newly created month sheet
function CopyData(monthlySheetName) {
 var sss = SpreadsheetApp.getActiveSpreadsheet(); //replace with source ID
  var activesheet = sss.getActiveSheet();
 var ss = sss.getSheetByName('Attendance Sheet'); //replace with source Sheet tab name
 var range = ss.getRange(1,4,200,150); //assign the range you want to copy
 var data = range.getValues();

 var tss = SpreadsheetApp.getActiveSpreadsheet(); //replace with destination ID
 var ts = tss.getSheetByName(monthlySheetName); //replace with destination Sheet tab name
 ts.getRange(1, 2,data.length, data[0].length).setValues(data); //you will need to define the size of the copied data see getRange()
}

//This function delete all the responses from form response sheet
function deleteResponsesForNewMonth() {
 
  var ss = SpreadsheetApp.getActiveSheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  var range = sheet.getDataRange(); 
  var data = range.getValues();
  var numberRows = range.getNumRows();
  var numberColumns = range.getNumColumns();
  var lastRow = sheet.getLastRow();
  
  //Logger.log(numberRows);
  
  sheet.deleteRows(2, numberRows);
}
//-----monthlyBackup function ends here-----