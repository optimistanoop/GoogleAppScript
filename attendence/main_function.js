//This function creates a custom menu item in sheet on opening. This helps users to call the allTogether Functionto process automation
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance Sheet');
  var date = new Date();
  var menuEntries = [];
  
   menuEntries.push({name: "Create Attendance Program", functionName: "allTogether"});
  ss.addMenu("Attendance", menuEntries);
   
}

 
function allTogether() {
  
 
 //Takes the values from Sheet "StudenList" and add to array
 
  
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('StudentList');
  var range = sheet.getDataRange(); 
  var data = range.getValues();
  var numberRows = range.getNumRows();
  var numberColumns = range.getNumColumns();
  var lastRow = sheet.getLastRow();
  var ui = SpreadsheetApp.getUi();
  if(ss.getSheets() [0].getName() == "Form Responses 1"){
     ui.alert("Dear User, This function has already been performed");
  }
  
  else{
    
  var form = FormApp.create(SpreadsheetApp.getActiveSpreadsheet().getName());
 
  //Logger.log("Last Row" +lastRow);
  var rowTitles = [];

     //getRange(starting Row, starting column, number of rows, number of columns)

    for(var i=0;i<(lastRow);i++)
    {
      rowTitles.push(data[i]);
      //Logger.log(data[i]);
    }
  
    //Takes the values from Sheet "ParticipationStatus" and add to array
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ParticipationStatus');
  var range = sheet.getDataRange(); 
  var data = range.getValues();
  var numberRows = range.getNumRows();
  var numberColumns = range.getNumColumns();
  var lastRow = sheet.getLastRow();
  //Logger.log("Last Row" +lastRow);
  var columnTitles = [];

     //getRange(starting Row, starting column, number of rows, number of columns)

    for(var i=0;i<(lastRow);i++)
    {
      columnTitles.push(data[i]);
      Logger.log(i+" " + data[i]);
    }
 
  
 //TimeSlot starts here 
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TimePeriod');
  var range = sheet.getDataRange(); 
  var data = range.getValues();
  var numberRows = range.getNumRows();
  var numberColumns = range.getNumColumns();
  var lastRow = sheet.getLastRow();
  
  var timeTitles = [];

     //getRange(starting Row, starting column, number of rows, number of columns)

    for(var i=0;i<(lastRow);i++)
    {
      timeTitles.push(data[i]);
      //Logger.log(data[i]);
    }

 //Form Builder start here
  
  //var form = FormApp.create('All together form 1');

// This adds a Multiple choice Item to newly created form and set values collected from sheet arrays .   
  form.addMultipleChoiceItem()
    .setTitle("TimeSlot") 
    .setHelpText("")
    .setChoiceValues(timeTitles)
    .setRequired(true);
  
  // This adds a Grid Item to newly created form and set values collected from sheet arrays. 
  form.addGridItem()
    .setTitle("Students") 
    .setHelpText("")
    .setRows(rowTitles)
    .setColumns(columnTitles)
    .setRequired(true);
   rowTitles = [];
  columnTitles = [];
  
 //Setting up form desitnation to this sheet 
 form.setDestination(FormApp.DestinationType.SPREADSHEET, SpreadsheetApp.getActiveSpreadsheet().getId());
 
  
 //Getting response sheet name for forumla string starts here. 
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  
  var responseSheet = sheet.getSheets() [0];
  
  //Get Response Sheet Name
  var sheetname = responseSheet.getName();
  
  //Logger.log(sheetname);
  
  //Building Formula String for Attendance Sheet Cell "D6"
  var attendanceRecordString = '=Transpose(INDEX(INDIRECT('+'"Form Responses 1!$C$2:$AZ1000",TRUE),,))'
  
  var DestinationCellToStoreRecordString = sheet.getSheetByName("Attendance Sheet").getRange("D6");
  
  DestinationCellToStoreRecordString.setValue(attendanceRecordString);
  
  var dateRecordString = '=Transpose(INDEX(INDIRECT('+'"Form Responses 1!A:A",TRUE),,))'
    
  var DestinationCellToStoreDateString = sheet.getSheetByName("Attendance Sheet").getRange("C5");
  
  DestinationCellToStoreDateString.setValue(dateRecordString);
    
  var slotRecordString = '=Transpose(INDEX(INDIRECT('+'"Form Responses 1!B:B",TRUE),,))'
    
  var DestinationCellToStoreSlotString = sheet.getSheetByName("Attendance Sheet").getRange("C4");
  
  DestinationCellToStoreSlotString.setValue(slotRecordString);
  }
    
  //Set Time trigger to take monthly backup
  ScriptApp.newTrigger("monthlyBackup")
   .timeBased()
   .onMonthDay(1)
   .create();
}


