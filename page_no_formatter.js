function formatPageNo() {
  var ss = SpreadsheetApp.getActiveSheet();
  var columnA = ss.getRange("A2:A505");
  var values = columnA.getValues();
  var result = [];
  for(var index in values){
    if(values[index]){
      var dataStr = values[index].toString();
      // make it uppercase
      var upperData = dataStr.toUpperCase();
      // trim and remove white space between worlds
      var trimmedData = upperData.trim();
      var noWhiteSpaceData = trimmedData.replace(" ", "");
      // replace ',' with '-'
      var noCommaData = noWhiteSpaceData.replace(",","-");
      
      var data = noCommaData ;
      if(isNaN(noCommaData) && noCommaData.indexOf("-") < 0){
        data = addUniqueIdentifier([noCommaData]);
      }else if(noCommaData.indexOf("-") >= 0){
        var arr = noCommaData.split("-");
        data = addUniqueIdentifier(arr);
      }
      result.push([data]);
   } 
  }
  var newRange = ss.getRange("D2:D505");
  newRange.setValues(result);
}

function addUniqueIdentifier(data){
  // add unique identifier for alpha numeric no
  var result = "";
  if(data.length == 1){
    data = data[0];
    result = data.substring(0,data.length -1) + " " + data.charAt(data.length-1);
  }else if(data.length > 1){
    for(var index in data){
      var d = data[index];
      if(isNaN(d) && d.indexOf("-") < 0){
        result = d.substring(0,d.length -1) + " " + d.charAt(d.length-1) + " ";
      }else{
        //result += d.substring(0,d.length -1) + "-" + d.charAt(d.length-1) + " ";
        result += d + " "
      }
    
    }
  }
  
  return result;
}


function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName('Sheet1');
  var menuEntries = [];
  menuEntries.push({name: "Format Page No", functionName: "formatPageNo"});
  ss.addMenu("DID", menuEntries);  
}