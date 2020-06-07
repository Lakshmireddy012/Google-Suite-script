function saveAsCSV(type) {
//  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
//  var sheet = ss.getSheets()[0];
  // create a folder from the name of the spreadsheet
  //var folder = DriveApp.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + new Date().getTime());
  var curDate = Utilities.formatDate(new Date(), "CEST", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var folder = folderCreation();
  // append ".csv" extension to the sheet name
  var fileName="Exported_"+ curDate + ".csv";
  var spreadsheetsIDs=allFileIds();
  var csvFile = "";
  for(var i = 0; i < spreadsheetsIDs.length; i++){
    var ss=SpreadsheetApp.openById(spreadsheetsIDs[i].toString());
    var sheets=ss.getSheets();
    for(var j = 0; j < sheets.length; j++){
      csvFile = csvFile+convertRangeToCsvFile_(fileName, sheets[j],type);
    }
  }
  
  if(type=="H"){
    fileName = "High_Priority_"+fileName;
  }
  if(type=="L"){
    fileName = "Low_Priority_"+fileName;
  }
  if(type=="A"){
    fileName = "Archive_"+fileName;
  }
  Logger.log("prepare before");
  // convert all available sheet data to csv format
  //var csvFile = convertRangeToCsvFile_(fileName, sheet,type);
  //Logger.log(csvFile);
  // create a file in the Docs List with the given name and the csv data
  var file = folder.createFile(fileName, csvFile);
  Logger.log("prepare after");
  //File downlaod
  var downloadURL = file.getDownloadUrl().slice(0, -8);
  showurl(downloadURL);
  
}
function showurl( url ){
  var html = HtmlService.createHtmlOutput('<html><script>'
                                          +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
                                          +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
                                          +'if(document.createEvent){'
                                          +'  var event=document.createEvent("MouseEvents");'
                                          +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
                                          +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
                                          +'}else{ a.click() }'
                                          +'close();'
                                          +'</script>'
                                          // Offer URL as clickable link in case above code fails.
                                          +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
                                          +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
                                          +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}

function convertRangeToCsvFile_(csvFileName, sheet , type) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  Logger.log("Read before");
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;
    if(type=="H"){
      data=data.filter(val=> {if(val.length> 2){ return val[5]==0}});
    }
    if(type=="L"){
      data=data.filter(val=> {if(val.length> 2){ return val[5]>0 && val[5] <=30}})
    }
    if(type=="A"){
      data=data.filter(val=> {if(val.length> 2){ return val[5]>30}})
    }
    //Logger.log(data);
    // loop through the data in the range and build a string with the csv data
    //if (data.length => 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }
        // remove unwanted columns
        data[row].splice(2,1);
        // index will adjusted by 1 after splice
        data[row].splice(4,2);
        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    //}
    Logger.log("Read complete");
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

function folderCreation(){
  var folderList=DriveApp.getFoldersByName("AmazonTrackerExportedFiles");
  Logger.log("folderList",folderList);
  if(folderList.hasNext()){
    var folder=folderList.next();
  }else{
    var folder=DriveApp.createFolder("AmazonTrackerExportedFiles");
  }
}

