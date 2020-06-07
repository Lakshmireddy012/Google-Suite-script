function assignEditPermissions() {
  var mailIDs=allMailIds();
  var spreadsheetsIDs=allFileIds();
  for(var i = 0; i < spreadsheetsIDs.length; i++){
    var ss=SpreadsheetApp.openById(spreadsheetsIDs[i][0].toString());
    for(var j = 0; j < mailIDs.length; j++){
      // give permission to active sheet as well
      var activeSheet = SpreadsheetApp.openById(getActiveSpreadSheetId());
      activeSheet.addEditor(mailIDs[j][0].toString());
      ss.addEditor(mailIDs[j][0].toString());
    }
  }
}

function allMailIds(){
  var id=getActiveSpreadSheetId();
  Logger.log(id)
  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheets()[1];
  var activeRange = sheet.getDataRange();
  var  ids= activeRange.getValues();
  return ids;
}


function pushCurrentEvnDetails(){
    var ss = SpreadsheetApp.openById(getActiveSpreadSheetId());
    var sheet = ss.getSheets()[2];
    var scriptProperties = PropertiesService.getScriptProperties();
    var amazonMWSConfigProperties = PropertiesService.getUserProperties();
    var SheetIDIndex=scriptProperties.getProperty('SheetIDIndex');
    var SheetIndex=scriptProperties.getProperty('SheetIndex');
    var SheetID=scriptProperties.getProperty('SheetID');
    var startRow=scriptProperties.getProperty('startRow');
    var config = amazonMWSConfigProperties.getProperty('amazonConfig');
    var configData = JSON.parse(config);
    var sellerID=configData.sellerID;
    var accessKey=configData.accessKey;
    var secretKey=configData.secretKey;
    var authToken=configData.authToken;
    var defaultMarket=configData.defaultMarket;
    sheet.getRange(1, 2).setValue(1);
    sheet.getRange(2, 2).setValue(SheetIndex);
    sheet.getRange(3, 2).setValue(SheetIDIndex);
    sheet.getRange(4, 2).setValue(SheetID);
    sheet.getRange(5, 2).setValue(startRow);
    sheet.getRange(6, 2).setValue(sellerID);
    sheet.getRange(7, 2).setValue(accessKey);
    sheet.getRange(8, 2).setValue(secretKey);
    sheet.getRange(9, 2).setValue(authToken);
    sheet.getRange(10, 2).setValue(defaultMarket);
    sheet.getRange(11, 2).setValue(ss.getId());
}


function doGet(e) {
  Logger.log(e.parameter.jobType);
  Logger.log(e.parameter.activeSheetId);
  if(e.parameter.jobType=="stopTriggers"){
    stopActiveTriggersInCurrentEngine();
  }
  if(e.parameter.jobType=="startBatch"){
    setActiveJobValuesToProps(e.parameter.activeSheetId);
    batchHighTrigger();
  }
}

function setActiveJobValuesToProps(activeSheetID){
    var ss = SpreadsheetApp.openById(activeSheetID)
    var sheet = ss.getSheets()[2];
    var SheetIndex=sheet.getRange(2, 2).getValue();
    var SheetIDIndex=sheet.getRange(3, 2).getValue();
    var SheetID=sheet.getRange(4, 2).getValue();
    var startRow=sheet.getRange(5, 2).getValue();
    var sellerID=sheet.getRange(6, 2).getValue();
    var accessKey=sheet.getRange(7, 2).getValue();
    var secretKey=sheet.getRange(8, 2).getValue();
    var authToken=sheet.getRange(9, 2).getValue();
    var defaultMarket=sheet.getRange(10, 2).getValue();
    
    var scriptProperties = PropertiesService.getScriptProperties();
    var amazonMWSConfigProperties = PropertiesService.getUserProperties();
    scriptProperties.setProperty('SheetIDIndex',SheetIDIndex);
    scriptProperties.setProperty('SheetIndex',SheetIndex);
    scriptProperties.setProperty('SheetID',SheetID);
    scriptProperties.setProperty('startRow',startRow);
    scriptProperties.setProperty("activeSheetID", activeSheetID);
    amazonMWSConfigProperties.setProperty('amazonConfig',JSON.stringify({'sellerID':sellerID,'accessKey':accessKey, 'secretKey':secretKey , 'authToken': authToken , 'defaultMarket':defaultMarket}));
}

function getRemoteEngineUrls(activeSheetID){
  var ss = SpreadsheetApp.openById(activeSheetID)
  var sheet = ss.getSheets()[2];
  var remoteLinks=sheet.getRange("B11:B12").getValues();
  return remoteLinks;
}

function getActiveSpreadSheetId(){
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty("activeSheetID").toString();
}

function setActiveSpreadSheetId(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("activeSheetID", ss.getId());
}