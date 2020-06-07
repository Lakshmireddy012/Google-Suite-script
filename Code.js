var amazonMWSConfigProperties = PropertiesService.getUserProperties();

function saveCredentials(formData) {
  var sheet = SpreadsheetApp.getActiveSheet();
  amazonMWSConfigProperties.setProperty('amazonConfig', JSON.stringify(formData));
}

function userProps(){
  Logger.log(amazonMWSConfigProperties.getProperties());
}

function loadCredentials() {
  var config = amazonMWSConfigProperties.getProperty('amazonConfig');
  return config;
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ss.toast('Loading... Please wait.'); 
  
  var menu = [  
    {name: "Amazon MWS Credential",     functionName: "showDialog"},
    null,
    {name: "Start ASIN Feed check", functionName: "batchHighTrigger"},
    {name: "Stop ASIN Feed check",  functionName: "stopASINTracking"}
  ];  
  
  var exportSubMenu= [{ name: "High Priority" , functionName: "exportHighPriority"},
                      { name: "Low Priority" , functionName: "exportLowPriority"},
                      { name: "Archive" , functionName: "exportArchive"}];
  ss.addMenu("Amazon ASIN Tracker", menu);
  ss.addMenu("Export to CSV", exportSubMenu);
  
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
  .setWidth(400)
  .setHeight(400);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
  .showModalDialog(html, 'Amazon MWS Credential');
}

function startASINTracking() {  
  try {
    stopASINTracking(true);
    
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('startRow',1);
    scriptProperties.setProperty('startRowLow',1);
    scriptProperties.setProperty('startRowArchive',1);
    startBatchHigh();
    
    ss.toast('The ASIN tracker is now active. You can now close this sheet.', '', -1); 
    
    ScriptApp.newTrigger('startBatchHigh')
    .timeBased()
    .everyMinutes(5)
    .create();  
    
    // Once per day
    ScriptApp.newTrigger('startBatchLowTrigger')
    .timeBased()
    .everyDays(2)
    .create();  
    
    // 4 weeks for archive
    ScriptApp.newTrigger('startBatchArchiveTrigger')
    .timeBased()
    .everyWeeks(4)
    .create();  
    
    return;
  } catch (e) {
    Browser.msgBox(e.toString());
  }
}

function stopASINTracking(e) {
  
  var triggers = ScriptApp.getProjectTriggers();
  
  for(var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  var ss = SpreadsheetApp.openById(getActiveSpreadSheetId());  
  var remoteLinks=getRemoteEngineUrls(ss.getId());
  remoteLinks.forEach((item,index)=> {
    Logger.log(item[0].toString()+"?jobType=stopTriggers");
    UrlFetchApp.fetch(item[0].toString()+"?jobType=stopTriggers");
  });
  if (!e) {
    ss.toast('The ASIN tracker is no longer active. You can restart the tracker anytime later from the same menu.', '', -1); 
  }  
}

function stopActiveTriggersInCurrentEngine(){
  var triggers = ScriptApp.getProjectTriggers();
  for(var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}



function exportLowPriority(){
  saveAsCSV("L");
}

function exportHighPriority(){
  saveAsCSV("H")
}

function exportArchive(){
  saveAsCSV("A")
}