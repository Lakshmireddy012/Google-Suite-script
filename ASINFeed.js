
var reqCount=20;
var highStatus=0;
var archiveLowThreshold=30;

function init(){
  // give permissions to all mail engines
  setActiveSpreadSheetId();
  assignEditPermissions();
}

function batchHighTrigger(){
  //  var scriptProperties = PropertiesService.getScriptProperties();
  //  scriptProperties.setProperty('startRow',0);
  init();
  ScriptApp.newTrigger('startBatch')
  .timeBased()
  .everyMinutes(1)
  .create(); 
}

function clearProperties(){
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
}
function startBatch(){
  var batchType="H";
  var scriptProperties = PropertiesService.getScriptProperties();
  var SheetIDIndex=scriptProperties.getProperty('SheetIDIndex');
  if(SheetIDIndex==null || SheetIDIndex==undefined){
    setCurrentSpreadSheet();
  }
  makeNewRequest(batchType);
}

function allFileIds(){
  var ss = SpreadsheetApp.openById(getActiveSpreadSheetId());
  var sheet = ss.getSheets()[0];
  var activeRange = sheet.getDataRange();
  var  ids= activeRange.getValues();
  return ids;
}

function setCurrentSpreadSheet(){
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('startRow',0);
  var spreadsheetsIDs=allFileIds();
  var IDsCount=spreadsheetsIDs.length;
  var SheetIDIndex=scriptProperties.getProperty('SheetIDIndex');
  if(SheetIDIndex==null || SheetIDIndex==undefined){
    scriptProperties.setProperty('SheetIDIndex',0);
    scriptProperties.setProperty('SheetIndex',0);
  }else{
    setActiveSheet(IDsCount);
  }
  var currentSheet=spreadsheetsIDs[parseInt(scriptProperties.getProperty('SheetIDIndex'))][0];
  Logger.log(scriptProperties.getProperties());
  scriptProperties.setProperty('SheetID',currentSheet);
}

function setActiveSheet(IDsCount){
  var scriptProperties = PropertiesService.getScriptProperties();
  
  Logger.log(scriptProperties.getProperty('SheetID').toString());
  var ss=SpreadsheetApp.openById(scriptProperties.getProperty('SheetID').toString());
  var sheets=ss.getSheets();
  var sheetsCount=sheets.length;
  var SheetIndex=scriptProperties.getProperty('SheetIndex');
  Logger.log("sheetsCount",sheetsCount);
  // if only one sheet exist in spreadsheet
  if(sheetsCount == 1){
    Logger.log("sheetsCount if",sheetsCount);
    scriptProperties.setProperty('SheetIndex',0);
    // move to next spread sheet
    if (parseInt(scriptProperties.getProperty('SheetIDIndex')) == IDsCount) {
      scriptProperties.setProperty('SheetIDIndex',0);
    }else{
      scriptProperties.setProperty('SheetIDIndex',parseInt(scriptProperties.getProperty('SheetIDIndex'))+1);
    }
  }else if(parseInt(scriptProperties.getProperty('SheetIndex')) < sheetsCount-1){
    Logger.log("sheetsCount else if");
    scriptProperties.setProperty('SheetIndex',parseInt(scriptProperties.getProperty('SheetIndex'))+1);
  }
  else{
    Logger.log("sheetsCount else");
    scriptProperties.setProperty('SheetIndex',0);
    if (parseInt(scriptProperties.getProperty('SheetIDIndex')) == IDsCount) {
      scriptProperties.setProperty('SheetIDIndex',0);
    }else{
      scriptProperties.setProperty('SheetIDIndex',parseInt(scriptProperties.getProperty('SheetIDIndex'))+1);
    }
  }
}


function fetchActiveSheetData(batchType){
  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetID=scriptProperties.getProperty('SheetID').toString();
  Logger.log('scriptProperties before',scriptProperties.getProperties());
  var ss=SpreadsheetApp.openById(sheetID);
  var sheetIndex=parseInt(scriptProperties.getProperty('SheetIndex'));
  Logger.log('sheetIndex',sheetIndex);
  var sheet = ss.getSheets()[sheetIndex];
  var activeRange = sheet.getDataRange();
  var data= activeRange.getValues();
  var ASINListData=[];
  var startRow=scriptProperties.getProperty('startRow');
  if(startRow==null || startRow==undefined){
    scriptProperties.setProperty('startRow',0);
  }
  startRow=parseInt(scriptProperties.getProperty('startRow'))+1;
  Logger.log("startRow",startRow);
  for(var i= startRow;i<=data.length;i++){
    var row=data[i-1].filter(n => n!= null);
    if(row){
      if(row.length > 2){
        if(row[5] == highStatus && batchType== "H"){
          ASINListData.push({values: row , row:i});
        } 
        if(row[5] > highStatus && row[5] <= archiveLowThreshold && batchType== "L"){
          ASINListData.push({values: row , row:i});
        }
        if(row[5] > archiveLowThreshold && batchType== "A"){
          ASINListData.push({values: row , row:i});
        }
      }else if(row.length >1){
        ASINListData.push({values: row , row:i});
      }
    }
    // contniue next time with last row
    if(ASINListData.length ==500){
      scriptProperties.setProperty('startRow',i);
      break;
    }
    // if sheet data is over, change active sheet
    if(i==data.length){
      Logger.log("Update sheet method triggered")
      setCurrentSpreadSheet();
    }
  }
  Logger.log("Started at row ",ASINListData[0].row ," Ended at ",ASINListData[ASINListData.length-1].row);
  return ASINListData;
}


function ASINListBatches(ASINListData){
  var ASINListBatchesArray=[];
  for(var i=0;i<=ASINListData.length/reqCount;i++){
    var ASINListBatch=[];
    var temp=(reqCount*i);
    for(var j=0+temp;j<reqCount+temp;j++){
      if(ASINListData[j]){
        ASINListBatch.push(ASINListData[j]);
      }
    }
    if(ASINListBatch.length > 0){
      ASINListBatchesArray.push(ASINListBatch);
    }
  }
  return ASINListBatchesArray;
}

function makeNewRequest(batchType) {
  var scriptProperties = PropertiesService.getScriptProperties();
  try {
    var config = amazonMWSConfigProperties.getProperty('amazonConfig');
    var configData = JSON.parse(config);
    var sellerID=configData.sellerID;
    var accessKey=configData.accessKey;
    var secretKey=configData.secretKey;
    var authToken=configData.authToken;
    var defaultMarket=configData.defaultMarket;
    var sheetId= scriptProperties.getProperty('SheetID').toString();
    var sheetIndex= parseInt(scriptProperties.getProperty('SheetIndex'));
    var ASINListData=fetchActiveSheetData(batchType);
    var ASINListBatchesVar=ASINListBatches(ASINListData);
    var ASINListParamKeyValues=ASINListArrayParamKeyValues(ASINListBatchesVar);
    
    for(var i=0;i<ASINListBatchesVar.length;i++){
      var paramStr=prepareParamStringForUrl(ASINListParamKeyValues[i]);
      var sortedParamStr=prepareParamStringForSign(ASINListParamKeyValues[i]);
      var url = 'https://mws.amazonservices.fr/Products/2011-10-01?';
      var today = new Date();
      var unsignedURL = 
          'POST\nmws.amazonservices.fr\n'+
            '/Products/2011-10-01\n'+
              sortedParamStr+
                '&AWSAccessKeyId=' +accessKey+
                  '&Action=GetLowestOfferListingsForASIN'+
                    '&MarketplaceId='+defaultMarket+
                      '&SellerId='+sellerID+
                        '&SignatureMethod=HmacSHA256'+
                          '&SignatureVersion=2'+
                            '&Timestamp='+encodeURIComponent(Utilities.formatDate(today, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'")) + 
                              '&Version=2011-10-01';
      //Logger.log(unsignedURL);
      var SignedRequest = calculatedSignature(unsignedURL, secretKey);
      var Encoded = Utilities.base64Encode(SignedRequest);
      var encodedSignedRequest=encodeURIComponent(SignedRequest);
      // update counter
      quotaCounter(scriptProperties);
      // Logger.log(encodedSignedRequest);
      var param = 
          'AWSAccessKeyId=' +accessKey+
            '&Action=GetLowestOfferListingsForASIN'+
              '&SellerId='+sellerID+
                '&SignatureVersion=2'+
                  '&Timestamp='+encodeURIComponent(Utilities.formatDate(today, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'")) + 
                    '&Version=2011-10-01'+
                      '&Signature='+encodedSignedRequest + 
                        '&SignatureMethod=HmacSHA256'+
                          '&MarketplaceId='+defaultMarket+
                            '&'+paramStr;
      
      var options = {
        "method" : "POST",
        "muteHttpExceptions" : true
      };
      var result = UrlFetchApp.fetch(url+param,options);
      //Logger.log(result);
      if (result.getResponseCode() == 200) {
        //Logger.log(i);
        var processedResults=processResult(result,ASINListBatchesVar[i],batchType);
        //Logger.log(processedResults);
        pushValuesToExcel(processedResults,sheetId,sheetIndex);
      }
    }
  } catch(err) {
    Logger.log(err);
//    var failedBatchesStr=scriptProperties.getProperty("FailedBatches");
//    if(failedBatchesStr!=null && failedBatchesStr!=undefined){
//      var failedBatches= JSON.parse(failedBatchesStr);
//      failedBatches.push({"batchType": batchType , "sheetId":sheetId,"sheetIndex":sheetIndex, "startRow":ASINListData[0].row ,"endRow":ASINListData[ASINListData.length-1].row ,"error":err});
//      scriptProperties.setProperty("FailedBatches", JSON.stringify(failedBatches));
//    }else{
//      var failedBatches= [{"batchType": batchType , "sheetId":sheetId,"sheetIndex":sheetIndex, "startRow":ASINListData[0].row ,"endRow":ASINListData[ASINListData.length-1].row,"error":err}];
//      scriptProperties.setProperty("FailedBatches", JSON.stringify(failedBatches));
//    }
  }
}



function calculatedSignature(url,secret) {
  var urlToSign = url;
  
  var byteSignature = Utilities.computeHmacSha256Signature(urlToSign, secret);
  // convert byte array to hex string
  var signature = byteSignature.reduce(function(str,chr){
    chr = (chr < 0 ? chr + 256 : chr).toString(16);
    return str + (chr.length==1?'0':'') + chr;
  },'');
  return Utilities.base64Encode(byteSignature);
}


function XML_to_JSON(xml) {
  var doc = XmlService.parse(xml);
  var result = {};
  var root = doc.getRootElement();
  result[root.getName()] = elementToJSON(root);
  return result;
}

function elementToJSON(element) {
  var result = {};
  // Attributes.
  element.getAttributes().forEach(function(attribute) {
    result[attribute.getName()] = attribute.getValue();
  });
  // Child elements.
  element.getChildren().forEach(function(child) {
    var key = child.getName();
    var value = elementToJSON(child);
    if (result[key]) {
      if (!(result[key] instanceof Array)) {
        result[key] = [result[key]];
      }
      result[key].push(value);
    } else {
      result[key] = value;
    }
  });
  // Text content.
  if (element.getText()) {
    result['Text'] = element.getText();
  }
  return result;
}

function xmlToJson_(xml) {
  
  // Create the return object
  var obj = {};
  
  // get type
  var type = '';
  try { type = xml.getType(); } catch(e){}
  
  if (type == 'ELEMENT') {
    // do attributes
    var attributes = xml.getAttributes();
    if (attributes.length > 0) {
      obj["@attributes"] = {};
      for (var j = 0; j < attributes.length; j++) {
        var attribute = attributes[j];
        obj["@attributes"][attribute.getName()] = attribute.getValue();
      }
    }
  } else if (type == 'TEXT') {
    obj = xml.getValue();
  }
  
  // get children
  var elements = [];
  try { elements = xml.getAllContent(); } catch(e){}
  
  
  // do children
  if (elements.length > 0) {
    for(var i = 0; i < elements.length; i++) {
      var item = elements[i];
      
      var nodeName = false;
      try { nodeName = item.getName(); } catch(e){}
      
      if (nodeName)
      {
        if (typeof(obj[nodeName]) == "undefined") {
          obj[nodeName] = xmlToJson_(item);
        } else {
          if (typeof(obj[nodeName].push) == "undefined") {
            var old = obj[nodeName];
            obj[nodeName] = [];
            obj[nodeName].push(old);
          }
          obj[nodeName].push(xmlToJson_(item));
        }                
      }
    }
  }
  return obj;
};

function ASINListArrayParamKeyValues(ASINListArray){
  var ASINListArrayParamKeyValuesVar=[];
  for(var i=0; i < ASINListArray.length ; i++){
    var ASINList=ASINListArray[i];
    var ASINListKeyValArray=[];
    for(var j=1; j <= ASINList.length ; j++){
      var keyVal={};
      var key = 'ASINList.ASIN.'+j;
      // condition will not execute for the first time as only code  and ASIN will present in the list
      if(ASINList[j-1].length > 2){
        keyVal[key]=ASINList[j-1].values[1];
        ASINListKeyValArray.push(keyVal);
      }else{
        // without condition for the first time
        keyVal[key]=ASINList[j-1].values[1];
        ASINListKeyValArray.push(keyVal);
      }
    }
    ASINListArrayParamKeyValuesVar.push(ASINListKeyValArray);
  }
  return ASINListArrayParamKeyValuesVar;
}

function prepareParamStringForSign(array){
  var sortedArray=array.sort(function(a,b){
    return (Object.keys(a)[0] > Object.keys(b)[0]) - 0.5;
  });
  return paramString(sortedArray);
}

function prepareParamStringForUrl(array){
  return paramString(array);
}

function paramString(array){
  var ASINListString="";
  var length=array.length;
  for(var i=0; i < length ; i++){
    var key=Object.keys(array[i])[0];
    ASINListString=ASINListString+key+'='+array[i][key];   
    if(i!=length-1){
      ASINListString=ASINListString+'&';
    }
  }
  return ASINListString;
}

function processResult(result,ASINListArray,batchType){
  //Logger.log(JSON.stringify(ASINListArray));
  // var xml = XmlService.parse(result); 
  var jsonResult=XML_to_JSON(result);
  var processedResults=[];
  var ASINResultsArrayJson=jsonResult.GetLowestOfferListingsForASINResponse.GetLowestOfferListingsForASINResult;
  //Logger.log("ASINResultsArrayJson.length" , ASINResultsArrayJson.length);
  var curDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  //Logger.log(ASINResultsArrayJson.length);
  //Logger.log(ASINResultsArrayJson);
  for(var i=0;i<ASINResultsArrayJson.length; i++){
    var product=ASINResultsArrayJson[i].Product;
    var productListings=product.LowestOfferListings.LowestOfferListing;
    var price=undefined;
    var merchant;
    // Logger.log("productListings.length" , productListings.length);
    var processedResult=[];
    var asin;
    // if result contain single product
    if(productListings instanceof Array){
      for(var j=0;j<productListings.length; j++){
        merchant=productListings[j].Qualifiers.FulfillmentChannel.Text;
        if(merchant== "Amazon"){
          if(price!=undefined){
            if(parseFloat(price) > parseFloat(productListings[j].Price.LandedPrice.Amount.Text)){
              price=productListings[j].Price.LandedPrice.Amount.Text;
            }
          }else{
            price=productListings[j].Price.LandedPrice.Amount.Text;
          }
        }
      }
    }else if(productListings){
      merchant=productListings.Qualifiers.FulfillmentChannel.Text;
      if(merchant== "Amazon"){
        if(price!=undefined){
          if(parseFloat(price) > parseFloat(productListings.Price.LandedPrice.Amount.Text)){
            price=productListings.Price.LandedPrice.Amount.Text;
          }
        }else{
          price=productListings.Price.LandedPrice.Amount.Text;
        }
      }
    }
    asin=product.Identifiers.MarketplaceASIN.ASIN.Text;
    processedResult.push(curDate);
    if(price==undefined){
      processedResult.push("N/A");
      processedResult.push("N/A");
      processedResult.push(1);
    }else{
      processedResult.push("in stock");
      processedResult.push(price);
      processedResult.push(0);
    }
    processedResult.push(asin);
    //Logger.log(processedResult);
    ASINListArray=rowValuesMapping(ASINListArray,processedResult,batchType);
    //Logger.log(asin);
  }
  return ASINListArray;
}

function rowValuesMapping(ASINListArray,processedResult,batchType){
  // for duplicates
  function getAllIndexes(arr, val) {
    var indexes = [], i;
    for(i = 0; i < arr.length; i++)
      if (arr[i].values[1] == val)
        indexes.push(i);
    return indexes;
  }
  
  var indexes = getAllIndexes(ASINListArray, processedResult[4]);
  
  for(var i=0; i< indexes.length ; i++){
    var index=indexes[i];
    if(ASINListArray[index].values.length > 2){
      // date and ASIN is same for all
      ASINListArray[index].values[2]=processedResult[0];
      ASINListArray[index].values[6]=processedResult[4];
      
      if(ASINListArray[index].values[5] == highStatus  && batchType== "H"){  
        ASINListArray[index].values[3]=processedResult[1];
        ASINListArray[index].values[4]=processedResult[2];
        ASINListArray[index].values[5]=processedResult[3];
      }
      // for low or archive
      if((ASINListArray[index].values[5] > highStatus && ASINListArray[index].values[5] <= archiveLowThreshold && batchType== "L") ||
         ASINListArray[index].values[5] > archiveLowThreshold && batchType== "A" ){
        // if product is back to stock 
        if(processedResult[1]=="in stock"){
          ASINListArray[index].values[3]=processedResult[1];
          ASINListArray[index].values[4]=processedResult[2];
          ASINListArray[index].values[5]=processedResult[3];
        }else{
          ASINListArray[index].values[5]=ASINListArray[index].values[5]+1;
          // Logger.log(ASINListArray[index]);
        }
      }
    }else{
      ASINListArray[index].values[2]=processedResult[0];
      ASINListArray[index].values[3]=processedResult[1];
      ASINListArray[index].values[4]=processedResult[2];
      ASINListArray[index].values[5]=processedResult[3];
      ASINListArray[index].values[6]=processedResult[4];
      //processedResults.push(processedResult);
    }
    //Logger.log('    ---',i,ASINListArray);
  }
  return ASINListArray;
}

function pushValuesToExcel(values,sheetId,sheetIndex){
  var ss=SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheets()[sheetIndex];
  for(var i=0;i<values.length;i++){
    var rowData=values[i];
    var rowNumber=rowData.row;
    //Logger.log(rowData.values);
    try {
      sheet.getRange("A"+rowNumber+":G"+rowNumber).setValues([rowData.values]);
      //Logger.log("Updated ",rowData.values);
    } catch(err) {
      //Logger.log(err,rowData.values);
    }
    
  }
  
}


function quotaCounter(scriptProperties){
  var fetchUrlCounter=scriptProperties.getProperty('fetchUrl');
  if(fetchUrlCounter==null || fetchUrlCounter==undefined){
    scriptProperties.setProperty('fetchUrl',0);
  }else{
    scriptProperties.setProperty('fetchUrl',parseInt(scriptProperties.getProperty('fetchUrl'))+1);
    // stop current engine
    if(parseInt(scriptProperties.getProperty('fetchUrl')) > 19500){
      stopActiveTriggersInCurrentEngine();
      startBatchInNewEngine();
      scriptProperties.setProperty('fetchUrl',0);
    }
  }
}

function startBatchInNewEngine(){
  var ss = SpreadsheetApp.openById(getActiveSpreadSheetId());  
  var remoteLinks=getRemoteEngineUrls(ss.getId());
  Logger.log("Engine one link ",remoteLinks[0][0]);
  UrlFetchApp.fetch(remoteLinks[0][0].toString()+"?jobType=startBatch&activeSheetId="+setActiveSpreadSheetId());
}