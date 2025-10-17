//Helper function to get a list of doc ids
function helper_getListOfDocIds(){
  try{
    const docIdtab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("docIDs")
    const docFolderId = getSettingValue("MainWEPfolderID")
    if(!docIdtab){throw "DocID tab is missing."}
    if(!docFolderId){throw "WEP Doc folder id missing or blank."}
    
    let rslt = getListOfDocIds(docFolderId,docIdtab)
  }
  catch(err){
    logIt({
      level: "severe",
      theMsg: "Error showing settings dialog",
      error: err
      })
    return false
  }
}


/**
 * Generates a list of docIds and titles to check
 * 
 * @param {string} FolderID the id of the folder with the docs
 * @param {string} sheet reference to sheet where the info is to entered
 * @return {boolean} done returns true if done
 */
function getListOfDocIds(FolderID, sheet){
  try{
    var fld = DriveApp.getFolderById(FolderID);
    var files = fld.getFiles();
    var rslts = [];
    while (files.hasNext()) {
     var file = files.next();
     rslts.push([file.getName(),file.getId()]);
  }
    rslts.sort();
    var rng = sheet.getRange(2, sheet.getLastColumn()+1, rslts.length, 2);
    rng.setValues(rslts);
    var headers = sheet.getRange(1,sheet.getLastColumn()-1,1,2);
    headers.setValues([['StudLastFirst','DocID']]);
    return true;
  }
  catch(err){
  return false;
  }
  
}

/**
 * Checks to see if an argument is missing
 * 
 * @param {arg} The argument to check
 * @return {boolean} done returns true if missing
 */
function isMissing(arg){
    Logger.log(typeof arg);
    if(arg === undefined || arg === null){ return true;}
    else{ return false};
}

//Time check helper function
//Google alows GSuite accounts scripts to run for 30 minutes!!  Free accounts are still 4 - 6 minutes
function isTimeUp_(start){
  var now = new Date();
  return now.getTime() - start.getTime() > 1500000; //5 min = 300000  25 min = 1500000
}

//Used to show prompt and return the input
function getAnswer_(ui,title,message) {
  var result = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  
  if(result.getSelectedButton()==ui.Button.OK && result.getResponseText()!=''){
    return result.getResponseText();
    }
  else if(result.getSelectedButton()==ui.Button.OK && result.getResponseText()==''){
    ui.alert("Please enter a value.");
    return;
  }
  else if(result.getSelectedButton()==ui.Button.CANCEL || result.getSelectedButton()==ui.Button.CLOSE){
    return;
  }
}
