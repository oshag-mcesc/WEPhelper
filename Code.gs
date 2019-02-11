/**
 * Creates a document based on a template for each row of data.
 * (NOTE: There is a timing function to catch if script runs to long)
 *
 * @param {obj} sheet the sheet with all the data
 * @param {array} theData the array of data values
 * @param {string} fileID the ID of the template file
 * @param {string} folderID the ID of the destination folder
 * @param {integer} wepType the type of wep 1 = Mid Year, 2 = Final Eval missing will be Initial
 * @return {list} the code will fill in the DocIDs column
 */


function createDocs(sheet, theData, fileID, folderID, wepType) {
  //Get the data and change to an object
  var data =ObjApp.rangeToObjects(theData);
  
  //get the template file and destination folder
  var file = DriveApp.getFileById(fileID);
  var folder = DriveApp.getFolderById(folderID);
  
  //set the WEP type
  var theType
  if(wepType){
    switch(wepType){
      case 1:
        theType = "Mid-Year Progress";
        break;
      case 2:
        theType = "Final Evaluation";
        break;
    }    
  }
  
  //more vars needed!!
  var docIDs = []; //array to hold all the document IDs
  var rowNum; //to hold the row number where stopped
  var firstRow; //pointer to where to start processing
  var start = new Date();  //set the start time for the script
  
  //Get script props and check
  var scriptProps = PropertiesService.getScriptProperties();
  if(scriptProps.getProperty("done")=="false"){
    rowNum = Number(scriptProps.getProperty("rowNum"));
  }
  
  if(!rowNum){rowNum=0}; // first time run no properties are set so we need to set this to 0
  
  firstRow = rowNum + 2; //ObjApp removes header from range so add 2
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Running....", "Creating", 2);
  
  for(rowNum; rowNum < data.length; rowNum++){
    //the info needed
    var firstName = data[rowNum].studfirst;
    var studName = data[rowNum].studlastfirst;
    var grade = data[rowNum].grade;
    var giftedArea = data[rowNum].giftedarea;
    var schoolCode = data[rowNum].schoolcode;
    var returnTo = data[rowNum].returnto;
    var studentnumber = data[rowNum].studentnumber;
    var fileName = studName + " " + studentnumber;
    
    //make copy of the template, get the body of the new doc
    //replace the text as needed
    var newFile = file.makeCopy(fileName, folder);
    var body = DocumentApp.openById(newFile.getId());
    body.replaceText('<<StudLastFirst>>', studName);
    body.replaceText('<<Grade>>',grade);
    body.replaceText('<<StudFirst>>',firstName);
    body.replaceText('<<GiftedArea>>',giftedArea);
    body.replaceText('<<SchoolCode>>',schoolCode);
    body.replaceText('<<ReturnTo>>',returnTo);
    body.replaceText('<<WEP type>>',theType);
    
    //load the docID array
    docIDs.push([newFile.getId()]);
    
    //check execution time ***NEED TO SET THIS TO THE CALLING SCRIPT!!!
    if(isTimeUp_(start)){
       scriptProps.setProperties({done:"false", rowNum: rowNum+1});
       Logger.log("Limit met!!!");
       doneFlag = 0; //set this for the message
       break; //get out if time is up!!
       }
    }
    var lcol = sheet.getLastColumn();
    if(!(sheet.getRange(1, lcol).getValue()=='DocID')){
       sheet.getRange(1, lcol+1).setValue('DocID');
       lcol +=1;
       }
    //fill in the IDs
    sheet.getRange(firstRow, lcol, docIDs.length).setValues(docIDs);
    
    //check if we are done
    if(rowNum >= data.length){
      scriptProps.setProperties({done: "true", rowNum: "0"});
    }
    //give a message to let the user know
    switch(scriptProps.getProperty("done")){
      case "true":
        SpreadsheetApp.getActiveSpreadsheet().toast("Got done in time!", "All done!", -1)
        break;
      case "false":
         SpreadsheetApp.getActiveSpreadsheet().toast("Exceeded time limit! Please run createDocs again.", "NOT done yet!", -1)
        break;
    } 
}





/*****

cleanGoal

This function will take a string and clean it up for
using it in creating a pre filled link to a from.

It takes care of %, /, New lines and &

USE: Select the range you want to clean then run it
OUTPUT: This will put the "clean" string in a new column
*/

function cleanGoal() {
  //get all the data needed
  var ss = SpreadsheetApp.getActiveSheet();
  var rng = ss.getActiveRange();
  var theData = ss.getActiveRange().getValues();
  var ui = SpreadsheetApp.getUi();
  
  //Get the column to put the results
  var result = ui.prompt("Which column?", "You are about to fix illegal charcters in a string.\nWhat column would you like the results in?\
                         \nNOTE: column A is 1, B is 2, etc", ui.ButtonSet.OK_CANCEL);
  
  if(result.getSelectedButton()==ui.Button.OK && result.getResponseText()!=''){
    var column = result.getResponseText();
    }
  else if(result.getSelectedButton()==ui.Button.OK && result.getResponseText()==''){
    ui.alert("Please enter a column number.");
    return;
  }
  else if(result.getSelectedButton()==ui.Button.CANCEL || result.getSelectedButton()==ui.Button.CLOSE){
    return;
  }
  var cleaned = [];
  
  
  for (var i = 0; i < theData.length;i++){
    var goal = theData[i].toString();
    goal = goal.replace(/\%/g,"%25");
    goal = goal.replace(/\//g,"-");
    goal = goal.replace(/\n/g," ")
    goal = goal.replace(/\&/g,"%26");
    cleaned.push([goal]);
  }
  Logger.log(cleaned[0]);
  ss.getRange(2, column,theData.length).setValues(cleaned);
}