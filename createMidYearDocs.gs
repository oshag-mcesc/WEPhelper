/**
 * Creates a document based on a template for each row of data.
 * (NOTE: There is a timing function to catch if script runs to long)
 *
 * @param {obj} sheet the sheet with all the data
 * @param {array} theData the array of data values
 * @param {string} fileID the ID of the template file
 * @param {string} folderID the ID of the destination folder
 * @return {object} the code will return {rowNum, done} where ronwNum is the next row to process and done if done
 */

const createMidYears1 = (()=>{
 const createMidYearDocs_ =(sheet, theData, fileID, folderID) =>{  
    //get the template file and destination folder
    let file = DriveApp.getFileById(fileID);
    let folder = DriveApp.getFolderById(folderID);
   
    //more lets needed!!
    let docIDs = []//array to hold all the document IDs
    let start = new Date();  //set the start time for the script
    let rslts = {} //object to return info to calling script
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Running....", "Creating Mid Year Docs", 2);
    //get rid of header row of the data
    let headers = theData.shift()
    
    //For each row create a doc
    theData.every((row, rowNum)=>{
      let [studentnumber,studlastfirst,studlast,studfirst,grade,giftedarea,schoolcode] = row
      let fileName = studlastfirst + " " + studentnumber + " - Mid Year"
      
      //make copy of the template, get the body of the new doc
      //replace the text as needed
      let newFile = file.makeCopy(fileName, folder);
      let body = DocumentApp.openById(newFile.getId());
      body.replaceText('<<StudLastFirst>>', studlastfirst);
      body.replaceText('<<Grade>>',grade);
      body.replaceText('<<StudFirst>>',studfirst);
      body.replaceText('<<StudentNumber>>',studentnumber);
      body.replaceText('<<GiftedArea>>',giftedarea);
      body.replaceText('<<SchoolCode>>',schoolcode);

      //load the docID array
      docIDs.push([newFile.getName(),newFile.getId()]);
      //check execution time ***NEED TO SET THIS TO THE CALLING SCRIPT!!!
      if(isTimeUp_(start)){
        rslts = {done:"false", rowNum: rowNum+1};
        //Logger.log("Limit met!!!");
        return false; //get out if time is up!!
        }  
      return true  
    })


    let lcol = sheet.getLastColumn();
    if(!(sheet.getRange(1, lcol).getValue()=='MidDocName')){
        sheet.getRange(1, lcol+1,1,2).setValues([['MidDocName','Mid DocID']]);
        lcol +=1;
        }
    //fill in the IDs
    sheet.getRange(2, lcol, docIDs.length,2).setValues(docIDs);
    
    //check if we are done
    if(typeof rslts.done === 'undefined'){
      rslts = {done: "true", rowNum: "0"};
    }
      
    //return all info
    return rslts;
  }
  
  return{
    createMidYearDocs : createMidYearDocs_
  }
})()

var createMidDocs = createMidYears1
 