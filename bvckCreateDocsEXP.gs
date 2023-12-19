const tester2 = () => {
  const ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1eh_O6C358dwzpIgbD9zmXd4rVID6JM7fAekU7ANDOBY/edit#gid=292274356")
  const sheet = ss.getSheetByName("forDocs")
  let studInfo = sheet.getDataRange().getValues()
  let saveFolderId = "1ZkfNQtGVxs2mZ40Eb5JBpL2pezfkjc9X"
  let templateId = "1bKhtliVNUglYLjJLX5zPWBqq5Rw9TjXqtMvPjrc2atM"
  let rowStart = 0
  let infoObj = {
    sheet: sheet,
    theData: studInfo,
    fileID: templateId,
    folderID: saveFolderId,
    startRow: rowStart
  }
  let rslt = createWEPs.create(infoObj)
  console.log(rslt)
}

/**
 * Change to arrow function with a namespace
 * Change to accept an opject then use spread operator
 * Change to "forEach" instead of for/loop
 * Change from setting a var for each replacement to setting the code in the replacement
 * 
 */

const createWEPs1 = (() => {
  const create_ = (theInfo) => {
    const { sheet, theData, fileID, folderID, startRow } = theInfo
    //Get the data and change to an object
    let data = ObjApp.rangeToObjects(theData);

    //get the template file and destination folder
    let file = DriveApp.getFileById(fileID);
    let folder = DriveApp.getFolderById(folderID);

    //set the start row
    let rowNum = startRow;

    //more info needed!!
    let docIDs = []; //array to hold all the document IDs
    let firstRow = startRow + 2; //ObjApp removes header from range so add 2
    let start = new Date();  //set the start time for the script
    let rslts = {} //object to return info to calling script

    //SpreadsheetApp.getActiveSpreadsheet().toast("Running....", "Creating", 2);

    data.every((row, index) => {
      //studentID#_Lastname_firstname_23_24_WEP
      //let fileName = row.studentnumber + "_" + row.studlast + "_" + row.studfirst + "_23_24_WEP"
      //make copy of the template, get the body of the new doc
      //replace the text as needed
      let newFile = file.makeCopy(row.filename, folder);
      let body = DocumentApp.openById(newFile.getId());
      body.replaceText('<<StudLastFirst>>', row.studlastfirst);
      body.replaceText('<<Grade>>', row.grade);
      body.replaceText('<<StudFirst>>', row.studfirst);
      body.replaceText('<<StudentNumber>>', row.studentnumber);
      body.replaceText('<<GiftedArea>>', row.giftedarea);
      body.replaceText('<<SchoolCode>>', row.schoolcode);
      //load the docID array
      docIDs.push([newFile.getId()]);

      //check execution time ***NEED TO SET THIS TO THE CALLING SCRIPT!!!
      if (isTimeUp_(start)) {
        rslts = { done: "false", rowNum: index + 1 };
        return false; //get out if time is up!!
      }
      return true
    })


    let lcol = sheet.getLastColumn();
    if (!(sheet.getRange(1, lcol).getValue() == 'DocID')) {
      sheet.getRange(1, lcol + 1).setValue('DocID');
      lcol += 1;
    }
    //fill in the IDs
    sheet.getRange(firstRow, lcol, docIDs.length).setValues(docIDs);

    //check if we are done
    if (docIDs.length >= data.length) {
      rslts = { done: "true", rowNum: "0" };
    }

    //return all info
    return rslts;
  }

  return {
    create: create_
  }
})()

//needed global let to be called from a library!
var createWEPs = createWEPs1










/**
 * Creates a document based on a template for each row of data.
 * (NOTE: There is a timing function to catch if script runs to long)
 *
 * @param {obj} sheet the sheet with all the data
 * @param {array} theData the array of data values
 * @param {string} fileID the ID of the template file
 * @param {string} folderID the ID of the destination folder
 * @param {integer} startRow the row to start processing
 * @return {object} the code will return {rowNum, done} where ronwNum is the next row to process and done if done
 */


function createDocs2(sheet, theData, fileID, folderID, startRow) {
  //Get the data and change to an object
  let data = ObjApp.rangeToObjects(theData);

  //get the template file and destination folder
  let file = DriveApp.getFileById(fileID);
  let folder = DriveApp.getFolderById(folderID);

  //set the start row
  let rowNum = startRow;

  //more lets needed!!
  let docIDs = []; //array to hold all the document IDs
  let firstRow = startRow + 2; //ObjApp removes header from range so add 2
  let start = new Date();  //set the start time for the script
  let rslts = {} //object to return info to calling script

  SpreadsheetApp.getActiveSpreadsheet().toast("Running....", "Creating", 2);

  let dataLength = data.length;
  for (rowNum; rowNum < dataLength; rowNum++) {
    //the info needed
    let firstName = data[rowNum].studfirst;
    let studName = data[rowNum].studlastfirst;
    let grade = data[rowNum].grade;
    let giftedArea = data[rowNum].giftedarea;
    let schoolCode = data[rowNum].schoolcode;
    let returnTo = data[rowNum].returnto;
    let studentnumber = data[rowNum].studentnumber;
    let fileName = studName + " " + studentnumber;

    //make copy of the template, get the body of the new doc
    //replace the text as needed
    let newFile = file.makeCopy(fileName, folder);
    let body = DocumentApp.openById(newFile.getId());
    body.replaceText('<<StudLastFirst>>', studName);
    body.replaceText('<<Grade>>', grade);
    body.replaceText('<<StudFirst>>', firstName);
    body.replaceText('<<StudentNumber>>', studentnumber);
    body.replaceText('<<GiftedArea>>', giftedArea);
    body.replaceText('<<SchoolCode>>', schoolCode);

    //load the docID array
    docIDs.push([newFile.getId()]);

    //check execution time ***NEED TO SET THIS TO THE CALLING SCRIPT!!!
    if (isTimeUp_(start)) {
      rslts = { done: "false", rowNum: rowNum + 1 };
      //Logger.log("Limit met!!!");
      break; //get out if time is up!!
    }
  }
  let lcol = sheet.getLastColumn();
  if (!(sheet.getRange(1, lcol).getValue() == 'DocID')) {
    sheet.getRange(1, lcol + 1).setValue('DocID');
    lcol += 1;
  }
  //fill in the IDs
  sheet.getRange(firstRow, lcol, docIDs.length).setValues(docIDs);

  //check if we are done
  if (rowNum >= dataLength) {
    rslts = { done: "true", rowNum: "0" };
  }

  //return all info
  return rslts;
}