const createWEPdocs = (() => {
  const createInitialWEPs_ = () => {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast("Creating Initial WEPs!", "Started!", -1)
      //Get the rowNum property, set it to 0 if it doesn't exist yet
      const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('forDocs')
      let props = PropertiesService.getScriptProperties()
      let infoObj = {
        sheet: ss,
        theData: ss.getDataRange().getValues(),
        fileID: props.getProperty("WEPtemplateID"),
        folderID: props.getProperty("MainWEPfolderID"),
        startRow: parseInt((props.getProperty("rowNum")) ? props.getProperty("rowNum") : 0)
      }
      let theRslts = createWEPs.create(infoObj)
      //saveSettings(theRslts);
      console.log(theRslts)

      switch (theRslts.done) {
        case "true":
          SpreadsheetApp.getActiveSpreadsheet().toast("Got done in time!", "All done!", -1)
          break;
        case "false":
          SpreadsheetApp.getActiveSpreadsheet().toast("Exceeded time limit! Please run createDocs again.", "NOT done yet!", -1)
          break;
      }
    }
    catch (err) {
      errorhandler_(err, "Creating docs");
    }
  }

  return {
    createInitialWEPs: createInitialWEPs_
  }
})()



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

