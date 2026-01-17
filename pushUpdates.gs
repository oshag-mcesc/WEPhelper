/**
 * Date: 11/17/2023 
 * This code is to push mid year and final updates to the MAIN wep document.  This elimanates the need
 * to create separate mid and final documents
 * 
 * TOTO: somehow communicate all is good or not!
 */

//These 2 functions are called from the menu.
//They use nsConfig info to send to the "middle man" function
const pushMidUpdate1 = ()=>{
  configPushUpdates(nsConfig.tabs.forMidYearPush)
}

const pushFinalUpdate1 = ()=>{
  configPushUpdates(nsConfig.tabs.forFinalPush)
}

//needed to be called from the library
var pushMidUpdate   = pushMidUpdate1
var pushFinalUpdate = pushFinalUpdate1

/**
 * This is the "middle man" function.  It gets the correct info
 * based on the config data it is sent
 * 
 * TODO: add parameter info
 */
const configPushUpdates = (updateObj) => {
  //desctructure the input parameter
  let {name , rngCell, rowTitle, statusCol} = updateObj
  //set a toast message that we started
  SpreadsheetApp.getActiveSpreadsheet().toast(`Working on ${rowTitle} updates.`,"Working!",5)
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const dataTab = ss.getSheetByName(name)
  let updateData = dataTab.getRange(rngCell).getDataRegion().getValues()
  
  let infoObject = {
    theData:updateData,
    rowTitle:updateData[0][3],  //Get the title for the row from the sheet
    theTab:dataTab,
    theStatusCol:statusCol
  }
  let rslt = pushTheUpdates(infoObject)
  if(rslt.anyErrors){
    let errorsTab = ss.getSheetByName("errors")
    if(!errorsTab){
      errorsTab = ss.insertSheet("errors")
      }
    errorsTab.getRange(errorsTab.getLastRow()+1,1,rslt.errors.length,rslt.errors[0].length).setValues(rslt.errors)
    SpreadsheetApp.getActiveSpreadsheet().toast("There has been an error.  Check the errors tab","Uh - oh!!",-1)
  }else{
    SpreadsheetApp.getActiveSpreadsheet().toast(`Done with updating ${rowTitle}`,"Done!", 5)
  }
  // console.log(rslt.anyErrors);
  // console.log(rslt.errors);
}

/**
 * The WORK horse!!  This function does the heavy lifting
 * of adding the info to the correct student's WEP
 * 
 * TODO: add paramenter info
 */
const pushTheUpdates = (updateInfo) => {
  let {theData, rowTitle, theTab, theStatusCol} = updateInfo
  //set and object to return info
  let results = {}
  results.errors = []  //this will hold an array of IDs and the errors
  //drop the header row
  let headers = theData.shift()

  //get a flat array of just the docIds from the data 
  let studDocIds = [...new Set(nsUtils.keepCols(theData, [0]).flat())]

  //work with each docId
  studDocIds.forEach((docId,idx) => {
    //First see if we can get a good document
    try {
      //open the doc and get the table index info
      let theDoc = DocumentApp.openById(docId)

      //filter the update info for the docId we are working with
      let updateInfo = theData.filter(row => docId == row[0])

      let tableInfo = nsUtils.getTableIndexes(theDoc)

      //SOMEhow get the right data appended to the right table!!
      updateInfo.forEach((row) => {
        //filter the index table info for the teacher name and class from update info... get just the table index (first row third column )
        let theIndex = tableInfo.filter(teacherName => teacherName[0] == row[1]).filter(className => className[1] == row[2])[0][2]

        //now get the table using theIndex, append a row, append the cell info and set background
        let tbl = theDoc.getBody().getTables()[theIndex]
        let tr = tbl.appendTableRow()
        tr.appendTableCell(rowTitle).setBackgroundColor("#E0E0E0")
        tr.appendTableCell(row[3])
        
      })

      theDoc.saveAndClose()
    }
    catch (err) {
      results.errors.push([nsUtils.getFormattedDateTime(),idx+1, docId, err.message])
      results.anyErrors = true
      logIt({"level":"severe","theMsg":"Error pushing updates","error":err})
    }

  })

  return results
}
