/**
 * Creates a "Mid Year Progress" form to be used to submit WEP goals 
 *
 * @return {obj} formID, formURL the ID and URL of the new form
 */

function createMidYearProgressForm() {
  //Create the form and move it to the correct folder
  //The following line should get the folder ID of where the spreadsheet is in
  var folders = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
  var folder = folders.next().getId();
  var form = FormApp.create(SpreadsheetApp.getActive().getName() +  " Mid Year Progress");
  var file = DriveApp.getFileById(form.getId());
  DriveApp.getFolderById(folder).addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  //Create the Form Items 
  form.addListItem().setTitle("Student Name");
  form.addTextItem().setTitle("Student ID");
  form.addTextItem().setTitle("Course Name");
  form.addTextItem().setTitle("Teacher Name");
  form.addParagraphTextItem().setTitle("Goal");
  form.addParagraphTextItem().setTitle("Progress");
  form.setRequireLogin(false);
  
  //return the info
  return {
    formID: form.getId(),
    formURL: form.getEditUrl()
  }
}

/**
 * Creates a pre filled URL for the mid year progress form
 *
 * @param {obj} sheet Reference to the sheet with the data
 * @param {string} formID The ID of the Mid Year Prgress form file
 * @return {boolean}  Returns true when done
 */

function preFilledURLs(sheet, theData, formID){
  //get the data into an object
  var allData = ObjApp.rangeToObjects(theData);
  //get the form and form items
  var form = FormApp.openById(formID)
  var items = form.getItems();
  var preFilledURL = [];  //needed to hold the URLs
  
  //Loop through the data and set the responses. 
  //AND add the response to the PrefilledURL arrat
  for(var i = 0; i < allData.length; i++){
    var fileName = allData[i].studlastfirst + " " + allData[i].studid;
    var resp = form.createResponse();
    resp.withItemResponse(items[0].asListItem().createResponse(fileName));
    resp.withItemResponse(items[1].asTextItem().createResponse(allData[i].studid));
    resp.withItemResponse(items[2].asTextItem().createResponse(allData[i].coursename));
    resp.withItemResponse(items[3].asTextItem().createResponse(allData[i].teacherlastfirst));
    resp.withItemResponse(items[4].asParagraphTextItem().createResponse(allData[i].goal));
    preFilledURL.push([resp.toPrefilledUrl()]);  
  }
  
  //Add the URLs to the spreadsheet.  
  var lcol = sheet.getLastColumn()+1;
  sheet.getRange(1, lcol).setValue('PreFilledURL');
   
  sheet.getRange(2, lcol, allData.length).setValues(preFilledURL);
  sheet.getRange(2, lcol + 1).setFormula("=arrayformula(HYPERLINK(G2:G,A2:A))"); //NEED a way to offset this!!
  
  return true;
}
