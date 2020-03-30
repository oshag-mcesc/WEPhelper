//POSSIBLE DELETION not needed anymore

/**
 * Creates a "MiddleMan" form to be used to submit WEP goals 
 *
 * @return {obj} formID, formURL the ID and URL of the new form
 */

function createMiddleManForm() {
  //Create the form and move it to the correct folder
  //The following line should get the folder ID of where the spreadsheet is in
  var folders = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
  var folder = folders.next().getId();
  var form = FormApp.create(SpreadsheetApp.getActive().getName() +  " [Middle Man]");
  var file = DriveApp.getFileById(form.getId());
  DriveApp.getFolderById(folder).addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  //Create the Form Items 
  form.addListItem().setTitle("Student Name");
  form.addTextItem().setTitle("Teacher");
  form.addTextItem().setTitle("Class/Course");
  form.addTextItem().setTitle("Gifted Area");
  form.addTextItem().setTitle("ODE Instructional Code and Setting");
  form.addParagraphTextItem().setTitle("Instructional Strategies for Service");
  form.addParagraphTextItem().setTitle("Measures to determine student growth");
  form.addParagraphTextItem().setTitle("Content Goal");
  form.setRequireLogin(true);
  
  //return the info
  return {
    formID: form.getId(),
    formURL: form.getEditUrl()
  }
}

/**
 * Creates a from response with data then submits it.
 *
 * @param {array} allData the array of data values
 * @param {string} MidManFormID the ID of the MiddleMan form file
 * @return {boolean} done returns true if done
 */

function submitMiddleManForm(allData, MidManFormID){
  //get range data into object data
  var theData = ObjApp.rangeToObjects(allData);
  //get the form and all the items on the form
  var frm = FormApp.openById(MidManFormID);
  var items = frm.getItems();
  
  var start = new Date(); //set the start time
  
  //NOTE: Need to uncomment/comment code for difference if doc has studNumber or not
  Logger.log(theData[0].odeisandcode);
  //Loop through each row and submit
  for(var i = 0; i < theData.length; i++){
  var filename = theData[i].studlastfirst + " " + theData[i].studentnumber;
  var resp = frm.createResponse();
  //set the values for the response
  resp.withItemResponse(items[0].asListItem().createResponse(filename));
  //resp.withItemResponse(items[0].asListItem().createResponse(theData[i].studlastfirst));
  resp.withItemResponse(items[1].asTextItem().createResponse(theData[i].teacherlastfirst));
  resp.withItemResponse(items[2].asTextItem().createResponse(theData[i].coursename));
  resp.withItemResponse(items[3].asTextItem().createResponse(theData[i].giftedarea));
  resp.withItemResponse(items[4].asTextItem().createResponse(theData[i].odeisandcode));
  resp.withItemResponse(items[5].asParagraphTextItem().createResponse(theData[i].instructionalstrategies)); 
  resp.withItemResponse(items[6].asParagraphTextItem().createResponse(theData[i].assesstools));
  resp.withItemResponse(items[7].asParagraphTextItem().createResponse(theData[i].goal)); 
  
  //go and submit it!!
  resp.submit();
  
  //check execution time
    if(isTimeUp_(start)){
       var theInfo = {
        done: false,
        lastStudent: theData[i].studlastfirst,
        row:i
        }
       break; //get out if time is up!!
       }
  }
  if(theInfo){
    return theInfo;
    }
    else{
    return {done:true}
    }
 
}
