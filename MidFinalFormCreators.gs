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
  form.addTextItem().setTitle("Student Name");
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
 * Creates a "Final Progress" form to be used to submit WEP goals 
 *
 * @return {obj} formID, formURL the ID and URL of the new form
 */

function createFinalProgressForm() {
  //Create the form and move it to the correct folder
  //The following line should get the folder ID of where the spreadsheet is in
  var folders = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
  var folder = folders.next().getId();
  var form = FormApp.create(SpreadsheetApp.getActive().getName() + " End of Year Evaluation");
  var file = DriveApp.getFileById(form.getId());
  DriveApp.getFolderById(folder).addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  //Create the Form Items 
  form.addTextItem().setTitle("Student Name");
  form.addTextItem().setTitle("Student ID");
  form.addTextItem().setTitle("Course Name");
  form.addTextItem().setTitle("Teacher Name");
  form.addParagraphTextItem().setTitle("Goal");
  form.addParagraphTextItem().setTitle("Mid Year Progress");
  form.addParagraphTextItem().setTitle("End of Year Evaluation");
  form.setRequireLogin(false);
  
  //return the info
  return {
    formID: form.getId(),
    formURL: form.getEditUrl()
  }
}

