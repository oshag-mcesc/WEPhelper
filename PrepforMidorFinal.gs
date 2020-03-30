/**
 * Creates a horizontal rule and adds a title for the next section
 *
 * @param {array} allData the array of data values
 * @return {obj} returns done (true or flase) if false then adds studlast (the last name of the student) and row (about where it got to)
 */

function prepForMidOrFinal(allData) {
  var ids = ObjApp.rangeToObjects(allData);
  var start = new Date();  //set the start time for the script
  
  var ui = SpreadsheetApp.getUi();
  var theTitle = getAnswer_(ui,"What's the title?", "Please enter the title for new section.  For example '2018-19 Mid Year Progress'.");
  //var startRow = parseInt(getAnswer_(ui,"What's the start row?", "Please enter the start row.  If this is first run enter 0."));
  var startRow = 0;
  
  for (var i= startRow; i < ids.length;  i++){
    var doc = DocumentApp.openById(ids[i].docid);
    var body = doc.getBody();
    
    body.appendHorizontalRule();
    body.appendParagraph(theTitle)
    .setHeading(DocumentApp.ParagraphHeading.HEADING3)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    doc.saveAndClose();
    
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
  return {done: true};
  }
}

/**
 * UPDATED!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
 * Creates a horizontal rule and adds a title for the next section
 *
 * @param {array} docIDs array of file IDs
 * @return {obj} returns done (true or flase) if false then adds studlast (the last name of the student) and row (about where it got to)
 */

function addMidOrFinalDivider(docIDs) { 
  var start = new Date();  //set the start time for the script
  
  var ui = SpreadsheetApp.getUi();
  var theTitle = getAnswer_(ui,"What's the title?", "Please enter the title for new section.  For example '2018-19 Mid Year Progress'.");
  //var startRow = parseInt(getAnswer_(ui,"What's the start row?", "Please enter the start row.  If this is first run enter 0."));
  var startRow = 0;
  var numIDs = docIDs.length;
  for (var i= startRow; i < numIDs;  i++){
    var doc = DocumentApp.openById(docIDs[i]);
    var body = doc.getBody();
    
    body.appendHorizontalRule();
    body.appendParagraph(theTitle)
    .setHeading(DocumentApp.ParagraphHeading.HEADING3)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    doc.saveAndClose();
    
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
  
  return {done: true};
  }
}
