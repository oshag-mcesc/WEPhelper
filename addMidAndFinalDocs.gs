/*
* Creates the Mid or Final folder then copies all the docs
* 
* @param {array} docIDs the array of file IDs
* @param {string} destination the ID of the PARENT folder 
* @param {string} fldrTitle this will be the title of the new folder (ie "Mid Year", "Final")
* @param {string} MidOrFinal this will be added to file name (ie "Mid Year", "Final")
* @param {string} sheet the sheet to return info to
* @param {string} folderID the ID of the folder IF it exists
* @return {obj} info this object will have "done" and "theFolderID"
*/

function createMidOrFinalDocs(docIDs,destination, fldrTitle, MidOrFinal,sheet, folderID){
  var info = {};
  info.done = false;
  var theFolderID = "";
  try{
    //do we need to create the folder
    if(!folderID){
      //If not id then create the new folder
      theFolderID = CreateFolder_(destination, fldrTitle);
      }
    else{
      //already exists
      theFolderID = folderID;
      }
    //copy docs and move to new folder
    var rslt = copyDocs_(docIDs, theFolderID, MidOrFinal);
    //Get a list of the docs
    getListOfDocIds(theFolderID, sheet);
    info.theFolderID = theFolderID;
    info.done = true;
  }
  catch(err){
    console.log(err, "Error in creating folder and mid year docs.");
    }
  return info;
}

/*
* Given a list of IDs this will delete everything after "District Policies"
* 
* @param {array} theIds the array of file IDs
* @return {booliean} done true if done, false otherwise
*/

function deleteText(theIDs){
  
  var done = false; //flag to tell if it is done
  
  try{
  theIDs.forEach(function(docID){
    var doc = DocumentApp.openById(docID);
    var body = doc.getBody();
    var paras = body.getParagraphs();
    var number = paras.length;
    for(var i = 10; i<number-1;i++){
      if(paras[i].getText() == "District Policies"){break;};
      }
    
    var paraIndex = body.getChildIndex(paras[i]);
    var totalChildren = body.getNumChildren();
    
    theDeleter_(doc,paraIndex);
  
  });
  done = true;
  } //end of Try
  catch(err){
    console.error("Removing text after 'District Policies' error %s", err);
  } //end of catch
  
  return done;
}

/*
* Given a list of IDs this will delete everything from the last HR on
* 
* @param {array} theIds the array of file IDs
* @return {booliean} done true if done, false otherwise
*/
//THIS needs to be in WEPhelpers... maybe need a better name?
function deleteTextHR(theIDs){
  
  var done = false; //flag to tell if it is done
  
  try {
    theIDs.forEach(function (docID) {
      var doc = DocumentApp.openById(docID);
      var body = doc.getBody();

      var paras = body.getParagraphs();
      var hzRule = [];
      var num = paras.length;
      for (var i = 0; i < num; i++) {
        if (paras[i].getNumChildren() > 0) {
          var fChild = paras[i].getChild(0);
          if (fChild.getType() === DocumentApp.ElementType.HORIZONTAL_RULE) {
            hzRule.push(i);
          }
        }
      }
      
      var stopIndex = doc.getBody().getChildIndex(paras[hzRule[hzRule.length-1]]);
      theDeleter_(doc,stopIndex)
      //console.log(stopIndex);
    });
    done = true;
  } //end of Try
  catch (err) {
    console.error("Removing text after HR error %s", err);
  } //end of catch

  return done;
}

//****************************************HELPER FUNCTIONS
//Creates a new folder

function CreateFolder_(folderID, fldrName){
  var newFolder = DriveApp.getFolderById(folderID).createFolder(fldrName);
  return newFolder.getId();
}

/*
* Given a list of IDs this will delete everything after "District Policies"
* 
* @param {array} theIds the array of file IDs
* @param {string} destFolderID the ID of the folder to copy into
* @param {string} MidFinal this will be added to file name (ie "Mid Year", "Final")
* @return {booliean} done true if done, false otherwise
*/


function copyDocs_(theIDs, destFolderID, MidFinal) {
  var done = false;
  try{
  var dest = DriveApp.getFolderById(destFolderID);
  theIDs.forEach(function(docID){
    var doc = DriveApp.getFileById(docID);
    doc.makeCopy(doc.getName() + " - " + MidFinal, dest);
    });
  done = true;
  } //end of Try
  catch(err){
    console.error("Making copies of the docs error: %s, %s", destFolderID, err);
    return done;
  } //of of catch
  
  return done;
}

//Will delete text from a document starting from the end and working backwards to a 'stopelement'
function theDeleter_(doc,stopElement) {
  
  var body = doc.getBody();
  //add a very last element (made it eaiser to delete... look at other comments
  body.appendParagraph("DELETE ME!!");
  //get number of childres
  var totChildren = body.getNumChildren();
  //Because of 0 index AND that you cannot delete the last element of a doc we have to start
  //at 2 less then the number of children.  Then work backwards to the current element
  //Had to go backwards because it automatically updates the number of children when one is added or removed
  for(var i = totChildren-2; i>=stopElement; i--){
    body.removeChild(body.getChild(i));
    var newNum = body.getNumChildren();
    }
  //Set the text of the las element to nothing
  body.getChild(body.getNumChildren()-1).asText().setText('');
}
