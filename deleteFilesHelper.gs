/**
 * Deletes files from a  (actually just trashes them)
 *
 * @param {string} folderID The id of the folder with the files to delete
 * @return {bool} boolean True if all were moved to trash
 */
 
function deleteFiles(folderID) {
  try{
    var fld = DriveApp.getFolderById(folderID);
    var files = fld.getFiles();
    
    while (files.hasNext()) {
     var file = files.next().setTrashed(true);
    }
    return true
   }
  catch(err){
   Logger.log(err);
   return false
  }
}

 