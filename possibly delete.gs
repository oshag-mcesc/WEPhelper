/**
 * Adds files to sub folders
 *
 * @param {array} studData the array of data values from the named range
 * @param {array} folderData the array of data values from the named range
 * @return {obj} theInfo has done, laststudent and row number
 */

function addFiletoFolder(studData, folderData, startRow){
    if(startRow ==='undefined'|| startRow === null){startRow = 1;} 
    var grade = studData[0].indexOf("Grade");
    var docID = studData[0].indexOf("DocID");
    var theFolder = folderData[0].indexOf("FolderID");
    var include = folderData[0].indexOf("Include");
    var studDataLength = studData.length;
    var folderDataLength = folderData.length;
    
    var start = new Date();
    
    for (var i = startRow; i < studDataLength ;i++){
      for(var j = 1; j < folderDataLength ; j++){
        if(studData[i][grade]==folderData[j][include]){
          
          DriveApp.getFolderById(folderData[j][theFolder])
             .addFile(DriveApp.getFileById(studData[i][docID]));
          break;
          }
        }
        //check execution time
        if(isTimeUp_(start)){
         var theInfo = {
         done: false,
         lastStudent: studData[i][1],
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




