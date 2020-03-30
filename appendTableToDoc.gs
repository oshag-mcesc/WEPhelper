/*
* Appends inforamtion to a document as a table.  This was created to replace docAppender
* 
* @param {array} info the array of data to append
* @return {obj} result this object will have "done" and "theFolderID"
*/

function appendTable(info){
 var titles = info.shift()
 var result = {};
 try{
  for(var i=0;i<=info.length-1;i++){
   // Logger.log("in loop")
    var cells = [];
    for(var j=0;j<=titles.length-2;j++){
      cells[j]=[];
      cells[j][0] = titles[j+1];
      cells[j][1] = info[i][j+1];
      }
   // Logger.log(i);
    var doc = DocumentApp.openById(info[i][0]);
    var body = doc.getBody();
    var tbl = body.appendTable(cells)
    tbl.setColumnWidth(0, 75);
    for(var k=0;k<tbl.getNumRows();k++){
      tbl.getCell(k, 0).setBackgroundColor("#E0E0E0");
    }
    doc.saveAndClose()
    }
    result.done = true;
    }//end try
 catch(err){
   console.info("Error in appending tables!");
   console.log(err);
   result.done = false;
   result.info = err;
   }
 return result;
}