/**
 * Add the student classes to the class table on the WEP document.
 * (NOTE: There is a timing function to catch if script runs to long)
 * @param {array} theData the array of data values
 * @param {integer} rowNum the row to start processing
 * @return {object} the code will return {rowNum, done} where ronwNum is the next row to process and done if done
 */

function addClasses(theData, rowNum) {
  //get the length of the data
  var length = theData.length;
  //create the style for the header row
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  //var for the results and timer
  var rslts = {};
  var start = new Date();
  //for loop WITH NO COUNTER... the counting is done in the loop
  for(var i = rowNum;i < length;){
    //Get new doc and then get the right table
    var tbl = getTheTable_(theData[i][0]);
    //do loop will run until i == length OR (see coments above the 'while' statement
    do{
      var tr = tbl.appendTableRow();
      for(var j = 1;j<4;j++){
        tr.appendTableCell(theData[i][j]);
      }
      i++  //the actual loop counter
      if(i == length){break};
      }
    //compare the next data to the previous data, if same keep looping
    while(theData[i-1][0] === theData[i][0]);
    tbl.getRow(0).setAttributes(headerStyle);
    //check execution time
    if(isTimeUp_(start)){
       rslts = {classRowNum: i };
       //Logger.log("Limit met!!!");
       break; //get out if time is up!!
       }    
    }
    //check if we are done
    if(i >= length){
      rslts = {classRowNum: "0"};
    }    
    return rslts;
}

//Helper function
function getTheTable_(docId){
  //get the "template" doc
  var doc = DocumentApp.openById(docId);
   //find the right table
  var tables = doc.getBody().getTables();
  
  var numberOfTables = tables.length;
  for(var i=0; i<=numberOfTables;i++){
    var cell = tables[i].getCell(0, 0).getText();
    if(cell == "Service Area"){
      //Get out!  i is the index of the table we want
      break;
    }
  }
 
  return tables[i];
}
