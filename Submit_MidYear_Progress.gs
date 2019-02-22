/**
 *  Creates a from response with data then submits it.
 *
 * @param {obj} sheet Reference to the sheet with the data
 * @param {string} formID The ID of the Mid Year Prgress form file
 * @return {boolean}  Returns true when done
 */

function submitMidYearProgress(theData, formID){
  //get the data into an object
  var allData = ObjApp.rangeToObjects(theData);
  //get the form and form items
  var form = FormApp.openById(formID)
  var items = form.getItems();
  
  var start = new Date(); //set the start time
  
  //Loop through the data and set the responses.
  for(var i = 0; i < allData.length; i++){
    var fileName = allData[i].studlastfirst + " " + allData[i].studid;
    var resp = form.createResponse();
    resp.withItemResponse(items[0].asListItem().createResponse(fileName));
    resp.withItemResponse(items[1].asTextItem().createResponse(allData[i].studid));
    resp.withItemResponse(items[2].asTextItem().createResponse(allData[i].coursename));
    resp.withItemResponse(items[3].asTextItem().createResponse(allData[i].teacherlastfirst));
    resp.withItemResponse(items[4].asParagraphTextItem().createResponse(allData[i].goal));
    resp.withItemResponse(items[5].asParagraphTextItem().createResponse(allData[i].progress));
   
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
  
