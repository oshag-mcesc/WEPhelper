/**
 * Creates a pre filled URL for the initial goals form
 *
 * @param {obj} sheet Reference to the sheet with the data
 * @param {string} formID The ID of the Initial Goals form file
 * @return {boolean}  Returns true when done
 */

function preFilledGoalsURLs(sheet, theData, formID){
  //get the data into an object
  var allData = ObjApp.rangeToObjects(theData);
  //get the form and form items
  var form = FormApp.openById(formID)
  var items = form.getItems();
  var preFilledURL = [];  //needed to hold the URLs
  
  //Loop through the data and set the responses. 
  //AND add the response to the PrefilledURL arrat
  for(var i = 0; i < allData.length; i++){
    var resp = form.createResponse();
    resp.withItemResponse(items[0].asTextItem().createResponse(allData[i].studlastfirst));
    resp.withItemResponse(items[1].asTextItem().createResponse(allData[i].studid));
    resp.withItemResponse(items[2].asTextItem().createResponse(allData[i].served));
    resp.withItemResponse(items[3].asTextItem().createResponse(allData[i].teacherlastfirst));
    resp.withItemResponse(items[4].asTextItem().createResponse(allData[i].course));
    preFilledURL.push([resp.toPrefilledUrl()]);  
  }
  
  //Add the URLs to the spreadsheet.  
  var lcol = sheet.getLastColumn()+1;
  sheet.getRange(1, lcol).setValue('PreFilledURL');
  sheet.getRange(2, lcol, allData.length).setValues(preFilledURL);
  
  sheet.getRange(1, lcol + 1).setValue('Hyperlink');
  sheet.getRange(2, lcol + 1).setFormula("=arrayformula(HYPERLINK(G2:G,A2:A))"); //NEED a way to offset this!!
  
  return true;
}