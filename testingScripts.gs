function testGoalsubmit() {
  var ss = SpreadsheetApp.openById("11tl5J7ddt3nhKKDcgZ3iSTBKmTVdeJIicnVccj8D7TA").getSheetByName("forGoals");
  
  var allData = ss.getRange(1,1,ss.getLastRow(),ss.getLastColumn()).getValues();
  Logger.log(allData[1][0]);
  
  var midMan = "1oBGS6kAHmGWyOY-I8kJ7zFbueaisU6YvFGv_VhrY7gc";
  
  var rslt = submitMiddleManForm(allData,midMan);
  
}

function logPractice(){
  // A simple INFO log message, using sprintf() formatting.
  //console.info('Timing the %s function (%d arguments)', 'myFunction', 1);

  // Log a JSON object at a DEBUG level. If the object contains a property called "message",
  // that is used as the summary in the log viewer, otherwise a stringified version of
  // the object is used as the summary.
  var parameters = {
    isValid: true,
    content: 'some string',
    timestamp: new Date()
  };
  //console.log(parameters);
  var x = 10;
  try{
  if(x >5){
    //console.warn();
    //throw "Number is bigger than 5.";
    DocumentApp.openById("jsjsjsjs");
    }
    else{
    console.info();
    }
  }
  catch(e){
    console.log(e);
    }

}