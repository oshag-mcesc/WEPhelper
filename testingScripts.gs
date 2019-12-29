function testGoalsubmit() {
  var ss = SpreadsheetApp.openById("11tl5J7ddt3nhKKDcgZ3iSTBKmTVdeJIicnVccj8D7TA").getSheetByName("forGoals");
  
  var allData = ss.getRange(1,1,ss.getLastRow(),ss.getLastColumn()).getValues();
  Logger.log(allData[1][0]);
  
  var midMan = "1oBGS6kAHmGWyOY-I8kJ7zFbueaisU6YvFGv_VhrY7gc";
  
  var rslt = submitMiddleManForm(allData,midMan);
  
}
