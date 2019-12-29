/*
*This is the main code to show a sidebar to toggle visibility of tabs
*/

function pickTabs(){
  var tmp = HtmlService.createTemplateFromFile('tablist');
  var cbos = createCheckBoxes();
  tmp.cboxes = cbos;
  var thePage = tmp.evaluate()
  //Logger.log(thePage.getContent());
  SpreadsheetApp.getUi().showSidebar(thePage);
}

function createCheckBoxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = ss.getSheets();
  var tmp = [];
  var cbo = tabs.map(function (tab){
    
    if(tab.isSheetHidden()){
      return ('<input type="checkbox" name="theTabs" value="' + tab.getName() + '" checked>' +tab.getName()+'<br>') ;
    }
    else{
      return ('<input type="checkbox" name="theTabs" value="' + tab.getName() + '">' +tab.getName()+'<br>');
    }   
    }).join(' ');

  return cbo;
}

function toggleEm(data){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  data.forEach(function(row){
    var sht = ss.getSheetByName(row[0]);
    toggle_(sht,row[1]);
  });
}

function toggle_(sheet,state){
  switch (state){
    case true:
      sheet.hideSheet();
      break;
    case false:
      sheet.showSheet();
      break;
  }
}

//get css and js into the web page
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}