/**
 * Creates the copiedMidYears and forFinalEvaluations tabs and sets the formulas.
 * Also updates the query funtion for the forDocs tab
 *
 * @param {SpreadSheet} aSpreadsheet the active spreadsheet
 * @return {boolean} returns done (true or flase) 
 */

function addSheetsAndUpdate(aSpreadsheet) {
  var ss = aSpreadsheet;
  //do copiedMidYears
  var name = "copiedMidYears";
  var sheetIndex = ss.getNumSheets();
  ss.insertSheet(name, sheetIndex)
  var formula = "=arrayformula(if(NOT(ISBLANK(E2:E)),VLOOKUP(E2:E,OFFSET(TeacherInfo,,1),4,false),))";
  var titles = [["New","teacherEmail"]];
  var copied = ss.getSheetByName(name);
  copied.getRange("H1:I1").setValues(titles);
  copied.getRange("I2").setFormula(formula);
  copied.setFrozenRows(1);
  
  var name2 = "forFinalEvaluation";
  var sheetIndex = ss.getNumSheets();
  ss.insertSheet(name2, sheetIndex)
  var formula = "=QUERY(copiedMidYears!B1:I,\"Select B, C, D, E, F, G, H, I where B is not null AND H = 'Y' order by E label B 'StudLastFirst', E 'teacherlastfirst', C 'studid', I 'teacheremail', F 'goal', D 'coursename'\",1)";


  ss.getSheetByName(name2).getRange("A1").setFormula(formula);
  ss.getSheetByName(name2).setFrozenRows(1);
  
  var theDocs = ss.getSheetByName("forDocs");
  if (!theDocs){
    ss.insertSheet("forDocs");
    theDocs = ss.getSheetByName("forDocs");
    }
  theDocs.getRange("A1").setFormula("=query(StudInfo,\"Select N, O, P, Q, R, S, T  where U = ''order by O\",1)");
  theDocs.setFrozenRows(1);
  
  return true;

  }
  