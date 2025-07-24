//Attempt to make this IIFE globally available
const settings = (() => {
  const ns = {};

  ns.tabs = {
      template1: 'Template 1',
      original: 'original',
      rawData: 'rawData',
      gifted: 'gifted',
      preQuery: 'preQuery',
      theGiftedData: 'theGiftedData',
      forGoalLinks: 'forGoalLinks',
      goalsSent: 'Goals Sent',
      tcg: 'TeacherCourseGifted',
      forDocs: 'forDocs',
      docIds: 'docIds',
      copiedGoals: 'copiedGoals',
      forClasses: 'forClasses',
      forGoals: 'forGoals',
      forMidYear: 'forMidYear',
      copiedMidYears: 'copiedMidYears',
      forMidYearClasses: 'forMidYearClasses',
      forMidYearPush: 'forMidYearPush',
      forFinalEvaluation: 'forFinalEvaluation',
      copiedFinals: 'copiedFinals',
      forFinalPush: 'forFinalPush',
      courseInfo: 'CourseInfo',
      fileIDs: 'fileIDs',
      emisReport: 'emisReport'
  }

  ns.resetValues = {
    maxCols: 26, //for reset, max columns
    startRow: 21, //for reset, delete from row 21 on
    clearRange: "A2:Z20", //the range to clear out
    noReset: [    //DON"T reset these tabs
      ns.tabs.template1,
      ns.tabs.tcg,
      ns.tabs.courseInfo,
      ns.tabs.emisReport
    ]
  }

  
  ns.combineTabNames = (tab1, tab2) => {
    return `${tab1} along with ${tab2}`;
  };

  return ns;
})();


// Testing code follows
const testIt = () => {
  let cols = settings.resetValues.maxCols
  let rows = settings.resetValues.startRow
  let arr = settings.resetValues.noReset
  
  console.log(`The cols are ${cols} and the rows are ${rows}`)
  console.log(arr)

  // let rslt = settings.combineTabNames(settings.tabs.rawData,settings.tabs.template1)
  // console.log(rslt)
  // SpreadsheetApp.getUi().alert(rslt)
}



//Global to be able to call from library
var testIt_g = testIt