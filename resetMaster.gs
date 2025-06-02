/**
 * This code was added 9/12/2024 by Michael O'Shaughnessy
 * A major attempt was done to add JSDOC documentation.  
 * The doc for the function and namespace is good BUT for the settings object it is not right...
 */
/**
 * @namespace reset
 * @description A namespace containing a function to reset spreadsheet tabs based on configurations.
 */
const reset = (() => {

  /**
   * Resets spreadsheet tabs based on the settings provided in the `settingsTabs.tabs` object.
   * Loops through each tab configuration in the settings object and performs the following actions for tabs marked for reset:
   *   - Gets the sheet based on the configuration's `name` property.
   *   - Clears the sheet content.
   *   - Deletes unnecessary rows and columns based on thresholds (configurable in settings).
   *   - Sets headers if provided in the configuration's `headers` property.
   *   - Sets formulas if provided in the configuration's `formulas` property.
   */
  const resetEm = () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet
    try{
        for (const tabName in settingsTabs.tabs) {
      const tab = settingsTabs.tabs[tabName];
      if (tab.reset) {
        let sheet = ss.getSheetByName(tab.name); // Get the sheet (tab) by name

        // Batch clear content and unnecessary rows/columns
        const maxRows = sheet.getMaxRows();
        const maxColumns = sheet.getMaxColumns();
        if (maxRows > 20) {
          sheet.deleteRows(21, maxRows - 20);
        }
        if (maxColumns > 26) {
          sheet.deleteColumns(27, maxColumns - 26);
        }
        sheet.clearContents();

        // Set headers (if provided)
        if (tab.headers) {
          sheet.getRange(1, 1, 1, tab.headers.length).setValues([tab.headers]);
        }

        // Set formulas (if provided)
        if (tab.formulas) {
          for (const cell in tab.formulas) {
            sheet.getRange(cell).setFormula(tab.formulas[cell]);
          }
        }
      }
    }
    }
    catch(err){
      logIt({
        level:"severe",
        theMsg:"Problem with resetting sheets",
        error: err})
    }
  
  };

  return {
    resetEm,
  };
})();

//so it can be called from the menu
var reset1= reset

//-----------------------------------
//the settings object for all the tabs
//-----------------------------------
//attempted to use JSDOC... but it is not all correct.
/**
 * @typedef {Object} TabConfiguration
 * @property {string} name
 * @property {boolean} reset
 * @property {string[]} [headers]
 * @property {Object.<string, string>} [formulas]
 */
const settingsTabs = (() => {
  const ns = {}
  /**
   * @property {TabConfiguration[]} tabs
   * @description An array of tab configurations.
   */
  ns.tabs = {
    template1: {
      name: 'Template 1',
      reset: false,
    },
    original: {
      name: 'ORIGINAL',
      reset: true,
      headers: ['SchoolCode', 'StudentId', 'StateStudentId', 'FirstName', 'LastName', 'Gender', 'GradeLevelCode', 'HomeSchoolIRN', 'HomeSchool', 'ProgramCode', 'ProgramName', 'Status', 'StudentStatusCode', 'CourseCode', 'CourseName', 'CourseTypeCode', 'CourseTypeDescription', 'SectionNumber', 'TermCode', 'TermName', 'LocationEx', 'TeacherCode', 'CalendarPeriodCode', 'RotationDays']
    },
    rawData: {
      name: 'rawData',
      reset: true,
      formulas: {
        A1: `=filter(original!A1:X,MATCH(original!N1:N&"",index(CourseInfo,,1),0))`
      }
    },
    gifted: {
      name: 'gifted',
      reset: true,
      headers: ['SchoolCode', 'StudentId', 'StateStudentId', 'FirstName', 'LastName', 'Gender', 'GradeLevelCode', 'HomeSchoolIRN', 'HomeSchool', 'ProgramCode', 'ProgramName', 'Status', 'StudentStatusCode', 'CourseCode', 'CourseName', 'CourseTypeCode', 'CourseTypeDescription', 'SectionNumber', 'TermCode', 'TermName', 'LocationEx', 'TeacherCode', 'CalendarPeriodCode', 'RotationDays'],
      formulas: {
        A2: `=filter(rawdata!A2:X,MATCH(rawdata!B2:B&"",index(StudInfo,,1),0))`
      }
    },
    preQuery: {
      name: 'preQuery',
      reset: true,
      formulas: {
        A1: `=query(gifted!A1:W,"select E, D, F, B, G, N, O, A, V,W where A is not null label G 'Grade', E 'StudLast', D 'StudFirst', B 'studentnumber', O 'Course', W 'Period'", 1)`,
        K1: `={"GiftedArea","GiftedCourse";arrayformula(if(isblank(D2:D),,VLOOKUP(D2:D,StudInfo,6,false))),arrayformula(if(isblank(F2:F),,VLOOKUP(F2:F,CourseInfo,4,false)))}`,
        N1: `={"StudLastFirst","TeacherLastFirst","TeacherEmail";arrayformula(if(isblank(D2:D),,VLOOKUP(D2:D,StudInfo,2,false))),arrayformula(if(isblank(I2:I),,VLOOKUP(I2:I,TeacherInfo,2,false))),arrayformula(if(isblank(I2:I),,VLOOKUP(I2:I,TeacherInfo,5,false)))}`
      }
    },
    theGiftedData: {
      name: 'theGiftedData',
      reset: true,
      formulas: {
        A1: `=query(preQuery!A1:P,"Select * where M <> 'NONE'and A is not null order by A",1)`,
        Q1: `="New"`
      }
    },
    forGoalLinks: {
      name: 'forGoalLinks',
      reset: true,
      formulas: {
        A1: `=query(theGiftedData!A1:Q,"Select N,D,M,O,G,P where A<>'' and Q ='Y' order by O,N",1)`
      }
    },
    tcg: {
      name: 'TeacherCourseGifted',
      reset: false,
    },
    forDocs: {
      name: 'forDocs',
      reset: true,
      formulas: {
        A1: `=query(StudInfo,"Select N, O, P, Q, R, S, T  where U is null order by O",1)`
      }
    },
    docIds: {
      name: 'docIds',
      reset: true,
    },
    copiedGoals: {
      name: 'copiedGoals',
      reset: true,
      formulas: {
        J1: `={"ODEISandcode";FILTER(ARRAYFORMULA(VLOOKUP(F2:F,OFFSET(CourseInfo,,1),2,false)),NOT(ISBLANK(C2:C)))}`,
        K1: `="New"`,
        L1: `={"teacherEmail";arrayformula(if(NOT(ISBLANK(E2:E)),VLOOKUP(E2:E,OFFSET(TeacherInfo,,1),4,false),))}`,
      }
    },
    forClasses: {
      name: 'forClasses',
      reset: true,
      formulas: {
        A1: `=query(copiedGoals!B1:K,"Select B, C, F, J where B is not null AND K ='Y' order by B", 1)`,
        E1: `={"codeDescription","docID","courseGifted";filter(ARRAYFORMULA(VLOOKUP(D2:D+0,emisDescriptions,3,false)),NOT(ISBLANK(A2:A))),filter(ARRAYFORMULA(VLOOKUP(B2:B&"",StudInfo,8,false)),NOT(ISBLANK(A2:A))),filter(ARRAYFORMULA(VLOOKUP(C2:C,offset(CourseInfo,,1),3,false)),NOT(ISBLANK(A2:A)))}`,
        I1: `=QUERY(A1:G,"Select F, G, C, E where F is not null order by A",1)`
      }
    },
    forGoals: {
      name: 'forGoals',
      reset: true,
      formulas: {
        A1: `=QUERY(copiedGoals!B1:K,"Select * where B is not null AND K='Y' order by B label B 'studlastfirst', C 'studentnumber', D 'giftedarea', E 'teacherlastfirst', F 'coursename',G 'instructionalstrategies', H 'assesstools', J 'odeisandcode'",1)`,
        L1: `={"docID";FILTER(ARRAYFORMULA(VLOOKUP(B2:B&"",StudInfo,8,false)),NOT(ISBLANK(A2:A)))}`,
        M1: `={"Teacher","Class / Course","Content Goal";D2:D,E2:E,H2:H}`
      }
    },
    forMidYear: {
      name: 'forMidYear',
      reset: true,
      formulas: {
        A1: `=QUERY(copiedGoals!B1:L,"Select B, C, E, F, I, L where B is not null AND K='Y' order by E label B 'StudLastFirst', E 'teacherlastfirst', C 'studentnumber', L 'teacheremail', I 'goal', F 'coursename'",1)`,
      }
    },
    copiedMidYears: {
      name: 'copiedMidYears',
      reset: true,
      formulas: {
        H1: `="New"`
      }
    },
    forMidYearPush: {
      name: 'forMidYearPush',
      reset: true,
      formulas: {
        A1: `=QUERY(copiedMidYears!B2:H,"Select B, C, D, E, F, G  where B is not null AND H = 'Y' order by B label B 'studlastfirst', C 'studentnumber', D 'coursename', E 'teacherlastfirst',F 'goal',G 'progress'",1)`,
        I1: `={"docID";FILTER(ARRAYFORMULA(VLOOKUP(B2:B&"",StudInfo,8,false)),NOT(ISBLANK(A2:A)))}`,
        J1: `={"Teacher","Class / Course","Progress";D2:D,C2:C,F2:F}`
      }
    },
    forFinalEvaluation: {
      name: 'forFinalEvaluation',
      reset: true,
      formulas: {
        A1: `=QUERY(copiedMidYears!B1:I,"Select B, C, D, E, F, G, H, I where B is not null AND H = 'Y' order by E label B 'StudLastFirst', E 'teacherlastfirst', C 'studentnumber', I 'teacheremail', F 'goal', D 'coursename'",1)`
      }
    },
    copiedFinals: {
      name: 'copiedFinals',
      reset: true,
      formulas: {
        I1: `="New"`
      }
    },
    forFinalPush: {
      name: 'forFinalPush',
      reset: true,
      formulas: {
        A1: `=QUERY(copiedFinals!B1:I,"Select B, C, D, E, F, G, H  where B is not null AND I = 'Y' order by B label B 'studlastfirst', C 'studentnumber', D 'coursename', E 'teacherlastfirst',G 'progress'",1)`,
        J1: `={"docID";FILTER(ARRAYFORMULA(VLOOKUP(B2:B&"",StudInfo,8,false)),NOT(ISBLANK(A2:A)))}`,
        K1: `={"Teacher","Class / Course","Progress";D2:D,C2:C,F2:F}`
      }
    },
  }

  return ns
})()