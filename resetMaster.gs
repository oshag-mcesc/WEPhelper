/**
 * Reset functionality for Google Sheets - prepares spreadsheets for new school year
 * Created: 9/12/2024 by Michael O'Shaughnessy
 * Optimized: Performance improvements with batched operations
 */

/**
 * @namespace reset
 * @description Namespace containing functions to reset spreadsheet tabs based on configurations.
 * Creates missing required tabs and resets existing tabs to their initial state for new school year.
 */
const reset = (() => {

  /**
   * Main reset function that prepares the spreadsheet for a new school year.
   * Creates missing required tabs (especially config) and resets existing tabs based on configuration.
   * Uses batched operations for optimal performance.
   * 
   * @function resetEm
   * @memberof reset
   * @description Resets spreadsheet tabs based on the settings provided in the `settingsTabs.tabs` object.
   * Performs the following actions:
   *   - Creates missing required tabs (config tab with proper setup)
   *   - Clears content from tabs marked for reset
   *   - Resizes sheets to standard dimensions (20 rows, 26 columns)
   *   - Sets headers if provided in configuration
   *   - Sets formulas if provided in configuration
   * 
   * @throws {Error} Logs severe errors via logIt function if reset process fails
   */
  const resetEm = () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    try {
      // Notify user that reset process has started
      SpreadsheetApp.getActiveSpreadsheet().toast("Resetting spreadsheet for new year...", "Starting Reset", -1);

      // Phase 1: Create any missing required tabs (especially config tab)
      createMissingTabs(ss);

      // Phase 2: Reset existing tabs with optimized batched operations
      resetExistingTabs(ss);

      // Notify user of successful completion
      SpreadsheetApp.getActiveSpreadsheet().toast("Reset complete!", "Finished", 3);

    } catch (err) {
      logIt({
        level: "severe",
        theMsg: "Problem with resetting sheets",
        error: err
      });
    }

    /**
     * Creates missing tabs that are marked as required in the configuration.
     * Checks for tabs that have either createIfMissing: true or reset: true.
     * 
     * @function createMissingTabs
     * @inner
     * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet
     */
    function createMissingTabs(ss) {
      for (const tabName in settingsTabs.tabs) {
        const tab = settingsTabs.tabs[tabName];

        // Create tab if it's marked as createIfMissing or required for reset
        if (tab.createIfMissing || tab.reset) {
          let sheet = ss.getSheetByName(tab.name);

          if (!sheet) {
            console.log(`Creating missing required tab: ${tab.name}`);
            sheet = ss.insertSheet(tab.name);

            // Special setup for config tab
            if (tabName === 'config') {
              setupConfigTab(sheet, tab);
            }
          }
        }
      }
    }

    /**
     * Processes all tabs marked for reset using batched operations for performance.
     * Collects all sheets first, then processes each with error handling per sheet.
     * 
     * @function resetExistingTabs
     * @inner
     * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet
     */
    function resetExistingTabs(ss) {
      const sheetsToProcess = [];

      // Step 1: Collect all sheets that need resetting and validate they exist
      for (const tabName in settingsTabs.tabs) {
        const tab = settingsTabs.tabs[tabName];
        if (tab.reset) {
          const sheet = ss.getSheetByName(tab.name);
          if (sheet) {
            sheetsToProcess.push({ sheet, config: tab });
          } else {
            console.log(`Warning: Tab "${tab.name}" marked for reset but not found`);
          }
        }
      }

      // Step 2: Process each sheet with batched operations and individual error handling
      sheetsToProcess.forEach(({ sheet, config }) => {
        try {
          resetSingleTab(sheet, config);
        } catch (err) {
          logIt({
            level: "severe",
            theMsg: `Error resetting tab: ${sheet.getName()}`,
            error: err
          });
        }
      });
    }

    /**
     * Resets a single tab with batched operations for optimal performance.
     * Handles resizing, clearing, and populating the sheet based on configuration.
     * 
     * @function resetSingleTab
     * @inner
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to reset
     * @param {TabConfiguration} tabConfig - Configuration object for the tab
     */
    function resetSingleTab(sheet, tabConfig) {
      // Batch operation 1: Resize sheet to standard dimensions (more efficient before clearing)
      const maxRows = sheet.getMaxRows();
      const maxColumns = sheet.getMaxColumns();

      if (maxRows > 20) {
        sheet.deleteRows(21, maxRows - 20);
      }
      if (maxColumns > 26) {
        sheet.deleteColumns(27, maxColumns - 26);
      }

      // Clear all existing content
      sheet.clearContents();

      // Batch operation 2: Set headers if provided in configuration
      if (tabConfig.headers) {
        sheet.getRange(1, 1, 1, tabConfig.headers.length).setValues([tabConfig.headers]);
      }

      // Batch operation 3: Set formulas if provided in configuration
      if (tabConfig.formulas) {
        for (const cell in tabConfig.formulas) {
          sheet.getRange(cell).setFormula(tabConfig.formulas[cell]);
        }
      }
      // Batch operation 4: Set notes if provided in configuration
      if (tabConfig.notes) {
        for (const cell in tabConfig.notes) {
          sheet.getRange(cell).setNote(tabConfig.notes[cell]);
        }
      }
    }

    /**
     * Sets up the config tab with proper structure, positioning, and visibility.
     * Creates the Settings/Value structure needed for the Settings class.
     * 
     * @function setupConfigTab
     * @inner
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The config sheet to set up
     * @param {TabConfiguration} tabConfig - Configuration object containing keys to populate
     */
    function setupConfigTab(sheet, tabConfig) {
      // Set up column headers for Settings class compatibility
      sheet.getRange('A1:B1').setValues([['Setting', 'Value']]);
      sheet.getRange('A1:B1').setFontWeight('bold');

      // Populate config keys (values will be set later via Settings class)
      for (let i = 0; i < tabConfig.keys.length; i++) {
        sheet.getRange(i + 2, 1).setValue(tabConfig.keys[i]);
      }

      // Resize config sheet to minimal necessary size (4 columns, 20 rows)
      const maxRows = sheet.getMaxRows();
      const maxColumns = sheet.getMaxColumns();

      if (maxRows > 20) {
        sheet.deleteRows(21, maxRows - 20);
      }
      if (maxColumns > 4) {
        sheet.deleteColumns(5, maxColumns - 4);
      }

      // Position config tab at the beginning for easy access during development
      sheet.getParent().moveActiveSheet(0);

      // Hide the config sheet to prevent accidental user modifications
      if (tabConfig.hideAfterReset) {
        sheet.hideSheet();
      }
    }
  };

  // Return public interface
  return {
    resetEm
  };
})();

// Global reference for menu system access
var reset1 = reset;

//=============================================================================
// TAB CONFIGURATION SETTINGS
//=============================================================================

/**
 * @typedef {Object} TabConfiguration
 * @property {string} name - The actual name of the sheet tab
 * @property {boolean} reset - Whether this tab should be reset during yearly reset
 * @property {boolean} [createIfMissing] - Whether to create this tab if it doesn't exist
 * @property {boolean} [hideAfterReset] - Whether to hide this tab after reset/creation
 * @property {string[]} [headers] - Array of column headers to set in row 1
 * @property {Object.<string, string>} [formulas] - Object mapping cell addresses to formulas
 * @property {string[]} [keys] - Array of setting keys (used for config tab)
 */

/**
 * @namespace settingsTabs
 * @description Configuration object defining all spreadsheet tabs and their reset behavior.
 * Used by the reset system to determine which tabs to create, reset, and how to populate them.
 */
const settingsTabs = (() => {
  const ns = {};

  /**
   * @property {Object.<string, TabConfiguration>} tabs
   * @description Object containing all tab configurations indexed by tab identifier.
   * Each tab configuration defines how that tab should be handled during reset operations.
   */
  ns.tabs = {
    config: {
      name: 'config',
      reset: false,
      createIfMissing: true,
      hideAfterReset: true,
      keys: ["WEPtemplateID", "MainWEPfolderID", "InitialGoalsFormID", "MidYearProgressFormID", "FinalProgressFormID", "rowNum", "subject", "greeting", "closing", "linktime"]
    },
    copiedFinals: {
      name: 'copiedFinals',
      reset: true,
      formulas: {
        I1: `="New"`
      },
      notes: {
        I1: `Use F for Final push`
      }
    },
    copiedGoals: {
      name: 'copiedGoals',
      reset: true,
      formulas: {
        J1: `={"ODEISandcode";FILTER(ARRAYFORMULA(VLOOKUP(F2:F,OFFSET(CourseInfo,,1),2,false)),NOT(ISBLANK(C2:C)))}`,
        K1: `="New"`,
        L1: `={"teacheremail";ARRAYFORMULA(IF(A2:A="","",IF(C2:C="","MISSING ID",IFERROR(VLOOKUP(E2:E, OFFSET(TeacherInfo,,1), 4, FALSE),"MISSING EMAIL"))))}}`,
        U1: `=QUERY(B1:L, "Select C, F, B, E where C is not null",1)`,
        K1: `=New`
      },
      notes: {
        K1: `Use Y to push classes and goals`
      }
    },
    copiedMidYears: {
      name: 'copiedMidYears',
      reset: true,
      formulas: {
        H1: `="New"`,
        I1: `={"teacheremail";ARRAYFORMULA(IF(A2:A="","",IF(C2:C="","MISSING ID",IFERROR(VLOOKUP(E2:E, OFFSET(TeacherInfo,,1), 4, FALSE),"MISSING EMAIL"))))}`,
        U1: `=QUERY(B1:L, "Select C, D, B, E where C is not null",1)`
      },
      notes: {
        H1: `Use F for final links\nUse M for midyear push`
      }
    },
    docIds: {
      name: 'docIds',
      reset: true
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
    forDocs: {
      name: 'forDocs',
      reset: true,
      formulas: {
        A1: `=query(StudInfo,"Select N, O, P, Q, R, S, T  where U is null order by O",1)`,
        H1: `={"filename";arrayformula(if(NOT(ISBLANK(A2:A)),B2:B & " " & A2:A,))}`
      }
    },
    forFinalEvaluation: {
      name: 'forFinalEvaluation',
      reset: true,
      formulas: {
        A1: `=QUERY(copiedMidYears!B1:I,"Select B, C, D, E, F, G, H, I where B is not null AND H = 'F' order by E label B 'StudLastFirst', E 'teacherlastfirst', C 'studentnumber', I 'teacheremail', F 'goal', D 'coursename'",1)`
      }
    },
    forFinalPush: {
      name: 'forFinalPush',
      reset: true,
      formulas: {
        A1: `=QUERY(copiedFinals!B1:I,"Select B, C, D, E, F, G, H  where B is not null AND I = 'F' order by B label B 'studlastfirst', C 'studentnumber', D 'coursename', E 'teacherlastfirst',G 'progress'",1)`,
        J1: `={"docID";FILTER(ARRAYFORMULA(VLOOKUP(B2:B&"",StudInfo,8,false)),NOT(ISBLANK(A2:A)))}`,
        K1: `={"Teacher","Class / Course","Final Evaluation";D2:D,C2:C,G2:G}`
      },
      rngCell: "J1",
      rowTitle: "Final Evaluation",
      statusCol: 15
    },
    forGoalLinks: {
      name: 'forGoalLinks',
      reset: true,
      formulas: {
        A1: `=query(theGiftedData!A1:Q,"Select N,D,M,O,G,P where A<>'' and Q ='Y' order by O,N",1)`
      }
    },
    forGoals: {
      name: 'forGoals',
      reset: true,
      formulas: {
        A1: `=QUERY(copiedGoals!B1:K,"Select * where B is not null AND K='Y' order by B label B 'studlastfirst', C 'studentnumber', D 'giftedarea', E 'teacherlastfirst', F 'coursename',G 'instructionalstrategies', H 'assesstools', J 'odeisandcode'",1)`,
        L1: `={"docID";FILTER(ARRAYFORMULA(VLOOKUP(B2:B&"",StudInfo,8,false)),NOT(ISBLANK(A2:A)))}`,
        M1: `={"Teacher","Class / Course","Instructional Strategies for Service","Measures to Determine Student Growth","Content Goal";D2:H}`
      }
    },
    forMidYear: {
      name: 'forMidYear',
      reset: true,
      formulas: {
        A1: `=QUERY(copiedGoals!B1:L,"Select B, C, E, F, I, L where B is not null AND K='Y' order by E label B 'StudLastFirst', E 'teacherlastfirst', C 'studentnumber', L 'teacheremail', I 'goal', F 'coursename'",1)`
      }
    },
    forMidYearPush: {
      name: 'forMidYearPush',
      reset: true,
      formulas: {
        A1: `=QUERY(copiedMidYears!B1:H,"Select B, C, D, E, F, G  where B is not null AND H = 'Y' order by B label B 'studlastfirst', C 'studentnumber', D 'coursename', E 'teacherlastfirst',F 'goal',G 'progress'",1)`,
        I1: `={"docID";FILTER(ARRAYFORMULA(VLOOKUP(B2:B&"",StudInfo,8,false)),NOT(ISBLANK(A2:A)))}`,
        J1: `={"Teacher","Class / Course","Progress";D2:D,C2:C,F2:F}`
      },
      rngCell: "I1",
      rowTitle: "Mid Year Progress",
      statusCol: 14
    },
    gifted: {
      name: 'gifted',
      reset: true,
      formulas: {
        A1: `=filter(rawdata!A1:X,{TRUE; MATCH(rawdata!B2:B&"",index(StudInfo,,1),0)})`
      }
    },
    original: {
      name: 'ORIGINAL',
      reset: true,
      headers: ['SchoolCode', 'StudentNumber', 'StateStudentId', 'FirstName', 'LastName', 'Gender', 'GradeLevelCode', 'HomeSchoolIRN', 'HomeSchool', 'ProgramCode', 'ProgramName', 'Status', 'StudentStatusCode', 'CourseCode', 'CourseName', 'CourseTypeCode', 'CourseTypeDescription', 'SectionNumber', 'TermCode', 'TermName', 'LocationEx', 'TeacherCode', 'CalendarPeriodCode', 'RotationDays']
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
    rawData: {
      name: 'rawData',
      reset: true,
      formulas: {
        A1: `=filter(original!A1:X,MATCH(original!N1:N&"",index(CourseInfo,,1),0))`
      }
    },
    tcg: {
      name: 'TeacherCourseGifted',
      reset: false
    },
    template1: {
      name: 'Template 1',
      reset: false
    },
    theGiftedData: {
      name: 'theGiftedData',
      reset: true,
      formulas: {
        A1: `=query(preQuery!A1:P,"Select * where M <> 'NONE'and A is not null order by A",1)`,
        Q1: `="New"`,
        U1: `={D1:D,G1:G,N1:O}`
      }
    }
  };

  return ns;
})();

