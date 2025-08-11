/**
 * Consolidated URL creation functionality for WEPhelper library
 * Replaces three separate functions with one configurable function
 */

// Configuration for URL creation types
const URL_CONFIG = {
  initial: {
    sheetName: "forGoalLinks",
    formIdSetting: "InitialGoalsFormID",
    titleAdd: "",
    linkFormula: `=arrayformula(if(A2:A<>"",HYPERLINK({preFilledCol}2:{preFilledCol},A2:A&" - "&E2:E),))`,
    fields: [
      { dataField: "studlastfirst", itemType: "TEXT" },
      { dataField: "studentnumber", itemType: "TEXT" },
      { dataField: "served", itemType: "TEXT" },
      { dataField: "teacherlastfirst", itemType: "TEXT" },
      { dataField: "course", itemType: "TEXT" }
    ]
  },
  mid: {
    sheetName: "forMidYear",
    formIdSetting: "MidYearProgressFormID",
    titleAdd: "Mid Year",
    linkFormula: `=filter(arrayformula(HYPERLINK({preFilledCol}2:{preFilledCol},A2:A&" - "&D2:D)),A2:A<>"")`,
    fields: [
      { dataField: "fileName", itemType: "TEXT" }, // Special case - constructed
      { dataField: "studentnumber", itemType: "TEXT" },
      { dataField: "coursename", itemType: "TEXT" },
      { dataField: "teacherlastfirst", itemType: "TEXT" },
      { dataField: "goal", itemType: "PARAGRAPH_TEXT" }
    ]
  },
  final: {
    sheetName: "forFinalEvaluation",
    formIdSetting: "FinalProgressFormID",
    titleAdd: "Final",
    linkFormula: `=filter(arrayformula(HYPERLINK({preFilledCol}2:{preFilledCol},A2:A&" - "&C2:C)),A2:A<>"")`,
    fields: [
      { dataField: "fileName", itemType: "TEXT" }, // Special case - constructed
      { dataField: "studentnumber", itemType: "TEXT" },
      { dataField: "coursename", itemType: "TEXT" },
      { dataField: "teacherlastfirst", itemType: "TEXT" },
      { dataField: "goal", itemType: "PARAGRAPH_TEXT" },
      { dataField: "progress", itemType: "PARAGRAPH_TEXT" }
    ]
  }
};

/**
 * Main URL creation namespace
 */
const nsURLCreator = (function() {
  
  /**
   * Unified function to create pre-filled URLs for any form type
   * @param {string} urlType - Type of URLs to create: 'initial', 'mid', 'final'
   * @returns {boolean} - Success status
   */
  function createPrefilledURLs(urlType) {
    try {
      // Validate URL type
      if (!URL_CONFIG[urlType]) {
        throw new Error(`Invalid URL type: ${urlType}. Must be 'initial', 'mid', or 'final'.`);
      }
      
      const config = URL_CONFIG[urlType];
      
      // Get settings using cSettings
      const cSettings = getSettingsInstance('config');
      const formID = cSettings.getSetting(config.formIdSetting);
      
      if (!formID) {
        throw new Error(`Form ID not found in settings: ${config.formIdSetting}`);
      }
      
      // Get the spreadsheet and data
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(config.sheetName);
      
      if (!sheet) {
        throw new Error(`Sheet not found: ${config.sheetName}`);
      }
      
      const data = sheet.getDataRange().getValues();
      
      logIt({
        level: "info",
        theMsg: `Starting URL creation for ${urlType} - ${data.length - 1} records`
      });
      
      // Show progress
      showProgress(true, `Creating ${urlType} URLs...`);
      
      // Create the URLs
      const success = processURLCreation(sheet, data, formID, config);
      
      // Hide progress
      showProgress(false);
      
      if (success) {
        logIt({
          level: "info",
          theMsg: `Successfully created ${urlType} URLs`
        });
        
        // Show completion message
        const message = `Pre-filled URLs for ${urlType} have been created!`;
        SpreadsheetApp.getActiveSpreadsheet().toast(message, "All Done", 3);
        
        return true;
      }
      
      return false;
      
    } catch (err) {
      showProgress(false);
      logIt({
        level: "severe",
        theMsg: `Error creating ${urlType} URLs`,
        error: err
      });
      throw err;
    }
  }
  
  /**
   * Process the actual URL creation
   * @param {Sheet} sheet - The target sheet
   * @param {Array} theData - The sheet data
   * @param {string} formID - The form ID
   * @param {Object} config - The configuration object
   * @returns {boolean} - Success status
   */
  function processURLCreation(sheet, theData, formID, config) {
    try {
      // Convert data to objects
      const allData = ObjApp.rangeToObjects(theData);
      
      // Get the form and form items
      const form = FormApp.openById(formID);
      const items = form.getItems();
      const preFilledURL = [];
      
      // Loop through the data and create responses
      for (let i = 0; i < allData.length; i++) {
        const resp = form.createResponse();
        
        // Process each field based on configuration
        config.fields.forEach((fieldConfig, index) => {
          let value;
          
          // Handle special fileName field
          if (fieldConfig.dataField === "fileName") {
            if (config.titleAdd) {
              value = `${allData[i].studlastfirst} ${allData[i].studentnumber} - ${config.titleAdd}`;
            } else {
              value = `${allData[i].studlastfirst} - ${config.titleAdd}`;
            }
          } else {
            value = allData[i][fieldConfig.dataField] || "";
          }
          
          // Create appropriate response based on item type
          if (fieldConfig.itemType === "PARAGRAPH_TEXT") {
            resp.withItemResponse(items[index].asParagraphTextItem().createResponse(value));
          } else {
            resp.withItemResponse(items[index].asTextItem().createResponse(value));
          }
        });
        
        preFilledURL.push([resp.toPrefilledUrl()]);
      }
      
      // Add the URLs to the spreadsheet
      const lcol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lcol).setValue('PreFilledURL');
      sheet.getRange(1, lcol + 1).setValue('Links');
      
      sheet.getRange(2, lcol, allData.length).setValues(preFilledURL);
      
      // Set the formula with dynamic column reference
      const preFilledColLetter = getColumnLetter(lcol);
      const formula = config.linkFormula.replace(/{preFilledCol}/g, preFilledColLetter);
      sheet.getRange(2, lcol + 1).setFormula(formula);
      
      return true;
      
    } catch (err) {
      logIt({
        level: "severe",
        theMsg: "Error in processURLCreation",
        error: err
      });
      throw err;
    }
  }
  
  /**
   * Convert column number to letter (1=A, 2=B, etc.)
   * @param {number} columnNumber - Column number
   * @returns {string} - Column letter
   */
  function getColumnLetter(columnNumber) {
    let result = "";
    while (columnNumber > 0) {
      columnNumber--;
      result = String.fromCharCode(65 + (columnNumber % 26)) + result;
      columnNumber = Math.floor(columnNumber / 26);
    }
    return result;
  }
  
  /**
   * Show/hide progress indicator
   * @param {boolean} show - Whether to show or hide progress
   * @param {string} message - Progress message
   */
  function showProgress(show, message = "") {
    try {
      if (show && message) {
        logIt({
          level: "info",
          theMsg: `Progress: ${message}`
        });
        // You can implement your progressInfo function here
        // progressInfo(1, message);
      } else if (!show) {
        logIt({
          level: "info",
          theMsg: "Progress completed"
        });
        // progressInfo(0);
      }
    } catch (err) {
      logIt({
        level: "warning",
        theMsg: "Error updating progress indicator",
        error: err
      });
    }
  }
  
  /**
   * Get URL creation configuration (for testing/debugging)
   * @param {string} urlType - The URL type
   * @returns {Object} - Configuration object
   */
  function getURLConfig(urlType) {
    return URL_CONFIG[urlType] || null;
  }
  
  // Return public interface
  return {
    createPrefilledURLs: createPrefilledURLs,
    getURLConfig: getURLConfig
  };
  
})();

/**
 * Namespace for URL creation helper functions
 * These functions handle the interaction between bound script and library functions
 */
var nsPropsURLCreator = (function() {
  
  /**
   * Helper function to create initial goal URLs
   * Called from bound script via callLibraryNS
   */
  function helperCreateInitialURLs() {
    try {
      return nsURLCreator.createPrefilledURLs('initial');
    } catch (error) {
      logIt({
        level: "severe",
        theMsg: "Error in helper create initial URLs function",
        error: error
      });
      throw new Error("Error creating initial URLs: " + error.message);
    }
  }
  
  /**
   * Helper function to create mid year URLs
   * Called from bound script via callLibraryNS
   */
  function helperCreateMidURLs() {
    try {
      return nsURLCreator.createPrefilledURLs('mid');
    } catch (error) {
      logIt({
        level: "severe",
        theMsg: "Error in helper create mid URLs function",
        error: error
      });
      throw new Error("Error creating mid year URLs: " + error.message);
    }
  }
  
  /**
   * Helper function to create final URLs
   * Called from bound script via callLibraryNS
   */
  function helperCreateFinalURLs() {
    try {
      return nsURLCreator.createPrefilledURLs('final');
    } catch (error) {
      logIt({
        level: "severe",
        theMsg: "Error in helper create final URLs function",
        error: error
      });
      throw new Error("Error creating final URLs: " + error.message);
    }
  }
  
  /**
   * Generic helper to create any type of URLs
   * Called from bound script via callLibraryNS
   */
  function helperCreateURLs(urlType) {
    try {
      return nsURLCreator.createPrefilledURLs(urlType);
    } catch (error) {
      logIt({
        level: "severe",
        theMsg: `Error in helper create URLs function for type: ${urlType}`,
        error: error
      });
      throw new Error(`Error creating ${urlType} URLs: ` + error.message);
    }
  }
  
  // Return public interface
  return {
    helperCreateInitialURLs: helperCreateInitialURLs,
    helperCreateMidURLs: helperCreateMidURLs,
    helperCreateFinalURLs: helperCreateFinalURLs,
    helperCreateURLs: helperCreateURLs
  };
  
})();