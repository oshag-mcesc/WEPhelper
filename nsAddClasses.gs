/**
 * Top-level function to trigger class-adding process.
 * Can be called from menu (WEPhelper.addClassesToTable).
 * @param {Array} [classData] - Optional array of data, else pulls from sheet.
 */
function addClassesToTable(classData) {
  const settings = getSettingsInstance("config");
  try {
    let theData;

    if (classData) {
      theData = classData;
    } else {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("forClasses");
      theData = sheet.getRange(2, 9, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    }
    const rowNum = parseInt(settings.getSetting("classRowNum") || "0");
    const results = nsAddClasses.addClasses(theData, rowNum);
    settings.setSetting("classRowNum", results.classRowNum);

    const ui = SpreadsheetApp.getActiveSpreadsheet();
    if (parseInt(results.classRowNum) === 0) {
      ui.toast("Got done in time!", "All done!", -1);
    } else {
      ui.toast("Exceeded time limit! Please run addClasses again.", "NOT done yet!", -1);
    }

  } catch (err) {
    logIt({
      level: "severe",
      theMsg: "Error in addClassesToTable",
      error: err.toString()
    });
  }
}


/** 
 * Namespace for functions related to adding classes to WEP documents 
 */
const nsAddClasses = (() => {

  /**
   * Add student classes to the table on their WEP document.
   * Uses timing logic to avoid timeouts.
   * 
   * @param {Array} theData - Array of student/class data.
   * @param {number} rowNum - Starting index to process.
   * @return {Object} - { classRowNum: string }
   */
  function addClasses(theData, rowNum) {
    const start = new Date();
    const headerStyle = { [DocumentApp.Attribute.BOLD]: true };
    const rowStyle = { [DocumentApp.Attribute.BOLD]: false };
    const length = theData.length;

    let results = {};

    try {
      for (let i = rowNum; i < length;) {
        const tbl = getTheTable_(theData[i][0]);  // doc ID in first column

        do {
          const tr = tbl.appendTableRow().setAttributes(rowStyle);
          for (let j = 1; j < 4; j++) {
            tr.appendTableCell(theData[i][j]);
          }
          i++;

          if (i === length) break;

        } while (theData[i - 1][0] === theData[i][0]); // same doc ID

        tbl.getRow(0).setAttributes(headerStyle);

        if (isTimeUp_(start)) {
          results = { classRowNum: i.toString() }; // Save progress
          logIt({
            level: "warning",
            theMsg: "Execution time limit reached in addClasses()",
            error: `Stopped at row ${i}`
          });
          break;
        }
      }

      // If all done
      if (!results.classRowNum) {
        results = { classRowNum: "0" };
      }

      return results;

    } catch (err) {
      logIt({
        level: "severe",
        theMsg: "Error in addClasses()",
        error: err.toString()
      });
      return { classRowNum: rowNum.toString() }; // return where it left off
    }
  }

  /**
   * Finds the target table in the WEP document.
   * @param {string} docId - The ID of the Google Doc.
   * @returns {GoogleAppsScript.Document.Table} - The matching table.
   */
  function getTheTable_(docId) {
    try {
      const doc = DocumentApp.openById(docId);
      const tables = doc.getBody().getTables();

      for (let i = 0; i < tables.length; i++) {
        const cell = tables[i].getCell(0, 0).getText();
        if (cell.trim() === "Service Area") {
          return tables[i];
        }
      }

      throw new Error(`No matching table found in doc: ${docId}`);
    } catch (err) {
      logIt({
        level: "severe",
        theMsg: `Error in getTheTable_() for docId: ${docId}`,
        error: err.toString()
      });
      throw err;
    }
  }

  return {
    addClasses
  };

})();
