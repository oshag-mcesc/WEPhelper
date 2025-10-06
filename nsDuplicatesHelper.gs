/**
 * @file This script provides utility functions for Google Sheets,
 * specifically for highlighting duplicate rows based on specified criteria.
 * @namespace SheetUtils
 */
var SheetUtils = SheetUtils || {}; // Ensure SheetUtils namespace exists

/**
 * @constant {Array<string>} REQUIRED_HEADERS
 * @description The column headers required for identifying duplicate rows.
 * These headers must exist in the sheet's first row.
 */
SheetUtils.REQUIRED_HEADERS = ["Student ID", "Course Name", "Teacher Name"];

/**
 * Finds the column indices for the required headers in the sheet.
 * 
 * @memberof SheetUtils
 * @param {Array<Array>} headerRow - The first row of the sheet containing column headers.
 * @returns {Object|null} An object mapping header names to their 0-indexed column positions,
 *                        or null if any required headers are missing.
 * @private
 * @example
 * // Returns: { "Student ID": 2, "Course Name": 3, "Teacher Name": 4 }
 */
SheetUtils.findHeaderIndices = (headerRow) => {
  const indices = {};
  const normalizedHeaders = headerRow.map(h => String(h).trim().toLowerCase());
  
  for (const requiredHeader of SheetUtils.REQUIRED_HEADERS) {
    const normalizedRequired = requiredHeader.trim().toLowerCase();
    const index = normalizedHeaders.indexOf(normalizedRequired);
    
    if (index === -1) {
      Logger.log(`SheetUtils.findHeaderIndices: Required header "${requiredHeader}" not found in sheet.`);
      return null;
    }
    
    indices[requiredHeader] = index;
  }
  
  return indices;
};

/**
 * Highlights duplicate rows in the active sheet based on a combination of
 * "Student ID", "Course Name", and "Teacher Name".
 * Duplicate rows (including the first occurrence of a duplicate key) will be
 * highlighted in a light orange color (`#FFD966`).
 *
 * The function dynamically locates the required columns by searching for their
 * header names in the first row. This allows the columns to be in any order
 * and in any position within the sheet.
 *
 * This function clears any existing background colors in the data range
 * before applying new highlights.
 *
 * @memberof SheetUtils
 * @returns {void}
 * @example
 * // To call this function from a sheet script after adding the library:
 * SheetUtils.highlightDuplicateRows();
 */
SheetUtils.highlightDuplicateRows = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet
  const sheet = ss.getActiveSheet(); // Get the active sheet
  const sheetName = sheet.getName(); // Get the name of the active sheet
  
  try {
    const range = sheet.getDataRange(); // Get the entire data range of the sheet
    const values = range.getValues(); // Get all the values in the sheet
    
    // If there's no data at all, log a message and exit.
    if (values.length === 0) {
      Logger.log("SheetUtils.highlightDuplicateRows: No data found in sheet. Exiting.");
      return;
    }
    
    // Find the column indices based on header names
    const headerRow = values[0];
    const headerIndices = SheetUtils.findHeaderIndices(headerRow);
    
    if (!headerIndices) {
      const missingHeaders = SheetUtils.REQUIRED_HEADERS.join('", "');
      const errorMsg = `Required column headers not found in sheet "${sheetName}". Please ensure the sheet has columns named: "${missingHeaders}".`;
      Logger.log(`SheetUtils.highlightDuplicateRows: ${errorMsg}`);
      SpreadsheetApp.getUi().alert(errorMsg);
      return;
    }
    
    Logger.log(`SheetUtils.highlightDuplicateRows: Processing sheet "${sheetName}". Found columns - Student ID: ${headerIndices["Student ID"]}, Course Name: ${headerIndices["Course Name"]}, Teacher Name: ${headerIndices["Teacher Name"]}`);
    
    // Clear existing background colors to ensure a clean highlight.
    // This is important to remove highlights from previous runs or manual highlighting.
    range.setBackground(null);
    
    // If there's only a header row, log a message and exit.
    if (values.length <= 1) {
      Logger.log("SheetUtils.highlightDuplicateRows: Only header row found, no data to process. Exiting.");
      return;
    }
    
    // A Map to store unique keys and the row number (1-indexed) of their first occurrence.
    const seenKeys = new Map();
    // A Set to store row numbers (1-indexed) of all rows identified as duplicates.
    const duplicateRows = new Set();
    
    // Iterate over each row, starting from the second row (index 1) to skip the header.
    values.forEach((row, index) => {
      // The `index` is 0-based, so for actual row numbers, add 1.
      const currentRowNumber = index + 1;
      if (index === 0) return; // Skip the header row (index 0)
      
      // Extract the values for the unique key combination using the dynamically found column indices.
      const studentId = row[headerIndices["Student ID"]];
      const courseName = row[headerIndices["Course Name"]];
      const teacherName = row[headerIndices["Teacher Name"]];
      
      // Create a unique key by concatenating the relevant fields.
      // Using a delimiter like '|~|' ensures that combinations are distinct even with special characters.
      const uniqueKey = `${studentId}|~|${courseName}|~|${teacherName}`;
      
      // Check if the unique key has been seen before.
      if (seenKeys.has(uniqueKey)) {
        // If the key exists, it means we've found a duplicate.
        // Add both the current row and the row where this key was first encountered to the duplicate set.
        duplicateRows.add(currentRowNumber); // Add current row number (1-indexed)
        duplicateRows.add(seenKeys.get(uniqueKey)); // Add previously seen row number (1-indexed)
      } else {
        // If the key is new, add it to the map with its corresponding row number.
        seenKeys.set(uniqueKey, currentRowNumber);
      }
    });
    
    // Apply highlighting to all identified duplicate rows.
    if (duplicateRows.size > 0) {
      const highlightColor = "#FFD966"; // Light orange color for duplicates
      duplicateRows.forEach(rowNum => {
        // Get the range for the entire duplicate row.
        // Parameters: startRow, startColumn, numRows, numColumns
        sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).setBackground(highlightColor);
      });
      Logger.log(`SheetUtils.highlightDuplicateRows: Highlighted ${duplicateRows.size} duplicate rows.`);
    } else {
      Logger.log("SheetUtils.highlightDuplicateRows: No duplicate rows found.");
    }
  } catch (error) {
    // Catch and log any errors that occur during script execution.
    // Use console.error for more detailed logging in the Apps Script logs.
    console.error("SheetUtils.highlightDuplicateRows: Failed to highlight duplicate rows.", error);
    // Alert the user that an error occurred for a better user experience.
    SpreadsheetApp.getUi().alert("An error occurred. Please check the script logs (Extensions > Apps Script > Editor > Execution log) for details.");
  }
};