/**
 * @file This script provides utility functions for Google Sheets,
 * specifically for highlighting duplicate rows based on specified criteria.
 * @namespace SheetUtils
 */
var SheetUtils = SheetUtils || {}; // Ensure SheetUtils namespace exists

/**
 * @typedef {Object} ColumnHeader
 * @property {string} name - The name of the column header (for readability/reference).
 * @property {number} colIndex - The 0-indexed column position in the sheet.
 */

/**
 * @constant {Object.<string, ColumnHeader>} HEADERS_FOR_DUPLICATES
 * @description Defines the column headers and their 0-indexed positions
 * used for identifying duplicate rows in evaluation data.
 * These positions are based on the expected header row:
 * "Timestamp", "Student Name", "Student ID", "Course Name",
 * "Teacher Name", "Goal", "Mid Year Progress", "End of Year Evaluation".
 */
SheetUtils.HEADERS_FOR_DUPLICATES = {
  STUDENT_ID: { name: "Student ID", colIndex: 2 }, // Column C (0-indexed: 2)
  COURSE_NAME: { name: "Course Name", colIndex: 3 }, // Column D (0-indexed: 3)
  TEACHER_NAME: { name: "Teacher Name", colIndex: 4 } // Column E (0-indexed: 4)
};

/**
 * Highlights duplicate rows in the active sheet based on a combination of
 * "Student ID", "Course Name", and "Teacher Name".
 * Duplicate rows (including the first occurrence of a duplicate key) will be
 * highlighted in a light orange color (`#FFD966`).
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

  try {
    const range = sheet.getDataRange(); // Get the entire data range of the sheet
    const values = range.getValues(); // Get all the values in the sheet

    // Clear existing background colors to ensure a clean highlight.
    // This is important to remove highlights from previous runs or manual highlighting.
    range.setBackground(null);

    // If there's only a header row or no data, log a message and exit.
    if (values.length <= 1) {
      Logger.log("SheetUtils.highlightDuplicateRows: No data or only header row found. Exiting.");
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

      // Extract the values for the unique key combination using the predefined column indices.
      const studentId = row[SheetUtils.HEADERS_FOR_DUPLICATES.STUDENT_ID.colIndex];
      const courseName = row[SheetUtils.HEADERS_FOR_DUPLICATES.COURSE_NAME.colIndex];
      const teacherName = row[SheetUtils.HEADERS_FOR_DUPLICATES.TEACHER_NAME.colIndex];

      // Create a unique key by concatenating the relevant fields.
      // Using a delimiter like '-' ensures that combinations like "123-ABC" and "12-3ABC" are distinct.
      const uniqueKey = `${studentId}-${courseName}-${teacherName}`;

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