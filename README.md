# WEPhelper

WEPhelper is a Google Apps Script library designed to streamline the management, creation, and updating of WEPs (Written Education Plans) for gifted education programs. The library automates document generation, updates, and workflow tasks within Google Sheets and Google Drive, eliminating manual effort and reducing errors.

## Features

- **Automated WEP Document Creation:**  
  Generates personalized WEP documents for students based on data in Google Sheets and a template document. Student information (name, grade, gifted area, etc.) is dynamically inserted into the template.

- **Push Updates to WEPs:**  
  Mid-year and final updates can be pushed directly to WEP documents from configured tabs in your spreadsheet, avoiding the need for separate update documents.

- **Utility Functions:**  
  Includes utilities for manipulating tabular data, cleaning strings for use in URLs or forms, and extracting information from documents or arrays.

- **Sidebar UI for Tab Management:**  
  Provides a sidebar interface to show/hide tabs in Google Sheets, making navigation and organization easier.

- **Logging and Error Handling:**  
  Integrates with custom logs to track errors and info messages, helping diagnose issues in automation processes.

## How It Works

1. **Setup:**
   - Prepare a Google Sheet with student data and a template Google Doc for WEPs.
   - Configure folder IDs and template IDs in your script or sheet.

2. **Document Generation:**
   - Use the `createDocs` function to loop through each student row in the sheet, copy the template, and replace placeholders (like `<<StudLastFirst>>`, `<<Grade>>`, etc.) with actual data.
   - The generated docs are saved to a designated folder and their IDs are recorded in the sheet.

3. **Updating WEPs:**
   - Use the `pushMidUpdate` and `pushFinalUpdate` functions to send updates from the sheet to the main WEP documents.
   - The update functions extract relevant data and push changes to each student's document.

4. **Utilities:**
   - Functions like `cleanGoal`, `keepCols_`, and data conversion utilities help format and process information for document creation and updates.

5. **UI & Logging:**
   - The sidebar UI facilitates tab visibility toggling.
   - Logging functions help track the success or failure of operations for troubleshooting.

## Example Usage

```javascript
// Generate WEP documents for all students in the sheet
createDocs(sheet, studentData, templateFileId, destinationFolderId, startRow);

// Push mid-year updates to WEPs
pushMidUpdate1();

// Clean a range of goals for safe URL usage
cleanGoal();
```

## Scripts Overview

- `Code.gs` – Main logic for document creation and data processing.
- `pushUpdates.gs` – Handles pushing updates to WEPs.
- `nsUtilities.gs` – Utility functions for array and document manipulation.
- `tabMain.gs` – Sidebar UI and tab toggling logic.
- `objService.gs` – Object/array conversion helpers.

## Requirements

- Google Workspace account with access to Google Sheets and Google Drive.
- A properly formatted Google Sheet and template Google Doc.
- Script installed in your Google Sheet (via Extensions > Apps Script).

## Notes

- This library is designed for use by school administrators or educators managing gifted education plans.
- Certain configuration (IDs, tab names) must be tailored to your district's workflow.
- Error handling and logging are integrated to help ensure smooth operation.

## License

See LICENSE file for details.
