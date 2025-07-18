# Script Overview: WEPhelper

This overview summarizes the main purpose and functionality of each script file in the WEPhelper repository.

---

## Code.gs

**Purpose:**  
Core logic for WEP document creation and utility functions.

**Key Features:**
- `createDocs`: Creates personalized WEP documents from Google Sheet data and a template Google Doc. Populates student info into the template and saves copies in Drive.
- `cleanGoal`: Cleans goal strings for use in URLs or forms (removes/encodes special characters).
- Additional helper functions for data processing and execution time checks.

---

## pushUpdates.gs

**Purpose:**  
Handles mid-year and final updates to existing WEP documents.

**Key Features:**
- `pushMidUpdate1` / `pushFinalUpdate1`: Entry points for pushing updates, configured via spreadsheet tabs.
- `configPushUpdates`: Middle-man function that orchestrates the update process based on configuration.
- `pushTheUpdates`: Main update engine that applies changes to student WEP documents and logs errors.

---

## nsUtilities.gs

**Purpose:**  
Provides utility functions for data manipulation and document indexing.

**Key Features:**
- `keepCols_`: Extracts specified columns from a 2D array.
- `getTableIndexes_`: Indexes tables within a document for later reference.

---

## tabMain.gs

**Purpose:**  
Implements Google Sheets sidebar UI for toggling tab visibility.

**Key Features:**
- `pickTabs`: Displays a sidebar to show/hide tabs.
- `createCheckBoxes`: Generates checkboxes for each tab.
- `toggleEm` / `toggle_`: Functions to hide/show individual tabs.
- `include`: Loads CSS/JS into the sidebar UI.

---

## objService.gs

**Purpose:**  
Data conversion utilities between arrays and objects.

**Key Features:**
- `objectToArray`: Converts object arrays to 2D arrays.
- `rangeToObjects`: Converts sheet ranges to arrays of objects.
- `splitRangesToObjects`, `camelArray`, `camelString`: Helpers for object/array transformation and string formatting.

---

## addMidAndFinalDocs.gs

**Purpose:**  
Manages folder creation and document copying for mid-year and final WEPs.

**Key Features:**
- `CreateFolder_`: Creates new folders in Google Drive.
- `copyDocs_`: Copies and renames documents for mid-year/final milestones.

---

## bvckCreateDocsEXP.gs

**Purpose:**  
Experimental backup logic for WEP document creation.

**Key Features:**
- `tester2`: Demonstrates alternative method for document creation and template filling.
- Contains experimental code for processing student info and generating docs.

---

## testingScripts.gs

**Purpose:**  
Contains test and practice scripts for development and debugging.

**Key Features:**
- `testGoalsubmit`: Tests form submission workflow.
- `logPractice`: Demonstrates logging and error handling.

---

## LoggingStuff.gs

**Purpose:**  
Custom logging integration for tracking errors and info messages.

**Key Features:**
- `logIt`: Logs messages to an external log sheet using severity levels.
- `logError`: Exposes logging for use in other scripts.

---

## README.md

**Purpose:**  
High-level documentation for project purpose, setup, features, and usage.

---

## Other Files

Any HTML, CSS, or JS files referenced (such as for sidebar UI) are included for user interface enhancements in Google Sheets.

---
