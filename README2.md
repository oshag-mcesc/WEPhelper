# WEPhelper

A Google Apps Script library designed to streamline the management, creation, and updating of WEPs (Written Education Plans) for gifted education programs. This library automates document generation, updates, and workflow tasks within Google Sheets and Google Drive, eliminating manual effort and reducing errors.

## Features

### 🚀 Automated WEP Document Creation
- Generates personalized WEP documents for students based on data in Google Sheets and a template document
- Student information (name, grade, gifted area, etc.) is dynamically inserted into the template
- Bulk document creation for entire student rosters

### 📝 Push Updates to WEPs
- Mid-year and final updates can be pushed directly to WEP documents from configured tabs in your spreadsheet
- Eliminates the need for separate update documents
- Maintains consistency across all student documents

### 🛠️ Utility Functions
- Manipulate tabular data with ease
- Clean strings for use in URLs or forms
- Extract information from documents or arrays
- Data conversion and formatting utilities

### 🎛️ Sidebar UI for Tab Management
- Provides a sidebar interface to show/hide tabs in Google Sheets
- Improves navigation and organization
- Streamlines workflow management

### 📊 Logging and Error Handling
- Integrates with custom logs to track errors and info messages
- Helps diagnose issues in automation processes
- Comprehensive error reporting for troubleshooting

## Prerequisites

- Google Workspace account with access to Google Sheets and Google Drive
- A properly formatted Google Sheet with student data
- A template Google Doc for WEPs
- Basic familiarity with Google Apps Script

## File Structure

- **`Code.gs`** - Main logic for document creation and data processing
- **`pushUpdates.gs`** - Handles pushing updates to WEPs
- **`nsUtilities.gs`** - Utility functions for array and document manipulation
- **`tabMain.gs`** - Sidebar UI and tab toggling logic
- **`objService.gs`** - Object/array conversion helpers
- **`appsscript.json`** - Google Apps Script manifest file containing project configuration and metadata
- **`README.md`** - Project documentation and usage guide
- **`LICENSE`** - MIT License file specifying terms of use and distribution

## Configuration

The library requires configuration of:
- Student data spreadsheet structure
- Template document placeholders
- Folder and file IDs
- Tab names specific to your district's workflow

## Target Users

This library is designed for:
- School administrators managing gifted education programs
- Educators working with Written Education Plans
- District personnel responsible for compliance documentation
- Anyone needing to automate Google Workspace document workflows

## Important Notes

- Certain configuration (IDs, tab names) must be tailored to your district's specific workflow
- Error handling and logging are integrated to help ensure smooth operation
- The library assumes familiarity with Google Apps Script and Google Workspace administration
- Always test with a small subset of data before running bulk operations

## Contributing

This is a specialized tool for educational administration. If you have suggestions or improvements:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request with detailed description of changes

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For questions about implementation or troubleshooting, please open an issue in this repository. Please note that this tool is designed for specific educational workflows and may require customization for your particular use case.
