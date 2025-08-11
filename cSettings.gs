/**
 * Constants for boolean settings values.
 */
const YES = 'Yes';
const NO = 'No';

/**
 * The Settings class provides a way to manage script parameters/settings
 * directly within a Google Sheet, making it accessible for non-technical users.
 */
class Settings {

    /**
     * Constructor initializes the settings sheet and map.
     * @param {string} [sheetName="Settings"] - Name of the sheet where settings are stored.
     */
    constructor(sheetName = "config") {
        this.sheetName = sheetName;
        this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

        // Wrap sheet retrieval in a try-catch block to handle potential errors
        try {
            this.settingsSheet = this.spreadsheet.getSheetByName(sheetName);
        } catch (e) {
            Logger.log(`Error retrieving settings sheet: ${e.message}`);
            this.settingsSheet = null;
        }

        this.settingsMap = this.initSettingsMap();
    }

    /**
     * Initializes the settings sheet if it doesn't exist.
     */
    init() {
        if (!this.settingsSheet) {
            // Wrap sheet creation in a try-catch block to handle potential errors
            try {
                this.settingsSheet = this.spreadsheet.insertSheet(this.sheetName);
                this.settingsSheet.appendRow(['Setting', 'Value']);
                this.settingsSheet.getRange('1:1').setFontWeight('bold');
            } catch (e) {
                Logger.log(`Error initializing settings sheet: ${e.message}`);
                this.settingsSheet = null;
            }
        }
    }

    /**
     * Initializes the settings map from the sheet data.
     * @returns {Map} - A map of settings.
     */
    initSettingsMap() {
        if (!this.settingsSheet) return new Map();

        // Wrap data retrieval in a try-catch block to handle potential errors
        try {
            const data = this.settingsSheet.getDataRange().getValues();
            const map = new Map();

            for (const [key, value] of data) {
                map.set(key, value);
            }

            return map;
        } catch (e) {
            Logger.log(`Error retrieving settings data: ${e.message}`);
            return new Map();
        }
    }

    /**
     * Sets or updates a setting in the sheet.
     * @param {string} settingName - Name of the setting.
     * @param {string} settingValue - Value of the setting.
     */
    setSetting(settingName, settingValue) {
        const rowIndex = [...this.settingsMap.keys()].indexOf(settingName) + 1;

        if (rowIndex > 0) {
            // Wrap range retrieval and value setting in a try-catch block to handle potential errors
            try {
                this.settingsSheet.getRange(rowIndex, 2).setValue(settingValue);
            } catch (e) {
                Logger.log(`Error setting value for ${settingName}: ${e.message}`);
            }
        } else {
            // Wrap row appending in a try-catch block to handle potential errors
            try {
                this.settingsSheet.appendRow([settingName, settingValue]);
            } catch (e) {
                Logger.log(`Error appending row for ${settingName}: ${e.message}`);
            }
        }

        this.settingsMap.set(settingName, settingValue);
    }

    /**
     * Retrieves a setting's value from the map.
     * @param {string} settingName - Name of the setting.
     * @returns {string|null} - Value of the setting or null if not found.
     */
    getSetting(settingName) {
        return this.settingsMap.get(settingName) || null;
    }

    /**
     * Retrieves a boolean setting's value.
     * @param {string} settingName - Name of the setting.
     * @returns {boolean|null} - True if 'Yes', False if 'No', or null if neither.
     */
    getBooleanSetting(settingName) {
        const settingValue = this.getSetting(settingName);

        if (settingValue === YES) return true;
        if (settingValue === NO) return false;

        Logger.log(`Setting value is not ${YES} or ${NO}: ${settingName}`);
        return null;
    }

    /**
     * Sets a setting in the script properties.
     * @param {string} settingName - Name of the setting.
     * @param {string} settingValue - Value of the setting.
     */
    setSettingInScriptProperties(settingName, settingValue) {
        PropertiesService.getScriptProperties().setProperty(settingName, settingValue);
    }

    /**
     * Retrieves a setting from the script properties.
     * @param {string} settingName - Name of the setting.
     * @returns {string} - Value of the setting.
     */
    getSettingFromScriptProperties(settingName) {
        return PropertiesService.getScriptProperties().getProperty(settingName);
    }
}

/**To make this avaialble in other scripts
 * You need to have this "factory function"
 */
// Expose the class by wrapping it in a function
function getSettingsInstance(sheetName) {
  return new Settings(sheetName);
}
