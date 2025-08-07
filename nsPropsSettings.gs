/**
 * Settings management for refactor1 library
 * Uses the Settings class to store configuration in a sheet instead of PropertiesService
 */

// Default configuration values
const DEFAULT_SETTINGS = {
  WEPtemplateID: "",
  MainWEPfolderID: "",
  InitialGoalsFormID: "",
  MidYearProgressFormID: "",
  FinalProgressFormID: "",
  subject: "The subject of the email",
  greeting: "The greeting that is before the list of links.",
  closing: "The closing after the list of links.",
  linktime: "initial"
};

/**
 * Main function to show the settings dialog
 * This should be called from the bound script
 */
function showSettingsDialog() {
  try {
    // Initialize settings if needed
    const settingsInstance = getSettingsInstance('config');
    if (!settingsInstance.settingsSheet) {
      settingsInstance.init();
      initializeDefaultSettings(settingsInstance);
    }
    
    // Load current settings
    const currentSettings = loadAllSettings(settingsInstance);
    
    // Create and show the dialog
    const html = HtmlService.createTemplateFromFile('settingsDialog');
    
    // Set all the template variables
    Object.keys(DEFAULT_SETTINGS).forEach(key => {
      html[key] = currentSettings[key] || DEFAULT_SETTINGS[key];
    });
    
    const htmlOutput = html.evaluate()
      .setWidth(700)
      .setHeight(800);
    
    SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, 'Application Settings');
      
  } catch (err) {
    logIt({
      level: "severe",
      theMsg: "Error showing settings dialog",
      error: err
    });
    throw new Error("Unable to show settings dialog: " + err.message);
  }
}

/**
 * Initialize default settings in the config sheet
 */
function initializeDefaultSettings(settingsInstance) {
  try {
    Object.keys(DEFAULT_SETTINGS).forEach(key => {
      // Only set if the setting doesn't already exist
      if (!settingsInstance.getSetting(key)) {
        settingsInstance.setSetting(key, DEFAULT_SETTINGS[key]);
      }
    });
    
    logIt({
      level: "info",
      theMsg: "Default settings initialized successfully"
    });
    
  } catch (err) {
    logIt({
      level: "severe",
      theMsg: "Error initializing default settings",
      error: err
    });
    throw err;
  }
}

/**
 * Load all settings from the config sheet
 */
function loadAllSettings(settingsInstance) {
  const settings = {};
  
  try {
    Object.keys(DEFAULT_SETTINGS).forEach(key => {
      const value = settingsInstance.getSetting(key);
      settings[key] = value !== null ? value : DEFAULT_SETTINGS[key];
    });
    
    return settings;
    
  } catch (err) {
    logIt({
      level: "severe",
      theMsg: "Error loading settings from config sheet",
      error: err
    });
    throw err;
  }
}

/**
 * Save ID settings (first tab of dialog)
 * Called from the HTML dialog
 */
function saveIdSettings(settingsData) {
  try {
    const settingsInstance = getSettingsInstance('config');
    const idFields = ['WEPtemplateID', 'MainWEPfolderID', 'InitialGoalsFormID', 
                      'MidYearProgressFormID', 'FinalProgressFormID'];
    
    idFields.forEach(field => {
      if (settingsData[field] !== undefined) {
        settingsInstance.setSetting(field, settingsData[field]);
      }
    });
    
    logIt({
      level: "info",
      theMsg: "ID settings saved successfully",
      error: null
    });
    
    return { success: true, message: "ID settings saved successfully!" };
    
  } catch (err) {
    logIt({
      level: "severe",
      theMsg: "Error saving ID settings",
      error: err
    });
    
    return { success: false, message: "Error saving ID settings: " + err.message };
  }
}

/**
 * Save email settings (second tab of dialog)
 * Called from the HTML dialog
 */
function saveEmailSettings(settingsData) {
  try {
    const settingsInstance = getSettingsInstance('config');
    const emailFields = ['subject', 'greeting', 'closing', 'linktime'];
    
    emailFields.forEach(field => {
      if (settingsData[field] !== undefined) {
        settingsInstance.setSetting(field, settingsData[field]);
      }
    });
    
    logIt({
      level: "info",
      theMsg: "Email settings saved successfully",
      error: null
    });
    
    return { success: true, message: "Email settings saved successfully!" };
    
  } catch (err) {
    logIt({
      level: "severe",
      theMsg: "Error saving email settings",
      error: err
    });
    
    return { success: false, message: "Error saving email settings: " + err.message };
  }
}

/**
 * Get a specific setting value
 * Utility function for other library functions
 */
function getSettingValue(settingName) {
  try {
    const settingsInstance = getSettingsInstance('config');
    const value = settingsInstance.getSetting(settingName);
    return value !== null ? value : DEFAULT_SETTINGS[settingName];
    
  } catch (err) {
    logIt({
      level: "warning",
      theMsg: `Error getting setting ${settingName}`,
      error: err
    });
    return DEFAULT_SETTINGS[settingName];
  }
}

/**
 * Get all settings as an object
 * Utility function for other library functions
 */
function getAllSettings() {
  try {
    const settingsInstance = getSettingsInstance('config');
    return loadAllSettings(settingsInstance);
    
  } catch (err) {
    logIt({
      level: "warning",
      theMsg: "Error getting all settings, returning defaults",
      error: err
    });
    return DEFAULT_SETTINGS;
  }
}

/**
 * Helper function to include HTML files from the library
 * This needs to be called from the library context
 */
function includeHtmlFile(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Namespace for settings-related helper functions
 * These functions handle the interaction between HTML and library functions
 */
var nsPropsSettings = (function() {
  
  /**
   * Helper function to save ID settings
   * Called from HTML dialog via bound script callLibraryNS
   */
  function helperSaveIdSettings(settingsData) {
    try {
      return saveIdSettings(settingsData);
    } catch (error) {
      return { 
        success: false, 
        message: "Error in helper function: " + error.message 
      };
    }
  }

  /**
   * Helper function to save email settings
   * Called from HTML dialog via bound script callLibraryNS
   */
  function helperSaveEmailSettings(settingsData) {
    try {
      return saveEmailSettings(settingsData);
    } catch (error) {
      return { 
        success: false, 
        message: "Error in helper function: " + error.message 
      };
    }
  }

  /**
   * Helper function to show settings dialog
   * Called from menu via bound script callLibraryNS
   */
  function helperShowSettingsDialog() {
    try {
      return showSettingsDialog();
    } catch (error) {
      throw new Error("Error showing settings dialog: " + error.message);
    }
  }

  /**
   * Helper function to get a specific setting value
   * Called from other functions via bound script callLibraryNS
   */
  function helperGetSettingValue(settingName) {
    try {
      return getSettingValue(settingName);
    } catch (error) {
      Logger.log('Error getting setting ' + settingName + ': ' + error.message);
      return null;
    }
  }

  /**
   * Helper function to get all settings
   * Called from other functions via bound script callLibraryNS
   */
  function helperGetAllSettings() {
    try {
      return getAllSettings();
    } catch (error) {
      Logger.log('Error getting all settings: ' + error.message);
      return {};
    }
  }

  // Return public interface
  return {
    helperSaveIdSettings: helperSaveIdSettings,
    helperSaveEmailSettings: helperSaveEmailSettings,
    helperShowSettingsDialog: helperShowSettingsDialog,
    helperGetSettingValue: helperGetSettingValue,
    helperGetAllSettings: helperGetAllSettings
  };
  
})();