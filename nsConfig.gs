/**
 * Global configuration object for the library
 * Central place to manage all config keys and provide easy renaming/updates
 */
const CONFIG = {
  // Sheet settings keys
  KEYS: {
    WEP_TEMPLATE_ID: 'WEPtemplateID',
    MAIN_WEP_FOLDER_ID: 'MainWEPfolderID', 
    INITIAL_GOALS_FORM_ID: 'InitialGoalsFormID',
    MID_YEAR_PROGRESS_FORM_ID: 'MidYearProgressFormID',
    FINAL_PROGRESS_FORM_ID: 'FinalProgressFormID',
    ROW_NUM: 'rowNum'  
  },
  
  // Sheet name
  SHEET_NAME: 'config'
};

/**
 * Helper function to get config values using the centralized keys
 * @param {string} configKey - Key from CONFIG.KEYS
 * @param {Object} settings - Optional existing settings instance
 * @returns {string} The config value
 */
function getConfigValue(configKey, settings = null) {
  const settingsInstance = settings || getSettingsInstance(CONFIG.SHEET_NAME);
  return settingsInstance.getSetting(configKey);
}

//TODO - Find out if this code is still being used.

const nsConfig = (()=>{
  return{
    tabs:{
      forMidYearPush:{
        name:"forMidYearPush",
        rngCell:"I1",
        rowTitle:"Mid Year Progress",
        statusCol:14 //column N
      },
      forFinalPush:{
        name:"forFinalPush",
        rngCell:"J1",
        rowTitle:"Final Evaluation",
        statusCol:15  // column O
      }

    }
  }
})()