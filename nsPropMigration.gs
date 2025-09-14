/**
 * @fileoverview One-time migration script to move properties from Properties Service
 * to the cSettings config sheet. This script should only be run once during the
 * transition from Properties Service to config sheet-based settings.
 */

/**
 * Properties to migrate from Properties Service to config sheet
 * @constant {Array<string>}
 */
const PROPERTIES_TO_MIGRATE = [
  'WEPtemplateID',
  'MainWEPfolderID', 
  'InitialGoalsFormID',
  'MidYearProgressFormID',
  'FinalProgressFormID',
  'subject',
  'greeting',
  'closing',
  'linktime'
];

/**
 * Name of the config sheet (adjust if different)
 * @constant {string}
 */
const CONFIG_SHEET_NAME = 'config'; // Change this if your config sheet has a different name

/**
 * Main function to migrate properties from Properties Service to config sheet
 * This is a one-time migration script that will:
 * 1. Read all specified properties from Properties Service
 * 2. Find matching properties in the config sheet
 * 3. Update config sheet values (overwriting existing values)
 * 4. Log the migration results
 * 
 * @returns {void}
 */
function migratePropertiesToConfig() {
  try {
    console.log("Starting Properties Service to Config Sheet migration...");
    
    // Get the config sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      const errorMsg = `ERROR: Config sheet '${CONFIG_SHEET_NAME}' not found. Please check the sheet name.`;
      console.error(errorMsg);
      SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Migration Error", -1);
      logIt({ level: "severe", theMsg: errorMsg });
      return;
    }
    
    // Get Properties Service data
    const scriptProperties = PropertiesService.getScriptProperties();
    const allProperties = scriptProperties.getProperties();
    
    console.log(`Found ${Object.keys(allProperties).length} total properties in Properties Service`);
    
    // Get config sheet data (assumes column A = property names, column B = values)
    const configData = configSheet.getDataRange().getValues();
    
    if (configData.length < 2) {
      const errorMsg = "ERROR: Config sheet appears to be empty or only has headers.";
      console.error(errorMsg);
      SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Migration Error", -1);
      logIt({ level: "severe", theMsg: errorMsg });
      return;
    }
    
    // Track migration results
    const migrationResults = {
      successful: [],
      notFound: [],
      errors: []
    };
    
    // Process each property we want to migrate
    for (const propertyKey of PROPERTIES_TO_MIGRATE) {
      try {
        // Check if property exists in Properties Service
        if (!(propertyKey in allProperties)) {
          migrationResults.notFound.push(propertyKey);
          console.warn(`Property '${propertyKey}' not found in Properties Service - skipping`);
          continue;
        }
        
        const propertyValue = allProperties[propertyKey];
        
        // Find the property in the config sheet
        let foundRow = -1;
        let currentValue = null;
        
        for (let i = 1; i < configData.length; i++) { // Start at 1 to skip header
          if (configData[i][0] && configData[i][0].toString().trim() === propertyKey) {
            foundRow = i + 1; // +1 because arrays are 0-indexed but sheets are 1-indexed
            currentValue = configData[i][1]; // Get current value
            break;
          }
        }
        
        if (foundRow === -1) {
          // Property not found in config sheet - ADD IT
          const newRow = configData.length + 1; // Add to end
          configSheet.getRange(newRow, 1).setValue(propertyKey);   // Column A: property name
          configSheet.getRange(newRow, 2).setValue(propertyValue); // Column B: property value
          
          migrationResults.successful.push({
            property: propertyKey,
            oldValue: 'NOT_FOUND',
            newValue: propertyValue,
            row: newRow,
            action: 'ADDED'
          });
          
          console.log(`✓ Added '${propertyKey}': '${propertyValue}' (New Row ${newRow})`);
          
        } else {
          // Property found in config sheet
          if (!currentValue || currentValue.toString().trim() === '') {
            // Current value is empty - UPDATE IT
            configSheet.getRange(foundRow, 2).setValue(propertyValue);
            
            migrationResults.successful.push({
              property: propertyKey,
              oldValue: currentValue || 'EMPTY',
              newValue: propertyValue,
              row: foundRow,
              action: 'UPDATED'
            });
            
            console.log(`✓ Updated '${propertyKey}': '${currentValue || 'EMPTY'}' → '${propertyValue}' (Row ${foundRow})`);
          } else {
            // Current value exists - SKIP IT
            migrationResults.notFound.push({
              property: propertyKey,
              reason: 'HAS_VALUE',
              currentValue: currentValue
            });
            
            console.log(`- Skipped '${propertyKey}': already has value '${currentValue}' (Row ${foundRow})`);
          }
        }
        
      } catch (err) {
        migrationResults.errors.push({
          property: propertyKey,
          error: err.message
        });
        
        console.error(`✗ Error migrating '${propertyKey}': ${err.message}`);
        logIt({ 
          level: "error", 
          theMsg: `Migration error for property '${propertyKey}': ${err.message}` 
        });
      }
    }
    
    // Generate migration report
    const reportSummary = generateMigrationReport(migrationResults);
    console.log("\n" + reportSummary);
    
    // Show user-friendly toast message
    const toastMsg = `Migration completed! ${migrationResults.successful.length} properties migrated successfully.`;
    SpreadsheetApp.getActiveSpreadsheet().toast(toastMsg, "Migration Complete", 5);
    
    // Log final summary
    logIt({ 
      level: "info", 
      theMsg: `Properties migration completed. Successful: ${migrationResults.successful.length}, Not found: ${migrationResults.notFound.length}, Errors: ${migrationResults.errors.length}` 
    });
    
  } catch (err) {
    const errorMsg = `Critical error during properties migration: ${err.message}`;
    console.error(errorMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Migration Error", -1);
    logIt({ level: "severe", theMsg: errorMsg, error: err });
  }
}

/**
 * Generates a detailed migration report
 * @param {Object} migrationResults - The results object from the migration process
 * @returns {string} Formatted report string
 * @private
 */
function generateMigrationReport(migrationResults) {
  let report = "=== PROPERTIES MIGRATION REPORT ===\n\n";
  
  // Successful migrations and additions
  if (migrationResults.successful.length > 0) {
    report += `✓ SUCCESSFULLY PROCESSED (${migrationResults.successful.length}):\n`;
    migrationResults.successful.forEach(item => {
      const action = item.action || 'UPDATED';
      const oldValue = item.oldValue === 'NOT_FOUND' ? 'NOT_FOUND' : `"${item.oldValue}"`;
      report += `  • ${item.property}: ${oldValue} → "${item.newValue}" (Row ${item.row}) [${action}]\n`;
    });
    report += "\n";
  }
  
  // Properties skipped (already have values or not found)
  if (migrationResults.notFound.length > 0) {
    report += `⚠ SKIPPED (${migrationResults.notFound.length}):\n`;
    migrationResults.notFound.forEach(item => {
      if (typeof item === 'string') {
        // Old format - property not in Properties Service
        report += `  • ${item} (not in Properties Service)\n`;
      } else {
        // New format - property has value or other reason
        const reason = item.reason === 'HAS_VALUE' ? `already has value "${item.currentValue}"` : item.reason;
        report += `  • ${item.property} (${reason})\n`;
      }
    });
    report += "\n";
  }
  
  // Errors
  if (migrationResults.errors.length > 0) {
    report += `✗ ERRORS (${migrationResults.errors.length}):\n`;
    migrationResults.errors.forEach(error => {
      report += `  • ${error.property}: ${error.error}\n`;
    });
    report += "\n";
  }
  
  report += `SUMMARY: ${migrationResults.successful.length} processed (${migrationResults.successful.filter(i => i.action === 'ADDED').length} added, ${migrationResults.successful.filter(i => i.action === 'UPDATED').length} updated), ${migrationResults.notFound.length} skipped, ${migrationResults.errors.length} errors`;
  
  return report;
}

/**
 * Helper function to view current Properties Service values (for debugging)
 * This function just displays what's currently in Properties Service
 * 
 * @returns {void}
 */
function viewCurrentProperties() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const allProperties = scriptProperties.getProperties();
    
    console.log("=== CURRENT PROPERTIES SERVICE VALUES ===");
    
    if (Object.keys(allProperties).length === 0) {
      console.log("No properties found in Properties Service");
      return;
    }
    
    for (const [key, value] of Object.entries(allProperties)) {
      console.log(`${key}: "${value}"`);
    }
    
    console.log(`\nTotal properties: ${Object.keys(allProperties).length}`);
    
  } catch (err) {
    console.error(`Error reading Properties Service: ${err.message}`);
  }
}

/**
 * Helper function to view current config sheet values (for debugging)
 * This function displays what's currently in the config sheet
 * 
 * @returns {void}
 */
function viewCurrentConfigSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      console.log(`Config sheet '${CONFIG_SHEET_NAME}' not found`);
      return;
    }
    
    const configData = configSheet.getDataRange().getValues();
    
    console.log("=== CURRENT CONFIG SHEET VALUES ===");
    
    for (let i = 0; i < configData.length; i++) {
      const property = configData[i][0] || '';
      const value = configData[i][1] || '';
      
      if (i === 0) {
        console.log(`HEADERS: "${property}" | "${value}"`);
      } else {
        console.log(`${property}: "${value}"`);
      }
    }
    
    console.log(`\nTotal rows: ${configData.length}`);
    
  } catch (err) {
    console.error(`Error reading config sheet: ${err.message}`);
  }
}