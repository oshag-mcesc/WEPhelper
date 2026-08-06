const CONDITIONAL_FORMAT_CONFIG = {
  sheetName: "Sheet1",
  rangeA1: "J2:J",
  backgroundColor: "#FFC7CE",
  fontColor: "#9C0006",
  
  // Choose ONE setup below to uncomment and use:
  
  // Option A: The "BAD ID" Custom Formula we made earlier
  criterionType: "CUSTOM_FORMULA",
  value: '=COUNTIF($J$2:$J, "BAD ID") > 0'
  
  // Option B: Built-in "Text Contains" (Uncomment to use instead)
  // criterionType: "TEXT_CONTAINS",
  // value: "BAD ID" 
};

function resetConditionalFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = CONDITIONAL_FORMAT_CONFIG;
  
  const sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) return;
  const range = sheet.getRange(cfg.rangeA1);
  
  sheet.clearConditionalFormatRules();
  
  // 1. Initialize the base rule builder
  let ruleBuilder = SpreadsheetApp.newConditionalFormatRule()
    .setBackground(cfg.backgroundColor)
    .setFontColor(cfg.fontColor)
    .setRanges([range]);
    
  // 2. Dynamically apply the rule type based on your config
  switch (cfg.criterionType) {
    case "CUSTOM_FORMULA":
      ruleBuilder.whenFormulaSatisfied(cfg.value);
      break;
      
    case "TEXT_CONTAINS":
      ruleBuilder.whenTextContains(cfg.value);
      break;
      
    case "TEXT_IS_EMPTY":
      ruleBuilder.whenCellIsEmpty();
      break;
      
    case "NUMBER_GREATER_THAN":
      ruleBuilder.whenNumberGreaterThan(Number(cfg.value));
      break;
      
    default:
      Logger.log("Error: Unknown criterionType provided.");
      return;
  }
  
  // 3. Build and save the rule
  const rule = ruleBuilder.build();
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
  
  Logger.log("Successfully applied " + cfg.criterionType + " rule.");
}
