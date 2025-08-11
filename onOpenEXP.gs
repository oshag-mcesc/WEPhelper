function onOpen() {
  let ui = SpreadsheetApp.getUi()
  ui.createMenu('EXPERIMENTAL')
    .addItem("Get Served Area", "WEPhelper.nsServedGiftedRouter.servedGiftedStandard")
    .addItem("Show Settings", "WEPhelper.nsPropsSettings.helperShowSettingsDialog")
    .addItem("Link Emailer", "WEPhelper.nsWEPemail.emailLinks")
    .addSubMenu(ui.createMenu('Prefilled URLs')
      .addItem('Create Initial URLs', 'WEPhelper.nsPropsURLCreator.helperCreateInitialURLs')
      .addItem('Create Mid URLs', 'WEPhelper.nsPropsURLCreator.helperCreateMidURLs')
      .addItem('Create Final URLs', 'WEPhelper.nsPropsURLCreator.helperCreateFinalURLs'))
    .addSubMenu(ui.createMenu('Initial WEPs')
      .addItem('Create Initial WEPs', 'WEPhelper.createWEPdocs.createInitialWEPs')
      .addItem('Add Classes Info', 'WEPhelper.addClassesToTable')
      .addItem('Push Initial Goals', 'WEPhelper.pushInitialGoals'))
    .addSeparator()
    .addItem("Show Alert", "WEPhelper.showAlert")
    .addItem("Push Mid Years", "WEPhelper.pushMidUpdate")
    .addItem("Push Final Eval", "WEPhelper.pushFinalUpdate")
    .addSeparator()
    .addSubMenu(ui.createMenu("Admin")
      .addItem("Reset Sheets", "WEPhelper.reset1.resetEm")
      .addItem("Run Test", "WEPhelper.testIt_g")
      .addItem("Test Log", "WEPhelper.testLog1"))
    .addSeparator()
    .addSubMenu(ui.createMenu("Specials")
      .addItem("HH Get Served", "WEPhelper.nsServedGiftedRouter.servedGiftedHH")
      .addItem("Run Test", "WEPhelper.testIt_g")
      .addItem("Test Log", "WEPhelper.testLog1"))
    .addSeparator()
    .addSubMenu(ui.createMenu('Helpers')
      .addItem('Get DocID list', 'getDocIdList')
      .addItem('Quick DocID list', 'quickDocList')
      .addItem('Delete files in folder', 'deleteTheFiles')
      .addItem('Toggle Tab Visibility', 'toggleTabVisibility'))
    .addToUi()
}


const showAlert_ = () => {
  SpreadsheetApp.getUi().alert("Yo! You did it!")
}

//needed globals for library work
var showAlert = showAlert_


/**
 * Some names that are not the same as the file names:
 * Name in Menu | file name or code name
 * nsWEPemail   | nsEmailer
 * createWEPdocs | createWEPdocs1
 */

function perms() {
  let file = DriveApp.getFileById("1M_gRvGl64WU27y6S0fvQ5xGt0fUf8dNds9QY9FmFfyU")
}

