function onOpen() {
  let ui = SpreadsheetApp.getUi()
  ui.createMenu('EXPERIMENTAL')
    .addItem("Get Served Area", "WEPhelper.nsServedGiftedRouter.servedGiftedStandard")
    .addItem("Show Settings", "WEPhelper.nsPropsSettings.helperShowSettingsDialog")
    .addSubMenu(ui.createMenu('Initial WEPs')
      .addItem('Create Initial WEPs', 'WEPhelper.createWEPdocs.createInitialWEPs')
      .addItem('Add Classes Info', 'WEPhelper.showAlert')
      .addItem('Push Initial Goals', 'WEPhelper.showAlert'))
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
    .addToUi()
}

const showAlert_ = () => {
  SpreadsheetApp.getUi().alert("Yo! You did it!")
}

//needed globals for library work
var showAlert = showAlert_


