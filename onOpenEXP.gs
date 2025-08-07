function onOpen() {
  let ui = SpreadsheetApp.getUi()
  ui.createMenu('EXPERIMENTAL')
    .addItem("Show Settings", "WEPhelper.nsPropsSettings.helperShowSettingsDialog")
    .addSubMenu(ui.createMenu('Initial WEPs')
      .addItem('Create the Docs', 'WEPhelper.createWEPdocs.createInitialWEPs')
      .addItem('Add Classes Info', 'WEPhelper.showAlert')
      .addItem('Push Initial Goals', 'WEPhelper.showAlert'))
    .addSeparator()
    .addItem("Create Initial WEPs", "WEPhelper.createWEPdocs.createInitialWEPs")
    .addItem("Show Alert", "WEPhelper.showAlert")
    .addItem("Push Mid Years", "WEPhelper.pushMidUpdate")
    .addItem("Push Final Eval", "WEPhelper.pushFinalUpdate")
    .addSeparator()
    .addSubMenu(ui.createMenu("Admin")
      .addItem("Reset Sheets", "WEPhelper.reset1.resetEm")
      .addItem("Run Test", "WEPhelper.testIt_g")
      .addItem("Test Log", "WEPhelper.testLog1"))
    .addToUi()
}

const showAlert_ = () => {
  SpreadsheetApp.getUi().alert("Yo! You did it!")
}

//needed globals for library work
var showAlert = showAlert_


