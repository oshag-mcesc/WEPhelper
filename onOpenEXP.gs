function onOpen() {
  let ui = SpreadsheetApp.getUi()
  ui.createMenu('EXPERIMENTAL')
    .addItem("Show Alert","WEPhelper.showAlert")
    .addToUi()
}

const showAlert_ = ()=>{
  SpreadsheetApp.getUi().alert("Yo! You did it!")
}

//needed globals for library work
var showAlert = showAlert_


