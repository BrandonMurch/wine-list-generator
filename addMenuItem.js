function onOpen() {
  SpreadsheetApp.getUi().createMenu("Wine List Generator").addItem("Generate", "createWineList").addToUi();
}

function createWineList() {
  WineListGenerator.createWineList();
}
