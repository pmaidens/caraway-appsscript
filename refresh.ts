
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Menu") // Name of the menu
    .addItem("Refresh Data", "refreshDate") // Add an item
    .addToUi();
}

function refreshDate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  if (!sheet) {
    Logger.log("Sheet 'Data' not found.");
    return;
  }

  const target = sheet.getRange("A2"); // Correct method to get a specific cell
  const currentDateTime = new Date(); // Get the current date and time
  target.setValue(currentDateTime);   // Write the current date and time into A2
}