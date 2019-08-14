  function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('JsonParser')
      .addItem('ParseJson', 'menuItem1')
      .addToUi();
    
function menuItem1() {

  var response = SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('parsing!', ui.ButtonSet.YES_NO);
  
// Process the user's response.
if (response == ui.Button.YES) {
  JsonToItemz();
} else {
  Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
}
}
}





