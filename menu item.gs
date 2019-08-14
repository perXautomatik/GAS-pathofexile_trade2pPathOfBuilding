 function onOpen() {
 var ui = SpreadsheetApp.getUi();

 ui.createMenu('JsonParser')
    .addItem('ParseJson', 'menuItem1')
    .addToUi(); 
}

function menuItem1() {
      JsonToItemz();
    }