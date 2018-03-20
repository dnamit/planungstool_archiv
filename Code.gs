function onOpen() {
  
  
 var ui = SpreadsheetApp.getUi();
  ui.createMenu('Planungstool')
  .addItem('Archive Project', 'CopyRange')
  .addItem('Testi', 'test')
  .addToUi();
  
  
} 


function copyRange() {
 
  // Zuweisung aktuelles Spreadsheet  
 var sss = SpreadsheetApp.getActiveSpreadsheet();
  
 // Zuweisung Projektplanung Sheet
  
 var ss = sss.getSheetByName('Projektplanung'); //replace with source Sheet tab name
  
 // Auswahl der aktiven Range 
 var range = sss.getActiveRange();
  
 // Bestimmung der Anfangs- und Endzeile der Auswahl 
  
 var start_row = range.getRow();
 var end_row = range.getLastRow();
 var last_col = range.getLastColumn();
 // Browser.msgBox(last_col);
 
 // Bestimmung der Zeilenanzahl 
 var diff = end_row - start_row + 1;
  
  



}


function getData(ss,row,range, last_col){
  
  var DATA_START = 13;
  
  for(var a=0;a<last_col;a++){
   
   x 
    
    
  }
  
}