function onOpen() {
  
  
 var ui = SpreadsheetApp.getUi();
  ui.createMenu('Planungstool')
  .addItem('Archive Project', 'copyRange')
  .addItem('Testi', 'test')
  .addToUi();
  
  
} 


function copyRange() {
 
  // Zuweisung aktuelles Spreadsheet  
 var sss = SpreadsheetApp.getActiveSpreadsheet();
  
 // Zuweisung Projektplanung Sheet
  
 var ss = sss.getSheetByName('Projektplanung'); //replace with source Sheet tab name
 var ss2 = sss.getActiveSheet();
  
 // Konstanten
  
  // Abstand von Zelle A bis letzte Zeile der Projekt Metadaten (alles bis zum Zeitstrahl)
  
  var PROJECT_DATA = 13;
  var PROJECT_META_RANGE = 3;
  
  // Auswahl der aktiven Range 
 var range = sss.getActiveRange();
 
  
 // Bestimmung der Anfangs- und Endzeile der Auswahl 
  
 var start_row = range.getRow();
 var end_row = range.getLastRow();
 var last_col = range.getLastColumn();
 // Browser.msgBox(last_col);
 
 // Bestimmung der Zeilenanzahl 
 var diff = end_row - start_row + 1;
  
  
 
 // Alle Daten werden in 2D Array eingetragen (mit Lücken)
 var data = range.getValues();
 
 // Deklaration Array für Index zwischenspeicher 
  var time_index = [];
   
  // Zählervariable
  var r=0; // Row
  var c=0; // Column
  
  
  // Beginn erste For-Schleife, die alle Zeilen durchgeht
  for(r=0;r<diff;r++){
    
    
    // Beginn For-Schleife, die alle Spalten durchgeht
    for(c=3;c<last_col;c++){
     
      // Abfrage, wenn Wert in Array hinterlegt ist wird Zeile (r) und Spalte (c) der Daten Matrix gespeichert.
      if(data[r][c] != ""){
         
        time_index.push([r, c, data[r][c]]); 
        
        var time_date = ss2.getRange(13,1+c).getCell(1,1).getValue();
    
        var rowContents = [data[r][0],data[r][1],data[r][2],time_date,data[r][c]];
    
        ss2.appendRow(rowContents);
        
        
        
        
      } // Ende If
    } // Ende For
  } // Ende For Haupt
  
 
}