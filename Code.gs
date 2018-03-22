function onOpen() {
  
  
 var ui = SpreadsheetApp.getUi();
  ui.createMenu('Planungstool')
  .addItem('Archive Project', 'archiveRange')
  .addToUi();
  
  
} 

// Neues Archivierungsskript

function archiveRange() {
 
  // Zuweisung aktuelles Spreadsheet  
 var sss = SpreadsheetApp.getActiveSpreadsheet();
  
 // Zuweisung Projektplanung Sheet
  
 var ss = sss.getSheetByName('Projektplanung'); //replace with source Sheet tab name
 var ss2 = sss.getActiveSheet();
 var sss2 = SpreadsheetApp.openById('1kNQIwcuAJY2KvOL0NQZp3IxuyUMZZjYZNWCPPsHkyr0');
 var ss3 = sss2.getSheetByName('1');
  
 // Konstanten
  
  // Abstand von Zelle A bis letzte Zeile der Projekt Metadaten (alles bis zum Zeitstrahl)
  
  var PROJECT_DATA = 12;
  var DATE_ROW = 3;
  
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
 var month_name = ["Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember"];
 
 // Deklaration Array für Index zwischenspeicher 
  var time_index = [];
   
  // Zählervariable
  var r=0; // Row
  var c=0; // Column
  
  
  // Beginn erste For-Schleife, die alle Zeilen durchgeht
  for(r=0;r<diff;r++){
    
    
    // Beginn For-Schleife, die alle Spalten durchgeht
    for(c=PROJECT_DATA;c<last_col;c++){
     
      // Abfrage, wenn Wert in Array hinterlegt ist wird Zeile (r) und Spalte (c) der Daten Matrix gespeichert.
      if(data[r][c] != ""){
                    
      
        var pnumber = data[r][0];
        var brand = data[r][1];
        var project = data[r][2];
        var phase = data[r][3];
        var po = data[r][4];
        var guild = data[r][5];
        var person = data[r][6];
        var task = data[r][7];
        var sold = data[r][10];
        var time_date = ss2.getRange(DATE_ROW,1+PROJECT_DATA+c).getCell(1,1).getValue();
        var time_year = time_date.getYear();
        var time_month = month_name[time_date.getMonth()] + " " + time_year;
        var dailyhours = data[r][c];
        
        var rowContents = [pnumber,brand,project,phase,po,guild,person,task,sold,time_month,time_year,time_date,dailyhours];
        
        ss3.appendRow(rowContents);
        
        
        
        
      } // Ende If
    } // Ende For
  } // Ende For Haupt
  
 
}