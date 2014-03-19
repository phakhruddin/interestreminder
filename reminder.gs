var ss = SpreadsheetApp.getActiveSpreadsheet();
var Bank = ss.getSheetName();

function exposed() {
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var content = " ";
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var startexposure = 0;
  var table = '<table style=\"font-family: Arial; line-height: 19px; width: 700px; border-collapse: collapse; font-size: 12px;\"><tr>';
  for (var j = 0; j < data[0].length; j++) {
    ( data[0][0] )? head = "Date" : data[0][j];
    table += '<th style=\"border: 1px solid rgb(221, 221, 221); padding: 10px; background-color: rgb(245, 245, 245); text-align: left; font-size: 1.1em;\">' + head + '</th>';
  }
  for (var i = 1; i < data.length; i++) {
    var celldate = data[i][0];    
    var lapse = 30 * 24 * 3600 * 1000;
    var exposed_day = today.getTime() - lapse;
    //var dayexp = new Date(exposed_day);
    ( data[i][5] > 0 ) ? needbgcolor = 1 : needbgcolor = 0;
    Logger.log("expose: "+data[i][5] +" needbgcolor: "+needbgcolor);
    if ( startexposure == 0 && data[i][0] ) {
      if ( needbgcolor == 1 ) {
        var dayexpinms = data[i][0].getTime() + lapse;
        var dayexp = new Date(dayexpinms).toLocaleDateString("en-US"); 
        var startexposure = 1; // done
      }
    } 
    Logger.log("day start exposed" +dayexp);
    if ( celldate > exposed_day ) {  
      table += '</tr><tr>';
      for (var j = 0; j < data[i].length; j++) {
        if ( typeof data[i][j] == "number" ) {
        var cell = data[i][j].toFixed(2);
        } else {
          if (data[i][j] instanceof Date) { 
            var cell = data[i][j].toLocaleDateString("en-US");
            } else { var cell = data[i][j] }     
        }  
        if ( needbgcolor == 1 ) { 
          table += '<td style=\"border: 1px solid rgb(221, 221, 221); padding: 7px;background-color: rgb(246, 204, 218);\">' + cell + '</td>'; 
          Logger.log("color background")
          } else {
          table += '<td style=\"border: 1px solid rgb(221, 221, 221); padding: 7px;\">' + cell + '</td>'; 
          Logger.log("DO NOT color bg");
        }
      }
      var x=1;
    }
  }
  table += '</tr></table>';
  var body = table +"<br /><br /><a href=\"" + ss.getUrl() + "\">Open spreadsheet</a><br />";
  var recipient = "phakhruddin@gmail.com";
  MailApp.sendEmail(recipient, 'Pay '+Bank+' before '+dayexp+' : ' + ss.getName(), '', {
        htmlBody: body
  });
}
