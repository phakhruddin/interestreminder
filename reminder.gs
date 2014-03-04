var ss = SpreadsheetApp.getActiveSpreadsheet();

function exposed() {
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var content = " ";
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var table = '<table style=\"font-family: Arial; line-height: 19px; width: 700px; border-collapse: collapse; font-size: 12px;\"><tr>';
  for (var j = 0; j < data[0].length; j++) {
    table += '<th style=\"border: 1px solid rgb(221, 221, 221); padding: 10px; background-color: rgb(245, 245, 245); text-align: left; font-size: 1.1em;\">' + data[0][j] + '</th>';
  }
  for (var i = 1; i < data.length; i++) {
    var celldate = data[i][0];
    ( data[i][5] > 0 ) ? needbgcolor = 1 : needbgcolor = 0;
    //Logger.log("expose: "+data[i][5] +" needbgcolor: "+needbgcolor);
    var lapse = 30 * 24 * 3600 * 1000;
    var exposed_day = today.getTime() - lapse; 
    var dayexp = exposed_day.getday();
    if ( celldate > exposed_day ) {
      exposed_content = data[i][1];    
      table += '</tr><tr>';
      for (var j = 0; j < data[i].length; j++) {
        if ( typeof data[i][j] == "number" ) {
        var cell = data[i][j].toFixed(2);
        } else {
        var cell = data[i][j];
        }  
        if ( needbgcolor == 1 ) { 
          table += '<td style=\"border: 1px solid rgb(221, 221, 221); padding: 7px;background-color: rgb(246, 204, 218);\">' + cell + '</td>'; 
          } else {
          table += '<td style=\"border: 1px solid rgb(221, 221, 221); padding: 7px;\">' + cell + '</td>'; 
        }
      }
      var x=1;
    }
  }
  table += '</tr></table>';
  var body = table +"<br /><br /><a href=\"" + ss.getUrl() + "\">Open spreadsheet</a><br />";
  var recipient = "phakhruddin@gmail.com";
  MailApp.sendEmail(recipient, 'Interest reminder for : ' + ss.getName(), '', {
        htmlBody: body
  });
}
