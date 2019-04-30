//function myEdit() { //onEdit trigger
  var veriSheet = SpreadsheetApp.getActiveSheet();//gets active sheet name
  if (veriSheet.getName() == "Verified Refunds"){ //if on the Verified sheet then:


//Email send function for Verified claims
function sendEmailVerified() {

var spreadsheet = SpreadsheetApp.getActive();//Get active spreadsheet

spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Verified Refunds'), true); //sets Verified Sheet as active sheet

//var sheet = spreadsheet.getActiveSheet(); //gets active sheet

var cell = spreadsheet.getActiveCell().getA1Notation();//get description of the range in A1 notation for easier identification of column for cellA
var activeRow = spreadsheet.getActiveCell().getRow();//gets row number of active cell for datat range

var cellA = ''
    if(cell.indexOf('A')!=-1){ // = if you edit data in col A
      var dataRange = spreadsheet.getRange('A' + activeRow + ':G' + activeRow);//gets row from active cell

var data = dataRange.getValues(); //gets values from columns based on position as below

// Defines variables based on position in columns
for (i in data) {
    var row = data[i];
    var atomRef = row[0];//Atomcore Reference
    var locoRef = row[1];//Loco2 Reference
    var rAmount = row[3];//Refund Amount
    var vDate = row[4];//Verification date
    var rReason = row[5]; //Verification reason
    var notes = row[6]; //Notes
}
    }
//HTML email body
var html =
    '<body>' +
      '<p>The following refund has been verified and a refund needs to be made to the customer:</p>' +
'<br>' +

  //Loco2 Reference includes link to search
'<p><b>Loco2 Reference</b> -  <a href="https://loco2.com/en/admin/orders?utf8=%E2%9C%93&filter%5Bfield%5D=reference&filter%5Bvalue%5D=' + locoRef + '&filter%5Bstart_time%5D=&filter%5Bend_time%5D=&commit=">' + locoRef + ' </a>' + '</p>' +
'<p><b>Atomised Reference</b> - ' + atomRef + '</p>' +
'<p><b>Amount</b> - ' + rAmount + '</p>' +
'<p><b>Refund reason</b> - </p>' +
'<p>' + rReason + '</p>' +
'<p><b>Notes</b> - </p>' +
'<p>' + notes + '</p>' +
'<br>' +
'<p><b><a href=' + spreadsheet.getUrl() + '>Remember to update the spreadsheet!</a></b></p>'
    '</body>'


  MailApp.sendEmail(
    'hello@loco2.com', // recipient
    'Verified Refunds Update - Atomised',    // subject
     '', {
         name: 'RefundBot',//sends from RefundBot
      htmlBody: html  //includes definied html body
    }
  );
};
  }

else if (veriSheet.getName() == "Rejected Refunds"){ //if on the Rejected sheet then:


//Email send function for Rejected claims
function sendEmailRejected() {

var spreadsheet = SpreadsheetApp.getActive();//Get active spreadsheet

spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Rejected Refunds'), true); //sets Rejected Sheet as active sheet

//var sheet = spreadsheet.getActiveSheet(); //gets active sheet

var cell = spreadsheet.getActiveCell().getA1Notation();//get description of the range in A1 notation for easier identification of column for cellA
var activeRow = spreadsheet.getActiveCell().getRow();//gets row number of active cell for datat range

var cellA = ''
    if(cell.indexOf('A')!=-1){ // = if you edit data in col A
      var dataRange = spreadsheet.getRange('A' + activeRow + ':H' + activeRow);//gets row from active cell

var data = dataRange.getValues(); //gets values from columns based on position as below

// Defines variables based on position in columns
for (i in data) {
    var row = data[i];
    var atomRef = row[0];//Atomcore Reference
    var locoRef = row[1];//Loco2 Reference
    var rAmount = row[4];//Refund Amount
    var vDate = row[5];//Verification date
    var rReason = row[6]; //Reason for rejection
    var notes = row[7]; //Notes
}
    }
//HTML email body
var html =
    '<body>' +
      '<p>The following refund has been rejected. Please update the customer as necessary:</p>' +
'<br>' +

  //Loco2 Reference includes link to search
'<p><b>Loco2 Reference</b> -  <a href="https://loco2.com/en/admin/orders?utf8=%E2%9C%93&filter%5Bfield%5D=reference&filter%5Bvalue%5D=' + locoRef + '&filter%5Bstart_time%5D=&filter%5Bend_time%5D=&commit=">' + locoRef + ' </a>' + '</p>' +
'<p><b>Atomised Reference</b> - ' + atomRef + '</p>' +
'<p><b>Amount</b> - ' + rAmount + '</p>' +
'<p><b>Reason for rejection</b> - </p>' +
'<p>' + rReason + '</p>' +
'<p><b>Notes</b> - </p>' +
'<p>' + notes + '</p>' +
'<br>' +
'<p><b><a href=' + spreadsheet.getUrl() + '>Remember to update the spreadsheet!</a></b></p>'
    '</body>'


  MailApp.sendEmail(
    'hello@loco2.com', // recipient
    'Rejected Refunds Update - Atomised',    // subject
     '', {
       name: 'RefundBot', //sends from RefundBot
      htmlBody: html  //includes definied html body
    }
  );
};
}
