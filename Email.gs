function Export_And_Email() {
  //Hide List
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i in sheets) {
    var sheetName = new String;
    sheetName = sheets[i].getName();
    if (sheetName.indexOf("List") !== -1) {
      sheets[i].hideSheet();
    }
  }

  //All requests must include id in the path and a format parameter
  //https://docs.google.com/spreadsheets/d/{SpreadsheetId}/export

  //FORMATS WITH NO ADDITIONAL OPTIONS
  //format=xlsx       //excel
  //format=ods        //Open Document Spreadsheet
  //format=zip        //html zipped          

  //CSV,TSV OPTIONS***********
  //format=csv        // comma seperated values
  //             tsv        // tab seperated values
  //gid=sheetId             // the sheetID you want to export, The first sheet will be 0. others will have a uniqe ID

  // PDF OPTIONS****************
  //format=pdf     
  //size=0,1,2..10             paper size. 0=letter, 1=tabloid, 2=Legal, 3=statement, 4=executive, 5=folio, 6=A3, 7=A4, 8=A5, 9=B4, 10=B5  
  //fzr=true/false             repeat row headers
  //portrait=true/false        false =  landscape
  //fitw=true/false            fit window or actual size
  //gridlines=true/false
  //printtitle=true/false
  //pagenum=CENTER/UNDEFINED      CENTER = show page numbers / UNDEFINED = do not show
  //attachment = true/false      dunno? Leave this as true
  //gid=sheetId                 Sheet Id if you want a specific sheet. The first sheet will be 0. others will have a uniqe ID. 
  // Leave this off for all sheets. 
  // EXPORT RANGE OPTIONS FOR PDF
  //need all the below to export a range
  //gid=sheetId                must be included. The first sheet will be 0. others will have a uniqe ID
  //ir=false                   seems to be always false
  //ic=false                   same as ir
  //r1=Start Row number - 1        row 1 would be 0 , row 15 wold be 14
  //c1=Start Column number - 1     column 1 would be 0, column 8 would be 7   
  //r2=End Row number
  //c2=End Column number

  var ssID = "<<spereadsheet ID>>";  //Spreadsheet ID *Remember to change it when doing backup or copy to edit

  var xlsx = "https://docs.google.com/spreadsheets/d/" + ssID + "/export" + "?format=xlsx&"
  var pdf = "https://docs.google.com/spreadsheets/d/" + ssID + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=false&" +
    "portrait=false&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true";
  var params = {
    method: "GET",
    headers: {
      "authorization": "Bearer " + ScriptApp.getOAuthToken()
    }
  };

  var pdf_file = UrlFetchApp.fetch(pdf, params).getBlob();
  var excel_file = UrlFetchApp.fetch(xlsx, params).getBlob();

  //send the pdf as email
  var email = [];
  //Recipients

  email[0] = "email0@gmail.com";
  email[1] = "email1@gmail.com";
  email[2] = "email2@gmail.com";
  email[3] = "email3@gmail.com";


  //Email Format
  var date = new Date();
  var Subjetdate = Utilities.formatDate(date, "GMT+8:00", "ddMMyyyy");
  var subject = "<<subject prefix>> " + Subjetdate;
  var body = "<<email body>>";
  MailApp.sendEmail(email, subject, body, {
    attachments:
      [
        {
          fileName: "<<spreadsheet prefix>>" + "(" + Subjetdate + ")" + ".pdf",
          content: pdf_file.getBytes(),
          mimeType: "application/pdf",
        },
        {
          fileName: "<<spreadsheet prefix>>" + "(" + Subjetdate + ")" + ".xlsx",
          content: excel_file.getBytes(),
          mimeType: "application/xlsx"
        }
      ]
  });
}
