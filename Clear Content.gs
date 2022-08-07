function protect(source, FormName, RangeName, Description) {
  var protection = source.getSheetByName(FormName).getRange(RangeName).protect().setDescription(Description);

  protection.removeEditors(protection.getEditors());
  protection.addEditor(["<<rightful editor>>"]);
}

function Clear_Content() {
  var source = SpreadsheetApp.openById("<<source spreadsheet id>>");
  var format = SpreadsheetApp.openById("<<format spreadsheet id>>");

  //Set Sheet variables
  var classroom_sheet = source.getSheetByName("<<Sheet name>>");
  var classroom_format = format.getSheetByName("<<Sheet format>>");

  //Replace Sheet
  source.deleteSheet(classroom_sheet);
  classroom_format.copyTo(source).setName("<<Sheet name>>");

  Utilities.sleep(500);
  protect(source, "<<Sheet name>>", "A4:A26", "Room");
  protect(source, "<<Sheet name>>", "A1:AD3", "ItemRow");

  protect(source, "<<Sheet name>>", "E19:I20", "Grey");
  protect(source, "<<Sheet name>>", "D24:V26", "Grey");

}
