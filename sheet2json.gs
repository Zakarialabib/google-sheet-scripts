function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet4");
  var range = sheet.getDataRange();
  var values = range.getValues();
  var headers = values.shift();
  var json = [];
  for (var row = 0; row < values.length; row++) {
    var obj = {};
    for (var col = 0; col < values[row].length; col++) {
      obj[headers[col]] = values[row][col];
    }
    json.push(obj);
  }
  var output = ContentService.createTextOutput(JSON.stringify(json));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
