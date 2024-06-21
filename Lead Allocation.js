function copyLeadsToOGV() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = spreadsheet.getSheetByName("Sheet Name");
  var leadsOGVSheet = SpreadsheetApp.openByUrl("link").getSheetByName("Leads");

  var dataRange = leadsSheet.getRange("B7:P" + leadsSheet.getLastRow());
  var dataValues = dataRange.getValues();

  for (var i = 0; i < dataValues.length; i++) {
    var leadData = dataValues[i];
    if (leadData[14] === "oGV") {
      // Check if "Our choice" column value is "oGV"
      var leadName = leadData[1];
      var leadPhone = leadData[4];
      var leadExists =
        leadsOGVSheet.getRange("B:B").getValues().flat().indexOf(leadName) !==
          -1 &&
        leadsOGVSheet.getRange("D:D").getValues().flat().indexOf(leadPhone) !==
          -1;

      if (!leadExists) {
        var emptyRow = findFirstEmptyRow(leadsOGVSheet);
        var leadInfo = [
          leadName, // Lead Name
          leadData[3], // Email Address
          leadPhone, // Phone Number
          leadData[5], // Faculty
          leadData[9], // Field of Study
          leadData[10], // Level of Study
        ];
        leadsOGVSheet
          .getRange(emptyRow, 2, 1, leadInfo.length)
          .setValues([leadInfo]);
        Logger.log("Lead '" + leadName + "' copied to Leads oGV sheet.");
      } else {
        Logger.log("Lead '" + leadName + "' exists.");
      }
    }
  }
}

function findFirstEmptyRow(sheet) {
  var range = sheet.getRange("B7:B");
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      return range.getRow() + i;
    }
  }
  return range.getLastRow() + 1;
}
