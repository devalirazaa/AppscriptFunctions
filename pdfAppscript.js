//this is code to print hole spreadsheet pdf
function exportSpreadsheetToPDF() {
  var blob, exportUrl, options, pdfFile, response, ss, ssID, url_base;

  // Get the active spreadsheet and its ID
  ss = SpreadsheetApp.getActiveSpreadsheet();
  ssID = ss.getId();
  url_base = ss.getUrl().replace(/edit$/, '');

  // Export URL to export the whole spreadsheet
  exportUrl = url_base + 'export?exportFormat=pdf&format=pdf' +
    '&id=' + ssID +
    '&size=A4' + // Paper size
    '&portrait=true' + // Orientation, false for landscape
    '&fitw=true' + // Fit to width
    '&sheetnames=true&printtitle=false&pagenumbers=true' + // Header and footer settings
    '&gridlines=false' + // Hide gridlines
    '&fzr=false'; // Do not repeat frozen rows on each page

  options = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
    },
    muteHttpExceptions: true
  };

  try {
    response = UrlFetchApp.fetch(exportUrl, options);

    if (response.getResponseCode() !== 200) {
      Logger.log("Error exporting spreadsheet to PDF. Response Code: " + response.getResponseCode());
      return;
    }

    blob = response.getBlob();
    blob.setName('Complete_Spreadsheet_Export.pdf');

    // Save the PDF file in Google Drive
    pdfFile = DriveApp.createFile(blob);
    Logger.log('PDF file created: ' + pdfFile.getId());
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}
// and this generate pdf for only one sheet
function exportRangeToPDf(range) {
  var blob,exportUrl,options,pdfFile,response,sheetTabNameToGet,sheetTabId,ss,ssID,url_base;

  range = range ? range : "C4:D4";//Set the default to whatever you want

  sheetTabNameToGet = "Sheet1";//Replace the name with the sheet tab name for your situation
  ss = SpreadsheetApp.getActiveSpreadsheet();//This assumes that the Apps Script project is bound to a G-Sheet
  ssID = ss.getId();
  sh = ss.getSheetByName(sheetTabNameToGet);
  sheetTabId = sh.getSheetId();
  url_base = ss.getUrl().replace(/edit$/,'');

  //Logger.log('url_base: ' + url_base)

  exportUrl = url_base + 'export?exportFormat=pdf&format=pdf' +

    '&gid=' + sheetTabId + '&id=' + ssID +
    '&range=' + range + 
    //'&range=NamedRange +
    '&size=A4' +     // paper size
    '&portrait=true' +   // orientation, false for landscape
    '&fitw=true' +       // fit to width, false for actual size
    '&sheetnames=true&printtitle=false&pagenumbers=true' + //hide optional headers and footers
    '&gridlines=false' + // hide gridlines
    '&fzr=false';       // do not repeat row headers (frozen rows) on each page

  //Logger.log('exportUrl: ' + exportUrl)

  options = {
    headers: {
      'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken(),
    }
  }

  options.muteHttpExceptions = true;//Make sure this is always set

  response = UrlFetchApp.fetch(exportUrl, options);

  //Logger.log(response.getResponseCode())

  if (response.getResponseCode() !== 200) {
    console.log("Error exporting Sheet to PDF!  Response Code: " + response.getResponseCode());
    return;

  }
  
  blob = response.getBlob();

  blob.setName('AAA_test.pdf')

  pdfFile = DriveApp.createFile(blob);//Create the PDF file
  //Logger.log('pdfFile ID: ' +pdfFile.getId())
}
// and this for gnerate more and seperately for every sheet
function exportRangeToPDF(range) {
  var blob, exportUrl, options, pdfFile, response, ss, ssID, url_base;

  // Set default range if not provided
  range = range || "A1:Z100";

  // List of sheet names to export
  var sheetTabNameToGet = ["Sheet1", "Sheet2"]; // Replace with your sheet tab names
  ss = SpreadsheetApp.getActiveSpreadsheet(); // Assume this Apps Script is bound to a spreadsheet
  ssID = ss.getId();
  url_base = ss.getUrl().replace(/edit$/, '');

  // Iterate through each sheet
  for (var i = 0; i < sheetTabNameToGet.length; i++) {
    var sh = ss.getSheetByName(sheetTabNameToGet[i]);
    if (!sh) {
      Logger.log("Sheet not found: " + sheetTabNameToGet[i]);
      continue; // Skip to the next sheet if not found
    }

    var sheetTabId = sh.getSheetId();

    // Construct export URL
    exportUrl =
      url_base +
      "export?exportFormat=pdf&format=pdf" +
      "&gid=" +
      sheetTabId +
      "&id=" +
      ssID +
      "&size=A4" + // Paper size
      "&portrait=true" + // Orientation: true for portrait, false for landscape
      "&fitw=true" + // Fit to width
      "&sheetnames=true&printtitle=false&pagenumbers=true" + // Optional headers/footers
      "&gridlines=false" + // Hide gridlines
      "&fzr=false"; // Do not repeat frozen rows

    Logger.log("Export URL: " + exportUrl);

    // Set up request options
    options = {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken(),
      },
      muteHttpExceptions: true, // Ensure all responses are caught
    };

    // Fetch the file
    response = UrlFetchApp.fetch(exportUrl, options);

    // Handle response
    if (response.getResponseCode() !== 200) {
      Logger.log("Error exporting sheet: " + sheetTabNameToGet[i]);
      continue; // Skip to the next sheet on error
    }

    // Save the response as a PDF
    blob = response.getBlob().setName(sheetTabNameToGet[i] + "_Export.pdf");
    pdfFile = DriveApp.createFile(blob);
    Logger.log("PDF created for sheet: " + sheetTabNameToGet[i] + ", File ID: " + pdfFile.getId());
  }
}
