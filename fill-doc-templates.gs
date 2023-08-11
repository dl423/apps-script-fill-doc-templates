LOG_DATA = false
FOLDER_ID = ''       //folder in which to generate all the filled docs
DOC_TEMPLATE_ID = ''   //the doc template that has several placeholders to be replaced


function Main() {
  //Set up the folder workspace
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var sheet = SpreadsheetApp.getActiveSheet();
  var pdfFolder = folder.createFolder('PDF_Files');

  var maxRow = GetFirstEmptyRowByColumnArray(sheet) - 1;
  Logger.log('Number of items in sheet: ' + maxRow);

  // Parse the spreadsheet item data
  data = sheet.getRange(1, 1, maxRow, 8).getValues();
  var items = [];
  for (var row = 1; row < data.length; row++) {
    var item = {
      fileName: data[row][0],
      content1: data[row][1],
      content2: data[row][2]
    };
    items.push(item);
  }
  
  if (LOG_DATA) {
    for (var j = 0; j < items.length; j++) {
      var itemData = items[j];
      Logger.log('item ' + (j + 1) + ':');
      Logger.log('Date: ', + itemData.date);
      Logger.log('item Title: ' + itemData.itemTitle);
      Logger.log('fileName: ' + itemData.fileName);
      Logger.log('content2: ' + itemData.content2);
      Logger.log('-------------------------');
    }
  }

  // Loop through each item
  for (var itemNum = 0; itemNum < items.length; itemNum++) {
    item = items[itemNum]; //current item to generate doc for

    FillInFile(item, folder, pdfFolder);
    Logger.log('Fill completed for item #' + (itemNum+1));

  }
}

function FillInFile(item, folder, pdfFolder) {
  // Copy the doc template, rename the copy, and move it to designated folder
  var templateDoc = DriveApp.getFileById(DOC_TEMPLATE_ID);
  var doc = templateDoc.makeCopy();
  doc.setName("New File" + item.fileName);
  doc.moveTo(folder);

  // Replace the placeholders with actual content for this item
  doc = DocumentApp.openById(doc.getId());
  body = doc.getBody();
  body.replaceText('{replacement 1}', item.fileName)
  body.replaceText('{replacement 2}', item.content1)
  body.replaceText('{replacement 3}', item.content2);
  doc.saveAndClose();

  CreatePDFFile(doc, pdfFolder);
}

// Converts a file (e.g. Google Docs) into a PDF
function CreatePDFFile(file, folder) {
    var theBlob = file.getBlob().getAs('application/pdf');
    var newPDFFile = folder.createFile(theBlob);

    var fileName = file.getName().replace(".", ""); //otherwise filename will be shortened after full stop    
    newPDFFile.setName(fileName + ".pdf");
    //newPDFFile.moveTo(folder);
}

// Obtains the index of the first empty row in column B
function GetFirstEmptyRowByColumnArray(sheet) {
  var values = sheet.getRange('B:B').getValues();
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}
