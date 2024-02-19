const docID = 'ID of the Google Doc Template'
const folderID = 'ID of the Drive Folder'

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Sheets to Docs PDF')
      .addItem('create New PDF', 'createNewPDF')
      .addToUi();
}

function createNewPDF() {
  const googleDocTemplate = DriveApp.getFileById(docID);
  const destinationFolder = DriveApp.getFolderById(folderID);

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Sheet1');

  const rows = sheet.getDataRange().getValues();
  rows.forEach(function(row, index){
    if (index == 0) return; 
    if (row[1]) return;

    const copy = googleDocTemplate.makeCopy(`${row[0]}`, destinationFolder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    body.replaceText('{{NAME}}', row[0]);
    doc.saveAndClose();

    const pdf = doc.getAs('application/pdf');
    const url = destinationFolder.createFile(pdf).getUrl();

    copy.setTrashed(true);
    
    sheet.getRange(index + 1, 2).setValue(url);
  })
}
