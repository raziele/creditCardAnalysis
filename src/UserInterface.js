//==================================
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Check For New Files', functionName: 'main'},
    {name: 'Configure Report Folder', functionName: 'doReportFolderConfig'},
    {name: 'Detect Analysis File', functionName: 'doDetectAnalysisFile'}
  ];
  spreadsheet.addMenu('Credit Cards Analysis', menuItems);
}

//==================================
function doReportFolderConfig(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      'Please enter the folder ID where the reports are to be stored',
      'You can find the ID by going into the folder on Google Drive \n' +  
      'The URL should look like https://drive.google.com/drive/folders/FOLDER_ID',
      ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var newReportFolderId = result.getResponseText();
  if (button == ui.Button.OK) {
    try{
     // const folderName = DriveApp.getFolderById(newId).getName();
      var folderName = DriveApp.getFolderById(newReportFolderId).getName();

    }
    catch(e){
      ui.alert('Error! Folder not found - make sure the ID is right');
    }
    try{
      PropertiesService.getDocumentProperties().setProperty('ID_REPORTS_FOLDER',newReportFolderId);
      ui.alert('Success! Report folder is set to: ' + folderName);
    }
    catch(e){
      ui.alert('Error! Something was not right within the script');
    }
  } else if (button == ui.Button.CANCEL) {
  } else if (button == ui.Button.CLOSE) {
  }
}

//==================================
function doDetectAnalysisFile(){
  var ui = SpreadsheetApp.getUi();
  try{
    const newAnalysisFileId = SpreadsheetApp.getActiveSpreadsheet().getId();
    fileName = DriveApp.getFileById(newAnalysisFileId).getName();
    PropertiesService.getDocumentProperties().setProperty('ID_ANALYSIS_FILE',newAnalysisFileId);
    ui.alert('Success! Analysis file is set to: ' + fileName);
  }
  catch(e){
    ui.alert('Error! file not found');
  }
}