function main(){
  const analysisFileId = PropertiesService.getDocumentProperties().getProperty('ID_ANALYSIS_FILE');
  const analysisFile = SpreadsheetApp.openById(analysisFileId);
  //const analysisFile = SpreadsheetApp.getActiveSpreadsheet();
  const analysisDbSheet = analysisFile.getSheetByName('DB');
  const analysisCategoriesSheet = analysisFile.getSheetByName('Categories');
  const analysisStatusSheet = analysisFile.getSheetByName('_status');
  const idReportFolder = PropertiesService.getDocumentProperties().getProperty('ID_REPORTS_FOLDER');
  const newFiles = get_new_files_by_list(idReportFolder, analysisStatusSheet);
  var categoriesTable = analysisCategoriesSheet.getDataRange().getValues();
  const transactionDetailsTemplate = new makeStruct("inputDate, fid, fname, type, nCard, billingMonth, transactionDate, name, amount, currency");


  for (var i=0; i<newFiles.length; i++){
    fid = newFiles[i];
    
    var fileType = get_credit_card_type(fid);
    src_fh = SpreadsheetApp.openById(fid);
    
    switch(fileType){
      case "Max":
        out = get_max_data(src_fh, fid, categoriesTable);
        break;
      case 'Isracard':
        out = get_isracard_data(src_fh, fid, categoriesTable);
        break;
      case 'Visa':
        out = get_visa_data(src_fh, fid, categoriesTable);
        break;
      default:
        out = [];
        break;
    }
    var data = [];
    for (j=0; j< out['nRow']; j++){
      data[j] = []
      for (var key in out['data'][j]){
        data[j].push(out['data'][j][key]);
      }
    }
    for (k=0; k< out['nRow']; k++){
      analysisDbSheet.appendRow(data[k]);
    }
  }
  
  Logger.log('done!');
}


