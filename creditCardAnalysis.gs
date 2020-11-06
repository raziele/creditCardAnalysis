/*==================================
This script analyzes Israeli credit cards 
and builds a db people can build nice dashboards from.
Instructions for a new Sheet:
- make sure the sheet has the following tabs:
--DB
--Status
--Categories
--_categories_flattened (this is a utility tab to hold temporary calculations)

DB headers. Rows for all headers but last (category) are created by this script.
-Date Added
-Card Type
-Card Number
-Billing Month
-Transaction Date
-Business Name
-Amount
-Currency
-Category
-- this column is created automatically by using the following formula:
"=arrayformula(iferror(vlookup(F2:F,'_categories_flattened'!A2:B,2,FALSE),"ללא סיווג"))"

Status headers. rows for all headers are created by this script.
-File Name
-File ID

Categories tab 
- This tab is handled manually (you can choose your own categories and where each business goes)
- The only mandatory column is the last one which is "ללא סיווג". Its first row should run the following formula:
"=QUERY(DB!F2:I,"select F where I='ללא סיווג'",0)"

_categories_flattened tab
This tab organizes the categories table from previous tab in a list.
Headers:
- Business name
-- "=filter(flatten(Categories!A2:O),NOT(ISBLANK(flatten(Categories!A2:O))))"
- Category
-- "=arrayformula(reverseLookup(A2:A,transpose(Categories!A:O)))"
*/

//==================================
//PARAMETERS
ID_ANALYSIS_FILE = ''
ID_REPORTS_FOLDER = ''

FILENAME_PREFIX_VISA = 'Transactions'
FILENAME_PREFIX_ISRACARD = 'Export'
FILENAME_PREFIX_MAX = 'transaction'

var transactionDetailsTemplate = new makeStruct("inputDate, fid, fname, type, nCard, billingMonth, transactionDate, name, amount, currency");

//==================================
//Get the list of files from a folder
function get_file_list(folderID){
  var folderHandle = DriveApp.getFolderById(folderID);
  return folderHandle.getFiles();
}

//==================================
//Create an arbitrary struct (for building the rows in the db)
function makeStruct(names = transactionDetailsTemplate) {
  var names = names.split(', ');
  var count = names.length;
  function constructor() {
    for (var i = 0; i < count; i++) {
      this[names[i]] = arguments[i];
    }
  }
  return constructor;
}

//==================================
// Detect credit card type by file name
function get_credit_card_type(fid){
  var file = DriveApp.getFileById(fid);
  var filename = file.getName();
  
  if(filename.startsWith(FILENAME_PREFIX_ISRACARD)){
    return 'Isracard';
  }
  else if(filename.startsWith(FILENAME_PREFIX_VISA)){
          return 'Visa';
          }
  else if(filename.startsWith(FILENAME_PREFIX_MAX)){
    return 'Max';
  }
  else{
    return 'Unknown';
  }
}

//==================================
// Parse MAX credit card files
function get_max_data(f_handler, fid, categoriesTable){
  sheets = f_handler.getSheets();
  var out_data = [];
  var nCard = '';
  var billingMonth = new Date(0);
  var inputDate = new Date();
  var fname = SpreadsheetApp.openById(fid).getName()
  
  for (var s = 0; s < sheets.length; s++){
    lastRow = sheets[s].getLastRow();
    lastCol = sheets[s].getLastColumn();
    var data = sheets[s].getRange(1,1,lastRow,lastCol).getValues();

        
    for (var r=0; r < lastRow; r++){
      var row = data[r];

      if (row[0].split('-').length == 3){
        var formattedRow = new transactionDetailsTemplate();
        formattedRow.inputDate = inputDate;
        formattedRow.type = 'Max';
        formattedRow.nCard = row[3];
        
        formattedRow.billingMonth = new Date(row[9].split('-')[2], row[9].split('-')[1]-1, row[9].split('-')[0]); //month is zero-based
        
        formattedRow.transactionDate = new Date(row[0].split('-')[2], row[0].split('-')[1]-1, row[0].split('-')[0]); //month is zero-based
         
        formattedRow.name = row[1];
        formattedRow.amount = row[5]
        formattedRow.currency = row[6]
        formattedRow.fid = fid;
        formattedRow.fname = fname;
                
        out_data.push(formattedRow);
        continue;
      }
      else{ continue;}
      }
    }
  return {'data': out_data, 'nRow': out_data.length, 'nCol': Object.keys(out_data[0]).length};
 }

//==================================
// Parse Isracard credit card files
function get_isracard_data(f_handler, fid, categoriesTable){
  sheet = f_handler.getSheets()[0];
  lastRow = sheet.getLastRow();
  lastCol = sheet.getLastColumn();
  var data = sheet.getRange(1,1,lastRow,lastCol).getValues();
  var out_data = [];
  var nCard = '';
  var billingMonth = new Date(0); 
  var inputDate = new Date();
  var abroadCharges = 0; 
  var fname = SpreadsheetApp.openById(fid).getName()
  
  for (var r=0; r < lastRow; r++){
    var row = data[r];
    
    if (row[1] == 'מועד חיוב'){
      nCard = row[0].split(' - ')[row[0].split(' - ').length -1];
      billingMonth.setFullYear(['20', row[2].split('/')[2]].join(''));
      billingMonth.setMonth(row[2].split('/')[1]-1);
      //billingMonth.setMonth(billingMonth.getMonth()-1); // because JS Date months are zero-based
      continue;
    }
    else if (row[0] == 'עסקאות בארץ'){
      abroadCharges = 0;
      continue;
    }
    else if (row[0] == 'עסקאות בחו˝ל'){
      abroadCharges = 1;
      continue;
    }
    else if (row[0].split('/').length == 3){
      var formattedRow = new transactionDetailsTemplate();
      formattedRow.inputDate = inputDate;
      formattedRow.type = 'Isracard';
      formattedRow.nCard = nCard;
      formattedRow.billingMonth = billingMonth;
      formattedRow.fid = fid;
      formattedRow.fname = fname;
      
      formattedRow.transactionDate = new Date(row[0].split('/')[2], row[0].split('/')[1]-1, row[0].split('/')[0]); //month is zero-based
      
      if (abroadCharges == 1){
        formattedRow.name = row[2];
        formattedRow.amount = row[5];
        formattedRow.currency = row[6];
      }
      else{
        formattedRow.name = row[1];
        formattedRow.amount = row[4];
        formattedRow.currency = row[5];
        }
      
      if(formattedRow.name == 'TOTAL FOR DATE'){ continue;}
      
      out_data.push(formattedRow);
      
      continue;
      
    }  
    else{ continue;}
  }
  return {'data': out_data, 'nRow': out_data.length, 'nCol': Object.keys(out_data[0]).length};
 }

//==================================
// Parse Visa credit card files 
function get_visa_data(f_handler, fid, categoriesTable){
  sheets = f_handler.getSheets();
  var out_data = [];
  var nCard = '';
  var billingMonth = new Date(0);
  var inputDate = new Date();
  var fname = SpreadsheetApp.openById(fid).getName()
  
  for (var s = 0; s < sheets.length; s++){
    lastRow = sheets[s].getLastRow();
    lastCol = sheets[s].getLastColumn();
    var data = sheets[s].getRange(1,1,lastRow,lastCol).getValues();

        
    for (var r=0; r < lastRow; r++){
      var row = data[r];
 
      if (Object.prototype.toString.call(row[0]) == "[object String]"){
        if (row[0].startsWith('פירוט עסקות לכרטיס')){
          var nCardRe = new RegExp(/(\d{3,4})/);
          nCard = nCardRe.exec(row[0])[0];
        }
        continue;
      }
      else if (Object.prototype.toString.call(row[0]) == "[object Date]"){

        var formattedRow = new transactionDetailsTemplate();
        formattedRow.inputDate = inputDate;
        formattedRow.type = 'Visa';
        formattedRow.nCard = nCard;
        
        billingMonth.setYear(row[0].getFullYear());
        billingMonth.setMonth(row[0].getMonth()+1);
        
        formattedRow.billingMonth = billingMonth.toLocaleDateString("en-US");
        formattedRow.transactionDate = row[0].toLocaleDateString("en-US");
        //formattedRow.transactionDate = new Date(row[0].split('/')[2], row[0].split('/')[1]-1, row[0].split('/')[0]); //month is zero-based

        formattedRow.name = row[1];
        formattedRow.amount = row[3]
        formattedRow.currency = '₪';
        formattedRow.fid = fid;
        formattedRow.fname = fname;
                     
        out_data.push(formattedRow);
        continue;
      }
      else{ continue;}
      }
    }
  return {'data': out_data, 'nRow': out_data.length, 'nCol': Object.keys(out_data[0]).length};
 }

//==================================
function update_status_sheet(files, sheet){
  for (m = 0; m < files.length; m++){
    fileId = files[m];
    fileName = SpreadsheetApp.openById(fileId).getName();
    sheet.appendRow([fileName, fileId]);
  }
  return;  
}

//==================================
function get_row_values(sheet, rowNum){
  return sheet.getRange([rowNum, rowNum].join(':')).getValues()[0];
}

//==================================
function get_column_letter_from_name(row, name){
  return String.fromCharCode(65 + row.indexOf(name));
}
  
//==================================
//Compare google folder content to a table and return files not appearing in table
function get_new_files_by_list(folderID, statusSheet){

  var FilesIdColumnLetter = get_column_letter_from_name(get_row_values(statusSheet,1), 'File ID');
  var listOfProcessesFiles = statusSheet.getRange([FilesIdColumnLetter, FilesIdColumnLetter].join(':')).getValues();
  var folderFileList = get_file_list(folderID);
  var listOfProcessesFiles1d = listOfProcessesFiles.map(x => x[0]);
  var newFilesArray = [];
  while(folderFileList.hasNext()){
    var fileID = folderFileList.next().getId();
    var exist = listOfProcessesFiles1d.indexOf(fileID);
    if (exist > -1){
      continue;
    }
    else{
      newFilesArray.push(fileID);
    }
  }
  return newFilesArray;
}


//==================================
function main(){
  var analysisFile = SpreadsheetApp.openById(ID_ANALYSIS_FILE);
  var analysisDbSheet = analysisFile.getSheetByName('DB');
  var analysisStatusSheet = analysisFile.getSheetByName('Status');
  var analysisCategoriesSheet = analysisFile.getSheetByName('Categories');
  
  var newFiles = get_new_files_by_list(ID_REPORTS_FOLDER, analysisStatusSheet);
  var categoriesTable = analysisCategoriesSheet.getDataRange().getValues();
  
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
      default:
        data = 0
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

//==================================
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Check For New Files', functionName: 'main'},
  ];
  spreadsheet.addMenu('Credit Cards Analysis', menuItems);
}

//==================================
/**
 * For a value in a table, teturn the column header where the value is found.
 *
 * @param {value} value.
 * @param {in_range} range to look.
 * @return The index.
 * @customfunction
 */
function REVERSELOOKUP(input, in_range){
  
  function map_value(value, range){
    return value.length > 0 ?
      range.filter(row => row.indexOf(value) > -1)[0][0] : 
      "";
  }
    
  return Array.isArray(input) ?
      input.map(cell => map_value(cell[0], in_range)) :
      map_value(input, in_range);
}
