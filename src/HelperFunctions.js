//==================================
function get_column_letter_from_name(row, name){
  return String.fromCharCode(65 + row.indexOf(name));
}

//==================================
function charToCurrencyCode(c){
  var code;
  
  switch(c){
    case String.fromCharCode(0x20aa): // Shekels
      code = "ILS";
      break;
    case String.fromCharCode(0x24): // USD
      code = "USD";
      break;
    case String.fromCharCode(0x20AC): // Euro
      code = "EUR"; 
      break;
    default:
      code = "NA";
      break;
  }
  return code;
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

//==================================
function get_row_values(sheet, rowNum){
  return sheet.getRange([rowNum, rowNum].join(':')).getValues()[0];
}

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