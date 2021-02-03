//==================================
//GLOBAL PARAMETERS
FILENAME_PREFIX_VISA = 'Transactions'
FILENAME_PREFIX_ISRACARD = 'Export'
FILENAME_PREFIX_MAX = 'transaction'

const transactionDetailsTemplate = new makeStruct("inputDate, fid, fname, type, nCard, billingMonth, transactionDate, name, amount, currency");
const dateRegex = new RegExp(/^(?:(?:31(\/|-|\.)(?:0?[13578]|1[02]))\1|(?:(?:29|30)(\/|-|\.)(?:0?[13-9]|1[0-2])\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:29(\/|-|\.)0?2\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\d|2[0-8])(\/|-|\.)(?:(?:0?[1-9])|(?:1[0-2]))\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$/g);

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
        formattedRow.currency = charToCurrencyCode(row[6]);
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
        formattedRow.currency = charToCurrencyCode(row[6]);
      }
      else{
        formattedRow.name = row[1];
        formattedRow.amount = row[4];
        formattedRow.currency = charToCurrencyCode(row[5]);
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

  const billMonthRegex = new RegExp(/\s(?<billMonth>[0-9]+\/[0-9]+)[^0-9\/]/);
  const nCardRegex = new RegExp(/\s([0-9]{4})/); 

  for (var s = 0; s < sheets.length; s++){
    lastRow = sheets[s].getLastRow();
    lastCol = sheets[s].getLastColumn();
    var data = sheets[s].getRange(1,1,lastRow,lastCol).getValues();

    for (var r=0; r < lastRow - 1; r++){
      var row = data[r];
      //var isFullDate = dateRegex.exec(row[0]);
      const dateCellFormat = typeof row[0];
      
      if (r < 3){ // first lines contain the billing date and card number
        const billDate = billMonthRegex.exec(row[0]);
        const nCardRe = nCardRegex.exec(row[0]);
        
        if (billDate != null){
        billingMonth.setYear((parseInt(billDate.groups.billMonth.split('/')[1])));
        billingMonth.setMonth(parseInt(billDate.groups.billMonth.split('/')[0])-1);
        }
        if (nCardRe != null ){
        nCard = nCardRe[0];
        }
      }
      //else if {dateCellFormat == "object" || isFullDate != null){
      else{

        var formattedRow = new transactionDetailsTemplate();

        //determine transaction date
        if (dateCellFormat == "object"){
           formattedRow.transactionDate = row[0].toLocaleDateString("en-IL");
        }
        else{
          var rSplit = row[0].split('/');
          formattedRow.transactionDate = new Date("20" + rSplit[2], rSplit[1]-1, rSplit[0]); //month is zero-based 
        } 

        formattedRow.inputDate = inputDate;
        formattedRow.type = 'Visa';
        formattedRow.nCard = nCard;
        formattedRow.billingMonth = billingMonth.toLocaleDateString("en-US");
    
        formattedRow.name = row[1];
        var amountRegex = new RegExp(/[0-9].{0,10}\.[0-9]{0,5}/);

        formattedRow.amount = amountRegex.exec(row[3])[0];
        
        formattedRow.currency = charToCurrencyCode('₪');
        formattedRow.fid = fid;
        formattedRow.fname = fname;
                     
        out_data.push(formattedRow);
        continue; 
      }
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