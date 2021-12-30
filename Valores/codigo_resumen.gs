var GLOBAL = {
  compra : "Compra",
  venta : "Venta",
  dividendo : "Dividendo",
  summaryId : "173B3q-LGgtJqXfvdV9GI5Ibr7KkKCvkuxUZi-ogo-Yw",  // You have to update this value
  companyId : "",  
}

function updateStocks() {
  var summarySpreadsheet = SpreadsheetApp.openById(GLOBAL.summaryId);
  var summarySheet = summarySpreadsheet.getActiveSheet();
  
  for(var summaryRow = 3; !isEmptyCell(summarySheet.getRange(summaryRow, 1).getValue()); summaryRow++) {
    var companySpreadsheet = SpreadsheetApp.openByUrl(summarySheet.getRange(summaryRow, 1).getValue());
    var companySheet = companySpreadsheet.getActiveSheet();
    var totalStock = 0;
    var totalCost = 0;
    var totalProfit = 0;
    var dividend = 0;
    
    //debugger;
  
    for(var companyRow = 2; !isEmptyCell(companySheet.getRange(companyRow, 1).getValue()); companyRow++) {
      if(!isEmptyCell(companySheet.getRange(companyRow, 8).getValue())) {
        totalStock = companySheet.getRange(companyRow, 8).getValue();
      }
      if(!isEmptyCell(companySheet.getRange(companyRow, 9).getValue())) {
        totalCost = companySheet.getRange(companyRow, 9).getValue();
      }
      if(!isEmptyCell(companySheet.getRange(companyRow, 12).getValue())) {
        totalProfit = companySheet.getRange(companyRow, 12).getValue();
      }
      if(stringEquals(companySheet.getRange(companyRow, 1).getValue(), GLOBAL.dividendo)) {
        dividend += companySheet.getRange(companyRow, 5).getValue();
      }
    }
    summarySheet.getRange(summaryRow, 3).setValue(totalStock);
    summarySheet.getRange(summaryRow, 4).setValue(totalCost);
    if(totalStock == 0) {
      summarySheet.getRange(summaryRow, 5).setValue("");
    } else {
      summarySheet.getRange(summaryRow, 5).setValue(totalCost / totalStock);
    }
    summarySheet.getRange(summaryRow, 6).setValue(dividend);
    summarySheet.getRange(summaryRow, 7).setValue(totalProfit);
    
    for(var col = 11; !isEmptyCell(summarySheet.getRange(2, col).getValue()); col++) {
      var profitPerYear = getProfitPerYear(summarySheet.getRange(2, col).getValue(),summarySheet.getRange(summaryRow, 1).getValue());
      if(profitPerYear == 0) {
        summarySheet.getRange(summaryRow, col).setValue("");
      } else {
        summarySheet.getRange(summaryRow, col).setValue(profitPerYear);
      }
    }
  }
}

function getProfitPerYear(year, url) {
  var companySpreadsheet = SpreadsheetApp.openByUrl(url);
  var companySheet = companySpreadsheet.getActiveSheet();
  var profit = 0;
  for(var companyRow = 2; !isEmptyCell(companySheet.getRange(companyRow, 1).getValue()); companyRow++) {
    if(stringEquals(companySheet.getRange(companyRow, 1).getValue(), GLOBAL.venta)) {
      if(year == getYear(companySheet.getRange(companyRow, 2).getValue())) {
        profit += companySheet.getRange(companyRow, 11).getValue();
      }
    }
  }
  return profit;
}

function getYear(date){
  var temp = date.getFullYear();
  debugger;
  return date.getFullYear();
}

function stringEquals(text1, text2) {
  //debugger;
  if(text1.length != text2.length) {
    return false;
  }
  if(text1.indexOf(text2)==-1){
    return false;
  }
  return true;
}

function isEmptyCell(cell) {
  //debugger;
  if(stringEquals(typeof cell, "string")) {
    if(cell.trim().length == 0) {
      //debugger;
      return true;
    }
  }
  //debugger;
  return false;
}

/*************************************************/
function updateYear() {
  var summarySpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var yearSheet = summarySpreadsheet.getActiveSheet();
  var year = parseInt(yearSheet.getName());
  if (isNaN(year)) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('No se ha podido obtener el año');
    return;
  } 
  
  yearSheet.deleteRows(3, yearSheet.getLastRow());
  
  var summarySheet = summarySpreadsheet.getSheetByName('Inversion');
  
  for(var summaryRow = 3; !isEmptyCell(summarySheet.getRange(summaryRow, 1).getValue()); summaryRow++) {
    var companySpreadsheet = SpreadsheetApp.openByUrl(summarySheet.getRange(summaryRow, 1).getValue());
    var companySheet = companySpreadsheet.getActiveSheet();
    
    var totalProfit = 0;
    var totalDividend = 0;
    var transOk = false;
    
    for(var companyRow = 2; !isEmptyCell(companySheet.getRange(companyRow, 1).getValue()); companyRow++) {
      var transType = companySheet.getRange(companyRow, 1).getValue();
      var transYear = getYear(companySheet.getRange(companyRow, 2).getValue());
      
      if(transYear == year) {
        if(stringEquals(transType, "Venta")) {
          totalProfit += companySheet.getRange(companyRow, 11).getValue();
          transOk = true;
        }
        if(stringEquals(transType, "Dividendo")) {
          totalDividend += companySheet.getRange(companyRow, 5).getValue();
          transOk = true;
        }
      }
    }
    
    if(transOk) {
      yearSheet.appendRow([summarySheet.getRange(summaryRow, 1).getValue(),summarySheet.getRange(summaryRow, 2).getValue(),totalProfit,totalDividend]);
    }
  }
  
}

