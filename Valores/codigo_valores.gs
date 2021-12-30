var GLOBAL = {
  compra : "Compra",
  venta : "Venta",
  dividendo : "Dividendo",
  formId : "1OkeqoNVFT4vPEl5sm-34vYeuPQyXpD1b0qetIFh8eBA",  // You have to update this value
  summaryId : "173B3q-LGgtJqXfvdV9GI5Ibr7KkKCvkuxUZi-ogo-Yw",  // You have to update this value 
  companyId : "",  
}


function onForm() {
  var form = FormApp.getActiveForm();
 
  var formId = form.getId();
  if(!stringEquals(form.getId(),GLOBAL.formId)) {
    return;
  }
  
  var summarySpreadsheet = SpreadsheetApp.openById(GLOBAL.summaryId);
  var summarySheet = summarySpreadsheet.getActiveSheet();
  var responses = form.getResponses();
  var length = responses.length;
  var lastResponse = responses[length-1];
  var itemResponses = lastResponse.getItemResponses();
  var date;
  var stockName;
  var type;
  var stock;
  var cost;
  
  for (var i = 0; i<itemResponses.length; i++) {
    var thisItem = itemResponses[i].getItem().getTitle();
    var thisResponse = itemResponses[i].getResponse();
    switch (thisItem) {
      case "Fecha":
        date = fromDateToString(new Date(thisResponse));
        break;
      case "Valor":
        stockName = thisResponse;
        break;
      case "Tipo de operación":
        type = thisResponse;
        break; 
      case "Número de acciones":
        stock = toInt(thisResponse);
        break;
      case "Importe de la operación":
        cost = toFloat(thisResponse);
        break;
    } 
  }
  
  if(getCompanyId(stockName)) {
    if(stringEquals(type, GLOBAL.compra)) {
      insertPurchase(date, stock, cost);
    } else if(stringEquals(type, GLOBAL.venta)) {
      insertSale(date, stock, cost);
    } else if(stringEquals(type, GLOBAL.dividendo)) {
      insertDividend(date, stock, cost);
    } 
  }
  
  
  /*insertPurchase("10/2/2020", 1800, 12876);
  insertPurchase("12/3/2020", 10, 76);
  insertSale("15/3/2020", 1500, 33976);
  insertPurchase("16/3/2020", 200, 476);
  insertSale("15/3/2020", 400, 3976);
  insertPurchase("8/01/2020",1000,	2442.25);
  insertSale("9/01/2020",1000,2478.84);
  insertPurchase("23/01/2018",1000,2271.10);
  insertDividend("1/1/2020",3000, 880,04);
  insertSale("13/02/2020",1000,	2482.95);
  insertPurchase("28/02/2018",1000,	2397.13);
  insertSale("13/02/2020",1000,	2412.86);*/
}


function getCompanyId(stockName){
  var summarySpreadsheet = SpreadsheetApp.openById(GLOBAL.summaryId);
  var summarySheet = summarySpreadsheet.getActiveSheet();
  for(var row = 2; !isEmptyCell(summarySheet.getRange(row, 2).getValue()); row++) {
    if(stringEquals(summarySheet.getRange(row, 2).getValue(),stockName)) {
      var companySpreadsheet = SpreadsheetApp.openByUrl(summarySheet.getRange(row, 1).getValue());
      if(companySpreadsheet != null) {
        GLOBAL.companyId = companySpreadsheet.getId();
        return true;
      }
    }
  }
  return false;
}

function insertPurchase(date, stock, cost){
  var companySpreadsheet = SpreadsheetApp.openById(GLOBAL.companyId);
  var companySheet = companySpreadsheet.getActiveSheet();
  var price = cost / stock;
  var totalStock = 0;
  var totalCost = 0;
  
  for(var row = 2; !isEmptyCell(companySheet.getRange(row, 1).getValue()); row++) {
    if(!isEmptyCell(companySheet.getRange(row, 8).getValue())) {
      totalStock = companySheet.getRange(row, 8).getValue();
    }
    if(!isEmptyCell(companySheet.getRange(row, 9).getValue())) {
      totalCost = companySheet.getRange(row, 9).getValue();
    }
  }
  
  totalStock += stock;
  totalCost += cost;
  companySheet.appendRow([GLOBAL.compra, date, stock, 0, cost, price,"", totalStock, totalCost]);
}

function insertDividend(date, stock, cost) {
  var companySpreadsheet = SpreadsheetApp.openById(GLOBAL.companyId);
  var companySheet = companySpreadsheet.getActiveSheet();
  var price = cost / stock;
  
  companySheet.appendRow([GLOBAL.dividendo, date, stock, "", cost, price]);
}

function insertSale(date, stock, cost) {
  var companySpreadsheet = SpreadsheetApp.openById(GLOBAL.companyId);
  var companySheet = companySpreadsheet.getActiveSheet();
  var price = cost / stock;
  var restStock = stock;
  var purchaseCost = 0;
  var totalStock = 0;
  var totalCost = 0;
  var profit = 0;
  var totalProfit = 0;
  
  for(var row = 2; !isEmptyCell(companySheet.getRange(row, 1).getValue()); row++) {
    if(!isEmptyCell(companySheet.getRange(row, 8).getValue())) {
      totalStock = companySheet.getRange(row, 8).getValue();
    }
    if(!isEmptyCell(companySheet.getRange(row, 9).getValue())) {
      totalCost = companySheet.getRange(row, 9).getValue();
    }
    if(!isEmptyCell(companySheet.getRange(row, 12).getValue())) {
      totalProfit = companySheet.getRange(row, 12).getValue();
    }
  }
  
  
  for(var row = 2; restStock>0; row++) {
    var type = companySheet.getRange(row, 1).getValue();
    if(isEmptyCell(type)) {
      return false;
    }
    if(stringEquals(type, GLOBAL.compra)){
      var purchaseStock = companySheet.getRange(row, 3).getValue();
      var saleStock = companySheet.getRange(row, 4).getValue();
      if(purchaseStock > saleStock) {
        var tempStock = purchaseStock - saleStock;
        if(tempStock > restStock) {
          tempStock = restStock;
        }
        companySheet.getRange(row, 4).setValue(saleStock+tempStock);
        purchaseCost += tempStock * companySheet.getRange(row, 6).getValue();
        restStock -= tempStock;
      }
    }   
  }
  
  totalStock -= stock;
  totalCost -= purchaseCost;
  profit = cost-purchaseCost;
  totalProfit += profit;
  companySheet.appendRow([GLOBAL.venta, date, stock, "", cost, price,"",totalStock, totalCost,purchaseCost,profit,totalProfit]);
  return true;
}



function stringEquals(text1, text2) {
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
  
  /*
  if(cell == null) {
    debugger;
    return true;
  }
  */
  
  if(stringEquals(typeof cell, "string")) {
    if(cell.trim().length == 0) {
      return true;
    }
  }
  
  return false;
}


function fromDateToString(date) {
  var texto = "";
  texto = texto.concat(date.getDate());
  texto = texto.concat("/");
  texto = texto.concat(date.getMonth()+1);
  texto = texto.concat("/");
  texto = texto.concat(date.getFullYear());
  
  return texto;
}

function toFloat(text) {
  return parseFloat(text);
}

function toInt(text) {
  return parseInt(text);
}
