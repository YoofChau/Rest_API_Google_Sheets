// Function of pasting array values into desired rows and columns.
function addArrayToSheetColumn(sheet, column, row, values) {
  const defined_row = row.toString().concat(":")
  const range = [column, defined_row, column, values.length + row - 1].join("");
  const fn = function(v) {
    return [ v ];
  };
  sheet.getRange(range).setValues(values.map(fn));
}

// Main Function
function myFunction() {
  
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Main");
  var noteSheet = ss.getSheetByName("Note")
  
  // Clear sheets contents before run.
  mainSheet.clearContents();

  // Get input currencies from "Note" tab 
  var num_currency = noteSheet.getRange("A11:A").getValues().flat();
  var lastrow_currency = num_currency.filter(String).length;
  
  // Input currencies to list for API call
  var currency_list = num_currency.slice(0,lastrow_currency);
  var currency = currency_list[0]
  for (let i = 1; i < currency_list.length; i++) {

    currency = currency.concat(",",currency_list[i]);

    }
  
  // Forming base url for API Call
  var symbols_url = '&symbols='
  var access_key = '' //own API key
  var url_string = "http://data.fixer.io/api/latest?access_key=";
  var url_string_final = url_string.concat(access_key, symbols_url, currency); 
  
  // Get response from url and convert it into json format
  var response = UrlFetchApp.fetch(url_string_final);
  var json = response.getContentText();
  var data = JSON.parse(json);
  
  // Get desired values from Dictionary. In this case, currency rates and currency symbols
  var symbols_list_source = Object.keys(data.rates);
  var currency_source = Object.values(data.rates);

  // Paste array value into desired rows and columns
  mainSheet.getRange(1, 1).setValue("Base Currency");
  mainSheet.getRange(1, 2).setValue(data.base);
  
  mainSheet.getRange(3, 1).setValue("Currency");
  mainSheet.getRange(3, 2).setValue("Rates");  
  addArrayToSheetColumn(mainSheet, 'A', 4, symbols_list_source)
  addArrayToSheetColumn(mainSheet, 'B', 4, currency_source)

  // Message alert
  SpreadsheetApp.getUi().alert("Conversion Rate Updated!");
}
