/*
Copyright 2017-2020, Dmitry Klimenko, All right reserved

MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

// UI

function onOpen() {

  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('CMC')
      .addItem('Update', 'updateFromCMC')
      .addToUi();
}

// DATA

function updateFromCMC() {
  
  const proKey = "00000000-0000-0000-0000-000000000000";
  const URL = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest?start=1&limit=5000&convert=USD&CMC_PRO_API_KEY=" + proKey; 
  
  const data = JSON.parse(UrlFetchApp.fetch(URL))["data"]; 
  
  const keys = ["symbol", "id", "name", "cmc_rank", "quote.USD.price", "quote.USD.percent_change_1h", "quote.USD.percent_change_24h", "quote.USD.percent_change_7d"];
  const numberKeys = ["quote.USD.price", "quote.USD.percent_change_1h", "quote.USD.percent_change_24h", "quote.USD.percent_change_7d"];
  
  const values = transformToTable(data, keys, numberKeys);
  
  var sheet = createOrClearSheet("CMC")

  sheet.getRange(1, 1, 1, keys.length).setValues([keys]);
  
  var range = sheet.getRange(2, 1, values.length, keys.length)
  range.setValues(values);
  setNameOfRange("CMC", range);
}

// DATA-HELPERS

function transformToTable(data, keys, numberKeys) {
  var table = new Array();
  for (var d = 0; d < data.length; d++) {
    var dataRow = data[d];
    var row = new Array();
    table.push(row);
    for (var k = 0; k < keys.length; k++) {
      var keyParts = keys[k].split(".");
      var value = dataRow[keyParts[0]];
      for (var kp = 1; kp < keyParts.length; kp++) {
        value = value[keyParts[kp]];
      }
      if (numberKeys.indexOf(keys[k]) > -1) {
        value = Number(value)
      }
      row.push(value);
    }
  }
  return table;
}  

// SHEET-HELPERS

function createOrClearSheet(name) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(name);
  if (sheet == null) {
    sheet = spreadSheet.insertSheet(name);
  }
  sheet.clearContents();
  return sheet;
}  

function setNameOfRange(name, range) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  namedRange = spreadSheet.getRangeByName(name);
  if (namedRange == null) {
    spreadSheet.setNamedRange(name, range);
  } else {
    namedRanges = spreadSheet.getNamedRanges();
    for (var i = 0; i < namedRanges.length; i++) {
      if (namedRanges[i].getName() == name) {
        namedRanges[i].setRange(range);
        break;
      }
    }
  }
}

