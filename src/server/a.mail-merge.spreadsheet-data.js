/// Copyright 2018 Joseph W. May. All Rights Reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.


/**
 * Base class for the data spreadsheet that contains the data to merge into
 * the document.
 * 
 * @constructor
 */
var DataSpreadsheet = function() {
  this.storage = new PropertyStore();
  this.data_sheet_name_key = 'DATA_SHEET_NAME';
  this.data_spreadsheet_id_key = 'DATA_SPREADSHEET_ID';
};


/**
 * Returns the stored id of the data spreadsheet.
 * 
 * @returns {string} The spreadsheet id.
 */
DataSpreadsheet.prototype.getId = function() {
  var id = this.storage.getProperty(this.data_spreadsheet_id_key);
  return id;
};


/**
 * Returns the title of the data spreadsheet, or null if no spreadsheet
 * id is stored.
 * 
 * @returns {string} The spreadsheet title, or null if no spreadsheet
 *    id is stored.
 */
DataSpreadsheet.prototype.getName = function() {
  var spreadsheet = this.getSpreadsheet();
  var name = (spreadsheet !== null ? spreadsheet.getName() : null);
  return name;
};


/**
 * Returns the number of records (rows minus the header) in the selected sheet.
 * 
 * A record is considered a row in the selected sheet, excluding the first row,
 * which is considered the header row.
 * 
 * @returns {integer} The number of records in the sheet.
 */
DataSpreadsheet.prototype.getRecordCount = function() {
  // Get the sheet, or return null if there is no selected sheet or spreadsheet.
  var sheet = this.getSheet();
  if (sheet === null) return null;

  // Return the number of rows with content, minus one for the header row.
  var numRows = sheet.getLastRow();
  var recordCount = numRows - 1;
  return recordCount;
};


/**
 * Returns the records in the selected sheet.
 * 
 * Returns a two-dimensional array of values, indexed by row, then by column.
 * The header row is included as the first row of values.
 * 
 * @returns {array[][]} A two-dimensional array of values.
 */
DataSpreadsheet.prototype.getRecords = function() {
  // Get the sheet, or return null if there is no selected sheet or spreadsheet.
  var sheet = this.getSheet();
  if (sheet === null) return null;

  // Return the records, which include the header row.
  var recordsRange = sheet.getDataRange();
  var records = recordsRange.getValues();
  return records;
};


/**
 * Returns the selected sheet as a Google Sheet object, or null if no sheet or
 * spreadsheet is selected.
 * 
 * @returns {Sheet} The selected sheet, or null if no sheet is selected.
 */
DataSpreadsheet.prototype.getSheet = function() {
  // Return null if there is no selected spreadsheet.
  var spreadsheet = this.getSpreadsheet();
  if (spreadsheet === null) return null;

  // Return the sheet, or null if there is no selected sheet.
  var sheetName = this.getSheetName();
  var sheet = (sheetName !== null ? spreadsheet.getSheetByName(sheetName) : null);
  return sheet;
};


/**
 * Returns an array of header values for the selected sheet.
 * 
 * If there is no selected sheet, the sheet has no data, or there is no header,
 * returns null. The header is considered the first row of the selected sheet.
 * 
 * @returns {array} The header names as strings, or null if there is no header.
 */
DataSpreadsheet.prototype.getSheetHeader = function() {
  // Get the sheet, or return null if there is no selected sheet or spreadsheet.
  var sheet = this.getSheet();
  if (sheet === null) return null;
  
  // Return null if there is no data in the sheet, otherwise, get the header.
  var maxCols = sheet.getLastColumn();
  if (maxCols === 0) return null;
  var range = sheet.getRange(1, 1, 1, maxCols);
  var header = range.getValues()[0];

  // Return null if there are no headers.
  header = header.removeEmpty();
  if (header.length < 1) return null;
  return header;
};


/**
 * Returns the stored sheet name.
 * 
 * @returns {string} The sheet name.
 */
DataSpreadsheet.prototype.getSheetName = function() {
  var name = this.storage.getProperty(this.data_sheet_name_key);
  return name;
};


/**
 * Returns an array of all the sheet names in the spreadsheet, or null if no
 * spreadsheet id is stored.
 * 
 * @returns {array} An array containing the sheet names as strings.
 */
DataSpreadsheet.prototype.getSheetNames = function() {
  var spreadsheet = this.getSpreadsheet();
  if (spreadsheet !== null) {
    var sheets = spreadsheet.getSheets();
    var sheetNames = [];
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      sheetNames.push(sheet.getName());
    }
    return sheetNames;
  } else {
    return null;
  }
};


/**
 * Returns the data spreadsheet as a Google Spreadsheet object, or null if
 * no spreadsheet id is stored.
 * 
 * @returns {Spreadsheet} The spreadsheet as a Google Spreadsheet object, or
 *    null if no spreadsheet id is stored.
 */
DataSpreadsheet.prototype.getSpreadsheet = function() {
  var id = this.getId();
  var spreadsheet = (id !== null ? SpreadsheetApp.openById(id) : null);
  return spreadsheet;
};


/**
 * Returns the url of the data spreadsheet, or null if no spreadsheet
 * id is stored.
 * 
 * @returns {string} The spreadsheet url, or null if no spreadsheet
 *    id is stored.
 */
DataSpreadsheet.prototype.getUrl = function() {
  var spreadsheet = this.getSpreadsheet();
  var url = (spreadsheet !== null ? spreadsheet.getUrl() : null);
  return url;
};


/**
 * Stores the data spreadsheet id.
 * 
 * @param {string} id The spreadsheet id.
 */
DataSpreadsheet.prototype.setId = function(id) {
  this.storage.setProperty(this.data_spreadsheet_id_key, id);
};


/**
 * Stores the data sheet name.
 * 
 * @param {string} name The sheet name.
 */
DataSpreadsheet.prototype.setSheetName = function(name) {
  this.storage.setProperty(this.data_sheet_name_key, name);
};