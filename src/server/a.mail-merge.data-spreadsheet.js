/// Copyright 2017 Joseph W. May. All Rights Reserved.
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
 * into the document.
 * 
 * @constructor
 */
var DataSpreadsheet = function() {
  this.storage = new PropertyStore();
};


/**
 * Return the stored id of the data spreadsheet.
 * 
 * @return {string} The spreadsheet id.
 */
DataSpreadsheet.prototype.getId = function() {
  var id = this.storage.getProperty('DATA_SPREADSHEET_ID');
  return id;
};


/**
 * Store the data spreadsheet id.
 * 
 * @param {string} id The spreadsheet id.
 */
DataSpreadsheet.prototype.setId = function(id) {
  this.storage.setProperty('DATA_SPREADSHEET_ID', id);
};


/**
 * Return the data spreadsheet as a Google Spreadsheet object, or null if
 * no spreadsheet id is stored.
 * 
 * @return {Spreadsheet} The spreadsheet as a Google Spreadsheet object, or
 *        null if no spreadsheet id is stored.
 */
DataSpreadsheet.prototype.getSpreadsheet = function() {
  var id = this.getId();
  if (id !== null) {
    var spreadsheet = SpreadsheetApp.openById(id);
    return spreadsheet;
  } else {
    return null;
  }
};


/**
 * Return the name of the data spreadsheet, or null if no spreadsheet
 * id is stored.
 * 
 * @return {string} The spreadsheet name, or null if no spreadsheet
 *        id is stored.
 */
DataSpreadsheet.prototype.getName = function() {
  var spreadsheet = this.getSpreadsheet();
  if (spreadsheet === null) return null;
  var name = spreadsheet.getName();
  return name;
};


/**
 * Return the url of the data spreadsheet, or null if no spreadsheet
 * id is stored.
 * 
 * @return {string} The spreadsheet url, or null if no spreadsheet
 *        id is stored.
 */
DataSpreadsheet.prototype.getUrl = function() {
  var spreadsheet = this.getSpreadsheet();
  if (spreadsheet === null) return null;
  var url = spreadsheet.getUrl();
  return url;
};


/**
 * Return an array of all the sheet names in the spreadsheet, or null if no
 * spreadsheet id is stored.
 * 
 * @return {array} An array containing the sheet names as strings.
 */
DataSpreadsheet.prototype.getSheetNames = function() {
  var spreadsheet = this.getSpreadsheet();
  if (spreadsheet === null) return null;
  var sheets = spreadsheet.getSheets();
  var sheetNames = [];
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    sheetNames.push(sheet.getName());
  }
  return sheetNames;
};