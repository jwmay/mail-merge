// Copyright 2017 Joseph W. May. All Rights Reserved.
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
 * Base class that assists in working with a Spreadsheet instance in a
 * standalone project that may be dynamically bound to a Spreadsheet (as an
 * add-on), but also may run separate (testing, Execution API, etc.).
 * 
 * @constructor
 * @param {grade-importer.json.Configuration} config An optional valid
 *     configuration object instance.
 */
var BaseSpreadsheet = function(config) {
  this.config = config !== undefined ? config : Configuration.getCurrent();
};


/**
 * Returns a Spreadsheet instance. Instance is retrieved by ID if an ID
 * has been set by test code in the configuration.
 * 
 * @return {object} A Spreadsheet instance.
 */
BaseSpreadsheet.prototype.getSpreadsheet = function() {
  var loadLocal = ((this.config.debugSpreadsheetId !== null) &&
      (this.config.debugSpreadsheetId !== undefined) &&
      (this.config.debugSpreadsheetId !== ''));
  if (loadLocal) {
    return SpreadsheetApp.openById(this.config.debugSpreadsheetId);
  } else {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
};


/**
 * Returns the Spreadsheet name.
 * 
 * @return {string} The name of the current Spreadsheet.
 */
BaseSpreadsheet.prototype.getSpreadsheetName = function() {
  var spreadsheet = this.getSpreadsheet();
  var spreadsheetName = spreadsheet.getName();
  return spreadsheetName;
};


/**
 * Returns an array containing the name of every sheet in the spreadsheet
 * as a string.
 * 
 * @return {array} An array of sheet names as strings.
 */
BaseSpreadsheet.prototype.getSheetNames = function() {
  var sheets = this.getSpreadsheet().getSheets();
  var sheetNames = [];
  for (var i = 0; i < sheets.length; i++) {
    sheetNames.push(sheets[i].getName());
  }
  return sheetNames;
};


/**
 * Returns the sheet id of the sheet with the given name.
 * 
 * @param {string} name The name of the sheet whose id will be returned.
 * @return {integer} The integer representing the sheet id of the given sheet. 
 */
BaseSpreadsheet.prototype.getSheetId = function(name) {
  var spreadsheet = this.getSpreadsheet();
  var sheet = spreadsheet.getSheetByName(name);
  var sheetId = sheet.getSheetId();
  return sheetId;
};


/**
 * Returns true if the given sheet name is found in the spreadsheet.
 * 
 * @param {string} name The sheet name.
 * @return {boolean} True if the sheet name is found, otherwise, false.
 */
BaseSpreadsheet.prototype.hasSheet = function(name) {
  // Google Sheets sheet names are not case sensistive. The comparisons here
  // are done in lower case to ensure no duplicates exist in the spreadsheet.
  var lowercaseName = name.toLowerCase();
  var sheetNamesArray = this.getSheetNames();  
  var sheetNames = sheetNamesArray.toLowerCase();
  var index = sheetNames.indexOf(lowercaseName);
  var found = index > -1 ? true : false;
  return found;
};


/**
 * Returns true if the current user is the owner of the current spreadsheet.
 * 
 * @return {boolean} True if the current user is the owner of the Drive file.
 */
BaseSpreadsheet.prototype.currentUserIsOwner = function() {
  var ownerEmail = this.getSpreadsheet().getOwner().getEmail();
  return (ownerEmail === Session.getEffectiveUser().getEmail());
};


/**
 * Returns a Google Sheet object with the given sheet id. If a sheet with the
 * given id is not found, return null.
 * 
 * @param {number} sheetId A Sheet id.
 * @return {object} A Sheet instance, or null if not found.
 */
BaseSpreadsheet.prototype.getSheetById = function(sheetId) {
  if (sheetId === undefined || sheetId === null) {
    return null;
  }
  var sheets = this.getSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === sheetId) {
      return sheets[i];
    }
  }
  return null;
};