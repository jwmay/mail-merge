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
 * An object representing the form responses sheet.
 * 
 * @constructor
 */
var FormResponses = function() {
  // Get the sheet name from the configuration object.
  this.config = Configuration.getCurrent();
  this.sheetName = this.config.sheets.formResponses.name;

  // Get the sheet id.
  var spreadsheet = new BaseSpreadsheet();
  var sheetId = spreadsheet.getSheetId(this.sheetName);

  // Inherit from BaseSheet.
  BaseSheet.call(this, sheetId);
};
inherit_(FormResponses, BaseSheet);


/**
 * Returns true if the form responses sheet is properly initialized with the
 * required headers.
 * 
 * @return {boolean} True if the sheet is initialized, otherwise, false.
 */
FormResponses.prototype.isInitialized = function() {
  // Ensure the required headers are set.
  var header = this.getHeader();
  var requiredHeaders = this.config.sheets.formResponses.headers;
  var lastColumn = this.sheet.getLastColumn();
  var headerSlice = header.slice(lastColumn - 2, lastColumn);
  var headersSet = arraysEqual(requiredHeaders, headerSlice);
  if (headersSet !== true) {
    this.initialize();
  }

  // Ensure a template file has been selected.
  var templateFile = new TemplateFile();
  var file = templateFile.getFile();
  if (file === null) {
    showAlert('Template File', 'Please select a template file');
    return false;
  }

  // Ensure a report folder has been specified.
  var reportsFolder = new ReportsFolder();
  var folder = reportsFolder.folder;
  if (folder === null) {
    showAlert('Reports Folder', 'Please provide a folder to store reports');
    return false;
  }

  // Ensure a report filename has been specified.
  var reports = new Reports();
  var filename = reports.getFilename();
  if (filename === null) {
    showAlert('Report Name', 'Please provide a report name');
    return false;
  }

  return true;
};


/**
 * Initializes the form responses sheet by adding the required headers to the
 * header row of the sheet. The added column numbers are stored in document
 * properties for later use.
 */
FormResponses.prototype.initialize = function() {
  var requiredHeaders = this.config.sheets.formResponses.headers;
  
  // Get the required header column numbers.
  var lastColumn = this.sheet.getLastColumn();
  var reportStatusColumn = lastColumn + 1;
  var pdfLinkColumn = reportStatusColumn + 1;

  // Add the headers and notes and resize the columns.
  this.addHeader(reportStatusColumn, requiredHeaders[0]);
  this.addHeader(pdfLinkColumn, requiredHeaders[1]);
  this.sheet.setColumnWidth(reportStatusColumn, 100);
  this.sheet.setColumnWidth(pdfLinkColumn, 600);
};


/**
 * Returns a Google Range object containing the header cells.
 * 
 * @return {Range} A Google Range object.
 */
FormResponses.prototype.getHeaderRange = function() {
  var headerRange = this.getRow(1, 1, this.sheet.getLastColumn());
  return headerRange;
};


/**
 * Returns an array containing the headers for the form responses sheet.
 * 
 * @return {array} An array of the header values.
 */
FormResponses.prototype.getHeader = function() {
  var headerRange = this.getHeaderRange();
  var header = headerRange.getValues();
  return header[0];
};


/**
 * Returns an array of header keys for use in JavaScript objects. The keys are
 * constructed by replacing spaces with an underscore and removing any special
 * characters from the string.
 * 
 * @return {array} An array of header keys.
 */
FormResponses.prototype.getHeaderKeys = function() {
  var headers = this.getHeader();

  var headerKeys = [];
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    var headerKey = header.replace(/\s+/g, '_').replace(/\W+/g, '');
    headerKeys.push(headerKey);
  }

  return headerKeys;
};


/**
 * Returns an array of header tags for use in the template file. The tags are
 * constructed by adding **...** around the header key.
 * 
 * @return {array} An array of header tags.
 */
FormResponses.prototype.getHeaderTags = function() {
  var headerKeys = this.getHeaderKeys();

  var headerTags = [];
  for (var i = 0; i < headerKeys.length; i++) {
    var headerKey = headerKeys[i];
    var headerTag = '**' + headerKey + '**';
    headerTags.push(headerTag);
  }

  return headerTags;
};


/**
 * Adds the header to the given colNum.
 * 
 * @param {integer} colNum The column number to add to.
 * @param {string} header The header to add.
 */
FormResponses.prototype.addHeader = function(colNum, header) {
  var headerCell = this.sheet.getRange(1, colNum);
  headerCell.setValue(header);
  headerCell.setNote('This column is required by the ' +
          'incident reporter plugin. Do not remove this column.');
};


/**
 * Returns the number of the last column with content in the header.
 * 
 * @return {integer} The last column's position.
 */
FormResponses.prototype.getMaxColumn = function() {
  var header = this.getHeader();
  var headerLength = header.length;
  return headerLength;
};


/**
 * Returns the 'Report Status' column's position.
 * 
 * @return {integer} The Report Status column position.
 */
FormResponses.prototype.getReportStatusColumn = function() {
  var header = this.getHeader();
  var requiredHeaders = this.config.sheets.formResponses.headers;
  var index = header.indexOf(requiredHeaders[0]);
  var column = index + 1;
  return column;
};


/**
 * Returns the 'PDF Link' column's position.
 * 
 * @return {integer} The PDF Link column position.
 */
FormResponses.prototype.getPdfLinkColumn = function() {
  var header = this.getHeader();
  var requiredHeaders = this.config.sheets.formResponses.headers;
  var index = header.indexOf(requiredHeaders[1]);
  var column = index + 1;
  return column;
};