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
 * Returns and HTML-formatted string containing the user-interface to display
 * on the sidebar.
 * 
 * @return {string} An HTML-formatted string to display the user-interface.
 */
function getSidebarDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var sheet = spreadsheet.getSheetName();

  var display = [];
  display.unshift(getSpreadsheetDisplay());
  display.unshift(getSheetSelectDisplay());
  if (sheet !== null) display.unshift(getHeaderSelectDisplay());
  
  var html = display.join('\n');
  return html;
}


/**
 * Stores the selected sheet name and returns the header select display.
 * 
 * @param {string} name The sheet name.
 */
function updateSelectedSheet(name) {
  var spreadsheet = new DataSpreadsheet();
  spreadsheet.setSheetName(name);
  var display = getHeaderSelectDisplay();
  return display;
}


/**
 * Returns an HTML-formatted string containing the spreadsheet file selector.
 * 
 * @return {string} An HTML-formatted string to display the spreadsheet
 *        file selector.
 */
function getSpreadsheetDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var id = spreadsheet.getId();

  // Get the link to display.
  var linkDisplay = 'No file selected';
  if (id !== null) {
    var name = spreadsheet.getName();
    var url = spreadsheet.getUrl();
    linkDisplay = Utilities.formatString('<a href="%s" target="_blank">%s</a>',
        url, name);
  }

  // Construct the display with filename and url.
  var display = '<div class="card" id="spreadsheetDisplay">' +
      '<h4>Spreadsheet</h4>' +
      '<div class="file">' +
        '<i class="fa fa-file" aria-hidden="true"></i> ' +
        '<span id="dataSpreadsheet">' + linkDisplay + '</span>' +
      '</div>' +
      '<div class="btn-bar">' +
        '<input type="button" value="Select file" onclick="selectSpreadsheet_onclick();">' +
      '</div>' +
    '</div>';
  return display;
}


/**
 * Returns an HTML-formatted string containing the sheet selector.
 * 
 * @return {string} An HTML-formatted string to display the sheet selector.
 */
function getSheetSelectDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var selected = spreadsheet.getSheetName();
  var sheetNames = spreadsheet.getSheetNames();
  if (sheetNames !== null) {
    var select = makeSelect(sheetNames, 'Select a sheet...', selected,
            'sheetSelector');
    var display = '<div class="card" id="sheetSelectDisplay">' +
        '<h4>Sheet</h4>' +
        '<div>' +
          select +
        '</div>' +
      '</div>';
    return display;
  }
}


/**
 * Returns an HTML-formatted string containing the data selector.
 * 
 * @return {string} An HTML-formatted string to display the data selector.
 */
function getHeaderSelectDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var sheetName = spreadsheet.getSheetName();
  var headers = spreadsheet.getSheetHeader(sheetName);

  if (headers !== null) {
    // Construct the list items.
    var list = [];
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      var item = '<li>' + header + '</li>';
      list.push(item);
    }
    var listDisplay = list.join('\n');
    
    // Construct the HTML display.
    var display = '<div class="card" id="headerSelectDisplay">' +
        '<h4>Select data</h4>' +
        '<div class="selector">' +
          '<ul>' +
            listDisplay +
          '</ul>' +
        '</div>' +
      '</div>';
    return display;
  } else {
    var error = '<div class="card alert alert-error" id="headerSelectDisplay">' +
            'No headers found in the selected sheet.' +
          '</div>';
    return error;
  }
}


/**
 * Displays an HTML Service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API. Used to select the report
 * template file.
 */
function showSpreadsheetPicker() {
  showDialog('a.mail-merge.data-spreadsheet-picker.view', 900, 550,
          'Select a spreadsheet');
}