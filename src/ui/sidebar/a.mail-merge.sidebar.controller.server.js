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


function getSidebarDisplay() {
  var display = [];
  display.unshift(getSpreadsheetDisplay());
  display.unshift(getSheetSelectDisplay());
  var html = display.join('\n');
  return html;
}


/**
 * 
 */
function getSpreadsheetDisplay() {
  var dataSpreadsheet = new DataSpreadsheet();
  var id = dataSpreadsheet.getId();

  // Get the link to display.
  var linkDisplay = 'No file selected';
  if (id !== null) {
    var name = dataSpreadsheet.getName();
    var url = dataSpreadsheet.getUrl();
    linkDisplay = Utilities.formatString('<a href="%s" target="_blank">%s</a>',
        url, name);
  }

  // Construct the display with filename and url.
  var display = '<div class="card">' +
      '<h4>Data Spreadsheet</h4>' +
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
 * 
 * @return {string} An HTML-formatted string to display in the sidebar.
 */
function getSheetSelectDisplay() {
  var dataSpreadsheet = new DataSpreadsheet();
  var sheetNames = dataSpreadsheet.getSheetNames();
  if (sheetNames !== null) {
    var select = makeSelect(sheetNames);
    var display = '<div class="card">' +
        '<h4>Select a sheet</h4>' +
        '<div>' +
          select +
        '</div>' +
      '</div>';
    return display;
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