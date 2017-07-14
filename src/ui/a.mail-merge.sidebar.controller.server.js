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
 * 
 */
function getDataSpreadsheetDisplay() {
  var dataSpreadsheet = new DataSpreadsheet();
  // dataSpreadsheet.setId('10E05jyBlG1cLj7OAbCL_U6YVYWheDLpeG5Mf83Bkvec');
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
        linkDisplay +
      '</div>' +
      '<div class="btn-bar">' +
        '<input type="button" value="Select file" onclick="selectFile_onclick();">' +
      '</div>' +
    '</div>';
  return display;
}


/**
 * Displays the file selector and returns an HTML-formatted string containing
 * the link to the currently selected template file.
 * 
 * @returns {string} An HTML-formatted string.
 */
function setDataSpreadsheetFile() {
  showFilePicker();
  return getDataSpreadsheetDisplay();
}


/**
 * Displays an HTML Service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API. Used to select the report
 * template file.
 */
function showFilePicker() {
  showDialog('a.mail-merge.file-picker.view', 900, 550,
          'Select a spreadsheet');
}