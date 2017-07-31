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
 * @return {displayObjects[]} An array of display objects to construct
 *        the user-interface.
 */
function getSidebarDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var sheetNames = spreadsheet.getSheetNames();
  var sheet = spreadsheet.getSheetName();
  var headers = spreadsheet.getSheetHeader();

  // Construct the individual displayObjects for each component
  // of the sidebar display and store them in an array.
  var displayObjects = [];
  displayObjects.push(getSpreadsheetDisplay());
  
  // Get the sheet selector display only if there are sheet names available.
  if (sheetNames !== null) {
    displayObjects.push(getSheetSelectDisplay());
  }

  // Get the merge field selector only if a sheet is selected.
  if (sheet !== null) {
    displayObjects.push(getMergeFieldDisplay());
  }

  // Get the merge button display only if merge fields are displayed.
  if (headers !== null) {
    displayObjects.push(getRunMergeDisplay());
  }
  return displayObjects;
}


/**
 * Stores the selected sheet name and returns the header select display and
 * run merge display if the sheet has headers, otherwise, displays an error.
 * 
 * @param {string} name The sheet name.
 */
function updateSelectedSheet(name) {
  var spreadsheet = new DataSpreadsheet();
  spreadsheet.setSheetName(name);

  // Get the merge field selector, or error display if there are
  // no headers in selected sheet.
  var displayObjects = [];
  displayObjects.push(getMergeFieldDisplay());

  // Get the run merge display only if there are headers in the selected sheet.
  var headers = spreadsheet.getSheetHeader();
  if (headers !== null) {
    displayObjects.push(getRunMergeDisplay());
  }
  return displayObjects;
}


/**
 * Returns a display object containing the spreadsheet file selector.
 * 
 * @return {displayObject} A dislpay object for the spreadsheet file selector.
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

  // Construct the display object with filename and url.
  var content = '<h4>Spreadsheet</h4>' +
      '<div class="file">' +
        '<i class="fa fa-file" aria-hidden="true"></i> ' +
        '<span id="dataSpreadsheet">' + linkDisplay + '</span>' +
      '</div>' +
      '<div class="btn-bar">' +
        '<input type="button" value="Select file" id="selectSpreadsheet">' +
      '</div>';
  var display = getDisplayObject('card', content, 'spreadsheetDisplay');
  return display;
}


/**
 * Returns a display object containing the sheet selector.
 * 
 * @return {displayObject} A display object for the sheet selector.
 */
function getSheetSelectDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var selected = spreadsheet.getSheetName();
  var sheetNames = spreadsheet.getSheetNames();  
  var select = makeSelect(sheetNames, 'Select a sheet...', selected,
          'sheetSelector');
  var content = '<h4>Sheet</h4>' +
      '<div>' +
        select +
      '</div>';
  var display = getDisplayObject('card', content, 'sheetSelectDisplay');
  return display;
}


/**
 * Returns a display object containing the merge field selector. If no headers
 * are found in the selected sheet, an error display object is returned.
 * 
 * @return {displayObject} A display object for the merge field selector.
 */
function getMergeFieldDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var headers = spreadsheet.getSheetHeader();

  if (headers !== null) {
    // Construct the list of merge field items.
    var list = [];
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      var item = '<li>' + header + '</li>';
      list.push(item);
    }
    var listDisplay = list.join('\n');
    
    // Construct the merge field display.
    var content = '<h4>Insert merge field</h4>' +
        '<div class="selector">' +
          '<ul>' +
            listDisplay +
          '</ul>' +
        '</div>';

    // Construct the rules selector display.
    content += '<h4>Rules</h4>' +
        '<div class="selector">' +
          '<ul>' +
            '<li>Next record</li>' +
          '</ul>' +
        '</div>';

    var display = getDisplayObject('card', content, 'mergeFieldSelectDisplay');
    return display;
  } else {
    var errorContent = 'No headers found in the selected sheet.';
    var errorDisplay = getDisplayObject('alert-error', errorContent);
    return errorDisplay;
  }
}


/**
 * Returns a display object containing the run merge button.
 * 
 * @return {displayObject} A display object for the run merge button.
 */
function getRunMergeDisplay() {
  var content = '<input type="button" class="btn action" value="Run merge" ' +
          'id="runMerge">';
  var dislpay = getDisplayObject('card', content, 'runMergeDisplay');
  return dislpay;
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