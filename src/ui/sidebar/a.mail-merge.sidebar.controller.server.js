// Copyright 2018 Joseph W. May. All Rights Reserved.
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
 * Returns an array of DisplayObjects containing the user-interface to display
 * on the sidebar.
 * 
 * The interface display will depend on which options have been selected and
 * saved. The interface loads in a 'bottom-up' approach: the spreasheet selector
 * will be displayed first and a stack will generate with each new selection.
 * Once a spreadsheet has been selected, the sheet selector will display. Once a
 * sheet has been selected, the insert merge field and merge control buttons
 * will be displayed.
 * 
 * @returns {DisplayObjects[]} An array of DisplayObject instances that make up
 *    the sidebar user interface.
 */
function getSidebarDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var sheetNames = spreadsheet.getSheetNames();
  var sheet = spreadsheet.getSheetName();
  var headers = spreadsheet.getSheetHeader();

  // Construct the individual DisplayObjects for each component
  // of the sidebar display and store them in an array
  var displayObjects = [];
  displayObjects.push(getSpreadsheetDisplay());
  
  // Get the sheet selector display only if there are sheet names available
  if (sheetNames !== null) {
    displayObjects.push(getSheetSelectDisplay());
  }

  // Get the merge field selector only if a sheet is selected
  if (sheet !== null) {
    displayObjects.push(getMergeFieldDisplay());
  }

  // Get the merge button display only if merge fields are displayed, which is
  // based on their being a header row in the selected sheet
  if (headers !== null) {
    displayObjects.push(getRunMergeDisplay());
  }
  return displayObjects;
}


/**
 * Stores the selected sheet name and returns the merge field select display
 * and run merge display if the sheet has headers, otherwise, displays an error.
 * 
 * @param {string} name The sheet name.
 * @returns {DisplayObjects[]} An array of DisplayObject instances for the merge
 *    field selector and merge control buttons.
 */
function updateSelectedSheet(name) {
  var spreadsheet = new DataSpreadsheet();
  spreadsheet.setSheetName(name);

  // Get the merge field selector, or error display if there are
  // no headers in selected sheet
  var displayObjects = [];
  displayObjects.push(getMergeFieldDisplay());

  // Get the run merge display only if there are headers in the selected sheet
  var headers = spreadsheet.getSheetHeader();
  if (headers !== null) {
    displayObjects.push(getRunMergeDisplay());
  }
  return displayObjects;
}


/**
 * Returns a display object containing the data spreadsheet file selector.
 * 
 * @returns {DisplayObject} A DislpayObject instance for the spreadsheet file
 *    selector.
 */
function getSpreadsheetDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var id = spreadsheet.getId();

  // Get the link to display
  var linkDisplay = 'No file selected';
  if (id !== null) {
    var name = spreadsheet.getName();
    var url = spreadsheet.getUrl();
    linkDisplay = Utilities.formatString('<a href="%s" target="_blank">%s</a>',
        url, name);
  }

  // Construct the display object with filename and url.
  var content = '' +
      '<h4>Spreadsheet</h4>' +
      '<div class="file">' +
        '<i class="fa fa-table" aria-hidden="true"></i> ' +
        '<span id="dataSpreadsheet">' + linkDisplay + '</span>' +
      '</div>' +
      '<div class="btn-bar">' +
        '<input type="button" value="Select file" id="selectSpreadsheet">' +
      '</div>';
  var display = getDisplayObject('card', content, 'spreadsheetDisplay');
  return display;
}


/**
 * Returns a DisplayObject instance for selecting a sheet from the data
 * spreadsheet.
 * 
 * @returns {DisplayObject} A DisplayObject instance for the sheet selector.
 */
function getSheetSelectDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var selected = spreadsheet.getSheetName();
  var sheetNames = spreadsheet.getSheetNames();  
  var select = makeSelect(sheetNames, 'Select a sheet...', selected,
      'sheetSelector');
  var content = '' +
      '<h4>Sheet</h4>' +
      '<div>' +
        select +
      '</div>';
  var display = getDisplayObject('card', content, 'sheetSelectDisplay');
  return display;
}


/**
 * Returns a DisplayObject instance containing the merge field selector or an
 * error message if no headers are found in the selected sheet.
 * 
 * @returns {DisplayObject} A DisplayObject instance for the merge field
 *    selector.
 */
function getMergeFieldDisplay() {
  var spreadsheet = new DataSpreadsheet();
  var headers = spreadsheet.getSheetHeader();

  if (headers !== null) {
    // Construct the list of merge field items
    var list = [];
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      var item = '<li>' + header + '</li>';
      list.push(item);
    }
    var listDisplay = list.join('\n');
    
    // Construct the merge field display
    var content = '' +
        '<h4>Insert merge field</h4>' +
        '<div class="selector">' +
          '<ul>' +
            listDisplay +
          '</ul>' +
        '</div>';

    var display = getDisplayObject('card', content, 'mergeFieldSelectDisplay');
    return display;
  } else {
    var message = 'No headers found in the selected sheet.';
    var error = getDisplayObject('alert-error', message);
    return error;
  }
}


/**
 * Returns a DisplayObject instance containing the merge control buttons.
 * 
 * @returns {DisplayObject} A DisplayObject instance for the merge control
 *    buttons.
 */
function getRunMergeDisplay() {
  var content = '' +
      '<h4>Merge</h4>' +
      '<div class="btn-bar">' +
        '<input type="button" class="btn action" value="Run merge" ' +
          'id="runMerge">' +
        '<input type="button" class="btn" value="Merge options" ' +
          'id="mergeOptions">' +
      '</div>';
  var dislpay = getDisplayObject('card', content, 'runMergeDisplay');
  return dislpay;
}


/**
 * Displays an HTML Service dialog in Google Sheets for setting the
 * merge options.
 */
function showMergeOptions() {
  showDialog('a.mail-merge.options.view', 500, 390, 'Merge options');
}


/**
 * Displays an HTML Service dialog in Google Sheets for the Google Picker API
 * thas is used to select the source spreadsheet.
 */
function showSpreadsheetPicker() {
  showDialog('a.mail-merge.spreadsheet-picker.view', 900, 550,
      'Select a spreadsheet');
}