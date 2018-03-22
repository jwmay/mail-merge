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
 * Sidebar initialization and user-response functions.
 */
$(function() {
  initialize();

  // Handle the select spreadsheet file button click.
  $(document).on('click', '#selectSpreadsheet', function() {
    selectSpreadsheet_onClick();
  });

  // Handle the run merge button click.
  $(document).on('click', '#runMerge', function() {
    runMerge_onClick();
  });

  // Update the merge field selector when the sheet selector changes.
  $(document).on('change', '#sheetSelector', function() {
    // Remove the merge field selector display, the run merge display, and any
    // alerts before replacement. Remove the default option in the sheet
    // selector.
    clearDisplay('alerts');
    $('#mergeFieldSelectDisplay').remove();
    $('#runMergeDisplay').remove();
    $('#sheetSelector option.default').remove();
    
    var sheet = $(this).val();
    if (sheet !== '') {
      google.script.run
        .withSuccessHandler(updateDisplay)
        .updateSelectedSheet(sheet);
    }
  });

  // Handle clicks to the insert merge field and rules selectors.
  var dataOptionSelector = '#mergeFieldSelectDisplay li';
  $(document).on('click', dataOptionSelector, function() {
    var data = $(this).html();
    google.script.run
        .withSuccessHandler(insertMergeFieldCallback)
        .insertMergeField(data);
  });
});


/**
 * Runs sidebar initialization functions to retrieve and display the primary
 * UI components based on selected and stored data.
 */
function initialize() {
  // Display the data spreadsheet and file selector button.
  google.script.run
    .withSuccessHandler(updateDisplay)
    .getSidebarDisplay();
}


/**
 * Handle the 'Select file' button click response. This handles the process of
 * allowing the user to select a data spreadsheet file, storing the file's id, 
 * and updating the sidebar display with the file's name and a link to the file.
 */
function selectSpreadsheet_onClick() {
  showLoading();
  google.script.run
    .withSuccessHandler(hideLoading)
    .showSpreadsheetPicker();
}


/**
 * Handle the 'Run merge' button click response.
 */
function runMerge_onClick() {
  showLoading();
  google.script.run
    .withSuccessHandler(updateDisplay)
    .runMerge();
}


/**
 * Handle the callback after a merge field has been inserted into the document
 * by displaying an error message if the insert failed.
 * 
 * @param {object} display A display object containing an error message,
 *        otherwise, null if the insert was successful.
 */
function insertMergeFieldCallback(display) {
  if (display !== undefined) {
    updateDisplay(display);
  }
}