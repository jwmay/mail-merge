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
 * Runs sidebar initialization and handles all sidebar user inputs.
 */
$(function() {
  initialize();

  // Handles the select spreadsheet file button click.
  $(document).on('click', '#selectSpreadsheet', function() {
    selectSpreadsheet_onClick();
  });

  // Handles the run merge button click.
  $(document).on('click', '#runMerge', function() {
    runMerge_onClick();
  });

  // Handles the merge options button click.
  $(document).on('click', '#mergeOptions', function() {
    mergeOptions_onClick();
  });

  // Updates the merge field selector when the sheet selector changes.
  $(document).on('change', '#sheetSelector', function() {
    // Removes the merge field selector, the merge controls, and any alerts
    // before replacement. Removes the default option in the sheet selector.
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

  // Handles clicks to the insert merge field selector.
  var mergeFieldSelector = '#mergeFieldSelectDisplay li';
  $(document).on('click', mergeFieldSelector, function() {
    var field = $(this).html();
    google.script.run
      .withSuccessHandler(insertMergeFieldCallback)
      .insertMergeField(field);
  });
});


/**
 * Runs sidebar initialization functions to retrieve and display the primary
 * user interface components for the sidebar.
 */
function initialize() {
  google.script.run
    .withSuccessHandler(updateDisplay)
    .getSidebarDisplay();
}


/**
 * Handles the 'Select file' button click response. This handles the process of
 * allowing the user to select a data spreadsheet file, storing the id of the
 * file, and updating the sidebar display with a hyperlinked name for the file.
 */
function selectSpreadsheet_onClick() {
  showLoading();
  google.script.run
    .withSuccessHandler(hideLoading)
    .showSpreadsheetPicker();
}


/**
 * Handles the 'Run merge' button click response.
 */
function runMerge_onClick() {
  showLoading();
  clearDisplay('alerts');
  google.script.run
    .withSuccessHandler(updateDisplay)
    .runMerge();
}


/**
 * Handles the 'Merge options' button click response.
 */
function mergeOptions_onClick() {
  google.script.run.showMergeOptions();
}


/**
 * Handles the callback after a merge field has been inserted into the document
 * by displaying an error message only if the insert failed.
 * 
 * @param {DisplayObject} display A DisplayObject instance containing an error
 *    message, or null if no display should be show.
 */
function insertMergeFieldCallback(display) {
  if (typeof display === 'object') {
    updateDisplay(display);
  }
}