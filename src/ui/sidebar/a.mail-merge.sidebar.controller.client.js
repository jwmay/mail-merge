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
 * Sidebar initialization and user-response functions.
 */
$(function() {
  initialize();

  // Handle changes to the sheet selector by replacing the data selector.
  $(document).on('change', '#sheetSelector', function() {
    // Remove the data selector display before replacement and the default
    // option in the sheet selector.
    $('#headerSelectDisplay').remove();
    $('#sheetSelector option.default').remove();
    
    var sheet = $(this).val();
    if (sheet !== '') {
      google.script.run
        .withSuccessHandler(updateDisplay)
        .updateSelectedSheet(sheet);
    }
  });

  // Handle clicks to the data selector.
  var dataOptionSelector = '#headerSelectDisplay li';
  $(document).on('click', dataOptionSelector, function() {
    var data = $(this).html();
    google.script.run
        .withSuccessHandler()
        .insertVariable(data);
  });
});


/**
 * Runs sidebar initialization functions to retrieve and display stored data.
 */
function initialize() {
  // Display the data spreadsheet and file selector button.
  google.script.run
    .withSuccessHandler(showSidebarDisplay)
    .getSidebarDisplay();
}


/**
 * Updates the sidebar display with an array of display objects containing
 * the add-on's primary user interface.
 * 
 * @param {displayObjects[]} displayObjects An array of display objects.
 */
function showSidebarDisplay(displayObjects) {
  displayObjects.forEach(function(display) {
    updateDisplay(display);
  });
}


/**
 * Handle the 'Select File' button click response. This handles the process of
 * allowing the user to select a data spreadsheet file, storing the file's id, 
 * and updating the sidebar display with the file's name and a link to the file.
 */
function selectSpreadsheet_onclick() {
  showLoading();
  google.script.run
    .withSuccessHandler(hideLoading)
    .showSpreadsheetPicker();
}