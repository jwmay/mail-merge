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
 * Sidebar initialization and user-response functions to load the sidebar.
 */
$(function() {
  initializeSidebar();
});


/**
 * Runs sidebar initialization functions to retrieve and display stored data.
 */
function initializeSidebar() {
  // Display the data spreadsheet and file selector button.
  google.script.run
    .withSuccessHandler(updateDisplay)
    .getSidebarDisplay();
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