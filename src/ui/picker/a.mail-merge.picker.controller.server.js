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
 * Process the selected file from the Google Picker API. Perform mime-type
 * validation and store the id of the selected file.
 * 
 * @param {array} files An array of JSON objects returned by Google Picker
 *     representing the selected file.
 */
function loadSpreadsheetFile(files) {
  var file = files[0];
  if (file.mimeType === MimeType.GOOGLE_SHEETS) {
    var dataSpreadsheet = new DataSpreadsheet();
    dataSpreadsheet.setId(file.id);
    var success_message = 'Spreadsheet file successfully updated.';
    // Refresh the sidebar to display the newly-selected file.
    onShowSidebar();
    return success_message;
  }
}