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


var DEVELOPER_KEY = 'AIzaSyCwQxkZsb6_OUI8LRHbEgR1UzxswOodGvM';
var DIALOG_DIMENSIONS = {width: 900, height: 550};
var pickerApiLoaded = false;


/**
 * Loads the Google Picker API for file selection.
 */
function onFilePickerApiLoad() {
  gapi.load('picker', {
    'callback': function() {
      pickerApiLoaded = true;
    }
  });
  google.script.run
    .withSuccessHandler(createFilePicker)
    .withFailureHandler(showError)
    .getOAuthToken();
}


/**
 * Loads the Google Picker API for folder selection.
 */
function onFolderPickerApiLoad() {
  gapi.load('picker', {
    'callback': function() {
      pickerApiLoaded = true;
    }
  });
  google.script.run
    .withSuccessHandler(createFolderPicker)
    .withFailureHandler(showError)
    .getOAuthToken();
}


/**
 * Creates a Picker that can access the user's documents and My Drive folder.
 * This Picker is used to select the report template file.
 *
 * @param {string} token An OAuth 2.0 access token that lets Picker access
 *     the file type specified in the addView call.
 */
function createFilePicker(token) {
  if (pickerApiLoaded && token) {
    var picker = new google.picker.PickerBuilder()
        
        // Instruct Picker to display Spreadsheets only.
        .addView(new google.picker.View(google.picker.ViewId.SPREADSHEETS))

        // Allow user to select files from Google Drive.
        .addView(new google.picker.DocsView()
            .setIncludeFolders(true)
            .setOwnedByMe(true))
        
        // Hide title bar since an Apps Script dialog already has a title.
        .hideTitleBar()
        
        .setOAuthToken(token)
        .setDeveloperKey(DEVELOPER_KEY)
        .setCallback(filePickerCallback)
        .setOrigin(google.script.host.origin)
        
        // Instruct Picker to fill the dialog, minus 2 px for the border.
        .setSize(DIALOG_DIMENSIONS.width - 2,
            DIALOG_DIMENSIONS.height - 2)
        .build();

    picker.setVisible(true);
  } else {
    showError('<div class="alert alert-error">' +
        'Unable to load the file picker. Please try again.' +
      '</div>' +
      closeButton());
  }
}


/**
 * Creates a Picker that can access the user's folders and My Drive folder.
 * This Picker is used to select the report storage folder.
 *
 * @param {string} token An OAuth 2.0 access token that lets Picker access
 *     the file type specified in the addView call.
 */
function createFolderPicker(token) {
  if (pickerApiLoaded && token) {
    var picker = new google.picker.PickerBuilder()

        // Allow user to select folders from their My Drive.
        .addView(new google.picker.DocsView()
            .setIncludeFolders(true)
            .setSelectFolderEnabled(true)
            .setOwnedByMe(true))
        
        // Hide title bar since an Apps Script dialog already has a title.
        .hideTitleBar()
        
        .setOAuthToken(token)
        .setDeveloperKey(DEVELOPER_KEY)
        .setCallback(folderPickerCallback)
        .setOrigin(google.script.host.origin)
        
        // Instruct Picker to fill the dialog, minus 2 px for the border.
        .setSize(DIALOG_DIMENSIONS.width - 2,
            DIALOG_DIMENSIONS.height - 2)
        .build();

    picker.setVisible(true);
  } else {
    showError('<div class="alert alert-error">' +
        'Unable to load the folder picker. Please try again.' +
      '</div>' +
      closeButton());
  }
}


/**
 * A callback function that extracts the chosen document's metadata from the
 * response object. For details on the response object, see
 * https://developers.google.com/picker/docs/results
 *
 * @param {object} data The Picker JSON-response object.
 */
function filePickerCallback(data) {
  if (data.action == google.picker.Action.PICKED) {
    google.script.run
        .withSuccessHandler(updateDisplay)
        .loadSelectedFile(data.docs);
  } else if (data.action == google.picker.Action.CANCEL) {
    google.script.host.close();
  }
}


/**
 * A callback function that extracts the chosen folders's metadata from the
 * response object. For details on the response object, see
 * https://developers.google.com/picker/docs/results
 *
 * @param {object} data The Picker JSON-response object.
 */
function folderPickerCallback(data) {
  if (data.action == google.picker.Action.PICKED) {
    updateDisplay('<em>Selecting...</em>');
    google.script.run
        .withSuccessHandler(updateDisplay)
        .loadSelectedFolder(data.docs);
  } else if (data.action == google.picker.Action.CANCEL) {
    google.script.host.close();
  }
}