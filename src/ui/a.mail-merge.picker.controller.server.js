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
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}


/**
 * Process the selected file from the Google Picker API. Perform mime-type
 * validation and return a display message based on the validation of the
 * selected file.
 * 
 * @param {array} files An array of JSON objects returned by Google Picker
 *     representing the selected file.
 * @return {string} An HTML-formatted string containing the results message.
 */
function loadSelectedFile(files) {
  var file = files[0];
  file.mime_type = DriveApp.getFileById(file.id).getMimeType();

  if (file.mime_type === MimeType.GOOGLE_DOCS) {
    var templateFile = new TemplateFile();
    templateFile.setFileId(file.id);
    var success_message = '<div class="msg msg-success">' +
              'Template file updated' +
            '</div>' +
            showCloseButton();
    // Refresh the sidebar to display the newly-selected file.
    onShowSidebar();
    return success_message;
  } else {
    var error_message = '<div class="msg msg-error">' +
              'Only Google Docs files can be used as a template' +
            '</div>' +
            showCloseButton();
    return error_message;
  }
}


/**
 * Process the selected folder from the Google Picker API.
 * 
 * @param {array} folders An array of JSON objects returned by Google Picker
 *     representing the selected folder.
 * @return {string} An HTML-formatted string containing the results message.
 */
function loadSelectedFolder(folders) {
  var folder = folders[0];

  if (folder.type === 'folder') {
    var reportsFolder = new ReportsFolder();
    reportsFolder.setFolderId(folder.id);
    var success_message = '<div class="msg msg-success">' +
              'Reports folder updated' +
            '</div>' +
            showCloseButton();
    // Refresh the sidebar to display the newly-selected folder.
    onShowSidebar();
    return success_message;
  } else {
    var error_message = '<div class="msg msg-error">' +
          'Only a folder can be selected to store reports' +
        '</div>' +
        showCloseButton();
    return error_message;
  }
}