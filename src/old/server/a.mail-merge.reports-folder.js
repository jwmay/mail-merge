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
 * Base class for the Reports storage folder.
 */
var ReportsFolder = function() {
  this.folderId = this.getFolderId();
  this.folder = this.getFolder();

  // Inherit from BaseFolder.
  BaseFolder.call(this, this.folderId);
};
inherit_(ReportsFolder, BaseFolder);


/**
 * Returns the folder id if one is stored, otherwise, returns null.
 * 
 * @return {string} The folder id.
 */
ReportsFolder.prototype.getFolderId = function() {
  var storage = new PropertyStore();
  var folderId = storage.getProperty('REPORTS_FOLDER_ID');
  if (folderId !== undefined && folderId !== null && folderId !== '') {
    return folderId;
  }
  return null;
};


/**
 * Stores the reports folder id in document properties.
 */
ReportsFolder.prototype.setFolderId = function(folderId) {
  var storage = new PropertyStore();
  storage.setProperty('REPORTS_FOLDER_ID', folderId);
};


/**
 * Returns the reports folder as a Google Folder object.
 * 
 * @return {Folder} A Google Folder object.
 */
ReportsFolder.prototype.getFolder = function() {
  var folder = this.getFolderById(this.folderId);
  return folder;
};


/**
 * Returns the report folder for the current month.
 *
 * @return {Folder} A Google Folder object.
 */
ReportsFolder.prototype.getCurrentMonthFolder = function() {
  var currentMonth = getCurrentMonth_();
  var monthFolders = this.folder.getFoldersByName(currentMonth);
  
  if (monthFolders.hasNext()) {
    return monthFolders.next();
  } else {
    return this.folder.createFolder(currentMonth);
  }
};