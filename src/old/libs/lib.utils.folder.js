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
 * Base class for working with Google Drive Folder objects.
 * 
 * @constructor
 * @param {string=} id The id of the folder. If one is not provided, the user's
 *     root 'MyDrive' folder will be used.
 */
var BaseFolder = function(id) {
  this.rootFolder = this.getRootFolder();
  if (id !== undefined && id !== null && id !== '') {
    this.folder = this.getFolderById(id);
  } else {
    this.folder = this.rootFolder;
  }
};


/**
 * Returns the users root 'MyDrive' folder.
 * 
 * @return {Folder} A Google Drive Folder object.
 */
BaseFolder.prototype.getRootFolder = function() {
  var root = DriveApp.getRootFolder();
  return root;
};


/**
 * Returns the folder with the given id, or null if folder not found.
 * 
 * @param {string} id The id of the folder.
 * @return {Folder} A Google Folder object.
 */
BaseFolder.prototype.getFolderById = function(id) {
  try {
    var folder = DriveApp.getFolderById(id);
    return folder;
  } catch(e) {
    return null;
  }
};


/**
 * Returns the folder with the given path.
 *
 * @param {string} path The path of the folder.
 * @return {Folder} A Google Folder object.
 */
BaseFolder.prototype.getFolderByPath = function(path) {
  var folders = path.split('/');
  var currentFolder = DriveApp.getRootFolder();
  for (var i = 0; i < folders.length; i++) {
    var folderName = folders[i];
    var childFolders = currentFolder.getFoldersByName(folderName);
    if (childFolders.hasNext()) {
      currentFolder = childFolders.next();
    }
  }
  return currentFolder;
};


/**
 * Returns the name of the folder.
 * 
 * @return {string} The name of the folder.
 */
BaseFolder.prototype.getName = function() {
  var name = this.folder.getName();
  return name;
};


/**
 * Returns the url of the folder.
 * 
 * @return {string} The url of the folder.
 */
BaseFolder.prototype.getUrl = function() {
  var url = this.folder.getUrl();
  return url;
};