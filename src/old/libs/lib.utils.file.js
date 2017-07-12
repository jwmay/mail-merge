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
 * Base class for workign with Google Drive File objects.
 * 
 * @constructor
 * @param {string} fileId The Id of the file.
 */
var BaseFile = function(fileId) {
  this.fileId = fileId;
  this.file = this.getFile();
  this.name = this.getName();
  this.url = this.getUrl();
};


/**
 * Returns a Google Drive File object with the given file id. If the file cannot
 * be found or you do not have permission to access the file, returns null.
 * 
 * @return {File} A Google File object.
 */
BaseFile.prototype.getFile = function() {
  try {
    var file = DriveApp.getFileById(this.fileId);
    return file; 
  } catch(e) {
    return null;
  }
};


/**
 * Returns the name of the file, or null if there is no file.
 * 
 * @return {string} The name of the file.
 */
BaseFile.prototype.getName = function() {
  if (this.file !== null) {
    var name = this.file.getName();
    return name;
  }
  return null;
};


/**
 * Returns the url of the file, or null if there is no file.
 * 
 * @return {string} The url of the file.
 */
BaseFile.prototype.getUrl = function() {
  if (this.file !== null) {
    var url = this.file.getUrl();
    return url;
  }
  return null;
};


/**
 * Returns the first parent folder of the file.
 * 
 * @return {Folder} A Google Folder object.
 */
BaseFile.prototype.getParent = function() {
  var file = this.file;
  var parents = file.getParents();
  var parent = parents.next();
  return parent;
};


/**
 * Copies the file to the given destination with the given name and returns
 * a BaseFile object representing the newly-created file.
 * 
 * @param {string} name The name of the duplicated file.
 * @param {Folder} destination A Google Folder object where the file will be
 *     stored.
 * @return {object} A BaseFile object representing the copied file.
 */
BaseFile.prototype.copyFile = function(name, destination) {
  var copiedFile = this.file.makeCopy(name, destination);
  var copiedFileId = copiedFile.getId();
  var copiedFileObject = new BaseFile(copiedFileId);
  return copiedFileObject;
};


/**
 * Creates a PDF file of the given Google Docs file and returns a Google Drive
 * File object representing the newly-created PDF.
 * 
 * @param {string} name The name for the PDF file.
 * @param {Folder} destination A Google Folder object.
 * @return {File} A Google File object of the PDF.
 */
BaseFile.prototype.createPdf = function(name, destination) {
  var document = DocumentApp.openById(this.fileId);
  var pdfBlob = document.getAs('application/pdf');

  pdfBlob.setName(name);
  var pdfFile = destination.createFile(pdfBlob);

  return pdfFile;
};


/**
 * Moves the file to the destination folder, removing from its parent folder.
 */
BaseFile.prototype.moveTo = function(destination) {
  var file = this.file;
  var parent = this.getParent();
  destination.addFile(file);
  parent.removeFile(file);
};