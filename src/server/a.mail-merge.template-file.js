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
 * Base class for the template file.
 * 
 * @constructor
 */
var TemplateFile = function() {
  this.fileId = this.getFileId();

  // Inherit from BaseFile.
  BaseFile.call(this, this.fileId);
};
inherit_(TemplateFile, BaseFile);


/**
 * Returns the folder id if one is stored, otherwise, returns null.
 * 
 * @return {string} The file id.
 */
TemplateFile.prototype.getFileId = function() {
  var storage = new PropertyStore();
  var fileId = storage.getProperty('TEMPLATE_FILE_ID');
  if (fileId !== undefined && fileId !== null && fileId !== '') {
    return fileId;
  }
  return null;
};


/**
 * Stores the template file id in document properties.
 */
TemplateFile.prototype.setFileId = function(fileId) {
  var storage = new PropertyStore();
  storage.setProperty('TEMPLATE_FILE_ID', fileId);
};


/**
 * Copies the template file to the reports folder and returns a BaseFile
 * object representing the copied file.
 * 
 * @return {BaseFile} A BaseFile object of the copied file.
 */
TemplateFile.prototype.copyTemplateFile = function() {
  var name = this.name + ' ' + getTime_();
  var reportsFolder = new ReportsFolder();
  var destination = reportsFolder.folder;
  var copiedTemplateFile = this.copyFile(name, destination);
  return copiedTemplateFile;
};


/**
 * Creates a report template file with all of the header keys.
 */
TemplateFile.prototype.generateFile = function() {
  var reportsFolder = new ReportsFolder();
  var destination = reportsFolder.folder;

  var filename = 'Incident Reporter Template File';
  var file = DocumentApp.create(filename);

  var body = file.getBody();

  var formResponses = new FormResponses();
  var headerKeys = formResponses.getHeaderKeys();

  body.appendParagraph('The following tags are placeholders for information ' +
          'from the form responses. You can arrange them and format them ' +
          'however you would like. You can even remove them. When the report ' +
          'is generated, they will be replaced by the actual form response.\n');
  for (var i = 0; i < headerKeys.length; i++) {
    var headerKey = '<<' + headerKeys[i] + '>>';
    body.appendParagraph(headerKey);
  }

  this.setFileId(file.getId());

  // Move the file to the reports folder.
  var templateFile = new BaseFile(file.getId());
  templateFile.moveTo(destination);
};