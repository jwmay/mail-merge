/// Copyright 2017 Joseph W. May. All Rights Reserved.
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
 * Helper function for inserting a merge field into the template document.
 * 
 * @param {string} field The merge field to insert.
 */
function insertMergeField(field) {
  var template = new TemplateDocument();
  var insert = template.insertMergeField(field);
  return insert;
}


/**
 * Base class for the template document that contains the layout and data
 * variables into which the data spreadsheet will be merged.
 * 
 * @constructor
 */
var TemplateDocument = function() {};


/**
 * Returns the template document as a Google Document object.
 * 
 * @return {Document} The template file as a Google Document object.
 */
TemplateDocument.prototype.getDocument = function() {
  var document = DocumentApp.getActiveDocument();
  return document;
};


/**
 * Returns the id of the template document.
 * 
 * @return {string} The id of the template document.
 */
TemplateDocument.prototype.getId = function() {
  var document = this.getDocument();
  var id = document.getId();
  return id;
};


/**
 * Returns the title of the template document.
 * 
 * @return {string} The title of the template document.
 */
TemplateDocument.prototype.getName = function() {
  var document = this.getDocument();
  var name = document.getName();
  return name;
};


/**
 * Creates a copy of the template file with the given output anme and
 * returns the id fo the new copy.
 * 
 * @param {string} outputName The filename that should be applied to
 *        the new copy.
 * @return {string} The id of the new copy.
 */
TemplateDocument.prototype.makeCopy = function(outputName) {
  var id = this.getId();
  var templateFile = DriveApp.getFileById(id);
  var outputFile = templateFile.makeCopy(outputName);
  var outputFileId = outputFile.getId();
  return outputFileId;
};


/**
 * Inserts the merge field into the template document as <<field>>. Returns null
 * if the field was successfully inserted, otherwise, returns a displayObject
 * containing the error message.
 * 
 * @param {string} field The merge field to insert.
 * @return {null|object} Returns null if the field was successfully inserted,
 *        otherwise, returns a displayObject containing the error message.
 */
TemplateDocument.prototype.insertMergeField = function(field) {
  var document = this.getDocument();
  var cursor = document.getCursor();
  var dataVariable = '<<' + field + '>>';
  var element = cursor.insertText(dataVariable);
  if (element !== null) {
    // Position the cursor at the end of the inserted merge field variable.
    var elementEnd = field.length + 4;
    var position = document.newPosition(element, elementEnd);
    document.setCursor(position);
    return null;
  } else {
    var error = getDisplayObject('alert-error',
            'There was an error inserting the merge field.');
    return error;
  }
};