/// Copyright 2018 Joseph W. May. All Rights Reserved.
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
 * Runs the merge and returns the results of the merge, either a success message
 * or an error message, for display in the UI.
 * 
 * @return {DisplayObject} The results of the merge for display in the UI.
 */
function runMerge() {
  var merge = new Merge();
  var results = merge.runMerge();
  return results;
}


/**
 * The Merge object performs the merge by using the data spreadsheet and
 * template document and is responsible for creating the output document.
 * 
 * @constructor
 */
var Merge = function() {
  this.spreadsheet = new DataSpreadsheet();
  this.template = new TemplateDocument();
  this.output = {};
};


/**
 * Returns an array of child Elements cotained within the containerElement.
 * 
 * @returns {Elements[]} An array of child elements.
 */
Merge.prototype.extractElements = function(containerElement) {
  var totalElements = containerElement.getNumChildren();
  var elements = [];
  for (var i = 0; i < totalElements; i++) {
    var element = containerElement.getChild(i).copy();
    elements.push(element);
  }
  return elements;
};


/**
 * Runs the mail merge by creating a copy of the template file body, replacing
 * the fields with their respective data, and appending the elements of the body
 * to the output document. Returns a DisplayObject containing a success message
 * with a link to the output document.
 * 
 * @return {DisplayObject} The results of the merge for display in the UI.
 */
Merge.prototype.runMerge = function() {
  // Create the output document as a copy of the template file
  this.output = this.getOutputDocument(); // hide this when testing
  // this.output = new OutputDocument('1Xa0vpSHrBCfdwo0QZxLN6ILM8nlLp6YHlTRXtNGnthQ');  // show this when testing
  this.output.clearBody();

  // Get the records (with header) and loop over each record
  var records = this.spreadsheet.getRecords();
  var fields = records[0];
  for (var recordNum = 1; recordNum < records.length; recordNum++) {
    // Get the current record
    var record = records[recordNum];
    log('>>> Current Record: ' + recordNum + ' <<<');

    // Get a copy of the template and loop over the fields to replace them in body copy
    var bodyCopy = this.template.getBodyCopy();
    for (var fieldNum = 0; fieldNum < fields.length; fieldNum++) {
      var field = fields[fieldNum];
      bodyCopy.replaceText('<<' + field + '>>', record[fieldNum]);
    }

    // Convert the body into an array of its child elements
    var bodyElements = this.extractElements(bodyCopy);
    
    // Add the modified template body elements to the output document
    var first = (recordNum === 1 ? true : false);
    var last = (recordNum == (records.length - 1) ? true : false);
    this.output.insertNewPage(bodyElements, first, last);
  }

  // Return a success message with a link to the output document
  var url = this.output.getUrl();
  var content = 'Merge done! ' +
          '<a href="' + url + '">Click here</a> to open the output document. ';
  var results = getDisplayObject('alert-success', content);
  return results;
};


/**
 * Returns the output file which is a copy of the template file.
 * 
 * @return {OutputDocument} The output document.
 */
Merge.prototype.getOutputDocument = function() {
  var outputFileName = '[Merge Output] ' + this.template.getName();
  var outputFileId = this.template.makeCopy(outputFileName);
  var outputFile = new OutputDocument(outputFileId);
  return outputFile;
};