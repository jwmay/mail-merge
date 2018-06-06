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
 * X
 * 
 * @return {object} A displayObject containing the results of the merge.
 */
function runMerge() {
  var merge = new Merge();
  var results = merge.runMerge();
  return results;
}


/**
 * Base class.
 */
var Merge = function() {
  this.spreadsheet = new DataSpreadsheet();
  this.template = new TemplateDocument();
  this.output = {};
};


/**
 * X
 * 
 * @return {object} A displayObject containing the results of the merge.
 */
Merge.prototype.runMerge = function() {
  this.output = this.getOutputDocument();
  // this.output = new OutputDocument('1LZDeMmFyTn-tqcx9rAy-cGydQpfZt83GyKeuYP1Pglg');  // for testing only

  // Return a success message.
  var results = getDisplayObject('alert-success', 'Merge done!');
  return results;
};


/**
 * 
 */
Merge.prototype.getOutputDocument = function() {
  var outputFileName = '[Merge Output] ' + this.template.getName();
  var outputFileId = this.template.makeCopy(outputFileName);
  var outputFile = new OutputDocument(outputFileId);
  return outputFile;
};