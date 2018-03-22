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
 * Base class for the output document that contains the output of the data
 * spreadsheet merged into the template file.
 * 
 * @constructor
 */
var OutputDocument = function(id) {
  this.id = id;
};


OutputDocument.prototype.getDocument = function() {
  var document = DocumentApp.openById(this.id);
  return document;
};


/**
 * 
 */
OutputDocument.prototype.getId = function() {
  return this.id;
};


/**
 * 
 */
OutputDocument.prototype.getNextRecordCount = function() {
  var nextRecordKey = '<<Next record>>';
  var nextRecordRegex = new RegExp(nextRecordKey, 'g');
  var bodyText = this.getDocument().getBody().getText();
  var nextRecordCount = bodyText.match(nextRecordRegex).length;
  return nextRecordCount;
};