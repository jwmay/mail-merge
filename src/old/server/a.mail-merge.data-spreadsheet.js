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
 * 
 * 
 * @constructor
 */
var DataSpreadsheet = function() {
  this._storage = new PropertyStore();
};


/**
 * 
 */
DataSpreadsheet.prototype.getId = function() {
  var id = this._storage.getProperty('DATA_SPREADSHEET_ID');
  return id;
};


/**
 * 
 */
DataSpreadsheet.prototype.setId = function(id) {
  this._storage.setProperty('DATA_SPREADSHEET_ID', id);
};


/**
 * 
 */
DataSpreadsheet.prototype.getSpreadsheet = function() {
  var id = this.getId();
  if (id !== null) {
    var spreadsheet = SpreadsheetApp.openById(id);
    return spreadsheet;
  } else {
    return null;
  }
};


/**
 * 
 */
DataSpreadsheet.prototype.getName = function() {
  var spreadsheet = this.getSpreadsheet();
  var name = spreadsheet.getName();
  return name;
};


/**
 * 
 */
DataSpreadsheet.prototype.getUrl = function() {
  var spreadsheet = this.getSpreadsheet();
  var url = spreadsheet.getUrl();
  return url;
};