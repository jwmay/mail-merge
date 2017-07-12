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
 * Empty base class for Reports.
 */
var Reports = function() {};


/**
 * Returns the report filename.
 * 
 * @return {string} The filename or null if no filename is stored.
 */
Reports.prototype.getFilename = function() {
  var storage = new PropertyStore();
  var filename = storage.getProperty('REPORT_FILENAME');
  if (filename !== null && filename !== undefined && filename !== '') {
    return filename;
  }
  return null;
};


/**
 * Stores the given filename.
 * 
 * @param {string} filename The filename to be stored.
 */
Reports.prototype.setFilename = function(filename) {
  var storage = new PropertyStore();
  storage.setProperty('REPORT_FILENAME', filename);
};


/**
 * Prompts the user for a new filename, updates the stored filename, and
 * returns the new filename.
 * 
 * @return {string} The newly-stored filename.
 */
Reports.prototype.updateFilename = function() {
  var filename = showPrompt('Enter a new report filename');
  if (filename !== '') {
    this.setFilename(filename);
  }
  return this.getFilename();
};


/**
 * Generates a PDF file for each form response.
 */
Reports.prototype.generateReports = function() {
  var formResponses = new FormResponses();
  var initialized = formResponses.isInitialized();
  
  if (initialized === true) {
    var maxCol = formResponses.getMaxColumn();

    var formResponseSheet = formResponses.sheet;
    var responses = formResponseSheet.getRange(2, 1,
            formResponseSheet.getLastRow() - 1, maxCol).getValues();

    for (var i = 0; i < responses.length; i++) {
      var response = responses[i];
      var row = i + 2;
      var incident = new Incident(row, response);
      
      if (incident.isSent() === false) {
        incident.createReport();      
      }
    }
  }
};


/**
 * Helper function to generate a PDF file for each form response. 
 */
function generateReports() {
  var reports = new Reports();
  reports.generateReports();
}