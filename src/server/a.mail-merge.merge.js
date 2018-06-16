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
 * @returns {DisplayObject} An instance of DisplayObject with the results of
 *    the merge for display in the UI.
 */
function runMerge() {
  var merge = new Merge();
  var results = merge.runMerge();
  return results;
}


/**
 * The Merge object performs all of the merge functions by using the data
 * spreadsheet and template document and is responsible for creating the
 * output document.
 * 
 * @constructor
 */
var Merge = function() {
  this.config = Configuration.getCurrent();
  this.spreadsheet = new DataSpreadsheet();
  this.template = new TemplateDocument();
  this.output = {};

  // @todo Create an Options class and have this.options.mergeType as a single
  //       call to eliminate this mess
  this.storage = new PropertyStore();
  this.mergeType = this.storage.getProperty('mergeType');
};


/**
 * Returns an array of all the child Elements cotained within the
 * containerElement.
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
 * Returns an OutputDocument instance after copying the template document.
 * 
 * Create the output document as a copy of the template file, or use a file id
 * from the configuration object if in debug mode so that the same document can
 * be reused during testing eliminating the creation of numerous documents.
 * 
 * @returns {OutputDocument} An OutputDocument instance of the output file.
 */
Merge.prototype.getOutputDocument = function() {
  var fileId = (this.mergeType === 'labels' ? this.config.outputFileIds.label : this.config.outputFileIds.letter);
  if (this.config.debug === false) {
    var fileName = this.config.outputFileNamePrefix + this.template.getName();
    fileId = this.template.makeCopy(fileName);
  }
  var output = new OutputDocument(fileId);
  return output;
};


/**
 * Runs the mail merge by creating a copy of the template file body, replacing
 * the fields with their respective data, and appending the elements of the body
 * to the output document. Returns a DisplayObject instance containing a success
 * message with a link to the output document, or an error message if the merge
 * failed.
 * 
 * @returns {DisplayObject} An instance of DisplayObject with the results of
 *    the merge for display in the UI.
 */
Merge.prototype.runMerge = function() {
  // Get the records and fields from the data spreadsheet
  var records = this.spreadsheet.getRecords();
  var fields = records[0]; // the fields are in the first row of the records

  // Run the selected merge type, or return an error DisplayObject instance
  var error = null; // store an error DisplayObject if an error occurs
  if (this.mergeType === 'letters') {
    error = this.runLetterMerge(records, fields);
  } else if (this.mergeType === 'labels') {
    error = this.runLabelMerge(records, fields);
  } else {
    var message = 'An unsupported merge type was selected.';
    error = getDisplayObject('alert-error', message);
  }

  // Return an error DisplayObject if the merge failed for any given record
  if (error !== null) return error;

  // Return a success message with a link to the output document
  var url = this.output.getUrl();
  var content = '' +
      'Merge done! ' +
      '<a href="' + url + '">Click here</a> to open the output document. ';
  var results = getDisplayObject('alert-success', content);
  return results;
};


/**
 * Runs the merge of the records from the data spreadsheet into the output
 * document as a new table cell for each record.
 * 
 * @param {array} records The records from the spreadsheet.
 * @param {array} fields The field names.
 * @returns {null|DisplayObject} Returns null if the merge was successful, or
 *    a DisplayObject instance with an error message.
 */
Merge.prototype.runLabelMerge = function(records, fields) {
  // Look for a table in the template, return error if there is no table
  var numTables = this.template.getNumTables();

  // Label merges only work when the template document has one table
  if (numTables === 1) {
    // The output document is only created if we can perform a merge
    this.output = this.getOutputDocument();
    var outputBody = this.output.document.getBody();

    // Use to determine the current row and column positions for each cell of
    // the output table; also stores the current TableRow for storing TableCells
    var table = this.template.getTableDimensions();
    
    // Clear the output document and insert a new table; the table will be
    // updated below by appending TableRows to this table
    this.output.clearBody();
    var outputTable = outputBody.appendTable();

    // Loop over each record in the data spreadsheet and add it to the output
    var numRecords = this.spreadsheet.getRecordCount();
    for (var recordNum = 1; recordNum < records.length; recordNum++) {
      var record = records[recordNum];
      
      // Replace all of the fields in a copy of the cell element of the template
      // document with their respective values
      var cellCopy = this.template.getCellCopy();
      for (var fieldNum = 0; fieldNum < fields.length; fieldNum++) {
        var field = fields[fieldNum];
        cellCopy = cellCopy.replaceText('<<' + field + '>>', record[fieldNum]);
      }
      
      // Get the current row and column position in the output table
      var currentRowNum = table.currentRow(recordNum);
      var currentColNum = table.currentCol(recordNum);

      // Add a new row when the current cell is in the first column position
      // and save the TableRow element in the tableDimensions object for the
      // next pass (record)
      if (currentColNum === 1) {
        table.currentTableRow = outputTable.appendTableRow();
      }

      // Add the modified cell element to the output document
      table.currentTableRow.appendTableCell(cellCopy);

      // Add a new page when table end is reached; needed because tables must
      // have a paragraph before and after a table, this ensures the table
      // appears in the same position on every page of the document
      log(' Row: ' + currentRowNum + '  Col: ' + currentColNum);
      if (currentRowNum === table.rows && currentColNum === table.cols) {
        outputBody.appendParagraph(''); // ensure same start position on new page
        outputTable = outputBody.appendTable();
      }

      // Fill in the last row with empty cells
      if (recordNum === numRecords) {
        var numCellsToAdd = (table.cols - currentColNum);
        if (numCellsToAdd > 0) {
          for (var i = 0; i < numCellsToAdd; i++) {
            var emptyCell = this.template.getCellCopy().clear();
            table.currentTableRow.appendTableCell(emptyCell);
          }
        }
      }
    }

    // No errors so return null
    return null;
  } else {
    // An incorrect number of tables were found in the template document
    var message = 'The <em>labels</em> merge type requires the template ' +
        'document have only one table. This document contains <strong>' +
        numTables + '</strong> tables.';
    var error = getDisplayObject('alert-error', message);
    return error;
  }
};


/**
 * Runs the merge of the records from the data spreadsheet into the output
 * document as a new page for each record.
 * 
 * @param {array} records The records from the spreadsheet.
 * @param {array} fields The field names.
 * @returns {null|DisplayObject} Returns null if the merge was successful, or
 *    a DisplayObject instance with an error message.
 */
Merge.prototype.runLetterMerge = function(records, fields) {
  this.output = this.getOutputDocument();
  this.output.clearBody();

  // Loop over each record in the data spreadsheet and add it to the output
  var numRecords = this.spreadsheet.getRecordCount();
  for (var recordNum = 1; recordNum < records.length; recordNum++) {
    // Get the current record
    var record = records[recordNum];
  
    // Replace all of the fields in a copy of the body element of the template
    // document with their respective values
    var bodyCopy = this.template.getBodyCopy();
    for (var fieldNum = 0; fieldNum < fields.length; fieldNum++) {
      var field = fields[fieldNum];
      bodyCopy.replaceText('<<' + field + '>>', record[fieldNum]);
    }
    
    // Convert the body into an array of its child elements
    var bodyElements = this.extractElements(bodyCopy);
    
    // Add the modified template body elements to the output document
    var page = {
      first: (recordNum === 1 ? true : false),
      last: (recordNum === numRecords ? true : false),
    };
    this.output.insertNewPage(bodyElements, page);
  }

  // Remove the extra empty paragraph elements added after each table
  this.output.removeTableParagraphs();

  // No errors so return null
  return null;
};