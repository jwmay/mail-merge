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
  return merge.runMerge();
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
  //    call to eliminate this mess
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
  var fileId = '';
  if (this.config.debug === false) {
    var fileName = this.config.outputFileNamePrefix + this.template.getName();
    fileId = this.template.makeCopy(fileName);
  } else {
    fileId = (this.mergeType === 'labels' ? this.config.outputFileIds.label : this.config.outputFileIds.letter);
  }
  return new OutputDocument(fileId);
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
  // Store an error DisplayObject if an error occurs
  var error = null;

  // Get the records and fields from the data spreadsheet
  var records = this.spreadsheet.getRecords();
  var fields = records[0]; // the fields are in the first row of the records

  // Return an error if there are no records
  if (records.length < 2) {
    return getDisplayObject('alert-error',
        'There are no records in the spreadsheet.');
  }

  // Run the selected merge type, or return an error DisplayObject instance
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
  return getDisplayObject('alert-success', content);
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
    this.output.clearBody();

    // Use to determine the current row and column positions for each cell of
    // the output table; contains max number of rows and columns in the table
    var tableInfo = this.template.getTableDimensions();

    // Get a copy of the template table (from the output document).
    // When using a copy of the table from the template document, any images
    // in the table cells would show up as blank images (likely due to an image
    // source error). Retrieving copies directly from the output document (which
    // is created as a copy of the template document file) preserves the images.
    var tableCopy = this.output.getTableCopy();
    
    var tableNum = 0; // count the number of appended tables
    var numRecords = this.spreadsheet.getRecordCount();
    for (var recordNum = 1; recordNum < records.length; recordNum++) {
      var record = records[recordNum];

      // Get the current row and column positions in the table
      var current = {
        rowNum: tableInfo.currentRow(recordNum),
        colNum: tableInfo.currentCol(recordNum)
      };

      // Get the current table cell; note that the cell row and column positions
      // are shifted by 1 because Google starts their indexes at 0
      var cellRow = (current.rowNum - 1);
      var cellCol = (current.colNum - 1);
      var currentTableCell = tableCopy.getCell(cellRow, cellCol);

      // Replace all of the fields in the current cell of the copied table
      for (var fieldNum = 0; fieldNum < fields.length; fieldNum++) {
        var field = fields[fieldNum];
        currentTableCell.replaceText('<<' + field + '>>', record[fieldNum]);
      }

      // Tasks for when the end of the table or the last record is reached
      if (current.rowNum === tableInfo.rows && current.colNum === tableInfo.cols || recordNum === numRecords) {
        // Tables must have a paragraph before and after them, this ensures the
        // table appears in the same position on every page of the document
        if (tableNum !== 0) {
          this.output.body.appendParagraph('');
        }
        
        // Clear any unused cells in the last row of the table
        if (recordNum === numRecords) {
          var numCellsToClear = (tableInfo.cols - current.colNum) + (tableInfo.rows - current.rowNum) * tableInfo.cols;
          if (numCellsToClear > 0) {
            for (var cellNum = 0; cellNum < numCellsToClear; cellNum++) {
              // Add the next record number to the cell number to clear all
              // remaining unused cells
              var currentCell = (cellNum + recordNum + 1);
              var rowNum = tableInfo.currentRow(currentCell);
              var colNum = tableInfo.currentCol(currentCell);
              // Shifts of -1 account for cell indexes starting at 0
              tableCopy.getCell((rowNum - 1), (colNum - 1)).clear();
            }
          }
        }
        
        // Append the table, get a new table copy, and update the table counter
        this.output.body.appendTable(tableCopy);
        tableCopy = this.output.getTableCopy();
        tableNum += 1;
      }
    }
    
    // No errors so return null
    return null;

  } else {
    // An incorrect number of tables were found in the template document
    var message = 'The <em>labels</em> merge type requires the template ' +
        'document have only one table. This document contains <strong>' +
        numTables + '</strong> tables.';
    return getDisplayObject('alert-error', message);
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
 * 
 * @todo To speed up this function, can we put the body elements into a table
 *    with no padding or borders (invisible) and then append tables to the
 *    output document?
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