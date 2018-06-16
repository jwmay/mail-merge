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
 * Helper function for inserting a merge field into the template document.
 * 
 * @param {string} field The merge field to insert.
 * @returns {DisplayObject|null} Null is returned if the insert was successful,
 *    otherwise, a DisplayObject instance is returned with an error message if
 *    the insert was not successful.
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
var TemplateDocument = function() {
  this.document = this.getDocument();
};


/**
 * Returns a detached, deep copy of the template document body element.
 * 
 * Any child elements present in the element are also copied. The new element
 * will not have a parent.
 * 
 * @returns {Body} A copy of the document body.
 */
TemplateDocument.prototype.getBodyCopy = function() {
  var body = this.document.getBody();
  var copy = body.copy();
  return copy;
};


/**
 * Returns a detached, deep copy of the first cell of the first table in the
 * document body element.
 * 
 * Any child elements present in the element are also copied. The new element
 * will not have a parent.
 * 
 * @returns {TableCell} A copy of the table cell.
 */
TemplateDocument.prototype.getCellCopy = function() {
  var tableCellCopy = this.getTable().getRow(0).getCell(0).copy();
  return tableCellCopy;
};


/**
 * Returns the template document as a Google Document object.
 * 
 * @returns {Document} The template file as a Google Document object.
 */
TemplateDocument.prototype.getDocument = function() {
  var document = DocumentApp.getActiveDocument();
  return document;
};


/**
 * Returns the id of the template document.
 * 
 * @returns {string} The id of the template document.
 */
TemplateDocument.prototype.getId = function() {
  var id = this.document.getId();
  return id;
};


/**
 * Returns the formatted merge field.
 * 
 * @static
 * @returns {string} The formatted merge field.
 */
TemplateDocument.getMergeField = function(field) {
  var mergeField = '<<' + field + '>>';
  return mergeField;
};


/**
 * Returns the title of the template document.
 * 
 * @returns {string} The title of the template document.
 */
TemplateDocument.prototype.getName = function() {
  var name = this.document.getName();
  return name;
};


/**
 * Returns the number of tables in the document.
 * 
 * @returns {integer} The number of tables in the document.
 */
TemplateDocument.prototype.getNumTables = function() {
  var tables = this.document.getBody().getTables();
  var numTables = tables.length;
  return numTables;
};


/**
 * Returns the first table in the document body element.
 * 
 * When running a label merge, the document should only contain one table. This
 * method retrieves the first table in the document regardless of there being
 * any additional tables.
 * 
 * @returns {Table} The first table in the document.
 */
TemplateDocument.prototype.getTable = function() {
  var tables = this.document.getBody().getTables();
  var table = tables[0];
  return table;
};


/**
 * Returns an object with the dimensions of the first table in the document as
 * well as methods for determining the current row and column numbers given a
 * cell index.
 * 
 * The object includes the keys `rows` and `cols`, which are the number of rows
 * and columns, respectively. The methods `currentRow` and `currentCol` are for
 * determining the current row and column numbers given a cell index. The
 * `currentTableRow` property is for storing the current TableRow object during
 * a label merge.
 * 
 * @returns {object} An object containing the dimensions of the first table in
 *    the document.
 */
TemplateDocument.prototype.getTableDimensions = function() {
  var table = this.getTable();
  var dimensions = {
    rows: table.getNumRows(),
    cols: table.getRow(1).getNumCells(),
    currentTableRow: {},
    /**
     * Returns the 2D row index given the 1D cell index.
     * 
     * @returns {integer} The row index.
     */
    currentRow: function(cell) {
      var currentRow = (Math.ceil(cell / this.cols) % this.rows);
      currentRow = (currentRow === 0 ? this.rows : currentRow);
      return currentRow;
    },
    /**
     * Returns the 2D column index given the 1D cell index.
     * 
     * @returns {integer} The column index.
     */
    currentCol: function(cell) {
      var currentCol = (cell % this.cols);
      currentCol = (currentCol === 0 ? this.cols : currentCol);
      return currentCol;
    },
  };
  return dimensions;
};


/**
 * Inserts the merge field into the template document as <<field>>. Returns null
 * if the field was successfully inserted, otherwise, returns a DisplayObject
 * instance containing the error message.
 * 
 * @param {string} field The merge field to insert.
 * @returns {null|object} Returns null if the field was successfully inserted,
 *    otherwise, returns a DisplayObject instance containing the error message.
 */
TemplateDocument.prototype.insertMergeField = function(field) {
  var cursor = this.document.getCursor();
  var dataVariable = TemplateDocument.getMergeField(field);
  var element = cursor.insertText(dataVariable);
  if (element !== null) {
    // Position the cursor at the end of the inserted merge field variable
    var elementEnd = (field.length + 4); // the 4 accounts for the '<<' and '>>'
    var position = this.document.newPosition(element, elementEnd);
    this.document.setCursor(position);
    return null;
  } else {
    var error = getDisplayObject('alert-error',
        'There was an error inserting the merge field.');
    return error;
  }
};


/**
 * Creates a copy of the template file with the given output name and
 * returns the id of the new copy.
 * 
 * @param {string} outputName The filename that should be applied to the new copy.
 * @returns {string} The id of the new copy.
 */
TemplateDocument.prototype.makeCopy = function(outputName) {
  var id = this.getId();
  var templateFile = DriveApp.getFileById(id);
  var outputFile = templateFile.makeCopy(outputName);
  var outputFileId = outputFile.getId();
  return outputFileId;
};