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
  return template.insertMergeField(field);
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
  return body.copy();
};


/**
 * Returns the template document as a Google Document object.
 * 
 * @returns {Document} The template file as a Google Document object.
 */
TemplateDocument.prototype.getDocument = function() {
  return DocumentApp.getActiveDocument();
};


/**
 * Returns the id of the template document.
 * 
 * @returns {string} The id of the template document.
 */
TemplateDocument.prototype.getId = function() {
  return this.document.getId();
};


/**
 * Returns the formatted merge field.
 * 
 * @static
 * @returns {string} The formatted merge field.
 */
TemplateDocument.getMergeField = function(field) {
  return '<<' + field + '>>';
};


/**
 * Returns the title of the template document.
 * 
 * @returns {string} The title of the template document.
 */
TemplateDocument.prototype.getName = function() {
  return this.document.getName();
};


/**
 * Returns the number of tables in the document.
 * 
 * @returns {integer} The number of tables in the document.
 */
TemplateDocument.prototype.getNumTables = function() {
  var tables = this.document.getBody().getTables();
  return tables.length;
};


/**
 * Returns the parent folder of the template file.
 * 
 * Files can be in more than one folder. This method only returns the first
 * parent folder containing the file. This is determined alphabetically.
 * 
 * @returns {Folder} The parent folder.
 */
TemplateDocument.prototype.getParentFolder = function() {
  var id = this.getId();
  var file = DriveApp.getFileById(id);
  return file.getParents().next();
};


/**
 * Returns a detached, deep copy of the first table in the document.
 * 
 * Any child elements present in the element are also copied. The new element
 * will not have a parent.
 * 
 * When running a label merge, the document should only contain one table. This
 * method retrieves the first table in the document regardless of there being
 * any additional tables.
 * 
 * @returns {Table} A copy of the first table in the document.
 */
TemplateDocument.prototype.getTableCopy = function() {
  var tables = this.document.getBody().getTables();
  return tables[0].copy();
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
  var table = this.getTableCopy();
  return {
    rows: table.getNumRows(),
    cols: table.getRow(0).getNumCells(),
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
  var selection = this.document.getSelection();
  var dataVariable = TemplateDocument.getMergeField(field);
  if (cursor !== null) {
    // Merge field will be inserted at the current cursor position
    var inserted = cursor.insertText(dataVariable);
    if (inserted !== null) {
      // Position the cursor at the end of the inserted merge field variable
      var elementEnd = (field.length + 4); // the 4 accounts for the '<<' and '>>'
      var newPosition = this.document.newPosition(inserted, elementEnd);
      this.document.setCursor(newPosition);
      return null;
    } else {
      return getDisplayObject('alert-error',
          'There was an error inserting the merge field. Please try again.');
    }
  } else if (selection !== null) {
    // The current text selection will be replaced with the merge field
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      var elementType = element.getElement().getType();

      // The following element types are not supported for selection replacement
      if (elementType === DocumentApp.ElementType.EQUATION ||
          elementType === DocumentApp.ElementType.EQUATION_FUNCTION ||
          elementType === DocumentApp.ElementType.EQUATION_FUNCTION_ARGUMENT_SEPARATOR ||
          elementType === DocumentApp.ElementType.EQUATION_SYMBOL ||
          elementType === DocumentApp.ElementType.HORIZONTAL_RULE ||
          elementType === DocumentApp.ElementType.INLINE_DRAWING ||
          elementType === DocumentApp.ElementType.INLINE_IMAGE ||
          elementType === DocumentApp.ElementType.PAGE_BREAK ||
          elementType === DocumentApp.ElementType.TABLE ||
          elementType === DocumentApp.ElementType.TABLE_ROW ||
          elementType === DocumentApp.ElementType.TABLE_CELL) {
        return getDisplayObject('alert-error',
            'Cannot replace selection with a merge field. ' +
            'Only text selections can be replaced.');
      }

      // Only modify elements that can be edited as text; skip images and other
      // non-text elements
      if (element.getElement().editAsText) {
        var text = element.getElement().editAsText();
        if (element.isPartial()) {
          text.deleteText(element.getStartOffset(), element.getEndOffsetInclusive())
              .insertText(element.getStartOffset(), dataVariable);
        } else {
          text.appendText(dataVariable).removeFromParent();
        }
      }
    }
  } else {
    // An unsupported element type was selected
    return getDisplayObject('alert-error',
        'Only text selections can be replaced.');
  }
  return null;
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
  return outputFile.getId();
};


/**
 * Checks the first child element of document body for any content and returns
 * true if the first line does not contain content (is blank) or false if there
 * is content.
 * 
 * @returns {boolean} True if the document starts with a blank line, otherwise,
 *    returns false.
 */
TemplateDocument.prototype.startsWithBlankLine = function() {
  var body = this.document.getBody();
  var firstElement = body.getChild(0);
  if (firstElement.getNumChildren() === 0) return true;
  return false;
};