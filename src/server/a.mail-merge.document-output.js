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
 * Base class for the output document that contains the output of the data
 * spreadsheet merged into the template file.
 * 
 * The output document itself is created by the Merge class, which makes a copy
 * of the template document and passes the id of that copied document to create
 * and instance of the OutputDocument.
 * 
 * @constructor
 * @param {string} id The id of the output document.
 */
var OutputDocument = function(id) {
  this.id = id;
  this.document = this.getDocument();
  this.body = this.document.getBody();
  this.bodyCopy = this.body.copy();
  this.config = Configuration.getCurrent();
};


/**
 * Clears the contents of the body element in the document itself. A preserved
 * copy of the original body document is saved in the body property.
 */
OutputDocument.prototype.clearBody = function() {
  this.body.clear();
};


/**
 * Returns a detached, deep copy of the body element.
 * 
 * Any child elements present in the element are also copied. The new element
 * will not have a parent.
 * 
 * @returns {Body} A copy of the body.
 */
OutputDocument.prototype.getBodyCopy = function() {
  var bodyCopy = this.bodyCopy.copy();
  return bodyCopy;
};


/**
 * Returns the output document as a Google Document object, which is a copy of
 * the template document.
 * 
 * @returns {Document} The output document.
 */
OutputDocument.prototype.getDocument = function() {
  var document = DocumentApp.openById(this.id);
  return document;
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
OutputDocument.prototype.getTableCopy = function() {
  var tableCopy = this.bodyCopy.getTables()[0].copy();
  return tableCopy;
};


/**
 * Returns the url of the output document.
 * 
 * @returns {string} The url of the output document.
 */
OutputDocument.prototype.getUrl = function() {
  var url = this.document.getUrl();
  return url;
};


/**
 * Returns true if the document has a header.
 * 
 * @todo FIX: This function always detects a header, even if one is not present
 * in the document.
 */
OutputDocument.prototype.hasHeader = function() {
  var header = this.document.getHeader();
  if (this.config.debug) log('  * Header: ' + header + ' *');
  // var headerText = header.findElement(DocumentApp.ElementType.PARAGRAPGH);
  var headerText = header.appendParagraph('This is just a test.');
  return headerText;
};


/**
 * Inserts a page break on the last page of the document and adds the given
 * content to the new page.
 * 
 * The page parameter accepts an object consisting of flags indicating if the
 * new page is the first or last page of the document. This removes an extra
 * paragraph on the first page of the document and prevents the addition of a
 * page break on the last page of the document.
 * 
 * @param {array} content An array of Element objects added to the new page.
 * @param {object} page Flags to control formatting of output document.
 */
OutputDocument.prototype.insertNewPage = function(content, page) {
  for (var i = 0; i < content.length; i++) {
    var element = content[i];
    var type = element.getType();
    if (type == DocumentApp.ElementType.PARAGRAPH) {
      this.body.appendParagraph(element);
    } else if (type == DocumentApp.ElementType.LIST_ITEM) {
      this.body.appendListItem(element);
    } else if (type == DocumentApp.ElementType.TABLE) {
      this.body.appendTable(element);
    }
  }
  
  // Remove the empty paragraph at the beginning of the document
  if (page.first === true) this.body.getChild(0).removeFromParent();

  // Only add a page break if it is not the last new page to be added
  if (page.last === false) this.body.appendPageBreak();
};


/**
 * Removes the extra paragraph element added after each table element.
 * 
 * This method should be run once creation of the output document is complete
 * to remove all added paragraph elements from any inserted table elements.
 * 
 * Note: this method is intended for us in the non-table-wrapped method of
 * running a letter merge.
 */
OutputDocument.prototype.removeTableParagraphs = function() {
  var tables = this.body.getTables();
  for (var i = 0; i < tables.length; i++) {
    var table = tables[i];
    table.getNextSibling().removeFromParent();
  }
};


/**
 * Reduces the top and bottom margins of the document by 0.05 inches to account
 * for the table wrapper around the body.
 * 
 * This method is for letter merges that use the table wrapping feature to speed
 * up the merge. Goolge requires a paragraph before and after the table when it
 * is at the beginning and/or end of a document. While those paragraph elements
 * have been reduced to a 1px font size, there is still some spacing issues that
 * this fixes. The shift of 0.05 inches (3.6 pts) was determined during testing.
 */
OutputDocument.prototype.shiftMargins = function() {
  var shift = 3.6;
  var topMargin = this.body.getMarginTop();
  var bottomMargin = this.body.getMarginBottom();
  this.body.setMarginTop(topMargin - shift);
  this.body.setMarginBottom(bottomMargin - shift);
};




/**
 * Helper class for wrapping the body content of the template in a table and
 * storing a copy (template) of that table.
 * 
 * This is the 'magic' behind speeding up the letter merge type. By creating
 * only one table, we can use the replaceText method on copies of the table, 
 * and append the table to the body of the output; this drastically reduces the
 * number of append calls (when appending each element in the template to the
 * output) which require the longest server call time.
 * 
 * @param {Body} body The body element of the output document.
 * @param {Element[]} content An array of elements to insert.
 */
var BodyWrapper = function(body, content) {
  this.body = body.copy().clear();
  this.topParagraph = this.body.getChild(0);
  this.table = this.body.appendTable();
  this.tableCell = this.table.appendTableRow()
      .appendTableCell()
      .setPaddingBottom(0)
      .setPaddingLeft(0)
      // Fix for right border of table elements 'disappearing'; this small
      // amount of padding prevents thier right border from being hidden by
      // the surrouing table wrapper
      .setPaddingRight(1)
      .setPaddingTop(0);
  this.bottomParagraph = this.body.getChild(2);

  // Insert the given content into the table cell
  this.insertContent(content);

  // Formatting adjustments to minimize the effects of the table wrapper; set
  // paragraphs before and after the table wrapper to have a font size of 1px
  // and the table to have a border width of 0
  this.topParagraph.setLineSpacing(1).editAsText().setFontSize(1);
  this.bottomParagraph.setLineSpacing(1).editAsText().setFontSize(1);
  this.table.setBorderWidth(0);

  // Save a copy of the table
  this.tableCopy = this.table.copy();
};


/**
 * Appends the currently stored table-wrapped template document to the output
 * document.
 * 
 * Also appends a paragraph before and after the table to account for Google's
 * requirement that document not begin or end with a table. By appending these
 * paragraph elements to every table in the output document, we ensure (with
 * high confidence) the page layout remains identical to that of the template.
 * 
 * The page parameter accepts an object consisting of flags indicating if the
 * new page is the first or last page of the document. This removes an extra
 * paragraph on the first page of the document and prevents the addition of a
 * page break on the last page of the document.
 * 
 * @param {Body} body The body element of the output document.
 * @param {object} page Flags to control formatting of output document.
 */
BodyWrapper.prototype.appendWrappedBody = function(body, page) {
  body.appendParagraph(this.topParagraph.copy());
  var tableParagraph = body.appendTable(this.table.copy()).getNextSibling();
  body.appendParagraph(this.bottomParagraph.copy());
  
  // Remove the default empty paragraph Google adds after a table
  tableParagraph.removeFromParent();

  // Remove the empty paragraph at the beginning of the document
  if (page.first === true) body.getChild(0).removeFromParent();

  // Only add a page break if it is not the last new page to be added; also
  // remove the extra paragraph element after the table
  if (page.last === false) {
    body.appendPageBreak().getParent().getPreviousSibling().removeFromParent();
  }
};


/**
 * Returns a copy of the table which wraps the template document body content.
 * 
 * @returns {Table} A copy of the table that wraps the body content.
 */
BodyWrapper.prototype.getTableCopy = function() {
  return this.tableCopy.copy();
};


/**
 * Inserts the given array of elements into the table cell.
 * 
 * @param {Element[]} content An array of elements to insert.
 */
BodyWrapper.prototype.insertContent = function(content) {
  for (var i = 0; i < content.length; i++) {
    var element = content[i];
    var type = element.getType();
    if (type == DocumentApp.ElementType.PARAGRAPH) {
      this.tableCell.appendParagraph(element);
    } else if (type == DocumentApp.ElementType.LIST_ITEM) {
      this.tableCell.appendListItem(element);
    } else if (type == DocumentApp.ElementType.TABLE) {
      this.tableCell.appendTable(element);
    }
  }
  // Remove the default empty first paragraph from the table cell
  this.tableCell.getChild(0).removeFromParent();
};