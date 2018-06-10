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
 * @constructor
 */
var OutputDocument = function(id) {
  this.id = id;
  this.document = this.getDocument();
};


/**
 * Returns the output document as a Google Document object, which is a copy of
 * the template document.
 * 
 * @return {Document} The output document.
 */
OutputDocument.prototype.getDocument = function() {
  var document = DocumentApp.openById(this.id);
  return document;
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
 * Clears the contents of the body element.
 * 
 * @return {Body} The Body element of the document.
 */
OutputDocument.prototype.clearBody = function() {
  var body = this.document.getBody();
  body = body.clear();
  return body;
};


/**
 * Inserts a page break on the last page of the document add adds the given
 * content to the new page.
 * 
 * @param {array} content An array of Element objects added to the new page.
 * @param {boolean} last Flag to control output of a page break at the end of
 *        the new content. Do not add a page break if true, otherwise, add a
 *        page break to the end of the content.
 */
OutputDocument.prototype.insertNewPage = function(content, first, last) {
  var body = this.document.getBody();
  log(' *** START COPY TEMPLATE *** ');
  for (var i = 0; i < content.length; i++) {
    var element = content[i];
    var type = element.getType();
    log('   * Element type: ' + type);
    if (type == DocumentApp.ElementType.PARAGRAPH) {
      body.appendParagraph(element);
    } else if (type == DocumentApp.ElementType.LIST_ITEM) {
      body.appendListItem(element);
    } else if (type == DocumentApp.ElementType.TABLE) {
      body.appendTable(element);
    }
  }
  
  // Remove the empty paragraph at the beginning of the document
  if (first === true) body.getChild(0).removeFromParent();

  // Only add a page break if it is not the last new page
  if (last === false) body.appendPageBreak();
  log(' *** END COPY TEMPLATE *** \n');
};


/**
 * Returns true if the document has a header.
 * 
 * @todo FIX: This function always detects a header, even if one is not present
 * in the document.
 */
OutputDocument.prototype.hasHeader = function() {
  var header = this.document.getHeader();
  log('  * Header: ' + header + ' *');
  // var headerText = header.findElement(DocumentApp.ElementType.PARAGRAPGH);
  var headerText = header.appendParagraph('This is just a test.');
  return headerText;
};