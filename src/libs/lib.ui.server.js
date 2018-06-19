// Copyright 2018 Joseph W. May. All Rights Reserved.
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
 * Returns an HTML-formatted string to display a 'Close' button.
 * 
 * @returns {string} An HTML-formatted string.
 */
function closeButton() {
  return '<div><input type="button" value="Close" class="btn" '+
      'onclick="google.script.host.close();"></div>';
}


/**
 * Displays an alert with a single OK button.
 * 
 * @param {string} title The title of the alert.
 * @param {string} message The message to display.
 */
function showAlert(title, message) {
  var ui = DocumentApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}


/**
 * Displays an alert with 'Yes' and 'No' buttons. Returns the user selection as
 * a 'yes' or 'no' string.
 * 
 * @param {string} title The title of the alert.
 * @param {string} message The message to display.
 * @returns {string} The response as a 'yes' or 'no'.
 */
function showConfirmation(title, message) {
  var ui = DocumentApp.getUi();
  var result = ui.alert(title, message, ui.ButtonSet.YES_NO);
  if (result === ui.Button.YES) return 'yes';
  return 'no';
}


/**
 * Opens a dialog window using an HTML template with the given dimensions
 * and title.
 * 
 * @param {string} source The name of the HTML template file.
 * @param {integer} width The width of dialog window.
 * @param {integer} height The height of dialog window.
 * @param {string} title The title to display on dialog window.
 */
function showDialog(source, width, height, title) {
  var ui = HtmlService.createTemplateFromFile(source)
      .evaluate()
      .setWidth(width)
      .setHeight(height);
  DocumentApp.getUi().showModalDialog(ui, title);
}


/**
 * Displays a prompt and return the user's response.
 * 
 * @param {string} message The message to display.
 * @returns A string representing the user's response, or null if
 *    no response is given.
 */
function showPrompt(message) {
  var ui = DocumentApp.getUi();
  var response = ui.prompt(message);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    return response.getResponseText();
  } else {
    return null;
  }
}


/**
 * Opens a sidebar using an HTML template.
 * 
 * @param {string} source The name of the HTML template file.
 * @param {string} title The title to display on sidebar.
 */
function showSidebar(source, title) {
  var ui = HtmlService.createTemplateFromFile(source)
      .evaluate()
      .setTitle(title);
  DocumentApp.getUi().showSidebar(ui);
}



/**
 * Returns an HTML-formatted select element with the given options. The name
 * of each option will be used as the option's value. Optional class names and
 * an id can be specified for the select element.
 * 
 * @todo If only one option is given, display that option only, do not display
 *    the 'Select an item...' message.
 * 
 * @param {array} options The options to add to the select element.
 * @param {array=} message The default message to display in the select element.
 *    Default is "Select an item...".
 * @param {string=} value The currently-selected value. Default is none.
 * @param {string=} id The id of the select element. Default is none.
 * @param {string=} classes The class name(s) of the select element. Default is
 *    none.
 * @return {string} An HTML-formatted string containing the select element.
 */
function makeSelect(options, message, value, id, classes) {
  // Default parameters are not supported by GAS server code, so define the
  // optional variables here if they are not defined
  message = message === undefined ? 'Select an item...' : message;
  value = value === undefined ? '' : value;
  id = id === undefined ? '' : id;
  classes = classes === undefined ? '' : classes;

  var select = [];
  select.push(Utilities.formatString('<select id="%s" class="%s">', id, classes));

  // Construct the default option if there is no assigned value
  if (value === '' || value === null || value === undefined) {
    select.push('<option value="" class="default">' + message + '</option>');
  }

  // Construct the options
  for (i = 0; i < options.length; i++) {
    var option = options[i];
    var selected = option === value ? 'selected' : '';
    select.push(Utilities.formatString('<option value="%s" %s>%s</option>',
            option, selected, option));
  }

  // Close and construct the select element
  select.push('</select>');
  return select.join('\n');
}


/**
 * Returns an HTML-formatted string of checkboxes for the given value-item
 * objects. The input must be an array of objects each with a value and label
 * property. 
 * 
 * @param {string} name The name for the checkbox group.
 * @param {array} items An array of value-item objects.
 * @param {boolean} allChecked If true, all checkboxes will be initially checked
 *    on display.
 * @returns {string} An HTML-formatted string.
 */
function showCheckboxes(name, items, allChecked) {
  var checkboxes = [];
  var checked = allChecked === true ? ' checked' : '';
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    checkboxes.push(
      Utilities.formatString(
        '<div><label><input type="checkbox" name="%s" value="%s"%s>%s</label></div>',
        name, item.value, checked, item.label
      )
    );
  }
  return checkboxes.join('');
}


/**
 * Returns the data structure for showCheckboxes() given an array of strings.
 * 
 * The returned data is an array of objects each with a value and label
 * property.
 * 
 * @private
 * @param {strings} columns An array of strings.
 * @returns {object[]} An array of objects.
 */
function constructColumnSelectItems_(columns, startIndex) {
  var items = [];
  var firstIndex = startIndex === undefined ? 1 : startIndex;
  for (var i = 0; i < columns.length; i++) {
    var col = columns[i];
    var colNum = i + firstIndex;
    var item = {
      value: colNum,
      label: col
    };
    items.push(item);
  }
  return items;
}


/**
 * Returns the data structure for showCheckboxes() given an array of Sheet
 * objects. The returned data is an array of objects each with a value and
 * label property.
 * 
 * @private
 * @param {array} sheets An array of Sheet objects.
 * @returns {object[]} An array of objects.
 */
function constructSheetSelectItems_(sheets) {
  var items = [];
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var item = {
        value: sheet.getSheetId(),
        label: sheet.getSheetName()
    };
    items.push(item);
  }
  return items;
}