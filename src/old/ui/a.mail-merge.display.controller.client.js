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
 * 
 */
function updateDisplay(content, reset) {
  var display = new Display();
  display.updateDisplay(content, reset);
}


/**
 * Base class for the HTML display on the sidebar.
 */
var Display = function() {
  this.display = $('#display');
  this.error = $('#error');
  this.loading = $('#loading');
};


/**
 * Update the HTML of the display to show the given content.
 * 
 * @param {string} content The HTML content to display.
 * @param {boolean=} reset Reset the display before updating if true,
 *         otherwise, append the HTML to the display. Default is false.
 */
Display.prototype.updateDisplay = function(content, reset) {
  var clear = reset !== undefined ? reset : false;
  if (clear === true) this.resetDisplay();
  $(content).hide().appendTo(this.display).slideDown('slow');  
};


/**
 * Clear the display by removing all of the HTML content.
 */
Display.prototype.resetDisplay = function() {
  this.display.html('');
};


/**
 * Update the user display to show loading message.
 */
Display.prototype.showLoading = function(message) {
  var msg = message || 'Loading...';
  var content = '<div id="loading" class="card loading">' +
      '<i class="fa fa-circle-o-notch fa-spin fa-2x fa-fw"></i> ' +
      '<span>' + msg + '</span>' +
    '</div>';
  this.updateDisplay(content, true);
};


/**
 * 
 */
Display.prototype.hideLoading = function() {
  this.loading.remove();
};


/**
 * Update the user display to show error message.
 *
 * @param {string} message The error message to display.
 * @param {string=} header The header to display. Default is 'Error'.
 */
Display.prototype.showError = function(message, header) {
  var head = header !== undefined ? header : 'Error';
  var msg = '<h4 class="alert-header">' + head + '</h4>' +
      '<p>' + message + '</p>';
  this.error.html(msg).removeClass('hidden');
};
function showError(message) {
  var display = $('#error');
  var msg = '<h4 class="alert-header">Error</h4>' +
      '<p>' + message + '</p>';
  display.html(msg).removeClass('hidden');
}


/**
 * Update the user display to hide the error mesage.
 */
Display.prototype.hideError = function() {
  this.error.addClass('hidden');
};
function hideError() {
  var display = $('#error');
  display.addClass('hidden');
}


/**
 * Returns an HTML-formatted string to display the 'Close' button.
 * 
 * @returns {string} An HTML-formatted string.
 */
function closeButton() {
  button = '<input type="button" value="Close" class="btn" onclick="google.script.host.close();">';
  return button;
}


/**
 * Returns the value of a field with the specified selector.
 * 
 * @private
 * @param {string} selector The input selector.
 * @returns {string} A string representing the user's response.
 */
function getValue_(selector) {
  var inputValue = $(selector).val();
  return inputValue;
}


/**
 * Returns the values of all fields with the specified selector.
 * 
 * @private
 * @param {string} selector The input selector.
 * @returns {array} An array of strings representing the user's response for
 *     each field with the given selector.
 */
function getValues_(selector) {
  var inputValues = [];
  $(selector).each(function() {
    inputValues.push($(this).val());
  });
  return inputValues;
}


/**
 * Returns an array of values of checked inputs.
 *
 * @private
 * @param {string} name The name attribute of the checkboxes.
 * @return {array} An array of values of checked inputs.
 */
function getCheckedBoxes_(name) {
  var selector = 'input[name="' + name + '"]:checked';
  var checked = [];
  $(selector).each(function() {
    checked.push($(this).val());
  });
  return checked;
}


/**
 * Returns an array of booleans representing the status of each checkbox
 * with the given name.
 * 
 * @private
 * @param {string} name The name attribute of the checkboxes.
 * @returns {array} An array of booleans.
 */
function getCheckboxStatus_(name) {
  var selector = 'input[name="' + name + '"]';
  var status = [];
  $(selector).each(function() {
    status.push($(this).prop('checked'));
  });
  return status;
}