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
 * Helper function for updating the user-interface display.
 * 
 * @param {string} content The HTML content to display.
 * @param {boolean=} reset Clear the display before updating if true,
 *         otherwise, append the content to the display. Default is false.
 */
function updateDisplay(content, reset) {
  var display = new Display();
  display.hideLoading();
  display.updateDisplay(content, reset);
}


/**
 * Helper function for updating the user-interface display with an error 
 * message.
 * 
 * @param {string} message The error message to display.
 * @param {string=} header The header to display. Default is 'Error'.
 */
function showError(message, header) {
  var display = new Display();
  display.showError(message, header);
}


/**
 * Helper function for clearing an error message from the user-interface
 * display.
 */
function hideError() {
  var display = new Display();
  display.hideError();
}


/**
 * Helper function for displaying the loading message on the user-interface
 * display.
 */
function showLoading() {
  var display = new Display();
  display.showLoading();
}


/**
 * Helper function for hiding the loading message from the user-interface
 * display.
 */
function hideLoading() {
  var display = new Display();
  display.hideLoading();
}


/**
 * Base class for the HTML display on the sidebar.
 */
var Display = function() {
  this.display = $('#display');
  this.error = $('#error');
  this.loading = $('#loading');
  this.loadingOverlay = $('#loading-overlay');
};


/**
 * Update the HTML of the display to show the given content.
 * 
 * @param {string} content The HTML content to display.
 * @param {boolean=} reset Clear the display before updating if true,
 *         otherwise, append the content to the display. Default is false.
 */
Display.prototype.updateDisplay = function(content, reset = false) {
  if (reset === true) this.resetDisplay();
  $(content).hide().appendTo(this.display).slideDown('slow');  
};


/**
 * Clear the display by removing all of the HTML content.
 */
Display.prototype.resetDisplay = function() {
  this.display.html('');
};


/**
 * Show loading message on the user display.
 */
Display.prototype.showLoading = function() {
  this.loading.removeClass('hidden');
  this.loadingOverlay.removeClass('hidden');
};


/**
 * Hide loading message on the user display.
 */
Display.prototype.hideLoading = function() {
  this.loading.addClass('hidden');
  this.loadingOverlay.addClass('hidden');
};


/**
 * Update the user display to show an error message.
 *
 * @param {string} message The error message to display.
 * @param {string=} header The header to display. Default is 'Error'.
 */
Display.prototype.showError = function(message, header = 'Error') {
  var msg = '<h4 class="alert-header">' + header + '</h4>' +
      '<p>' + message + '</p>';
  this.error.html(msg).removeClass('hidden');
};


/**
 * Update the user display to hide the error mesage.
 */
Display.prototype.hideError = function() {
  this.error.addClass('hidden');
};


/**
 * Returns an HTML-formatted string to display the 'Close' button.
 * 
 * @returns {string} An HTML-formatted string.
 */
function closeButton() {
  button = '<input type="button" value="Close" class="btn" onclick="google.script.host.close();">';
  return button;
}