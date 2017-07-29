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
 * @param {boolean=} prepend If true, the content will be added to the top of
 *        the display, otherwise, it will be added to the bottom. Default is
 *        true.
 * @param {boolean=} reset Clear the display before updating if true,
 *         otherwise, append the content to the display. Default is false.
 */
function updateDisplay(content, prepend, reset) {
  var display = new Display();
  display.hideLoading();
  display.updateDisplay(content, prepend, reset);
}


/**
 * Helper function for updating the user-interface display with a success 
 * message.
 * 
 * @param {string} message The error message to display.
 * @param {boolean=} close Flag to display close button. Default is True.
 */
function showSuccess(message) {
  var display = new Display();
  display.showSuccess(message, close);
}


/**
 * Helper function for clearing a success message from the user-interface
 * display.
 */
function hideSuccess() {
  var display = new Display();
  display.hideSuccess();
}


/**
 * Helper function for updating the user-interface display with an error 
 * message.
 * 
 * @param {string} message The error message to display.
 * @param {string=} header The header to display. Default is 'Error'.
 * @param {boolean=} close Flag to display close button. Default is True.
 */
function showError(message, header, close) {
  var display = new Display();
  display.showError(message, header, close);
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
 * Returns an HTML-formatted string to display the 'Close' button.
 * 
 * @returns {string} An HTML-formatted string.
 */
function closeButton() {
  button = '<input type="button" value="Close" class="btn" onclick="google.script.host.close();">';
  return button;
}


/**
 * Base class for the HTML display on the sidebar.
 */
var Display = function() {
  this.display = $('#display');
  this.success = $('#success');
  this.error = $('#error');
  this.loading = $('#loading');
  this.loadingOverlay = $('#loading-overlay');
};


/**
 * Update the HTML of the display to show the given content.
 * 
 * @param {string} content The HTML content to display.
 * @param {boolean=} prepend If true, the content will be added to the top of
 *        the display, otherwise, it will be added to the bottom. Default is
 *        true.
 * @param {boolean=} reset Clear the display before updating if true,
 *         otherwise, append the content to the display. Default is false.
 */
Display.prototype.updateDisplay = function(content, prepend=true, reset=false) {
  if (reset === true) this.resetDisplay();
  if (prepend) {
    $(content).hide().prependTo(this.display).slideDown('slow');
  } else {
    $(content).hide().appendTo(this.display).slideDown('slow');  
  }
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
 * Update the user display to show a success message. An optional close button
 * is also displayed by default.
 *
 * @param {string} message The success message to display.
 * @param {boolean=} close Display a close button. Default is True.
 */
Display.prototype.showSuccess = function(message, close=true) {
  var success = '<div class="alert alert-success border-0">' +
            message +
          '</div>';
  if (close) success += closeButton();
  this.success.html(success).removeClass('hidden');
};


/**
 * Update the user display to hide the success mesage.
 */
Display.prototype.hideSuccess = function() {
  this.success.html('').addClass('hidden');
};


/**
 * Update the user display to show an error message. An optional header can be
 * specified with the default being 'Error'. An optional close button is also
 * displayed by default.
 *
 * @param {string} message The error message to display.
 * @param {string=} header The header to display. Default is 'Error'.
 * @param {boolean=} close Display a close button. Default is True.
 */
Display.prototype.showError = function(message, header='Error', close=true) {
  var error = '<div class="alert alert-error border-0">' +
            '<h4 class="alert-header">' + header + '</h4>' +
            '<p>' + message + '</p>' +
          '</div>';
  if (close) error += closeButton();
  this.error.html(error).removeClass('hidden');
};


/**
 * Update the user display to hide the error mesage.
 */
Display.prototype.hideError = function() {
  this.error.html('').addClass('hidden');
};