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
 * Sidebar initialization and user-response functions to load the sidebar.
 */
$(function() {
  initializeSidebar();

  // Handle the 'x' click response on email addresses.
  // $(document).on('click', '.fa-close', function() {
  //   $(this).removeClass('fa-close').addClass('fa-spinner fa-spin');
  //   var email = $(this).parents('li').data('email');
  //   google.script.run
  //     .withSuccessHandler(updateEmailAddresses)
  //     .removeEmailAddress(email);
  // });
});


/**
 * Runs sidebar initialization functions to retrieve and display stored data.
 */
function initializeSidebar() {
  google.script.run
    .withSuccessHandler(updateDisplay)
    .getDataSpreadsheetDisplay();
}


/**
 * Handle the selectFile button click response.
 */
function selectFile_onclick() {
  google.script.run
    .withSuccessHandler(updateDisplay)
    .setDataSpreadsheetFile();
}


/**
 * Updates the email address display on the sidebar.
 * 
 * @param {string} emails An HTML-formatted string containing the list of email
 *         addresses to be displayed.
 */
function updateEmailAddresses(emails) {
  $('#emails').html(emails);
}


/**
 * Handle the addEmails button click response.
 */
function addEmails_onclick() {
  google.script.run
    .withSuccessHandler(updateEmailAddresses)
    .addEmailAddresses();
}


/**
 * Handle the generateReports button click response.
 */
function generateReports_onclick() {
  var $button = $('#generate');
  $button.prop('disabled', true);
  $('.loading').show();
  google.script.run
    .withSuccessHandler(
      function() {
        $button.prop('disabled', false);
        $('.loading').hide();
      }
    )
    .generateReports();
}