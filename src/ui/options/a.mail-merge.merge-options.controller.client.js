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
 * Modal window initialization.
 */
$(function() {
  initialize();

  // Handle the merge type selectors.
  var selectors = '#letters, #labels';
  $(document).on('click', selectors, function() {
    $(selectors).toggleClass('selected');
  });

  // Handle the save options button click.
  $(document).on('click', '#optionsSave', function() {
    saveOptions_onClick();
  });

  // Handle the cancel button click.
  $(document).on('click', '#optionsCancel', function() {
    cancelOptions_onClick();
  });
});


/**
 * Dialog initialization and user-response functions.
 */
function initialize() {
  google.script.run
    .withSuccessHandler(updateDisplay)
    .getOptionsDisplay();
}


/**
 * Handle the 'Save options' button click response.
 */
function saveOptions_onClick() {
  var options = {};
  options.mergeType = $('#mergeType').children('.selected').attr('id');
  google.script.run.saveMergeOptions(options);
  google.script.host.close();
}


/**
 * Handle the 'Cancel' button click response.
 */
function cancelOptions_onClick() {
  google.script.host.close();
}