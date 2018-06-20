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
 * Runs dialog initialization and handles all dialog user inputs.
 */
$(function() {
  initialize();

  // Handle the merge type selectors
  var selectors = '#letters, #labels';
  $(document).on('click', selectors, function() {
    $(selectors).toggleClass('selected');
  });

  // Handle the save options button click
  $(document).on('click', '#optionsSave', function() {
    saveOptions_onClick();
  });

  // Handle the cancel button click
  $(document).on('click', '#optionsCancel', function() {
    cancelOptions_onClick();
  });

  // Handle the advanced options button click
  $(document).on('click', '#optionsAdvanced', function() {
    $('#advancedOptions').toggleClass('hidden');
    scrollTo('#advancedOptions');
  });
});


/**
 * Runs dialog initialization functions to retrieve and display the primary
 * user interface components for the dialog.
 */
function initialize() {
  google.script.run
      .withSuccessHandler(updateDisplay)
      .getOptionsDisplay();
}


/**
 * Handles the 'Save options' button click response.
 */
function saveOptions_onClick() {
  showLoading('Saving...');
  var options = {
    mergeType: $('#mergeType').children('.selected').attr('id'),
    outputFileName: $('#outputFileName').val(),
    outputFileType: $('#outputFileType option:selected').val(),
    numOutputFiles: $('input[name="numOutputFiles"]:checked').val(),
    tableWrapMerge: ($('#tableWrapMerge').prop('checked') === true ? 'enable' : 'disable'),
  };
  clearDisplay();
  google.script.run
      .withSuccessHandler(function() {
        google.script.host.close();
      })
      .saveMergeOptions(options);
}


/**
 * Handles the 'Cancel' button click response.
 */
function cancelOptions_onClick() {
  google.script.host.close();
}