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
 * Returns an array of DisplayObject instances for creating the options dialog.
 * 
 * This is a helper function called from the client by google.script.run.
 * 
 * @returns {DisplayObjects[]} An array of DisplayObject instances.
 */
function getOptionsDisplay() {
  var optionsDisplay = new OptionsDisplay();
  return optionsDisplay.getOptionsDisplay();
}


/**
 * Stores all of the options provided in options parameter. Options must be
 * stored as key/value pairs in an object where the option name is the key.
 * 
 * This is a helper function called from the client by google.script.run.
 * 
 * @param {object} options An object of key/value pairs where the option name
 *    is the key.
 */
function saveMergeOptions(options) {
  var optionsDisplay = new OptionsDisplay();
  optionsDisplay.saveMergeOptions(options);
}



/**
 * Class for creating the options dialog display.
 */
var OptionsDisplay = function() {
  this.options = new Options();
};


/**
 * Returns an instance of DisplayObject containing the advanced options.
 * 
 * @returns {DisplayObject} A DisplayObject instance for the advanced options.
 */
OptionsDisplay.prototype.getAdvancedOptionsDisplay = function() {
  var options = this.options.getOptions();
  var selected = {
    filename: (options.outputFileName !== '' ? options.outputFileName: ''),
    googleDoc: (options.outputFileType === 'google_doc' ? 'selected' : ''),
    pdf: (options.outputFileType === 'pdf' ? 'selected' : ''),
    singleFile: (options.numOutputFiles === 'single' ? 'checked' : ''),
    multiFile: (options.numOutputFiles === 'multi' ? 'checked' : ''),
    tableWrapMerge: (options.tableWrapMerge === 'enable' ? 'checked' : ''),
  };
  var content = '' +
      '<div id="advancedOptions" class="option hidden">' +
        '<form>' +
          '<h3>Advanced Options</h3>' +
          '<div class="form-group">' +
            '<label for="outputFileName">' +
              'Output File Name' +
              '<span class="help-text">Specify an output file name to override the default shown below.</span>' +
            '</label>' +
            '<input type="text" class="form-control" id="outputFileName" placeholder="[Merge Output] Simply Mail Merge - Dev - Letter Template with Header" value="' + selected.filename + '">' +
          '</div>' +
          '<div class="form-row">' +
            '<div class="form-group">' +
              '<label for="outputFileType">' +
                'Output File Type' +
                '<span class="help-text">Select the type of output file to generate.</span>' +
              '</label>' +
              '<select class="form-control" id="outputFileType">' +
                '<option value="google_doc" ' + selected.googleDoc + '>Google Doc</option>' +
                '<option value="pdf" ' + selected.pdf + '>PDF</option>' +
              '</select>' +
            '</div>' +
            '<div class="form-group">' +
              '<h4 class="form-check-header">Output Files</h4>' +
              '<span class="help-text">Select the number of output files to generate.</span>' +
              '<div class="form-check">' +
                '<input class="form-check-input" type="radio" name="numOutputFiles" id="numOutputFilesSingle" value="single" ' + selected.singleFile + '>' +
                '<label class="form-check-label" for="numOutputFilesSingle">' +
                  'Single file' +
                  '<span class="help-text">One file will be created with all records</span>' +
                '</label>' +
              '</div>' +
              '<div class="form-check">' +
                '<input class="form-check-input" type="radio" name="numOutputFiles" id="numOutputFilesMulti" value="multi" ' + selected.multiFile + '>' +
                '<label class="form-check-label" for="numOutputFilesMulti">' +
                  'Multiple files' +
                  '<span class="help-text">One file will be created for each record</span>' +
                '</label>' +
              '</div>' +
            '</div>' +
          '</div>' +
          '<div class="form-group">' +
            '<h4 class="form-check-header">Letter Merge Method</h4>' +
            '<span class="help-text">' +
              'Uncheck to disable this method of speeding up letter merges, which places the ' +
              'output into tables with no padding or border. Disabling this should only be needed ' +
              'when the formatting of the output document fails to match that of the template document.' +
            '</span>' +
            '<div class="form-check">' +
              '<input class="form-check-input" type="checkbox" value="tableWrapMergeEnable" id="tableWrapMerge" ' + selected.tableWrapMerge + '>' +
              '<label class="form-check-label" for="tableWrapMerge">' +
                'Speed up letter merge' +
              '</label>' +
            '</div>' +
          '</div>' +
        '</form>' +
      '</div>';
  return getDisplayObject('card', content, 'advancedOptionsDisplay', 'bottom');
};


/**
 * Returns an instance of DisplayObject containing the option for selecting the
 * merge type.
 * 
 * The available merge types include 'letters' and 'labels'. Letters will create
 * a copy of each page in the template document for each record in the source
 * spreadsheet. Labels uses a table in the template document to place multiple
 * records on the same page; each record will be placed in a cell of the table.
 * 
 * @returns {DisplayObject} A DisplayObject instance for the selector.
 */
OptionsDisplay.prototype.getMergeTypeSelector = function() {
  var mergeType = this.options.getOption('mergeType');
  var selected = {
    letters: (mergeType == 'letters' ? 'selected' : ''),
    labels: (mergeType == 'labels' ? 'selected' : '')
  };
  var content = '' +
      '<div class="option">' +
      '<h4>Merge Type</h4>' +
      '<ul id="mergeType" class="block-selectors">' +
        '<li id="letters" class="' + selected.letters + '">' +
          '<div class="selector-icon">' +
            '<i class="fa fa-file-text fa-3x"></i>' +
          '</div>' +
          '<div class="selector-content">' +
            '<h4 class="selector-title">Letters</h4>' +
            '<p class="selector-description">' +
              'Letters will create a new copy of the template for each record ' +
              'in the spreadsheet.' +
            '</p>' +
          '</div>' +
        '</li>' +
        '<li id="labels" class="' + selected.labels + '">' +
          '<div class="selector-icon">' +
            '<i class="fa fa-tags fa-3x"></i>' +
          '</div>' +
          '<div class="selector-content">' +
            '<h4 class="selector-title">Labels</h4>' +
            '<p class="selector-description">' +
              'Labels require the template contain a single table with any ' +
              'number of identical cells, where each cell represents a ' +
              'different record in the spreadsheet.' +
            '</p>' +
          '</div>' +
        '</li>' +
      '</ul>' +
    '</div>';
  return getDisplayObject('card', content, 'mergeTypeSelector', 'bottom');
};


/**
 * Returns an array containing all of the DisplayObject instances for creating
 * the options dialog display.
 * 
 * @returns {DisplayObject[]} An array of DisplayObject instances.
 */
OptionsDisplay.prototype.getOptionsDisplay = function() {
  return [
    this.getMergeTypeSelector(),
    this.getAdvancedOptionsDisplay(),
    this.getSaveOptionsDisplay(),
  ];
};


/**
 * Returns an instance of DisplayObject containing the save and cancel buttons.
 * 
 * @returns {DisplayObject} A DisplayObject instance for the buttons.
 */
OptionsDisplay.prototype.getSaveOptionsDisplay = function() {
  var content = '' +
      '<div class="btn-bar">' +
        '<input type="button" class="btn action" value="Save options" ' +
            'id="optionsSave">' +
        '<input type="button" class="btn" value="Cancel" id="optionsCancel">' +
        '<input type="button" class="btn btn-link" value="Advanced options" ' +
            'id="optionsAdvanced">' +
      '</div>';
  var dislpay = getDisplayObject('card', content, 'saveOptionsDisplay', 'bottom');
  return dislpay;
};


/**
 * Stores all of the options provided in options parameter. Options must be
 * stored as key/value pairs in an object where the option name is the key.
 * 
 * @param {object} options An object of key/value pairs where the option name
 *    is the key.
 */
OptionsDisplay.prototype.saveMergeOptions = function(options) {
  for (var option in options) {
    this.options.setOption(option, options[option]);
  }
};