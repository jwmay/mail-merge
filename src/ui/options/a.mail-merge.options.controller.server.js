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
 * Returns an array of DisplayObject instances for all options to be displayed
 * in the dialog. Includes a display for the save and cancel buttons.
 * 
 * @returns {DisplayObjects[]} An array of DisplayObject instances.
 */
function getOptionsDisplay() {
  setDefaultOptions();
  var displays = [
    getMergeTypeSelector(),
    getSaveOptionsDisplay()
  ];
  return displays;
}


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
function getMergeTypeSelector() {
  var storage = new PropertyStore();
  var mergeType = storage.getProperty('mergeType');
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
              'Letters will create a new copy of the template for each record' +
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
              'Labels require the template contain a single table with any' +
              'number of identical cells, where each cell represents a' +
              'different record in the spreadsheet.' +
            '</p>' +
          '</div>' +
        '</li>' +
      '</ul>' +
    '</div>';
  var display = getDisplayObject('card', content, 'mergeTypeSelector', 'bottom');
  return display;
}


/**
 * Returns an instance of DisplayObject containing the save and cancel buttons.
 * 
 * @returns {DisplayObject} A DisplayObject instance for the buttons.
 */
function getSaveOptionsDisplay() {
  var content = '' +
      '<div class="btn-bar">' +
        '<input type="button" class="btn action" value="Save options" ' +
          'id="optionsSave">' +
        '<input type="button" class="btn" value="Cancel" ' +
          'id="optionsCancel">' +
      '</div>';
  var dislpay = getDisplayObject('card', content, 'saveOptionsDisplay', 'bottom');
  return dislpay;
}


/**
 * Stores all of the options provided in options parameter. Options must be
 * stored as key/value pairs in an object where the option name is the key.
 * 
 * @param {object} options An object of key/value pairs where the option name
 *    is the key.
 */
function saveMergeOptions(options) {
  var storage = new PropertyStore();
  for (var option in options) {
    storage.setProperty(option, options[option]);
  }
}


/**
 * Checks if an option is saved and sets a default option if there is no
 * saved value.
 */
function setDefaultOptions() {
  var storage = new PropertyStore();

  // Set default merge type to 'letters' if no option is set
  var mergeType = storage.getProperty('mergeType');
  if (mergeType === null) mergeType = 'letters';
  storage.setProperty('mergeType', mergeType);
}