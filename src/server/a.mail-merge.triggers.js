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
 * Base class for report generation triggers.
 */
var Triggers = function() {
  this.defaultTrigger = 'manual';
  this.trigger = this.getTrigger();
};


/**
 * Returns the report generation trigger as 'automatic' or 'manual'.
 * 
 * @return {string} The current trigger.
 */
Triggers.prototype.getTrigger = function() {
  var storage = new PropertyStore();
  var trigger = storage.getProperty('REPORT_TRIGGER');
  if (trigger !== undefined && trigger !== null && trigger !== '') {
    return trigger;
  }
  return this.defaultTrigger;
};


/**
 * Stores the report trigger in the document properties.
 * 
 * @param {string} trigger The trigger to be stored.
 */
Triggers.prototype.setTrigger = function(trigger) {
  var storage = new PropertyStore();
  storage.setProperty('REPORT_TRIGGER', trigger);
};


/**
 * Activates the currently selected trigger and returns the trigger name as
 * either 'automatic' or 'manual'.
 * 
 * @return {string} The current trigger.
 */
Triggers.prototype.activateCurrentTrigger = function() {
  var trigger = this.getTrigger();
  var scriptTriggers = ScriptApp.getProjectTriggers();
  
  if (trigger === 'automatic') {
    this.enableFormResponseTrigger();
  }
  if (trigger === 'manual') {
    this.disableTrigger('generateReports');
  }
  return this.trigger;
};


/**
 * Adds a trigger to generate reports on each form submission.
 */
Triggers.prototype.enableFormResponseTrigger = function() {
  try {
    ScriptApp.newTrigger('generateReports')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onFormSubmit()
      .create();
  } catch(e) {
    showAlert('Installable trigger error', '[enableFormResponseTrigger] ' +
            'Installable triggers cannot be tested.\n' + e);
  }
};


/**
 * Removes the trigger with the given trigger function name.
 * 
 * @param {string} triggerFunctionName The name of the function the trigger
 *     calls.
 */
Triggers.prototype.disableTrigger = function(triggerFunctionName) {
  var scriptTriggers = ScriptApp.getProjectTriggers();
  try {
    for (var i = 0; i < scriptTriggers.length; i++) {
      var trigger = scriptTriggers[i];
      if (trigger.getHandlerFunction() === triggerFunctionName) {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  } catch(e) {
    showAlert('Installable trigger error', '[deleteTrigger] Installable ' +
            'triggers cannot be tested.\n' + e);
  }
};