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
 * Returns a human-readable string representing
 * the current month. The string will be three
 * characters in length.
 *
 * @return {string} A three-character string
 *     representing the current month.
 */
function getCurrentMonth_() {
  var date = new Date();
  var month = date.getMonth();
  var months = ['Jan', 'Feb', 'Mar',
                'Apr', 'May', 'Jun',
                'Jul', 'Aug', 'Sep',
                'Oct', 'Nov', 'Dec'];
  return months[month];
}


/**
 * Returns an integer representing the number
 * of milliseconds between midnight of
 * January 1, 1970 and the specified date.
 *
 * @return {integer} The time in milliseconds
 *   between midnight of January 1, 1970 and
 *   the specified date.
 */
function getTime_() {
  var date = new Date();
  var time = date.getTime();
  return time;
}