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
 * Removes all occurrances of the given element from the array. If the all flag
 * is set to false, only the first occurrance of the element will be removed.
 * This method does alter the array.
 * 
 * @param {string} element The element to be removed.
 * @param {boolean} all If true, all matched elements will be removed. If false,
 *         only the first matched element will be removed. Default is true.
 */
Array.prototype.removeElement = function(element, all) {
  var removeAll = all === undefined ? true : all;
  var index = this.indexOf(element);
  if (removeAll === true) {
    while (index !== -1) {
      this.splice(index, 1);
      index = this.indexOf(element);
    }
  } else {
    this.splice(index, 1);
  }
};


/**
 * Returns and array with all duplicate elements removed from the origianl
 * array. This method does not alter the original array.
 * 
 * @return {array} An array without any duplicate elements.
 */
Array.prototype.removeDuplicates = function() {
  var seen = {};
  var out = [];
  var len = this.length;
  var j = 0;
  for(var i = 0; i < len; i++) {
    var item = this[i];
    if(seen[item] !== 1) {
      seen[item] = 1;
      out[j++] = item;
    }
  }
  return out;
};


/**
 * Returns the array with all empty, null, and undefined elements removed. This
 * does not modify the original array.
 * 
 * @returns An array with no empty, null, or undefined elements.
 */
Array.prototype.removeEmpty = function() {
  var cleanedArray = this.filter(function(n) {
      return (n !== undefined && n !== null && n !== '');
  });
  return cleanedArray;
};


/**
 * Returns the array of strings with each element converted to lower case. This
 * does not modify the origianl array.
 * 
 * @returns An array of strings converted to lower case.
 */
Array.prototype.toLowerCase = function() {
  var lowerCaseArray = this.map(
    function(x) {
      return x.toLowerCase();
    }
  );
  return lowerCaseArray;
};


/**
 * Compares two arrays to determine if they are equal. If sort is set to true,
 * the arrays are sorted before comparison. The default sort is false, so the
 * order of the arrays is taken into account during the comparison.
 * 
 * @param {array} array1 The first array to compare.
 * @param {array} array2 The second array to compare.
 * @param {boolean=} sort True if the arrays are to be sorted.
 *    Default is false, the array order does matter.
 * @returns 
 */
function arraysEqual(array1, array2, sort) {
  var a = array1;
  var b = array2;

  if (a === b) return true;
  if (a === null || b === null) return false;
  if (a.length != b.length) return false;

  // If sort set to true, then the arrays are sorted
  // and compared, otherwise, the arrays will be compared
  // taking into account the order of the elements.
  var doSort = sort === undefined ? false : sort;
  if (doSort) {
    a.sort();
    b.sort();
  }

  for (var i = 0; i < a.length; ++i) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}