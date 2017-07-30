/**
 * Returns a display object to be used in the Display class.
 *
 * Two types of display types can be generated: an alert-{type} or a card,
 * where {type} can be success, warning, error, or information. The content
 * of the display can be an HTML-formatted string or just a plain string.
 * An optional id for a card can be specified. Cards can be added to the top
 * or bottom of other cards with the optional position argument. By default
 * cards are positioned at the top. An option reset flag is available to clear
 * the display of all alerts and cards before updating the display. A flag to
 * close the current display is also available.
 *
 * @param {string} type The display type, either alert-{type} or card,
 *         where {type} can be success, warning, error, or information.
 * @param {string} content The content to display.
 * @param {string=} id The id of the card. Default is none.
 * @param {string=} position The position of the card. Either top or bottom.
 *         Default is top.
 * @param {boolean=} reset Clear the display of all alerts and cards before
 *         updating. Default is false.
 * @param {boolean=} close Close the current display. Default is false.
 * @return {object} The display object.
 */
var getDisplayObject = function(type, content, id, position, reset, close) {
  // Assign default values as Google's server 'gs' does not support
  // ES6 default parameter values in the function definition.
  id = undefined === id ? '' : id;
  position = undefined === position ? 'top' : position;
  reset = undefined === reset ? false : reset;
  close = undefined === close ? false : close;
  
  var displayObject = {
    type: type,
    content: content,
    id: id,
    position: position,
    reset: reset,
    close: close,
  };
  return displayObject;
};


/**
 * Returns the content of the given file to be displayed in an HTML template.
 *
 * @param {string} file The file name to include.
 * @return {string} The file content.
 */
function include(file) {
  var content = HtmlService.createHtmlOutputFromFile(file).getContent();
  return content;
}