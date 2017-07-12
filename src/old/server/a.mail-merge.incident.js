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
 * Base class for an Incident.
 *
 * @constructor
 * @param {integer} row The row in the response sheet where
 *     the incident report is located.
 * @param {array} data An array containing the incident data.
 */
var Incident = function(row, data) {
  this.report = new Reports();
  this.responses = new FormResponses();
  this.config = Configuration.getCurrent();
  
  this.row = row;
  this.pdfUrl = '';
  this.pdfFile = {};
  this.data = this.getData(data);
};


/**
 * Returns an object containing the incident data. Converts header values to 
 * header tags and stores the incident data with each of the respective keys.
 * 
 * @param {array} data An array containing the incident data.
 * @return {object} An object with the incident data stored using the
 *     header tags.
 */
Incident.prototype.getData = function(data) {
  var headerKeys = this.responses.getHeaderKeys();

  var incidentData = {};
  for (var i = 0; i < headerKeys.length; i++) {
    var key = headerKeys[i];
    var value = data[i];
    incidentData[key] = value;
  }

  return incidentData;
};


/**
 * Returns true if the incident report was sent.
 * 
 * @return {boolean} True if the incident report was sent, otherwise, false.
 */
Incident.prototype.isSent = function() {
  var reportStatusHeader = this.config.sheets.formResponses.headers[0];
  var reportStatusKey = reportStatusHeader.replace(/\s+/g, '_');
  var status = this.data[reportStatusKey];

  if (status === 'sent') {
    return true;
  }
  return false;
};


/**
 * Creates a report for the given incident in the current
 * month's report folder and marks the incdient as sent in
 * the form response sheet.
 */
Incident.prototype.createReport = function() {
  var templateFile = new TemplateFile();
  var reportFile = templateFile.copyTemplateFile();
  var reportFileDocument = DocumentApp.openById(reportFile.fileId);
  var reportFileBody = reportFileDocument.getBody();
  
  // Replace all occurances of the header tag with the incident data value.
  for (var header in this.data) {
    if (this.data.hasOwnProperty(header)) {
      var headerTag = '<<' + header + '>>';
      var headerValue = this.data[header];
      reportFileBody.replaceText(headerTag, headerValue);
    }
  }

  reportFileDocument.saveAndClose();

  var reportsFolder = new ReportsFolder();
  var destination = reportsFolder.getCurrentMonthFolder();
  var pdfFileName = this.getPdfFilename();
  var pdfFile = reportFile.createPdf(pdfFileName, destination);
  
  reportFile.file.setTrashed(true);
  
  this.setPdfUrl(pdfFile.getUrl());
  this.pdfFile = pdfFile;
  this.setReportStatus();
  this.emailReport();
};


/**
 * Sets the value of the 'Report Status' box to 'sent' for
 * the current incident.
 */
Incident.prototype.setReportStatus = function() {
  var responseSheet = this.responses.sheet;
  var reportStatusColumn = this.responses.getReportStatusColumn();
  var responseSheetSentCell = responseSheet.getRange(this.row,
          reportStatusColumn);
  responseSheetSentCell.setValue('sent').setHorizontalAlignment('center');
};


/**
 * Returns a string representing the PDF filename of the incident.
 * 
 * @return {string} The PDF filename.
 */
Incident.prototype.getPdfFilename = function() {
  var headers = this.responses.getHeaderKeys();
  var filenameString = this.report.getFilename();
  var filename = filenameString.split('**');

  var pdfFilename = [];
  for (var i = 0; i < filename.length; i++) {
    var name = filename[i];
    if (headers.indexOf(name) === -1) {
      pdfFilename.push(name);
    } else {
      pdfFilename.push(this.data[name]);
    }
  }

  return pdfFilename.join('');
};


/**
 * Sets the value of the 'PDF Link' box to the PDF file's url with hyperlink
 * for the current incident.
 *
 * @param {string} url The url of the PDF file.
 */
Incident.prototype.setPdfUrl = function(url) {
  this.pdfUrl = url;
  
  var formResponses = new FormResponses();
  var formResponseSheet = formResponses.sheet;
  var pdfLinkColumn = formResponses.getPdfLinkColumn();
  var formResponseUrlCell = formResponseSheet.getRange(this.row, pdfLinkColumn);
  formResponseUrlCell.setValue(url);
};


/**
 * Sends an email of the report to all email addresses stored in
 * document properties.
 */
Incident.prototype.emailReport = function() {
  var emailAddresses = new EmailAddresses();
  var emails = emailAddresses.getEmailAddresses();

  if (emails !== null) {
    for (var i = 0; i < emails.length; i++) {
      var recipient = emails[i];
      var subject = this.getPdfFilename();
      var body = this.getEmailBody();
      var options = {
        attachments: [this.pdfFile.getAs(MimeType.PDF)],
        htmlBody: body,
        name: 'Dean\'s Office Incident Reporter'
      };
      GmailApp.sendEmail(recipient, subject, body, options);
    }
  }
};


/**
 * Returns and HTML-formatted string containing the body content of the email.
 * 
 * @return {string} An HTML-formatted string.
 */
Incident.prototype.getEmailBody = function() {
  var message = [];

  message.push('<h2 style="color:green;font-family:helvetica,arial,sans-serif;font-weight:400;">' +
                 'Dean\'s Office Incident Report' +
               '</h2>');
  message.push('<p style="font-size:16px;font-weight:200;font-family:helvetica,arial,sans-serif;">' +
                 'Please find the attached incident report titled ' +
                 '<em>' + this.getPdfFilename() + '</em>. ' +
               '</p>');
  message.push('<p style="color:#a5a5a5;font-family:helvetica,arial,sans-serif;"><small><em>' +
                 'You are receiving this email because your email address was added to ' +
                 'the Mojave High School Dean\'s Office Incident Reporter. Please email ' +
                 '<a href="mailto:mayj3@nv.ccsd.net">mayj3@nv.ccsd.net</a> if you have ' +
                 'any questions or no longer wish to receive these emails.' +
               '</em></small></p>');

  return message.join('');
};