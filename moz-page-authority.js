/** 
 * @file MOZ Page Authority and Link Equity Google Spreadsheet Script. To use this script place your fully qualified URLs in the first column of your spreadsheet. Fill out the variables below with your own information and run the script from the spreadsheet's menu. 
 * @version 1.0  
 *
 * @author Matthew Rosenberg [matt@dynamicwebsolutions.com]
 * @copyright 2014 Dynamic Web Solutions
 * @license GPL-2.0+
 *
 * This program is free software; you can redistribute it and/or 
 * modify it under the terms of the GNU General Public License 
 * as published by the Free Software Foundation; either version 2 
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful, 
 * but WITHOUT ANY WARRANTY; without even the implied warranty of 
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the 
 * GNU General Public License for more details.
 */



/*********************************************
 *                                           *
 * You need to fill out the varaibles below. *
 *                                           *
 *********************************************/

    /** 
     * Your MOZ Access ID Number. For more information about who to get credentials visit {@link https://moz.com/products/api/keys}.
     * @type {string} 
     */
var accessId    = '',

    /** 
     * MOZ secret key used to salt the hash.
     * @type {string}  
     */
    secret      = 'your-secret',

    /** 
     * The number of rows the script should skip from the top. 
     * This allows you to set headers and sheet titles. 
     * @type {number} 
     */        
    rowOffset   = 3,

    /** 
     * The column where you would like to write your page authority data. 
     * @type {string} 
     */
    pageAuthorityColumn       = '',
    domainAuthorityColumn       = '',

    /** 
     * The column where you would like to write your link equity data.
     * @type {string}  
     */
    externalEquityLinksColumn = '',  
    externalDomainLinksColumn = '';



/***********************************************
 *                                             *
 * Stop! No need to touch anything below here. *
 *                                             *
 ***********************************************/    


    /** 
     * MOZ API endpoint
     * @type {string}  
     */    
var apiBase     = 'https://lsapi.seomoz.com/linkscape/',  

    /** 
     * We'll use this later to hold data.
     * @type {array}  
     */
    requestURLs = [];  


/**
 * @function mozAuthenticate
 * @summary Generates the needed MOZ API security token.
 * Moz documentation: {@link http://apiwiki.moz.com/signed-authentication}
 * 
 * @returns {string} 
 */
function mozAuthenticate() {
  var expiry, key, signature;
  
  /** 
   * A UNIX formatted number a few seconds in the future. 
   * @type {number} 
   */
  expiry = Math.round(new Date().getTime() / 1000) + 300;

  /**
   * The new line character is required to generate a proper hash value. 
   * @type {string} 
   */
  key = this.accessId +"\n"+ expiry;

  /** 
   * The secret must be passed in as a salt to the algorithm.    
   * Some examples show it being included in the key which will not work. 
   * Proper usage {@link https://developers.google.com/apps-script/reference/utilities/utilities#computeHmacSignature(MacAlgorithm,String,String)}.
   */
  signature = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, key, this.secret);

  return 'AccessID='+this.accessId+'&Expires='+expiry+'&Signature='+encodeURIComponent(Utilities.base64Encode(signature));  
}


/**
 * @function mozRequest
 * @summary Generates the request to MOZ's api endpoint
 * Moz documentation: {@link http://apiwiki.moz.com/url-metrics}
 */
function mozRequest() {
  var self, bitFlags, url, requests;
  
  self = this;

  /**
   * The sum of the MOZ Page Authority & External Equity Links bitwise identifiers.
   * @type {number}  
   */
  bitFlags = (34359738368+32+68719476736+1024);

  /** 
   * Our completed URL including authentication. 
   * @type {string} 
   */
  url = this.apiBase + 'url-metrics/?Cols=' + bitFlags + '&' + this.mozAuthenticate();

  /**
   * Free access to the MOZ API is limited to one request every ten seconds.
   * You are allowed to batch up to ten URLs in one request. 
   * @type {string[]} 
   */
  requests = this.chunkRequests();    
  
  /** Loop over request arrays. */
  requests.forEach(function(request, i) {
    var payload, options, response;
    
    /** 
     * The URLs contained in the individual arrays.
     * @type {string[]}  
     */
    payload = request.map(function(request) { return request.url });
    
    /**
     * Setups up the XHR request. 
     * @type {object} 
     */
    options = {
      'method': 'post',
      'payload': JSON.stringify(payload)
    };
  
    /**
     * The results of our request to MOZ.
     * @type {object[]}  
     */
    response = UrlFetchApp.fetch(url, options);
    
    /** Parse the JSON reponse from MOZ and add the values back to the request object. */
    Utilities.jsonParse(response.getContentText()).forEach(function(result, index) {
      request[index]['pageAuthorityValue'] = result.upa;
      request[index]['domainAuthorityValue'] = result.pda;
      request[index]['externalEquityLinksValue'] = result.ueid;
      request[index]['externalDomainLinksValue'] = result.uipl;
    });    
    
    /** 
      * Write the data back to our spreadsheet.
      * @this Is cached to self so we can access proeprties outside the loop.
     */    
    self.writeRows(request);
    
    /** 
     * Free access to the MOZ API is rate limited to one request every ten seconds. 
     * Google Apps Script's sleep utility is set to eleven seconds, just to be safe. 
     * {@link https://developers.google.com/apps-script/reference/utilities/utilities#sleep(Integer)} 
    */
    Utilities.sleep(11000);
  });
}


/**
 * @function chunkRequests
 * @summary Breaks down large requests into groups of ten.
 * Moz documentation: {@link http://apiwiki.moz.com/free-vs-paid-access}
 */
function chunkRequests() {
  var i, j, 
      chunk  = 10,
      result = [];
  
  for (i = 0, j = this.requestURLs.length; i < j; i+=chunk) {
    result.push(this.requestURLs.slice(i, i + chunk));
  }

  return result;
}


/**
 * @function writeRows
 * @summary Writes the data back to the spread sheet.
 * Moz documentation: {@link http://apiwiki.moz.com/free-vs-paid-access}
 *
 * @param {object[]} data Row data to write back to the spreadsheet.
 */
function writeRows(data) {
  var doc; 

  /** @type {object} */
  doc = SpreadsheetApp.getActiveSpreadsheet();
  
  data.forEach(function(row) {

    /** Write the data. */
    doc.getRange(row.pageAuthorityCell).setValue(row.pageAuthorityValue);
    doc.getRange(row.domainAuthorityCell).setValue(row.domainAuthorityValue);
    doc.getRange(row.externalEquityLinksCell).setValue(row.externalEquityLinksValue);
    doc.getRange(row.externalDomainLinksCell).setValue(row.externalDomainLinksValue);
  });
}


/**
 * @function readRows
 * @summary Writes the data back to the spread sheet.
 * Spreadsheet API documentation: {@link https://developers.google.com/apps-script/service_spreadsheet}
 */ 
function readRows() {
  var sheet, rows, numRows, values;
  
  /** @type {object} */
  sheet   = SpreadsheetApp.getActiveSheet();

  /** @type {object} */
  rows    = sheet.getDataRange();

  /** @type {number} */
  numRows = rows.getNumRows();

  /** @type {string[]} */  
  values  = rows.getValues();

  for (var i = this.rowOffset; i <= numRows - 1; i++) {
    var row, rowNumber;
    
    /** @type {string[]} */  
    row = values[i];

    /** 
     * Advanced by one. Remember 'i' starts at zero. 
     * @type {number} 
     */  
    rowNumber = (i + 1);

    /** Create the object and push it into the global varaible. */
    this.requestURLs.push({
      'url': row[0],
      'pageAuthorityCell': this.pageAuthorityColumn+rowNumber,
      'domainAuthorityCell': this.domainAuthorityColumn+rowNumber,
      'externalEquityLinksCell': this.externalEquityLinksColumn+rowNumber,
      'externalDomainLinksCell': this.externalDomainLinksColumn+rowNumber,      
      'pageAuthorityValue': null,
      'domainAuthorityValue': null,
      'externalEquityLinksValue': null,
      'externalDomainLinksValue': null
    });
  }
  
  /** Fire off the requests to MOZ! */
  this.mozRequest();
};


/**
 * @function onOpen
 * @summary Adds a custom menu to the active spreadsheet.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * {@link https://developers.google.com/apps-script/service_spreadsheet}
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Page Authority & Link Equity",
    functionName : "readRows"
  }];
  spreadsheet.addMenu("MOZ API", entries);
};
