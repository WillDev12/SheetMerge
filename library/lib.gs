/**
 * Displays a dialog with a QR Code based on the provided text.
 * @param {string}  text - The text to be converted into a QR Code.
 */
function generateQRCodeDialog(text) {
  var htmlOutput = HtmlService.createHtmlOutput('<img src="' + generateQRCodeUrl(text) + '">')
    .setWidth(200)
    .setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'QR Code');
}

/**
 * Generates the QR Code URL based on the provided text.
 * @param {string}  text - The text to be converted into a QR Code.
 * @returns {string} chartUrl - The generated QR Code URL.
 */
function generateQRCodeUrl(text) {
  var encodedText = encodeURIComponent(text);
  var chartUrl = 'https://chart.googleapis.com/chart?cht=qr&chs=200x200&chl=' + encodedText;
  Logger.log(chartUrl)
  return chartUrl;
}

/**
 * Generates the base64-encoded image of the QR Code based on the provided text and logs it.
 * @param {string}  text - The text to be converted into a QR Code.
 */
function generateQRCodeAndLogBase64(text) {
  var qrCodeBase64 = generateQRCodeBase64(text);
  Logger.log("QR Code Base64: " + qrCodeBase64);
}

/**
 * Generates the base64-encoded image of the QR Code based on the provided text.
 * @param {string}  text - The text to be converted into a QR Code.
 * @returns {string} qrCodeBase64 - The base64-encoded image of the generated QR Code.
 */
function generateQRCodeBase64(text) {
  var encodedText = encodeURIComponent(text);
  var chartUrl = 'https://chart.googleapis.com/chart?cht=qr&chs=200x200&chl=' + encodedText;
  var imageUrl = UrlFetchApp.fetch(chartUrl).getBlob().getAs('image/png');
  var qrCodeBase64 = Utilities.base64Encode(imageUrl.getBytes());
  return qrCodeBase64;
}
/**
 * Displays a pop-up/alert in the spreadsheet.
 * @param {string} message - The desired pop-up message.
 * @param {string} status - The status of the message. You can replace it with "warning", "danger", or "success".
 */
function alertMaker(message, status) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(status, message, ui.ButtonSet.OK);
}

/**
 * Spins and replaces text within curly braces.
 * @param {string} text - The text to be spun and replaced.
 * @returns {string} - The spun and replaced text.
 */
function spinText(text) {
  var spunText = text.replace(/{([^{}]*)}/g, function(match, choices) {
    var choicesArray = choices.split("|");
    var randomIndex = Math.floor(Math.random() * choicesArray.length);
    return choicesArray[randomIndex];
  });
  return spunText;
}

/**
 * Removes duplicate rows from the current sheet.
 */
function removeDuplicates() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const uniqueData = {};
  for (let row of data) {
    const key = row.join();
    uniqueData[key] = uniqueData[key] || row;
  }
  sheet.clearContents();
  const newData = Object.values(uniqueData);
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

/**
 * Checks the status of a website URL or server (online or offline).
 * @param {string} url - The URL of the website to be checked.
 * @returns {boolean} - true for online, false for offline.
 */
function checkServer(url) {
  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var responseCode = response.getResponseCode();
    
    if (responseCode == 200) {
      return true; // Server online
    } else {
      return false; // Server offline or invalid response
    }
  } catch (error) {
    return false; // Error occurred while pinging
  }
}

/**
 * Sets the value of a column in a spreadsheet without linking to the sheet.
 * @param {string} worksheet - Your worksheet name.
 * @param {string} cellcolomn - The column you want to change.
 * @param {string} value - The value to set.
 */
function setValue(worksheet, cellcolomn, value) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(worksheet);
  var cell = sheet.getRange(cellcolomn);
  cell.setValue(value);
}

/**
 * Converts the spreadsheet to a PDF file and saves it in Google Drive.
 * @param {string} spreadsheetId - The ID of the spreadsheet.
 * @param {string} sheetName - The name of the sheet to convert.
 * @param {string} folderId - The ID of the folder to save the PDF file in.
 * @returns {string} idfile - The ID of the file used for downloading the file.
 */
function convertSpreadsheetToPDF(spreadsheetId, sheetName, folderId) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?exportFormat=pdf";
  url += "&format=pdf";
  url += "&gid=" + sheet.getSheetId();
  url += "&size=A4";
  url += "&portrait=true";
  url += "&scale=4";
  url += "&sheetnames=false&printtitle=false&pagenumbers=false";
  url += "&gridlines=false";
  url += "&fzr=false";

  var pdfBlob = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    }
  }).getBlob();

  var timestamp = new Date().getTime();
  var pdfFileName = sheetName + "_" + timestamp + ".pdf";

  var folder = DriveApp.getFolderById(folderId);
  var pdfFile = folder.createFile(pdfBlob).setName(pdfFileName);

  var idfile = pdfFile.getId();
  return idfile;
}
function copySpreadsheet(newName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var newSpreadsheet = spreadsheet.copy(newName);
  var newSpreadsheetUrl = newSpreadsheet.getUrl();
  return newSpreadsheetUrl;
}
function copySpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var newSpreadsheet = spreadsheet.copy("Copy of " + spreadsheet.getName());
  var newSpreadsheetUrl = newSpreadsheet.getUrl();
  return newSpreadsheetUrl;
}
function getSheetUrl() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUrl = spreadsheet.getUrl();
  return sheetUrl;
}
function createSpreadsheet(name) {
    return SpreadsheetApp.create(name).getUrl();
}

function getActiveSheet(link) {
    var cache = CacheService.getScriptCache();
    var ss = SpreadsheetApp.openByUrl(link);
    return ss.getActiveSheet().getName();
}

function getSpreadsheetName(link) {
    return SpreadsheetApp.openByUrl(link).getName();
}

function getSpreadsheetID(link) {
    return SpreadsheetApp.openByUrl(link).getId();
}

function getSheetId(link) {
    return SpreadsheetApp.openByUrl(link).getActiveSheet().getSheetId();
}

function getSheetName(link) {
    return SpreadsheetApp.openByUrl(link).getActiveSheet().getSheetName();
}

function getURL(id) {
    return SpreadsheetApp.openById(id).getUrl();
}

function getViewers(link) {
    return SpreadsheetApp.openByUrl(link).getViewers();
}

function getOwner(link) {
    return SpreadsheetApp.openByUrl(link).getOwner();
}

function getA1(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getA1Notation();
}

function getCellBorder(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getBorder();
}

function getFontColor(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getFontColor();
}

function getFontSize(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getFontSize();
}

function getCellRow(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getRow();
}

function getCellColumn(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getColumn();
}

function getFontStyle(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getFontStyle();
}

function getFontFamily(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getFontFamily();
}

function getFontWeight(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getFontWeight();
}

function getFormula(range, link) {
    if (SpreadsheetApp.openByUrl(link).getRange(range).getFormula() == "") {
        return "none";
    } else {
        return SpreadsheetApp.openByUrl(link).getRange(range).getFormula();
    }
}

function getLastColumn(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getLastColumn();
}

function getLastRow(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getLastRow();
}

function getValue(range, link) {
    return SpreadsheetApp.openByUrl(link).getRange(range).getValue();
}

function setCell(cell, text, link) {
    var ss = SpreadsheetApp.openByUrl(link);
    var sheet = ss.getActiveSheet();
    sheet.getRange(cell).setValue(text);
}

function setCellFontSize(cell, size, link) {
    var ss = SpreadsheetApp.openByUrl(link);
    var sheet = ss.getActiveSheet();
    sheet.getRange(cell).setFontSize(size);
}

function clearRange(range, link) {
    var ss = SpreadsheetApp.openByUrl(link);
    var sheet = ss.getActiveSheet();
    sheet.getRange(range).clear();
}

function deleteRow(rowNum, link) {
    SpreadsheetApp.openByUrl(link).deleteRow(rowNum);
}

function deleteRows(row, howMany, link) {
    SpreadsheetApp.openByUrl(link).deleteRows(row, howMany);
}

function deleteColumn(columnNum, link) {
    SpreadsheetApp.openByUrl(link).deleteColumn(columnNum);
}

function deleteColumns(column, howMany, link) {
    SpreadsheetApp.openByUrl(link).deleteColumns(column, howMany);
}

function setRowHeight(row, height, link) {
    SpreadsheetApp.openByUrl(link).setRowHeight(row, height);
}

function setColumnWidth(column, width, link) {
    SpreadsheetApp.openByUrl(link).setColumnWidth(column, width);
}

function setCellBorder(range, top, left, bottom, right, link) {
    SpreadsheetApp.openByUrl(link).getRange(range).setBorder(top, left, bottom, right);
}

function setFontColor(range, color, link) {
    SpreadsheetApp.openByUrl(link).getRange(range).setFontColor(color);
}

function setFontStyle(range, style, link) {
    SpreadsheetApp.openByUrl(link).getRange(range).setFontStyle(style);
}

function setFontFamily(range, font, link) {
    SpreadsheetApp.openByUrl(link).getRange(range).setFontFamily(font);
}

function setFontWeight(range, weight, link) {
    SpreadsheetApp.openByUrl(link).getRange(range).setFontWeight();
}

function setFormula(range, formula, link) {
    SpreadsheetApp.openByUrl(link).getRange(range).setFormula(formula)
}

function setTabColor(color, link) {
    SpreadsheetApp.openByUrl(link).getActiveSheet().setTabColor(color);
}

function setTabColor(color, link) {
    SpreadsheetApp.openByUrl(link).getActiveSheet();
}

function copySheet(newName, link) {
    SpreadsheetApp.openByUrl(link).copy(newName);
}

function deleteActiveSheet(link) {
    var ss = SpreadsheetApp.openByUrl(link),
        sheet = ss.getActiveSheet();
    SpreadsheetApp.openByUrl(link).deleteSheet(sheet);
}

function deleteSheet(sheetName, link) {
    var sheet = SpreadsheetApp.openByUrl(link).getSheetByName(sheetName);
    if (sheet != null) {
        SpreadsheetApp.openByUrl(link).deleteSheet(sheet);
    } else {
        throw Error("Sheet does not exist");
    }
}

function duplicateActiveSheet(link) {
    SpreadsheetApp.openByUrl(link).duplicateActiveSheet();
}

function renameActiveSheet(newName, link) {
    SpreadsheetApp.openByUrl(link).renameActiveSheet(newName);
}

function insertSheet(name, link) {
    SpreadsheetApp.openByUrl(link).insertSheet(name);
}
