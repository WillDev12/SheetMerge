// Ⓒ Willdevv12 2023 //
function createSpreadsheet(name) {return SpreadsheetApp.create(name).getUrl();}
function getActiveSheet(link) {var cache = CacheService.getScriptCache();var ss = SpreadsheetApp.openByUrl(link);return ss.getActiveSheet().getName();}
function getSpreadsheetName(link) {return SpreadsheetApp.openByUrl(link).getName();}
function getSpreadsheetID(link) {return SpreadsheetApp.openByUrl(link).getId();}
function getSheetId(link) {return SpreadsheetApp.openByUrl(link).getActiveSheet().getSheetId();}
function getSheetName(link) {return SpreadsheetApp.openByUrl(link).getActiveSheet().getSheetName();}
function getURL(id) {return SpreadsheetApp.openById(id).getUrl();}
function getViewers(link) {return SpreadsheetApp.openByUrl(link).getViewers();}
function getOwner(link) {return SpreadsheetApp.openByUrl(link).getOwner();}
function getA1(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getA1Notation();}
function getCellBorder(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getBorder();}
function getFontColor(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getFontColor();}
function getFontSize(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getFontSize();}
function getCellRow(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getRow();}
function getCellColumn(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getColumn();}
function getFontStyle(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getFontStyle();}
function getFontFamily(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getFontFamily();}
function getFontWeight(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getFontWeight();}
function getFormula(range, link) {if (SpreadsheetApp.openByUrl(link).getRange(range).getFormula() == "") {return "none";} else {return SpreadsheetApp.openByUrl(link).getRange(range).getFormula();}}
function getLastColumn(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getLastColumn();}
function getLastRow(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getLastRow();}
function getValue(range, link) {return SpreadsheetApp.openByUrl(link).getRange(range).getValue();}
function setCell(cell, text, link) {var ss = SpreadsheetApp.openByUrl(link);var sheet = ss.getActiveSheet();sheet.getRange(cell).setValue(text);}
function setCellFontSize(cell, size, link) {var ss = SpreadsheetApp.openByUrl(link);var sheet = ss.getActiveSheet();sheet.getRange(cell).setFontSize(size);}
function clearRange(range, link) {var ss = SpreadsheetApp.openByUrl(link);var sheet = ss.getActiveSheet();sheet.getRange(range).clear();}
function deleteRow(rowNum, link) {SpreadsheetApp.openByUrl(link).deleteRow(rowNum);}
function deleteRows(row, howMany, link) {SpreadsheetApp.openByUrl(link).deleteRows(row, howMany);}
function deleteColumn(columnNum, link) {SpreadsheetApp.openByUrl(link).deleteColumn(columnNum);}
function deleteColumns(column, howMany, link) {SpreadsheetApp.openByUrl(link).deleteColumns(column, howMany);}
function setRowHeight(row, height, link) {SpreadsheetApp.openByUrl(link).setRowHeight(row, height);}
function setColumnWidth(column, width, link) {SpreadsheetApp.openByUrl(link).setColumnWidth(column, width);}
function setCellBorder(range, top, left, bottom, right, link) {SpreadsheetApp.openByUrl(link).getRange(range).setBorder(top, left, bottom, right);}
function setFontColor(range, color, link) {SpreadsheetApp.openByUrl(link).getRange(range).setFontColor(color);}
function setFontStyle(range, style, link) {SpreadsheetApp.openByUrl(link).getRange(range).setFontStyle(style);}
function setFontFamily(range, font, link) {SpreadsheetApp.openByUrl(link).getRange(range).setFontFamily(font);}
function setFontWeight(range, weight, link) {SpreadsheetApp.openByUrl(link).getRange(range).setFontWeight();}
function setFormula(range, formula, link) {SpreadsheetApp.openByUrl(link).getRange(range).setFormula(formula)}
function setTabColor(color, link) {SpreadsheetApp.openByUrl(link).getActiveSheet().setTabColor(color);}
function setTabColor(color, link) {SpreadsheetApp.openByUrl(link).getActiveSheet();}
function copySheet(newName, link) {SpreadsheetApp.openByUrl(link).copy(newName);}
function deleteActiveSheet(link) {var ss = SpreadsheetApp.openByUrl(link),sheet = ss.getActiveSheet();SpreadsheetApp.openByUrl(link).deleteSheet(sheet);}
function deleteSheet(sheetName, link) {var sheet = SpreadsheetApp.openByUrl(link).getSheetByName(sheetName);if (sheet != null) {SpreadsheetApp.openByUrl(link).deleteSheet(sheet);} else {throw Error("Sheet does not exist");}}
function duplicateActiveSheet(link) {SpreadsheetApp.openByUrl(link).duplicateActiveSheet();}
function renameActiveSheet(newName, link) {SpreadsheetApp.openByUrl(link).renameActiveSheet(newName);}
function insertSheet(name, link) {SpreadsheetApp.openByUrl(link).insertSheet(name);}
