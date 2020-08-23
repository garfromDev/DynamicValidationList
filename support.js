//======== Support functions =============
// could be used in other application

/**
*  Make the A1 notation for a given range in a given sheet 
* @param {Range} range : the range (can be in a different sheet)
* @param {Sheet} sheet : the  sheet
* @return {string} : a location in the sheet ( "'sheetName!A5A'")
*/
function locationForRangeInSheet(range, sheetName){
  return "'"+ sheetName.getName() +"'!"+ range.getA1Notation();
}


// retoune le Range par son nom (plage nommées)
function myRange(customName ){
   s = SpreadsheetApp.getActiveSpreadsheet();
   var r = s.getRangeByName(customName);
  if(r == null){
    throw new Error("Plage nommée '"+customName+"' non trouvée dans le classeur");
  }
  return r;
}


// retourne la valeur du champ nommé
// la valeur de la cellule en haut à gauche si le champ est grand
function value(namedRange){
  return myRange(namedRange).getValue();
}


// affiche une boite d'alerte avec le message
function alert(prompt){
   SpreadsheetApp.getUi().alert(prompt);
}

/**
* @param {String} message : the message to display
* @return {String} : the text input by the user, empty if "close" clicked
*/
function prompt(message){
  var ui = SpreadsheetApp.getUi();
  return ui.prompt(message, ui.ButtonSet.OK).getResponseText();
}


// affiche le message dans le coin en bas à droite
function toast(msg){
  SpreadsheetApp.getActiveSpreadsheet().toast(msg);   
}


/** copie la feuille vers le classeur et efface toute les autres feuilles
* @param {Sheet} fromSheet
* @param {Spreadsheet} targetSpreadsheet
*/
function copyUniqueTo(fromSheet, targetSpreadsheet){
 var targetSheet = fromSheet.copyTo(targetSpreadsheet);
 targetSheet.setName(fromSheet.getName());
 var sheets = targetSpreadsheet.getSheets(); 
 sheets.forEach(function(s){
    if(s.getName() != targetSheet.getName()){
      targetSpreadsheet.deleteSheet(s);
    }
 }); 
}


// make the string for a formula that create an hyperlink to link
// displaying "display" in the cell
function getHyperlinkFormulaToWithDisplay(link, display){
 return  '=HYPERLINK("'+link+'"; "'+display+'")'; 
}


/** add an hyperlink to link to the cell
* it will display the current cell content
* @param {Cell} cell
* @param {String} link
*/
function addHyperlinkToCell(cell, link){
  cell.setFormula( 
    getHyperlinkFormulaToWithDisplay(
      link, cell.getValue())
    );
}


/** add an hyperlink to link to the cell
* it will display the current cell content which is a date, formatted as dd/mm/YY
* @param {Cell} cell
* @param {String} link
*/
function addHyperlinkToDateCell(cell, link){
  cell.setFormula( 
    getHyperlinkFormulaToWithDisplay(
      link, Utilities.formatDate(cell.getValue(), "GMT", "yyyy-MM-dd"))
    );
}

/** return the URL of a given sheet in this spreadsheet (for direct access, without opening a new tab)
* @param {Sheet} sheet
* @return {String}
*/
function getLinkToSheet(sheet){
  return "#gid="+sheet.getSheetId();
}


/** return the active sheet of the active spreadsheet
* @return {Sheet}
*/
function activeSheet(){
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}


/** return the sheet in this spreadsheet with given name (null if doesn't exist)
* @param {String} name
* @return {Sheet}
*/
function getSheet(name){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}


/** return the value of the range in active sheet  in this spreadsheet with given 
* location (in A1 notation)
* @param {String} A1notation
* @return {Value}
*/
function getValueForRange(A1notation){
  return activeSheet().getRange(A1notation).getValue();
}


/** return the value of the range in active sheet  in this spreadsheet 
* In case the range is calculated through custom function, force the refreshing
* trick from https://issuetracker.google.com/issues/36754498
* NOT USED TODAY, kest for future need
* @param {Range} the range
* @return {Value} value of first cell of the range
*/
function getFilteredValue(range){
  var v = range.getValue();
  for(var counter=15; v.toString() == "#N/A" && counter>0; counter--) { 
      v = range.getValue(); 
  }
  return v;
}


/** return the row from an expression like A14:G13 ou F5
* @param {String} rangeA1
* @return {Int}
*/
function getRowFromA1(rangeA1){
  return rangeA1.match(/\d+/)[0];
}


/** the user time zone from active spreadsheet
* @return {String}
*/
function getUserTimeZone() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone();
}


/**
* remove the protection of existing range of the sheet
* @param {Sheet} sheet
* @return {unprotectedRanges :[Range], sheet:{Sheet}} : the original unprotected ranges, empty array ih wasn't  protected
*/
function unprotectSheet(sheet){
  var protections=sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if(protections.length<1){
    return {unprotectedRanges : [], sheet : sheet};
  }
   var originalUnprotected = protections[0].getUnprotectedRanges();
   protections[0].setUnprotectedRanges([sheet.getDataRange()]);
  return {unprotectedRanges : originalUnprotected, sheet : sheet};
}


/**
* @param {unprotectedRanges:[Range],sheet: {Sheet}} originalUnprotected 
* @return {Protection} for chaining
*/
function restoreProtection(originalUnprotected){
  var protections=originalUnprotected.sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if(protections.length<1){
    return protections;
  }
  return protections[0].setUnprotectedRanges(originalUnprotected.unprotectedRanges);
}
 

/**
* Copy the sheet protection from one sheet to another one
* @param {Sheet} fromSheet
* @param {Sheet} toSheet
* @return {Protection} : the protection object of the new sheet, null if fromSheet was not protected
* NOTE : only the first sheet protection is copied, including unprotected ranges
* CAUTION : no check done, may throw if sheets do not exist or function executed with unsuficient privilege
*/
function copyProtectionFromSheetToSheet(fromSheet, toSheet){
  var protections = fromSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if(protections.length<1){return null;}
  return copyProtectiontoSheet(protections[0],
                        toSheet);
}


/**
* Copy the given protection to another sheet
* @param {Protection} protection
* @param {Sheet} targetSheet
* @return {Protection} : the protection object of the new sheet
* NOTE : the protection is copied, including unprotected ranges
* CAUTION : no check done, may throw if sheets do not exist or function executed with unsuficient privilege
*/
function copyProtectiontoSheet(protection, targetSheet){
  var ur = protection.getUnprotectedRanges();
  // convert range into same range in new sheet
  var targetUr=[];
  for(i=0;i<ur.length;i++){
    targetUr.push(targetSheet.getRange(ur[i].getA1Notation()));
  }
  // set description to sheet name and copy unprotected ranges
  var newProtection= targetSheet.protect()
    .setDescription(targetSheet.getSheetName())
    .setUnprotectedRanges(targetUr);
  // allowed editors are those from original protection    
   return  newProtection.removeEditors(newProtection.getEditors())
    .addEditors(protection.getEditors());
}



/**
* force a sheet to refresh (when using query, Index(), custom function
* @param {Spreadsheet} the spreadsheet to whic the sheet belongs
* @param {Sheet} the sheet to refresh
* @return 
* NOTE : max 10 attempt done if legitimate #N/A in cell
* CAUTION : no check done, may throw if sheets do not exist or function executed with unsuficient privilege
*/
function refreshSheet(spreadsheet, sheet) {
  var dataArrayRange = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  var dataArray = dataArrayRange.getValues(); // necessary to refresh custom functions
  var nanFound = true;
  var cpt = 10;
  while(nanFound && cpt > 0) {
    for(var i = 0; i < dataArray.length; i++) {
      if(dataArray[i].indexOf('#N/A') >= 0) {
        nanFound = true;
        dataArray = dataArrayRange.getValues();
        cpt--; // to avoid looping when formula result in #N/A legitimely
        break;
      } // end if
      else if(i == dataArray.length - 1) nanFound = false;
    } // end for
  } // end while
}



/* ======== discussion =============
I tried to declare a class Support fo avoid global function
and allow auto-completion but app script doesn't allow class declaration
if the class is instancied through function constructor ( function Support(){ this.getTruc = function () { ...}
we need to instanciate a global support object, but auto-completion doesn't work
*/


//============ PROV ==================
function testgetRowFromA1(){
  alert( getRowFromA1("A1"));
}
function testgetValueForRange(){
  alert(getValueForRange("A2"));
}

function testgetSheet(){
 alert(getSheet("toto") == null); 
}

function testuserTimeZOne(){
  Logger.log(getUserTimeZone()); 
}
