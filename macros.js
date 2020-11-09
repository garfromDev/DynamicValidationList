const COL_COCHE = 28;

function Verrouillageligne() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var noligne=spreadsheet.getActiveSheet().getActiveCell().getRow();
  var coche = spreadsheet.getActiveSheet().getRange(noligne, COL_COCHE).getValue();
  if(! coche){ return}
  var rangeToProtect = "A" + noligne + ":AB" + noligne;
  spreadsheet.getRange(rangeToProtect).activate();
  var protection = spreadsheet.getRange(rangeToProtect).protect();
  spreadsheet.getActiveRangeList().setFontColor('#999999');
};