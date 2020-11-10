const HEADER_MARKER = "@Ligne des en-têtes"
const DYNAMIC_VALIDATION_MARKER = "@ListeDynamique"; // must be written into a note in first line of column

/*
USAGE
=====
the dynamic validation list is a validation list 'from range'
each validation list refers to a different range, which is calculated from the value from which the list depends
This script dynamically creates data validation rules when a 'triggering' cell is edited
Each column that contains dynamic dropdown list must be marked with a note containg the marker in the topmost cell,
and pointing to the triggering cell for eample @ListeDynamique(Site demandeur,LexiqueDemandeur)
*/

var headerLine = 0;
var triggeringColumns = {}; // dictionary with 5: [12, 14] meaning that clomumn 5 triggers validation list change of col 12, 14
var rangeSource = {};  // dictionary no_col: "nom_de_l'onglet"


/**
 * Initialize the document properties according to the marker in the cell's note
 * NOTE : any change in markers fater opening the doc is not taken into account until next opening
 * @param {*} e eventObject
 * modify : docProperty headerLine, triggeringColumns, rangeSource
 */
function onOpen(e) {
  let sh = e.source.getActiveSheet();
  let docProperties = PropertiesService.getDocumentProperties();
  // find header line, which is indicated by a note in first column
  for(l = 1; l <= sh.getLastRow(); l++){
    if(sh.getRange(l,1).getNote().includes(HEADER_MARKER)){
      headerLine = l;
      docProperties.setProperty('headerLine', headerLine);
      break
    }
  }

  //find columns triggering validation list update
  var headers = sh.getRange(headerLine, 1, 1, sh.getLastColumn()).getValues();
  for(c=1; c <= headers[0].length; c++){ // horizontal range, headers[0][n]
    dependency = decodeDependencyNote(sh.getRange(1,c).getNote());
    if(dependency){
      let triggerCol = headers[0].indexOf(dependency.triggerName) + 1;
      if(triggeringColumns[triggerCol]){
        triggeringColumns[triggerCol].push(c);
      }else{
        triggeringColumns[triggerCol] = [c];
      }
      rangeSource[c] = dependency.rangeSource;
    }
  }
  docProperties.setProperty('triggeringColumns', JSON.stringify(triggeringColumns));
  docProperties.setProperty('rangeSource', JSON.stringify(rangeSource));
}


/**
 * decode the text with markers DYNAMIC_VALIDATION_MARKER(triggerName,rangeSource)
 * @param {*} noteText 
 * @returns {triggerName, rangeSource}, null if no valid marker in the note
 *  triggerName : the header of the column triggering this validationList
 *  rangeSource : the name of the sheet where to find data list according triggering value
 * NOTE : no check done for validity of range
 */
function decodeDependencyNote(noteText){
  regex = DYNAMIC_VALIDATION_MARKER + "\\(([^,]*),(.*)\\)";
  reg = new RegExp(regex, 'gm');
  result = reg.exec(noteText);
  return result ? {triggerName: result[1], rangeSource: result[2]} : null;
}


/*
* each time a cell is edited, check if this must trigger a change in validation List
*/
function onEdit(e){
  updateDynamicValidationListIfNeeded(e);
}


/**
 * check if something has to be done and triggers update of dependant columns
 * @param {*} e 
 * NOTE : multi-cell selection no taken into account
 */
function updateDynamicValidationListIfNeeded(e){
  headerLine = PropertiesService.getDocumentProperties().getProperty('headerLine')
  triggeringColumns = JSON.parse(PropertiesService.getDocumentProperties().getProperty('triggeringColumns'))
  rangeSource = JSON.parse(PropertiesService.getDocumentProperties().getProperty('rangeSource'))
  if(e.range.getRow() <= headerLine){return} // do not handle headre line or above
  if(e.value == undefined){
    return;
  } // TODO: multi-selection à traiter
  targetCols = triggeringColumns[e.range.getColumn()];
  if(targetCols){
    for(c=0; c<targetCols.length;c++){ col = targetCols[c];                              
      updateColValidationListFromSourceForValue(col, rangeSource[col], e);
    }
  }
}


/**
 * set new validation rule
 * @param {*} col : no of col to update 1=A (line is taken from e)
 * @param {*} sheetName : sheet name where data list is
 * @param {*} e 
 */
function updateColValidationListFromSourceForValue(col, sheetName, e) 
{
  const range = getRangeForValueFromSheet(e.value, getSheet(sheetName));
  if(!range){return};
  const line = e.range.getRow();
  const sheet = e.range.getSheet();
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  sheet.getRange(line, col).setDataValidation(rule);
}


/**
 * get the range corresponding to the key value in the sheet
 * @param {*} triggerValue : key (value to which the data list is linked)
 * @param {*} sh : Spreadsheet where the data list is
 */
function getRangeForValueFromSheet(triggerValue, sh)
{
  for(var c=2; c <= sh.getLastColumn(); c++){
    if(sh.getRange(1, c).getValue() == triggerValue){
      return sh.getRange(2,c,sh.getLastRow() - 1, 1);
    }
  }
}




// ================================================ TESTS ============================================
// those test are not unit test, they are intended to interactively test the setup by adjusting values
function test_onOpen(){
  onOpen({source: SpreadsheetApp.getActiveSpreadsheet()});
    headerLine = PropertiesService.getDocumentProperties().getProperty('headerLine')
  triggeringColumns = JSON.parse(PropertiesService.getDocumentProperties().getProperty('triggeringColumns'))
  rangeSource = JSON.parse(PropertiesService.getDocumentProperties().getProperty('rangeSource'))
  alert("headerline " + headerLine);
  alert(Object.keys(triggeringColumns) + "   " + Object.values(triggeringColumns));
  alert(Object.keys(rangeSource));
}


function testgetRangeForValueFromSheet() {
  res = getRangeForValueFromSheet('Ceva Libourne', getSheet('LexiqueDemandeur'));
  alert(res.getA1Notation());
  alert(res.getCell(1,1).getValue() + "  " + res.getCell(2,1).getValue());
  
}

function testupdateColValidationListFromSourceForValue(){
  updateColValidationListFromSourceForValue(
    4,
    'LexiqueDemandeur',
    {value: 'Ceva Libourne', range: activeSheet().getRange("C3")}
  )
}

