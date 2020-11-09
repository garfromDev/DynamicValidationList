// Column in the souchier where the souche reference is
const COL_SOUCHE = 3
// sheet where the auto_import manager store its information
const AUTO_IMPORT_MNG = "Auto_Import"
// sheet where the request mus be added
const AUTO_IMPORT_TARGET = "Demandes"

// parameters for the line to insert into the request list
function toInsert() {
    return [
        new Data(column=2,  value = new Date(), dataVal=false),
        new Data(column=3,  value = "Ceva Biovac", dataVal=true),                    // request site
        new Data(column=4,  value = "magali.bossiere@ceva.com", dataVal=false),      // requester
        new Data(column=5,  value = ImportManager.getNextSouche(), dataVal=false),   // souche reference
        new Data(column=12, value = "Identification", dataVal=false),
        new Data(column=13, value = "Identification_malditof", dataVal=true),        // analysis requested
        new Data(column=14, value = "Ceva Biovac", dataVal=false),
        new Data(column=17, value = "Import auto depuis le souchier", dataVal=false)
    ];
}


class Data {
    constructor(column, value, dataVal) {
        this.column = column;
        this.value = value;
        this.dataVal = dataVal; // true if this column is trigering dataValidation list in other columns
    }
}


var ImportManager = {

    init : function() {
        this._aim = getSheet(AUTO_IMPORT_MNG);
        this._target_sheet = getSheet(AUTO_IMPORT_TARGET);
        this._lastMngLine = this._aim.getLastRow();
        // we look for column E, because others columns have empty row agt the top
        this._line_to_insert = getFirstEmptyRow(this._target_sheet.getRange("E:E"));
        let id = this._aim.getRange(this._lastMngLine, 1).getValue();
        let name = this._aim.getRange(this._lastMngLine, 2).getValue();
        this._import_from_sheet = SpreadsheetApp.openById(id).getSheetByName(name);
        this._line_from = this._aim.getRange(this._lastMngLine, 5).getValue() + 1;
        this._max_line_from = getFirstEmptyRow(this._import_from_sheet.getRange("C:C")) - 1;
    },

    performImport : function(){
       this.startImport();
       while(!this.import_finished()){
           this.addRequest();
       } 
       this.endImport();
    },

    addRequest : function(){
        cols = toInsert();  // generate the row content
        cols.forEach(data => {
            let rng = this._target_sheet.getRange(this._line_to_insert, data.column)
            rng.setValue(data.value);
            if(data.dataVal){ // force update of validation list because onEdit may not trigger properly
                updateDynamicValidationListIfNeeded({value: data.value, range:rng});
            }
        });

        this._line_to_insert++;
    },

    getNextSouche : function(){
        return this._import_from_sheet.getRange(this._line_from++, COL_SOUCHE).getValue();
    },

    import_finished : function(){
        return this._line_from > this._max_line_from;
    },

    startImport : function(){
        let cell = this._aim.getRange;
        let nextLine = this._lastMngLine + 1;
        cell(nextLine, 1).setValue(cell(this._lastMngLine, 1).getValue());  //File
        cell(nextLine, 2).setValue(cell(this._lastMngLine, 2).getValue());  //sheet
        cell(nextLine, 3).setValue(new Date());
        cell(nextLine, 4).setValue(this._line_from);
        cell(nextLine, 5).setValue(this._max_line_from);
        cell(nextLine, 6).setValue("STARTED...");
    },

    endImport(){
        let cell = this._aim.getRange;
        let nextLine = this._lastMngLine + 1;
        cell(nextLine, 6).setValue("DONE");
        this._lastMngLine = nextLine;
    }

}


function importFromSouchier(){
    ImportManager.init();
    ImportManager.performImport();
}


// ================================================ TESTS ============================================
// those test are not unit test, they are intended to interactively test the setup by adjusting values
function testInit() {
    let im = ImportManager;
    im.init();
    if(im._import_from_sheet.getName() != "Souchier Ceva Biovac") { alert("wrong name  " + im._import_from_sheet.getName())}
    if(im._line_from != 17824) {alert("wrong line_from : " + im._line_from)}
    alert("max line from.  " + im._max_line_from);
    if(im._line_to_insert != 10) {alert("wrong _lineToInsert " + im._line_to_insert)}
    if(im._lastMngLine != 2) {alert("wrong _lastMngLine " + im._lastMngLine)}
}

function testLastRow(){
  alert(activeSheet().getLastRow());
}

function testStart() {
    let im = ImportManager;
    im.init();
    im.startImport(); 
}

function testEnd(){
    let im = ImportManager;
    im.init();
    im.endImport();
}

function testAddRequest() {
    let im = ImportManager;
    im.init();
    im.addRequest();
    if(im._line_to_insert != 11) {alert("wrong _lineToInsert " + im._line_to_insert)}

}