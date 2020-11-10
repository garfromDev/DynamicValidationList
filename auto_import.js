// Column in the souchier where the souche reference is
const COL_SOUCHE = 3
// sheet where the auto_import manager store its information
const AUTO_IMPORT_MNG = "Auto_Import"
// sheet where the request mus be added
const AUTO_IMPORT_TARGET = "Demandes"

// parameters for the line to insert into the request list
function toInsert() {
    return [
        new Data(column=2,  value = new Date()),
        new Data(column=3,  value = "Ceva Biovac"),                    // request site
        new Data(column=4,  value = "magali.bossiere@ceva.com"),      // requester
        new Data(column=5,  value = ImportManager.getNextSouche()),   // souche reference
        new Data(column=12, value = "Identification"),
        new Data(column=13, value = "Identification_malditof"),        // analysis requested
        new Data(column=14, value = "Ceva Biovac"),
        new Data(column=17, value = "Import auto depuis le souchier")
    ];
}


class Data {
    constructor(column, value, dataVal) {
        this.column = column;
        this.value = value;
    }
}


var ImportManager = {

    /**
     * get the value from the sheet AUTO_IMPORT_MNG to prep the ImportManager singleton
     */
    init : function() {
        // no error handling as there is no way to report error to user in a time triggered script
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

    /**
     * perform the creation of request for each souche in souchier, added a the end of
     * AUTO_IMPORT_TARGET sheet  
     * start line to import is taken from previous import in AUTO_IMPORT_MNG sheet  
     * **Note** `init()` must have been called before
     */
    performImport : function(){
       this.startImport();
       while(!this.import_finished()){
           this.addRequest();
       } 
       this.endImport();
    },

    /**
     * add a request line at _line_to_insert, based on information provided by toInsert()  
     * this is an array of Data, pair of (column, value)
     * set _line_to_insert to next line
     */
    addRequest : function(){
        cols = toInsert();  // generate the row content
        cols.forEach(data => {
            let rng = this._target_sheet.getRange(this._line_to_insert, data.column)
            rng.setValue(data.value);
        });
        this._line_to_insert++;
    },

    /**
     * return the souche reference from the souchier and increment pointer to the next
     */
    getNextSouche : function(){
        return this._import_from_sheet.getRange(this._line_from++, COL_SOUCHE).getValue();
    },

    /**
     * true if _max_line_from has been processed, no more souches to import
     */
    import_finished : function(){
        return this._line_from > this._max_line_from;
    },

    /**
     * Create a new porcess line with all information in the AUTO_IMPORT_MNG sheet  
     * **Note** `init()` must have been called before
     */
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

    /**
     * mark the current process line in AUTO_IMPORT_MNG as DONE and point to the next one
     */
    endImport(){
        let cell = this._aim.getRange;
        let nextLine = this._lastMngLine + 1;
        cell(nextLine, 6).setValue("DONE");
        this._lastMngLine = nextLine;
    }

}

/**
 * Main entry point  
 * is called by the dayly time trigger  
 * Will create request for all new souche in souchier since last execution
 */
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