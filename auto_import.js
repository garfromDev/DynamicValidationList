// Column in the souchier where the souche reference is
const COL_SOUCHE = 3
// sheet where the auto_import manager store its information
const AUTO_IMPORT_MNG = "Auto_Import"
// sheet where the request mus be added
const AUTO_IMPORT_TARGET = "Demandes"

class Data {
    constructor(column, value) {
        this.column = column;
        this.value = value;
    }
}

function toInsert() {
    return [
        new Data(column=2,  value = new Date()), 
        new Data(column=3,  value = "Ceva Biovac"),                  // request site
        new Data(column=4,  value = "magali.bossiere@ceva.com"),     // requester
        new Data(column=5,  value = ImportManager.getNextSouche()),  // souche reference
        new Data(column=17, value = "Import auto depuis le souchier")
    ];
}


var ImportManager = {
    _target_sheet : getSheet(AUTO_IMPORT_TARGET),

    init : function(){
        this._aim = getSheet(AUTO_IMPORT_MNG);
        this._lastMngLine = this._aim.getLastRow();
        this._line_to_insert = getSheet(AUTO_IMPORT_TARGET).getLastRow() + 1;
        let id = this._aim.getRange(this._lastMngLine, 1).getValue();
        this.gid = this._aim.getRange(this._lastMngLine, 1).getValue();
        this._import_from_sheet = SpreadsheetApp.openById(id).getSheetByName("Souchier Ceva Biovac");
        this._line_from = this._aim.getRange(this._lastMngLine, 5).getValue() + 1;
        this._max_line_from = this._import_from_sheet.getLastRow();
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
            this._target_sheet.getRange(line, data.column).setValue(data.value);
        });
        this._line_to_insert++;
    },

    getNextSouche : function(){
        return this._from_sheet.getRange(this._line_from++, COL_SOUCHE).getValue();
    },

    import_finished : function(){
        return this._line_from > this._max_line_from;
    },

    startImport : function(){
        let cell = this.aim.GetRange;
        let nextLine = this._lastMngLine + 1;
        cell(nextLine, 1).setValue(cell(this._lastMngLine, 1).getValue());  //File
        cell(nextLine, 2).setValue(cell(this._lastMngLine, 2).getValue());  //sheet
        cell(nextLine, 3).setValue(new Date());
        cell(nextLine, 4).setValue(this._line_from);
        cell(nextLine, 5).setValue(this._max_line_from);
        cell(nextLine, 6).setValue("STARTED...");
    },

    endImport(){
        let cell = this.aim.GetRange;
        let nextLine = this._lastMngLine++;
        cell(nextLine, 6).setValue("DONE");
    }

}