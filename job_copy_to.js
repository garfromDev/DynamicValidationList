var ExportManager = {
    init : function(source_sheet_name, target_spreadsheet_id, target_sheet_name, fields_to_map, 
        target_not_empty_column = "B", 
        source_not_empty_field = "Demandeur",
        fields_with_date = [], fields_with_text = [],
        must_be_true = [], must_be_false = [],
        unique_id, 
        force_last_line,
        max_header_line = 3) {
        /**
         * @param {Sheet} source_sheet
         * @param {Sheet} target_sheet
         * @param {String: String} fields_to_map dict of source_field_name: target_fields_name where name is column header
         */
        if(!(source_sheet_name && target_spreadsheet_id && fields_to_map)){
            throw 'Need source_sheet, target_sheet and fields_to_map';
        };
        this.source_sheet = getSheet(source_sheet_name);
        this.target_sheet = SpreadsheetApp.openById(target_spreadsheet_id).getSheetByName(target_sheet_name);
        console.log("target sheet. :  ",this.target_sheet);
        this.target_not_empty_column = target_not_empty_column;
        this.must_be_true = must_be_true;
        this.must_be_false = must_be_false;
        this.source_not_empty_field = source_not_empty_field;
        this.fields_to_map = fields_to_map;
        this.fields_with_date = fields_with_date;
        this.fields_with_text = fields_with_text;
        this.max_header_line = max_header_line;
        this.unique_id = unique_id;
        this.set_last_line(force_last_line);
        this.get_column_for_fields();
    },
    

    find_header : function(field, sheet){
        console.log("find header <",field,"> in ",sheet.getName());
        if(!(field && sheet)){return null;}
        for(var line=1;line <= this.max_header_line; line++){
            for(var col=1; col <= sheet.getLastColumn(); col++){
                if(sheet.getRange(line,col).getValue() == field){
                    console.log("find-header found  ",field," at line ",line," col ", col, " of ",sheet.getName());
                    return col;
                }
            }
        }
        return null;
    },


    find_columns_for : function(fields_to_find, sheet = this.source_sheet){
        var col_list = [];
        for( const field of fields_to_find){
            var col = this.find_header(field, sheet);
            if(col){
                col_list.push(col);
            }
        }
        return col_list;
    },


    get_column_for_fields : function(){
        this.col_must_be_true = this.find_columns_for(this.must_be_true);
        console.log("col must be true : ", this.col_must_be_true);
        this.col_must_be_false = this.find_columns_for(this.must_be_false);
        console.log("col must be false : ", this.col_must_be_false);
        this.source_not_empty_col = this.find_columns_for([this.source_not_empty_field])
        console.log("fields with date. :  ", this.fields_with_date);
        this.col_with_date = this.find_columns_for(this.fields_with_date, this.target_sheet)
        console.log("col with date : ", this.col_with_date);
        //var col_for_field_with_text = this.find_columns_for(this.fields_with_text, this.target_sheet);
        this.col_for_field_with_text = {};
        for(const field in this.fields_with_text){
            this.col_for_field_with_text[field] = {'col': this.find_columns_for([field], this.target_sheet)
                                                              .find(x=>x!==undefined),
                                                   'value': this.fields_with_text[field]
                                                  }
        }
        console.log("col_for_field_with_text : ", this.col_for_field_with_text);
        this.col_mapping = {}
        for(const field in this.fields_to_map){
            this.col_mapping[field] = {
                'source_col' : this.find_columns_for([field]).find(x=>x!==undefined),
                'target_col' : this.find_columns_for([this.fields_to_map[field]], this.target_sheet).find(x=>x!==undefined)
            }
        }
    },


    get_first_free_line_of_target : function(){
        /**
         * returns first free line in column target_not_empty_column of target sheet
         */
        const letter = this.target_not_empty_column;
        return getLastRowForColumn(this.target_sheet.getRange(letter + ":" + letter)) + 1;
    },

    set_last_line(force_last_line){
        var last_line_key = 'ExportManager.' + this.unique_id;
        var last_line = PropertiesService.getScriptProperties().getProperty(last_line_key);
        if(force_last_line > last_line){
            PropertiesService.getScriptProperties().setProperty(last_line_key, force_last_line);
        }
    },

    copy_line_to_target : function(source_line, target_line){
        /**
         * Copy mapped filed from source line to traget line of target sheet and set date field and fixed text field
         * @param source_line {int} the line in source sheet to copy to target line in target sheet 
         */
        for(const field in this.fields_to_map){
            var maping = this.col_mapping[field];
            this.target_sheet.getRange(target_line, maping.target_col)
            .setValue(this.source_sheet.getRange(source_line, maping.source_col).getValue());
        }
        for(const field in this.fields_with_text){
            var maping = this.col_for_field_with_text[field];
            this.target_sheet.getRange(target_line, maping.col).setValue(maping.value);
        }
        for(const col of this.col_with_date){
            this.target_sheet.getRange(target_line, col).setValue(new Date);
        }
    },


    end_of_data_reached : function(line){
        return this.source_sheet.getRange(line, this.source_not_empty_col).getValue() == "";
    },


    must_be_exported : function(line){
        var result = true;
        result = result && this.col_must_be_true.every(col => this.source_sheet.getRange(line, col).getValue());
        result = result && ! this.col_must_be_false.some(col => this.source_sheet.getRange(line, col).getValue());
        return result;
    },


    run_export : function() {
        var last_line_key = 'ExportManager.' + this.unique_id;
        var last_line_exported = parseInt(PropertiesService.getScriptProperties().getProperty(last_line_key)) || 1;
        last_line_exported++;
        var target_line = this.get_first_free_line_of_target();
        while(!this.end_of_data_reached(last_line_exported)){
            if(this.must_be_exported(last_line_exported)){
                this.copy_line_to_target(last_line_exported, target_line++);
            }
            last_line_exported++;
        }
        PropertiesService.getScriptProperties().setProperty(last_line_key, last_line_exported - 1);
    }
}

/** fonction a appeller depuis le declencheur horaire */
function export_vers_demande_repiquage(){
    // TODO: gestion erreurs
    // NE PAS CHANGER L'ORDRE DES PARAMETRES, NE PAS EN SUPPRIMER !!!
    ExportManager.init(
        'Demandes',                                     // source sheet name
        '113g_b6dqVSrTRSjRNYGKSyXMWk8oXmM4GNDTgRqhmqo', // target spreadsheet ID
        'Demandes repiquages',                          // target sheet name
        // maping des champs "source": "target" séparé pas  des virgules
        {
            "Demandeur": "Demandeur / origine demande",
            "Référence souche demandeur\n(N°Cl si souchotèque Ceva Biovac)" : "n°CL FMP12",
            "N° demande" : "N° demande\n(si applicable)",
            "Labo\n(Auto)" : "Labo destinataire de l'envoi\n(si applicable)",
            "Milieu de repiquage\n\n\n(si demande particulière)" : "Milieu\n(si demande particulière)",
            "Condition d'incubation\n\n\n(si demande particulière)" : "Incubation\n(si demande particulière)",
            "Nb de boîtes à repiquer\n\n\n(si demande particulière)" : "Nb de boîtes (si demande particulière)",
            "Commentaire labo bactério" : "Commentaires (interne labo)",
            "Commentaire à l'attention de Ceva Biovac" : "Commentaires demandeur\n(si applicable)"
        },
        //======!! make sure this field is mandatory !!!!!========
        "A",                                            // column letter to detect end of data in target file
        //======!! make sure this field is mandatory !!!!!========
        "Demandeur",                                    // fields to detect end of data in source file
        ['Date'],                                       // target field filled with current date ['nom1', 'nom2'] ou [] si aucun
        // target field filled with raw text "target_field_name": "text". {} si aucun champ de type texte
        {
            "Destination repiquage": "Labo bactério"
        },
        ['Demande de repiquage'],['Validation identification\n(Auto)']      //source field(s) that must all be true
        ['Annuler demande'],                                                //source field(s) that must all be false   
        'export_vers_demande_repiquage',                                    //ID unique du script pour stocker les lignes atteintes
        49                                               // no de la ligne a laquelle commencer, est ignoré si zéro ou si cette ligne est dépassée
    );
    ExportManager.run_export();
}

function export_vers_planning_malditof(){
    // TODO: gestion erreurs
    // NE PAS CHANGER L'ORDRE DES PARAMETRES, NE PAS EN SUPPRIMER !!!
    ExportManager.init(
        'Demandes',                                                // source sheet name
        '1mzA8VvlIOsF8eOoEwsa8keyCZI-74FRIeveB6JsPCl8', // target spreadsheet ID
        'AV',                                                      // target sheet name
        // maping des champs "source": "target" séparé pas  des virgules
        {
            "N° demande": "Origine demande",
            "Date demande\n(envoi email Ctrl+Alt+Shift+1)": "Date de la demande",
            "Référence souche demandeur\n(N°Cl si souchotèque Ceva Biovac)" : "Référence",
            "GEB client\n(Auto)" : "GEB client\n(si applicable)",
            "GEB Biovac\n(Auto)" : "GEB Biovac\n(si applicable)"
        },
        //======!! make sure this field is mandatory !!!!!========
        "B",                                         // column letter to detect end of data in target file
        //======!! make sure this field is mandatory !!!!!========
        "Demandeur",                                          // fields to detect end of data in source file
        ['Date transfert auto demande'],                      // target field filled with current date ['nom1', 'nom2'] ou [] si aucun
        // target field filled with raw text "target_field_name": "text". {} si aucun champ de type texte
        {},
        ['Malditof'],                                                     //source field(s) that must all be true
        ['Annuler demande'],                                              //source field(s) that must all be false   
        'export_vers_planning_malditof',                 //ID unique du script pour stocker les lignes atteintes
         0                                              // no de la ligne a laquelle commencer, est ignoré si zéro ou si cette ligne est dépassée
    );
    ExportManager.run_export();
}

function export_vers_suivi_analyses_externes(){
    // TODO: gestion erreurs
    // NE PAS CHANGER L'ORDRE DES PARAMETRES, NE PAS EN SUPPRIMER !!!
    ExportManager.init(
        'Demandes',                                                              // source sheet name
        '1BWZ33Uw4vOfhAviM1xcDw74Ph9w5j3TxDjib9P5jvPA',                          // target spreadsheet ID
        'Suivi des analyses externes',                                           // target sheet name
        // maping des champs "source": "target" séparé pas  des virgules
        {
            "N° demande": "N° demande\n(si applicable)",
            "Demandeur": "Demandeur",
            "Référence souche demandeur\n(N°Cl si souchotèque Ceva Biovac)" : "Ref souche Biovac",
            "Ref souche client\n(Auto)" : "Ref souche client\n(si applicable)",
            "GEB Biovac\n(Auto)" : "GEB Biovac\n(si applicable)",
            "Sérotype Biovac\n(Auto)" : "Sérotype Biovac\n(si applicable)",
            "GEB client\n(Auto)" : "GEB client\n(si applicable)",
        },
        //======!! make sure this field is mandatory !!!!!========
        "B",                                         // column letter to detect end of data in target file
        //======!! make sure this field is mandatory !!!!!========
        "Demandeur",                                          // fields to detect end of data in source file
        ['Date transfert auto demande'],                      // target field filled with current date ['nom1', 'nom2'] ou [] si aucun
        // target field filled with raw text "target_field_name": "text". {} si aucun champ de type texte
        {},
        ['Malditof'],                                                     //source field(s) that must all be true
        ['Annuler demande'],                                              //source field(s) that must all be false   
        'export_vers_suivi_analyses_externes',                 //ID unique du script pour stocker les lignes atteintes
         0                                              // no de la ligne a laquelle commencer, est ignoré si zéro ou si cette ligne est dépassée
    );
    ExportManager.run_export();
}
// ================================================ TESTS ============================================
// those test are not unit test, they are intended to interactively test the setup by adjusting values
function reset_last_line(){
  PropertiesService.getScriptProperties().setProperty("ExportManager.export_vers_planning_malditof", 30);
}

function check_last_line(){
  console.log("export_vers_planning_malditof : ",
    parseInt(PropertiesService.getScriptProperties().getProperty("ExportManager.export_vers_planning_malditof"))
  );
  console.log("export_vers_suivi_analyses_externes : ",
    parseInt(PropertiesService.getScriptProperties().getProperty("ExportManager.export_vers_suivi_analyses_externes"))
  );
  console.log("export_vers_demande_repiquage : ",
    parseInt(PropertiesService.getScriptProperties().getProperty("ExportManager.export_vers_demande_repiquage"))
  );
}
function test_init(){
      ExportManager.init(
    'Demandes',                                     // source sheet name
    '113g_b6dqVSrTRSjRNYGKSyXMWk8oXmM4GNDTgRqhmqo', // target spreadsheet ID
    'Demandes repiquages',                          // target sheet name
    // maping des champs "source": "target"
    {
        "Référence souche demandeur (N°Cl si souchotèque Ceva Biovac)": "n°CL FMP12",
        "Date d'envoi ou transfert de la souche au labo bactériologie\n\n(N/A si souchotèque)" : "Commentaires"
    },
    //======!! make sure this field is mandatory !!!!!========
    "B",                                            // column to detect end of data in target file
    //======!! make sure this field is mandatory !!!!!========
    "Demandeur",                                    // fields to detect end of data in source file
    ['Date'],                                       // target field filled with current date
    // target field filled with raw text "target_field_name": "text"
    {
        "Demandeur / origine demande": "Demande d'analyse (auto)",
        "Destination repiquage": "Labo bactério"
    },
    ['Demande de 1er repiquage'],                   //field that must be true
    ['Annuler demande'],                             //field that must be false
    'testValue',
    7
    );
    console.log("source sheet name. ", ExportManager.source_sheet.getName());
    console.log("target sheet name. ", ExportManager.target_sheet.getName());
    console.log("target not empty column. ", ExportManager.target_not_empty_column);
    console.log("col_must_be_true  ", ExportManager.col_must_be_true);
    console.log("col_must_be_false  ", ExportManager.col_must_be_false);
    console.log("col_with date  ", ExportManager.col_with_date);
    console.log("col_for_field_with_text   ", ExportManager.col_for_field_with_text);
    console.log("col_mapping   ", ExportManager.col_mapping);
    console.log("find header 'Demande de 1er repiquage'", ExportManager.find_header('Demande de 1er repiquage', ExportManager.source_sheet));
    console.log("1st line of target ", ExportManager.get_first_free_line_of_target());
    console.log("end of data reach 48  ", ExportManager.end_of_data_reached(48));
    console.log("must be exported 48  ", ExportManager.must_be_exported(48));
    console.log("last_line", PropertiesService.getScriptProperties().getProperty('ExportManager.testValue'));
    //ExportManager.copy_line_to_target(48,ExportManager.get_first_free_line_of_target())
}