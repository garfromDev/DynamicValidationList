const COL_COCHE = 28;

function Verrouillageligne() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var noligne=spreadsheet.getActiveSheet().getActiveCell().getRow();
  var coche = spreadsheet.getActiveSheet().getRange(noligne, COL_COCHE).getValue();
  if(! coche){ return}
  var rangeToProtect = "A" + noligne + ":AV" + noligne;
  spreadsheet.getRange(rangeToProtect).activate();
  var protection = spreadsheet.getRange(rangeToProtect).protect();
  spreadsheet.getActiveRangeList().setFontColor('#999999');
};

//fonction pour récupérer l'url d'une feuille
function URL_FEUILLE() {
  var spreadsheet = SpreadsheetApp.getActiveSheet();  
  URL_FEUILLE = spreadsheet.getFormUrl();
}

// fonction appelée par le bouton
function envoiMailDemandeAnalyse() {
  sendEmail("magali.bossiere@ceva.com, alexandre.brechet@ceva.com, romain.skowron@ceva.com, lea.legrand@ceva.com, karine.gauvin@ceva.com, amandine.laigle@ceva.com, nelly.lesceau@ceva.com, aline.fromont@ceva.com","Nouvelle demande d'analyse",36);
}


/**
Envoi un email depuis le compte de l'utilisateur courant
to : l'adresse (ou les adresses séparés par des virgules) de destination
subject : le sujet du mail
fromCol : le no de la colonne dans laquelle on trouve le contenu du message
*/
function sendEmail(to, subject, fromCol)
{
  //1 on récupère le nO de ligne de la cellule sélectionnée
  l = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  // 2 le contenu du mail est dans la cellule de la même ligne, en colonne fromCol
  contenu = SpreadsheetApp.getActiveSheet().getRange(l, fromCol).getValue();
  
// 3 Display a dialog box with a title, message, and "Yes" and "No" buttons. The
// user can also close the dialog by clicking the close button in its title bar.
var ui = SpreadsheetApp.getUi();
var response = ui.alert('Confirmer envoi email ?', contenu, ui.ButtonSet.YES_NO);

// Process the user's response.
if (response == ui.Button.YES)
  MailApp.sendEmail(to, subject, contenu);
else if (response == ui.Button.NO)
  return;// on arrête tout
else
  Logger.log('The user clicked the close button in the dialog\'s title bar.');
return; // on arrête tout
}  
