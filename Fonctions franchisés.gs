//                      Projet Suivi de gestion 2024
//              Gestion des fonctions franchisés accessibles depuis le google sheet référentiel
//                                 26/06/2024 
//                                    V2.1
//                                
//                        
//




//Création du menu custom dans la barre d'outils de GSheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Fonctions franchisés')
    .addItem('Envoyer le mail au point de vente', 'sendIndividualEmail')
    .addItem('Créer un nouveau tableau de suivi', 'userInput')
    .addToUi();
}


//Envoi individuel de mail
function sendIndividualEmail() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeStoresSheet = spreadsheet.getSheetByName('Magasins');
  var emailAddress = activeStoresSheet.getRange('B4').getValue();
  var emailTitle = "Tableau de suivi 2024";
  var sheetUrl = activeStoresSheet.getRange('D4').getValue();

  GmailApp.sendEmail(emailAddress, emailTitle, 'Bonjour,\n\nVeuillez trouver ci-après le lien vers votre tableau de suivi 2024.\n\n' + sheetUrl + '\n\nBien cordialement');

  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, "Europe/Paris", "yyyy-MM-dd HH:mm:ss");
  var logMessage = "Dernier email envoyé à : " + formattedDate;
  activeStoresSheet.getRange('F4').setValue(logMessage);
}


// Gestion de l'interface du formulaire de création de PDV
function userInput() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('UserInterface')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Créer un tableau');
}


// Récupération des points de vente dans l'onglet BDD Stores pour l'interface du formulaire de création de PDV
function getStores() {
  var storeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FILTERED_STORES');
  var values = storeSheet.getRange('A:B').getValues();
  var stores = [];
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][1]) {
      stores.push({ name: values[i][0], id: values[i][1] });
    }
  }
  return stores;
}


// Création de tableau suite à l'envoi du formulaire
function processInputs(data) {
  Logger.log("Point de vente: " + data.store);
  Logger.log("Store ID: " + data.storeId);
  Logger.log("Email: " + data.email);

  //Variables
  var folder = DriveApp.getFolderById('1AyZWN7xB6lmArFtPv1SQkkS-dxOMcFaE');
  var template = DriveApp.getFileById('1IjTEw9F82AMVdLFe5xmYFUZaaoVJQ_YeF0iEBJpCh_E');
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeStoresSheet = spreadsheet.getSheetByName('Magasins');
  //var name = data.store;

  

  var copy = template.makeCopy(data.store, folder);
  copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  var copyId = copy.getId();
  Logger.log("Tableur créé: " + copyId);

  activeStoresSheet.appendRow([data.store, data.email, data.storeId, copyId]);

  var newSheetUrl = "https://docs.google.com/spreadsheets/d/" + copyId;

  var emailTitle = "Tableau de suivi 2024";
  GmailApp.sendEmail(data.email, emailTitle, 'Bonjour,\n\nVeuillez trouver ci-après le lien vers votre tableau de suivi 2024.\n\n' + newSheetUrl + '\n\nBien cordialement');
  

}
