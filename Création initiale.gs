//                      Projet Suivi de gestion 2024
//                      Création initiale des tableaux
//                              02/05/2024 
//                                 V1.1
//                                
//

function createCopiesAndInsertIds() {
  
  // Tableur récapitulatif
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Magasins');

  // Colonne A - noms des magasins
  var dataRange = sheet.getRange('A:A');
  var values = dataRange.getValues();

  // Dossier de dépôt
  var folder = DriveApp.getFolderById('1AyZWN7xB6lmArFtPv1SQkkS-dxOMcFaE');

  // Boucle indexée sur dataRange
  for (var i = 7; i < values.length; i++) {
    var name = values[i][0];

    // Check si nom renseigné
    if (name !== "") {

      // Copie du modèle
      var template = DriveApp.getFileById('1IjTEw9F82AMVdLFe5xmYFUZaaoVJQ_YeF0iEBJpCh_E');
      var copy = template.makeCopy(name, folder);
      copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

      // Ajout de l'ID
      var copyId = copy.getId();
      sheet.getRange(i + 1, 4).setValue(copyId);

      Logger.log("Tableau créé pour : " + name);
    }
  }
}
