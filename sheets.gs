function doPost(e) {
  // Ouvre la feuille active
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  try {
    // Récupère les données envoyées par votre page web
    var data = JSON.parse(e.postData.contents);
    
    // Vérification du token
    if (data.token !== "7RdFw0qL1fMHbS6mjhr1") {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Accès refusé"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // Ajoute une nouvelle ligne (Date actuelle + UID)
    var date = new Date();
    // Formatage simple de la date pour le tableur
    var dateString = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    sheet.appendRow([dateString, data.uid]);
    
    // Répond que tout s'est bien passé
    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.message}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
