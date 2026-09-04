const SHEET_ID = 'VOTRE_ID_DE_FICHIER_SPREADSHEET'; // À remplacer (trouvable dans l'URL)

function doGet(e) {
  // Routage basique pour afficher différentes pages
  let page = e.parameter.page || 'emprunt';
  let html = HtmlService.createTemplateFromFile(page === 'membre' ? 'AjoutMembre' : 'Emprunt');
  return html.evaluate().setTitle('Gestion Plongée').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ajouterMembre(uid, nom, prenom) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Membres');
  sheet.appendRow([uid, nom, prenom]);
  return "Membre ajouté avec succès.";
}

function enregistrerEmprunt(itemIdentifiant, memberUid) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // 1. Trouver le membre
  const membresSheet = ss.getSheetByName('Membres');
  const membresData = membresSheet.getDataRange().getValues();
  let membreNom = "Inconnu";
  for (let i = 1; i < membresData.length; i++) {
    if (membresData[i][0] === memberUid) {
      membreNom = membresData[i][1] + ' ' + membresData[i][2]; // Nom Prénom
      break;
    }
  }

  // 2. Trouver l'équipement (recherche dans les 3 feuilles)
  const categories = ['Détendeurs', 'Stabs', 'Blocs'];
  let itemTrouve = false;
  let categorieTrouvee = "";
  
  for (const cat of categories) {
    let sheet = ss.getSheetByName(cat);
    let data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      // On cherche soit par l'ID (col A), soit par l'UID NFC (col B)
      if (data[i][0].toString() === itemIdentifiant || data[i][1] === itemIdentifiant) {
        // Mise à jour de la disponibilité et du dernier emprunteur
        // Hypothèse : Dispo est l'avant-dernière colonne, Dernier emprunteur la dernière.
        // Adaptez les index selon votre tableau exact.
        let dispoCol = data[0].indexOf('Dispo ?') + 1;
        let emprunteurCol = data[0].indexOf('Dernier emprunteur') + 1;
        
        sheet.getRange(i + 1, dispoCol).setValue('non');
        sheet.getRange(i + 1, emprunteurCol).setValue(membreNom);
        
        itemTrouve = true;
        categorieTrouvee = cat;
        break;
      }
    }
    if (itemTrouve) break;
  }

  if (!itemTrouve) throw new Error("Équipement introuvable.");

  // 3. Ajouter à la liste des emprunts
  const empruntsSheet = ss.getSheetByName('Liste des emprunts');
  const dateJour = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  empruntsSheet.appendRow([dateJour, "", categorieTrouvee, itemIdentifiant, membreNom]);

  return `Matériel (${categorieTrouvee}) attribué à ${membreNom}.`;
}