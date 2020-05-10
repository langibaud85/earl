// Cahier des charges
// On cherche le premier changement et on envoie un email pcp et un email secondaire
// il y a mail de verification de fonctiopnnement tout les soirs à 21h00
// script avec trigger toutes les heures

//*****************************************
// Fonction Metier
//*****************************************
function LectureTable(FileRef , tabCour) { 
       
       retour= -1 ;
  
      for (j=0 ; j < 9 ; j++) {
        sTmp01='A'+glbCpt01.toString() ;
        sTmp02='B'+glbCpt01.toString() ;
        sTmp03='C'+glbCpt01.toString() ;
        sTmp04='D'+glbCpt01.toString() ;
        sTmp05='E'+glbCpt01.toString() ;
        
        if (  FileRef.getRange(sTmp01).getValue() == 0  ) { // il n'y a pas de valeur prealable
          // On affecte les valeurs pour la semaine
          FileRef.getRange(sTmp01).setValue(tabCour.getCell(0, j+1).editAsText().getText() );
          // On affecte les valeurs pour Bete LS
          FileRef.getRange(sTmp02).setValue(tabCour.getCell(1, j+1).editAsText().getText() );
          // On affecte les valeurs pour Bete TRAD
          FileRef.getRange(sTmp03).setValue(tabCour.getCell(4, j+1).editAsText().getText() );
        }else
        {
          // on regarde si il y a une difference de valeur pour les 2 types de betes
          if ( FileRef.getRange(sTmp02).getValue() != tabCour.getCell(1, j+1).editAsText().getText() ) {
             // en mode debug, on affiche a cote
            // a garder : FileRef.getRange(sTmp04).setValue(tabCour.getCell(1, j+1).editAsText().getText() );
            // en prod, on met la nouvelle valeur et on envoi un email
            FileRef.getRange(sTmp02).setValue( tabCour.getCell(1, j+1).editAsText().getText()) ;
            retour =  tabCour.getCell(0, j+1).editAsText().getText() ;
          }
          if ( FileRef.getRange(sTmp03).getValue() != tabCour.getCell(4, j+1).editAsText().getText() ) {
            // en mode debug, on affiche a cote
            // a garder :  FileRef.getRange(sTmp05).setValue(tabCour.getCell(4, j+1).editAsText().getText() );
          // en prod, on met la nouvelle valeur et on envoi un email
            FileRef.getRange(sTmp02).setValue(tabCour.getCell(1, j+1).editAsText().getText()) ;
            retour =  tabCour.getCell(0, j+1).editAsText().getText() ;
          }
        }
        glbCpt01 ++ ;
      }
    return retour ;
}


//*****************************************
// Fonction Metier
//*****************************************
function LireDoc(FileRef ) {
  
    // ouverture de copie achat vivant
  var cstURL = 'https://docs.google.com/document/d/1zfs7Q5I1W7t6barkJbyAJDVlsBjucK1mW4L5pCCqSTA/edit' ;
  
  var doc = DocumentApp.openByUrl(cstURL);

  var allTables = doc.getBody().getTables() ;
  var change = -1 ;
  
  glbCpt01 = 1 ;
  // on balaye toutes les tables
  for (i=0 ; i< allTables.length ; i++) {
    var tableCour = allTables[i];
    //  **************************
    //  **************************
    // on verifie que c'est bien une table de données
    if ( tableCour.getCell(0, 0).editAsText().getText() == 'SEMAINE' ) {
      change = LectureTable  (FileRef , tableCour) ;
      if ( change !=-1) return change ;
    }
  }
  return change ;
}

//*****************************************
// Fonction Principale
//*****************************************

function Main() {
 
   var cstFile = 'RefAchatVivant' ;
  
   var d = new Date();
   
  // Mail de controle à 19H UTC
  if (( d.getUTCHours() == 19 ) && ( d.getUTCMinutes() < 5 ) ){
      MailApp.sendEmail('laurent.mci85@orange.fr', 'Control Application AchatVivant', "OK");
  }
   
  var haBDs  = DriveApp.getFilesByName(cstFile)
  var bInit = haBDs.hasNext() ;
  if(! bInit){
     var FileRef= SpreadsheetApp.create(cstFile) ;
  }
  else{
    var FileRef = SpreadsheetApp.open(haBDs.next()) ;
  }
  
  var change = LireDoc ( FileRef ) ;
  if ( change != -1 ) {
    MailApp.sendEmail('laurent.mci85@orange.fr', 'Changement semaine : '+ change.toString() , "OK");
   }
}

