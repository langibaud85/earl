// Cahier des charges
// On cherche le premier changement et on envoie un email pcp et un email secondaire
// il y a mail de verification de fonctiopnnement tout les soirs à 21h00
// il faut detecter le changement d'année : 01/2020





//***************************************
//************************** BIBLIOTHEQUE
//***************************************
// Returns the ISO week of the date.
getWeek = function() {
  var date = new Date();
  date.setHours(0, 0, 0, 0);
  // Thursday in current week decides the year.
  date.setDate(date.getDate() + 3 - (date.getDay() + 6) % 7);
  // January 4 is always in week 1.
  var week1 = new Date(date.getFullYear(), 0, 4);
  // Adjust to Thursday in week 1 and count number of weeks from date to week1.
  return 1 + Math.round(((date.getTime() - week1.getTime()) / 86400000
                        - 3 + (week1.getDay() + 6) % 7) / 7);
}


//*****************************************
// Fonction Metier
//*****************************************


function LectureTable(FileRef , tabCour) { 

  
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
             FileRef.getRange(sTmp04).setValue(tabCour.getCell(1, j+1).editAsText().getText() );
          }
          if ( FileRef.getRange(sTmp03).getValue() != tabCour.getCell(4, j+1).editAsText().getText() ) {
             FileRef.getRange(sTmp05).setValue(tabCour.getCell(4, j+1).editAsText().getText() );
          }
        }
        glbCpt01 ++ ;
      }
    
}


//*****************************************
// Fonction Metier
//*****************************************
function LireDoc(FileRef ) {
  
    // ouverture de copie achat vivant
  var cstURL = 'https://docs.google.com/document/d/1zfs7Q5I1W7t6barkJbyAJDVlsBjucK1mW4L5pCCqSTA/edit' ;
   
  
  
  var cstWeek = getWeek() ;
  
  var doc = DocumentApp.openByUrl(cstURL);

  var allTables = doc.getBody().getTables() ;
  
  glbCpt01 = 1 ;
  // on balaye toutes les tables
  for (i=0 ; i< allTables.length ; i++) {
    var tableCour = allTables[i];
    // onverifie que c'est bien une table de données
    if ( tableCour.getCell(0, 0).editAsText().getText() == 'SEMAINE' ) {
      LectureTable  (FileRef , tableCour) ;
    }
  }
}


//*****************************************
// Fonction Metier
//*****************************************

function Main() {
 
   var cstFile = 'RefAchatVivant' ;
  
   var d = new Date();
   
  // Mail de control à 19H UTC
  if ( d.getUTCHours() == 19 ) {
    if ( d.getUTCMinutes() < 5 ) {
      MailApp.sendEmail('laurent.mci85@orange.fr', 'Control Application AchatVivant', "OK");
    }
  }
   
   var haBDs  = DriveApp.getFilesByName(cstFile)
  //Does not exist
  var bInit = haBDs.hasNext() ;
  if(! bInit){
     var FileRef= SpreadsheetApp.create(cstFile) ;
  }
  //Does exist
  else{
    var FileRef = SpreadsheetApp.open(haBDs.next()) ;
  }
  
  LireDoc ( FileRef ) ;
  
}

