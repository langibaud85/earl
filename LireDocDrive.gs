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

function LireDoc(FileRef ) {
  
    // ouverture de copie achat vivant
  var cstURL = 'https://docs.google.com/document/d/1zfs7Q5I1W7t6barkJbyAJDVlsBjucK1mW4L5pCCqSTA/edit' ;
  
  
  var cstWeek = getWeek() ;
  
  var doc = DocumentApp.openByUrl(cstURL);

  var allTables = doc.getBody().getTables() ;
  
  var Cpt01 = 1 ;
  // on balaye toutes les tables
  for (i=0 ; i< allTables.length ; i++) {
    var tableCour = allTables[i];
    for (j=0 ; j < 7 ; j++) {
      sTmp01='A'+Cpt01.toString() ;
      sTmp02='B'+Cpt01.toString() ;
      sTmp03='C'+Cpt01.toString() ;
  
      
      if (  FileRef.getRange(sTmp01).getValue() == 0  ) {
        // On affecte les valeurs pour la semaine
        FileRef.getRange(sTmp01).setValue(tableCour.getCell(0, j+2).editAsText().getText() );
        // On affecte les valeurs pour Bete LS
        FileRef.getRange(sTmp02).setValue(tableCour.getCell(1, j+2).editAsText().getText() );
        // On affecte les valeurs pour Bete TRAD
        FileRef.getRange(sTmp03).setValue(tableCour.getCell(4, j+2).editAsText().getText() );
       
      }else
      {
        // on regarde si le réfAchatVivant = 0 sinon on passe
        if ( FileRef.getRange(sTmp02).getValue() == 0 ) {
        }
      }
      Cpt01 ++ ;
    }
  }
}

  // modif du 30/04/2020 ACER MCI  001
  // modif du 30/04/2020 ACER MCI  002


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
     var FileTmp = SpreadsheetApp.create(cstFile) ;
  }
  //Does exist
  else{
    var FileRef = SpreadsheetApp.open(haBDs.next()) ;
  }
  
  LireDoc ( FileRef ) ;
  
}

