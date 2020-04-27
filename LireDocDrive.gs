function LireDoc() {
  
    // ouverture de PFE
  var cstFilePFE = '1RHB1CT1X5cw7o4RC11oKpRdq7T3OvdmyVniBT3wDKPQ' ;
  var doc = DocumentApp.openById(cstFilePFE);

// Log the number of elements in the document.
  var body = doc.getBody() ;
  
  var toto = body.getTables() ;
  
  var titi = toto[0] ;

// modif linuix 27/04/2020

  // lolo win
  
  var tutu = titi.getCell(0, 0).editAsText().getText()  ;
  
  Logger.log (tutu) ;
}
