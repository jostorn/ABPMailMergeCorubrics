 //////**variables globals**/////

  //CC
  var resultatsCC = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("resultats");
  var descripcionsCC = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("Rúbrica");

  //
  //M05 UFs i RAs
  //
  //M05_UF1_RA2 
  var resultatsM05_UF1_RA2 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("resultats");
  var descripcionsM05_UF1_RA2 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("Rúbrica");

  //M05_UF1_RA3 
  var resultatsM05_UF1_RA3 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("resultats");  
  var descripcionsM05_UF1_RA3 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("Rúbrica");

  

  
  //
  //M03 UFs i RAs
  //
  //M03_UF1_RA1
  var resultatsM03_UF1_RA1 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("resultats");
  var descripcionsM03_UF1_RA1 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("Rúbrica");

  //M03_UF1_RA2 
  var resultatsM03_UF1_RA2 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("resultats");
  var descripcionsM03_UF1_RA2 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("Rúbrica");

  //M03_UF3_RA1 
  var resultatsM03_UF3_RA1 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("resultats");  
  var descripcionsM03_UF3_RA1 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("Rúbrica");

  //M03_UF4_RA1 
  var resultatsM03_UF4_RA1 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("resultats");
  var descripcionsM03_UF4_RA1 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("Rúbrica");
  
  //M03_UF6_RA1 
  var resultatsM03_UF6_RA1 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa-aaaaaaaaaaaaaaaaaaaaaaaaaaa-7JVY").getSheetByName("resultats");
  var descripcionsM03_UF6_RA1 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa-aaaaaaaaaaaaaaaaaaaaaaaaaaa-7JVY").getSheetByName("Rúbrica");

  //M03_UF6_RA3 
  var resultatsM03_UF6_RA3 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa-aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("resultats");
  var descripcionsM03_UF6_RA3 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa-aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("Rúbrica");


/*  Angles */

//M04_UF1_RA1 
  var resultatsM04_UF1_RA1 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa-aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("resultats");
  var descripcionsM04_UF1_RA1 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa-aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("Rúbrica");


//M04_UF1_RA2 
  var resultatsM04_UF1_RA2 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("resultats");
  var descripcionsM04_UF1_RA2 = SpreadsheetApp.openById("aaaaaaaaaaaaaaaaaaaaaaaaaaa").getSheetByName("Rúbrica");



  
  //pantilla
  //var informePlantilla = DriveApp.getFileById("aaaaaaaaaaaaaaaaaaaaaaaaaaa-aaaaaaaaaaaaaaaaaaaaaaaaaaa");
  var informePlantilla = DriveApp.getFileById("aaaaaaaaaaaaaaaaaaaaaaaaaaa-aaaaaaaaaaaaaaaaaaaaaaaaaaa");
  
  //destinació
  var carpetaDesti = DriveApp.getFolderById("aaaaaaaaaaaaaaaaaaaaaaaaaaa-Yn");


/**************************************/

 function obtenir_resultats_RA(fullresultats,fulldescripcions,i,body,descOffsetFil,resulOffset,nota,item)
 {

   var it,valor; 
   var valorC = fullresultats.getRange(i, resulOffset).getValue();
   var valorA = fullresultats.getRange(i, resulOffset+1).getValue();
   var valorP = fullresultats.getRange(i, resulOffset+2).getValue();

   
   valor = valorC*0.2+valorA*0.2 + valorP*0.6 ;
    
   if (valor>=1 && valor<2)
   {
     valor="Novell";
   }
   else if(valor>=2 && valor<3)
   {
       valor="Aprenent";
   }
   else if(valor>=3 && valor<4)
   {
       valor="Avançat";
   }      
   else if(valor==4)
   {
       valor="Expert";
   }  
   else
   {
     valor="No definit";
   }
  
   it = fulldescripcions.getRange(descOffsetFil,1).getValue();
              
    
   body.replaceText(nota,valor);
   body.replaceText(item,it);
        
  }  

 function obtenir_resultats_CC(fullresultats,fulldescripcions,i,body,descOffsetCol,descOffsetFil,resulOffset,nota,descripcio)
 {

   var desc,valor; 
   var valorC = fullresultats.getRange(i, resulOffset).getValue();   
   var valorA = fullresultats.getRange(i, resulOffset+1).getValue();
   var valorP = fullresultats.getRange(i, resulOffset+2).getValue();
   
   valor = (valorC*0.2+valorA*0.2 + valorP*0.6);
    
   if (valor>=1 && valor<=4)
   {
      desc= fulldescripcions.getRange(descOffsetFil,descOffsetCol-valor.toFixed(0)).getValue();
   }
   else
   {
      desc="No definit";
   }
      
     
   valor=valor/4*10;
      
   body.replaceText(nota,valor);
   body.replaceText(descripcio,desc);
        
  }  


//

function myFunction() {


  var descOffsetCol;
  var descOffsetFil;
  var resulOffset;
  var resultatRA;
  var ca,ra,caitem;
  var num;
  var fullresul,fulldescrip;
  
  // Iterar per tots els alumnes del full de resultats (ha ha 26 columnes)
  for (i = 4; i< 19;i++)
  {

    // Definim variables per a cada alumne
    var nomAlumne = resultatsCC.getRange(i, 2).getValue();
    
    //Fem una còpia de la plantilla amb el nom de l'alumne
    var idDoc = informePlantilla.makeCopy('SMX1_P2_InformeCompetencial_'+nomAlumne, carpetaDesti).getId();
      
    var body = DocumentApp.openById(idDoc).getBody(); //Obtenir el cos del document com una variable
        
    body.replaceText('##alumne##',nomAlumne); //Inserir el nom de l'alumne
   
  
    
    /******************** C O M P E T E N C I E S   C L A U************************************/
    fullresul=resultatsCC;
    fulldescrip=descripcionsCC;   
    
    descOffsetCol = 6;
    descOffsetFil = 3;
    resulOffset=6;
    obtenir_resultats_CC(fullresul,fulldescrip,i,body,descOffsetCol,descOffsetFil,resulOffset,
                      '##C1CC##','##DescripcioC1CC##');
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_CC(fullresul,fulldescrip,i,body,descOffsetCol,descOffsetFil,resulOffset,
                      '##C2CC##','##DescripcioC2CC##');
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_CC(fullresul,fulldescrip,i,body,descOffsetCol,descOffsetFil,resulOffset,
                      '##C3CC##','##DescripcioC3CC##');
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_CC(fullresul,fulldescrip,i,body,descOffsetCol,descOffsetFil,resulOffset,
                      '##C4CC##','##DescripcioC4CC##');
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_CC(fullresul,fulldescrip,i,body,descOffsetCol,descOffsetFil,resulOffset,
                      '##C5CC##','##DescripcioC5CC##');
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_CC(fullresul,fulldescrip,i,body,descOffsetCol,descOffsetFil,resulOffset,
                      '##C6CC##','##DescripcioC6CC##');
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_CC(fullresul,fulldescrip,i,body,descOffsetCol,descOffsetFil,resulOffset,
                      '##C7CC##','##DescripcioC7CC##');
   
      /****/
      
      
      
      
    /* X A R X E S  */
    /*********************M05_UF1_RA2*************************/
    
    fullresul=resultatsM05_UF1_RA3;
    fulldescrip=descripcionsM05_UF1_RA3;
    
    descOffsetFil = 3; //on comença les descripcions
    resulOffset=6;     // del full resultats
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M05_UF1_RA2_1##','##itemM05_UF1_RA2_1##');

    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M05_UF1_RA2_2##','##item05_UF1_RA2_2##');
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M05_UF1_RA2_3##','##item05_UF1_RA2_3##');

    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M05_UF1_RA2_4##','##item05_UF1_RA2_4##');

    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M05_UF1_RA2_5##','##item05_UF1_RA2_5##');


    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M05_UF1_RA2_6##','##item05_UF1_RA2_6##');
    
   
    /*************************************************************/
      
      
      
    /*********************M05_UF1_RA3*************************/
    fullresul=resultatsM01_UF1_RA2;
    fulldescrip=descripcionsM01_UF1_RA2;

    descOffsetFil = 3; //on comença les descripcions
    resulOffset=6;     // del full resultats
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M05_UF1_RA3_1##','##itemM05_UF1_RA3_1##');
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M05_UF1_RA3_2##','##itemM05_UF1_RA3_2##');

    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M05_UF1_RA3_3##','##itemM05_UF1_RA3_3##');
      
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M05_UF1_RA3_4##','##itemM05_UF1_RA3_4##');
      

          
      /*************************************************************/
          

    /*** F I   X A R X E S.  ***/
   

    
    /* O F I M A T I C A  */
    
    //M03_UF1_RA1
    ra="M03_UF1_RA1";
    fullresul="resultats"+ra;
    fulldescrip="descripcions"+ra;
    
    descOffsetFil = 3; 
    resulOffset = 6;

    ca="##"+ra+"_";
    caitem="##item"+ra+"_";
      
    num=1;
    do {
      obtenir_resultats_RA(eval(fullresul),eval(fulldescrip),i,body,descOffsetFil,resulOffset,ca+num+"##",caitem+num+"##");
      descOffsetFil++;
      resulOffset=resulOffset+3;
      num++;
    }while (num <=6); 

    
         /*************************************************************/
   
    ra="M03_UF1_RA2";
    fullresul="resultats"+ra;
    fulldescrip="descripcions"+ra;
    
    descOffsetFil = 3; 
    resulOffset = 6;

    ca="##"+ra+"_";
    caitem="##item"+ra+"_";
      
    num=1;
    do {
      obtenir_resultats_RA(eval(fullresul),eval(fulldescrip),i,body,descOffsetFil,resulOffset,ca+num+"##",caitem+num+"##");
      descOffsetFil++;
      resulOffset=resulOffset+3;
      num++;
    }while (num <=2); 

    
         /*************************************************************/
    ra="M03_UF3_RA1";
    fullresul="resultats"+ra;
    fulldescrip="descripcions"+ra;
    
    descOffsetFil = 3; 
    resulOffset = 6;

    ca="##"+ra+"_";
    caitem="##item"+ra+"_";
      
    num=1;
    do {
      obtenir_resultats_RA(eval(fullresul),eval(fulldescrip),i,body,descOffsetFil,resulOffset,ca+num+"##",caitem+num+"##");
      descOffsetFil++;
      resulOffset=resulOffset+3;
      num++;
    }while (num <=1); 

    
         /*************************************************************/
    ra="M03_UF4_RA1";
    fullresul="resultats"+ra;
    fulldescrip="descripcions"+ra;
    
    descOffsetFil = 3; 
    resulOffset = 6;

    ca="##"+ra+"_";
    caitem="##item"+ra+"_";
      
    num=1;
    do {
      obtenir_resultats_RA(eval(fullresul),eval(fulldescrip),i,body,descOffsetFil,resulOffset,ca+num+"##",caitem+num+"##");
      descOffsetFil++;
      resulOffset=resulOffset+3;
      num++;
    }while (num <=3); 

    
         /*************************************************************/
    ra="M03_UF6_RA1";
    fullresul="resultats"+ra;
    fulldescrip="descripcions"+ra;
    
    descOffsetFil = 3; 
    resulOffset = 6;

    ca="##"+ra+"_";
    caitem="##item"+ra+"_";
      
    num=1;
    do {
      obtenir_resultats_RA(eval(fullresul),eval(fulldescrip),i,body,descOffsetFil,resulOffset,ca+num+"##",caitem+num+"##");
      descOffsetFil++;
      resulOffset=resulOffset+3;
      num++;
    }while (num <=4); 

    
         /*************************************************************/
    ra="M03_UF6_RA3";
    fullresul="resultats"+ra;
    fulldescrip="descripcions"+ra;
    
    descOffsetFil = 3; 
    resulOffset = 6;

    ca="##"+ra+"_";
    caitem="##item"+ra+"_";
      
    num=1;
    do {
      obtenir_resultats_RA(eval(fullresul),eval(fulldescrip),i,body,descOffsetFil,resulOffset,ca+num+"##",caitem+num+"##");
      descOffsetFil++;
      resulOffset=resulOffset+3;
      num++;
    }while (num <=1); 

    
         /*************************************************************/


  /* A N  G L E S */

    /*********************M04_UF1_RA1*************************/
  
    fullresul=resultatsM04_UF1_RA1;
    fulldescrip=descripcionsM04_UF1_RA1;
    
    descOffsetFil = 3; //on comença les descripcions
    resulOffset=6;     // del full resultats
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M04_UF1_RA1_1##','##itemM04_UF1_RA1_1##');

    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M04_UF1_RA1_2##','##itemM04_UF1_RA1_2##');
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M04_UF1_RA1_3##','##itemM04_UF1_RA1_3##');

    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M04_UF1_RA1_4##','##itemM04_UF1_RA1_4##');
    
    /*************************************************************/
      
      
    /*********************M04_UF1_RA2*************************/
    fullresul=resultatsM04_UF1_RA2;
    fulldescrip=descripcionsM04_UF1_RA2;

    descOffsetFil = 3; //on comença les descripcions
    resulOffset=6;     // del full resultats
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M04_UF1_RA2_1##','##itemM04_UF1_RA2_1##');
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M04_UF1_RA2_2##','##itemM04_UF1_RA2_2##');

    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M04_UF1_RA2_3##','##itemM04_UF1_RA2_3##');
      
    descOffsetFil++;
    resulOffset=resulOffset+3;
    obtenir_resultats_RA(fullresul,fulldescrip,i,body,descOffsetFil,resulOffset,
                      '##M04_UF1_RA2_4##','##itemM04_UF1_RA2_4##');
      
          
      /*************************************************************/
          

    /*** F I   A N G L E S  ***/

       


   }
}



