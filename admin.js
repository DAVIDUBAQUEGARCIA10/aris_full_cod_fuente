/*=======================================================================

     =====================  Listar - Uno o Todos Etiquetas  =======================

==============================================================================*/

function listarEtiquetas(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Riesgos");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:M "+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][12]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion1:datos[i][3],
        descripcion2:datos[i][4],
        descripcion3:datos[i][5],
        descripcion4:datos[i][6],
        descripcion5:datos[i][7],
        descripcion6:datos[i][8],
        descripcion7:datos[i][9],
        descripcion8:datos[i][10],
        descripcion9:datos[i][11]        
   
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][12]=="A"){
      
      if(rol=="ADMIN"){
      
        var botones = "<a class='btn btn-primary posicionBoton' title='Editar Etiqueta' "+
                       "data-bs-toggle='modal' data-bs-target='#editarEtiquetas' "+
                       "onclick='verEditarEtiqueta("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar Etiqueta' "+
                       "onclick='eliminarEtiqueta("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
        
      }else if(rol=="SPECIAL"){
      
        var botones = "<a class='btn btn-primary posicionBoton' title='Editar Etiqueta' "+
                       "data-bs-toggle='modal' data-bs-target='#editarEtiquetas' "+
                       "onclick='verEditarEtiqueta("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      }else if(rol=="GESTOR"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar Etiqueta' "+
                       "data-bs-toggle='modal' data-bs-target='#editarEtiqueta' "+
                       "onclick='verEditarEtiqueta("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
      }
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion1:datos[i][3],
        descripcion2:datos[i][4],
        descripcion3:datos[i][5],
        descripcion4:datos[i][6],
        descripcion5:datos[i][7],
        descripcion6:datos[i][8],
        descripcion7:datos[i][9],
        descripcion8:datos[i][10],
        descripcion9:datos[i][11],        
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Etiqueta  =======================

==============================================================================*/

function crearEtiqueta(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Riesgos");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearEtiqueta!="" && form.descipcionCrearEtiqueta!="" && form.descipcionCrearEtiqueta1!=""&& form.descipcionCrearEtiqueta2!=""&& form.descipcionCrearEtiqueta3!=""&& form.descipcionCrearEtiqueta4!=""&& form.descipcionCrearEtiqueta5!=""&& form.descipcionCrearEtiqueta6!=""&& form.descipcionCrearEtiqueta7!=""&& form.descipcionCrearEtiqueta8!=""&& form.descipcionCrearEtiqueta9!=""&& form.descipcionCrearEtiqueta10!="" && form.txtAccion=="crearEtiqueta"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearEtiqueta);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearEtiqueta);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearEtiqueta1);
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearEtiqueta2);
    hoja.getRange("F"+(fila+1)).setValue(form.descipcionCrearEtiqueta3);
    hoja.getRange("G"+(fila+1)).setValue(form.descipcionCrearEtiqueta4); 
    hoja.getRange("H"+(fila+1)).setValue(form.descipcionCrearEtiqueta5);
    hoja.getRange("I"+(fila+1)).setValue(form.descipcionCrearEtiqueta6);
    hoja.getRange("J"+(fila+1)).setValue(form.descipcionCrearEtiqueta7);
    hoja.getRange("K"+(fila+1)).setValue(form.descipcionCrearEtiqueta8);  
    hoja.getRange("L"+(fila+1)).setValue(form.descipcionCrearEtiqueta9);                                 
    hoja.getRange("M"+(fila+1)).setValue("A");
      
  
    return "crearEtiquetaOk";
  }else if(form.nombreEditarEtiqueta!="" && form.descipcionEditarEtiqueta!="" && form.descipcionEditarEtiqueta1!=""&& form.descipcionEditarEtiqueta2!=""&& form.descipcionEditarEtiqueta3!=""&& form.descipcionEditarEtiqueta4!=""&& form.descipcionEditarEtiqueta5!=""&& form.descipcionEditarEtiqueta6!=""&& form.descipcionEditarEtiqueta7!=""&& form.descipcionEditarEtiqueta8!=""&& form.descipcionEditarEtiqueta9!="" && form.txtAccion=="editarEtiqueta" && form.idEditarEtiqueta!=""){
     
    var datos = hoja.getRange("A2:M"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarEtiqueta==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarEtiqueta);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarEtiqueta);
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarEtiqueta1); 
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditarEtiqueta2); 
        hoja.getRange("F"+(i+2)).setValue(form.descipcionEditarEtiqueta3); 
        hoja.getRange("G"+(i+2)).setValue(form.descipcionEditarEtiqueta4); 
        hoja.getRange("H"+(i+2)).setValue(form.descipcionEditarEtiqueta5); 
        hoja.getRange("I"+(i+2)).setValue(form.descipcionEditarEtiqueta6); 
        hoja.getRange("J"+(i+2)).setValue(form.descipcionEditarEtiqueta7); 
        hoja.getRange("K"+(i+2)).setValue(form.descipcionEditarEtiqueta8);
        hoja.getRange("L"+(i+2)).setValue(form.descipcionEditarEtiqueta9);         
                                                 
        return "editarEtiquetaOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}




/*=======================================================================

     =====================  Listar - Uno o Todos Controles  =======================

==============================================================================*/

function listarAreas(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Controles");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:AY"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][33]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6],
        descripcion6:datos[i][7],
        descripcion7:datos[i][8],
        descripcion8:datos[i][9],
        descripcion9:datos[i][10],
        descripcion10:datos[i][11],
        descripcion11:datos[i][12],
        descripcion12:datos[i][13],                                                                
        descripcion13:datos[i][14],
        descripcion14:datos[i][15],     
        descripcion15:datos[i][16],
        descripcion16:datos[i][17],
        descripcion17:datos[i][18],
        descripcion18:datos[i][19],
        descripcion19:datos[i][20],
        descripcion20:datos[i][21],
        descripcion21:datos[i][22],
        descripcion22:datos[i][23],
        descripcion23:datos[i][24],
        descripcion24:datos[i][25],
        descripcion25:datos[i][26],
        descripcion26:datos[i][27],
        descripcion27:datos[i][28],
        descripcion28:datos[i][29],
        descripcion29:datos[i][30],
        descripcion30:datos[i][31],
        descripcion31:datos[i][32]                                                              
              
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][33]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar Área' "+
                       "data-bs-toggle='modal' data-bs-target='#editarAreas' "+
                       "onclick='verEditarArea("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar Área' "+
                       "onclick='eliminarArea("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar Area' "+
                       "data-bs-toggle='modal' data-bs-target='#editarAreas' "+
                       "onclick='verEditarArea("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }else if(rol=="GESTOR"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar Área' "+
                       "data-bs-toggle='modal' data-bs-target='#editarAreas' "+
                       "onclick='verEditarArea("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6], 
        descripcion6:datos[i][7],  
        descripcion7:datos[i][8],  
        descripcion8:datos[i][9],  
        descripcion9:datos[i][10],  
        descripcion10:datos[i][11],  
        descripcion11:datos[i][12],  
        descripcion12:datos[i][13],  
        descripcion13:datos[i][14], 
        descripcion14:datos[i][15], 
        descripcion15:datos[i][16], 
        descripcion16:datos[i][17], 
        descripcion17:datos[i][18], 
        descripcion18:datos[i][19], 
        descripcion19:datos[i][20], 
        descripcion20:datos[i][21], 
        descripcion21:datos[i][22], 
        descripcion22:datos[i][23], 
        descripcion23:datos[i][24], 
        descripcion24:datos[i][25], 
        descripcion25:datos[i][26], 
        descripcion26:datos[i][27], 
        descripcion27:datos[i][28], 
        descripcion28:datos[i][29], 
        descripcion29:datos[i][30], 
        descripcion30:datos[i][31], 
        descripcion31:datos[i][32],
        acciones:botones              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Controles  =======================

==============================================================================*/

function crearArea(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Controles");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearArea!="" && form.descipcionCrearArea!="" && form.descipcionCrearArea2!="" && form.descipcionCrearArea3!="" && form.descipcionCrearArea4!=""&& form.descipcionCrearArea5!="" && form.descipcionCrearArea6!="" && form.descipcionCrearArea7!="" && form.descipcionCrearArea8!="" && form.descipcionCrearArea9!="" && form.descipcionCrearArea10!="" && form.descipcionCrearArea11!="" && form.descipcionCrearArea12!="" && form.descipcionCrearArea13!="" && form.descipcionCrearArea14!="" && form.descipcionCrearArea15!="" && form.descipcionCrearArea16!="" && form.descipcionCrearArea17!="" && form.descipcionCrearArea18!="" && form.descipcionCrearArea19!="" && form.descipcionCrearArea20!="" && form.descipcionCrearArea21!="" && form.descipcionCrearArea22!="" && form.descipcionCrearArea23!="" && form.descipcionCrearArea24!="" && form.descipcionCrearArea25!="" && form.descipcionCrearArea26!="" && form.descipcionCrearArea27!="" && form.descipcionCrearArea28!="" && form.descipcionCrearArea29!="" && form.descipcionCrearArea30!="" && form.descipcionCrearArea31!="" && form.txtAccion=="crearArea"){
     

    var factoresSeleccionados18 = form.descipcionCrearArea18;
    if (Array.isArray(factoresSeleccionados18)) {
      factoresSeleccionados18 = factoresSeleccionados18.join(",");
    }

    var factoresSeleccionados29 = form.descipcionCrearArea29;
    if (Array.isArray(factoresSeleccionados29)) {
      factoresSeleccionados29 = factoresSeleccionados29.join(",");
    }  

    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearArea);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearArea);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearArea2);
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearArea3);
    hoja.getRange("F"+(fila+1)).setValue(form.descipcionCrearArea4); 
    hoja.getRange("G"+(fila+1)).setValue(form.descipcionCrearArea5);  
    hoja.getRange("H"+(fila+1)).setValue(form.descipcionCrearArea6);  
    hoja.getRange("I"+(fila+1)).setValue(form.descipcionCrearArea7);  
    hoja.getRange("J"+(fila+1)).setValue(form.descipcionCrearArea8);  
    hoja.getRange("K"+(fila+1)).setValue(form.descipcionCrearArea9);  
    hoja.getRange("L"+(fila+1)).setValue(form.descipcionCrearArea10);  
    hoja.getRange("M"+(fila+1)).setValue(form.descipcionCrearArea11);  
    hoja.getRange("N"+(fila+1)).setValue(form.descipcionCrearArea12);  
    hoja.getRange("O"+(fila+1)).setValue(form.descipcionCrearArea13); 
    hoja.getRange("P"+(fila+1)).setValue(form.descipcionCrearArea14);  
    hoja.getRange("Q"+(fila+1)).setValue(form.descipcionCrearArea15);
    hoja.getRange("R"+(fila+1)).setValue(form.descipcionCrearArea16);
    hoja.getRange("S"+(fila+1)).setValue(form.descipcionCrearArea17);
    hoja.getRange("T"+(fila+1)).setValue(factoresSeleccionados18);
    hoja.getRange("U"+(fila+1)).setValue(form.descipcionCrearArea19);
    hoja.getRange("V"+(fila+1)).setValue(form.descipcionCrearArea20);
    hoja.getRange("W"+(fila+1)).setValue(form.descipcionCrearArea21);
    hoja.getRange("X"+(fila+1)).setValue(form.descipcionCrearArea22);
    hoja.getRange("Y"+(fila+1)).setValue(form.descipcionCrearArea23);
    hoja.getRange("Z"+(fila+1)).setValue(form.descipcionCrearArea24);
    hoja.getRange("AA"+(fila+1)).setValue(form.descipcionCrearArea25);
    hoja.getRange("AB"+(fila+1)).setValue(form.descipcionCrearArea26);
    hoja.getRange("AC"+(fila+1)).setValue(form.descipcionCrearArea27);
    hoja.getRange("AD"+(fila+1)).setValue(form.descipcionCrearArea28);
    hoja.getRange("AE"+(fila+1)).setValue(factoresSeleccionados29);
    hoja.getRange("AF"+(fila+1)).setValue(form.descipcionCrearArea30);
    hoja.getRange("AG"+(fila+1)).setValue(form.descipcionCrearArea31);        
    hoja.getRange("AH"+(fila+1)).setValue("A");
  
    return "crearAreaOk";
    
    
  }else if(form.nombreEditarAreas!="" && form.descipcionEditarAreas!=""&& form.descipcionEditarAreas2!="" && form.descipcionEditarAreas3!="" && form.descipcionEditarAreas4!="" && form.descipcionEditarAreas5!=""&& form.descipcionEditarAreas6!=""&& form.descipcionEditarAreas7!=""&& form.descipcionEditarAreas8!=""&& form.descipcionEditarAreas9!=""&& form.descipcionEditarAreas10!=""&& form.descipcionEditarAreas11!=""&& form.descipcionEditarAreas12!=""&& form.descipcionEditarAreas13!=""&& form.descipcionEditarAreas14!=""&& form.descipcionEditarAreas15!=""&& form.descipcionEditarAreas16!=""&& form.descipcionEditarAreas17!=""&& form.descipcionEditarAreas18!=""&& form.descipcionEditarAreas19!=""&& form.descipcionEditarAreas20!=""&& form.descipcionEditarAreas21!=""&& form.descipcionEditarAreas22!=""&& form.descipcionEditarAreas23!=""&& form.descipcionEditarAreas24!=""&& form.descipcionEditarAreas25!=""&& form.descipcionEditarAreas26!=""&& form.descipcionEditarAreas27!=""&& form.descipcionEditarAreas28!=""&& form.descipcionEditarAreas29!=""&& form.descipcionEditarAreas30!=""&& form.descipcionEditarAreas31!="" && form.txtAccion=="editarAreas" && form.idEditarAreas!=""){

    var factoresSeleccionados18editar = form.descipcionEditarAreas18;
    if (Array.isArray(factoresSeleccionados18editar)) {
      factoresSeleccionados18editar = factoresSeleccionados18editar.join(",");
    }

    var factoresSeleccionados29editar = form.descipcionEditarAreas29;
    if (Array.isArray(factoresSeleccionados29editar)) {
      factoresSeleccionados29editar = factoresSeleccionados29editar.join(",");
    }


    var datos = hoja.getRange("A2:AY"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarAreas==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarAreas);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarAreas); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarAreas2);
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditarAreas3);
        hoja.getRange("F"+(i+2)).setValue(form.descipcionEditarAreas4);
        hoja.getRange("G"+(i+2)).setValue(form.descipcionEditarAreas5);  
        hoja.getRange("H"+(i+2)).setValue(form.descipcionEditarAreas6);  
        hoja.getRange("I"+(i+2)).setValue(form.descipcionEditarAreas7);  
        hoja.getRange("J"+(i+2)).setValue(form.descipcionEditarAreas8);  
        hoja.getRange("K"+(i+2)).setValue(form.descipcionEditarAreas9);  
        hoja.getRange("L"+(i+2)).setValue(form.descipcionEditarAreas10);  
        hoja.getRange("M"+(i+2)).setValue(form.descipcionEditarAreas11);  
        hoja.getRange("N"+(i+2)).setValue(form.descipcionEditarAreas12);  
        hoja.getRange("O"+(i+2)).setValue(form.descipcionEditarAreas13); 
        hoja.getRange("P"+(i+2)).setValue(form.descipcionEditarAreas14); 
        hoja.getRange("Q"+(i+2)).setValue(form.descipcionEditarAreas15);
        hoja.getRange("R"+(i+2)).setValue(form.descipcionEditarAreas16);
        hoja.getRange("S"+(i+2)).setValue(form.descipcionEditarAreas17);
        hoja.getRange("T"+(i+2)).setValue(factoresSeleccionados18editar);
        hoja.getRange("U"+(i+2)).setValue(form.descipcionEditarAreas19);
        hoja.getRange("V"+(i+2)).setValue(form.descipcionEditarAreas20);
        hoja.getRange("W"+(i+2)).setValue(form.descipcionEditarAreas21);
        hoja.getRange("X"+(i+2)).setValue(form.descipcionEditarAreas22);
        hoja.getRange("Y"+(i+2)).setValue(form.descipcionEditarAreas23);
        hoja.getRange("Z"+(i+2)).setValue(form.descipcionEditarAreas24);
        hoja.getRange("AA"+(i+2)).setValue(form.descipcionEditarAreas25);
        hoja.getRange("AB"+(i+2)).setValue(form.descipcionEditarAreas26);
        hoja.getRange("AC"+(i+2)).setValue(form.descipcionEditarAreas27);
        hoja.getRange("AD"+(i+2)).setValue(form.descipcionEditarAreas28);
        hoja.getRange("AE"+(i+2)).setValue(factoresSeleccionados29editar);
        hoja.getRange("AF"+(i+2)).setValue(form.descipcionEditarAreas30);                      
        hoja.getRange("AG"+(i+2)).setValue(form.descipcionEditarAreas31);                                
        return "editarAreaOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Controles  =======================

==============================================================================*/
function eliminarAreaAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Controles");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:AH"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("AH"+(i+2)).setValue("D");
      
      return "eliminarAreaOK";
    }      
  } 

}


/*=======================================================================

     =====================  Listar - Uno o Todos Testeos  =======================

==============================================================================*/

function listartest(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Controles");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:AY"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][50]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6],
        descripcion6:datos[i][7],
        descripcion7:datos[i][8],
        descripcion8:datos[i][9],
        descripcion9:datos[i][10],
        descripcion10:datos[i][11],
        descripcion11:datos[i][12],
        descripcion12:datos[i][13],                                                                
        descripcion13:datos[i][14],
        descripcion14:datos[i][15],
        descripcion15:datos[i][16],
        descripcion16:datos[i][17],
        descripcion17:datos[i][18],
        descripcion18:datos[i][19],
        descripcion19:datos[i][20],
        descripcion20:datos[i][21],
        descripcion21:datos[i][22],
        descripcion22:datos[i][23],
        descripcion23:datos[i][24],
        descripcion24:datos[i][25],
        descripcion25:datos[i][26],
        descripcion26:datos[i][27],
        descripcion27:datos[i][28],
        descripcion28:datos[i][29],
        descripcion29:datos[i][30],
        descripcion30:datos[i][31],
        descripcion31:datos[i][32],                                
        descripcion32:datos[i][33],
        descripcion33:datos[i][34],
        descripcion34:datos[i][35],
        descripcion35:datos[i][36],
        descripcion36:datos[i][37],
        descripcion37:datos[i][38],
        descripcion38:datos[i][39],
        descripcion39:datos[i][40],
        descripcion40:datos[i][41],
        descripcion41:datos[i][42],
        descripcion42:datos[i][43],
        descripcion42:datos[i][44],
        descripcion44:datos[i][45],
        descripcion45:datos[i][46],
        descripcion46:datos[i][47],
        descripcion47:datos[i][48],
        descripcion48:datos[i][49]                                        
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][50]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar test' "+
                       "data-bs-toggle='modal' data-bs-target='#editartest' "+
                       "onclick='verEditartest("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar test' "+
                       "onclick='eliminartest("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar test' "+
                       "data-bs-toggle='modal' data-bs-target='#editartest' "+
                       "onclick='verEditartest("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                            
      }else if(rol=="EVALUATOR"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar test' "+
                       "data-bs-toggle='modal' data-bs-target='#editartest' "+
                       "onclick='verEditartest("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>   </a>"+
                       "<a class='btn btn-danger' title='Eliminar test' "+
                       "onclick='eliminartest("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
                            
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6], 
        descripcion6:datos[i][7],  
        descripcion7:datos[i][8],  
        descripcion8:datos[i][9],  
        descripcion9:datos[i][10],  
        descripcion10:datos[i][11],  
        descripcion11:datos[i][12],  
        descripcion12:datos[i][13],  
        descripcion13:datos[i][14], 
        descripcion14:datos[i][15], 
        descripcion15:datos[i][16],
        descripcion16:datos[i][17],
        descripcion17:datos[i][18],
        descripcion18:datos[i][19],
        descripcion19:datos[i][20],
        descripcion20:datos[i][21],
        descripcion21:datos[i][22],
        descripcion22:datos[i][23],
        descripcion23:datos[i][24],
        descripcion24:datos[i][25],
        descripcion25:datos[i][26],
        descripcion26:datos[i][27],
        descripcion27:datos[i][28],
        descripcion28:datos[i][29],
        descripcion29:datos[i][30],
        descripcion30:datos[i][31],
        descripcion31:datos[i][32],                                
        descripcion32:datos[i][33], 
        descripcion33:datos[i][34], 
        descripcion34:datos[i][35], 
        descripcion35:datos[i][36],
        descripcion36:datos[i][37],
        descripcion37:datos[i][38],
        descripcion38:datos[i][39],
        descripcion39:datos[i][40],
        descripcion40:datos[i][41],
        descripcion41:datos[i][42],
        descripcion42:datos[i][43],
        descripcion42:datos[i][44],
        descripcion44:datos[i][45],
        descripcion45:datos[i][46],
        descripcion46:datos[i][47],
        descripcion47:datos[i][48],
        descripcion48:datos[i][49],                                         
        acciones:botones     
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Testeo  =======================

==============================================================================*/

function creartest(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Controles");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCreartest!="" && form.descipcionCreartest!="" && form.descipcionCreartest2!="" && form.descipcionCreartest3!="" && form.descipcionCreartest4!=""&& form.descipcionCreartest5!="" && form.descipcionCreartest6!="" && form.descipcionCreartest7!="" && form.descipcionCreartest8!="" && form.descipcionCreartest9!="" && form.descipcionCreartest10!="" && form.descipcionCreartest11!="" && form.descipcionCreartest12!="" && form.descipcionCreartest13!=""&& form.descipcionCreartest14!=""&& form.descipcionCreartest15!=""&& form.descipcionCreartest16!=""&& form.descipcionCreartest17!=""&& form.descipcionCreartest18!=""&& form.descipcionCreartest19!=""&& form.descipcionCreartest20!=""&& form.descipcionCreartest21!=""&& form.descipcionCreartest22!=""&& form.descipcionCreartest23!=""&& form.descipcionCreartest24!=""&& form.descipcionCreartest25!=""&& form.descipcionCreartest26!=""&& form.descipcionCreartest27!=""&& form.descipcionCreartest28!=""&& form.descipcionCreartest29!=""&& form.descipcionCreartest30!=""&& form.descipcionCreartest31!=""&& form.descipcionCreartest32!=""&& form.descipcionCreartest33!=""&& form.descipcionCreartest34!=""&& form.descipcionCreartest35!=""&& form.descipcionCreartest36!=""&& form.descipcionCreartest37!=""&& form.descipcionCreartest38!=""&& form.descipcionCreartest39!=""&& form.descipcionCreartest40!=""&& form.descipcionCreartest41!=""&& form.descipcionCreartest42!=""&& form.descipcionCreartest43!=""&& form.descipcionCreartest44!=""&& form.descipcionCreartest45!=""&& form.descipcionCreartest46!=""&& form.descipcionCreartest47!=""&& form.descipcionCreartest48!="" && form.txtAccion=="creartest"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCreartest);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCreartest);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCreartest2);
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCreartest3);
    hoja.getRange("F"+(fila+1)).setValue(form.descipcionCreartest4); 
    hoja.getRange("G"+(fila+1)).setValue(form.descipcionCreartest5);  
    hoja.getRange("H"+(fila+1)).setValue(form.descipcionCreartest6);  
    hoja.getRange("I"+(fila+1)).setValue(form.descipcionCreartest7);  
    hoja.getRange("J"+(fila+1)).setValue(form.descipcionCreartest8);  
    hoja.getRange("K"+(fila+1)).setValue(form.descipcionCreartest9);  
    hoja.getRange("L"+(fila+1)).setValue(form.descipcionCreartest10);  
    hoja.getRange("M"+(fila+1)).setValue(form.descipcionCreartest11);  
    hoja.getRange("N"+(fila+1)).setValue(form.descipcionCreartest12);  
    hoja.getRange("O"+(fila+1)).setValue(form.descipcionCreartest13); 
    hoja.getRange("P"+(fila+1)).setValue(form.descipcionCreartest14);  
    hoja.getRange("Q"+(fila+1)).setValue(form.descipcionCreartest15); 
    hoja.getRange("R"+(fila+1)).setValue(form.descipcionCreartest16); 
    hoja.getRange("S"+(fila+1)).setValue(form.descipcionCreartest17); 
    hoja.getRange("T"+(fila+1)).setValue(form.descipcionCreartest18); 
    hoja.getRange("U"+(fila+1)).setValue(form.descipcionCreartest19); 
    hoja.getRange("V"+(fila+1)).setValue(form.descipcionCreartest20); 
    hoja.getRange("W"+(fila+1)).setValue(form.descipcionCreartest21); 
    hoja.getRange("X"+(fila+1)).setValue(form.descipcionCreartest22); 
    hoja.getRange("Y"+(fila+1)).setValue(form.descipcionCreartest23); 
    hoja.getRange("Z"+(fila+1)).setValue(form.descipcionCreartest24); 
    hoja.getRange("AA"+(fila+1)).setValue(form.descipcionCreartest25); 
    hoja.getRange("AB"+(fila+1)).setValue(form.descipcionCreartest26); 
    hoja.getRange("AC"+(fila+1)).setValue(form.descipcionCreartest27); 
    hoja.getRange("AD"+(fila+1)).setValue(form.descipcionCreartest28); 
    hoja.getRange("AE"+(fila+1)).setValue(form.descipcionCreartest29); 
    hoja.getRange("AF"+(fila+1)).setValue(form.descipcionCreartest30); 
    hoja.getRange("AG"+(fila+1)).setValue(form.descipcionCreartest31); 
    hoja.getRange("AH"+(fila+1)).setValue(form.descipcionCreartest32);
    hoja.getRange("AI"+(fila+1)).setValue(form.descipcionCreartest33);
    hoja.getRange("AJ"+(fila+1)).setValue(form.descipcionCreartest34);
    hoja.getRange("AK"+(fila+1)).setValue(form.descipcionCreartest35);
    hoja.getRange("AL"+(fila+1)).setValue(form.descipcionCreartest36);
    hoja.getRange("AM"+(fila+1)).setValue(form.descipcionCreartest37);
    hoja.getRange("AN"+(fila+1)).setValue(form.descipcionCreartest38);
    hoja.getRange("AO"+(fila+1)).setValue(form.descipcionCreartest39);
    hoja.getRange("AP"+(fila+1)).setValue(form.descipcionCreartest40);
    hoja.getRange("AQ"+(fila+1)).setValue(form.descipcionCreartest41);
    hoja.getRange("AR"+(fila+1)).setValue(form.descipcionCreartest42);
    hoja.getRange("AS"+(fila+1)).setValue(form.descipcionCreartest43);
    hoja.getRange("AT"+(fila+1)).setValue(form.descipcionCreartest44);
    hoja.getRange("AU"+(fila+1)).setValue(form.descipcionCreartest45);
    hoja.getRange("AV"+(fila+1)).setValue(form.descipcionCreartest46);
    hoja.getRange("AW"+(fila+1)).setValue(form.descipcionCreartest47);
    hoja.getRange("AX"+(fila+1)).setValue(form.descipcionCreartest48) 
    hoja.getRange("AY"+(fila+1)).setValue("A");
    return "creartestOk";
    
    
  }else if(form.nombreEditartest!="" && form.descipcionEditartest!=""&& form.descipcionEditartest2!="" && form.descipcionEditartest3!="" && form.descipcionEditartest4!="" && form.descipcionEditartest5!=""&& form.descipcionEditartest6!=""&& form.descipcionEditartest7!=""&& form.descipcionEditartest8!=""&& form.descipcionEditartest9!=""&& form.descipcionEditartest10!=""&& form.descipcionEditartest11!=""&& form.descipcionEditartest12!=""&& form.descipcionEditartest13!=""&& form.descipcionEditartest14!=""&& form.descipcionEditartest15!=""&& form.descipcionEditartest16!=""&& form.descipcionEditartest17!=""&& form.descipcionEditartest18!=""&& form.descipcionEditartest19!=""&& form.descipcionEditartest20!=""&& form.descipcionEditartest21!=""&& form.descipcionEditartest22!=""&& form.descipcionEditartest23!=""&& form.descipcionEditartest24!=""&& form.descipcionEditartest25!=""&& form.descipcionEditartest26!=""&& form.descipcionEditartest27!=""&& form.descipcionEditartest28!=""&& form.descipcionEditartest29!=""&& form.descipcionEditartest30!=""&& form.descipcionEditartest31!=""&& form.descipcionEditartest32!="" && form.descipcionEditartest33!=""&& form.descipcionEditartest34!=""&& form.descipcionEditartest35!=""&& form.descipcionEditartest36!=""&& form.descipcionEditartest37!=""&& form.descipcionEditartest38!=""&& form.descipcionEditartest39!=""&& form.descipcionEditartest40!=""&& form.descipcionEditartest41!=""&& form.descipcionEditartest42!=""&& form.descipcionEditartest43!=""&& form.descipcionEditartest44!=""&& form.descipcionEditartest45!=""&& form.descipcionEditartest46!=""&& form.descipcionEditartest47!=""&& form.descipcionEditartest48!=""&& form.txtAccion=="editartest" && form.idEditartest!=""){
     
    var datos = hoja.getRange("A2:AY"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditartest==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditartest);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditartest); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditartest2);
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditartest3);
        hoja.getRange("F"+(i+2)).setValue(form.descipcionEditartest4);
        hoja.getRange("G"+(i+2)).setValue(form.descipcionEditartest5);  
        hoja.getRange("H"+(i+2)).setValue(form.descipcionEditartest6);  
        hoja.getRange("I"+(i+2)).setValue(form.descipcionEditartest7);  
        hoja.getRange("J"+(i+2)).setValue(form.descipcionEditartest8);  
        hoja.getRange("K"+(i+2)).setValue(form.descipcionEditartest9);  
        hoja.getRange("L"+(i+2)).setValue(form.descipcionEditartest10);  
        hoja.getRange("M"+(i+2)).setValue(form.descipcionEditartest11);  
        hoja.getRange("N"+(i+2)).setValue(form.descipcionEditartest12);  
        hoja.getRange("O"+(i+2)).setValue(form.descipcionEditartest13); 
        hoja.getRange("P"+(i+2)).setValue(form.descipcionEditartest14); 
        hoja.getRange("Q"+(i+2)).setValue(form.descipcionEditartest15);
        hoja.getRange("R"+(i+2)).setValue(form.descipcionEditartest16);
        hoja.getRange("S"+(i+2)).setValue(form.descipcionEditartest17);
        hoja.getRange("T"+(i+2)).setValue(form.descipcionEditartest18);
        hoja.getRange("U"+(i+2)).setValue(form.descipcionEditartest19);
        hoja.getRange("V"+(i+2)).setValue(form.descipcionEditartest20);
        hoja.getRange("W"+(i+2)).setValue(form.descipcionEditartest21);
        hoja.getRange("X"+(i+2)).setValue(form.descipcionEditartest22);
        hoja.getRange("Y"+(i+2)).setValue(form.descipcionEditartest23);
        hoja.getRange("Z"+(i+2)).setValue(form.descipcionEditartest24);
        hoja.getRange("AA"+(i+2)).setValue(form.descipcionEditartest25);
        hoja.getRange("AB"+(i+2)).setValue(form.descipcionEditartest26);
        hoja.getRange("AC"+(i+2)).setValue(form.descipcionEditartest27);
        hoja.getRange("AD"+(i+2)).setValue(form.descipcionEditartest28);
        hoja.getRange("AE"+(i+2)).setValue(form.descipcionEditartest29);
        hoja.getRange("AF"+(i+2)).setValue(form.descipcionEditartest30);
        hoja.getRange("AG"+(i+2)).setValue(form.descipcionEditartest31);
        hoja.getRange("AH"+(i+2)).setValue(form.descipcionEditartest32);
        hoja.getRange("AI"+(i+2)).setValue(form.descipcionEditartest33);
        hoja.getRange("AJ"+(i+2)).setValue(form.descipcionEditartest34);
        hoja.getRange("AK"+(i+2)).setValue(form.descipcionEditartest35); 
        hoja.getRange("AL"+(i+2)).setValue(form.descipcionEditartest36); 
        hoja.getRange("AM"+(i+2)).setValue(form.descipcionEditartest37); 
        hoja.getRange("AN"+(i+2)).setValue(form.descipcionEditartest38); 
        hoja.getRange("AO"+(i+2)).setValue(form.descipcionEditartest39); 
        hoja.getRange("AP"+(i+2)).setValue(form.descipcionEditartest40); 
        hoja.getRange("AQ"+(i+2)).setValue(form.descipcionEditartest41); 
        hoja.getRange("AR"+(i+2)).setValue(form.descipcionEditartest42); 
        hoja.getRange("AS"+(i+2)).setValue(form.descipcionEditartest43); 
        hoja.getRange("AT"+(i+2)).setValue(form.descipcionEditartest44); 
        hoja.getRange("AU"+(i+2)).setValue(form.descipcionEditartest45); 
        hoja.getRange("AV"+(i+2)).setValue(form.descipcionEditartest46); 
        hoja.getRange("AW"+(i+2)).setValue(form.descipcionEditartest47); 
        hoja.getRange("AX"+(i+2)).setValue(form.descipcionEditartest48);        
        return "editartestOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar testeos  =======================

==============================================================================*/
function eliminartestAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Controles");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:AY"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("AY"+(i+2)).setValue("D");
      
      return "eliminartestOK";
    }      
  } 

}


/*=======================================================================

     =====================  Listar - Uno o Todos Causas  =======================

==============================================================================*/

function listarcausas(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Causas");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:J"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][9]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6],
        descripcion6:datos[i][7],
        descripcion7:datos[i][8]                              
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][9]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar causas' "+
                       "data-bs-toggle='modal' data-bs-target='#editarcausas' "+
                       "onclick='verEditarcausas("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar causas' "+
                       "onclick='eliminarcausas("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar causas' "+
                       "data-bs-toggle='modal' data-bs-target='#editarcausas' "+
                       "onclick='verEditarcausas("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      }else if(rol=="GESTOR"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar causas' "+
                       "data-bs-toggle='modal' data-bs-target='#editarcausas' "+
                       "onclick='verEditarcausas("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],                
        descripcion5:datos[i][6],
        descripcion6:datos[i][7],  
        descripcion7:datos[i][8],                     
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Causas  =======================

==============================================================================*/

function crearcausas(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Causas");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearcausas!="" && form.descipcionCrearcausas!="" && form.descipcionCrearcausas2!="" && form.descipcionCrearcausas3!="" && form.descipcionCrearcausas4!="" && form.descipcionCrearcausas5!="" && form.descipcionCrearcausas6!="" && form.descipcionCrearcausas7!="" && form.txtAccion=="crearcausas"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearcausas);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearcausas);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearcausas2);
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearcausas3);
    hoja.getRange("F"+(fila+1)).setValue(form.descipcionCrearcausas4); 
    hoja.getRange("G"+(fila+1)).setValue(form.descipcionCrearcausas5); 
    hoja.getRange("H"+(fila+1)).setValue(form.descipcionCrearcausas6);   
    hoja.getRange("I"+(fila+1)).setValue(form.descipcionCrearcausas7);                     
    hoja.getRange("J"+(fila+1)).setValue("A");
  
    return "crearcausasOk";
    
    
  }else if(form.nombreEditarcausas!="" && form.descipcionEditarcausas!=""&& form.descipcionEditarcausas2!="" && form.descipcionEditarcausas3!="" && form.descipcionEditarcausas4!="" && form.descipcionEditarcausas5!=""&& form.descipcionEditarcausas6!="" && form.descipcionEditarcausas7!="" && form.txtAccion=="editarcausas" && form.idEditarcausas!=""){
     
    var datos = hoja.getRange("A2:J"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarcausas==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarcausas);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarcausas); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarcausas2);
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditarcausas3);
        hoja.getRange("F"+(i+2)).setValue(form.descipcionEditarcausas4);
        hoja.getRange("G"+(i+2)).setValue(form.descipcionEditarcausas5);
        hoja.getRange("H"+(i+2)).setValue(form.descipcionEditarcausas6);
        hoja.getRange("I"+(i+2)).setValue(form.descipcionEditarcausas7);                        
        return "editarcausasOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Causas  =======================

==============================================================================*/
function eliminarcausasAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Causas");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:J"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("J"+(i+2)).setValue("D");
      
      return "eliminarcausasOK";
    }      
  } 
}




/*=======================================================================

     =====================  Listar - Uno o Todos Actividad Economica  =======================

==============================================================================*/

function listaracti(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Actividad Economica");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:F"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][5]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4]                             
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][5]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar acti' "+
                       "data-bs-toggle='modal' data-bs-target='#editaracti' "+
                       "onclick='verEditaracti("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar acti' "+
                       "onclick='eliminaracti("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar acti' "+
                       "data-bs-toggle='modal' data-bs-target='#editaracti' "+
                       "onclick='verEditaracti("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],                     
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Actividadades Economicas  =======================

==============================================================================*/

function crearacti(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Actividad Economica");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearacti!="" && form.descipcionCrearacti!="" && form.descipcionCrearacti2!="" && form.descipcionCrearacti3!="" && form.txtAccion=="crearacti"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearacti);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearacti);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearacti2);
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearacti3);                    
    hoja.getRange("F"+(fila+1)).setValue("A");
  
    return "crearactiOk";
    
    
  }else if(form.nombreEditaracti!="" && form.descipcionEditaracti!=""&& form.descipcionEditaracti2!="" && form.descipcionEditaracti3!="" && form.txtAccion=="editaracti" && form.idEditaracti!=""){
     
    var datos = hoja.getRange("A2:F"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditaracti==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditaracti);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditaracti); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditaracti2);
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditaracti3);                       
        return "editaractiOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Actividad Economica  =======================

==============================================================================*/
function eliminaractiAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Actividad Economica");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:F"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("F"+(i+2)).setValue("D");
      
      return "eliminaractiOK";
    }      
  } 
}



/*=======================================================================

     =====================  Listar - Uno o Todos Nacionalidades  =======================

==============================================================================*/

function listarpais(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Nacionalidad_real");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:F"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][5]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4]                                     
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][5]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar pais' "+
                       "data-bs-toggle='modal' data-bs-target='#editarpais' "+
                       "onclick='verEditarpais("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar pais' "+
                       "onclick='eliminarpais("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar pais' "+
                       "data-bs-toggle='modal' data-bs-target='#editarpais' "+
                       "onclick='verEditarpais("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],   
        descripcion3:datos[i][4],                            
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Nacionalidad  =======================

==============================================================================*/

function crearpais(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Nacionalidad_real");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearpais!="" && form.descipcionCrearpais!="" && form.descipcionCrearpais2!=""&& form.descipcionCrearpais3!="" && form.txtAccion=="crearpais"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearpais);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearpais);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearpais2);
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearpais3);                        
    hoja.getRange("F"+(fila+1)).setValue("A");
  
    return "crearpaisOk";
    
    
  }else if(form.nombreEditarpais!="" && form.descipcionEditarpais!="" && form.descipcionEditarpais2!="" && form.descipcionEditarpais3!="" && form.txtAccion=="editarpais" && form.idEditarpais!=""){
     
    var datos = hoja.getRange("A2:F"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarpais==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarpais);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarpais); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarpais2); 
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditarpais3);                                
        return "editarpaisOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Nacionalidades  =======================

==============================================================================*/
function eliminarpaisAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Nacionalidad_real");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:F"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("F"+(i+2)).setValue("D");
      
      return "eliminarpaisOK";
    }      
  } 
}




/*=======================================================================

     =====================  Listar - Uno o Todos Jurisdicciones de colombia  =======================

==============================================================================*/

function listarjurisd(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Jurisdicciones");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:G"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][6]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],  
        descripcion4:datos[i][5]                                             
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][6]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar jurisd' "+
                       "data-bs-toggle='modal' data-bs-target='#editarjurisd' "+
                       "onclick='verEditarjurisd("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar jurisd' "+
                       "onclick='eliminarjurisd("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar jurisd' "+
                       "data-bs-toggle='modal' data-bs-target='#editarjurisd' "+
                       "onclick='verEditarjurisd("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],   
        descripcion3:datos[i][4],    
        descripcion4:datos[i][5],                                    
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}











/*=======================================================================

     =====================  Crear-Editar Nacionalidad  =======================

==============================================================================*/

function crearjurisd(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Jurisdicciones");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearjurisd!="" && form.descipcionCrearjurisd!="" && form.descipcionCrearjurisd2!=""&& form.descipcionCrearjurisd3!=""&& form.descipcionCrearjurisd4!="" && form.txtAccion=="crearjurisd"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearjurisd);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearjurisd);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearjurisd2);
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearjurisd3);  
    hoja.getRange("F"+(fila+1)).setValue(form.descipcionCrearjurisd4);                            
    hoja.getRange("G"+(fila+1)).setValue("A");
  
    return "crearjurisdOk";
    
    
  }else if(form.nombreEditarjurisd!="" && form.descipcionEditarjurisd!="" && form.descipcionEditarjurisd2!="" && form.descipcionEditarjurisd3!=""&& form.descipcionEditarjurisd4!="" && form.txtAccion=="editarjurisd" && form.idEditarjurisd!=""){
     
    var datos = hoja.getRange("A2:G"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarjurisd==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarjurisd);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarjurisd); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarjurisd2); 
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditarjurisd3);
        hoja.getRange("F"+(i+2)).setValue(form.descipcionEditarjurisd4);                                        
        return "editarjurisdOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Nacionalidades  =======================

==============================================================================*/
function eliminarjurisdAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Jurisdicciones");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:G"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("G"+(i+2)).setValue("D");
      
      return "eliminarjurisdOK";
    }      
  } 
}





/*=======================================================================

     =====================  Listar - Uno o Todos Requerimientos  =======================

==============================================================================*/

function listarrequ(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Requerimientos");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:I"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][8]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6],
        descripcion6:datos[i][7]              
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][8]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar requ' "+
                       "data-bs-toggle='modal' data-bs-target='#editarrequ' "+
                       "onclick='verEditarrequ("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar requ' "+
                       "onclick='eliminarrequ("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar requ' "+
                       "data-bs-toggle='modal' data-bs-target='#editarrequ' "+
                       "onclick='verEditarrequ("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }else if(rol=="GESTOR"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar requ' "+
                       "data-bs-toggle='modal' data-bs-target='#editarrequ' "+
                       "onclick='verEditarrequ("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6],   
        descripcion6:datos[i][7],                                                 
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Requerimientos  =======================

==============================================================================*/

function crearrequ(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Requerimientos");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearrequ!="" && form.descipcionCrearrequ!="" && form.descipcionCrearrequ2!=""&& form.descipcionCrearrequ3!=""&& form.descipcionCrearrequ4!=""&& form.descipcionCrearrequ5!="" && form.descipcionCrearrequ6!=""&& form.txtAccion=="crearrequ"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearrequ);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearrequ);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearrequ2); 
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearrequ3); 
    hoja.getRange("F"+(fila+1)).setValue(form.descipcionCrearrequ4); 
    hoja.getRange("G"+(fila+1)).setValue(form.descipcionCrearrequ5);    
    hoja.getRange("H"+(fila+1)).setValue(form.descipcionCrearrequ6);                                 
    hoja.getRange("I"+(fila+1)).setValue("A");
  
    return "crearrequOk";
    
    
  }else if(form.nombreEditarrequ!="" && form.descipcionEditarrequ!=""&& form.descipcionEditarrequ2!=""&& form.descipcionEditarrequ3!=""&& form.descipcionEditarrequ4!=""&& form.descipcionEditarrequ5!="" && form.descipcionEditarrequ6!=""&& form.txtAccion=="editarrequ" && form.idEditarrequ!=""){
     
    var datos = hoja.getRange("A2:I"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarrequ==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarrequ);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarrequ); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarrequ2);
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditarrequ3);
        hoja.getRange("F"+(i+2)).setValue(form.descipcionEditarrequ4);
        hoja.getRange("G"+(i+2)).setValue(form.descipcionEditarrequ5);  
        hoja.getRange("H"+(i+2)).setValue(form.descipcionEditarrequ6);                                                        
        return "editarrequOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Requerimientos  =======================

==============================================================================*/
function eliminarrequAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Requerimientos");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:I"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("I"+(i+2)).setValue("D");
      
      return "eliminarrequOK";
    }      
  } 
}


/*=======================================================================

     =====================  Listar - Tablero  =======================

==============================================================================*/

function listardash(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tablero");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:O"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][9]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6],
        descripcion6:datos[i][7],
        descripcion7:datos[i][8]                                                        
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][9]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar dash' "+
                       "data-bs-toggle='modal' data-bs-target='#editardash' "+
                       "onclick='verEditardash("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar dash' "+
                       "onclick='eliminardash("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar dash' "+
                       "data-bs-toggle='modal' data-bs-target='#editardash' "+
                       "onclick='verEditardash("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";             
      
      }else if(rol=="GESTOR"){   
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar dash' "+
                       "data-bs-toggle='modal' data-bs-target='#editardash' "+
                       "onclick='verEditardash("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";             
      }
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],    
        descripcion3:datos[i][4],    
        descripcion4:datos[i][5],    
        descripcion5:datos[i][6],    
        descripcion6:datos[i][7],    
        descripcion7:datos[i][8],    
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Nacionalidad  =======================

==============================================================================*/

function creardash(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tablero");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCreardash!="" && form.descipcionCreardash!="" && form.descipcionCreardash2!="" && form.descipcionCreardash3!="" && form.descipcionCreardash4!="" && form.descipcionCreardash5!="" && form.descipcionCreardash6!="" && form.descipcionCreardash7!="" && form.txtAccion=="creardash"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCreardash);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCreardash);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCreardash2); 
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCreardash3); 
    hoja.getRange("F"+(fila+1)).setValue(form.descipcionCreardash4); 
    hoja.getRange("G"+(fila+1)).setValue(form.descipcionCreardash5); 
    hoja.getRange("H"+(fila+1)).setValue(form.descipcionCreardash6); 
    hoja.getRange("I"+(fila+1)).setValue(form.descipcionCreardash7); 
    hoja.getRange("J"+(fila+1)).setValue("A");
  
    return "creardashOk";
    
    
  }else if(form.nombreEditardash!="" && form.descipcionEditardash!=""&& form.descipcionEditardash2!=""&& form.descipcionEditardash3!=""&& form.descipcionEditardash4!=""&& form.descipcionEditardash5!=""&& form.descipcionEditardash6!=""&& form.descipcionEditardash7!="" && form.txtAccion=="editardash" && form.idEditardash!=""){
     
    var datos = hoja.getRange("A2:O"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditardash==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditardash);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditardash); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditardash2);
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditardash3);
        hoja.getRange("F"+(i+2)).setValue(form.descipcionEditardash4);
        hoja.getRange("G"+(i+2)).setValue(form.descipcionEditardash5);
        hoja.getRange("H"+(i+2)).setValue(form.descipcionEditardash6);
        hoja.getRange("I"+(i+2)).setValue(form.descipcionEditardash7);                                     
        return "editardashOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Nacionalidades  =======================

==============================================================================*/
function eliminardashAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tablero");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:O"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("J"+(i+2)).setValue("D");
      
      return "eliminardashOK";
    }      
  } 
}


/*=======================================================================

     =====================  Listar - Uno o Todos Matriz  =======================

==============================================================================*/

function listarmatriz(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Matriz_final");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:F"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][5]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3], 
        descripcion3:datos[i][4]                                     
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][5]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar matriz' "+
                       "data-bs-toggle='modal' data-bs-target='#editarmatriz' "+
                       "onclick='verEditarmatriz("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar matriz' "+
                       "onclick='eliminarmatriz("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar matriz' "+
                       "data-bs-toggle='modal' data-bs-target='#editarmatriz' "+
                       "onclick='verEditarmatriz("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      }else if(rol=="GESTOR"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar matriz' "+
                       "data-bs-toggle='modal' data-bs-target='#editarmatriz' "+
                       "onclick='verEditarmatriz("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],                           
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Matriz  =======================

==============================================================================*/

function crearmatriz(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Matriz_final");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearmatriz!="" && form.descipcionCrearmatriz!="" && form.descipcionCrearmatriz2!="" &&form.descipcionCrearmatriz3!="" && form.txtAccion=="crearmatriz"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearmatriz);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearmatriz);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearmatriz2); 
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearmatriz3);                        
    hoja.getRange("F"+(fila+1)).setValue("A");
  
    return "crearmatrizOk";
    
    
  }else if(form.nombreEditarmatriz!="" && form.descipcionEditarmatriz!=""&& form.descipcionEditarmatriz2!="" &&form.descipcionEditarmatriz3!="" && form.txtAccion=="editarmatriz" && form.idEditarmatriz!=""){
     
    var datos = hoja.getRange("A2:F"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarmatriz==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarmatriz);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarmatriz); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarmatriz2);
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditarmatriz3);                                
        return "editarmatrizOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar MAtriz  =======================

==============================================================================*/
function eliminarmatrizAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Matriz_final");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:F"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("F"+(i+2)).setValue("D");
      
      return "eliminarmatrizOK";
    }      
  } 
}











/*=======================================================================

     =====================  Listar - Uno o Todos Profesiones  =======================

==============================================================================*/

function listarprofes(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Profesion");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:D"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][3]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2]                                     
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][3]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar profes' "+
                       "data-bs-toggle='modal' data-bs-target='#editarprofes' "+
                       "onclick='verEditarprofes("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar profes' "+
                       "onclick='eliminarprofes("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar profes' "+
                       "data-bs-toggle='modal' data-bs-target='#editarprofes' "+
                       "onclick='verEditarprofes("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      }else if(rol=="GESTOR"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar profes' "+
                       "data-bs-toggle='modal' data-bs-target='#editarprofes' "+
                       "onclick='verEditarprofes("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],                        
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Profesiones  =======================

==============================================================================*/

function crearprofes(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Profesion");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearprofes!="" && form.descipcionCrearprofes!="" &&  form.txtAccion=="crearprofes"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearprofes);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearprofes);                      
    hoja.getRange("D"+(fila+1)).setValue("A");
  
    return "crearprofesOk";
    
    
  }else if(form.nombreEditarprofes!="" && form.descipcionEditarprofes!=""&& form.txtAccion=="editarprofes" && form.idEditarprofes!=""){
     
    var datos = hoja.getRange("A2:D"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarprofes==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarprofes);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarprofes);                                
        return "editarprofesOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Profesiones  =======================

==============================================================================*/
function eliminarprofesAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Profesion");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:D"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("D"+(i+2)).setValue("D");
      
      return "eliminarprofesOK";
    }      
  } 
}











/*=======================================================================

     =====================  Listar - Uno o Todos Ocupacion  =======================

==============================================================================*/

function listarocupa(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Ocupacion_real");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:D"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][3]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2]                                     
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][3]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar ocupa' "+
                       "data-bs-toggle='modal' data-bs-target='#editarocupa' "+
                       "onclick='verEditarocupa("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar ocupa' "+
                       "onclick='eliminarocupa("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar ocupa' "+
                       "data-bs-toggle='modal' data-bs-target='#editarocupa' "+
                       "onclick='verEditarocupa("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      }else if(rol=="GESTOR"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar ocupa' "+
                       "data-bs-toggle='modal' data-bs-target='#editarocupa' "+
                       "onclick='verEditarocupa("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],                        
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Ocupacion  =======================

==============================================================================*/

function crearocupa(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Ocupacion_real");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearocupa!="" && form.descipcionCrearocupa!="" &&  form.txtAccion=="crearocupa"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearocupa);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearocupa);                      
    hoja.getRange("D"+(fila+1)).setValue("A");
  
    return "crearocupaOk";
    
    
  }else if(form.nombreEditarocupa!="" && form.descipcionEditarocupa!=""&& form.txtAccion=="editarocupa" && form.idEditarocupa!=""){
     
    var datos = hoja.getRange("A2:D"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarocupa==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarocupa);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarocupa);                                
        return "editarocupaOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Ocupacion  =======================

==============================================================================*/
function eliminarocupaAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Ocupacion_real");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:D"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("D"+(i+2)).setValue("D");
      
      return "eliminarocupaOK";
    }      
  } 
}

















/*=======================================================================

     =====================  Listar - Uno o Todos person  =======================

==============================================================================*/

function listarperson(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Usuarios");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:H"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][7]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][3],
        descripcion2:datos[i][6]                             
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][7]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar person' "+
                       "data-bs-toggle='modal' data-bs-target='#editarperson' "+
                       "onclick='verEditarperson("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar person' "+
                       "onclick='eliminarperson("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar person' "+
                       "data-bs-toggle='modal' data-bs-target='#editarperson' "+
                       "onclick='verEditarperson("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][3],
        descripcion2:datos[i][6],                   
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar person  =======================

==============================================================================*/

function crearperson(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Usuarios");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearperson!="" && form.descipcionCrearperson!="" && form.descipcionCrearperson2!="" && form.txtAccion=="crearperson"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearperson);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearperson);
    hoja.getRange("G"+(fila+1)).setValue(form.descipcionCrearperson2);                    
    hoja.getRange("H"+(fila+1)).setValue("A");
  
    return "crearpersonOk";
    
    
  }else if(form.nombreEditarperson!="" && form.descipcionEditarperson!=""&& form.descipcionEditarperson2!="" && form.txtAccion=="editarperson" && form.idEditarperson!=""){
     
    var datos = hoja.getRange("A2:H"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarperson==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarperson);
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarperson); 
        hoja.getRange("G"+(i+2)).setValue(form.descipcionEditarperson2);                        
        return "editarpersonOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Person  =======================

==============================================================================*/
function eliminarpersonAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Usuarios");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:E"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("H"+(i+2)).setValue("D");
      
      return "eliminarpersonOK";
    }      
  } 
}



/*=======================================================================

     =====================  Listar - Uno o Todos Transacciones  =======================

==============================================================================*/

function listartrans(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tipo_transaccion");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:E"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][4]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3]                             
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][4]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar trans' "+
                       "data-bs-toggle='modal' data-bs-target='#editartrans' "+
                       "onclick='verEditartrans("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar trans' "+
                       "onclick='eliminartrans("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar trans' "+
                       "data-bs-toggle='modal' data-bs-target='#editartrans' "+
                       "onclick='verEditartrans("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],                   
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Transaccionalidad  =======================

==============================================================================*/

function creartrans(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tipo_transaccion");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCreartrans!="" && form.descipcionCreartrans!="" && form.descipcionCreartrans2!="" && form.txtAccion=="creartrans"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCreartrans);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCreartrans);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCreartrans2);                    
    hoja.getRange("E"+(fila+1)).setValue("A");
  
    return "creartransOk";
    
    
  }else if(form.nombreEditartrans!="" && form.descipcionEditartrans!=""&& form.descipcionEditartrans2!="" && form.txtAccion=="editartrans" && form.idEditartrans!=""){
     
    var datos = hoja.getRange("A2:E"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditartrans==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditartrans);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditartrans); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditartrans2);                        
        return "editartransOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Transacciones  =======================

==============================================================================*/
function eliminartransAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tipo_transaccion");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:E"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("E"+(i+2)).setValue("D");
      
      return "eliminartransOK";
    }      
  } 
}


/*=======================================================================

     =====================  Listar - Uno o Todos Canales  =======================

==============================================================================*/

function listarcanal(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Canal");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:E"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][4]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3]                             
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][4]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar canal' "+
                       "data-bs-toggle='modal' data-bs-target='#editarcanal' "+
                       "onclick='verEditarcanal("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar canal' "+
                       "onclick='eliminarcanal("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar canal' "+
                       "data-bs-toggle='modal' data-bs-target='#editarcanal' "+
                       "onclick='verEditarcanal("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],                   
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Nacionalidad  =======================

==============================================================================*/

function crearcanal(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Canal");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearcanal!="" && form.descipcionCrearcanal!="" && form.descipcionCrearcanal2!="" && form.txtAccion=="crearcanal"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearcanal);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearcanal);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearcanal2);                    
    hoja.getRange("E"+(fila+1)).setValue("A");
  
    return "crearcanalOk";
    
    
  }else if(form.nombreEditarcanal!="" && form.descipcionEditarcanal!=""&& form.descipcionEditarcanal2!="" && form.txtAccion=="editarcanal" && form.idEditarcanal!=""){
     
    var datos = hoja.getRange("A2:E"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarcanal==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarcanal);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarcanal); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarcanal2);                        
        return "editarcanalOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Canal  =======================

==============================================================================*/
function eliminarcanalAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Canal");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:E"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("E"+(i+2)).setValue("D");
      
      return "eliminarcanalOK";
    }      
  } 
}


/*=======================================================================

     =====================  Listar - Uno o Todos Productos  =======================

==============================================================================*/

function listarprod(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Productos");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:I"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][8]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6],
        descripcion6:datos[i][7]                                                             
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][8]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar prod' "+
                       "data-bs-toggle='modal' data-bs-target='#editarprod' "+
                       "onclick='verEditarprod("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar prod' "+
                       "onclick='eliminarprod("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar prod' "+
                       "data-bs-toggle='modal' data-bs-target='#editarprod' "+
                       "onclick='verEditarprod("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6],
        descripcion6:datos[i][7],                                                   
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Productos  =======================

==============================================================================*/

function crearprod(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Productos");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearprod!="" && form.descipcionCrearprod!="" && form.descipcionCrearprod2!="" && form.descipcionCrearprod3!="" &&form.descipcionCrearprod4!="" &&form.descipcionCrearprod5!="" &&form.descipcionCrearprod6!="" && form.txtAccion=="crearprod"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearprod);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearprod);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearprod2);
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearprod3);
    hoja.getRange("F"+(fila+1)).setValue(form.descipcionCrearprod4);
    hoja.getRange("G"+(fila+1)).setValue(form.descipcionCrearprod5);
    hoja.getRange("H"+(fila+1)).setValue(form.descipcionCrearprod6);                                    
    hoja.getRange("I"+(fila+1)).setValue("A");
  
    return "crearprodOk";
    
    
  }else if(form.nombreEditarprod!="" && form.descipcionEditarprod!=""&& form.descipcionEditarprod2!=""&& form.descipcionEditarprod3!=""&& form.descipcionEditarprod4!=""&& form.descipcionEditarprod5!=""&& form.descipcionEditarprod6!="" && form.txtAccion=="editarprod" && form.idEditarprod!=""){
     
    var datos = hoja.getRange("A2:I"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarprod==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarprod);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarprod); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarprod2);
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditarprod3);
        hoja.getRange("F"+(i+2)).setValue(form.descipcionEditarprod4);
        hoja.getRange("G"+(i+2)).setValue(form.descipcionEditarprod5);
        hoja.getRange("H"+(i+2)).setValue(form.descipcionEditarprod6);                                                        
        return "editarprodOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Productos  =======================

==============================================================================*/
function eliminarprodAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Productos");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:I"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("I"+(i+2)).setValue("D");
      
      return "eliminarprodOK";
    }      
  } 
}



/*=======================================================================

     =====================  Listar - Uno o Todos Listas   =======================

==============================================================================*/

function listarlistas(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Listas");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:E"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][4]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3]                             
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][4]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar listas' "+
                       "data-bs-toggle='modal' data-bs-target='#editarlistas' "+
                       "onclick='verEditarlistas("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar listas' "+
                       "onclick='eliminarlistas("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar listas' "+
                       "data-bs-toggle='modal' data-bs-target='#editarlistas' "+
                       "onclick='verEditarlistas("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],                   
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Listas  =======================

==============================================================================*/

function crearlistas(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Listas");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearlistas!="" && form.descipcionCrearlistas!="" && form.descipcionCrearlistas2!="" && form.txtAccion=="crearlistas"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearlistas);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearlistas);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearlistas2);                    
    hoja.getRange("E"+(fila+1)).setValue("A");
  
    return "crearlistasOk";
    
    
  }else if(form.nombreEditarlistas!="" && form.descipcionEditarlistas!=""&& form.descipcionEditarlistas2!="" && form.txtAccion=="editarlistas" && form.idEditarlistas!=""){
     
    var datos = hoja.getRange("A2:E"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarlistas==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarlistas);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarlistas); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarlistas2);                        
        return "editarlistasOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Listas  =======================

==============================================================================*/
function eliminarlistasAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Listas");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:E"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("E"+(i+2)).setValue("D");
      
      return "eliminarlistasOK";
    }      
  } 
}




/*=======================================================================

     =====================  Listar - Uno o Todos Alertas  =======================

==============================================================================*/

function listaralerta(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Alertas");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:AA"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][26]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6],
        descripcion6:datos[i][7],
        descripcion7:datos[i][8],
        descripcion8:datos[i][9],
        descripcion9:datos[i][10],
        descripcion10:datos[i][11],
        descripcion11:datos[i][12],
        descripcion12:datos[i][13],
        descripcion13:datos[i][14],
        descripcion14:datos[i][15],
        descripcion15:datos[i][16],
        descripcion16:datos[i][17],
        descripcion17:datos[i][18],
        descripcion18:datos[i][19],
        descripcion19:datos[i][20],
        descripcion20:datos[i][21],
        descripcion21:datos[i][22],
        descripcion22:datos[i][23],
        descripcion23:datos[i][24],
        descripcion24:datos[i][25]       
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][26]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar alerta' "+
                       "data-bs-toggle='modal' data-bs-target='#editaralerta' "+
                       "onclick='verEditaralerta("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar alerta' "+
                       "onclick='eliminaralerta("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar alerta' "+
                       "data-bs-toggle='modal' data-bs-target='#editaralerta' "+
                       "onclick='verEditaralerta("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }else if(rol=="GESTOR"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar alerta' "+
                       "data-bs-toggle='modal' data-bs-target='#editaralerta' "+
                       "onclick='verEditaralerta("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],                
        descripcion5:datos[i][6],
        descripcion6:datos[i][7],  
        descripcion7:datos[i][8],
        descripcion8:datos[i][9],
        descripcion9:datos[i][10],
        descripcion10:datos[i][11],
        descripcion11:datos[i][12],
        descripcion12:datos[i][13],
        descripcion13:datos[i][14],
        descripcion14:datos[i][15],
        descripcion15:datos[i][16],
        descripcion16:datos[i][17],
        descripcion17:datos[i][18],
        descripcion18:datos[i][19],
        descripcion19:datos[i][20],
        descripcion20:datos[i][21],
        descripcion21:datos[i][22],
        descripcion22:datos[i][23],
        descripcion23:datos[i][24],    
        descripcion24:datos[i][25],                           
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Alerta  =======================

==============================================================================*/

function crearalerta(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Alertas");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearalerta!="" && form.descipcionCrearalerta!="" && form.descipcionCrearalerta2!="" && form.descipcionCrearalerta3!="" && form.descipcionCrearalerta4!="" && form.descipcionCrearalerta5!="" && form.descipcionCrearalerta6!="" && form.descipcionCrearalerta7!="" && form.descipcionCrearalerta8!=""&& form.descipcionCrearalerta9!=""&& form.descipcionCrearalerta10!=""&& form.descipcionCrearalerta11!=""&& form.descipcionCrearalerta12!=""&& form.descipcionCrearalerta13!=""&& form.descipcionCrearalerta14!=""&& form.descipcionCrearalerta15!=""&& form.descipcionCrearalerta16!=""&& form.descipcionCrearalerta17!=""&& form.descipcionCrearalerta18!=""&& form.descipcionCrearalerta19!=""&& form.descipcionCrearalerta20!=""&& form.descipcionCrearalerta21!=""&& form.descipcionCrearalerta22!=""&& form.descipcionCrearalerta23!=""&& form.descipcionCrearalerta24!=""&& form.txtAccion=="crearalerta"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearalerta);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearalerta);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearalerta2);
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearalerta3);
    hoja.getRange("F"+(fila+1)).setValue(form.descipcionCrearalerta4); 
    hoja.getRange("G"+(fila+1)).setValue(form.descipcionCrearalerta5); 
    hoja.getRange("H"+(fila+1)).setValue(form.descipcionCrearalerta6);   
    hoja.getRange("I"+(fila+1)).setValue(form.descipcionCrearalerta7);
    hoja.getRange("J"+(fila+1)).setValue(form.descipcionCrearalerta8);
    hoja.getRange("K"+(fila+1)).setValue(form.descipcionCrearalerta9);
    hoja.getRange("L"+(fila+1)).setValue(form.descipcionCrearalerta10);
    hoja.getRange("M"+(fila+1)).setValue(form.descipcionCrearalerta11);
    hoja.getRange("N"+(fila+1)).setValue(form.descipcionCrearalerta12);
    hoja.getRange("O"+(fila+1)).setValue(form.descipcionCrearalerta13);
    hoja.getRange("P"+(fila+1)).setValue(form.descipcionCrearalerta14);
    hoja.getRange("Q"+(fila+1)).setValue(form.descipcionCrearalerta15);
    hoja.getRange("R"+(fila+1)).setValue(form.descipcionCrearalerta16);
    hoja.getRange("S"+(fila+1)).setValue(form.descipcionCrearalerta17);
    hoja.getRange("T"+(fila+1)).setValue(form.descipcionCrearalerta18);
    hoja.getRange("U"+(fila+1)).setValue(form.descipcionCrearalerta19);
    hoja.getRange("V"+(fila+1)).setValue(form.descipcionCrearalerta20);
    hoja.getRange("W"+(fila+1)).setValue(form.descipcionCrearalerta21);
    hoja.getRange("X"+(fila+1)).setValue(form.descipcionCrearalerta22);
    hoja.getRange("Y"+(fila+1)).setValue(form.descipcionCrearalerta23);
    hoja.getRange("Z"+(fila+1)).setValue(form.descipcionCrearalerta24);
    hoja.getRange("AA"+(fila+1)).setValue("A");
  
    return "crearalertaOk";
    
    
  }else if(form.nombreEditaralerta!="" && form.descipcionEditaralerta!=""&& form.descipcionEditaralerta2!="" && form.descipcionEditaralerta3!="" && form.descipcionEditaralerta4!="" && form.descipcionEditaralerta5!=""&& form.descipcionEditaralerta6!="" && form.descipcionEditaralerta7!=""&& form.descipcionEditaralerta8!=""&& form.descipcionEditaralerta9!=""&& form.descipcionEditaralerta10!=""&& form.descipcionEditaralerta11!=""&& form.descipcionEditaralerta12!=""&& form.descipcionEditaralerta13!=""&& form.descipcionEditaralerta14!=""&& form.descipcionEditaralerta15!=""&& form.descipcionEditaralerta16!=""&& form.descipcionEditaralerta17!=""&& form.descipcionEditaralerta18!=""&& form.descipcionEditaralerta19!=""&& form.descipcionEditaralerta20!=""&& form.descipcionEditaralerta21!=""&& form.descipcionEditaralerta22!=""&& form.descipcionEditaralerta23!=""&& form.descipcionEditaralerta24!="" && form.txtAccion=="editaralerta" && form.idEditaralerta!=""){
     
    var datos = hoja.getRange("A2:AA"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditaralerta==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditaralerta);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditaralerta); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditaralerta2);
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditaralerta3);
        hoja.getRange("F"+(i+2)).setValue(form.descipcionEditaralerta4);
        hoja.getRange("G"+(i+2)).setValue(form.descipcionEditaralerta5);
        hoja.getRange("H"+(i+2)).setValue(form.descipcionEditaralerta6);
        hoja.getRange("I"+(i+2)).setValue(form.descipcionEditaralerta7);
        hoja.getRange("J"+(i+2)).setValue(form.descipcionEditaralerta8);  
        hoja.getRange("K"+(i+2)).setValue(form.descipcionEditaralerta9);  
        hoja.getRange("L"+(i+2)).setValue(form.descipcionEditaralerta10);  
        hoja.getRange("M"+(i+2)).setValue(form.descipcionEditaralerta11);  
        hoja.getRange("N"+(i+2)).setValue(form.descipcionEditaralerta12);  
        hoja.getRange("O"+(i+2)).setValue(form.descipcionEditaralerta13);  
        hoja.getRange("P"+(i+2)).setValue(form.descipcionEditaralerta14);  
        hoja.getRange("Q"+(i+2)).setValue(form.descipcionEditaralerta15);  
        hoja.getRange("R"+(i+2)).setValue(form.descipcionEditaralerta16);  
        hoja.getRange("S"+(i+2)).setValue(form.descipcionEditaralerta17);  
        hoja.getRange("T"+(i+2)).setValue(form.descipcionEditaralerta18);  
        hoja.getRange("U"+(i+2)).setValue(form.descipcionEditaralerta19);  
        hoja.getRange("V"+(i+2)).setValue(form.descipcionEditaralerta20);  
        hoja.getRange("W"+(i+2)).setValue(form.descipcionEditaralerta21);  
        hoja.getRange("X"+(i+2)).setValue(form.descipcionEditaralerta22);  
        hoja.getRange("Y"+(i+2)).setValue(form.descipcionEditaralerta23);  
        hoja.getRange("Z"+(i+2)).setValue(form.descipcionEditaralerta24);  
                                                           
        return "editaralertaOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}

/*=======================================================================

     =====================  Eliminar Alertas  =======================

==============================================================================*/
function eliminaralertaAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Alertas");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:AA"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("AA"+(i+2)).setValue("D");
      
      return "eliminaralertaOK";
    }      
  } 

}






/*=======================================================================

     =====================  Listar - Uno o Todos Riesgos  =======================

==============================================================================*/

function listarriesgos(id,rol) {
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Riesgos");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:M"+fila).getValues();
  var resultado = [];
  
  if(id!=null){
     
    for(var i = 0; i<datos.length; i++){
    
    if(datos[i][12]=="A" && datos[i][0]== id){
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],
        descripcion5:datos[i][6],
        descripcion6:datos[i][7],
        descripcion7:datos[i][8],
        descripcion8:datos[i][9],
        descripcion9:datos[i][10],
        descripcion10:datos[i][11]                                                     
      })
      
      break;
    }  
  
  }
    
  }else{
  
  
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][12]=="A"){
        
      if(rol=="ADMIN"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar riesgos' "+
                       "data-bs-toggle='modal' data-bs-target='#editarriesgos' "+
                       "onclick='verEditarriesgos("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar riesgos' "+
                       "onclick='eliminarriesgos("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      
      }else if(rol=="SPECIAL"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar riesgos' "+
                       "data-bs-toggle='modal' data-bs-target='#editarriesgos' "+
                       "onclick='verEditarriesgos("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      
      }else if(rol=="GESTOR"){
         
         var botones = "<a class='btn btn-primary posicionBoton' title='Editar riesgos' "+
                       "data-bs-toggle='modal' data-bs-target='#editarriesgos' "+
                       "onclick='verEditarriesgos("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
      }
      
      
      resultado.push({
         
        id:datos[i][0],
        nombre:datos[i][1],
        descripcion:datos[i][2],
        descripcion2:datos[i][3],
        descripcion3:datos[i][4],
        descripcion4:datos[i][5],                
        descripcion5:datos[i][6],
        descripcion6:datos[i][7],  
        descripcion7:datos[i][8],
        descripcion8:datos[i][9],  
        descripcion9:datos[i][10],  
        descripcion10:datos[i][11],                                                        
        acciones:botones
              
      })
    }  
  
  }  
  
 }
   return resultado;
  
}


/*=======================================================================

     =====================  Crear-Editar Riesgos  =======================

==============================================================================*/

function crearriesgos(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Riesgos");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.nombreCrearriesgos!="" && form.descipcionCrearriesgos!="" && form.descipcionCrearriesgos2!="" && form.descipcionCrearriesgos3!="" && form.descipcionCrearriesgos4!="" && form.descipcionCrearriesgos5!="" && form.descipcionCrearriesgos6!="" && form.descipcionCrearriesgos7!=""&& form.descipcionCrearriesgos8!=""&& form.descipcionCrearriesgos9!=""&& form.descipcionCrearriesgos10!="" && form.txtAccion=="crearriesgos"){
     
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreCrearriesgos);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionCrearriesgos);
    hoja.getRange("D"+(fila+1)).setValue(form.descipcionCrearriesgos2);
    hoja.getRange("E"+(fila+1)).setValue(form.descipcionCrearriesgos3);
    hoja.getRange("F"+(fila+1)).setValue(form.descipcionCrearriesgos4); 
    hoja.getRange("G"+(fila+1)).setValue(form.descipcionCrearriesgos5); 
    hoja.getRange("H"+(fila+1)).setValue(form.descipcionCrearriesgos6);   
    hoja.getRange("I"+(fila+1)).setValue(form.descipcionCrearriesgos7);  
    hoja.getRange("J"+(fila+1)).setValue(form.descipcionCrearriesgos8);  
    hoja.getRange("K"+(fila+1)).setValue(form.descipcionCrearriesgos9);  
    hoja.getRange("L"+(fila+1)).setValue(form.descipcionCrearriesgos10);                                 
    hoja.getRange("M"+(fila+1)).setValue("A");
  
    return "crearriesgosOk";
    
    
  }else if(form.nombreEditarriesgos!="" && form.descipcionEditarriesgos!=""&& form.descipcionEditarriesgos2!="" && form.descipcionEditarriesgos3!="" && form.descipcionEditarriesgos4!="" && form.descipcionEditarriesgos5!=""&& form.descipcionEditarriesgos6!="" && form.descipcionEditarriesgos7!="" && form.descipcionEditarriesgos8!=""&& form.descipcionEditarriesgos9!=""&& form.descipcionEditarriesgos10!=""&& form.txtAccion=="editarriesgos" && form.idEditarriesgos!=""){
     
    var datos = hoja.getRange("A2:M"+fila).getValues();
    
    for(var i=0; i<datos.length; i++){
      
      if(form.idEditarriesgos==datos[i][0]){
      
        hoja.getRange("B"+(i+2)).setValue(form.nombreEditarriesgos);
        hoja.getRange("C"+(i+2)).setValue(form.descipcionEditarriesgos); 
        hoja.getRange("D"+(i+2)).setValue(form.descipcionEditarriesgos2);
        hoja.getRange("E"+(i+2)).setValue(form.descipcionEditarriesgos3);
        hoja.getRange("F"+(i+2)).setValue(form.descipcionEditarriesgos4);
        hoja.getRange("G"+(i+2)).setValue(form.descipcionEditarriesgos5);
        hoja.getRange("H"+(i+2)).setValue(form.descipcionEditarriesgos6);
        hoja.getRange("I"+(i+2)).setValue(form.descipcionEditarriesgos7);        
        hoja.getRange("J"+(i+2)).setValue(form.descipcionEditarriesgos8);
        hoja.getRange("K"+(i+2)).setValue(form.descipcionEditarriesgos9);    
        hoja.getRange("L"+(i+2)).setValue(form.descipcionEditarriesgos10);      
        return "editarriesgosOk";
      }
    
    }
     
    
  }else{
  
    return "faltaInformacion";
  }

}


/*=======================================================================

     =====================  Eliminar Riesgos  =======================

==============================================================================*/
function eliminarriesgosAdmin(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Riesgos");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:M"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
    
    if(id==datos[i][0]){
      
      hoja.getRange("M"+(i+2)).setValue("D");
      
      return "eliminarriesgosOK";
    }      
  } 

}




/*=======================================================================

     =====================  Listar Usuario Todos-Uno Etiqueta  =======================

==============================================================================*/
function listarUsuarios(idUsuario, rol){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Usuarios");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:H"+fila).getValues();
  var resultado = [];

  //  Datos de Areas
  var hojaAreas = libro.getSheetByName("Controles");
  var filaAreas = hojaAreas.getLastRow();
  var datosAreas = hojaAreas.getRange("A2:G"+fila).getValues();
  
  if(idUsuario!=null){
  
    for(var i = 0; i<datos.length; i++){
      
      if(idUsuario==datos[i][0]){
        
        var Areas = [];        
          
//        Area Actual persona
        for(var h = 0; h<datosAreas.length; h++){ 
          
          if(datos[i][6]==datosAreas[h][0]){
          
             Areas.push("<option value='"+datosAreas[h][0]+"'>"+datosAreas[h][1]+"</option>");  
          } 
       }
        
//        Listado de Areas
        for(var j = 0; j<datosAreas.length; j++){ 
          
           Areas.push("<option value='"+datosAreas[j][0]+"'>"+datosAreas[j][1]+"</option>");     
      
      }
        
        
//        Perfiles de usuario
        var perfil = "<option value='"+datos[i][6]+"'>"+datos[i][6]+"</option>"+
                     "<option value='ADMIN'>ADMIN</option>"+
                     "<option value='SPECIAL'>SPECIAL</option>"+
                     "<option value='EVALUATOR'>EVALUATOR</option>"+                     
                     "<option value='STANDAR'>STANDAR</option>";
        
 
          resultado.push({
           
            id:datos[i][0],
            nombre:datos[i][1],
            correo:datos[i][3],
            telefono:datos[i][4],
            area:Areas,
            perfil:perfil
            
          })
        
          break;

      }    
    }
    
    return resultado;
  
  }else{
  
//  Recorrer areas para saber el nombre
  for(var i = 0; i<datos.length; i++){
    
    if(datos[i][6]=="A"){
      
      for(var j = 0; j<datosAreas.length; j++){
         
        if(datosAreas[j][0]==datos[i][6]){
        
          var Area = datosAreas[j][1];
          
          break;          
        }
      
      }
    
    
    if(datos[i][2]!=""){
      var img  = "<div class='text-center'>"+
                    "<img src='https://raw.githubusercontent.com/DAVIDUBAQUEGARCIA10/PROYECTO_ARIS/refs/heads/main/BANNER/ChatGPT%20Image%201%20jul%202025%2C%2000_08_18.png' class='img-thumbnail img-fluid imagenes'> "+
                 "</div>";
    }else{
    
      var img  = "<div class='text-center'>"+
                    "<img src='https://raw.githubusercontent.com/DAVIDUBAQUEGARCIA10/PROYECTO_ARIS/refs/heads/main/BANNER/ChatGPT%20Image%201%20jul%202025%2C%2000_08_18.png' class='img-thumbnail img-fluid imagenes'> "+
                 "</div>";  
    }
    
      
      if(rol=="ADMIN"){
      
        var botones = "<a class='btn btn-primary posicionBoton' title='Editar Usuario' "+
                       "data-bs-toggle='modal' data-bs-target='#editarUsuario' "+
                       "onclick='verEditarUsuario("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>"+
                       "<a class='btn btn-danger' title='Eliminar Usuario' "+
                       "onclick='eliminarUsuario("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-trash-can'></i> </a>";
      }else if(rol=="SPECIAL"){
      
        var botones = "<a class='btn btn-primary posicionBoton' title='Editar Usuario' "+
                       "data-bs-toggle='modal' data-bs-target='#editarUsuario' "+
                       "onclick='verEditarUsuario("+datos[i][0]+");'> "+
                       "<i class='fa-solid fa-pencil'></i>  </a>";
                      
      }
     
    resultado.push({
       
      id:datos[i][0],
      nombre:datos[i][1],
      foto:img,
      correo:datos[i][3],
      telefono:datos[i][4],
      area:Area,
      perfil:datos[i][6],
      acciones:botones
    });    
    
    }
    
  }
  
  return resultado;
    
  }
  
}


/*============================================================

   ==================== Listar Areas Select  =================
   
==============================================================*/
function listarAreasSelect(){
    
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Controles");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:G"+fila).getValues();
  var arrayAreas = [];
    
   arrayAreas.push("<option value='0'>Seleccionar</option>");
  
  for(var i=0; i<datos.length; i++){
    
    if(datos[i][6]=="A"){
      
      arrayAreas.push("<option value='"+datos[i][0]+"'>"+datos[i][1]+"</option>");
    }
  }  
  
//  Roles de Usuario
  
  var hojaUsuario = libro.getSheetByName("Usuarios");
  var filaUsuario = hojaUsuario.getLastRow();
  var datosUsuario = hojaUsuario.getRange("A2:H"+filaUsuario).getValues();
  var usuarioActivo = Session.getActiveUser().getEmail();
  var arrayUsuario = [];
  
  for(var i = 0; i<datosUsuario.length; i++) {
    if(usuarioActivo == datosUsuario[i][3]){
      
      arrayUsuario.push({
        nombre:datosUsuario[i][1],
        foto:datosUsuario[i][2],
        rol:datosUsuario[i][6]      
      });      
    }  
  }
       
  var resultado = [];
  
  resultado.push({
    arrayAreas:arrayAreas,
    arrayUsuario:arrayUsuario
  });
  
  return resultado;
}


/*============================================================

   ==================== Crear - Editar Usuario =================
   
==============================================================*/
function crearUsuario(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Usuarios");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  
  if(form.txtAccion=="crearUsuarios" && form.nombreUsuario !="" && form.correoUsuario != "" && 
     form.telefonoUsuario!="" && form.areaUsuario !="0" && form.perfilUsuario !="0"){
    
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.nombreUsuario);
    hoja.getRange("C"+(fila+1)).setValue(form.extensionFotoUsuario);
    
    var extension = hoja.getRange("C"+(fila+1)).getValue();
    
    if(extension == "jpg" || extension =="JPG" || extension =="png" || 
       extension =="jpeg" || extension =="PNG" || extension =="jfif"){
    
      var carpeta = DriveApp.getFolderById("1cskX4ZHqj5Plo_6ywmYKNf1_98otDydp");
      var archivo = carpeta.createFile(form.fotoUsuario);
      var idfoto = archivo.getId();
      hoja.getRange("C"+(fila+1)).setValue(idfoto);
    
    }   
    
    hoja.getRange("D"+(fila+1)).setValue(form.correoUsuario);
    hoja.getRange("E"+(fila+1)).setValue(form.telefonoUsuario);
    hoja.getRange("F"+(fila+1)).setValue(form.areaUsuario);
    hoja.getRange("G"+(fila+1)).setValue(form.perfilUsuario);
    hoja.getRange("H"+(fila+1)).setValue("A");
    
    return "crearUsuarioOK"
  }else if(form.txtAccion=="editarUsuario" && form.idEditarUsuarios!="" && form.nombreUsuarioEditar !="" &&
           form.correoUsuarioEditar !="" && form.telefonoUsuarioEditar !=""){
      
    var datos = hoja.getRange("A2:H"+fila).getValues();
    
    for(var i = 0; i<datos.length; i++){
      
      if(datos[i][0]==form.idEditarUsuarios){
         
        hoja.getRange("B"+(i+2)).setValue(form.nombreUsuarioEditar);
        
//        Editar Imagen
        var fotoAnterior = hoja.getRange("C"+(i+2)).getValue();
        var pegarExtension = hoja.getRange("C"+(i+2)).setValue(form.extensionFotoUsuarioEditar);
        var extension = hoja.getRange("C"+(i+2)).getValue();
        
        if(extension == "jpg" || extension =="JPG" || extension =="png" || 
           extension =="jpeg" || extension =="PNG" || extension =="jfif"){
        
          var carpeta = DriveApp.getFolderById("1cskX4ZHqj5Plo_6ywmYKNf1_98otDydp");
          var archivo = carpeta.createFile(form.fotoUsuarioEditar);
          var idfoto = archivo.getId();
          hoja.getRange("C"+(i+2)).setValue(idfoto);
          
          if(fotoAnterior!=""){
            
            var archivoBorrar = DriveApp.getFileById(fotoAnterior);
            
            archivoBorrar.setTrashed(true);
            
          }         
        
        }else{
        
          hoja.getRange("C"+(i+2)).setValue(fotoAnterior);
        
        }
        
      
        hoja.getRange("D"+(i+2)).setValue(form.correoUsuarioEditar);
        hoja.getRange("E"+(i+2)).setValue(form.telefonoUsuarioEditar);
        hoja.getRange("F"+(i+2)).setValue(form.areaUsuarioEditar);
        hoja.getRange("G"+(i+2)).setValue(form.perfilUsuarioEditar);
      
//         break;
        
        return "editarUsuarioOK";
      }     
      
    }   
    
  }else{
    
    return "faltaInformacion";
  
  }
}


/*============================================================

   ==================== Eliinar Usuario =================
   
==============================================================*/
function eliminarUsuarioAdmin(id){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Usuarios");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:H"+fila).getValues();
  
  for(var i= 0; i<datos.length; i++){
    
    if(id==datos[i][0]){
       
      hoja.getRange("H"+(i+2)).setValue("D");
      
      return "eliminarUsuarioOK";
    
    }
  
  }

}


/*============================================================

   ==================== Listar Tickets =================
   
==============================================================*/
function listarTickets(accion,rol){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tickets");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:M"+fila).getValues();
  var resultado = [];
  var usuarioActivo = Session.getActiveUser().getEmail();
  
//  Llamado a las datos de usuario
  var hojaUsuarios = libro.getSheetByName("Usuarios");
  var filaUsuarios = hojaUsuarios.getLastRow();
  var datosUsuarios = hojaUsuarios.getRange("A2:D"+filaUsuarios).getValues();
  
//  Llamado a los datos de area
  
  var hojaAreas = libro.getSheetByName("Controles");
  var filaAreas = hojaAreas.getLastRow();
  var datosAreas = hojaAreas.getRange("A2:B"+filaAreas).getValues();
  
  
//  Llamado a los datos de Etiquetas
  
  var hojaEtiquetas = libro.getSheetByName("Riesgos");
  var filaEtiquetas = hojaEtiquetas.getLastRow();
  var datosEtiquetas = hojaEtiquetas.getRange("A2:B"+filaEtiquetas).getValues();  
  
  
  for(var i=0; i<datos.length; i++){
    
//    Color de Boton Segun Estado
    if(datos[i][4]=="Nuevo"){
       
      var estadoB = "<center><a class='btn btn-warning btn-sm tamano' title='Estado'>"+datos[i][4]+"</a></center>"
    
    }else if(datos[i][4]=="Abierto"){
       
      var estadoB = "<center><a class='btn btn-danger btn-sm tamano' title='Estado'>"+datos[i][4]+"</a></center>"
    
    }else if(datos[i][4]=="En Curso"){
       
      var estadoB = "<center><a class='btn btn-info btn-sm tamano' title='Estado'>"+datos[i][4]+"</a></center>"
    
    }else if(datos[i][4]=="Cerrado"){
       
      var estadoB = "<center><a class='btn btn-success btn-sm tamano' title='Estado'>"+datos[i][4]+"</a></center>"
    
    }else if(datos[i][4]=="Borrado"){
       
      var estadoB = "<center><a class='btn btn-secondary btn-sm tamano' title='Estado'>"+datos[i][4]+"</a></center>"
    
    }
    
//    Recorremos datos de usuario para identificar el agente 
    
    if(datos[i][6]!=""){
       
        for(var j=0; j<datosUsuarios.length; j++){
        
        if(datosUsuarios[j][3]==datos[i][6]){
           
          var nombreAgente = datosUsuarios[j][1];        
          break;        
        }
         
      }
    
    }else{
    
      var nombreAgente = "--"; 
    }
    
    
    //    Recorremos datos de usuario para identificar el Cliente 
    
    for(var h=0; h<datosUsuarios.length; h++){
      
      if(datosUsuarios[h][3]==datos[i][7]){
         
        var nombreCliente = datosUsuarios[h][1];        
        break;        
      }
       
    }
    
//    Recorremos datos de Areas para identificar el nombre
    
    if(datos[i][8]!=""){
      
      for(var a =0; a<datosAreas.length; a++ ){       
          
        if(datos[i][8]==datosAreas[a][0]){
            
          var nombreArea = datosAreas[a][1];
          break;        
        }        
      }
    }else{
    
      var nombreArea = "--";
    
    }
    
//    Recorremos datos de Etiquetas para identificar el nombre
    
    if(datos[i][9]!=""){
      
      for(var b = 0; b<datosEtiquetas.length; b++ ){       
          
        if(datos[i][9]==datosEtiquetas[b][0]){
            
          var nombreEtiqueta = datosEtiquetas[b][1];
          break;        
        }        
      }
    }else{
    
      var nombreEtiqueta = "--";
    
    }    
    
    if(accion=="listarTicketsAll" && datos[i][4]!="Borrado" && datos[i][4]!="Cerrado"){
      
      if(rol=="ADMIN"){
      
        var botones = "<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoSolicitud("+datos[i][0]+")' "+
                 "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a>"+
                 "<a class='btn btn-danger btn-sm' onclick='eliminarTicket("+datos[i][0]+")' "+
                 "title='Eliminar Ticket'><i class='fa-solid fa-trash-can'></i></a></center>";
      }else if(rol=="SPECIAL"){
      
        var botones = "<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoSolicitud("+datos[i][0]+")' "+
                 "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a></center>";
      } 
      resultado.push({
      
        id:datos[i][0],
        asunto:datos[i][1],
        estado:estadoB,
        idRespuesta:datos[i][5],
        agente:nombreAgente,
        cliente:nombreCliente,
        area:nombreArea,
        etiqueta:nombreEtiqueta,
        fechaI:datos[i][10],
        acciones:botones     
      });
     
    }else if(accion=="listarTicketsNuevos" && datos[i][4]=="Nuevo"){
       
      if(rol=="ADMIN"){
         
        var botones = "<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoSolicitud("+datos[i][0]+")' "+
                 "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a>"+
                 "<a class='btn btn-danger btn-sm' onclick='eliminarTicket("+datos[i][0]+")' "+
                 "title='Eliminar Ticket'><i class='fa-solid fa-trash-can'></i></a></center>";
      
      }else if(rol=="SPECIAL"){
         
        var botones = "<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoSolicitud("+datos[i][0]+")' "+
                 "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a></center>";
      
      }
      
      
      resultado.push({
      
        id:datos[i][0],
        asunto:datos[i][1],
        estado:estadoB,
        idRespuesta:datos[i][5],
        agente:nombreAgente,
        cliente:nombreCliente,
        area:nombreArea,
        etiqueta:nombreEtiqueta,
        fechaI:datos[i][10],
        acciones:botones    
      });
     
    }else if(accion=="listarTicketsPendientes" && datos[i][4]=="En Curso" ){
        
      if(rol=="ADMIN"){
         
        var botones = "<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoSolicitud("+datos[i][0]+")' "+
                       "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a>"+
                       "<a class='btn btn-danger btn-sm' onclick='eliminarTicket("+datos[i][0]+")' "+
                       "title='Eliminar Ticket'><i class='fa-solid fa-trash-can'></i></a></center>";
      
      }else if(rol=="SPECIAL"){
         
        var botones = "<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoSolicitud("+datos[i][0]+")' "+
                       "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a></center>";
      
      }
      
      
      resultado.push({
      
        id:datos[i][0],
        asunto:datos[i][1],
        estado:estadoB,
        idRespuesta:datos[i][5],
        agente:nombreAgente,
        cliente:nombreCliente,
        area:nombreArea,
        etiqueta:nombreEtiqueta,
        fechaI:datos[i][10],
        acciones:botones      
      });
     
    }else if(accion=="listarTicketsResueltos" && datos[i][4]=="Cerrado" ){
       
      if(rol=="ADMIN"){
      
         var botones = "<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoSolicitud("+datos[i][0]+")' "+
                 "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a>"+
                 "<a class='btn btn-danger btn-sm' onclick='eliminarTicket("+datos[i][0]+")' "+
                 "title='Eliminar Ticket'><i class='fa-solid fa-trash-can'></i></a></center>";
      }else if(rol=="SPECIAL"){
      
         var botones = "<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoSolicitud("+datos[i][0]+")' "+
                 "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a></center>";
      }
      
      
      
      resultado.push({
      
        id:datos[i][0],
        asunto:datos[i][1],
        estado:estadoB,
        idRespuesta:datos[i][5],
        agente:nombreAgente,
        cliente:nombreCliente,
        area:nombreArea,
        etiqueta:nombreEtiqueta,
        fechaI:datos[i][11],
        acciones:botones    
      });
     
    }else if(accion=="listarTicketsBorrados" && datos[i][4]=="Borrado" ){
    
      resultado.push({
      
        id:datos[i][0],
        asunto:datos[i][1],
        estado:estadoB,
        idRespuesta:datos[i][5],
        agente:nombreAgente,
        cliente:nombreCliente,
        area:nombreArea,
        etiqueta:nombreEtiqueta,
        fechaI:datos[i][11],
        acciones:"<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoSolicitud("+datos[i][0]+")' "+
                 "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a></center>"      
      });
     
    }else if(accion=="listarMisTickets" && datos[i][7]==usuarioActivo ){
      
      if(datos[i][4]=="Borrado" || datos[i][4] == "Cerrado"){
        
       var  fechaT = datos[i][11];
       }else{
       
        var  fechaT = datos[i][10]
        
       }
    
      resultado.push({
      
        id:datos[i][0],
        asunto:datos[i][1],
        estado:estadoB,
        idRespuesta:datos[i][5],
        agente:nombreAgente,
        cliente:nombreCliente,
        area:nombreArea,
        etiqueta:nombreEtiqueta,
        fechaI:fechaT,
        acciones:"<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoMisSolicitud("+datos[i][0]+")' "+
                 "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a></center>"      
      });
     
    }else if(accion=="listarMisTareas" && datos[i][6]==usuarioActivo && datos[i][4]!="Borrado" && datos[i][4]!="Cerrado"){
        
      resultado.push({
      
        id:datos[i][0],
        asunto:datos[i][1],
        estado:estadoB,
        idRespuesta:datos[i][5],
        agente:nombreAgente,
        cliente:nombreCliente,
        area:nombreArea,
        etiqueta:nombreEtiqueta,
        fechaI:datos[i][10],
        acciones:"<center><a class='btn btn-primary btn-sm posicionBoton' onclick='verInfoSolicitud("+datos[i][0]+")' "+
                 "title='Gestionar Ticket'><i class='fa-solid fa-pencil'></i></a></center>"      
      });
     
    }

  } 

   return resultado;  
}


/*============================================================

   ==================== Listar Tickets =================
   
==============================================================*/
function crearTickets(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tickets");
  var fila = hoja.getLastRow();
  var idActual = hoja.getRange("A"+fila).getValue();
  var usuarioActivo = Session.getActiveUser().getEmail();
  
  var fecha = new Date();
  var dia = fecha.getDate();
  var mes = fecha.getMonth()+1;
  var year = fecha.getFullYear();
  
  if(dia<10){
    
    dia = "0"+dia;
  
  }
  
  
  if(mes==1){mes ="Enero"}else if(mes==2){mes ="Febrero"}else if(mes==3){mes ="Marzo"}
  else if(mes==4){mes ="Abril"}else if(mes==5){mes ="Mayo"}else if(mes==6){mes ="Junio"}
  else if(mes==7){mes ="Julio"}else if(mes==8){mes ="Agosto"}else if(mes==9){mes ="Septiembre"}
  else if(mes==10){mes ="Octubre"}else if(mes==11){mes ="Noviembre"}else{ mes ="Diciembre"}
  
  if(form.txtAccion=="crearTickets" && form.asuntoTicket!="" && form.descipcionTicket!=""){
    
    hoja.getRange("A"+(fila+1)).setValue(idActual+1);
    hoja.getRange("B"+(fila+1)).setValue(form.asuntoTicket);
    hoja.getRange("C"+(fila+1)).setValue(form.descipcionTicket);
    hoja.getRange("E"+(fila+1)).setValue("Nuevo");
    
    hoja.getRange("D"+(fila+1)).setValue(form.extensionAdjunto);
    var extension = hoja.getRange("D"+(fila+1)).getValue();
    
    if(extension=="jpg" || extension=="JPG" || extension=="png" || extension=="jpeg" || extension=="PNG" || extension=="jfif" ){
      
      var carpeta = DriveApp.getFolderById("1-c9tZvVxQLnZrm96J75dmlJUWfjmxceN");
      var archivo = carpeta.createFile(form.adjunto);
      var idArchivo = archivo.getId();
      hoja.getRange("D"+(fila+1)).setValue(idArchivo);
    
    }
 
    hoja.getRange("H"+(fila+1)).setValue(usuarioActivo);
    hoja.getRange("K"+(fila+1)).setValue(dia+"-"+mes+"-"+year).setNumberFormat("@");
    
    hoja.getRange("N"+fila).autoFill(hoja.getRange("N"+fila+":N"+(fila+1)), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    
    
    return "okCrearTicket";
  }else{
  
     return "faltaInformacion";
  
  }
}

/*============================================================

   ==================== Ver Info Tickets =================
   
==============================================================*/
function infoTicket(id){
  
  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tickets");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:M"+fila).getValues();
  
  
  var ticket = [];
  var selectEstado = [];
  
  var hojaUsuarios = libro.getSheetByName("Usuarios");
  var filaAgentes = hojaUsuarios.getLastRow();
  var datosUsuarios = hojaUsuarios.getRange("A2:H"+filaAgentes).getValues();
  
  for(var i = 0; i<datos.length; i++ ){
    
    if(id==datos[i][0]){
      
      
      var selectAgentes = [];
      
      for(var j = 0; j<datosUsuarios.length; j++){
        if(datos[i][6]==datosUsuarios[j][3]){
          selectAgentes.push("<option value='"+datosUsuarios[j][3]+"'>"+datosUsuarios[j][1]+"</option>");
        }
      }
      
      if(datos[i][6]==""){
          selectAgentes.push("<option value='0'>Seleccionar Agente</option>");
          
          for(var j = 0; j<datosUsuarios.length; j++){
          if(datosUsuarios[j][6]=="SPECIAL" || datosUsuarios[j][6]=="ADMIN"){
            selectAgentes.push("<option value='"+datosUsuarios[j][3]+"'>"+datosUsuarios[j][1]+"</option>");
          }
        }        
      }else{

        for(var j = 0; j<datosUsuarios.length; j++){
          if(datosUsuarios[j][6]=="SPECIAL" && datos[i][6]!=datosUsuarios[j][3]  || datosUsuarios[j][6]=="ADMIN" && datos[i][6]!=datosUsuarios[j][3]){
            selectAgentes.push("<option value='"+datosUsuarios[j][3]+"'>"+datosUsuarios[j][1]+"</option>");
          }
        }
      
      }
      
      var hojaEtiquetas = libro.getSheetByName("Riesgos");
      var filaEtiquetas = hojaEtiquetas.getLastRow();
      var datosEtiquetas = hojaEtiquetas.getRange("A2:B"+filaEtiquetas).getValues();
      var selectEtiquetas = [];
      
      for(var k = 0; k<datosEtiquetas.length; k++){
        if(datos[i][9]==datosEtiquetas[k][0]){
           selectEtiquetas.push("<option value='"+datosEtiquetas[k][0]+"'>"+datosEtiquetas[k][1]+"</option>"); 
         }
      }
      
      if(datos[i][9]==""){
        selectEtiquetas.push("<option value='0'>Seleccionar T Soporte</option>");
        for(var k = 0; k<datosEtiquetas.length; k++){
          selectEtiquetas.push("<option value='"+datosEtiquetas[k][0]+"'>"+datosEtiquetas[k][1]+"</option>"); 
         }      
      }else{
        for(var k = 0; k<datosEtiquetas.length; k++){
          if(datos[i][9]!=datosEtiquetas[k][0]){
             selectEtiquetas.push("<option value='"+datosEtiquetas[k][0]+"'>"+datosEtiquetas[k][1]+"</option>") 
           }
        } 
      }
      
      var selectPrioridad = [];
      
      if(datos[i][12]==""){
        
        selectPrioridad.push("<option value='0'>Seleccionar Prioridad</option>"+
                           "<option value='BAJA'>Baja</option>"+
                           "<option value='NORMAL'>Normal</option>"+
                           "<option value='ALTA'>Alta</option>"+
                           "<option value='URGENTE'>Urgente</option>");
      
      }else{      
      selectPrioridad.push("<option value='"+datos[i][12]+"'>"+datos[i][12]+"</option>"+
                           "<option value='BAJA'>Baja</option>"+
                           "<option value='NORMAL'>Normal</option>"+
                           "<option value='ALTA'>Alta</option>"+
                           "<option value='URGENTE'>Urgente</option>");
        
      }
      
      
      var hojaAreas = libro.getSheetByName("Controles");
      var filaAreas = hojaAreas.getLastRow();
      var datosAreas = hojaAreas.getRange("A2:B"+filaAreas).getValues();
      var selectAreas = [];
      
      for(var l = 0; l<datosAreas.length; l++){
        if(datos[i][8]==datosAreas[l][0]){
          selectAreas.push("<option value='"+datosAreas[l][0]+"'>"+datosAreas[l][1]+"</option>");
        }
      }
      
      if(datos[i][8]==""){
        selectAreas.push("<option value='0'>Seleccionar Área</option>");
        for(var l = 0; l<datosAreas.length; l++){
          selectAreas.push("<option value='"+datosAreas[l][0]+"'>"+datosAreas[l][1]+"</option>");
        }
      }else{
        for(var l = 0; l<datosAreas.length; l++){
          if(datos[i][8]!=datosAreas[l][0]){
            selectAreas.push("<option value='"+datosAreas[l][0]+"'>"+datosAreas[l][1]+"</option>");
          }
        }
      }
      
      selectEstado.push("<option value='"+datos[i][4]+"'>"+datos[i][4]+"</option>"+
                       "<option value='Abierto'>Abierto</option>"+
                       "<option value='En Curso'>En Curso</option>"+
                       "<option value='Cerrado'>Cerrado</option>");
      
      
      
      for(var g =0; g<datosUsuarios.length; g++){
        if(datosUsuarios[g][3]==datos[i][7]){          
          var fotoCliente = datosUsuarios[g][2];
          var nombreCliente = datosUsuarios[g][1];        
        }
      }
      
      ticket.push({
        id:id,
        asunto:datos[i][1],
        descripcion:datos[i][2],
        adjunto:datos[i][3],
        estado:datos[i][4],
        estadoForm:selectEstado,
        respuesta:datos[i][5],
        agente:selectAgentes,
        cliente:datos[i][7],
        area:selectAreas,
        etiqueta:selectEtiquetas,
        prioridad:selectPrioridad,
        fechaI:datos[i][10],
        fechaF:datos[i][11],
        fotoCliente:fotoCliente,  
        nombreCliente:nombreCliente
      });      
    }    
  }
  
  
  var hojaRespuestas =  libro.getSheetByName("Respuestas");
  var filaRespuestas = hojaRespuestas.getLastRow();
  var datosRespuestas = hojaRespuestas.getRange("A2:F"+filaRespuestas).getValues();
  var respuestas = [];
  
  for(var h=0; h<datosRespuestas.length; h++){
    
    if(datosRespuestas[h][1] == id ){
      
      for(var n = 0; n<datosUsuarios.length; n++){        
        if(datosUsuarios[n][3]==datosRespuestas[h][4]){
          var nombreAgente = datosUsuarios[n][1];
          var fotoAgente = datosUsuarios[n][2];
          
        }     
      }
            
      respuestas.push({        
        idRespuestas:datosRespuestas[h][0],
        adjuntoRespuesta:datosRespuestas[h][2],
        descripcionRespuesta:datosRespuestas[h][3],
        usuario:datosRespuestas[h][4],
        fecha:datosRespuestas[h][5],
        nombreAgente:nombreAgente,
        fotoAgente:fotoAgente
      });
    
    }
   
  } 
  
  
  var respuestaFinal = [];
  
  respuestaFinal.push({
    respuestaA:ticket,
    respuestas:respuestas
  });
   
  return respuestaFinal;
}


/*=======================================================================

     =====================  Cerrar Ticket  =======================

==============================================================================*/
function cerrarTicketAdmin(form){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tickets");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:M"+fila).getValues();
  
  for(var i = 0; i<datos.length; i++){
    if(form.txtIdCerrarTicket == datos[i][0]){
      
      if(form.txtEstadoTicket == "Cerrado"){
        
        if(!!form.txtAgenteAsignadoTicket && form.txtAgenteAsignadoTicket != "0" &&
           !!form.txtTipoSoporteTicket && form.txtTipoSoporteTicket != "0" &&
           !!form.txtPrioridadTicket && form.txtPrioridadTicket !="0" &&
           !!form.txtEstadoTicket &&
           !! form.txtAreaTicket && form.txtAreaTicket != "0" && 
           Boolean(form.txtRespuestaTicket)){
          
          hoja.getRange("E"+(i+2)).setValue(form.txtEstadoTicket);
          hoja.getRange("G"+(i+2)).setValue(form.txtAgenteAsignadoTicket);
          hoja.getRange("I"+(i+2)).setValue(form.txtAreaTicket);
          hoja.getRange("J"+(i+2)).setValue(form.txtTipoSoporteTicket);
          
          var fecha = new Date();
          var dia = fecha.getDate();
          var mes = fecha.getMonth()+1;
          var year = fecha.getFullYear();
          
          if(mes==1){mes="Enero"}else if(mes==2){mes="Febrero"}else if(mes==3){mes="Marzo"}
          else if(mes==4){mes="Abril"}else if(mes==5){mes="Mayo"}else if(mes==6){mes="Junio"}
          else if(mes==7){mes="Julio"}else if(mes==8){mes="Agosto"}else if(mes==9){mes="Septiembre"}
          else if(mes==10){mes="Octubre"}else if(mes==11){mes="Noviembre"}else if(mes==12){mes="Diciembre"}
          
          hoja.getRange("L"+(i+2)).setValue(dia+"-"+mes+"-"+year).setNumberFormat("@");
          hoja.getRange("M"+(i+2)).setValue(form.txtPrioridadTicket);
        
        }else{
          
          return "faltaInfoCerrarTicket";
        
        }      
      }else{
        
        if(form.txtAgenteAsignadoTicket != "0" && !!form.txtAgenteAsignadoTicket){
          
          if(form.txtEstadoTicket=="Nuevo"){
            hoja.getRange("E"+(i+2)).setValue("Abierto");
          }else{
            hoja.getRange("E"+(i+2)).setValue(form.txtEstadoTicket);
          }
          hoja.getRange("G"+(i+2)).setValue(form.txtAgenteAsignadoTicket);
        }
        
        if(form.txtAreaTicket != "0" && !!form.txtAreaTicket){
          hoja.getRange("I"+(i+2)).setValue(form.txtAreaTicket);
        }
        
        if(form.txtTipoSoporteTicket != "0" && !!form.txtTipoSoporteTicket ){        
        hoja.getRange("J"+(i+2)).setValue(form.txtTipoSoporteTicket);
        }  
          
        if(form.txtPrioridadTicket != "0" && !!form.txtPrioridadTicket){
        hoja.getRange("M"+(i+2)).setValue(form.txtPrioridadTicket);
        }  
      
      } 
      
      if(Boolean(form.txtRespuestaTicket) || !!form.txtExtensionEvidenciaTicket){
          
          var hojaRespuesta = libro.getSheetByName("Respuestas");
          var filaRespuestas = hojaRespuesta.getLastRow();
          var idActualRespuesta = hoja.getRange("A"+filaRespuestas).getValue();
         
          hojaRespuesta.getRange("A"+(filaRespuestas+1)).setValue(idActualRespuesta+1);
          hojaRespuesta.getRange("B"+(filaRespuestas+1)).setValue(form.txtIdCerrarTicket);
        
          
         if(!!form.txtExtensionEvidenciaTicket){
             var carpeta = DriveApp.getFolderById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
             var archivo = carpeta.createFile(form.txtEvidenciaTicket);
             var idArchivo = archivo.getId();
             hojaRespuesta.getRange("C"+(filaRespuestas+1)).setValue(idArchivo);
              
          }
        
        if(Boolean(form.txtRespuestaTicket)){
         
          hojaRespuesta.getRange("D"+(filaRespuestas+1)).setValue(form.txtRespuestaTicket);
          
        }
         var usuarioActivo = Session.getActiveUser().getEmail();
          hojaRespuesta.getRange("E"+(filaRespuestas+1)).setValue(usuarioActivo);
        
          var fecha = new Date();
          var dia = fecha.getDate();
          var mes = fecha.getMonth()+1;
          var year = fecha.getFullYear();
          
          if(mes==1){mes="Enero"}else if(mes==2){mes="Febrero"}else if(mes==3){mes="Marzo"}
          else if(mes==4){mes="Abril"}else if(mes==5){mes="Mayo"}else if(mes==6){mes="Junio"}
          else if(mes==7){mes="Julio"}else if(mes==8){mes="Agosto"}else if(mes==9){mes="Septiembre"}
          else if(mes==10){mes="Octubre"}else if(mes==11){mes="Noviembre"}else if(mes==12){mes="Diciembre"}
          
          hojaRespuesta.getRange("F"+(filaRespuestas+1)).setValue(dia+"-"+mes+"-"+year).setNumberFormat("@");
        
        
      }
    }
  }
       
  
  if(form.txtAccion == "respuestaCliente"){
     
    for(var j=0; j<datos.length; j++){      
      if(form.txtIdCerrarTicket == datos[j][0]){
        if(datos[j][4]=="Cerrado" || datos[j][4]=="Borrado"){
        
          hoja.getRange("E"+(j+2)).setValue("Abierto");
          
        }        
      }
    }
    
    return "okCliente"
  }else{
    return "okCerrarTicket";
  }
}


/*=======================================================================

     =====================  Eliminar Ticket  =======================

==============================================================================*/
function eliminarTicketAdmin(id){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Tickets");
  var fila = hoja.getLastRow();
  var datos = hoja.getRange("A2:E"+fila).getValues();
  
  for(var i=0; i<datos.length; i++){
     
    if(datos[i][0] == id){
       
      hoja.getRange("E"+(i+2)).setValue("Borrado");
      
      return "ticketEliminado";
    } 
  }
}


/*=======================================================================

     =====================  Ver Graficas  =======================

==============================================================================*/
function verGraficasAdmin(){

  var libro = SpreadsheetApp.openById("1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E");
  var hoja = libro.getSheetByName("Smart");
  
  var tickets = hoja.getRange("B3:B5").getValues();
  
  var resultado = [];
  var arrayTickets = [];
  
  arrayTickets.push({
    
    tickets:tickets
  });
  
  
  var fecha = new Date();
  var mes = fecha.getMonth()+1;
  
  if(mes==1){    
    var ticketMes = hoja.getRange("D3:G3").getValues();  
  }else if(mes==2){    
   var ticketMes = hoja.getRange("D3:G4").getValues();  
  }else if(mes==3){    
    var ticketMes = hoja.getRange("D3:G5").getValues();  
  }else if(mes==4){    
    var ticketMes = hoja.getRange("D3:G6").getValues();  
  }else if(mes==5){    
    var ticketMes = hoja.getRange("D3:G7").getValues();  
  }else if(mes==6){    
    var ticketMes = hoja.getRange("D3:G8").getValues();  
  }else if(mes==7){    
    var ticketMes = hoja.getRange("D3:G9").getValues();  
  }else if(mes==8){    
    var ticketMes = hoja.getRange("D3:G10").getValues();  
  }else if(mes==9){    
    var ticketMes = hoja.getRange("D3:G11").getValues();  
  }else if(mes==10){    
    var ticketMes = hoja.getRange("D3:G12").getValues();  
  }else if(mes==11){    
    var ticketMes = hoja.getRange("D3:G13").getValues();  
  }else if(mes==12){    
    var ticketMes = hoja.getRange("D3:G14").getValues();  
  } 
  
  var arrayTicketsMes = [];
  
  for(var i = 0; i<ticketMes.length; i++){
   
    arrayTicketsMes.push({
      mes:ticketMes[i][0],
      abiertos:ticketMes[i][1],
      enCurso:ticketMes[i][2],
      cerrados:ticketMes[i][3]    
    })  
  }
  

// Tickets Agentes 

  var hoja2 = libro.getSheetByName("Smart2");
  var filaAgentes = hoja2.getLastRow()
  var ticketsAgentes  = hoja2.getRange("B3:E"+filaAgentes).getValues();
  var arrayAgentes = [];
  
  for(var j=0; j<ticketsAgentes.length; j++){
    
    if(ticketsAgentes[j][0]!=""){
      arrayAgentes.push({
        nombre:ticketsAgentes[j][0],
        abiertos:ticketsAgentes[j][1],
        enCurso:ticketsAgentes[j][2],
        cerrados:ticketsAgentes[j][3]
      
      })
    }  
  } 
  
  
//  Tickets By Etiquetas
  
  var etiquetas = hoja2.getRange("H3:I7").getValues();
  var arrayEtiquetas = [];
  var colores = ['rgb(255, 99, 132)', 'rgb(255, 159, 64)',
                 'rgb(255, 205, 86)', 'rgb(75, 192, 192)',
                 'rgb(54, 162, 235)'];
  
  for(var e = 0; e<etiquetas.length; e++){
    
    if(etiquetas[e][0]!=""){
      
      arrayEtiquetas.push({
    
      nombre:etiquetas[e][0],
      tickets:etiquetas[e][1],
      colores:colores[e]    
    });
    
    }  
  }
  
//  Tickets By Areas
  
  var areas = hoja2.getRange("K3:L7").getValues();
  var arrayAreas = [];
  
  for(var a=0; a<areas.length; a++){
   
    if(areas[a][0]!=""){
      arrayAreas.push({
        nombre:areas[a][0],
        tickets:areas[a][1]        
      });
    }  
  }
  
  
//  Cantidad de Tickets por usuario
  
  var usuarios = hoja2.getRange("O3:Q7").getValues();
  var arrayUsuarios = [];
  
  for(var u=0; u<usuarios.length; u++){
    if(usuarios[u][0]!=""){
      arrayUsuarios.push({
       
        nombre:usuarios[u][0],
        foto:usuarios[u][1],
        tickets:usuarios[u][2]
      });
    }
  }
  
  resultado.push({
    
    arrayTickets:arrayTickets,
    arrayTicketsMes:arrayTicketsMes,
    arrayAgentes:arrayAgentes,
    arrayEtiquetas:arrayEtiquetas,
    arrayAreas:arrayAreas,
    arrayUsuarios:arrayUsuarios
     
  })
  
//  Logger.log(resultado);
  
  return resultado;
}















