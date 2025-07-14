/*============================================================

   ==================== WEB  ==========================
   
==============================================================*/

function doGet() {
  
  var html = HtmlService.createTemplateFromFile("index");
  
  return html.evaluate().setTitle("ARIS")
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  .addMetaTag("viewport", "width=device-width, initial-scale=1")
  .setFaviconUrl("https://raw.githubusercontent.com/DAVEUBAQUE1996/dave/main/icon-bolivar-conmigo.png");
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Index'); // Suponiendo que tienes un archivo HTML llamado Index.html
  
  // Configurar cabeceras de seguridad
  var headers = {
    'X-Content-Type-Options': 'nosniff',
    'X-XSS-Protection': '1; mode=block',
    'Strict-Transport-Security': 'max-age=31536000; includeSubDomains',
    'Content-Security-Policy': "default-src 'self'; script-src 'self' 'unsafe-inline'; object-src 'none'",
    'Referrer-Policy': 'no-referrer'
  };
  
  // Aplicar las cabeceras de seguridad
  var response = HtmlService.createHtmlOutput(htmlOutput.getContent())
                   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  for (var header in headers) {
    response.append(header + ": " + headers[header] + "\n");
  }
  
  return response;

  
}


function getMenuData() {
  return {
    link1: 'https://www.superfinanciera.gov.co/publicaciones/10086577/industrias-supervisadassistema-de-administracion-de-riesgo-de-lavado-de-activos-y-financiacion-del-terrorismo-10086577/',
    link2: 'https://www.supersociedades.gov.co/web/asuntos-economicos-societarios/sagrilaft',
    link3: 'https://www.supersociedades.gov.co/web/asuntos-economicos-societarios/ptee',
    link4: 'https://docs.supersalud.gov.co/PortalWeb/Juridica/CircularesExterna/CIRCULAR%20EXTERNA%20202117000000055.pdf',
    link5: 'https://www.supervigilancia.gov.co/publicaciones/10532/capacitacion-sarlaft-20/'
  };
}

/*============================================================

   ==================== Paginador  ==========================
   
==============================================================*/

function incluirArchivo(pagina){

  var html = HtmlService.createTemplateFromFile(pagina);
  
  return html.evaluate().getContent();


}


/*============================================================

   ==================== Envio PQR  ==========================
   
==============================================================*/
function enviarPQRControl(form){

  var html = HtmlService.createHtmlOutputFromFile("plantillaPQR").getContent();
  
  var fecha = new Date();
  var dia = fecha.getDate();
  var mes = fecha.getMonth()+1;
  var year = fecha.getFullYear();
  
  html = html.replace("|*tipoPQR*|", form.itemPQR);
  html = html.replace("|*descripcionPQR*|", form.descripcionPQR);
  html = html.replace("|*dia*|", dia);
  html = html.replace("|*mes*|", mes);
  html = html.replace("|*year*|", year);
  
  GmailApp.sendEmail("javier.ubaque@segurosbolivar.com", "Nueva Denuncia!!", html, {htmlBody:html} );

   return "okPQR";

}


/*=======================================================================

     =====================  Controlador de Acciones  =======================

==============================================================================*/

function controlador(form, accion, id){
 
  if(form!=""){
    
    accion = form.txtAccion;
  }
  
  var rol;
  
  if(accion[0]=="listarEtiquetas" || accion[0]=="listarAreas"|| accion[0]=="listarcausas"||accion[0]=="listaracti"|| accion[0]=="listarlistas" ||accion[0]=="listarpais"||accion[0]=="listarjurisd"||accion[0]=="listarrequ"||accion[0]=="listardash"||accion[0]=="listarmoni"||accion[0]=="listarmatriz"||accion[0]=="listarprofes"|| accion[0]=="listarperson"||accion[0]=="listartrans"||accion[0]=="listarcanal"||accion[0]=="listarprod"  ||accion[0]=="listaralerta"||accion[0]=="listartest"||accion[0]=="listarriesgos"  || accion[0] =="listarUsuarios" || accion[0]=="listarocupa" || accion[0]=="listarTicketsNuevos" || accion[0]=="listarTicketsAll" || accion[0]=="listarTicketsPendientes" ||
     accion[0]=="listarTicketsResueltos"){
    
    rol = accion[1];
    accion = accion[0];
  
  }
  try{
    
    switch(accion){
      
      case "listarEtiquetas":
        return listarEtiquetas(id,rol);
        break;
     case "crearEtiqueta":
        return crearEtiqueta(form);
        break;
     case "editarEtiqueta":
        return crearEtiqueta(form);
        break;
     case "listarAreas":
        return listarAreas(id,rol);
        break; 
     case "crearArea":
        return crearArea(form);
        break; 
     case "editarAreas":
        return crearArea(form);
        break; 
     case "listartest":
        return listartest(id,rol);
        break; 
     case "creartest":
        return creartest(form);
        break; 
     case "editartest":
        return creartest(form);
        break;         
     case "listarcausas":
        return listarcausas(id,rol);
        break;         
     case "crearcausas":
        return crearcausas(form);
        break; 
     case "editarcausas":
        return crearcausas(form);
        break;
     case "listarmoni":
        return listarmoni(id,rol);
        break;         
     case "crearmoni":
        return crearmoni(form);
        break; 
     case "editarmoni":
        return crearmoni(form);
        break;        
     case "listaracti":
        return listaracti(id,rol);
        break;         
     case "crearacti":
        return crearacti(form);
        break; 
     case "editaracti":
        return crearacti(form);
        break;        
     case "listarlistas":
        return listarlistas(id,rol);
        break;         
     case "crearlistas":
        return crearlistas(form);
        break; 
     case "editarlistas":
        return crearlistas(form);
        break;           
     case "listarpais":
        return listarpais(id,rol);
        break;         
     case "crearpais":
        return crearpais(form);
        break; 
     case "editarpais":
        return crearpais(form);
        break; 
     case "listarjurisd":
        return listarjurisd(id,rol);
        break;         
     case "crearjurisd":
        return crearjurisd(form);
        break; 
     case "editarjurisd":
        return crearjurisd(form);
        break;  
     case "listarrequ":
        return listarrequ(id,rol);
        break;         
     case "crearrequ":
        return crearrequ(form);
        break; 
     case "editarrequ":
        return crearrequ(form);
        break;         
     case "listardash":
        return listardash(id,rol);
        break;         
     case "creardash":
        return creardash(form);
        break; 
     case "editardash":
        return creardash(form);
        break;         
     case "listarmatriz":
        return listarmatriz(id,rol);
        break;         
     case "crearmatriz":
        return crearmatriz(form);
        break; 
     case "editarmatriz":
        return crearmatriz(form);
        break;  
     case "listarprofes":
        return listarprofes(id,rol);
        break;         
     case "crearprofes":
        return crearprofes(form);
        break; 
     case "editarprofes":
        return crearprofes(form);
        break;  

     case "listarocupa":
        return listarocupa(id,rol);
        break;         
     case "crearocupa":
        return crearocupa(form);
        break; 
     case "editarocupa":
        return crearocupa(form);
        break;        

     case "listarperson":
        return listarperson(id,rol);
        break;         
     case "crearperson":
        return crearperson(form);
        break; 
     case "editarperson":
        return crearperson(form);
        break;         
     case "listartrans":
        return listartrans(id,rol);
        break;         
     case "creartrans":
        return creartrans(form);
        break; 
     case "editartrans":
        return creartrans(form);
        break;         
     case "listarcanal":
        return listarcanal(id,rol);
        break;         
     case "crearcanal":
        return crearcanal(form);
        break; 
     case "editarcanal":
        return crearcanal(form);
        break;         
     case "listarprod":
        return listarprod(id,rol);
        break;         
     case "crearprod":
        return crearprod(form);
        break; 
     case "editarprod":
        return crearprod(form);
        break;                  
     case "listaralerta":
        return listaralerta(id,rol);
        break;         
     case "crearalerta":
        return crearalerta(form);
        break; 
     case "editaralerta":
        return crearalerta(form);
        break;         
     case "listarriesgos":
        return listarriesgos(id,rol);
        break;         
     case "crearriesgos":
        return crearriesgos(form);
        break; 
     case "editarriesgos":
        return crearriesgos(form);
        break;               
     case "listarUsuarios":
        return listarUsuarios(id,rol);
        break;
    case "listarSelectAreas":
        return listarAreasSelect();
        break;   
   case "crearUsuarios":
        return crearUsuario(form);
        break;  
  case "editarUsuario":
        return crearUsuario(form);
        break;   
  case "listarTicketsAll":
        return listarTickets(accion,rol);
        break;
  case "listarTicketsNuevos":
        return listarTickets(accion,rol);
        break;
  case "listarTicketsPendientes":
        return listarTickets(accion,rol);
        break;
  case "listarTicketsResueltos":
        return listarTickets(accion,rol);
        break;
  case "listarTicketsBorrados":
        return listarTickets(accion);
        break; 
  case "listarMisTickets":
        return listarTickets(accion);
        break; 
  case "listarMisTareas":
        return listarTickets(accion);
        break;      
     
  case "crearTickets":
        return crearTickets(form);
        break; 
  case "verInfoTicket":
        return infoTicket(id);
        break;
  case "cerrarTicket":
        return cerrarTicketAdmin(form);
        break;
  case "respuestaCliente":
        return cerrarTicketAdmin(form);
        break;
  case "verGraficasAdmin":
        return verGraficasAdmin();
        break;    
        
        
      
    }   
  
  }catch(e){
    
    return "Error en el Controldor "+e;
  
  }
}

/*=======================================================================

     =====================  Obener informacion de la hoja de calculo para tablero en front   =======================

==============================================================================*/

function obtenerCifrasARIS() {
  const sheet = SpreadsheetApp.openById('1Zgs6gpF2UBd3eB3LR_vm3pyBwwk6mMLNmH_bi-YwE8E')
                             .getSheetByName('CIFRAS');
  const valores = sheet.getRange("B1:B200").getValues().flat(); // Aplana la matriz en arreglo
  return valores;
}