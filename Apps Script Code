/////////////////////////////////////////////////////////////////////////////
/////// FIRST MOMENT OF EXECUTION: GENERATE LIST OF FILES TO DOWNLOAD ///////
/////////////////////////////////////////////////////////////////////////////

//Definir función para generar lista de archivos a importar
function List_Thumbnails() {
/* Adapted from Code written by @hubgit https://gist.github.com/hubgit/3755293
Updated since DocsList is deprecated  https://ctrlq.org/code/19854-list-files-in-google-drive-folder
*/


  // List all files in a single folder on Google Drive
  // declare an array of folders
  var folderNames = ['Archivos de caso','Archivos de resolución (Demandas)', 'Archivos de resolución (Convenios)',
                     'Archivos de resolución (Laborales)', 'Archivos de resolución (Derechos Humanos)', 
                     'Archivos de resolución (Recursos de revisión)', 'Archivos de resolución (Carpetas penales)', 
                     'Archivos de resolución (Transparencia)']

  // declare this sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName('Importador');
  // clear any existing contents
  sheet.clear();
  // append a header row
  sheet.appendRow(["Folder","Name", "Date Last Updated", "Size", "URL", "ID", "Description", "Type", "Download", "Date Created", "ID de área", "Tabla contenedora"]);


  // getFoldersByName = Gets a collection of all folders in the user's Drive that have the given name.
  // folders is a "Folder Iterator" but there is only one unique folder name called, so it has only one value (next)

  for (var i=0;i<folderNames.length;i++){

    var folders = DriveApp.getFoldersByName(folderNames[i]);
    var foldersnext = folders.next();
    // Logger.log("THE FOLDER IS "+foldersnext);// DEBUG

    // declare an array to push data into the spreadsheet
    var data = [];

    // list files in this folder
    // myfiles is a File Iterator
    var myfiles = foldersnext.getFiles();

    // Logger.log("FILES IN THIS FOLDER"); DEBUG

    // loop through files in this folder
    while (myfiles.hasNext()) {
      var myfile = myfiles.next();
      var fname = myfile.getName();
      var fdate = myfile.getLastUpdated(); 
      var fsize = myfile.getSize();
      var furl = myfile.getUrl();
      var fid = myfile.getId();
      var fdesc = myfile.getDescription();
      var ftype = myfile.getMimeType();
      var fdown = myfile.getDownloadUrl();
      var fcreate = myfile.getDateCreated()
      
      //Logger.log("File Name is "+myfile.getName()); //Logger.log("Date is "+myfile.getLastUpdated()); //Logger.log("Size is "+myfile.getSize());
      //Logger.log("URL is "+myfile.getUrl()); //Logger.log("ID is "+myfile.getId()); //Logger.log("Description is "+myfile.getDescription());
      //Logger.log("File Type is "+myfile.getMimeType());

      // Populate the array for this file
      data = [ 
        foldersnext,
        fname,
        fdate,
        fsize,
        furl,
        fid,
        fdesc,
        ftype,
        fdown,
        fcreate
      ];
      //Logger.log("data = "+data); //DEBUG
      sheet.appendRow(data);
    }// Completes listing of the files in the named folder
  } // Completes loop of Folder names
}

////////////////////////////////////////////////////////////////////////////////////////
/////// SECOND MOMENT OF EXECUTION: REMOVE PREVIOUSLY DOWNLOADED FILES FROM LIST ///////
////////////////////////////////////////////////////////////////////////////////////////

//Definir una función para eliminar registros de importación de archivos descargados previamente
function RemovePrevs(){
  //Definir entorno de búsqueda
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Importador");
  var chck = sheet.getRange(2, 6, sheet.getLastRow()-1, 1).getValues().flat();
  var hist = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Histórico");
  var excls = hist.getRange(2, 6, hist.getLastRow()-1, 1).getValues().flat();

  //Mapear archivos descargados anteriormente
  var output = chck.map(row =>{
    //Definir cuáles están presentes en hoja de importación
    var index = excls.indexOf(row);
    //Incorporar control de verificación
    if(index >= 0){
      //Casos descargados anteriormente
      return ["Descargado"];
    } else{
      return ["Undefined"];
    }
  })

  //Incorporar marca de descarga en hoja de trabajo
  var mk = sheet.getRange(2,sheet.getLastColumn()+1,output.length,1);
  mk.setValues(output);

  //Eliminar filas de archivos ya descargados
  for (var i = sheet.getLastRow(); i > 0; i--) {
    var data = sheet.getRange(i,sheet.getLastColumn()).getValue();
    if (data == 'Descargado') {
      sheet.deleteRow(i);
    }
  }

  //Remover la columna de verificación
  var clear_range = sheet.getRange(2,sheet.getLastColumn(),sheet.getLastRow(),1);
  clear_range.clearContent();
}

///////////////////////////////////////////////////////////////////////////////////////////////
/////// THIRD MOMENT OF EXECUTION: GENERATE ID AREA KEYS FOR THE FILES TO BE DOWNLOADED ///////
///////////////////////////////////////////////////////////////////////////////////////////////

//Definir una función para mapear a qué área pertenece cada archivo
function MapeadorArea() {
  //Definir entorno de búsqueda
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Importador");
  var area = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues();

  //Deposito de ID
  var ids = ""; //ID de tabla contenedora
  var nash = ""; //Nombre de hoja en tabla contenedora
  var resls = [];
  var nals = []; 

  //Verificar el ID de cada área
  for(var i = 0; i < area.length-1; i++){
    switch(area[i][0]){
      //Demandas
      case "Archivos de resolución (Demandas)":
      ids = "*********************************";
      nash = "Demandas";
      break;
      //Juicios laborales
      case "Archivos de resolución (Laborales)":
      ids = "*********************************";
      nash = "Laborales";
      break;
      //Recursos de revisión
      case "Archivos de resolución (Recursos de revisión)":
      ids = "*********************************";
      nash = "Recursos";
      break;
      //Solicitudes de transparencia
      case "Archivos de resolución (Transparencia)":
      ids = "*********************************"; 
      nash = "Solicitudes";
      break;
      //Carpetas penales
      case "Archivos de resolución (Carpetas penales)":
      ids = "*********************************";
      nash = "Carpetas"
      break;
      //Derechos humanoos
      case "Archivos de resolución (Derechos Humanos)":
      ids = "*********************************";
      nash = "Carpetas"
      break;
      //Convenios
      case "Archivos de resolución (Convenios)":
      ids = "*********************************";
      nash = "Convenios"
      break;
      //Carpeta base (root)
      case "Archivos de caso":
      ids = "Undefined";
      nash = "Undefined"

    }
  //Incorporar ID de formulario a lista
  resls.push([ids]);
  nals.push([nash]);
  }

  //Definir área de impresión de valores
  var printRes = sheet.getRange(2, 11, resls.length, 1);
  var printSh = sheet.getRange(2, 12, nals.length, 1);

  //Imprimir lista de form IDs en hoja
  printRes.setValues(resls);
  printSh.setValues(nals);
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////// FOURTH MOMENT OF EXECUTION: DOWNLOAOD DRIVE FILES IN LOCAL PATHS AND GENERATE ACCESS KEYS FOR EACH FILE ///////
//////////////////////////////////////////////////// IN PYTHON ////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////////////////////
/////// FIFTH MOMENT OF EXECUTION: MATCH LOCAL PATHS WITH DRIVE URLS AND SUBSTITUTE THE LATTER ///////
//////////////////////////////////////////////////////////////////////////////////////////////////////

//Definir función para mapear path local en columna de archivo de la tabla contenedora
function PathMatcher(){
  //Definir entorno de búsqueda
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Importador");
  var path = sheet.getRange(2, sheet.getLastColumn(), sheet.getLastRow(), 1).getValues();
  var id = sheet.getRange(2, 6, sheet.getLastRow(), 1).getValues();
  var she = sheet.getRange(2, 11, sheet.getLastRow(), 1).getValues();
  var nam = sheet.getRange(2, 12, sheet.getLastRow(), 1).getValues();
  
  //Implementar depuración
  for(i = 0; i < she.length-1; i++){
    //Abrir tabla contenedora del archivo
    var tab = SpreadsheetApp.openById(she[i][0]).getSheetByName(nam[i][0]);
    //Definir rango de búsqueda del ID de archivo
    var idr = tab.getRange(2,tab.getLastColumn()-1,tab.getLastRow(),1);
    //Construir clave de URL de reemplazo
    const url = "https://drive.google.com/open?id="+id[i][0];
    //console.log(url);
    //Buscar y reemplazar texto de url por el path
    idr.createTextFinder(url).replaceAllWith(path[i][0]);
  }
}

///////////////////////////////////////////////////////////////////////////
/////// SIXTH MOMENT OF EXECUTION: CLEAN THE LIST OF FILE DOWNLOADS ///////
///////////////////////////////////////////////////////////////////////////

//Definir función copiar registros de histórico de importaciones
function SheetCleaner(){
  //Definir entorno de búsqueda
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Importador");
  var hist = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Histórico");
  
  //Definir rango de eliminación
  start = 2
  end = sheet.getLastRow() - start + 1;

  //Definir valores a copiar 
  var sortRange = sheet.getSheetValues(2, 1, sheet.getLastRow(), sheet.getLastColumn());

  //Hacer copiado y pegado de valores
  for(var i = 1; i <= sortRange.length; i++){
    //Definir variable de filtrado
    let filt = sheet.getRange(i, 6).getValue();
    let contr = sheet.getRange(i, 13).getValue();

    //Copiar si se confirmó la eliminación
    if(filt == contr){
      let rowVals = sheet.getRange(i, 1, 1, 14).getValues();
      hist.getRange(hist.getLastRow()+1, 1, 1, 14).setValues(rowVals);
    }
  }

  //Eliminar datos de respuestas
  sheet.deleteRows(start, end);
}

///////////////////////////////////////////////////////////////////////////////
/////// SEVENTH MOMENT OF EXECUTION: REMOVE DOWNLOADED FILES FROM DRIVE ///////
///////////////////////////////////////////////////////////////////////////////

//Definir función para eliminar archivos de carpetas de Drive
function DriveCleaner(){
  //Definir entorno de búsqueda: Fill list with names of folders
  var folderNames = ['Testing']

  //Revisar iterativamente folders de área
  for (var i=0;i<folderNames.length;i++){

    //Definir folder a revisar
    var folder = DriveApp.getFoldersByName(folderNames[i]);
    var foldersnext = folder.next();
  
    //Obtener archivos almacenados en folder
    var myfiles = foldersnext.getFiles();

    //Enviar archivos a la papelera
    while (myfiles.hasNext()) {
        const file = myfiles.next();
        Logger.log('Moving file to trash: ', file);
        file.setTrashed(true);
        // Instrucción para borrar permanentemente
        //Drive.Files.remove(file.getId())
    } 
  }
}




