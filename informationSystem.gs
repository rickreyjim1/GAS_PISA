// función que permite la generación de un menú adicional en la hoja de cálculo, con varios submenús
function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu('Evaluación')
  .addItem('Recibir Documento1','documento1')
  // crear funcion documento1
  .addToUi();
}
//Prueba para escoger un valor determinado
function myFunction() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respuestas");
  var value = sheet.getRange(sheet.getLastRow(), 14).getValue();
  if(value == 'Agua residual tratamiento/Wastewater treatment'){
    sheet.setActiveRange(sheet.getRange(sheet.getLastRow(), 21)).setValue('Yes');
  } else {
    sheet.setActiveRange(sheet.getRange(sheet.getLastRow(), 21)).setValue('No');
  }
}
//Recibe la información, genera las columnas de las 5 palabras clave, genera carpeta de usuario
function leerEstructura() {//function name
  var ss = SpreadsheetApp.getActiveSpreadsheet();//some gBook: -> this one gBook
  var hojaGeneradora = ss.getSheetByName("GenForm");//the form generator gSheet
  var lastColumn = hojaGeneradora.getLastColumn();//for counter
  var formTitle = hojaGeneradora.getRange("B24").getValue().toString();
  var form;
  var formExist = hojaGeneradora.getRange("A25").getValue().toString();
  if(formExist=="urlFormWillBeHere" || formExist == ""){
    form = FormApp.create(formTitle);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  } else {
     form = FormApp.openByUrl(formExist); 
  }
  form.setTitle(formTitle);
  //form.setCollectEmails(true); Can't use because will be public, not limited view
  var formURL = form.getPublishedUrl();
  hojaGeneradora.getRange("B25").setValue(formURL);
  //var url = UrlShortener.Url.insert({"longUrl": hojaGeneradora.getRange("B25")});
  //var urlID = url.getId();
  //var shortURL = hojaGeneradora.getRange("B26").setValue(urlID);
  //Logger.log(urlID);
  if(hojaGeneradora.getRange("B27").getValue().toString()!="" && hojaGeneradora.getRange("A27").getValue().toString() != "formLogoUrlHere"){
    var img = UrlFetchApp.fetch(hojaGeneradora.getRange("B1").getValue().toString());
    form.addImageItem()
     //.setTitle('Logo')
     .setHelpText(formTitle + " logo") 
     .setImage(img);
  }
  for(i=2; i<=lastColumn; i++){
   var inputValues = hojaGeneradora.getRange(2, i, 11, 1).getValues();
   if(parseInt(inputValues[8][0])==0){
     var inputType = inputValues[2][0];
     switch (inputType) {
       case "text":
         TextInputItem(inputValues, form);
         break;
       case "paragraph":
         TextAreaItem(inputValues, form);
         break;
       case "multiple":
          MultipleChoiceItem(inputValues, form);
         break;
       case "multiple+other":
         MultipleChoiceOtherItem(inputValues, form);
         break;
       case "checkbox":
         CheckboxItem(inputValues, form);
         break;
       case "checkbox+other":
         CheckboxOtherItem(inputValues, form);
         break;
       case "list":
         ListItem(inputValues, form);
         break;
       case "grid":
          GridItem(inputValues, form);
         break;
       case "breakPage":
         BreakItem(inputValues, form);
     }
     hojaGeneradora.getRange(10,i,1,1).getCell(1,1).setValue(1);
   }
  }
}

// función que crea selección de lista
function ListItem(valores, form){
  var item = form.addListItem();
  var lista = valores[7][0].split(";");
  item.setRequired(true);
  item.setTitle(valores[1][0]);
  item.setChoiceValues(lista);
  item.setHelpText(valores[4][0]);
}

// función que crea el campo de texto
function TextInputItem(valores, form){
  var item = form.addTextItem();
  item.setTitle(valores[1][0]);
  item.setRequired(!!parseInt(valores[3][0]));
  item.setHelpText(valores[4][0]);
  Logger.log(item.getId());
}

// función que crea un campo de texto más amplio
function TextAreaItem(valores, form){
  var item = form.addParagraphTextItem();
  item.setTitle(valores[1][0]);
  item.setRequired(!!parseInt(valores[3][0]));
  item.setHelpText(valores[4][0]);
}

function MultipleChoiceOtherItem(valores, form){
  var item = form.addMultipleChoiceItem();
  var lista = valores[7][0].split(";");
  item.setTitle(valores[1][0]);
  item.setRequired(true);
  item.setChoiceValues(lista);
  item.setHelpText(valores[4][0]);
  item.showOtherOption(true);
}

function MultipleChoiceItem(valores, form){
  var item = form.addMultipleChoiceItem();
  var lista = valores[7][0].split(";");
  item.setTitle(valores[1][0]);
  item.setRequired(true);
  item.setChoiceValues(lista);
  item.setHelpText(valores[4][0]);
}

function CheckboxOtherItem(valores, form){
  var item = form.addCheckboxItem();
  var lista = valores[7][0].split(";");
  item.setTitle(valores[1][0]);
  item.setRequired(true);
  item.setHelpText(valores[4][0]);
  item.setChoiceValues(lista);
  item.showOtherOption(true);
}

function CheckboxItem(valores, form){
  var item = form.addCheckboxItem();
  var lista = valores[7][0].split(";");
  item.setTitle(valores[1][0]);
  item.setHelpText(valores[4][0]);
  item.setRequired(true);
  item.setChoiceValues(lista);
}

function GridItem(valores, form){
  var item = form.addGridItem();
  var fila = valores[9][0].split(";");
  var columna = valores[10][0].split(";");
  item.setTitle(valores[1][0]);
  item.setRequired(true);
  item.setRows(fila);
  item.setColumns(columna);
  item.setHelpText(valores[4][0]);
}

//función que genera la carta de confirmación de inicio de trámite y la carpeta del usuario.
function inicio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaProceso = ss.getSheetByName("Respuestas");//some sheets of this gBook
  var hojaAdicional = ss.getSheetByName("Adicional");
  var hojaAuxiliar = ss.getSheetByName("Auxiliar");
  hojaAuxiliar.deleteRows(1, 2);//clearing the contents of the auxiliar gSheet
  var keyWordInicial = hojaProceso.getRange(hojaProceso.getLastRow(), 10).getCell(1, 1);
  keyWordInicial.copyTo(hojaAuxiliar.getRange(1, 1));
  var cell = hojaAuxiliar.getRange("A2");
  cell.setFormula('=SPLIT(A1; ",";true)');
  hojaAuxiliar.getRange("A2:E2").copyTo(hojaProceso.getRange(hojaProceso.getLastRow(), 14), {contentsOnly:true});
  hojaAuxiliar.deleteRows(1, 2);
  var codigoInicial = hojaAdicional.getRange("C2").getValue();
  var idAsistente = codigoInicial + "_" + (hojaProceso.getLastRow()-1);
  var mainFolder = "Sistema de Evaluación";//root folder
  var folders = DriveApp.getFoldersByName(mainFolder);//folder iterator for mainFolder
  var userFolder = (folders.hasNext()) ? folders.next() : DriveApp.createFolder(mainFolder);
  folders = DriveApp.getFoldersByName(idAsistente);//folder iterator for particularFolder 
  var folder = (folders.hasNext()) ? folders.next() : userFolder.createFolder(idAsistente); 
  var folderUserUrl = folder.getUrl();
  var folderUserId = folder.getId();
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);//sharing link for everyOne
  hojaProceso.getRange(hojaProceso.getLastRow(), 12).getCell(1,1).setValue(idAsistente);
  hojaProceso.getRange(hojaProceso.getLastRow(), 13).getCell(1, 1).setValue(folderUserUrl);
  var dataRangeDatos = hojaProceso.getRange(hojaProceso.getLastRow(), 1, 1, hojaProceso.getLastColumn());
  var dataProceso = dataRangeDatos.getValues();
  for (var i = 0; i < dataProceso.length; ++i) {
    var column = dataProceso[i];
    var submitDate = column[0];
    var submitDateFormatted = Utilities.formatDate(submitDate, "GMT", "dd/MM/yyyy");
    var nameUser = column[1];
    var identificationTypeUser = column[2];
    var identificationNumberUser = column[3];
    var emailUser = column[4];
    var tutorUser = column[5];
    var tutorEmail = column[6];
    var thesisName = column[7];
    var thesisAbstract = column[8];
    var keyWords;
    var careerUser = column[10];
    var codeUser = column[11];
    var driveUser = column[12];
    var key1 = column[13];
    var key2 = column[14];
    var key3 = column[15];
    var key4 = column[16];
    var key5 = column[17];
    if (!key2 && !key3 && !key4 && !key5) {
      keyWords = key1;
    } else if (!key3 && !key4 && !key5) {
      keyWords = key1 + ",\n                        " + key2;
    } else if (!key4 && !key5) {
      keyWords = key1 + ",\n                        " + key2 + ",\n                        " + key3;
    } else if (!key5) {
      keyWords = key1 + ",\n                        " + key2 + ",\n                        " + key3 + ",\n                        " + key4;
    } else {
      keyWords = key1 + ",\n                        " + key2 + ",\n                        " + key3 + ",\n                        " + key4 + ",\n                        " + key5;
    }
    var nombreCartaConfirmacion = ("Carta inicio de proceso de " + nameUser + ".");
    var confirmacionId = DriveApp.getFileById('PUT_YOUR_GDOC_ID').makeCopy(nombreCartaConfirmacion).getId();//confirmation letter
    var confirmacion = DocumentApp.openById(confirmacionId);
    var cuerpoConfirmacion = confirmacion.getActiveSection();
    cuerpoConfirmacion.replaceText("%submitDateFormatted%", submitDateFormatted);
    cuerpoConfirmacion.replaceText("%nameUser%", nameUser);
    cuerpoConfirmacion.replaceText("%careerUser%", careerUser);
    cuerpoConfirmacion.replaceText("%thesisName%", thesisName);
    cuerpoConfirmacion.replaceText("%thesisAbstract%", thesisAbstract);
    cuerpoConfirmacion.replaceText("%keyWords%", keyWords);
    cuerpoConfirmacion.replaceText("%tutorUser%", tutorUser);
    cuerpoConfirmacion.replaceText("%tutorEmail%", tutorEmail);
    cuerpoConfirmacion.replaceText("%codeUser%", codeUser);
    confirmacion.saveAndClose();
    var pdfConfirmacion = DriveApp.createFile(confirmacion.getAs("application/pdf"));
    var pdfConfirmacionId = pdfConfirmacion.getId();
    var cartaConfirmacion = DriveApp.getFileById(pdfConfirmacionId).makeCopy(pdfConfirmacion.getName(), folder);
    cartaConfirmacion.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var cartaConfirmacionURL = cartaConfirmacion.getUrl();
    DriveApp.getFileById(pdfConfirmacionId).setTrashed(true);
    DriveApp.getFileById(confirmacionId).setTrashed(true);
  }
  var inicioSubjectUser = ("Inicio del proceso de finalización de estudios de " + nameUser);
  var inicioBodyUser = ("Respetad@ " + nameUser + ", reciba un cordial saludo de parte de PISA.\nEn la dirección\n" + cartaConfirmacionURL + "\npuede descargar la constancia de inicio de su trámite.");
  MailApp.sendEmail(emailUser, inicioSubjectUser, inicioBodyUser);
}


