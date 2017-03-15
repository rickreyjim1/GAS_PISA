// función que permite la generación de un menú adicional en la hoja de cálculo, con varios submenús
function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu('PISA_Eval')
  //.addItem('Crear formulario','leerEstructura')
  .addItem('Recibir Documento Original','tesisDownload')
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
function leerEstructura() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaGeneradora = ss.getSheetByName("GenForm");
  var lastColumn = hojaGeneradora.getLastColumn();
  var formTitle = hojaGeneradora.getRange("B24").getValue().toString();
  Logger.log(formTitle);
  var form;
  var formExist = hojaGeneradora.getRange("A25").getValue().toString();
  if(formExist=="urlFormWillBeHere" || formExist == ""){
    form = FormApp.create(formTitle);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  } else {
     form = FormApp.openByUrl(formExist); 
  }
  form.setTitle(formTitle);
  //form.setCollectEmails(true);
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

//Recibe la información, genera las columnas de las 5 palabras clave, genera carpeta de usuario
//genera la carta de confirmación de inicio de trámite y solicita el envío de adjunto como pdf.
function inicio() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaProceso = ss.getSheetByName("Respuestas");
  var hojaAdicional = ss.getSheetByName("Adicional");
  var hojaAuxiliar = ss.getSheetByName("Auxiliar");
  var keyWordInicial = hojaProceso.getRange(hojaProceso.getLastRow(), 24).getCell(1, 1).getValues();
  Logger.log(keyWordInicial);
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy");
  var studentPreviousCode = hojaProceso.getRange(hojaProceso.getLastRow(), 9).getCell(1,1).getValue();
  if (studentPreviousCode == "Anteproyecto/Proposal") {
    var studentPISACode = "PISA_Proposal_" + formattedDate + "_" + (hojaProceso.getLastRow()-1);
  }
  else if (studentPreviousCode == "Trabajo final de investigación o tesis/Thesis") {
    var studentPISACode = "PISA_Thesis_" + formattedDate + "_" + (hojaProceso.getLastRow()-1);
  }
  hojaProceso.getRange(hojaProceso.getLastRow(), 27).getCell(1,1).setValue(studentPISACode);
  Logger.log(studentPISACode);
  var studentPISACode;
  var mainFolder = "Sistema de Evaluación";
  var folders = DriveApp.getFoldersByName(mainFolder);
  var userFolder = (folders.hasNext()) ? folders.next() : DriveApp.createFolder(mainFolder);
  folders = DriveApp.getFoldersByName(studentPISACode);
  var folder = (folders.hasNext()) ? folders.next() : userFolder.createFolder(studentPISACode); 
  var folderUserUrl = folder.getUrl();
  var folderUserId = folder.getId();
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  hojaProceso.getRange(hojaProceso.getLastRow(), 28).getCell(1, 1).setValue(folderUserUrl);
  var dataRangeDatos = hojaProceso.getRange(hojaProceso.getLastRow(), 1, 1, hojaProceso.getLastColumn());
  var dataProceso = dataRangeDatos.getValues();
  for (var i = 0; i < dataProceso.length; ++i) {
    var column = dataProceso[i];
    var submitDate = column[0];
    var submitDateFormatted = Utilities.formatDate(submitDate, "GMT", "dd/MM/yyyy");
    var emailStudent = column[1];
    var nameStudent = column[2];
    var surnameStudent = column[3];
    var nationalityStudent = column[4];
    var identificationStudent = column[5];
    var identificationNumberStudent = column[6];
    var univalleStudent = column[7];
    var workType = column[8];
    var programStudent = column[9];
    var supervisor = column[10];
    var emailSupervisor = column[11];
    var otherSupervisors = column[12];
    var supervisor2 = column[13];
    var emailSupervisor2 = column[14];
    var advisor = column[15];
    var emailAdvisor = column[16];
    var title = column[21];
    var summary = column[22];
    var keyWords = column[23];
    var mainAreaStudent = column[24];
    var subAreaStudent = column[25];
    var studentPISACode = column[26];
    var studentDrive = column[27];
    /*var key1 = column[28];
    var key2 = column[29];
    var key3 = column[30];
    var key4 = column[31];
    var key5 = column[32];
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
    */var nombreCartaConfirmacion = ("Carta inicio de proceso de " + nameStudent + ".");
    if (otherSupervisors == "No tengo codirector ni asesor/None of the above") {
    var confirmacionId = DriveApp.getFileById('17J0uN1jRRiF3vC1fkOsyjvALpYddoyTLWMXCkjmDWZo').makeCopy(nombreCartaConfirmacion).getId();//crea la carta de confimación
    }
    else if (otherSupervisors == "Codirector y asesor/Supervisor 2 and advisor") {
    var confirmacionId = DriveApp.getFileById('1jwWXPYljTJy8vaQ9Oy3m2H4h8smAArs8Y3mA3yUnCLY').makeCopy(nombreCartaConfirmacion).getId();//crea la carta de confimación
    }
    else if (otherSupervisors == "Solo codirector/Just supervisor 2") {
    var confirmacionId = DriveApp.getFileById('15W5XOeXATW3o5a2UHwe07qY5VpgMjQiLb8JbKqv7sq8').makeCopy(nombreCartaConfirmacion).getId();//crea la carta de confimación
    }
    else if (otherSupervisors == "Solo asesor/Just advisor") {
    var confirmacionId = DriveApp.getFileById('1QJMc3TQAzhPpTxnP4V7a1og1X5Srz0zhC81lug7HOoQ').makeCopy(nombreCartaConfirmacion).getId();//crea la carta de confimación
    }
    var confirmacion = DocumentApp.openById(confirmacionId);
    var cuerpoConfirmacion = confirmacion.getActiveSection();
    cuerpoConfirmacion.replaceText("%submitDateFormatted%", submitDateFormatted);
    cuerpoConfirmacion.replaceText("%nameStudent%", nameStudent);
    cuerpoConfirmacion.replaceText("%surnameStudent%", surnameStudent);
    cuerpoConfirmacion.replaceText("%programStudent%", programStudent);
    cuerpoConfirmacion.replaceText("%title%", title);
    cuerpoConfirmacion.replaceText("%summary%", summary);
    cuerpoConfirmacion.replaceText("%keyWords%", keyWords);
    cuerpoConfirmacion.replaceText("%supervisor%", supervisor);
    cuerpoConfirmacion.replaceText("%emailSupervisor%", emailSupervisor);
    cuerpoConfirmacion.replaceText("%supervisor2%", supervisor2);
    cuerpoConfirmacion.replaceText("%emailSupervisor2%", emailSupervisor2);
    cuerpoConfirmacion.replaceText("%advisor%", advisor);
    cuerpoConfirmacion.replaceText("%emailAdvisor%", emailAdvisor);
    cuerpoConfirmacion.replaceText("%studentPISACode%", studentPISACode);
    confirmacion.saveAndClose();
    var pdfConfirmacion = DriveApp.createFile(confirmacion.getAs("application/pdf"));
    var pdfConfirmacionId = pdfConfirmacion.getId();
    var cartaConfirmacion = DriveApp.getFileById(pdfConfirmacionId).makeCopy(pdfConfirmacion.getName(), folder);
    cartaConfirmacion.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var cartaConfirmacionURL = cartaConfirmacion.getUrl();
    DriveApp.getFileById(pdfConfirmacionId).setTrashed(true);
    DriveApp.getFileById(confirmacionId).setTrashed(true);
  }
  //  var inicioSubjectUser = ("Inicio del proceso de finalización de estudios de " + nameStudent);
//  var inicioBodyUser = ("Respetad@ " + nameStudent + ", reciba un cordial saludo de parte de PISA.\nEn la dirección\n" + cartaConfirmacionURL + 
//                        "puede descargar la constancia de inicio de su trámite.\n\nPara continuar con su trámite, favor responder este correo adjuntando el documento completo, en formato PDF, sin cambiar el asunto." + 
//                        "\n\nMuchas gracias por su colaboración.");
//  MailApp.sendEmail(emailStudent, inicioSubjectUser, inicioBodyUser);
}

function tesisDownload(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaProceso = ss.getSheetByName("Respuestas");
  var dataRange = hojaProceso.getRange(1, 1, hojaProceso.getLastRow(), hojaProceso.getLastColumn());
  var data = dataRange.getValues();
  for (var i = 1; i < data.length; i++) {
    var column = data[i];
    var nameStudent = column[2];
    var emailStudent = column[1];
    var studentPISACode = column[26];
    var threads = GmailApp.search('is:unread subject:"Inicio del proceso de finalización de estudios de "' + nameStudent +' has:attachment from:' +emailStudent);
    for(var c = 0; c < threads.length; c++ ){
      var foldersUser = DriveApp.getFoldersByName(studentPISACode);
      var folderUserDrive;
      if(foldersUser.hasNext()){
        folderUserDrive = foldersUser.next();
      }
      else {
        folderUserDrive = DriveApp.createFolder(studentPISACode);
      }
      var emails = threads[c].getMessages();
      for(var e = 0; e < emails.length; e++){
        var adjuntos = emails[e].getAttachments();
        for(var a = 0; a < adjuntos.length; a++){
          folderUserDrive.createFile(adjuntos[a].copyBlob());
        }
      }
      threads[c].markRead();
     }
  }
}
