function doGet(e) {
  var output = HtmlService.createTemplateFromFile('index.html').evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  var email = Session.getActiveUser().getEmail();
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('user',email);
  scriptProperties.setProperty('test',"hjflyljy");
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');  
  return output;

}
function getUserName(formObject){
  var email = Session.getActiveUser().getEmail();
  return email
}
function uploadFiles(formObject) {
  var sheet = SpreadsheetApp.openById("1ULyuxW6w34asVIr-PeQh8vaMst1UocnvYTmmSnrRa24").getSheetByName('Log');
  var fila = sheet.getLastRow()
  var fileUrl = "";
  var thumbnail = "https://drive.google.com/thumbnail?id="
  try {
    var folder = DriveApp.getFolderById("1oR1oA-StMAPuqH-aIITf9qXQ0AqYb9WG");
    var emp = formObject.emp
    
    //Upload file if exists and update the file url
    if (formObject.myFile.length > 0) {
      var blob = formObject.myFile;
      var file = folder.createFile(blob);
      fileUrl = file.getUrl();
      var id = file.getId()
      file.setName(formObject.fecha)
      fileName = file.getName();
      var fila = sheet.getLastRow()
      var lote = keyToLot(fila+1)
      // var emp = formObject.fecha
      sheet.appendRow([fila+1,formObject.fecha,formObject.nota,fileUrl,thumbnail+id,formObject.cantidad,formObject.estilo,formObject.fermentador,"Activo",formObject.color])
      savePasoV(formObject.fecha,formObject.fermentador,"LLENAR FERMENTADOR",formObject.nota,emp,"",formObject.cantidad,lote,formObject.estilo)
      editFermentador(vesselToKey(formObject.fermentador),formObject.fecha,formObject.nota,"Activo","LLENADO",formObject.cantidad,lote,formObject.estilo)
      // sheet = SpreadsheetApp.openById("1gI7FaJxPMtfHrIqfZL3KS6Ursc3RxbjzUbz9hI5wpQw").getSheetByName('Log');
      // sheet.getRange("B"+fila).setValue(fecha)
      // sheet.getRange("C"+fila).setValue(nota)
      // sheet.getRange("H"+fila).setValue(status)
      // sheet.getRange("L"+fila).setValue(sts)
      // sheet.getRange("F"+fila).setValue(litros)
      // sheet.getRange("I"+fila).setValue(lote)
      // sheet.getRange("K"+fila).setValue(sts2)
      // sheet.appendRow([fila+1,
      //   formObject.fecha,
      //   formObject.nota,
      //   fileUrl,
      //   "https://drive.google.com/thumbnail?id="+id,
      //   formObject.peso,
      //   formObject.material,
      //   "Activo"],
      //   formObject.color);
    } else{
      fileUrl = "Record saved without a file";
      var fila = sheet.getLastRow()
      var lote = keyToLot(fila+1)
      sheet.appendRow([fila+1,formObject.fecha,formObject.nota,fileUrl,"",formObject.cantidad,formObject.estilo,formObject.fermentador,"Activo",formObject.color])
      savePasoV(formObject.fecha,formObject.fermentador,"LLENAR FERMENTADOR",formObject.nota,emp,"",formObject.cantidad,lote,formObject.estilo)
      editFermentador(vesselToKey(formObject.fermentador),formObject.fecha,formObject.nota,"Activo","LLENADO",formObject.cantidad,lote,formObject.estilo)
      // sheet.appendRow([fila+1,
      //   formObject.fecha,
      //   formObject.nota,
      //   fileUrl,
      //   "https://drive.google.com/thumbnail?id="+id,
      //   formObject.peso,
      //   formObject.material,
      //   "Activo"],
      //   formObject.color);
    }
    //Saving records to Google Sheet
    
    // Return the URL of the saved file
    return fileUrl;
    
  } catch (error) {
    return error.toString();
  }
}
function vesselToKey(fermentador){
  var categoria
  if (fermentador.includes("FV")){categoria="FV"} else {categoria="BT"}
  fermentador = fermentador.replace(categoria,"")
  fermentador = parseInt(fermentador)
  // alert(fermentador)
  return fermentador
}
function keyToLot(key){
  var code =  "LOTE"
  switch(key.toString().length) {
    case 1:
      code = "00" + key
      break;
    case 2:
      code = "0" + key.toString()
      break;
    case 3:
      code = "" + key
      break;
    case 4:
      code = "" + key
      break;
    case 5:
      code = "" + key
      break;
    case 6:
      code = key
      break;
  }  
  code = "LOTE" + code
  return code
}
function crearEmpleado(formObject) {
  var sheet = SpreadsheetApp.openById("1Fh9rTnwggwJvrJ3xSn-qGsv9II-cawhkspMuOUquft8").getSheetByName('Log');

  try {
    var folder = DriveApp.getFolderById("1pORFb1Fnyd187Ww-YoKZR03mJv520HFC");
    var fileUrl = "";
    var fileName = "";

    //Upload file if exists and update the file url
    if (formObject.myFileEmp.length > 0) {
      var auth = 1
      if (formObject.puesto=="CARDERO"){auth=1} else 
      if (formObject.puesto=="BATIENTERO"){auth=1} else 
      if (formObject.puesto=="OPENERO"){auth=1} else 
      if (formObject.puesto=="MECANICO"){auth=1} else 
      if (formObject.puesto=="ADMIN"){auth=4} 
      var blob = formObject.myFileEmp;
      var file = folder.createFile(blob);
      fileUrl = file.getUrl();
      var id = file.getId()
      file.setName(formObject.nombre+"_"+formObject.apellidos)
      fileName = file.getName();
      var fila = sheet.getLastRow()
      sheet.appendRow([fila+1,
        formObject.fecha,
        formObject.nombre.toUpperCase()+" "+formObject.apellidos.toUpperCase(),
        formObject.puesto,
        fileUrl,
        "https://drive.google.com/thumbnail?id="+id,
        "Activo",auth]);
    } else{
      fileUrl = "Record saved without a file";
      var fila = sheet.getLastRow()
      sheet.appendRow([fila+1,
        formObject.fecha,
        formObject.nombre.toUpperCase()+" "+formObject.apellidos.toUpperCase(),
        formObject.puesto,
        fileUrl,
        "https://drive.google.com/thumbnail?id="+id,
        "Activo",auth]);
    }

    //Saving records to Google Sheet
    
    // Return the URL of the saved file
    return fileUrl;
    
  } catch (error) {
    return error.toString();
  }
}
function editPaka(formObject) {
  var sheet = SpreadsheetApp.openById("1XJjKyHv_U8tBkdxer5EpUL8_EcHk7plyW5M6kYCe9e0").getSheetByName('Log');

  try {
    var folder = DriveApp.getFolderById("1pORFb1Fnyd187Ww-YoKZR03mJv520HFC");
    var fileUrl = "";
    var fileName = "";

    //Upload file if exists and update the file url
    if (formObject.archivo.length > 0) {
      var blob = formObject.archivo;
      var file = folder.createFile(blob);
      fileUrl = file.getUrl();
      var fileId = file.getId()
      fileName = file.getName();
      var fila = formObject.keyEdit
      sheet.getRange("C"+fila).setValue(formObject.nota)
      sheet.getRange("D"+fila).setValue(fileUrl)
      sheet.getRange("E"+fila).setValue("https://drive.google.com/thumbnail?id="+fileId)
      sheet.getRange("F"+fila).setValue(formObject.peso)
      sheet.getRange("G"+fila).setValue(formObject.material)
      sheet.getRange("I"+fila).setValue(formObject.color)
    } else {
      var fila = formObject.keyEdit
      sheet.getRange("C"+fila).setValue(formObject.nota)
      sheet.getRange("F"+fila).setValue(formObject.peso)
      sheet.getRange("G"+fila).setValue(formObject.material)
      sheet.getRange("I"+fila).setValue(formObject.color)
    }
    return fileUrl;
    
  } catch (error) {
    return error.toString();
  }
}
function editEmpleado(formObject) {
  var sheet = SpreadsheetApp.openById("1Fh9rTnwggwJvrJ3xSn-qGsv9II-cawhkspMuOUquft8").getSheetByName('Log');

  try {
    var folder = DriveApp.getFolderById("1pORFb1Fnyd187Ww-YoKZR03mJv520HFC");
    var fileUrl = "";

    //Upload file if exists and update the file url
    if (formObject.archivo.length > 0) {
      var blob = formObject.archivo;
      var file = folder.createFile(blob);
      var fileName = "";
      fileUrl = file.getUrl();
      var fileId = file.getId()
      fileName = file.getName();
      var fila = formObject.keyEdit
      sheet.getRange("C"+fila).setValue(formObject.nombre)
      sheet.getRange("D"+fila).setValue(formObject.puesto)
      sheet.getRange("E"+fila).setValue(fileUrl)
      sheet.getRange("F"+fila).setValue("https://drive.google.com/thumbnail?id="+fileId)
    } else {
      var fila = formObject.keyEdit
      sheet.getRange("C"+fila).setValue(formObject.nombre)
      sheet.getRange("D"+fila).setValue(formObject.puesto)
    }
    return fileUrl;
    
  } catch (error) {
    return error.toString();
  }
}
function editFermentador(fila,fecha,nota,status,sts,litros,lote,sts2) {
  var sheet = SpreadsheetApp.openById("1gI7FaJxPMtfHrIqfZL3KS6Ursc3RxbjzUbz9hI5wpQw").getSheetByName('Log');
  sheet.getRange("B"+fila).setValue(fecha)
  sheet.getRange("C"+fila).setValue(nota)
  sheet.getRange("H"+fila).setValue(status)
  sheet.getRange("L"+fila).setValue(sts)
  sheet.getRange("F"+fila).setValue(litros)
  sheet.getRange("I"+fila).setValue(lote)
  sheet.getRange("K"+fila).setValue(sts2)
}
function savePaso(fecha,paka,paso,peso,material,nota,emp,color,cobro,moneda,metodo,lote,fuenteLote,fuenteEstilo){
  var pasos = SpreadsheetApp.openById("1-zGHeNPOyt-4umBS7RIB-pl1IqLTvtPacaI411rf0dE").getSheetByName('Log')
  var filas = pasos.getLastRow()
  pasos.appendRow([filas+1,fecha,paka,paso,peso,material,nota,emp,"Activo",color,cobro,moneda,metodo,lote,fuenteLote,fuenteEstilo])
}
function savePasoV(fecha,fermentador,paso,nota,emp,keg,litros,lote,estilo){
  var pasos = SpreadsheetApp.openById("1l5mXjbLjBrajqku6WnG_XZ3nkddd9R-YOYwp-SSXo3s").getSheetByName('Log')
  var filas = pasos.getLastRow()
  pasos.appendRow([filas+1,fecha,fermentador,paso,nota,emp,"Activo",keg,litros,lote,estilo])
}
function updPasosKeg(dato,range){
  var sheet = SpreadsheetApp.openById("1-zGHeNPOyt-4umBS7RIB-pl1IqLTvtPacaI411rf0dE").getSheetByName('Log');
  sheet.getRange(range).setValue(dato)
}
function updPasosFermentador(dato,range){
  var sheet = SpreadsheetApp.openById("1l5mXjbLjBrajqku6WnG_XZ3nkddd9R-YOYwp-SSXo3s").getSheetByName('Log');
  sheet.getRange(range).setValue(dato)
}
function updKeg(dato,range){
  var sheet = SpreadsheetApp.openById("1XJjKyHv_U8tBkdxer5EpUL8_EcHk7plyW5M6kYCe9e0").getSheetByName('Log');
  sheet.getRange(range).setValue(dato)
}
function updVessel(dato,range){
  var sheet = SpreadsheetApp.openById("1gI7FaJxPMtfHrIqfZL3KS6Ursc3RxbjzUbz9hI5wpQw").getSheetByName('Log');
  sheet.getRange(range).setValue(dato)
}
function updEmp(dato,range){
  var sheet = SpreadsheetApp.openById("1Fh9rTnwggwJvrJ3xSn-qGsv9II-cawhkspMuOUquft8").getSheetByName('Log');
  sheet.getRange(range).setValue(dato)
}
function borrarPaka(){
  var sheet = SpreadsheetApp.openById("1e46r8fT5OGcdtToc6XvxcMUvnMsirC3ien89yySff1Y").getSheetByName('Log')
  sheet.getRange("A1").setValue("")
}
function clearPaka(){
  var sheet = SpreadsheetApp.openById("1e46r8fT5OGcdtToc6XvxcMUvnMsirC3ien89yySff1Y").getSheetByName('Log')
  sheet.getRange("A1").setValue("")
}
function clearEmpleado(){
  var sheet = SpreadsheetApp.openById("1soXANrgHHGQrFniwyEuzh1B0iaQV4GZRXxlzddN8zPg").getSheetByName('Log')
  sheet.getRange("A1").setValue("")
}
function generatePdf() {
  var remisionSheet = SpreadsheetApp.openById("18yEZwTPekOWTD4wLpVZT2IEpddoTQ-VXOR0ugjPA-0g").getSheetByName('Numero')
  var sourceSheet = SpreadsheetApp.openById("1rEPJzaNFaMYP8AciOqCu_d-TUpKT1fbBjnag3wqGW8k").getSheetByName('Remision')
  var numRemision = remisionSheet.getRange("A1").getValue()
  sourceSheet.getRange("H9").setValue(numRemision+1)
  var pdfName = "Remision"+numRemision+".pdf"; // Set the output filename as SheetName.
  var folder = DriveApp.getFolderById("1FnnAIKgWZna2UKNED9Ysip93uzMTHMzn")
  var theBlob = createblobpdf(pdfName);
  var newFile = folder.createFile(theBlob);
  var email = Session.getActiveUser().getEmail() || 'ing.adrian.echavarria@gmail.com';
  var custemail = sourceSheet.getRange('B15').getValue();
  email = custemail + ","  + email
  const subject = 'Pax Brewing Remision #'+(numRemision).toString();
  // const subject = `Your subject Attachement`;
  const body = "Anexa se encuentra la remisión. Cualquier duda favor de comunicarse con nosotros. Gracias por su compra ";
  GmailApp.sendEmail(email, subject, body, {
    htmlBody: body,
    attachments: [theBlob]
  });
  remisionSheet.getRange("A1").setValue(numRemision+1)
}

function createblobpdf(pdfName) {
  var sourceSheet = SpreadsheetApp.openById("1rEPJzaNFaMYP8AciOqCu_d-TUpKT1fbBjnag3wqGW8k").getSheetByName('Remision');
  var url = 'https://docs.google.com/spreadsheets/d/1rEPJzaNFaMYP8AciOqCu_d-TUpKT1fbBjnag3wqGW8k/export?exportFormat=pdf&format=pdf' // export as pdf / csv / xls / xlsx
    +    '&size=A4' // paper size legal / letter / A4
    +    '&portrait=true' // orientation, false for landscape
    +    '&fitw=true' // fit to page width, false for actual size
    +    '&sheetnames=true&printtitle=false' // hide optional headers and footers
    +    '&pagenum=RIGHT&gridlines=false' // hide page numbers and gridlines
    +    '&fzr=false' // do not repeat row headers (frozen rows) on each page
    +    '&horizontal_alignment=CENTER' //LEFT/CENTER/RIGHT
    +    '&vertical_alignment=TOP' //TOP/MIDDLE/BOTTOM
    +    '&gid=' + sourceSheet.getSheetId(); // the sheet's Id
  var token = ScriptApp.getOAuthToken();
  // request export url
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  var theBlob = response.getBlob().setName(pdfName);
  return theBlob;
};
function setClienteData(nombre,mail,bar,dire,litros,monto,keg,estilo,disc,moneda,metodo,amDue){
  var sourceSheet = SpreadsheetApp.openById("1rEPJzaNFaMYP8AciOqCu_d-TUpKT1fbBjnag3wqGW8k").getSheetByName('Remision');
  sourceSheet.getRange("B11").setValue(nombre)
  sourceSheet.getRange("B15").setValue(mail)
  sourceSheet.getRange("B10").setValue(bar)
  sourceSheet.getRange("B12").setValue(dire)
  sourceSheet.getRange("C35").setValue(keg)
  sourceSheet.getRange("B21").setValue(estilo+" "+litros+" L")
  sourceSheet.getRange("G4").setValue(monto)
  sourceSheet.getRange("G21").setValue(monto)
  sourceSheet.getRange("H21").setValue(monto)
  sourceSheet.getRange("H23").setValue(monto)
  sourceSheet.getRange("H25").setValue(monto)
  sourceSheet.getRange("B36").setValue(disc)
  sourceSheet.getRange("H3").setValue("("+moneda+")")
  sourceSheet.getRange("G25").setValue(amDue)
  sourceSheet.getRange("H13").setValue(metodo)

}
function setUserData(nombre,phone,mail){
  var sourceSheet = SpreadsheetApp.openById("1rEPJzaNFaMYP8AciOqCu_d-TUpKT1fbBjnag3wqGW8k").getSheetByName('Remision');
  sourceSheet.getRange("H48").setValue(nombre)
  sourceSheet.getRange("H49").setValue(phone)
  sourceSheet.getRange("H50").setValue(mail)
}


