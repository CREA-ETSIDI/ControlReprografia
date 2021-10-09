function Nuevalinea() {
  let sheet = SpreadsheetApp.openById("1thXqSrujMZXH_Qurpey84J2cHShauHY2K2XPUl6hAgU").getSheetByName("Respuestas de formulario 1");
  let fila = sheet.getLastRow();
  
  sheet.getRange(fila, 13).insertCheckboxes();
  sheet.getRange(fila, 17, 1, 4).insertCheckboxes();
  
  if(sheet.getRange(fila, 10) != "No")
  {
    ComprobarSocio(sheet.getRange(fila, 2).getValue(),sheet.getRange(fila, 4).getValue(),fila);
    Utilities.sleep(500);
    if((sheet.getRange(fila, 13).getValue() == false) && (sheet.getRange(fila, 10) == "Sí"))
    {
      EnviarMail(sheet.getRange(fila, 2).getValue());
    }
  }
  sheet.getRange(fila-1,21).activate();
  sheet.getActiveRange().autoFill(sheet.getRange(fila-1,21,2), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  sheet.getRange(fila-1,16).activate();
  sheet.getActiveRange().autoFill(sheet.getRange(fila-1,16,2), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  sheet.getRange(fila-1,22).activate();
  sheet.getActiveRange().autoFill(sheet.getRange(fila-1,22,2), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  sheet.getRange(fila-1,22,2).activate();
  
}

function ComprobarSocio(email,telefono,fila) {
  let sheet = SpreadsheetApp.openById("1thXqSrujMZXH_Qurpey84J2cHShauHY2K2XPUl6hAgU").getSheetByName("Respuestas de formulario 1");
  
  let lista = SpreadsheetApp.openById("14x0k5Xt6SJGJr0EQ-U2OwCC0zD-DUOVuQRyKG-RqEQ0").getSheetByName("Respuestas de formulario 1");
  let numero = lista.getLastRow();
  
  for(let i = 1; i <= numero; i++)
  {
    if(lista.getRange(i, 2).getValue()==email || lista.getRange(i, 8).getValue()==telefono)
    {
      sheet.getRange(fila, 13).setValue(true);
    }
  }
}

function EnviarMail(email) {
  GmailApp.sendEmail(email, "Socix no reconocido”, “En el formulario en el que has solicitado la impresión de tu pieza has marcado que eres socix del CREA, sin embargo nuestro robot no te ha reconocido como tal, si crees que se ha cometido un error, responde a este email");
}