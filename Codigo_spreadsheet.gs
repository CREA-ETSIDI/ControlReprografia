function ActualizarColores() {
  //Declaramos la variable sheet y la inicializamos con la hoja de las respuestas
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tarifas_y_colores");
  //Y obtenemos la longitud de la lista asumiendo que hay más colores que tarifas
  var size = sheet.getLastRow();

  //Ordenamos alfabéticamente los colorinchis (a petición de Jorge Lord de las impresoras 15/02/2021)
  sheet.getRange(5,1,size-4,2).sort(1);
  
  //Declaramos la variable form y la inicializamos con el formulario
  var form = FormApp.openById("1APSOoixKlKzkuuIS64RemlEOK9jtU7_Y6ynGFFQkfEk");
  //Almacenamos la ID de la pregunta a modificar ¡EN CASO DE MODIFICAR EL FORMULARIO HAY QUE CORREGIR ESTE VALOR!
  var QID = 1535240816;
  //Y guardamos en la variable item el objeto de la pregunta del formulario
  var item = form.getItemById(QID);
  
  var choices = ["Indiferente"];
  var j = 0;
  for(var i = 2; i <= size+1; i++)
  {
    if(sheet.getRange(i, 2).getValue())
    {
      choices[j]=sheet.getRange(i, 1).getValue();
      j++;
    }
  }
  for (j = 0; j < choices.lenght; j++)
  {
    Logger.log(choices[j]);
  }
  item.asListItem().setChoiceValues(choices)
}

function NuevaImpresora() {
  let actv = SpreadsheetApp.getActiveSpreadsheet();
  let hojas = actv.getSheets();
  //User proofing
  let nombre = hojas[1].getRange(9,11).getValue().toString();
  let claro = hojas[1].getRange(10,11).getValue().toString();
  let oscuro = hojas[1].getRange(11,11).getValue().toString();
  if(nombre.length > 0)
  {
    if(claro.length > 0)
    {
      if(oscuro.length > 0)
      {
          if(claro.indexOf("#") > -1)
          {
            if(oscuro.indexOf("#") > -1)
            {
              if(claro.length == 7)
              {
                if(oscuro.length == 7)
                {
                  //Todo bien, todo correcto y yo que me alegro

                  let formulaClaro = '=MOD(ROW(L2);2)*(L1="'+nombre+'")';
                  let formulaOscuro = '=MOD(ROW(L1);2)*(L1="'+nombre+'")';
                  let rango = hojas[0].getRange("L:L");
                  let reglaClaro = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formulaClaro).setBackground(claro).setRanges([rango]).build();
                  let reglaOscuro = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formulaOscuro).setBackground(oscuro).setRanges([rango]).build();

                  let reglas = hojas[0].getConditionalFormatRules();
                  reglas.push(reglaClaro);
                  reglas.push(reglaOscuro);
                  hojas[0].setConditionalFormatRules(reglas);

                  //Fin de creación de formato condicional

                  //Inicio de la limpieza

                  hojas[1].getRange(9,11).setValue("");
                  hojas[1].getRange(10,11).setValue("");
                  hojas[1].getRange(11,11).setValue("");
                }
                else if(oscuro.length < 7)
                {
                actv.toast('Al color oscuro le faltan caracteres',"Error: Datos incorrectos");
                }
                else if(oscuro.length > 7)
                {
                  actv.toast('Al color oscuro le sobran caracteres',"Error: Datos incorrectos");
                }
                else
                {
                  actv.toast('Color oscuro no válido',"Error: Datos incorrectos");
                }
              }
              else if(claro.length < 7)
              {
                actv.toast('Al color claro le faltan caracteres',"Error: Datos incorrectos");
              }
              else if(claro.length > 7)
              {
                actv.toast('Al color claro le sobran caracteres',"Error: Datos incorrectos");
              }
              else
              {
                actv.toast('Color claro no válido',"Error: Datos incorrectos");
              }
            }
            else
            {
              actv.toast('falta el "#" del color oscuro',"Error: Datos incorrectos");
            }
          }
          else
          {
            actv.toast('falta el "#" del color claro',"Error: Datos incorrectos");
          }
      }
      else
      {
        actv.toast("No has introducido un color oscuro para la nueva impresora","Error: Faltan datos");
      }
    }
    else
    {
    actv.toast("No has introducido un color claro para la nueva impresora","Error: Faltan datos");
    }
  }
  else
  {
    actv.toast("No has introducido un nombre para la nueva impresora","Error: Faltan datos");
  }
  ///hoja[0].getRange('L:L');
}

function Separador() {
  let actv = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = actv.getSheets()[0];
  let ultimaFila = hoja.getLastRow();

  hoja.getRange("A:W").setBorder(false,false,false,false,false,false);
  let i = 0;
  for(i = ultimaFila; (hoja.getRange(i,21).getValue() != 1) || i < 1; i--)
  {}
  hoja.getRange(i,1,1,23).setBorder(false,false,true,false,false,false,'#000000',SpreadsheetApp.BorderStyle.SOLID_THICK);
}

function GetItemsID() {
  var form = FormApp.openById("1APSOoixKlKzkuuIS64RemlEOK9jtU7_Y6ynGFFQkfEk");
  var items = form.getItems();
  for (var i in items)
  {
    Logger.log(i+' '+items[i].getTitle()+': '+items[i].getId());
  }
}

