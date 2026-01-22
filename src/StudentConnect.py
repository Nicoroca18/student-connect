
function doPost(e) {
  const spreadsheet = SpreadsheetApp.getActive();
  const hojaControl = spreadsheet.getSheetByName("Control");
  const hojaValidaciones = spreadsheet.getSheetByName("Validacion Didit");
  const hojaForm = spreadsheet.getSheetByName("Input Form Web");

  const enProceso = hojaControl.getRange("B1").getValue();
  Logger.log("La var enProceso es igual a " + enProceso);

  if (enProceso != 1) {
    hojaControl.getRange("B1").setValue(1);
    hojaControl.getRange("C2").setValue(new Date());
    SpreadsheetApp.flush();

    let iVal = hojaValidaciones.getRange(1, 13).getValue();
    let totalValidaciones = hojaValidaciones.getRange(2, 13).getValue();

    while ((iVal - 1) < totalValidaciones) {
      try {
        const filaAactivar = hojaValidaciones.getRange(iVal, 10).getValue();
        const country = hojaValidaciones.getRange(iVal, 7).getValue();
        Logger.log('iVal = ' + iVal);
        Logger.log('Country = ' + country);

        if (country === "spain" || country === "france" || country === "uk" || country === "flexisim") {
          activarFilasWeb();
        } else {
          actman();
        }

        //hojaControl.getRange("B1").setValue(0);
        hojaControl.getRange("E2").setValue(new Date());
        SpreadsheetApp.flush();

        iVal = hojaValidaciones.getRange(1, 13).getValue();
        totalValidaciones = hojaValidaciones.getRange(2, 13).getValue();
      } catch (error) {
        Logger.log("Error capturado: " + error);
        //hojaControl.getRange("B1").setValue(0);
        hojaControl.getRange("E2").setValue(new Date());
        SpreadsheetApp.flush();

        /*
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: error.toString() })
        ).setMimeType(ContentService.MimeType.JSON);
        */
      }
    }

    hojaControl.getRange("B1").setValue(0);

    /*
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success' })
    ).setMimeType(ContentService.MimeType.JSON);
    */
  } else {
    Logger.log("Proceso ya en ejecución");
    /*
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'busy' })
    ).setMimeType(ContentService.MimeType.JSON);
    */
  }
}



function activarFilasWeb() {

  spreadsheet = SpreadsheetApp.getActive();
  var hojaForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Form Web");
  var hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  var hojaListados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listados");
  var hojaValidaciones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validacion Didit");
  var hojaPlanes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Planes");


  var iAct = hojaActivar.getRange("AC1").getValue();
  var iVal = hojaValidaciones.getRange(1,13).getValue();
  var totalValidaciones = hojaValidaciones.getRange(2,13).getValue();

  var hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  //var hoyFormateado = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy");

  //Quitar este while
  // while (iVal - 1 < totalValidaciones){

  var filaAactivar = hojaValidaciones.getRange(iVal,10).getValue();
    
  var accionfilaValidar = hojaValidaciones.getRange(iVal,2).getValue();
  var esPortabilidad = hojaForm.getRange(filaAactivar,20).getValue();

  hojaValidaciones.getRange(1,13).setValue(iVal+1);

  SpreadsheetApp.flush();
  Utilities.sleep(2000);

  if (esPortabilidad === "yes"){

    Logger.log('Es una Protabilidad');
    hojaValidaciones.getRange(iVal,11).setValue("Línea no activada porque pq es Portabilidad");

    avisarPortabilidad();

  }
  else{

    if (accionfilaValidar === "Approved" || accionfilaValidar === "Extension"){

      //var filaAactivar = hojaValidaciones.getRange(iVal,10).getValue();

      Logger.log('La accion de la fila iVal es igual a Approved = '+accionfilaValidar+'. La mandamos a activar. La fila a activar es '+filaAactivar);

      //Definimos variables de la hojaForm
      var timestamp = hojaForm.getRange(filaAactivar,1).getValue();
      var country = hojaForm.getRange(filaAactivar,2).getValue();
      var name = hojaForm.getRange(filaAactivar,3).getValue();
      var useremail = hojaForm.getRange(filaAactivar,4).getValue();
      

      //Fecha de activación
      const activationDateCell = hojaForm.getRange(filaAactivar, 5).getValue();

      let activationDateObj;
      if (typeof activationDateCell === 'string') {
        const partes = activationDateCell.split("-");
        activationDateObj = new Date(partes[0], partes[1] - 1, partes[2]);
      } else {
        activationDateObj = activationDateCell ? new Date(activationDateCell) : hoy;
      }
      // Limpiar horas para comparar solo fechas
      activationDateObj.setHours(0, 0, 0, 0);

      // Si fecha de activación es anterior a hoy, usamos hoy
      if (activationDateObj < hoy) {
        activationDateObj = hoy;
      }
      activationdateF = Utilities.formatDate(activationDateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");


      var sim_esim = hojaForm.getRange(filaAactivar,6).getValue();
      var icc = "89" + hojaForm.getRange(filaAactivar,7).getValue();
      var productaction = hojaForm.getRange(filaAactivar,8).getValue();
      var plantypeGB = hojaForm.getRange(filaAactivar,9).getValue();
      var plandurationmonths = hojaForm.getRange(filaAactivar,10).getValue();
      var referral = hojaForm.getRange(filaAactivar,11).getValue();
      var coupon = hojaForm.getRange(filaAactivar,12).getValue();
      var ref_phonenumber = hojaForm.getRange(filaAactivar,13).getValue();
      var idstripe = hojaForm.getRange(filaAactivar,14).getValue();
      var importepagadoTexto = hojaForm.getRange(filaAactivar,15).getValue();
      var importepagado = parseFloat(importepagadoTexto);
      var currentNumberText = hojaForm.getRange(filaAactivar,19).getValue();
      var currentNumber = Number(currentNumberText);
      var fechaPagoText = hojaForm.getRange(filaAactivar,16).getValue();
      var fechaPago = new Date(fechaPagoText);
      var recommended = hojaForm.getRange(filaAactivar,26).getValue();
      var currency = hojaForm.getRange(filaAactivar,27).getValue();


      //miramos si es Activacion o Extension

      if (productaction === "activation"){

        hojaActivar.getRange("AC1").setValue(iAct + 1);

        Logger.log('Es una activación');
        
        //Llenamos una nueva fila de la hoja activaciones
        hojaActivar.getRange(iAct,1).setValue(country);
        hojaActivar.getRange(iAct,2).setValue(icc);
        hojaActivar.getRange(iAct,4).setValue(plantypeGB);
        hojaActivar.getRange(iAct,5).setValue(sim_esim);
        hojaActivar.getRange(iAct,6).setValue(activationdateF);

        hojaActivar.getRange(iAct, 7).setFormula('=if(F' + iAct + '="";""; date(year(F' + iAct + ');month(F' + iAct + ')+H' + iAct + ';day(F' + iAct + ')))'); //Fecha Caducidad

        hojaActivar.getRange(iAct,8).setFormula("=VLOOKUP(D" + iAct + ";Listados!D:E;2;FALSE)");

        //hojaActivar.getRange(iAct,10).setFormula('=IF(I' + iAct + '="Activa";IF($AA$1-1<G' + iAct + ';"ok";if(I' + iAct + '="Suspendida";"ok";"suspender"));if(I' + iAct + '="Suspendida";if($AA$1-1<G' + iAct + ';"Activar";"Ok");""))');

        hojaActivar.getRange(iAct,10).setFormula('=IF(I' + iAct + '="Activa";IF($AA$1-1<G' + iAct + ';"ok";if(I' + iAct + '="Suspendida";"ok";"suspender"));if(I' + iAct + '="Suspendida";if($AA$1-1<G' + iAct + ';"Activar";if(and($AA$1>G' + iAct + '+6;A' + iAct + '="france");"Cancelar";"Ok"));""))');


        hojaActivar.getRange(iAct,11).setValue(name);
        hojaActivar.getRange(iAct,12).setValue(useremail);
        hojaActivar.getRange(iAct,13).setValue(idstripe);
        hojaActivar.getRange(iAct,14).setValue(fechaPago);
        hojaActivar.getRange(iAct,15).setValue(importepagado);


        if (currency === "USD"){
          var cambioUSDEUR = hojaPlanes.getRange(2,21).getValue();
          importepagado = importepagado*cambioUSDEUR;
        }
        else if (currency === "GBP"){
          var cambioGBPEUR = hojaPlanes.getRange(3,21).getValue();
          importepagado = importepagado*cambioGBPEUR;
        }

        //Aplicamos IVA SOLO a SPAIN
        if (country === "spain") {
          hojaActivar.getRange(iAct,16).setFormula("=O" + iAct + "/(1+VLOOKUP(A" + iAct + ";Planes!Q:R;2;FALSE))");
        }
        else{
          hojaActivar.getRange(iAct,16).setValue(importepagado);
        }

        // Guardar el valor original del referral
        var referralOriginal = referral;

        // Si hay cupón y no hay referral válido, buscarlo en la hoja Cupon-Referral
        if (coupon) {
          var hojaCupon = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cupon-Referral");
          var datosCupones = hojaCupon.getRange("A2:B" + hojaCupon.getLastRow()).getValues();

          var referralEncontrado = null;

          for (var j = 0; j < datosCupones.length; j++) {
            if (datosCupones[j][0] === coupon) {  // Columna A: cupones
              referralEncontrado = datosCupones[j][1];  // Columna B: referral
              Logger.log("El cupón es: " + coupon + " y el referral obtenido del cupón es: " + referralEncontrado);
              break;
            }
          }

          // Solo actualizar si se encontró un referral válido
          if (referralEncontrado && referralEncontrado !== "") {
            referral = referralEncontrado;
          } else {
            referral = referralOriginal; // Mantener el original si no se encontró nada
          }
        }

        hojaActivar.getRange(iAct,18).setValue(referral);
        hojaActivar.getRange(iAct,19).setValue(ref_phonenumber);
        hojaActivar.getRange(iAct,25).setValue(productaction);
        hojaActivar.getRange(iAct,31).setValue(recommended);

        //Si tiene numero de referal pone friend en referal
        var stringfriendreferal = String(ref_phonenumber).length;
        Logger.log("La longitud del friendreferal es "+stringfriendreferal);
    
        if (stringfriendreferal > 0) {
          hojaActivar.getRange(iAct,18).setValue("friend");
          Logger.log("Hemos puesto friend en la fila " +iAct);
        }

        SpreadsheetApp.flush();
        Utilities.sleep(5000); // esperamos 5 segundos

        
        if (sim_esim === "SIM"){
          Logger.log("es SIM");
          // Llamar función API activación SIM física, pedir que devuelva nº telefono (igual hay que llamar a otra api para esto)

          if (country === "spain" || country === "flexisim") {
            var respuestaActivacion = activarAPISIMSpain(iAct);
          }
          else if (country === "france") {
            if (activationDateObj > hoy) {
              hojaActivar.getRange(iAct, 9).setValue("PreActivada");
              var respuestaActivacion = "PreActivada";
            } else {
              var respuestaActivacion = activarAPISIMFrance(iAct);
            }
          }
          else if (country === "uk") {
            //SIM en UK lo hacemos Manual
            var respuestaActivacion = activarAPISIMuk();

          }

          // Si la activación con la API se ha hecho bien:
          if(respuestaActivacion === "ok"){
            // Para francia dejamos el Estado WIP
            if (country != "france") {
              hojaActivar.getRange(iAct,9).setValue("Activa"); //Ponemos el estado actual como Activa
            } else {
              Logger.log('Para francia dejamos el Estado WIP');
            }
          } else if (respuestaActivacion === "PreActivada") {
            hojaActivar.getRange(iAct,9).setValue("PreActivada"); //Ponemos el estado actual como PreActivada
          }
          else{
            // Llamar función enviar email error interno activacion SIM fisica
            hojaActivar.getRange(iAct,9).setValue("Error"); //Ponemos el estado actual como Error
            Logger.log('La respuesta de activación de la SIM fisica con icc = '+icc+' de la fila '+iAct+ ' NO es ok');
          }
        }

        else if (sim_esim === "eSIM"){
          Logger.log("es eSIM");
          // Llamar función API activación eSIM, pedir que devuelva nº telefono (igual hay que llamar a otra api para esto)
          // Llamar funcion API para pedir QR
          
          if (country === "spain" || country === "flexisim") {
            if (activationDateObj > hoy) {
              hojaActivar.getRange(iAct, 9).setValue("PreActivada");
              var respuestaActivacion = "PreActivada";
            } else {
              var respuestaActivacion = activarAPIeSIMSpainLIKES(iAct);
            }
          }
          else if (country === "france") {
            if (activationDateObj > hoy) {
              var respuestaActivacion = preActivarAPIeSIMFrance(iAct);
              //hojaActivar.getRange(iAct, 9).setValue("PreActivada");
            } else {
              var respuestaActivacion = activarAPIeSIMFrance(iAct);
            }
          }
    
          else if (country === "uk") {
            if (activationDateObj > hoy) {
              var respuestaActivacion = preActivarAPIeSIMuk(iAct);
              //hojaActivar.getRange(iAct, 9).setValue("PreActivada");
            } else {
              var respuestaActivacion = activarAPIeSIMuk(iAct);
            }
          }

          // Si la activación con la API se ha hecho bien:

          //DFM - Cuando Likes esté funcionando
          if (respuestaActivacion === "ok") {
            const esFrance = country === "france";
            const esSpainESIM = country === "spain" && sim_esim === "eSIM";
            if (esFrance || esSpainESIM) {
              Logger.log(`Para ${country} eSIM dejamos el estado en WIP`);
            } else {
              hojaActivar.getRange(iAct, 9).setValue("Activa"); // Columna I
            }
          } else if (respuestaActivacion === "PreActivada") {
            hojaActivar.getRange(iAct,9).setValue("PreActivada"); //Ponemos el estado actual como PreActivada
          }

          //if(respuestaActivacion === "ok"){
          //  // Para francia dejamos el Estado WIP
          //  if (country != "france") {
          //    hojaActivar.getRange(iAct,9).setValue("Activa"); //Ponemos el estado actual como Activa
          //  } else {
          //    Logger.log('Para francia dejamos el Estado WIP');
          //  }
          //  // Llamar funcion enviar email confirmación + QR
          //}
          else{
            //Llamar función enviar email error interno activacion eSIM
            Logger.log('La respuesta de activación de la eSIM con icc = '+icc+' de la fila '+iAct+ 'NO es ok');

            var destinatario = "studentconnect@connectivityglobal.com";
            var mensaje = 'Country = '+ country + '. La respuesta de activación de la eSIM con icc = '+icc+' de la fila '+iAct+ 'NO es ok';
            var asunto = "Error Activación "+ country;
            var copiaOculta = "sistemas@connectivity.es";

            enviarEmail(destinatario, copiaOculta, mensaje, asunto);
          }
        } 
      }

      else if (productaction === "extension"){

        Logger.log('Es una extension');
        Logger.log('icc = ' + icc);
        Logger.log('currentNumber = ' + currentNumber);
        
        // Mejorar rendimiento: obtener todos los números una vez
        var numerosAntiguos = hojaActivar.getRange(2, 3, iAct - 1).getValues(); // columna 3 desde fila 2 hasta iAct-1

        for (var i = numerosAntiguos.length - 1; i >= 0; i--) {
          var numeroAntiguo = numerosAntiguos[i][0];
          if (numeroAntiguo === currentNumber) {
            var filaEncontrada = i + 2; // ajuste porque empezamos en fila 2

            Logger.log('El numero a extender = '+currentNumber+' es igual al numeroAntiguo encontrado = '+numeroAntiguo+' de la fila = '+filaEncontrada);

            if (hojaActivar.getRange(filaEncontrada,9).getValue() === "Activa" || hojaActivar.getRange(filaEncontrada,9).getValue() === "Suspendida"){

              Logger.log('La fila a extender = '+filaEncontrada+' está activa. Procedemos a extender');

              hojaActivar.getRange("AC1").setValue(iAct + 1);
              hojaActivar.getRange(filaEncontrada,9).setValue('Extendida');

              var fechacaducidadextendida = hojaActivar.getRange(filaEncontrada,7).getValue();
              var partnerextendida = hojaActivar.getRange(filaEncontrada,17).getValue();
              var iccextendida = hojaActivar.getRange(filaEncontrada,2).getValue();
              var proveedor = hojaActivar.getRange(filaEncontrada,23).getValue();
              var referral = hojaActivar.getRange(filaEncontrada,18).getValue();

              hojaActivar.getRange(iAct,2).setValue(iccextendida);
              hojaActivar.getRange(iAct,3).setValue(currentNumber);
              hojaActivar.getRange(iAct,6).setValue(fechacaducidadextendida);
              hojaActivar.getRange(iAct,7).setFormula('=if(F' + iAct + '="";""; date(year(F' + iAct + ');month(F' + iAct + ')+H' + iAct + ';day(F' + iAct + ')))');
              hojaActivar.getRange(iAct,4).setValue(plantypeGB);
              hojaActivar.getRange(iAct,12).setValue(useremail);
              hojaActivar.getRange(iAct,5).setValue(sim_esim);
              //hojaActivar.getRange(iAct,10).setFormula('=IF(I' + iAct + '="Activa";IF($AA$1-1<G' + iAct + ';"ok";if(I' + iAct + '="Suspendida";"ok";"suspender"));if(I' + iAct + '="Suspendida";if($AA$1-1<G' + iAct + ';"Activar";"Ok");""))');
              hojaActivar.getRange(iAct,10).setFormula('=IF(I' + iAct + '="Activa";IF($AA$1-1<G' + iAct + ';"ok";if(I' + iAct + '="Suspendida";"ok";"suspender"));if(I' + iAct + '="Suspendida";if($AA$1-1<G' + iAct + ';"Activar";if(and($AA$1>G' + iAct + '+6;A' + iAct + '="france");"Cancelar";"Ok"));""))');
              hojaActivar.getRange(iAct,13).setValue(idstripe);
              hojaActivar.getRange(iAct,15).setValue(importepagado);
              hojaActivar.getRange(iAct,18).setValue(referral);
              hojaActivar.getRange(iAct,8).setFormula("=VLOOKUP(D" + iAct + ";Listados!D:E;2;FALSE)");
              hojaActivar.getRange(iAct,1).setValue(country);
              hojaActivar.getRange(iAct,17).setValue(partnerextendida);

              if (currency === "USD"){
                var cambioUSDEUR = hojaPlanes.getRange(2,21).getValue();
                importepagado = importepagado*cambioUSDEUR;
              }
              else if (currency === "GBP"){
                var cambioGBPEUR = hojaPlanes.getRange(3,21).getValue();
                importepagado = importepagado*cambioGBPEUR;
              }

              //Aplicamos IVA SOLO a SPAIN
              if (country === "spain") {
                hojaActivar.getRange(iAct,16).setFormula("=O" + iAct + "/(1+VLOOKUP(A" + iAct + ";Planes!Q:R;2;FALSE))");
              }
              else{
                hojaActivar.getRange(iAct,16).setValue(importepagado);
              }

              hojaActivar.getRange(iAct,19).setValue(ref_phonenumber);
              hojaActivar.getRange(iAct,14).setValue(fechaPago);
              hojaActivar.getRange(iAct,23).setValue(proveedor);
              hojaActivar.getRange(iAct,25).setValue(productaction);
              hojaActivar.getRange(iAct,31).setValue(recommended);

              var stringfriendreferal = String(ref_phonenumber).length;
              if (stringfriendreferal > 0) {
                hojaActivar.getRange(iAct,18).setValue("friend");
                Logger.log("Hemos puesto friend en la fila " +iAct);
              }

              if (country === "spain" && sim_esim === "SIM") {
                var lastTwoChars = plantypeGB.slice(-2);
                if (lastTwoChars == "-L") {
                  datosTarifa = 81920;
                } else if (lastTwoChars == "-M") {
                  datosTarifa = 40960;
                } else {
                  datosTarifa = 204800;
                }
                modificarDatos(iAct,datosTarifa);
              }

              if (hojaActivar.getRange(filaEncontrada,9).getValue() === "Suspendida") {
                if (country === "spain" && sim_esim === "SIM") {
                  var respuestaReactivacion = reactivarSIMSpain(currentNumber);
                  // respuestaReactivacion = "Activa" o "Error"
                } else if (country === "france"){
                  var respuestaReactivacion = reactivarSIMFrance(iAct);
                  // respuestaReactivacion = "WIP" o "Error"
                //} else if (country === "uk"){
                //  var respuestaReactivacion = reactivarSIMuk(iAct);
                } else {
                  var respuestaReactivacion = "Activa";
                }
              } else {
                var respuestaReactivacion = "Activa";
              }

              hojaActivar.getRange(iAct,9).setValue(respuestaReactivacion);

              // Si la activación con la API se ha hecho bien:
              if(respuestaReactivacion === "Error") {
                Logger.log('La linea a extender de la fila '+filaEncontrada+' no se ha podido extender por un error en la reactivación.');

                var destinatario = "studentconnect@connectivityglobal.com";
                var mensaje = 'Country = '+ country + '. No se ha podido hacer la extensión por un error en la reactivación. Mirar la fila ' + iAct + " de Activaciones en STUDENTCONNECT. En la hoja Validacion Didit, mirar la fila: " + iVal;
                var asunto = "Student Connect - Extensión Fallida " +country;
                var copiaOculta = "sistemas@connectivity.es";

                enviarEmail(destinatario, copiaOculta, mensaje, asunto);
              }



            } else {
              Logger.log('La linea a extender de la fila '+filaEncontrada+' no se ha podido extender. Su estado es '+hojaActivar.getRange(filaEncontrada,9).getValue());

              var destinatario = "studentconnect@connectivityglobal.com";
              var mensaje = 'Country = '+ country + '. No se ha podido hacer la extensión porque la linea no tiene estado Activa ni Suspendida. Mirar la fila ' + filaEncontrada + " de Activaciones en STUDENTCONNECT. En la hoja Validacion Didit, mirar la fila: " + iVal;
              var asunto = "Student Connect - Extensión Fallida " +country;
              var copiaOculta = "sistemas@connectivity.es";

              enviarEmail(destinatario, copiaOculta, mensaje, asunto);
            }

            break; // Salimos tras encontrar
          }
        }
      }

      else {
        Logger.log('ERROR: No es ni activation ni extension.')
        // Enviar email interno
      }

      var iAct = hojaActivar.getRange("AC1").getValue();
      var iVal = hojaValidaciones.getRange(1,13).getValue();
      var totalValidaciones = hojaValidaciones.getRange(2,13).getValue();

    }
    else{
      Logger.log('Accion validar no es Approved, no procedemos a activar. Enviamos mail a Ana/Luke');
      hojaValidaciones.getRange(iVal,11).setValue("Línea no activada pq accionvalidar no es approved");

      revisarValidacion();

      var iVal = hojaValidaciones.getRange(1,13).getValue();
      var iAct = hojaActivar.getRange("AC1").getValue();
    }
  }
}
    


function actman() {

  var hojaValidaciones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validacion Didit");
  var iVal = hojaValidaciones.getRange(1,13).getValue();
  var country = hojaValidaciones.getRange(iVal,7).getValue();

  Logger.log('funcion actman');

  //ENVIAR EMAIL A LUKE PARA HACER MANUAL
  var destinatario = "studentconnect@connectivityglobal.com";
  var mensaje = "Hay una nueva línea de " + country + " pendiente de activar en STUDENTCONNECT. En la hoja Validacion Didit, mirar la fila: " + iVal;
  var asunto = "Student Connect - Activación Manual " + country;
  var copiaOculta = "sistemas@connectivity.es";

  enviarEmail(destinatario, copiaOculta, mensaje, asunto);

  hojaValidaciones.getRange(1,13).setValue(iVal+1);
}


function revisarValidacion() {

  var hojaValidaciones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validacion Didit");
  var iVal = hojaValidaciones.getRange(1,13).getValue() - 1;

  //ENVIAR EMAIL A LUKE PARA HACER MANUAL
  var destinatario = "studentconnect@connectivityglobal.com";
  var mensaje = "Hay una nueva línea pendiente de activar en Student Connect. La validación NO ha sido aprovada. En la hoja Validacion Didit, mirar la fila: " + iVal;
  var asunto = "Student Connect - Revisar Validacion";
  var copiaOculta = "sistemas@connectivity.es";

  enviarEmail(destinatario, copiaOculta, mensaje, asunto);

  //hojaValidaciones.getRange(1,13).setValue(iVal+1);
}


function avisarPortabilidad() {

  var hojaValidaciones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validacion Didit");
  var iVal = hojaValidaciones.getRange(1,13).getValue() - 1;

  //ENVIAR EMAIL A LUKE PARA HACER MANUAL
  var destinatario = "studentconnect@connectivityglobal.com";
  var mensaje = 'Hay formulario de portabilidad en Student Connect de la fila: ' + iVal;
  var asunto = "Student Connect - Revisar Portabilidad";
  var copiaOculta = "sistemas@connectivity.es";

  enviarEmail(destinatario, copiaOculta, mensaje, asunto);

  hojaValidaciones.getRange(1,13).setValue(iVal+1);
}


function activarFilasWebMANUAL() {

  spreadsheet = SpreadsheetApp.getActive();
  var hojaForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Form Web");
  var hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  var hojaListados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listados");
  var hojaValidaciones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validacion Didit");
  var hojaControl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control");
  var hojaPlanes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Planes");

  var iAct = hojaActivar.getRange("AC1").getValue();
  var iVal = hojaValidaciones.getRange(1,13).getValue();
  var totalValidaciones = hojaValidaciones.getRange(2,13).getValue();

  
  //miramos si ya está activando
  var enProceso = hojaControl.getRange("B1").getValue();
  Logger.log("La var enProceso es igual a " + enProceso);

  if (enProceso != 1) {
    enProceso = 1; // Marcar que la función proceso está en ejecución      
    var horaActual = new Date();
    hojaControl.getRange("B1").setValue(enProceso);
    hojaControl.getRange("C2").setValue(horaActual);
    Logger.log("Ponemos enProceso igual a 1");
    SpreadsheetApp.flush();

    var hoy = new Date();
    //var hoyFormateado = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy");

    SpreadsheetApp.flush();
    Utilities.sleep(2000);

    var filaAactivar = hojaValidaciones.getRange("Q2").getValue();
    var numerotelefono = hojaValidaciones.getRange("Q3").getValue();

    hojaValidaciones.getRange("Q5").clearContent();

    if (!isNaN(filaAactivar) && !isNaN(numerotelefono) && filaAactivar > 0 && numerotelefono > 0){

      Logger.log('Activamos manualmente la fila ' + filaAactivar);

      //Definimos variables de la hojaForm
      var timestamp = hojaForm.getRange(filaAactivar,1).getValue();
      var country = hojaForm.getRange(filaAactivar,2).getValue();
      var name = hojaForm.getRange(filaAactivar,3).getValue();
      var useremail = hojaForm.getRange(filaAactivar,4).getValue();

      // Fecha de activación
      var activationdate = hojaForm.getRange(filaAactivar, 5).getValue();
      var activationdateF;

      if (activationdate) {
        var activationDateObj;

        if (typeof activationdate === 'string') {
          var partes = activationdate.split("-");
          activationDateObj = new Date(partes[0], partes[1] - 1, partes[2]);
        } else {
          activationDateObj = new Date(activationdate);
        }

        // Limpiar horas para comparar solo fechas
        activationDateObj.setHours(0, 0, 0, 0);
        hoy.setHours(0, 0, 0, 0);

        // Si la fecha de activación es anterior a hoy, usamos la fecha de hoy
        if (activationDateObj < hoy) {
          activationdateF = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else {
          activationdateF = Utilities.formatDate(activationDateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }

      } else {
        activationdateF = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }

      var sim_esim = hojaForm.getRange(filaAactivar,6).getValue();
      var icc = hojaForm.getRange(filaAactivar,7).getValue();
      var productaction = hojaForm.getRange(filaAactivar,8).getValue();
      var plantypeGB = hojaForm.getRange(filaAactivar,9).getValue();
      var plandurationmonths = hojaForm.getRange(filaAactivar,10).getValue();
      var referral = hojaForm.getRange(filaAactivar,11).getValue();
      var coupon = hojaForm.getRange(filaAactivar,12).getValue();
      var ref_phonenumber = hojaForm.getRange(filaAactivar,13).getValue();
      var idstripe = hojaForm.getRange(filaAactivar,14).getValue();
      var importepagadoTexto = hojaForm.getRange(filaAactivar,15).getValue();
      var importepagado = parseFloat(importepagadoTexto);
      var fechaPagoText = hojaForm.getRange(filaAactivar,16).getValue();
      var fechaPago = new Date(fechaPagoText);
      var recommended = hojaForm.getRange(filaAactivar,26).getValue();
      var currency = hojaForm.getRange(filaAactivar,27).getValue();

      //miramos si es Activacion o Extension

      if (productaction === "activation"){

        hojaActivar.getRange("AC1").setValue(iAct + 1);
        Logger.log('Es una activación');
        
        //Llenamos una nueva fila de la hoja activaciones
        hojaActivar.getRange(iAct,1).setValue(country);
        hojaActivar.getRange(iAct,2).setValue(icc);
        hojaActivar.getRange(iAct,4).setValue(plantypeGB);
        hojaActivar.getRange(iAct,5).setValue(sim_esim);
        hojaActivar.getRange(iAct,6).setValue(activationdateF);

        hojaActivar.getRange(iAct, 7).setFormula('=if(F' + iAct + '="";""; date(year(F' + iAct + ');month(F' + iAct + ')+H' + iAct + ';day(F' + iAct + ')))'); //Fecha Caducidad

        hojaActivar.getRange(iAct,8).setFormula("=VLOOKUP(D" + iAct + ";Listados!D:E;2;FALSE)");

        //hojaActivar.getRange(iAct,10).setFormula('=IF(I' + iAct + '="Activa";IF($AA$1-1<G' + iAct + ';"ok";if(I' + iAct + '="Suspendida";"ok";"suspender"));if(I' + iAct + '="Suspendida";if($AA$1-1<G' + iAct + ';"Activar";"Ok");""))');
        hojaActivar.getRange(iAct,10).setFormula('=IF(I' + iAct + '="Activa";IF($AA$1-1<G' + iAct + ';"ok";if(I' + iAct + '="Suspendida";"ok";"suspender"));if(I' + iAct + '="Suspendida";if($AA$1-1<G' + iAct + ';"Activar";if(and($AA$1>G' + iAct + '+6;A' + iAct + '="france");"Cancelar";"Ok"));""))');

        hojaActivar.getRange(iAct,11).setValue(name);
        hojaActivar.getRange(iAct,12).setValue(useremail);
        hojaActivar.getRange(iAct,13).setValue(idstripe);
        hojaActivar.getRange(iAct,14).setValue(fechaPago);
        hojaActivar.getRange(iAct,15).setValue(importepagadoTexto);

        if (currency === "USD"){
          var cambioUSDEUR = hojaPlanes.getRange(2,21).getValue();
          importepagado = importepagado*cambioUSDEUR;
        }
        else if (currency === "GBP"){
          var cambioGBPEUR = hojaPlanes.getRange(3,21).getValue();
          importepagado = importepagado*cambioGBPEUR;
        }

        //Aplicamos IVA SOLO a SPAIN
        if (country === "spain") {
          hojaActivar.getRange(iAct,16).setFormula("=O" + iAct + "/(1+VLOOKUP(A" + iAct + ";Planes!Q:R;2;FALSE))");
        }
        else{
          hojaActivar.getRange(iAct,16).setValue(importepagado);
        }

        // Guardar el valor original del referral
        var referralOriginal = referral;

        // Si hay cupón y no hay referral válido, buscarlo en la hoja Cupon-Referral
        if (coupon) {
          var hojaCupon = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cupon-Referral");
          var datosCupones = hojaCupon.getRange("A2:B" + hojaCupon.getLastRow()).getValues();

          var referralEncontrado = null;

          for (var j = 0; j < datosCupones.length; j++) {
            if (datosCupones[j][0] === coupon) {  // Columna A: cupones
              referralEncontrado = datosCupones[j][1];  // Columna B: referral
              Logger.log("El cupón es: " + coupon + " y el referral obtenido del cupón es: " + referralEncontrado);
              break;
            }
          }

          // Solo actualizar si se encontró un referral válido
          if (referralEncontrado && referralEncontrado !== "") {
            referral = referralEncontrado;
          } else {
            referral = referralOriginal; // Mantener el original si no se encontró nada
          }
        }

        hojaActivar.getRange(iAct,18).setValue(referral);
        hojaActivar.getRange(iAct,19).setValue(ref_phonenumber);
        hojaActivar.getRange(iAct,25).setValue(productaction);
        hojaActivar.getRange(iAct,31).setValue(recommended);

        //Si tiene numero de referal pone friend en referal
        var stringfriendreferal = String(ref_phonenumber).length;
        Logger.log("La longitud del friendreferal es "+stringfriendreferal);
    
        if (stringfriendreferal > 0) {
          hojaActivar.getRange(iAct,18).setValue("friend");
          Logger.log("Hemos puesto friend en la fila " +iAct);
        }
      
        hojaActivar.getRange(iAct,3).setValue(numerotelefono); //numero de telefono recibido de la API
        hojaActivar.getRange(iAct,9).setValue("Activa"); //Ponemos el estado actual como Activa
        // Llamar función enviar Email de confirmación

        hojaValidaciones.getRange("Q5").setValue("Fila activada correctamente. FilaAactivar = " +filaAactivar+" y numerotelefono = "+numerotelefono);
      }


      else if (productaction === "extension") {
        Logger.log('Es una extensión');

        // Cargamos toda la columna 3 (número teléfono) y columna 9 (estado) de la hoja Activaciones
        var datosActivaciones = hojaActivar.getRange(2, 1, iAct, hojaActivar.getLastColumn()).getValues();

        // Buscamos desde abajo la última fila que tenga ese número y esté activa o suspendida
        var filaExtendida = -1;
        for (var i = datosActivaciones.length - 1; i >= 0; i--) {
          var fila = datosActivaciones[i];
          var numeroAntiguo = fila[2]; // Columna 3 (índice 2)
          var estado = fila[8]; // Columna 9 (índice 8)

          if (numeroAntiguo === numerotelefono && (estado === "Activa" || estado === "Suspendida")) {
            filaExtendida = i + 2; // +2 porque datosActivaciones empieza en fila 2
            break;
          }
        }

        if (filaExtendida !== -1) {
          Logger.log('Extensión encontrada en fila ' + filaExtendida);
          hojaActivar.getRange("AC1").setValue(iAct + 1);

          hojaActivar.getRange(filaExtendida, 9).setValue('Extendida');
          var fechacaducidadextendida = hojaActivar.getRange(filaExtendida, 7).getValue();
          var partnerextendida = hojaActivar.getRange(filaExtendida, 17).getValue();
          var iccextendida = hojaActivar.getRange(filaExtendida, 2).getValue();
          var proveedor = hojaActivar.getRange(filaExtendida, 23).getValue();
          var referral = hojaActivar.getRange(filaEncontrada,18).getValue();

          hojaActivar.getRange(iAct, 2).setValue(iccextendida);
          hojaActivar.getRange(iAct, 6).setValue(fechacaducidadextendida);

          hojaActivar.getRange(iAct, 7).setFormula('=if(F' + iAct + '="";""; date(year(F' + iAct + ');month(F' + iAct + ')+H' + iAct + ';day(F' + iAct + ')))');

          hojaActivar.getRange(iAct, 4).setValue(plantypeGB);
          hojaActivar.getRange(iAct, 12).setValue(useremail);
          hojaActivar.getRange(iAct, 5).setValue(sim_esim);
          //hojaActivar.getRange(iAct, 10).setFormula('=IF(I' + iAct + '="Activa";IF($AA$1-1<G' + iAct + ';"ok";if(I' + iAct + '="Suspendida";"ok";"suspender"));if(I' + iAct + '="Suspendida";if($AA$1-1<G' + iAct + ';"Activar";"Ok");""))');
          hojaActivar.getRange(iAct,10).setFormula('=IF(I' + iAct + '="Activa";IF($AA$1-1<G' + iAct + ';"ok";if(I' + iAct + '="Suspendida";"ok";"suspender"));if(I' + iAct + '="Suspendida";if($AA$1-1<G' + iAct + ';"Activar";if(and($AA$1>G' + iAct + '+6;A' + iAct + '="france");"Cancelar";"Ok"));""))');
          hojaActivar.getRange(iAct, 13).setValue(idstripe);
          hojaActivar.getRange(iAct, 15).setValue(importepagado);
          hojaActivar.getRange(iAct, 18).setValue(referral);
          hojaActivar.getRange(iAct, 8).setFormula("=VLOOKUP(D" + iAct + ";Listados!D:E;2;FALSE)");
          hojaActivar.getRange(iAct, 1).setValue(country);
          hojaActivar.getRange(iAct, 17).setValue(partnerextendida);


          if (currency === "USD"){
            var cambioUSDEUR = hojaPlanes.getRange(2,21).getValue();
            importepagado = importepagado*cambioUSDEUR;
          }
          else if (currency === "GBP"){
            var cambioGBPEUR = hojaPlanes.getRange(3,21).getValue();
            importepagado = importepagado*cambioGBPEUR;
          }

          //Aplicamos IVA SOLO a SPAIN
          if (country === "spain") {
            hojaActivar.getRange(iAct,16).setFormula("=O" + iAct + "/(1+VLOOKUP(A" + iAct + ";Planes!Q:R;2;FALSE))");
          }
          else{
            hojaActivar.getRange(iAct,16).setValue(importepagado);
          }

          hojaActivar.getRange(iAct, 19).setValue(ref_phonenumber);
          hojaActivar.getRange(iAct, 14).setValue(fechaPago);
          hojaActivar.getRange(iAct,23).setValue(proveedor);
          hojaActivar.getRange(iAct,25).setValue(productaction);
          hojaActivar.getRange(iAct,31).setValue(recommended);

          var stringfriendreferal = String(ref_phonenumber).length;
          Logger.log("La longitud del friendreferal es " + stringfriendreferal);
          if (stringfriendreferal > 0) {
            hojaActivar.getRange(iAct, 18).setValue("friend");
            Logger.log("Hemos puesto friend en la fila " + iAct);
          }

          hojaActivar.getRange(iAct, 3).setValue(numerotelefono);
          hojaActivar.getRange(iAct, 9).setValue("Activa");

          hojaValidaciones.getRange("Q5").setValue("Fila extendida correctamente. FilaAactivar = " + filaAactivar + " y numerotelefono = " + numerotelefono);
        } else {
          Logger.log('No se encontró una línea activa o suspendida para extender.');
          hojaValidaciones.getRange("Q5").setValue("No se encontró una línea activa o suspendida para extender.");
        }
      }

      else {
        Logger.log('ERROR: No es ni activation ni extension.');
        hojaValidaciones.getRange("Q5").setValue("ERROR: No es ni activation ni extension.");
        // Enviar email interno
      }

    }
    else {
      Logger.log('Falta rellenar Q2 (fila input form web) o Q3 (numero telefono)');
      hojaValidaciones.getRange("Q5").setValue("ERROR: Falta rellenar Q2 (fila input form web) o Q3 (numero telefono).");
    }

    Utilities.sleep(10000);
    hojaValidaciones.getRange("Q2").clearContent();
    hojaValidaciones.getRange("Q3").clearContent();

    var iAct = hojaActivar.getRange("AC1").getValue();
    var iVal = hojaValidaciones.getRange(1,13).getValue();
    var totalValidaciones = hojaValidaciones.getRange(2,13).getValue();

    enProceso = 0;
    Logger.log("acaba el proceso y pone enproceso = 0");
    hojaControl.getRange("B1").setValue(enProceso);
    horaActual = new Date();
    hojaControl.getRange("E2").setValue(horaActual);
  }
  else {
    Logger.log('ERROR: En proceso = 1. Esperar a que termine el proceso en curso para realizar acción Manual');
    hojaValidaciones.getRange("Q5").setValue('ERROR: En proceso = 1. Esperar a que termine el proceso en curso para realizar acción Manual');
  }
}


function activarPreactivadas() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  const data = hoja.getDataRange().getValues();
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0); // limpiamos hora para comparar solo fecha

  for (let i = 1; i < data.length; i++) { // empezamos en 1 para saltar encabezados
    const estado = data[i][8]; // columna I (estado)
    const country = data[i][0].toLowerCase(); // columna A
    const sim_esim = data[i][4]; // columna E
    const fechaActivacion = new Date(data[i][5]); // columna F (fecha activación)
    fechaActivacion.setHours(0, 0, 0, 0); // limpiar hora

    if (estado === "PreActivada" && fechaActivacion <= hoy) {
      const fila = i + 1; // porque i empieza en 0, y data en fila 2

      if (country === "france" && sim_esim === "SIM") {
        activarAPISIMFrance(fila);
      } else if (country === "france" && sim_esim === "eSIM") {
        activarPreActAPIeSIMFrance(fila);
      } else if (country === "uk" && sim_esim === "eSIM") {
        activarPreActAPIeSIMuk(fila);
      } else if (country === "spain" && sim_esim === "eSIM") {
        activarAPIeSIMSpainLIKES(fila);
      }
    }
  }
}


function suspenderLineas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaActivar = ss.getSheetByName("Activaciones");
  var hojaPlanes = ss.getSheetByName("Planes");

  var hoy = new Date();
  var unaSemanaDespRaw = new Date(hoy);
  unaSemanaDespRaw.setDate(hoy.getDate() + 7);
  var unaSemanaDesp = Utilities.formatDate(unaSemanaDespRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");

  var dosDiasDespRaw = new Date(hoy);
  dosDiasDespRaw.setDate(hoy.getDate() + 2);
  var dosDiasDesp = Utilities.formatDate(dosDiasDespRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Leer datos una vez
  var datosActivar = hojaActivar.getDataRange().getValues();
  var datosPlanes = hojaPlanes.getDataRange().getValues();

  var numFilas = datosActivar.length;

  for (var i = 1; i < numFilas; i++) {  // Comienza en 1 para saltar encabezados
    var fila = datosActivar[i];
    var accion = fila[9];   // Columna J
    var estado = fila[8];   // Columna I
    var fechaRevRaw = fila[6]; // Columna G
    var country = fila[0];  // Columna A
    var proveedor = fila [22] // Columna W

    var fechaRev = null;
    // ✅ Validar fecha antes de formatear
    if (fechaRevRaw instanceof Date) {
      fechaRev = Utilities.formatDate(fechaRevRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } else {
      Logger.log(`⚠️ Fila ${i + 1}: fechaRevRaw inválida -> ${fechaRevRaw}`);
      // O podrías continuar con la siguiente fila si la fecha es esencial:
      continue;
    }

    if (accion === "suspender") {
      if (country.toLowerCase() === "spain") {
        suspenderSIMSpain(i + 1, hojaActivar);  // +1 porque en hoja empieza en fila 2
      } else if (country.toLowerCase() === "france" && proveedor === "Airmob") {
        suspenderSIMFrance(i + 1);
      } else if (country.toLowerCase() === "uk") {
        suspenderSIMuk(i + 1);
      }  else {
        var mensaje = `Country = ${country}. Hay una nueva línea pendiente de suspender en STUDENTCONNECT. En la hoja Activaciones, mirar la fila: ${i + 1}`;
        enviarEmail("studentconnect@connectivityglobal.com", "sistemas@connectivity.es", mensaje, "Student Connect - Suspensión manual " + country);
      }

    } else if (accion === "Cancelar") {
      if (country.toLowerCase() === "france" && proveedor === "Airmob") {
        cancelarSIMFrance(i + 1);
      } else {
        var mensaje = `Country = ${country}. Hay una nueva línea pendiente de cancelar en STUDENTCONNECT. En la hoja Activaciones, mirar la fila: ${i + 1}`;
        enviarEmail("studentconnect@connectivityglobal.com", "sistemas@connectivity.es", mensaje, "Student Connect - Baja manual " + country);
      }
    } else if (accion === "ok" && estado === "Activa") {
      // Variables comunes
      var name = fila[10];         // Columna K
      var numero = fila[2];        // Columna C
      var producto = fila[3];      // Columna D
      var simEsim = fila[4];       // Columna E
      var emailCliente = fila[11]; // Columna L
      var partesProducto = producto.split("-");
      var plan = partesProducto.length > 1 ? partesProducto[1] : "";

      // Buscar link
      var link = "";
      for (var j = 0; j < datosPlanes.length; j++) {
        var planFila = datosPlanes[j];
        if (planFila[0] == country && planFila[1] == simEsim && planFila[2] == plan) {
          link = planFila[3];
          break;
        }
      }

      if (fechaRev === unaSemanaDesp) {
        hojaActivar.getRange(i + 1, 20).setValue("mail aviso 7 días enviado " + hoy.toLocaleDateString());
        var mensaje = generarMensaje(name, numero, fechaRev, "in a week", link);
        enviarEmailHTML2(emailCliente, "soporte@connectivity.es", mensaje, "Your mobile line subscription is about to end next week");
        Logger.log(`Email 1 enviado correctamente a: `+emailCliente);

      } else if (fechaRev === dosDiasDesp) {
        hojaActivar.getRange(i + 1, 20).setValue("mail aviso 2 días enviado " + hoy.toLocaleDateString());
        var mensaje = generarMensaje(name, numero, fechaRev, "in 2 days", link);
        enviarEmailHTML2(emailCliente, "soporte@connectivity.es", mensaje, "Your mobile line subscription is about to end in 2 days");
        Logger.log(`Email 2 enviado correctamente a: `+emailCliente);
      }
    }
  }
}

function generarMensaje(name, numero, fechaRev, diasTexto, link) {
  return `
    <html>
      <body style="font-family: Arial, sans-serif; background-color: #fce5cd; padding: 30px;">
        <div style="background-color: #ffffff; max-width: 600px; margin: auto; padding: 25px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05);">
          <h2 style="color: #ff6d01;">Hello ${name}! 👋</h2>
          
          <p style="font-size: 16px; color: #333;">
            Your line <strong>${numero}</strong> is set to expire <strong>${diasTexto}</strong> on <strong>${fechaRev}</strong> 📅
          </p>

          <p style="font-size: 16px; color: #333;">
            If you'd like to keep enjoying your plan's benefits, you need to renew it through our website:
          </p>

          <p style="text-align: center; margin: 30px 0;">
            <a href="${link}" style="background-color: #ff6d01; color: #ffffff; text-decoration: none; padding: 12px 24px; border-radius: 8px; font-weight: bold;">
              Student Connect 🌐
            </a>
          </p>

          <p style="font-size: 16px; color: #cc0000;">
            You have 3 business days after your expiration to renew your line, otherwise <strong>you won't be able to keep the same phone number!</strong> ❌
          </p>

          <p style="font-size: 16px; color: #333;">
            If you have any questions, let us know! 
          </p>

          <br>

          <p style="font-size: 16px; color: #333;">
            Kind regards,<br><br>
            <strong>Student Connect Team</strong><br>
            📞 +34 681 91 63 22
          </p>
        </div>
      </body>
    </html>
  `;
}


function suspenderSIMSpain(i, hoja) {
  var hojaListados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listados");

  try {
    //ObtenerTokenFi();  // Solo si no la has obtenido antes

    const fila = hoja.getRange(i, 1, 1, 12).getValues()[0];
    const simEsim = fila[4];
    const numero = fila[2];
    const proveedor = fila[22];

    var tarifa = hoja.getRange(i,4).getValue();
    var lastTwoChars = tarifa.slice(-2);

    let filalistados = 1;


    if (simEsim === "SIM") {
      ObtenerTokenFi();  // Solo si no la has obtenido antes
      const id = obtenerIdSpainNico(numero);
      if (id !== "error") {
        const estado = suspenderSIMSpainAPI(id);
        if (estado === "Suspendida") {
          hoja.getRange(i, 9).setValue(estado);
          eliminarBancoDatos(i);
          
          //Devolvemos a listados
          if (lastTwoChars == "-L"){                                            //Si es tarifa grande ponemos el numero en la columna I
            while (hojaListados.getRange(filalistados,9).getValue() !== ""){
              filalistados++;
            }
            hojaListados.getRange(filalistados,9).setValue(numero);
          }
          else{                                                                 //Si es tarifa pequeña ponemos el numero en la columna H
            while (hojaListados.getRange(filalistados,8).getValue() !== ""){
              filalistados++;
            }
            hojaListados.getRange(filalistados,8).setValue(numero);
          }

        } else {
          hoja.getRange(i, 9).setValue("error API, no se ha suspendido");
          enviarAvisoManual("Spain", i);
        }
      } else {
        hoja.getRange(i, 9).setValue("error id, no se ha suspendido");
        enviarAvisoManual("Spain", i);
      }
    } else if (simEsim === "eSIM" && proveedor === "Likes") {
      const estado = suspendereSIMSpainLikes(numero, i);
      hoja.getRange(i, 9).setValue(estado);
    } else {
      enviarAvisoManual("Spain", i);
    }

  } catch (e) {
    hoja.getRange(i, 9).setValue("error crítico");
    enviarAvisoManual("Spain", i, e.message);
  }
}

function enviarAvisoManual(pais, fila, error = "") {
  const destinatario = "studentconnect@connectivityglobal.com";
  const copiaOculta = "sistemas@connectivity.es";
  const asunto = `Student Connect - Suspensión manual Country = ${pais}`;
  const mensaje = `Hay una línea pendiente de suspender manualmente en STUDENTCONNECT.\n\n` +
                  `Hoja: Activaciones\nFila: ${fila}\n${error ? "Error: " + error : ""}`;

  enviarEmail(destinatario, copiaOculta, mensaje, asunto);
}

//Antiguo - Antes del Proxy
/*
function suspendereSIMSpainLikes(numero, fila) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
    const token = obtenerTokenLikes();  // Llamada a la función que devuelve el JWT

    const url = "https://api.likestelecom.com/line/block";
    const body = {
      action: "BLOCK",
      lineNumber: numero
    };

    const options = {
      method: "put",
      contentType: "application/json",
      payload: JSON.stringify(body),
      headers: {
        "Authorization": "Bearer " + token
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (response.getResponseCode() === 200 && data.success === true) {
      hoja.getRange(fila, 9).setValue("Suspendida"); // columna I
      return "Suspendida";
    } else {
      Logger.log("❌ Error suspensión Likes: " + JSON.stringify(data));
      hoja.getRange(fila, 9).setValue("Error");
      hoja.getRange(fila, 20).setValue(data.message || "Error al suspender eSIM"); // columna T
      return "Error";
    }

  } catch (e) {
    Logger.log("⚠️ Excepción en suspendereSIMSpainLikes: " + e.message);
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
    hoja.getRange(fila, 9).setValue("Error");
    hoja.getRange(fila, 20).setValue(e.message); // columna T
    return "Error";
  }
}
*/

function suspendereSIMSpainLikes(numero, fila) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");

  try {
    if (!numero) {
      Logger.log("❌ No hay MSISDN para suspender en fila " + fila);
      hoja.getRange(fila, 9).setValue("Error");                 // Col I
      hoja.getRange(fila, 20).setValue("Sin número para suspender"); // Col T
      return "Error";
    }

    // 🔹 Ahora llamamos al PROXY, no directo a Likes
    const url = "http://207.154.252.56:3010/es/likes/block";

    const body = {
      action: "BLOCK",
      lineNumber: numero
    };

    const options = {
      method: "post",                          // 👈 POST al proxy
      contentType: "application/json",
      payload: JSON.stringify(body),
      muteHttpExceptions: true                // 👈 para loguear aunque haya error HTTP
    };

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const text = response.getContentText();
    Logger.log("🔎 Respuesta proxy Likes block (" + statusCode + "): " + text);

    let data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      Logger.log("❌ Error parseando JSON proxy block: " + e);
      hoja.getRange(fila, 9).setValue("Error");
      hoja.getRange(fila, 20).setValue("JSON inválido desde proxy Likes");
      return "Error";
    }

    // El proxy devuelve { status: 'ok', raw: { ...respuesta Likes... } }
    const raw = data.raw || {};
    const successRaw = raw.success === true;

    if (statusCode === 200 && (data.status === "ok" || successRaw)) {
      hoja.getRange(fila, 9).setValue("Suspendida");  // Col I
      hoja.getRange(fila, 20).setValue("Suspendida en Likes vía proxy");
      return "Suspendida";
    } else {
      const mensajeError = data.message || raw.message || text || "Error al suspender eSIM via proxy";
      Logger.log("❌ Error suspensión Likes via proxy: " + mensajeError);
      hoja.getRange(fila, 9).setValue("Error");
      hoja.getRange(fila, 20).setValue(mensajeError); // Col T
      return "Error";
    }

  } catch (e) {
    Logger.log("⚠️ Excepción en suspendereSIMSpainLikes (proxy): " + e.message);
    hoja.getRange(fila, 9).setValue("Error");
    hoja.getRange(fila, 20).setValue(e.message); // Col T
    return "Error";
  }
}



////////////////////////////////////////////////////////////////
//Pruebanicomails. Este email de aqui abajo es un ejemplo de como quedaría el mail de activaciones. No esta implementado porque el mail de activaciones se envia con la funcion enviaremail

function generarMensajeHTML(name, numero, fechaRev, diasTexto, link) {
  var name = "NICO";
  var numero ="123456789";
  var plan = "PLANPLAN";
  var fechaFin = "1234321"

  return `
<html>
  <body style="font-family: Arial, sans-serif; background-color: #fce5cd; padding: 30px;">
    <div style="background-color: #ffffff; max-width: 600px; margin: auto; padding: 25px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05);">
      
      <h2 style="color: #ff6d01;">Hi ${name}! </h2>

      <p style="font-size: 16px; color: #333;">
        Your <strong>Spain Connect</strong> activation is now complete. You can start enjoying your plan immediately!🎉
      </p>

      <p style="font-size: 16px; color: #333;">
        ✅ <strong>${plan}</strong>GB every month<br>
        📅 Until: <strong>${fechaFin}</strong><br>
        📱 New phone number: <strong>${numero}</strong>
      </p>

      <hr style="margin: 30px 0;">

      <h3 style="color: #ff6d01;">📶 No internet?</h3>
      <p style="font-size: 16px; color: #333;">
        Don’t worry — it’s usually due to the APN settings. Just follow these steps:
      </p>

      <p style="font-size: 16px; color: #333;"><strong>For Android:</strong></p>
      <ul style="font-size: 16px; color: #333; padding-left: 20px;">
        <li>Go to: <strong>Settings > Mobile Networks > Access Point Names (APN)</strong></li>
        <li>Create a new APN with the following settings:</li>
        <ul>
          <li><strong>Name:</strong> FI</li>
          <li><strong>APN:</strong> fi.omv.es</li>
          <li><strong>APN Type:</strong> default,supl,dun</li>
          <li><strong>MVNO Type:</strong> IMSI</li>
          <li><strong>MVNO Value:</strong> 2140606</li>
        </ul>
      </ul>

      <p style="font-size: 16px; color: #333;"><strong>For iPhone (iOS):</strong></p>
      <ul style="font-size: 16px; color: #333; padding-left: 20px;">
        <li>Go to: <strong>Settings > Mobile Data > Mobile Data Options</strong></li>
        <li>Configure the following:</li>
        <ul>
          <li><strong>Mobile Data</strong><br>
            APN: <code>fi.omv.es</code><br>
            Username: <em>(leave blank)</em><br>
            Password: <em>(leave blank)</em>
          </li>
          <li><strong>MMS</strong><br>
            APN: <code>fi.omv.es</code><br>
            Username: <em>(leave blank)</em><br>
            Password: <em>(leave blank)</em>
          </li>
          <li><strong>Internet Sharing</strong><br>
            APN: <code>fi.omv.es</code><br>
            Username: <em>(leave blank)</em><br>
            Password: <em>(leave blank)</em>
          </li>
          <li><strong>Optional LTE Field</strong><br>
            APN: <code>fi.omv.es</code><br>
            Username: <em>(leave blank)</em><br>
            Password: <em>(leave blank)</em>
          </li>
        </ul>
      </ul>

      <p style="font-size: 16px; color: #333;">
        🔁 Then, turn off your phone, wait 5 seconds, and turn it back on.<br>
        📲 Still not working? Try inserting the SIM into another phone to test connectivity.
      </p>

      <hr style="margin: 30px 0;">

      <p style="font-size: 16px; color: #333;">
        We're here if you need us!
      </p>

      <p style="text-align: center; margin: 30px 10px;">
        <a href="https://wa.me/34681916322" style="background-color: #25D366; color: #ffffff; text-decoration: none; padding: 12px 24px; border-radius: 8px; font-weight: bold; margin: 5px; display: inline-block;">
          Contact via WhatsApp 📲
        </a>
        <a href="mailto:spainconnect@connectivity.es" style="background-color: #ff6d01; color: #ffffff; text-decoration: none; padding: 12px 24px; border-radius: 8px; font-weight: bold; margin: 5px; display: inline-block;">
          Contact via Gmail 📧
        </a>
      </p>

      <p style="font-size: 16px; color: #333;">
        You'll receive more info about your plan in the next business day.
      </p>

      <br>
      <p style="font-size: 16px; color: #333;">
        Kind regards,<br><br>
        <strong>The Spain Connect Team</strong>
      </p>

    </div>
  </body>
</html>
`;
}


function enviarCorreoPrueba() {
  const name = "Nico Prueba";
  const numero = "1234567890";
  const fechaRev = new Date("2025-07-26");
  const diasTexto = "in a week";
  const link = "https://www.connectivityglobal.com/es/how-to-install-your-esim/";

  const mensajeHTML = generarMensajeHTML(name, numero, fechaRev, diasTexto, link);

  MailApp.sendEmail({
    to: "nicolas@connectivityglobal.com",  // ← pon aquí tu email para pruebas
    subject: "Your France Connect line is expiring soon 📅",
    htmlBody: mensajeHTML,
    from: "studentconnect@connectivityglobal.com", // <-- remitente personalizado
    name: "France Connect Team"
  });
}


//////////////////////////////////////////////////////////////////////////////

//Funcion nueva para enviar emails con HTML
function enviarEmailHTML2(destinatarioCliente, destinatarioOculto, mensaje, asunto) {

  MailApp.sendEmail({
    to: destinatarioCliente,  // ← pon aquí tu email para pruebas
    subject: asunto,
    htmlBody: mensaje,
    from: "studentconnect@connectivityglobal.com", // remitente personalizado
    bcc: destinatarioOculto,
    name: "Student Connect Team"
  });
}

/////////////////////////////////////////////////////////////////////////////

function emailfriends(){
  //HAY que distinguir por país?? Ahora solo distinguimos si tiene Friend
  
  var hojaAct = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");

  for (var i = 0; i<hojaAct.length; i++){
    var friendreferal = hojaAct.getRange(i,19).getValue();
    var stringfriendreferal = String(friendreferal).length;
    
    var fechaalta = hojaAct.getRange(i,6).getValue();
    var fechaalta7 = fechaalta + 7;

    var fechacaducidad = hojaAct.getRange(i,7).getValue();
    var fechacaducidadM5 = fechacaducidad - 5;
    
    var hoy = new Date();
    var hoyTexto = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy");

    const whatsappLink = "https://api.whatsapp.com/send?text=Hey%21%20I%27ve%20got%20an%20awesome%20deal%20for%20you%20%F0%9F%8E%89%21%20You%20can%20get%20a%20%E2%82%AC5%20discount%20on%20SIM%20and%20eSIM%20in%20Spain.%20Just%20use%20the%20code%20FRIENDS%20when%20you%20buy%20and%20introduce%20my%20spanish%20phone%20number%20in%20the%20purchase.%20Don%E2%80%99t%20miss%20out%21%20%F0%9F%9A%80%F0%9F%93%B1%20https%3A%2F%2Fwww.connectivityglobal.com%2Fproduct%2Fesim-unlimited%2F";

    var name = hojaAct.getRange(i,11).getValue();

    if (hoyTexto === fechaalta7){

      var mensaje = `
          <html>
            <body style="font-family: Arial, sans-serif; background-color: #fce5cd; padding: 30px;">
              <div style="background-color: #ffffff; max-width: 600px; margin: auto; padding: 25px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05);">
              
                <h2 style="color: #ff6d01;">Hello ${name}! 👋</h2>

                <p style="font-size: 16px; color: #333;">
                  Thanks for activating your Connectivity eSIM/SIM. As a token of appreciation, we'd love to introduce you to our <strong>FRIENDS Promo</strong> — a way to get <strong>money back</strong> just by sharing! 🧡
                </p>

                <p style="font-size: 16px; color: #333;">
                  Imagine using your SIM for <strong>free</strong> 🤩
                </p>

                <p style="font-size: 16px; color: #333;">
                  Every time a friend buys a Connectivity plan using your Spanish number and the code <strong>FRIENDS</strong>, you both earn:<br>
                <div style="max-width: 400px; margin: 0 auto; padding-left: 60px; display: flex; justify-content: center; gap: 20px; flex-wrap: wrap;">
                  <div style="background-color: #e6f4ea; border-radius: 10px; padding: 16px; text-align: center; width: 120px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                    <div style="font-size: 24px;">💰</div>
                    <div style="font-weight: bold; font-size: 14px; color: #1b5e20;">5€ for you</div>
                  </div>
                  <div style="background-color: #fff4e6; border-radius: 10px; padding: 16px; text-align: center; width: 120px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                    <div style="font-size: 24px;">🎁</div>
                    <div style="font-weight: bold; font-size: 14px; color: #e65100;">5€ off for them</div>
                  </div>
                </div>
                  <em>(Valid on plans of at least 4 months)</em>
                </p>

                <p style="font-size: 16px; color: #333;"><strong>Here’s how it works:</strong></p>

                <p style="font-size: 16px; color: #333;">⭐ Invite a friend to get a Connectivity SIM/eSIM for 4+ months</p>
                <p style="font-size: 16px; color: #333;">⭐ They use the code <strong>FRIENDS</strong> and enter your Spanish number at checkout</p>
                <p style="font-size: 16px; color: #333;">⭐ You earn 5€, and they save 5€</p>
                <p style="font-size: 16px; color: #333;">
                  You can earn up to the amount you paid for your own plan — for example, if your plan cost 25€, you can earn up to 25€ back! 🙌
                </p>

                <p style="font-size: 16px; color: #333;">
                  Ready to share?
                </p>

                <p style="text-align: center; margin: 30px 0;">
                  <a href="${whatsappLink}" style="background-color: #25D366; color: #ffffff; text-decoration: none; padding: 12px 24px; border-radius: 8px; font-weight: bold;">
                    Share via WhatsApp 📲
                  </a>
                </p>

                <p style="font-size: 16px; color: #333;">
                  Or simply forward this email to your friends 🚀
                </p>

                <br>

                <p style="font-size: 16px; color: #333;">
                  Kind regards,<br><br>
                  <strong>Spain Connect Team</strong>
                </p>
              </div>
            </body>
          </html>
        `;

      var destinatarioCliente = hojaAct.getRange(i,12).getValue();
      var destinatarioOculto = "";
      enviarEmailHTML2(destinatarioCliente, destinatarioOculto, mensaje, "Make your line free by bringing friends to Connectivity!");

      Logger.log("Email enviado 7 dias despues inicio. Línea " + i);
      Logger.log(`Email friends1 enviado correctamente a: `+destinatarioCliente);

    }
    else if (hoyTexto === fechacaducidadM5){
      var mensaje =
          `
      <html>
        <body style="font-family: Arial, sans-serif; background-color: #fce5cd; padding: 30px;">
          <div style="background-color: #ffffff; max-width: 600px; margin: auto; padding: 25px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05);">
          
            <h2 style="color: #ff6d01;">Hello ${name}! 👋</h2>

            <p style="font-size: 16px; color: #333;">
              You are leaving Spain and your Connectivity mobile line, but you can get back the money it cost you! We want to explain to you our <strong>FRIENDS promo</strong> 🤩
            </p>

            <p style="font-size: 16px; color: #333;">
              Sharing the coupon code <strong>"FRIENDS"</strong> with your friends coming to Spain, you will earn <strong>5€</strong> for each line they activate with Connectivity!
            </p>

            <p style="font-size: 16px; color: #333;">
              They will also get <strong>5€ off</strong> their purchase <em>(only valid for plans with a minimum duration of 4 months)</em>. Everyone wins!
            </p>

            <div style="max-width: 400px; margin: 0 auto; padding-left: 60px; display: flex; justify-content: center; gap: 20px; flex-wrap: wrap;">
              <div style="background-color: #e6f4ea; border-radius: 10px; padding: 16px; text-align: center; width: 120px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-size: 24px;">💰</div>
                <div style="font-weight: bold; font-size: 14px; color: #1b5e20;">5€ for you</div>
              </div>
              <div style="background-color: #fff4e6; border-radius: 10px; padding: 16px; text-align: center; width: 120px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div style="font-size: 24px;">🎁</div>
                <div style="font-weight: bold; font-size: 14px; color: #e65100;">5€ off for them</div>
              </div>
            </div>
            <br>

            <p style="font-size: 16px; color: #333;"><strong>How it works:</strong></p>

            <p style="font-size: 16px; color: #333;">⭐ Invite a friend to get a Connectivity SIM/eSIM for 4+ months</p>
            <p style="font-size: 16px; color: #333;">⭐ They use the code <strong>FRIENDS</strong> and enter your Spanish number at checkout</p>
            <p style="font-size: 16px; color: #333;">⭐ You earn 5€, and they save 5€</p>

            <p style="font-size: 16px; color: #333;">
              You can earn up to the amount you paid for your own plan — for example, if your plan cost 25€, you can earn up to 25€ back! 🙌
            </p>

            <p style="font-size: 16px; color: #333;">
              Ready to share?
            </p>

            <p style="text-align: center; margin: 30px 0;">
              <a href="${whatsappLink}" style="background-color: #25D366; color: #ffffff; text-decoration: none; padding: 12px 24px; border-radius: 8px; font-weight: bold;">
                Share via WhatsApp 📲
              </a>
            </p>

            <p style="font-size: 16px; color: #333;">
              Or simply forward this email to your friends 🚀
            </p>

            <br>

            <p style="font-size: 16px; color: #333;">
              Kind regards,<br><br>
              <strong>Spain Connect Team</strong>
            </p>
          </div>
        </body>
      </html>
      `;

      var destinatarioCliente = hojaAct.getRange(i,12).getValue();
      var destinatarioOculto = "soporte@connectivity.es";
      enviarEmailHTML2(destinatarioCliente, destinatarioOculto, mensaje, "Earn Back Your Connectivity Plan Cost!");

      Logger.log("Email enviado 5 dias antes fin. Línea " + i);
      Logger.log(`Email friends2 enviado correctamente a: `+destinatarioCliente);

    }
    
  }
}


//Funcion que se ejecute cada dia
function ejecutarSiUltimoDiaDelMes() {
  if (esUltimoDiaLaborable()) {
    Logger.log("Hoy es el último día laborable del mes, ejecutando función...");
    cancelacionmasiva();
    
  } else {
    Logger.log("Hoy no es el último día laborable del mes, no se ejecuta la función.");
  }
}


function esUltimoDiaLaborable() {
  var hoy = new Date();  // Fecha de hoy
  var ultimoDiaMes = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0);  // Último día del mes

  // Si el último día del mes es un sábado o domingo, retrocede hasta el último día laborable
  while (ultimoDiaMes.getDay() === 6 || ultimoDiaMes.getDay() === 0) {
    ultimoDiaMes.setDate(ultimoDiaMes.getDate() - 1);
  }

  // Compara la fecha de hoy con el último día laborable
  if (hoy.toDateString() === ultimoDiaMes.toDateString()) {
    return true;  // Hoy es el último día laborable del mes
  } else {
    return false;  // Hoy no es el último día laborable
  }
}


//*******************************************************************************************************************************************
//************     FUNCIONES GENERICAS          *********************************************************************************************
//*******************************************************************************************************************************************


function enviarEmailHTML(destinatario, copiaOculta, mensaje, asunto) {

  //envío con el alias

  // Dirección de correo electrónico del remitente
  //para saber si el mail que lanza el proceso
  var emailSesion = Session.getActiveUser().getEmail();
  Logger.log("El correo electrónico del usuario que ejecuta el script es: " + emailSesion);

  //if (emailSesion == "f.fabregas@call2world.es") {
  //  var remitente = "sistemas@connectivity.es";

  //} 
  //else if (emailSesion == "n.roca@call2world.es"){
  //  var remitente = "studentconnect@connectivityglobal.com";
  //}
  //else {
  //  var remitente = emailSesion;
  //}

  var remitente = "studentconnect@connectivityglobal.com"; // Fijo
  // var remitente = "sistemas@connectivity.es"; // Fijo

  // Configuración del correo electrónico
  var opciones = {
    htmlBody: mensaje,
    from: remitente,
    bcc: copiaOculta
  };

  // Envío del correo electrónico
  GmailApp.sendEmail(destinatario, asunto, mensaje, opciones);
}



function pruebaEmail() {     
  
  destinatario = "studentconnect@connectivityglobal.com";
  asunto = "Test";

  let mensaje =
  "Hi,\n\n" +
  "Your Spain Connect activation has been processed.\n\n" +
  "You can now enjoy 40Gb each month until XXXXX " + 
  ". Your new phone number is XXXXX " + ".\n\n" +
  "All the best,\n\n" +
  "Spain Connect team\n" +
  "spainconnect@connectivity.es\n\n" +

  "--------------------------------------------\n" +
  "TROUBLE SHOOTING\n" +
  "--------------------------------------------\n\n" +

  "If you don't get a data connection, it might be that your APN has not been set automatically. This is the most common issue when switching carriers, so don't worry. Just proceed as follows:\n\n" +

  "1. Check your APN settings.\n\n" +

  "SETTING FOR ANDROID:\n" +
  "- Go to: Settings > Mobile Networks > APN (Access Point Name).\n" +
  "- Create a new APN with the following settings:\n" +
  "  * Name: FI\n" +
  "  * APN: fi.omv.es\n" +
  "  * APN Type: default,supl,dun\n" +
  "  * MVNO Type: IMSI\n" +
  "  * MVNO Value: 2140606\n\n" +

  "SETTING FOR iOS (iPhone):\n" +
  "- Go to: Settings > Mobile Data > Mobile Data Options\n" +
  "- Configure the following:\n\n" +
  "MOBILE DATA field:\n" +
  "  * APN: fi.omv.es\n" +
  "  * Username: (leave it blank)\n" +
  "  * Password: (leave it blank)\n\n" +

  "MMS field:\n" +
  "  * APN: fi.omv.es\n" +
  "  * Username: (leave it blank)\n" +
  "  * Password: (leave it blank)\n\n" +

  "INTERNET SHARING field:\n" +
  "  * APN: fi.omv.es\n" +
  "  * Username: (leave it blank)\n" +
  "  * Password: (leave it blank)\n\n" +

  "Optional LTE field:\n" +
  "  * APN: fi.omv.es\n" +
  "  * Username: (leave it blank)\n" +
  "  * Password: (leave it blank)\n\n" +

  "1.2. Turn off the phone, wait 5 seconds, and turn it back on.\n\n" +
  "2. If the issue persists, try inserting the SIM card into another device to see if it picks up a signal.\n\n" +
  "3. If the issue continues, contact us at spainconnect@connectivity.es or send a WhatsApp message to +34 681 916 322.\n\n" +
  "Thank you for your patience!\n";


  copiaOculta = "d.ferreon@call2world.es";

  enviarEmail(destinatario, copiaOculta, mensaje, asunto);


}


function enviarEmail(destinatario, copiaoculta, mensaje, asunto) {
  // Dirección de correo electrónico del remitente
  //para saber si el mail que lanza el proceso
  var emailSesion = Session.getActiveUser().getEmail();
  Logger.log("El correo electrónico del usuario que ejecuta el script es: " + emailSesion);


  //if (emailSesion == "f.fabregas@call2world.es") {
  //  var remitente = "sistemas@connectivity.es";
  //
  //} 
  //else if (emailSesion == "n.roca@call2world.es"){
  //  var remitente = "studentconnect@connectivityglobal.com";
  //}
  //else {
  //  var remitente = emailSesion;
  //}
  ////envío con el alias
  //GmailApp.sendEmail(destinatario, asunto, mensaje, { 'from': remitente, "bcc": copiaoculta });
  
  //var remitente = "sistemas@connectivity.es"; // Fijo
  var remitente = "studentconnect@connectivityglobal.com"; // Fijo

  try {
    GmailApp.sendEmail(destinatario, asunto, mensaje, {
      from: remitente, // o quita esto si no está autorizado como alias
      bcc: copiaoculta
    });
    Logger.log('enviado OK');
  } catch (e) {
    Logger.log("Error al enviar: " + e.message);
  }

}


function buscarUltimaFila(columnaBuscar) {
  spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange(columnaBuscar + 1).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  var lastFila = SpreadsheetApp.getActiveRange().getRow();
  return lastFila
}


//*******************************************************************************************************************************************
//************       SPAIN          *********************************************************************************************************
//*******************************************************************************************************************************************


// DFM - hay que revisar que los campos esten bien 
function activarAPISIMSpain(iAct) {

  var hojaForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Form Web");
  var hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  var hojaListados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listados");

  var tarifa = hojaActivar.getRange(iAct,4).getValue();
  var lastTwoChars = tarifa.slice(-2);

  var numerocolumnalistados = 0;

  if (lastTwoChars == "-L") {
    // Número de 50GB ya que la tarifa es la L
    datosTarifa = 81920;
    var planGB = 80;
    hojaListados.activate();
    var n = buscarUltimaFila("B");
    hojaActivar.getRange(iAct, 3).setValue(hojaListados.getRange(n, 2).getValue());
    hojaListados.getRange(n, 2).setValue("");
    numerocolumnalistados = 2;
    
    // Envía mail si faltan líneas
    if (n < hojaListados.getRange("B2").getValue()) {
      mensaje = "En el listado spain Connect quedan " + n + "líneas de 50GB, hay que poner más";
      destinatarioCliente = "sistemas@call2world.es";
      destinatarioOculto = "studentconnect@connectivityglobal.com";
      enviarEmail(destinatarioCliente, destinatarioOculto, mensaje, "Aviso líneas disponibles Spain Connect de 50GB");
    }

  }else {
    // Número de 10GB
    if (lastTwoChars == "-M") {
      datosTarifa = 40960;
      var planGB = 40;
    } else {
      datosTarifa = 204800;
      var planGB = 200;
    }
    //datosTarifa = 40960;
    hojaListados.activate();
    var n = buscarUltimaFila("A");
    hojaActivar.getRange(iAct, 3).setValue(hojaListados.getRange(n, 1).getValue());
    hojaListados.getRange(n, 1).setValue("");
    numerocolumnalistados = 1;
    
    // Envía mail si faltan líneas
    if (n < hojaListados.getRange("A2").getValue()) {
      mensaje = "En el listado spain Connect quedan " + n + "líneas, hay que poner más";
      destinatarioCliente = "sistemas@call2world.es";
      destinatarioOculto = "studentconnect@connectivityglobal.com";
      enviarEmail(destinatarioCliente, destinatarioOculto, mensaje, "Aviso líneas disponibles Spain Connect");
    }
  }
  numeroLinea = hojaActivar.getRange(iAct, 3).getValue();
  iccnuevo = hojaActivar.getRange(iAct, 2).getValue();
  hojaActivar.activate();
  
  var respDuplicado = duplicarSIMSpain(numeroLinea, iccnuevo);
  
  if (respDuplicado == "ok") {
  
    // se mete en el banco de datos
    pasarDatos(iAct,datosTarifa);
  
    // El duplicado es correcto
    SpreadsheetApp.flush();
    Utilities.sleep(5000); // esperamos 5 segundos
    //esto es para intentar detectar el error Exception: The parameters (String,String,String) don't match the method signature for Utilities.formatDate.
    Logger.log("El duplicado ok y vamos a enviar mail")
    Logger.log("1er parametro = " + hojaActivar.getRange(iAct,4).getValue());
    Logger.log("2o parametro es = " + Session.getScriptTimeZone());
    Logger.log("3er parámetro es = " + 'dd/MM/yyyy');
    
        //Mensaje sin formato
    let mensaje =
    "Hi " + hojaActivar.getRange(iAct,11).getValue() + ",\n\n" +
    "Your Spain Connect activation has been processed.\n\n" +
    "You can now enjoy " + planGB + "GB each month until " + Utilities.formatDate(hojaActivar.getRange(iAct,7).getValue(), Session.getScriptTimeZone(), 'dd/MM/yyyy') + 
    ". Your new phone number is " + hojaActivar.getRange(iAct,3).getValue() + ".\n\n" +

    "--------------------------------------------\n" +
    "TROUBLE SHOOTING\n" +
    "--------------------------------------------\n\n" +

    "If you don't get a data connection, it might be that your APN has not been set automatically. This is the most common issue when switching carriers, so don't worry. Just proceed as follows:\n\n" +

    "1. Check your APN settings.\n\n" +

    "SETTING FOR ANDROID:\n" +
    "- Go to: Settings > Mobile Networks > APN (Access Point Name).\n" +
    "- Create a new APN with the following settings:\n" +
    "  * Name: FI\n" +
    "  * APN: fi.omv.es\n" +
    "  * APN Type: default,supl,dun\n" +
    "  * MVNO Type: IMSI\n" +
    "  * MVNO Value: 2140606\n\n" +

    "SETTING FOR iOS (iPhone):\n" +
    "- Go to: Settings > Mobile Data > Mobile Data Options\n" +
    "- Configure the following:\n\n" +
    "MOBILE DATA field:\n" +
    "  * APN: fi.omv.es\n" +
    "  * Username: (leave it blank)\n" +
    "  * Password: (leave it blank)\n\n" +

    "MMS field:\n" +
    "  * APN: fi.omv.es\n" +
    "  * Username: (leave it blank)\n" +
    "  * Password: (leave it blank)\n\n" +

    "INTERNET SHARING field:\n" +
    "  * APN: fi.omv.es\n" +
    "  * Username: (leave it blank)\n" +
    "  * Password: (leave it blank)\n\n" +

    "Optional LTE field:\n" +
    "  * APN: fi.omv.es\n" +
    "  * Username: (leave it blank)\n" +
    "  * Password: (leave it blank)\n\n" +

    "1.2. Turn off the phone, wait 5 seconds, and turn it back on.\n\n" +
    "2. If the issue persists, try inserting the SIM card into another device to see if it picks up a signal.\n\n" +
    "3. If the issue continues, contact us at spainconnect@connectivity.es or send a WhatsApp message to +34 681 916 322.\n\n" +

    "--------------------------------------------\n\n" +

    "You will be receiving more detailed informations regarding your plan in the next business day.\n\n";
    "Best regards\n\n";
    "Connectivity Team!\n";
    

    //Mensaje HTML
    /*
    mensaje = mensaje = "<p>Hi,</p>" +
    "<p>Your Spain Connect activation has been processed.</p>" +
    "<p>You can now enjoy 40Gb each month until  <strong>" + Utilities.formatDate(hojaActivar.getRange(iAct,4).getValue(), Session.getScriptTimeZone(), 'dd-MM-yyyy') + "</strong>. Your new phone number is <strong>" + hojaActivar.getRange(iAct,2).getValue() + "</strong>.</p>" +
    "<p>All the best.</p>" +
    "<br>" +
    "<p> Spain Connect team.<br><br>spainconnect@connectivity.es</p>" +
    "<br>" +
    "<ul>" +
    "<p><strong><br><br>Trouble shooting:</strong> If you don't get a data connection, it might be that your APN has not been set automatically. This is the most common issue when switching carriers, so don't worry. Just proceed as follows:</p>" +
    "<h3><br>1. Check your APN. This is the most common issue when switching carriers, so don't worry.</h3>" +
    "<h4>SETTING FOR ANDROID:</h4>" +
    "<p>Go to: Settings / Mobile Networks / APN (Access Point Name).</p>" +
    "<p>Create a new APN with the following settings:</p>" +
    "<ul>" +
    "<li>Name: FI</li>" +
    "<li>APN: fi.omv.es</li>" +
    "<li>APN Type: default,supl,dun</li>" +
    "<li>MVNO Type: IMSI</li>" +
    "<li>MVNO Value: 2140606</li>" +
    "</ul>" +
    "<h4>SETTING FOR iOS IPHONE:</h4>" +
    "<p>Go to: Settings / Mobile Data / Mobile Data Options, and configure:</p>" +
    "<p><strong>MOBILE DATA field:</strong></p>" +
    "<ul>" +
    "<li>APN: fi.omv.es</li>" +
    "<li>Username: (leave it blank)</li>" +
    "<li>Password: (leave it blank)</li>" +
    "</ul>" +
    "<p><strong>MMS field:</strong></p>" +
    "<ul>" +
    "<li>APN: fi.omv.es</li>" +
    "<li>Username: (leave it blank)</li>" +
    "<li>Password: (leave it blank)</li> " +
    "</ul>" +
    "<p><strong>INTERNET SHARING field:</strong></p>" +
    "<ul>" +
    "<li>APN: fi.omv.es</li>" +
    "<li>Username: (leave it blank)</li>" +
    "<li>Password: (leave it blank)</li>" +
    "</ul>" +
    "<p><strong>Optional LTE field:</strong></p>" +
    "<ul>" +
    "<li>APN: fi.omv.es</li>" +
    "<li>Username: (leave it blank)</li>" +
    "<li>Password: (leave it blank)</li>" +
    "</ul>" +
    "<p>1.2. Turn off the phone, leave it off for 5 seconds, and turn it back on.</p>" +
    "<h3>2. If the issue persists, try inserting the SIM card into another device to see if it picks up a signal.</h3>" +
    "<h3>3. If the issue persists, you can email us at spainconnect@connectivity.es or send a WhatsApp message to this phone number tel:+34681916322.</h3>" +
    "<p>Thank you for your patience!</p>"
    */

    destinatarioCliente = hojaActivar.getRange(iAct,12).getValue();
    destinatarioOculto = "soporte@connectivity.es";
    //Email HTML
    //enviarEmailHTML(destinatarioCliente, destinatarioOculto, mensaje, "Prepaid Registration Request Response");
    //Email sencillo
    enviarEmail(destinatarioCliente, destinatarioOculto, mensaje, "Prepaid Registration Request Response");


  } else {

    // Metemos el numero otra vez en listados
    hojaListados.getRange(n, numerocolumnalistados).setValue(numeroLinea);

    // El duplicado NO es correcto
    Logger.log("Error al hacer el duplicado")
    mensaje = "Hola.\n\nNo se ha ejecutado correctamente el Alta Spain Connect del ICC: " + iccnuevo + ". El error es" + respDuplicado;
    destinatarioCliente = "studentconnect@connectivityglobal.com";
    destinatarioOculto = "sistemas@connectivity.es";
    enviarEmail(destinatarioCliente, destinatarioOculto, mensaje, "Prepaid Registration Request Response");
  }

  return respDuplicado;

}


function activarAPIeSIMSpain (){

  //eSIM en Spain de momento es Manual
  
  //No podemos llamar directamente al actman por los contadores
  //actman();

  var hojaValidaciones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validacion Didit");
  // restamos una posición a iVal porque ya se ha hecho el +1
  var iVal = hojaValidaciones.getRange(1,13).getValue() - 1;

  //hojaActivar.getRange(iAct,2).setValue(999999999);

  //ENVIAR EMAIL A LUKE PARA HACER MANUAL
  var destinatario = "studentconnect@connectivityglobal.com";
  var mensaje = "Hay una nueva eSIM de Spain Connect pendiente de activar. En la hoja Validacion Didit, mirar la fila: " + iVal;
  var asunto = "Student Connect - eSIM SPAIN - Activación Manual";
  var copiaOculta = "sistemas@connectivity.es";

  enviarEmail(destinatario, copiaOculta, mensaje, asunto);

  // no hacemos el +1 porque ya se ha hecho
  //hojaValidaciones.getRange(1,13).setValue(iVal+1);

  var respuesta = "ok";
  return respuesta;

}


function obtenerTokenLikes() {
  const url = "https://api.likestelecom.com/token";
  const email = "vicky@connectivity.es";     // ⚠️ Sustituir por tus credenciales reales
  const password = "vicky19LIKES";            // ⚠️ Sustituir por tus credenciales reales
  Logger.log("Entramos en obtener token Likes");

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      email: email,
      password: password
    }),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());

  Logger.log("token = " + result.token);

  return result.token || null;
}


function activarAPIeSIMSpainLIKES(iAct) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  const plan = hoja.getRange(iAct, 4).getValue();      // Columna D
  const email = hoja.getRange(iAct, 12).getValue();    // Columna L

  const fiscalId = "B66148917"; // ⚠️ CIF de Call2World
  //const fiscalId = "B66347634"; // ⚠️ CIF de Connectivity
  //const fiscalId = "B66437872"; // ⚠️ CIF de Unique

  //El token se hace en el Proxy de DigitalOcean
  //const token = obtenerTokenLikes();
  //Logger.log("Salimos del obtener token Likes");
  //if (!token) {
  //  Logger.log("Error al obtener token");
  //  return "Error token";
  //}

  // Detectar ProductID según sufijo del plan
  let productId = null;
  if (plan.endsWith("-M")) productId = "2401";
  else if (plan.endsWith("-L")) productId = "1446";
  else if (plan.endsWith("-XLINT")) productId = "2413";
  else {
    Logger.log("Plan no reconocido: " + plan);
    hoja.getRange(iAct, 9).setValue("Error");                // Col I
    hoja.getRange(iAct, 20).setValue("Plan Likes inválido"); // Col T
    return "Plan inválido";
  }

  Logger.log("activarAPIeSIMSpainLIKES via proxy");
  Logger.log("fiscalId: " + fiscalId);
  Logger.log("productId: " + productId);
  Logger.log("email: " + email);

  // Antes del Proxy
  //const payload = {
  //  fiscalId: fiscalId,
  //  digitalSignature: false,
  //  products: [{
  //    productId: productId,
  //    family: "Mobile", 
  //    portability: false,
  //    eSim: true,
  //    eSimEmail: email
  //  }]
  //};

  const payload = {
    fiscalId: fiscalId,
    eSimEmail: email,
    productId: productId,
    portability: false
  };

  //const url = "https://api.likestelecom.com/signupv2";
  const url = "http://207.154.252.56:3010/es/likes/activate";
  const options = {
    method: "post",
    contentType: "application/json",
    //headers: {
    //  Authorization: "Bearer " + token
    //},
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  //const response = UrlFetchApp.fetch(url, options);
  //const res = JSON.parse(response.getContentText());
  //
  //if (res.orderId) {
  //  hoja.getRange(iAct, 24).setValue(res.orderId);  // Columna X = orderId
  //  hoja.getRange(iAct, 9).setValue("WIP");         // Columna I = Estado
  //  hoja.getRange(iAct, 23).setValue("Likes");         // Columna W = Proveedor
  //  return "ok";
  //} else {
  //  Logger.log("Error activación: " + response.getContentText());
  //  return "Error activación";
  //}

  let response;
  let bodyText;

  try {
    response = UrlFetchApp.fetch(url, options);
    bodyText = response.getContentText();
    Logger.log("Respuesta proxy Likes activate: " + bodyText);
  } catch (e) {
    Logger.log("Excepción al llamar al proxy Likes: " + e);
    hoja.getRange(iAct, 9).setValue("Error");
    hoja.getRange(iAct, 20).setValue(e.toString());
    return "Error activación";
  }

  const statusCode = response.getResponseCode();
  let res;

  try {
    res = JSON.parse(bodyText);
  } catch (e) {
    Logger.log("Error parseando JSON del proxy: " + e + " / body: " + bodyText);
    hoja.getRange(iAct, 9).setValue("Error");
    hoja.getRange(iAct, 20).setValue("Respuesta inválida del proxy");
    return "Error activación";
  }

  if (statusCode === 200 && res.status === "ok" && res.orderId) {
    hoja.getRange(iAct, 24).setValue(res.orderId);  // Columna X = orderId
    hoja.getRange(iAct, 9).setValue("WIP");         // Columna I = Estado
    hoja.getRange(iAct, 23).setValue("Likes");      // Columna W = Proveedor
    return "ok";
  } else {
    const msg = res.message || bodyText || "Error al activar eSIM via proxy";
    Logger.log("Error activación Likes via proxy: " + msg);
    hoja.getRange(iAct, 9).setValue("Error");
    hoja.getRange(iAct, 20).setValue(msg);
    return "Error activación";
  }

}


/*
function checkSpaineSIMActivations() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  const data = hoja.getDataRange().getValues();
  const token = obtenerTokenLikes();

  if (!token) {
    Logger.log("Error al obtener token de Likes");
    return;
  }

  for (let i = 1; i < data.length; i++) {
    const country = data[i][0].toString().toLowerCase(); // Col A
    const tipo = data[i][4];                             // Col E
    const estado = data[i][8];                           // Col I
    const orderId = data[i][23];                         // Col X
    const email = data[i][11];                           // Col L                          
    const name = data[i][10];                            // Col K
    const expiryDateRaw = data[i][6];                    // Col G (Fecha caducidad)

    // Formatear fecha de caducidad a DD/MM/YYYY
    let expiryDate = "";
    if (expiryDateRaw instanceof Date) {
      const day = ("0" + expiryDateRaw.getDate()).slice(-2);
      const month = ("0" + (expiryDateRaw.getMonth() + 1)).slice(-2);
      const year = expiryDateRaw.getFullYear();
      expiryDate = `${day}/${month}/${year}`;
    } else {
      expiryDate = expiryDateRaw; // Por si ya viene como texto
    }

    if (country === "spain" && tipo === "eSIM" && estado === "WIP" && orderId) {
      Logger.log("Consultando orderId: " + orderId);

      // --- Versión simple del endpoint ---
      const url = `https://api.likestelecom.com/draft-order-v2?orderId=${orderId}`;

      const options = {
        method: "get",
        headers: {
          Authorization: "Bearer " + token
        },
        muteHttpExceptions: true
      };

      try {
        const response = UrlFetchApp.fetch(url, options);
        const raw = response.getContentText();
        Logger.log("Raw response: " + raw);

        let res;
        try {
          res = JSON.parse(raw);
        } catch (err) {
          Logger.log("Error al parsear respuesta JSON: " + err);
          continue;
        }

        Logger.log("res.status = " + res.status);

        if (res.status === "COMPLETED") {
          const product = res.products?.[0] || {};
          const msisdn = product.lineNumber || "";
          const icc = product.icc || "";

          if (icc)   hoja.getRange(i + 1, 2).setValue(icc);     // Col B (ICC)
          if (msisdn)hoja.getRange(i + 1, 3).setValue(msisdn);  // Col C (MSISDN)
          hoja.getRange(i + 1, 9).setValue("Activa");           // Col I

          // Email
          // === PDFs adjuntos ===
          const pdf1 = DriveApp.getFileById("1ACFDBsgG48OQgh5BoDkVjZyDiDrKPErj").getAs("application/pdf");
          const pdf2 = DriveApp.getFileById("1TmHgaOD8XuyNMI9L5E9_IAtg_Sfyur6X").getAs("application/pdf");

          // === Mensaje del email ===
          const asunto = "Your Spanish eSIM is now active!";
          const mensaje =
            "Hi " + name + ",\n\n" +
            "Welcome to Student Connect!!! In this email you have the data and instructions to install your eSIM on iPhone and Android devices.\n\n" +
            "You have just received your eSIM with the QR code, PIN and PUK.\n\n" +
            "Your new Spanish number is: " + msisdn + "\n" +
            "Your plan will be active until the date: " + expiryDate + "\n\n" +
            "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
            "For roaming activation, please fill in this forms:\n" +
            "https://docs.google.com/forms/d/e/1FAIpQLSf_VL0oxXlCZJ5A8-rkp7cyL5mNOJq6StHX-aTtR1s7CWhOjQ/viewform\n\n" +
            "TELL YOUR FRIENDS AND GET PAID! Find more information here:\n" +
            "https://www.connectivityglobal.com/spain-connect-esim/#friends\n\n" +
            "If you have any questions, just reply to this email — we're here to help you.\n\n" +
            "Enjoy your plan!\n\n" +
            "Student Connect Team";

          // === Envío con adjuntos ===
          GmailApp.sendEmail(email, asunto, mensaje, {
            from: "studentconnect@connectivityglobal.com",
            attachments: [pdf1, pdf2],
            bcc: "support@connectivityglobal.com",
          });

          Logger.log(`Línea activada fila ${i + 1}: ${msisdn} - email enviado`);
        } else {
          Logger.log(`Fila ${i + 1}: orden aún en estado ${res.status}`);
        }

      } catch (e) {
        Logger.log(`Error en fila ${i + 1}: ${e}`);
      }
    }
  }
}
*/

function checkSpaineSIMActivations() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  const data = hoja.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const country = data[i][0].toString().toLowerCase(); // Col A
    const tipo = data[i][4];                             // Col E
    const estado = data[i][8];                           // Col I
    const orderId = data[i][23];                         // Col X
    const email = data[i][11];                           // Col L
    const name = data[i][10];                            // Col K
    const expiryDateRaw = data[i][6];                    // Col G (Fecha caducidad)

    // Formatear fecha de caducidad a DD/MM/YYYY
    let expiryDate = "";
    if (expiryDateRaw instanceof Date) {
      const day = ("0" + expiryDateRaw.getDate()).slice(-2);
      const month = ("0" + (expiryDateRaw.getMonth() + 1)).slice(-2);
      const year = expiryDateRaw.getFullYear();
      expiryDate = `${day}/${month}/${year}`;
    } else {
      expiryDate = expiryDateRaw; // Por si ya viene como texto
    }

    if (country === "spain" && tipo === "eSIM" && estado === "WIP" && orderId) {
      Logger.log("Consultando orderId (via proxy): " + orderId);

      const url = "http://207.154.252.56:3010/es/likes/order-draft";
      const payload = { orderId: orderId };

      const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      try {
        const response = UrlFetchApp.fetch(url, options);
        const raw = response.getContentText();
        Logger.log("Raw response proxy order-draft: " + raw);

        let json;
        try {
          json = JSON.parse(raw);
        } catch (err) {
          Logger.log("Error al parsear JSON del proxy order-draft: " + err);
          continue;
        }

        if (json.status !== "ok") {
          Logger.log("Proxy devolvió error en order-draft: " + (json.message || raw));
          continue;
        }

        // A partir de aquí usamos la respuesta original de Likes
        const res = json.raw;

        Logger.log("res.status = " + res.status);

        if (res.status === "COMPLETED") {
          const product = res.products?.[0] || {};
          const msisdn = product.lineNumber || "";
          const icc = product.icc || "";

          if (icc)   hoja.getRange(i + 1, 2).setValue(icc);    // Col B (ICC)
          if (msisdn)hoja.getRange(i + 1, 3).setValue(msisdn); // Col C (MSISDN)
          hoja.getRange(i + 1, 9).setValue("Activa");          // Col I

          // === PDFs adjuntos ===
          const pdf1 = DriveApp.getFileById("1ACFDBsgG48OQgh5BoDkVjZyDiDrKPErj").getAs("application/pdf");
          const pdf2 = DriveApp.getFileById("1TmHgaOD8XuyNMI9L5E9_IAtg_Sfyur6X").getAs("application/pdf");

          // === Mensaje del email ===
          const asunto = "Your Spanish eSIM is now active!";
          const mensaje =
            "Hi " + name + ",\n\n" +
            "Welcome to Student Connect!!! In this email you have the data and instructions to install your eSIM on iPhone and Android devices.\n\n" +
            "You have just received your eSIM with the QR code, PIN and PUK.\n\n" +
            "Your new Spanish number is: " + msisdn + "\n" +
            "Your plan will be active until the date: " + expiryDate + "\n\n" +
            "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
            "For roaming activation, please fill in this form:\n" +
            "https://docs.google.com/forms/d/e/1FAIpQLSf_VL0oxXlCZJ5A8-rkp7cyL5mNOJq6StHX-aTtR1s7CWhOjQ/viewform\n\n" +
            "TELL YOUR FRIENDS AND GET PAID! Find more information here:\n" +
            "https://www.connectivityglobal.com/spain-connect-esim/#friends\n\n" +
            "If you have any questions, just reply to this email — we're here to help you.\n\n" +
            "Enjoy your plan!\n\n" +
            "Student Connect Team";

          // === Envío con adjuntos ===
          GmailApp.sendEmail(email, asunto, mensaje, {
            from: "studentconnect@connectivityglobal.com",
            attachments: [pdf1, pdf2],
            bcc: "support@connectivityglobal.com",
          });

          Logger.log(`Línea activada fila ${i + 1}: ${msisdn} - email enviado`);
        } else {
          Logger.log(`Fila ${i + 1}: orden aún en estado ${res.status}`);
        }

      } catch (e) {
        Logger.log(`Error en fila ${i + 1}: ${e}`);
      }
    }
  }
}




/*
//NUEVA FUNCION x NICO con el mail mejorado, adjuntando PDFs, fecha de caducidad y todo corregido. 
function checkSpaineSIMActivations() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  const data = hoja.getDataRange().getValues();
  const token = obtenerTokenLikes();

  if (!token) {
    Logger.log("Error al obtener token de Likes");
    return;
  }

  for (let i = 1; i < data.length; i++) {
    const country = data[i][0].toString().toLowerCase(); // Col A
    const tipo = data[i][4];                             // Col E
    const estado = data[i][8];                           // Col I
    const orderId = data[i][23].toString().trim();       // Col X
    const email = data[i][11];                           // Col L
    const name = data[i][10];                            // Col K
    const expiryDateRaw = data[i][6];                    // Col G (Fecha caducidad)

    // Formatear fecha de caducidad a DD/MM/YYYY
    let expiryDate = "";
    if (expiryDateRaw instanceof Date) {
      const day = ("0" + expiryDateRaw.getDate()).slice(-2);
      const month = ("0" + (expiryDateRaw.getMonth() + 1)).slice(-2);
      const year = expiryDateRaw.getFullYear();
      expiryDate = `${day}/${month}/${year}`;
    } else {
      expiryDate = expiryDateRaw; // Por si ya viene como texto
    }

    if (country === "spain" && tipo === "eSIM" && estado === "WIP" && orderId) {
      Logger.log("Consultando orderId: " + orderId);

      const url = "https://api.likestelecom.com/order-draft?orderId=" + encodeURIComponent(orderId);
      const options = {
        method: "get",
        headers: {
          Authorization: "Bearer " + token
        },
        muteHttpExceptions: true
      };

      try {
        const response = UrlFetchApp.fetch(url, options);
        const raw = response.getContentText();
        Logger.log("Raw response: " + raw);

        let res;
        try {
          res = JSON.parse(raw);
        } catch (err) {
          Logger.log("Error al parsear respuesta JSON: " + err);
          continue;
        }

        Logger.log("res.status = " + res.status);

        if (res.status === "COMPLETED") {
          const product = res.products?.[0] || {};
          const msisdn = product.lineNumber || "";
          const icc = product.icc || "";

          if (msisdn) hoja.getRange(i + 1, 3).setValue(msisdn); // Col C
          if (icc) hoja.getRange(i + 1, 2).setValue(icc);       // Col B
          hoja.getRange(i + 1, 9).setValue("Activa");           // Col I

          // === PDFs adjuntos ===
          const pdf1 = DriveApp.getFileById("1wQYVwkktb0Qw94xpHfIMCLECzRsmvBaa").getAs("application/pdf");
          const pdf2 = DriveApp.getFileById("1-38Vx2Yuzpk00JcPemjL2ySPWObK4X27").getAs("application/pdf");

          // === Mensaje del email ===
          const asunto = "Your Spanish eSIM is now active!";
          const mensaje =
            "Hi " + name + ",\n\n" +
            "Welcome to Student Connect!!! In this email you have the data and instructions to install your eSIM on iPhone and Android devices.\n\n" +
            "You have just received your eSIM with the QR code, PIN and PUK.\n\n" +
            "Your new Spanish number is: " + msisdn + "\n" +
            "Your plan will be active until the date: " + expiryDate + "\n\n" +
            "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
            "For roaming activation, please fill in this forms:\n" +
            "https://docs.google.com/forms/d/e/1FAIpQLSf_VL0oxXlCZJ5A8-rkp7cyL5mNOJq6StHX-aTtR1s7CWhOjQ/viewform\n\n" +
            "TELL YOUR FRIENDS AND GET PAID! Find more information here:\n" +
            "https://www.connectivityglobal.com/spain-connect-esim/#friends\n\n" +
            "If you have any questions, just reply to this email — we're here to help you.\n\n" +
            "Enjoy your plan!\n\n" +
            "Student Connect Team";

          // === Envío con adjuntos ===
          GmailApp.sendEmail(email, asunto, mensaje, {
            from: "soporte@connectivity.es",
            attachments: [pdf1, pdf2]
          });

          Logger.log(`Línea activada fila ${i + 1}: ${msisdn} - email enviado`);
        } else {
          Logger.log(`Fila ${i + 1}: orden aún en estado ${res.status}`);
        }

      } catch (e) {
        Logger.log(`Error en fila ${i + 1}: ${e}`);
      }
    }
  }
}
*/



function pasarDatos(filaOrigen,datos) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones").activate();
  var hojaOrigen = SpreadsheetApp.getActiveSheet();
  var docGoogleHoja2 = SpreadsheetApp.openById("1Z1qOg8K0M-krp2bHwdZXVRnzWWppqBnKQ63Eg8c3X20");
  var hojaDestino = docGoogleHoja2.getSheetByName("Banco datos puntual");

  var filaDestino = hojaDestino.getRange("g2").getValue() + 1;
  hojaDestino.getRange("a" + filaDestino).setValue(hojaOrigen.getRange("c" + filaOrigen).getValue());
  hojaDestino.getRange("b" + filaDestino).setValue("2");
  hojaDestino.getRange("c" + filaDestino).setValue("SpainConnect");
  hojaDestino.getRange("d" + filaDestino).setValue(datos);
  hojaDestino.getRange("e" + filaDestino).setValue("Alta");
}


function eliminarBancoDatos(filaOrigen) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones").activate();
  var hojaOrigen = SpreadsheetApp.getActiveSheet();
  var docGoogleHoja2 = SpreadsheetApp.openById("1Z1qOg8K0M-krp2bHwdZXVRnzWWppqBnKQ63Eg8c3X20");
  var hojaDestino = docGoogleHoja2.getSheetByName("Banco datos puntual");

  var filaDestino = hojaDestino.getRange("g2").getValue() + 1;
  hojaDestino.getRange("a" + filaDestino).setValue(hojaOrigen.getRange("c" + filaOrigen).getValue());
  hojaDestino.getRange("b" + filaDestino).setValue("2");
  hojaDestino.getRange("c" + filaDestino).setValue("SpainConnect");
  hojaDestino.getRange("e" + filaDestino).setValue("Baja");
}

function modificarDatos(filaOrigen,datos) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones").activate();
  var hojaOrigen = SpreadsheetApp.getActiveSheet();
  var docGoogleHoja2 = SpreadsheetApp.openById("1Z1qOg8K0M-krp2bHwdZXVRnzWWppqBnKQ63Eg8c3X20");
  var hojaDestino = docGoogleHoja2.getSheetByName("Banco datos puntual");

  var filaDestino = hojaDestino.getRange("g2").getValue() + 1;
  hojaDestino.getRange("a" + filaDestino).setValue(hojaOrigen.getRange("c" + filaOrigen).getValue());
  hojaDestino.getRange("b" + filaDestino).setValue("2");
  hojaDestino.getRange("c" + filaDestino).setValue("SpainConnect");
  hojaDestino.getRange("d" + filaDestino).setValue(datos);
  hojaDestino.getRange("e" + filaDestino).setValue("Modificación");
}



function duplicarSIMSpain(numLinea, iccnew) {
  var url;
  var respuesta;
  var response;
  var data;
  var jsonresp;
  var options;
  var body;

  ObtenerTokenFi();

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones").activate();
  var hojaActivar = SpreadsheetApp.getActiveSheet();

  var iccold = ObtenerIccOld(numLinea);
  var body = "{\"msisdn\": " + numLinea + ", \"newIccId\": " + iccnew + ", \"oldIccId\": " + iccold + "}";


  //realiza duplicado
  options = {
    "method": "get",
    "payload": body,
    "headers": {
      "Content-type": "application/json",
      "Accept": "*/*",
      "Authorization": Token
    },
    "muteHttpExceptions": true
  }

  url = "https://gateway-operators.finetwork.com/api/v2/sims/swap";

  response = UrlFetchApp.fetch(url, options);
  data = response.getContentText();
  jsonresp = JSON.parse(data)


  if (jsonresp["error"] == null) {
    respuesta = "ok";

  } else {
    respuesta = jsonresp["error"]["message"];
    Logger.log("error al duplicar = " + respuesta);
  }
  return respuesta;
}



function obtenrIdSpain(num) {
  try {
    Logger.log("MSISDN que se está enviando: " + num);
    ObtenerTokenFi();  // Asegúrate de que la variable Token sea global

    const url = `https://gateway-operators.finetwork.com/api/v2/phoneline/searchByMsisdn?msisdn=${num}`;
    const options = {
      "method": "get",
      "muteHttpExceptions": true,  // para poder capturar errores 404 o 500
      "headers": {
        "Content-type": "application/json",
        "Accept": "*/*",
        "Authorization": Token
      }
    };

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const json = JSON.parse(response.getContentText());

    if (statusCode !== 200 || json.error) {
      Logger.log(`Error al obtener ID: ${json.error?.message || "Status " + statusCode}`);
      return "error";
    }

    return json.id;
  } catch (e) {
    Logger.log("Excepción al obtener ID Spain: " + e.message);
    return "error";
  }
}

function suspenderSIMSpainAPI(idLinea) {
  var url;
  var jsonresp;
  var response;
  var data;
  var options;
  var errorSuspension;

  try {
    body = "{\"id\": " + idLinea + "}";

    options = {
      "method": "patch",
      "payload": body,
      "headers": {
        "Content-type": "application/json",
        "Accept": "*/*",
        "Authorization": Token
      }
    }

    url = "https://gateway-operators.finetwork.com/api/v2/services/suspend";

    response = UrlFetchApp.fetch(url, options);
    data = response.getContentText();
    jsonresp = JSON.parse(data)

    if (jsonresp["error"] == null) {

      errorSuspension = "Suspendida";

    } else {
      errorSuspension = "error";
    }

  } catch (error) {
    Logger.log("Ha capturado error el try de suspender" + error);
    errorSuspension = "error";
  }
  return errorSuspension;
}


function reactivarSIMSPainAPI(idLinea) {
  var url;
  var jsonresp;
  var response;
  var data;
  var options;
  var errorSuspension;

  try {

    body = "{\"id\": " + idLinea + "}";

    options = {
      "method": "patch",
      "payload": body,
      "headers": {
        "Content-type": "application/json",
        "Accept": "*/*",
        "Authorization": Token
      }
    }

    url = "https://gateway-operators.finetwork.com/api/v2/services/unsuspend";

    response = UrlFetchApp.fetch(url, options);
    data = response.getContentText();
    jsonresp = JSON.parse(data)

    if (jsonresp["error"] == null) {

      errorSuspension = "Activa";

    } else {
      errorSuspension = "error";

    }
    Logger.log("el resultado del intento de reactivación es: " + jsonresp["error"])
  } catch (error) {
    Logger.log("Ha capturado error el try de reactivar" + error);
    errorSuspension = "error";
  }
  return errorSuspension;
}


function ObtenerIccOld(num) {
  var url;
  var jsonresp;
  var response;
  var data;
  var jsonresp;
  var options;
  var sim;

  //ObtenerTokenFi();

  options = {
    "method": "get",
    "headers": {
      "Content-type": "application/json",
      "Accept": "*/*",
      "Authorization": Token
    }
  }

  url = "https://gateway-operators.finetwork.com/api/v2/phoneline/searchByMsisdn?msisdn=" + num;

  response = UrlFetchApp.fetch(url, options);
  data = response.getContentText();
  jsonresp = JSON.parse(data);

  if (jsonresp["error"] == null) {

    sim = jsonresp["card"]["iccid"];
    idlinea = jsonresp["id"];

    //si está suspendida la reactivamos
    var estadoLinea = jsonresp["currentStatus"]["name"];
    if(estadoLinea == "SUSPENDIDO IMPAGO"){
      estadoLinea = reactivarSIMSPainAPI(idlinea);      
    }

  } else {
    sim = "error";
  }

  return sim;

}

function ObtenerTokenFi() {
  spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('a1').activate();

  var hojaToken = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Token");
  var cell = hojaToken.getRange("B1");

  var jsonString = cell.getValue();
  var payload = JSON.parse(jsonString);

  var options = {
    "method": "post",
    "payload": JSON.stringify(payload),
    "headers": {
      "Content-Type": "application/json"
    }

  };

  var response = UrlFetchApp.fetch("https://gateway-operators.finetwork.com/api/v2/token", options);
  var data = response.getContentText();
  var respuesta = JSON.parse(data)
  hojaToken.getRange("b3").setValue(data)

  //var targetCell = sheet.getRange("B" + sheet.getRange("b12").getValue());
  var targetCell = hojaToken.getRange("B8");
  Token = respuesta.access_token;

  targetCell.setValue(Token);

  // esto sirve como ocntador de procesos
  hojaToken.getRange("b12").setValue(1 + hojaToken.getRange("b12").getValue());
  TimeToken = new Date().getTime();
}


function reactivarSIMSpain(num) {
  var url;
  var jsonresp;
  var response;
  var data;
  var jsonresp;
  var options;
  var sim;

  ObtenerTokenFi();

  options = {
    "method": "get",
    "headers": {
      "Content-type": "application/json",
      "Accept": "*/*",
      "Authorization": Token
    }
  }

  url = "https://gateway-operators.finetwork.com/api/v2/phoneline/searchByMsisdn?msisdn=" + num;

  response = UrlFetchApp.fetch(url, options);
  data = response.getContentText();
  jsonresp = JSON.parse(data);

  if (jsonresp["error"] == null) {

    sim = jsonresp["card"]["iccid"];
    idlinea = jsonresp["id"];

    //si está suspendida la reactivamos
    var estadoLinea = jsonresp["currentStatus"]["name"];
    if(estadoLinea == "SUSPENDIDO IMPAGO"){
      estadoLinea = reactivarSIMSPainAPI(idlinea);      
    }

  } else {
    estadoLinea = "error";
  }

  return estadoLinea;

}

//*******************************************************************************************************************************************
//************       FRANCE          ********************************************************************************************************
//*******************************************************************************************************************************************

function activarAPISIMFrance(iAct) {
  const hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  
  const icc = hojaActivar.getRange(iAct, 2).getValue();
  let planOriginal = hojaActivar.getRange(iAct, 4).getValue();
  const sim_esim = hojaActivar.getRange(iAct, 5).getValue();
  const name = hojaActivar.getRange(iAct, 11).getValue();
  const email = hojaActivar.getRange(iAct, 12).getValue();
  
  // Solo SIM física
  if (sim_esim !== "SIM") {
    Logger.log("No es SIM física, se omite la activación automática");
    return "skip";
  }
  
  // Traducir plan
  let pack;
  if (planOriginal.endsWith("-M")) {
    pack = "Connect30";
  } else if (planOriginal.endsWith("-L")) {
    pack = "Connect80";
  } else if (planOriginal.endsWith("-XL")) {
    pack = "Connect130";
  } else {
    Logger.log("Plan no reconocido: " + planOriginal);
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue("Plan no reconocido");
    return "error";
  }
  
  // Construimos body
  const body = {
    iccid: icc,
    name: name,
    pack: pack
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  };
  
  try {
    const url = "http://207.154.252.56:3010/activate"; // Sustituir con tu IP real y puerto
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.status === "ok") {
      hojaActivar.getRange(iAct, 9).setValue("WIP");
      hojaActivar.getRange(iAct, 23).setValue("Airmob");
      hojaActivar.getRange(iAct, 24).setValue(json.ticket); // ← Aquí guardamos la respuesta
      Logger.log("Ticket guardado y estado WIP. Ticket: " + json.ticket);
      return "ok";
    } else {
      hojaActivar.getRange(iAct, 9).setValue("Error");
      hojaActivar.getRange(iAct, 20).setValue(json.message); 
      Logger.log("Error activación: " + json.message);
      return "error";
    }
  } catch (error) {
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue(error.toString());
    Logger.log("Excepción: " + error);
    return "error";
  }
}


function activarAPIeSIMFrance(iAct) {
  const hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  const hojaESIM = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ICC France Airmob");
  
  // Datos fila activación
  let planOriginal = hojaActivar.getRange(iAct, 4).getValue();
  const name = hojaActivar.getRange(iAct, 11).getValue();
  const email = hojaActivar.getRange(iAct, 12).getValue();
  
  // Buscar ICC libre
  const data = hojaESIM.getRange(2, 1, hojaESIM.getLastRow(), 2).getValues();
  let iccLibre = null;
  let filaLibre = null;

  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === "libre") {
      iccLibre = data[i][0];
      filaLibre = i + 2; // data empieza en fila 2
      break;
    }
  }
  
  if (!iccLibre) {
    Logger.log("No hay ICC eSIM libre disponible.");
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue("No hay ICC eSIM libre");
    return "error";
  }
  
  // Marcar ICC como usado
  hojaESIM.getRange(filaLibre, 2).setValue("usado");
  
  // Traducir plan
  let pack;
  if (planOriginal.endsWith("-M")) {
    pack = "Connect30";
  } else if (planOriginal.endsWith("-L")) {
    pack = "Connect80";
  } else if (planOriginal.endsWith("-XL")) {
    pack = "Connect130";
  } else {
    Logger.log("Plan no reconocido: " + planOriginal);
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue("Plan no reconocido");
    return "error";
  }
  
  // Construir body
  const body = {
    iccid: iccLibre,
    name: name,
    pack: pack
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  };
  
  try {
    const url = "http://207.154.252.56:3010/activate"; 
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.status === "ok") {
      hojaActivar.getRange(iAct, 9).setValue("WIP");
      hojaActivar.getRange(iAct, 2).setValue(iccLibre); // Guardamos ICC en hoja activaciones
      hojaActivar.getRange(iAct, 23).setValue("Airmob");
      hojaActivar.getRange(iAct, 24).setValue(json.ticket); // ← Aquí guardamos la respuesta
      Logger.log("Ticket guardado y estado WIP. Ticket: " + json.ticket);
      return "ok";
      
    } else {
      hojaActivar.getRange(iAct, 9).setValue("Error");
      hojaActivar.getRange(iAct, 20).setValue(json.message);
      Logger.log("Error activación: " + json.message);
      return "error";
    }
  } catch (error) {
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue(error.toString());
    Logger.log("Excepción: " + error);
    return "error";
  }
}


function activarPreActAPIeSIMFrance(iAct) {
  const hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");

  // Datos fila activación
  let planOriginal = hojaActivar.getRange(iAct, 4).getValue();
  const name = hojaActivar.getRange(iAct, 11).getValue();
  const email = hojaActivar.getRange(iAct, 12).getValue();
  const icc = hojaActivar.getRange(iAct, 2).getValue();  // ya estaba asignado
  const msisdn = hojaActivar.getRange(iAct, 3).getValue();
  const expiryDateRaw = hojaActivar.getRange(iAct, 7).getValue();

  const expiryDate = Utilities.formatDate(expiryDateRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Validar ICC existente
  if (!icc || icc === "") {
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue("ICC vacío al activar desde PreAct");
    Logger.log("Error: ICC vacío al intentar activar desde PreActivada");
    return "error";
  }

  // Traducir plan
  let pack;
  if (planOriginal.endsWith("-M")) {
    pack = "Connect30";
  } else if (planOriginal.endsWith("-L")) {
    pack = "Connect80";
  } else if (planOriginal.endsWith("-XL")) {
    pack = "Connect130";
  } else {
    Logger.log("Plan no reconocido: " + planOriginal);
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue("Plan no reconocido");
    return "error";
  }

  // Construir body
  const body = {
    iccid: icc,
    name: name,
    pack: pack
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  };

  try {
    const url = "http://207.154.252.56:3010/activate"; 
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.status === "ok") {
      hojaActivar.getRange(iAct, 9).setValue("WIP_PRE");  // Diferencia importante
      hojaActivar.getRange(iAct, 23).setValue("Airmob");
      hojaActivar.getRange(iAct, 24).setValue(json.ticket); // guardar ticket
      Logger.log("Activación desde PreAct OK. Ticket: " + json.ticket);

      // Email
        // === PDFs adjuntos ===
        const pdf1 = DriveApp.getFileById("1jMJR6m3YC_LAdjOGk-UFuqe2Di7fmsa6").getAs("application/pdf");
        const pdf2 = DriveApp.getFileById("1Q66l1m-63jxewu4-PikKla3TGPNL22Ux").getAs("application/pdf");

        // === Mensaje del email ===
        const asunto = "Your french eSIM is now active!";
        const mensaje = 
          "Hi " + name + ",\n\n" +
          "Your France eSIM has been activated! Attached you have our France manuals to assist you with your eSIM.\n\n" +
          "Your French number is: " + msisdn + "\n" +
          "Your plan will be active until the date: " + expiryDate + "\n\n" +
          "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
          "If you have any questions, just reply to this email — we're here to help you.\n\n" +
          "Enjoy your plan!\n\n" +
          "Student Connect Team";

        // === Envío con adjuntos ===
        GmailApp.sendEmail(email, asunto, mensaje, {
          from: "studentconnect@connectivityglobal.com",
          attachments: [pdf1, pdf2],
          bcc: "support@connectivityglobal.com",
        });

      return "ok";
    } else {
      hojaActivar.getRange(iAct, 9).setValue("Error");
      hojaActivar.getRange(iAct, 20).setValue(json.message);
      Logger.log("Error activando desde PreAct: " + json.message);
      return "error";
    }
  } catch (error) {
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue(error.toString());
    Logger.log("Excepción activando desde PreAct: " + error);
    return "error";
  }
}



function preActivarAPIeSIMFrance(iAct) {
  const ss = SpreadsheetApp.getActive();
  const hojaActivar = ss.getSheetByName("Activaciones");
  const hojaESIM = ss.getSheetByName("ICC France Airmob");

  const data = hojaESIM.getRange(2, 1, hojaESIM.getLastRow() - 1, 2).getValues();
  let iccLibre = null;
  let filaLibre = null;
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === "libre") {
      iccLibre = data[i][0];
      filaLibre = i + 2;
      break;
    }
  }
  if (!iccLibre) {
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue("No ICC libre");
    return "error";
  }

  hojaESIM.getRange(filaLibre, 2).setValue("usado");
  hojaActivar.getRange(iAct, 2).setValue(iccLibre);
  //hojaActivar.getRange(iAct, 9).setValue("PreActivada");

  const name = hojaActivar.getRange(iAct, 11).getValue();
  const email = hojaActivar.getRange(iAct, 12).getValue();
  const activationdateRaw = hojaActivar.getRange(iAct, 6).getValue(); // esto devuelve un objeto Date
  const activationdate = Utilities.formatDate(activationdateRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");

  enviarQREmailSinNumero(name, email, iccLibre, activationdate); // sin MSISDN

  return "PreActivada";
}


function checkFranceActivations() {
  const hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  const lastRow = hojaActivar.getLastRow();
  const data = hojaActivar.getRange(2, 1, lastRow - 1, 24).getValues();
  
  for (let i = 0; i < data.length; i++) {
    const country = data[i][0];
    const estado = data[i][8];
    const ticket = data[i][23];
    const sim_esim = data[i][4];
    const name = data[i][10];
    const email = data[i][11];
    const expiryDate = data[i][6];

    if (country === "france" && ticket && (estado === "WIP" || estado === "WIP_PRE")) {
      try {
        // Llamada al proxy /ticket
        const body = { noticket: ticket }; // (igual que ahora; el proxy añade apikey)
        const options = {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(body),
          muteHttpExceptions: true,
        };
        const ticketUrl = "http://207.154.252.56:3010/ticket";
        const response = UrlFetchApp.fetch(ticketUrl, options);

        // CHANGE: log defensivo para ver qué llega realmente
        const raw = response.getContentText();
        Logger.log("Respuesta raw /ticket: " + raw);

        // CHANGE: parseo según el ejemplo real de Airmob (objeto con clave numérica)
        const json = JSON.parse(raw);
        const firstKey = Object.keys(json)[0];               // "1", "2", etc.
        const node = firstKey ? json[firstKey] : null;       // { ticket, msisdn, code, ... }
        const code = node && node.code ? node.code : "KO";   // "OK" esperado
        const msisdn = node && node.msisdn ? node.msisdn : null;

        // CHANGE: validar antes de actualizar hoja / enviar correos
        if (code === "OK" && msisdn) {
          // Normalizar número: quitar prefijo 33 si existe
          let msisdnFinal = msisdn;
          if (msisdn && msisdn.startsWith("33")) {
            msisdnFinal = msisdn.substring(2); // elimina los dos primeros caracteres
          }
          // Guardar MSISDN
          hojaActivar.getRange(i + 2, 3).setValue(msisdnFinal);

          const pdf1 = DriveApp.getFileById("1jMJR6m3YC_LAdjOGk-UFuqe2Di7fmsa6").getAs("application/pdf");
          const pdf2 = DriveApp.getFileById("1Q66l1m-63jxewu4-PikKla3TGPNL22Ux").getAs("application/pdf");

          // Enviar emails
          if (sim_esim === "SIM") {
            const asunto = "Your SIM is activated";
            const mensaje = "Hi " + name + ",\n\nYour SIM has been successfully activated!\n\nYour new French number is: " + msisdnFinal + ".\n\nEnjoy your plan!\n\nStudent Connect Team";
            enviarEmail(email, "soporte@connectivity.es", mensaje, asunto);
          } else if (sim_esim === "eSIM") {
            if (estado === "WIP_PRE") {
              // CHANGE: en preactivadas solo informar número (sin reenviar QR)
              const asunto = "Your eSIM is activated";

              const mensaje = 
                "Hi " + name + ",\n\n" +
                "Your France eSIM has been activated! Attached you have our french manuals to assist you with your eSIM.\n\n" +
                "Your French number is: " + msisdnFinal + "\n" +
                "Your plan will be active until the date: " + expiryDate + "\n\n" +
                "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
                "If you have any questions, just reply to this email — we're here to help you.\n\n" +
                "Enjoy your plan!\n\n" +
                "Student Connect Team";

                // === Envío con adjuntos ===
                GmailApp.sendEmail(email, asunto, mensaje, {
                  from: "studentconnect@connectivityglobal.com",
                  attachments: [pdf1, pdf2],
                  bcc: "support@connectivityglobal.com",
                });

            } else {
              enviarQREmailConNumero(name, email, data[i][1], msisdnFinal, expiryDate);
            }
          }

          // Marcar estado
          hojaActivar.getRange(i + 2, 9).setValue("Activa");
          Logger.log("Ticket actualizado y activación completada. MSISDN: " + msisdnFinal);
        } else {
          // CHANGE: no marcar Activa ni enviar correo si no hay msisdn o code != OK
          Logger.log("Ticket no OK o sin msisdn. code=" + code + ", msisdn=" + msisdn + ", fila=" + (i + 2));
        }

      } catch (error) {
        Logger.log("Error procesando ticket fila " + (i + 2) + ": " + error);
      }
    }
  }
}



/*
//NUEVA funcion NICO con los cambios de email, pdfs adjuntos, etc. De momento la quito porque debe haber algun error al enviarse desde Dopost, con los archivos, etc.

function checkFranceActivations() {
  const hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  const lastRow = hojaActivar.getLastRow();
  const data = hojaActivar.getRange(2, 1, lastRow - 1, 24).getValues();

  for (let i = 0; i < data.length; i++) {
    const country = data[i][0];     // Col A
    const estado = data[i][8];      // Col I
    const ticket = data[i][23];     // Col X
    const sim_esim = data[i][4];    // Col E
    const name = data[i][10];       // Col K
    const email = data[i][11];      // Col L
    const expiryDateRaw = data[i][6]; // Col G

    // Formatear fecha de caducidad
    let expiryDate = "";
    if (expiryDateRaw instanceof Date) {
      const day = ("0" + expiryDateRaw.getDate()).slice(-2);
      const month = ("0" + (expiryDateRaw.getMonth() + 1)).slice(-2);
      const year = expiryDateRaw.getFullYear();
      expiryDate = `${day}/${month}/${year}`;
    } else {
      expiryDate = expiryDateRaw;
    }

    if (country === "france" && ticket && (estado === "WIP" || estado === "WIP_PRE")) {
      try {
        // Llamada al proxy /ticket
        const body = { noticket: ticket };
        const options = {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(body),
          muteHttpExceptions: true,
        };
        const ticketUrl = "http://207.154.252.56:3010/ticket";
        const response = UrlFetchApp.fetch(ticketUrl, options);
        const json = JSON.parse(response.getContentText());
        const msisdn = json[0][ticket]?.msisdn || "not found";

        // Guardar número
        hojaActivar.getRange(i + 2, 3).setValue(msisdn);

        // Enviar emails
        if (sim_esim === "SIM") {
          // Nuevo email para SIM
          const asunto = "Your French SIM is now active!";
          const mensaje =
            "Hi " + name + ",\n\n" +
            "Welcome to Student Connect!!! In this email you have the QR code, the data and instructions to install your eSIM on iPhone and Android devices.\n\n" +
            "Your new French number is: " + msisdn + "\n" +
            "Your plan will be active until the date: " + expiryDate + "\n" +
            "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
            "If you have any questions, just reply to this email — we're here to help you.\n\n" +
            "Enjoy your plan!\n\n" +
            "Best regards,\n" +
            "The Student Connect Team";

          enviarEmail(email, "soporte@connectivity.es", mensaje, asunto);

        } else if (sim_esim === "eSIM") {
          if (estado === "WIP_PRE") {
            // Email original para WIP_PRE
            const asunto = "Your SIM is activated";
            const mensaje = "Hi " + name + ",\n\nYour SIM has been successfully activated!\n\nYour new French number is: " + msisdn + ".\n\nEnjoy your plan!\n\nStudent Connect Team";
            enviarEmail(email, "soporte@connectivity.es", mensaje, asunto);
          } else {
            // Email con QR y número + adjuntos
            const pdf1 = DriveApp.getFileById("17GkruusuoqC0cJh1WT1eaLUFJdPBhiyO").getAs("application/pdf");
            const pdf2 = DriveApp.getFileById("1-FgRe-8Lc57T2bl2XnUho-oDsbUo8XAA").getAs("application/pdf");

            const asunto = "Your French eSIM is now active!";
            const mensaje =
              "Hi " + name + ",\n\n" +
              "Welcome to Student Connect!!! In this email you have the QR code, the data and instructions to install your eSIM on iPhone and Android devices.\n\n" +
              "Your new French number is: " + msisdn + "\n" +
              "Your plan will be active until the date: " + expiryDate + "\n" +
              "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
              "If you have any questions, just reply to this email — we're here to help you.\n\n" +
              "Enjoy your plan!\n\n" +
              "Best regards,\n" +
              "The Student Connect Team";

            GmailApp.sendEmail(email, asunto, mensaje, {
              from: "soporte@connectivity.es",
              attachments: [pdf1, pdf2]
            });
          }
        }

        // Marcar estado como Activa
        hojaActivar.getRange(i + 2, 9).setValue("Activa");
        Logger.log("Ticket actualizado y activación completada. MSISDN: " + msisdn);

      } catch (error) {
        Logger.log("Error procesando ticket fila " + (i + 2) + ": " + error);
      }
    }
  }
}
*/

function enviarQREmailConNumero(name, email, icc, msisdn, expiryDate) {
  const carpetaID = '1uEQmtKU72a7i-2gaUVjVq5KI-ClYjcH-';
  const carpeta = DriveApp.getFolderById(carpetaID);
  const nombreArchivo = icc + '-qr.png';
  const archivos = carpeta.getFilesByName(nombreArchivo);

  const pdf1 = DriveApp.getFileById("1jMJR6m3YC_LAdjOGk-UFuqe2Di7fmsa6").getAs("application/pdf");
  const pdf2 = DriveApp.getFileById("1Q66l1m-63jxewu4-PikKla3TGPNL22Ux").getAs("application/pdf");

  if (archivos.hasNext()) {
    const archivo = archivos.next();
    const cuerpo = "Hi " + name + ",\n\n" +
    "Welcome to Student Connect!!!\n\n" +
    "In this email you have the QR code, the data, and instructions to install your eSIM on iPhone and Android devices.\n\n" +
    "Your French number is: " + msisdn + "\n" +
    "Your plan will be active until: " + expiryDate + "\n" +
    "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
    "If you have any questions, just reply to this email — we're here to help you.\n\n" +
    "Enjoy your plan!\n\n" +
    "Best regards,\n" +
    "The Student Connect Team";


    GmailApp.sendEmail(email,
      "Your eSIM is ready",
      cuerpo,
      {
        attachments: [archivo.getBlob(), pdf1, pdf2],
        name: "Student Connect",
        from: "studentconnect@connectivityglobal.com",
        bcc: "support@connectivityglobal.com"
      }
    );
    Logger.log("Email eSIM enviado con QR y número: " + msisdn);
  } else {
    Logger.log("QR no encontrado para ICC: " + icc);
  }
}


function enviarQREmailSinNumero(name, email, icc, activationdate) {
  const carpetaID = '1uEQmtKU72a7i-2gaUVjVq5KI-ClYjcH-';
  const carpeta = DriveApp.getFolderById(carpetaID);
  const nombreArchivo = icc + '-qr.png';
  const archivos = carpeta.getFilesByName(nombreArchivo);

  if (archivos.hasNext()) {
    const archivo = archivos.next();
    const cuerpo = 
    "Hi " + name + ",\n\n" +
    "Your France eSIM has been pre-activated and on " + activationdate + " it will be fully activated with your purchased plan. You can find the QR code attached to this email.\n\n" +
    "Save the QR code for when your eSIM is ready to be installed!\n\n" +
    "If you have any questions, just reply to this email — we're here to help you.\n\n" +
    "Best regards,\n\n" +
    "Student Connect Team";

    GmailApp.sendEmail(email,
      "Your eSIM is ready",
      cuerpo,
      {
        from: "studentconnect@connectivityglobal.com",
        attachments: [archivo.getBlob()],
        name: "Student Connect",
        bcc: "support@connectivityglobal.com"
      }
    );
    Logger.log("Email eSIM enviado con QR");
  } else {
    Logger.log("QR no encontrado para ICC: " + icc);
  }
}


function reactivarSIMFrance(iAct) {
  const hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");

  const country = hojaActivar.getRange(iAct, 1).getValue();
  const msisdn = hojaActivar.getRange(iAct, 3).getValue();
  const sim_esim = hojaActivar.getRange(iAct, 5).getValue();

  if (country !== "france" || sim_esim !== "SIM") {
    Logger.log("No es Francia o no es SIM física. Se omite reactivación.");
    return "skip";
  }

  const body = {
    msisdn: msisdn
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  };

  try {
    const url = "http://207.154.252.56:3010/restore"; // 
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.status === "ok") {
      hojaActivar.getRange(iAct, 24).setValue(json.ticket); // columna X
      hojaActivar.getRange(iAct, 9).setValue("WIP");
      Logger.log("Reactivación solicitada. Ticket: " + json.ticket + " Estado: WIP");
      return "WIP";
    } else {
      hojaActivar.getRange(iAct, 9).setValue("Error");
      hojaActivar.getRange(iAct, 20).setValue(json.message);
      Logger.log("Error reactivación: " + json.message);
      return "Error";
    }
  } catch (error) {
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue(error.toString());
    Logger.log("Excepción reactivación: " + error);
    return "Error";
  }
}



function suspenderSIMFrance(iAct) {
  const hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");

  //const country = hojaActivar.getRange(iAct, 1).getValue();
  //const msisdn = hojaActivar.getRange(iAct, 3).getValue();
  const msisdn = "33" + String(hojaActivar.getRange(iAct, 3).getValue());
  //const sim_esim = hojaActivar.getRange(iAct, 5).getValue();

  //if (country !== "france" || sim_esim !== "SIM") {
  //  Logger.log("No es Francia o no es SIM física. Se omite suspensión.");
  //  return "skip";
  //}

  const body = {
    msisdn: msisdn
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  };

  try {
    const url = "http://207.154.252.56:3010/suspend"; 
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.status === "ok") {
      hojaActivar.getRange(iAct, 24).setValue(json.ticket); // columna X
      hojaActivar.getRange(iAct, 9).setValue("Suspendida"); // columna I
      Logger.log("Suspensión solicitada. Ticket: " + json.ticket);
      return "ok";
    } else {
      hojaActivar.getRange(iAct, 9).setValue("Error");
      hojaActivar.getRange(iAct, 20).setValue(json.message); // columna T
      Logger.log("Error suspensión: " + json.message);
      return "error";
    }
  } catch (error) {
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue(error.toString());
    Logger.log("Excepción suspensión: " + error);
    return "error";
  }
}


function cancelarSIMFrance(iAct) {
  const hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");

  const country = hojaActivar.getRange(iAct, 1).getValue();
  const msisdn = hojaActivar.getRange(iAct, 3).getValue();
  const sim_esim = hojaActivar.getRange(iAct, 5).getValue();

  //if (country !== "france" || sim_esim !== "SIM") {
  //  Logger.log("No es Francia o no es SIM física. Se omite cancelación.");
  //  return "skip";
  //}

  const body = {
    msisdn: msisdn
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  };

  try {
    const url = "http://207.154.252.56:3010/delete"; // Sustituye con tu IP real
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.status === "ok") {
      hojaActivar.getRange(iAct, 24).setValue(json.ticket); // columna X
      hojaActivar.getRange(iAct, 9).setValue("Baja"); // columna I
      Logger.log("Cancelación completada. Ticket: " + json.ticket);
      return "ok";
    } else {
      hojaActivar.getRange(iAct, 9).setValue("Error");
      hojaActivar.getRange(iAct, 20).setValue(json.message);
      Logger.log("Error al cancelar línea: " + json.message);
      return "error";
    }
  } catch (error) {
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue(error.toString());
    Logger.log("Excepción al cancelar línea: " + error);
    return "error";
  }
}



//*******************************************************************************************************************************************
//************       UK              ********************************************************************************************************
//*******************************************************************************************************************************************


function activarAPISIMuk (){

  //SIM en UK de momento es Manual
  
  //No podemos llamar directamente al actman por los contadores
  //actman();

  var hojaValidaciones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Validacion Didit");
  // restamos una posición a iVal porque ya se ha hecho el +1
  var iVal = hojaValidaciones.getRange(1,13).getValue() - 1;

  //ENVIAR EMAIL A LUKE PARA HACER MANUAL
  var destinatario = "studentconnect@connectivityglobal.com";
  var mensaje = "Hay una nueva SIM de UK Connect pendiente de activar. En la hoja Validacion Didit, mirar la fila: " + iVal;
  var asunto = "Student Connect - SIM UK - Activación Manual";
  var copiaOculta = "sistemas@connectivity.es";

  enviarEmail(destinatario, copiaOculta, mensaje, asunto);

  // no hacemos el +1 porque ya se ha hecho
  //hojaValidaciones.getRange(1,13).setValue(iVal+1);

  var respuesta = "ok";
  return respuesta;

}

function activarAPIeSIMuk(iAct) {
  const hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  
  // Datos fila activación
  let planOriginal = hojaActivar.getRange(iAct, 4).getValue();
  const name = hojaActivar.getRange(iAct, 11).getValue();
  const email = hojaActivar.getRange(iAct, 12).getValue();
  const fechacad = hojaActivar.getRange(iAct, 7).getValue();
  const activationDate = hojaActivar.getRange(iAct,6).getValue();


  Logger.log("planOriginal: " + planOriginal);

  // Traducir plan
  let item;
  if (planOriginal.endsWith("-M")) {
    item = "domp_20GB_ULSMS_ULMIN_1M_V1";
  } else if (planOriginal.endsWith("-L")) {
    item = "domp_50GB_ULSMS_ULMIN_1M_V1";
  } else if (planOriginal.endsWith("-XL")) {
    item = "domp_ULGB_ULSMS_ULMIN_1M_V1";
  } else if (planOriginal.endsWith("-S")) {
    item = "domp_1GB_ULSMS_ULMIN_1M_V1";
  } else {
    Logger.log("Plan no reconocido: " + planOriginal);
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue("Plan no reconocido");
    return "error";
  }

  Logger.log("item: " + item);
  
  // Body final
  const body = {
    "type": "transaction",
    "assign": true,
    "order": [
      {
        "type": "bundle",
        "quantity": 1,
        "item": item,
        "allowReassign": false
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    headers: {
      "X-API-Key": "hlL-26ORuHSlTmNE_E1tGTQNou6gNgqvlEdMZ_nI" // ⚠️ Tu API Key real
    },
    muteHttpExceptions: true
  };

  try {
    const url = "https://domestic-uk.api.esim-go.com/v2.5/orders"
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    // Usar variable intermedia para simplificar
    const esimData = 
      json && json.order && json.order[0] &&
      json.order[0].esims && json.order[0].esims[0];

    Logger.log("response.getContentText(): " + response.getContentText());
    Logger.log("esimData.iccid: " + esimData.iccid);
    Logger.log("esimData.smdpAddress: " + esimData.smdpAddress);
    Logger.log("esimData.matchingId: " + esimData.matchingId);

    if (esimData && esimData.iccid && esimData.smdpAddress && esimData.matchingId) {
      const icc = esimData.iccid;
      const smdpAddress = esimData.smdpAddress;
      const matchingID = esimData.matchingId;
      
      // Guardar ICCID
      hojaActivar.getRange(iAct, 2).setValue(icc);
      
      // Construir LPA
      const lpa = "LPA:1$" + smdpAddress + "$" + matchingID;
      hojaActivar.getRange(iAct, 24).setValue(lpa);
      
      Logger.log("Activación eSIM UK correcta. ICCID: " + icc + ". LPA: " + lpa);
      
      // Obtener MSISDN
      const msisdn = conseguirMsisdnUK(icc);
      hojaActivar.getRange(iAct, 3).setValue(msisdn);
      
      // Enviar QR
      enviarQRCodeUK(name, email, lpa, msisdn, iAct, fechacad, activationDate);
      
      // Marcar estado como Activa
      hojaActivar.getRange(iAct, 9).setValue("Activa");
      return "ok";
    } else {
      Logger.log("Error al activar: " + response.getContentText());
      hojaActivar.getRange(iAct, 9).setValue("Error");
      hojaActivar.getRange(iAct, 20).setValue(response.getContentText());
      return "error";
    }
  } catch (error) {
    Logger.log("Excepción: " + error);
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue(error.toString());
    return "error";
  }
}


function preActivarAPIeSIMuk(iAct) {
  const hojaActivar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  
  // Datos fila activación
  let planOriginal = hojaActivar.getRange(iAct, 4).getValue();
  const name = hojaActivar.getRange(iAct, 11).getValue();
  const email = hojaActivar.getRange(iAct, 12).getValue();
  const fechacad = hojaActivar.getRange(iAct, 7).getValue();
  const activationDate = hojaActivar.getRange(iAct,6).getValue();


  Logger.log("planOriginal: " + planOriginal);

  // Traducir plan. Forzamos el plan pequeño para la PreActivación
  item = "domp_1GB_ULSMS_ULMIN_1M_V1";

  Logger.log("item: " + item);
  
  // Body final
  const body = {
    "type": "transaction",
    "assign": true,
    "order": [
      {
        "type": "bundle",
        "quantity": 1,
        "item": item,
        "allowReassign": false
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    headers: {
      "X-API-Key": "hlL-26ORuHSlTmNE_E1tGTQNou6gNgqvlEdMZ_nI" // ⚠️ Tu API Key real
    },
    muteHttpExceptions: true
  };

  try {
    const url = "https://domestic-uk.api.esim-go.com/v2.5/orders"
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    // Usar variable intermedia para simplificar
    const esimData = 
      json && json.order && json.order[0] &&
      json.order[0].esims && json.order[0].esims[0];

    Logger.log("response.getContentText(): " + response.getContentText());
    Logger.log("esimData.iccid: " + esimData.iccid);
    Logger.log("esimData.smdpAddress: " + esimData.smdpAddress);
    Logger.log("esimData.matchingId: " + esimData.matchingId);

    if (esimData && esimData.iccid && esimData.smdpAddress && esimData.matchingId) {
      const icc = esimData.iccid;
      const smdpAddress = esimData.smdpAddress;
      const matchingID = esimData.matchingId;
      
      // Guardar ICCID
      hojaActivar.getRange(iAct, 2).setValue(icc);
      
      // Construir LPA
      const lpa = "LPA:1$" + smdpAddress + "$" + matchingID;
      hojaActivar.getRange(iAct, 24).setValue(lpa);
      
      Logger.log("Activación eSIM UK correcta. ICCID: " + icc + ". LPA: " + lpa);
      
      // Obtener MSISDN
      const msisdn = conseguirMsisdnUK(icc);
      hojaActivar.getRange(iAct, 3).setValue(msisdn);
      
      // Enviar QR
      enviarQRCodeUK(name, email, lpa, msisdn, iAct, fechacad, activationDate);
      
      // Marcar estado como Activa
      hojaActivar.getRange(iAct, 9).setValue("PreActivada");
      return "PreActivada";
    } else {
      Logger.log("Error al activar: " + response.getContentText());
      hojaActivar.getRange(iAct, 9).setValue("Error");
      hojaActivar.getRange(iAct, 20).setValue(response.getContentText());
      return "error";
    }
  } catch (error) {
    Logger.log("Excepción: " + error);
    hojaActivar.getRange(iAct, 9).setValue("Error");
    hojaActivar.getRange(iAct, 20).setValue(error.toString());
    return "error";
  }
}


function conseguirMsisdnUK(icc) {
  const url = "https://domestic-uk.api.esim-go.com/v2.5/esims/" + icc;

  const options = {
    method: "get",
    headers: {
      "X-API-Key": "hlL-26ORuHSlTmNE_E1tGTQNou6gNgqvlEdMZ_nI",
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.msisdn) {
      Logger.log("MSISDN obtenido: " + json.msisdn);
      return json.msisdn;
    } else {
      Logger.log("No se encontró msisdn en la respuesta");
      return null;
    }
  } catch (error) {
    Logger.log("Error en conseguirMsisdnUK: " + error);
    return null;
  }
}



function enviarQRCodeUK(name, email, lpa, msisdn, filaAactivar, fechacad, activationDate) {
  var encodedLPA = encodeURIComponent(lpa);
  var url = "https://quickchart.io/qr?text=" + encodedLPA + "&size=300";

  var hoyRaw = new Date();
  var hoy = Utilities.formatDate(hoyRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");

  var activationDateF = Utilities.formatDate(activationDate, Session.getScriptTimeZone(), "dd/MM/yyyy");

  var fechacadF = Utilities.formatDate(fechacad, Session.getScriptTimeZone(), "dd/MM/yyyy");

  if (activationDateF < hoy) {
    activationDateF = hoy;
  }
  

  var response = UrlFetchApp.fetch(url);
  var imageBlob = response.getBlob();
  imageBlob.setName("QRCode.png");

  var subject = "Connectivity - Your UK eSIM is ready!";

  if (activationDateF > hoy) {
    const message = 
      "Hi " + name + ",\n\n" +
      "Welcome to Student Connect!\n\n" +
      "Your UK eSIM has been pre-activated and on " + activationDateF + " it will be fully activated with your purchased plan. You can find the QR code attached to this email.\n\n" +
      "Your UK number is: " + msisdn + "\n\n" +
      "Note that you can only activate your eSIM in the UK!\n\n" +
      "If you have any questions, just reply to this email — we're here to help you.\n\n" +
      "Best regards,\n\n" +
      "Student Connect Team";

    var remitente = "studentconnect@connectivityglobal.com";
    //var remitente = "sistemas@connectivity.es";
    var copiaOculta = "soporte@connectivity.es";

    var opciones = {
      from: remitente,
      bcc: copiaOculta,
      attachments: [imageBlob]
    };

    GmailApp.sendEmail(email, subject, message, opciones);
  }
  else{

    // === PDFs adjuntos ===
    const pdf1 = DriveApp.getFileById("15-UohFy3bwHH8Ks1lYAH9IWhFn5k8KVI").getAs("application/pdf");
    const pdf2 = DriveApp.getFileById("1NhGyL04e1B5OVGTpCXs_hetRDhVkQ5hc").getAs("application/pdf");
    
    const message = 
      "Hi " + name + ",\n\n" +
      "Welcome to Student Connect!\n\n" +
      "In this email you have the QR code, the data and instructions to install your eSIM on iPhone and Android devices.\n\n" +
      "Your UK number is: " + msisdn + "\n\n" +
      "Your plan will be active until the date: " + fechacadF + "\n\n" +
      "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
      "Note that you can only activate your eSIM in the UK!\n\n" +
      "If you have any questions, just reply to this email — we're here to help you.\n\n" +
      "Enjoy your plan!\n\n" +
      "Student Connect Team";

    var remitente = "studentconnect@connectivityglobal.com";
    //var remitente = "sistemas@connectivity.es";
    var copiaOculta = "soporte@connectivity.es";

    var opciones = {
      from: remitente,
      bcc: copiaOculta,
      attachments: [imageBlob, pdf1, pdf2]
    };

    GmailApp.sendEmail(email, subject, message, opciones);
  }
  Logger.log("Correo enviado a: " + email);
}



/*
//Nueva funcion NICO para enviar email bien, pdfs adjuntos, etc. De momento la quito porque debe haber algun error al enviarse desde un dopost, con los archivos, etc.

function enviarQRCodeUK(name, email, lpa, msisdn, filaAactivar, fechacad) {
  var encodedLPA = encodeURIComponent(lpa);
  var url = "https://quickchart.io/qr?text=" + encodedLPA + "&size=300";

  var response = UrlFetchApp.fetch(url);
  var imageBlob = response.getBlob();
  imageBlob.setName("QRCode.png");

  // Formatear fecha de caducidad si es Date
  var expiryDate = "";
  if (fechacad instanceof Date) {
    var day = ("0" + fechacad.getDate()).slice(-2);
    var month = ("0" + (fechacad.getMonth() + 1)).slice(-2);
    var year = fechacad.getFullYear();
    expiryDate = day + "/" + month + "/" + year;
  } else {
    expiryDate = fechacad;
  }

  var subject = "Your UK eSIM is now active!";

  var message =
    "Hi " + name + ",\n\n" +
    "Welcome to Student Connect!!! In this email you have the QR code, the data and instructions to install your eSIM on iPhone and Android devices.\n\n" +
    "Your new UK number is: " + msisdn + "\n" +
    "Your plan will be active until the date: " + expiryDate + "\n" +
    "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
    "Note that you can only activate your eSIM in the UK!\n\n" +
    "If you have any questions, just reply to this email — we're here to help you.\n\n" +
    "Enjoy your plan!\n\n" +
    "Best regards,\n" +
    "The Student Connect Team";

  // Adjuntar PDFs desde Google Drive
  var pdf1 = DriveApp.getFileById("1yRgptWW8afOcAJtVv3lyGh-rcK8cCN8W").getAs("application/pdf");
  var pdf2 = DriveApp.getFileById("1tyH9HDEFV7-OGUk8pNH3qVpN7f4ULiuL").getAs("application/pdf");

  var remitente = "sistemas@connectivity.es";
  var copiaOculta = "soporte@connectivity.es";

  var opciones = {
    htmlBody: message.replace(/\n/g, "<br>"), // Para que el email respete saltos de línea
    from: remitente,
    bcc: copiaOculta,
    attachments: [imageBlob, pdf1, pdf2]
  };

  GmailApp.sendEmail(email, subject, "", opciones);

  Logger.log("Correo enviado a: " + email);
}
*/


function activarPreActAPIeSIMuk(iAct) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");
  const iccid = hoja.getRange(iAct, 2).getValue();
  const plan = hoja.getRange(iAct, 4).getValue();
  const msisdn = hoja.getRange(iAct, 3).getValue();
  const expiryDateRaw = hoja.getRange(iAct, 7).getValue();
  const nameclient = hoja.getRange(iAct,11).getValue();

  const expiryDate = Utilities.formatDate(expiryDateRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");

  let name;
  if (plan.endsWith("-M")) {
    name = "domp_20GB_ULSMS_ULMIN_1M_V1";
  } else if (plan.endsWith("-L")) {
    name = "domp_50GB_ULSMS_ULMIN_1M_V1";
  } else if (plan.endsWith("-XL")) {
    name = "domp_ULGB_ULSMS_ULMIN_1M_V1";
  } else {
    hoja.getRange(iAct, 9).setValue("Error");
    hoja.getRange(iAct, 20).setValue("Plan no reconocido");
    return "error";
  }

  const apiKey = "hlL-26ORuHSlTmNE_E1tGTQNou6gNgqvlEdMZ_nI";
  const deleteUrl = `https://domestic-uk.api.esim-go.com/v2.5/esims/${iccid}/bundles/domp_1GB_ULSMS_ULMIN_1M_V1`;
  const applyUrl = "https://domestic-uk.api.esim-go.com/v2.5/esims/apply";

  try {
    UrlFetchApp.fetch(deleteUrl, {
      method: "delete",
      headers: { "X-API-Key": apiKey }
    });

    const payload = {
      iccid: iccid,
      name: name,
      allowReassign: false,
      repeat: 1
    };

    const response = UrlFetchApp.fetch(applyUrl, {
      method: "post",
      contentType: "application/json",
      headers: { "X-API-Key": apiKey },
      payload: JSON.stringify(payload)
    });

    Logger.log("Apply status: " + response.getResponseCode());
    Logger.log("Apply body: " + response.getContentText());

    const json = JSON.parse(response.getContentText());
    const esimResult = json.esims && json.esims[0];

    if (esimResult && esimResult.status === "Successfully Applied Bundle" && esimResult.bundle === name) {
      hoja.getRange(iAct, 9).setValue("Activa");

      // Email
        // === PDFs adjuntos ===
        const pdf1 = DriveApp.getFileById("15-UohFy3bwHH8Ks1lYAH9IWhFn5k8KVI").getAs("application/pdf");
        const pdf2 = DriveApp.getFileById("1NhGyL04e1B5OVGTpCXs_hetRDhVkQ5hc").getAs("application/pdf");

        // === Mensaje del email ===
        const asunto = "Your UK eSIM is now active!";
        const mensaje = 
          "Hi " + nameclient + ",\n\n" +
          "Your eSIM has now been fully activated! Attached you have our UK manuals to assist you with your eSIM.\n\n" +
          "Your UK number is: " + msisdn + "\n" +
          "Your plan will be active until the date: " + expiryDate + "\n\n" +
          "We'll remind you of your expiration date, and you can easily renew your subscription to stay connected to our service!\n\n" +
          "Note that you can only activate your eSIM in the UK!\n\n" +
          "If you have any questions, just reply to this email — we're here to help you.\n\n" +
          "Enjoy your plan!\n\n" +
          "Student Connect Team";


        // === Envío con adjuntos ===
        GmailApp.sendEmail(email, asunto, mensaje, {
          from: "studentconnect@connectivityglobal.com",
          attachments: [pdf1, pdf2],
          bcc: "support@connectivityglobal.com",
        });



      return "ok";
    } else {
      hoja.getRange(iAct, 9).setValue("Error");
      hoja.getRange(iAct, 20).setValue(json.message || "Error en activación apply");
      return "error";
    }
  } catch (err) {
    hoja.getRange(iAct, 9).setValue("Error");
    hoja.getRange(iAct, 20).setValue(err.message);
    return "error";
  }
}



function suspenderSIMuk(iAct) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activaciones");

  try {
    const country  = String(hoja.getRange(iAct, 1).getValue() || "").toLowerCase(); // A
    const iccid    = String(hoja.getRange(iAct, 2).getValue() || "");               // B
    const planOrig = String(hoja.getRange(iAct, 4).getValue() || "");               // D
    const tipo     = String(hoja.getRange(iAct, 5).getValue() || "");               // E

    if (country !== "uk") {
      Logger.log("suspenderSIMuk: fila " + iAct + " no es UK. Se omite.");
      return "skip";
    }
    if (!iccid) {
      Logger.log("suspenderSIMuk: ICCID vacío en fila " + iAct);
      hoja.getRange(iAct, 20).setValue("Error suspensión UK: ICCID vacío");
      enviarAvisoManual("UK", iAct, "ICCID vacío");
      hoja.getRange(iAct, 9).setValue("Error");
      return "error";
    }

    // --- 1) Resolver el nombre EXACTO del bundle (igual que en activarAPIeSIMuk) ---
    // -M => 20GB, -L => 50GB, -XL => Unlimited, -S => 1GB
    let bundleName;
    if (planOrig.endsWith("-M")) {
      bundleName = "domp_20GB_ULSMS_ULMIN_1M_V1";
    } else if (planOrig.endsWith("-L")) {
      bundleName = "domp_50GB_ULSMS_ULMIN_1M_V1";
    } else if (planOrig.endsWith("-XL")) {
      bundleName = "domp_ULGB_ULSMS_ULMIN_1M_V1";
    } else if (planOrig.endsWith("-S")) {
      bundleName = "domp_1GB_ULSMS_ULMIN_1M_V1";
    } else {
      const msg = "Plan UK no reconocido para suspensión: " + planOrig;
      Logger.log("suspenderSIMuk: " + msg);
      hoja.getRange(iAct, 20).setValue(msg);
      enviarAvisoManual("UK", iAct, msg);
      hoja.getRange(iAct, 9).setValue("Error");
      return "error";
    }

    // --- 2) GET assignment (para obtener assignmentId de la ÚLTIMA asignación de ese bundle) ---
    const apiKey = "hlL-26ORuHSlTmNE_E1tGTQNou6gNgqvlEdMZ_nI"; // ← tu API key ya usada en UK
    const base   = "https://domestic-uk.api.esim-go.com/v2.5";
    const getUrl = base + "/esims/" + encodeURIComponent(iccid) + "/bundles/" + encodeURIComponent(bundleName);

    let getResp, getJson;
    try {
      getResp = UrlFetchApp.fetch(getUrl, {
        method: "get",
        headers: {
          "X-API-Key": apiKey,
          "Content-Type": "application/json",
          "Accept": "application/json"
        },
        muteHttpExceptions: true
      });
      getJson = JSON.parse(getResp.getContentText() || "{}");
    } catch (e) {
      hoja.getRange(iAct, 20).setValue("Error GET bundle UK: " + e.message);
      enviarAvisoManual("UK", iAct, "Error consultando bundle: " + e.message);
      hoja.getRange(iAct, 9).setValue("Error");
      return "error";
    }

    // El endpoint devuelve un objeto con "assignments" (tomamos la última o la 1ª si ya viene filtrada)
    const assignments = (getJson && getJson.assignments) ? getJson.assignments : [];
    if (!assignments.length) {
      const msg = "No se encontró assignment para el bundle '" + bundleName + "' en ICCID " + iccid;
      Logger.log("suspenderSIMuk: " + msg);
      hoja.getRange(iAct, 20).setValue(msg);
      enviarAvisoManual("UK", iAct, msg);
      hoja.getRange(iAct, 9).setValue("Error");
      return "error";
    }

    // El doc habla de "assignmentId" en el DELETE; en la respuesta del GET aparece "assignmentReference"
    // (y algunos esquemas incluyen 'id'). Probamos primero 'assignmentReference' y si no, 'id'.
    const latest = assignments[assignments.length - 1];
    const assignmentId = latest.assignmentReference || latest.id;
    if (!assignmentId) {
      const msg = "No se pudo determinar assignmentId (assignmentReference/id) en respuesta GET.";
      Logger.log("suspenderSIMuk: " + msg + " Respuesta: " + getResp.getContentText());
      hoja.getRange(iAct, 20).setValue(msg);
      enviarAvisoManual("UK", iAct, msg);
      hoja.getRange(iAct, 9).setValue("Error");
      return "error";
    }

    // --- 3) DELETE (revocar assignment) ---
    const delUrl = base
      + "/esims/" + encodeURIComponent(iccid)
      + "/bundles/" + encodeURIComponent(bundleName)
      + "/assignments/" + encodeURIComponent(assignmentId)
      + "?type=transaction"; // ejecuta la revocación (no solo validar)

    let delResp, delJson;
    try {
      delResp = UrlFetchApp.fetch(delUrl, {
        method: "delete",
        headers: {
          "X-API-Key": apiKey,
          "Content-Type": "application/json",
          "Accept": "application/json"
        },
        muteHttpExceptions: true
      });
      delJson = JSON.parse(delResp.getContentText() || "{}");
    } catch (e) {
      hoja.getRange(iAct, 20).setValue("Error DELETE bundle UK: " + e.message);
      enviarAvisoManual("UK", iAct, "Error revocando bundle: " + e.message);
      hoja.getRange(iAct, 9).setValue("Error");
      return "error";
    }

    // Respuesta “OK” típica: { status: "..." }
    const code = delResp.getResponseCode();
    if (code >= 200 && code < 300) {
      hoja.getRange(iAct, 9).setValue("Suspendida"); // Col I (Estado)
      hoja.getRange(iAct, 20).setValue("UK suspendida OK. assignmentId=" + assignmentId);
      Logger.log("suspenderSIMuk: OK. ICCID " + iccid + " bundle " + bundleName + " assignmentId " + assignmentId);
      return "ok";
    } else {
      const bodyTxt = delResp.getContentText();
      const msg = "DELETE UK no-2xx (" + code + "): " + bodyTxt;
      hoja.getRange(iAct, 20).setValue(msg);
      enviarAvisoManual("UK", iAct, msg);
      hoja.getRange(iAct, 9).setValue("Error");
      return "error";
    }

  } catch (e) {
    Logger.log("suspenderSIMuk: excepción en fila " + iAct + " => " + e.message);
    try { hoja.getRange(iAct, 20).setValue("Excepción suspensión UK: " + e.message); } catch (_) {}
    enviarAvisoManual("UK", iAct, e.message);
    hoja.getRange(iAct, 9).setValue("Error");
    return "error";
  }
}



/**
 * Solicita a eSIM Go el código de portabilidad (PAC/STAC) para un ICCID.
 * - Lee el ICCID de UK desde: hoja "UK eSIM Go Portabilidad", celda C3
 * - Escribe el portCode devuelto en E3 y el raw response en F3 (debug).
 * - Devuelve el portCode si todo va bien; si no, devuelve "error".
 */
function requestPACuk() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UK eSIM Go Portabilidad");
  if (!hoja) {
    Logger.log('No existe la hoja "UK eSIM Go Portabilidad".');
    return "error";
  }

  const iccid = String(hoja.getRange("C3").getValue()).trim();
  const portType = "PAC"; // Por defecto PAC

  if (!iccid) {
    Logger.log("ICC vacío en C3.");
    hoja.getRange("C5").setValue("error");
    hoja.getRange("C6").setValue("ICC vacío en C3");
    return "error";
  }

  const url = `https://domestic-uk.api.esim-go.com/v2.5/esims/${iccid}/porting/out?portType=${portType}`;

  const options = {
    method: "post",
    headers: {
      "X-API-Key": "hlL-26ORuHSlTmNE_E1tGTQNou6gNgqvlEdMZ_nI", // ⚠️ Sustituye si usas otra key
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };

  try {
    const resp = UrlFetchApp.fetch(url, options);
    const status = resp.getResponseCode();
    const body = resp.getContentText();
    Logger.log("requestPCuk status: " + status);
    Logger.log("requestPCuk raw: " + body);

    let json;
    try { json = JSON.parse(body); } catch (e) { json = {}; }

    const portCode = json && (json.portCode || json.portcode);
    if (status >= 200 && status < 300 && portCode) {
      hoja.getRange("C5").setValue(portCode);
      hoja.getRange("C6").clearContent();
      Logger.log("Porting out OK. portType=" + portType + " | portCode=" + portCode);
      return portCode;
    } else {
      hoja.getRange("C5").setValue("error");
      hoja.getRange("C6").setValue(body || ("HTTP " + status));
      Logger.log("Fallo en porting out. HTTP " + status);
      return "error";
    }
  } catch (err) {
    hoja.getRange("C5").setValue("error");
    hoja.getRange("C6").setValue(String(err));
    Logger.log("Excepción requestPCuk: " + err);
    return "error";
  }
}



//*******************************************************************************************************************************************
//************       MAILCHIMP       ********************************************************************************************************
//*******************************************************************************************************************************************


/***********************
 * CONFIG & HELPERS
 ***********************/
function getMailchimpConfig_() {
  const props = PropertiesService.getScriptProperties();
  const API_KEY = props.getProperty('MAILCHIMP_API_KEY');   // p.ej. 123abc-us21
  const LIST_ID = props.getProperty('MAILCHIMP_LIST_ID');   // Audience ID
  if (!API_KEY || !LIST_ID) {
    throw new Error('Faltan MAILCHIMP_API_KEY o MAILCHIMP_LIST_ID en Script Properties.');
  }
  const dcParts = API_KEY.split('-');
  if (dcParts.length < 2) throw new Error('API Key inválida. Debe terminar en -usXX (datacenter).');
  const DC = dcParts[1];
  return { API_KEY, LIST_ID, DC };
}

function md5Lower_(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(str).trim().toLowerCase());
  return raw.map(b => (b + 256) % 256).map(b => b.toString(16).padStart(2, '0')).join('');
}

function mcRequest_(method, path, bodyObj) {
  const { API_KEY, DC } = getMailchimpConfig_();
  const url = `https://${DC}.api.mailchimp.com/3.0${path}`;
  const options = {
    method,
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode('any:' + API_KEY),
      'Content-Type': 'application/json'
    }
  };
  if (bodyObj !== undefined) options.payload = JSON.stringify(bodyObj);
  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const text = resp.getContentText() || '';
  let json = {};
  try { json = text ? JSON.parse(text) : {}; } catch (e) {}
  const headersOut = resp.getAllHeaders && resp.getAllHeaders();
  return { code, json, text, headersOut, url, method };
}

//function getLogSheet_() {
//  const ss = SpreadsheetApp.getActiveSpreadsheet();
//  const sh = ss.getSheetByName('Mailchimp Sync Logs') || ss.insertSheet('Mailchimp Sync Logs');
//  if (sh.getLastRow() === 0) {
//    sh.appendRow(['Timestamp', 'Email', 'Acción', 'Resultado/Detalle']);
//  }
//  return sh;
//}
function getLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Mailchimp Sync Logs') || ss.insertSheet('Mailchimp Sync Logs');

  // Asegura 4 columnas (A:D). Si hay más, las elimina.
  const maxCols = sh.getMaxColumns();
  if (maxCols > 4) sh.deleteColumns(5, maxCols - 4);

  // (Opcional) limita filas para no crecer infinito: deja 50.000 filas máximo
  const maxRows = sh.getMaxRows();
  if (maxRows > 50000) sh.deleteRows(50001, maxRows - 50000);

  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Email', 'Acción', 'Resultado/Detalle']);
  }
  return sh;
}


function formatDateForMailchimp(v) {
  if (!v) return null;                          // para limpiar si está vacío
  // Si ya es Date
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    const mm = String(v.getMonth() + 1).padStart(2, '0');
    const dd = String(v.getDate()).padStart(2, '0');
    const yyyy = v.getFullYear();
    return `${mm}/${dd}/${yyyy}`;               // MM/DD/YYYY
  }
  // Si viene como string
  const s = String(v).trim();
  if (!s) return null;

  // dd/MM/yyyy -> MM/DD/YYYY
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m1) {
    const dd = m1[1].padStart(2, '0');
    const mm = m1[2].padStart(2, '0');
    const yyyy = m1[3];
    // detecta si ya está en MM/DD o en DD/MM según tu hoja; asumimos DD/MM -> invertimos
    // Si tu hoja ya trae MM/DD, comenta las dos líneas siguientes:
    return `${mm}/${dd}/${yyyy}`;               // MM/DD/YYYY
  }

  // yyyy-MM-dd -> MM/DD/YYYY
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) {
    const yyyy = m2[1], mm = m2[2], dd = m2[3];
    return `${mm}/${dd}/${yyyy}`;
  }

  // último intento: que Date lo parsee
  const d = new Date(s);
  if (!isNaN(d)) {
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
  }
  return null; // si nada funcionó
}


/****************************************
 * LECTURA DE "Mailchimp Export" (A:K)
 * Columnas esperadas:
 * A email | B Cliente | C Country | D Producto | E Soporte
 * F Fecha | G Fecha Caducidad | H Meses | I Estado | J Numero | K ICC
 ****************************************/
function readExportRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Mailchimp Export');
  if (!sh) throw new Error('No existe la hoja "Mailchimp Export".');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // Lee todas las filas y devuelve objetos con claves claras
  const values = sh.getRange(2, 1, lastRow - 1, 12).getValues();
  return values
    .map(r => ({
      email: (r[0] || '').toString().trim().toLowerCase(),
      cliente: r[1] || '',
      country: r[2] || '',
      producto: r[3] || '',
      soporte: r[4] || '',
      fecha: r[5] || '',
      fechaCad: r[6] || '',
      meses: r[7] || '',
      estado: r[8] || '',
      numero: r[9] || '',
      icc: r[10] || '',
      aviso15: r[11] || ''
    }))
    .filter(o => o.email); // sin email no se sincroniza
}

/****************************************
 * UPSERT a Mailchimp (crea o actualiza)
 ****************************************/
function upsertFromExport_() {
  const logSh = getLogSheet_();
  const rows = readExportRows_();
  const { LIST_ID } = getMailchimpConfig_();

  const currentEmails = new Set();

  rows.forEach((r, idx) => {
    const email = r.email;
    currentEmails.add(email);

    const subscriber_hash = md5Lower_(email);

    const fechaMC = formatDateForMailchimp(r.fecha);
    const fechaCadMC = formatDateForMailchimp(r.fechaCad);

    // Meses es Number en Mailchimp: manda número o null
    const mesesNum = (r.meses === '' || r.meses === null || r.meses === undefined)
      ? null
      : Number(r.meses);

    const payload = {
      email_address: email,
      status_if_new: 'subscribed',
      status: 'subscribed',
      merge_fields: {
        MMERGE11: String(r.cliente || ''),   // Cliente (text)
        MMERGE13: String(r.country || ''),   // Country (text)
        MMERGE17: String(r.producto || ''),  // Producto (text)
        MMERGE1:  String(r.soporte || ''),   // Soporte (text)
        MMERGE2:  fechaMC,                   // Fecha (date MM/DD/YYYY o null)
        MMERGE3:  fechaCadMC,                // Fecha Caducidad (date MM/DD/YYYY o null)
        MMERGE4:  mesesNum,                  // Meses (number o null)
        MMERGE5:  String(r.estado || ''),    // Estado (text)
        MMERGE6:  String(r.numero || ''),    // Número (text)
        MMERGE8:  String(r.icc || ''),        // ICC (text)
        AVISO15:  String(r.aviso15 || '')
      }
    };

    const { code, json, text } = mcRequest_('PUT', `/lists/${LIST_ID}/members/${subscriber_hash}`, payload);

    if (code >= 200 && code < 300) {
      logSh.appendRow([new Date(), email, 'UPSERT', 'OK (subscribed)']);
    } else {
      // Si está unsubscribed/cleaned, Mailchimp no permite resuscribir por API (compliance)
      const detail = (json && json.detail) ? json.detail : text;
      logSh.appendRow([new Date(), email, 'UPSERT', `ERROR (${code}) ${detail}`]);
    }

    // Pequeño respiro si hay muchas filas (evita 429)
    if ((idx + 1) % 400 === 0) Utilities.sleep(1000);
  });

}

/***********************************************************
 * ARCHIVADO: saca del billing lo que no está en la export
 ***********************************************************/

function archiveMembersNotInExport_() {
  const { LIST_ID } = getMailchimpConfig_();
  const logSh = getLogSheet_();

  // 1) Emails vigentes desde la hoja (normalizados)
  const currentEmails = new Set(
    readExportRows_().map(r => String(r.email || '').trim().toLowerCase())
  );

  let offset = 0;
  const count = 1000;
  let totalArchived = 0;

  while (true) {
    // 2) Listar miembros (paginado)
    const listResp = mcRequest_(
      'GET',
      `/lists/${LIST_ID}/members?count=${count}&offset=${offset}&fields=members.email_address,total_items`
    );
    if (listResp.code !== 200) {
      logSh.appendRow([new Date(), '', 'LISTAR', `ERROR (${listResp.code}) ${listResp.text}`]);
      break;
    }

    const members = listResp.json?.members || [];
    if (!members.length) break;

    // 3) Archivar (DELETE) los que NO están en la hoja
    for (let i = 0; i < members.length; i++) {
      const em = String(members[i].email_address || '').trim().toLowerCase();
      if (em && !currentEmails.has(em)) {
        const hash = md5Lower_(em);

        // DELETE /lists/{list_id}/members/{subscriber_hash}  -> ARCHIVE
        const res = mcRequest_('DELETE', `/lists/${LIST_ID}/members/${hash}`);

        if (res.code >= 200 && res.code < 300) {
          totalArchived++;
          logSh.appendRow([new Date(), em, 'ARCHIVE', 'OK']);
        } else {
          const detail = res.json?.detail || res.text || '';
          logSh.appendRow([new Date(), em, 'ARCHIVE', `ERROR (${res.code}) ${detail}`]);
        }

        // Pequeña pausa anti rate-limit cada 400
        if ((i + 1) % 400 === 0) Utilities.sleep(1000);
      }
    }

    offset += members.length;
    if (members.length < count) break;
  }

  logSh.appendRow([new Date(), '', 'RESUMEN', `Archivados: ${totalArchived}`]);
}

/*****************************************
 * ORQUESTA + TRIGGER DIARIO
 *****************************************/
function dailyMailchimpSync() {
  // 1) Upsert de los que están en "Mailchimp Export"
  upsertFromExport_();
  // 2) Archivar los que ya no están (reduce billing)
  archiveMembersNotInExport_();
}


function debugGetMemberByEmail_(email) {
  const { LIST_ID } = getMailchimpConfig_();
  const hash = md5Lower_(email);
  const { code, json, text } = mcRequest_('GET', `/lists/${LIST_ID}/members/${hash}`, undefined);
  Logger.log('CODE: ' + code);
  Logger.log('MERGE_FIELDS: ' + JSON.stringify(json?.merge_fields, null, 2));
}

function debugListMergeFields() {
  const { LIST_ID } = getMailchimpConfig_();
  const { code, json } = mcRequest_('GET', `/lists/${LIST_ID}/merge-fields?count=100`, undefined);
  Logger.log('CODE: ' + code);
  Logger.log(JSON.stringify(json, null, 2));
}

/**
 * Archiva un contacto concreto en Mailchimp
 */
function archiveOne_TEST() {
  const email = "d.ferreon@call2world.es";  // ← pon aquí el email a archivar
  const subscriberHash = md5Lower_(email);
  const { LIST_ID } = getMailchimpConfig_();
  const resp = mcRequest_("POST", `/lists/${LIST_ID}/members/${subscriberHash}/actions/archive`);
  Logger.log(resp);
}


/**
 * Comprueba si un email está en el set actual de Mailchimp (audience)
 */
function debugInCurrentSet_TEST() {
  const email = "d.ferreon@call2world.es";  // ← pon aquí el email a comprobar
  const subscriberHash = md5Lower_(email);
  const { LIST_ID } = getMailchimpConfig_();
  const resp = mcRequest_("GET", `/lists/${LIST_ID}/members/${subscriberHash}`);
  Logger.log(resp);
}



