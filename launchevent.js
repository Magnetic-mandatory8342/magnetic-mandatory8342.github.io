/*
 * Aviso de desconexión laboral - complemento de Outlook (Smart Alerts / OnMessageSend)
 * Agrega el aviso al inicio del correo cuando se envía fuera del horario laboral:
 *   - Antes de las 7:45 a.m. o después de las 5:15 p.m. (hora del equipo)
 *   - Sábados y domingos (todo el día)
 *
 * Se escribe con sintaxis sencilla (var, function, indexOf) para que también
 * funcione en el motor JavaScript de Outlook clásico para Windows.
 */

// ====== CONFIGURACIÓN ======
var HORA_INICIO = { h: 7, m: 45 };   // 7:45 a.m.
var HORA_FIN = { h: 10, m: 15 };     // 10:15 a.m.

// Fragmento sin tildes para detectar si el aviso ya está (evita duplicados en respuestas/reenvíos
// y evita doble aviso si el equipo también tiene la macro VBA).
var MARCADOR = "tu jornada laboral y tu derecho";

var AVISO_HTML =
  "<p style='color:#666666;font-size:10pt;font-family:Calibri, Arial, sans-serif;'>" +
  "<b>La empresa respeta tu jornada laboral y tu derecho a la desconexión.</b><br>" +
  "Cualquier envío de correos fuera de este horario es voluntario y bajo responsabilidad del colaborador." +
  "</p><br>";
// ===========================

function esFueraDeHorario(fecha) {
  var dia = fecha.getDay(); // 0 = domingo, 6 = sábado
  if (dia === 0 || dia === 6) {
    return true;
  }
  var minutos = fecha.getHours() * 60 + fecha.getMinutes();
  var inicio = HORA_INICIO.h * 60 + HORA_INICIO.m;
  var fin = HORA_FIN.h * 60 + HORA_FIN.m;
  return minutos < inicio || minutos > fin;
}

function onMessageSendHandler(event) {
  // Dentro del horario: se deja enviar sin tocar nada.
  if (!esFueraDeHorario(new Date())) {
    event.completed({ allowEvent: true });
    return;
  }

  var item = Office.context.mailbox.item;

  item.body.getAsync(Office.CoercionType.Html, function (resultado) {
    // Si no se pudo leer el cuerpo, no se bloquea el envío.
    if (resultado.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true });
      return;
    }

    var cuerpo = resultado.value || "";
    if (cuerpo.toLowerCase().indexOf(MARCADOR) !== -1) {
      event.completed({ allowEvent: true }); // ya tiene el aviso
      return;
    }

    item.body.prependAsync(
      AVISO_HTML,
      { coercionType: Office.CoercionType.Html },
      function () {
        // Se permite el envío aunque el prepend falle, para no frenar el correo.
        event.completed({ allowEvent: true });
      }
    );
  });
}

// Conecta el nombre declarado en el manifiesto con la función.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
