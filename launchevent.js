/*
 * Doble check de envío — complemento Smart Alerts para Outlook (nuevo) y Outlook web.
 *
 *  - Solo destinatarios internos  -> 1 confirmación.
 *  - Algún destinatario externo   -> 2 confirmaciones.
 *
 * Este archivo se carga desde commands.html (Outlook nuevo / web / Mac) y directamente
 * como JavaScript en Outlook clásico. Por eso NO usa imports, ni DOM, ni async/await.
 */

// ---------------------------------------------------------------------------
// CONFIGURACIÓN
// ---------------------------------------------------------------------------
var DC_CONFIG = {
  // Dominios que se consideran internos.
  internalDomains: ["creacontrol.es"],

  // true: "algo.creacontrol.es" también cuenta como interno.
  includeSubdomains: false,

  // false: en Outlook clásico de Windows el complemento no hace nada (ahí ya actúa el script VBA).
  // Ponlo a true si retiras el VBA y quieres el mismo comportamiento en los dos Outlook.
  runInClassicOutlook: false,

  // true: al pulsar "Confirmar externos" se relanza el envío y aparece sola la 2ª confirmación.
  // false: tras la 1ª confirmación hay que volver a pulsar Enviar para ver la 2ª.
  autoResend: true,

  // Tiempo máximo entre la 1ª y la 2ª confirmación. Pasado este tiempo se vuelve a empezar.
  stageTtlMs: 120000,

  // Límite de Outlook para el texto del aviso.
  maxMessageChars: 500
};

var DC_STAGE_KEY = "dcExternalStage";
var DC_COMMAND_ID = "dcConfirmButton"; // id del <Control> en manifest.xml
var DC_CONFIRM_LABEL = "Confirmar externos"; // máx. 20 caracteres

// ---------------------------------------------------------------------------
// CLASIFICACIÓN DE DESTINATARIOS
// ---------------------------------------------------------------------------
function dcDomainOf(address) {
  var a = String(address || "").replace(/^\s+|\s+$/g, "").toLowerCase();
  var at = a.lastIndexOf("@");
  if (at < 1 || at === a.length - 1) {
    return "";
  }
  return a.substring(at + 1);
}

function dcIsInternal(address) {
  var domain = dcDomainOf(address);
  if (!domain) {
    return false; // sin dirección reconocible -> se trata como externo (lo más prudente)
  }
  for (var i = 0; i < DC_CONFIG.internalDomains.length; i++) {
    var internal = String(DC_CONFIG.internalDomains[i]).toLowerCase();
    if (domain === internal) {
      return true;
    }
    if (DC_CONFIG.includeSubdomains) {
      var suffix = "." + internal;
      if (domain.length > suffix.length && domain.slice(-suffix.length) === suffix) {
        return true;
      }
    }
  }
  return false;
}

// Lee Para, CC y CCO (en ese orden) y devuelve [{address, name, field}].
function dcGetRecipients(item, done) {
  var fields = [
    ["to", "Para"],
    ["cc", "CC"],
    ["bcc", "CCO"]
  ];
  var all = [];

  function next(index) {
    if (index >= fields.length) {
      done(null, all);
      return;
    }
    var field = item[fields[index][0]];
    if (!field || typeof field.getAsync !== "function") {
      next(index + 1);
      return;
    }
    field.getAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        done(result.error || { message: "No se pudo leer el campo " + fields[index][1] }, all);
        return;
      }
      var values = result.value || [];
      for (var i = 0; i < values.length; i++) {
        all.push({
          address: String(values[i].emailAddress || "").replace(/^\s+|\s+$/g, ""),
          name: String(values[i].displayName || ""),
          field: fields[index][1]
        });
      }
      next(index + 1);
    });
  }

  next(0);
}

// Separa internos/externos (sin duplicados) y calcula una "firma" de la lista de destinatarios.
function dcAnalyze(recipients) {
  var seen = {};
  var keys = [];
  var internal = [];
  var external = [];

  for (var i = 0; i < recipients.length; i++) {
    var r = recipients[i];
    var key = r.address ? r.address.toLowerCase() : "?" + r.name.toLowerCase();
    if (seen[key]) {
      continue;
    }
    seen[key] = true;
    keys.push(key);

    var label = r.address || (r.name ? r.name + " (sin dirección)" : "(destinatario sin dirección)");
    if (r.field !== "Para") {
      label += " (" + r.field + ")";
    }
    if (dcIsInternal(r.address)) {
      internal.push(label);
    } else {
      external.push(label);
    }
  }

  keys.sort();
  return {
    internal: internal,
    external: external,
    total: internal.length + external.length,
    signature: keys.join(";")
  };
}

// ---------------------------------------------------------------------------
// TEXTO DE LOS AVISOS (máx. 500 caracteres)
// ---------------------------------------------------------------------------
function dcEscapeMarkdown(text) {
  return String(text).replace(/([\\*_\[\]])/g, "\\$1");
}

function dcPlural(n, singular, plural) {
  return n + " " + (n === 1 ? singular : plural);
}

// Monta cabecera + lista + pie sin pasar del límite; si no caben todos, añade "… y N más".
function dcBuildMessage(head, items, tail, markdown) {
  var budget = DC_CONFIG.maxMessageChars - head.length - tail.length;
  var parts = [];
  var used = 0;
  var shown = 0;

  for (var i = 0; i < items.length; i++) {
    var text = markdown ? "- " + dcEscapeMarkdown(items[i]) + "\r" : (i > 0 ? ", " : "") + items[i];
    var remaining = items.length - (i + 1);
    var reserve = remaining > 0 ? 24 : 0; // hueco para "… y N más"
    if (used + text.length + reserve > budget) {
      break;
    }
    parts.push(text);
    used += text.length;
    shown++;
  }

  var hidden = items.length - shown;
  if (hidden > 0) {
    var more = shown === 0 ? dcPlural(hidden, "destinatario", "destinatarios") : "… y " + hidden + " más";
    parts.push(markdown ? "- " + more + "\r" : (shown > 0 ? ", " : "") + more);
  }

  return head + parts.join("") + tail;
}

function dcInternalMessage(analysis, markdown) {
  var n = analysis.internal.length;
  var title = "Envío interno";
  var head =
    (markdown ? "**" + title + "**" : title) +
    " · " +
    dcPlural(n, "destinatario", "destinatarios") +
    " de la organización:" +
    (markdown ? "\n\n" : " ");
  var tail = markdown ? "\n\n¿Enviar el mensaje?" : ". ¿Enviar el mensaje?";
  return dcBuildMessage(head, analysis.internal, tail, markdown);
}

function dcExternalMessage(analysis, step, markdown) {
  var n = analysis.external.length;
  var title = step === 1 ? "ATENCIÓN: DESTINATARIOS EXTERNOS" : "SEGUNDA CONFIRMACIÓN";
  var head =
    (markdown ? "**" + title + "**" : title) +
    " · " +
    step +
    " de 2" +
    (markdown ? "\n\n" : ". ") +
    "El mensaje sale de la organización hacia " +
    dcPlural(n, "externo", "externos") +
    ":" +
    (markdown ? "\n\n" : " ");

  var sep = markdown ? "\n\n" : ". ";
  var tail = "";
  if (analysis.internal.length > 0) {
    tail += sep + "Además, " + dcPlural(analysis.internal.length, "interno", "internos") + ".";
    sep = markdown ? "\n\n" : " ";
  }
  if (step === 1) {
    var button = markdown ? "**" + DC_CONFIRM_LABEL + "**" : '"' + DC_CONFIRM_LABEL + '"';
    tail += sep + "Si son correctos, pulsa " + button;
    tail += DC_CONFIG.autoResend ? "." : " y después vuelve a pulsar Enviar.";
  } else {
    tail += sep + "¿Seguro que quieres enviarlo?";
  }
  return dcBuildMessage(head, analysis.external, tail, markdown);
}

// ---------------------------------------------------------------------------
// ESTADO ENTRE LA 1ª Y LA 2ª CONFIRMACIÓN (sessionData del mensaje en redacción)
// ---------------------------------------------------------------------------
function dcReadStage(item, done) {
  try {
    item.sessionData.getAsync(DC_STAGE_KEY, function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
        done(null);
        return;
      }
      try {
        done(JSON.parse(result.value));
      } catch (e) {
        done(null);
      }
    });
  } catch (e) {
    done(null);
  }
}

function dcClearStage(item, done) {
  try {
    item.sessionData.removeAsync(DC_STAGE_KEY, function () {
      done();
    });
  } catch (e) {
    done();
  }
}

function dcStageIsValid(stage, signature) {
  if (!stage || stage.sig !== signature || typeof stage.t !== "number") {
    return false;
  }
  var age = Date.now() - stage.t;
  return age >= 0 && age <= DC_CONFIG.stageTtlMs;
}

// ---------------------------------------------------------------------------
// UTILIDADES
// ---------------------------------------------------------------------------
function dcIsClassicOutlook() {
  try {
    return (
      Office.context.mailbox.diagnostics.hostName === "Outlook" &&
      Office.context.platform === Office.PlatformType.PC
    );
  } catch (e) {
    return false;
  }
}

// Aviso con "Enviar de todos modos" / "No enviar".
function dcPrompt(event, plain, markdown) {
  var enums = Office.MailboxEnums || {};
  event.completed({
    allowEvent: false,
    errorMessage: plain,
    errorMessageMarkdown: markdown,
    sendModeOverride: enums.SendModeOverride ? enums.SendModeOverride.PromptUser : "promptUser"
  });
}

function dcFail(event, error) {
  var detail = error && error.message ? " (" + error.message + ")" : "";
  var text = "No se han podido comprobar los destinatarios" + detail + ". Revísalos a mano antes de enviar.";
  dcPrompt(event, text.substring(0, DC_CONFIG.maxMessageChars), text.substring(0, DC_CONFIG.maxMessageChars));
}

function dcNotify(item, message, done) {
  try {
    item.notificationMessages.replaceAsync(
      "dcInfo",
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: String(message).substring(0, 150),
        icon: "Icon.16x16",
        persistent: false
      },
      function () {
        done();
      }
    );
  } catch (e) {
    done();
  }
}

// ---------------------------------------------------------------------------
// EVENTO: se pulsa Enviar
// ---------------------------------------------------------------------------
function onMessageSendHandler(event) {
  try {
    if (!DC_CONFIG.runInClassicOutlook && dcIsClassicOutlook()) {
      event.completed({ allowEvent: true });
      return;
    }

    var item = Office.context.mailbox.item;
    dcGetRecipients(item, function (error, recipients) {
      try {
        if (error) {
          dcFail(event, error);
          return;
        }

        var analysis = dcAnalyze(recipients);

        // Sin destinatarios: Outlook ya avisa por su cuenta.
        if (analysis.total === 0) {
          event.completed({ allowEvent: true });
          return;
        }

        // Solo internos: una confirmación.
        if (analysis.external.length === 0) {
          dcPrompt(event, dcInternalMessage(analysis, false), dcInternalMessage(analysis, true));
          return;
        }

        // Hay externos: dos confirmaciones.
        dcReadStage(item, function (stage) {
          try {
            if (dcStageIsValid(stage, analysis.signature)) {
              // 2ª confirmación. La marca es de un solo uso: si se cancela, se empieza de nuevo.
              dcClearStage(item, function () {
                dcPrompt(event, dcExternalMessage(analysis, 2, false), dcExternalMessage(analysis, 2, true));
              });
            } else {
              // 1ª confirmación. Sin "Enviar de todos modos": solo "Confirmar externos" o "No enviar".
              event.completed({
                allowEvent: false,
                errorMessage: dcExternalMessage(analysis, 1, false),
                errorMessageMarkdown: dcExternalMessage(analysis, 1, true),
                cancelLabel: DC_CONFIRM_LABEL,
                commandId: DC_COMMAND_ID,
                contextData: JSON.stringify({ source: "smartAlert" })
              });
            }
          } catch (e) {
            dcFail(event, e);
          }
        });
      } catch (e) {
        dcFail(event, e);
      }
    });
  } catch (e) {
    dcFail(event, e);
  }
}

// ---------------------------------------------------------------------------
// COMANDO: botón "Confirmar externos" del aviso (y botón de la cinta)
// ---------------------------------------------------------------------------
function dcGetInitContext(item, done) {
  try {
    item.getInitializationContextAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
        done(null);
        return;
      }
      try {
        var context = result.value;
        // Según el cliente llega como texto JSON (a veces doblemente codificado) o como objeto.
        for (var i = 0; i < 2 && typeof context === "string"; i++) {
          context = JSON.parse(context);
        }
        done(context && typeof context === "object" ? context : null);
      } catch (e) {
        done(null);
      }
    });
  } catch (e) {
    done(null);
  }
}

function dcConfirmExternal(event) {
  var item = Office.context.mailbox.item;

  function finish() {
    event.completed();
  }

  try {
    dcGetInitContext(item, function (context) {
      dcGetRecipients(item, function (error, recipients) {
        if (error) {
          dcNotify(item, "No se han podido leer los destinatarios.", finish);
          return;
        }
        var analysis = dcAnalyze(recipients);

        // Pulsado desde la cinta (no desde el aviso): solo informa, no confirma nada.
        if (!context || context.source !== "smartAlert") {
          var summary =
            "Destinatarios: " +
            dcPlural(analysis.internal.length, "interno", "internos") +
            ", " +
            dcPlural(analysis.external.length, "externo", "externos") +
            (analysis.external.length ? " (" + analysis.external.join(", ") + ")" : "") +
            ".";
          dcNotify(item, summary.length > 150 ? summary.substring(0, 149) + "…" : summary, finish);
          return;
        }

        // Pulsado desde el aviso: registra la 1ª confirmación para esta lista exacta de destinatarios.
        var stage = JSON.stringify({ sig: analysis.signature, t: Date.now() });
        item.sessionData.setAsync(DC_STAGE_KEY, stage, function (result) {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            dcNotify(item, "No se ha podido registrar la confirmación. Inténtalo de nuevo.", finish);
            return;
          }
          if (DC_CONFIG.autoResend && typeof item.sendAsync === "function") {
            // Relanza el envío: vuelve a saltar OnMessageSend y se muestra la 2ª confirmación.
            item.sendAsync(function () {
              finish();
            });
          } else {
            dcNotify(item, "1ª confirmación hecha. Pulsa Enviar de nuevo para la 2ª.", finish);
          }
        });
      });
    });
  } catch (e) {
    finish();
  }
}

// Asocia los nombres del manifiesto con las funciones.
if (typeof Office !== "undefined" && Office.actions && typeof Office.actions.associate === "function") {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("dcConfirmExternal", dcConfirmExternal);
}

// IMPRESCINDIBLE en Outlook nuevo / web: Office.js no entrega el evento al manejador hasta que la
// página llama a Office.onReady(). Sin esta línea el envío se queda en "está procesando el mensaje".
if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(function () {});
}
