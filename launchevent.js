/*
 * Doble check de envío — complemento Smart Alerts para Outlook (nuevo) y Outlook web.
 *
 *  1. Siempre: aviso con TODOS los destinatarios (externos e internos).
 *  2. Si hay algún externo: segunda confirmación.
 *  3. Si hay algún externo y el asunto o el texto contienen palabras del diccionario
 *     de palabras malsonantes: aviso extra con las palabras encontradas.
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

  // true: si hay destinatarios externos, busca palabras malsonantes en el asunto y el texto.
  checkLanguage: true,

  // false: en respuestas y reenvíos solo se revisa lo que has escrito tú, no el hilo citado.
  // true: se revisa todo el cuerpo, incluidos los mensajes anteriores.
  checkQuotedText: false,

  // false: en Outlook clásico de Windows el complemento no hace nada (ahí ya actúa el script VBA).
  // Ponlo a true si retiras el VBA y quieres el mismo comportamiento en los dos Outlook.
  runInClassicOutlook: false,

  // true: al confirmar un aviso se relanza el envío y aparece solo el siguiente aviso.
  // false: tras confirmar hay que volver a pulsar Enviar para ver el siguiente.
  autoResend: true,

  // Tiempo máximo entre un aviso y el siguiente. Pasado este tiempo se vuelve a empezar.
  stageTtlMs: 120000,

  // Límite de Outlook para el texto del aviso.
  maxMessageChars: 500
};

// ---------------------------------------------------------------------------
// DICCIONARIO DE PALABRAS MALSONANTES
// ---------------------------------------------------------------------------
// - Da igual mayúsculas/minúsculas y las tildes ("imbecil" detecta "Imbécil").
// - Solo coinciden palabras completas: "puta" NO salta con "computadora" ni "disputa".
// - Un asterisco al final detecta cualquier terminación: "gilipoll*" -> gilipollas, gilipollez...
// - Se pueden poner expresiones de varias palabras: "hijo de puta".
var DC_BAD_WORDS = [
  // Insultos
  "gilipoll*", "gilipuert*", "imbecil*", "idiot*", "estupid*", "subnormal*", "capull*",
  "cabron*", "hijo de puta", "hijos de puta", "hija de puta", "hijoputa*", "hijaputa*",
  "hdp", "mamon", "mamona", "mamones", "maricon*", "zorra", "zorras", "payaso", "payasos", "payasa",
  "cretin*", "desgraciado", "desgraciada", "desgraciados", "desgraciadas", "asqueros*",
  "sinverguenza*", "malnacid*", "bastard*", "tarado", "tarada", "tarados", "taradas",
  "zopenc*", "mentecat*", "pendej*", "boludo*", "pelotud*", "huevon*", "lameculos", "pringad*",
  "mierdos*", "comemierda*", "tontopolla*", "tontolaba", "soplapollas", "chupapollas",
  // Lenguaje vulgar
  "mierda", "mierdas", "puta", "putas", "puto", "putos", "putada", "putadas", "puteo", "putead*",
  "joder", "jodido", "jodida", "jodidos", "jodidas", "jodete", "jodanse", "jodiendo", "joda",
  "coño", "coños", "hostia", "hostias", "cojon*", "polla", "pollas", "culo", "carajo",
  "me cago", "cagada", "cagadas", "cagarla", "cagaste", "puñeter*",
  "a tomar por culo", "vete a la mierda", "iros a la mierda",
  // Inglés
  "fuck*", "motherfuck*", "shit", "shitty", "bullshit", "bitch", "bitches", "asshole*", "arsehole*",
  "moron", "morons", "stupid", "dumbass", "jackass", "crap", "wtf"
];

var DC_STAGE_KEY = "dcStage";
var DC_COMMAND_ID = "dcConfirmButton"; // id del <Control> en manifest.xml
var DC_LABEL_RECIPIENTS = "Confirmar"; // texto del botón, máx. 20 caracteres
var DC_LABEL_EXTERNAL = "Confirmar externos"; // máx. 20 caracteres

// ---------------------------------------------------------------------------
// CLASIFICACIÓN DE DESTINATARIOS
// ---------------------------------------------------------------------------
function dcTrim(text) {
  return String(text == null ? "" : text).replace(/^\s+|\s+$/g, "");
}

function dcDomainOf(address) {
  var a = dcTrim(address).toLowerCase();
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
          address: dcTrim(values[i].emailAddress),
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
// REVISIÓN DEL LENGUAJE
// ---------------------------------------------------------------------------
var DC_ACCENTS = {
  "á": "a", "à": "a", "ä": "a", "â": "a",
  "é": "e", "è": "e", "ë": "e", "ê": "e",
  "í": "i", "ì": "i", "ï": "i", "î": "i",
  "ó": "o", "ò": "o", "ö": "o", "ô": "o",
  "ú": "u", "ù": "u", "ü": "u", "û": "u"
};

// Minúsculas y sin tildes. La ñ se conserva ("coño" no es "cono").
function dcNormalize(text) {
  return String(text == null ? "" : text)
    .toLowerCase()
    .replace(/[áàäâéèëêíìïîóòöôúùüû]/g, function (c) {
      return DC_ACCENTS[c];
    });
}

var dcBadWordPatterns = null;

function dcCompileBadWords() {
  if (dcBadWordPatterns) {
    return dcBadWordPatterns;
  }
  var letter = "a-z0-9ñç";
  var patterns = [];
  for (var i = 0; i < DC_BAD_WORDS.length; i++) {
    var entry = dcTrim(dcNormalize(DC_BAD_WORDS[i]));
    var anyEnding = entry.charAt(entry.length - 1) === "*";
    if (anyEnding) {
      entry = entry.slice(0, -1);
    }
    if (!entry) {
      continue;
    }
    var words = entry.split(/\s+/);
    for (var w = 0; w < words.length; w++) {
      words[w] = words[w].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
    var body = words.join("\\s+") + (anyEnding ? "[" + letter + "]*" : "");
    patterns.push(new RegExp("(^|[^" + letter + "])(" + body + ")(?![" + letter + "])", "g"));
  }
  dcBadWordPatterns = patterns;
  return patterns;
}

// Devuelve las palabras o expresiones del diccionario que aparecen en el texto, sin repetir
// y en orden de aparición.
function dcFindBadWords(text) {
  var normalized = dcNormalize(text);
  var patterns = dcCompileBadWords();
  var hits = [];
  var i;
  var k;

  for (i = 0; i < patterns.length && hits.length < 500; i++) {
    var pattern = patterns[i];
    pattern.lastIndex = 0;
    var match;
    while ((match = pattern.exec(normalized)) !== null && hits.length < 500) {
      var start = match.index + match[1].length;
      hits.push({ word: match[2].replace(/\s+/g, " "), start: start, end: start + match[2].length });
    }
  }

  // Si una coincidencia está dentro de otra más larga ("puta" en "hijo de puta"), vale la larga.
  hits.sort(function (a, b) {
    return b.end - b.start - (a.end - a.start) || a.start - b.start;
  });
  var kept = [];
  for (i = 0; i < hits.length; i++) {
    var inside = false;
    for (k = 0; k < kept.length; k++) {
      if (hits[i].start >= kept[k].start && hits[i].end <= kept[k].end) {
        inside = true;
        break;
      }
    }
    if (!inside) {
      kept.push(hits[i]);
    }
  }

  kept.sort(function (a, b) {
    return a.start - b.start;
  });
  var seen = {};
  var found = [];
  for (i = 0; i < kept.length; i++) {
    if (!seen["w:" + kept[i].word]) {
      seen["w:" + kept[i].word] = true;
      found.push(kept[i].word);
    }
  }
  return found;
}

var DC_ENTITIES = {
  nbsp: " ", amp: "&", lt: "<", gt: ">", quot: '"', apos: "'",
  aacute: "á", eacute: "é", iacute: "í", oacute: "ó", uacute: "ú", uuml: "ü", ntilde: "ñ", ccedil: "ç",
  Aacute: "Á", Eacute: "É", Iacute: "Í", Oacute: "Ó", Uacute: "Ú", Uuml: "Ü", Ntilde: "Ñ", Ccedil: "Ç",
  iexcl: "¡", iquest: "¿"
};

// Marcas con las que los distintos Outlook separan lo escrito del hilo citado.
var DC_QUOTE_HTML =
  /<[^>]+(?:id\s*=\s*["']?(?:divRplyFwdMsg|appendonsend|mail-editor-reference-message-container|OLK_SRC_BODY_SECTION)\b|class\s*=\s*["']?gmail_quote\b|border-top\s*:\s*solid\s+#(?:E1E1E1|B5C4DF))/i;
var DC_QUOTE_TEXT = [
  /(^|\n)[ \t>]*(?:De|From)\s*:[^\n]*\n(?:[^\n]*\n){0,3}?[ \t>]*(?:Enviado(?: el)?|Sent|Fecha|Date)\s*:/i,
  /(^|\n)[ \t>]*-{3,}\s*(?:Mensaje original|Original Message|Mensaje reenviado|Forwarded message)\s*-{3,}/i
];

// Convierte el cuerpo HTML en texto y, salvo que se pida lo contrario, descarta el hilo citado.
function dcHtmlToText(html) {
  var h = String(html == null ? "" : html);

  if (!DC_CONFIG.checkQuotedText) {
    var cut = h.search(DC_QUOTE_HTML);
    if (cut >= 0) {
      h = h.substring(0, cut);
    }
  }

  h = h
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<(style|script|head|title)\b[\s\S]*?<\/\1\s*>/gi, " ")
    .replace(/<(?:br|hr|\/p|\/div|\/li|\/tr|\/h[1-6]|\/blockquote)\b[^>]*>/gi, "\n")
    .replace(/<[^>]*>/g, " ")
    .replace(/&#(\d+);/g, function (m, code) {
      return String.fromCharCode(parseInt(code, 10));
    })
    .replace(/&#x([0-9a-f]+);/gi, function (m, code) {
      return String.fromCharCode(parseInt(code, 16));
    })
    .replace(/&([a-zA-Z]+);/g, function (m, name) {
      return Object.prototype.hasOwnProperty.call(DC_ENTITIES, name) ? DC_ENTITIES[name] : " ";
    })
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t ]+/g, " ");

  if (!DC_CONFIG.checkQuotedText) {
    for (var i = 0; i < DC_QUOTE_TEXT.length; i++) {
      var at = h.search(DC_QUOTE_TEXT[i]);
      if (at >= 0) {
        h = h.substring(0, at);
      }
    }
  }
  return h;
}

// Lee asunto y cuerpo y busca palabras del diccionario. done({ words: [...], failed: bool })
function dcCheckLanguage(item, done) {
  if (!DC_CONFIG.checkLanguage) {
    done({ words: [], failed: false });
    return;
  }

  var finished = false;
  function finish(result) {
    if (!finished) {
      finished = true;
      done(result);
    }
  }

  try {
    item.subject.getAsync(function (subjectResult) {
      try {
        var failed = subjectResult.status !== Office.AsyncResultStatus.Succeeded;
        var subject = failed ? "" : String(subjectResult.value || "");

        var options = {};
        var enums = Office.MailboxEnums || {};
        if (enums.BodyMode && enums.BodyMode.HostConfig) {
          // En respuestas, pide a Outlook nuevo / web solo el mensaje actual cuando lo permite.
          options.bodyMode = enums.BodyMode.HostConfig;
        }

        item.body.getAsync(Office.CoercionType.Html, options, function (bodyResult) {
          try {
            var body = "";
            if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
              body = dcHtmlToText(bodyResult.value);
            } else {
              failed = true;
            }
            finish({ words: dcFindBadWords(subject + "\n" + body), failed: failed });
          } catch (e) {
            finish({ words: [], failed: true });
          }
        });
      } catch (e) {
        finish({ words: [], failed: true });
      }
    });
  } catch (e) {
    finish({ words: [], failed: true });
  }
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

function dcBold(text, markdown) {
  return markdown ? "**" + text + "**" : text;
}

function dcTitle(text, step, steps, markdown) {
  var counter = steps > 1 ? " · " + step + " de " + steps : "";
  return dcBold(text, markdown) + counter + (markdown ? "\n\n" : ". ");
}

// Lista de direcciones que cabe en "budget" caracteres; si no caben todas añade "… y N más".
function dcFitList(items, budget, markdown) {
  var parts = [];
  var used = 0;
  var shown = 0;

  for (var i = 0; i < items.length; i++) {
    var text = markdown ? "- " + dcEscapeMarkdown(items[i]) + "\r" : (i > 0 ? ", " : "") + items[i];
    var reserve = i < items.length - 1 ? 24 : 0; // hueco para "… y N más"
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
  return parts.join("");
}

function dcNextStepHint(label, markdown) {
  var button = markdown ? "**" + label + "**" : '"' + label + '"';
  return "pulsa " + button + (DC_CONFIG.autoResend ? "." : " y después vuelve a pulsar Enviar.");
}

// Aviso 1: todos los destinatarios, externos primero.
function dcRecipientsMessage(analysis, steps, markdown) {
  var hasExternal = analysis.external.length > 0;
  var hasInternal = analysis.internal.length > 0;
  var gap = markdown ? "\n\n" : " ";

  var head = dcTitle("Destinatarios", 1, steps, markdown);
  var externalHead = hasExternal
    ? dcBold((markdown ? "Externos" : "EXTERNOS") + " (" + analysis.external.length + "):", markdown) + gap
    : "";
  var internalHead = hasInternal
    ? (hasExternal ? (markdown ? "\n\n" : ". ") : "") +
      dcBold("Internos (" + analysis.internal.length + "):", markdown) +
      gap
    : "";
  var tail =
    (markdown ? "\n\n" : ". ") +
    (hasExternal ? "Si son correctos, " + dcNextStepHint(DC_LABEL_RECIPIENTS, markdown) : "¿Enviar el mensaje?");

  var available = DC_CONFIG.maxMessageChars - head.length - externalHead.length - internalHead.length - tail.length;
  var externalList = "";
  var internalList = "";

  if (hasExternal && hasInternal) {
    // Los externos tienen preferencia, pero se deja sitio para los internos.
    var internalFull = dcFitList(analysis.internal, 100000, markdown).length;
    var externalBudget = Math.max(available - internalFull, Math.floor(available * 0.6));
    externalList = dcFitList(analysis.external, externalBudget, markdown);
    internalList = dcFitList(analysis.internal, available - externalList.length, markdown);
  } else if (hasExternal) {
    externalList = dcFitList(analysis.external, available, markdown);
  } else {
    internalList = dcFitList(analysis.internal, available, markdown);
  }

  return head + externalHead + externalList + internalHead + internalList + tail;
}

// Aviso 2: confirmación de que el mensaje sale de la organización.
function dcExternalMessage(analysis, step, steps, languageFailed, markdown) {
  var isLast = step === steps;
  var head =
    dcTitle("SEGUNDA CONFIRMACIÓN", step, steps, markdown) +
    "El mensaje sale de la organización hacia " +
    dcPlural(analysis.external.length, "externo", "externos") +
    ":" +
    (markdown ? "\n\n" : " ");

  var sep = markdown ? "\n\n" : ". ";
  var tail =
    sep + (isLast ? "¿Seguro que quieres enviarlo?" : "Si es correcto, " + dcNextStepHint(DC_LABEL_EXTERNAL, markdown));
  if (languageFailed) {
    tail += (markdown ? "\n\n" : " ") + "(No se ha podido revisar el lenguaje del mensaje.)";
  }

  var budget = DC_CONFIG.maxMessageChars - head.length - tail.length;
  return head + dcFitList(analysis.external, budget, markdown) + tail;
}

// Aviso 3: palabras malsonantes encontradas.
function dcLanguageMessage(words, step, steps, markdown) {
  var head =
    dcTitle("REVISA EL LENGUAJE", step, steps, markdown) +
    "El mensaje va a destinatarios externos y contiene " +
    (words.length === 1 ? "esta expresión" : "estas expresiones") +
    ": ";
  var tail = (markdown ? ".\n\n" : ". ") + "¿Enviarlo de todos modos?";

  var budget = DC_CONFIG.maxMessageChars - head.length - tail.length;
  var parts = [];
  var used = 0;
  for (var i = 0; i < words.length; i++) {
    var word = (i > 0 ? ", " : "") + dcBold(markdown ? dcEscapeMarkdown(words[i]) : words[i], markdown);
    var reserve = i < words.length - 1 ? 16 : 0; // hueco para "… y N más"
    if (used + word.length + reserve > budget) {
      parts.push((i > 0 ? ", " : "") + "… y " + (words.length - i) + " más");
      break;
    }
    parts.push(word);
    used += word.length;
  }
  return head + parts.join("") + tail;
}

// ---------------------------------------------------------------------------
// ESTADO ENTRE UN AVISO Y EL SIGUIENTE (sessionData del mensaje en redacción)
// ---------------------------------------------------------------------------
// Huella corta de los destinatarios y de las palabras encontradas: si cambian, se empieza de nuevo.
function dcHash(text) {
  var h = 5381;
  for (var i = 0; i < text.length; i++) {
    h = ((h << 5) + h + text.charCodeAt(i)) | 0;
  }
  return (h >>> 0).toString(36) + "." + text.length;
}

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

// Número de avisos ya confirmados para este mensaje tal y como está ahora (0 si no hay ninguno válido).
function dcConfirmedSteps(stage, key, totalSteps) {
  if (!stage || stage.key !== key || typeof stage.t !== "number" || typeof stage.done !== "number") {
    return 0;
  }
  var age = Date.now() - stage.t;
  if (age < 0 || age > DC_CONFIG.stageTtlMs) {
    return 0;
  }
  if (stage.done < 1 || stage.done >= totalSteps || stage.done !== Math.floor(stage.done)) {
    return 0;
  }
  return stage.done;
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

// Aviso final: "Enviar de todos modos" / "No enviar".
function dcPrompt(event, plain, markdown) {
  var enums = Office.MailboxEnums || {};
  event.completed({
    allowEvent: false,
    errorMessage: plain,
    errorMessageMarkdown: markdown,
    sendModeOverride: enums.SendModeOverride ? enums.SendModeOverride.PromptUser : "promptUser"
  });
}

// Aviso intermedio: botón propio (pasa al siguiente aviso) / "No enviar".
function dcAskToContinue(event, plain, markdown, label, key, done) {
  event.completed({
    allowEvent: false,
    errorMessage: plain,
    errorMessageMarkdown: markdown,
    cancelLabel: label,
    commandId: DC_COMMAND_ID,
    contextData: JSON.stringify({ source: "smartAlert", key: key, done: done })
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

        // Solo internos: un único aviso con la lista.
        if (analysis.external.length === 0) {
          dcPrompt(event, dcRecipientsMessage(analysis, 1, false), dcRecipientsMessage(analysis, 1, true));
          return;
        }

        // Hay externos: destinatarios -> confirmación de externos -> (lenguaje, si procede).
        dcCheckLanguage(item, function (language) {
          try {
            var steps = language.words.length > 0 ? 3 : 2;
            var key = dcHash(analysis.signature + "|" + language.words.join(","));

            dcReadStage(item, function (stage) {
              try {
                var step = dcConfirmedSteps(stage, key, steps) + 1;
                var plain;
                var markdown;

                if (step === 1) {
                  plain = dcRecipientsMessage(analysis, steps, false);
                  markdown = dcRecipientsMessage(analysis, steps, true);
                } else if (step === 2) {
                  plain = dcExternalMessage(analysis, step, steps, language.failed, false);
                  markdown = dcExternalMessage(analysis, step, steps, language.failed, true);
                } else {
                  plain = dcLanguageMessage(language.words, step, steps, false);
                  markdown = dcLanguageMessage(language.words, step, steps, true);
                }

                if (step === steps) {
                  // Último aviso. La marca es de un solo uso: si se cancela, se empieza de nuevo.
                  dcClearStage(item, function () {
                    dcPrompt(event, plain, markdown);
                  });
                } else {
                  dcAskToContinue(
                    event,
                    plain,
                    markdown,
                    step === 1 ? DC_LABEL_RECIPIENTS : DC_LABEL_EXTERNAL,
                    key,
                    step
                  );
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
    });
  } catch (e) {
    dcFail(event, e);
  }
}

// ---------------------------------------------------------------------------
// COMANDO: botón de confirmación de los avisos (y botón de la cinta)
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
      var fromAlert =
        context && context.source === "smartAlert" && typeof context.key === "string" && typeof context.done === "number";

      // Pulsado desde la cinta (no desde un aviso): solo informa, no confirma nada.
      if (!fromAlert) {
        dcGetRecipients(item, function (error, recipients) {
          if (error) {
            dcNotify(item, "No se han podido leer los destinatarios.", finish);
            return;
          }
          var analysis = dcAnalyze(recipients);
          var summary =
            "Destinatarios: " +
            dcPlural(analysis.internal.length, "interno", "internos") +
            ", " +
            dcPlural(analysis.external.length, "externo", "externos") +
            (analysis.external.length ? " (" + analysis.external.join(", ") + ")" : "") +
            ".";
          dcNotify(item, summary.length > 150 ? summary.substring(0, 149) + "…" : summary, finish);
        });
        return;
      }

      // Pulsado desde un aviso: lo da por confirmado para el mensaje tal y como estaba en ese momento.
      var stage = JSON.stringify({ key: context.key, done: context.done, t: Date.now() });
      item.sessionData.setAsync(DC_STAGE_KEY, stage, function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          dcNotify(item, "No se ha podido registrar la confirmación. Inténtalo de nuevo.", finish);
          return;
        }
        if (DC_CONFIG.autoResend && typeof item.sendAsync === "function") {
          // Relanza el envío: vuelve a saltar OnMessageSend y se muestra el siguiente aviso.
          item.sendAsync(function () {
            finish();
          });
        } else {
          dcNotify(item, "Confirmado. Pulsa Enviar de nuevo para el siguiente aviso.", finish);
        }
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
