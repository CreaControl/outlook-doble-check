// Pruebas de la lógica con un Office.js simulado.  Ejecutar: npm test
const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const SOURCE = fs.readFileSync(path.join(__dirname, "..", "launchevent.js"), "utf8");

function ok(value) {
  return { status: "succeeded", value };
}
const FAILED = { status: "failed", error: { message: "boom" } };

// Crea un entorno con un mensaje en redacción simulado.
function createEnv({
  to = [],
  cc = [],
  bcc = [],
  subject = "Asunto de prueba",
  body = "<div>Hola, te adjunto el informe.</div>",
  hostName = "newOutlookWindows",
  platform = "OfficeOnline"
} = {}) {
  const session = {};
  const log = { sendAsync: 0, notifications: [], bodyOptions: null };
  const recipients = (list) => ({
    getAsync: (cb) => cb(ok(list.map((a) => (typeof a === "string" ? { emailAddress: a, displayName: a } : a))))
  });

  const item = {
    to: recipients(to),
    cc: recipients(cc),
    bcc: recipients(bcc),
    subject: { getAsync: (cb) => cb(ok(env.subject)) },
    body: {
      getAsync: (coercion, options, cb) => {
        log.bodyOptions = { coercion, ...options };
        cb(ok(env.body));
      }
    },
    sessionData: {
      getAsync: (key, cb) => (key in session ? cb(ok(session[key])) : cb(FAILED)),
      setAsync: (key, value, cb) => {
        session[key] = value;
        cb(ok());
      },
      removeAsync: (key, cb) => {
        delete session[key];
        cb(ok());
      }
    },
    notificationMessages: {
      replaceAsync: (key, data, cb) => {
        log.notifications.push(data.message);
        cb(ok());
      }
    },
    getInitializationContextAsync: (cb) => cb(ok(env.initContext)),
    sendAsync: (cb) => {
      log.sendAsync++;
      cb(ok());
    }
  };

  const Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    PlatformType: { PC: "PC", OfficeOnline: "OfficeOnline" },
    CoercionType: { Html: "html", Text: "text" },
    MailboxEnums: {
      SendModeOverride: { PromptUser: "promptUser" },
      BodyMode: { FullBody: "fullBody", HostConfig: "hostConfig" },
      ItemNotificationMessageType: { InformationalMessage: "informationalMessage" }
    },
    actions: { associate: (name, fn) => (env.associated[name] = fn) },
    onReady: () => (env.onReadyCalls = (env.onReadyCalls || 0) + 1),
    context: { platform, mailbox: { item, diagnostics: { hostName } } }
  };

  const env = { session, log, item, subject, body, initContext: "", associated: {}, last: null };
  env.ctx = vm.createContext({ Office, JSON, Date, String, Object, Math, RegExp, parseInt });
  vm.runInContext(SOURCE, env.ctx);

  // El usuario pulsa Enviar (o el complemento relanza el envío).
  env.send = () => {
    let result;
    env.ctx.onMessageSendHandler({ completed: (options) => (result = options) });
    assert.ok(result, "el manejador debe llamar a event.completed");
    for (const key of ["errorMessage", "errorMessageMarkdown"]) {
      if (result[key] !== undefined) {
        assert.ok(result[key].length <= 500, `${key} supera 500 caracteres: ${result[key].length}`);
      }
    }
    if (result.cancelLabel !== undefined) {
      assert.ok(result.cancelLabel.length <= 20, "el texto del botón supera 20 caracteres");
    }
    env.last = result;
    return result;
  };
  const runCommand = () => {
    let done = false;
    env.ctx.dcConfirmExternal({ completed: () => (done = true) });
    assert.ok(done, "el comando debe llamar a event.completed");
  };
  // El usuario pulsa el botón propio del último aviso mostrado.
  env.clickAlertButton = () => {
    assert.ok(env.last && env.last.commandId, "el último aviso no tiene botón propio");
    env.initContext = env.last.contextData;
    runCommand();
  };
  // El usuario pulsa el botón de la cinta.
  env.clickRibbon = () => {
    env.initContext = "";
    runCommand();
  };
  return env;
}

const isFinal = (r) => r.allowEvent === false && r.sendModeOverride === "promptUser" && r.commandId === undefined;
const isIntermediate = (r) => r.allowEvent === false && r.sendModeOverride === undefined && r.commandId === "dcConfirmButton";

// ---------------------------------------------------------------------------
// Carga
// ---------------------------------------------------------------------------
test("registra los dos nombres de función del manifiesto", () => {
  const env = createEnv();
  assert.deepStrictEqual(Object.keys(env.associated).sort(), ["dcConfirmExternal", "onMessageSendHandler"]);
});

test("llama a Office.onReady al cargar (sin esto Outlook nuevo se queda 'procesando el mensaje')", () => {
  assert.strictEqual(createEnv().onReadyCalls, 1);
});

// ---------------------------------------------------------------------------
// Destinatarios
// ---------------------------------------------------------------------------
test("clasifica dominios internos y externos", () => {
  const { ctx } = createEnv();
  assert.strictEqual(ctx.dcIsInternal("Ana@CreaControl.es"), true);
  assert.strictEqual(ctx.dcIsInternal("ana@creacontrol.es.evil.com"), false);
  assert.strictEqual(ctx.dcIsInternal("ana@sub.creacontrol.es"), false);
  assert.strictEqual(ctx.dcIsInternal("ana@micreacontrol.es"), false);
  assert.strictEqual(ctx.dcIsInternal("sin-arroba"), false);
  assert.strictEqual(ctx.dcIsInternal(""), false);
});

test("solo internos: un único aviso con la lista y 'Enviar de todos modos'", () => {
  const env = createEnv({ to: ["ana@creacontrol.es"], cc: ["luis@creacontrol.es"] });
  const r = env.send();
  assert.ok(isFinal(r));
  assert.match(r.errorMessageMarkdown, /^\*\*Destinatarios\*\*\n\n\*\*Internos \(2\):\*\*/);
  assert.match(r.errorMessageMarkdown, /- ana@creacontrol\.es\r- luis@creacontrol\.es \(CC\)\r/);
  assert.match(r.errorMessageMarkdown, /¿Enviar el mensaje\?$/);
  assert.doesNotMatch(r.errorMessageMarkdown, /Externos| de \d/);
  assert.strictEqual(r.errorMessage, "Destinatarios. Internos (2): ana@creacontrol.es, luis@creacontrol.es (CC). ¿Enviar el mensaje?");
});

test("solo internos: no se revisa el lenguaje aunque haya palabrotas", () => {
  const env = createEnv({ to: ["ana@creacontrol.es"], body: "<p>Vaya mierda de día</p>" });
  assert.ok(isFinal(env.send()));
  assert.strictEqual(env.log.bodyOptions, null, "no debe ni leer el cuerpo");
});

test("con externos: aviso 1 con TODOS los destinatarios y aviso 2 de confirmación", () => {
  const env = createEnv({ to: ["ana@creacontrol.es", "cliente@acme.com"], bcc: ["otro@gmail.com"] });

  const first = env.send();
  assert.ok(isIntermediate(first));
  assert.strictEqual(first.cancelLabel, "Confirmar");
  assert.match(first.errorMessageMarkdown, /^\*\*Destinatarios\*\* · 1 de 2\n\n/);
  assert.match(first.errorMessageMarkdown, /\*\*Externos \(2\):\*\*\n\n- cliente@acme\.com\r- otro@gmail\.com \(CCO\)\r/);
  assert.match(first.errorMessageMarkdown, /\*\*Internos \(1\):\*\*\n\n- ana@creacontrol\.es\r/);
  assert.ok(first.errorMessageMarkdown.indexOf("Externos") < first.errorMessageMarkdown.indexOf("Internos"));
  assert.match(first.errorMessage, /EXTERNOS \(2\): cliente@acme\.com, otro@gmail\.com \(CCO\)\. Internos \(1\): ana@creacontrol\.es\./);

  // Sin pulsar "Confirmar", reintentar el envío vuelve al aviso 1.
  assert.match(env.send().errorMessageMarkdown, /1 de 2/);

  env.clickAlertButton();
  assert.strictEqual(env.log.sendAsync, 1, "debe relanzar el envío");

  const second = env.send();
  assert.ok(isFinal(second));
  assert.match(second.errorMessageMarkdown, /^\*\*SEGUNDA CONFIRMACIÓN\*\* · 2 de 2/);
  assert.match(second.errorMessageMarkdown, /hacia 2 externos:\n\n- cliente@acme\.com\r/);
  assert.doesNotMatch(second.errorMessageMarkdown, /ana@creacontrol\.es/);
  assert.match(second.errorMessageMarkdown, /¿Seguro que quieres enviarlo\?$/);

  // La marca es de un solo uso: si se cancela el aviso 2, se empieza otra vez.
  assert.match(env.send().errorMessageMarkdown, /1 de 2/);
});

test("si cambian los destinatarios tras confirmar, se vuelve a empezar", () => {
  const env = createEnv({ to: ["cliente@acme.com"] });
  env.send();
  env.clickAlertButton();
  env.item.to = { getAsync: (cb) => cb(ok([{ emailAddress: "cliente@acme.com" }, { emailAddress: "nuevo@otro.com" }])) };
  assert.match(env.send().errorMessageMarkdown, /1 de 2/);
});

test("la confirmación caduca", () => {
  const env = createEnv({ to: ["cliente@acme.com"] });
  env.send();
  env.clickAlertButton();
  const stage = JSON.parse(env.session.dcStage);
  stage.t -= 10 * 60 * 1000;
  env.session.dcStage = JSON.stringify(stage);
  assert.match(env.send().errorMessageMarkdown, /1 de 2/);
});

test("una marca manipulada o incoherente no salta avisos", () => {
  const env = createEnv({ to: ["cliente@acme.com"] });
  const key = JSON.parse(env.send().contextData).key;
  for (const done of [2, 5, -1, 1.5, "1"]) {
    env.session.dcStage = JSON.stringify({ key, done, t: Date.now() });
    assert.match(env.send().errorMessageMarkdown, /1 de 2/, `done=${done}`);
  }
});

test("el botón de la cinta solo informa: no confirma ni envía", () => {
  const env = createEnv({ to: ["cliente@acme.com", "ana@creacontrol.es"] });
  env.send();
  env.clickRibbon();
  assert.strictEqual(env.log.sendAsync, 0);
  assert.deepStrictEqual(Object.keys(env.session), []);
  assert.match(env.log.notifications[0], /1 interno, 1 externo/);
  assert.match(env.send().errorMessageMarkdown, /1 de 2/);
});

test("con autoResend desactivado pide volver a pulsar Enviar", () => {
  const env = createEnv({ to: ["cliente@acme.com"] });
  env.ctx.DC_CONFIG.autoResend = false;
  assert.match(env.send().errorMessageMarkdown, /vuelve a pulsar Enviar/);
  env.clickAlertButton();
  assert.strictEqual(env.log.sendAsync, 0);
  assert.match(env.log.notifications[0], /Pulsa Enviar de nuevo/);
  assert.match(env.send().errorMessageMarkdown, /SEGUNDA CONFIRMACIÓN/);
});

test("muchos destinatarios: caben en 500 caracteres y se ven externos e internos", () => {
  const externals = Array.from({ length: 120 }, (_, i) => `persona_numero_${i}@empresa-externa-${i}.com`);
  const internals = Array.from({ length: 80 }, (_, i) => `empleado.numero.${i}@creacontrol.es`);

  const env = createEnv({ to: externals, cc: internals });
  const first = env.send(); // env.send ya comprueba el límite de 500
  assert.match(first.errorMessageMarkdown, /Externos \(120\)/);
  assert.match(first.errorMessageMarkdown, /Internos \(80\)/);
  assert.match(first.errorMessageMarkdown, /persona\\_numero\\_0@/, "los guiones bajos se escapan en Markdown");
  assert.match(first.errorMessageMarkdown, /empleado\.numero\.0@creacontrol\.es/, "debe quedar sitio para algún interno");
  assert.strictEqual((first.errorMessageMarkdown.match(/… y \d+ más/g) || []).length, 2);
  env.clickAlertButton();
  assert.match(env.send().errorMessageMarkdown, /hacia 120 externos/);

  // Pocos internos: se muestran todos y el resto del espacio es para los externos.
  const few = createEnv({ to: externals, cc: ["ana@creacontrol.es"] }).send();
  assert.match(few.errorMessageMarkdown, /- ana@creacontrol\.es \(CC\)\r/);
  assert.ok(few.errorMessageMarkdown.length > 440, "debe aprovechar el espacio disponible");

  createEnv({ to: internals }).send();
});

test("destinatario sin dirección reconocible cuenta como externo", () => {
  const env = createEnv({ to: ["ana@creacontrol.es", { emailAddress: "", displayName: "Lista rara" }] });
  const r = env.send();
  assert.match(r.errorMessageMarkdown, /Externos \(1\):\*\*\n\n- Lista rara \(sin dirección\)/);
});

test("un mismo destinatario repetido en Para y CC cuenta una vez", () => {
  const env = createEnv({ to: ["Cliente@acme.com"], cc: ["cliente@ACME.com"] });
  assert.match(env.send().errorMessageMarkdown, /Externos \(1\)/);
});

test("en Outlook clásico de Windows no interviene (ahí actúa el VBA)", () => {
  const env = createEnv({ to: ["cliente@acme.com"], hostName: "Outlook", platform: "PC" });
  assert.deepStrictEqual({ ...env.send() }, { allowEvent: true });
  env.ctx.DC_CONFIG.runInClassicOutlook = true;
  assert.match(env.send().errorMessageMarkdown, /1 de 2/);
});

test("si falla la lectura de destinatarios, avisa en vez de enviar en silencio", () => {
  const env = createEnv({ to: ["ana@creacontrol.es"] });
  env.item.cc = { getAsync: (cb) => cb(FAILED) };
  const r = env.send();
  assert.strictEqual(r.allowEvent, false);
  assert.match(r.errorMessage, /No se han podido comprobar/);
});

// ---------------------------------------------------------------------------
// Lenguaje
// ---------------------------------------------------------------------------
test("diccionario: palabras completas, sin distinguir mayúsculas ni tildes", () => {
  const { ctx } = createEnv();
  const find = (text) => Array.from(ctx.dcFindBadWords(text));

  assert.deepStrictEqual(find("Eres un GILIPOLLAS y un Imbécil."), ["gilipollas", "imbecil"]);
  assert.deepStrictEqual(find("menuda mierda, ¡joder!"), ["mierda", "joder"]);
  assert.deepStrictEqual(find("mierda mierda MIERDA"), ["mierda"]);
  assert.deepStrictEqual(find("es una estupidez"), ["estupidez"]);
  assert.deepStrictEqual(find("What the fuck, this is bullshit"), ["fuck", "bullshit"]);

  // No deben saltar palabras normales que contienen una malsonante.
  const clean =
    "La computadora, la disputa y la reputación del diputado. Un cono de señalización. " +
    "Desgraciadamente el pedido va retrasado. El cálculo del artículo es estupendo. " +
    "Espero que te den el alta. Reunión en Hostalric con el Sr. Pollard. Scrap y Shitake.";
  assert.deepStrictEqual(find(clean), []);

  assert.deepStrictEqual(find("coño"), ["coño"]);
  assert.deepStrictEqual(find("cono"), []);
});

test("diccionario: expresiones de varias palabras y coincidencias solapadas", () => {
  const { ctx } = createEnv();
  const find = (text) => Array.from(ctx.dcFindBadWords(text));
  assert.deepStrictEqual(find("es un hijo de   puta"), ["hijo de puta"]);
  assert.deepStrictEqual(find("es un hijo de puta y además una puta vergüenza"), ["hijo de puta", "puta"]);
  assert.deepStrictEqual(find("vete a la mierda"), ["vete a la mierda"]);
  assert.deepStrictEqual(find("Me cago en todo"), ["me cago"]);
});

test("diccionario: se puede ampliar", () => {
  const env = createEnv();
  env.ctx.DC_BAD_WORDS.push("Inútil", "chapuc*");
  assert.deepStrictEqual(Array.from(env.ctx.dcFindBadWords("Sois unos inutiles... no: eres un INÚTIL y un chapucero")), [
    "inutil",
    "chapucero"
  ]);
});

test("con externos y palabras malsonantes: tercer aviso con las palabras encontradas", () => {
  const env = createEnv({
    to: ["cliente@acme.com", "ana@creacontrol.es"],
    subject: "Vuestra mierda de servicio",
    body: "<div>Sois unos <b>incompetentes</b> y unos gilipollas.</div><div>No me <i>jodáis</i> más, idiotas.</div>"
  });

  const first = env.send();
  assert.ok(isIntermediate(first));
  assert.match(first.errorMessageMarkdown, /^\*\*Destinatarios\*\* · 1 de 3/);
  assert.strictEqual(first.cancelLabel, "Confirmar");
  env.clickAlertButton();

  const second = env.send();
  assert.ok(isIntermediate(second), "el aviso 2 ya no es el último");
  assert.match(second.errorMessageMarkdown, /^\*\*SEGUNDA CONFIRMACIÓN\*\* · 2 de 3/);
  assert.match(second.errorMessageMarkdown, /pulsa \*\*Confirmar externos\*\*\.$/);
  assert.strictEqual(second.cancelLabel, "Confirmar externos");
  env.clickAlertButton();

  const third = env.send();
  assert.ok(isFinal(third));
  assert.match(third.errorMessageMarkdown, /^\*\*REVISA EL LENGUAJE\*\* · 3 de 3/);
  assert.match(third.errorMessageMarkdown, /estas expresiones: \*\*mierda\*\*, \*\*gilipollas\*\*, \*\*idiotas\*\*\./);
  assert.match(third.errorMessageMarkdown, /¿Enviarlo de todos modos\?$/);
  assert.match(third.errorMessage, /contiene estas expresiones: mierda, gilipollas, idiotas\. ¿Enviarlo de todos modos\?$/);
  assert.strictEqual(env.log.sendAsync, 2);

  // Cancelar el último aviso obliga a empezar de nuevo.
  assert.match(env.send().errorMessageMarkdown, /1 de 3/);
});

test("si se corrige el texto a mitad de camino, se vuelve a empezar con un aviso menos", () => {
  const env = createEnv({ to: ["cliente@acme.com"], body: "<p>Menuda mierda</p>" });
  assert.match(env.send().errorMessageMarkdown, /1 de 3/);
  env.clickAlertButton();
  assert.match(env.send().errorMessageMarkdown, /2 de 3/);
  env.body = "<p>Menudo problema</p>";
  assert.match(env.send().errorMessageMarkdown, /1 de 2/);
});

test("una sola palabra encontrada se anuncia en singular", () => {
  const env = createEnv({ to: ["cliente@acme.com"], body: "<p>Esto es una putada</p>" });
  env.send();
  env.clickAlertButton();
  env.send();
  env.clickAlertButton();
  assert.match(env.send().errorMessageMarkdown, /contiene esta expresión: \*\*putada\*\*\./);
});

test("muchas palabras encontradas: el aviso no supera 500 caracteres", () => {
  const env = createEnv({ to: ["cliente@acme.com"] });
  const words = Array.from({ length: 80 }, (_, i) => `palabrotalarguisima${i}`);
  env.ctx.DC_BAD_WORDS.push(...words);
  env.body = `<p>${words.join(" ")}</p>`;
  env.send();
  env.clickAlertButton();
  env.send();
  env.clickAlertButton();
  const third = env.send();
  assert.match(third.errorMessageMarkdown, /palabrotalarguisima0\*\*, /);
  assert.match(third.errorMessageMarkdown, /… y \d+ más\./);
});

test("en respuestas no se revisa el hilo citado (salvo que se pida)", () => {
  const quotedOwa =
    '<div>Gracias por tu mensaje.</div><div id="appendonsend"></div><hr><div id="divRplyFwdMsg">' +
    "<b>De:</b> Cliente<br><b>Enviado:</b> lunes</div><div>Sois unos gilipollas</div>";
  const quotedText =
    "<div>Gracias por tu mensaje.</div><div><br></div><div>De: Cliente &lt;c@acme.com&gt;</div>" +
    "<div>Enviado el: lunes, 3 de marzo</div><div>Para: Ana</div><div>Asunto: X</div><div>Menuda mierda</div>";
  const quotedOriginal = "<p>Te lo reenvío.</p><p>-----Mensaje original-----</p><p>Eres idiota</p>";

  for (const body of [quotedOwa, quotedText, quotedOriginal]) {
    const env = createEnv({ to: ["cliente@acme.com"], body });
    assert.match(env.send().errorMessageMarkdown, /1 de 2/, body);
    assert.strictEqual(env.log.bodyOptions.coercion, "html");
    assert.strictEqual(env.log.bodyOptions.bodyMode, "hostConfig");

    env.ctx.DC_CONFIG.checkQuotedText = true;
    assert.match(env.send().errorMessageMarkdown, /1 de 3/, body);
  }

  // Lo escrito por encima del hilo citado sí se revisa.
  const own = createEnv({ to: ["cliente@acme.com"], body: "<div>Sois unos capullos.</div>" + quotedOwa });
  assert.match(own.send().errorMessageMarkdown, /1 de 3/);
});

test("el HTML se convierte bien a texto (etiquetas, estilos y entidades)", () => {
  const { ctx } = createEnv();
  const text = ctx.dcHtmlToText(
    "<html><head><style>.mierda{color:red}</style></head><body><!-- idiota -->" +
      '<p class="puta">Imb&eacute;cil&nbsp;total</p><p>co&#241;o</p><a href="https://x.com/joder">enlace</a></body></html>'
  );
  assert.deepStrictEqual(Array.from(ctx.dcFindBadWords(text)), ["imbecil", "coño"]);
  // Palabras partidas por etiquetas de bloque no se pegan entre sí.
  assert.deepStrictEqual(Array.from(ctx.dcFindBadWords(ctx.dcHtmlToText("<div>com</div><div>puta</div>"))), ["puta"]);
});

test("si no se puede leer el texto, se avisa en la confirmación en vez de callarlo", () => {
  const env = createEnv({ to: ["cliente@acme.com"] });
  env.item.body.getAsync = (coercion, options, cb) => cb(FAILED);
  assert.match(env.send().errorMessageMarkdown, /1 de 2/);
  env.clickAlertButton();
  const second = env.send();
  assert.ok(isFinal(second));
  assert.match(second.errorMessageMarkdown, /No se ha podido revisar el lenguaje/);
});

test("la revisión del lenguaje se puede desactivar", () => {
  const env = createEnv({ to: ["cliente@acme.com"], body: "<p>Menuda mierda</p>" });
  env.ctx.DC_CONFIG.checkLanguage = false;
  assert.match(env.send().errorMessageMarkdown, /1 de 2/);
  assert.strictEqual(env.log.bodyOptions, null);
});
