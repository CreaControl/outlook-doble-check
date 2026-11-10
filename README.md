# Doble check de envío para Outlook nuevo

Complemento que, al pulsar **Enviar**, revisa los destinatarios (Para, CC y CCO) y:

1. **siempre** muestra un aviso con todos los destinatarios, externos e internos;
2. si hay algún **externo** (no `@creacontrol.es`), pide una segunda confirmación;
3. si hay algún externo y el asunto o el texto contienen **palabras malsonantes**, muestra un tercer aviso con las palabras encontradas.

Es el equivalente del script VBA de Outlook clásico. Outlook nuevo no admite VBA ni complementos COM, así que la única vía es un **complemento web** que usa la función *Smart Alerts* (evento `OnMessageSend`).

## Cómo se ve

| Caso | Aviso | Botones |
| --- | --- | --- |
| Solo internos | «Destinatarios» con la lista de internos | **Enviar de todos modos** / **No enviar** |
| Con externos, aviso 1 | «Destinatarios · 1 de 2» con los externos primero y los internos después | **Confirmar** / **No enviar** |
| Con externos, aviso 2 | «SEGUNDA CONFIRMACIÓN · 2 de 2» con la lista de externos | **Enviar de todos modos** / **No enviar** |
| Con externos y palabras malsonantes, aviso 3 | «REVISA EL LENGUAJE · 3 de 3» con las palabras encontradas | **Enviar de todos modos** / **No enviar** |

Cuando hay tercer aviso, los contadores pasan a ser «1 de 3», «2 de 3» y «3 de 3», y el botón del aviso 2 es **Confirmar externos**. El mensaje solo sale al pulsar **Enviar de todos modos** en el último aviso.

Los avisos son los cuadros de diálogo propios de Outlook: el título es el nombre del complemento, el texto admite como mucho 500 caracteres (si no cabe todo se muestra «… y N más») y los textos «Enviar de todos modos» y «No enviar» los pone Outlook.

Si entre un aviso y el siguiente cambian los destinatarios o las palabras encontradas, o pasan más de 2 minutos, o se cancela el último aviso, el proceso empieza de nuevo.

## Diccionario de palabras malsonantes

Está en [launchevent.js](launchevent.js), en la lista `DC_BAD_WORDS`, y se puede ampliar o recortar libremente:

- No distingue mayúsculas ni tildes: `"imbecil"` detecta «Imbécil».
- Solo coinciden palabras completas: `"puta"` no salta con «computadora» ni «disputa».
- Un asterisco al final admite cualquier terminación: `"gilipoll*"` detecta «gilipollas», «gilipollez»…
- Se admiten expresiones de varias palabras: `"hijo de puta"`.

Se revisa el asunto y el texto del mensaje, y solo cuando hay destinatarios externos. En respuestas y reenvíos se intenta revisar únicamente lo que has escrito tú, no el hilo citado (opción `checkQuotedText`). Esa separación se hace reconociendo las marcas que pone Outlook al citar; si en algún caso no las reconoce, el efecto es que también revisa el hilo, nunca que deje de revisar tu texto.

He dejado fuera a propósito palabras que en un correo de trabajo suelen ser inocentes («inútil», «retrasado», «anormal», «tonto»…). Añádelas a la lista si las quieres.

## Qué hace falta

1. **Buzón de Microsoft 365 / Exchange Online.** En cuentas IMAP o Gmail dentro de Outlook nuevo los complementos no funcionan.
2. **Un sitio HTTPS donde alojar los archivos del complemento** (`commands.html`, `launchevent.js` y los iconos). Outlook los descarga de ahí cada vez; no vale una ruta de disco.
3. **Instalar el [manifest.xml](manifest.xml)** en el buzón: a mano para uno mismo, o desde el centro de administración de Microsoft 365 para toda la empresa.

## Archivos

| Archivo | Para qué sirve |
| --- | --- |
| [manifest.xml](manifest.xml) | Describe el complemento a Outlook. Es lo que se instala. |
| [launchevent.js](launchevent.js) | Toda la lógica. La configuración está al principio (`DC_CONFIG`). |
| [commands.html](commands.html) | Página vacía que carga Office.js y el script. |
| [assets/](assets/) | Iconos. |
| [test/logic.test.js](test/logic.test.js) | Pruebas de la lógica con un Outlook simulado (`npm test`). |
| [package.json](package.json) | Comandos `npm test`, `npm run certs` y `npm run serve`. |

Todo está en un único repositorio de GitHub, que sirve a la vez de copia de seguridad y de alojamiento. Outlook solo usa `commands.html`, `launchevent.js` y los iconos; el resto se publica también, pero no interviene.

## Puesta en marcha

### 1. Alojamiento (ya hecho)

El repositorio está publicado con GitHub Pages:

- Repositorio (público): <https://github.com/CreaControl/outlook-doble-check>
- URL que usa Outlook: `https://creacontrol.github.io/outlook-doble-check/`
- URL del manifiesto: `https://creacontrol.github.io/outlook-doble-check/manifest.xml`

Para publicar un cambio, desde la carpeta del repositorio:

```powershell
npm test
git add -A
git commit -m "Descripción del cambio"
git push
```

GitHub tarda alrededor de un minuto en servir la versión nueva.

### 2. URL en el manifiesto (ya hecho)

`manifest.xml` ya apunta a la URL anterior. Si algún día se mueve el alojamiento, hay que sustituir `https://creacontrol.github.io/outlook-doble-check` por la URL nueva en todo el archivo y volver a instalarlo.

### 3. Instalar el complemento

**Solo para ti (instalación manual):**

1. Abre <https://aka.ms/olksideload>. Se abre Outlook web con la ventana *Complementos para Outlook*.
2. **Mis complementos** → **Complementos personalizados** → **Agregar un complemento personalizado** → **Agregar desde un archivo** → elige `manifest.xml`.
3. Cierra y vuelve a abrir Outlook nuevo. El complemento se instala en el buzón, así que aparece también en Outlook web.

Si la opción no aparece, el administrador tiene bloqueada la instalación de complementos personalizados; hay que usar la vía siguiente.

**Para toda la empresa (administrador de Microsoft 365):**

Centro de administración → **Configuración** → **Aplicaciones integradas** → **Cargar aplicaciones personalizadas** → tipo **Complemento de Office** → subir `manifest.xml` o indicar su URL (la del apartado 1) → elegir a qué usuarios se asigna.

La URL solo se lee al instalar: si después cambia el manifiesto, hay que actualizar el complemento desde ese mismo sitio.

### 4. Comprobar que funciona

1. Mensaje solo a direcciones `@creacontrol.es` → un aviso con la lista; con **Enviar de todos modos** se envía.
2. Mensaje con algún externo → aviso 1 con todos los destinatarios; al pulsar **Confirmar** aparece el aviso 2, y con **Enviar de todos modos** se envía.
3. Mensaje con externo en CCO → igual que el caso 2, con la dirección marcada «(CCO)».
4. Mensaje con algún externo y una palabra del diccionario en el texto → tres avisos; el tercero muestra la palabra.
5. Respuesta a un correo cuyo hilo citado contiene una palabra del diccionario, sin escribirla tú → solo dos avisos.

El paso de un aviso al siguiente lo hace el complemento relanzando el envío. Si en algún cliente el mensaje saliera sin mostrar el aviso siguiente, pon `autoResend: false` en `launchevent.js`: entonces, tras confirmar un aviso, hay que pulsar **Enviar** otra vez para ver el siguiente.

## Configuración

Al principio de [launchevent.js](launchevent.js):

| Opción | Valor por defecto | Significado |
| --- | --- | --- |
| `internalDomains` | `["creacontrol.es"]` | Dominios internos. Se pueden añadir más. |
| `includeSubdomains` | `false` | Si `x@algo.creacontrol.es` cuenta como interno. |
| `checkLanguage` | `true` | Activa la revisión de palabras malsonantes cuando hay externos. |
| `checkQuotedText` | `false` | Con `true`, en respuestas y reenvíos se revisa también el hilo citado. |
| `runInClassicOutlook` | `false` | El complemento se instala en el buzón y por tanto también se carga en Outlook clásico. Con `false` ahí no hace nada, para no duplicar los avisos del VBA. Ponlo a `true` si quitas el VBA. |
| `autoResend` | `true` | Ver el apartado anterior. |
| `stageTtlMs` | `120000` | Tiempo máximo entre un aviso y el siguiente. |

La lista `DC_BAD_WORDS` (diccionario) está justo debajo de `DC_CONFIG`.

Tras cambiar el script basta con volver a publicarlo (`git push`); Outlook puede tardar un rato en recoger la versión nueva (cerrar y abrir Outlook lo acelera). Si se cambia `manifest.xml`, hay que volver a instalarlo.

## Limitaciones que conviene conocer

- **No se pueden mostrar cuadros de diálogo propios.** Outlook solo permite su aviso estándar al enviar, con un mensaje y esos botones; por eso las confirmaciones se hacen encadenando avisos.
- **El diccionario es una lista de palabras, no entiende el contexto.** No detecta palabras disfrazadas («m1erda», «m*erda») ni insultos que no estén en la lista, y puede avisar por una palabra usada en sentido inocente.
- **Hace falta conexión.** Si el alojamiento no responde, Outlook avisa de que el complemento no está disponible y deja enviar igualmente (modo *SoftBlock*). Si se prefiere que sin complemento no se pueda enviar, cambia `SendMode="SoftBlock"` por `SendMode="Block"` en el manifiesto.
- **Permisos.** El manifiesto pide `ReadWriteMailbox` porque es lo que exige Outlook para relanzar el envío tras confirmar un aviso. El complemento lee los destinatarios y, si hay externos, el asunto y el texto para buscar las palabras del diccionario; todo se procesa dentro de Outlook y no se envía a ningún sitio.
- **Botón en la cinta.** Aparece un botón «Revisar destinatarios» al redactar. Es obligatorio para que existan los botones de confirmación de los avisos; pulsado desde la cinta solo muestra cuántos internos y externos hay.
- **Solo correos.** Las convocatorias de reunión no se comprueban.
- **Listas de distribución.** Se clasifican por la dirección de la lista, no por sus miembros.

## Probar en local (opcional)

```powershell
npm run certs   # instala un certificado de desarrollo para https://localhost (pide permiso)
npm run serve   # sirve esta carpeta en https://localhost:3000
```

Para probar así hace falta una copia de `manifest.xml` con `https://creacontrol.github.io/outlook-doble-check` sustituido por `https://localhost:3000` (y un `<Id>` distinto, para que no choque con el complemento ya instalado). Solo funciona en ese equipo y mientras el servidor esté arrancado.
