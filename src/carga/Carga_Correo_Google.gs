/**
 * Carga docente EPS (#28) · adaptador de los correos a los servicios de Google: MailApp, el activador diario
 * (ScriptApp) y la hoja Correos. La lógica está en Carga_Correo.gs; acá sólo se conectan los servicios.
 *
 * Funciones para la cuenta propietaria, desde el editor:
 *   activarCorreosCarga       crea (o recrea) el activador diario de enviarCorreosCarga a las 8 (Buenos Aires).
 *                             Apps Script lo ejecuta en algún minuto entre las 8:00 y las 9:00, el mismo cada día:
 *                             «a las 8:00» es la hora deseada, no un horario garantizado al minuto.
 *   desactivarCorreosCarga    borra ese activador. Los correos dejan de salir; la carga sigue funcionando.
 *   previsualizarCorreosCarga lo que se enviaría ahora, sin anotar ni enviar nada (asuntos, destinatarios, tamaños).
 *   enviarCorreosCarga        la ejecuta el activador. A mano, sirve para reintentar el mismo día lo que falló o
 *                             quedó sin procesar: nunca manda un segundo correo del día a quien ya lo recibió.
 *
 * Los correos salen de la cuenta propietaria (la que crea el activador), con el nombre «Equipo EPS». Si una
 * ejecución tiene problemas, errores de envío o mensajes sin procesar, termina con error: Google lo informa por
 * correo a esa cuenta y el detalle queda en Ejecuciones y en la hoja Correos.
 *
 * enviarCorreosCarga no termina en "_" (un activador no puede llamar a una función privada), así que la página
 * también podría invocarla: sólo corre si la llama el activador de este proyecto (su triggerUid) o la propietaria.
 */

function enviarCorreosCarga(e) {
  autorizarCorreos_(e);
  // Un lock de usuario: excluye otra ejecución de los correos (activador o manual, ambas de la propietaria) sin
  // competir con el lock del guardado docente, que es el del script.
  const lock = LockService.getUserLock();
  if (!lock.tryLock(1000)) { console.warn('Hay otra ejecución de los correos en curso: ésta no hace nada.'); return null; }
  let r;
  try {
    r = enviarCorreosDia_(serviciosCorreo_(), {});
  } finally {
    lock.releaseLock();
  }
  informarCorreos_(r);
  const fallas = r.problemas.length + r.errores.length + r.sinProcesar.length;
  if (fallas) {
    throw new Error('Correos del ' + r.dia + ': ' + r.enviados.length + ' enviados, ' + r.errores.length + ' con error, ' + r.sinProcesar.length +
      ' sin procesar, ' + r.problemas.length + ' problemas. Detalle en el registro de esta ejecución y en la hoja ' + CORREO.HOJA + '.');
  }
  return resumenCorreos_(r);
}

/** Lo que se enviaría ahora, sin anotar ni enviar. */
function previsualizarCorreosCarga() {
  soloPropietario_();
  const r = enviarCorreosDia_(serviciosCorreo_(), { simular: true });
  informarCorreos_(r);
  return resumenCorreos_(r);
}

function activarCorreosCarga() {
  soloPropietario_();
  const faltan = problemasPropiedadesCorreo_(PropertiesService.getScriptProperties());
  const admin = serviciosCarga_().admin();
  const hoja = admin.getSheetByName(CORREO.HOJA);
  if (!hoja) faltan.push('Falta la hoja ' + CORREO.HOJA + ' en la planilla de administración (ejecutar prepararAdministracionCarga).');
  else { try { correosValidar_(hoja); } catch (e) { faltan.push(e.message); } }
  if (faltan.length) throw new Error('No se activaron los correos: ' + faltan.join(' '));
  const borrados = borrarActivadoresCorreo_();
  const t = ScriptApp.newTrigger(CORREO.FUNCION).timeBased().atHour(CORREO.HORA_ENVIO).everyDays(1).inTimezone(CARGA.ZONA).create();
  console.log('Correos activados (activador ' + t.getUniqueId() + (borrados ? '; reemplaza ' + borrados + ' anterior' + (borrados > 1 ? 'es' : '') : '') +
    '). Corre todos los días entre las ' + CORREO.HORA_ENVIO + ':00 y las ' + (CORREO.HORA_ENVIO + 1) + ':00; sólo envía en días hábiles.');
}

function desactivarCorreosCarga() {
  soloPropietario_();
  console.log('Activadores de correos borrados: ' + borrarActivadoresCorreo_() + '.');
}

// ---------------------------------------------------------------- auxiliares

/** La llamada viene del activador de este proyecto (por su triggerUid) o de la cuenta propietaria. */
function autorizarCorreos_(e) {
  const uid = e && typeof e === 'object' && (typeof e.triggerUid === 'string' || typeof e.triggerUid === 'number') ? String(e.triggerUid) : '';
  if (uid && ScriptApp.getProjectTriggers().some(t => t.getUniqueId() === uid && t.getHandlerFunction() === CORREO.FUNCION)) return;
  soloPropietario_();
}

function borrarActivadoresCorreo_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === CORREO.FUNCION) { ScriptApp.deleteTrigger(t); n++; } });
  return n;
}

function serviciosCorreo_() {
  const base = serviciosCarga_();
  const propiedades = PropertiesService.getScriptProperties();
  const inicio = Date.now();
  return Object.assign(base, {
    correos: () => {
      const h = base.admin().getSheetByName(CORREO.HOJA);
      if (!h) throw new Error('Falta la hoja ' + CORREO.HOJA + ' en la planilla de administración (ejecutar prepararAdministracionCarga).');
      return h;
    },
    urlCarga: () => propiedades.getProperty(CORREO.PROPIEDAD_URL) || '',
    enviar: m => MailApp.sendEmail({ to: m.to, subject: m.subject, body: m.body, htmlBody: m.htmlBody, name: m.name }),
    cuota: () => MailApp.getRemainingDailyQuota(),
    ms: () => Date.now() - inicio
  });
}

function problemasPropiedadesCorreo_(propiedades) {
  const p = [];
  const clave = propiedades.getProperty(ACCESO.PROPIEDAD_CLAVE);
  if (!clave || clave.length < ACCESO.MIN_CLAVE) p.push('Falta la propiedad ' + ACCESO.PROPIEDAD_CLAVE + ' (ejecutar prepararAdministracionCarga).');
  if (!CORREO.URL_APP.test(txt_(propiedades.getProperty(CORREO.PROPIEDAD_URL)))) {
    p.push('La propiedad ' + CORREO.PROPIEDAD_URL + ' debe tener la URL /exec de la app web publicada (https://script.google.com/…/exec).');
  }
  return p;
}

/** Una línea JSON por evento en el registro de ejecución. Nunca incluye enlaces ni el contenido de los mensajes. */
function informarCorreos_(r) {
  console.log(JSON.stringify(resumenCorreos_(r)));
  r.enviados.forEach(x => console.log(JSON.stringify({ correo: 'enviado', a: x.correo, tipos: x.tipos, periodos: x.periodos, pendientes: x.pendientes, bytes: x.bytes })));
  r.mensajes.forEach(x => console.log(JSON.stringify({ correo: 'simulado', a: x.correo, tipos: x.tipos, periodos: x.periodos, pendientes: x.pendientes, asunto: x.asunto, bytes: x.bytes })));
  r.errores.forEach(x => console.error(JSON.stringify({ correo: 'error', a: x.correo, error: x.error })));
  r.sinProcesar.forEach(c => console.error(JSON.stringify({ correo: 'sin procesar', a: c, motivo: 'tiempo o cuota agotados: volver a ejecutar enviarCorreosCarga hoy' })));
  r.problemas.forEach(p => console.error(JSON.stringify({ correo: 'problema', detalle: p })));
}

function resumenCorreos_(r) {
  return { dia: r.dia, simulado: r.simulado, motivo: r.motivo, enviados: r.enviados.length, simulados: r.mensajes.length, errores: r.errores.length,
    sinProcesar: r.sinProcesar.length, yaRecibieron: r.yaRecibieron.length, problemas: r.problemas.length };
}

/** Problemas de acceso y correos para verificarCargaDocente. */
function problemasAccesoCorreo_(servicios) {
  const problemas = problemasPropiedadesCorreo_(PropertiesService.getScriptProperties());
  try {
    const datos = servicios.datos();
    problemasEnlaces_(configuracion_(datos, servicios.reloj()), datos.enlaces).forEach(p => problemas.push(p));
  } catch (e) { problemas.push(e.message); }
  try { correosValidar_(servicios.correos()); } catch (e) { problemas.push(e.message); }
  try {
    const activadores = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === CORREO.FUNCION).length;
    if (activadores !== 1) problemas.push(activadores ? 'Hay ' + activadores + ' activadores de correos: ejecutar activarCorreosCarga para dejar uno.' : 'Los correos no están activados (activarCorreosCarga).');
  } catch (e) { problemas.push('No se pudieron leer los activadores: ' + e.message); }
  return problemas;
}
