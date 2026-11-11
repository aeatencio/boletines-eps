/**
 * Carga docente EPS (#26) · adaptador a los servicios de Google del núcleo (Carga_Nucleo.gs).
 *
 * Propiedades del script:
 *   ADMIN_ID      planilla de administración de la carga: Ventanas, Feriados, Asignaciones y Registro
 *   GENERADOR_ID  planilla del generador; sólo se lee Control!B10:B14 (la grilla de cada período)
 *
 * Todo se lee de nuevo en cada operación: el Equipo EPS puede cambiar ventanas, feriados, asignaciones y fuentes
 * con la carga abierta. Las grillas se abren por el ID que figura en Control, nunca por otro camino.
 *
 * consultarCarga_ y guardarCampoCarga_ reciben el correo del docente ya identificado y, por terminar en "_",
 * google.script.run no puede invocarlas. La página (#27, Carga_Web.gs) las llama después de identificar al docente
 * en el servidor; la identificación es #28. Las demás funciones sin "_" de este archivo son para el propietario,
 * desde el editor.
 */

function serviciosCarga_() {
  const propiedades = PropertiesService.getScriptProperties();
  const adminId = propiedades.getProperty('ADMIN_ID');
  const generadorId = propiedades.getProperty('GENERADOR_ID');
  if (!adminId || !generadorId) throw new Error('Faltan las propiedades del script ADMIN_ID o GENERADOR_ID.');
  let admin = null, generador = null;
  const abrirAdmin = () => admin || (admin = SpreadsheetApp.openById(adminId));
  // Abiertas una vez por ejecución: la revalidación dentro del lock sólo vuelve a leer sus celdas.
  const abrirGenerador = () => generador || (generador = SpreadsheetApp.openById(generadorId));
  const libros = {};
  const hoja = (fuente, orientacion) => {
    if (!CARGA.HOJAS_ORIGEN.includes(orientacion)) throw rechazo_('pedido', 'Orientación desconocida.');
    const libro = libros[fuente] || (libros[fuente] = SpreadsheetApp.openById(fuente));
    const h = libro.getSheetByName(orientacion);
    if (!h) throw rechazo_('grilla', 'No se encontró la hoja ' + orientacion + ' en la grilla. Avisá al Equipo EPS.');
    return h;
  };
  return {
    reloj: () => new Date(),
    datos: () => datosCarga_(abrirAdmin(), abrirGenerador()),
    hoja,
    valoresHoja: (fuente, orientacion) => hoja(fuente, orientacion).getDataRange().getDisplayValues(),
    registro: () => {
      const r = abrirAdmin().getSheetByName(CARGA.HOJAS.registro);
      if (!r) throw rechazo_('registro', 'Falta la hoja ' + CARGA.HOJAS.registro + ' en la planilla de administración. Avisá al Equipo EPS.');
      return r;
    },
    lock: () => LockService.getScriptLock(),
    flush: () => SpreadsheetApp.flush()
  };
}

/**
 * Datos crudos de la configuración. Las fechas de Ventanas y Feriados se convierten a 'aaaa-mm-dd' con la zona
 * horaria de la planilla de administración (la misma con la que Sheets las muestra); el texto queda como está.
 */
function datosCarga_(admin, generador) {
  const zona = admin.getSpreadsheetTimeZone();
  const tabla = nombre => {
    const h = admin.getSheetByName(nombre);
    if (!h) return null;
    return h.getDataRange().getValues().map(f => f.map(v =>
      Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, zona, 'yyyy-MM-dd') : v));
  };
  const control = generador.getSheetByName(CARGA.HOJA_CONTROL);
  if (!control) throw rechazo_('configuracion', 'No se encontró la hoja Control del generador. Avisá al Equipo EPS.');
  // Una sola lectura de las celdas de las fuentes (B10:B14), como las lee el generador (valor mostrado).
  if (CARGA.PERIODOS.some(p => CARGA.CELDAS_FUENTE[p][0] !== 'B')) throw new Error('Las celdas de las fuentes deben estar en la columna B de Control.');
  const filas = CARGA.PERIODOS.map(p => Number(CARGA.CELDAS_FUENTE[p].slice(1)));
  const primera = Math.min.apply(null, filas);
  const valores = control.getRange(primera, 2, Math.max.apply(null, filas) - primera + 1, 1).getDisplayValues();
  const fuentes = {};
  CARGA.PERIODOS.forEach((p, i) => { fuentes[p] = valores[filas[i] - primera][0]; });
  return { ventanas: tabla(CARGA.HOJAS.ventanas), feriados: tabla(CARGA.HOJAS.feriados), asignaciones: tabla(CARGA.HOJAS.asignaciones), fuentes };
}

/** Lo que el docente puede ver y cargar ahora (ver consultarDocente_). */
function consultarCarga_(correo) {
  try {
    return consultarDocente_(serviciosCarga_(), correo);
  } catch (e) {
    console.error(JSON.stringify({ consulta: 'error', error: e.message }));
    return { estado: 'error', motivo: e.rechazo ? e.message : 'No se pudo leer la carga. Volvé a intentarlo; si persiste, avisá al Equipo EPS.' };
  }
}

/** Guarda un campo (ver guardarDocente_). Agrega la duración y deja una línea en el registro de ejecución. */
function guardarCampoCarga_(correo, pedido) {
  const inicio = Date.now();
  let respuesta;
  try {
    respuesta = guardarDocente_(serviciosCarga_(), correo, pedido);
  } catch (e) {
    respuesta = { estado: 'reintentar', motivo: 'No se pudo guardar. Se reintenta; si persiste, avisá al Equipo EPS.', error: e.message };
  }
  respuesta.ms = Date.now() - inicio;
  console.log(JSON.stringify({ guardado: respuesta.estado, codigo: respuesta.codigo, ms: respuesta.ms, lockMs: respuesta.lockMs, error: respuesta.error }));
  return respuesta;
}

// ---------------------------------------------------------------- para el propietario, desde el editor

/** Crea en la planilla de administración las hojas que falten, sólo con sus encabezados. Nunca toca una existente. */
function prepararAdministracionCarga() {
  soloPropietario_();
  const admin = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('ADMIN_ID'));
  [[CARGA.HOJAS.ventanas, CARGA.ENCABEZADOS.ventanas], [CARGA.HOJAS.feriados, CARGA.ENCABEZADOS.feriados],
    [CARGA.HOJAS.asignaciones, CARGA.ENCABEZADOS.asignaciones], [CARGA.HOJAS.registro, CARGA.REGISTRO]].forEach(([nombre, enc]) => {
    if (admin.getSheetByName(nombre)) return;
    admin.insertSheet(nombre).getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold');
  });
  console.log('Listo. Completar Ventanas, Feriados y Asignaciones; proteger el Registro para que sólo lo edite esta cuenta.');
}

/** Problemas de configuración, de las grillas de las ventanas vigentes y del Registro (ver verificarConfiguracion_). */
function verificarCargaDocente() {
  soloPropietario_();
  const r = verificarConfiguracion_(serviciosCarga_());
  console.log(['Ventanas vigentes o próximas:'].concat(r.ventanas.length ? r.ventanas : ['(ninguna)'], '', 'Problemas:',
    r.problemas.length ? r.problemas : ['(ninguno)']).join('\n'));
  return r;
}

/** Resuelve, mirando la celda, las filas 'pendiente' del Registro con más de 10 minutos (ver resolverPendientes_). */
function resolverPendientesCarga() {
  soloPropietario_();
  const r = resolverPendientes_(serviciosCarga_(), 10 * 60 * 1000);
  console.log(JSON.stringify(r, null, 1));
  return r;
}

// Un visitante de una app web no es el propietario: getActiveUser es vacío o distinto.
function soloPropietario_() {
  const propietario = Session.getEffectiveUser().getEmail();
  if (!propietario || Session.getActiveUser().getEmail() !== propietario) throw new Error('Sólo el propietario del proyecto puede ejecutar esta función.');
}
