/**
 * Carga docente EPS (#26) · adaptador a los servicios de Google del núcleo (Carga_Nucleo.gs).
 *
 * Propiedades del script:
 *   ADMIN_ID       planilla de administración de la carga: Ventanas, Asignaciones, Enlaces, Registro y Correos
 *   GENERADOR_ID   planilla del generador; sólo se lee Control!B10:B14 (la grilla de cada período)
 *   CLAVE_ENLACES  secreto de los enlaces individuales (#28); lo crea prepararAdministracionCarga si falta
 *   URL_CARGA      URL /exec de la app web publicada (#28); la usan los correos para armar los enlaces
 *
 * Todo se lee de nuevo en cada operación: el Equipo EPS puede cambiar ventanas, asignaciones y fuentes con la carga
 * abierta. Los feriados están en el código (Carga_Calendario.gs); una hoja Feriados que haya quedado no se lee. Las grillas se abren por el ID que figura en Control, nunca por otro camino.
 *
 * consultarCarga_ y guardarCampoCarga_ reciben el correo del docente ya identificado y, por terminar en "_",
 * google.script.run no puede invocarlas. La página (#27, Carga_Web.gs) las llama después de identificar al docente
 * en el servidor con su enlace (#28). Los correos de #28 tienen su adaptador en Carga_Correo_Google.gs. Las demás
 * funciones sin "_" de este archivo son para el propietario, desde el editor.
 */

function serviciosCarga_() {
  const propiedades = PropertiesService.getScriptProperties();
  const adminId = propiedades.getProperty('ADMIN_ID');
  const generadorId = propiedades.getProperty('GENERADOR_ID');
  if (!adminId || !generadorId) throw new Error('Faltan las propiedades del script ADMIN_ID o GENERADOR_ID.');
  let admin = null, generador = null, leidos = null;
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
    // reutilizarDatos: la lectura con la que se identificó al docente sirve como la primera lectura de la operación,
    // que es casi simultánea; toda lectura posterior (la revalidación dentro del lock) es nueva.
    datos: () => {
      if (leidos) { const d = leidos; leidos = null; return d; }
      return datosCarga_(abrirAdmin(), abrirGenerador());
    },
    reutilizarDatos: d => { leidos = d; },
    firmar: firmaEnlaces_(propiedades),
    hoja,
    valoresHoja: (fuente, orientacion) => hoja(fuente, orientacion).getDataRange().getDisplayValues(),
    registro: () => {
      const r = abrirAdmin().getSheetByName(CARGA.HOJAS.registro);
      if (!r) throw rechazo_('registro', 'Falta la hoja ' + CARGA.HOJAS.registro + ' en la planilla de administración. Avisá al Equipo EPS.');
      return r;
    },
    admin: abrirAdmin,
    lock: () => LockService.getScriptLock(),
    flush: () => SpreadsheetApp.flush()
  };
}

/**
 * firmar(texto) de los enlaces (#28): HMAC-SHA256 con la propiedad CLAVE_ENLACES, en base64url sin relleno. Sin la
 * propiedad (o demasiado corta), lanza: ningún enlace se puede verificar ni emitir.
 */
function firmaEnlaces_(propiedades) {
  return texto => {
    const clave = propiedades.getProperty(ACCESO.PROPIEDAD_CLAVE);
    if (!clave || clave.length < ACCESO.MIN_CLAVE) throw new Error('Falta la propiedad del script ' + ACCESO.PROPIEDAD_CLAVE + ' (ejecutar prepararAdministracionCarga).');
    return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(texto, clave, Utilities.Charset.UTF_8)).replace(/=+$/, '');
  };
}

/**
 * Datos crudos de la configuración. Las fechas de Ventanas se convierten a 'aaaa-mm-dd' con la zona
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
  return { ventanas: tabla(CARGA.HOJAS.ventanas), asignaciones: tabla(CARGA.HOJAS.asignaciones), enlaces: tabla(CARGA.HOJAS.enlaces), fuentes };
}

/** Lo que el docente puede ver y cargar ahora (ver consultarDocente_). servicios: los de la identificación, si los hay. */
function consultarCarga_(correo, servicios) {
  try {
    return consultarDocente_(servicios || serviciosCarga_(), correo);
  } catch (e) {
    console.error(JSON.stringify({ consulta: 'error', error: e.message }));
    return { estado: 'error', motivo: e.rechazo ? e.message : 'No se pudo leer la carga. Volvé a intentarlo; si persiste, avisá al Equipo EPS.' };
  }
}

/**
 * Guarda un campo (ver guardarDocente_). Agrega la duración y deja una línea en el registro de ejecución.
 * servicios: los de la identificación (con su gancho «acceso»), si los hay.
 */
function guardarCampoCarga_(correo, pedido, servicios) {
  const inicio = Date.now();
  let respuesta;
  try {
    respuesta = guardarDocente_(servicios || serviciosCarga_(), correo, pedido);
  } catch (e) {
    respuesta = { estado: 'reintentar', motivo: 'No se pudo guardar. Se reintenta; si persiste, avisá al Equipo EPS.', error: e.message };
  }
  respuesta.ms = Date.now() - inicio;
  console.log(JSON.stringify({ guardado: respuesta.estado, codigo: respuesta.codigo, ms: respuesta.ms, lockMs: respuesta.lockMs, error: respuesta.error }));
  return respuesta;
}

// ---------------------------------------------------------------- para el propietario, desde el editor

/**
 * Crea en la planilla de administración las hojas que falten, sólo con sus encabezados, y el secreto de los
 * enlaces (CLAVE_ENLACES) si falta. Nunca toca una hoja existente ni reemplaza el secreto: borrarlo y volver a
 * ejecutar esta función invalida a la vez todos los enlaces.
 */
function prepararAdministracionCarga() {
  soloPropietario_();
  const propiedades = PropertiesService.getScriptProperties();
  const admin = serviciosCarga_().admin(); // sin ADMIN_ID o GENERADOR_ID, un error que las nombra
  [[CARGA.HOJAS.ventanas, CARGA.ENCABEZADOS.ventanas],
    [CARGA.HOJAS.asignaciones, CARGA.ENCABEZADOS.asignaciones], [CARGA.HOJAS.enlaces, CARGA.ENCABEZADOS.enlaces],
    [CARGA.HOJAS.registro, CARGA.REGISTRO], [CARGA.HOJAS.correos, CORREO.COLUMNAS]].forEach(([nombre, enc]) => {
    if (admin.getSheetByName(nombre)) return;
    admin.insertSheet(nombre).getRange(1, 1, 1, enc.length).setValues([enc]).setFontWeight('bold');
  });
  if (!propiedades.getProperty(ACCESO.PROPIEDAD_CLAVE)) {
    // getUuid usa un generador aleatorio seguro: dos UUID v4 son 244 bits.
    propiedades.setProperty(ACCESO.PROPIEDAD_CLAVE, (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''));
    console.log('Se creó la propiedad ' + ACCESO.PROPIEDAD_CLAVE + ' (secreto de los enlaces). No la copies a otro lugar.');
  }
  console.log('Listo. Completar Ventanas y Asignaciones; proteger Registro y Correos para que sólo los edite esta cuenta. ' +
    'Los feriados están en el código: calendario de ' + aniosCalendario_().join(', ') + '. ' +
    'Para los correos: publicar la app web, guardar su URL /exec en la propiedad ' + CORREO.PROPIEDAD_URL + ' y ejecutar activarCorreosCarga.');
}

/**
 * Problemas de configuración, de las grillas de las ventanas vigentes, del Registro, de los enlaces y de los correos.
 * La primera línea del registro dice qué años tiene el calendario de feriados del código.
 */
function verificarCargaDocente() {
  soloPropietario_();
  const servicios = serviciosCorreo_();
  const r = verificarConfiguracion_(servicios);
  r.problemas = r.problemas.concat(problemasAccesoCorreo_(servicios));
  console.log(['Calendario de feriados: ' + aniosCalendario_().join(', '), '', 'Ventanas vigentes o próximas:'].concat(r.ventanas.length ? r.ventanas : ['(ninguna)'], '', 'Grillas de esas ventanas:',
    r.grillas.length ? r.grillas : ['(ninguna)'], '', 'Problemas:', r.problemas.length ? r.problemas : ['(ninguno)']).join('\n'));
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
