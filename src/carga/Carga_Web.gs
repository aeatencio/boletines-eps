/**
 * Carga docente EPS (#27) · app web: la página docente y sus dos llamadas.
 *
 * ┌──────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
 * │ LÍMITE PENDIENTE DE #28. NO PUBLICAR ESTE PROYECTO COMO APP WEB HASTA RESOLVERLO.                              │
 * │                                                                                                                │
 * │ La identidad del docente la decide sólo el servidor, en identificarDocente_(credencial), a partir de la        │
 * │ credencial opaca que llega en cada llamada (hoy, el parámetro «k» del enlace, que doGet pasa a la página).     │
 * │ La página nunca envía un correo ni otra identidad, y nada de lo que envía se usa como identidad sin esa         │
 * │ verificación. #28 debe implementar identificarDocente_ (enlace individual, estable durante la ventana y         │
 * │ revocable) y decidir el despliegue (ejecutar como la cuenta propietaria; quién puede abrir la app).            │
 * │                                                                                                                │
 * │ Mientras tanto identificarDocente_ rechaza toda credencial: la página carga, pero la consulta y el guardado     │
 * │ responden «acceso» sin abrir ninguna planilla. Para desarrollar y probar la interfaz, tests/carga-dev.js y      │
 * │ tests/carga-interfaz.js sustituyen sólo esta función por credenciales aleatorias creadas por el propio          │
 * │ servidor local, sobre datos ficticios. El núcleo no cambia: sigue recibiendo el correo ya identificado.        │
 * └──────────────────────────────────────────────────────────────────────────────────────────────────────────────┘
 *
 * Funciones invocables por google.script.run (sin "_"): cargaConsultar y cargaGuardar. Ambas identifican al
 * docente en cada llamada, antes de tocar ninguna planilla. El pedido de guardado se reduce a sus campos
 * conocidos: un campo de más (por ejemplo, un correo) nunca llega al núcleo.
 *
 * Archivos HTML del proyecto: Carga_Pagina (la página) y Carga_Cliente (su lógica sin DOM: cola de guardado,
 * situación de la Calificación y modelo de la consulta), que la página incluye.
 */
const CARGA_WEB = {
  PARAMETRO: 'k',
  MAX_CREDENCIAL: 512,
  ACCESO: 'No se pudo verificar tu acceso. Abrí la carga desde el enlace de tu último correo.',
  CAMPOS_PEDIDO: ['operacion', 'periodo', 'fuente', 'orientacion', 'materia', 'dni', 'campo', 'esperado', 'nuevo']
};

function doGet(e) {
  const pagina = HtmlService.createTemplateFromFile('Carga_Pagina');
  pagina.credencial = credencialDe_(e && e.parameter && e.parameter[CARGA_WEB.PARAMETRO]);
  return pagina.evaluate()
    .setTitle('Carga de valoraciones EPS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Contenido de otro archivo HTML del proyecto, para incluirlo en la página: <?!= incluirCarga_('Carga_Cliente'); ?> */
function incluirCarga_(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

// ---------------------------------------------------------------- llamadas de la página

/** Lo que el docente identificado por la credencial puede ver y cargar ahora (ver consultarDocente_). */
function cargaConsultar(credencial) {
  return consultarInterfaz_(identificarDocente_, consultarCarga_, credencial);
}

/** Guarda un campo para el docente identificado por la credencial (ver guardarDocente_). */
function cargaGuardar(credencial, pedido) {
  return guardarInterfaz_(identificarDocente_, guardarCampoCarga_, credencial, pedido);
}

/**
 * LÍMITE #28. Docente verificado para una credencial opaca: { correo } o null.
 *
 * Hoy no existe ningún mecanismo de acceso: rechaza todo, sin mirar la credencial. #28 la reemplaza. Requisitos
 * para el reemplazo: verificar en el servidor, en cada llamada, que la credencial es vigente y de ese docente; no
 * aceptar nunca una identidad que el cliente pueda elegir; que una credencial revocada deje de valer en la llamada
 * siguiente. Las asignaciones vigentes siguen decidiendo qué puede cargar (lo comprueba el núcleo, también dentro
 * del lock). Si la revocación tiene que valer también durante la espera del lock, el núcleo ya ofrece el lugar:
 * el gancho revalidar de guardarDocente_.
 */
function identificarDocente_(credencial) {
  return null;
}

// ---------------------------------------------------------------- lógica sin servicios de Google (probada en Node)

function credencialDe_(valor) {
  const t = typeof valor === 'string' ? valor.trim() : '';
  return t.length <= CARGA_WEB.MAX_CREDENCIAL ? t : '';
}

/** { correo } del docente verificado, o null. Una credencial vacía no se consulta. */
function docenteVerificado_(identificar, credencial) {
  const c = credencialDe_(credencial);
  if (!c) return null;
  const docente = identificar(c);
  const correo = docente && typeof docente.correo === 'string' ? correo_(docente.correo) : '';
  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo) ? { correo } : null;
}

function consultarInterfaz_(identificar, consultar, credencial) {
  let docente;
  try {
    docente = docenteVerificado_(identificar, credencial);
  } catch (e) {
    console.error(JSON.stringify({ acceso: 'error', error: e.message }));
    return { estado: 'error', motivo: 'No se pudo verificar tu acceso. Volvé a intentarlo.' };
  }
  if (!docente) return { estado: 'rechazado', codigo: 'acceso', motivo: CARGA_WEB.ACCESO };
  const respuesta = consultar(docente.correo);
  if (respuesta && respuesta.estado === 'ok') respuesta.limites = { observaciones: CARGA.MAX_OBSERVACIONES };
  return respuesta;
}

/**
 * Sin identidad verificada, «rechazado» con código «acceso»: como todo rechazo, garantiza que no se intentó
 * escribir. Un error al verificar es «reintentar»: tampoco se llegó a escribir.
 */
function guardarInterfaz_(identificar, guardar, credencial, pedido) {
  let docente;
  try {
    docente = docenteVerificado_(identificar, credencial);
  } catch (e) {
    console.error(JSON.stringify({ acceso: 'error', error: e.message }));
    return { estado: 'reintentar', motivo: 'No se pudo verificar tu acceso. Se reintenta.' };
  }
  if (!docente) return { estado: 'rechazado', codigo: 'acceso', motivo: CARGA_WEB.ACCESO };
  return guardar(docente.correo, pedidoDePagina_(pedido));
}

/** Sólo los campos conocidos del pedido y sólo si son texto; lo demás no llega al núcleo (que valida el resto). */
function pedidoDePagina_(pedido) {
  if (!pedido || typeof pedido !== 'object') return null;
  const limpio = {};
  CARGA_WEB.CAMPOS_PEDIDO.forEach(k => { if (typeof pedido[k] === 'string') limpio[k] = pedido[k]; });
  return limpio;
}
