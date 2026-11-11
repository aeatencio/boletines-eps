/**
 * Carga docente EPS (#27, #28) · app web: la página docente y sus dos llamadas.
 *
 * Identidad (#28). La decide sólo el servidor, en cada llamada, con la credencial opaca del enlace individual
 * (parámetro «k», que doGet pasa escapado a la página y la página devuelve en cada llamada). identificarDocente_ la
 * verifica contra los enlaces vigentes (Carga_Acceso.gs) antes de abrir ninguna grilla; una credencial mal formada se
 * descarta sin leer nada. La página nunca envía un correo ni otra identidad, y nada de lo que envía se usa como
 * identidad sin esa verificación. En el guardado, la credencial se vuelve a verificar dentro del lock (gancho
 * «acceso» de guardarDocente_): una revocación ocurrida mientras se esperaba el lock rechaza el guardado sin escribir.
 * Las asignaciones vigentes siguen decidiendo qué puede ver y cargar cada docente (lo comprueba el núcleo).
 *
 * Despliegue (appsscript.json, sección webapp): se ejecuta como la cuenta propietaria; los docentes no necesitan
 * acceso a las grillas ni a la administración. Quién puede abrir la app («access») se confirma en el piloto en Google.
 *
 * Funciones invocables por google.script.run (sin "_"): cargaConsultar y cargaGuardar. El pedido de guardado se
 * reduce a sus campos conocidos: un campo de más (por ejemplo, un correo) nunca llega al núcleo.
 *
 * Archivos HTML del proyecto: Carga_Pagina (la página) y Carga_Cliente (su lógica sin DOM: cola de guardado,
 * situación de la Calificación y modelo de la consulta), que la página incluye.
 */
const CARGA_WEB = {
  PARAMETRO: 'k',
  MAX_CREDENCIAL: 512,
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
  const acceso = accesoCarga_();
  return consultarInterfaz_(acceso.identificar, correo => consultarCarga_(correo, acceso.servicios), credencial);
}

/** Guarda un campo para el docente identificado por la credencial (ver guardarDocente_). */
function cargaGuardar(credencial, pedido) {
  const acceso = accesoCarga_();
  return guardarInterfaz_(acceso.identificar, (correo, p) => guardarCampoCarga_(correo, p, acceso.servicios), credencial, pedido);
}

/**
 * Identificación de una llamada: los servicios de Google se crean sólo si la credencial tiene la forma de una
 * credencial, y la misma ejecución los reutiliza para consultar o guardar.
 */
function accesoCarga_() {
  const a = { servicios: null };
  a.identificar = credencial => {
    if (!ACCESO.FORMATO.test(credencial)) return null;
    a.servicios = serviciosCarga_();
    return identificarDocente_(a.servicios, credencial);
  };
  return a;
}

/**
 * #28. Docente verificado para una credencial opaca: { correo } o null. Lee la configuración (Ventanas, Feriados,
 * Asignaciones, Enlaces) y nunca una grilla. Si la credencial vale:
 *  - la consulta o el guardado reutilizan esa misma lectura (servicios.reutilizarDatos), sin leer dos veces;
 *  - servicios.acceso queda listo para que el núcleo la vuelva a comprobar con cada configuración que lea, también
 *    dentro del lock: si dejó de valer (revocada, vencida o el docente ya no tiene asignaciones), rechazo «acceso».
 */
function identificarDocente_(servicios, credencial) {
  if (typeof credencial !== 'string' || !ACCESO.FORMATO.test(credencial)) return null;
  const datos = servicios.datos();
  const correo = docenteDeCredencial_(configuracion_(datos, servicios.reloj()), datos.enlaces, credencial, servicios.firmar, servicios.reloj());
  if (!correo) return null;
  servicios.reutilizarDatos(datos);
  servicios.acceso = (d, config) => {
    if (!credencialVigentePara_(config, d.enlaces, credencial, correo, servicios.firmar, servicios.reloj())) throw rechazo_('acceso', ACCESO.MOTIVO);
  };
  return { correo };
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
  if (!docente) return { estado: 'rechazado', codigo: 'acceso', motivo: ACCESO.MOTIVO };
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
  if (!docente) return { estado: 'rechazado', codigo: 'acceso', motivo: ACCESO.MOTIVO };
  return guardar(docente.correo, pedidoDePagina_(pedido));
}

/** Sólo los campos conocidos del pedido y sólo si son texto; lo demás no llega al núcleo (que valida el resto). */
function pedidoDePagina_(pedido) {
  if (!pedido || typeof pedido !== 'object') return null;
  const limpio = {};
  CARGA_WEB.CAMPOS_PEDIDO.forEach(k => { if (typeof pedido[k] === 'string') limpio[k] = pedido[k]; });
  return limpio;
}
