/**
 * Carga docente EPS (#28) · acceso individual por enlace. Sin servicios de Google: la firma (HMAC) llega por
 * parámetro, así que corre igual en Apps Script y en las pruebas locales (tests/carga-acceso.js).
 *
 * El enlace de cada docente lleva una credencial opaca (parámetro «k»):
 *
 *     credencial = base64url( HMAC-SHA256( CLAVE_ENLACES, «correo · versión · tramo» ) )      (43 caracteres)
 *
 *  - CLAVE_ENLACES es una propiedad del script: sólo la ve la cuenta propietaria. La credencial no se guarda en
 *    ninguna planilla; se recalcula en cada llamada.
 *  - correo: el de Asignaciones. La credencial no lo contiene: el servidor prueba con los docentes asignados.
 *  - versión: la de la hoja Enlaces (1 si el docente no figura). Para revocar un enlace, el Equipo EPS agrega al
 *    docente o aumenta su versión: el enlace anterior deja de valer en la llamada siguiente, también si un guardado
 *    está esperando el lock, y el docente recibe el nuevo en el próximo correo (Carga_Correo.gs).
 *  - tramo: el «Desde» más antiguo del tramo de ventanas vigente (ventanas que se superponen forman un solo tramo;
 *    2º cuatrimestre y anual, con las mismas fechas, son un tramo). El enlace es estable durante todo el tramo, sigue
 *    valiendo después del cierre (la página muestra que no hay cargas abiertas y el núcleo no deja escribir) y vence
 *    solo cuando empieza el tramo siguiente. Cambiar el «Desde» de una ventana ya abierta cambia los enlaces.
 *
 * Fallas siempre cerradas: sin la hoja Enlaces, o con una versión ilegible para un docente, ese acceso no vale.
 * El enlace acredita posesión, no identidad (decisión de #13): quien tenga el enlace actúa como ese docente, y las
 * asignaciones vigentes siguen decidiendo qué puede ver y cargar.
 */
const ACCESO = {
  PROPIEDAD_CLAVE: 'CLAVE_ENLACES',
  MIN_CLAVE: 32,
  FORMATO: /^[A-Za-z0-9_-]{43}$/,
  CONTEXTO: 'boletines-eps/carga/enlace/v1',
  MOTIVO: 'El enlace no es válido, fue reemplazado o es de una carga anterior. Abrí la carga desde el enlace de tu último correo.'
};

/**
 * Versiones de la hoja Enlaces: { versiones: { correo: n }, bloqueados: { correo: motivo }, problema, advertencias }.
 * Sin hoja (o sin sus columnas), problema: ningún enlace vale (borrar la hoja no debe deshacer una revocación).
 * Una fila con versión ilegible bloquea a ese docente; dos filas del mismo docente valen por la mayor versión.
 */
function versionesEnlaces_(filas) {
  const r = { versiones: {}, bloqueados: {}, problema: '', advertencias: [] };
  let tabla;
  try { tabla = tablaAdministracion_(filas, 'enlaces'); } catch (e) { r.problema = e.message; return r; }
  tabla.forEach(f => {
    const correo = correo_(f['Correo']);
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) {
      r.advertencias.push('Enlaces, fila ' + f.fila + ': correo no reconocido ("' + f['Correo'] + '"). Esa fila no revoca nada.');
      return;
    }
    const t = txt_(f['Versión']);
    const v = /^\d{1,6}$/.test(t) ? Number(t) : NaN;
    if (!(v >= 1)) {
      r.bloqueados[correo] = 'Enlaces, fila ' + f.fila + ': la versión de ' + correo + ' debe ser un número entero desde 1 ("' + t + '"). Su enlace no vale hasta corregirla.';
      return;
    }
    r.versiones[correo] = Math.max(r.versiones[correo] || 1, v);
  });
  return r;
}

/**
 * Tramo de acceso vigente: { desde, inicio, fin } del último tramo ya empezado, o null si no empezó ninguno.
 * Un tramo une las ventanas que se superponen o se tocan (de su apertura a su cierre). Si un cierre no se pudo
 * calcular (faltan feriados), para unir se usa el día siguiente a «Hasta».
 */
function tramoAcceso_(config, ahora) {
  const ventanas = CARGA.PERIODOS.map(n => config.periodos[n].ventana).filter(Boolean)
    .map(v => ({ desde: v.desde, inicio: v.apertura.getTime(), fin: (v.cierre || instanteLocal_(diaSiguiente_(v.hasta), CARGA.HORA_CIERRE)).getTime() }))
    .sort((a, b) => a.inicio - b.inicio || (a.desde < b.desde ? -1 : 1));
  let actual = null;
  ventanas.forEach(v => {
    if (v.inicio > ahora.getTime()) return;
    if (actual && v.inicio <= actual.fin) actual.fin = Math.max(actual.fin, v.fin);
    else actual = { desde: v.desde, inicio: v.inicio, fin: v.fin };
  });
  return actual;
}

/** Credencial del docente para una versión y un tramo. firmar(texto) → base64url sin relleno de HMAC-SHA256. */
function credencialDocente_(firmar, correo, version, tramo) {
  return firmar([ACCESO.CONTEXTO, correo_(correo), String(version), tramo.desde].join('\n'));
}

/** Docentes que pueden tener enlace: los de las filas válidas de Asignaciones, con el nombre de su primera fila. */
function docentesAsignados_(config) {
  const vistos = {};
  const docentes = [];
  config.asignaciones.forEach(a => {
    if (vistos[a.correo]) return;
    vistos[a.correo] = true;
    docentes.push({ correo: a.correo, docente: a.docente });
  });
  return docentes;
}

/**
 * Enlace vigente de un docente: { credencial, version, tramo } o { problema }. Lo usan los correos; un docente que
 * no está en Asignaciones, bloqueado en Enlaces o sin tramo empezado no tiene enlace.
 */
function enlaceVigente_(config, filasEnlaces, firmar, correo, ahora) {
  const c = correo_(correo);
  if (!docentesAsignados_(config).some(d => d.correo === c)) return { problema: c + ' no figura en Asignaciones.' };
  const versiones = versionesEnlaces_(filasEnlaces);
  if (versiones.problema) return { problema: versiones.problema };
  if (versiones.bloqueados[c]) return { problema: versiones.bloqueados[c] };
  const tramo = tramoAcceso_(config, ahora);
  if (!tramo) return { problema: 'No empezó ninguna ventana de carga.' };
  const version = versiones.versiones[c] || 1;
  return { credencial: credencialDocente_(firmar, c, version, tramo), version, tramo };
}

/**
 * Correo del docente dueño de la credencial, o '' si no corresponde a ningún enlace vigente. Una credencial mal
 * formada se descarta sin calcular nada. Compara todas las candidatas sin cortar en la primera.
 */
function docenteDeCredencial_(config, filasEnlaces, credencial, firmar, ahora) {
  if (typeof credencial !== 'string' || !ACCESO.FORMATO.test(credencial)) return '';
  const versiones = versionesEnlaces_(filasEnlaces);
  if (versiones.problema) return '';
  const tramo = tramoAcceso_(config, ahora);
  if (!tramo) return '';
  let hallado = '';
  docentesAsignados_(config).forEach(d => {
    if (versiones.bloqueados[d.correo]) return;
    if (igualesSeguro_(credencialDocente_(firmar, d.correo, versiones.versiones[d.correo] || 1, tramo), credencial)) hallado = d.correo;
  });
  return hallado;
}

/** La credencial sigue siendo la vigente de ese correo (una sola firma). Para revalidar dentro del lock. */
function credencialVigentePara_(config, filasEnlaces, credencial, correo, firmar, ahora) {
  const e = enlaceVigente_(config, filasEnlaces, firmar, correo, ahora);
  return !e.problema && igualesSeguro_(e.credencial, credencial);
}

/** Comparación de textos sin cortar en la primera diferencia. */
function igualesSeguro_(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

/** Problemas de la hoja Enlaces para el Equipo EPS (verificarCargaDocente): filas ilegibles o que no revocan a nadie. */
function problemasEnlaces_(config, filasEnlaces) {
  const v = versionesEnlaces_(filasEnlaces);
  if (v.problema) return [v.problema + ' Sin ella ningún enlace vale.'];
  const asignados = {};
  docentesAsignados_(config).forEach(d => { asignados[d.correo] = true; });
  return v.advertencias
    .concat(Object.keys(v.bloqueados).map(c => v.bloqueados[c]))
    .concat(Object.keys(v.versiones).filter(c => !asignados[c]).map(c => 'Enlaces: ' + c + ' no figura en Asignaciones (¿error de tipeo?). Esa fila no revoca a nadie.'));
}
