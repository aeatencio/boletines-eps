/**
 * Carga docente EPS (#32) · preparación de las grillas de 2º cuatrimestre y anual y del borrador de Asignaciones.
 *
 * Los borradores de esas dos grillas se hicieron copiando la de 1º cuatrimestre y después se editaron a mano (entre
 * otras cosas, vaciando las calificaciones copiadas). Una sola ejecución desde el editor (prepararPeriodosCarga, en
 * Carga_Preparacion_Google.gs) los deja listos para la carga docente y deja en la planilla de administración un
 * informe de lo hecho y de lo que necesita una persona:
 *  - cursadas: CURSA donde 3º bimestre, el último período calificado y con la matrícula vigente, tiene una
 *    calificación ordinaria y el borrador dice que no cursa (vacía o NO CURSA, equivalentes) o conserva la
 *    calificación copiada de 1º cuatrimestre. Una celda vacía no es una decisión de dejar de cursar: los borradores se
 *    limpiaron vaciando. ACREDITÓ, EQUIVALENCIA y CURSA sin calificar en 3º bimestre no son evidencia. Anual se
 *    prepara con la misma evidencia: 2º cuatrimestre todavía no está calificado y no se usa;
 *  - restos de la copia: Calidad, Participación y Observaciones iguales a las de 1º cuatrimestre se borran. Ninguna
 *    calificación anterior pasa a ser una valoración nueva: una nota que no tiene evidencia se informa;
 *  - asignaciones: una fila por orientación y materia con cursantes en «Asignaciones (borrador)», con orientación,
 *    nivel y materia de la grilla. El Equipo EPS elige el docente de una lista (la hoja «Docentes», su nómina) y el
 *    correo sale de esa nómina por fórmula. No hay una fuente fiable para proponer docentes: no se adivinan;
 *  - revisión: matrícula, identidad, grupos, orientación, materias y niveles de cada borrador contra 3º bimestre y
 *    1º cuatrimestre, y de los dos borradores entre sí. El informe agrupa los casos por tipo.
 *
 * Conservación (se puede repetir sin deshacer trabajo ni duplicar filas):
 *  - no cambia ACREDITÓ, EQUIVALENCIA, CURSA, una nota que no viene de la copia ni un valor no reconocido;
 *  - una cursada que la preparación ya puso (hoja «Preparación · cambios») y alguien quitó después no se vuelve a poner;
 *  - nunca inserta, borra ni mueve filas o columnas, ni toca ORIENTACIÓN y los campos posteriores;
 *  - no toca a un estudiante en una materia en la que la carga docente ya escribió (Registro);
 *  - no usa evidencia de un estudiante que no se identifica en 3º bimestre (DNI ausente o repetido, otra orientación)
 *    ni de una materia que allí falta o tiene otro nivel;
 *  - con una ventana cargada para 2º cuatrimestre o anual, no modifica nada;
 *  - en el borrador de Asignaciones nunca cambia Orientación, Nivel, Materia, Docente ni Correo de una fila con datos;
 *    no escribe la hoja Asignaciones que lee la carga.
 *
 * Este archivo no usa servicios de Google: recibe las lecturas hechas por el adaptador y devuelve el plan, que las
 * pruebas locales (tests/carga-preparacion.js) ejercitan directamente.
 */
const PREPARACION = {
  OBJETIVOS: ['2º cuatrimestre', 'Anual'],
  // Grilla de la que se copiaron los borradores: lo que sigue igual a ella es resto de la copia.
  COPIA: '1º cuatrimestre',
  // Último período calificado, con la matrícula vigente: su calificación ordinaria es la evidencia de cursada.
  EVIDENCIA: '3º bimestre',
  HOJA_INFORME: 'Preparación',
  HOJA_CAMBIOS: 'Preparación · cambios',
  HOJA_BORRADOR: 'Asignaciones (borrador)',
  // Nómina de docentes del Equipo EPS. Columnas fijas (A y B): las usan la lista y la fórmula del borrador.
  HOJA_DOCENTES: 'Docentes',
  NOMINA: ['Docente', 'Correo'],
  // Columnas que la preparación agrega al borrador. La carga las ignora (lee por encabezado).
  EXTRA_BORRADOR: ['Cursantes', 'Revisión'],
  CAMBIOS: ['Fecha', 'Estado', 'Detalle', 'Período', 'Grilla', 'Hoja', 'Celda', 'DNI', 'Estudiante', 'Nivel', 'Materia', 'Campo', 'Anterior', 'Nuevo', 'Motivo'],
  INFORME: ['Prioridad', 'Tipo', 'Período', 'Hoja', 'Nivel', 'Materia', 'Fila', 'DNI', 'Estudiante', 'Detalle'],
  RESUMEN_TIPOS: ['Prioridad', 'Tipo', 'Casos', 'Qué hacer', 'Ejemplos'],
  // Bloqueo: impide preparar (todo, una hoja o un estudiante). Decidir: necesita una persona. Revisar: se conservó lo
  // del borrador, pero difiere de la evidencia o de lo esperado.
  PRIORIDADES: ['Bloqueo', 'Decidir', 'Revisar'],
  // Cada caso del informe tiene un tipo: [prioridad, título, qué hacer]. El informe los resume por tipo.
  TIPOS: {
    configuracion: ['Bloqueo', 'Configuración o grilla faltante', 'Corregir lo indicado y volver a ejecutar.'],
    hoja: ['Bloqueo', 'Hoja de grilla no reconocida', 'Corregir su estructura; hasta entonces esa hoja no se prepara.'],
    estudianteBloqueado: ['Bloqueo', 'Estudiante con DNI repetido o bloque incompleto en el borrador', 'Corregir la fila en la grilla; hasta entonces no se prepara.'],
    hojaAdministracion: ['Bloqueo', 'Borrador de asignaciones o nómina con columnas no reconocidas', 'Corregir los encabezados de la fila 1; hasta entonces no se escribe el borrador.'],
    notaSinEvidencia: ['Decidir', 'Calificación de 1º cuatrimestre sin evidencia de cursada en 3º bimestre', 'Reemplazarla en la grilla por CURSA, NO CURSA, ACREDITÓ o EQUIVALENCIA.'],
    notaPropia: ['Decidir', 'Calificación en el borrador que no viene de 1º cuatrimestre', 'Si no es una valoración del período, reemplazarla por CURSA o NO CURSA.'],
    valorNoReconocido: ['Decidir', 'Calificación no reconocida', 'Corregirla (CURSA, NO CURSA, ACREDITÓ o EQUIVALENCIA).'],
    sinDni: ['Decidir', 'Fila con nombre y sin DNI', 'Completar el DNI o quitar la fila.'],
    estudianteSinEvidencia: ['Decidir', 'Estudiante no identificado en 3º bimestre, con calificaciones de 1º cuatrimestre', 'Revisar DNI y orientación; decidir CURSA o NO CURSA.'],
    materiaSinEvidencia: ['Decidir', 'Materia sin evidencia en 3º bimestre (falta o tiene otro nivel)', 'Decidir CURSA o NO CURSA en esa columna; revisar el nivel.'],
    nivelEntrePeriodos: ['Decidir', 'Materia con nivel distinto en 2º cuatrimestre y Anual', 'Corregir el nivel en la grilla: la asignación es una sola.'],
    sinDocente: ['Decidir', 'Materias sin docente elegido', 'Elegir el docente en la columna Docente de «Asignaciones (borrador)» (lista de la hoja Docentes).'],
    asignacion: ['Decidir', 'Fila de asignación incompleta o inconsistente', 'Corregir la fila indicada de «Asignaciones (borrador)».'],
    nomina: ['Decidir', 'Nómina de docentes con datos repetidos o incompletos', 'Corregir la hoja Docentes.'],
    administrativo3b: ['Revisar', 'Vacía en el borrador; ACREDITÓ o EQUIVALENCIA en 3º bimestre', 'La preparación no copia estados administrativos: registrarlos en la grilla si siguen vigentes.'],
    administrativoBorrador: ['Revisar', 'ACREDITÓ o EQUIVALENCIA en el borrador; calificación ordinaria en 3º bimestre', 'Confirmar el estado administrativo.'],
    cursaSinEvidencia: ['Revisar', 'CURSA en el borrador sin calificación ordinaria en 3º bimestre', 'Confirmar que cursa.'],
    quitadaDespues: ['Revisar', 'Cursada que la preparación puso y después se quitó', 'Se respeta; confirmar que no fue un error.'],
    sinCursantes: ['Revisar', 'Materia con estudiantes calificados en 3º bimestre y sin cursantes en el borrador', 'Confirmar que la materia no se dicta en el período.'],
    complementario: ['Revisar', 'Calidad, Participación u Observaciones que no vienen de 1º cuatrimestre', 'Si no corresponde al período, borrarlo: el docente lo verá como cargado.'],
    matricula: ['Revisar', 'Diferencia de matrícula con 3º bimestre', 'Agregar o corregir estudiantes en el borrador si corresponde.'],
    identidad: ['Revisar', 'Nombre, grupo u orientación distintos de 3º bimestre o 1º cuatrimestre', 'Corregir si es un error.'],
    periodos: ['Revisar', 'Diferencia entre 2º cuatrimestre y Anual', 'Igualar las grillas si corresponde.'],
    materia: ['Revisar', 'Materia que falta en una de las grillas', 'Confirmar la estructura de las grillas.'],
    camposGenerales: ['Revisar', 'Campos generales u operativos iguales a 1º cuatrimestre', 'Revisarlos antes de generar boletines.'],
    registro: ['Revisar', 'Celda en la que la carga docente ya escribió', 'No se toca: revisarla si hace falta.']
  },
  MAX_FILAS_TEXTO: 15,
  MAX_TEXTO: 60,
  MAX_EJEMPLOS: 3
};

// ================================================================ lectura de una grilla

/**
 * Una hoja de una grilla, con la estructura que exigen la carga y el generador (estructuraGrilla_) y sus estudiantes
 * por DNI. Lanza si la estructura no se reconoce: sobre una hoja así no se prepara nada.
 */
function analizarHojaPrep_(valores, periodo) {
  const estructura = estructuraGrilla_(valores, periodo);
  const e1 = valores[0] || [], e2 = valores[1] || [];
  const cola = {};
  CARGA.ENCABEZADOS_COLA.forEach(enc => {
    const cols = [];
    for (let c = estructura.inicioCola + 1; c < Math.max(e1.length, e2.length); c++) if (norm_(e1[c]) === norm_(enc) || norm_(e2[c]) === norm_(enc)) cols.push(c);
    if (cols.length === 1) cola[enc] = cols[0];
  });
  const materias = {};
  estructura.materias.forEach(m => { materias[m.clave] = m; });
  const lista = [], porDni = {}, sinDni = [];
  for (let r = 2; r < valores.length; r++) {
    const f = valores[r] || [];
    const dni = dni_(f[2]);
    if (!dni) {
      const valoracion = estructura.materias.some(m => etiqueta_(f[m.columnaValor - 1]) === 'VALORACION');
      if (valoracion && !/^EPS\d+$/.test(norm_(f[0])) && (txt_(f[0]) || txt_(f[1]))) sinDni.push({ fila: r, nombre: [txt_(f[0]), txt_(f[1])].filter(Boolean).join(', ') });
      continue;
    }
    let bloque = '';
    try { validarBloque_(valores, estructura, r); } catch (e) { bloque = sinAvisoDocente_(e.message); }
    const est = { dni, fila: r, nombre: [txt_(f[0]), txt_(f[1])].filter(Boolean).join(', '), grupo: nivelDeFila_(valores, r), orientacion: txt_(f[estructura.inicioCola]), bloque };
    lista.push(est);
    (porDni[dni] = porDni[dni] || []).push(est);
  }
  return { valores, estructura, materias, cola, lista, porDni, sinDni };
}

/** Grilla de un período: { periodo, fuente, problema, hojas: { EPSI: análisis | null }, problemasHoja }. */
function analizarGrillaPrep_(periodo, conf, lectura) {
  const g = { periodo, fuente: conf.fuente, problema: '', hojas: {}, problemasHoja: {} };
  if (conf.problemasFuente.length) g.problema = conf.problemasFuente.join(' ');
  else if (!conf.fuente) g.problema = 'Falta la grilla de ' + periodo + ' en Control!' + CARGA.CELDAS_FUENTE[periodo] + ' del generador.';
  else if (!lectura || lectura.problema) g.problema = (lectura && lectura.problema) || 'No se leyó la grilla de ' + periodo + '.';
  if (g.problema) return g;
  CARGA.HOJAS_ORIGEN.forEach(h => {
    const valores = lectura.hojas[h];
    if (!valores) { g.hojas[h] = null; g.problemasHoja[h] = 'La grilla de ' + periodo + ' no tiene la hoja ' + h + '.'; return; }
    try { g.hojas[h] = analizarHojaPrep_(valores, periodo); } catch (e) { g.hojas[h] = null; g.problemasHoja[h] = sinAvisoDocente_(e.message); }
  });
  return g;
}

function sinAvisoDocente_(m) { return String(m).replace(/ (Volvé a abrir la carga|Avisá al Equipo EPS)\.$/, ''); }

/** Todas las apariciones de un DNI en una grilla: [{ hoja, est }]. */
function buscarPrep_(g, dni) {
  const r = [];
  if (!g || g.problema) return r;
  CARGA.HOJAS_ORIGEN.forEach(h => { const a = g.hojas[h]; if (a && a.porDni[dni]) a.porDni[dni].forEach(est => r.push({ hoja: h, est })); });
  return r;
}

/** El estudiante en esa hoja de la grilla, si figura una sola vez en toda la grilla y en esa hoja. */
function unicoEnHojaPrep_(g, hoja, dni) {
  const u = buscarPrep_(g, dni);
  return u.length === 1 && u[0].hoja === hoja ? u[0].est : null;
}

function celdaPrep_(a, est, m, desp) { return txt_((a.valores[est.fila + desp] || [])[m.columnaValor]); }
function quedanPrep_(n) { return (n === 1 ? 'Quedó ' : 'Quedaron ') + plural_(n, 'calificación', 'calificaciones') + ' de ' + PREPARACION.COPIA + ' sin preparar'; }
function recortePrep_(t) { t = txt_(t); return t.length > PREPARACION.MAX_TEXTO ? t.slice(0, PREPARACION.MAX_TEXTO - 1) + '…' : t; }
function filasTextoPrep_(filas) {
  const l = filas.slice(0, PREPARACION.MAX_FILAS_TEXTO).join(', ');
  return (filas.length === 1 ? 'fila ' : 'filas ') + l + (filas.length > PREPARACION.MAX_FILAS_TEXTO ? '…' : '');
}


// ================================================================ decisión por celda

/**
 * Calificación de un estudiante en una materia del borrador del período P.
 * D: valor del borrador. C0: el de 1º cuatrimestre (null si el estudiante o la materia no están allí: no viene de
 * la copia). E0: el de 3º bimestre, o null con «falla» si no hay evidencia comparable. quitada: la preparación ya
 * había puesto CURSA en esta celda y ahora no lo tiene (alguien lo quitó).
 * Devuelve { cursa, motivo } (cambiar a CURSA), { tipo, detalle } (informar, sin cambiar), { sinPreparar } (una nota
 * copiada sin evidencia comparable; la informa la línea del estudiante o de la materia) o {} (nada que hacer).
 */
function decidirCalificacionPrep_(P, D, C0, E0, falla, quitada) {
  const copia = PREPARACION.COPIA, ev = PREPARACION.EVIDENCIA;
  const d = situacionCalificacion_(P, D);
  const c = C0 === null ? 'ausente' : situacionCalificacion_(copia, C0);
  const residuo = c !== 'ausente' && (d === 'no cursa' ? c === 'no cursa' : norm_(D) === norm_(C0));
  if (d === 'a revisar') return { tipo: 'valorNoReconocido', detalle: 'Calificación «' + recortePrep_(D) + '» no reconocida para ' + P + '.' };
  if (d === 'valorada' && !residuo) {
    return { tipo: 'notaPropia', detalle: 'Calificación «' + D + '» en el borrador, que no viene de ' + copia + (c === 'ausente' ? '' : ' (allí: «' + (C0 || 'vacía') + '»)') +
      '. La carga la mostraría como ya calificada.' };
  }
  if (falla) return { sinPreparar: d === 'valorada' };
  const e = situacionCalificacion_(ev, E0);
  const enEv = 'en ' + ev + ' ' + evidenciaTextoPrep_(e, E0);
  if (d === 'valorada') {
    if (e === 'valorada') return { cursa: true, motivo: 'Calificación «' + D + '» copiada de ' + copia + '; en ' + ev + ' tiene «' + E0 + '».' };
    return { tipo: 'notaSinEvidencia', detalle: 'Quedó la calificación «' + D + '» de ' + copia + ' y ' + enEv + '.' };
  }
  if (d === 'pendiente') {
    if (e === 'administrativo' || e === 'no cursa') return { tipo: 'cursaSinEvidencia', detalle: 'CURSA en el borrador' + (residuo ? ' (copiado de ' + copia + ')' : '') + ', pero ' + enEv + '.' };
    return {};
  }
  if (d === 'administrativo') {
    if (e === 'valorada') return { tipo: 'administrativoBorrador', detalle: '«' + D + '» en el borrador, pero ' + enEv + '.' };
    return {};
  }
  // d === 'no cursa' (vacía o NO CURSA, equivalentes): con calificación ordinaria en 3º bimestre, cursa. Que en
  // 1º cuatrimestre tuviera una nota no cambia nada: vaciarla fue limpiar la copia, no decidir que dejó de cursar.
  if (e === 'valorada') {
    if (quitada) return { tipo: 'quitadaDespues', detalle: (D || 'Vacía') + ' en el borrador después de que la preparación puso CURSA; ' + enEv + '. Se conservó.' };
    const origen = c === 'ausente' ? 'la fila o la columna no venía de ' + copia : c === 'no cursa' ? 'en ' + copia + ' no la cursaba' : 'en ' + copia + ' tenía «' + C0 + '», que se había vaciado';
    return { cursa: true, motivo: 'En ' + ev + ' tiene «' + E0 + '»; ' + origen + '.' };
  }
  if (e === 'administrativo') return { tipo: 'administrativo3b', detalle: (D || 'Vacía') + ' en el borrador; ' + enEv + '.' };
  return {};
}

function evidenciaTextoPrep_(e, valor) {
  if (e === 'valorada') return 'tiene «' + valor + '»';
  if (e === 'administrativo') return 'figura «' + valor + '»';
  if (e === 'pendiente') return 'quedó CURSA, sin calificar';
  if (e === 'no cursa') return valor ? 'dice ' + valor : 'no la cursaba (vacía)';
  return 'tiene «' + recortePrep_(valor) + '», que no se reconoce';
}

// ================================================================ plan

/**
 * Plan completo, sin escribir nada. entrada: { ahora, datos (como datosCarga_), grillas: { período: { hojas: { EPSI:
 * valores | null, EPSG }, problema } }, borrador, nomina, cambiosPrevios, registro (matrices de esas hojas, o null) }.
 * Filas y columnas en base 0.
 */
function planificarPreparacion_(entrada) {
  const P_ = PREPARACION;
  const plan = { ahora: entrada.ahora, avisos: [], cambios: [], bloqueos: [], puedeAplicar: false, fuentes: {}, porPeriodo: {}, proyectado: {}, asignaciones: null };
  const aviso = x => {
    const t = P_.TIPOS[x.tipo];
    if (!t) throw new Error('Tipo de aviso desconocido: ' + x.tipo);
    plan.avisos.push(Object.assign({ periodo: '', hoja: '', nivel: '', materia: '', fila: '', dni: '', estudiante: '', detalle: '' }, x, { prioridad: t[0] }));
  };
  const general = detalle => { plan.bloqueos.push(detalle); aviso({ tipo: 'configuracion', detalle }); };

  const config = configuracion_(entrada.datos, entrada.ahora);
  config.problemas.forEach(general);
  ventanasObjetivoPrep_(entrada.datos).forEach(general);

  const g = {};
  [P_.COPIA, P_.EVIDENCIA].concat(P_.OBJETIVOS).forEach(p => {
    g[p] = analizarGrillaPrep_(p, config.periodos[p], (entrada.grillas || {})[p]);
    plan.fuentes[p] = g[p].fuente;
    const papel = p === P_.COPIA ? ' (de la que se copiaron los borradores)' : p === P_.EVIDENCIA ? ' (evidencia de cursada)' : '';
    if (g[p].problema) general(g[p].problema + (papel ? ' Es la grilla' + papel + ': sin ella no se prepara nada.' : ''));
    CARGA.HOJAS_ORIGEN.forEach(h => {
      if (!g[p].problemasHoja[h]) return;
      aviso({ tipo: 'hoja', periodo: p, hoja: h, detalle: g[p].problemasHoja[h] + (P_.OBJETIVOS.includes(p) ? ' No se prepara esta hoja.' : ' No se prepara la hoja ' + h + ' de los borradores.') });
    });
  });

  // Celdas en las que la carga docente ya escribió: Registro con período, fuente, orientación, DNI y materia.
  const tocadas = clavesTablaPrep_(entrada.registro, f => P_.OBJETIVOS.includes(normalizarPeriodo_(f('Período'))),
    f => [normalizarPeriodo_(f('Período')), txt_(f('Fuente')), normalizarHojaOrigen_(f('Orientación')), dni_(f('DNI')), clave_(f('Materia'))]);
  // Cursadas que una aplicación anterior ya puso: si ahora no están, alguien las quitó y no se vuelven a poner.
  const preparadas = clavesTablaPrep_(entrada.cambiosPrevios,
    f => ['aplicado', 'inferido'].includes(txt_(f('Estado'))) && norm_(f('Campo')) === norm_(CARGA.CAMPOS[0].nombre) && norm_(f('Nuevo')) === CARGA.PENDIENTE,
    f => [normalizarPeriodo_(f('Período')), txt_(f('Grilla')), normalizarHojaOrigen_(f('Hoja')), dni_(f('DNI')), clave_(f('Materia'))]);

  P_.OBJETIVOS.forEach(P => { if (!g[P].problema) prepararObjetivoPrep_(P, g, tocadas, preparadas, plan, aviso); });
  if (P_.OBJETIVOS.every(P => !g[P].problema)) compararObjetivosPrep_(g, plan, aviso);

  plan.asignaciones = planAsignacionesPrep_({ g, proyectado: plan.proyectado, borrador: entrada.borrador, nomina: entrada.nomina });
  plan.asignaciones.avisos.forEach(aviso);

  plan.g = g;
  plan.puedeAplicar = !plan.bloqueos.length;
  const orden = a => P_.PRIORIDADES.indexOf(a.prioridad) * 1000 + Object.keys(P_.TIPOS).indexOf(a.tipo);
  plan.avisos.sort((a, b) => orden(a) - orden(b));
  plan.conteo = {};
  P_.PRIORIDADES.forEach(p => { plan.conteo[p] = plan.avisos.filter(a => a.prioridad === p).length; });
  return plan;
}

/** Claves de las filas de una tabla con encabezado (Registro, Preparación · cambios) que cumplen la condición. */
function clavesTablaPrep_(matriz, condicion, clave) {
  const r = {};
  if (!matriz || matriz.length < 2) return r;
  const enc = matriz[0].map(norm_);
  matriz.slice(1).forEach(f => {
    const v = n => { const i = enc.indexOf(norm_(n)); return i < 0 ? '' : f[i]; };
    if (condicion(v)) r[clave(v).join('|')] = true;
  });
  return r;
}

/** Filas de Ventanas de los períodos a preparar: con una, la preparación no modifica nada. */
function ventanasObjetivoPrep_(datos) {
  let filas = [];
  try { filas = tablaAdministracion_(datos.ventanas, 'ventanas'); } catch (e) { return []; } // configuracion_ ya lo informa
  return filas.filter(v => PREPARACION.OBJETIVOS.includes(normalizarPeriodo_(v['Período']))).map(v => 'Ventanas, fila ' + v.fila + ': hay una ventana de carga de ' +
    normalizarPeriodo_(v['Período']) + '. La preparación se hace antes de cargar las ventanas de esos períodos: no se modificó nada.');
}

/**
 * Dentro del lock, antes de escribir: la configuración sigue sin problemas ni ventanas de esos períodos, y Control
 * apunta a las mismas grillas que se leyeron. Devuelve '' o el motivo para no escribir.
 */
function revalidarPrep_(datos, ahora, fuentes) {
  const config = configuracion_(datos, ahora);
  const motivos = config.problemas.concat(ventanasObjetivoPrep_(datos));
  Object.keys(fuentes).forEach(p => { if (config.periodos[p].fuente !== fuentes[p]) motivos.push('La grilla de ' + p + ' en Control del generador cambió desde la lectura.'); });
  return motivos.length ? motivos.join(' ') + ' No se modificó nada; volver a ejecutar.' : '';
}

/** Cursadas, restos de la copia, matrícula y materias del borrador de un período. */
function prepararObjetivoPrep_(P, g, tocadas, preparadas, plan, aviso) {
  const P_ = PREPARACION;
  const T = g[P], C = g[P_.COPIA], E = g[P_.EVIDENCIA];
  const res = { estudiantes: 0, aCursa: 0, aCursaVaciadas: 0, aCursaNotas: 0, borrados: 0, yaCursa: 0, administrativos: 0, sinPreparar: 0, iguales: 0, comparadas: 0, altas: 0, bajas: 0, hojas: {} };
  plan.porPeriodo[P] = res;
  const proy = plan.proyectado[P] = {};
  const sinPrepEst = {}, sinPrepMat = {}, calificados3b = {};

  CARGA.HOJAS_ORIGEN.forEach(h => {
    const t = T.hojas[h];
    if (!t) return;
    res.hojas[h] = t.lista.length;
    const c = C.problema ? null : C.hojas[h];
    const ev = E.problema ? null : E.hojas[h];
    const base = est => ({ periodo: P, hoja: h, fila: est.fila + 1, dni: est.dni, estudiante: est.nombre });

    t.lista.forEach(est => {
      res.estudiantes++;
      t.estructura.materias.forEach(m => { proy[h + '|' + est.dni + '|' + m.clave] = celdaPrep_(t, est, m, 0); });
      const enT = buscarPrep_(T, est.dni);
      if (enT.length > 1) {
        aviso(Object.assign(base(est), { tipo: 'estudianteBloqueado', detalle: 'El DNI figura ' + enT.length + ' veces en la grilla (' + enT.map(u => u.hoja + ' fila ' + (u.est.fila + 1)).join(', ') +
          '): ni la carga ni el generador pueden usarlo.' }));
        return;
      }
      if (est.bloque) { aviso(Object.assign(base(est), { tipo: 'estudianteBloqueado', detalle: est.bloque })); return; }
      if (!c || !ev) return;
      const enC = unicoEnHojaPrep_(C, h, est.dni);
      const enE = buscarPrep_(E, est.dni);
      const fallaEst = !enE.length ? 'sin estudiante' : enE.length > 1 ? 'ambiguo' : enE[0].hoja !== h ? 'otra hoja' : '';
      t.estructura.materias.forEach(m => {
        const mc = enC && c.materias[m.clave];
        const me = ev.materias[m.clave];
        const falla = fallaEst || (!me ? 'sin materia' : me.nivel !== m.nivel ? 'otro nivel' : '');
        const D = celdaPrep_(t, est, m, 0);
        const C0 = mc ? celdaPrep_(c, enC, mc, 0) : null;
        const E0 = falla ? null : celdaPrep_(ev, enE[0].est, me, 0);
        if (E0 !== null && situacionCalificacion_(P_.EVIDENCIA, E0) === 'valorada') calificados3b[h + '|' + m.clave] = (calificados3b[h + '|' + m.clave] || 0) + 1;
        if (C0 !== null) { res.comparadas++; if (norm_(D) === norm_(C0)) res.iguales++; }
        const lugar = Object.assign(base(est), { nivel: m.nivel, materia: m.nombre });
        const cambios = [], lineas = [];
        const cambio = (desp, nuevo, motivo) => cambios.push({
          periodo: P, fuente: T.fuente, hoja: h, fila: est.fila + desp, columna: m.columnaValor, a1: a1_(est.fila + desp, m.columnaValor), desp, clave: m.clave,
          dni: est.dni, estudiante: est.nombre, nivel: m.nivel, materia: m.nombre, campo: CARGA.CAMPOS[desp].nombre, anterior: celdaPrep_(t, est, m, desp), nuevo, motivo
        });
        const clave = [P, T.fuente, h, est.dni, m.clave].join('|');
        const dc = decidirCalificacionPrep_(P, D, C0, E0, falla, !!preparadas[clave]);
        if (dc.cursa) cambio(0, CARGA.PENDIENTE, dc.motivo);
        if (dc.tipo) lineas.push({ tipo: dc.tipo, detalle: dc.detalle });
        if (dc.sinPreparar) {
          const k = fallaEst ? est.dni : h + '|' + m.clave;
          (fallaEst ? sinPrepEst : sinPrepMat)[k] = ((fallaEst ? sinPrepEst : sinPrepMat)[k] || 0) + 1;
        }
        [1, 2, 3].forEach(desp => {
          const Df = celdaPrep_(t, est, m, desp);
          if (!Df) return;
          const Cf = mc ? celdaPrep_(c, enC, mc, desp) : null;
          if (Cf !== null && norm_(Df) === norm_(Cf)) cambio(desp, '', 'Copiado de ' + P_.COPIA + '.');
          else lineas.push({ tipo: 'complementario', detalle: CARGA.CAMPOS[desp].nombre + ' «' + recortePrep_(Df) + '» en el borrador, que no viene de ' + P_.COPIA + '. Se conservó.' });
        });
        if (tocadas[clave]) {
          if (cambios.length || lineas.length) aviso(Object.assign(lugar, { tipo: 'registro', detalle: 'La carga docente ya guardó cambios de este estudiante en esta materia (Registro): la preparación no la tocó.' }));
          return;
        }
        lineas.forEach(l => aviso(Object.assign({}, lugar, l)));
        cambios.forEach(x => {
          plan.cambios.push(x);
          if (x.desp > 0) { res.borrados++; return; }
          proy[h + '|' + est.dni + '|' + m.clave] = CARGA.PENDIENTE;
          res.aCursa++;
          const sd = situacionCalificacion_(P, D);
          if (sd === 'valorada') res.aCursaNotas++;
          else if (C0 !== null && situacionCalificacion_(P_.COPIA, C0) !== 'no cursa') res.aCursaVaciadas++;
        });
        if (dc.sinPreparar) res.sinPreparar++;
        const s = situacionCalificacion_(P, D);
        if (s === 'pendiente') res.yaCursa++;
        if (s === 'administrativo') res.administrativos++;
      });
    });

    t.sinDni.forEach(x => aviso({ tipo: 'sinDni', periodo: P, hoja: h, fila: x.fila + 1, estudiante: x.nombre, detalle: 'Fila con nombre y sin DNI: ni la carga ni el generador la ven.' }));
    if (!ev) return;
    // Materias del borrador contra las de 3º bimestre (misma hoja): presencia, nivel y cursantes.
    t.estructura.materias.forEach(m => {
      const me = ev.materias[m.clave];
      const n = sinPrepMat[h + '|' + m.clave] || 0;
      const cursantes = t.lista.filter(est => situacionCalificacion_(P, proy[h + '|' + est.dni + '|' + m.clave]) !== 'no cursa').length;
      const lugar = { periodo: P, hoja: h, nivel: m.nivel, materia: m.nombre };
      const quedan = n ? ' ' + quedanPrep_(n) + '.' : '';
      if (!me) {
        if (n) aviso(Object.assign(lugar, { tipo: 'materiaSinEvidencia', detalle: 'No figura en la hoja ' + h + ' de ' + P_.EVIDENCIA + ': no hay evidencia de cursada.' + quedan }));
        else if (cursantes) aviso(Object.assign(lugar, { tipo: 'materia', detalle: 'No figura en la hoja ' + h + ' de ' + P_.EVIDENCIA + '; se conservó lo que tiene el borrador (' + plural_(cursantes, 'cursante', 'cursantes') + ').' }));
      } else if (me.nivel !== m.nivel) {
        aviso(Object.assign(lugar, { tipo: 'materiaSinEvidencia', detalle: 'Es de ' + m.nivel + ' en el borrador y de ' + me.nivel + ' en ' + P_.EVIDENCIA + ': no se usó esa evidencia.' + quedan }));
      } else if (!cursantes && calificados3b[h + '|' + m.clave]) {
        aviso(Object.assign(lugar, { tipo: 'sinCursantes', detalle: plural_(calificados3b[h + '|' + m.clave], 'estudiante calificado', 'estudiantes calificados') + ' en ' + P_.EVIDENCIA + ' y nadie la cursa en el borrador.' }));
      }
    });
    ev.estructura.materias.filter(me => !t.materias[me.clave]).forEach(me => {
      const n = ev.lista.filter(est => situacionCalificacion_(P_.EVIDENCIA, celdaPrep_(ev, est, me, 0)) === 'valorada').length;
      aviso({ tipo: 'materia', periodo: P, hoja: h, nivel: me.nivel, materia: me.nombre, detalle: 'Figura en ' + P_.EVIDENCIA + ' (' + plural_(n, 'estudiante calificado', 'estudiantes calificados') + ') y no en el borrador.' });
    });
    // Campos generales y operativos iguales a 1º cuatrimestre: se informan, nunca se tocan.
    if (!c) return;
    const listo = [], exportado = [], generales = [];
    const deCola = (a, est, enc) => a.cola[enc] === undefined ? '' : txt_((a.valores[est.fila] || [])[a.cola[enc]]);
    t.lista.forEach(est => {
      const ec = unicoEnHojaPrep_(C, h, est.dni);
      if (!ec) return;
      if (norm_(deCola(t, est, 'LISTO PARA EXPORTAR')) === 'SI' && norm_(deCola(c, ec, 'LISTO PARA EXPORTAR')) === 'SI') listo.push(est.fila + 1);
      if (norm_(deCola(t, est, 'EXPORTADO')) === 'SI' && norm_(deCola(c, ec, 'EXPORTADO')) === 'SI') exportado.push(est.fila + 1);
      if (CARGA.ENCABEZADOS_COLA.slice(0, 4).some(enc => deCola(t, est, enc) && norm_(deCola(t, est, enc)) === norm_(deCola(c, ec, enc)))) generales.push(est.fila + 1);
    });
    const cola = (filas, texto) => { if (filas.length) aviso({ tipo: 'camposGenerales', periodo: P, hoja: h, detalle: plural_(filas.length, 'estudiante', 'estudiantes') + texto + ' (' + filasTextoPrep_(filas) + '). La preparación no toca esos campos.' }); };
    cola(listo, ' con LISTO PARA EXPORTAR = SÍ igual que en ' + P_.COPIA + ': el generador los tomaría como listos para ' + P);
    cola(exportado, ' con EXPORTADO = SÍ igual que en ' + P_.COPIA);
    cola(generales, ' con observaciones generales, de asistencia, de convivencia o PTF iguales a las de ' + P_.COPIA + ': aparecerían en el boletín de ' + P);
  });

  matriculaPrep_(P, g, sinPrepEst, res, aviso);
}

/** Matrícula e identidad del borrador contra 3º bimestre (matrícula vigente) y 1º cuatrimestre (origen de la copia). */
function matriculaPrep_(P, g, sinPrepEst, res, aviso) {
  const P_ = PREPARACION;
  const T = g[P], C = g[P_.COPIA], E = g[P_.EVIDENCIA];
  const todos = G => { const r = {}; if (G && !G.problema) CARGA.HOJAS_ORIGEN.forEach(h => { const a = G.hojas[h]; if (a) a.lista.forEach(est => { (r[est.dni] = r[est.dni] || []).push({ hoja: h, est }); }); }); return r; };
  const enT = todos(T), enE = todos(E), enC = todos(C);
  const porNombre = mapa => { const r = {}; Object.keys(mapa).forEach(d => { const k = clave_(mapa[d][0].est.nombre); if (k) (r[k] = r[k] || []).push(d); }); return r; };
  const nombresE = porNombre(enE), nombresT = porNombre(enT);
  const lugar = u => ({ periodo: P, hoja: u.hoja, fila: u.est.fila + 1, dni: u.est.dni, estudiante: u.est.nombre });
  if (!C.problema) {
    res.altas = Object.keys(enT).filter(d => !enC[d]).length;
    res.bajas = Object.keys(enC).filter(d => !enT[d]).length;
  }
  Object.keys(enT).forEach(dni => {
    const u = enT[dni][0];
    if (enT[dni].length > 1) return; // ya informado como bloqueo
    const n = sinPrepEst[dni] || 0;
    const tipo = n ? 'estudianteSinEvidencia' : 'matricula';
    const quedan = n ? ' ' + quedanPrep_(n) + '.' : '';
    if (!E.problema) {
      const e = enE[dni] || [];
      if (!e.length) {
        const otro = (nombresE[clave_(u.est.nombre)] || []).find(d => !enT[d]);
        aviso(Object.assign(lugar(u), { tipo, detalle: 'No figura en ' + P_.EVIDENCIA +
          (otro ? ': allí figura «' + enE[otro][0].est.nombre + '» con DNI ' + otro + ' (¿DNI mal escrito?).' : ' (la matrícula no debía cambiar en ese período).') + quedan }));
      } else if (e.length > 1) {
        aviso(Object.assign(lugar(u), { tipo, detalle: 'Figura ' + e.length + ' veces en ' + P_.EVIDENCIA + ': no se usó como evidencia.' + quedan }));
      } else if (e[0].hoja !== u.hoja) {
        aviso(Object.assign(lugar(u), { tipo, detalle: 'En ' + P_.EVIDENCIA + ' está en la hoja ' + e[0].hoja + ' (cambio de orientación): no se usó como evidencia.' + quedan }));
      } else {
        if (e[0].est.grupo !== u.est.grupo) aviso(Object.assign(lugar(u), { tipo: 'identidad', detalle: 'Grupo ' + (u.est.grupo || '(sin grupo)') + ' en el borrador y ' + (e[0].est.grupo || '(sin grupo)') + ' en ' + P_.EVIDENCIA + '.' }));
        if (u.est.orientacion && e[0].est.orientacion && norm_(u.est.orientacion) !== norm_(e[0].est.orientacion)) {
          aviso(Object.assign(lugar(u), { tipo: 'identidad', detalle: 'ORIENTACIÓN «' + u.est.orientacion + '» en el borrador y «' + e[0].est.orientacion + '» en ' + P_.EVIDENCIA + '.' }));
        }
      }
    }
    const nombres = [['el borrador', u.est.nombre]].concat((enE[dni] || []).map(x => [P_.EVIDENCIA, x.est.nombre]), (enC[dni] || []).map(x => [P_.COPIA, x.est.nombre]));
    if (nombres.some(x => clave_(x[1]) !== clave_(u.est.nombre))) {
      aviso(Object.assign(lugar(u), { tipo: 'identidad', detalle: 'El DNI tiene nombres distintos: ' + nombres.map(x => x[0] + ' «' + x[1] + '»').join('; ') + '.' }));
    }
  });
  if (!E.problema) {
    Object.keys(enE).filter(d => !enT[d]).forEach(dni => {
      const u = enE[dni][0];
      const otro = (nombresT[clave_(u.est.nombre)] || []).find(d => !enE[d]);
      aviso({ tipo: 'matricula', periodo: P, hoja: u.hoja, dni, estudiante: u.est.nombre, detalle: 'Figura en ' + P_.EVIDENCIA + ' (' + u.hoja + ', ' + (u.est.grupo || 'sin grupo') + ', fila ' + (u.est.fila + 1) +
        ') y no en el borrador' + (otro ? ', donde figura «' + enT[otro][0].est.nombre + '» con DNI ' + otro + ' (¿DNI mal escrito?).' : '. Si sigue cursando, agregarlo.') });
    });
  }
}

/** Los dos borradores entre sí (después de la preparación): estudiantes, grupos, materias, niveles y cursadas. */
function compararObjetivosPrep_(g, plan, aviso) {
  const [A, B] = PREPARACION.OBJETIVOS;
  const ambos = A + ' y ' + B;
  const todos = G => { const r = {}; CARGA.HOJAS_ORIGEN.forEach(h => { const a = G.hojas[h]; if (a) a.lista.forEach(est => { (r[est.dni] = r[est.dni] || []).push({ hoja: h, est }); }); }); return r; };
  const ea = todos(g[A]), eb = todos(g[B]);
  const lugar = u => ({ tipo: 'periodos', periodo: ambos, hoja: u.hoja, dni: u.est.dni, estudiante: u.est.nombre });
  Object.keys(ea).filter(d => !eb[d]).forEach(d => aviso(Object.assign(lugar(ea[d][0]), { detalle: 'Figura en ' + A + ' (fila ' + (ea[d][0].est.fila + 1) + ') y no en ' + B + '.' })));
  Object.keys(eb).filter(d => !ea[d]).forEach(d => aviso(Object.assign(lugar(eb[d][0]), { detalle: 'Figura en ' + B + ' (fila ' + (eb[d][0].est.fila + 1) + ') y no en ' + A + '.' })));
  Object.keys(ea).filter(d => eb[d] && ea[d].length === 1 && eb[d].length === 1).forEach(d => {
    const a = ea[d][0], b = eb[d][0];
    if (a.hoja !== b.hoja) return aviso(Object.assign(lugar(a), { detalle: 'Está en la hoja ' + a.hoja + ' en ' + A + ' y en ' + b.hoja + ' en ' + B + '.' }));
    if (a.est.grupo !== b.est.grupo) aviso(Object.assign(lugar(a), { detalle: 'Grupo ' + a.est.grupo + ' en ' + A + ' y ' + b.est.grupo + ' en ' + B + '.' }));
    const ta = g[A].hojas[a.hoja], tb = g[B].hojas[a.hoja];
    ta.estructura.materias.forEach(m => {
      const mb = tb.materias[m.clave];
      if (!mb || mb.nivel !== m.nivel) return;
      const va = plan.proyectado[A][a.hoja + '|' + d + '|' + m.clave], vb = plan.proyectado[B][a.hoja + '|' + d + '|' + m.clave];
      const ca = situacionCalificacion_(A, va) !== 'no cursa', cb = situacionCalificacion_(B, vb) !== 'no cursa';
      if (ca !== cb) {
        aviso(Object.assign(lugar(a), { fila: a.est.fila + 1, nivel: m.nivel, materia: m.nombre,
          detalle: 'Cursa en ' + (ca ? A : B) + ' («' + (ca ? va : vb) + '») y no en ' + (ca ? B : A) + ' («' + ((ca ? vb : va) || 'vacía') + '»).' }));
      }
    });
  });
  CARGA.HOJAS_ORIGEN.forEach(h => {
    const ta = g[A].hojas[h], tb = g[B].hojas[h];
    if (!ta || !tb) return;
    ta.estructura.materias.forEach(m => {
      const mb = tb.materias[m.clave];
      if (!mb) aviso({ tipo: 'materia', periodo: ambos, hoja: h, nivel: m.nivel, materia: m.nombre, detalle: 'Figura en ' + A + ' y no en ' + B + '.' });
      else if (mb.nivel !== m.nivel) aviso({ tipo: 'nivelEntrePeriodos', periodo: ambos, hoja: h, nivel: m.nivel, materia: m.nombre, detalle: 'Es de ' + m.nivel + ' en ' + A + ' y de ' + mb.nivel + ' en ' + B + ': la carga rechazaría uno de los dos períodos.' });
    });
    tb.estructura.materias.filter(m => !ta.materias[m.clave]).forEach(m => aviso({ tipo: 'materia', periodo: ambos, hoja: h, nivel: m.nivel, materia: m.nombre, detalle: 'Figura en ' + B + ' y no en ' + A + '.' }));
  });
}

// ================================================================ asignaciones

/**
 * Nómina de docentes (hoja Docentes): Docente en A y Correo en B, con encabezado en la fila 1. Devuelve
 * { problema, docentes: { clave del nombre: { docente, correo, fila, repetido } }, cantidad, avisos }.
 */
function leerNominaPrep_(matriz) {
  const P_ = PREPARACION;
  const n = { problema: '', docentes: {}, cantidad: 0, avisos: [] };
  if (!matriz || !matriz.some(f => f.some(v => txt_(v)))) return n;
  if (norm_((matriz[0] || [])[0]) !== norm_(P_.NOMINA[0]) || norm_((matriz[0] || [])[1]) !== norm_(P_.NOMINA[1])) {
    n.problema = 'La hoja ' + P_.HOJA_DOCENTES + ' debe tener «' + P_.NOMINA[0] + '» en A1 y «' + P_.NOMINA[1] + '» en B1 (la lista y la fórmula del borrador usan esas columnas).';
    return n;
  }
  const correos = {};
  const problema = (fila, texto) => n.avisos.push({ tipo: 'nomina', fila, detalle: P_.HOJA_DOCENTES + ', fila ' + fila + ': ' + texto });
  for (let r = 1; r < matriz.length; r++) {
    const docente = txt_(matriz[r][0]), correo = correo_(matriz[r][1]);
    if (!docente && !correo) continue;
    if (!docente) { problema(r + 1, 'el correo ' + correo + ' no tiene nombre.'); continue; }
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) problema(r + 1, '«' + docente + '» no tiene un correo válido.');
    const k = clave_(docente);
    if (n.docentes[k]) { n.docentes[k].repetido = true; problema(r + 1, '«' + docente + '» repite el nombre de la fila ' + n.docentes[k].fila + ': en la lista no se distinguirían.'); continue; }
    if (correo && correos[correo]) problema(r + 1, 'el correo ' + correo + ' ya es de «' + correos[correo] + '».');
    n.docentes[k] = { docente, correo, fila: r + 1 };
    if (correo) correos[correo] = correos[correo] || docente;
    n.cantidad++;
  }
  return n;
}

function leerBorradorPrep_(matriz) {
  const estandar = CARGA.ENCABEZADOS.asignaciones;
  const todas = estandar.concat(PREPARACION.EXTRA_BORRADOR);
  const vacia = !matriz || !matriz.some(f => f.some(v => txt_(v)));
  const b = { existe: !!matriz, vacia, encabezado: vacia ? todas.slice() : matriz[0].map(txt_), col: {}, filas: [], problema: '' };
  if (vacia) { todas.forEach((n, i) => { b.col[n] = i; }); return b; }
  todas.forEach(nombre => {
    const pos = b.encabezado.map((e, i) => (norm_(e) === norm_(nombre) ? i : -1)).filter(i => i >= 0);
    if (pos.length > 1) b.problema = b.problema || 'La hoja ' + PREPARACION.HOJA_BORRADOR + ' tiene más de una columna "' + nombre + '".';
    else if (pos.length) b.col[nombre] = pos[0];
    else if (estandar.includes(nombre)) b.problema = b.problema || 'La hoja ' + PREPARACION.HOJA_BORRADOR + ' no tiene la columna "' + nombre + '" en la fila 1.';
  });
  PREPARACION.EXTRA_BORRADOR.forEach(n => { if (b.col[n] === undefined) { b.col[n] = b.encabezado.length; b.encabezado.push(n); } });
  if (b.problema) return b;
  for (let r = 1; r < matriz.length; r++) {
    const f = matriz[r];
    const v = n => txt_(f[b.col[n]]);
    if (!estandar.some(n => v(n))) continue;
    b.filas.push({ fila: r + 1, orientacion: normalizarHojaOrigen_(v('Orientación')), nivel: normalizarNivel_(v('Nivel')), materia: v('Materia'), clave: clave_(v('Materia')),
      docente: v('Docente'), correo: correo_(v('Correo')) });
  }
  return b;
}

/**
 * Plan del borrador de Asignaciones: una fila por orientación y materia con cursantes después de la preparación
 * (las que faltan se agregan, con el Docente vacío para elegirlo de la nómina y el Correo por fórmula), y las columnas
 * propias (Cursantes y Revisión) de todas las filas. entrada: { g, proyectado, borrador, nomina }. Devuelve
 * { problema, crear, encabezado, col, actualizar, nuevas, conFormula, avisos, resumen }.
 */
function planAsignacionesPrep_(entrada) {
  const P_ = PREPARACION;
  const g = entrada.g;
  const b = leerBorradorPrep_(entrada.borrador);
  const nomina = leerNominaPrep_(entrada.nomina);
  const plan = { problema: b.problema || nomina.problema, crear: b.vacia, encabezado: b.encabezado, col: b.col, actualizar: [], nuevas: [], conFormula: [], avisos: [],
    resumen: { materias: 0, existentes: b.filas.length, nuevas: 0, conDocente: 0, sinDocente: 0, aCorregir: 0, nomina: nomina.cantidad } };
  if (plan.problema) { plan.avisos.push({ tipo: 'hojaAdministracion', detalle: plan.problema }); return plan; }
  nomina.avisos.forEach(a => plan.avisos.push(a));

  // Materias de los dos borradores, con nivel y cursantes (después de la preparación) por período.
  const materias = [], porClave = {};
  P_.OBJETIVOS.forEach(P => {
    const T = g[P];
    if (!T || T.problema) return;
    CARGA.HOJAS_ORIGEN.forEach(h => {
      const t = T.hojas[h];
      if (!t) return;
      t.estructura.materias.forEach(m => {
        const k = h + '|' + m.clave;
        let x = porClave[k];
        if (!x) { x = porClave[k] = { orientacion: h, clave: m.clave, nombre: m.nombre, niveles: {}, cursantes: {}, administrativos: {} }; materias.push(x); }
        x.niveles[P] = m.nivel;
        const s = t.lista.map(est => situacionCalificacion_(P, (entrada.proyectado[P] || {})[h + '|' + est.dni + '|' + m.clave])).filter(v => v !== 'no cursa');
        x.cursantes[P] = s.length;
        x.administrativos[P] = s.filter(v => v === 'administrativo').length;
      });
    });
  });
  const conCursantes = x => P_.OBJETIVOS.some(P => x.cursantes[P] > 0);
  const soloAdministrativos = x => P_.OBJETIVOS.every(P => !x.cursantes[P] || x.cursantes[P] === x.administrativos[P]);
  plan.resumen.materias = materias.filter(conCursantes).length;
  const cursantesTexto = x => P_.OBJETIVOS.filter(P => x.cursantes[P] !== undefined).map(P => P + ': ' + x.cursantes[P]).join(' · ');

  // Notas de una fila. Docente vacío es lo esperable al empezar: se cuenta aparte (sinDocente), no como problema.
  const notas = (x, fila) => {
    const n = [];
    const niveles = P_.OBJETIVOS.map(P => x.niveles[P]).filter(Boolean).filter((v, i, l) => l.indexOf(v) === i);
    if (niveles.length > 1) n.push('El nivel es distinto en ' + P_.OBJETIVOS.map(P => P + ' (' + x.niveles[P] + ')').join(' y ') + ': corregir la grilla.');
    else if (fila.nivel !== niveles[0]) n.push('El nivel de la materia en las grillas es ' + niveles[0] + '.');
    if (!conCursantes(x)) n.push('Sin cursantes en ' + P_.OBJETIVOS.join(' ni en ') + '.');
    if (fila.docente) {
      const d = nomina.docentes[clave_(fila.docente)];
      if (!d) n.push('«' + fila.docente + '» no está en la hoja ' + P_.HOJA_DOCENTES + '.');
      else if (d.repetido) n.push('«' + fila.docente + '» está repetido en la hoja ' + P_.HOJA_DOCENTES + '.');
      else if (fila.correo !== d.correo) n.push(fila.correo ? 'El correo no es el de ' + P_.HOJA_DOCENTES + ' (' + (d.correo || 'vacío') + ').' : 'Falta el correo: la columna Correo perdió su fórmula.');
    } else if (fila.correo) n.push('Hay correo y falta el docente: elegirlo de la lista.');
    return n;
  };
  const registrar = (x, fila, n, donde) => {
    if (n.length) {
      plan.resumen.aCorregir++;
      plan.avisos.push({ tipo: 'asignacion', hoja: x ? x.orientacion : fila.orientacion, nivel: fila.nivel, materia: x ? x.nombre : fila.materia, fila: fila.fila || '', detalle: P_.HOJA_BORRADOR + ', ' + donde + ': ' + n.join(' ') });
    }
    if (x && conCursantes(x)) { if (fila.docente && !n.length) plan.resumen.conDocente++; else if (!fila.docente) plan.resumen.sinDocente++; }
    const revision = (fila.docente ? [] : ['Falta elegir el docente.']).concat(n);
    if (x && soloAdministrativos(x) && conCursantes(x)) revision.push('Sólo ACREDITÓ o EQUIVALENCIA: el docente sólo completaría Calidad, Participación u Observaciones.');
    return revision.join(' ');
  };

  // Filas existentes: nunca se cambian sus cinco columnas; sólo Cursantes y Revisión. Una fila sin docente ni correo
  // recibe la fórmula del correo (no hay nada que perder).
  const vistas = {};
  b.filas.forEach(f => { if (f.orientacion && f.clave) (vistas[f.orientacion + '|' + f.clave] = vistas[f.orientacion + '|' + f.clave] || []).push(f); });
  b.filas.forEach(f => {
    const k = f.orientacion + '|' + f.clave;
    const x = porClave[k];
    let n;
    if (!f.orientacion || !f.clave) n = ['No se reconoce la orientación (EPSI o EPSG) o la materia.'];
    else if (!x) n = ['La materia no figura en las grillas de ' + P_.OBJETIVOS.join(' ni de ') + ' (' + f.orientacion + ').'];
    else {
      n = notas(x, f);
      if (vistas[k].length > 1) n.unshift('Hay ' + vistas[k].length + ' filas para esta materia (filas ' + vistas[k].map(v => v.fila).join(', ') + '): dejar una.');
    }
    if (!f.docente && !f.correo) plan.conFormula.push(f.fila);
    plan.actualizar.push({ fila: f.fila, 'Cursantes': x ? cursantesTexto(x) : '', 'Revisión': registrar(x, f, n, 'fila ' + f.fila) });
  });

  // Filas nuevas: una por materia con cursantes sin fila en el borrador, en el orden de las grillas.
  materias.filter(x => conCursantes(x) && !vistas[x.orientacion + '|' + x.clave]).forEach(x => {
    const nivel = x.niveles[P_.OBJETIVOS[0]] || x.niveles[P_.OBJETIVOS[1]];
    const fila = { nivel, docente: '', correo: '' };
    plan.nuevas.push({ 'Orientación': x.orientacion, 'Nivel': nivel, 'Materia': x.nombre, 'Docente': '', 'Correo': '', 'Cursantes': cursantesTexto(x), 'Revisión': registrar(x, fila, notas(x, fila), 'fila nueva') });
  });
  plan.resumen.nuevas = plan.nuevas.length;
  if (plan.resumen.sinDocente) {
    plan.avisos.push({ tipo: 'sinDocente', detalle: plural_(plan.resumen.sinDocente, 'materia con cursantes no tiene', 'materias con cursantes no tienen') + ' docente elegido en ' + P_.HOJA_BORRADOR +
      ' (filtrar la columna Docente vacía).' + (nomina.cantidad ? '' : ' La hoja ' + P_.HOJA_DOCENTES + ' está vacía: cargar primero la nómina (Docente y Correo).') });
  }
  return plan;
}

// ================================================================ aplicación (comprobaciones puras)

/**
 * Antes de escribir una hoja, con su lectura recién hecha dentro del lock: sólo se aplica un cambio si la hoja tiene
 * la misma estructura, el estudiante sigue en su fila con el bloque entero y la celda todavía muestra lo que se leyó.
 */
function confirmarCambiosPrep_(valores, periodo, cambios) {
  let estructura;
  try { estructura = estructuraGrilla_(valores, periodo); } catch (e) {
    return { aplicables: [], descartados: cambios.map(c => ({ cambio: c, detalle: 'La hoja cambió de forma desde la lectura: ' + sinAvisoDocente_(e.message) })) };
  }
  const aplicables = [], descartados = [];
  cambios.forEach(c => {
    const base = c.fila - c.desp;
    const m = estructura.materias.find(x => x.columnaValor === c.columna);
    let detalle = '';
    if (!m || m.clave !== c.clave) detalle = 'La columna de la materia se movió desde la lectura.';
    else if (dni_((valores[base] || [])[2]) !== c.dni) detalle = 'La fila del estudiante se movió desde la lectura.';
    else { try { validarBloque_(valores, estructura, base); } catch (e) { detalle = 'El bloque del estudiante cambió desde la lectura.'; } }
    if (!detalle) {
      const actual = txt_((valores[c.fila] || [])[c.columna]);
      if (actual !== c.anterior) detalle = 'La celda cambió desde la lectura: ahora dice «' + recortePrep_(actual) + '».';
    }
    if (detalle) descartados.push({ cambio: c, detalle }); else aplicables.push(c);
  });
  return { aplicables, descartados };
}

/** Después de escribir: 'aplicado' si la celda muestra el valor nuevo en su lugar ('inferido' si Sheets informó un error), si no 'revisar'. */
function verificarCambiosPrep_(valores, cambios, error) {
  return cambios.map(c => {
    const actual = txt_((valores[c.fila] || [])[c.columna]);
    const enSuLugar = dni_((valores[c.fila - c.desp] || [])[2]) === c.dni;
    if (enSuLugar && actual === c.nuevo) return { cambio: c, estado: error ? 'inferido' : 'aplicado', detalle: error ? 'Sheets informó un error (' + error + '), pero la celda muestra el valor nuevo.' : '' };
    return { cambio: c, estado: 'revisar', detalle: (error ? 'Error al escribir (' + error + '). ' : '') + (enSuLugar ? 'La celda muestra «' + recortePrep_(actual) + '».' : 'La fila del estudiante se movió: revisar la celda ' + c.a1 + '.') };
  });
}

function filaCambioPrep_(fecha, estado, detalle, c) {
  return [fecha, estado, detalle, c.periodo, c.fuente, c.hoja, c.a1, c.dni, c.estudiante, c.nivel, c.materia, c.campo, c.anterior, c.nuevo, c.motivo];
}

// ================================================================ informe

/**
 * Informe para la hoja Preparación y el registro de ejecución: resumen de lo hecho, una tabla con un renglón por tipo de
 * caso (cantidad, qué hacer y ejemplos) y, debajo, todos los casos. aplicacion: null (vista previa o con bloqueos) o
 * { error, aplicados, inferidos, revisar, descartados, filasCambios, borrador }.
 */
function informePreparacion_(plan, aplicacion, aplicar) {
  const P_ = PREPARACION;
  const hecho = aplicar && aplicacion && !aplicacion.error;
  const L = [];
  let modo;
  if (!aplicar) modo = 'VISTA PREVIA: no se modificó ninguna grilla ni el borrador de Asignaciones.';
  else if (!plan.puedeAplicar) modo = 'NO SE APLICÓ: hay bloqueos generales (ver abajo). No se modificó ninguna grilla ni el borrador de Asignaciones.';
  else if (aplicacion && aplicacion.error) modo = 'NO SE TERMINÓ DE APLICAR: ' + aplicacion.error;
  else modo = 'APLICADA.';
  L.push('Preparación de ' + listaY_(P_.OBJETIVOS) + ' · ' + fechaHoraTexto_(plan.ahora) + ' · ' + modo);
  const periodos = [P_.COPIA, P_.EVIDENCIA].concat(P_.OBJETIVOS);
  L.push('Grillas (Control del generador): ' + periodos.map(p => p + (p === P_.COPIA ? ' (origen de la copia)' : p === P_.EVIDENCIA ? ' (evidencia de cursada)' : '') + ' ' + (plan.fuentes[p] || '(falta)')).join(' · '));
  P_.OBJETIVOS.forEach(P => {
    const r = plan.porPeriodo[P];
    if (!r) { L.push(P + ': no se preparó (ver bloqueos).'); return; }
    L.push(P + ': ' + plural_(r.estudiantes, 'estudiante', 'estudiantes') + ' (' + CARGA.HOJAS_ORIGEN.map(h => h + ' ' + (r.hojas[h] || 0)).join(', ') + '). ' +
      (hecho ? 'Se pusieron ' : 'Se pondrían ') + r.aCursa + ' CURSA con evidencia de ' + P_.EVIDENCIA + ' (' + r.aCursaNotas + ' en lugar de notas copiadas de ' + P_.COPIA + ', ' +
      r.aCursaVaciadas + ' en celdas que se habían vaciado al limpiar la copia, ' + (r.aCursa - r.aCursaNotas - r.aCursaVaciadas) + ' donde no la cursaba en ' + P_.COPIA + ' o no figuraba) y se ' +
      (hecho ? 'borraron ' : 'borrarían ') + r.borrados + ' Calidad, Participación u Observaciones copiadas. Se conservan ' + r.yaCursa + ' CURSA ya puestos y ' + r.administrativos +
      ' ACREDITÓ o EQUIVALENCIA. Notas copiadas sin evidencia comparable: ' + r.sinPreparar + '.');
    L.push('   Antes de preparar, ' + r.iguales + ' de ' + r.comparadas + ' calificaciones comparables eran iguales a las de ' + P_.COPIA + '. ' +
      'Respecto de ' + P_.COPIA + ': ' + plural_(r.altas, 'estudiante nuevo', 'estudiantes nuevos') + ' y ' + plural_(r.bajas, 'que ya no está', 'que ya no están') + '.');
  });
  const a = plan.asignaciones;
  if (a && !a.problema) {
    L.push(P_.HOJA_BORRADOR + ': ' + plural_(a.resumen.materias, 'materia', 'materias') + ' con cursantes; ' + a.resumen.conDocente + ' con docente, ' + a.resumen.sinDocente + ' sin docente, ' +
      a.resumen.aCorregir + ' filas a corregir; ' + plural_(a.resumen.nuevas, 'fila nueva', 'filas nuevas') + '. Nómina (hoja ' + P_.HOJA_DOCENTES + '): ' +
      plural_(a.resumen.nomina, 'docente', 'docentes') + '. La hoja Asignaciones no se toca: se copia del borrador al abrir (manual, paso E1).');
  }
  if (hecho) {
    L.push('Escritura: ' + aplicacion.aplicados + ' celdas aplicadas y verificadas' + (aplicacion.inferidos ? ', ' + aplicacion.inferidos + ' inferidas' : '') + ', ' + aplicacion.revisar + ' a revisar, ' +
      aplicacion.descartados + ' no aplicadas porque cambiaron desde la lectura. Detalle por celda: hoja «' + P_.HOJA_CAMBIOS + '»' + (aplicacion.filasCambios ? ', filas ' + aplicacion.filasCambios : '') + '.');
    if (aplicacion.borrador) L.push(P_.HOJA_BORRADOR + ': ' + aplicacion.borrador);
  }
  L.push('A resolver: ' + P_.PRIORIDADES.map(p => p + ' ' + plan.conteo[p]).join(' · ') +
    '. Bloqueo: impide preparar. Decidir: lo resuelve una persona. Revisar: se conservó el borrador, pero difiere de la evidencia o de lo esperado.');

  // Un renglón por tipo, en orden de prioridad: lo rutinario y repetido ocupa una línea; lo excepcional, también.
  const tipos = [];
  plan.avisos.forEach(x => {
    let t = tipos.find(y => y.tipo === x.tipo);
    if (!t) { t = { tipo: x.tipo, prioridad: x.prioridad, casos: [] }; tipos.push(t); }
    t.casos.push(x);
  });
  const ejemplo = x => [x.periodo, x.hoja, x.materia, x.estudiante].filter(Boolean).join(' · ') || recortePrep_(x.detalle);
  const filasTipos = tipos.map(t => [t.prioridad, P_.TIPOS[t.tipo][1], t.casos.length, P_.TIPOS[t.tipo][2],
    t.casos.slice(0, P_.MAX_EJEMPLOS).map(ejemplo).join(' | ') + (t.casos.length > P_.MAX_EJEMPLOS ? ' | …' : '')]);

  const ancho = P_.INFORME.length;
  const pad = f => f.concat(Array(Math.max(0, ancho - f.length)).fill(''));
  const filas = L.map(l => pad([l])).concat([pad([]), pad(['Casos por tipo'])], [pad(P_.RESUMEN_TIPOS.slice())], filasTipos.map(pad),
    [pad([]), pad(['Todos los casos (usar una vista de filtro por Prioridad o Tipo)'])]);
  const encabezado = filas.length;
  filas.push(P_.INFORME.slice());
  plan.avisos.forEach(x => filas.push([x.prioridad, P_.TIPOS[x.tipo][1], x.periodo, x.hoja, x.nivel, x.materia, x.fila === '' ? '' : x.fila, x.dni, x.estudiante, x.detalle]));
  const texto = L.concat(['', 'Casos por tipo:'], filasTipos.length ? filasTipos.map(f => '  ' + f[0] + ' · ' + f[1] + ': ' + f[2] + '. ' + f[3] + (f[4] ? ' Ej.: ' + f[4] : '')) : ['  (ninguno)'],
    ['', 'Detalle de cada caso en la hoja «' + P_.HOJA_INFORME + '» de la planilla de administración.']).join('\n');
  return { filas, texto, encabezado, encabezadoTipos: L.length + 2 };
}
