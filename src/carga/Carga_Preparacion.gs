/**
 * Carga docente EPS (#32) · preparación de las grillas de 2º cuatrimestre y anual y del borrador de Asignaciones.
 *
 * Los borradores de esas dos grillas se hicieron copiando la de 1º cuatrimestre y después se editaron a mano. Una
 * sola ejecución desde el editor (prepararPeriodosCarga, en Carga_Preparacion_Google.gs) los deja listos para la
 * carga docente y deja en la planilla de administración un informe de lo hecho y de lo que necesita una persona:
 *  - cursadas: CURSA donde 3º bimestre, el último período calificado y con la matrícula vigente, tiene una
 *    calificación ordinaria. ACREDITÓ, EQUIVALENCIA, CURSA sin calificar o vacío no habilitan CURSA. Anual se prepara
 *    con la misma evidencia: 2º cuatrimestre todavía no está calificado y no se usa;
 *  - restos de la copia: una calificación de 1º cuatrimestre que quedó en el borrador se reemplaza por CURSA sólo
 *    con esa evidencia (si no, se informa); Calidad, Participación y Observaciones iguales a las de 1º cuatrimestre
 *    se borran. Ninguna calificación anterior pasa a ser una valoración nueva;
 *  - asignaciones: una fila por orientación y materia con cursantes en «Asignaciones (borrador)», con orientación,
 *    nivel y nombre de la grilla. El correo se completa sólo si todas las protecciones por columna de las grillas
 *    (las de la carga directa anterior) que cubren esa materia nombran a la misma persona; el nombre, sólo si ese
 *    correo ya tiene uno escrito por el Equipo EPS;
 *  - revisión: matrícula, identidad, grupos, orientación, materias y niveles de cada borrador contra 3º bimestre y
 *    1º cuatrimestre, y de los dos borradores entre sí.
 *
 * Conservación (se puede repetir sin deshacer trabajo ni duplicar filas):
 *  - sólo cambia una celda que todavía tiene lo que vino de la copia (igual a 1º cuatrimestre; vacía y NO CURSA son
 *    equivalentes) o una calificación vacía o NO CURSA en una fila o columna que no venía de la copia. Todo lo demás se conserva y, si contradice la evidencia,
 *    se informa para que lo decida el Equipo EPS;
 *  - nunca inserta, borra ni mueve filas o columnas, ni toca ORIENTACIÓN y los campos posteriores;
 *  - no toca a un estudiante en una materia en la que la carga docente ya escribió (Registro);
 *  - con una ventana cargada para 2º cuatrimestre o anual, no modifica nada;
 *  - en el borrador de Asignaciones nunca cambia Orientación, Nivel, Materia, Docente ni Correo de una fila existente.
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
  // Grillas cuyas protecciones por columna pueden identificar al docente de cada materia.
  FUENTES_CORREO: ['3º bimestre', '2º cuatrimestre', 'Anual', '1º cuatrimestre'],
  HOJA_INFORME: 'Preparación',
  HOJA_CAMBIOS: 'Preparación · cambios',
  HOJA_BORRADOR: 'Asignaciones (borrador)',
  // Columnas que la preparación agrega al borrador. La carga las ignora (lee por encabezado).
  EXTRA_BORRADOR: ['Cursantes', 'Origen del correo', 'Revisión'],
  CAMBIOS: ['Fecha', 'Estado', 'Detalle', 'Período', 'Grilla', 'Hoja', 'Celda', 'DNI', 'Estudiante', 'Nivel', 'Materia', 'Campo', 'Anterior', 'Nuevo', 'Motivo'],
  INFORME: ['Prioridad', 'Período', 'Hoja', 'Nivel', 'Materia', 'Fila', 'DNI', 'Estudiante', 'Detalle'],
  // Bloqueo: impide preparar (todo o una hoja o un estudiante). Decidir: quedó sin preparar y necesita una persona.
  // Revisar: se conservó lo del borrador, pero difiere de la evidencia o de lo esperado.
  PRIORIDADES: ['Bloqueo', 'Decidir', 'Revisar'],
  MAX_FILAS_TEXTO: 15,
  MAX_TEXTO: 60,
  MAX_LINEAS_REGISTRO: 40,
  // Tiempo de la ejecución después del cual ya no se leen protecciones (el límite de Apps Script es de 6 minutos).
  MS_PROTECCIONES: 180000
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
 * la copia). E0: el de 3º bimestre, o null con «falla» si no hay evidencia comparable.
 * Devuelve { cursa, motivo } (cambiar a CURSA), { prioridad, detalle } (informar, sin cambiar), { sinPreparar }
 * (un resto de la copia sin evidencia; lo informa la línea del estudiante o de la materia) o {} (nada que hacer).
 */
function decidirCalificacionPrep_(P, D, C0, E0, falla) {
  const copia = PREPARACION.COPIA, ev = PREPARACION.EVIDENCIA;
  const d = situacionCalificacion_(P, D);
  const c = C0 === null ? 'ausente' : situacionCalificacion_(copia, C0);
  const residuo = c !== 'ausente' && (d === 'no cursa' ? c === 'no cursa' : norm_(D) === norm_(C0));
  if (d === 'a revisar') return { prioridad: 'Decidir', detalle: 'Calificación «' + recortePrep_(D) + '» no reconocida para ' + P + ': corregirla (CURSA, NO CURSA, ACREDITÓ o EQUIVALENCIA).' };
  if (d === 'valorada' && !residuo) {
    return { prioridad: 'Decidir', detalle: 'Calificación «' + D + '» en el borrador, que no viene de ' + copia + (c === 'ausente' ? '' : ' (allí: «' + (C0 || 'vacía') + '»)') +
      '. La carga la mostraría como ya calificada: si no corresponde, poner CURSA o NO CURSA.' };
  }
  if (falla) return { sinPreparar: d === 'valorada' };
  const e = situacionCalificacion_(ev, E0);
  const enEv = 'en ' + ev + ' ' + evidenciaTextoPrep_(e, E0);
  if (d === 'valorada') {
    if (e === 'valorada') return { cursa: true, motivo: 'Calificación «' + D + '» copiada de ' + copia + '; en ' + ev + ' tiene «' + E0 + '».' };
    return { prioridad: 'Decidir', detalle: 'Quedó la calificación «' + D + '» de ' + copia + ' y ' + enEv + ': decidir CURSA, NO CURSA, ACREDITÓ o EQUIVALENCIA.' };
  }
  if (d === 'pendiente') {
    if (e === 'administrativo' || e === 'no cursa') return { prioridad: 'Revisar', detalle: 'CURSA en el borrador' + (residuo ? ' (copiado de ' + copia + ')' : '') + ', pero ' + enEv + '.' };
    return {};
  }
  if (d === 'administrativo') {
    if (e === 'valorada') return { prioridad: 'Revisar', detalle: '«' + D + '» en el borrador, pero ' + enEv + '.' };
    return {};
  }
  // d === 'no cursa': vacía o NO CURSA, equivalentes (#13). Pasa a CURSA si no cursaba tampoco en la copia o si la
  // fila o la columna no venía de ella; si en 1º cuatrimestre cursaba, alguien la quitó a mano y se conserva.
  if (e === 'valorada') {
    if (residuo || c === 'ausente') {
      return { cursa: true, motivo: 'En ' + ev + ' tiene «' + E0 + '»' + (residuo ? '; en ' + copia + ' no la cursaba.' : '; la fila o la columna no venía de ' + copia + '.') };
    }
    return { prioridad: 'Revisar', detalle: (D || 'Vacía') + ' en el borrador (en ' + copia + ': «' + C0 + '»), pero ' + enEv + '. Se conservó.' };
  }
  if (e === 'administrativo') return { prioridad: 'Revisar', detalle: (D || 'Vacía') + ' en el borrador, pero ' + enEv + ': lo decide el Equipo EPS.' };
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
 * valores | null, EPSG }, problema } }, protecciones: { período: { EPSI: [{ filas: [r1, r2], columnas: [c1, c2],
 * editores }] | { problema } } }, borrador (matriz de Asignaciones (borrador) o null), registro (matriz o null),
 * usuario (la cuenta propietaria) }. Filas y columnas en base 0, inclusivas.
 */
function planificarPreparacion_(entrada) {
  const P_ = PREPARACION;
  const plan = { ahora: entrada.ahora, avisos: [], cambios: [], bloqueos: [], puedeAplicar: false, fuentes: {}, porPeriodo: {}, proyectado: {}, asignaciones: null };
  const aviso = x => plan.avisos.push(Object.assign({ prioridad: 'Revisar', periodo: '', hoja: '', nivel: '', materia: '', fila: '', dni: '', estudiante: '', detalle: '' }, x));
  const general = detalle => { plan.bloqueos.push(detalle); aviso({ prioridad: 'Bloqueo', detalle }); };

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
      aviso({ prioridad: 'Bloqueo', periodo: p, hoja: h, detalle: g[p].problemasHoja[h] + (P_.OBJETIVOS.includes(p) ? ' No se prepara esta hoja.' : ' No se prepara la hoja ' + h + ' de los borradores.') });
    });
  });

  // Celdas en las que la carga docente ya escribió: Registro con período, fuente, orientación, DNI y materia.
  const tocadas = {};
  if (entrada.registro && entrada.registro.length > 1) {
    const enc = entrada.registro[0].map(norm_);
    const col = n => enc.indexOf(norm_(n));
    entrada.registro.slice(1).forEach(f => {
      const periodo = normalizarPeriodo_(f[col('Período')]);
      if (!P_.OBJETIVOS.includes(periodo)) return;
      tocadas[[periodo, txt_(f[col('Fuente')]), normalizarHojaOrigen_(f[col('Orientación')]), dni_(f[col('DNI')]), clave_(f[col('Materia')])].join('|')] = true;
    });
  }

  P_.OBJETIVOS.forEach(P => { if (!g[P].problema) prepararObjetivoPrep_(P, g, tocadas, plan, aviso); });
  if (P_.OBJETIVOS.every(P => !g[P].problema)) compararObjetivosPrep_(g, plan, aviso);

  plan.asignaciones = planAsignacionesPrep_({ g, proyectado: plan.proyectado, borrador: entrada.borrador, asignaciones: entrada.datos.asignaciones, protecciones: entrada.protecciones, usuario: entrada.usuario });
  if (plan.asignaciones.problema) aviso({ prioridad: 'Bloqueo', detalle: plan.asignaciones.problema });
  plan.asignaciones.avisos.forEach(aviso);

  plan.g = g;
  plan.puedeAplicar = !plan.bloqueos.length;
  const orden = p => P_.PRIORIDADES.indexOf(p);
  plan.avisos.sort((a, b) => orden(a.prioridad) - orden(b.prioridad));
  plan.conteo = {};
  P_.PRIORIDADES.forEach(p => { plan.conteo[p] = plan.avisos.filter(a => a.prioridad === p).length; });
  return plan;
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
function prepararObjetivoPrep_(P, g, tocadas, plan, aviso) {
  const P_ = PREPARACION;
  const T = g[P], C = g[P_.COPIA], E = g[P_.EVIDENCIA];
  const res = { estudiantes: 0, aCursa: 0, aCursaSinCopia: 0, borrados: 0, yaCursa: 0, administrativos: 0, sinPreparar: 0, iguales: 0, comparadas: 0, altas: 0, bajas: 0, hojas: {} };
  plan.porPeriodo[P] = res;
  const proy = plan.proyectado[P] = {};
  const sinPrepEst = {}, sinPrepMat = {};

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
        aviso(Object.assign(base(est), { prioridad: 'Bloqueo', detalle: 'El DNI figura ' + enT.length + ' veces en la grilla (' + enT.map(u => u.hoja + ' fila ' + (u.est.fila + 1)).join(', ') +
          '): ni la carga ni el generador pueden usarlo. No se preparó.' }));
        return;
      }
      if (est.bloque) { aviso(Object.assign(base(est), { prioridad: 'Bloqueo', detalle: est.bloque + ' No se preparó este estudiante.' })); return; }
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
        if (C0 !== null) { res.comparadas++; if (norm_(D) === norm_(C0)) res.iguales++; }
        const lugar = Object.assign(base(est), { nivel: m.nivel, materia: m.nombre });
        const cambios = [], lineas = [];
        const cambio = (desp, nuevo, motivo) => cambios.push({
          periodo: P, fuente: T.fuente, hoja: h, fila: est.fila + desp, columna: m.columnaValor, a1: a1_(est.fila + desp, m.columnaValor), desp, clave: m.clave,
          dni: est.dni, estudiante: est.nombre, nivel: m.nivel, materia: m.nombre, campo: CARGA.CAMPOS[desp].nombre, anterior: celdaPrep_(t, est, m, desp), nuevo, motivo
        });
        const dc = decidirCalificacionPrep_(P, D, C0, E0, falla);
        if (dc.cursa) cambio(0, CARGA.PENDIENTE, dc.motivo);
        if (dc.detalle) lineas.push({ prioridad: dc.prioridad, detalle: dc.detalle });
        if (dc.sinPreparar) {
          const k = fallaEst ? est.dni : h + '|' + m.clave;
          (fallaEst ? sinPrepEst : sinPrepMat)[k] = ((fallaEst ? sinPrepEst : sinPrepMat)[k] || 0) + 1;
        }
        [1, 2, 3].forEach(desp => {
          const Df = celdaPrep_(t, est, m, desp);
          if (!Df) return;
          const Cf = mc ? celdaPrep_(c, enC, mc, desp) : null;
          if (Cf !== null && norm_(Df) === norm_(Cf)) cambio(desp, '', 'Copiado de ' + P_.COPIA + '.');
          else lineas.push({ prioridad: 'Revisar', detalle: CARGA.CAMPOS[desp].nombre + ' «' + recortePrep_(Df) + '» en el borrador, que no viene de ' + P_.COPIA + '. Se conservó: el docente lo verá como ya cargado.' });
        });
        if (tocadas[[P, T.fuente, h, est.dni, m.clave].join('|')]) {
          if (cambios.length || lineas.length) aviso(Object.assign(lugar, { prioridad: 'Revisar', detalle: 'La carga docente ya guardó cambios de este estudiante en esta materia (Registro): la preparación no la tocó.' }));
          return;
        }
        lineas.forEach(l => aviso(Object.assign({}, lugar, l)));
        cambios.forEach(x => {
          plan.cambios.push(x);
          if (x.desp === 0) { proy[h + '|' + est.dni + '|' + m.clave] = CARGA.PENDIENTE; res.aCursa++; if (C0 === null) res.aCursaSinCopia++; } else res.borrados++;
        });
        if (dc.sinPreparar) res.sinPreparar++;
        const s = situacionCalificacion_(P, D);
        if (s === 'pendiente') res.yaCursa++;
        if (s === 'administrativo') res.administrativos++;
      });
    });

    t.sinDni.forEach(x => aviso({ prioridad: 'Decidir', periodo: P, hoja: h, fila: x.fila + 1, estudiante: x.nombre, detalle: 'Fila con nombre y sin DNI: ni la carga ni el generador la ven. Completar el DNI o quitarla.' }));
    if (!ev) return;
    // Materias del borrador contra las de 3º bimestre (misma hoja): presencia y nivel.
    t.estructura.materias.forEach(m => {
      const me = ev.materias[m.clave];
      const n = sinPrepMat[h + '|' + m.clave] || 0;
      const cursantes = t.lista.filter(est => situacionCalificacion_(P, proy[h + '|' + est.dni + '|' + m.clave]) !== 'no cursa').length;
      const lugar = { periodo: P, hoja: h, nivel: m.nivel, materia: m.nombre };
      const quedan = n ? ' ' + quedanPrep_(n) + ': decidir si corresponde CURSA o NO CURSA.' : '';
      if (!me) {
        if (n || cursantes) aviso(Object.assign(lugar, { prioridad: n ? 'Decidir' : 'Revisar', detalle: 'No figura en la hoja ' + h + ' de ' + P_.EVIDENCIA + ': no hay evidencia de cursada.' + (quedan || ' Se conservó lo que tiene el borrador (' + plural_(cursantes, 'cursante', 'cursantes') + ').') }));
      } else if (me.nivel !== m.nivel) {
        aviso(Object.assign(lugar, { prioridad: 'Decidir', detalle: 'Es de ' + m.nivel + ' en el borrador y de ' + me.nivel + ' en ' + P_.EVIDENCIA + ': no se usó esa evidencia.' + quedan }));
      }
    });
    ev.estructura.materias.filter(me => !t.materias[me.clave]).forEach(me => {
      const n = ev.lista.filter(est => situacionCalificacion_(P_.EVIDENCIA, celdaPrep_(ev, est, me, 0)) === 'valorada').length;
      aviso({ periodo: P, hoja: h, nivel: me.nivel, materia: me.nombre, prioridad: 'Revisar', detalle: 'Figura en ' + P_.EVIDENCIA + ' (' + plural_(n, 'estudiante calificado', 'estudiantes calificados') + ') y no en el borrador.' });
    });
    // Campos generales y operativos copiados de 1º cuatrimestre: se informan, nunca se tocan.
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
    const cola = (filas, texto) => { if (filas.length) aviso({ periodo: P, hoja: h, prioridad: 'Revisar', detalle: plural_(filas.length, 'estudiante', 'estudiantes') + texto + ' (' + filasTextoPrep_(filas) + '). La preparación no toca esos campos: los administra el Equipo EPS.' }); };
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
    const quedan = n ? ' ' + quedanPrep_(n) + ': decidir CURSA o NO CURSA.' : '';
    if (!E.problema) {
      const e = enE[dni] || [];
      if (!e.length) {
        const otro = (nombresE[clave_(u.est.nombre)] || []).find(d => !enT[d]);
        aviso(Object.assign(lugar(u), { prioridad: n ? 'Decidir' : 'Revisar', detalle: 'No figura en ' + P_.EVIDENCIA +
          (otro ? ': allí figura «' + enE[otro][0].est.nombre + '» con DNI ' + otro + ' (¿DNI mal escrito?).' : ' (la matrícula no debía cambiar en ese período).') + quedan }));
      } else if (e.length > 1) {
        aviso(Object.assign(lugar(u), { prioridad: n ? 'Decidir' : 'Revisar', detalle: 'Figura ' + e.length + ' veces en ' + P_.EVIDENCIA + ': no se usó como evidencia.' + quedan }));
      } else if (e[0].hoja !== u.hoja) {
        aviso(Object.assign(lugar(u), { prioridad: n ? 'Decidir' : 'Revisar', detalle: 'En ' + P_.EVIDENCIA + ' está en la hoja ' + e[0].hoja + ' (cambio de orientación): no se usó como evidencia.' + quedan }));
      } else {
        if (e[0].est.grupo !== u.est.grupo) aviso(Object.assign(lugar(u), { detalle: 'Grupo ' + (u.est.grupo || '(sin grupo)') + ' en el borrador y ' + (e[0].est.grupo || '(sin grupo)') + ' en ' + P_.EVIDENCIA + '.' }));
        if (u.est.orientacion && e[0].est.orientacion && norm_(u.est.orientacion) !== norm_(e[0].est.orientacion)) {
          aviso(Object.assign(lugar(u), { detalle: 'ORIENTACIÓN «' + u.est.orientacion + '» en el borrador y «' + e[0].est.orientacion + '» en ' + P_.EVIDENCIA + '.' }));
        }
      }
    }
    const nombres = [['el borrador', u.est.nombre]].concat((enE[dni] || []).map(x => [P_.EVIDENCIA, x.est.nombre]), (enC[dni] || []).map(x => [P_.COPIA, x.est.nombre]));
    if (nombres.some(x => clave_(x[1]) !== clave_(u.est.nombre))) {
      aviso(Object.assign(lugar(u), { detalle: 'El DNI tiene nombres distintos: ' + nombres.map(x => x[0] + ' «' + x[1] + '»').join('; ') + '.' }));
    }
  });
  if (!E.problema) {
    Object.keys(enE).filter(d => !enT[d]).forEach(dni => {
      const u = enE[dni][0];
      const otro = (nombresT[clave_(u.est.nombre)] || []).find(d => !enE[d]);
      aviso({ periodo: P, hoja: u.hoja, dni, estudiante: u.est.nombre, prioridad: 'Revisar', detalle: 'Figura en ' + P_.EVIDENCIA + ' (' + u.hoja + ', ' + (u.est.grupo || 'sin grupo') + ', fila ' + (u.est.fila + 1) +
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
  const lugar = u => ({ periodo: ambos, hoja: u.hoja, dni: u.est.dni, estudiante: u.est.nombre });
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
      if (!mb) aviso({ periodo: ambos, hoja: h, nivel: m.nivel, materia: m.nombre, detalle: 'Figura en ' + A + ' y no en ' + B + '.' });
      else if (mb.nivel !== m.nivel) aviso({ periodo: ambos, hoja: h, nivel: m.nivel, materia: m.nombre, prioridad: 'Decidir', detalle: 'Es de ' + m.nivel + ' en ' + A + ' y de ' + mb.nivel + ' en ' + B + ': la asignación es una sola, y la carga rechazaría uno de los dos períodos.' });
    });
    tb.estructura.materias.filter(m => !ta.materias[m.clave]).forEach(m => aviso({ periodo: ambos, hoja: h, nivel: m.nivel, materia: m.nombre, detalle: 'Figura en ' + B + ' y no en ' + A + '.' }));
  });
}

// ================================================================ asignaciones

/**
 * Correos que las protecciones por columna de cada grilla asocian a cada materia: { 'EPSI|CLAVE': { período: [correos] } }.
 * Una protección cuenta si cubre la columna de valoración de la materia en alguna fila de estudiantes. Se descartan
 * la cuenta propietaria y quienes figuran en todas esas protecciones de la grilla (el Equipo EPS), si hay al menos dos.
 */
function candidatosCorreoPrep_(g, protecciones, usuario) {
  const r = {};
  PREPARACION.FUENTES_CORREO.forEach(S => {
    const G = g[S];
    if (!G || G.problema) return;
    CARGA.HOJAS_ORIGEN.forEach(h => {
      const a = G.hojas[h];
      const prots = ((protecciones || {})[S] || {})[h];
      if (!a || !Array.isArray(prots)) return;
      const cubre = (p, col) => p.columnas[0] <= col && col <= p.columnas[1] && p.filas[1] >= 2;
      const deMaterias = prots.filter(p => a.estructura.materias.some(m => cubre(p, m.columnaValor)));
      const generales = deMaterias.length >= 2 ? deMaterias.map(p => p.editores.map(correo_)).reduce((acc, l) => acc.filter(x => l.includes(x))) : [];
      const fuera = generales.concat([correo_(usuario)]);
      a.estructura.materias.forEach(m => {
        const cubren = deMaterias.filter(p => cubre(p, m.columnaValor));
        if (!cubren.length) return;
        const k = h + '|' + m.clave;
        const lista = [];
        cubren.forEach(p => p.editores.map(correo_).forEach(x => { if (x && !fuera.includes(x) && !lista.includes(x)) lista.push(x); }));
        (r[k] = r[k] || {})[S] = lista.sort();
      });
    });
  });
  return r;
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
 * Plan del borrador de Asignaciones: filas nuevas para las materias con cursantes que no tienen fila, y las columnas
 * propias (Cursantes y Revisión) de todas las filas existentes. entrada: { g, proyectado, borrador, asignaciones,
 * protecciones, usuario }. Devuelve { problema, crear, encabezado, col, actualizar, nuevas, avisos, resumen }.
 */
function planAsignacionesPrep_(entrada) {
  const P_ = PREPARACION;
  const g = entrada.g;
  const b = leerBorradorPrep_(entrada.borrador);
  const plan = { problema: b.problema, crear: b.vacia, encabezado: b.encabezado, col: b.col, actualizar: [], nuevas: [], avisos: [],
    resumen: { materias: 0, existentes: b.filas.length, nuevas: 0, conCorreo: 0, aCompletar: 0 } };
  if (b.problema) return plan;

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
        if (!x) { x = porClave[k] = { orientacion: h, clave: m.clave, nombre: m.nombre, niveles: {}, cursantes: {} }; materias.push(x); }
        x.niveles[P] = m.nivel;
        x.cursantes[P] = t.lista.filter(est => situacionCalificacion_(P, entrada.proyectado[P] && entrada.proyectado[P][h + '|' + est.dni + '|' + m.clave]) !== 'no cursa').length;
      });
    });
  });
  const conCursantes = x => P_.OBJETIVOS.some(P => x.cursantes[P] > 0);
  plan.resumen.materias = materias.filter(conCursantes).length;
  const cursantesTexto = x => P_.OBJETIVOS.filter(P => x.cursantes[P] !== undefined).map(P => P + ': ' + x.cursantes[P]).join(' · ');

  const candidatos = candidatosCorreoPrep_(g, entrada.protecciones, entrada.usuario);
  P_.FUENTES_CORREO.forEach(S => CARGA.HOJAS_ORIGEN.forEach(h => {
    const x = ((entrada.protecciones || {})[S] || {})[h];
    if (x && x.problema) plan.avisos.push({ prioridad: 'Revisar', periodo: S, hoja: h, detalle: 'No se pudieron leer las protecciones de esta hoja (' + x.problema + '): no se usaron para buscar correos.' });
  }));
  const propuesta = k => {
    const porFuente = candidatos[k] || {};
    const todos = [];
    Object.keys(porFuente).forEach(S => porFuente[S].forEach(x => { if (!todos.includes(x)) todos.push(x); }));
    const detalle = P_.FUENTES_CORREO.filter(S => porFuente[S] && porFuente[S].length).map(S => S + ': ' + porFuente[S].join(', ')).join('; ');
    if (todos.length === 1) return { correo: todos[0], origen: 'Protecciones de ' + listaY_(P_.FUENTES_CORREO.filter(S => porFuente[S] && porFuente[S].includes(todos[0]))), detalle };
    return { correo: '', detalle };
  };
  // Nombres que el Equipo EPS ya escribió para cada correo (borrador y Asignaciones). Un correo con dos nombres no se usa.
  const nombres = {};
  const anotar = (correo, docente) => { if (!correo || !docente) return; nombres[correo] = nombres[correo] === undefined || clave_(nombres[correo]) === clave_(docente) ? docente : null; };
  b.filas.forEach(f => anotar(f.correo, f.docente));
  try { tablaAdministracion_(entrada.asignaciones, 'asignaciones').forEach(a => anotar(correo_(a['Correo']), a['Docente'])); } catch (e) { /* sin Asignaciones legible: sólo el borrador */ }

  const notas = (x, fila, prop) => {
    const n = [];
    const niveles = P_.OBJETIVOS.map(P => x && x.niveles[P]).filter(Boolean).filter((v, i, l) => l.indexOf(v) === i);
    if (niveles.length > 1) n.push('El nivel es distinto en ' + P_.OBJETIVOS.map(P => P + ' (' + x.niveles[P] + ')').join(' y ') + ': corregir la grilla.');
    else if (fila.nivel !== niveles[0]) n.push('El nivel de la materia en las grillas es ' + niveles[0] + '.');
    if (!conCursantes(x)) n.push('Sin cursantes en ' + P_.OBJETIVOS.join(' ni en ') + '.');
    if (!fila.correo) {
      n.push('Falta el correo' + (prop.correo ? ': las protecciones de las grillas indican ' + prop.correo + '.' :
        prop.detalle ? ' (las protecciones de las grillas nombran a varias personas: ' + prop.detalle + ').' : '.'));
    }
    else if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(fila.correo)) n.push('El correo no es válido.');
    else if (prop.correo && prop.correo !== fila.correo) n.push('Las protecciones de las grillas indican ' + prop.correo + '.');
    if (!fila.docente) n.push('Falta el nombre del docente.');
    return n;
  };

  // Filas existentes: nunca se cambian sus cinco columnas; sólo Cursantes y Revisión.
  const vistas = {};
  b.filas.forEach(f => { if (f.orientacion && f.clave) (vistas[f.orientacion + '|' + f.clave] = vistas[f.orientacion + '|' + f.clave] || []).push(f); });
  b.filas.forEach(f => {
    const k = f.orientacion + '|' + f.clave;
    const x = porClave[k];
    let n;
    if (!f.orientacion || !f.clave) n = ['No se reconoce la orientación (EPSI o EPSG) o la materia.'];
    else if (!x) n = ['La materia no figura en las grillas de ' + P_.OBJETIVOS.join(' ni de ') + ' (' + f.orientacion + ').'];
    else {
      n = notas(x, f, propuesta(k));
      if (vistas[k].length > 1) n.unshift('Hay ' + vistas[k].length + ' filas para esta materia (filas ' + vistas[k].map(v => v.fila).join(', ') + '): dejar una.');
    }
    plan.actualizar.push({ fila: f.fila, 'Cursantes': x ? cursantesTexto(x) : '', 'Revisión': n.join(' ') });
    if (n.length) plan.avisos.push({ prioridad: 'Decidir', hoja: f.orientacion, nivel: f.nivel, materia: f.materia, detalle: P_.HOJA_BORRADOR + ', fila ' + f.fila + ': ' + n.join(' ') });
  });

  // Filas nuevas: una por materia con cursantes sin fila en el borrador, en el orden de las grillas.
  materias.filter(x => conCursantes(x) && !vistas[x.orientacion + '|' + x.clave]).forEach(x => {
    const prop = propuesta(x.orientacion + '|' + x.clave);
    const nivel = x.niveles[P_.OBJETIVOS[0]] || x.niveles[P_.OBJETIVOS[1]];
    const docente = prop.correo && nombres[prop.correo] ? nombres[prop.correo] : '';
    const n = notas(x, { nivel, correo: prop.correo, docente, nueva: true }, prop);
    plan.nuevas.push({ 'Orientación': x.orientacion, 'Nivel': nivel, 'Materia': x.nombre, 'Docente': docente, 'Correo': prop.correo, 'Cursantes': cursantesTexto(x),
      'Origen del correo': prop.correo ? prop.origen : '', 'Revisión': n.join(' ') });
    if (prop.correo) plan.resumen.conCorreo++;
    if (n.length) {
      plan.resumen.aCompletar++;
      plan.avisos.push({ prioridad: 'Decidir', hoja: x.orientacion, nivel, materia: x.nombre, detalle: P_.HOJA_BORRADOR + ', fila nueva: ' + n.join(' ') });
    }
  });
  plan.resumen.nuevas = plan.nuevas.length;
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
 * Informe para la hoja Preparación y el registro de ejecución. aplicacion: null (vista previa o con bloqueos) o
 * { error, aplicados, inferidos, revisar, descartados, filasCambios, borrador }.
 */
function informePreparacion_(plan, aplicacion, aplicar) {
  const P_ = PREPARACION;
  const L = [];
  let modo;
  if (!aplicar) modo = 'VISTA PREVIA: no se modificó ninguna grilla ni el borrador de Asignaciones.';
  else if (!plan.puedeAplicar) modo = 'NO SE APLICÓ: hay bloqueos generales (ver abajo). No se modificó ninguna grilla ni el borrador de Asignaciones.';
  else if (aplicacion && aplicacion.error) modo = 'NO SE TERMINÓ DE APLICAR: ' + aplicacion.error;
  else modo = 'APLICADA.';
  L.push('Preparación de ' + listaY_(P_.OBJETIVOS) + ' · ' + fechaHoraTexto_(plan.ahora) + ' · ' + modo);
  L.push('Grillas (Control del generador): ' + [P_.COPIA + ' (origen de la copia)', P_.EVIDENCIA + ' (evidencia de cursada)'].concat(P_.OBJETIVOS)
    .map((p, i) => p + ' ' + (plan.fuentes[[P_.COPIA, P_.EVIDENCIA].concat(P_.OBJETIVOS)[i]] || '(falta)')).join(' · '));
  P_.OBJETIVOS.forEach(P => {
    const r = plan.porPeriodo[P];
    if (!r) { L.push(P + ': no se preparó (ver bloqueos).'); return; }
    L.push(P + ': ' + plural_(r.estudiantes, 'estudiante', 'estudiantes') + ' (' + CARGA.HOJAS_ORIGEN.map(h => h + ' ' + (r.hojas[h] || 0)).join(', ') + '). ' +
      (aplicar && aplicacion && !aplicacion.error ? 'Se pasaron' : 'Se pasarían') + ' a CURSA ' + r.aCursa + ' calificaciones' + (r.aCursaSinCopia ? ' (' + r.aCursaSinCopia + ' en filas o columnas que no venían de ' + P_.COPIA + ')' : '') +
      ' y se ' + (aplicar && aplicacion && !aplicacion.error ? 'borraron ' : 'borrarían ') + r.borrados + ' Calidad, Participación u Observaciones copiadas de ' + P_.COPIA + '. ' +
      'Se conservan ' + r.yaCursa + ' CURSA ya puestos y ' + r.administrativos + ' ACREDITÓ o EQUIVALENCIA. Sin preparar por falta de evidencia: ' + r.sinPreparar + '.');
    L.push('   Antes de preparar, ' + r.iguales + ' de ' + r.comparadas + ' calificaciones comparables eran iguales a las de ' + P_.COPIA + '. ' +
      'Respecto de ' + P_.COPIA + ': ' + plural_(r.altas, 'estudiante nuevo', 'estudiantes nuevos') + ' y ' + plural_(r.bajas, 'que ya no está', 'que ya no están') + '.');
  });
  const a = plan.asignaciones;
  if (a && !a.problema) {
    L.push(P_.HOJA_BORRADOR + ': ' + plural_(a.resumen.materias, 'materia', 'materias') + ' con cursantes; ' + plural_(a.resumen.existentes, 'fila existente', 'filas existentes') + ' (se conservan); ' +
      plural_(a.resumen.nuevas, 'fila nueva', 'filas nuevas') + ', ' + a.resumen.conCorreo + ' con correo de las protecciones de las grillas y ' + a.resumen.aCompletar + ' a completar. ' +
      'La hoja Asignaciones no se toca: se copia del borrador al abrir (manual, paso E1).');
  }
  if (aplicacion && !aplicacion.error) {
    L.push('Escritura: ' + aplicacion.aplicados + ' celdas aplicadas y verificadas' + (aplicacion.inferidos ? ', ' + aplicacion.inferidos + ' inferidas' : '') + ', ' + aplicacion.revisar + ' a revisar, ' +
      aplicacion.descartados + ' no aplicadas porque cambiaron desde la lectura. Detalle por celda: hoja «' + P_.HOJA_CAMBIOS + '»' + (aplicacion.filasCambios ? ', filas ' + aplicacion.filasCambios : '') + '.');
    if (aplicacion.borrador) L.push(P_.HOJA_BORRADOR + ': ' + aplicacion.borrador);
  }
  L.push('A resolver: ' + P_.PRIORIDADES.map(p => p + ' ' + plan.conteo[p]).join(' · ') +
    '. Bloqueo: impide preparar. Decidir: quedó sin preparar y lo decide una persona. Revisar: se conservó el borrador, pero difiere de la evidencia o de lo esperado.');

  const ancho = P_.INFORME.length;
  const pad = f => f.concat(Array(Math.max(0, ancho - f.length)).fill(''));
  const filas = L.map(l => pad([l])).concat([pad([]), P_.INFORME.slice()],
    plan.avisos.map(x => [x.prioridad, x.periodo, x.hoja, x.nivel, x.materia, x.fila === '' ? '' : x.fila, x.dni, x.estudiante, x.detalle]));
  const lineas = plan.avisos.slice(0, P_.MAX_LINEAS_REGISTRO).map(x => x.prioridad + ' · ' + [x.periodo, x.hoja, x.nivel, x.materia, x.fila ? 'fila ' + x.fila : '', x.estudiante].filter(Boolean).join(' · ') + ': ' + x.detalle);
  const texto = L.concat([''], lineas, plan.avisos.length > lineas.length ? ['… y ' + (plan.avisos.length - lineas.length) + ' más en la hoja «' + P_.HOJA_INFORME + '» de la planilla de administración.'] : []).join('\n');
  return { filas, texto, encabezado: L.length + 1 };
}
