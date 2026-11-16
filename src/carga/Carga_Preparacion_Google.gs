/**
 * Carga docente EPS (#32) · adaptador de la preparación de 2º cuatrimestre y anual (Carga_Preparacion.gs) a los
 * servicios de Google.
 *
 * Funciones para la cuenta propietaria, desde el editor (las dos escriben el informe en la hoja «Preparación» de la
 * planilla de administración y un resumen en el registro de ejecución):
 *   previsualizarPeriodosCarga  arma el plan sin tocar las grillas ni el borrador de Asignaciones. Para revisar antes
 *                               de la primera aplicación y después de cambios grandes en los borradores.
 *   prepararPeriodosCarga       arma el plan y, si no hay bloqueos generales, lo aplica: escribe CURSA y borra los
 *                               restos de la copia en las grillas, anota cada celda en «Preparación · cambios» y
 *                               completa «Asignaciones (borrador)»: una fila por materia con cursantes, la columna
 *                               Docente con la lista de la hoja «Docentes» (la crea vacía si falta) y la columna
 *                               Correo con una fórmula que lo toma de esa nómina. Se puede repetir: lo ya preparado
 *                               no cambia y las decisiones posteriores se respetan.
 *
 * Ninguna toca Ventanas, Asignaciones, Enlaces, Registro ni Correos, ni abre ventanas o envía correos. Las grillas se
 * abren por los ID de Control del generador, como en la carga. La escritura se hace con el lock del script (el del
 * guardado docente) y vuelve a leer cada hoja antes de escribirla: una celda que alguien cambió desde la lectura
 * no se toca.
 */

function previsualizarPeriodosCarga() {
  soloPropietario_();
  return ejecutarPreparacion_(false);
}

function prepararPeriodosCarga() {
  soloPropietario_();
  return ejecutarPreparacion_(true);
}

function ejecutarPreparacion_(aplicar) {
  const inicio = Date.now();
  const servicios = serviciosCarga_();
  const admin = servicios.admin();
  const entrada = leerPreparacion_(servicios, admin);
  const plan = planificarPreparacion_(entrada);
  const aplicacion = aplicar && plan.puedeAplicar ? aplicarPreparacion_(servicios, admin, plan) : null;
  const informe = informePreparacion_(plan, aplicacion, aplicar);
  escribirInformePrep_(admin, informe);
  console.log(informe.texto + '\n\n(' + Math.round((Date.now() - inicio) / 1000) + ' s; informe completo en la hoja «' + PREPARACION.HOJA_INFORME + '» de la planilla de administración)');
  const resumen = { aplicar, puedeAplicar: plan.puedeAplicar, cambios: plan.cambios.length, conteo: plan.conteo, aplicacion };
  if (aplicar && (!plan.puedeAplicar || (aplicacion && aplicacion.error))) {
    throw new Error('La preparación no se aplicó' + (aplicacion && aplicacion.error ? ' completa: ' + aplicacion.error : ': ' + plan.bloqueos.join(' ')) + ' Ver la hoja «' + PREPARACION.HOJA_INFORME + '».');
  }
  return resumen;
}

/** Lecturas que necesita el plan: configuración, las cuatro grillas, el borrador, la nómina, los cambios previos y el Registro. */
function leerPreparacion_(servicios, admin) {
  const datos = servicios.datos();
  const ahora = servicios.reloj();
  const config = configuracion_(datos, ahora);
  const grillas = {};
  [PREPARACION.COPIA, PREPARACION.EVIDENCIA].concat(PREPARACION.OBJETIVOS).forEach(p => {
    const id = config.periodos[p].fuente;
    if (!id) return;
    let libro;
    try { libro = SpreadsheetApp.openById(id); } catch (e) { grillas[p] = { problema: 'No se pudo abrir la grilla de ' + p + ' (' + e.message + ').' }; return; }
    grillas[p] = { hojas: {} };
    CARGA.HOJAS_ORIGEN.forEach(h => {
      const hoja = libro.getSheetByName(h);
      grillas[p].hojas[h] = hoja ? hoja.getDataRange().getDisplayValues() : null;
    });
  });
  const tabla = nombre => { const h = admin.getSheetByName(nombre); return h ? h.getDataRange().getDisplayValues() : null; };
  return { ahora, datos, grillas, borrador: tabla(PREPARACION.HOJA_BORRADOR), nomina: tabla(PREPARACION.HOJA_DOCENTES),
    cambiosPrevios: tabla(PREPARACION.HOJA_CAMBIOS), registro: tabla(CARGA.HOJAS.registro) };
}

/**
 * Aplica el plan con el lock del script. Antes de escribir cada hoja la vuelve a leer y descarta lo que cambió
 * (confirmarCambiosPrep_); anota cada celda en «Preparación · cambios» como pendiente antes de escribirla y la marca
 * después de leerla de nuevo (verificarCambiosPrep_). Nunca lanza: un error queda en «error».
 */
function aplicarPreparacion_(servicios, admin, plan) {
  const r = { error: '', aplicados: 0, inferidos: 0, revisar: 0, descartados: 0, filasCambios: '', borrador: '' };
  const lock = servicios.lock();
  if (!lock.tryLock(CARGA.ESPERA_LOCK_MS)) { r.error = 'hay otra operación de la carga en curso. No se modificó nada; volver a ejecutar en unos minutos.'; return r; }
  const filas = [];
  try {
    r.error = revalidarPrep_(servicios.datos(), servicios.reloj(), plan.fuentes);
    if (r.error) return r;
    const log = hojaCambiosPrep_(admin);
    const grupos = {};
    plan.cambios.forEach(c => { const k = c.fuente + '|' + c.hoja; (grupos[k] = grupos[k] || []).push(c); });
    Object.keys(grupos).forEach(k => {
      const cambios = grupos[k];
      const hoja = SpreadsheetApp.openById(cambios[0].fuente).getSheetByName(cambios[0].hoja);
      const conf = confirmarCambiosPrep_(hoja ? hoja.getDataRange().getDisplayValues() : [], cambios[0].periodo, cambios);
      const ahora = servicios.reloj();
      if (conf.descartados.length) filas.push(anotarCambiosPrep_(log, conf.descartados.map(d => filaCambioPrep_(ahora, 'no aplicado', d.detalle, d.cambio))));
      r.descartados += conf.descartados.length;
      if (!conf.aplicables.length) return;
      const primera = anotarCambiosPrep_(log, conf.aplicables.map(c => filaCambioPrep_(ahora, 'pendiente', '', c)));
      filas.push(primera, primera + conf.aplicables.length - 1);
      let error = '';
      try {
        trozosPrep_(conf.aplicables.filter(c => c.nuevo).map(c => c.a1)).forEach(l => hoja.getRangeList(l).setValue("'" + CARGA.PENDIENTE));
        trozosPrep_(conf.aplicables.filter(c => !c.nuevo).map(c => c.a1)).forEach(l => hoja.getRangeList(l).clearContent());
        servicios.flush();
      } catch (e) { error = e.message; }
      const verificados = verificarCambiosPrep_(hoja.getDataRange().getDisplayValues(), conf.aplicables, error);
      log.getRange(primera, 2, verificados.length, 2).setValues(verificados.map(v => [textoPrep_(v.estado), textoPrep_(v.detalle)]));
      verificados.forEach(v => { r[v.estado === 'aplicado' ? 'aplicados' : v.estado === 'inferido' ? 'inferidos' : 'revisar']++; });
    });
    r.borrador = escribirBorradorPrep_(admin, plan);
  } catch (e) {
    r.error = 'error inesperado (' + e.message + '). Revisar «' + PREPARACION.HOJA_CAMBIOS + '»: las filas «pendiente» son celdas que pudieron cambiar.';
  } finally {
    try { servicios.flush(); } finally { lock.releaseLock(); }
  }
  if (filas.length) r.filasCambios = Math.min.apply(null, filas) + ' a ' + Math.max.apply(null, filas);
  return r;
}

/** La hoja «Preparación · cambios», creada si falta. Si sus columnas no son las esperadas, lanza antes de escribir nada. */
function hojaCambiosPrep_(admin) {
  let hoja = admin.getSheetByName(PREPARACION.HOJA_CAMBIOS);
  if (!hoja) {
    hoja = admin.insertSheet(PREPARACION.HOJA_CAMBIOS);
    hoja.getRange(1, 1, 1, PREPARACION.CAMBIOS.length).setValues([PREPARACION.CAMBIOS]).setFontWeight('bold');
    return hoja;
  }
  const enc = hoja.getRange(1, 1, 1, PREPARACION.CAMBIOS.length).getDisplayValues()[0];
  if (PREPARACION.CAMBIOS.some((c, i) => norm_(enc[i]) !== norm_(c))) throw new Error('La hoja «' + PREPARACION.HOJA_CAMBIOS + '» no tiene las columnas esperadas en la fila 1');
  return hoja;
}

function anotarCambiosPrep_(hoja, filas) {
  const primera = Math.max(hoja.getLastRow(), 1) + 1;
  hoja.getRange(primera, 1, filas.length, PREPARACION.CAMBIOS.length).setValues(filas.map(f => f.map(textoPrep_)));
  return primera;
}

/**
 * Borrador de Asignaciones, releído y replanificado dentro del lock: agrega las columnas propias que falten y las
 * filas nuevas, y reescribe Cursantes y Revisión de las existentes. En las filas nuevas y en las existentes sin docente
 * ni correo, el Correo es una fórmula que lo busca en la nómina; la columna Docente ofrece la lista de la nómina.
 * Nunca escribe Orientación, Nivel, Materia, Docente ni Correo de una fila que tenga datos.
 */
function escribirBorradorPrep_(admin, plan) {
  const P_ = PREPARACION;
  let docentes = admin.getSheetByName(P_.HOJA_DOCENTES);
  let hoja = admin.getSheetByName(P_.HOJA_BORRADOR);
  const a = planAsignacionesPrep_({ g: plan.g, proyectado: plan.proyectado, borrador: hoja ? hoja.getDataRange().getDisplayValues() : null,
    nomina: docentes ? docentes.getDataRange().getDisplayValues() : null });
  if (a.problema) return 'no se modificó: ' + a.problema;
  if (!docentes) {
    docentes = admin.insertSheet(P_.HOJA_DOCENTES);
    docentes.getRange(1, 1, 1, P_.NOMINA.length).setValues([P_.NOMINA]).setFontWeight('bold');
  }
  if (!hoja) hoja = admin.insertSheet(P_.HOJA_BORRADOR);
  const ancho = a.encabezado.length;
  const previas = a.crear ? 0 : hoja.getDataRange().getDisplayValues()[0].length;
  if (ancho > previas) hoja.getRange(1, previas + 1, 1, ancho - previas).setValues([a.encabezado.slice(previas).map(textoPrep_)]).setFontWeight('bold');
  const ultima = a.actualizar.reduce((m, u) => Math.max(m, u.fila), 1);
  ['Cursantes', 'Revisión'].forEach(nombre => {
    if (ultima < 2) return;
    const valores = Array.from({ length: ultima - 1 }, () => ['']);
    a.actualizar.forEach(u => { valores[u.fila - 2][0] = textoPrep_(u[nombre]); });
    hoja.getRange(2, a.col[nombre] + 1, ultima - 1, 1).setValues(valores);
  });
  let formulas = a.conFormula.slice();
  if (a.nuevas.length) {
    const desde = Math.max(hoja.getLastRow(), 1) + 1;
    hoja.getRange(desde, 1, a.nuevas.length, ancho).setValues(a.nuevas.map(n => {
      const f = Array(ancho).fill('');
      Object.keys(n).forEach(c => { f[a.col[c]] = textoPrep_(n[c]); });
      return f;
    }));
    formulas = formulas.concat(a.nuevas.map((n, i) => desde + i));
  }
  const colDocente = a.col['Docente'] + 1, colCorreo = a.col['Correo'] + 1;
  formulas.forEach(f => hoja.getRange(f, colCorreo).setFormula(formulaCorreoPrep_(a1_(f - 1, colDocente - 1))));
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila >= 2) {
    const lista = SpreadsheetApp.newDataValidation().requireValueInRange(docentes.getRange('A2:A'), true).setAllowInvalid(false)
      .setHelpText('Elegir un docente de la hoja ' + P_.HOJA_DOCENTES + '. El correo se completa solo.').build();
    hoja.getRange(2, colDocente, ultimaFila - 1, 1).setDataValidation(lista);
  }
  return plural_(a.nuevas.length, 'fila nueva', 'filas nuevas') + '; Cursantes y Revisión actualizadas en ' + plural_(a.actualizar.length, 'fila existente', 'filas existentes') +
    '; fórmula del correo en ' + plural_(formulas.length, 'fila', 'filas') + '; lista de docentes en la columna Docente.';
}

/** Correo de la nómina para el docente de esa celda; vacío si no está elegido o no figura. */
function formulaCorreoPrep_(celdaDocente) {
  return '=IFERROR(VLOOKUP(' + celdaDocente + ',' + PREPARACION.HOJA_DOCENTES + '!$A:$B,2,FALSE),"")';
}

function escribirInformePrep_(admin, informe) {
  const hoja = admin.getSheetByName(PREPARACION.HOJA_INFORME) || admin.insertSheet(PREPARACION.HOJA_INFORME);
  hoja.clearContents();
  hoja.getRange(1, 1, informe.filas.length, PREPARACION.INFORME.length).setValues(informe.filas.map(f => f.map(textoPrep_)));
  [1, informe.encabezadoTipos + 1, informe.encabezado + 1].forEach(f => hoja.getRange(f, 1, 1, PREPARACION.INFORME.length).setFontWeight('bold'));
}

// Texto con apóstrofo inicial, como el Registro: Sheets lo guarda tal cual (ni fórmulas, ni números, ni fechas).
function textoPrep_(v) {
  if (v === undefined || v === null || v === '') return '';
  return typeof v === 'string' ? "'" + v : v;
}

function trozosPrep_(lista) {
  const r = [];
  for (let i = 0; i < lista.length; i += 300) r.push(lista.slice(i, i + 300));
  return r;
}
