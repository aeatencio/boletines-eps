/**
 * Carga docente EPS (#28) · correos de apertura, seguimiento y cierre. Sin servicios de Google: el envío, la hoja
 * Correos, el reloj y la firma de los enlaces llegan por parámetro (Carga_Correo_Google.gs los conecta), así que
 * corre igual en Apps Script y en las pruebas locales (tests/carga-correo.js).
 *
 * Una ejecución por día (activador diario, Carga_Correo_Google.gs). En cada una, enviarCorreosDia_:
 *  1. no hace nada antes de las 8:00 (hora de Buenos Aires), ni sin ventanas en curso o cierres por informar, ni en
 *     un día no hábil de CABA (sábado, domingo o feriado de Carga_Calendario.gs). Si hay algo que informar y el año no
 *     tiene calendario, no adivina: no envía e informa el problema;
 *  2. lee la configuración vigente y la hoja Correos (lo ya enviado);
 *  3. arma a lo sumo un mensaje por docente (planCorreos_), con TODOS sus períodos relevantes de ese día:
 *     - apertura: el primer día hábil de la ventana, a todos los docentes asignados, con el enlace y el estado
 *       inicial. Un docente que no la recibió (asignado más tarde, o con un envío fallido) la recibe el siguiente día
 *       hábil en que el trabajo corre, mientras la ventana siga abierta;
 *     - seguimiento: los días hábiles siguientes, sólo a quien tiene calificaciones CURSA pendientes (o, como mejora
 *       no bloqueante de #28, desaprobaciones con Calidad, Participación u Observaciones sin completar). Sin
 *       pendientes no hay correo; si aparecen nuevos, vuelve a haberlo;
 *     - cierre: el primer día hábil posterior a la ventana (el del cierre de las 8:00), a todos los asignados, con
 *       el resumen final por período. Si ese día el trabajo no corrió o falló, se recupera hasta MAX_DIAS_CIERRE
 *       días después. No certifica que la carga esté completa: el bloqueo lo garantiza el núcleo (#26);
 *     - enlace nuevo: si el Equipo EPS revocó el enlace de un docente (hoja Enlaces), el nuevo le llega el siguiente
 *       día hábil, aunque no tenga pendientes;
 *  4. anota cada mensaje en Correos ('pendiente'), lo envía y lo marca 'enviado' o 'error'. Un docente con un
 *     mensaje 'enviado' o 'pendiente' ese día no recibe otro; con 'error', una nueva ejecución del mismo día reintenta.
 *
 * Contenido: se lee de las grillas, como la página (materiasDocente_), sólo para las asignaciones vigentes del
 * destinatario. El modelo del mensaje (seccionCorreo_) se arma sin DNI y sin el contenido de Calidad,
 * Participación ni Observaciones: sólo nombre, grupo y Calificación de cada estudiante.
 */
const CORREO = {
  FUNCION: 'enviarCorreosCarga',
  HOJA: 'Correos',
  COLUMNAS: ['Fecha', 'Día', 'Estado', 'Detalle', 'Correo', 'Docente', 'Tipo', 'Acceso', 'Eventos', 'Pendientes', 'Asunto'],
  // Hora deseada (y mínima) de envío. El activador diario de Apps Script corre en algún minuto entre las 8:00 y las
  // 9:00; antes de las 8:00 nunca se envía (además, el cierre de las 8:00 tiene que haber ocurrido).
  HORA_ENVIO: 8,
  MAX_DIAS_CIERRE: 7,
  // Margen dentro de los 6 minutos de una ejecución de Apps Script: lo que no se envía queda para otra ejecución.
  LIMITE_MS: 4.5 * 60 * 1000,
  REMITENTE: 'Equipo EPS',
  ASUNTO: 'Carga de valoraciones EPS',
  PROPIEDAD_URL: 'URL_CARGA',
  URL_APP: /^https:\/\/script\.google\.com\/(?:a\/macros\/[A-Za-z0-9.-]+\/s|macros\/s)\/[A-Za-z0-9_-]+\/exec$/,
  // Gmail recorta los mensajes de más de ~102 KB: por encima de esto, el detalle sólo lista lo que requiere atención.
  MAX_HTML: 90000,
  DIAS: ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado']
};

// ================================================================ calendario (días hábiles de CABA)

function anioConFeriados_(dia, feriados) {
  const anio = dia.slice(0, 4);
  return Object.keys(feriados).some(f => f.slice(0, 4) === anio);
}

/** Lunes a viernes y no feriado. feriados: feriadosCalendario_(). Para un año sin calendario, lanza: no se adivina. */
function diaHabil_(dia, feriados) {
  if (!anioConFeriados_(dia, feriados)) throw new Error(sinCalendario_(dia.slice(0, 4), 'no se puede saber si ' + dia + ' es hábil'));
  const semana = new Date(dia + 'T12:00:00Z').getUTCDay();
  return semana !== 0 && semana !== 6 && !feriados[dia];
}

function primerDiaHabilDesde_(dia, feriados) {
  for (let i = 0, d = dia; i < CARGA.MAX_DIAS_HASTA_CIERRE; i++, d = diaSiguiente_(d)) if (diaHabil_(d, feriados)) return d;
  throw new Error('No hay un día hábil en el mes posterior al ' + dia + '.');
}

function diasEntre_(desde, hasta) { return Math.round((Date.parse(hasta + 'T00:00:00Z') - Date.parse(desde + 'T00:00:00Z')) / 86400000); }

function horaLocal_(instante) { return new Date(instante.getTime() + Number(CARGA.DESFASE.slice(0, 3)) * 3600000); }

/** 'lunes 7/12/2026' */
function fechaTexto_(dia) {
  const f = new Date(dia + 'T12:00:00Z');
  return CORREO.DIAS[f.getUTCDay()] + ' ' + f.getUTCDate() + '/' + (f.getUTCMonth() + 1) + '/' + f.getUTCFullYear();
}

/** 'lunes 7/12/2026, 8:05' (hora de Buenos Aires) */
function fechaHoraTexto_(instante) {
  const l = horaLocal_(instante);
  return fechaTexto_(l.toISOString().slice(0, 10)) + ', ' + l.getUTCHours() + ':' + String(l.getUTCMinutes()).padStart(2, '0');
}

// ================================================================ hoja Correos (lo ya enviado)

function correosValidar_(hoja) {
  const enc = hoja.getRange(1, 1, 1, CORREO.COLUMNAS.length).getDisplayValues()[0];
  if (CORREO.COLUMNAS.some((c, i) => norm_(enc[i]) !== norm_(c))) throw new Error('La hoja ' + CORREO.HOJA + ' no tiene las columnas esperadas.');
}

/** Filas de Correos: { fila, dia, correo, estado, acceso, eventos[] }. */
function correosLeer_(hoja) {
  correosValidar_(hoja);
  const ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  const col = Object.fromEntries(CORREO.COLUMNAS.map((c, i) => [c, i]));
  return hoja.getRange(2, 1, ultima - 1, CORREO.COLUMNAS.length).getDisplayValues().map((f, i) => ({
    fila: i + 2, dia: txt_(f[col['Día']]), correo: correo_(f[col['Correo']]), estado: txt_(f[col['Estado']]),
    acceso: txt_(f[col['Acceso']]), eventos: txt_(f[col['Eventos']]).split(/\s*;\s*/).filter(Boolean)
  })).filter(h => h.dia && h.correo);
}

function correosArreglo_(datos) {
  return CORREO.COLUMNAS.map(c => {
    const v = datos[c];
    if (v === undefined || v === null || v === '') return '';
    return typeof v === 'string' ? "'" + v : v;
  });
}

/** Fila de (día, correo): la indicada si sigue ahí, o la única que la tenga (el Equipo EPS pudo ordenar la hoja). */
function correosUbicar_(hoja, fila, dia, correo) {
  const col = c => CORREO.COLUMNAS.indexOf(c) + 1;
  const es = f => txt_(hoja.getRange(f, col('Día')).getDisplayValue()) === dia && correo_(hoja.getRange(f, col('Correo')).getDisplayValue()) === correo;
  if (fila && es(fila)) return fila;
  const filas = correosLeer_(hoja).filter(h => h.dia === dia && h.correo === correo).map(h => h.fila);
  if (filas.length !== 1) throw new Error('No se pudo ubicar en ' + CORREO.HOJA + ' la fila de ' + correo + ' del ' + dia + '.');
  return filas[0];
}

/** Anota el mensaje como 'pendiente' (reutiliza la fila 'error' del mismo día). Devuelve la fila. */
function correosAnotar_(hoja, filaPrevia, datos) {
  const arreglo = correosArreglo_(datos);
  if (filaPrevia) {
    const f = correosUbicar_(hoja, filaPrevia, datos['Día'], datos['Correo']);
    hoja.getRange(f, 1, 1, arreglo.length).setValues([arreglo]);
    return f;
  }
  hoja.appendRow(arreglo);
  return correosUbicar_(hoja, hoja.getLastRow(), datos['Día'], datos['Correo']);
}

function correosMarcar_(hoja, fila, dia, correo, estado, detalle) {
  const f = correosUbicar_(hoja, fila, dia, correo);
  hoja.getRange(f, CORREO.COLUMNAS.indexOf('Estado') + 1, 1, 2).setValues([["'" + estado, detalle ? "'" + detalle : '']]);
}

// ================================================================ plan del día

function escalaNumerica_(periodo) { return CARGA.ESCALA_DOCENTE[periodo].indexOf('10') >= 0; }

/** Desaprobación: 5 o menos en los períodos numéricos; «En proceso» en los conceptuales. */
function desaprobada_(periodo, estudiante) {
  if (estudiante.situacion !== 'valorada') return false;
  const v = estudiante.campos.calificacion;
  return escalaNumerica_(periodo) ? Number(v) <= 5 : norm_(v) === 'EN PROCESO';
}

/** Desaprobación con Calidad, Participación u Observaciones vacías (sólo si están vacías: nunca su contenido). */
function complementariosFaltan_(periodo, estudiante) {
  return desaprobada_(periodo, estudiante) && ['calidad', 'participacion', 'observaciones'].some(c => !txt_(estudiante.campos[c]));
}

/**
 * Sección del mensaje para un período. Es lo único que llega a la redacción: sin DNI y sin el contenido de los
 * campos complementarios (sólo si faltan, en una desaprobación).
 */
function seccionCorreo_(p, evento, materias) {
  const secc = {
    periodo: p.periodo, evento, desde: p.ventana.desde, hasta: p.ventana.hasta, cierre: p.ventana.cierre,
    pendientes: 0, completar: 0,
    materias: materias.map(m => {
      const estudiantes = m.problema ? [] : m.estudiantes.map(e => ({
        nombre: e.nombre, grupo: e.grupo, situacion: e.situacion,
        calificacion: e.situacion === 'pendiente' ? '' : e.campos.calificacion,
        completar: complementariosFaltan_(p.periodo, e)
      }));
      return {
        orientacion: m.orientacion, nivel: m.nivel, materia: m.materia, problema: m.problema || '',
        conteo: m.conteo || null, completar: estudiantes.filter(e => e.completar).length, estudiantes
      };
    })
  };
  secc.materias.forEach(m => { if (m.conteo) { secc.pendientes += m.conteo.pendiente; secc.completar += m.completar; } });
  return secc;
}

/**
 * Qué mensajes corresponden hoy. config: configuracion_; historial: correosLeer_; hoja: lectorHojas_;
 * enlace(correo): enlaceVigente_ ({ credencial, version, tramo } o { problema }).
 * Devuelve { dia, motivo, envios[], problemas[], yaRecibieron[] }. Nunca lanza por un problema de datos.
 */
function planCorreos_(config, ahora, historial, hoja, enlace) {
  const dia = fechaLocalDe_(ahora);
  const plan = { dia, motivo: '', envios: [], problemas: [], yaRecibieron: [] };
  if (horaLocal_(ahora).getUTCHours() < CORREO.HORA_ENVIO) { plan.motivo = 'Antes de las ' + CORREO.HORA_ENVIO + ':00 no se envían correos.'; return plan; }
  if (config.problemas.length) {
    plan.problemas = config.problemas.map(x => 'Configuración: ' + x + ' No se envía ningún correo hasta resolverlo.');
    return plan;
  }
  // Sin ventanas en curso ni cierres recientes no hay nada que informar, y el calendario no hace falta (#29).
  const relevante = p => p.ventana && (p.estado === 'abierta' ||
    (p.estado === 'cerrada' && p.ventana.diaCierre && diasEntre_(p.ventana.diaCierre, dia) >= 0 && diasEntre_(p.ventana.diaCierre, dia) <= CORREO.MAX_DIAS_CIERRE));
  if (!CARGA.PERIODOS.some(n => relevante(config.periodos[n]))) { plan.motivo = 'No hay ventanas en curso ni cierres para informar.'; return plan; }
  try {
    if (!diaHabil_(dia, config.feriados)) { plan.motivo = fechaTexto_(dia) + ' no es día hábil: no se envían correos.'; return plan; }
  } catch (e) { plan.problemas.push(e.message); return plan; }

  // Períodos que se informan hoy: abiertos desde su primer día hábil, y cerrados hace poco.
  const abiertas = [], cierres = [];
  CARGA.PERIODOS.forEach(n => {
    const p = config.periodos[n];
    if (!p.ventana) return;
    if (p.estado === 'abierta') {
      let primero;
      try { primero = primerDiaHabilDesde_(p.ventana.desde, config.feriados); } catch (e) { plan.problemas.push(n + ': ' + e.message); return; }
      if (primero > dia) return;
      if (p.problemas.length) return plan.problemas.push(n + ': no se envían sus correos hasta resolver: ' + p.problemas.join(' '));
      abiertas.push({ p, clave: 'abierta|' + n + '|' + p.ventana.desde });
    } else if (p.estado === 'cerrada' && p.ventana.diaCierre) {
      const atraso = diasEntre_(p.ventana.diaCierre, dia);
      if (atraso < 0 || atraso > CORREO.MAX_DIAS_CIERRE) return;
      if (p.problemas.length) return plan.problemas.push(n + ': no se envía su cierre hasta resolver: ' + p.problemas.join(' '));
      cierres.push({ p, clave: 'cierre|' + n + '|' + p.ventana.desde + '|' + p.ventana.hasta });
    }
  });
  if (!abiertas.length && !cierres.length) { plan.motivo = 'No hay ventanas en curso ni cierres para informar.'; return plan; }

  docentesAsignados_(config).forEach(d => {
    const propios = historial.filter(h => h.correo === d.correo);
    const hoy = propios.find(h => h.dia === dia);
    const vigentes = propios.filter(h => h.estado === 'enviado' || h.estado === 'pendiente');
    const informado = {};
    vigentes.forEach(h => h.eventos.forEach(ev => { informado[ev] = true; }));
    const propias = config.asignaciones.filter(a => a.correo === d.correo);

    const secciones = abiertas.map(a => {
      const s = seccionCorreo_(a.p, '', materiasDocente_(propias, a.p, hoja));
      s.evento = !informado[a.clave] ? 'apertura' : s.pendientes || s.completar ? 'seguimiento' : 'en curso';
      s.clave = a.clave;
      return s;
    }).concat(cierres.filter(c => !informado[c.clave]).map(c => {
      const s = seccionCorreo_(c.p, 'cierre', materiasDocente_(propias, c.p, hoja));
      s.clave = c.clave;
      return s;
    }));

    let acceso = null, nuevoEnlace = false;
    if (abiertas.length) {
      try { acceso = enlace(d.correo); } catch (e) { acceso = { problema: e.message }; }
      if (acceso.problema) acceso = Object.assign(acceso, { clave: '' });
      else {
        acceso.clave = acceso.tramo.desde + '|v' + acceso.version;
        nuevoEnlace = !vigentes.some(h => h.acceso === acceso.clave);
      }
    }
    const tipos = [];
    if (secciones.some(s => s.evento === 'apertura')) tipos.push('apertura');
    if (secciones.some(s => s.evento === 'seguimiento')) tipos.push('seguimiento');
    if (secciones.some(s => s.evento === 'cierre')) tipos.push('cierre');
    if (nuevoEnlace && !tipos.includes('apertura')) tipos.push('enlace');
    if (!tipos.length) return;
    if (hoy && hoy.estado !== 'error') return plan.yaRecibieron.push(d.correo);
    if (acceso && acceso.problema) return plan.problemas.push(d.correo + ': no se le envía el correo porque no se puede armar su enlace: ' + acceso.problema);
    plan.envios.push({
      correo: d.correo, docente: d.docente, dia, tipos, secciones, credencial: acceso ? acceso.credencial : '', acceso: acceso ? acceso.clave : '',
      eventos: secciones.map(s => s.clave), pendientes: secciones.filter(s => s.evento !== 'cierre').reduce((n, s) => n + s.pendientes, 0),
      // Para la hoja Correos: también lo que quedó CURSA en los períodos que cierran.
      pendientesTotal: secciones.reduce((n, s) => n + s.pendientes, 0),
      filaError: hoy ? hoy.fila : 0
    });
  });
  return plan;
}

// ================================================================ redacción

function esc_(t) { return String(t).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

function listaY_(l) { return l.length < 2 ? l.join('') : l.slice(0, -1).join(', ') + ' y ' + l[l.length - 1]; }

function plural_(n, uno, varios) { return n + ' ' + (n === 1 ? uno : varios); }

/** 'cierra el lunes 14/12/2026 a las 8:00' o 'cerró el …' */
function plazoTexto_(s) {
  const dia = fechaLocalDe_(s.cierre);
  return (s.evento === 'cierre' ? 'Cerró el ' : 'Cierra el ') + fechaTexto_(dia) + ' a las ' + CARGA.HORA_CIERRE.replace(/^0/, '') + '.';
}

const ORDEN_SITUACION_ = { pendiente: 0, 'a revisar': 1, valorada: 2, administrativo: 3 };

function calificacionTexto_(e) {
  if (e.situacion === 'pendiente') return 'Pendiente';
  const extra = e.situacion === 'administrativo' ? ' (Equipo EPS)' : e.situacion === 'a revisar' ? ' (a revisar con el Equipo EPS)' : '';
  return e.calificacion + extra + (e.completar ? ' · completar Calidad, Participación u Observaciones' : '');
}

/**
 * Asunto, HTML y texto del mensaje de un envío. enlaceUrl: '' si no hay períodos abiertos. resumido: el detalle
 * sólo lista lo que requiere atención (pendientes, a revisar y desaprobaciones por completar).
 */
function redactarCorreo_(envio, ahora, enlaceUrl, resumido) {
  const abiertas = envio.secciones.filter(s => s.evento !== 'cierre');
  const cerradas = envio.secciones.filter(s => s.evento === 'cierre');
  const aperturas = abiertas.filter(s => s.evento === 'apertura').map(s => s.periodo);
  const pendientes = abiertas.reduce((n, s) => n + s.pendientes, 0);
  const completar = abiertas.reduce((n, s) => n + s.completar, 0);

  const partes = [];
  if (aperturas.length) partes.push('abrió ' + listaY_(aperturas));
  if (abiertas.length && pendientes) {
    const proximo = abiertas.map(s => s.cierre).sort((a, b) => a - b)[0];
    const l = horaLocal_(proximo);
    partes.push(plural_(pendientes, 'calificación pendiente', 'calificaciones pendientes') + ', cierra el ' + l.getUTCDate() + '/' + (l.getUTCMonth() + 1));
  } else if (abiertas.length && completar) partes.push(plural_(completar, 'valoración por completar', 'valoraciones por completar'));
  if (cerradas.length) partes.push('cerró ' + listaY_(cerradas.map(s => s.periodo)) + ', resumen final');
  if (envio.tipos.includes('enlace')) partes.push('tu nuevo enlace');
  if (!partes.length) partes.push('estado de tu carga');
  const asunto = CORREO.ASUNTO + ': ' + partes.join(' · ');

  const intro = [];
  if (aperturas.length) intro.push('Ya podés cargar las valoraciones de ' + listaY_(aperturas) + '.');
  if (envio.tipos.includes('enlace')) intro.push('Tu enlace de acceso cambió: usá el de este correo. El anterior ya no funciona.');
  if (abiertas.length) intro.push(pendientes ? 'Tenés ' + plural_(pendientes, 'calificación pendiente', 'calificaciones pendientes') + '.' : 'No tenés calificaciones pendientes.');
  if (abiertas.length && completar) intro.push(plural_(completar, 'desaprobación tiene', 'desaprobaciones tienen') + ' Calidad, Participación u Observaciones sin completar.');
  if (cerradas.length) intro.push('Cerró la carga de ' + listaY_(cerradas.map(s => s.periodo)) + '. Abajo está el resumen final de lo que figuraba guardado en las grillas.');
  const estadoAl = 'Estado de las grillas al ' + fechaHoraTexto_(ahora) + ' (hora de Buenos Aires).';

  const titulo = s => s.periodo + ' · ' + (s.evento === 'cierre' ? 'resumen final' : s.evento === 'apertura' ? 'abrió la carga' : 'carga abierta');
  const nombreGrupo = (e, m) => e.nombre + (e.grupo && e.grupo !== m.nivel ? ' (' + e.grupo + ')' : '');
  const visibles = m => {
    const l = m.estudiantes.slice().sort((a, b) => ORDEN_SITUACION_[a.situacion] - ORDEN_SITUACION_[b.situacion] || (b.completar ? 1 : 0) - (a.completar ? 1 : 0) || a.nombre.localeCompare(b.nombre, 'es'));
    return resumido ? l.filter(e => e.situacion === 'pendiente' || e.situacion === 'a revisar' || e.completar) : l;
  };
  const cuenta = m => m.conteo ? [m.conteo.valorada, m.conteo.pendiente, m.conteo.administrativo, m.conteo['a revisar']] : null;

  // ---- HTML (estilos en línea: los clientes de correo ignoran <style>)
  const td = 'style="padding:4px 8px;border-bottom:1px solid #e3e3e3;text-align:left"';
  const tdn = 'style="padding:4px 8px;border-bottom:1px solid #e3e3e3;text-align:right"';
  const h = [];
  h.push('<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.45;color:#1f1f1f;max-width:680px">');
  h.push('<p>Hola, ' + esc_(envio.docente) + '.</p>');
  intro.forEach(t => h.push('<p style="margin:0 0 8px">' + esc_(t) + '</p>'));
  h.push('<p style="margin:0 0 12px;color:#555">' + esc_(estadoAl) + '</p>');
  if (enlaceUrl) h.push('<p style="margin:16px 0"><a href="' + esc_(enlaceUrl) + '" style="background:#0b57d0;color:#ffffff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:bold">Abrir mi carga</a></p>');
  envio.secciones.forEach(s => {
    h.push('<h2 style="font-size:17px;margin:24px 0 2px">' + esc_(titulo(s)) + '</h2>');
    h.push('<p style="margin:0 0 8px;color:#444">' + esc_(plazoTexto_(s)) + (s.evento !== 'cierre' ? ' ' + esc_(plural_(s.pendientes, 'pendiente', 'pendientes')) + '.' : '') + '</p>');
    h.push('<table cellspacing="0" style="border-collapse:collapse;font-size:13px"><tr><th ' + td + '>Materia</th><th ' + tdn + '>Valoradas</th><th ' + tdn + '>Pendientes</th><th ' + tdn + '>Equipo EPS</th><th ' + tdn + '>A revisar</th></tr>');
    s.materias.forEach(m => {
      const c = cuenta(m);
      h.push('<tr><td ' + td + '>' + esc_(m.orientacion + ' ' + m.nivel + ' · ' + m.materia) + '</td>' + (c
        ? c.map((n, i) => '<td ' + tdn + '>' + (i === 1 && n ? '<b>' + n + '</b>' : n) + '</td>').join('')
        : '<td colspan="4" style="padding:4px 8px;border-bottom:1px solid #e3e3e3;color:#b3261e">' + esc_(m.problema) + '</td>') + '</tr>');
    });
    h.push('</table>');
    s.materias.filter(m => !m.problema && m.estudiantes.length).forEach(m => {
      const l = visibles(m);
      h.push('<h3 style="font-size:14px;margin:16px 0 4px">' + esc_(m.orientacion + ' ' + m.nivel + ' · ' + m.materia) + '</h3>');
      if (m.completar) h.push('<p style="margin:0 0 4px;color:#8a5300">' + esc_(plural_(m.completar, 'desaprobación sin', 'desaprobaciones sin') + ' Calidad, Participación u Observaciones completas.') + '</p>');
      if (!l.length) { h.push('<p style="margin:0;color:#555">Nada que requiera atención.</p>'); return; }
      h.push('<table cellspacing="0" style="border-collapse:collapse;font-size:13px">' + l.map(e => '<tr><td ' + td + '>' + esc_(nombreGrupo(e, m)) + '</td><td ' + td + '>' +
        (e.situacion === 'pendiente' ? '<b>Pendiente</b>' : esc_(calificacionTexto_(e))) + '</td></tr>').join('') + '</table>');
    });
  });
  const pie = [];
  if (resumido) pie.push('El detalle completo no entra en un correo: este sólo lista lo que requiere atención. Todo lo demás está en la carga.');
  if (enlaceUrl) pie.push('El enlace es personal y sirve durante toda esta carga: no reenvíes este correo. Si creés que alguien más lo usó, avisá al Equipo EPS para reemplazarlo.');
  if (cerradas.length) pie.push('El resumen final no certifica que la carga esté completa. Las correcciones fuera de término se gestionan con el Equipo EPS.');
  pie.push('ACREDITÓ y EQUIVALENCIA los establece el Equipo EPS. Este correo se envía como máximo una vez por día hábil.');
  h.push('<p style="color:#666;font-size:12px;margin-top:24px">' + pie.map(esc_).join('<br>') + '</p></div>');

  // ---- texto
  const t = ['Hola, ' + envio.docente + '.', ''].concat(intro, [estadoAl, '']);
  if (enlaceUrl) t.push('Abrir mi carga: ' + enlaceUrl, '');
  envio.secciones.forEach(s => {
    t.push('== ' + titulo(s) + ' ==', plazoTexto_(s));
    s.materias.forEach(m => {
      const c = cuenta(m);
      t.push('', '- ' + m.orientacion + ' ' + m.nivel + ' · ' + m.materia + ': ' + (c ? c[0] + ' valoradas, ' + c[1] + ' pendientes, ' + c[2] + ' del Equipo EPS, ' + c[3] + ' a revisar' : m.problema));
      if (!m.problema) visibles(m).forEach(e => t.push('    ' + nombreGrupo(e, m) + ': ' + calificacionTexto_(e)));
    });
    t.push('');
  });
  t.push(pie.join('\n'));
  return { asunto, html: h.join(''), texto: t.join('\n') };
}

// ================================================================ ejecución del día

/**
 * Trabajo diario. s: { reloj, datos, valoresHoja, firmar, correos() (hoja Correos), urlCarga(), enviar(mensaje),
 * cuota(), flush, ms() (tiempo transcurrido) }. opciones.simular: arma y redacta sin anotar ni enviar.
 * Devuelve { dia, motivo, enviados[], errores[], sinProcesar[], yaRecibieron[], problemas[], mensajes[] }: el
 * adaptador lo registra y, si hubo problemas, errores o envíos sin procesar, hace fallar la ejecución para que
 * Google avise a la cuenta propietaria.
 */
function enviarCorreosDia_(s, opciones) {
  const simular = !!(opciones && opciones.simular);
  const ahora = s.reloj();
  const r = { dia: fechaLocalDe_(ahora), al: ahora.toISOString(), simulado: simular, motivo: '', enviados: [], errores: [], sinProcesar: [], yaRecibieron: [], problemas: [], mensajes: [] };
  if (horaLocal_(ahora).getUTCHours() < CORREO.HORA_ENVIO) { r.motivo = 'Antes de las ' + CORREO.HORA_ENVIO + ':00 no se envían correos.'; return r; }

  const datos = s.datos();
  const config = configuracion_(datos, ahora);
  const hoja = s.correos();
  const historial = correosLeer_(hoja);
  const plan = planCorreos_(config, ahora, historial, lectorHojas_(s), correo => enlaceVigente_(config, datos.enlaces, s.firmar, correo, ahora));
  r.motivo = plan.motivo;
  r.problemas = plan.problemas.slice();
  r.yaRecibieron = plan.yaRecibieron;
  const url = txt_(s.urlCarga());
  const urlValida = CORREO.URL_APP.test(url);

  plan.envios.forEach(envio => {
    const enlaceUrl = envio.credencial ? url + '?' + CARGA_WEB.PARAMETRO + '=' + envio.credencial : '';
    if (enlaceUrl && !urlValida) {
      return r.problemas.push(envio.correo + ': no se le envía el correo porque la propiedad ' + CORREO.PROPIEDAD_URL + ' no tiene la URL /exec de la app web.');
    }
    let m = redactarCorreo_(envio, ahora, enlaceUrl, false);
    if (m.html.length > CORREO.MAX_HTML) m = redactarCorreo_(envio, ahora, enlaceUrl, true);
    const resumen = { correo: envio.correo, tipos: envio.tipos, periodos: envio.secciones.map(x => x.periodo + ' (' + x.evento + ')'), pendientes: envio.pendientes, asunto: m.asunto, bytes: m.html.length };
    if (simular) { r.mensajes.push(Object.assign(resumen, { html: m.html, texto: m.texto })); return; }
    if (s.ms() > CORREO.LIMITE_MS) return r.sinProcesar.push(envio.correo);
    if (s.cuota() < 1) return r.sinProcesar.push(envio.correo);

    const fila = {
      'Fecha': ahora, 'Día': envio.dia, 'Estado': 'pendiente', 'Detalle': '', 'Correo': envio.correo, 'Docente': envio.docente,
      'Tipo': envio.tipos.join(', '), 'Acceso': envio.acceso, 'Eventos': envio.eventos.join('; '), 'Pendientes': envio.pendientesTotal, 'Asunto': m.asunto
    };
    // Lo anotado precede al envío: si la ejecución se corta después de enviar, la fila 'pendiente' evita un duplicado.
    const f = correosAnotar_(hoja, envio.filaError, fila);
    s.flush();
    let error = null;
    try {
      s.enviar({ to: envio.correo, subject: m.asunto, body: m.texto, htmlBody: m.html, name: CORREO.REMITENTE });
    } catch (e) { error = e; }
    try {
      correosMarcar_(hoja, f, envio.dia, envio.correo, error ? 'error' : 'enviado', error ? 'MailApp: ' + error.message : '');
      s.flush();
    } catch (e) {
      r.problemas.push(envio.correo + ': no se pudo marcar el envío en ' + CORREO.HOJA + ' (queda «pendiente»; no se reenvía hoy): ' + e.message);
    }
    if (error) r.errores.push({ correo: envio.correo, error: error.message });
    else r.enviados.push(resumen);
  });
  return r;
}
