/**
 * Generador de informes individuales EPS v0.1.1
 * Corrección menor: Año y división muestra Informática 1 / Gastronomía 1.
 *
 * Esta versión mantiene el alcance de v0.1: un informe individual de 1º bimestre.
 */
const CONFIG_EPS = {
  HOJA_CONTROL: 'Control',
  HOJA_PLANTILLA: 'Plantilla',
  HOJA_SALIDA: 'Informe_actual',
  CELDA_FUENTE: 'B5',
  CELDA_HOJA_ORIGEN: 'B6',
  CELDA_DNI: 'B7',
  CELDA_CICLO: 'B9',
  ESTADO_FECHA: 'F2',
  ESTADO_RESULTADO: 'F3',
  ESTADO_ADVERTENCIAS: 'F4',
  FILA_INICIO_MATERIAS: 9,
  FILAS_MATERIAS_PLANTILLA: 8,
  FILA_OBS_GENERALES_BASE: 20,
  FILA_OBS_ASISTENCIA_BASE: 23,
  FILA_OBS_CONVIVENCIA_BASE: 26,
  FILA_PTF_BASE: 29
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Informes EPS')
    .addItem('Generar informe actual', 'generarInformeActual')
    .addToUi();
}

function generarInformeActual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const control = ss.getSheetByName(CONFIG_EPS.HOJA_CONTROL);
  const advertencias = [];
  try {
    if (!control) throw new Error('No se encontró la hoja Control.');
    const fuente = texto_(control.getRange(CONFIG_EPS.CELDA_FUENTE).getDisplayValue());
    const hojaOrigen = normalizarHojaOrigen_(control.getRange(CONFIG_EPS.CELDA_HOJA_ORIGEN).getDisplayValue());
    const dni = normalizarDni_(control.getRange(CONFIG_EPS.CELDA_DNI).getDisplayValue());
    const ciclo = texto_(control.getRange(CONFIG_EPS.CELDA_CICLO).getDisplayValue()) || '2026';
    if (!fuente) throw new Error('Falta la URL o ID de la planilla de valoraciones.');
    if (!hojaOrigen) throw new Error('La hoja origen debe ser EPSI o EPSG.');
    if (!dni) throw new Error('Falta completar el DNI del estudiante.');
    const registro = leerFuenteBimestre_(fuente, hojaOrigen, dni, advertencias);
    crearInforme_(ss, registro, ciclo, advertencias);
    registrarEstado_(control, 'Generado: ' + registro.nombreCompleto, advertencias);
    ui.alert('Informe generado correctamente.' + (advertencias.length ? '\n\nAdvertencias:\n- ' + advertencias.join('\n- ') : ''));
  } catch (err) {
    if (control) registrarEstado_(control, 'Error: ' + err.message, []);
    ui.alert('No se pudo generar el informe.\n\n' + err.message);
    throw err;
  }
}

function leerFuenteBimestre_(fuenteIngresada, hojaOrigen, dniBuscado, advertencias) {
  const idFuente = extraerSpreadsheetId_(fuenteIngresada);
  const libroFuente = SpreadsheetApp.openById(idFuente);
  const hoja = libroFuente.getSheetByName(hojaOrigen);
  if (!hoja) throw new Error('No se encontró la hoja ' + hojaOrigen + ' en la planilla de valoraciones.');
  const valores = hoja.getDataRange().getDisplayValues();
  const estructura = detectarEstructuraFuente_(valores);
  const filas = buscarFilasPorDni_(valores, dniBuscado);
  if (filas.length === 0) throw new Error('No se encontró el DNI indicado.');
  if (filas.length > 1) throw new Error('El DNI indicado aparece más de una vez.');
  return construirRegistroEstudiante_(valores, filas[0], estructura, hojaOrigen, advertencias);
}

function detectarEstructuraFuente_(valores) {
  const h1 = valores[0] || [];
  const h2 = valores[1] || [];
  const materias = [];
  const especiales = {};
  const maxCols = Math.max(h1.length, h2.length);
  for (let c = 0; c < maxCols; c++) {
    const combinado = normalizarTexto_([h1[c], h2[c]].filter(Boolean).join(' '));
    if (combinado.includes('ORIENTACION')) especiales.orientacion = c;
    if (combinado.includes('OBSERVACIONES GENERALES')) especiales.observacionesGenerales = c;
    if (combinado.includes('ASISTENCIA')) especiales.observacionesAsistencia = c;
    if (combinado.includes('CONVIVENCIA')) especiales.observacionesConvivencia = c;
    if (combinado.includes('PLAN DE TRABAJO') || combinado.includes('PTF')) especiales.ptf = c;
  }
  for (let c = 3; c < maxCols - 1; c++) {
    const nombre = texto_(h2[c]) || texto_(h1[c]);
    const n = normalizarTexto_(nombre);
    if (!nombre || /^EPS\d+$/.test(n)) continue;
    if (n.includes('ORIENTACION') || n.includes('OBSERVACIONES') || n.includes('ASISTENCIA') || n.includes('CONVIVENCIA') || n.includes('PLAN DE TRABAJO') || n.includes('PTF')) continue;
    materias.push({ nombre, clave: claveMateria_(nombre), columnaEtiqueta: c, columnaValor: c + 1 });
  }
  return { materias, especiales };
}

function buscarFilasPorDni_(valores, dniBuscado) {
  const filas = [];
  for (let r = 0; r < valores.length; r++) if (normalizarDni_(valores[r][2]) === dniBuscado) filas.push(r);
  return filas;
}

function construirRegistroEstudiante_(valores, fila, estructura, hojaOrigen, advertencias) {
  const row = valores[fila] || [];
  const apellido = texto_(row[0]);
  const nombres = texto_(row[1]);
  const nivel = detectarNivelPorFila_(valores, fila);
  const orientacion = obtenerValor_(valores, fila, estructura.especiales.orientacion);
  const materias = [];
  estructura.materias.forEach(m => {
    const valor = obtenerValor_(valores, fila, m.columnaValor);
    const c = clasificarValorBimestre_(valor);
    if (c.estado === 'valido') materias.push({ nombre: m.nombre, valor: c.valor });
    else if (c.estado === 'invalido') advertencias.push('Valor inesperado en "' + m.nombre + '": "' + valor + '".');
  });
  return {
    nombreCompleto: [apellido, nombres].filter(Boolean).join(', ') || '(sin nombre en la fuente)',
    dni: normalizarDni_(row[2]),
    nivelVisual: construirNivelVisual_(hojaOrigen, nivel, orientacion),
    materias,
    observacionesGenerales: obtenerValor_(valores, fila, estructura.especiales.observacionesGenerales),
    observacionesAsistencia: obtenerValor_(valores, fila, estructura.especiales.observacionesAsistencia),
    observacionesConvivencia: obtenerValor_(valores, fila, estructura.especiales.observacionesConvivencia),
    ptf: obtenerValor_(valores, fila, estructura.especiales.ptf)
  };
}

function clasificarValorBimestre_(valor) {
  const crudo = texto_(valor);
  const norm = normalizarTexto_(crudo);
  if (!norm || norm === 'NO CURSA' || norm === 'CURSA') return { estado: 'excluido' };
  const mapa = { 'EN PROCESO': 'En proceso', 'SUFICIENTE': 'Suficiente', 'AVANZADO': 'Avanzado', 'ACREDITO': 'ACREDITÓ' };
  if (mapa[norm]) return { estado: 'valido', valor: mapa[norm] };
  return { estado: 'invalido' };
}

function crearInforme_(ss, registro, ciclo, advertencias) {
  const plantilla = ss.getSheetByName(CONFIG_EPS.HOJA_PLANTILLA);
  if (!plantilla) throw new Error('No se encontró la hoja Plantilla.');
  const salidaExistente = ss.getSheetByName(CONFIG_EPS.HOJA_SALIDA);
  if (salidaExistente) ss.deleteSheet(salidaExistente);
  const salida = plantilla.copyTo(ss);
  salida.setName(CONFIG_EPS.HOJA_SALIDA);
  const n = Math.max(registro.materias.length, 1);
  const base = CONFIG_EPS.FILAS_MATERIAS_PLANTILLA;
  const inicio = CONFIG_EPS.FILA_INICIO_MATERIAS;
  salida.getRange('A1:A3').clearContent();
  salida.getRange('J3').setValue('CICLO LECTIVO: ' + ciclo);
  salida.getRange('E4').setValue(registro.nombreCompleto);
  salida.getRange('E5').setValue(registro.dni);
  salida.getRange('E6').setValue(registro.nivelVisual);
  if (n > base) {
    salida.insertRowsAfter(inicio + base - 1, n - base);
    salida.getRange(inicio + base - 1, 1, 1, 12).copyTo(salida.getRange(inicio + base, 1, n - base, 12), { formatOnly: true });
  } else if (n < base) {
    salida.deleteRows(inicio + n, base - n);
  }
  const delta = n - base;
  salida.getRange(inicio, 1, n, 12).clearContent();
  if (registro.materias.length === 0) {
    salida.getRange(inicio, 3).setValue('Sin materias con valoración válida para incluir');
    advertencias.push('No se detectaron materias con valoración válida para incluir.');
  } else {
    registro.materias.forEach((m, i) => {
      const r = inicio + i;
      salida.getRange(r, 3).setValue(m.nombre);
      salida.getRange(r, 5).setValue(m.valor);
      salida.getRange(r, 6, 1, 7).clearContent();
    });
  }
  escribirCajaTexto_(salida, CONFIG_EPS.FILA_OBS_GENERALES_BASE + delta, registro.observacionesGenerales);
  escribirCajaTexto_(salida, CONFIG_EPS.FILA_OBS_ASISTENCIA_BASE + delta, registro.observacionesAsistencia);
  escribirCajaTexto_(salida, CONFIG_EPS.FILA_OBS_CONVIVENCIA_BASE + delta, registro.observacionesConvivencia);
  escribirCajaTexto_(salida, CONFIG_EPS.FILA_PTF_BASE + delta, registro.ptf);
}

function escribirCajaTexto_(hoja, fila, valor) { hoja.getRange(fila, 4).setValue(texto_(valor)).setWrap(true); hoja.setRowHeight(fila, 42); }
function detectarNivelPorFila_(valores, fila) { for (let r = fila; r >= 0; r--) { const p = normalizarTexto_(valores[r][0]); if (/^EPS\d+$/.test(p)) return p; } return ''; }
function construirNivelVisual_(hojaOrigen, nivel, orientacion) { const numeroNivel = nivel ? nivel.replace(/\D/g, '') : ''; const codigoOrigen = normalizarTexto_(hojaOrigen).replace(/\s+/g, ''); const nombresPorHoja = { EPSI: 'Informática', EPSG: 'Gastronomía' }; const nombreCurso = nombresPorHoja[codigoOrigen] || texto_(orientacion) || texto_(hojaOrigen); return numeroNivel ? nombreCurso + ' ' + numeroNivel : nombreCurso; }
function normalizarHojaOrigen_(valor) { const hoja = normalizarTexto_(valor).replace(/\s+/g, ''); return ['EPSI', 'EPSG'].includes(hoja) ? hoja : ''; }
function extraerSpreadsheetId_(entrada) { const t = texto_(entrada); const u = t.match(/\/d\/([a-zA-Z0-9-_]+)/); if (u) return u[1]; if (/^[a-zA-Z0-9-_]{20,}$/.test(t)) return t; throw new Error('No pude reconocer el ID de la planilla fuente.'); }
function obtenerValor_(valores, fila, columna) { if (columna === undefined || columna === null || columna < 0) return ''; return texto_((valores[fila] || [])[columna]); }
function claveMateria_(nombre) { return normalizarTexto_(nombre).replace(/[^A-Z0-9]/g, ''); }
function normalizarDni_(valor) { return texto_(valor).replace(/\D/g, ''); }
function normalizarTexto_(valor) { return texto_(valor).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/\s+/g, ' ').trim(); }
function texto_(valor) { if (valor === null || valor === undefined) return ''; return String(valor).trim(); }
function registrarEstado_(control, resultado, advertencias) { try { control.getRange(CONFIG_EPS.ESTADO_FECHA).setValue(new Date()); control.getRange(CONFIG_EPS.ESTADO_RESULTADO).setValue(resultado); control.getRange(CONFIG_EPS.ESTADO_ADVERTENCIAS).setValue(advertencias.length ? advertencias.join('\n') : 'Sin advertencias').setWrap(true); } catch (err) {} }
