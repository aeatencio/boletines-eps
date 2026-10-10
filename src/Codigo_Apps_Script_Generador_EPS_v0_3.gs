/**
 * Generador de informes individuales EPS v0.3
 * 1º bimestre + 1º cuatrimestre + 3º bimestre · generación individual
 *
 * La app no modifica los archivos fuente ni la hoja Plantilla.
 * La salida se genera en una hoja revisable llamada Informe_actual.
 */
const CONFIG_EPS = {
  HOJA_CONTROL: 'Control',
  HOJA_PLANTILLA: 'Plantilla',
  HOJA_SALIDA: 'Informe_actual',

  CELDA_PERIODO: 'B5',
  CELDA_HOJA_ORIGEN: 'B6',
  CELDA_DNI: 'B7',
  CELDA_CICLO: 'B8',
  CELDA_FUENTE_1B: 'B10',
  CELDA_FUENTE_1C: 'B12',
  CELDA_FUENTE_3B: 'B13',

  ESTADO_FECHA: 'F2',
  ESTADO_RESULTADO: 'F3',
  ESTADO_ADVERTENCIAS: 'F4',

  FILA_INICIO_MATERIAS: 9,
  FILAS_MATERIAS_PLANTILLA: 8,
  FILA_OBS_GENERALES_BASE: 20,
  FILA_OBS_ASISTENCIA_BASE: 23,
  FILA_OBS_CONVIVENCIA_BASE: 26,
  FILA_PTF_BASE: 29,

  PERIODOS: {
    BIMESTRE_1: '1º bimestre',
    CUATRIMESTRE_1: '1º cuatrimestre',
    BIMESTRE_3: '3º bimestre'
  },

  COLUMNAS_INFORME: {
    '1º bimestre': 5,      // E
    '1º cuatrimestre': 6,  // F
    '3º bimestre': 7       // G
  }
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

  try {
    if (!control) throw new Error('No se encontró la hoja Control.');

    const config = leerConfiguracion_(control);
    const advertencias = [];

    let informe;
    if (config.periodo === CONFIG_EPS.PERIODOS.BIMESTRE_1) {
      informe = generarDatosBimestre_(config, advertencias);
    } else if (config.periodo === CONFIG_EPS.PERIODOS.CUATRIMESTRE_1) {
      informe = generarDatosCuatrimestre_(config, advertencias);
    } else if (config.periodo === CONFIG_EPS.PERIODOS.BIMESTRE_3) {
      informe = generarDatosBimestre3_(config, advertencias);
    } else {
      throw new Error('Período no implementado: ' + config.periodo);
    }

    crearInforme_(ss, informe.estudiante, informe.materias, config.periodo, config.cicloLectivo, advertencias);
    registrarEstado_(control, 'Generado: ' + informe.estudiante.nombreCompleto, advertencias);
    ui.alert('Informe generado correctamente.' + (advertencias.length ? '\n\nAdvertencias:\n- ' + advertencias.join('\n- ') : ''));
  } catch (err) {
    if (control) registrarEstado_(control, 'Error: ' + err.message, []);
    ui.alert('No se pudo generar el informe.\n\n' + err.message);
    throw err;
  }
}

function leerConfiguracion_(control) {
  const periodo = normalizarPeriodo_(control.getRange(CONFIG_EPS.CELDA_PERIODO).getDisplayValue());
  const hojaOrigen = normalizarHojaOrigen_(control.getRange(CONFIG_EPS.CELDA_HOJA_ORIGEN).getDisplayValue());
  const dni = normalizarDni_(control.getRange(CONFIG_EPS.CELDA_DNI).getDisplayValue());
  const cicloLectivo = texto_(control.getRange(CONFIG_EPS.CELDA_CICLO).getDisplayValue()) || '2026';
  const fuente1B = texto_(control.getRange(CONFIG_EPS.CELDA_FUENTE_1B).getDisplayValue());
  const fuente1C = texto_(control.getRange(CONFIG_EPS.CELDA_FUENTE_1C).getDisplayValue());
  const fuente3B = texto_(control.getRange(CONFIG_EPS.CELDA_FUENTE_3B).getDisplayValue());

  if (!periodo) throw new Error('Falta indicar el período a generar.');
  if (![CONFIG_EPS.PERIODOS.BIMESTRE_1, CONFIG_EPS.PERIODOS.CUATRIMESTRE_1, CONFIG_EPS.PERIODOS.BIMESTRE_3].includes(periodo)) {
    throw new Error('El período indicado no está implementado en esta versión: ' + periodo);
  }
  if (!hojaOrigen) throw new Error('La hoja origen debe ser EPSI o EPSG.');
  if (!dni) throw new Error('Falta completar el DNI del estudiante.');
  if (periodo === CONFIG_EPS.PERIODOS.BIMESTRE_1 && !fuente1B) {
    throw new Error('Falta la URL o ID del archivo de valoraciones del 1º bimestre.');
  }
  if (periodo === CONFIG_EPS.PERIODOS.CUATRIMESTRE_1 && !fuente1C) {
    throw new Error('Falta la URL o ID del archivo de valoraciones del 1º cuatrimestre.');
  }
  if (periodo === CONFIG_EPS.PERIODOS.BIMESTRE_3 && !fuente3B) {
    throw new Error('Falta la URL o ID del archivo de valoraciones del 3º bimestre.');
  }

  return { periodo, hojaOrigen, dni, cicloLectivo, fuente1B, fuente1C, fuente3B };
}

function generarDatosBimestre_(config, advertencias) {
  const fuente1B = leerFuente_(config.fuente1B, config.hojaOrigen, config.dni, CONFIG_EPS.PERIODOS.BIMESTRE_1, true, advertencias);
  return {
    estudiante: fuente1B.estudiante,
    materias: construirMateriasAcumuladas_([fuente1B], [CONFIG_EPS.PERIODOS.BIMESTRE_1], advertencias)
  };
}

function generarDatosCuatrimestre_(config, advertencias) {
  const fuente1C = leerFuente_(config.fuente1C, config.hojaOrigen, config.dni, CONFIG_EPS.PERIODOS.CUATRIMESTRE_1, true, advertencias);
  let fuentes = [fuente1C];
  let ordenPeriodos = [CONFIG_EPS.PERIODOS.CUATRIMESTRE_1];

  if (config.fuente1B) {
    const fuente1B = leerFuente_(config.fuente1B, config.hojaOrigen, config.dni, CONFIG_EPS.PERIODOS.BIMESTRE_1, false, advertencias);
    if (fuente1B) {
      fuentes = [fuente1C, fuente1B];
      ordenPeriodos = [CONFIG_EPS.PERIODOS.CUATRIMESTRE_1, CONFIG_EPS.PERIODOS.BIMESTRE_1];
    }
  } else {
    advertencias.push('No se indicó archivo de 1º bimestre. El informe se generó sin valores anteriores.');
  }

  return {
    estudiante: fuente1C.estudiante,
    materias: construirMateriasAcumuladas_(fuentes, ordenPeriodos, advertencias)
  };
}

function generarDatosBimestre3_(config, advertencias) {
  const fuente3B = leerFuente_(config.fuente3B, config.hojaOrigen, config.dni, CONFIG_EPS.PERIODOS.BIMESTRE_3, true, advertencias);
  let fuentes = [fuente3B];
  let ordenPeriodos = [CONFIG_EPS.PERIODOS.BIMESTRE_3];

  if (config.fuente1C) {
    const fuente1C = leerFuente_(config.fuente1C, config.hojaOrigen, config.dni, CONFIG_EPS.PERIODOS.CUATRIMESTRE_1, false, advertencias);
    if (fuente1C) {
      fuentes.push(fuente1C);
      ordenPeriodos.push(CONFIG_EPS.PERIODOS.CUATRIMESTRE_1);
    }
  } else {
    advertencias.push('No se indicó archivo de 1º cuatrimestre. El informe se generó sin valores anteriores de 1º cuatrimestre.');
  }

  if (config.fuente1B) {
    const fuente1B = leerFuente_(config.fuente1B, config.hojaOrigen, config.dni, CONFIG_EPS.PERIODOS.BIMESTRE_1, false, advertencias);
    if (fuente1B) {
      fuentes.push(fuente1B);
      ordenPeriodos.push(CONFIG_EPS.PERIODOS.BIMESTRE_1);
    }
  } else {
    advertencias.push('No se indicó archivo de 1º bimestre. El informe se generó sin valores anteriores de 1º bimestre.');
  }

  return {
    estudiante: fuente3B.estudiante,
    materias: construirMateriasAcumuladas_(fuentes, ordenPeriodos, advertencias)
  };
}

function leerFuente_(fuenteIngresada, hojaOrigen, dniBuscado, periodo, esPrincipal, advertencias) {
  try {
    const idFuente = extraerSpreadsheetId_(fuenteIngresada);
    const libroFuente = SpreadsheetApp.openById(idFuente);
    const hoja = libroFuente.getSheetByName(hojaOrigen);
    if (!hoja) throw new Error('No se encontró la hoja ' + hojaOrigen + ' en la fuente de ' + periodo + '.');

    const valores = hoja.getDataRange().getDisplayValues();
    const estructura = detectarEstructuraFuente_(valores, periodo);
    const filas = buscarFilasPorDni_(valores, dniBuscado);

    if (filas.length === 0) throw new Error('No se encontró el DNI en la fuente de ' + periodo + '.');
    if (filas.length > 1) throw new Error('El DNI aparece más de una vez en la fuente de ' + periodo + '.');

    const estudiante = construirRegistroEstudiante_(valores, filas[0], estructura, hojaOrigen, periodo, advertencias);
    return {
      periodo,
      estudiante,
      materiasPorClave: estudiante.materiasPorClave,
      ordenMaterias: estudiante.ordenMaterias,
      ordenGrilla: estructura.materias.map(m => m.clave),
      nombresGrilla: Object.fromEntries(estructura.materias.map(m => [m.clave, m.nombre]))
    };
  } catch (err) {
    if (esPrincipal) throw err;
    advertencias.push(err.message + ' Se continuará con la fuente principal.');
    return null;
  }
}

/**
 * La columna ORIENTACIÓN separa las dos zonas de la grilla:
 * - materias: desde la columna D hasta antes de ORIENTACIÓN, como pares etiqueta/valoración;
 * - campos generales y operativos: desde ORIENTACIÓN en adelante, nunca materias.
 * Si la estructura no puede identificarse de manera segura, se informa con un error.
 */
function detectarEstructuraFuente_(valores, periodo) {
  const encabezado1 = valores[0] || [];
  const encabezado2 = valores[1] || [];
  const maxCols = Math.max(encabezado1.length, encabezado2.length);
  const enFuente = ' en la fuente de ' + periodo + '.';

  const columnasOrientacion = [];
  for (let c = 3; c < maxCols; c++) {
    if (normalizarTexto_(encabezado1[c]) === 'ORIENTACION' || normalizarTexto_(encabezado2[c]) === 'ORIENTACION') {
      columnasOrientacion.push(c);
    }
  }
  if (columnasOrientacion.length === 0) {
    throw new Error('No se encontró la columna ORIENTACIÓN que separa las materias de los campos generales' + enFuente);
  }
  if (columnasOrientacion.length > 1) {
    throw new Error('La columna ORIENTACIÓN aparece más de una vez (columnas ' + columnasOrientacion.map(letraColumna_).join(', ') + ')' + enFuente);
  }
  const inicioCola = columnasOrientacion[0];

  const especiales = { orientacion: inicioCola };
  for (let c = inicioCola + 1; c < maxCols; c++) {
    const combinado = normalizarTexto_([texto_(encabezado1[c]), texto_(encabezado2[c])].filter(Boolean).join(' '));
    if (combinado.includes('OBSERVACIONES GENERALES')) especiales.observacionesGenerales = c;
    if (combinado.includes('ASISTENCIA')) especiales.observacionesAsistencia = c;
    if (combinado.includes('CONVIVENCIA')) especiales.observacionesConvivencia = c;
    if (combinado.includes('PLAN DE TRABAJO') || combinado.includes('PTF')) especiales.ptf = c;
  }

  // Formación Profesional: el encabezado de grupo está en fila 1 y cada materia en fila 2.
  const materias = [];
  const columnasPorClave = {};
  for (let c = 3; c < inicioCola; c++) {
    const nombre = texto_(encabezado2[c]) || texto_(encabezado1[c]);
    if (!nombre) continue;
    if (/^EPS\d+$/.test(normalizarTexto_(nombre))) continue;
    if (c + 1 >= inicioCola || texto_(encabezado1[c + 1]) || texto_(encabezado2[c + 1])) {
      throw new Error('La materia "' + nombre + '" (columna ' + letraColumna_(c) + ') no tiene a continuación su columna de valoración' + enFuente);
    }
    const clave = claveMateria_(nombre);
    if (!columnasPorClave[clave]) columnasPorClave[clave] = [];
    columnasPorClave[clave].push(c);
    materias.push({ nombre, clave, columnaEtiqueta: c, columnaValor: c + 1 });
  }

  const duplicadas = Object.keys(columnasPorClave).filter(clave => columnasPorClave[clave].length > 1);
  if (duplicadas.length) {
    const detalle = duplicadas.map(clave => {
      const nombre = materias.find(m => m.clave === clave).nombre;
      return '"' + nombre + '" (columnas ' + columnasPorClave[clave].map(letraColumna_).join(', ') + ')';
    });
    throw new Error('Materia repetida en la fuente de ' + periodo + ': ' + detalle.join('; ') + '. No se puede determinar cuál columna contiene la valoración correcta.');
  }

  return { materias, especiales, inicioCola };
}

function buscarFilasPorDni_(valores, dniBuscado) {
  const filas = [];
  for (let r = 0; r < valores.length; r++) {
    if (normalizarDni_(valores[r][2]) === dniBuscado) filas.push(r);
  }
  return filas;
}

function construirRegistroEstudiante_(valores, fila, estructura, hojaOrigen, periodo, advertencias) {
  const row = valores[fila] || [];
  const apellido = texto_(row[0]);
  const nombres = texto_(row[1]);
  const dni = normalizarDni_(row[2]);
  const nivel = detectarNivelPorFila_(valores, fila);
  const orientacion = obtenerValor_(valores, fila, estructura.especiales.orientacion);

  // Cada estudiante ocupa un bloque de cuatro filas; la del DNI debe ser la fila "Valoración".
  // Se controla por fila y no por celda: en las grillas reales algunas etiquetas quedan vacías.
  const esFilaValoracion = estructura.materias.some(m =>
    normalizarTexto_(obtenerValor_(valores, fila, m.columnaEtiqueta)).replace(/:$/, '') === 'VALORACION');
  if (!esFilaValoracion) {
    throw new Error('La fila del DNI no es la fila "Valoración" del bloque del estudiante en la fuente de ' + periodo + '.');
  }

  const materiasPorClave = {};
  const ordenMaterias = [];

  estructura.materias.forEach(m => {
    const valorCrudo = obtenerValor_(valores, fila, m.columnaValor);
    if (!valorCrudo) return;
    const clasificacion = clasificarValor_(periodo, valorCrudo);
    if (clasificacion.estado === 'valido') {
      materiasPorClave[m.clave] = { nombre: m.nombre, clave: m.clave, valores: { [periodo]: clasificacion.valor } };
      ordenMaterias.push(m.clave);
    } else if (clasificacion.estado === 'invalido') {
      advertencias.push('Valor inesperado en ' + periodo + ' para "' + m.nombre + '": "' + valorCrudo + '".');
    }
  });

  return {
    apellido,
    nombres,
    dni,
    nombreCompleto: [apellido, nombres].filter(Boolean).join(', ') || '(sin nombre en la fuente)',
    nivelVisual: construirNivelVisual_(hojaOrigen, nivel, orientacion),
    observacionesGenerales: obtenerValor_(valores, fila, estructura.especiales.observacionesGenerales),
    observacionesAsistencia: obtenerValor_(valores, fila, estructura.especiales.observacionesAsistencia),
    observacionesConvivencia: obtenerValor_(valores, fila, estructura.especiales.observacionesConvivencia),
    ptf: obtenerValor_(valores, fila, estructura.especiales.ptf),
    materiasPorClave,
    ordenMaterias
  };
}

function construirMateriasAcumuladas_(fuentes, ordenPeriodos, advertencias) {
  const acumuladas = {};

  fuentes.forEach(fuente => {
    if (!fuente) return;
    fuente.ordenMaterias.forEach(clave => {
      const mat = fuente.materiasPorClave[clave];
      if (!mat) return;
      if (!acumuladas[clave]) acumuladas[clave] = { nombre: mat.nombre, valores: {} };
      Object.assign(acumuladas[clave].valores, mat.valores);
    });
  });

  // 1. Orden y nombre de la grilla principal, aunque la valoración venga de una fuente anterior.
  const principal = fuentes[0];
  const orden = principal.ordenGrilla.filter(clave => acumuladas[clave]);
  orden.forEach(clave => { acumuladas[clave].nombre = principal.nombresGrilla[clave]; });

  // 2. Al final, con el nombre de su fuente, materias que no existen en la grilla principal.
  for (let i = 1; i < fuentes.length; i++) {
    const fuente = fuentes[i];
    if (!fuente) continue;
    fuente.ordenMaterias.forEach(clave => {
      if (!orden.includes(clave)) orden.push(clave);
    });
  }

  return orden.map(clave => acumuladas[clave]);
}

function crearInforme_(ss, estudiante, materias, periodoGenerado, cicloLectivo, advertencias) {
  const plantilla = ss.getSheetByName(CONFIG_EPS.HOJA_PLANTILLA);
  if (!plantilla) throw new Error('No se encontró la hoja Plantilla.');

  const salidaExistente = ss.getSheetByName(CONFIG_EPS.HOJA_SALIDA);
  if (salidaExistente) ss.deleteSheet(salidaExistente);

  const salida = plantilla.copyTo(ss);
  salida.setName(CONFIG_EPS.HOJA_SALIDA);
  ss.setActiveSheet(salida);

  salida.getRange('A1:A3').clearContent();
  salida.getRange('J3').setValue('CICLO LECTIVO: ' + cicloLectivo);
  salida.getRange('E4').setValue(estudiante.nombreCompleto);
  salida.getRange('E5').setValue(estudiante.dni);
  salida.getRange('E6').setValue(estudiante.nivelVisual);

  const filasNecesarias = Math.max(materias.length, 1);
  const filasBase = CONFIG_EPS.FILAS_MATERIAS_PLANTILLA;
  const filaInicio = CONFIG_EPS.FILA_INICIO_MATERIAS;

  if (filasNecesarias > filasBase) {
    salida.insertRowsAfter(filaInicio + filasBase - 1, filasNecesarias - filasBase);
    const origenFormato = salida.getRange(filaInicio + filasBase - 1, 1, 1, 12);
    const destinoFormato = salida.getRange(filaInicio + filasBase, 1, filasNecesarias - filasBase, 12);
    origenFormato.copyTo(destinoFormato, { formatOnly: true });
  } else if (filasNecesarias < filasBase) {
    salida.deleteRows(filaInicio + filasNecesarias, filasBase - filasNecesarias);
  }

  const delta = filasNecesarias - filasBase;
  const rangoMaterias = salida.getRange(filaInicio, 1, filasNecesarias, 12);
  rangoMaterias.clearContent();

  if (materias.length === 0) {
    salida.getRange(filaInicio, 3).setValue('Sin materias con calificación válida para incluir');
    advertencias.push('No se detectaron materias con calificación válida para incluir.');
  } else {
    materias.forEach((mat, idx) => {
      const row = filaInicio + idx;
      salida.getRange(row, 3).setValue(mat.nombre);
      Object.keys(CONFIG_EPS.COLUMNAS_INFORME).forEach(periodoKey => {
        const col = CONFIG_EPS.COLUMNAS_INFORME[periodoKey];
        if (mat.valores[periodoKey]) {
          salida.getRange(row, col).setValue(mat.valores[periodoKey]);
        }
      });
    });
  }

  escribirCajaTexto_(salida, CONFIG_EPS.FILA_OBS_GENERALES_BASE + delta, estudiante.observacionesGenerales);
  escribirCajaTexto_(salida, CONFIG_EPS.FILA_OBS_ASISTENCIA_BASE + delta, estudiante.observacionesAsistencia);
  escribirCajaTexto_(salida, CONFIG_EPS.FILA_OBS_CONVIVENCIA_BASE + delta, estudiante.observacionesConvivencia);
  escribirCajaTexto_(salida, CONFIG_EPS.FILA_PTF_BASE + delta, estudiante.ptf);
}

function escribirCajaTexto_(hoja, fila, valor) {
  hoja.getRange(fila, 4).setValue(texto_(valor));
  hoja.getRange(fila, 4, 1, 7).setWrap(true);
  hoja.setRowHeight(fila, 42);
}

function clasificarValor_(periodo, valor) {
  const crudo = texto_(valor);
  const norm = normalizarTexto_(crudo);
  if (!norm || norm === 'NO CURSA' || norm === 'CURSA') return { estado: 'excluido', valor: '' };
  if (norm === 'EQUIVALENCIA') return { estado: 'valido', valor: 'EQUIVALENCIA' };

  if (periodo === CONFIG_EPS.PERIODOS.BIMESTRE_1 || periodo === CONFIG_EPS.PERIODOS.BIMESTRE_3) {
    const mapa = {
      'EN PROCESO': 'En proceso',
      'SUFICIENTE': 'Suficiente',
      'AVANZADO': 'Avanzado',
      'ACREDITO': 'ACREDITÓ'
    };
    if (mapa[norm]) return { estado: 'valido', valor: mapa[norm] };
    return { estado: 'invalido', valor: crudo };
  }

  if (periodo === CONFIG_EPS.PERIODOS.CUATRIMESTRE_1) {
    if (norm === 'ACREDITO') return { estado: 'valido', valor: 'ACREDITÓ' };
    if (/^([1-9]|10)$/.test(crudo)) return { estado: 'valido', valor: crudo };
    return { estado: 'invalido', valor: crudo };
  }

  return { estado: 'invalido', valor: crudo };
}

function detectarNivelPorFila_(valores, fila) {
  for (let r = fila; r >= 0; r--) {
    const posible = normalizarTexto_(valores[r][0]);
    if (/^EPS\d+$/.test(posible)) return posible;
  }
  return '';
}

function construirNivelVisual_(hojaOrigen, nivel, orientacion) {
  const numeroNivel = nivel ? nivel.replace(/\D/g, '') : '';
  const codigoOrigen = normalizarTexto_(hojaOrigen).replace(/\s+/g, '');
  const nombresPorHoja = { EPSI: 'Informática', EPSG: 'Gastronomía' };
  const nombreCurso = nombresPorHoja[codigoOrigen] || texto_(orientacion) || texto_(hojaOrigen);
  return numeroNivel ? nombreCurso + ' ' + numeroNivel : nombreCurso;
}

function normalizarPeriodo_(valor) {
  const crudo = texto_(valor);
  const t = crudo.toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[º°]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (['1 bimestre', '1er bimestre', 'primer bimestre'].includes(t)) return CONFIG_EPS.PERIODOS.BIMESTRE_1;
  if (['1 cuatrimestre', '1er cuatrimestre', 'primer cuatrimestre'].includes(t)) return CONFIG_EPS.PERIODOS.CUATRIMESTRE_1;
  if (['3 bimestre', '3er bimestre', 'tercer bimestre'].includes(t)) return CONFIG_EPS.PERIODOS.BIMESTRE_3;
  return crudo;
}

function normalizarHojaOrigen_(valor) {
  const hoja = normalizarTexto_(valor).replace(/\s+/g, '');
  return ['EPSI', 'EPSG'].includes(hoja) ? hoja : '';
}

function extraerSpreadsheetId_(entrada) {
  const texto = texto_(entrada);
  const matchUrl = texto.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (matchUrl) return matchUrl[1];
  const matchId = texto.match(/^[a-zA-Z0-9-_]{20,}$/);
  if (matchId) return texto;
  throw new Error('No pude reconocer el ID de la planilla fuente. Pegá el link completo de Google Sheets o el ID del archivo.');
}

function obtenerValor_(valores, fila, columna) {
  if (columna === undefined || columna === null || columna < 0) return '';
  return texto_((valores[fila] || [])[columna]);
}

function letraColumna_(indice) {
  let letras = '';
  for (let n = indice + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    letras = String.fromCharCode(65 + (n - 1) % 26) + letras;
  }
  return letras;
}

function claveMateria_(nombre) {
  return normalizarTexto_(nombre).replace(/[^A-Z0-9]/g, '');
}

function normalizarDni_(valor) {
  return texto_(valor).replace(/\D/g, '');
}

function normalizarTexto_(valor) {
  return texto_(valor)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function texto_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}

function registrarEstado_(control, resultado, advertencias) {
  try {
    control.getRange(CONFIG_EPS.ESTADO_FECHA).setValue(new Date());
    control.getRange(CONFIG_EPS.ESTADO_RESULTADO).setValue(resultado);
    control.getRange(CONFIG_EPS.ESTADO_ADVERTENCIAS).setValue(advertencias.length ? advertencias.join('\n') : 'Sin advertencias');
    control.getRange(CONFIG_EPS.ESTADO_ADVERTENCIAS).setWrap(true);
  } catch (err) {
    // El registro de estado no debe impedir la generación del informe.
  }
}
