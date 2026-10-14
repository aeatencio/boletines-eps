/**
 * Generador de informes individuales EPS v0.3
 * 1º bimestre + 1º cuatrimestre + 3º bimestre · generación individual y por curso
 *
 * La app no modifica los archivos fuente ni la hoja Plantilla.
 * La salida individual se genera en una hoja revisable llamada Informe_actual;
 * la generación por curso crea un libro nuevo con una hoja por estudiante y una hoja Resumen.
 */
const CONFIG_EPS = {
  HOJA_CONTROL: 'Control',
  HOJA_PLANTILLA: 'Plantilla',
  HOJA_SALIDA: 'Informe_actual',
  HOJA_RESUMEN: 'Resumen',

  CELDA_PERIODO: 'B5',
  CELDA_HOJA_ORIGEN: 'B6',
  CELDA_DNI: 'B7',
  CELDA_CICLO: 'B8',
  CELDA_NIVEL: 'B9',
  CELDA_FUENTE_1B: 'B10',
  CELDA_FUENTE_1C: 'B11',
  CELDA_FUENTE_3B: 'B12',

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
  },

  // Encabezados que identifican cada campo de la cola (después de ORIENTACIÓN).
  // Se comparan normalizados (sin tildes ni diferencias de mayúsculas o espacios), nunca por posición.
  ENCABEZADOS_GENERALES: {
    observacionesGenerales: 'OBSERVACIONES GENERALES SOBRE DESEMPEÑO ACADÉMICO',
    observacionesAsistencia: 'OBSERVACIONES SOBRE ASISTENCIA',
    observacionesConvivencia: 'OBSERVACIONES SOBRE CONVIVENCIA ESCOLAR',
    ptf: 'PLAN DE TRABAJO FORMATIVO (PTF)'
  },
  ENCABEZADOS_OPERATIVOS: {
    listoParaExportar: 'LISTO PARA EXPORTAR',
    exportado: 'EXPORTADO'
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Informes EPS')
    .addItem('Generar informe actual', 'generarInformeActual')
    .addItem('Generar boletines por curso', 'generarBoletinesPorCurso')
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
    const informe = generarDatosInforme_(config, advertencias);

    crearInforme_(ss, informe.estudiante, informe.materias, config.periodo, config.cicloLectivo, advertencias);
    registrarEstado_(control, 'Generado: ' + informe.estudiante.nombreCompleto, advertencias);
    ui.alert('Informe generado correctamente.' + (advertencias.length ? '\n\nAdvertencias:\n- ' + advertencias.join('\n- ') : ''));
  } catch (err) {
    if (control) registrarEstado_(control, 'Error: ' + err.message, []);
    ui.alert('No se pudo generar el informe.\n\n' + err.message);
    throw err;
  }
}

/**
 * Generación por curso: un libro nuevo con una hoja por estudiante del nivel indicado
 * cuya fila en la fuente principal tiene LISTO PARA EXPORTAR = SÍ.
 * Cada fuente se lee una sola vez; las fuentes no se modifican.
 */
function generarBoletinesPorCurso() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const control = ss.getSheetByName(CONFIG_EPS.HOJA_CONTROL);

  try {
    if (!control) throw new Error('No se encontró la hoja Control.');

    const config = leerConfiguracion_(control, true);
    const tanda = generarTandaCurso_(ss, config);
    const conteo = 'Generados: ' + tanda.contar('Generado') + ' · Omitidos: ' + tanda.contar('Omitido') + ' · Con error: ' + tanda.contar('Error') + '.';

    if (!tanda.libro) {
      // Sin libro no hay hoja Resumen: el detalle por estudiante queda en Control y en el diálogo.
      const detalle = tanda.advertencias.concat(tanda.estudiantes.map(e => e.estado + ' · fila ' + (e.fila + 1) + ' · ' + e.nombre + ': ' + e.detalle));
      registrarEstado_(control, 'Sin libro: ' + tanda.motivo + ' ' + conteo, detalle);
      ui.alert('No se creó ningún libro.\n\n' + tanda.motivo + '\n' + conteo + (detalle.length ? '\n\n- ' + detalle.join('\n- ') : ''));
      return;
    }

    registrarEstado_(control, 'Libro creado: ' + tanda.libro.getName() + ' · ' + conteo + ' ' + tanda.libro.getUrl(), tanda.advertencias);
    ui.alert('Libro creado: ' + tanda.libro.getName() + '\n' + tanda.libro.getUrl() + '\n\n' + conteo +
      '\nEl detalle por estudiante está en la hoja ' + CONFIG_EPS.HOJA_RESUMEN + ' del libro' +
      (tanda.contar('Generado') ? ', oculta para que no se imprima ni se exporte.' : '.') +
      (tanda.advertencias.length ? '\n\nAdvertencias generales:\n- ' + tanda.advertencias.join('\n- ') : ''));
  } catch (err) {
    if (control) registrarEstado_(control, 'Error: ' + err.message, []);
    ui.alert('No se pudo generar el libro del curso.\n\n' + err.message);
    throw err;
  }
}

/**
 * Errores globales (configuración, estructura de la fuente principal, falta de LISTO PARA EXPORTAR)
 * interrumpen la tanda antes de crear el libro. El error de un estudiante sólo afecta su hoja.
 * Si ningún estudiante puede generarse, no se crea el libro.
 */
function generarTandaCurso_(ss, config) {
  const plantilla = ss.getSheetByName(CONFIG_EPS.HOJA_PLANTILLA);
  if (!plantilla) throw new Error('No se encontró la hoja Plantilla.');

  const advertencias = [];
  const fuentes = leerFuentes_(config, advertencias);
  const estudiantes = seleccionarEstudiantesCurso_(fuentes.principal, config.nivel);
  const tanda = {
    libro: null,
    motivo: '',
    advertencias,
    estudiantes,
    contar: estado => estudiantes.filter(e => e.estado === estado).length
  };

  estudiantes.filter(e => e.estado === 'Seleccionado').forEach(e => {
    try {
      e.informe = construirDatosEstudiante_(fuentes, config.hojaOrigen, e.dni, e.advertencias);
    } catch (err) {
      e.estado = 'Error';
      e.detalle = err.message;
    }
  });

  const preparados = estudiantes.filter(e => e.informe);
  if (!preparados.length) {
    const curso = config.hojaOrigen + ' ' + config.nivel + ' (' + config.periodo + ')';
    tanda.motivo = tanda.contar('Error')
      ? 'Ningún estudiante seleccionado de ' + curso + ' pudo prepararse.'
      : 'No hay estudiantes con LISTO PARA EXPORTAR = SÍ en ' + curso + '.';
    return tanda;
  }

  const libro = SpreadsheetApp.create(nombreLibroCurso_(ss, config));
  const resumen = libro.getSheets()[0];
  resumen.setName(CONFIG_EPS.HOJA_RESUMEN);
  asignarNombresHojas_(preparados);

  preparados.forEach(e => {
    let hoja = null;
    try {
      hoja = plantilla.copyTo(libro);
      hoja.setName(e.hoja);
      completarInforme_(hoja, e.informe.estudiante, e.informe.materias, config.cicloLectivo, e.advertencias);
      e.estado = 'Generado';
    } catch (err) {
      if (hoja) {
        try { libro.deleteSheet(hoja); } catch (errBorrado) { /* la hoja incompleta queda señalada en el Resumen */ }
      }
      e.estado = 'Error';
      e.detalle = 'No se pudo crear la hoja: ' + err.message;
    }
  });

  escribirResumen_(resumen, config, estudiantes, advertencias);
  // Oculta para que no se imprima ni se exporte con el libro completo.
  // Google Sheets exige una hoja visible: si no se generó ningún boletín, Resumen queda visible.
  if (tanda.contar('Generado')) resumen.hideSheet();
  tanda.libro = libro;
  return tanda;
}

/**
 * Estudiantes del nivel según la fuente principal, en el orden de la grilla.
 * Sólo LISTO PARA EXPORTAR = SÍ queda seleccionado (con la interpretación de leerEstadoOperativo_);
 * NO, vacío o inválido se omite. EXPORTADO se lee para informarlo, nunca para seleccionar.
 */
function seleccionarEstudiantesCurso_(principal, nivel) {
  const valores = principal.valores;
  const cola = principal.cola;
  const periodo = principal.periodo;
  const encabezados = CONFIG_EPS.ENCABEZADOS_OPERATIVOS;

  if (cola.listoParaExportar === undefined) {
    throw new Error('No se encontró una única columna "' + encabezados.listoParaExportar + '" después de ORIENTACIÓN en la hoja ' +
      principal.hojaOrigen + ' de la fuente de ' + periodo + '. Es la fuente principal del período a generar (' + periodo + ', elegido en ' +
      CONFIG_EPS.HOJA_CONTROL + '!' + CONFIG_EPS.CELDA_PERIODO + '): la generación por curso la necesita para seleccionar estudiantes y no usa otra columna en su lugar. ' +
      'Las fuentes anteriores no la necesitan.');
  }

  const estudiantes = [];
  let nivelFila = '';
  let nivelEncontrado = false;
  // Las dos primeras filas son encabezados; el nivel de cada fila es el último marcador EPSn de la columna A.
  for (let r = 2; r < valores.length; r++) {
    const marca = normalizarTexto_(valores[r][0]);
    if (/^EPS\d+$/.test(marca)) {
      nivelFila = marca;
      if (marca === nivel) nivelEncontrado = true;
    }
    if (nivelFila !== nivel) continue;

    const dni = normalizarDni_(valores[r][2]);
    const nombre = [texto_(valores[r][0]), texto_(valores[r][1])].filter(Boolean).join(', ') || '(sin nombre en la fuente)';
    const avisos = [];
    const listo = leerEstadoOperativo_(valores, r, cola.listoParaExportar, encabezados.listoParaExportar, periodo, avisos);
    const exportado = leerEstadoOperativo_(valores, r, cola.exportado, encabezados.exportado, periodo, []);
    const estudiante = { fila: r, dni, nombre, exportado: exportado || '', estado: '', detalle: '', hoja: '', advertencias: [] };

    if (!dni) {
      if (listo !== 'SÍ') continue;
      estudiante.estado = 'Error';
      estudiante.detalle = 'La fila tiene LISTO PARA EXPORTAR = SÍ pero no tiene DNI.';
    } else if (listo === 'SÍ') {
      // Las advertencias del estudiante, incluida una forma no canónica de SÍ, salen de su generación.
      estudiante.estado = 'Seleccionado';
    } else {
      estudiante.estado = 'Omitido';
      estudiante.detalle = listo === 'NO' ? 'LISTO PARA EXPORTAR = NO.' : avisos.join(' ');
    }
    estudiantes.push(estudiante);
  }

  if (!nivelEncontrado) {
    throw new Error('No se encontró el nivel ' + nivel + ' (marcador en la columna A) en la hoja ' + principal.hojaOrigen + ' de la fuente de ' + periodo + '.');
  }
  return estudiantes;
}

function nombreLibroCurso_(ss, config) {
  const fecha = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
  return ['Boletines EPS', config.hojaOrigen, config.nivel, config.periodo + ' ' + config.cicloLectivo, fecha].join(' - ');
}

/**
 * Nombre de hoja: "Apellido, Nombres". Si varios estudiantes producen el mismo nombre
 * (sin distinguir mayúsculas, como Google Sheets), todos llevan además su DNI.
 */
function asignarNombresHojas_(estudiantes) {
  const clave = nombre => nombre.toLowerCase();
  const bases = estudiantes.map(e => limpiarNombreHoja_(e.informe.estudiante.nombreCompleto) || 'Estudiante');
  const repeticiones = {};
  bases.forEach(base => { repeticiones[clave(base)] = (repeticiones[clave(base)] || 0) + 1; });

  const usados = new Set([clave(CONFIG_EPS.HOJA_RESUMEN)]);
  estudiantes.forEach((e, i) => {
    const base = repeticiones[clave(bases[i])] > 1 ? bases[i] + ' (' + e.dni + ')' : bases[i];
    let nombre = base;
    for (let n = 2; usados.has(clave(nombre)); n++) nombre = base + ' (' + n + ')';
    usados.add(clave(nombre));
    e.hoja = nombre;
  });
}

function limpiarNombreHoja_(nombre) {
  return texto_(nombre)
    .replace(/[\[\]*?\/\\:]/g, ' ')
    .replace(/^'+|'+$/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80)
    .trim();
}

function escribirResumen_(hoja, config, estudiantes, advertencias) {
  const columnas = 7;
  const filas = [
    ['Boletines EPS por curso'],
    ['Hoja origen', config.hojaOrigen],
    ['Nivel', config.nivel],
    ['Período', config.periodo],
    ['Ciclo lectivo', config.cicloLectivo],
    ['Fecha de generación', new Date()],
    ['Advertencias generales', advertencias.length ? advertencias.join('\n') : 'Sin advertencias'],
    [],
    ['Estado', 'Hoja', 'Estudiante', 'DNI', 'Fila en la fuente', 'EXPORTADO en la grilla', 'Motivo o advertencias']
  ].concat(estudiantes.map(e => [
    e.estado, e.hoja, e.nombre, e.dni, e.fila + 1, e.exportado,
    [e.detalle].concat(e.advertencias).filter(Boolean).join('\n')
  ])).map(fila => fila.concat(Array(columnas - fila.length).fill('')));

  hoja.getRange(1, 1, filas.length, columnas).setValues(filas);
  hoja.getRange(1, 1, filas.length, columnas).setWrap(true);
}

function leerConfiguracion_(control, porCurso) {
  const periodo = normalizarPeriodo_(control.getRange(CONFIG_EPS.CELDA_PERIODO).getDisplayValue());
  const hojaOrigen = normalizarHojaOrigen_(control.getRange(CONFIG_EPS.CELDA_HOJA_ORIGEN).getDisplayValue());
  const dni = porCurso ? '' : normalizarDni_(control.getRange(CONFIG_EPS.CELDA_DNI).getDisplayValue());
  const nivel = porCurso ? normalizarNivel_(control.getRange(CONFIG_EPS.CELDA_NIVEL).getDisplayValue()) : '';
  const cicloLectivo = texto_(control.getRange(CONFIG_EPS.CELDA_CICLO).getDisplayValue()) || '2026';
  const fuente1B = texto_(control.getRange(CONFIG_EPS.CELDA_FUENTE_1B).getDisplayValue());
  const fuente1C = texto_(control.getRange(CONFIG_EPS.CELDA_FUENTE_1C).getDisplayValue());
  const fuente3B = texto_(control.getRange(CONFIG_EPS.CELDA_FUENTE_3B).getDisplayValue());

  if (!periodo) throw new Error('Falta indicar el período a generar.');
  if (![CONFIG_EPS.PERIODOS.BIMESTRE_1, CONFIG_EPS.PERIODOS.CUATRIMESTRE_1, CONFIG_EPS.PERIODOS.BIMESTRE_3].includes(periodo)) {
    throw new Error('El período indicado no está implementado en esta versión: ' + periodo);
  }
  if (!hojaOrigen) throw new Error('La hoja origen debe ser EPSI o EPSG.');
  if (porCurso && !nivel) throw new Error('Falta indicar el nivel / curso en la celda ' + CONFIG_EPS.CELDA_NIVEL + ' de Control (EPS1, EPS2, EPS3 o EPS4).');
  if (!porCurso && !dni) throw new Error('Falta completar el DNI del estudiante.');
  if (periodo === CONFIG_EPS.PERIODOS.BIMESTRE_1 && !fuente1B) {
    throw new Error('Falta la URL o ID del archivo de valoraciones del 1º bimestre.');
  }
  if (periodo === CONFIG_EPS.PERIODOS.CUATRIMESTRE_1 && !fuente1C) {
    throw new Error('Falta la URL o ID del archivo de valoraciones del 1º cuatrimestre.');
  }
  if (periodo === CONFIG_EPS.PERIODOS.BIMESTRE_3 && !fuente3B) {
    throw new Error('Falta la URL o ID del archivo de valoraciones del 3º bimestre.');
  }

  return { periodo, hojaOrigen, dni, nivel, cicloLectivo, fuente1B, fuente1C, fuente3B };
}

/**
 * La generación individual y la generación por curso comparten las dos etapas:
 * leer cada fuente una vez (leerFuentes_) y construir los datos de un estudiante (construirDatosEstudiante_).
 */
function generarDatosInforme_(config, advertencias) {
  return construirDatosEstudiante_(leerFuentes_(config, advertencias), config.hojaOrigen, config.dni, advertencias);
}

/**
 * Cada generación tiene una única fuente principal: la del período solicitado.
 * De ella salen el estudiante, los campos generales (orientación, observaciones, PTF)
 * y los campos operativos. Las fuentes anteriores sólo aportan materias y valoraciones.
 * Una fuente anterior que no se indicó o no puede leerse se advierte y se omite.
 */
function leerFuentes_(config, advertencias) {
  const P = CONFIG_EPS.PERIODOS;
  const fuentePorPeriodo = { [P.BIMESTRE_1]: config.fuente1B, [P.CUATRIMESTRE_1]: config.fuente1C, [P.BIMESTRE_3]: config.fuente3B };
  // Fuentes anteriores de cada período, en orden de acumulación, con la advertencia si no se indicaron.
  const anterioresPorPeriodo = {
    [P.BIMESTRE_1]: [],
    [P.CUATRIMESTRE_1]: [
      [P.BIMESTRE_1, 'No se indicó archivo de 1º bimestre. El informe se generó sin valores anteriores.']
    ],
    [P.BIMESTRE_3]: [
      [P.CUATRIMESTRE_1, 'No se indicó archivo de 1º cuatrimestre. El informe se generó sin valores anteriores de 1º cuatrimestre.'],
      [P.BIMESTRE_1, 'No se indicó archivo de 1º bimestre. El informe se generó sin valores anteriores de 1º bimestre.']
    ]
  };
  if (!anterioresPorPeriodo[config.periodo]) throw new Error('Período no implementado: ' + config.periodo);

  const principal = abrirFuentePrincipal_(fuentePorPeriodo[config.periodo], config.hojaOrigen, config.periodo, advertencias);
  const anteriores = [];
  anterioresPorPeriodo[config.periodo].forEach(([periodo, advertenciaSinFuente]) => {
    if (!fuentePorPeriodo[periodo]) {
      advertencias.push(advertenciaSinFuente);
      return;
    }
    const fuente = abrirFuenteAnterior_(fuentePorPeriodo[periodo], config.hojaOrigen, periodo, advertencias);
    if (fuente) anteriores.push(fuente);
  });
  return { principal, anteriores };
}

function construirDatosEstudiante_(fuentes, hojaOrigen, dniBuscado, advertencias) {
  const principal = leerFuentePrincipal_(fuentes.principal, hojaOrigen, dniBuscado, advertencias);
  const leidas = [principal];
  fuentes.anteriores.forEach(fuente => {
    const anterior = leerFuenteHistorica_(fuente, dniBuscado, advertencias);
    if (anterior) leidas.push(anterior);
  });

  return {
    estudiante: principal.estudiante,
    camposOperativos: principal.camposOperativos,
    materias: construirMateriasAcumuladas_(leidas, leidas.map(f => f.periodo), advertencias)
  };
}

function errorFuentePrincipal_(err, periodo) {
  // Distingue este error del de una fuente anterior, que sólo genera una advertencia.
  return new Error(err.message + ' Es la fuente principal del período a generar (' + periodo + '), por eso no se puede continuar.');
}

/**
 * Fuente del período solicitado. Cualquier problema estructural interrumpe la generación.
 * Es la única lectura que interpreta la cola de la grilla (campos generales y operativos).
 */
function abrirFuentePrincipal_(fuenteIngresada, hojaOrigen, periodo, advertencias) {
  let fuente;
  try {
    fuente = abrirGrillaFuente_(fuenteIngresada, hojaOrigen, periodo);
  } catch (err) {
    throw errorFuentePrincipal_(err, periodo);
  }
  fuente.cola = detectarCamposCola_(fuente.valores, fuente.estructura, periodo, advertencias);
  return fuente;
}

/**
 * Fuente de un período anterior: sólo aporta materias y valoraciones.
 * No lee la cola de la grilla. Si no puede leerse, se advierte y se continúa con la principal.
 */
function abrirFuenteAnterior_(fuenteIngresada, hojaOrigen, periodo, advertencias) {
  try {
    return abrirGrillaFuente_(fuenteIngresada, hojaOrigen, periodo);
  } catch (err) {
    advertencias.push(err.message + ' Se continuará con la fuente principal.');
    return null;
  }
}

function leerFuentePrincipal_(fuente, hojaOrigen, dniBuscado, advertencias) {
  const periodo = fuente.periodo;
  let grilla;
  try {
    grilla = leerGrillaEstudiante_(fuente, dniBuscado, advertencias);
  } catch (err) {
    throw errorFuentePrincipal_(err, periodo);
  }
  const valores = grilla.valores;
  const fila = grilla.fila;
  const cola = fuente.cola;
  const row = valores[fila] || [];
  const apellido = texto_(row[0]);
  const nombres = texto_(row[1]);
  const orientacion = obtenerValor_(valores, fila, cola.orientacion);

  return {
    periodo,
    estudiante: {
      apellido,
      nombres,
      dni: normalizarDni_(row[2]),
      nombreCompleto: [apellido, nombres].filter(Boolean).join(', ') || '(sin nombre en la fuente)',
      orientacion,
      nivelVisual: construirNivelVisual_(hojaOrigen, detectarNivelPorFila_(valores, fila), orientacion),
      observacionesGenerales: obtenerValor_(valores, fila, cola.observacionesGenerales),
      observacionesAsistencia: obtenerValor_(valores, fila, cola.observacionesAsistencia),
      observacionesConvivencia: obtenerValor_(valores, fila, cola.observacionesConvivencia),
      ptf: obtenerValor_(valores, fila, cola.ptf)
    },
    materiasPorClave: grilla.materiasPorClave,
    ordenMaterias: grilla.ordenMaterias,
    ordenGrilla: grilla.ordenGrilla,
    nombresGrilla: grilla.nombresGrilla,
    // Se leen y validan; nunca se escriben en la fuente. La selección por curso los lee en seleccionarEstudiantesCurso_.
    camposOperativos: {
      listoParaExportar: leerEstadoOperativo_(valores, fila, cola.listoParaExportar, CONFIG_EPS.ENCABEZADOS_OPERATIVOS.listoParaExportar, periodo, advertencias),
      exportado: leerEstadoOperativo_(valores, fila, cola.exportado, CONFIG_EPS.ENCABEZADOS_OPERATIVOS.exportado, periodo, advertencias)
    }
  };
}

/**
 * Estudiante en una fuente anterior ya leída: sólo materias y valoraciones.
 * Si el DNI no puede ubicarse, se advierte y se continúa con la principal.
 */
function leerFuenteHistorica_(fuente, dniBuscado, advertencias) {
  try {
    const grilla = leerGrillaEstudiante_(fuente, dniBuscado, advertencias);
    return {
      periodo: fuente.periodo,
      materiasPorClave: grilla.materiasPorClave,
      ordenMaterias: grilla.ordenMaterias,
      ordenGrilla: grilla.ordenGrilla,
      nombresGrilla: grilla.nombresGrilla
    };
  } catch (err) {
    advertencias.push(err.message + ' Se continuará con la fuente principal.');
    return null;
  }
}

// Única lectura de una grilla fuente: sólo se consultan valores, nunca se escribe.
function abrirGrillaFuente_(fuenteIngresada, hojaOrigen, periodo) {
  const idFuente = extraerSpreadsheetId_(fuenteIngresada);
  const libroFuente = SpreadsheetApp.openById(idFuente);
  const hoja = libroFuente.getSheetByName(hojaOrigen);
  if (!hoja) throw new Error('No se encontró la hoja ' + hojaOrigen + ' en la fuente de ' + periodo + '.');

  const valores = hoja.getDataRange().getDisplayValues();
  return { periodo, hojaOrigen, valores, estructura: detectarEstructuraFuente_(valores, periodo) };
}

function leerGrillaEstudiante_(fuente, dniBuscado, advertencias) {
  const valores = fuente.valores;
  const estructura = fuente.estructura;
  const periodo = fuente.periodo;
  const filas = buscarFilasPorDni_(valores, dniBuscado);

  if (filas.length === 0) throw new Error('No se encontró el DNI en la fuente de ' + periodo + '.');
  if (filas.length > 1) throw new Error('El DNI aparece más de una vez en la fuente de ' + periodo + '.');

  const materias = leerMateriasEstudiante_(valores, filas[0], estructura, periodo, advertencias);
  return {
    valores,
    estructura,
    fila: filas[0],
    materiasPorClave: materias.materiasPorClave,
    ordenMaterias: materias.ordenMaterias,
    ordenGrilla: estructura.materias.map(m => m.clave),
    nombresGrilla: Object.fromEntries(estructura.materias.map(m => [m.clave, m.nombre]))
  };
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

  // Los encabezados de la cola nunca son materias, aunque aparezcan antes de ORIENTACIÓN.
  const encabezadosCola = Object.values(CONFIG_EPS.ENCABEZADOS_GENERALES)
    .concat(Object.values(CONFIG_EPS.ENCABEZADOS_OPERATIVOS))
    .map(normalizarTexto_);

  // Formación Profesional: el encabezado de grupo está en fila 1 y cada materia en fila 2.
  const materias = [];
  const columnasPorClave = {};
  for (let c = 3; c < inicioCola; c++) {
    const nombre = texto_(encabezado2[c]) || texto_(encabezado1[c]);
    if (!nombre) continue;
    if (/^EPS\d+$/.test(normalizarTexto_(nombre))) continue;
    if (encabezadosCola.includes(normalizarTexto_(nombre))) {
      throw new Error('El encabezado "' + nombre + '" (columna ' + letraColumna_(c) + ') es un campo general u operativo y está antes de ORIENTACIÓN' + enFuente);
    }
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

  return { materias, inicioCola };
}

/**
 * Columnas de la cola de la fuente principal, identificadas por encabezado exacto.
 * Un campo general sin encabezado identificable se advierte y queda vacío: no se deduce por posición.
 * Las columnas operativas pueden faltar en grillas anteriores a su estandarización, sin advertencia.
 */
function detectarCamposCola_(valores, estructura, periodo, advertencias) {
  const encabezado1 = valores[0] || [];
  const encabezado2 = valores[1] || [];
  const maxCols = Math.max(encabezado1.length, encabezado2.length);
  const enFuente = ' en la fuente de ' + periodo + '.';

  const buscar = (encabezado, advertirAusencia) => {
    const buscado = normalizarTexto_(encabezado);
    const columnas = [];
    for (let c = estructura.inicioCola + 1; c < maxCols; c++) {
      if (normalizarTexto_(encabezado1[c]) === buscado || normalizarTexto_(encabezado2[c]) === buscado) columnas.push(c);
    }
    if (columnas.length > 1) {
      advertencias.push('El encabezado "' + encabezado + '" aparece más de una vez (columnas ' + columnas.map(letraColumna_).join(', ') + ')' + enFuente + ' No se leerá ese campo.');
      return undefined;
    }
    if (columnas.length === 0 && advertirAusencia) {
      advertencias.push('No se encontró el encabezado "' + encabezado + '" después de ORIENTACIÓN' + enFuente + ' Ese campo quedará vacío en el informe.');
    }
    return columnas[0];
  };

  const cola = { orientacion: estructura.inicioCola };
  Object.keys(CONFIG_EPS.ENCABEZADOS_GENERALES).forEach(campo => {
    cola[campo] = buscar(CONFIG_EPS.ENCABEZADOS_GENERALES[campo], true);
  });
  Object.keys(CONFIG_EPS.ENCABEZADOS_OPERATIVOS).forEach(campo => {
    cola[campo] = buscar(CONFIG_EPS.ENCABEZADOS_OPERATIVOS[campo], false);
  });
  return cola;
}

/**
 * LISTO PARA EXPORTAR y EXPORTADO: valores canónicos SÍ / NO.
 * Devuelve 'SÍ', 'NO' o null (columna ausente, celda vacía o valor no admitido).
 */
function leerEstadoOperativo_(valores, fila, columna, encabezado, periodo, advertencias) {
  if (columna === undefined) return null;
  const enFuente = ' en la fuente de ' + periodo + '.';
  const crudo = obtenerValor_(valores, fila, columna);
  if (!crudo) {
    advertencias.push(encabezado + ' está vacío' + enFuente + ' Se esperaba SÍ o NO.');
    return null;
  }
  const canonico = crudo.normalize('NFC');
  if (canonico === 'SÍ' || canonico === 'NO') return canonico;

  const norm = normalizarTexto_(crudo);
  if (norm === 'SI' || norm === 'NO') {
    const estado = norm === 'SI' ? 'SÍ' : 'NO';
    advertencias.push(encabezado + ' tiene "' + crudo + '"' + enFuente + ' Se interpretó como ' + estado + ', pero la forma canónica es "' + estado + '".');
    return estado;
  }
  advertencias.push(encabezado + ' tiene un valor no admitido' + enFuente + ' Valor: "' + crudo + '". Se esperaba SÍ o NO; el estado queda sin determinar.');
  return null;
}

function buscarFilasPorDni_(valores, dniBuscado) {
  const filas = [];
  for (let r = 0; r < valores.length; r++) {
    if (normalizarDni_(valores[r][2]) === dniBuscado) filas.push(r);
  }
  return filas;
}

function leerMateriasEstudiante_(valores, fila, estructura, periodo, advertencias) {
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

  return { materiasPorClave, ordenMaterias };
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
  completarInforme_(salida, estudiante, materias, cicloLectivo, advertencias);
}

// Completa una copia de Plantilla: la usan la generación individual y cada hoja de la generación por curso.
function completarInforme_(salida, estudiante, materias, cicloLectivo, advertencias) {
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

// Nivel / curso como marcador de la grilla: "EPS2" o "2" → "EPS2".
function normalizarNivel_(valor) {
  const match = normalizarTexto_(valor).replace(/\s+/g, '').match(/^(?:EPS)?(\d+)$/);
  return match ? 'EPS' + Number(match[1]) : '';
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
